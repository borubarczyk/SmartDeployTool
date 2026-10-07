# Renderuje podgląd wszystkich okien Smart Tool for Deployment do plików PNG - bez klikania.
#
# Skrypt wczytuje narzędzie w trybie testowym (bez logowania i bez pokazywania okna głównego),
# a następnie po kolei otwiera każde okno (główne, komunikaty, Ustawienia, logi, deinstalator...)
# z prawdziwymi danymi. Globalny handler zdarzenia Loaded renderuje każde otwarte okno do PNG
# i od razu je zamyka, więc wywołania ShowDialog() nie czekają na użytkownika.
#
# Użycie (Windows PowerShell 5.1, wątek STA):
#   powershell -NoProfile -STA -ExecutionPolicy Bypass -File tools\Render-UiPreview.ps1 -Theme Dark
#   powershell -NoProfile -STA -ExecutionPolicy Bypass -File tools\Render-UiPreview.ps1 -Theme Light
# Wynik: ui-preview\<motyw>\NN-nazwa.png (uruchamiane też automatycznie w GitHub Actions,
# workflow "UI Preview" - pliki są do pobrania jako artefakt).
param(
    [ValidateSet('Dark', 'Light')][string]$Theme = 'Dark',
    [string]$OutDir,
    # Wypisuje do konsoli mapę układu (typ, nazwa, tekst, pozycja i rozmiar kontrolek) każdego okna.
    [switch]$DumpLayout,
    # Kończy skrypt kodem 1, jeśli w którymkolwiek oknie wykryto ucięty tekst.
    [switch]$FailOnClipping
)

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path $repoRoot ("ui-preview\" + $Theme.ToLower()) }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$global:UiShotDir = (Resolve-Path $OutDir).Path
$global:UiShotIndex = 0
$global:UiShotName = 'okno'
$global:UiPending = New-Object System.Collections.ArrayList
$global:UiClipping = New-Object System.Collections.ArrayList
$global:UiDumpLayout = [bool]$DumpLayout

# Tryb testowy: narzędzie nie pokaże okien logowania/powitania ani nie wywoła ShowDialog okna głównego.
$global:PesterTesting = $true
. (Join-Path $repoRoot 'SmartToolforDeployment.ps1')

function global:Save-UiElementPng {
    param([System.Windows.FrameworkElement]$Element, [System.Windows.Media.Brush]$Background, [string]$Name)
    $Element.UpdateLayout()
    $w = [int][Math]::Ceiling($Element.ActualWidth)
    $h = [int][Math]::Ceiling($Element.ActualHeight)
    if ($w -le 0 -or $h -le 0) { Write-Host "SKIP $Name (rozmiar $w x $h)"; return }
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap -ArgumentList $w, $h, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
    if ($null -ne $Background) {
        $dv = New-Object System.Windows.Media.DrawingVisual
        $dc = $dv.RenderOpen()
        $dc.DrawRectangle($Background, $null, (New-Object System.Windows.Rect 0, 0, $w, $h))
        $dc.Close()
        $rtb.Render($dv)
    }
    $rtb.Render($Element)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $global:UiShotIndex++
    $safe = ($Name -replace '[^\w\-]', '_')
    $file = Join-Path $global:UiShotDir ('{0:D2}-{1}.png' -f $global:UiShotIndex, $safe)
    $fs = [System.IO.File]::Create($file)
    try { $enc.Save($fs) } finally { $fs.Close() }
    Write-Host ("OK   {0} ({1} x {2})" -f (Split-Path $file -Leaf), $w, $h)
}

# Wykrywanie uciętego tekstu: dla każdego widocznego TextBlocka liczymy naturalny rozmiar tekstu
# (FormattedText) i porównujemy z miejscem, które dostał - zarówno sam TextBlock, jak i każdy jego
# rodzic, który przycina zawartość (np. przycisk ze sztywną wysokością albo zbyt wąska kolumna).
# Przycinanie przez ScrollViewer (przewijanie) jest zamierzone i pomijane.
function global:Get-UiElementLabel {
    param($Element)
    $label = $Element.GetType().Name
    if ($Element -is [System.Windows.FrameworkElement] -and -not [string]::IsNullOrWhiteSpace($Element.Name)) { $label += "#$($Element.Name)" }
    return $label
}

function global:Find-UiClipping {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
    $pixelsPerDip = 1.0
    try { $pixelsPerDip = [System.Windows.Media.VisualTreeHelper]::GetDpi($Win).PixelsPerDip } catch {}
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $v = $stack.Pop()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) {
            $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
        if ($v -isnot [System.Windows.Controls.TextBlock]) { continue }
        $tb = $v
        if ([string]::IsNullOrWhiteSpace($tb.Text) -or $tb.ActualWidth -le 0) { continue }

        # Najbliższy "właściciel" tekstu do opisu (przycisk, pole, etykieta...).
        $owner = $null
        $p = [System.Windows.Media.VisualTreeHelper]::GetParent($tb)
        while ($null -ne $p -and $null -eq $owner) {
            if ($p -is [System.Windows.Controls.Control]) { $owner = $p }
            $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p)
        }
        $ownerLabel = if ($null -ne $owner) { Get-UiElementLabel $owner } else { 'okno' }
        $textShort = ($tb.Text -replace '\s+', ' ').Trim()
        if ($textShort.Length -gt 60) { $textShort = $textShort.Substring(0, 60) + '…' }

        # 1) Czy tekst mieści się w samym TextBlocku.
        $typeface = New-Object System.Windows.Media.Typeface -ArgumentList $tb.FontFamily, $tb.FontStyle, $tb.FontWeight, $tb.FontStretch
        $ft = New-Object System.Windows.Media.FormattedText -ArgumentList $tb.Text, ([System.Globalization.CultureInfo]::CurrentUICulture), $tb.FlowDirection, $typeface, $tb.FontSize, ([System.Windows.Media.Brushes]::Black), $pixelsPerDip
        if (-not [double]::IsNaN($tb.LineHeight)) { $ft.LineHeight = $tb.LineHeight }
        $innerW = $tb.ActualWidth - $tb.Padding.Left - $tb.Padding.Right
        $innerH = $tb.ActualHeight - $tb.Padding.Top - $tb.Padding.Bottom
        if ($tb.TextWrapping -ne [System.Windows.TextWrapping]::NoWrap) {
            # Zawijanie liczymy tylko, gdy tekst wyraźnie nie mieści się w jednej linii - FormattedText
            # mierzy emoji (czcionka zastępcza) o 1-3 px inaczej niż TextBlock, co dawało fałszywe alarmy.
            if ($innerW -gt 0 -and $ft.WidthIncludingTrailingWhitespace -gt $innerW + 4) { $ft.MaxTextWidth = $innerW }
            if ($ft.Height -gt $innerH + 1.5) {
                [void]$results.Add(("{0} | {1} | '{2}' | tekst wyższy niż miejsce: {3:N0}px > {4:N0}px" -f $WindowName, $ownerLabel, $textShort, $ft.Height, $innerH))
            }
        } else {
            if ($tb.TextTrimming -eq [System.Windows.TextTrimming]::None -and $ft.WidthIncludingTrailingWhitespace -gt $innerW + 1.5) {
                [void]$results.Add(("{0} | {1} | '{2}' | ucięty w poziomie: tekst {3:N0}px, miejsce {4:N0}px" -f $WindowName, $ownerLabel, $textShort, $ft.WidthIncludingTrailingWhitespace, $innerW))
            }
            if ($ft.Height -gt $innerH + 1.5) {
                [void]$results.Add(("{0} | {1} | '{2}' | ucięty w pionie: tekst {3:N0}px, miejsce {4:N0}px" -f $WindowName, $ownerLabel, $textShort, $ft.Height, $innerH))
            }
        }

        # 2) Czy któryś z rodziców nie przycina TextBlocka (do najbliższego ScrollViewera).
        $tbRect = New-Object System.Windows.Rect 0, 0, $tb.ActualWidth, $tb.ActualHeight
        $a = [System.Windows.Media.VisualTreeHelper]::GetParent($tb)
        while ($null -ne $a -and $a -ne $root) {
            if ($a -is [System.Windows.Controls.ScrollContentPresenter]) { break }
            if ($a -is [System.Windows.FrameworkElement]) {
                $clip = [System.Windows.Controls.Primitives.LayoutInformation]::GetLayoutClip($a)
                $clipRect = $null
                if ($null -ne $clip) { $clipRect = $clip.Bounds }
                elseif ($a.ClipToBounds) { $clipRect = New-Object System.Windows.Rect 0, 0, $a.ActualWidth, $a.ActualHeight }
                if ($null -ne $clipRect -and -not $clipRect.IsEmpty) {
                    try {
                        $r = $tb.TransformToAncestor($a).TransformBounds($tbRect)
                        $cutX = [Math]::Max(0, $clipRect.Left - $r.Left) + [Math]::Max(0, $r.Right - $clipRect.Right)
                        $cutY = [Math]::Max(0, $clipRect.Top - $r.Top) + [Math]::Max(0, $r.Bottom - $clipRect.Bottom)
                        if ($cutX -gt 1.5 -or $cutY -gt 1.5) {
                            [void]$results.Add(("{0} | {1} | '{2}' | przycięty przez {3}: brakuje {4:N0}px w poziomie, {5:N0}px w pionie" -f $WindowName, $ownerLabel, $textShort, (Get-UiElementLabel $a), $cutX, $cutY))
                            break
                        }
                    } catch {}
                }
            }
            $a = [System.Windows.Media.VisualTreeHelper]::GetParent($a)
        }
    }
    return ,$results
}

# Tekstowa "mapa" okna: widoczne kontrolki z pozycją i rozmiarem względem okna.
function global:Write-UiLayoutMap {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
    Write-Host ("MAP  == {0} ({1:N0} x {2:N0}) ==" -f $WindowName, $root.ActualWidth, $root.ActualHeight)
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $v = $stack.Pop()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        $isControl = $v -is [System.Windows.Controls.Button] -or $v -is [System.Windows.Controls.TextBox] -or $v -is [System.Windows.Controls.PasswordBox] -or
            $v -is [System.Windows.Controls.ComboBox] -or $v -is [System.Windows.Controls.CheckBox] -or $v -is [System.Windows.Controls.ListView] -or
            $v -is [System.Windows.Controls.ListBox] -or $v -is [System.Windows.Controls.ProgressBar]
        $isLooseText = $v -is [System.Windows.Controls.TextBlock] -and $null -eq $v.TemplatedParent -and -not [string]::IsNullOrWhiteSpace($v.Text)
        if ($isControl -or $isLooseText) {
            try {
                $r = $v.TransformToAncestor($root).TransformBounds((New-Object System.Windows.Rect 0, 0, $v.ActualWidth, $v.ActualHeight))
                $text = ''
                if ($v -is [System.Windows.Controls.TextBlock]) { $text = $v.Text }
                elseif ($v -is [System.Windows.Controls.ContentControl] -and $v.Content -is [string]) { $text = $v.Content }
                elseif ($v -is [System.Windows.Controls.TextBox]) { $text = $v.Text }
                elseif ($v -is [System.Windows.Controls.ComboBox]) { $text = [string]$v.Text }
                $text = ($text -replace '\s+', ' ').Trim()
                if ($text.Length -gt 50) { $text = $text.Substring(0, 50) + '…' }
                Write-Host ("MAP  {0,-34} x={1,4:N0} y={2,4:N0} w={3,4:N0} h={4,3:N0} fs={5} '{6}'" -f (Get-UiElementLabel $v), $r.X, $r.Y, $r.Width, $r.Height, $v.FontSize, $text)
            } catch {}
            if ($isControl -and $v -isnot [System.Windows.Controls.ListView] -and $v -isnot [System.Windows.Controls.ListBox]) { continue }
        }
        for ($i = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v) - 1; $i -ge 0; $i--) {
            $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
    }
}

# Jeden "widok" okna: lint ucinania, mapa układu, PNG i (gdy jest przewijanie) PNG całej zawartości.
function global:Save-UiView {
    param([System.Windows.Window]$Win, [string]$Name)
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
    $root.UpdateLayout()
    foreach ($issue in (Find-UiClipping -Win $Win -WindowName $Name)) {
        [void]$global:UiClipping.Add($issue)
        Write-Host "CLIP $issue"
    }
    if ($global:UiDumpLayout) { Write-UiLayoutMap -Win $Win -WindowName $Name }
    Save-UiElementPng -Element $root -Background $null -Name $Name
    # Jeśli okno ma przewijaną zawartość dłuższą niż widok - dodatkowo cała zawartość.
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($root)
    while ($queue.Count -gt 0) {
        $v = $queue.Dequeue()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        if ($v -is [System.Windows.Controls.ScrollViewer] -and $v.ExtentHeight -gt ($v.ViewportHeight + 1) -and $v.ViewportHeight -gt 0 -and $v.TemplatedParent -isnot [System.Windows.Controls.Primitives.TextBoxBase] -and $v.TemplatedParent -isnot [System.Windows.Controls.ItemsControl]) {
            # Zawartość dłuższa niż widoczny obszar = trzeba przewijać (np. lista zadań w oknie głównym).
            Write-Host ("SCROLL {0} | {1} | zawartość {2:N0}px w widoku {3:N0}px" -f $Name, (Get-UiElementLabel $v.Content), $v.ExtentHeight, $v.ViewportHeight)
        }
        if ($v -is [System.Windows.Controls.ScrollViewer] -and $v.Content -is [System.Windows.FrameworkElement] -and $v.ExtentHeight -gt ($v.ViewportHeight + 40) -and $v.ViewportHeight -gt 100) {
            Save-UiElementPng -Element $v.Content -Background $Win.Background -Name "$Name-cala-zawartosc"
            continue
        }
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) {
            $queue.Enqueue([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
    }
}

# Okno z zakładkami: każda zakładka osobno (w drzewie wizualnym jest tylko wybrana).
function global:Save-UiTabbedView {
    param([System.Windows.Window]$Win, [string]$Name)
    $tabs = $null
    foreach ($tabName in 'tabSettings', 'tabSysInfo') { if ($null -eq $tabs) { $tabs = $Win.FindName($tabName) } }
    if ($null -eq $tabs) { Save-UiView -Win $Win -Name $Name; return }
    for ($i = 0; $i -lt $tabs.Items.Count; $i++) {
        $tabs.SelectedIndex = $i
        $viewName = "{0}-{1:D2}" -f $Name, ($i + 1)
        Write-Host ("TAB  {0} = {1}" -f $viewName, $tabs.Items[$i].Header)
        Save-UiView -Win $Win -Name $viewName
    }
    $tabs.SelectedIndex = 0
}

function global:Save-UiWindowShot {
    param([System.Windows.Window]$Win)
    try {
        $name = $global:UiShotName
        Save-UiTabbedView -Win $Win -Name $name
        # Okno główne i Ustawienia dodatkowo w minimalnym rozmiarze - tam najłatwiej coś uciąć.
        if ($name -in @('okno-glowne', 'ustawienia') -and $Win.MinWidth -gt 0 -and $Win.MinHeight -gt 0) {
            $Win.Width = $Win.MinWidth
            $Win.Height = $Win.MinHeight
            Save-UiTabbedView -Win $Win -Name "$name-min"
        }
    } catch {
        Write-Host "ERR  $($global:UiShotName): $($_.Exception.Message)"
    }
}

# Każde okno po wczytaniu (Loaded) czeka, aż dispatcher skończy pracę (dane, układ), po czym jest
# renderowane do PNG i zamykane.
$global:UiLoadedHandler = [System.Windows.RoutedEventHandler]{
    param($sender, $e)
    [void]$global:UiPending.Add($sender)
    [void]$sender.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::ApplicationIdle, [Action]{
        while ($global:UiPending.Count -gt 0) {
            $w = $global:UiPending[0]
            $global:UiPending.RemoveAt(0)
            Save-UiWindowShot -Win $w
            try { $w.Close() } catch {}
        }
    })
}
[System.Windows.EventManager]::RegisterClassHandler([System.Windows.Window], [System.Windows.FrameworkElement]::LoadedEvent, $global:UiLoadedHandler)

# Przetwarza zdarzenia dispatchera, aż warunek będzie spełniony (dla okien niemodalnych).
function global:Wait-UiCondition {
    param([scriptblock]$Until, [int]$TimeoutMs = 20000)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not (& $Until) -and $sw.ElapsedMilliseconds -lt $TimeoutMs) {
        $frame = New-Object System.Windows.Threading.DispatcherFrame
        [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::SystemIdle, [Action]{ $frame.Continue = $false })
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        Start-Sleep -Milliseconds 20
    }
}

function Invoke-UiShot {
    param([string]$Name, [scriptblock]$Action)
    $global:UiShotName = $Name
    try { & $Action | Out-Null } catch { Write-Host "ERR  ${Name}: $($_.Exception.Message)" }
}

# --- Motyw ---
$script:isDarkTheme = ($Theme -eq 'Dark')
Set-AppTheme
foreach ($w in @($authWindow, $welcomeWindow)) { if ($null -ne $w) { Update-WindowTheme $w } }

# --- Przykładowe wpisy logu (różne rodzaje), żeby przeglądarka logów nie była pusta ---
Write-Log "Rozpoczęto walidację konfiguracji i plików przed wdrożeniem..."
Write-Log "Mało wolnego miejsca na dysku C: (12.4 GB). Instalacja dużych programów może zakończyć się błędem."
Write-Log "[DRY-RUN] Zainstalowano by Google Chrome przez Winget: winget install --id `"Google.Chrome`" -e --silent"
Write-Log "Instalator 7-Zip zakończył się błędem (kod wyjścia: 1603)." -IsError
Write-Log "7-Zip zainstalowany." -Context "Automat"
Write-Log "Zastosowano profil wdrożenia: Standard (Wybrano programów: 3)" -Context "Użytkownik"

# --- Okna ---
Get-AppSelection
Load-Profiles
$cfg = Get-Config

Invoke-UiShot 'logowanie' { $authWindow.ShowDialog() }
Invoke-UiShot 'powitanie' { $welcomeWindow.ShowDialog() }
Invoke-UiShot 'okno-glowne' { $Window.ShowDialog() }

Invoke-UiShot 'komunikat-pytanie' { Show-ThemedMessageBox -Message "Czy na pewno chcesz rozpocząć konfigurację?" -Title "Potwierdzenie" -Button "YesNo" -Image "Question" }
Invoke-UiShot 'komunikat-blad' { Show-ThemedMessageBox -Message ("Wykryto bledy walidacji:`r`n- Brak sciezki sieciowej (InstallSourcePaths.network).`r`n- Nie znaleziono pliku dla 'Google Chrome': \\serwer\udzial\ChromeSetup.exe`r`n- Brak DomainJoin.DomainName w config.json.") -Title "Walidacja" -Button "OK" -Image "Error" }
Invoke-UiShot 'komunikat-ostrzezenie' { Show-ThemedMessageBox -Message "Czy na pewno chcesz przerwać wdrożenie?" -Title "Przerwij" -Button "YesNoCancel" -Image "Warning" }
Invoke-UiShot 'wpisz-tekst' { Show-InputDialog -Title "Zmiana nazwy komputera" -Message "Podaj nową nazwę komputera (maks. 15 znaków). Obecna nazwa: DESKTOP-AB12CD3" -DefaultText "PC-5CG1234XYZ" }
Invoke-UiShot 'wpisz-haslo' { Show-InputDialog -Title "Konto lokalnego administratora" -Message "Podaj hasło dla nowego konta 'LokalnyIT' (co najmniej 8 znaków):" -Password }
Invoke-UiShot 'pin' { Show-PinPrompt }
Invoke-UiShot 'wybor-aplikacji' { Show-AppSelectionWindow }
Invoke-UiShot 'informacja' { Show-CustomInfoDialog -Title "Zakończono" -Message "Przetwarzanie deinstalacji zakończone.`n`nPoprawnie odinstalowane aplikacje zostały automatycznie usunięte z listy.`nJeśli instalator zwrócił błąd, aplikacja pozostała na liście oznaczona krzyżykiem (❌)." -ShowCopy -HtmlData "<html></html>" }
Invoke-UiShot 'potwierdzenie-deinstalacji' { Show-UninstallConfirmDialog -AppCount 3 -AppListText "- 7-Zip 23.01 (x64)`n- Google Chrome`n- Microsoft OneDrive" }
Invoke-UiShot 'logi' { Show-LogWindow; Wait-UiCondition -Until { $null -eq $script:ActiveLogWindow -or -not $script:ActiveLogWindow.IsVisible } }
Invoke-UiShot 'szczegoly-wpisu-logu' { Show-LogEntryDetail -Entry ([PSCustomObject]@{ RodzajText = '❌ Błąd'; DataText = '06.10.2026 09:12:44'; Kontekst = 'Automat'; Informacja = "Instalator 7-Zip zakończył się błędem (kod wyjścia: 1603).`r`nSzczegóły: Fatal error during installation."; RowColorHex = (Get-ThemeColor 'ThemeDanger') }) }
Invoke-UiShot 'informacje-o-systemie' { Show-SystemInfoWindow }
Invoke-UiShot 'deinstalator' { Show-SoftwareUninstaller }
Invoke-UiShot 'ustawienia' { Show-ConfigEditor }
Invoke-UiShot 'edycja-programu' { Show-ProgramEditDialog -IsNew $false -ProgramName 'Google Chrome' -ProgramData ([PSCustomObject]@{ Enabled = $true; FileName = 'Google.Chrome'; SilentArgs = ''; DownloadUrl = 'https://dl.google.com/chrome/install/ChromeStandaloneSetup64.exe' }) }
Invoke-UiShot 'edycja-rejestru' { Show-RegistryEditDialog -IsNew $false -RegData ([PSCustomObject]@{ Path = 'HKLM:\SOFTWARE\MojaFirma'; Name = 'WdrozenieZakonczone'; Value = '1'; PropertyType = 'DWord' }) }
Invoke-UiShot 'edycja-skryptu' { Show-ScriptEditDialog -IsNew $true -ScriptPath 'SkryptDrukarki.ps1' }
Invoke-UiShot 'edycja-profilu' { Show-ProfileEditDialog -IsNew $false -ProfileName 'Standard' -ProfileApps @('7-Zip', 'Google Chrome') -config $cfg }

Write-Host "Gotowe: $global:UiShotIndex plików w $global:UiShotDir"
$report = Join-Path $global:UiShotDir 'uciety-tekst.txt'
if ($global:UiClipping.Count -gt 0) {
    $global:UiClipping | Set-Content -Path $report -Encoding UTF8
    Write-Host "UCIĘTY TEKST: $($global:UiClipping.Count) miejsc (lista w $report)"
    if ($FailOnClipping) { exit 1 }
} else {
    'Brak uciętego tekstu.' | Set-Content -Path $report -Encoding UTF8
    Write-Host "Brak uciętego tekstu."
}
