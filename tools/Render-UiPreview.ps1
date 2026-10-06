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
    [string]$OutDir
)

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path $repoRoot ("ui-preview\" + $Theme.ToLower()) }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$global:UiShotDir = (Resolve-Path $OutDir).Path
$global:UiShotIndex = 0
$global:UiShotName = 'okno'
$global:UiPending = New-Object System.Collections.ArrayList

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

function global:Save-UiWindowShot {
    param([System.Windows.Window]$Win)
    try {
        $name = $global:UiShotName
        # Okno Ustawień: rozwijamy wszystkie sekcje, żeby było widać wszystkie pola.
        foreach ($panelName in 'panelSrc', 'panelDom', 'panelTv', 'panelAv', 'panelWifi') {
            $panel = $Win.FindName($panelName)
            if ($null -ne $panel) { $panel.Visibility = [System.Windows.Visibility]::Visible }
        }
        $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
        Save-UiElementPng -Element $root -Background $null -Name $name
        # Jeśli okno ma przewijaną zawartość dłuższą niż widok - dodatkowo cała zawartość.
        $queue = New-Object System.Collections.Queue
        $queue.Enqueue($root)
        while ($queue.Count -gt 0) {
            $v = $queue.Dequeue()
            if ($v -is [System.Windows.Controls.ScrollViewer] -and $v.Content -is [System.Windows.FrameworkElement] -and $v.ExtentHeight -gt ($v.ViewportHeight + 40) -and $v.ViewportHeight -gt 100) {
                Save-UiElementPng -Element $v.Content -Background $Win.Background -Name "$name-cala-zawartosc"
                continue
            }
            for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) {
                $queue.Enqueue([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
            }
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
Invoke-UiShot 'programy' { Show-ProgramsManager -config $cfg }
Invoke-UiShot 'edycja-programu' { Show-ProgramEditDialog -IsNew $false -ProgramName 'Google Chrome' -ProgramData ([PSCustomObject]@{ Enabled = $true; FileName = 'Google.Chrome'; SilentArgs = ''; DownloadUrl = 'https://dl.google.com/chrome/install/ChromeStandaloneSetup64.exe' }) }
Invoke-UiShot 'rejestr' { Show-RegistryManager -config $cfg }
Invoke-UiShot 'edycja-rejestru' { Show-RegistryEditDialog -IsNew $false -RegData ([PSCustomObject]@{ Path = 'HKLM:\SOFTWARE\MojaFirma'; Name = 'WdrozenieZakonczone'; Value = '1'; PropertyType = 'DWord' }) }
Invoke-UiShot 'domyslne-zadania' { Show-DefaultTasksEditor -config $cfg }
Invoke-UiShot 'skrypty' { Show-PostInstallScriptsManager -config $cfg }
Invoke-UiShot 'edycja-skryptu' { Show-ScriptEditDialog -IsNew $true -ScriptPath 'SkryptDrukarki.ps1' }
Invoke-UiShot 'profile' { Show-ProfilesManager -config $cfg }
Invoke-UiShot 'edycja-profilu' { Show-ProfileEditDialog -IsNew $false -ProfileName 'Standard' -ProfileApps @('7-Zip', 'Google Chrome') -config $cfg }

Write-Host "Gotowe: $global:UiShotIndex plików w $global:UiShotDir"
