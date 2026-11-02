# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Render-StdUiPreview.ps1 - podgląd wszystkich okien STD 4 w plikach PNG, bez klikania (SPEC 9).
#
# Skrypt wczytuje program w trybie podglądu (SmartToolforDeployment.ps1 -PreviewMode: moduł STD i okna,
# bez podnoszenia uprawnień, bez kopii i bez pokazywania okien), a potem po kolei pokazuje każde okno
# z przykładowymi danymi (-PreviewData / -PreviewState, bez serwera, sieci i silnika):
#   - „Połącz z siecią” (lista sieci, sieć otwarta i firmowa, błąd, połączono, brak sieci);
#   - „Logowanie” (adres z bootstrap.json, puste pola z adresem do wpisania, błąd z pracą offline, w toku);
#   - okno główne w każdym stanie: wybór, brak wyboru, błąd konfiguracji, plan („Sprawdź”), wdrożenie
#     w trakcie (z dziennikiem), pauza, błąd, koniec z podsumowaniem, koniec z restartem, powiadomienia;
#   - potwierdzenie „Wdróż” (prawdziwe okno z listą kroków planu);
#   - okna komunikatu każdego rodzaju (informacja, pytanie, ostrzeżenie, błąd, sukces, długi tekst);
#   - ikonę STD (każdy rozmiar z pliku .ico).
# Globalny handler zdarzenia Loaded zapisuje każde pokazane okno do PNG, sprawdza je i zamyka, więc
# ShowDialog() nie czeka na użytkownika. Błąd jednego okna jest zapisywany jako ERR, a reszta idzie dalej.
#
# Okno główne, „Połącz z siecią” i „Logowanie” są zapisywane też w minimalnym rozmiarze okna oraz na ekranie
# 1366x768 przy skalowaniu 150% bez paska zadań (911 x 472 DIP) - tak jak w podglądzie STD 3.1.
#
# Użycie (Windows PowerShell 5.1, wątek STA):
#   powershell -NoProfile -STA -ExecutionPolicy Bypass -File v4\Tools\Render-StdUiPreview.ps1 -Theme Dark
#   powershell -NoProfile -STA -ExecutionPolicy Bypass -File v4\Tools\Render-StdUiPreview.ps1 -Theme Light -FailOnFindings
# Wynik: v4\ui-preview\<motyw>\NN-nazwa.png (folder jest w .gitignore). W GitHub Actions uruchamia go
# workflow „v4-ui-preview” - pliki są do pobrania jako artefakt v4-ui-preview.
#
# Kontrole (lista w konsoli, na końcu logu i w plikach obok PNG):
#   CLIP      ucięty tekst                                           -> uciety-tekst.txt
#   STALE     kolor z poprzedniego motywu po przełączeniu w locie    -> uwagi-motyw.txt
#   GLYPH     znak ikony (U+E000-U+F8FF) w zwykłej czcionce          -> uwagi-motyw.txt
#   EMOJI     emoji w tekście okna                                   -> uwagi-motyw.txt
#   CONTRAST  tekst za słabo odcina się od tła (< 3:1)               -> uwagi-motyw.txt
#   ERR       błąd budowania / renderowania okna, nieoczekiwane okno
#             (np. komunikat o błędzie), błąd w logu okien           -> bledy-renderowania.txt
# Ta sama uwaga w kilku oknach (np. nagłówek okna głównego w każdym stanie) jest na liście raz,
# z dopiskiem „także: ...”.
#
# Funkcje kontroli (Find-UiClipping, Find-UiStaleColors, Find-UiContrast, ...) są przeniesione z podglądu
# STD 3.1 (tools\Render-UiPreview.ps1) - ten skrypt go nie wczytuje.
param(
    [ValidateSet('Dark', 'Light')][string]$Theme = 'Dark',
    [string]$OutDir,
    # Wypisuje do konsoli mapę układu (typ, nazwa, tekst, pozycja i rozmiar kontrolek) każdego okna.
    [switch]$DumpLayout,
    # Kończy skrypt kodem 1, jeśli w którymkolwiek oknie wykryto ucięty tekst.
    [switch]$FailOnClipping,
    # Kończy skrypt kodem 1 przy jakiejkolwiek uwadze (CLIP, STALE, GLYPH, EMOJI, CONTRAST) albo błędzie (ERR); obejmuje -FailOnClipping.
    [switch]$FailOnFindings,
    # Tylko okna o nazwach pasujących do wzorców (np. 'okno-glowne*') - do szybkiego sprawdzania zmian.
    [string[]]$Only
)
if ($FailOnFindings) { $FailOnClipping = [switch]$true }

$pvV4Root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path (Join-Path $pvV4Root 'ui-preview') $Theme.ToLowerInvariant() }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

# Cały stan podglądu w jednym obiekcie globalnym: handlery okien (Loaded, timery) działają później,
# w pętli komunikatów ShowDialog(), i nie mogą zależeć od zmiennych lokalnych.
$global:StdPv = @{
    Theme         = $Theme
    Dir           = (Resolve-Path -LiteralPath $OutDir).ProviderPath
    Index         = 0
    DumpLayout    = [bool]$DumpLayout
    Only          = @($Only | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    Issues        = New-Object System.Collections.Specialized.OrderedDictionary   # klucz -> @{ Kind; Line; Windows }
    Errors        = New-Object 'System.Collections.Generic.List[string]'
    Log           = New-Object 'System.Collections.Generic.List[object]'
    Pending       = New-Object System.Collections.ArrayList
    Active        = $false
    ShotName      = ''
    ShotWindow    = $null
    Spec          = $null
    ChildName     = $null
    ChildSized    = $false
    ChildSeen     = $false
    Unexpected    = 0
    Shots         = 0
    Kinds         = @{}
    MinContrast   = 3.0
    SmallWidth    = 911
    SmallHeight   = 472
}

function global:Add-StdPvError {
    param([string]$Text)
    [void]$global:StdPv.Errors.Add($Text)
    Write-Host "ERR  $Text"
}

function global:Get-StdPvErrorText {
    # Komunikat wyjątku bez opakowania "Exception calling ...", w jednej linii.
    param($ErrorObject)
    $ex = $null
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ex = $ErrorObject.Exception } elseif ($ErrorObject -is [System.Exception]) { $ex = $ErrorObject }
    if ($null -eq $ex) { return ([string]$ErrorObject -replace '\s+', ' ').Trim() }
    while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
    $text = [string]$ex.Message
    if ($null -ne $ex.InnerException -and -not [string]::IsNullOrEmpty($ex.InnerException.Message) -and -not $text.Contains($ex.InnerException.Message)) { $text += ' ' + $ex.InnerException.Message }
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord] -and $null -ne $ErrorObject.InvocationInfo -and $ErrorObject.InvocationInfo.ScriptLineNumber -gt 0) {
        $text += ' (' + [System.IO.Path]::GetFileName([string]$ErrorObject.InvocationInfo.ScriptName) + ':' + $ErrorObject.InvocationInfo.ScriptLineNumber + ')'
    }
    return ($text -replace '\s+', ' ').Trim()
}

# Uwaga trafia na listę (i od razu do konsoli) raz. Ta sama uwaga w innym oknie (np. nagłówek okna głównego
# w każdym stanie) dopisuje tylko nazwę okna do „także: ...". $Line = "<okno> | <reszta>".
function global:Add-StdPvIssue {
    param([string]$Kind, [string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    $window = ''
    $rest = $Line
    $at = $Line.IndexOf(' | ')
    if ($at -gt 0) { $window = $Line.Substring(0, $at); $rest = $Line.Substring($at + 3) }
    $key = $Kind + ' | ' + $rest
    $issues = $global:StdPv.Issues
    if ($issues.Contains($key)) {
        $entry = $issues[$key]
        if (-not $entry.Windows.Contains($window)) { $entry.Windows.Add($window) }
        return
    }
    $windows = New-Object 'System.Collections.Generic.List[string]'
    $windows.Add($window)
    $issues[$key] = @{ Kind = $Kind; Line = ('{0} {1}' -f $Kind, $Line); Windows = $windows }
    Write-Host ('{0} {1}' -f $Kind, $Line)
}

# Uwaga z kontroli 3.1 ("KIND okno | ..."): rodzaj z pierwszego słowa.
function global:Add-UiFinding {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $space = $Text.IndexOf(' ')
    if ($space -le 0) { Add-StdPvIssue -Kind 'UWAGA' -Line $Text; return }
    Add-StdPvIssue -Kind $Text.Substring(0, $space) -Line $Text.Substring($space + 1)
}

function global:Get-StdPvIssueLines {
    # Lista uwag do raportu: -Kind CLIP albo uwagi motywu (wszystko poza CLIP).
    param([switch]$Clip)
    $wantClip = [bool]$Clip.IsPresent
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($entry in $global:StdPv.Issues.Values) {
        if ($wantClip -ne ($entry.Kind -eq 'CLIP')) { continue }
        $line = $entry.Line
        if ($entry.Windows.Count -gt 1) {
            $others = New-Object 'System.Collections.Generic.List[string]'
            for ($i = 1; $i -lt $entry.Windows.Count -and $i -le 8; $i++) { $others.Add($entry.Windows[$i]) }
            $line += ' | także: ' + ($others -join ', ')
            if ($entry.Windows.Count -gt 9) { $line += (' i {0} innych' -f ($entry.Windows.Count - 9)) }
        }
        $lines.Add($line)
    }
    return , $lines
}

function global:Write-StdPvTextFile {
    # Plik raportu: UTF-8 bez BOM (SPEC 0.1).
    param([string]$Path, [AllowEmptyCollection()][string[]]$Lines)
    $text = (@($Lines) -join [System.Environment]::NewLine) + [System.Environment]::NewLine
    [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

function Write-StdPvReport {
    # Pliki z uwagami, podsumowanie na końcu logu i (w GitHub Actions) adnotacja oraz podsumowanie zadania.
    $pv = $global:StdPv
    $clipLines = Get-StdPvIssueLines -Clip
    $findingLines = Get-StdPvIssueLines
    $errorLines = $pv.Errors.ToArray()
    $clipReport = Join-Path $pv.Dir 'uciety-tekst.txt'
    $findingsReport = Join-Path $pv.Dir 'uwagi-motyw.txt'
    $errorsReport = Join-Path $pv.Dir 'bledy-renderowania.txt'
    if ($clipLines.Count -gt 0) { Write-StdPvTextFile -Path $clipReport -Lines $clipLines } else { Write-StdPvTextFile -Path $clipReport -Lines @('Brak uciętego tekstu.') }
    if ($findingLines.Count -gt 0) { Write-StdPvTextFile -Path $findingsReport -Lines $findingLines } else { Write-StdPvTextFile -Path $findingsReport -Lines @('Brak uwag (STALE / GLYPH / EMOJI / CONTRAST).') }
    if ($errorLines.Count -gt 0) { Write-StdPvTextFile -Path $errorsReport -Lines $errorLines } else { Write-StdPvTextFile -Path $errorsReport -Lines @('Brak błędów renderowania.') }

    $count = @{ STALE = 0; GLYPH = 0; EMOJI = 0; CONTRAST = 0 }
    foreach ($entry in $pv.Issues.Values) { if ($count.ContainsKey($entry.Kind)) { $count[$entry.Kind]++ } }
    $summary = New-Object 'System.Collections.Generic.List[string]'
    $summary.Add(('==== Podgląd okien STD 4, motyw {0}: uwagi ====' -f $pv.Theme))
    $summary.Add(('PNG   zapisane pliki                 : {0} (okien: {1})' -f $pv.Index, $pv.Shots))
    $summary.Add(('CLIP  ucięty tekst                   : {0}' -f $clipLines.Count))
    $summary.Add(('STALE kolor z poprzedniego motywu    : {0}' -f $count.STALE))
    $summary.Add(('GLYPH ikona w zwykłej czcionce       : {0}' -f $count.GLYPH))
    $summary.Add(('EMOJI emoji w tekście okna           : {0}' -f $count.EMOJI))
    $summary.Add(('CONTRAST za słaby kontrast tekstu    : {0}' -f $count.CONTRAST))
    $summary.Add(('ERR   błąd przy renderowaniu okna    : {0}' -f $errorLines.Count))
    foreach ($line in $clipLines) { $summary.Add($line) }
    foreach ($line in $findingLines) { $summary.Add($line) }
    foreach ($line in $errorLines) { $summary.Add("ERR  $line") }
    foreach ($record in $pv.Log) {
        if ($record.Level -ne 'Warn') { continue }
        $summary.Add(('LOG  [{0}] {1}: {2} (okno: {3})' -f $record.Level, $record.Source, $record.Message, $record.Shot))
    }
    $summary.Add(('==== koniec uwag (pliki: {0}, {1}, {2}) ====' -f $clipReport, $findingsReport, $errorsReport))
    foreach ($line in $summary) { Write-Host $line }

    if ($env:GITHUB_ACTIONS -eq 'true') {
        $totals = 'CLIP {0}, STALE {1}, GLYPH {2}, EMOJI {3}, CONTRAST {4}, ERR {5}' -f $clipLines.Count, $count.STALE, $count.GLYPH, $count.EMOJI, $count.CONTRAST, $errorLines.Count
        if (($clipLines.Count + $findingLines.Count + $errorLines.Count) -gt 0) {
            Write-Host ('::warning title=Podgląd okien STD 4 ({0})::{1} - lista w logu kroku i w v4/ui-preview/{2}/' -f $pv.Theme, $totals, $pv.Theme.ToLowerInvariant())
        }
        if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_STEP_SUMMARY)) {
            try {
                $md = New-Object System.Text.StringBuilder
                [void]$md.AppendLine(('### Podgląd okien STD 4 - motyw {0}' -f $pv.Theme))
                [void]$md.AppendLine('')
                [void]$md.AppendLine(('PNG: {0} | {1}' -f $pv.Index, $totals))
                $details = New-Object 'System.Collections.Generic.List[string]'
                foreach ($line in $summary) { if ($line -match '^(CLIP|STALE|GLYPH|EMOJI|CONTRAST|ERR) ') { $details.Add($line) } }
                if ($details.Count -gt 0) {
                    [void]$md.AppendLine('')
                    [void]$md.AppendLine('```text')
                    $shown = 0
                    foreach ($line in $details) {
                        if ($shown -ge 300) { [void]$md.AppendLine(('... i {0} kolejnych (pełna lista w artefakcie)' -f ($details.Count - $shown))); break }
                        [void]$md.AppendLine($line)
                        $shown++
                    }
                    [void]$md.AppendLine('```')
                }
                [void]$md.AppendLine('')
                [System.IO.File]::AppendAllText($env:GITHUB_STEP_SUMMARY, $md.ToString(), (New-Object System.Text.UTF8Encoding($false)))
            } catch { Write-Host ('Nie udało się zapisać podsumowania zadania: {0}' -f $_.Exception.Message) }
        }
    }
    return ($clipLines.Count), ($findingLines.Count + $errorLines.Count)
}

function Exit-StdPvFatal {
    # Bez okien nie ma czego sprawdzać: raport z błędem i kod 1 (także bez -FailOnFindings).
    param([string]$Text)
    Add-StdPvError -Text $Text
    [void](Write-StdPvReport)
    if (-not [string]::IsNullOrEmpty($pvDataRoot)) { try { Remove-Item -LiteralPath $pvDataRoot -Recurse -Force -ErrorAction Stop } catch { } }
    exit 1
}

# ---------------------------------------------------------------------------------------------
# Warunki: Windows, wątek STA
# ---------------------------------------------------------------------------------------------

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
    Exit-StdPvFatal -Text 'Podgląd okien działa tylko w Windows (WPF, .NET Framework 4.8, Windows PowerShell 5.1).'
}
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA) {
    Exit-StdPvFatal -Text 'Okna WPF wymagają wątku STA - uruchom: powershell -NoProfile -STA -ExecutionPolicy Bypass -File v4\Tools\Render-StdUiPreview.ps1'
}

# Dane lokalne STD (prefs.json, cache, logi) w folderze tymczasowym: podgląd nie rusza %ProgramData%\STD.
$pvOldDataRoot = $env:STD_DATA_ROOT
$pvOldUiPreview = $env:STD_UI_PREVIEW
$pvDataRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('std-ui-preview-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
[void][System.IO.Directory]::CreateDirectory($pvDataRoot)
$env:STD_DATA_ROOT = $pvDataRoot
# Powiadomienia bez animacji (Show-StdToast): zrzut robimy od razu, animacja zostawiłaby krycie 0.
$env:STD_UI_PREVIEW = '1'

# ---------------------------------------------------------------------------------------------
# Program w trybie podglądu (moduł STD + okna, bez pokazywania)
# ---------------------------------------------------------------------------------------------

$pvEntry = Join-Path $pvV4Root 'SmartToolforDeployment.ps1'
try {
    . $pvEntry -PreviewMode
} catch {
    Exit-StdPvFatal -Text ('Nie udało się wczytać programu w trybie podglądu ({0}): {1}' -f $pvEntry, (Get-StdPvErrorText $_))
}
foreach ($pvCommand in @('New-StdWindow', 'New-StdMessageWindow', 'Show-StdMessage', 'Show-StdToast', 'Set-StdTheme', 'Show-StdConnectWindow', 'Show-StdLoginWindow', 'Show-StdMainWindow', 'Get-StdUiPreviewSample')) {
    if (-not (Get-Command -Name $pvCommand -ErrorAction SilentlyContinue)) {
        Exit-StdPvFatal -Text ('Po wczytaniu programu brakuje funkcji {0} (GUI\Std.UI.ps1 albo GUI\Screens).' -f $pvCommand)
    }
}

# Wpisy logu okien: ostrzeżenia trafiają do podsumowania, błędy (np. błąd w handlerze okna) - jako ERR.
Register-StdLogSink -Name 'StdUiPreview' -ScriptBlock {
    $pvRecord = $args[0]
    $pvLevel = [string]$pvRecord.Level
    if ($pvLevel -ne 'Warn' -and $pvLevel -ne 'Error') { return }
    $pvEntryLog = @{ Level = $pvLevel; Source = [string]$pvRecord.Source; Message = [string]$pvRecord.Message; Shot = [string]$global:StdPv.ShotName }
    [void]$global:StdPv.Log.Add($pvEntryLog)
    if ($pvLevel -eq 'Error') {
        Add-StdPvError -Text ('{0}: błąd w logu okien ({1}): {2}' -f $pvEntryLog.Shot, $pvEntryLog.Source, $pvEntryLog.Message)
    }
}

# ---------------------------------------------------------------------------------------------
# PNG
# ---------------------------------------------------------------------------------------------

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
    $global:StdPv.Index++
    $safe = ($Name -replace '[^\w\-]', '_')
    $file = Join-Path $global:StdPv.Dir ('{0:D2}-{1}.png' -f $global:StdPv.Index, $safe)
    $fs = [System.IO.File]::Create($file)
    try { $enc.Save($fs) } finally { $fs.Close() }
    Write-Host ('OK   {0} ({1} x {2})' -f (Split-Path $file -Leaf), $w, $h)
}

# ---------------------------------------------------------------------------------------------
# Kontrole wyglądu (przeniesione z podglądu STD 3.1)
# ---------------------------------------------------------------------------------------------

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
        $ownerLabel = 'okno'
        if ($null -ne $owner) { $ownerLabel = Get-UiElementLabel $owner }
        $textShort = ($tb.Text -replace '\s+', ' ').Trim()
        if ($textShort.Length -gt 60) { $textShort = $textShort.Substring(0, 60) + '…' }

        # 1) Czy tekst mieści się w samym TextBlocku.
        $typeface = New-Object System.Windows.Media.Typeface -ArgumentList $tb.FontFamily, $tb.FontStyle, $tb.FontWeight, $tb.FontStretch
        # Pomiar w trybie, w którym TextBlock się rysuje (okna STD: Display - szerokości liter
        # zaokrąglane do pikseli); domyślny Ideal różnił się o kilka px i dawał fałszywe alarmy.
        # Konstruktor z 9 argumentami (NumberSubstitution = $null) jest w .NET od 4.6.2.
        $mode = [System.Windows.Media.TextOptions]::GetTextFormattingMode($tb)
        $ft = New-Object System.Windows.Media.FormattedText -ArgumentList $tb.Text, ([System.Globalization.CultureInfo]::CurrentUICulture), $tb.FlowDirection, $typeface, $tb.FontSize, ([System.Windows.Media.Brushes]::Black), $null, $mode, $pixelsPerDip
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
    return , $results
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

# ---------- STALE / GLYPH / EMOJI / CONTRAST ----------
# Wzorzec emoji jak w podglądzie i testach STD 3.1; strzałki sortowania (U+25B2/25BC) celowo poza nim -
# to zwykłe kształty z Segoe UI.
$global:UiEmojiPattern = '[\u2300-\u23FF\u2600-\u27BF\u2B00-\u2BFF\u25B6\u25BE\u2139\uFE0F]|[\uD83C-\uDBFF][\uDC00-\uDFFF]'

# Widoczne elementy drzewa wizualnego (ukryte gałęzie pominięte). Zwracamy ,$lista - bez @() wokół
# listy generycznej (w PS 5.1 to błąd "Argument types do not match").
function global:Get-UiVisibleVisuals {
    param($Root)
    $list = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Stack
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $v = $stack.Pop()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        $list.Add($v)
        for ($i = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v) - 1; $i -ge 0; $i--) {
            $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
    }
    return , $list
}

function global:Get-UiShortText {
    param([string]$Text, [int]$Max = 50)
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max) + '…' }
    return $t
}

# Opis elementu do raportu: Typ#nazwa; element bez własnej nazwy z XAML okna (np. część szablonu
# albo wiersza listy) dostaje jeszcze najbliższego nazwanego przodka ("Border w ListView#lvSteps"),
# a TextBlock - swój tekst.
function global:Get-UiElementPath {
    param($Element)
    $label = Get-UiElementLabel $Element
    $isNamed = $Element -is [System.Windows.FrameworkElement] -and -not [string]::IsNullOrWhiteSpace($Element.Name) -and $null -eq $Element.TemplatedParent
    if (-not $isNamed) {
        $p = [System.Windows.Media.VisualTreeHelper]::GetParent($Element)
        while ($null -ne $p) {
            if ($p -is [System.Windows.FrameworkElement] -and -not [string]::IsNullOrWhiteSpace($p.Name) -and $null -eq $p.TemplatedParent) { $label += " w $(Get-UiElementLabel $p)"; break }
            $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p)
        }
    }
    if ($Element -is [System.Windows.Controls.TextBlock] -and -not [string]::IsNullOrWhiteSpace($Element.Text)) { $label += " '$(Get-UiShortText $Element.Text)'" }
    return $label
}

# Kod znaku do raportu (U+XXXX) - konsola CI w PS 5.1 nie zawsze wypisze sam znak.
function global:Get-UiCodePoint {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if ($Text.Length -ge 2 -and [char]::IsSurrogatePair($Text[0], $Text[1])) { return ('U+{0:X4}' -f [char]::ConvertToUtf32($Text[0], $Text[1])) }
    return ('U+{0:X4}' -f [int]$Text[0])
}

# GLYPH: znak ikony (obszar prywatny Unicode, U+E000-U+F8FF) w TextBlocku, który nie ma czcionki ikon
# (Segoe Fluent Icons / Segoe MDL2 Assets) - rysuje się jako pusty prostokąt.
function global:Find-UiGlyphFont {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        if ($v -isnot [System.Windows.Controls.TextBlock] -or [string]::IsNullOrEmpty($v.Text)) { continue }
        $m = [regex]::Match($v.Text, '[\uE000-\uF8FF]')
        if ($m.Success -and $v.FontFamily.Source -notmatch 'MDL2|Fluent Icons') {
            $results.Add(("GLYPH {0} | {1} | {2} | czcionka '{3}'" -f $WindowName, (Get-UiElementPath $v), (Get-UiCodePoint $m.Value), $v.FontFamily.Source))
        }
    }
    return , $results
}

# EMOJI: emoji w widocznym tekście okna (WPF rysuje je czcionką zastępczą: inny rozmiar i kolor,
# w motywie jasnym bywają nieczytelne). Zamiast nich ikony z czcionki ikon albo zwykły tekst.
function global:Find-UiEmoji {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        if ($v -isnot [System.Windows.Controls.TextBlock] -or [string]::IsNullOrEmpty($v.Text)) { continue }
        $m = [regex]::Match($v.Text, $global:UiEmojiPattern)
        if ($m.Success) {
            $results.Add(("EMOJI {0} | {1} | {2}" -f $WindowName, (Get-UiElementPath $v), (Get-UiCodePoint $m.Value)))
        }
    }
    return , $results
}

# CONTRAST: widoczny tekst, którego kolor ma za mały kontrast z faktycznym tłem (pierwsze nieprzezroczyste
# tło w górę drzewa wizualnego, z mieszaniem półprzezroczystych warstw). Łapie np. czarny tekst
# odziedziczony z domyślnego stylu Windows na ciemnej karcie. Pomijamy tekst celowo przygaszony
# (kolor ThemeFaint: podpowiedzi w polach, wersja) i tekst w wyłączonych / półprzezroczystych elementach.
function global:Get-UiRelativeLuminance {
    param([System.Windows.Media.Color]$Color)
    $ch = foreach ($c in @($Color.R, $Color.G, $Color.B)) {
        $v = $c / 255.0
        if ($v -le 0.03928) { $v / 12.92 } else { [Math]::Pow(($v + 0.055) / 1.055, 2.4) }
    }
    return 0.2126 * $ch[0] + 0.7152 * $ch[1] + 0.0722 * $ch[2]
}
function global:Get-UiContrastRatio {
    param([System.Windows.Media.Color]$A, [System.Windows.Media.Color]$B)
    $la = Get-UiRelativeLuminance $A; $lb = Get-UiRelativeLuminance $B
    $hi = [Math]::Max($la, $lb); $lo = [Math]::Min($la, $lb)
    return ($hi + 0.05) / ($lo + 0.05)
}
# Tło pod elementem: pędzle Background (Panel, Border, Control - TextBlock też) w górę drzewa aż do
# pierwszego nieprzezroczystego; półprzezroczyste nakładamy na nie od dołu. $null = tło z obrazu.
function global:Get-UiEffectiveBackground {
    param($Element, [System.Windows.Window]$Win)
    $layers = New-Object System.Collections.Generic.List[object]
    $v = $Element
    while ($null -ne $v) {
        $brush = $null
        if ($v -is [System.Windows.Controls.Panel]) { $brush = $v.Background }
        elseif ($v -is [System.Windows.Controls.Border]) { $brush = $v.Background }
        elseif ($v -is [System.Windows.Controls.TextBlock]) { $brush = $v.Background }
        elseif ($v -is [System.Windows.Window]) { $brush = $v.Background }
        $c = $null
        if ($brush -is [System.Windows.Media.SolidColorBrush]) { $c = $brush.Color }
        elseif ($brush -is [System.Windows.Media.GradientBrush] -and $brush.GradientStops.Count -gt 0) {
            # Gradient (np. kafelek logo w nagłówku): średnia kolorów punktów - znak stoi na środku kafelka.
            $r = 0; $g = 0; $bl = 0; $al = 0
            foreach ($stop in $brush.GradientStops) { $r += $stop.Color.R; $g += $stop.Color.G; $bl += $stop.Color.B; $al += $stop.Color.A }
            $n = $brush.GradientStops.Count
            $c = [System.Windows.Media.Color]::FromArgb([byte][Math]::Round($al / $n), [byte][Math]::Round($r / $n), [byte][Math]::Round($g / $n), [byte][Math]::Round($bl / $n))
        }
        elseif ($null -ne $brush) {
            # Obraz / VisualBrush: tła nie da się policzyć z kolorów, więc tekstu nie oceniamy.
            return $null
        }
        if ($null -ne $c) {
            $a = $c.A * $brush.Opacity
            if ($a -ge 255) { $layers.Add($c); break }
            if ($a -gt 0) { $layers.Add([System.Windows.Media.Color]::FromArgb([byte][Math]::Round($a), $c.R, $c.G, $c.B)) }
        }
        $v = [System.Windows.Media.VisualTreeHelper]::GetParent($v)
    }
    $result = [System.Windows.Media.Colors]::White
    if ($layers.Count -gt 0 -and $layers[$layers.Count - 1].A -ge 255) { $result = $layers[$layers.Count - 1] }
    elseif ($Win.Background -is [System.Windows.Media.SolidColorBrush]) { $result = $Win.Background.Color }
    for ($i = $layers.Count - 1; $i -ge 0; $i--) {
        $l = $layers[$i]
        if ($l.A -ge 255) { $result = $l; continue }
        $f = $l.A / 255.0
        $result = [System.Windows.Media.Color]::FromRgb([byte][Math]::Round($l.R * $f + $result.R * (1 - $f)), [byte][Math]::Round($l.G * $f + $result.G * (1 - $f)), [byte][Math]::Round($l.B * $f + $result.B * (1 - $f)))
    }
    return $result
}
function global:Find-UiContrast {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    $faint = ([string](Get-StdThemePalette)['ThemeFaint']).ToUpperInvariant()
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        if ($v -isnot [System.Windows.Controls.TextBlock] -or [string]::IsNullOrWhiteSpace($v.Text)) { continue }
        if ($v.Foreground -isnot [System.Windows.Media.SolidColorBrush]) { continue }
        $fg = $v.Foreground.Color
        $fgHex = ('#{0:X2}{1:X2}{2:X2}' -f $fg.R, $fg.G, $fg.B)
        if ($fgHex -eq $faint) { continue }
        $dimmed = $false
        $p = $v
        while ($null -ne $p) {
            if (($p -is [System.Windows.UIElement]) -and ($p.Opacity -lt 0.99 -or -not $p.IsEnabled)) { $dimmed = $true; break }
            $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p)
        }
        if ($dimmed) { continue }
        $bg = Get-UiEffectiveBackground -Element $v -Win $Win
        if ($null -eq $bg) { continue }
        $ratio = Get-UiContrastRatio $fg $bg
        if ($ratio -lt $global:StdPv.MinContrast) {
            $results.Add(("CONTRAST {0} | {1} | tekst {2} na tle #{3:X2}{4:X2}{5:X2} = {6:N2}:1 (min {7})" -f $WindowName, (Get-UiElementPath $v), $fgHex, $bg.R, $bg.G, $bg.B, $ratio, $global:StdPv.MinContrast))
        }
    }
    return , $results
}

# STALE: po przełączeniu motywu w locie żaden widoczny element nie może mieć koloru, który jest tylko
# w poprzedniej palecie (stare wartości minus nowe). Taki kolor to znak {StaticResource Theme...},
# koloru HEX w danych wiersza albo pędzla przypisanego w kodzie zamiast SetResourceReference
# (Set-StdUiColor). Palety jako IDictionary: [ordered] nie wiąże się z parametrem [hashtable].
# Foreground sprawdzamy tylko tam, gdzie jest ustawiony (lokalnie, stylem, szablonem) - wartość
# dziedziczona trafia do raportu raz, przy elemencie, który ją ustawia.
function global:Find-UiStaleColors {
    param(
        [System.Windows.Window]$Win,
        [string]$WindowName,
        [System.Collections.IDictionary]$OldPalette,
        [System.Collections.IDictionary]$NewPalette
    )
    $results = New-Object System.Collections.Generic.List[string]
    $newHex = @{}
    foreach ($k in $NewPalette.Keys) { $newHex[([string]$NewPalette[$k]).ToUpperInvariant()] = $true }
    $staleHex = @{}
    foreach ($k in $OldPalette.Keys) {
        $hex = ([string]$OldPalette[$k]).ToUpperInvariant()
        if ($newHex.ContainsKey($hex)) { continue }
        if ($staleHex.ContainsKey($hex)) { $staleHex[$hex] = $staleHex[$hex] + ', ' + $k } else { $staleHex[$hex] = [string]$k }
    }
    if ($staleHex.Count -eq 0) { return , $results }

    $fgDp = [System.Windows.Documents.TextElement]::ForegroundProperty
    $counts = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        $pairs = @()
        if ($v -is [System.Windows.Controls.Control] -or $v -is [System.Windows.Controls.Border]) { $pairs = @(@('Background', $v.Background), @('BorderBrush', $v.BorderBrush)) }
        elseif ($v -is [System.Windows.Controls.Panel] -or $v -is [System.Windows.Controls.TextBlock]) { $pairs = @(, @('Background', $v.Background)) }
        elseif ($v -is [System.Windows.Shapes.Shape]) { $pairs = @(@('Fill', $v.Fill), @('Stroke', $v.Stroke)) }
        if ($v -is [System.Windows.FrameworkElement]) {
            $src = [System.Windows.DependencyPropertyHelper]::GetValueSource($v, $fgDp).BaseValueSource
            if ($src -ne [System.Windows.BaseValueSource]::Inherited -and $src -ne [System.Windows.BaseValueSource]::Default) { $pairs += , @('Foreground', $v.GetValue($fgDp)) }
        }
        foreach ($pair in $pairs) {
            $brush = $pair[1]
            if ($brush -isnot [System.Windows.Media.SolidColorBrush]) { continue }
            $col = $brush.Color
            if ($col.A -eq 0) { continue }
            $hex = '#{0:X2}{1:X2}{2:X2}' -f $col.R, $col.G, $col.B
            if (-not $staleHex.ContainsKey($hex)) { continue }
            $line = "STALE {0} | {1} | {2} | {3} ({4})" -f $WindowName, (Get-UiElementPath $v), $pair[0], $hex, $staleHex[$hex]
            if ($counts.ContainsKey($line)) { $counts[$line]++ } else { $counts[$line] = 1; $order.Add($line) }
        }
    }
    foreach ($line in $order) {
        if ($counts[$line] -gt 1) { $results.Add("$line | x$($counts[$line])") } else { $results.Add($line) }
    }
    return , $results
}

# Przełącza motyw w locie tak jak przycisk motywu (Set-StdTheme: wszystkie otwarte okna STD, bez zapisu
# w prefs.json), szuka kolorów z poprzedniego motywu i zawsze wraca do motywu wyjściowego. PNG jest już zrobiony.
function global:Test-UiThemeSwitch {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $wasDark = Test-StdThemeDark
    $oldPalette = Get-StdThemePalette
    $found = $null
    try {
        Set-StdTheme -Dark (-not $wasDark) -NoSave
        $Win.UpdateLayout()
        $found = Find-UiStaleColors -Win $Win -WindowName $WindowName -OldPalette $oldPalette -NewPalette (Get-StdThemePalette)
    } catch {
        Add-StdPvError -Text ('{0} (przełączenie motywu): {1}' -f $WindowName, (Get-StdPvErrorText $_))
    } finally {
        try { Set-StdTheme -Dark $wasDark -NoSave } catch { Add-StdPvError -Text ('{0} (powrót do motywu): {1}' -f $WindowName, (Get-StdPvErrorText $_)) }
        $Win.UpdateLayout()
    }
    if ($null -ne $found) { foreach ($issue in $found) { Add-UiFinding $issue } }
}

# Jeden "widok" okna: kontrola ucinania, mapa układu, PNG i (gdy jest przewijanie) PNG całej zawartości.
# -Lint: dodatkowo GLYPH / EMOJI / CONTRAST i przełączenie motywu w locie (STALE) - dla zwykłych zrzutów
# i stanów dodatkowych; zrzuty w innym rozmiarze sprawdzają tylko ucinanie (reszta się od rozmiaru nie zmienia).
function global:Save-UiView {
    param([System.Windows.Window]$Win, [string]$Name, [switch]$Lint)
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
    $root.UpdateLayout()
    foreach ($issue in (Find-UiClipping -Win $Win -WindowName $Name)) { Add-StdPvIssue -Kind 'CLIP' -Line $issue }
    if ($global:StdPv.DumpLayout) { Write-UiLayoutMap -Win $Win -WindowName $Name }
    Save-UiElementPng -Element $root -Background $null -Name $Name
    # Jeśli okno ma przewijaną zawartość dłuższą niż widok - dodatkowo cała zawartość.
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($root)
    while ($queue.Count -gt 0) {
        $v = $queue.Dequeue()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        if ($v -is [System.Windows.Controls.ScrollViewer] -and $v.ExtentHeight -gt ($v.ViewportHeight + 1) -and $v.ViewportHeight -gt 0 -and $v.TemplatedParent -isnot [System.Windows.Controls.Primitives.TextBoxBase] -and $v.TemplatedParent -isnot [System.Windows.Controls.ItemsControl]) {
            # Zawartość dłuższa niż widoczny obszar = trzeba przewijać.
            Write-Host ("SCROLL {0} | {1} | zawartość {2:N0}px w widoku {3:N0}px" -f $Name, (Get-UiElementLabel $v.Content), $v.ExtentHeight, $v.ViewportHeight)
        }
        # Cała zawartość tylko dla zwykłych przewijanych paneli: lista (ListView, ListBox) i pole tekstowe
        # przewijają się wirtualnie, więc ich "cała zawartość" to i tak tylko widoczny fragment.
        if ($v -is [System.Windows.Controls.ScrollViewer] -and $v.Content -is [System.Windows.FrameworkElement] -and $v.ExtentHeight -gt ($v.ViewportHeight + 40) -and $v.ViewportHeight -gt 100 -and $v.TemplatedParent -isnot [System.Windows.Controls.Primitives.TextBoxBase] -and $v.TemplatedParent -isnot [System.Windows.Controls.ItemsControl]) {
            Save-UiElementPng -Element $v.Content -Background $Win.Background -Name "$Name-cala-zawartosc"
            continue
        }
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) {
            $queue.Enqueue([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
    }
    if ($Lint) {
        foreach ($issue in (Find-UiGlyphFont -Win $Win -WindowName $Name)) { Add-UiFinding $issue }
        foreach ($issue in (Find-UiEmoji -Win $Win -WindowName $Name)) { Add-UiFinding $issue }
        foreach ($issue in (Find-UiContrast -Win $Win -WindowName $Name)) { Add-UiFinding $issue }
        Test-UiThemeSwitch -Win $Win -WindowName $Name
    }
}

# ---------------------------------------------------------------------------------------------
# Rozmiary okna: minimalny i ekran 1366x768 przy 150%
# ---------------------------------------------------------------------------------------------

function global:Update-StdPvLayout {
    # Dwa przebiegi układu: zmiana rozmiaru wywołuje SizeChanged, a ten zmienia szerokości kolumn tabel.
    param([System.Windows.Window]$Win)
    $Win.UpdateLayout()
    $Win.UpdateLayout()
}

# Zrzut w podanym rozmiarze okna (DIP). Okno o stałym rozmiarze: MinWidth/MinHeight zerujemy, bo mały ekran
# jest mniejszy niż minimalny rozmiar okna - New-StdWindow przycina je wtedy do obszaru roboczego. Okno
# dopasowane do treści (SizeToContent): ekran ogranicza je tylko z góry (MaxWidth / MaxHeight).
function global:Save-UiSizedView {
    param([System.Windows.Window]$Win, [string]$Name, [double]$Width, [double]$Height)
    if ($Win.SizeToContent -eq [System.Windows.SizeToContent]::Manual) {
        $Win.MinWidth = 0
        $Win.MinHeight = 0
        $Win.Width = $Width
        $Win.Height = $Height
    } else {
        $Win.MaxWidth = $Width
        $Win.MaxHeight = $Height
        if (-not [double]::IsNaN($Win.Width) -and $Win.Width -gt $Width) { $Win.Width = $Width }
    }
    Update-StdPvLayout -Win $Win
    Save-UiView -Win $Win -Name $Name
}

function global:Save-StdPvSizes {
    # Minimalny rozmiar (okno o stałym rozmiarze z MinWidth/MinHeight) i ekran 911 x 472 DIP; potem powrót
    # do rozmiaru wyjściowego, żeby stany dodatkowe były w zwykłym rozmiarze.
    param([System.Windows.Window]$Win, [string]$Name)
    $saved = @{ Width = $Win.Width; Height = $Win.Height; MinWidth = $Win.MinWidth; MinHeight = $Win.MinHeight; MaxWidth = $Win.MaxWidth; MaxHeight = $Win.MaxHeight }
    try {
        if ($Win.SizeToContent -eq [System.Windows.SizeToContent]::Manual -and $Win.MinWidth -gt 0 -and $Win.MinHeight -gt 0) {
            $Win.Width = $Win.MinWidth
            $Win.Height = $Win.MinHeight
            Update-StdPvLayout -Win $Win
            Save-UiView -Win $Win -Name "$Name-min"
        }
        # Ekran 1366x768 przy skalowaniu 150% bez paska zadań: 911 x 472 DIP. Tu nie może być żadnego CLIP
        # (ucięty tekst albo przycisk); SCROLL (przewijana lista) jest oczekiwany.
        Save-UiSizedView -Win $Win -Name "$Name-1366x768-150proc" -Width $global:StdPv.SmallWidth -Height $global:StdPv.SmallHeight
    } catch {
        Add-StdPvError -Text ('{0} (inny rozmiar okna): {1}' -f $Name, (Get-StdPvErrorText $_))
    } finally {
        try {
            $Win.MaxWidth = $saved.MaxWidth
            $Win.MaxHeight = $saved.MaxHeight
            if ($Win.SizeToContent -eq [System.Windows.SizeToContent]::Manual) {
                $Win.Width = $saved.Width
                $Win.Height = $saved.Height
            }
            $Win.MinWidth = $saved.MinWidth
            $Win.MinHeight = $saved.MinHeight
            Update-StdPvLayout -Win $Win
        } catch { }
    }
}

# ---------------------------------------------------------------------------------------------
# Okna: pokazanie, zrzut, zamknięcie
# ---------------------------------------------------------------------------------------------

function global:Close-StdPvWindow {
    param([System.Windows.Window]$Win)
    if ($null -eq $Win) { return }
    $state = $Win.Tag
    if ($state -is [hashtable] -and $state.ContainsKey('ForceClose')) { $state.ForceClose = $true }
    try { $Win.Close() } catch { }
}

function global:Save-StdPvShot {
    # Okno zrzutu: przygotowanie (np. powiadomienia), zwykły zrzut z kontrolami, inne rozmiary, stany
    # dodatkowe na tym samym oknie (okna zamkniętego nie da się pokazać drugi raz) i okna podrzędne
    # (np. potwierdzenie „Wdróż”) otwierane z tego okna.
    param([System.Windows.Window]$Win, [string]$Name)
    $pv = $global:StdPv
    $spec = $pv.Spec
    if ($null -ne $spec.Prepare) {
        & $spec.Prepare $Win | Out-Null
        Update-StdPvLayout -Win $Win
    }
    Save-UiView -Win $Win -Name $Name -Lint
    if ($spec.Sized) { Save-StdPvSizes -Win $Win -Name $Name }
    foreach ($view in @($spec.Views)) {
        if ($null -eq $view) { continue }
        try {
            & $view.Apply $Win | Out-Null
            Update-StdPvLayout -Win $Win
            Save-UiView -Win $Win -Name $view.Name -Lint
        } catch {
            Add-StdPvError -Text ('{0}: {1}' -f $view.Name, (Get-StdPvErrorText $_))
        }
    }
    foreach ($child in @($spec.Children)) {
        if ($null -eq $child) { continue }
        $pv.ChildName = $child.Name
        $pv.ChildSized = [bool]$child.Sized
        $pv.ChildSeen = $false
        try {
            & $child.Open $Win | Out-Null
            if (-not $pv.ChildSeen) { Add-StdPvError -Text ('{0}: okno się nie pokazało.' -f $child.Name) }
        } catch {
            Add-StdPvError -Text ('{0}: {1}' -f $child.Name, (Get-StdPvErrorText $_))
        } finally {
            $pv.ChildName = $null
        }
    }
}

function global:Invoke-StdPvPending {
    # Okna pokazane od ostatniego razu (Loaded): zrzut, kontrole, zamknięcie. Wywoływane przez dispatcher
    # (ApplicationIdle), gdy okno ma już układ i treść.
    $pv = $global:StdPv
    while ($pv.Pending.Count -gt 0) {
        $entry = $pv.Pending[0]
        $pv.Pending.RemoveAt(0)
        $win = $entry.Window
        try {
            switch ($entry.Kind) {
                'shot' { Save-StdPvShot -Win $win -Name $entry.Name }
                'child' {
                    Save-UiView -Win $win -Name $entry.Name -Lint
                    if ($entry.Sized) { Save-StdPvSizes -Win $win -Name $entry.Name }
                }
                default {
                    # Okno, którego podgląd nie zamawiał - zwykle komunikat o błędzie z handlera okna.
                    $text = ''
                    try { $box = $win.FindName('txtMessage'); if ($null -ne $box) { $text = [string]$box.Text } } catch { }
                    Add-StdPvError -Text ("{0}: nieoczekiwane okno '{1}': {2}" -f $pv.ShotName, $win.Title, (Get-UiShortText -Text $text -Max 400))
                    Save-UiView -Win $win -Name $entry.Name
                }
            }
        } catch {
            Add-StdPvError -Text ('{0}: {1}' -f $entry.Name, (Get-StdPvErrorText $_))
        } finally {
            Close-StdPvWindow -Win $win
        }
    }
}

# Każde okno po wczytaniu (Loaded) czeka, aż dispatcher skończy pracę (dane, układ), po czym jest
# zapisywane do PNG i zamykane. Pierwsze okno zrzutu to okno zrzutu, okno zapowiedziane w ChildName -
# okno podrzędne, każde inne - nieoczekiwane (ERR).
$global:StdPvLoadedHandler = [System.Windows.RoutedEventHandler]{
    param($pvSender, $pvEventArgs)
    try {
        $pv = $global:StdPv
        if (-not $pv.Active) { return }
        $entry = @{ Window = $pvSender; Kind = 'unexpected'; Name = ''; Sized = $false }
        if ($null -eq $pv.ShotWindow) {
            $pv.ShotWindow = $pvSender
            $entry.Kind = 'shot'
            $entry.Name = $pv.ShotName
        } elseif (-not [string]::IsNullOrEmpty($pv.ChildName)) {
            $entry.Kind = 'child'
            $entry.Name = $pv.ChildName
            $entry.Sized = $pv.ChildSized
            $pv.ChildName = $null
            $pv.ChildSeen = $true
        } else {
            $pv.Unexpected++
            $entry.Name = '{0}-nieoczekiwane-okno-{1}' -f $pv.ShotName, $pv.Unexpected
        }
        [void]$pv.Pending.Add($entry)
        [void]$pvSender.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::ApplicationIdle, [Action]{ Invoke-StdPvPending })
    } catch {
        Write-Host ('ERR  obsługa Loaded: {0}' -f $_.Exception.Message)
    }
}
[System.Windows.EventManager]::RegisterClassHandler([System.Windows.Window], [System.Windows.FrameworkElement]::LoadedEvent, $global:StdPvLoadedHandler)

function Test-StdPvSelected {
    param([string]$Name)
    if ($global:StdPv.Only.Count -eq 0) { return $true }
    foreach ($pattern in $global:StdPv.Only) { if ($Name -like $pattern) { return $true } }
    return $false
}

function Invoke-StdPvShot {
    <#
    Jedno okno podglądu: -Show pokazuje okno (Show-Std*Window / Show-StdMessage z danymi przykładowymi),
    handler Loaded robi zrzuty i je zamyka. -Kind: rodzaj okna do kontroli pokrycia (Connect, Login, Main,
    Message). -Prepare { param($Win) }: przed zrzutem. -Views @(@{ Name; Apply = { param($Win) } }): stany
    dodatkowe. -Children @(@{ Name; Open = { param($Win) }; Sized }): okna podrzędne. -Sized: także
    minimalny rozmiar i ekran 911 x 472.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Kind,
        [Parameter(Mandatory = $true)][scriptblock]$Show,
        [scriptblock]$Prepare,
        [object[]]$Views,
        [object[]]$Children,
        [switch]$Sized
    )
    if (-not (Test-StdPvSelected -Name $Name)) { return }
    $pv = $global:StdPv
    $pv.Spec = @{ Prepare = $Prepare; Views = $Views; Children = $Children; Sized = [bool]$Sized }
    $pv.ShotName = $Name
    $pv.ShotWindow = $null
    $pv.ChildName = $null
    $pv.Unexpected = 0
    $pv.Pending.Clear()
    $errorsBefore = $pv.Errors.Count
    Write-Host ('---- {0} ----' -f $Name)
    # Straż: okno, które się nie zamknęło (np. zawieszony handler), zamykamy po 120 s z błędem ERR,
    # zamiast czekać do limitu czasu zadania w CI.
    $watchdog = New-Object System.Windows.Threading.DispatcherTimer
    $watchdog.Interval = [TimeSpan]::FromSeconds(120)
    $watchdog.Tag = @{ Name = $Name }
    $watchdog.Add_Tick({
            param($pvTimer, $pvTimerArgs)
            try {
                $pvTimer.Stop()
                Add-StdPvError -Text ('{0}: okno nie zamknęło się w 120 s - zamknięte siłą.' -f $pvTimer.Tag.Name)
                # Get-StdOpenWindows zwraca listę jako jeden obiekt (bez @(), które w PS 5.1 owinęłoby ją w tablicę).
                $pvOpen = Get-StdOpenWindows
                if ($null -ne $pvOpen) {
                    for ($pvI = $pvOpen.Count - 1; $pvI -ge 0; $pvI--) { if ($pvOpen[$pvI].IsVisible) { Close-StdPvWindow -Win $pvOpen[$pvI] } }
                }
            } catch { }
        })
    $pv.Active = $true
    $watchdog.Start()
    try {
        & $Show | Out-Null
    } catch {
        Add-StdPvError -Text ('{0}: {1}' -f $Name, (Get-StdPvErrorText $_))
    } finally {
        $watchdog.Stop()
        $pv.Active = $false
    }
    if ($null -ne $pv.ShotWindow) {
        $pv.Shots++
        $pv.Kinds[$Kind] = $true
    } elseif ($pv.Errors.Count -eq $errorsBefore) {
        Add-StdPvError -Text ('{0}: okno się nie pokazało.' -f $Name)
    }
    # Gdyby coś zostało otwarte (np. okno bez ShowDialog), zamykamy, żeby nie przeszkadzało następnym.
    $pvLeftOpen = Get-StdOpenWindows
    if ($null -ne $pvLeftOpen) {
        foreach ($pvLeft in $pvLeftOpen.ToArray()) { if ($pvLeft.IsVisible) { Close-StdPvWindow -Win $pvLeft } }
    }
    $pv.ShotWindow = $null
    $pv.Spec = $null
}

# ---------------------------------------------------------------------------------------------
# Motyw i dane przykładowe
# ---------------------------------------------------------------------------------------------

Set-StdTheme -Dark ($Theme -eq 'Dark') -NoSave
# Ostatnio używany wybór w prefs.json (folder tymczasowy): listy w oknie głównym mają wtedy „ostatnio” na górze.
Save-StdUiRecent -ClientId 'lizard' -SiteId 'krakow' -ProfileName 'Standard'

# Dziennik wdrożenia w trakcie: wpisy przykładowe i więcej wpisów każdego poziomu (długie teksty, ścieżki).
$pvRunLog = New-Object 'System.Collections.Generic.List[object]'
foreach ($pvLogEntry in @((Get-StdUiPreviewSample -State Running).Log)) { if ($null -ne $pvLogEntry) { $pvRunLog.Add($pvLogEntry) } }
$pvLogTime = (Get-Date).Date.AddHours(10).AddMinutes(16)
$pvLogLines = @(
    @('Debug', 'Invoke-StdProcess', 'Start: msiexec.exe /i "\\srv-files\STD$\std\files\lizard\TeamViewer_Host.msi" /qn /norestart (limit 30 min).'),
    @('Info', 'Apps.Install', 'Instalacja 1 z 5: TeamViewer Host (plik TeamViewer_Host.msi, SHA-256 zgodny).'),
    @('Info', 'Apps.Install', 'Instalacja 2 z 5: Antywirus (eset.msi).'),
    @('Error', 'Invoke-StdProcess', 'eset.msi: kod wyjścia 1603 po 12,4 s - instalator zgłosił błąd krytyczny. Szczegóły w pliku C:\ProgramData\STD\Logs\20261102-101500-ab12\eset-install.log.'),
    @('Info', 'Apps.Install', 'Instalacja 3 z 5: Google Chrome (winget install --id Google.Chrome -e --silent --accept-package-agreements --accept-source-agreements).'),
    @('Warn', 'Apps.Install', 'Pobieranie z udziału \\srv-files\STD$\std\files\lizard\ trwa dłużej niż zwykle (12 MB/min). Gdy udział nie odpowiada, STD przechodzi do następnego źródła z listy lokalizacji: https://pobieranie.example.com/std/files/lizard/.')
)
foreach ($pvLine in $pvLogLines) {
    $pvLogTime = $pvLogTime.AddSeconds(7)
    $pvRunLog.Add(@{ Time = $pvLogTime; Level = $pvLine[0]; Source = $pvLine[1]; Message = $pvLine[2] })
}
for ($pvI = 1; $pvI -le 24; $pvI++) {
    $pvLogTime = $pvLogTime.AddSeconds(3)
    $pvRunLog.Add(@{ Time = $pvLogTime; Level = 'Info'; Source = 'Apps.Install'; Message = ('Pobieranie ChromeStandaloneSetup64.exe: {0}% ({1} z 112 MB).' -f ($pvI * 4), [int]($pvI * 4.6)) })
}
$pvRunLogArray = $pvRunLog.ToArray()

$pvRebootResult = @{ Status = 'Finished'; RebootRequired = $true; Error = $null }
$pvToastLong = 'Nie udało się pobrać konfiguracji klienta z serwera: https://pobieranie.example.com/std/clients/lizard/release.sdt nie odpowiada (limit czasu 30 s). Użyto kopii z cache (rewizja 42).'

# ---------------------------------------------------------------------------------------------
# „Połącz z siecią”
# ---------------------------------------------------------------------------------------------

Invoke-StdPvShot -Name 'polacz' -Kind Connect -Sized -Show { Show-StdConnectWindow -OriginRoot $pvDataRoot -PreviewData @{ UsbProfiles = 2 } }
Invoke-StdPvShot -Name 'polacz-siec-otwarta' -Kind Connect -Show {
    Show-StdConnectWindow -PreviewData @{
        SelectedSsid = 'Firma-Goscie z bardzo długą nazwą sieci do sprawdzenia'; Status = 'connecting'
        Message = 'Łączenie z siecią Firma-Goscie z bardzo długą nazwą sieci do sprawdzenia...'; MessageKind = 'Info'
    }
}
Invoke-StdPvShot -Name 'polacz-siec-firmowa' -Kind Connect -Show {
    Show-StdConnectWindow -PreviewData @{
        SelectedSsid = 'Firma-802.1X'; UsbProfiles = 3
        Message = 'Do sieci firmowej (802.1X) STD nie łączy się samo. Zaimportuj profil z pendrive''a albo podłącz kabel sieciowy.'; MessageKind = 'Warn'
    }
}
Invoke-StdPvShot -Name 'polacz-blad' -Kind Connect -Show {
    Show-StdConnectWindow -PreviewData @{ Message = 'Nie udało się połączyć z siecią Firma: hasło jest nieprawidłowe albo sieć jest poza zasięgiem. Sprawdź hasło i spróbuj ponownie.'; MessageKind = 'Crit' }
}
Invoke-StdPvShot -Name 'polacz-polaczono' -Kind Connect -Show {
    Show-StdConnectWindow -PreviewData @{ Status = 'connected'; Message = 'Połączono z siecią Firma.'; MessageKind = 'Ok' }
}
Invoke-StdPvShot -Name 'polacz-brak-sieci' -Kind Connect -Sized -Show { Show-StdConnectWindow -PreviewData @{ Networks = @(); UsbProfiles = 0 } }

# ---------------------------------------------------------------------------------------------
# „Logowanie”
# ---------------------------------------------------------------------------------------------

Invoke-StdPvShot -Name 'logowanie' -Kind Login -Sized -Show { Show-StdLoginWindow -PreviewData @{} }
Invoke-StdPvShot -Name 'logowanie-pusty' -Kind Login -Sized -Show { Show-StdLoginWindow -PreviewData @{ Server = ''; UserName = '' } }
Invoke-StdPvShot -Name 'logowanie-blad' -Kind Login -Sized -Show {
    Show-StdLoginWindow -PreviewData @{
        Error   = "Nieprawidłowy login lub hasło.`n`nDostępna jest kopia listy klientów (cache) - możesz pracować offline."
        Offline = $true
    }
}
Invoke-StdPvShot -Name 'logowanie-w-toku' -Kind Login -Show { Show-StdLoginWindow -PreviewData @{ Busy = 'Logowanie do serwera https://pobieranie.example.com/std/...' } }

# ---------------------------------------------------------------------------------------------
# Okno główne
# ---------------------------------------------------------------------------------------------

Invoke-StdPvShot -Name 'okno-glowne' -Kind Main -Sized -Show { Show-StdMainWindow -PreviewData @{} -PreviewState Select } -Views @(
    @{ Name = 'okno-glowne-zakladka-plan'; Apply = { param($Win) $Win.Tag.C.tabSelect.SelectedIndex = 1 } }
)
Invoke-StdPvShot -Name 'okno-glowne-bez-wyboru' -Kind Main -Show { Show-StdMainWindow -PreviewData @{ ClientId = '' } -PreviewState Select }
Invoke-StdPvShot -Name 'okno-glowne-blad-konfiguracji' -Kind Main -Show {
    Show-StdMainWindow -PreviewState Select -PreviewData @{
        ReleaseError = 'Serwer STD nie odpowiada (HTTP 503), a kopia w cache jest starsza niż ostatnio widziana rewizja 42 - STD jej nie użyje. Spróbuj ponownie za chwilę albo poproś administratora o sprawdzenie serwera.'
    }
}
Invoke-StdPvShot -Name 'okno-glowne-plan' -Kind Main -Sized -Show { Show-StdMainWindow -PreviewData @{} -PreviewState Plan } -Views @(
    @{ Name = 'okno-glowne-plan-z-wylaczonymi'; Apply = { param($Win) $Win.Tag.C.chkShowDisabled.IsChecked = $true } }
) -Children @(
    # Prawdziwe potwierdzenie „Wdróż” z listą kroków planu (Confirm-StdMainDeploy); zamknięcie = „Nie”.
    @{ Name = 'potwierdzenie-wdrozenia'; Sized = $true; Open = { param($Win) Confirm-StdMainDeploy -Ui $Win.Tag } }
)
Invoke-StdPvShot -Name 'okno-glowne-wdrozenie' -Kind Main -Sized -Show { Show-StdMainWindow -PreviewData @{ Log = $pvRunLogArray } -PreviewState Running } -Views @(
    @{ Name = 'okno-glowne-wdrozenie-filtr-dziennika'; Apply = { param($Win) $Win.Tag.C.cmbLogLevel.SelectedIndex = 2 } },
    @{
        # Po pierwszym „Przerwij”: przycisk „Przerwij natychmiast” i karta „Przerywanie po bieżącym kroku...”.
        Name  = 'okno-glowne-przerywanie'
        Apply = { param($Win) $pvUi = $Win.Tag; $pvUi.CancelRequested = $true; Set-StdMainState -Ui $pvUi -State Running }
    }
)
Invoke-StdPvShot -Name 'okno-glowne-pauza' -Kind Main -Sized -Show { Show-StdMainWindow -PreviewData @{} -PreviewState Paused }
Invoke-StdPvShot -Name 'okno-glowne-blad' -Kind Main -Sized -Show { Show-StdMainWindow -PreviewData @{} -PreviewState Error }
Invoke-StdPvShot -Name 'okno-glowne-koniec' -Kind Main -Sized -Show { Show-StdMainWindow -PreviewData @{} -PreviewState Finished }
Invoke-StdPvShot -Name 'okno-glowne-koniec-restart' -Kind Main -Show { Show-StdMainWindow -PreviewData @{ Result = $pvRebootResult } -PreviewState Finished }
Invoke-StdPvShot -Name 'okno-glowne-powiadomienia' -Kind Main -Show { Show-StdMainWindow -PreviewData @{} -PreviewState Select } -Prepare {
    param($Win)
    # -Seconds 60: powiadomienia nie znikną przed zrzutem.
    Show-StdToast -Window $Win -Message 'Skopiowano do schowka.' -Kind Success -Seconds 60
    Show-StdToast -Window $Win -Message 'Konfiguracja klienta jest aktualna (rewizja 42).' -Kind Info -Seconds 60
    Show-StdToast -Window $Win -Message 'Czekam na zatrzymanie bieżącego kroku - okno zamknie się samo.' -Kind Warning -Seconds 60
    Show-StdToast -Window $Win -Message $pvToastLong -Kind Error -Seconds 60
}

# ---------------------------------------------------------------------------------------------
# Okna komunikatu (Show-StdMessage) - zamknięcie bez kliknięcia = przycisk Anuluj / Nie / OK
# ---------------------------------------------------------------------------------------------

Invoke-StdPvShot -Name 'komunikat-informacja' -Kind Message -Show {
    Show-StdMessage -Title 'Brak kroków do wykonania' -Image Information -Message 'Dla tego profilu nie ma żadnego włączonego kroku do wykonania. Sprawdź profil albo wybierz inny.'
}
Invoke-StdPvShot -Name 'komunikat-pytanie' -Kind Message -Show {
    Show-StdMessage -Title 'Ponowić nieudane kroki?' -Image Question -Button YesNo -Message 'Kroki z błędem i przerwane (2) zostaną uruchomione jeszcze raz. Kroki zakończone poprawnie nie będą powtarzane.'
}
Invoke-StdPvShot -Name 'komunikat-ostrzezenie' -Kind Message -Show {
    # Tekst w apostrofach: PowerShell traktuje „ i ” jak cudzysłów, więc w "..." zakończyłyby napis.
    Show-StdMessage -Title 'Przerwać wdrożenie?' -Image Warning -Button YesNo -Danger -Message ('Wdrożenie zostanie przerwane po zakończeniu bieżącego kroku. Pozostałe kroki dostaną stan „Przerwano” i będzie można je uruchomić przyciskiem „Ponów nieudane”.{0}{0}Przerwać wdrożenie?' -f "`n")
}
Invoke-StdPvShot -Name 'komunikat-blad' -Kind Message -Show {
    Show-StdMessage -Title 'Błąd' -Image Error -Message ("Nie udało się pobrać konfiguracji klienta.`n`n[STD-SIG-INVALID] Podpis konfiguracji jest nieprawidłowy (clients/lizard/release.sdt).`n`nSzczegóły są w dzienniku STD.")
}
Invoke-StdPvShot -Name 'komunikat-sukces' -Kind Message -Show {
    Show-StdMessage -Title 'Wdrożenie zakończone' -Image Success -Message 'Wszystkie kroki zakończyły się poprawnie. Podsumowanie jest w schowku.'
}
Invoke-StdPvShot -Name 'komunikat-tak-nie-anuluj' -Kind Message -Show {
    Show-StdMessage -Title 'Zapisać adres serwera?' -Image Warning -Button YesNoCancel -Message 'Adres https://pobieranie.example.com/std/ nie jest zapisany na tym komputerze. Zapisać go dla kolejnych uruchomień STD?'
}
Invoke-StdPvShot -Name 'komunikat-ok-anuluj' -Kind Message -Show {
    Show-StdMessage -Title 'Import profili Wi-Fi' -Image Information -Button OKCancel -Message 'Na pendrivie są 3 profile Wi-Fi (folder wifi). Zaimportować je do systemu Windows?'
}
$pvLongMessage = New-Object System.Text.StringBuilder
[void]$pvLongMessage.AppendLine('Wykryto błędy w konfiguracji klienta:')
[void]$pvLongMessage.AppendLine('')
for ($pvI = 1; $pvI -le 30; $pvI++) {
    [void]$pvLongMessage.AppendLine(('•  Aplikacja {0}: brak pliku instalatora \\srv-files\STD$\std\files\lizard\instalator-{0}.msi (sprawdź SHA-256 w konfiguracji).' -f $pvI))
}
[void]$pvLongMessage.Append('Popraw konfigurację w edytorze administratora i opublikuj ją ponownie.')
$pvLongMessageText = $pvLongMessage.ToString()
Invoke-StdPvShot -Name 'komunikat-dlugi' -Kind Message -Show { Show-StdMessage -Title 'Walidacja konfiguracji' -Image Error -Message $pvLongMessageText } -Sized

# ---------------------------------------------------------------------------------------------
# Ikona STD: każdy rozmiar z pliku .ico w skali 1:1 na tle motywu - wersja ze znakiem z czcionki
# i zapasowa rysowana kreską - oraz sam plik ikona.ico do obejrzenia w Eksploratorze.
# ---------------------------------------------------------------------------------------------

if (Test-StdPvSelected -Name 'ikona') {
    $global:StdPv.ShotName = 'ikona'
    try {
        $pvIconBytes = New-StdAppIconBytes
        [System.IO.File]::WriteAllBytes((Join-Path $global:StdPv.Dir 'ikona.ico'), $pvIconBytes)
        $pvTextBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString((Get-StdThemeColor 'ThemeText')))
        $pvIconPanel = New-Object System.Windows.Controls.StackPanel
        $pvIconPanel.Margin = '12'
        $pvVariants = @(
            @{ Label = 'Ikona STD (znak z czcionki Segoe Fluent Icons / Segoe MDL2 Assets)'; Bytes = $pvIconBytes },
            @{ Label = 'Zapas bez czcionki z ikonami (paczka rysowana kreską)'; Bytes = (New-StdAppIconBytes -DrawnMark) }
        )
        foreach ($pvVariant in $pvVariants) {
            $pvCaption = New-Object System.Windows.Controls.TextBlock
            $pvCaption.Text = $pvVariant.Label
            $pvCaption.Foreground = $pvTextBrush
            $pvCaption.Margin = '0,8,0,6'
            [void]$pvIconPanel.Children.Add($pvCaption)
            $pvRow = New-Object System.Windows.Controls.StackPanel
            $pvRow.Orientation = 'Horizontal'
            $pvDecoder = New-Object System.Windows.Media.Imaging.IconBitmapDecoder -ArgumentList (New-Object System.IO.MemoryStream (, [byte[]]$pvVariant.Bytes)), ([System.Windows.Media.Imaging.BitmapCreateOptions]::None), ([System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
            # Od najmniejszego: dekoder podaje klatki we własnej kolejności.
            foreach ($pvFrame in @($pvDecoder.Frames | Sort-Object PixelWidth)) {
                $pvImage = New-Object System.Windows.Controls.Image
                $pvImage.Source = $pvFrame
                $pvImage.Width = $pvFrame.PixelWidth
                $pvImage.Height = $pvFrame.PixelHeight
                $pvImage.Stretch = 'Fill'
                $pvImage.VerticalAlignment = 'Bottom'
                $pvImage.Margin = '0,0,12,0'
                [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($pvImage, [System.Windows.Media.BitmapScalingMode]::NearestNeighbor)
                [void]$pvRow.Children.Add($pvImage)
            }
            [void]$pvIconPanel.Children.Add($pvRow)
        }
        $pvIconPanel.Measure((New-Object System.Windows.Size ([double]::PositiveInfinity), ([double]::PositiveInfinity)))
        $pvIconPanel.Arrange((New-Object System.Windows.Rect $pvIconPanel.DesiredSize))
        $pvBackBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString((Get-StdThemeColor 'ThemeBackground')))
        Save-UiElementPng -Element $pvIconPanel -Background $pvBackBrush -Name 'ikona'
    } catch {
        Add-StdPvError -Text ('ikona: {0}' -f (Get-StdPvErrorText $_))
    }
}

# ---------------------------------------------------------------------------------------------
# Pokrycie: każde okno STD 4 (New-Std*Window) musi mieć podgląd w tym skrypcie
# ---------------------------------------------------------------------------------------------

if ($global:StdPv.Only.Count -eq 0) {
    $pvKnown = @{ 'New-StdWindow' = $null; 'New-StdConnectWindow' = 'Connect'; 'New-StdLoginWindow' = 'Login'; 'New-StdMainWindow' = 'Main'; 'New-StdMessageWindow' = 'Message' }
    foreach ($pvFunction in @(Get-Command -Name 'New-Std*Window' -CommandType Function -ErrorAction SilentlyContinue)) {
        $pvName = [string]$pvFunction.Name
        if (-not $pvKnown.ContainsKey($pvName)) {
            Add-StdPvError -Text ('Okno {0} nie ma podglądu - dopisz je (z danymi przykładowymi) w v4\Tools\Render-StdUiPreview.ps1.' -f $pvName)
            continue
        }
        $pvKind = $pvKnown[$pvName]
        if ($null -ne $pvKind -and -not $global:StdPv.Kinds.ContainsKey($pvKind)) {
            Add-StdPvError -Text ('Okno {0} nie zostało zapisane w żadnym zrzucie (wcześniejsze błędy ERR).' -f $pvName)
        }
    }
}

Write-Host ('Gotowe: {0} plików PNG w {1}' -f $global:StdPv.Index, $global:StdPv.Dir)

# ---------------------------------------------------------------------------------------------
# Raporty, sprzątanie, kod wyjścia
# ---------------------------------------------------------------------------------------------

try { Unregister-StdLogSink -Name 'StdUiPreview' } catch { }
$pvCounts = Write-StdPvReport
$env:STD_DATA_ROOT = $pvOldDataRoot
$env:STD_UI_PREVIEW = $pvOldUiPreview
try { Remove-Item -LiteralPath $pvDataRoot -Recurse -Force -ErrorAction Stop } catch { }

if ($FailOnClipping -and $pvCounts[0] -gt 0) { exit 1 }
if ($FailOnFindings -and ($pvCounts[0] + $pvCounts[1]) -gt 0) { exit 1 }
exit 0
