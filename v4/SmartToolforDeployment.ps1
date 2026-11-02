# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# SmartToolforDeployment.ps1 - program STD 4 (SPEC 8). Zwykle uruchamia go Start-STD.cmd:
#
#     powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File SmartToolforDeployment.ps1 [parametry]
#
# Parametry:
#   -Server <adres>     adres serwera STD (https://.../std/); ma pierwszeństwo przed bootstrap.json;
#   -OriginRoot <folder> folder startowy (pendrive): tam są bootstrap.json, offline\ i wifi\.
#                       Domyślnie folder tego skryptu. Przy przejściu do kopii w App\<wersja> program
#                       podaje go sam;
#   -NoCopy             praca prosto z bieżącego folderu, bez kopii do %ProgramData%\STD\App\<wersja>;
#   -PreviewMode        podgląd okien i testy: bez podnoszenia uprawnień, bez kopii i bez pokazywania okien.
#                       Skrypt wczytuje moduł i okna, ustawia $script:StdPreview = $true i wraca
#                       (narzędzie podglądu wczytuje go przez dot-source);
#   -StdRelaunch <...>  parametr wewnętrzny: które ponowne uruchomienia już były (admin, sta, copy).
#                       Chroni przed pętlą ponownych uruchomień - nie trzeba go podawać.
#
# Kolejność startu:
# 1. Brak uprawnień administratora -> ponowne uruchomienie jako administrator (okno UAC) i wyjście.
#    Odmowa w oknie UAC -> komunikat „STD wymaga uprawnień administratora” i wyjście. Start z dysku
#    sieciowego: proces administratora dostaje ścieżki UNC (nie widzi liter dysków zmapowanych przez
#    zwykłego użytkownika); gdy ich nie da się ustalić albo proces od razu się kończy - komunikat, co zrobić.
# 2. Wątek nie STA (okna WPF tego wymagają) -> ponowne uruchomienie z -STA.
#    Potem: folder %ProgramData%\STD tylko dla Administratorów i SYSTEM (Protect-StdDataFolder) i blokada
#    jednej instancji (Enter-StdAppInstance) - gdy STD już działa, komunikat i wyjście.
# 3. Uruchomienie spoza %ProgramData%\STD\App (np. z pendrive'a) i bez -NoCopy -> kopia narzędzia do
#    App\<wersja> (pomijana, gdy pliki są identyczne) i ponowne uruchomienie stamtąd z -OriginRoot.
#    Dzięki temu pendrive można wyjąć w trakcie wdrożenia, a po restarcie STD działa z dysku.
#    Gdy kopia się nie uda, STD działa dalej z bieżącego folderu.
# 4. Moduł STD i okna, protokoły TLS, motyw z prefs.json, dziennik okien (Logs\ui-<data>.log), potem ekrany:
#    „Połącz z siecią” (tylko gdy komputer nie ma sieci) -> „Logowanie” -> okno główne (z plikiem startowym,
#    którego użyło logowanie).
#
# Każdy nieobsłużony błąd trafia do %TEMP%\STD-crash.log i do prostego okna komunikatu (trap na dole
# pliku działa dla całego skryptu). Dane logowania do domeny i hasło lokalnego administratora są
# wpisywane w oknach i nigdy nie trafiają do wiersza poleceń, logu ani pliku z błędem.

param(
    [string]$Server,
    [string]$OriginRoot,
    [switch]$NoCopy,
    [switch]$PreviewMode,
    [string]$StdRelaunch
)

# ---------------------------------------------------------------------------------------------
# Pomocnicze (przed wczytaniem okien; bez modułu STD albo z nim)
# ---------------------------------------------------------------------------------------------

function Test-StartupIsWindows {
    return ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
}

function Protect-StartupText {
    # Maskowanie sekretów (gdy moduł STD jest już wczytany) i skrócenie długiego tekstu.
    param([AllowNull()][string]$Text, [int]$MaxLength = 0)
    if ($null -eq $Text) { $Text = '' }
    if (Get-Command -Name 'Protect-StdLogText' -ErrorAction SilentlyContinue) {
        try { $Text = Protect-StdLogText -Text $Text } catch { }
    }
    if ($MaxLength -gt 0 -and $Text.Length -gt $MaxLength) { $Text = $Text.Substring(0, $MaxLength) + '...' }
    return $Text
}

function Show-StartupMessage {
    # Proste okno komunikatu (bez motywu STD): przed wczytaniem okien i przy błędzie krytycznym.
    # Bez okien (Linux, testy) komunikat idzie na standardowe wyjście błędów.
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Title = 'STD (Smart Tool for Deployment)',
        [ValidateSet('Information', 'Warning', 'Error')][string]$Icon = 'Information'
    )
    if (Test-StartupIsWindows) {
        try {
            Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
            [void][System.Windows.MessageBox]::Show($Message, $Title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]$Icon)
            return
        } catch { }
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            [void][System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]$Icon)
            return
        } catch { }
    }
    try { [Console]::Error.WriteLine(('{0}: {1}' -f $Title, $Message)) } catch { }
}

function Write-StartupCrashLog {
    # Dopisuje opis błędu do %TEMP%\STD-crash.log (UTF-8 bez BOM; plik większy niż 1 MB zaczynamy od nowa).
    # Zwraca ścieżkę pliku albo $null.
    param([Parameter(Mandatory = $true)]$ErrorRecord)
    $tempDir = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($tempDir)) { $tempDir = [System.IO.Path]::GetTempPath() }
    $path = [System.IO.Path]::Combine($tempDir, 'STD-crash.log')

    $version = 'nieznana'
    if (Get-Command -Name 'Get-StdVersion' -ErrorAction SilentlyContinue) {
        try { $version = [string](Get-StdVersion) } catch { }
    }
    $exception = $null
    $position = ''
    $stack = ''
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        $exception = $ErrorRecord.Exception
        if ($null -ne $ErrorRecord.InvocationInfo) { $position = [string]$ErrorRecord.InvocationInfo.PositionMessage }
        $stack = [string]$ErrorRecord.ScriptStackTrace
    } elseif ($ErrorRecord -is [System.Exception]) {
        $exception = $ErrorRecord
    }
    $message = [string]$ErrorRecord
    $details = ''
    if ($null -ne $exception) {
        $message = [string]$exception.Message
        $details = [string]$exception.ToString()
    }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add(('==== {0} ====' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)))
    $lines.Add(('STD: {0}' -f $version))
    $lines.Add(('PowerShell: {0} ({1}), proces 64-bitowy: {2}' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, [System.Environment]::Is64BitProcess))
    $lines.Add(('System: {0}' -f [System.Environment]::OSVersion.VersionString))
    $lines.Add(('Wątek: {0}' -f [System.Threading.Thread]::CurrentThread.ApartmentState))
    $lines.Add(('Skrypt: {0}' -f $PSCommandPath))
    $lines.Add(('Błąd: {0}' -f $message))
    if (-not [string]::IsNullOrWhiteSpace($position)) { $lines.Add(('Miejsce: {0}' -f $position.Trim())) }
    if (-not [string]::IsNullOrWhiteSpace($stack)) { $lines.Add(('Stos skryptu:{0}{1}' -f [System.Environment]::NewLine, $stack)) }
    if (-not [string]::IsNullOrWhiteSpace($details)) { $lines.Add(('Wyjątek:{0}{1}' -f [System.Environment]::NewLine, $details)) }
    $lines.Add('')
    $text = Protect-StartupText -Text (($lines -join [System.Environment]::NewLine) + [System.Environment]::NewLine)

    try {
        $encoding = New-Object System.Text.UTF8Encoding($false)
        if ([System.IO.File]::Exists($path) -and (New-Object System.IO.FileInfo($path)).Length -gt 1048576) {
            [System.IO.File]::WriteAllText($path, $text, $encoding)
        } else {
            [System.IO.File]::AppendAllText($path, $text, $encoding)
        }
        return $path
    } catch {
        return $null
    }
}

function Get-StartupMarker {
    # Etapy ponownego uruchomienia, które już były (parametr -StdRelaunch, np. "admin,copy").
    param([AllowNull()][string]$Text)
    $result = New-Object 'System.Collections.Generic.List[string]'
    if ([string]::IsNullOrWhiteSpace($Text)) { return , $result.ToArray() }
    foreach ($part in ($Text -split '[,;\s]+')) {
        $name = $part.Trim().ToLowerInvariant()
        if ($name.Length -gt 0 -and -not $result.Contains($name)) { $result.Add($name) }
    }
    return , $result.ToArray()
}

function Invoke-StartupRelaunch {
    # Ponowne uruchomienie programu (Start-StdElevated: powershell.exe -STA, jako administrator) z tymi samymi
    # parametrami, dopisanym etapem w -StdRelaunch i parametrami z -Add. Zwraca $true, gdy proces wystartował.
    # -PassThru: wynik Start-StdElevated -PassThru (obiekt procesu albo $true; $false przy odmowie UAC).
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$BoundParameters,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Done,
        [System.Collections.IDictionary]$Add,
        [string]$WorkingDirectory,
        [switch]$PassThru
    )
    $stages = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $Done) { if (-not $stages.Contains($item)) { $stages.Add($item) } }
    if (-not $stages.Contains($Stage)) { $stages.Add($Stage) }
    $extra = @{ StdRelaunch = ($stages -join ',') }
    if ($null -ne $Add) { foreach ($key in @($Add.Keys)) { $extra[[string]$key] = $Add[$key] } }
    $arguments = Get-StdRelaunchArgument -BoundParameters $BoundParameters -Add $extra
    Write-StdLog -Level Info -Source 'SmartToolforDeployment' -Message ("Ponowne uruchomienie STD (etap: {0}): {1}" -f $Stage, $ScriptPath)
    $startParams = @{ ScriptPath = $ScriptPath; Arguments = $arguments }
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $startParams.WorkingDirectory = $WorkingDirectory }
    if ($PassThru) { return (Start-StdElevated @startParams -PassThru) }
    return [bool](Start-StdElevated @startParams)
}

function Get-StartupEarlyExitCode {
    # Kod wyjścia procesu, który zakończył się w ciągu -Milliseconds; $null, gdy nadal działa albo kodu nie da się
    # odczytać. Zamyka obiekt procesu (sam proces działa dalej).
    param([Parameter(Mandatory = $true)][System.Diagnostics.Process]$Process, [int]$Milliseconds = 5000)
    try {
        if ($Process.WaitForExit($Milliseconds)) { return [int]$Process.ExitCode }
    } catch {
    } finally {
        $Process.Dispose()
    }
    return $null
}

function Select-StartupResult {
    # Wynik okna: ostatni element różny od $null (gdyby funkcja okna wypisała wcześniej coś jeszcze).
    param([AllowNull()][AllowEmptyCollection()][object[]]$Output)
    if ($null -eq $Output) { return $null }
    for ($i = $Output.Count - 1; $i -ge 0; $i--) {
        if ($null -ne $Output[$i]) { return , $Output[$i] }
    }
    return $null
}

function Get-StartupBootstrap {
    # Plik startowy (adres serwera): -Server, <OriginRoot>\bootstrap.json, zapisany adres. Błędny adres
    # nie zatrzymuje programu: jeden komunikat i kolejne źródła - bootstrap.json (gdy błędny był -Server)
    # i adres zapisany na tym komputerze (Get-StdPath Bootstrap); bez nich technik wpisze adres w oknie logowania.
    param([AllowNull()][string]$ServerUrl, [Parameter(Mandatory = $true)][string]$Root)
    try {
        return (Get-StdBootstrap -Server $ServerUrl -ScriptRoot $Root)
    } catch {
        $text = Protect-StartupText -Text $_.Exception.Message -MaxLength 600
    }
    Write-StdLog -Level Warn -Source 'SmartToolforDeployment' -Message $text
    $fallback = $null
    if (-not [string]::IsNullOrWhiteSpace($ServerUrl)) {
        # Zły adres w -Server: próbujemy jeszcze bootstrap.json i zapisanego adresu.
        try { $fallback = Get-StdBootstrap -ScriptRoot $Root } catch { }
    }
    if ($null -eq $fallback) {
        # Błędny bootstrap.json na pendrive'ie: adres zapisany wcześniej na tym komputerze (bez pliku z pendrive'a).
        try { $fallback = Get-StdBootstrap } catch { }
    }
    if ($null -ne $fallback) {
        $message = "Nie udało się odczytać adresu serwera STD:`n`n$text`n`nSTD użyje adresu $($fallback.server)."
    } else {
        $message = "Nie udało się odczytać adresu serwera STD:`n`n$text`n`nPodaj adres serwera w oknie logowania."
    }
    try {
        [void](Show-StdMessage -Message $message -Title 'STD' -Image Warning)
    } catch {
        Show-StartupMessage -Message $message -Icon Warning
    }
    return $fallback
}

# ---------------------------------------------------------------------------------------------
# Start
# ---------------------------------------------------------------------------------------------

$script:StdPreview = [bool]$PreviewMode
$stdAppIsWindows = Test-StartupIsWindows
$stdAppScriptPath = $PSCommandPath
$stdAppScriptRoot = $PSScriptRoot
$stdAppDone = Get-StartupMarker -Text $StdRelaunch
$stdAppBound = @{}
foreach ($stdAppKey in @($PSBoundParameters.Keys)) { $stdAppBound[$stdAppKey] = $PSBoundParameters[$stdAppKey] }

$stdAppOrigin = $OriginRoot
if ([string]::IsNullOrWhiteSpace($stdAppOrigin)) { $stdAppOrigin = $stdAppScriptRoot }
$stdAppOrigin = [string]$stdAppOrigin
# Pełna ścieżka: ponownie uruchomiony proces ma inny folder bieżący, więc ścieżka względna by się zgubiła.
try { $stdAppOrigin = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($stdAppOrigin) } catch { }
if ($stdAppOrigin.Length -gt 3) { $stdAppOrigin = $stdAppOrigin.TrimEnd('\', '/') }

Import-Module -Name (Join-Path $stdAppScriptRoot 'STD.psd1') -DisableNameChecking -ErrorAction Stop

if (-not $PreviewMode) {
    try { $Host.UI.RawUI.WindowTitle = ('STD {0} - nie zamykaj tego okna (zamknięcie kończy STD)' -f (Get-StdVersion)) } catch { }
}

# 1. Uprawnienia administratora.
if (-not $PreviewMode -and -not (Test-StdIsAdmin)) {
    if ($stdAppDone -contains 'admin') {
        Show-StartupMessage -Icon Error -Message ('STD wymaga uprawnień administratora. Ponowne uruchomienie nie dało tych uprawnień - ' +
            'zaloguj się na konto administratora albo uruchom Start-STD.cmd poleceniem „Uruchom jako administrator”.')
        exit 1
    }
    # Start z dysku sieciowego: proces administratora nie widzi liter dysków zmapowanych przez zwykłego użytkownika
    # (osobna sesja logowania). Dostaje więc ścieżki UNC i folder roboczy na dysku lokalnym.
    $stdAppRelaunchPath = $stdAppScriptPath
    $stdAppRelaunchDir = $stdAppScriptRoot
    $stdAppRelaunchAdd = $null
    $stdAppNetworkHelp = ('Skopiuj folder STD na pendrive albo na dysk lokalny tego komputera i uruchom Start-STD.cmd stamtąd. ' +
        'Możesz też uruchomić Start-STD.cmd z wiersza poleceń otwartego jako administrator.')
    $stdAppNetworkScript = @{ IsNetwork = $false; UncPath = $null }
    $stdAppNetworkOrigin = @{ IsNetwork = $false; UncPath = $null }
    if ($stdAppIsWindows) {
        $stdAppNetworkScript = Get-StdNetworkPath -Path $stdAppScriptPath
        $stdAppNetworkOrigin = Get-StdNetworkPath -Path $stdAppOrigin
    }
    $stdAppNetwork = ($stdAppNetworkScript.IsNetwork -or $stdAppNetworkOrigin.IsNetwork)
    if ($stdAppNetwork) {
        $stdAppNetworkUnknown = $null
        if ($stdAppNetworkScript.IsNetwork -and $null -eq $stdAppNetworkScript.UncPath) { $stdAppNetworkUnknown = $stdAppScriptRoot }
        elseif ($stdAppNetworkOrigin.IsNetwork -and $null -eq $stdAppNetworkOrigin.UncPath) { $stdAppNetworkUnknown = $stdAppOrigin }
        if ($null -ne $stdAppNetworkUnknown) {
            Show-StartupMessage -Icon Error -Message ((('STD uruchomiono z dysku sieciowego ({0}), którego proces administratora nie zobaczy, ' +
                'a nie udało się ustalić, na jaki udział sieciowy wskazuje ten dysk.') -f $stdAppNetworkUnknown) + "`n`n" + $stdAppNetworkHelp)
            exit 1
        }
        if ($stdAppNetworkScript.IsNetwork) { $stdAppRelaunchPath = [string]$stdAppNetworkScript.UncPath }
        $stdAppRelaunchOrigin = $stdAppOrigin
        if ($stdAppNetworkOrigin.IsNetwork) { $stdAppRelaunchOrigin = [string]$stdAppNetworkOrigin.UncPath }
        $stdAppRelaunchAdd = @{ OriginRoot = $stdAppRelaunchOrigin }
        $stdAppRelaunchDir = $env:SystemRoot
    }
    $stdAppStarted = Invoke-StartupRelaunch -ScriptPath $stdAppRelaunchPath -Stage 'admin' -BoundParameters $stdAppBound -Done $stdAppDone `
        -Add $stdAppRelaunchAdd -WorkingDirectory $stdAppRelaunchDir -PassThru
    if (-not $stdAppStarted) {
        Show-StartupMessage -Icon Warning -Message ('STD wymaga uprawnień administratora. Uruchom STD ponownie i wybierz „Tak” w oknie ' +
            'Kontrola konta użytkownika.')
        exit 1
    }
    if ($stdAppStarted -is [System.Diagnostics.Process]) {
        $stdAppExitCode = $null
        # Z udziału sieciowego proces administratora może nie mieć dostępu do pliku (inne konto, hasło do udziału
        # wpisane tylko w tej sesji) - wtedy kończy się od razu i bez komunikatu. Czekamy chwilę na taki koniec.
        if ($stdAppNetwork) { $stdAppExitCode = Get-StartupEarlyExitCode -Process $stdAppStarted -Milliseconds 5000 } else { $stdAppStarted.Dispose() }
        if ($null -ne $stdAppExitCode -and $stdAppExitCode -ne 0) {
            Write-StdLog -Level Warn -Source 'SmartToolforDeployment' -Message ("STD jako administrator zakończył się od razu (kod {0}) - start z udziału sieciowego {1}." -f $stdAppExitCode, $stdAppRelaunchPath)
            Show-StartupMessage -Icon Error -Message ((('STD nie uruchomił się jako administrator z udziału sieciowego {0} (kod wyjścia {1}). ' +
                'Konto administratora prawdopodobnie nie ma dostępu do tego udziału.') -f $stdAppRelaunchPath, $stdAppExitCode) + "`n`n" + $stdAppNetworkHelp)
            exit 1
        }
    }
    exit 0
}

# 2. Wątek STA (bez niego okna WPF nie działają). Poza Windows okien i tak nie ma.
if (-not $PreviewMode -and $stdAppIsWindows -and [System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA) {
    if ($stdAppDone -contains 'sta') {
        Show-StartupMessage -Icon Error -Message 'Okna STD wymagają trybu STA, a ponowne uruchomienie go nie włączyło. Uruchom STD przez Start-STD.cmd.'
        exit 1
    }
    $stdAppStarted = Invoke-StartupRelaunch -ScriptPath $stdAppScriptPath -Stage 'sta' -BoundParameters $stdAppBound -Done $stdAppDone -WorkingDirectory $stdAppScriptRoot
    if (-not $stdAppStarted) {
        Show-StartupMessage -Icon Error -Message 'Nie udało się uruchomić STD w trybie STA. Uruchom STD przez Start-STD.cmd.'
        exit 1
    }
    exit 0
}

# Folder danych STD (%ProgramData%\STD) tylko dla Administratorów i SYSTEM: kod z App\ działa jako
# administrator, więc zwykły użytkownik nie może tam niczego dopisać ani podmienić.
if (-not $PreviewMode -and $stdAppIsWindows) {
    try {
        Protect-StdDataFolder -Path (Get-StdDataRoot)
    } catch {
        Write-StdLog -Level Warn -Source 'SmartToolforDeployment' -Message (Protect-StartupText -Text $_.Exception.Message -MaxLength 600)
    }
}

# Jedna instancja STD na komputerze (przed kopią do App\: druga kopia naraz zepsułaby pierwszą). Blokadę trzyma
# proces, który pokazuje okna, do końca pracy. Proces z kopii w App\<wersja> czeka chwilę, bo proces
# z pendrive'a oddaje blokadę dopiero po jego uruchomieniu.
if (-not $PreviewMode) {
    $stdAppInstanceWait = 0
    if ($stdAppDone -contains 'copy') { $stdAppInstanceWait = 15 }
    if (-not (Enter-StdAppInstance -WaitSec $stdAppInstanceWait)) {
        Show-StartupMessage -Icon Information -Message "STD już działa na tym komputerze.`n`nPrzełącz się do otwartego okna STD (pasek zadań)."
        exit 0
    }
}

# 3. Kopia do %ProgramData%\STD\App\<wersja> i ponowne uruchomienie stamtąd.
if (-not $PreviewMode -and -not $NoCopy) {
    if (Test-StdInAppFolder -Path $stdAppScriptRoot) {
        Write-StdLog -Level Debug -Source 'SmartToolforDeployment' -Message "STD działa z kopii $stdAppScriptRoot."
    } elseif ($stdAppDone -contains 'copy') {
        # Ten proces miał już działać z kopii, a nie działa: nie kopiujemy drugi raz (bez pętli).
        Write-StdLog -Level Warn -Source 'SmartToolforDeployment' -Message "STD po przejściu do kopii działa z $stdAppScriptRoot, a nie z folderu App - pracuję z bieżącego folderu."
    } else {
        # Każdy problem z kopią albo ponownym uruchomieniem: STD działa dalej z bieżącego folderu.
        $stdAppCopyProblem = $null
        $stdAppTarget = $null
        try {
            $stdAppTarget = Sync-StdAppFolder -Source $stdAppScriptRoot
        } catch {
            $stdAppCopyProblem = 'Nie udało się skopiować STD na dysk komputera: ' + $_.Exception.Message
        }
        if ($null -ne $stdAppTarget) {
            $stdAppTargetScript = Join-Path $stdAppTarget 'SmartToolforDeployment.ps1'
            $stdAppStarted = $false
            try {
                if (-not (Test-Path -LiteralPath $stdAppTargetScript -PathType Leaf)) {
                    throw "Brak pliku $stdAppTargetScript."
                }
                $stdAppStarted = Invoke-StartupRelaunch -ScriptPath $stdAppTargetScript -Stage 'copy' -BoundParameters $stdAppBound -Done $stdAppDone `
                    -Add @{ OriginRoot = $stdAppOrigin } -WorkingDirectory $stdAppTarget
                if (-not $stdAppStarted) { $stdAppCopyProblem = 'Nie udało się uruchomić STD z kopii na dysku komputera.' }
            } catch {
                $stdAppCopyProblem = 'Nie udało się uruchomić STD z kopii na dysku komputera: ' + $_.Exception.Message
            }
            if ($stdAppStarted) {
                # Blokadę jednej instancji przejmuje proces z kopii w App\<wersja>.
                Exit-StdAppInstance
                exit 0
            }
        }
        if ($null -ne $stdAppCopyProblem) {
            $stdAppCopyProblem = Protect-StartupText -Text $stdAppCopyProblem -MaxLength 600
            Write-StdLog -Level Warn -Source 'SmartToolforDeployment' -Message $stdAppCopyProblem
            Show-StartupMessage -Icon Warning -Message ("$stdAppCopyProblem`n`n" +
                'STD działa dalej z bieżącego folderu - nie wyjmuj pendrive''a do końca wdrożenia.')
        }
    }
}

# 4. Moduł jest wczytany; teraz protokoły TLS, okna i motyw.
Initialize-StdNetwork
. (Join-Path (Join-Path $stdAppScriptRoot 'GUI') 'Std.UI.ps1')
[void](Initialize-StdTheme)

# 5. Podgląd okien i testy: wszystko wczytane, bez pokazywania okien.
if ($PreviewMode) {
    $script:StdPreview = $true
    return
}

# Dziennik okien: błędy okien (Show-StdUiError pokazuje tę ścieżkę) i komunikaty programu trafiają do
# Logs\ui-<data>.log. Bez niego wpisy z okien logowania i sieci nie trafiałyby do żadnego pliku.
$script:StdUiLogPath = $null
try {
    $stdAppUiLog = Join-Path (Get-StdPath -Kind Logs) ('ui-{0}.log' -f (Get-Date).ToString('yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture))
    Register-StdFileLogSink -Name 'StdUi' -Path $stdAppUiLog
    $script:StdUiLogPath = $stdAppUiLog
} catch {
    Write-StdLog -Level Warn -Source 'SmartToolforDeployment' -Message ('Nie udało się otworzyć dziennika okien: ' + $_.Exception.Message)
}

foreach ($stdAppScreen in @('Show-StdConnectWindow', 'Show-StdLoginWindow', 'Show-StdMainWindow')) {
    if (-not (Get-Command -Name $stdAppScreen -ErrorAction SilentlyContinue)) {
        throw "[STD-UI-SCREEN] Brak ekranu $stdAppScreen (folder GUI\Screens). Skopiuj pełną wersję STD na pendrive."
    }
}

Write-StdLog -Level Info -Source 'SmartToolforDeployment' -Message ("STD {0} uruchomiony z {1} (folder startowy: {2})." -f (Get-StdVersion), $stdAppScriptRoot, $stdAppOrigin)

$stdAppBootstrap = Get-StartupBootstrap -ServerUrl $Server -Root $stdAppOrigin

if (-not (Test-StdNetworkAvailable)) {
    $stdAppConnect = [string](Select-StartupResult -Output @(Show-StdConnectWindow -OriginRoot $stdAppOrigin))
    if ($stdAppConnect -ne 'Connected' -and $stdAppConnect -ne 'Skipped') {
        # Okno zamknięte bez wyboru: technik kończy pracę.
        return
    }
}

$stdAppLogin = Select-StartupResult -Output @(Show-StdLoginWindow -Bootstrap $stdAppBootstrap -OriginRoot $stdAppOrigin)
if ($null -eq $stdAppLogin) {
    # Okno logowania zamknięte: koniec pracy.
    return
}

# Plik startowy dla okna głównego: ten, którego użyło logowanie (także adres wpisany w oknie). Bez ponownego
# czytania bootstrap.json - błędny plik na pendrive'ie zasłoniłby adres wpisany przy logowaniu.
if ($stdAppLogin -is [System.Collections.IDictionary] -and $stdAppLogin.Contains('Bootstrap') -and $null -ne $stdAppLogin['Bootstrap']) {
    $stdAppBootstrap = $stdAppLogin['Bootstrap']
} elseif ($null -eq $stdAppBootstrap) {
    # Adres wpisany w oknie logowania jest już zapisany (Save-StdBootstrapLocal) - tylko ten plik, bez pendrive'a.
    try {
        $stdAppBootstrap = Get-StdBootstrap
    } catch {
        Write-StdLog -Level Warn -Source 'SmartToolforDeployment' -Message (Protect-StartupText -Text $_.Exception.Message -MaxLength 600)
    }
}

$null = Show-StdMainWindow -Session $stdAppLogin.Session -Index $stdAppLogin.Index -IndexSource $stdAppLogin.Source `
    -Bootstrap $stdAppBootstrap -OriginRoot $stdAppOrigin

# ---------------------------------------------------------------------------------------------
# Siatka bezpieczeństwa: każdy nieobsłużony błąd w tym skrypcie (także w oknach i funkcjach, które woła)
# trafia tutaj. W trybie podglądu błąd idzie dalej do narzędzia podglądu (i do testów).
# ---------------------------------------------------------------------------------------------
trap {
    if ($PreviewMode) { break }
    $stdAppCrash = $_
    $stdAppCrashPath = $null
    try { $stdAppCrashPath = Write-StartupCrashLog -ErrorRecord $stdAppCrash } catch { }
    $stdAppCrashText = 'nieznany błąd'
    try { $stdAppCrashText = Protect-StartupText -Text $stdAppCrash.Exception.Message -MaxLength 800 } catch { }
    $stdAppCrashMessage = "STD napotkał nieoczekiwany błąd i musi zakończyć pracę.`n`n$stdAppCrashText"
    if ($null -ne $stdAppCrashPath) {
        $stdAppCrashMessage += "`n`nSzczegóły zapisano w pliku:`n$stdAppCrashPath`n`nDołącz ten plik do zgłoszenia problemu."
    }
    try { Write-StdLog -Level Error -Source 'SmartToolforDeployment' -Message $stdAppCrashText } catch { }
    try { Show-StartupMessage -Message $stdAppCrashMessage -Title 'STD - błąd krytyczny' -Icon Error } catch { }
    exit 1
}
