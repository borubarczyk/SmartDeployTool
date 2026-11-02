# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.UI.ps1 - wczytanie okien STD 4 (SPEC 7). Program (SmartToolforDeployment.ps1), testy okien
# i podgląd okien dot-source'ują ten plik po imporcie modułu STD:
#
#     Import-Module "$PSScriptRoot\STD.psd1"
#     . "$PSScriptRoot\GUI\Std.UI.ps1"
#
# Plik:
# 1. wczytuje biblioteki WPF i Windows Forms (bez nich [STD-UI-WPF]);
# 2. importuje moduł STD, jeśli nikt go jeszcze nie zaimportował;
# 3. wczytuje motyw (Std.Theme.ps1);
# 4. definiuje zadania w tle dla okien (Start-StdUiTask, SPEC 7.2) i obsługę błędów w handlerach
#    zdarzeń (Invoke-StdUiAction);
# 5. wczytuje wszystkie ekrany z GUI\Screens\*.ps1 (brak folderu = brak ekranów).
#
# Zasady dla kodu okien:
# - każdy handler zdarzenia (Add_Click, Add_SelectionChanged, ...) opakowujemy w Invoke-StdUiAction,
#   żeby błąd pokazał się w oknie komunikatu i trafił do logu, zamiast zamknąć okno;
# - dłuższa praca (sieć, pobieranie konfiguracji, netsh, CIM) idzie przez Start-StdUiTask - okno
#   nigdy się nie zawiesza. W kodzie okien nie czekamy (bez usypiania wątku okna) i nie przetwarzamy
#   ręcznie kolejki zdarzeń okna (SPEC 7.2);
# - handlery widzą zmienne funkcji, która wywołała ShowDialog(), tylko dopóki ta funkcja działa.
#   Stan okna najlepiej trzymać w jednym obiekcie [hashtable] (np. $ui) tworzonym przed ShowDialog().

# ---------------------------------------------------------------------------------------------
# Biblioteki i moduł
# ---------------------------------------------------------------------------------------------

$script:StdUiRoot = $PSScriptRoot
$script:StdUiModuleManifest = Join-Path (Split-Path -Parent $PSScriptRoot) 'STD.psd1'

foreach ($stdUiAssembly in @('PresentationFramework', 'PresentationCore', 'WindowsBase', 'System.Xaml', 'System.Windows.Forms', 'System.Drawing')) {
    try {
        Add-Type -AssemblyName $stdUiAssembly -ErrorAction Stop
    } catch {
        throw "[STD-UI-WPF] Nie udało się wczytać biblioteki $stdUiAssembly. Okna STD działają tylko w Windows (.NET Framework 4.8, Windows PowerShell 5.1). Szczegóły: $($_.Exception.Message)"
    }
}
Remove-Variable -Name stdUiAssembly -ErrorAction SilentlyContinue
try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }

if (-not (Get-Command -Name 'Write-StdLog' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $script:StdUiModuleManifest -ErrorAction Stop
}

. (Join-Path $PSScriptRoot 'Std.Theme.ps1')

if (-not (Get-Variable -Name 'StdUiTaskCounter' -Scope Script -ErrorAction SilentlyContinue)) {
    $script:StdUiTaskCounter = 0
}
# Plik dziennika okien (ustawia go program, gdy rejestruje odbiornik logu do pliku). Pusty = okna z błędem
# mówią tylko o dzienniku STD, bez ścieżki.
if (-not (Get-Variable -Name 'StdUiLogPath' -Scope Script -ErrorAction SilentlyContinue)) {
    $script:StdUiLogPath = $null
}

# ---------------------------------------------------------------------------------------------
# Błędy w oknach
# ---------------------------------------------------------------------------------------------

function Format-StdUiErrorText {
    <#
    .SYNOPSIS
    Krótki, zamaskowany tekst błędu dla okna komunikatu i logu: komunikat wyjątku bez opakowania
    "Exception calling ... with N argument(s)", z najgłębszą przyczyną (np. błąd certyfikatu).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        $ErrorObject,
        [int]$MaxLength = 600
    )
    $text = ''
    $ex = $null
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ex = $ErrorObject.Exception }
    elseif ($ErrorObject -is [System.Exception]) { $ex = $ErrorObject }
    if ($null -ne $ex) {
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $text = [string]$ex.Message
        $innermost = $ex
        while ($null -ne $innermost.InnerException) { $innermost = $innermost.InnerException }
        if (-not [object]::ReferenceEquals($innermost, $ex) -and -not [string]::IsNullOrEmpty($innermost.Message) -and -not $text.Contains($innermost.Message)) {
            $text = $text + ' ' + $innermost.Message
        }
    } elseif ($null -ne $ErrorObject) {
        $text = [string]$ErrorObject
    }
    if ($null -eq $text) { $text = '' }
    $text = ($text -replace '\s+', ' ').Trim()
    if ($text.Length -gt $MaxLength) { $text = $text.Substring(0, $MaxLength) + '...' }
    return (Protect-StdLogText -Text $text)
}

function Show-StdUiError {
    <#
    .SYNOPSIS
    Zapisuje błąd w logu (poziom Error, ze śladem stosu skryptu) i pokazuje go w oknie komunikatu.
    Gdy okna nie da się pokazać, zostaje sam wpis w logu. Okno podaje plik dziennika ($script:StdUiLogPath),
    gdy program go ustawił.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        $ErrorObject,
        [string]$Context = 'Wystąpił nieoczekiwany błąd.',
        [string]$Source = 'UI',
        [System.Windows.Window]$Owner,
        [string]$Title = 'Błąd'
    )
    # Zmienne prywatne: w trakcie ShowDialog() działają handlery i timery innych okien, a ich bloki
    # skryptu widzą zmienne z całego stosu wywołań (np. $Title zasłoniłby zmienną ekranu).
    $private:stdErrText = Format-StdUiErrorText -ErrorObject $ErrorObject
    $private:stdErrLog = '{0} {1}' -f $Context, $private:stdErrText
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord] -and -not [string]::IsNullOrEmpty($ErrorObject.ScriptStackTrace)) {
        $private:stdErrLog = '{0} | {1}' -f $private:stdErrLog, (($ErrorObject.ScriptStackTrace -replace '\r?\n', ' <- ').Trim())
    }
    try { Write-StdLog -Level Error -Source $Source -Message $private:stdErrLog } catch { }
    $private:stdErrWhere = 'Szczegóły są w dzienniku STD.'
    if (-not [string]::IsNullOrWhiteSpace([string]$script:StdUiLogPath)) {
        $private:stdErrWhere = "Szczegóły są w dzienniku STD:`n{0}" -f [string]$script:StdUiLogPath
    }
    $private:stdErrMessage = "{0}`n`n{1}`n`n{2}" -f $Context, $private:stdErrText, $private:stdErrWhere
    $private:stdErrTitle = $Title
    $private:stdErrOwner = $Owner
    Remove-Variable -Name ErrorObject, Context, Source, Owner, Title -Scope Local -ErrorAction SilentlyContinue
    try {
        [void](Show-StdMessage -Message $private:stdErrMessage -Title $private:stdErrTitle -Image Error -Owner $private:stdErrOwner)
    } catch {
        try { Write-StdLog -Level Warn -Source 'UI' -Message ('Nie udało się pokazać okna z błędem: {0}' -f $_.Exception.Message) } catch { }
    }
}

function Invoke-StdUiAction {
    <#
    .SYNOPSIS
    Uruchamia kod handlera zdarzenia okna w try/catch. Błąd trafia do logu i do okna komunikatu
    (Show-StdUiError), a okno działa dalej. Używać w KAŻDYM handlerze:

        $btnLogin.Add_Click({ Invoke-StdUiAction -Action { ... } })

    .DESCRIPTION
    Kod akcji działa z $ErrorActionPreference = 'Stop', więc także zwykły błąd polecenia
    (np. Copy-Item na brakującym pliku) przerywa akcję i jest pokazany, zamiast przejść po cichu.
    Akcja widzi zmienne handlera i funkcji okna ($this, $_, parametry handlera), a nie zmienne tej
    funkcji. Zwraca to, co zwróciła akcja (po błędzie: nic).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [scriptblock]$Action,
        [string]$Source = 'UI',
        [System.Windows.Window]$Owner,
        [string]$ErrorTitle = 'Błąd',
        [string]$ErrorContext = 'Nie udało się wykonać tej czynności.'
    )
    # Zmienne prywatne: akcja nie może zobaczyć parametrów tej funkcji zamiast zmiennych okna.
    $private:stdActBlock = $Action
    $private:stdActSource = $Source
    $private:stdActOwner = $Owner
    $private:stdActTitle = $ErrorTitle
    $private:stdActContext = $ErrorContext
    Remove-Variable -Name Action, Source, Owner, ErrorTitle, ErrorContext -Scope Local -ErrorAction SilentlyContinue
    $ErrorActionPreference = 'Stop'
    try {
        & $private:stdActBlock
    } catch {
        Show-StdUiError -ErrorObject $_ -Context $private:stdActContext -Source $private:stdActSource -Owner $private:stdActOwner -Title $private:stdActTitle
    }
}

# ---------------------------------------------------------------------------------------------
# Zadania w tle (SPEC 7.2)
# ---------------------------------------------------------------------------------------------

# Skrypt uruchamiany w runspace zadania. Wynik i błąd zapisuje we wspólnej tabeli ($StdUiTaskShared),
# wpisy logu modułu STD z tego runspace'u - do kolejki, z której okno przepisuje je do swojego logu
# (maski i odbiorniki logu okna działają więc także dla zadań). Skrypt zadania widzi $Arguments;
# może też zadeklarować param($Arguments).
$script:StdUiTaskWrapper = @'
param($StdUiTaskShared, $StdUiTaskScript, $StdUiTaskModule, $Arguments)
try {
    Import-Module -Name $StdUiTaskModule -ErrorAction Stop
    $StdUiTaskLogQueue = $StdUiTaskShared.Log
    $StdUiTaskLogSink = { $StdUiTaskLogQueue.Enqueue($args[0]) }.GetNewClosure()
    Register-StdLogSink -Name 'StdUiTask' -ScriptBlock $StdUiTaskLogSink
    $StdUiTaskBlock = [scriptblock]::Create($StdUiTaskScript)
    $StdUiTaskHasParam = $false
    if ($null -ne $StdUiTaskBlock.Ast.ParamBlock) {
        foreach ($StdUiTaskParam in $StdUiTaskBlock.Ast.ParamBlock.Parameters) {
            if ($StdUiTaskParam.Name.VariablePath.UserPath -eq 'Arguments') { $StdUiTaskHasParam = $true }
        }
    }
    if ($StdUiTaskHasParam) {
        $StdUiTaskOutput = @(& $StdUiTaskBlock -Arguments $Arguments)
    } else {
        $StdUiTaskOutput = @(& $StdUiTaskBlock)
    }
    if ($StdUiTaskOutput.Count -eq 1) {
        $StdUiTaskShared.Result = $StdUiTaskOutput[0]
    } elseif ($StdUiTaskOutput.Count -gt 1) {
        $StdUiTaskShared.Result = $StdUiTaskOutput
    }
} catch {
    $StdUiTaskShared.Error = $_
} finally {
    try { Unregister-StdLogSink -Name 'StdUiTask' } catch { }
}
'@

function New-StdUiTaskRunspace {
    <#
    .SYNOPSIS
    Część Start-StdUiTask bez WPF: nowy runspace z modułem STD, start skryptu zadania (asynchronicznie).
    Zwraca uchwyt zadania ([hashtable]). Wynik odbiera Complete-StdUiTaskRun.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$ScriptText,
        [hashtable]$Arguments,
        [string]$Name = 'Zadanie w tle'
    )
    if (-not (Test-Path -LiteralPath $script:StdUiModuleManifest -PathType Leaf)) {
        throw "[STD-UI-TASK] Brak pliku modułu STD: $($script:StdUiModuleManifest)"
    }
    if ($null -eq $Arguments) { $Arguments = @{} }
    $script:StdUiTaskCounter = [int]$script:StdUiTaskCounter + 1
    $shared = [hashtable]::Synchronized(@{
        Result = $null
        Error  = $null
        Log    = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    })
    $runspace = $null
    $ps = $null
    try {
        $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        try { $runspace.ApartmentState = [System.Threading.ApartmentState]::MTA } catch { }
        $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $runspace.Open()
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript($script:StdUiTaskWrapper)
        [void]$ps.AddArgument($shared)
        [void]$ps.AddArgument($ScriptText)
        [void]$ps.AddArgument($script:StdUiModuleManifest)
        [void]$ps.AddArgument($Arguments)
        $async = $ps.BeginInvoke()
    } catch {
        if ($null -ne $ps) { try { $ps.Dispose() } catch { } }
        if ($null -ne $runspace) { try { $runspace.Dispose() } catch { } }
        throw "[STD-UI-TASK] Nie udało się uruchomić zadania w tle '$Name': $($_.Exception.Message)"
    }
    return @{
        Id          = $script:StdUiTaskCounter
        Name        = $Name
        PowerShell  = $ps
        Runspace    = $runspace
        AsyncResult = $async
        Shared      = $shared
        StartedAt   = Get-Date
        FinishedAt  = $null
        Status      = 'Running'     # Running, Completed, Failed, Stopped
        Finished    = $false        # runspace zwolniony
        Abandoned   = $false        # Stop-StdUiTask: wynik nie trafi do OnComplete / OnError
        OwnerClosed = $false
        OnComplete  = $null
        OnError     = $null
        Owner       = $null
        Timer       = $null
    }
}

function Receive-StdUiTaskLog {
    # Przepisuje wpisy logu z runspace'u zadania do logu okna (najwyżej -Max naraz, żeby nie
    # blokować okna przy bardzo gadatliwym zadaniu).
    param(
        [Parameter(Mandatory = $true)][hashtable]$StdUiTask,
        [int]$Max = 500
    )
    $queue = $StdUiTask.Shared.Log
    if ($null -eq $queue) { return }
    $record = $null
    $count = 0
    while ($count -lt $Max -and $queue.TryDequeue([ref]$record)) {
        $count++
        try {
            $level = [string]$record.Level
            if ($level -notin @('Info', 'Warn', 'Error', 'Debug')) { $level = 'Info' }
            $source = [string]$record.Source
            if ([string]::IsNullOrEmpty($source)) { $source = $StdUiTask.Name }
            Write-StdLog -Level $level -Source $source -Message ([string]$record.Message)
        } catch { }
    }
}

function Complete-StdUiTaskRun {
    <#
    .SYNOPSIS
    Odbiera wynik zakończonego zadania i zwalnia jego runspace. Zwraca
    @{ Status = 'Completed'|'Failed'|'Stopped'; Result; ErrorRecord }. Bez WPF (testowalne).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [hashtable]$StdUiTask
    )
    $outcome = @{ Status = 'Completed'; Result = $null; ErrorRecord = $null }
    if ($StdUiTask.Finished) {
        $outcome.Status = $StdUiTask.Status
        return $outcome
    }
    $ps = $StdUiTask.PowerShell
    $endError = $null
    try { [void]$ps.EndInvoke($StdUiTask.AsyncResult) } catch { $endError = $_ }
    Receive-StdUiTaskLog -StdUiTask $StdUiTask -Max ([int]::MaxValue)
    # Błędy, które nie przerwały skryptu (np. Get-Item bez -ErrorAction Stop): tylko do logu.
    try {
        foreach ($nonTerminating in @($ps.Streams.Error)) {
            Write-StdLog -Level Warn -Source $StdUiTask.Name -Message (Format-StdUiErrorText -ErrorObject $nonTerminating)
        }
    } catch { }
    $shared = $StdUiTask.Shared
    $invocation = $ps.InvocationStateInfo
    if ($StdUiTask.Abandoned -or $invocation.State -eq [System.Management.Automation.PSInvocationState]::Stopped) {
        $outcome.Status = 'Stopped'
    } elseif ($null -ne $shared.Error) {
        $outcome.Status = 'Failed'
        $outcome.ErrorRecord = $shared.Error
    } elseif ($invocation.State -eq [System.Management.Automation.PSInvocationState]::Failed -or $null -ne $endError) {
        $outcome.Status = 'Failed'
        $reason = $invocation.Reason
        if ($null -eq $reason -and $null -ne $endError) { $reason = $endError.Exception }
        if ($reason -is [System.Management.Automation.IContainsErrorRecord]) {
            $outcome.ErrorRecord = $reason.ErrorRecord
        } else {
            $outcome.ErrorRecord = New-Object System.Management.Automation.ErrorRecord -ArgumentList $reason, 'StdUiTaskFailed', ([System.Management.Automation.ErrorCategory]::NotSpecified), $null
        }
    } else {
        $outcome.Result = $shared.Result
    }
    try { $ps.Dispose() } catch { }
    try { $StdUiTask.Runspace.Dispose() } catch { }
    $shared.Result = $null
    $StdUiTask.Finished = $true
    $StdUiTask.Status = $outcome.Status
    $StdUiTask.FinishedAt = Get-Date
    return $outcome
}

function Stop-StdUiTask {
    <#
    .SYNOPSIS
    Porzuca zadanie w tle: wynik nie trafi do OnComplete / OnError, a runspace jest zatrzymywany
    i zwalniany (bez czekania - okno się nie zawiesza; zwolnienie kończy timer zadania).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [hashtable]$Task
    )
    if ($Task.Finished) { return }
    $Task.Abandoned = $true
    if ($null -ne $Task.AsyncResult -and $Task.AsyncResult.IsCompleted) {
        if ($null -ne $Task.Timer) { try { $Task.Timer.Stop() } catch { } }
        [void](Complete-StdUiTaskRun -StdUiTask $Task)
        return
    }
    try { [void]$Task.PowerShell.BeginStop($null, $null) } catch { }
}

function Invoke-StdUiTaskPoll {
    # Jeden "tik" timera zadania (co 200 ms, w wątku okna): log, porzucenie po zamknięciu okna
    # właściciela, a po zakończeniu - OnComplete($result) albo OnError($errorRecord).
    # Zwraca $true, gdy zadanie jest już obsłużone.
    param([Parameter(Mandatory = $true)][hashtable]$StdUiTask)
    if ($StdUiTask.Finished) {
        if ($null -ne $StdUiTask.Timer) { try { $StdUiTask.Timer.Stop() } catch { } }
        return $true
    }
    Receive-StdUiTaskLog -StdUiTask $StdUiTask
    if ($StdUiTask.OwnerClosed -and -not $StdUiTask.Abandoned) {
        Write-StdLog -Level Debug -Source 'Start-StdUiTask' -Message ("Okno zamknięte - wynik zadania '{0}' zostanie pominięty." -f $StdUiTask.Name)
        Stop-StdUiTask -Task $StdUiTask
        if ($StdUiTask.Finished) { return $true }
    }
    if (-not $StdUiTask.AsyncResult.IsCompleted) { return $false }
    if ($null -ne $StdUiTask.Timer) { try { $StdUiTask.Timer.Stop() } catch { } }
    $private:stdUiOutcome = Complete-StdUiTaskRun -StdUiTask $StdUiTask
    if ($StdUiTask.Abandoned) { return $true }

    # Zmienne prywatne: OnComplete / OnError widzą zmienne okna, nie tej funkcji.
    $private:stdUiTaskRef = $StdUiTask
    Remove-Variable -Name StdUiTask -Scope Local -ErrorAction SilentlyContinue
    $ErrorActionPreference = 'Stop'
    if ($private:stdUiOutcome.Status -eq 'Completed') {
        try {
            [void](& $private:stdUiTaskRef.OnComplete $private:stdUiOutcome.Result)
        } catch {
            Show-StdUiError -ErrorObject $_ -Context ("Nie udało się obsłużyć wyniku zadania '{0}'." -f $private:stdUiTaskRef.Name) -Source 'Start-StdUiTask' -Owner $private:stdUiTaskRef.Owner
        }
        return $true
    }
    if ($null -ne $private:stdUiTaskRef.OnError) {
        try {
            Write-StdLog -Level Error -Source $private:stdUiTaskRef.Name -Message ("Zadanie '{0}' zakończyło się błędem: {1}" -f $private:stdUiTaskRef.Name, (Format-StdUiErrorText -ErrorObject $private:stdUiOutcome.ErrorRecord))
        } catch { }
        try {
            [void](& $private:stdUiTaskRef.OnError $private:stdUiOutcome.ErrorRecord)
        } catch {
            Show-StdUiError -ErrorObject $_ -Context ("Nie udało się obsłużyć błędu zadania '{0}'." -f $private:stdUiTaskRef.Name) -Source 'Start-StdUiTask' -Owner $private:stdUiTaskRef.Owner
        }
    } else {
        Show-StdUiError -ErrorObject $private:stdUiOutcome.ErrorRecord -Context ("Zadanie '{0}' nie powiodło się." -f $private:stdUiTaskRef.Name) -Source $private:stdUiTaskRef.Name -Owner $private:stdUiTaskRef.Owner
    }
    return $true
}

function New-StdUiTaskOwnerClosedHandler {
    # Handler Closed okna właściciela: oznacza zadanie do porzucenia (domknięcie zabiera tylko zadanie).
    param([Parameter(Mandatory = $true)][hashtable]$StdUiTask)
    return {
        param($s, $e)
        $StdUiTask.OwnerClosed = $true
    }.GetNewClosure()
}

function Start-StdUiTask {
    <#
    .SYNOPSIS
    Wykonuje tekst skryptu w osobnym runspace (z modułem STD) i po zakończeniu woła OnComplete($result)
    albo OnError($errorRecord) w wątku okna. Okno w tym czasie działa normalnie.
    .DESCRIPTION
    - Skrypt NIE widzi zmiennych okna: wszystko, czego potrzebuje, przekazujemy w -Arguments
      (w skrypcie: $Arguments, np. $Arguments.Session). Hasła tylko jako SecureString / PSCredential.
    - Wynik: to, co skrypt wypisał (jeden obiekt -> ten obiekt, kilka -> tablica, nic -> $null).
    - Błąd: wyjątek w skrypcie (throw albo -ErrorAction Stop). Zwykłe błędy poleceń trafiają tylko
      do logu. Bez -OnError błąd jest zapisany w logu i pokazany w oknie komunikatu.
    - Wpisy Write-StdLog ze skryptu trafiają do logu okna.
    - DispatcherTimer co 200 ms sprawdza zakończenie, więc OnComplete / OnError działają w wątku okna
      i mogą zmieniać kontrolki. Błąd w nich jest pokazany jak w Invoke-StdUiAction.
    - -Owner: po zamknięciu tego okna wynik jest pomijany, a runspace zatrzymywany.
    Zwraca uchwyt zadania (dla Stop-StdUiTask).
    .EXAMPLE
    Start-StdUiTask -Name 'Logowanie' -Owner $win -Arguments @{ Session = $session } -ScriptText {
        Get-StdClientIndex -Session $Arguments.Session
    } -OnComplete { param($Result) $ui.Index = $Result } -OnError { param($ErrorRecord) $ui.Error.Text = $ErrorRecord.Exception.Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$ScriptText,
        [hashtable]$Arguments,
        [Parameter(Mandatory = $true)]
        [scriptblock]$OnComplete,
        [scriptblock]$OnError,
        [System.Windows.Window]$Owner,
        [string]$Name = 'Zadanie w tle'
    )
    $task = New-StdUiTaskRunspace -ScriptText $ScriptText -Arguments $Arguments -Name $Name
    $task.OnComplete = $OnComplete
    $task.OnError = $OnError
    $task.Owner = $Owner
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(200)
    # Stan zadania w Tag timera: zmienne tej funkcji znikną, zanim timer zadziała.
    $timer.Tag = $task
    $task.Timer = $timer
    $timer.Add_Tick({
        param($stdUiTimerSender, $stdUiTimerArgs)
        try {
            [void](Invoke-StdUiTaskPoll -StdUiTask $stdUiTimerSender.Tag)
        } catch {
            try { $stdUiTimerSender.Stop() } catch { }
            try { Write-StdLog -Level Error -Source 'Start-StdUiTask' -Message ('Błąd obsługi zadania w tle: {0}' -f $_.Exception.Message) } catch { }
        }
    })
    if ($null -ne $Owner) {
        $Owner.Add_Closed((New-StdUiTaskOwnerClosedHandler -StdUiTask $task))
    }
    $timer.Start()
    Write-StdLog -Level Debug -Source 'Start-StdUiTask' -Message ("Start zadania w tle '{0}' (nr {1})." -f $Name, $task.Id)
    return $task
}

# ---------------------------------------------------------------------------------------------
# Ekrany (GUI\Screens\*.ps1)
# ---------------------------------------------------------------------------------------------

$stdUiScreensDir = Join-Path $PSScriptRoot 'Screens'
if (Test-Path -LiteralPath $stdUiScreensDir -PathType Container) {
    foreach ($stdUiScreenFile in @(Get-ChildItem -LiteralPath $stdUiScreensDir -Filter '*.ps1' -File | Where-Object { $_.Name -notlike '*.Tests.ps1' } | Sort-Object -Property Name)) {
        . $stdUiScreenFile.FullName
    }
}
Remove-Variable -Name stdUiScreensDir, stdUiScreenFile -ErrorAction SilentlyContinue
