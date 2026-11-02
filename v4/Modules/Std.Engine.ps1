# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.Engine.ps1 - silnik wdrożenia STD (SPEC 6): rejestr kroków, plan („Sprawdź”) i wdrożenie („Wdróż”)
# w osobnym runspace, pauza i przerywanie, plik stanu state.json, log wdrożenia (deploy.log, events.jsonl),
# uruchamianie programów (Invoke-StdProcess), wstrzymanie usypiania (Set-StdKeepAwake) oraz informacje
# o komputerze (Get-StdComputerInfo, SPEC 7.3).
#
# Wymaga Std.Core.ps1 (log, maski, JSON, ścieżki), Std.Crypto.ps1 (pomocnicze odczyty pól i SecureString)
# i Std.Config.ps1 (lista identyfikatorów zadań) - STD.psm1 wczytuje je wcześniej.
#
# Jak to działa:
# - Start-StdEngine (wątek okna) tworzy nowy runspace, ustawia w nim zmienną $StdEngineStart i uruchamia
#   asynchronicznie (BeginInvoke) krótki skrypt startowy: Import-Module STD.psd1 (ten sam folder co ten
#   plik), dot-source każdego pliku Steps\*.ps1 (bloki skryptu kroków nie przechodzą między runspace'ami),
#   a potem Invoke-StdEngineLoop.
# - Jedyny kanał silnik -> okno to kolejka Shared.Events (ConcurrentQueue[object]); okno -> silnik to flagi
#   Shared.PauseRequested / CancelRequested / ForceCancel. Okno o nic nie pyta i nie wywołuje kroków.
# - W runspace silnika jest osobna kopia modułu, więc silnik od nowa rejestruje maski logu dla każdej
#   wartości z Secrets, zanim cokolwiek zapisze do logu.
# - state.json nie zawiera haseł ani niczego z Secrets: zapisujemy tylko pola z białej listy (SPEC 6.3).
#
# Funkcje pomocnicze bez "-Std" w nazwie nie są eksportowane z modułu (Export-ModuleMember *-Std*).

# ---------------------------------------------------------------------------------------------
# Stan wspólny (zakres modułu). Przy ponownym dot-source tego samego pliku nie kasujemy kroków.
# ---------------------------------------------------------------------------------------------
if (-not (Get-Variable -Name 'StdSteps' -Scope Script -ErrorAction SilentlyContinue) -or $null -eq $script:StdSteps) {
    $script:StdSteps = [ordered]@{}
    $script:StdStepSequence = 0
}
if (-not (Get-Variable -Name 'StdKeepAwakeActive' -Scope Script -ErrorAction SilentlyContinue)) {
    $script:StdKeepAwakeActive = $false
}
# Bieżące uruchomienie silnika w tym runspace (jedno naraz); używa go odbiornik logu.
$script:StdEngineRun = $null
$script:StdEngineRoot = $null
if (-not [string]::IsNullOrEmpty($PSCommandPath)) {
    $script:StdEngineRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
}

$script:StdEngineStepIdPattern = '^[A-Za-z][A-Za-z0-9]*(\.[A-Za-z0-9][A-Za-z0-9_-]*)+$'
$script:StdEngineDeploymentIdPattern = '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$'
$script:StdEngineDetectStates = @('Done', 'Todo', 'Blocked')
$script:StdEngineRunResults = @('Ok', 'Error', 'Skipped', 'Cancelled')
$script:StdEngineDetailMax = 2000
$script:StdEngineDefaultSuccessCodes = @(0, 1641, 3010)
$script:StdEngineEventsSink = 'StdEngine.Events'
$script:StdEngineFileSink = 'StdEngine.Deploy'

# Skrypt startowy runspace'u silnika. Jest tekstem, bo bloki skryptu nie przechodzą między runspace'ami.
# Kroki są dot-source'owane na poziomie tego skryptu (zakres globalny runspace'u), więc funkcje pomocnicze
# z plików kroków są widoczne dla bloków Detect / Run przez cały czas działania silnika.
$script:StdEngineBootstrap = @'
$stdStart = $StdEngineStart
function Clear-EngineBootstrapSecrets {
    # Awaria poza pętlą silnika: Complete-EngineRun się nie wykonał, więc Secrets czyścimy tutaj (SPEC 6.3).
    # Loop.Secrets to [hashtable] wywołującego (Synchronized działa na tym samym obiekcie).
    try { if ($null -ne $stdStart.Loop.Secrets) { $stdStart.Loop.Secrets.Clear() } } catch { }
    try { if ($null -ne $stdStart.Loop.ExtraSecrets) { $stdStart.Loop.ExtraSecrets.Clear() } } catch { }
}
function Send-EngineBootstrapFailure {
    # Awaria poza pętlą silnika (moduł się nie wczytał albo pętla rzuciła wyjątek): czyszczenie Secrets,
    # stan Failed i zdarzenie Finished, żeby okno nie czekało w nieskończoność.
    param([string]$Message)
    Clear-EngineBootstrapSecrets
    $shared = $stdStart.Shared
    $now = Get-Date
    $counts = @{ Ok = 0; Skipped = 0; Error = 0; Cancelled = 0 }
    $shared.Result = @{ Status = 'Failed'; Steps = @(); Plan = $null; Counts = $counts; RebootRequired = $false; Error = $Message }
    $shared.Status = 'Failed'
    $shared.Events.Enqueue((New-Object -TypeName PSObject -Property ([ordered]@{ Type = 'Log'; Time = $now; Level = 'Error'; Message = $Message; Source = 'Std.Engine' })))
    $shared.Events.Enqueue((New-Object -TypeName PSObject -Property ([ordered]@{ Type = 'Finished'; Time = $now; Status = 'Failed'; Counts = $counts; RebootRequired = $false; Error = $Message })))
}
try {
    Import-Module -Name $stdStart.ManifestPath -Force -DisableNameChecking -ErrorAction Stop
} catch {
    Send-EngineBootstrapFailure -Message ('[STD-ENGINE-START] Nie udało się wczytać modułu STD w wątku silnika: ' + $_.Exception.Message)
    return
}
$stdLoadErrors = New-Object 'System.Collections.Generic.List[string]'
foreach ($stdStepFile in @($stdStart.StepFiles)) {
    try {
        . $stdStepFile
    } catch {
        $stdLoadErrors.Add(('{0}: {1}' -f [System.IO.Path]::GetFileName($stdStepFile), $_.Exception.Message))
    }
}
$stdLoop = $stdStart.Loop
try {
    Invoke-StdEngineLoop @stdLoop -StepLoadErrors $stdLoadErrors.ToArray()
} catch {
    Clear-EngineBootstrapSecrets
    $stdStatus = [string]$stdStart.Shared.Status
    if ($stdStatus -eq 'Running' -or $stdStatus -eq 'Paused') {
        Send-EngineBootstrapFailure -Message ('[STD-ENGINE-FAILED] Silnik zakończył się błędem: ' + (Protect-StdLogText -Text $_.Exception.Message))
    }
}
'@

# Wstrzymanie usypiania: SetThreadExecutionState działa dla wątku, który ją wywołał (wątek silnika).
$script:StdEnginePowerTypeName = 'StdEngine.PowerState'
$script:StdEnginePowerSource = @'
using System;
using System.Runtime.InteropServices;
namespace StdEngine {
    public static class PowerState {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint SetThreadExecutionState(uint esFlags);
        private const uint EsContinuous = 0x80000000;
        private const uint EsSystemRequired = 0x00000001;
        private const uint EsDisplayRequired = 0x00000002;
        public static bool KeepAwake() { return SetThreadExecutionState(EsContinuous | EsSystemRequired | EsDisplayRequired) != 0; }
        public static bool Restore() { return SetThreadExecutionState(EsContinuous) != 0; }
    }
}
'@

# ---------------------------------------------------------------------------------------------
# Pomocnicze (prywatne)
# ---------------------------------------------------------------------------------------------

function Get-EngineRoot {
    # Folder v4 (ten, w którym są STD.psd1 i Steps\).
    if (-not [string]::IsNullOrEmpty($script:StdEngineRoot)) { return $script:StdEngineRoot }
    $module = $MyInvocation.MyCommand.Module
    if ($null -ne $module -and -not [string]::IsNullOrEmpty($module.ModuleBase)) { return $module.ModuleBase }
    throw '[STD-ENGINE-START] Nie udało się ustalić folderu STD (brak ścieżki pliku Std.Engine.ps1).'
}

function Get-EngineUtcText {
    # Czas UTC w ISO 8601 (do state.json).
    param([AllowNull()]$Time)
    if ($null -eq $Time -or -not ($Time -is [datetime])) { $Time = Get-Date }
    return $Time.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-EngineInvariantNumber {
    param([double]$Value)
    return $Value.ToString('0.###', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Protect-EngineDetail {
    # Tekst szczegółów kroku do zdarzeń, wyniku i state.json: zamaskowany i przycięty.
    param([AllowNull()]$Text)
    if ($null -eq $Text) { return '' }
    $value = Protect-StdLogText -Text ([string]$Text)
    if ($null -eq $value) { return '' }
    if ($value.Length -gt $script:StdEngineDetailMax) { $value = $value.Substring(0, $script:StdEngineDetailMax) + '...' }
    return $value
}

function Get-EngineFlag {
    # Flaga z Shared bez wyliczania kluczy (okno zmienia wartości z innego wątku - wyliczanie kluczy
    # zsynchronizowanego [hashtable] w tym czasie może rzucić wyjątek; odczyt przez indeks jest bezpieczny).
    param([AllowNull()]$Shared, [Parameter(Mandatory = $true)][string]$Name)
    if ($null -eq $Shared) { return $false }
    if ($Shared -is [System.Collections.IDictionary]) { return ($Shared[$Name] -eq $true) }
    $property = $Shared.PSObject.Properties[$Name]
    if ($null -ne $property) { return ($property.Value -eq $true) }
    return $false
}

function Get-EngineAllowedValue {
    # Wartość z listy dozwolonych (bez względu na wielkość liter) w zapisie z listy albo $null.
    param([AllowNull()][string]$Value, [Parameter(Mandatory = $true)][string[]]$Allowed)
    if ([string]::IsNullOrEmpty($Value)) { return $null }
    foreach ($candidate in $Allowed) {
        if ([string]::Equals($candidate, $Value, [System.StringComparison]::OrdinalIgnoreCase)) { return $candidate }
    }
    return $null
}

function Select-EngineResultObject {
    # Ostatni obiekt z wyjścia bloku kroku, który ma pole -Key (krok może przypadkiem wypisać coś jeszcze).
    param([AllowNull()][object[]]$Output, [Parameter(Mandatory = $true)][string]$Key)
    if ($null -eq $Output) { return $null }
    for ($i = $Output.Count - 1; $i -ge 0; $i--) {
        $item = $Output[$i]
        if ($item -is [System.Collections.IDictionary]) {
            foreach ($candidate in @($item.Keys)) {
                if ([string]::Equals([string]$candidate, $Key, [System.StringComparison]::OrdinalIgnoreCase)) { return $item }
            }
        } elseif ($item -is [System.Management.Automation.PSCustomObject]) {
            if ($null -ne $item.PSObject.Properties[$Key]) { return $item }
        }
    }
    return $null
}

function New-EngineDeploymentId {
    # yyyyMMdd-HHmmss-<4 znaki> (małe litery i cyfry).
    $alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789'
    $bytes = New-CryptoRandomBytes -Count 4
    $suffix = New-Object System.Text.StringBuilder
    foreach ($b in $bytes) { [void]$suffix.Append($alphabet[[int]$b % $alphabet.Length]) }
    return ('{0}-{1}' -f (Get-Date).ToString('yyyyMMdd-HHmmss', [System.Globalization.CultureInfo]::InvariantCulture), $suffix.ToString())
}

function Get-EngineSortedSteps {
    # Zarejestrowane kroki w kolejności: faza, Order, kolejność rejestracji.
    $values = @($script:StdSteps.get_Values())
    if ($values.Count -eq 0) { return , @() }
    $sorted = @($values | Sort-Object -Property Phase, Order, Seq)
    return , $sorted
}

function Get-EngineStepInfo {
    # Dane kroku dla Context.Step (bez bloków skryptu).
    param([Parameter(Mandatory = $true)]$Step)
    return @{
        Id         = $Step.Id
        Title      = $Step.Title
        Phase      = $Step.Phase
        Order      = $Step.Order
        TaskId     = $Step.TaskId
        TimeoutMin = $Step.TimeoutMin
        Critical   = [bool]$Step.Critical
    }
}

function Get-EngineStepEnabled {
    # Czy krok jest włączony: -OnlySteps (gdy podane) i Config.Tasks[TaskId] = $true (gdy krok ma TaskId).
    param([Parameter(Mandatory = $true)]$Step, [AllowNull()]$Config, [AllowNull()][string[]]$OnlySteps)
    if ($null -ne $OnlySteps) {
        $selected = $false
        foreach ($id in $OnlySteps) {
            if ([string]::Equals([string]$id, [string]$Step.Id, [System.StringComparison]::OrdinalIgnoreCase)) { $selected = $true; break }
        }
        if (-not $selected) { return @{ Enabled = $false; Reason = 'Krok nie jest na liście wybranych kroków.' } }
    }
    if (-not [string]::IsNullOrEmpty([string]$Step.TaskId)) {
        $tasks = Get-CryptoMember $Config 'Tasks'
        $value = Get-CryptoMember $tasks ([string]$Step.TaskId)
        if (-not ($value -is [bool] -and $value)) {
            return @{ Enabled = $false; Reason = ('Zadanie {0} jest wyłączone.' -f $Step.TaskId) }
        }
    }
    return @{ Enabled = $true; Reason = '' }
}

function Write-EngineUnknownOnlySteps {
    param([AllowNull()][object[]]$Steps, [AllowNull()][string[]]$OnlySteps)
    if ($null -eq $OnlySteps) { return }
    foreach ($id in $OnlySteps) {
        $known = $false
        foreach ($step in @($Steps)) {
            if ([string]::Equals([string]$id, [string]$step.Id, [System.StringComparison]::OrdinalIgnoreCase)) { $known = $true; break }
        }
        if (-not $known) {
            Write-StdLog -Level Warn -Source 'Std.Engine' -Message ("Nieznany krok '{0}' na liście wybranych kroków - pomijam go." -f $id)
        }
    }
}

function Get-EnginePercent {
    param([int]$Done, [int]$Total)
    if ($Total -le 0) { return 100 }
    $value = [int][System.Math]::Floor(100.0 * $Done / $Total)
    if ($value -lt 0) { return 0 }
    if ($value -gt 100) { return 100 }
    return $value
}

function Get-EngineCounts {
    param([AllowNull()]$Records)
    $counts = @{ Ok = 0; Skipped = 0; Error = 0; Cancelled = 0 }
    if ($null -ne $Records) {
        foreach ($record in $Records) {
            $state = [string]$record.State
            if ($counts.ContainsKey($state)) { $counts[$state] = $counts[$state] + 1 }
        }
    }
    return $counts
}

function Register-EngineSecretMasks {
    # Maski logu dla każdej wartości z Secrets: hasło i login PSCredential, SecureString, tekst
    # (np. odszyfrowane sekrety konfiguracji), także w zagnieżdżonych słownikach i listach.
    param([AllowNull()]$Value, [int]$Depth = 0)
    if ($null -eq $Value -or $Depth -gt 8) { return }
    try {
        if ($Value -is [System.Management.Automation.PSCredential]) {
            if ($null -ne $Value.Password -and $Value.Password.Length -gt 0) {
                Register-StdLogMask -Value (ConvertFrom-CryptoSecureString -SecureString $Value.Password)
            }
            # Login do domeny też nie może trafić do logu (decyzja właściciela) - maskujemy pełny login
            # i jego części (DOMENA\login, login@domena).
            $user = [string]$Value.UserName
            if (-not [string]::IsNullOrEmpty($user)) {
                Register-StdLogMask -Value $user
                if ($user.Contains('\')) { Register-StdLogMask -Value ($user.Substring($user.LastIndexOf('\') + 1)) }
                if ($user.Contains('@')) { Register-StdLogMask -Value ($user.Substring(0, $user.IndexOf('@'))) }
            }
            return
        }
        if ($Value -is [System.Net.NetworkCredential]) {
            Register-StdLogMask -Value $Value.Password
            Register-StdLogMask -Value $Value.UserName
            return
        }
        if ($Value -is [System.Security.SecureString]) {
            if ($Value.Length -gt 0) { Register-StdLogMask -Value (ConvertFrom-CryptoSecureString -SecureString $Value) }
            return
        }
        if ($Value -is [string]) {
            Register-StdLogMask -Value $Value
            return
        }
        if ($Value -is [System.Collections.IDictionary]) {
            foreach ($key in @($Value.Keys)) { Register-EngineSecretMasks -Value $Value[$key] -Depth ($Depth + 1) }
            return
        }
        if ($Value -is [System.Collections.IEnumerable]) {
            foreach ($item in $Value) { Register-EngineSecretMasks -Value $item -Depth ($Depth + 1) }
        }
    } catch {
        # Uszkodzona wartość (np. zwolniony SecureString) nie może zatrzymać silnika.
    }
}

function Send-EngineEvent {
    # Zdarzenie do kolejki okna i do events.jsonl. Pola tekstowe muszą być już zamaskowane.
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [AllowNull()][System.Collections.IDictionary]$Fields,
        [AllowNull()]$Time
    )
    $run = $script:StdEngineRun
    if ($null -eq $Time) { $Time = Get-Date }
    $data = [ordered]@{ Type = $Type; Time = $Time }
    if ($null -ne $Fields) {
        foreach ($key in @($Fields.Keys)) { $data[[string]$key] = $Fields[$key] }
    }
    $event = New-Object -TypeName PSObject -Property $data
    if ($null -eq $run) { return }
    try { $run.Shared.Events.Enqueue($event) } catch { }
    if ($null -ne $run.EventsWriter) {
        try { $run.EventsWriter.WriteLine((ConvertTo-StdJson -InputObject $event -Compress)) } catch { }
    }
}

function Send-EngineLogRecord {
    # Odbiornik logu silnika: każdy wpis (już zamaskowany przez Write-StdLog) jako zdarzenie Log.
    param([AllowNull()]$Record)
    if ($null -eq $Record) { return }
    $fields = [ordered]@{ Level = [string]$Record.Level; Message = [string]$Record.Message; Source = [string]$Record.Source }
    Send-EngineEvent -Type 'Log' -Fields $fields -Time $Record.Time
}

function Get-EngineResultSnapshot {
    # Wynik dla Get-StdEngineResult (nowy obiekt przy każdej zmianie - okno czyta go bez blokad).
    param([Parameter(Mandatory = $true)]$Run, [Parameter(Mandatory = $true)][string]$Status)
    $steps = New-Object 'System.Collections.Generic.List[object]'
    foreach ($record in $Run.Records) {
        $steps.Add(@{
                Id         = $record.Id
                Title      = $record.Title
                Phase      = $record.Phase
                State      = $record.State
                Detail     = $record.Detail
                StartedAt  = $record.StartedAt
                FinishedAt = $record.FinishedAt
            })
    }
    $plan = $null
    if ($null -ne $Run.Plan) {
        $planList = New-Object 'System.Collections.Generic.List[object]'
        foreach ($entry in $Run.Plan) { $planList.Add($entry.Clone()) }
        $plan = $planList.ToArray()
    }
    return @{
        Status         = $Status
        Steps          = $steps.ToArray()
        Plan           = $plan
        Counts         = (Get-EngineCounts $Run.Records)
        RebootRequired = [bool]$Run.RebootRequired
        Error          = $Run.Error
    }
}

function Update-EngineSnapshot {
    param([Parameter(Mandatory = $true)]$Run)
    try { $Run.Shared.Result = Get-EngineResultSnapshot -Run $Run -Status ([string]$Run.Shared.Status) } catch { }
}

function Send-EngineStepState {
    param([Parameter(Mandatory = $true)]$Run, [Parameter(Mandatory = $true)]$Record)
    Update-EngineSnapshot -Run $Run
    $fields = [ordered]@{ StepId = $Record.Id; State = $Record.State; Detail = $Record.Detail; Title = $Record.Title }
    Send-EngineEvent -Type 'StepState' -Fields $fields
}

function Get-EngineAnswers {
    # Odpowiedzi technika do state.json - tylko pola z białej listy (bez haseł, nawet gdyby ktoś je tam wpisał).
    param([AllowNull()]$Config)
    $answers = Get-CryptoMember $Config 'Answers'
    $result = [ordered]@{ computerName = $null; apps = @(); tasks = [ordered]@{} }
    if ($null -eq $answers) { return $result }
    $computerName = Get-CryptoMember $answers 'computerName'
    if ($computerName -is [string] -and -not [string]::IsNullOrWhiteSpace($computerName)) { $result['computerName'] = $computerName }
    $apps = Get-CryptoMember $answers 'apps'
    if ($null -ne $apps) {
        $list = New-Object 'System.Collections.Generic.List[string]'
        foreach ($app in @($apps)) { if ($app -is [string] -and $app.Length -gt 0) { $list.Add($app) } }
        $result['apps'] = $list.ToArray()
    }
    $tasks = Get-CryptoMember $answers 'tasks'
    if ($tasks -is [System.Collections.IDictionary]) {
        foreach ($key in @($tasks.Keys)) {
            $value = $tasks[$key]
            if ($value -is [bool]) { $result['tasks'][[string]$key] = [bool]$value }
        }
    }
    return $result
}

function Get-EnginePreviousSteps {
    # Kroki z poprzedniego zapisu stanu tego samego wdrożenia (tylko state / detail / at).
    param([AllowNull()]$PreviousState)
    $result = [ordered]@{}
    $steps = Get-CryptoMember $PreviousState 'steps'
    if (-not ($steps -is [System.Collections.IDictionary])) { return $result }
    foreach ($key in @($steps.Keys)) {
        $entry = $steps[$key]
        if (-not ($entry -is [System.Collections.IDictionary])) { continue }
        $result[[string]$key] = [ordered]@{
            state  = [string](Get-CryptoMember $entry 'state')
            detail = (Protect-EngineDetail (Get-CryptoMember $entry 'detail'))
            at     = (Get-CryptoMember $entry 'at')
        }
    }
    return $result
}

function New-EngineStateDoc {
    param([Parameter(Mandatory = $true)]$Run, [AllowNull()]$PreviousState)
    $config = $Run.Config
    $site = Get-CryptoMember $config 'Site'
    $siteId = Get-CryptoMember $site 'id'
    if ($null -eq $siteId) { $siteId = Get-CryptoMember $config 'SiteId' }
    $startedAt = $null
    if ($null -ne $PreviousState) { $startedAt = Get-CryptoMember $PreviousState 'startedAt' }
    if ([string]::IsNullOrEmpty([string]$startedAt)) { $startedAt = Get-EngineUtcText }
    return [ordered]@{
        deploymentId   = $Run.DeploymentId
        clientId       = (Get-CryptoMember $config 'ClientId')
        siteId         = $siteId
        profile        = (Get-CryptoMember $config 'ProfileName')
        revision       = (Get-CryptoMember $config 'Revision')
        startedAt      = [string]$startedAt
        updatedAt      = (Get-EngineUtcText)
        phase          = 1
        status         = 'Running'
        rebootRequired = $false
        steps          = [ordered]@{}
        answers        = (Get-EngineAnswers $config)
    }
}

function Save-EngineState {
    # Atomowy zapis state.json (tylko tryb Run). Błąd zapisu nie przerywa wdrożenia.
    param([Parameter(Mandatory = $true)]$Run, [string]$Status)
    if ($Run.Mode -ne 'Run' -or $null -eq $Run.StateDoc) { return }
    if ([string]::IsNullOrEmpty($Status)) { $Status = [string]$Run.Shared.Status }
    $doc = $Run.StateDoc
    $doc['updatedAt'] = Get-EngineUtcText
    $doc['phase'] = [int]$Run.Phase
    $doc['status'] = $Status
    $doc['rebootRequired'] = [bool]$Run.RebootRequired
    $steps = [ordered]@{}
    foreach ($key in @($Run.PreviousSteps.Keys)) { $steps[[string]$key] = $Run.PreviousSteps[$key] }
    foreach ($record in $Run.Records) {
        $at = $record.At
        if ($null -ne $record.FinishedAt) { $at = Get-EngineUtcText $record.FinishedAt }
        elseif ($null -ne $record.StartedAt) { $at = Get-EngineUtcText $record.StartedAt }
        $steps[[string]$record.Id] = [ordered]@{ state = $record.State; detail = $record.Detail; at = $at }
    }
    $doc['steps'] = $steps
    try {
        Write-StdJsonFile -Path (Get-StdPath -Kind State) -InputObject $doc
    } catch {
        Write-StdLog -Level Warn -Source 'Std.Engine' -Message ('Nie udało się zapisać stanu wdrożenia: {0}' -f (Format-CoreErrorText $_))
    }
}

function Open-EngineLogs {
    # Logs\<deploymentId>\deploy.log (odbiornik plikowy) i events.jsonl (każde zdarzenie w osobnej linii)
    # oraz odbiornik, który przekazuje każdy wpis logu do kolejki okna.
    param([Parameter(Mandatory = $true)]$Run)
    # Blok skryptu napisany w module działa w zakresie modułu - widzi funkcje prywatne i $script:.
    Register-StdLogSink -Name $script:StdEngineEventsSink -ScriptBlock { Send-EngineLogRecord -Record $args[0] }
    try {
        $dir = Join-Path (Get-StdPath -Kind Logs) $Run.DeploymentId
        if (-not [System.IO.Directory]::Exists($dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
        $Run.LogDir = $dir
        $Run.LogPath = Join-Path $dir 'deploy.log'
        $Run.EventsPath = Join-Path $dir 'events.jsonl'
    } catch {
        Write-StdLog -Level Warn -Source 'Std.Engine' -Message ('Nie udało się utworzyć folderu logu wdrożenia: {0}' -f (Format-CoreErrorText $_))
        return
    }
    try {
        Register-StdFileLogSink -Name $script:StdEngineFileSink -Path $Run.LogPath
    } catch {
        Write-StdLog -Level Warn -Source 'Std.Engine' -Message ('Nie udało się otworzyć pliku deploy.log: {0}' -f (Format-CoreErrorText $_))
    }
    try {
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($Run.EventsPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, $share)
        $writer = New-Object System.IO.StreamWriter($stream, (New-CoreUtf8NoBom))
        $writer.AutoFlush = $true
        $Run.EventsWriter = $writer
    } catch {
        Write-StdLog -Level Warn -Source 'Std.Engine' -Message ('Nie udało się otworzyć pliku events.jsonl: {0}' -f (Format-CoreErrorText $_))
    }
}

function Close-EngineLogs {
    param([Parameter(Mandatory = $true)]$Run)
    foreach ($name in @($script:StdEngineEventsSink, $script:StdEngineFileSink)) {
        try { Unregister-StdLogSink -Name $name } catch { }
    }
    if ($null -ne $Run.EventsWriter) {
        try {
            try { $Run.EventsWriter.Flush() } finally { $Run.EventsWriter.Dispose() }
        } catch { }
        $Run.EventsWriter = $null
    }
}

function Clear-EngineSecrets {
    # Po wdrożeniu w pamięci nie zostaje nic z Secrets (SPEC 6.3).
    param([Parameter(Mandatory = $true)]$Run)
    try { if ($null -ne $Run.Secrets) { $Run.Secrets.Clear() } } catch { }
    try { if ($null -ne $Run.ExtraSecrets) { $Run.ExtraSecrets.Clear() } } catch { }
}

function Invoke-EngineDetect {
    # Detect kroku -> @{ State = Done|Todo|Blocked; Detail }. Wyjątek albo zły wynik = Blocked.
    param([Parameter(Mandatory = $true)]$Step, [Parameter(Mandatory = $true)]$Context)
    if ($null -eq $Step.Detect) { return @{ State = 'Todo'; Detail = '' } }
    try {
        $output = @(& $Step.Detect $Context)
    } catch {
        return @{ State = 'Blocked'; Detail = ('Błąd sprawdzania: {0}' -f (Format-CoreErrorText $_ 1000)) }
    }
    $result = Select-EngineResultObject -Output $output -Key 'State'
    if ($null -eq $result) {
        return @{ State = 'Blocked'; Detail = 'Sprawdzenie kroku nie zwróciło stanu (brak pola State).' }
    }
    $value = [string](Get-CryptoMember $result 'State')
    $state = Get-EngineAllowedValue -Value $value -Allowed $script:StdEngineDetectStates
    if ($null -eq $state) {
        return @{ State = 'Blocked'; Detail = ("Sprawdzenie kroku zwróciło nieznany stan '{0}'." -f (Protect-EngineDetail $value)) }
    }
    return @{ State = $state; Detail = (Protect-EngineDetail (Get-CryptoMember $result 'Detail')) }
}

function Set-EngineRemainingCancelled {
    # Przerwanie: wszystkie kroki, które jeszcze czekają, dostają Cancelled.
    param([Parameter(Mandatory = $true)]$Run)
    $count = 0
    foreach ($record in $Run.Records) {
        if ($record.State -ne 'Pending') { continue }
        $record.State = 'Cancelled'
        $record.Detail = 'Przerwano przed rozpoczęciem kroku.'
        $count++
        Send-EngineStepState -Run $Run -Record $record
    }
    Write-StdLog -Level Warn -Source 'Std.Engine' -Message ('Przerwano wdrożenie na żądanie. Niewykonane kroki: {0}.' -f $count)
}

function Wait-EnginePause {
    # Pauza między krokami: czekamy na wznowienie albo przerwanie (Start-Sleep jest tu dozwolony -
    # to wątek silnika, nie okna).
    param([Parameter(Mandatory = $true)]$Run)
    $shared = $Run.Shared
    $shared.Status = 'Paused'
    Update-EngineSnapshot -Run $Run
    Send-EngineEvent -Type 'Paused'
    Write-StdLog -Source 'Std.Engine' -Message 'Pauza: wdrożenie czeka na wznowienie.'
    Save-EngineState -Run $Run
    while ($shared.PauseRequested -eq $true -and $shared.CancelRequested -ne $true) {
        Start-Sleep -Milliseconds 200
    }
    $shared.Status = 'Running'
    Update-EngineSnapshot -Run $Run
    if ($shared.CancelRequested -ne $true) {
        Send-EngineEvent -Type 'Resumed'
        Write-StdLog -Source 'Std.Engine' -Message 'Wznowiono wdrożenie.'
    }
    Save-EngineState -Run $Run
}

function Invoke-EngineStep {
    # Jeden krok w trybie Run: Detect (Done -> Skipped, Blocked -> Error), potem Run.
    param(
        [Parameter(Mandatory = $true)]$Run,
        [Parameter(Mandatory = $true)]$Record,
        [AllowNull()]$Step,
        [int]$Index,
        [int]$Total
    )
    $shared = $Run.Shared
    $context = $Run.Context
    $Run.Phase = [int]$Record.Phase
    $Record.State = 'Running'
    $Record.Detail = ''
    $Record.StartedAt = Get-Date
    Send-EngineStepState -Run $Run -Record $Record
    Send-EngineEvent -Type 'Progress' -Fields ([ordered]@{ Percent = (Get-EnginePercent ($Index - 1) $Total); Text = ('Krok {0} z {1}: {2}' -f $Index, $Total, $Record.Title) })
    Write-StdLog -Source 'Std.Engine' -Message ('Krok {0} z {1}: {2} ({3}).' -f $Index, $Total, $Record.Title, $Record.Id)
    Save-EngineState -Run $Run

    $state = 'Error'
    $detail = ''
    $reboot = $false
    if ($null -eq $Step) {
        $detail = 'Krok nie jest zarejestrowany.'
    } else {
        $context.Step = Get-EngineStepInfo -Step $Step
        try {
            $detect = Invoke-EngineDetect -Step $Step -Context $context
            switch ($detect.State) {
                'Done' {
                    $state = 'Skipped'
                    $detail = 'Już zrobione.'
                    if (-not [string]::IsNullOrEmpty($detect.Detail)) { $detail = 'Już zrobione: {0}' -f $detect.Detail }
                }
                'Blocked' {
                    $state = 'Error'
                    $detail = 'Nie można wykonać kroku.'
                    if (-not [string]::IsNullOrEmpty($detect.Detail)) { $detail = 'Nie można wykonać kroku: {0}' -f $detect.Detail }
                }
                default {
                    try {
                        $output = @(& $Step.Run $context)
                        $result = Select-EngineResultObject -Output $output -Key 'Result'
                        if ($null -eq $result) {
                            $state = 'Error'
                            $detail = 'Krok nie zwrócił wyniku (brak pola Result).'
                        } else {
                            $value = [string](Get-CryptoMember $result 'Result')
                            $normalized = Get-EngineAllowedValue -Value $value -Allowed $script:StdEngineRunResults
                            if ($null -eq $normalized) {
                                $state = 'Error'
                                $detail = "Krok zwrócił nieznany wynik '{0}'." -f (Protect-EngineDetail $value)
                            } else {
                                $state = $normalized
                                $detail = [string](Get-CryptoMember $result 'Detail')
                            }
                            $rebootValue = Get-CryptoMember $result 'Reboot'
                            $reboot = ($rebootValue -is [bool] -and $rebootValue)
                        }
                    } catch {
                        $state = 'Error'
                        $detail = Format-CoreErrorText $_ 1000
                    }
                }
            }
        } finally {
            $context.Step = $null
        }
    }
    # Krok, który nie skończył się dobrze, gdy przyszło żądanie przerwania, liczymy jako przerwany.
    if ($state -ne 'Ok' -and $state -ne 'Skipped' -and $shared.CancelRequested -eq $true) {
        $state = 'Cancelled'
        if ([string]::IsNullOrEmpty($detail)) { $detail = 'Przerwano.' }
    }
    $Record.State = $state
    $Record.Detail = Protect-EngineDetail $detail
    $Record.FinishedAt = Get-Date
    $Record.RebootRequested = [bool]$reboot
    Send-EngineStepState -Run $Run -Record $Record

    $level = 'Info'
    if ($state -eq 'Error') { $level = 'Error' } elseif ($state -eq 'Cancelled') { $level = 'Warn' }
    $message = 'Wynik kroku {0}: {1}' -f $Record.Title, $state
    if (-not [string]::IsNullOrEmpty($Record.Detail)) { $message = '{0} - {1}' -f $message, $Record.Detail }
    if ($reboot) { $message = '{0} (wymaga restartu)' -f $message }
    Write-StdLog -Level $level -Source $Record.Id -Message $message
}

function Invoke-EnginePlanMode {
    # „Sprawdź”: Enabled + Detect dla każdego zarejestrowanego kroku, zdarzenie Plan. Nic nie jest zmieniane.
    param([Parameter(Mandatory = $true)]$Run, [AllowNull()][string[]]$OnlySteps)
    $all = Get-EngineSortedSteps
    Write-EngineUnknownOnlySteps -Steps $all -OnlySteps $OnlySteps
    $entries = New-Object 'System.Collections.Generic.List[object]'
    $total = $all.Count
    $index = 0
    $stateNames = @{ Done = 'już zrobione'; Todo = 'do zrobienia'; Blocked = 'problem'; Disabled = 'wyłączony' }
    foreach ($step in $all) {
        $index++
        if ($Run.Shared.CancelRequested -eq $true) {
            $Run.Cancelled = $true
            Write-StdLog -Level Warn -Source 'Std.Engine' -Message 'Przerwano sprawdzanie na żądanie.'
            break
        }
        $enabled = Get-EngineStepEnabled -Step $step -Config $Run.Config -OnlySteps $OnlySteps
        if (-not $enabled.Enabled) {
            $entries.Add(@{ Id = $step.Id; Title = $step.Title; Phase = $step.Phase; Enabled = $false; State = 'Disabled'; Detail = $enabled.Reason })
            continue
        }
        Send-EngineEvent -Type 'Progress' -Fields ([ordered]@{ Percent = (Get-EnginePercent ($index - 1) $total); Text = ('Sprawdzanie: {0}' -f $step.Title) })
        $Run.Context.Step = Get-EngineStepInfo -Step $step
        try {
            $detect = Invoke-EngineDetect -Step $step -Context $Run.Context
        } finally {
            $Run.Context.Step = $null
        }
        $entries.Add(@{ Id = $step.Id; Title = $step.Title; Phase = $step.Phase; Enabled = $true; State = $detect.State; Detail = $detect.Detail })
        $message = 'Sprawdzenie: {0} - {1}' -f $step.Title, $stateNames[$detect.State]
        if (-not [string]::IsNullOrEmpty($detect.Detail)) { $message = '{0}: {1}' -f $message, $detect.Detail }
        $level = 'Info'
        if ($detect.State -eq 'Blocked') { $level = 'Warn' }
        Write-StdLog -Level $level -Source $step.Id -Message $message
        $Run.Plan = $entries.ToArray()
        Update-EngineSnapshot -Run $Run
    }
    $Run.Plan = $entries.ToArray()
    Update-EngineSnapshot -Run $Run
    if (-not $Run.Cancelled) {
        Send-EngineEvent -Type 'Plan' -Fields ([ordered]@{ Steps = $Run.Plan })
    }
}

function Enable-EngineKeepAwake {
    # Zadanie SuspendSleep włączone: wstrzymanie usypiania w wątku silnika na starcie każdego wdrożenia, także
    # przy „Ponów nieudane” (-OnlySteps) i -Resume, w których krok Power.SuspendSleep nie jest uruchamiany
    # ponownie (SetThreadExecutionState działa tylko w wątku, który ją wywołał). Przywraca Complete-EngineRun.
    param([Parameter(Mandatory = $true)]$Run)
    $tasks = Get-CryptoMember $Run.Config 'Tasks'
    $value = Get-CryptoMember $tasks 'SuspendSleep'
    if (-not ($value -is [bool] -and $value)) { return }
    try {
        if (Set-StdKeepAwake -Enabled $true) {
            Write-StdLog -Source 'Std.Engine' -Message 'Wstrzymano usypianie komputera do końca wdrożenia (zadanie SuspendSleep).'
        }
    } catch { }
}

function Invoke-EngineRunMode {
    # „Wdróż”: włączone kroki w kolejności faza / Order; pauza i przerwanie przed każdym krokiem,
    # -Critical, restart po fazie 1, -Resume.
    param([Parameter(Mandatory = $true)]$Run, [AllowNull()][string[]]$OnlySteps, [bool]$Resume)
    $shared = $Run.Shared
    $all = Get-EngineSortedSteps
    Write-EngineUnknownOnlySteps -Steps $all -OnlySteps $OnlySteps
    foreach ($step in $all) {
        $enabled = Get-EngineStepEnabled -Step $step -Config $Run.Config -OnlySteps $OnlySteps
        if (-not $enabled.Enabled) { continue }
        $record = @{
            Id              = $step.Id
            Title           = $step.Title
            Phase           = [int]$step.Phase
            Order           = [int]$step.Order
            Critical        = [bool]$step.Critical
            State           = 'Pending'
            Detail          = ''
            StartedAt       = $null
            FinishedAt      = $null
            At              = $null
            Resumed         = $false
            RebootRequested = $false
            CriticalBlocked = $false
        }
        if ($Resume) {
            $previous = Get-CryptoMember $Run.PreviousSteps ([string]$step.Id)
            $previousState = Get-EngineAllowedValue -Value ([string](Get-CryptoMember $previous 'state')) -Allowed @('Ok', 'Skipped')
            if ($null -ne $previousState) {
                $record.State = $previousState
                $record.Detail = Protect-EngineDetail (Get-CryptoMember $previous 'detail')
                $record.At = Get-CryptoMember $previous 'at'
                $record.Resumed = $true
            }
        }
        $Run.Records.Add($record)
    }
    $total = $Run.Records.Count
    Write-StdLog -Source 'Std.Engine' -Message ('Kroki do wykonania: {0}.' -f $total)
    if ($Resume) {
        Write-StdLog -Source 'Std.Engine' -Message 'Wznowienie: kroki zakończone wcześniej (OK / Pominięto) nie będą uruchamiane ponownie.'
    }
    Enable-EngineKeepAwake -Run $Run
    Save-EngineState -Run $Run
    foreach ($record in $Run.Records) { Send-EngineStepState -Run $Run -Record $record }
    Send-EngineEvent -Type 'Progress' -Fields ([ordered]@{ Percent = 0; Text = 'Start wdrożenia' })

    $criticalByPhase = @{}
    $rebootAfterPhase1 = $false
    $index = 0
    foreach ($record in $Run.Records) {
        $index++
        if ($record.Resumed) {
            Write-StdLog -Source $record.Id -Message ('Krok {0}: {1} z poprzedniego uruchomienia - nie uruchamiam go ponownie.' -f $record.Title, $record.State)
            continue
        }
        if ($shared.CancelRequested -eq $true) {
            Set-EngineRemainingCancelled -Run $Run
            break
        }
        if ($rebootAfterPhase1 -and $record.Phase -gt 1) {
            # Faza 2 ruszy po restarcie komputera (etap 4) - krok zostaje w stanie Pending.
            $record.Detail = 'Czeka na restart komputera.'
            Send-EngineStepState -Run $Run -Record $record
            continue
        }
        if ($criticalByPhase.ContainsKey($record.Phase)) {
            # Krok się nie wykonał: Cancelled (nie Skipped), żeby „Ponów nieudane” (Error / Cancelled) i -Resume
            # (pomija tylko Ok / Skipped) uruchomiły go ponownie. Stanu wdrożenia nie zmienia na Cancelled.
            $record.State = 'Cancelled'
            $record.CriticalBlocked = $true
            $record.Detail = 'Pominięto, bo krok krytyczny „{0}” zakończył się błędem.' -f $criticalByPhase[$record.Phase]
            $record.FinishedAt = Get-Date
            Send-EngineStepState -Run $Run -Record $record
            Write-StdLog -Level Warn -Source $record.Id -Message ('Wynik kroku {0}: Cancelled - {1}' -f $record.Title, $record.Detail)
            Save-EngineState -Run $Run
            continue
        }
        if ($shared.PauseRequested -eq $true) {
            Wait-EnginePause -Run $Run
            if ($shared.CancelRequested -eq $true) {
                Set-EngineRemainingCancelled -Run $Run
                break
            }
        }
        $step = $null
        if ($script:StdSteps.Contains($record.Id)) { $step = $script:StdSteps[$record.Id] }
        Invoke-EngineStep -Run $Run -Record $record -Step $step -Index $index -Total $total
        if ($record.State -eq 'Error' -and $record.Critical) {
            $criticalByPhase[$record.Phase] = $record.Title
            Write-StdLog -Level Error -Source 'Std.Engine' -Message ('Krok krytyczny {0} zakończył się błędem - pozostałe kroki fazy {1} zostaną pominięte.' -f $record.Title, $record.Phase)
        }
        if ($record.RebootRequested) {
            $Run.RebootRequired = $true
            if ($record.Phase -eq 1 -and -not $rebootAfterPhase1) {
                $rebootAfterPhase1 = $true
                Write-StdLog -Source 'Std.Engine' -Message ('Krok {0} wymaga restartu: silnik dokończy fazę 1, faza 2 ruszy po restarcie.' -f $record.Title)
            }
        }
        Save-EngineState -Run $Run
        Send-EngineEvent -Type 'Progress' -Fields ([ordered]@{ Percent = (Get-EnginePercent $index $total); Text = ('Krok {0} z {1}: {2}' -f $index, $total, $record.Title) })
    }
    if ($rebootAfterPhase1) { $Run.Phase = 2 }
}

function Complete-EngineRun {
    # Koniec pracy silnika: stan końcowy, przywrócenie usypiania, czyszczenie Secrets (tryb Run),
    # wynik, zdarzenie Finished i zamknięcie plików logu.
    param([Parameter(Mandatory = $true)]$Run)
    $shared = $Run.Shared
    $status = 'Finished'
    if ($Run.Failed) {
        $status = 'Failed'
    } elseif ($Run.Cancelled) {
        $status = 'Cancelled'
    } else {
        foreach ($record in $Run.Records) {
            # Kroki niewykonane z powodu błędu kroku krytycznego nie oznaczają przerwania wdrożenia.
            if ($record.State -eq 'Cancelled' -and -not $record.CriticalBlocked) { $status = 'Cancelled'; break }
        }
    }
    foreach ($record in $Run.Records) {
        if ($record.State -ne 'Running') { continue }
        $record.State = 'Error'
        $record.Detail = Protect-EngineDetail ('Krok przerwany przez błąd silnika: {0}' -f $Run.Error)
        $record.FinishedAt = Get-Date
        try { Send-EngineStepState -Run $Run -Record $record } catch { }
    }
    if ($Run.Mode -eq 'Run') {
        try {
            if ($script:StdKeepAwakeActive) {
                if (Set-StdKeepAwake -Enabled $false) {
                    Write-StdLog -Source 'Std.Engine' -Message 'Przywrócono zwykłe usypianie komputera.'
                }
            }
        } catch { }
        Clear-EngineSecrets -Run $Run
    }
    try { Save-EngineState -Run $Run -Status $status } catch { }

    $counts = Get-EngineCounts $Run.Records
    try {
        if ($Run.Mode -eq 'Plan') {
            $planCounts = @{ Done = 0; Todo = 0; Blocked = 0; Disabled = 0 }
            foreach ($entry in @($Run.Plan)) {
                if ($null -ne $entry -and $planCounts.ContainsKey([string]$entry.State)) { $planCounts[[string]$entry.State] = $planCounts[[string]$entry.State] + 1 }
            }
            Write-StdLog -Source 'Std.Engine' -Message ('Koniec sprawdzania: do zrobienia {0}, już zrobione {1}, problem {2}, wyłączone {3}.' -f $planCounts.Todo, $planCounts.Done, $planCounts.Blocked, $planCounts.Disabled)
        } else {
            $statusNames = @{ Finished = 'zakończono'; Cancelled = 'przerwano'; Failed = 'błąd silnika' }
            $summary = 'Koniec wdrożenia {0}: {1}. OK: {2}, pominięto: {3}, błąd: {4}, przerwano: {5}.' -f $Run.DeploymentId, $statusNames[$status], $counts.Ok, $counts.Skipped, $counts.Error, $counts.Cancelled
            if ($Run.RebootRequired) { $summary = '{0} Wymagany restart komputera.' -f $summary }
            $level = 'Info'
            if ($status -eq 'Failed' -or $counts.Error -gt 0) { $level = 'Error' } elseif ($status -eq 'Cancelled') { $level = 'Warn' }
            Write-StdLog -Level $level -Source 'Std.Engine' -Message $summary
        }
    } catch { }
    if ($status -ne 'Failed') {
        try { Send-EngineEvent -Type 'Progress' -Fields ([ordered]@{ Percent = 100; Text = 'Zakończono' }) } catch { }
    }
    # Najpierw wynik, potem stan i zdarzenie Finished - okno po Finished od razu czyta Get-StdEngineResult.
    try { $shared.Result = Get-EngineResultSnapshot -Run $Run -Status $status } catch { }
    $shared.Status = $status
    try {
        Send-EngineEvent -Type 'Finished' -Fields ([ordered]@{ Status = $status; Counts = $counts; RebootRequired = [bool]$Run.RebootRequired; Error = $Run.Error })
    } catch { }
    Close-EngineLogs -Run $Run
    $script:StdEngineRun = $null
}

function Initialize-EnginePowerType {
    # Typ z SetThreadExecutionState (tylko Windows). Jeden na proces, więc sprawdzamy, czy już jest.
    if (-not (Test-CoreIsWindows)) { return $false }
    if ($null -ne ($script:StdEnginePowerTypeName -as [type])) { return $true }
    try {
        Add-Type -TypeDefinition $script:StdEnginePowerSource -Language CSharp -ErrorAction Stop
    } catch {
        if ($null -eq ($script:StdEnginePowerTypeName -as [type])) {
            Write-StdLog -Level Warn -Source 'Set-StdKeepAwake' -Message ('Nie udało się przygotować wstrzymania usypiania: {0}' -f (Format-CoreErrorText $_))
            return $false
        }
    }
    return ($null -ne ($script:StdEnginePowerTypeName -as [type]))
}

function Stop-EngineProcessTree {
    # Zabija proces razem z procesami potomnymi: Windows - taskkill /T /F, inne systemy - Kill(true) z .NET.
    param([Parameter(Mandatory = $true)][System.Diagnostics.Process]$Process)
    $processId = $Process.Id
    if (Test-CoreIsWindows) {
        $killer = $null
        try {
            $taskkill = 'taskkill.exe'
            if (-not [string]::IsNullOrEmpty($env:SystemRoot)) {
                $candidate = Join-Path $env:SystemRoot 'System32\taskkill.exe'
                if ([System.IO.File]::Exists($candidate)) { $taskkill = $candidate }
            }
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $taskkill
            $psi.Arguments = '/T /F /PID {0}' -f $processId
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $killer = [System.Diagnostics.Process]::Start($psi)
            [void]$killer.StandardOutput.ReadToEnd()
            [void]$killer.StandardError.ReadToEnd()
            [void]$killer.WaitForExit(15000)
        } catch {
        } finally {
            if ($null -ne $killer) { $killer.Dispose() }
        }
    } else {
        $treeKill = $Process.GetType().GetMethod('Kill', [type[]]@([bool]))
        if ($null -ne $treeKill) {
            try { $Process.Kill($true) } catch { }
        }
    }
    try { if (-not $Process.HasExited) { $Process.Kill() } } catch { }
    try { [void]$Process.WaitForExit(10000) } catch { }
}

function Get-EngineCleanText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ($text.Length -eq 0) { return $null }
    return $text
}

function Get-EngineSerialText {
    # Numer seryjny bez typowych zaślepek producentów płyt („To be filled by O.E.M.” itp.).
    param([AllowNull()]$Value)
    $text = Get-EngineCleanText $Value
    if ($null -eq $text) { return $null }
    $placeholders = @('To be filled by O.E.M.', 'Default string', 'System Serial Number', 'Not Specified', 'Not Applicable', 'None', 'N/A', '0', '00000000', '0123456789', '123456789')
    foreach ($placeholder in $placeholders) {
        if ([string]::Equals($text, $placeholder, [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
    }
    return $text
}

function Get-EngineCimInstances {
    # Wszystkie obiekty klasy CIM (object[], także pusta) albo $null, gdy CIM nie działa
    # (inny system, brak uprawnień, brak przestrzeni nazw).
    param([Parameter(Mandatory = $true)][string]$ClassName, [string]$Namespace = 'root/cimv2', [string]$Filter)
    try {
        $query = @{ ClassName = $ClassName; Namespace = $Namespace; OperationTimeoutSec = 20; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrEmpty($Filter)) { $query.Filter = $Filter }
        $items = @(Get-CimInstance @query)
        return , $items
    } catch {
        return $null
    }
}

function Get-EngineFirstCim {
    param([Parameter(Mandatory = $true)][string]$ClassName, [string]$Namespace = 'root/cimv2')
    $items = Get-EngineCimInstances -ClassName $ClassName -Namespace $Namespace
    if ($null -eq $items -or $items.Count -eq 0) { return $null }
    return $items[0]
}

function Get-EngineDiskType {
    # Rodzaj dysku pod woluminem: Win32_LogicalDisk -> Win32_DiskPartition -> Win32_DiskDrive -> MSFT_PhysicalDisk.
    param([Parameter(Mandatory = $true)]$LogicalDisk, [AllowNull()][object[]]$PhysicalDisks)
    if ($null -eq $PhysicalDisks -or $PhysicalDisks.Count -eq 0) { return $null }
    foreach ($partition in @(Get-CimAssociatedInstance -InputObject $LogicalDisk -ResultClassName 'Win32_DiskPartition' -ErrorAction Stop)) {
        foreach ($drive in @(Get-CimAssociatedInstance -InputObject $partition -ResultClassName 'Win32_DiskDrive' -ErrorAction Stop)) {
            $index = [string]$drive.Index
            foreach ($physical in $PhysicalDisks) {
                if ([string]$physical.DeviceId -ne $index) { continue }
                $type = $null
                switch ([int]$physical.MediaType) {
                    3 { $type = 'HDD' }
                    4 { $type = 'SSD' }
                    5 { $type = 'SCM' }
                }
                if ([int]$physical.BusType -eq 17) {
                    if ($null -eq $type) { $type = 'NVMe' } else { $type = '{0} NVMe' -f $type }
                }
                return $type
            }
        }
    }
    return $null
}

function Get-EngineDiskList {
    $logical = Get-EngineCimInstances -ClassName 'Win32_LogicalDisk' -Filter 'DriveType = 3'
    if ($null -eq $logical) { return $null }
    $physical = Get-EngineCimInstances -ClassName 'MSFT_PhysicalDisk' -Namespace 'root/Microsoft/Windows/Storage'
    $list = New-Object 'System.Collections.Generic.List[object]'
    foreach ($disk in $logical) {
        $entry = @{ Name = $null; SizeGB = $null; FreeGB = $null; Type = $null }
        try { $entry.Name = Get-EngineCleanText $disk.DeviceID } catch { }
        try { if ($null -ne $disk.Size) { $entry.SizeGB = [System.Math]::Round([double]$disk.Size / 1GB, 1) } } catch { }
        try { if ($null -ne $disk.FreeSpace) { $entry.FreeGB = [System.Math]::Round([double]$disk.FreeSpace / 1GB, 1) } } catch { }
        try { $entry.Type = Get-EngineDiskType -LogicalDisk $disk -PhysicalDisks $physical } catch { }
        $list.Add($entry)
    }
    return , $list.ToArray()
}

function Get-EngineTpmInfo {
    $items = Get-EngineCimInstances -ClassName 'Win32_Tpm' -Namespace 'root/cimv2/Security/MicrosoftTpm'
    if ($null -eq $items) { return $null }
    $tpm = @{ Present = $false; Ready = $null; Version = $null }
    if ($items.Count -eq 0) {
        $tpm.Ready = $false
        return $tpm
    }
    $device = $items[0]
    $tpm.Present = $true
    try {
        $spec = Get-EngineCleanText $device.SpecVersion
        if ($null -ne $spec) { $tpm.Version = Get-EngineCleanText (($spec -split ',')[0]) }
    } catch { }
    try {
        $ready = Invoke-CimMethod -InputObject $device -MethodName 'IsReady' -ErrorAction Stop
        if ($null -ne $ready -and $null -ne $ready.IsReady) { $tpm.Ready = [bool]$ready.IsReady }
    } catch { }
    if ($null -eq $tpm.Ready) {
        try {
            if ($null -ne $device.IsEnabled_InitialValue -and $null -ne $device.IsActivated_InitialValue) {
                $tpm.Ready = ([bool]$device.IsEnabled_InitialValue -and [bool]$device.IsActivated_InitialValue)
            }
        } catch { }
    }
    return $tpm
}

function Get-EngineNetworkList {
    $configs = Get-EngineCimInstances -ClassName 'Win32_NetworkAdapterConfiguration' -Filter 'IPEnabled = TRUE'
    if ($null -eq $configs) { return $null }
    $adapters = Get-EngineCimInstances -ClassName 'Win32_NetworkAdapter'
    $list = New-Object 'System.Collections.Generic.List[object]'
    foreach ($config in $configs) {
        $entry = @{ Name = $null; Ipv4 = $null; Mac = $null }
        try {
            $name = $null
            if ($null -ne $adapters) {
                foreach ($adapter in $adapters) {
                    if ($adapter.Index -eq $config.Index) { $name = Get-EngineCleanText $adapter.NetConnectionID; break }
                }
            }
            if ($null -eq $name) { $name = Get-EngineCleanText $config.Description }
            $entry.Name = $name
        } catch { }
        try {
            foreach ($address in @($config.IPAddress)) {
                $text = [string]$address
                if ($text -match '^\d{1,3}(\.\d{1,3}){3}$' -and -not $text.StartsWith('169.254.')) { $entry.Ipv4 = $text; break }
            }
        } catch { }
        try { $entry.Mac = Get-EngineCleanText $config.MACAddress } catch { }
        $list.Add($entry)
    }
    return , $list.ToArray()
}

# ---------------------------------------------------------------------------------------------
# Kroki (SPEC 6.1)
# ---------------------------------------------------------------------------------------------

function Register-StdStep {
    <#
    .SYNOPSIS
    Rejestruje krok wdrożenia (wołane z plików Steps\*.ps1). Id unikalne (Grupa.Nazwa), Phase 1 (przed
    restartem) lub 2 (po), Order w fazie, TaskId z listy zadań (krok włączony, gdy Config.Tasks[TaskId] = $true;
    bez TaskId zawsze włączony). Detect niczego nie zmienia i zwraca @{ State = 'Done'|'Todo'|'Blocked'; Detail },
    Run zwraca @{ Result = 'Ok'|'Error'|'Skipped'; Detail; Reboot } albo rzuca wyjątek (= Error).
    -Critical: błąd kroku pomija resztę fazy. Złe dane -> [STD-ENGINE-STEP].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [int]$Phase,
        [Parameter(Mandatory = $true)]
        [int]$Order,
        [string]$TaskId,
        [double]$TimeoutMin = 30,
        [switch]$Critical,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Detect,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Run
    )
    if ([string]::IsNullOrWhiteSpace($Id) -or $Id -notmatch $script:StdEngineStepIdPattern) {
        throw ("[STD-ENGINE-STEP] Nieprawidłowy identyfikator kroku '{0}' - wymagany zapis Grupa.Nazwa (litery i cyfry)." -f $Id)
    }
    if ([string]::IsNullOrWhiteSpace($Title)) {
        throw ('[STD-ENGINE-STEP] Krok {0} nie ma tytułu.' -f $Id)
    }
    if ($Phase -ne 1 -and $Phase -ne 2) {
        throw ('[STD-ENGINE-STEP] Krok {0}: faza musi być 1 (przed restartem) albo 2 (po restarcie), jest {1}.' -f $Id, $Phase)
    }
    if ($TimeoutMin -le 0 -or $TimeoutMin -gt 1440) {
        throw ('[STD-ENGINE-STEP] Krok {0}: TimeoutMin musi być w zakresie 0-1440 minut.' -f $Id)
    }
    $canonicalTaskId = $null
    if (-not [string]::IsNullOrWhiteSpace($TaskId)) {
        $known = $script:StdConfigTaskIds
        if ($null -eq $known -or @($known).Count -eq 0) {
            throw '[STD-ENGINE-STEP] Brak listy identyfikatorów zadań (Std.Config.ps1 nie jest wczytany).'
        }
        $canonicalTaskId = Get-EngineAllowedValue -Value $TaskId -Allowed $known
        if ($null -eq $canonicalTaskId) {
            throw ("[STD-ENGINE-STEP] Krok {0}: nieznany identyfikator zadania '{1}'. Dozwolone: {2}." -f $Id, $TaskId, ($known -join ', '))
        }
    }
    if ($script:StdSteps.Contains($Id)) {
        throw ('[STD-ENGINE-STEP] Krok {0} jest już zarejestrowany.' -f $Id)
    }
    $script:StdStepSequence = [int]$script:StdStepSequence + 1
    $script:StdSteps[$Id] = [pscustomobject]@{
        Id         = $Id
        Title      = $Title
        Phase      = $Phase
        Order      = $Order
        TaskId     = $canonicalTaskId
        TimeoutMin = $TimeoutMin
        Critical   = [bool]$Critical
        Detect     = $Detect
        Run        = $Run
        Source     = $MyInvocation.ScriptName
        Seq        = $script:StdStepSequence
    }
}

function Get-StdStep {
    <#
    .SYNOPSIS
    Zarejestrowane kroki (faza, Order, kolejność rejestracji). -Id: jeden krok albo nic.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$Id
    )
    if (-not [string]::IsNullOrEmpty($Id)) {
        if ($script:StdSteps.Contains($Id)) { return $script:StdSteps[$Id] }
        return
    }
    foreach ($step in (Get-EngineSortedSteps)) { $step }
}

function Clear-StdSteps {
    <#
    .SYNOPSIS
    Usuwa wszystkie zarejestrowane kroki (testy).
    #>
    [CmdletBinding()]
    param()
    $script:StdSteps.Clear()
    $script:StdStepSequence = 0
}

function Test-StdEngineCancel {
    <#
    .SYNOPSIS
    $true, gdy okno zażądało przerwania wdrożenia (Shared.CancelRequested). Kroki sprawdzają to w długich pętlach.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [object]$Context
    )
    $shared = Get-CryptoMember $Context 'Shared'
    return (Get-EngineFlag $shared 'CancelRequested')
}

# ---------------------------------------------------------------------------------------------
# Silnik (SPEC 6.2, 6.3)
# ---------------------------------------------------------------------------------------------

function Invoke-StdEngineLoop {
    <#
    .SYNOPSIS
    Pętla silnika - uruchamia ją Start-StdEngine w osobnym runspace (po wczytaniu modułu i kroków).
    Tryb Plan: Enabled + Detect każdego kroku i zdarzenie Plan. Tryb Run: kroki w kolejności faza / Order,
    state.json po każdym kroku, deploy.log i events.jsonl w Logs\<deploymentId>\. Nie rzuca wyjątków:
    błąd silnika kończy pracę ze stanem Failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Plan', 'Run')]
        [string]$Mode,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Config,
        [Parameter(Mandatory = $true)]
        [hashtable]$Shared,
        [AllowNull()]
        [hashtable]$Secrets,
        [Parameter(Mandatory = $true)]
        [string]$DeploymentId,
        [AllowNull()]
        [AllowEmptyCollection()]
        [string[]]$OnlySteps,
        [AllowNull()]
        [System.Collections.IDictionary]$PreviousState,
        [switch]$Resume,
        [AllowNull()]
        [System.Collections.IDictionary]$ExtraSecrets,
        [AllowNull()]
        [AllowEmptyCollection()]
        [string[]]$StepLoadErrors
    )
    if ($Mode -eq 'Plan') { $Mode = 'Plan' } else { $Mode = 'Run' }
    if ($null -eq $Secrets) { $Secrets = [hashtable]::Synchronized(@{}) }
    $run = @{
        Mode           = $Mode
        Config         = $Config
        Shared         = $Shared
        Secrets        = $Secrets
        ExtraSecrets   = $ExtraSecrets
        DeploymentId   = $DeploymentId
        LogDir         = $null
        LogPath        = $null
        EventsPath     = $null
        EventsWriter   = $null
        Records        = (New-Object 'System.Collections.Generic.List[object]')
        Plan           = $null
        RebootRequired = $false
        Failed         = $false
        Cancelled      = $false
        Error          = $null
        Phase          = 1
        PreviousSteps  = (Get-EnginePreviousSteps $PreviousState)
        StateDoc       = $null
        Context        = $null
    }
    $script:StdEngineRun = $run
    $Shared.Status = 'Running'
    try {
        # Maski przed pierwszym wpisem do logu: ta kopia modułu nie zna masek okna.
        Register-EngineSecretMasks -Value $Secrets
        Register-EngineSecretMasks -Value $ExtraSecrets
        Open-EngineLogs -Run $run
        $run.StateDoc = New-EngineStateDoc -Run $run -PreviousState $PreviousState
        $run.Context = @{
            Config       = $Config
            State        = $run.StateDoc
            Secrets      = $Secrets
            Shared       = $Shared
            DeploymentId = $DeploymentId
            LogPath      = $run.LogPath
            Step         = $null
        }
        $modeName = 'wdrożenie'
        if ($Mode -eq 'Plan') { $modeName = 'sprawdzenie' }
        Write-StdLog -Source 'Std.Engine' -Message ('STD {0}: {1} {2} - klient {3}, lokalizacja {4}, profil {5}, rewizja {6}.' -f (Get-StdVersion), $modeName, $DeploymentId, $run.StateDoc['clientId'], $run.StateDoc['siteId'], $run.StateDoc['profile'], $run.StateDoc['revision'])
        if ($null -ne $StepLoadErrors -and $StepLoadErrors.Count -gt 0) {
            foreach ($loadError in $StepLoadErrors) {
                Write-StdLog -Level Error -Source 'Std.Engine' -Message ('Błąd w pliku kroku {0}' -f $loadError)
            }
            throw ('[STD-ENGINE-STEP] Nie udało się wczytać kroków: {0}' -f ($StepLoadErrors -join '; '))
        }
        if ($Mode -eq 'Plan') {
            Invoke-EnginePlanMode -Run $run -OnlySteps $OnlySteps
        } else {
            Invoke-EngineRunMode -Run $run -OnlySteps $OnlySteps -Resume ([bool]$Resume)
        }
    } catch {
        $run.Failed = $true
        $run.Error = Format-CoreErrorText $_ 1000
        try { Write-StdLog -Level Error -Source 'Std.Engine' -Message ('Błąd silnika: {0}' -f $run.Error) } catch { }
    } finally {
        Complete-EngineRun -Run $run
    }
}

function Start-StdEngine {
    <#
    .SYNOPSIS
    Uruchamia silnik w osobnym runspace (MTA): import STD.psd1 z folderu tego modułu, dot-source kroków
    z -StepsPath (domyślnie <v4>\Steps) i Invoke-StdEngineLoop asynchronicznie. Zwraca uchwyt
    @{ Id; Mode; Shared; PowerShell; AsyncResult; Runspace }. -Secrets: [hashtable] z SecureString /
    PSCredential (tylko w pamięci; po wdrożeniu silnik go czyści). -Resume: identyfikator i wyniki z state.json,
    kroki OK / Pominięto nie są uruchamiane ponownie. Błędy: [STD-ENGINE-ARGS], [STD-ENGINE-RESUME], [STD-ENGINE-START].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Plan', 'Run')]
        [string]$Mode,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Config,
        [string]$StepsPath,
        [AllowNull()]
        [System.Collections.IDictionary]$Secrets,
        [string]$DeploymentId,
        [AllowNull()]
        [AllowEmptyCollection()]
        [string[]]$OnlySteps,
        [switch]$Resume
    )
    if ($Mode -eq 'Plan') { $Mode = 'Plan' } else { $Mode = 'Run' }
    $root = Get-EngineRoot
    $manifest = Join-Path $root 'STD.psd1'
    if (-not [System.IO.File]::Exists($manifest)) {
        throw ('[STD-ENGINE-START] Brak pliku modułu {0}.' -f $manifest)
    }
    if ([string]::IsNullOrWhiteSpace($StepsPath)) { $StepsPath = Join-Path $root 'Steps' }
    $stepsFull = Resolve-CoreFullPath -Path $StepsPath
    if (-not [System.IO.Directory]::Exists($stepsFull)) {
        throw ('[STD-ENGINE-ARGS] Nie ma folderu kroków {0}.' -f $stepsFull)
    }
    # -Filter '*.ps1' w Windows pasuje też do *.ps1xml (stara reguła 8.3) - dlatego jeszcze sprawdzenie rozszerzenia.
    $stepFiles = @(Get-ChildItem -LiteralPath $stepsFull -Filter '*.ps1' -File -ErrorAction Stop |
            Where-Object { $_.Extension -eq '.ps1' } | Sort-Object -Property Name | ForEach-Object { $_.FullName })

    $previousState = $null
    if ($Mode -eq 'Run') {
        $statePath = Get-StdPath -Kind State
        $saved = $null
        try {
            $saved = Read-StdJsonFile -Path $statePath
        } catch {
            if ($Resume) { throw ('[STD-ENGINE-RESUME] Nie udało się odczytać stanu wdrożenia: {0}' -f (Format-CoreErrorText $_)) }
        }
        $savedId = $null
        if ($saved -is [System.Collections.IDictionary]) { $savedId = [string](Get-CryptoMember $saved 'deploymentId') }
        if ($Resume) {
            if ([string]::IsNullOrWhiteSpace($savedId)) {
                throw ('[STD-ENGINE-RESUME] Nie ma zapisanego stanu wdrożenia do wznowienia ({0}).' -f $statePath)
            }
            if (-not [string]::IsNullOrWhiteSpace($DeploymentId) -and -not [string]::Equals($DeploymentId, $savedId, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw ('[STD-ENGINE-RESUME] Zapisany stan dotyczy wdrożenia {0}, a nie {1}.' -f $savedId, $DeploymentId)
            }
            $DeploymentId = $savedId
        }
        if (-not [string]::IsNullOrWhiteSpace($DeploymentId) -and [string]::Equals($DeploymentId, $savedId, [System.StringComparison]::OrdinalIgnoreCase)) {
            # To samo wdrożenie (wznowienie albo „Ponów nieudane”): wyniki pozostałych kroków zostają w state.json.
            $previousState = $saved
        }
    }
    if ([string]::IsNullOrWhiteSpace($DeploymentId)) {
        $DeploymentId = New-EngineDeploymentId
    } elseif ($DeploymentId -notmatch $script:StdEngineDeploymentIdPattern) {
        throw ("[STD-ENGINE-ARGS] Nieprawidłowy identyfikator wdrożenia '{0}' (litery, cyfry, '.', '_', '-', do 64 znaków)." -f $DeploymentId)
    }

    # Secrets: zsynchronizowany [hashtable]. Opakowanie Synchronized działa na tym samym obiekcie, więc
    # krok, który usuwa hasło ($Context.Secrets.Remove(...)), usuwa je też u wywołującego.
    $secretsSync = $null
    $extraSecrets = $null
    if ($null -eq $Secrets) {
        $secretsSync = [hashtable]::Synchronized(@{})
    } elseif ($Secrets -is [hashtable]) {
        if ($Secrets.IsSynchronized) { $secretsSync = $Secrets } else { $secretsSync = [hashtable]::Synchronized($Secrets) }
    } else {
        $copy = @{}
        foreach ($key in @($Secrets.Keys)) { $copy[$key] = $Secrets[$key] }
        $secretsSync = [hashtable]::Synchronized($copy)
        $extraSecrets = $Secrets
    }

    $onlyList = $null
    if ($PSBoundParameters.ContainsKey('OnlySteps') -and $null -ne $OnlySteps) {
        $selected = New-Object 'System.Collections.Generic.List[string]'
        foreach ($id in $OnlySteps) { if (-not [string]::IsNullOrWhiteSpace($id)) { $selected.Add($id.Trim()) } }
        $onlyList = $selected.ToArray()
    }

    $shared = [hashtable]::Synchronized(@{
            Events          = (New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]')
            PauseRequested  = $false
            CancelRequested = $false
            ForceCancel     = $false
            Status          = 'Running'
            Result          = $null
        })
    $loop = @{
        Mode          = $Mode
        Config        = $Config
        Shared        = $shared
        Secrets       = $secretsSync
        DeploymentId  = $DeploymentId
        OnlySteps     = $onlyList
        PreviousState = $previousState
        Resume        = [bool]$Resume
        ExtraSecrets  = $extraSecrets
    }
    $startArgs = @{ ManifestPath = $manifest; StepFiles = $stepFiles; Shared = $shared; Loop = $loop }

    $runspace = $null
    $powerShell = $null
    try {
        $iss = $null
        try {
            $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
        } catch {
            $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        }
        if (Test-CoreIsWindows) {
            try { $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass } catch { }
        }
        $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
        try { $runspace.ApartmentState = [System.Threading.ApartmentState]::MTA } catch { }
        try { $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread } catch { }
        $runspace.Open()
        $runspace.SessionStateProxy.SetVariable('StdEngineStart', $startArgs)
        $powerShell = [System.Management.Automation.PowerShell]::Create()
        $powerShell.Runspace = $runspace
        [void]$powerShell.AddScript($script:StdEngineBootstrap)
        $async = $powerShell.BeginInvoke()
    } catch {
        $reason = Format-CoreErrorText $_
        if ($null -ne $powerShell) { try { $powerShell.Dispose() } catch { } }
        if ($null -ne $runspace) { try { $runspace.Dispose() } catch { } }
        throw ('[STD-ENGINE-START] Nie udało się uruchomić wątku silnika: {0}' -f $reason)
    }
    return @{
        Id          = $DeploymentId
        Mode        = $Mode
        Shared      = $shared
        PowerShell  = $powerShell
        AsyncResult = $async
        Runspace    = $runspace
    }
}

function Get-EngineHandleShared {
    param([AllowNull()]$Engine)
    $shared = Get-CryptoMember $Engine 'Shared'
    if (-not ($shared -is [hashtable])) {
        throw '[STD-ENGINE-ARGS] Nieprawidłowy uchwyt silnika (oczekiwany wynik Start-StdEngine).'
    }
    return $shared
}

function Receive-StdEngineEvent {
    <#
    .SYNOPSIS
    Zdarzenia silnika z kolejki (najstarsze pierwsze, najwyżej -Max), bez czekania. Wynik do potoku.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [object]$Engine,
        [ValidateRange(1, 1000000)]
        [int]$Max = 500
    )
    $shared = Get-EngineHandleShared $Engine
    $queue = $shared.Events
    $item = $null
    $count = 0
    while ($count -lt $Max -and $queue.TryDequeue([ref]$item)) {
        $count++
        $item
    }
}

function Suspend-StdEngine {
    <#
    .SYNOPSIS
    Pauza po bieżącym kroku (zdarzenie Paused, gdy silnik się zatrzyma).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][object]$Engine)
    $shared = Get-EngineHandleShared $Engine
    $shared.PauseRequested = $true
}

function Resume-StdEngine {
    <#
    .SYNOPSIS
    Wznowienie po pauzie (zdarzenie Resumed).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][object]$Engine)
    $shared = Get-EngineHandleShared $Engine
    $shared.PauseRequested = $false
}

function Stop-StdEngine {
    <#
    .SYNOPSIS
    Przerwanie po bieżącym kroku (pozostałe kroki: Cancelled). -Force: także zabicie bieżącego procesu
    uruchomionego przez Invoke-StdProcess (razem z procesami potomnymi).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][object]$Engine,
        [switch]$Force
    )
    $shared = Get-EngineHandleShared $Engine
    $shared.CancelRequested = $true
    if ($Force) { $shared.ForceCancel = $true }
}

function Complete-EngineInvocation {
    # EndInvoke raz po zakończeniu (zwalnia wynik potoku); błędy potoku zostają w InvocationStateInfo.
    param([Parameter(Mandatory = $true)]$Engine)
    $powerShell = Get-CryptoMember $Engine 'PowerShell'
    $async = Get-CryptoMember $Engine 'AsyncResult'
    if ($null -eq $powerShell -or $null -eq $async -or -not $async.IsCompleted) { return }
    try { [void]$powerShell.EndInvoke($async) } catch { }
}

function Wait-StdEngine {
    <#
    .SYNOPSIS
    Czeka na koniec pracy silnika (testy, tryb bez okna) i zwraca Get-StdEngineResult.
    Przekroczenie -TimeoutSec -> [STD-ENGINE-TIMEOUT] (silnik pracuje dalej).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][object]$Engine,
        [ValidateRange(1, 604800)]
        [int]$TimeoutSec = 600
    )
    $null = Get-EngineHandleShared $Engine
    $async = Get-CryptoMember $Engine 'AsyncResult'
    if ($null -eq $async) { throw '[STD-ENGINE-ARGS] Uchwyt silnika nie ma AsyncResult.' }
    $done = $async.IsCompleted
    if (-not $done) { $done = $async.AsyncWaitHandle.WaitOne([int]([int64]$TimeoutSec * 1000)) }
    if (-not $done) {
        throw ('[STD-ENGINE-TIMEOUT] Silnik nie zakończył pracy w ciągu {0} s.' -f $TimeoutSec)
    }
    Complete-EngineInvocation -Engine $Engine
    return (Get-StdEngineResult -Engine $Engine)
}

function Get-StdEngineResult {
    <#
    .SYNOPSIS
    Bieżący wynik silnika: @{ Status = 'Running'|'Paused'|'Finished'|'Cancelled'|'Failed'; Steps (lista @{ Id; Title;
    Phase; State; Detail; StartedAt; FinishedAt }); Plan (tryb Plan); Counts @{ Ok; Skipped; Error; Cancelled };
    RebootRequired; Error }. Nie czeka i nie blokuje silnika.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][object]$Engine)
    $shared = Get-EngineHandleShared $Engine
    $status = [string]$shared.Status
    $snapshot = $shared.Result
    $result = @{
        Status         = $status
        Steps          = @()
        Plan           = $null
        Counts         = @{ Ok = 0; Skipped = 0; Error = 0; Cancelled = 0 }
        RebootRequired = $false
        Error          = $null
    }
    if ($snapshot -is [System.Collections.IDictionary]) {
        $steps = New-Object 'System.Collections.Generic.List[object]'
        foreach ($step in @($snapshot['Steps'])) { if ($step -is [hashtable]) { $steps.Add($step.Clone()) } }
        $result.Steps = $steps.ToArray()
        if ($null -ne $snapshot['Plan']) {
            $plan = New-Object 'System.Collections.Generic.List[object]'
            foreach ($entry in @($snapshot['Plan'])) { if ($entry -is [hashtable]) { $plan.Add($entry.Clone()) } }
            $result.Plan = $plan.ToArray()
        }
        if ($snapshot['Counts'] -is [hashtable]) { $result.Counts = $snapshot['Counts'].Clone() }
        $result.RebootRequired = [bool]$snapshot['RebootRequired']
        $result.Error = $snapshot['Error']
    }
    # Wątek silnika skończył się, a stan nie jest końcowy - awaria poza pętlą silnika.
    $async = Get-CryptoMember $Engine 'AsyncResult'
    if (($status -eq 'Running' -or $status -eq 'Paused') -and $null -ne $async -and $async.IsCompleted) {
        $reason = $null
        $powerShell = Get-CryptoMember $Engine 'PowerShell'
        try {
            if ($null -ne $powerShell.InvocationStateInfo.Reason) { $reason = Format-CoreErrorText $powerShell.InvocationStateInfo.Reason }
            elseif ($powerShell.Streams.Error.Count -gt 0) { $reason = Format-CoreErrorText $powerShell.Streams.Error[0] }
        } catch { }
        if ([string]::IsNullOrEmpty($reason)) { $reason = 'brak szczegółów' }
        $result.Status = 'Failed'
        $result.Error = '[STD-ENGINE-FAILED] Wątek silnika zakończył się nieoczekiwanie: {0}' -f $reason
    }
    return $result
}

function Close-StdEngine {
    <#
    .SYNOPSIS
    Zwalnia runspace silnika. Gdy silnik jeszcze pracuje: przerwanie z -Force, do 15 s czekania, potem Stop.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][object]$Engine)
    $shared = Get-EngineHandleShared $Engine
    $powerShell = Get-CryptoMember $Engine 'PowerShell'
    $runspace = Get-CryptoMember $Engine 'Runspace'
    $async = Get-CryptoMember $Engine 'AsyncResult'
    if ($null -ne $async -and -not $async.IsCompleted) {
        $shared.CancelRequested = $true
        $shared.ForceCancel = $true
        $finished = $false
        try { $finished = $async.AsyncWaitHandle.WaitOne(15000) } catch { }
        if (-not $finished -and $null -ne $powerShell) {
            try { $powerShell.Stop() } catch { }
        }
    }
    Complete-EngineInvocation -Engine $Engine
    if ($null -ne $powerShell) { try { $powerShell.Dispose() } catch { } }
    if ($null -ne $runspace) {
        try { $runspace.Close() } catch { }
        try { $runspace.Dispose() } catch { }
    }
}

# ---------------------------------------------------------------------------------------------
# Programy i zasilanie (SPEC 6.4, 9)
# ---------------------------------------------------------------------------------------------

function Invoke-StdProcess {
    <#
    .SYNOPSIS
    Uruchamia program bez okna i czeka na koniec: co 500 ms sprawdza Context.Shared.ForceCancel i limit czasu;
    przy przerwaniu / przekroczeniu czasu zabija drzewo procesów (taskkill /T /F, na Linuksie Kill(true)).
    Wynik @{ ExitCode; TimedOut; Cancelled; Success (kod w -SuccessCodes, domyślnie 0, 1641, 3010); DurationSec }.
    Polecenie trafia do logu zamaskowane. Program się nie uruchomił -> [STD-PROC-START].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,
        [Parameter(Position = 1)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ArgumentList,
        [ValidateRange(0.001, 1440)]
        [double]$TimeoutMin = 30,
        [AllowNull()]
        [object]$Context,
        [string]$WorkingDirectory,
        [int[]]$SuccessCodes = @(0, 1641, 3010)
    )
    if ($null -eq $SuccessCodes -or $SuccessCodes.Count -eq 0) { $SuccessCodes = $script:StdEngineDefaultSuccessCodes }
    $shared = Get-CryptoMember $Context 'Shared'
    $commandText = $FilePath
    if (-not [string]::IsNullOrEmpty($ArgumentList)) { $commandText = '{0} {1}' -f $FilePath, $ArgumentList }
    $source = 'Invoke-StdProcess'

    if (Get-EngineFlag $shared 'ForceCancel') {
        Write-StdLog -Level Warn -Source $source -Message ('Nie uruchamiam {0} - wdrożenie zostało przerwane.' -f $commandText)
        return @{ ExitCode = $null; TimedOut = $false; Cancelled = $true; Success = $false; DurationSec = 0 }
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    if (-not [string]::IsNullOrEmpty($ArgumentList)) { $psi.Arguments = $ArgumentList }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $psi.WorkingDirectory = Resolve-CoreFullPath -Path $WorkingDirectory }

    Write-StdLog -Source $source -Message ('Uruchamiam: {0} (limit {1} min).' -f $commandText, (Get-EngineInvariantNumber $TimeoutMin))
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $process = $null
    try {
        $process = [System.Diagnostics.Process]::Start($psi)
    } catch {
        throw ('[STD-PROC-START] Nie udało się uruchomić {0}: {1}' -f (Protect-StdLogText -Text $FilePath), (Format-CoreErrorText $_))
    }
    if ($null -eq $process) {
        throw ('[STD-PROC-START] Nie udało się uruchomić {0}.' -f (Protect-StdLogText -Text $FilePath))
    }
    try {
        # Windows PowerShell 5.1: bez odczytu Handle zaraz po starcie ExitCode bywa pusty.
        $null = $process.Handle
        $timedOut = $false
        $cancelled = $false
        $limitMs = $TimeoutMin * 60000.0
        while (-not $process.WaitForExit(500)) {
            if (Get-EngineFlag $shared 'ForceCancel') { $cancelled = $true; break }
            if ($watch.Elapsed.TotalMilliseconds -ge $limitMs) { $timedOut = $true; break }
        }
        if ($timedOut -or $cancelled) {
            Stop-EngineProcessTree -Process $process
        } else {
            # Bez parametru: czeka też na zamknięcie strumieni procesu.
            $process.WaitForExit()
        }
        $watch.Stop()
        $exitCode = $null
        try { if ($process.HasExited) { $exitCode = [int]$process.ExitCode } } catch { }
        $success = (-not $timedOut) -and (-not $cancelled) -and ($null -ne $exitCode) -and ($SuccessCodes -contains $exitCode)
        $duration = [System.Math]::Round($watch.Elapsed.TotalSeconds, 1)
        if ($timedOut) {
            Write-StdLog -Level Warn -Source $source -Message ('{0}: przekroczono limit czasu ({1} min) - proces zatrzymany po {2} s.' -f $FilePath, (Get-EngineInvariantNumber $TimeoutMin), (Get-EngineInvariantNumber $duration))
        } elseif ($cancelled) {
            Write-StdLog -Level Warn -Source $source -Message ('{0}: przerwano na żądanie - proces zatrzymany po {1} s.' -f $FilePath, (Get-EngineInvariantNumber $duration))
        } else {
            $level = 'Info'
            if (-not $success) { $level = 'Warn' }
            Write-StdLog -Level $level -Source $source -Message ('{0}: kod wyjścia {1} po {2} s.' -f $FilePath, $exitCode, (Get-EngineInvariantNumber $duration))
        }
        return @{
            ExitCode    = $exitCode
            TimedOut    = $timedOut
            Cancelled   = $cancelled
            Success     = [bool]$success
            DurationSec = $duration
        }
    } finally {
        try { $process.Dispose() } catch { }
    }
}

function Set-StdKeepAwake {
    <#
    .SYNOPSIS
    Wstrzymuje (-Enabled $true) albo przywraca (-Enabled $false) usypianie komputera i wygaszanie ekranu
    przez SetThreadExecutionState w bieżącym wątku (wątek silnika). Zwraca $true, gdy się udało; poza
    Windows zawsze $false. Silnik sam przywraca usypianie na końcu wdrożenia.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [bool]$Enabled
    )
    if (-not (Initialize-EnginePowerType)) { return $false }
    $type = $script:StdEnginePowerTypeName -as [type]
    $ok = $false
    try {
        if ($Enabled) { $ok = [bool]$type::KeepAwake() } else { $ok = [bool]$type::Restore() }
    } catch {
        $ok = $false
    }
    if ($ok) { $script:StdKeepAwakeActive = $Enabled }
    return $ok
}

# ---------------------------------------------------------------------------------------------
# Informacje o komputerze (SPEC 7.3)
# ---------------------------------------------------------------------------------------------

function Get-StdComputerInfo {
    <#
    .SYNOPSIS
    Informacje o komputerze przez CIM: @{ ComputerName; Manufacturer; Model; SerialNumber; Bios; Os; OsBuild; Cpu;
    RamGB; Disks (@{ Name; SizeGB; FreeGB; Type }); Tpm (@{ Present; Ready; Version }); Network (@{ Name; Ipv4; Mac }) }.
    Każde pole osobno - brak danych to $null, nie błąd. Bez WPF; okno woła ją w tle.
    #>
    [CmdletBinding()]
    param()
    $info = @{
        ComputerName = $null
        Manufacturer = $null
        Model        = $null
        SerialNumber = $null
        Bios         = $null
        Os           = $null
        OsBuild      = $null
        Cpu          = $null
        RamGB        = $null
        Disks        = $null
        Tpm          = $null
        Network      = $null
    }
    try { $info.ComputerName = Get-EngineCleanText ([System.Environment]::MachineName) } catch { }

    $system = Get-EngineFirstCim -ClassName 'Win32_ComputerSystem'
    try { $info.Manufacturer = Get-EngineCleanText $system.Manufacturer } catch { }
    try { $info.Model = Get-EngineCleanText $system.Model } catch { }
    try {
        # Suma kości pamięci (16 GB), a gdy jej brak - pamięć widziana przez system (np. 15,8 GB).
        $modules = Get-EngineCimInstances -ClassName 'Win32_PhysicalMemory'
        $total = 0.0
        if ($null -ne $modules) {
            foreach ($module in $modules) { if ($null -ne $module.Capacity) { $total += [double]$module.Capacity } }
        }
        if ($total -le 0 -and $null -ne $system -and $null -ne $system.TotalPhysicalMemory) { $total = [double]$system.TotalPhysicalMemory }
        if ($total -gt 0) { $info.RamGB = [System.Math]::Round($total / 1GB, 1) }
    } catch { }

    $bios = Get-EngineFirstCim -ClassName 'Win32_BIOS'
    try { $info.SerialNumber = Get-EngineSerialText $bios.SerialNumber } catch { }
    try {
        if ($null -ne $bios) {
            $parts = New-Object 'System.Collections.Generic.List[string]'
            foreach ($part in @($bios.Manufacturer, $bios.SMBIOSBIOSVersion)) {
                $text = Get-EngineCleanText $part
                if ($null -ne $text) { $parts.Add($text) }
            }
            if ($bios.ReleaseDate -is [datetime]) {
                $parts.Add(('({0})' -f $bios.ReleaseDate.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)))
            }
            if ($parts.Count -gt 0) { $info.Bios = $parts -join ' ' }
        }
    } catch { }

    $os = Get-EngineFirstCim -ClassName 'Win32_OperatingSystem'
    $currentVersion = $null
    try { $currentVersion = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop } catch { }
    try {
        $caption = Get-EngineCleanText $os.Caption
        if ($null -ne $caption) {
            $display = $null
            if ($null -ne $currentVersion) {
                $display = Get-EngineCleanText $currentVersion.DisplayVersion
                if ($null -eq $display) { $display = Get-EngineCleanText $currentVersion.ReleaseId }
            }
            if ($null -ne $display) { $caption = '{0} {1}' -f $caption, $display }
            $info.Os = $caption
        }
    } catch { }
    try {
        $build = Get-EngineCleanText $os.BuildNumber
        if ($null -eq $build -and $null -ne $currentVersion) { $build = Get-EngineCleanText $currentVersion.CurrentBuildNumber }
        if ($null -ne $build) {
            if ($null -ne $currentVersion -and $null -ne $currentVersion.UBR) { $build = '{0}.{1}' -f $build, $currentVersion.UBR }
            $info.OsBuild = $build
        }
    } catch { }

    try {
        $processors = Get-EngineCimInstances -ClassName 'Win32_Processor'
        if ($null -ne $processors) {
            $names = New-Object 'System.Collections.Generic.List[string]'
            foreach ($processor in $processors) {
                $name = Get-EngineCleanText $processor.Name
                if ($null -eq $name) { continue }
                $name = $name -replace '\s+', ' '
                if (-not $names.Contains($name)) { $names.Add($name) }
            }
            if ($names.Count -gt 0) { $info.Cpu = $names -join '; ' }
        }
    } catch { }

    try { $info.Disks = Get-EngineDiskList } catch { }
    try { $info.Tpm = Get-EngineTpmInfo } catch { }
    try { $info.Network = Get-EngineNetworkList } catch { }
    return $info
}
