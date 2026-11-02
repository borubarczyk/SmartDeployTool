# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy Modules\Std.Engine.ps1 (Pester 5, Windows PowerShell 5.1). Bez administratora i bez sieci.
# Silnik działa naprawdę - w osobnym runspace, z fałszywymi krokami zapisanymi w TestDrive (SPEC 9).
# Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' - nawiasy kwadratowe trzeba poprzedzić
# znakiem ` (w -like oznaczają zbiór znaków).

Describe 'Std.Engine' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'
        $script:IsWindowsOs = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
        $script:Utf8Bom = New-Object System.Text.UTF8Encoding($true)
        $script:StatePath = Join-Path $env:STD_DATA_ROOT 'state.json'
        $script:StepsRoot = Join-Path $PSScriptRoot '..\Steps'

        function New-TestFolder {
            param([Parameter(Mandatory = $true)][string]$Name)
            $path = Join-Path $TestDrive ('{0}-{1}' -f $Name, [guid]::NewGuid().ToString('N').Substring(0, 8))
            [void][System.IO.Directory]::CreateDirectory($path)
            return $path
        }

        function New-TestSteps {
            # Folder z plikami kroków; każdy element -Content to osobny plik (UTF-8 z BOM, jak w repozytorium).
            param([Parameter(Mandatory = $true)][string[]]$Content)
            $dir = New-TestFolder 'steps'
            for ($i = 0; $i -lt $Content.Count; $i++) {
                [System.IO.File]::WriteAllText((Join-Path $dir ('{0:D2}-kroki.ps1' -f ($i + 1))), $Content[$i], $script:Utf8Bom)
            }
            return $dir
        }

        function New-TestStepText {
            # Tekst Register-StdStep dla kroku testowego. Run najpierw dopisuje Id kroku do runs.txt w WorkDir,
            # potem wykonuje -Run (tekst PowerShella, ostatnia wartość to wynik).
            param(
                [Parameter(Mandatory = $true)][string]$Id,
                [int]$Phase = 1,
                [int]$Order = 10,
                [string]$TaskId,
                [switch]$Critical,
                [string]$Detect = "@{ State = 'Todo'; Detail = 'do zrobienia' }",
                [string]$Run = "@{ Result = 'Ok'; Detail = 'zrobione' }"
            )
            $extra = ''
            if (-not [string]::IsNullOrEmpty($TaskId)) { $extra += " -TaskId '$TaskId'" }
            if ($Critical) { $extra += ' -Critical' }
            return @"
Register-StdStep -Id '$Id' -Title 'Krok $Id' -Phase $Phase -Order $Order$extra -Detect {
    param(`$Context)
    $Detect
} -Run {
    param(`$Context)
    [System.IO.File]::AppendAllText((Join-Path `$Context.Config.WorkDir 'runs.txt'), '$Id' + [Environment]::NewLine)
    $Run
}
"@
        }

        # Run czekający na plik go.txt w WorkDir (pauza / przerwanie w trakcie kroku).
        $script:WaitForGo = @'
$go = Join-Path $Context.Config.WorkDir 'go.txt'
$limit = [DateTime]::UtcNow.AddSeconds(60)
while (-not [System.IO.File]::Exists($go) -and [DateTime]::UtcNow -lt $limit) { Start-Sleep -Milliseconds 50 }
@{ Result = 'Ok'; Detail = 'doczekano się' }
'@
        # Run, który kończy się błędem, gdy w WorkDir jest plik fail.txt.
        $script:FailWhenMarked = @'
if ([System.IO.File]::Exists((Join-Path $Context.Config.WorkDir 'fail.txt'))) { throw 'Celowy błąd kroku' }
@{ Result = 'Ok'; Detail = 'zrobione' }
'@

        function New-TestConfig {
            param([Parameter(Mandatory = $true)][string]$WorkDir, [hashtable]$Tasks = @{}, [hashtable]$Extra = @{})
            $config = @{
                ClientId    = 'testowy'
                ClientName  = 'Klient testowy'
                Revision    = 7
                Site        = @{ id = 'krk'; name = 'Kraków' }
                ProfileName = 'Biuro'
                Sources     = @{}
                Catalog     = @{}
                Tasks       = $Tasks
                Answers     = @{ computerName = 'KRK-TEST01'; apps = @('7zip'); tasks = @{ RunScripts = $true } }
                WorkDir     = $WorkDir
            }
            foreach ($key in @($Extra.Keys)) { $config[$key] = $Extra[$key] }
            return $config
        }

        function Receive-TestEvents {
            param([Parameter(Mandatory = $true)]$Engine, [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Into)
            foreach ($item in @(Receive-StdEngineEvent -Engine $Engine)) { $Into.Add($item) }
        }

        function Wait-TestEvent {
            # Zbiera zdarzenia do -Into, aż któreś spełni -Where (param($e)).
            param(
                [Parameter(Mandatory = $true)]$Engine,
                [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Into,
                [Parameter(Mandatory = $true)][scriptblock]$Where,
                [int]$TimeoutSec = 60
            )
            $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSec)
            $checked = 0
            while ([DateTime]::UtcNow -lt $deadline) {
                Receive-TestEvents -Engine $Engine -Into $Into
                while ($checked -lt $Into.Count) {
                    $candidate = $Into[$checked]
                    $checked++
                    if (& $Where $candidate) { return $candidate }
                }
                Start-Sleep -Milliseconds 50
            }
            throw "Nie doczekano się zdarzenia w ciągu $TimeoutSec s."
        }

        function Invoke-TestEngine {
            # Start, czekanie na koniec, wszystkie zdarzenia, zamknięcie runspace'u.
            param([Parameter(Mandatory = $true)][hashtable]$Start)
            $engine = Start-StdEngine @Start
            try {
                $result = Wait-StdEngine -Engine $engine -TimeoutSec 120
                $events = New-Object 'System.Collections.Generic.List[object]'
                Receive-TestEvents -Engine $engine -Into $events
                return @{ Id = $engine.Id; Result = $result; Events = $events.ToArray() }
            } finally {
                Close-StdEngine -Engine $engine
            }
        }

        function Get-TestStep {
            param([Parameter(Mandatory = $true)]$Result, [Parameter(Mandatory = $true)][string]$Id)
            foreach ($step in @($Result.Steps)) { if ($step.Id -eq $Id) { return $step } }
            return $null
        }

        function Get-TestPlanEntry {
            param([Parameter(Mandatory = $true)]$Plan, [Parameter(Mandatory = $true)][string]$Id)
            foreach ($entry in @($Plan)) { if ($entry.Id -eq $Id) { return $entry } }
            return $null
        }

        function Get-TestRuns {
            param([Parameter(Mandatory = $true)][string]$WorkDir)
            $path = Join-Path $WorkDir 'runs.txt'
            if (-not [System.IO.File]::Exists($path)) { return @() }
            return @([System.IO.File]::ReadAllLines($path) | Where-Object { $_ -ne '' })
        }

        function Get-TestPowerShellExe {
            return (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
        }

        function Get-TestExitCommand {
            param([int]$Code)
            if ($script:IsWindowsOs) { return @{ File = (Join-Path $env:SystemRoot 'System32\cmd.exe'); Args = "/c exit $Code" } }
            return @{ File = '/bin/sh'; Args = "-c `"exit $Code`"" }
        }

        function Get-TestSleepCommand {
            param([int]$Seconds)
            if ($script:IsWindowsOs) { return @{ File = (Get-TestPowerShellExe); Args = "-NoProfile -NonInteractive -Command Start-Sleep -Seconds $Seconds" } }
            return @{ File = '/bin/sh'; Args = "-c `"sleep $Seconds`"" }
        }

        function Get-TestTreeCommand {
            # Proces, który uruchamia proces potomny (zapisuje jego PID do -PidFile) i czeka.
            param([Parameter(Mandatory = $true)][string]$PidFile, [int]$Seconds)
            if ($script:IsWindowsOs) {
                $exe = Get-TestPowerShellExe
                $inner = "`$p = Start-Process -FilePath '$exe' -ArgumentList '-NoProfile','-NonInteractive','-Command','Start-Sleep -Seconds $Seconds' -WindowStyle Hidden -PassThru; [System.IO.File]::WriteAllText('$PidFile', [string]`$p.Id); Start-Sleep -Seconds $Seconds"
                $encoded = [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($inner))
                return @{ File = $exe; Args = "-NoProfile -NonInteractive -EncodedCommand $encoded" }
            }
            return @{ File = '/bin/sh'; Args = "-c `"sleep $Seconds & echo `$! > '$PidFile'; wait`"" }
        }

        function Test-TestProcessAlive {
            param([int]$ProcessId)
            $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
            if ($null -eq $process) { return $false }
            if (-not $script:IsWindowsOs) {
                # Na Linuksie zabity proces może zostać „zombie” do czasu odebrania przez init - to już nie działa.
                try {
                    $stat = [System.IO.File]::ReadAllText("/proc/$ProcessId/stat")
                    $state = $stat.Substring($stat.LastIndexOf(')') + 2, 1)
                    if ($state -eq 'Z' -or $state -eq 'X') { return $false }
                } catch {
                    return $false
                }
            }
            try { return (-not $process.HasExited) } catch { return $true }
        }

        function Wait-TestProcessGone {
            param([int]$ProcessId, [int]$TimeoutSec = 15)
            $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSec)
            while ([DateTime]::UtcNow -lt $deadline) {
                if (-not (Test-TestProcessAlive -ProcessId $ProcessId)) { return $true }
                Start-Sleep -Milliseconds 200
            }
            return $false
        }
    }

    AfterAll {
        Clear-StdSteps
        Clear-StdLogMasks
        Remove-Item Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
    }

    Context 'Rejestr kroków' {
        BeforeEach { Clear-StdSteps }
        AfterAll { Clear-StdSteps }

        It 'Get-StdStep zwraca kroki w kolejności faza, Order, rejestracja' {
            Register-StdStep -Id 'Test.Late' -Title 'Późny' -Phase 2 -Order 1 -Detect { @{ State = 'Todo' } } -Run { @{ Result = 'Ok' } }
            Register-StdStep -Id 'Test.B' -Title 'B' -Phase 1 -Order 20 -Detect { @{ State = 'Todo' } } -Run { @{ Result = 'Ok' } }
            Register-StdStep -Id 'Test.A' -Title 'A' -Phase 1 -Order 10 -TaskId 'runscripts' -Critical -Detect { @{ State = 'Todo' } } -Run { @{ Result = 'Ok' } }
            Register-StdStep -Id 'Test.A2' -Title 'A2' -Phase 1 -Order 10 -Detect { @{ State = 'Todo' } } -Run { @{ Result = 'Ok' } }
            (@(Get-StdStep) | ForEach-Object { $_.Id }) -join ',' | Should -Be 'Test.A,Test.A2,Test.B,Test.Late'
            $step = Get-StdStep -Id 'Test.A'
            $step.TaskId | Should -Be 'RunScripts'
            $step.Critical | Should -BeTrue
            $step.TimeoutMin | Should -Be 30
            Get-StdStep -Id 'Test.Brak' | Should -BeNullOrEmpty
        }

        It 'drugi krok o tym samym Id -> [STD-ENGINE-STEP]' {
            Register-StdStep -Id 'Test.Dup' -Title 'Raz' -Phase 1 -Order 1 -Detect { @{ State = 'Todo' } } -Run { @{ Result = 'Ok' } }
            { Register-StdStep -Id 'Test.Dup' -Title 'Dwa' -Phase 1 -Order 2 -Detect { @{ State = 'Todo' } } -Run { @{ Result = 'Ok' } } } |
                Should -Throw -ExpectedMessage '*`[STD-ENGINE-STEP`]*'
        }

        It 'zły Id, faza albo TaskId -> [STD-ENGINE-STEP]' {
            { Register-StdStep -Id 'BezKropki' -Title 'X' -Phase 1 -Order 1 -Detect { } -Run { } } | Should -Throw -ExpectedMessage '*`[STD-ENGINE-STEP`]*'
            { Register-StdStep -Id 'Test.Faza' -Title 'X' -Phase 3 -Order 1 -Detect { } -Run { } } | Should -Throw -ExpectedMessage '*`[STD-ENGINE-STEP`]*'
            { Register-StdStep -Id 'Test.Zadanie' -Title 'X' -Phase 1 -Order 1 -TaskId 'Nieznane' -Detect { } -Run { } } | Should -Throw -ExpectedMessage '*`[STD-ENGINE-STEP`]*'
            { Register-StdStep -Id 'Test.Tytul' -Title ' ' -Phase 1 -Order 1 -Detect { } -Run { } } | Should -Throw -ExpectedMessage '*`[STD-ENGINE-STEP`]*'
        }

        It 'Clear-StdSteps usuwa wszystkie kroki' {
            Register-StdStep -Id 'Test.X' -Title 'X' -Phase 1 -Order 1 -Detect { @{ State = 'Todo' } } -Run { @{ Result = 'Ok' } }
            @(Get-StdStep).Count | Should -Be 1
            Clear-StdSteps
            @(Get-StdStep).Count | Should -Be 0
        }

        It 'kroki etapu 2 (Steps\Network.ps1, Steps\Power.ps1) rejestrują się z właściwym zadaniem' {
            . (Join-Path $script:StepsRoot 'Network.ps1')
            . (Join-Path $script:StepsRoot 'Power.ps1')
            $network = Get-StdStep -Id 'Network.WaitForNetwork'
            $network.TaskId | Should -Be 'WaitForNetwork'
            $network.Phase | Should -Be 1
            $network.TimeoutMin | Should -Be 10
            $power = Get-StdStep -Id 'Power.SuspendSleep'
            $power.TaskId | Should -Be 'SuspendSleep'
            $power.Order | Should -BeLessThan $network.Order
        }

        It 'Test-StdEngineCancel czyta Shared.CancelRequested' {
            $shared = [hashtable]::Synchronized(@{ CancelRequested = $false })
            Test-StdEngineCancel -Context @{ Shared = $shared } | Should -BeFalse
            $shared.CancelRequested = $true
            Test-StdEngineCancel -Context @{ Shared = $shared } | Should -BeTrue
            Test-StdEngineCancel -Context $null | Should -BeFalse
        }
    }

    Context 'Plan („Sprawdź”)' {
        It 'pokazuje Done / Todo / Blocked / Disabled, wyjątek w Detect = Blocked, niczego nie uruchamia' {
            $work = New-TestFolder 'plan'
            $dir = New-TestSteps -Content @(
                (New-TestStepText -Id 'Test.Done' -Order 10 -Detect "@{ State = 'Done'; Detail = 'już jest' }"),
                (New-TestStepText -Id 'Test.Todo' -Order 20 -TaskId 'RunScripts'),
                (New-TestStepText -Id 'Test.Off' -Order 30 -TaskId 'HardwareAudit'),
                ((New-TestStepText -Id 'Test.Boom' -Order 40 -Detect "throw 'Detect wybuchł'") + "`n" +
                    (New-TestStepText -Id 'Test.Later' -Phase 2 -Order 5 -Detect "@{ State = 'Blocked'; Detail = 'brak dysku' }"))
            )
            if (Test-Path -LiteralPath $script:StatePath) { Remove-Item -LiteralPath $script:StatePath -Force }
            $config = New-TestConfig -WorkDir $work -Tasks @{ RunScripts = $true; HardwareAudit = $false }
            $run = Invoke-TestEngine -Start @{ Mode = 'Plan'; Config = $config; StepsPath = $dir }

            $run.Result.Status | Should -Be 'Finished'
            $plan = @($run.Result.Plan)
            ($plan | ForEach-Object { $_.Id }) -join ',' | Should -Be 'Test.Done,Test.Todo,Test.Off,Test.Boom,Test.Later'
            (Get-TestPlanEntry $plan 'Test.Done').State | Should -Be 'Done'
            (Get-TestPlanEntry $plan 'Test.Done').Detail | Should -Be 'już jest'
            (Get-TestPlanEntry $plan 'Test.Todo').State | Should -Be 'Todo'
            (Get-TestPlanEntry $plan 'Test.Todo').Enabled | Should -BeTrue
            (Get-TestPlanEntry $plan 'Test.Off').State | Should -Be 'Disabled'
            (Get-TestPlanEntry $plan 'Test.Off').Enabled | Should -BeFalse
            (Get-TestPlanEntry $plan 'Test.Boom').State | Should -Be 'Blocked'
            (Get-TestPlanEntry $plan 'Test.Boom').Detail | Should -BeLike '*Detect wybuchł*'
            (Get-TestPlanEntry $plan 'Test.Later').State | Should -Be 'Blocked'
            (Get-TestPlanEntry $plan 'Test.Later').Phase | Should -Be 2

            $planEvents = @($run.Events | Where-Object { $_.Type -eq 'Plan' })
            $planEvents.Count | Should -Be 1
            @($planEvents[0].Steps).Count | Should -Be 5
            @($run.Events)[-1].Type | Should -Be 'Finished'
            @(Get-TestRuns $work).Count | Should -Be 0
            @($run.Result.Steps).Count | Should -Be 0
            Test-Path -LiteralPath $script:StatePath | Should -BeFalse
        }

        It 'domyślne kroki z folderu Steps ładują się w wątku silnika' {
            $work = New-TestFolder 'plan-default'
            $config = New-TestConfig -WorkDir $work -Tasks @{ WaitForNetwork = $true; SuspendSleep = $false }
            $run = Invoke-TestEngine -Start @{ Mode = 'Plan'; Config = $config }
            $run.Result.Status | Should -Be 'Finished'
            $run.Result.Error | Should -BeNullOrEmpty
            $network = Get-TestPlanEntry $run.Result.Plan 'Network.WaitForNetwork'
            $network.Enabled | Should -BeTrue
            @('Done', 'Todo') | Should -Contain $network.State
            (Get-TestPlanEntry $run.Result.Plan 'Power.SuspendSleep').State | Should -Be 'Disabled'
        }
    }

    Context 'Wdrożenie („Wdróż”)' {
        It 'uruchamia kroki w kolejności: faza, potem Order (z wielu plików)' {
            $work = New-TestFolder 'order'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.P2' -Phase 2 -Order 5) + "`n" + (New-TestStepText -Id 'Test.P1b' -Order 20)),
                (New-TestStepText -Id 'Test.P1a' -Order 10)
            )
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = (New-TestConfig -WorkDir $work); StepsPath = $dir }
            $run.Result.Status | Should -Be 'Finished'
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.P1a,Test.P1b,Test.P2'
            (@($run.Result.Steps) | ForEach-Object { $_.Id }) -join ',' | Should -Be 'Test.P1a,Test.P1b,Test.P2'
            $run.Result.Counts.Ok | Should -Be 3
            $running = @($run.Events | Where-Object { $_.Type -eq 'StepState' -and $_.State -eq 'Running' } | ForEach-Object { $_.StepId })
            $running -join ',' | Should -Be 'Test.P1a,Test.P1b,Test.P2'
            foreach ($step in @($run.Result.Steps)) {
                $step.StartedAt | Should -BeOfType [datetime]
                $step.FinishedAt | Should -BeOfType [datetime]
            }
            $finished = @($run.Events | Where-Object { $_.Type -eq 'Finished' })
            $finished.Count | Should -Be 1
            $finished[0].Status | Should -Be 'Finished'
            $finished[0].Counts.Ok | Should -Be 3
        }

        It 'Detect Done -> Skipped, Blocked -> Error, wyjątek -> Error, zły wynik -> Error' {
            $work = New-TestFolder 'detect'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.Done' -Order 10 -Detect "@{ State = 'Done'; Detail = 'zainstalowane' }") + "`n" +
                    (New-TestStepText -Id 'Test.Blocked' -Order 20 -Detect "@{ State = 'Blocked'; Detail = 'brak źródła' }") + "`n" +
                    (New-TestStepText -Id 'Test.DetectThrow' -Order 30 -Detect "throw 'Detect padł'") + "`n" +
                    (New-TestStepText -Id 'Test.RunThrow' -Order 40 -Run "throw 'Run padł'") + "`n" +
                    (New-TestStepText -Id 'Test.Bad' -Order 50 -Run "'tylko tekst'") + "`n" +
                    (New-TestStepText -Id 'Test.Skip' -Order 60 -Run "Write-Output 'śmieci'; @{ Result = 'skipped'; Detail = 'nie dotyczy' }"))
            )
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = (New-TestConfig -WorkDir $work); StepsPath = $dir }
            $result = $run.Result
            $result.Status | Should -Be 'Finished'
            (Get-TestStep $result 'Test.Done').State | Should -Be 'Skipped'
            (Get-TestStep $result 'Test.Done').Detail | Should -BeLike 'Już zrobione*zainstalowane*'
            (Get-TestStep $result 'Test.Blocked').State | Should -Be 'Error'
            (Get-TestStep $result 'Test.Blocked').Detail | Should -BeLike '*brak źródła*'
            (Get-TestStep $result 'Test.DetectThrow').State | Should -Be 'Error'
            (Get-TestStep $result 'Test.DetectThrow').Detail | Should -BeLike '*Detect padł*'
            (Get-TestStep $result 'Test.RunThrow').State | Should -Be 'Error'
            (Get-TestStep $result 'Test.RunThrow').Detail | Should -BeLike '*Run padł*'
            (Get-TestStep $result 'Test.Bad').State | Should -Be 'Error'
            (Get-TestStep $result 'Test.Skip').State | Should -Be 'Skipped'
            (Get-TestStep $result 'Test.Skip').Detail | Should -Be 'nie dotyczy'
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.RunThrow,Test.Bad,Test.Skip'
            $result.Counts.Error | Should -Be 4
            $result.Counts.Skipped | Should -Be 2
        }

        It 'zadanie wyłączone w Tasks -> krok poza listą i nieuruchomiony' {
            $work = New-TestFolder 'tasks'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.On' -Order 10 -TaskId 'RunScripts') + "`n" + (New-TestStepText -Id 'Test.Off' -Order 20 -TaskId 'EnableBitLocker') +
                    "`n" + (New-TestStepText -Id 'Test.Always' -Order 30))
            )
            $config = New-TestConfig -WorkDir $work -Tasks @{ RunScripts = $true; EnableBitLocker = $false }
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = $config; StepsPath = $dir }
            (@($run.Result.Steps) | ForEach-Object { $_.Id }) -join ',' | Should -Be 'Test.On,Test.Always'
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.On,Test.Always'
        }

        It 'błąd kroku -Critical pomija resztę fazy, następna faza działa' {
            $work = New-TestFolder 'critical'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.Normal' -Order 10 -Run "throw 'Zwykły błąd'") + "`n" +
                    (New-TestStepText -Id 'Test.Crit' -Order 20 -Critical -Run "@{ Result = 'Error'; Detail = 'krytyczny problem' }") + "`n" +
                    (New-TestStepText -Id 'Test.After' -Order 30) + "`n" +
                    (New-TestStepText -Id 'Test.Next' -Phase 2 -Order 10))
            )
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = (New-TestConfig -WorkDir $work); StepsPath = $dir }
            $result = $run.Result
            $result.Status | Should -Be 'Finished'
            (Get-TestStep $result 'Test.Normal').State | Should -Be 'Error'
            (Get-TestStep $result 'Test.Crit').State | Should -Be 'Error'
            (Get-TestStep $result 'Test.Crit').Detail | Should -Be 'krytyczny problem'
            (Get-TestStep $result 'Test.After').State | Should -Be 'Skipped'
            (Get-TestStep $result 'Test.After').Detail | Should -BeLike '*krytyczny*Krok Test.Crit*'
            (Get-TestStep $result 'Test.Next').State | Should -Be 'Ok'
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.Normal,Test.Crit,Test.Next'
            $result.Counts.Error | Should -Be 2
            $result.Counts.Skipped | Should -Be 1
            $result.Counts.Ok | Should -Be 1
        }

        It 'Reboot w fazie 1: reszta fazy 1 działa, faza 2 czeka, RebootRequired' {
            $work = New-TestFolder 'reboot'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.Rename' -Order 10 -Run "@{ Result = 'Ok'; Detail = 'nowa nazwa'; Reboot = `$true }") + "`n" +
                    (New-TestStepText -Id 'Test.Join' -Order 20) + "`n" +
                    (New-TestStepText -Id 'Test.Apps' -Phase 2 -Order 10))
            )
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = (New-TestConfig -WorkDir $work); StepsPath = $dir; DeploymentId = 'reboot-test' }
            $result = $run.Result
            $result.Status | Should -Be 'Finished'
            $result.RebootRequired | Should -BeTrue
            (Get-TestStep $result 'Test.Rename').State | Should -Be 'Ok'
            (Get-TestStep $result 'Test.Join').State | Should -Be 'Ok'
            (Get-TestStep $result 'Test.Apps').State | Should -Be 'Pending'
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.Rename,Test.Join'
            @($run.Events | Where-Object { $_.Type -eq 'Finished' })[0].RebootRequired | Should -BeTrue
            $state = Read-StdJsonFile -Path $script:StatePath
            $state.deploymentId | Should -Be 'reboot-test'
            $state.rebootRequired | Should -BeTrue
            $state.phase | Should -Be 2
            $state.steps['Test.Apps'].state | Should -Be 'Pending'
            $state.steps['Test.Join'].state | Should -Be 'Ok'
        }

        It 'pauza po bieżącym kroku i wznowienie (Paused / Resumed)' {
            $work = New-TestFolder 'pause'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.Wait' -Order 10 -Run $script:WaitForGo) + "`n" + (New-TestStepText -Id 'Test.Second' -Order 20))
            )
            $engine = Start-StdEngine -Mode Run -Config (New-TestConfig -WorkDir $work) -StepsPath $dir
            try {
                $events = New-Object 'System.Collections.Generic.List[object]'
                $null = Wait-TestEvent -Engine $engine -Into $events -Where { param($e) $e.Type -eq 'StepState' -and $e.StepId -eq 'Test.Wait' -and $e.State -eq 'Running' }
                Suspend-StdEngine -Engine $engine
                [System.IO.File]::WriteAllText((Join-Path $work 'go.txt'), 'go')
                $null = Wait-TestEvent -Engine $engine -Into $events -Where { param($e) $e.Type -eq 'Paused' }
                Start-Sleep -Milliseconds 700
                $middle = Get-StdEngineResult -Engine $engine
                $middle.Status | Should -Be 'Paused'
                (Get-TestStep $middle 'Test.Wait').State | Should -Be 'Ok'
                (Get-TestStep $middle 'Test.Second').State | Should -Be 'Pending'
                (Get-TestRuns $work) -join ',' | Should -Be 'Test.Wait'
                (Read-StdJsonFile -Path $script:StatePath).status | Should -Be 'Paused'

                Resume-StdEngine -Engine $engine
                $result = Wait-StdEngine -Engine $engine -TimeoutSec 60
                Receive-TestEvents -Engine $engine -Into $events
                $result.Status | Should -Be 'Finished'
                $result.Counts.Ok | Should -Be 2
                $types = @($events | ForEach-Object { if ($_.Type -eq 'StepState') { '{0}:{1}' -f $_.StepId, $_.State } else { $_.Type } })
                $paused = [array]::IndexOf($types, 'Paused')
                $resumed = [array]::IndexOf($types, 'Resumed')
                $second = [array]::IndexOf($types, 'Test.Second:Running')
                $paused | Should -BeGreaterThan -1
                $resumed | Should -BeGreaterThan $paused
                $second | Should -BeGreaterThan $resumed
            } finally {
                Close-StdEngine -Engine $engine
            }
        }

        It 'przerwanie po bieżącym kroku: pozostałe kroki Cancelled' {
            $work = New-TestFolder 'cancel'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.Wait' -Order 10 -Run $script:WaitForGo) + "`n" + (New-TestStepText -Id 'Test.Second' -Order 20) +
                    "`n" + (New-TestStepText -Id 'Test.Third' -Phase 2 -Order 10))
            )
            $engine = Start-StdEngine -Mode Run -Config (New-TestConfig -WorkDir $work) -StepsPath $dir
            try {
                $events = New-Object 'System.Collections.Generic.List[object]'
                $null = Wait-TestEvent -Engine $engine -Into $events -Where { param($e) $e.Type -eq 'StepState' -and $e.StepId -eq 'Test.Wait' -and $e.State -eq 'Running' }
                Stop-StdEngine -Engine $engine
                [System.IO.File]::WriteAllText((Join-Path $work 'go.txt'), 'go')
                $result = Wait-StdEngine -Engine $engine -TimeoutSec 60
                Receive-TestEvents -Engine $engine -Into $events
            } finally {
                Close-StdEngine -Engine $engine
            }
            $result.Status | Should -Be 'Cancelled'
            (Get-TestStep $result 'Test.Wait').State | Should -Be 'Ok'
            (Get-TestStep $result 'Test.Second').State | Should -Be 'Cancelled'
            (Get-TestStep $result 'Test.Third').State | Should -Be 'Cancelled'
            $result.Counts.Cancelled | Should -Be 2
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.Wait'
            @($events | Where-Object { $_.Type -eq 'Finished' })[0].Status | Should -Be 'Cancelled'
            (Read-StdJsonFile -Path $script:StatePath).status | Should -Be 'Cancelled'
        }

        It '-OnlySteps uruchamia tylko podane kroki, wyniki pozostałych zostają w state.json' {
            $work = New-TestFolder 'only'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.A' -Order 10) + "`n" + (New-TestStepText -Id 'Test.B' -Order 20 -Run $script:FailWhenMarked) +
                    "`n" + (New-TestStepText -Id 'Test.C' -Order 30))
            )
            $config = New-TestConfig -WorkDir $work
            [System.IO.File]::WriteAllText((Join-Path $work 'fail.txt'), 'x')
            $first = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = $config; StepsPath = $dir; DeploymentId = 'only-test' }
            (Get-TestStep $first.Result 'Test.B').State | Should -Be 'Error'

            Remove-Item -LiteralPath (Join-Path $work 'fail.txt')
            $retry = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = $config; StepsPath = $dir; DeploymentId = 'only-test'; OnlySteps = @('Test.B') }
            $retry.Result.Status | Should -Be 'Finished'
            (@($retry.Result.Steps) | ForEach-Object { $_.Id }) -join ',' | Should -Be 'Test.B'
            (Get-TestStep $retry.Result 'Test.B').State | Should -Be 'Ok'
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.A,Test.B,Test.C,Test.B'
            $state = Read-StdJsonFile -Path $script:StatePath
            $state.steps['Test.A'].state | Should -Be 'Ok'
            $state.steps['Test.B'].state | Should -Be 'Ok'
            $state.steps['Test.C'].state | Should -Be 'Ok'
        }

        It '-Resume bierze identyfikator ze state.json i nie uruchamia ponownie kroków Ok / Skipped' {
            $work = New-TestFolder 'resume'
            $dir = New-TestSteps -Content @(
                ((New-TestStepText -Id 'Test.A' -Order 10) + "`n" + (New-TestStepText -Id 'Test.B' -Order 20 -Run $script:FailWhenMarked) +
                    "`n" + (New-TestStepText -Id 'Test.C' -Order 30 -Detect "@{ State = 'Done'; Detail = 'już jest' }"))
            )
            $config = New-TestConfig -WorkDir $work
            [System.IO.File]::WriteAllText((Join-Path $work 'fail.txt'), 'x')
            $first = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = $config; StepsPath = $dir; DeploymentId = 'resume-test' }
            (Get-TestStep $first.Result 'Test.B').State | Should -Be 'Error'
            (Get-TestStep $first.Result 'Test.C').State | Should -Be 'Skipped'

            Remove-Item -LiteralPath (Join-Path $work 'fail.txt')
            $second = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = $config; StepsPath = $dir; Resume = $true }
            $second.Id | Should -Be 'resume-test'
            $second.Result.Status | Should -Be 'Finished'
            (Get-TestStep $second.Result 'Test.A').State | Should -Be 'Ok'
            (Get-TestStep $second.Result 'Test.B').State | Should -Be 'Ok'
            (Get-TestStep $second.Result 'Test.C').State | Should -Be 'Skipped'
            (Get-TestRuns $work) -join ',' | Should -Be 'Test.A,Test.B,Test.B'
            $running = @($second.Events | Where-Object { $_.Type -eq 'StepState' -and $_.State -eq 'Running' } | ForEach-Object { $_.StepId })
            $running -join ',' | Should -Be 'Test.B'
        }

        It '-Resume bez zapisanego stanu -> [STD-ENGINE-RESUME]' {
            $work = New-TestFolder 'resume-none'
            $dir = New-TestSteps -Content @((New-TestStepText -Id 'Test.A'))
            if (Test-Path -LiteralPath $script:StatePath) { Remove-Item -LiteralPath $script:StatePath -Force }
            { Start-StdEngine -Mode Run -Config (New-TestConfig -WorkDir $work) -StepsPath $dir -Resume } |
                Should -Throw -ExpectedMessage '*`[STD-ENGINE-RESUME`]*'
        }

        It 'zły identyfikator wdrożenia albo brak folderu kroków -> [STD-ENGINE-ARGS]' {
            $work = New-TestFolder 'args'
            $dir = New-TestSteps -Content @((New-TestStepText -Id 'Test.A'))
            { Start-StdEngine -Mode Run -Config (New-TestConfig -WorkDir $work) -StepsPath $dir -DeploymentId '..\..\Windows' } |
                Should -Throw -ExpectedMessage '*`[STD-ENGINE-ARGS`]*'
            { Start-StdEngine -Mode Plan -Config (New-TestConfig -WorkDir $work) -StepsPath (Join-Path $work 'brak') } |
                Should -Throw -ExpectedMessage '*`[STD-ENGINE-ARGS`]*'
        }

        It 'błąd w pliku kroku -> Failed z [STD-ENGINE-STEP]' {
            $work = New-TestFolder 'broken'
            $dir = New-TestSteps -Content @((New-TestStepText -Id 'Test.A'), "throw 'zepsuty plik kroku'")
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = (New-TestConfig -WorkDir $work); StepsPath = $dir }
            $run.Result.Status | Should -Be 'Failed'
            $run.Result.Error | Should -BeLike '*`[STD-ENGINE-STEP`]*zepsuty plik kroku*'
            $finished = @($run.Events | Where-Object { $_.Type -eq 'Finished' })
            $finished[0].Status | Should -Be 'Failed'
            @(Get-TestRuns $work).Count | Should -Be 0
        }

        It 'zapisuje deploy.log i events.jsonl (każda linia to zdarzenie JSON)' {
            $work = New-TestFolder 'logs'
            $dir = New-TestSteps -Content @((New-TestStepText -Id 'Test.Log' -Run "Write-StdLog -Message 'Wiadomość z kroku'; @{ Result = 'Ok'; Detail = 'zrobione' }"))
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = (New-TestConfig -WorkDir $work); StepsPath = $dir; DeploymentId = 'logs-test' }
            $logDir = Join-Path $env:STD_DATA_ROOT 'Logs\logs-test'
            $deployLog = [System.IO.File]::ReadAllText((Join-Path $logDir 'deploy.log'))
            $deployLog | Should -Match '\[INFO\] Wiadomość z kroku'
            $deployLog | Should -Match 'Koniec wdrożenia logs-test'
            $lines = @([System.IO.File]::ReadAllLines((Join-Path $logDir 'events.jsonl')) | Where-Object { $_ -ne '' })
            $lines.Count | Should -Be @($run.Events).Count
            $parsed = @($lines | ForEach-Object { ConvertFrom-StdJson -Json $_ })
            $parsed[0].Type | Should -Be 'Log'
            $parsed[-1].Type | Should -Be 'Finished'
            @($parsed | Where-Object { $_.Type -eq 'StepState' -and $_.StepId -eq 'Test.Log' -and $_.State -eq 'Ok' }).Count | Should -Be 1
            @($run.Events | Where-Object { $_.Type -eq 'Log' -and $_.Message -eq 'Wiadomość z kroku' }).Count | Should -Be 1
            @($run.Events | Where-Object { $_.Type -eq 'Progress' })[-1].Percent | Should -Be 100
        }
    }

    Context 'Sekrety i maski' {
        It 'state.json, deploy.log, events.jsonl i zdarzenia nie zawierają haseł; silnik czyści Secrets' {
            # Maski okna nie przechodzą do runspace'u - silnik musi je zarejestrować sam.
            Clear-StdLogMasks
            $work = New-TestFolder 'secrets'
            $plain = @{
                DomainPassword = 'Haslo-Domeny-123!'
                DomainLogin    = 'jan.tajny'
                LocalPassword  = 'Lokalne-Haslo-456!'
                ConfigToken    = 'TV-token-789xyz'
            }
            $secrets = @{
                DomainJoin         = (New-Object System.Management.Automation.PSCredential('FIRMA\jan.tajny', (ConvertTo-SecureString $plain.DomainPassword -AsPlainText -Force)))
                LocalAdminPassword = (ConvertTo-SecureString $plain.LocalPassword -AsPlainText -Force)
                ConfigSecrets      = @{ 'tv-token' = $plain.ConfigToken }
            }
            $useSecrets = @'
$cred = $Context.Secrets['DomainJoin']
$password = $cred.GetNetworkCredential().Password
$login = $cred.UserName
$local = (New-Object System.Net.NetworkCredential('x', $Context.Secrets['LocalAdminPassword'])).Password
$token = $Context.Secrets['ConfigSecrets']['tv-token']
Write-StdLog -Level Warn -Message ('Hasło: {0}, login: {1}, lokalne: {2}, token: {3}' -f $password, $login, $local, $token)
$null = $Context.Secrets.Remove('DomainJoin')
@{ Result = 'Ok'; Detail = ('Użyto {0} / {1} / {2} / {3}' -f $password, $login, $local, $token) }
'@
            $detectSecrets = "@{ State = 'Todo'; Detail = ('Sprawdzono ' + `$Context.Secrets['ConfigSecrets']['tv-token']) }"
            $dir = New-TestSteps -Content @((New-TestStepText -Id 'Test.Secret' -Detect $detectSecrets -Run $useSecrets))
            $config = New-TestConfig -WorkDir $work
            $config.Answers['password'] = $plain.DomainPassword
            $config.Answers['DomainUser'] = 'FIRMA\jan.tajny'
            $run = Invoke-TestEngine -Start @{ Mode = 'Run'; Config = $config; StepsPath = $dir; Secrets = $secrets; DeploymentId = 'secrets-test' }

            $run.Result.Status | Should -Be 'Finished'
            (Get-TestStep $run.Result 'Test.Secret').State | Should -Be 'Ok'
            (Get-TestStep $run.Result 'Test.Secret').Detail | Should -Be 'Użyto *** / *** / *** / ***'
            $secrets.Count | Should -Be 0

            $logDir = Join-Path $env:STD_DATA_ROOT 'Logs\secrets-test'
            $texts = @{
                'state.json'   = [System.IO.File]::ReadAllText($script:StatePath)
                'deploy.log'   = [System.IO.File]::ReadAllText((Join-Path $logDir 'deploy.log'))
                'events.jsonl' = [System.IO.File]::ReadAllText((Join-Path $logDir 'events.jsonl'))
                'zdarzenia'    = (@($run.Events | ForEach-Object { ConvertTo-StdJson -InputObject $_ -Compress }) -join "`n")
                'wynik'        = (ConvertTo-StdJson -InputObject $run.Result -Compress)
            }
            foreach ($name in @($texts.Keys)) {
                foreach ($value in @($plain.Values)) {
                    $texts[$name].Contains($value) | Should -BeFalse -Because "$name nie może zawierać sekretu"
                }
            }
            $texts['deploy.log'] | Should -Match 'Hasło: \*\*\*, login: \*\*\*, lokalne: \*\*\*, token: \*\*\*'
            $state = Read-StdJsonFile -Path $script:StatePath
            $state.answers.computerName | Should -Be 'KRK-TEST01'
            @($state.answers.Keys) -join ',' | Should -Be 'computerName,apps,tasks'
            $state.clientId | Should -Be 'testowy'
            $state.siteId | Should -Be 'krk'
            $state.profile | Should -Be 'Biuro'
            $state.revision | Should -Be 7
        }

        It 'Invoke-StdProcess maskuje argumenty w logu' {
            Clear-StdLogMasks
            Register-StdLogMask -Value 'TajnyArgument-55'
            $captured = New-Object 'System.Collections.Generic.List[string]'
            $sink = { $captured.Add([string]$args[0].Message) }.GetNewClosure()
            Register-StdLogSink -Name 'Test.Capture' -ScriptBlock $sink
            try {
                if ($script:IsWindowsOs) {
                    $result = Invoke-StdProcess -FilePath (Join-Path $env:SystemRoot 'System32\cmd.exe') -ArgumentList '/c rem TajnyArgument-55'
                } else {
                    $result = Invoke-StdProcess -FilePath '/bin/sh' -ArgumentList '-c ": TajnyArgument-55"'
                }
            } finally {
                Unregister-StdLogSink -Name 'Test.Capture'
                Clear-StdLogMasks
            }
            $result.ExitCode | Should -Be 0
            $all = $captured -join "`n"
            $all | Should -Match 'Uruchamiam: .*\*\*\*'
            $all.Contains('TajnyArgument-55') | Should -BeFalse
        }
    }

    Context 'Invoke-StdProcess' {
        It 'zwraca kod wyjścia i Success według -SuccessCodes' {
            $command = Get-TestExitCommand -Code 3
            $result = Invoke-StdProcess -FilePath $command.File -ArgumentList $command.Args -TimeoutMin 2
            $result.ExitCode | Should -Be 3
            $result.Success | Should -BeFalse
            $result.TimedOut | Should -BeFalse
            $result.Cancelled | Should -BeFalse
            $result.DurationSec | Should -BeLessThan 60
            (Invoke-StdProcess -FilePath $command.File -ArgumentList $command.Args -SuccessCodes 0, 3).Success | Should -BeTrue
            $zero = Get-TestExitCommand -Code 0
            (Invoke-StdProcess -FilePath $zero.File -ArgumentList $zero.Args).Success | Should -BeTrue
            if ($script:IsWindowsOs) {
                # Na Linuksie kod wyjścia to tylko 0-255, więc 3010 („wymagany restart”) sprawdzamy w Windows.
                $reboot = Get-TestExitCommand -Code 3010
                $rebootResult = Invoke-StdProcess -FilePath $reboot.File -ArgumentList $reboot.Args
                $rebootResult.ExitCode | Should -Be 3010
                $rebootResult.Success | Should -BeTrue
            }
        }

        It 'limit czasu zabija proces' {
            $command = Get-TestSleepCommand -Seconds 60
            $result = Invoke-StdProcess -FilePath $command.File -ArgumentList $command.Args -TimeoutMin 0.05
            $result.TimedOut | Should -BeTrue
            $result.Success | Should -BeFalse
            $result.Cancelled | Should -BeFalse
            $result.DurationSec | Should -BeLessThan 40
        }

        It 'limit czasu zabija całe drzewo procesów' {
            $pidFile = Join-Path (New-TestFolder 'tree') 'child.pid'
            $command = Get-TestTreeCommand -PidFile $pidFile -Seconds 120
            $result = Invoke-StdProcess -FilePath $command.File -ArgumentList $command.Args -TimeoutMin 0.25
            $result.TimedOut | Should -BeTrue
            [System.IO.File]::Exists($pidFile) | Should -BeTrue
            $childId = [int]([System.IO.File]::ReadAllText($pidFile).Trim())
            Wait-TestProcessGone -ProcessId $childId | Should -BeTrue
        }

        It 'ForceCancel zabija proces' {
            $shared = [hashtable]::Synchronized(@{ CancelRequested = $false; ForceCancel = $false })
            $setter = [System.Management.Automation.PowerShell]::Create()
            [void]$setter.AddScript('param($s) Start-Sleep -Milliseconds 1500; $s.CancelRequested = $true; $s.ForceCancel = $true').AddArgument($shared)
            $handle = $setter.BeginInvoke()
            try {
                $command = Get-TestSleepCommand -Seconds 60
                $result = Invoke-StdProcess -FilePath $command.File -ArgumentList $command.Args -TimeoutMin 5 -Context @{ Shared = $shared }
            } finally {
                try { [void]$setter.EndInvoke($handle) } catch { }
                $setter.Dispose()
            }
            $result.Cancelled | Should -BeTrue
            $result.TimedOut | Should -BeFalse
            $result.Success | Should -BeFalse
            $result.DurationSec | Should -BeLessThan 40
            # Przy już przerwanym wdrożeniu program w ogóle nie startuje.
            $again = Invoke-StdProcess -FilePath $command.File -ArgumentList $command.Args -Context @{ Shared = $shared }
            $again.Cancelled | Should -BeTrue
            $again.ExitCode | Should -BeNullOrEmpty
        }

        It 'Stop-StdEngine -Force przerywa proces w bieżącym kroku' {
            $work = New-TestFolder 'force'
            $sleep = Get-TestSleepCommand -Seconds 90
            $procRun = @'
$r = Invoke-StdProcess -FilePath $Context.Config.SleepFile -ArgumentList $Context.Config.SleepArgs -TimeoutMin 5 -Context $Context
$state = 'Ok'
if (-not $r.Success) { $state = 'Error' }
@{ Result = $state; Detail = ('Cancelled={0} TimedOut={1}' -f $r.Cancelled, $r.TimedOut) }
'@
            $dir = New-TestSteps -Content @(((New-TestStepText -Id 'Test.Proc' -Order 10 -Run $procRun) + "`n" + (New-TestStepText -Id 'Test.After' -Order 20)))
            $config = New-TestConfig -WorkDir $work -Extra @{ SleepFile = $sleep.File; SleepArgs = $sleep.Args }
            $engine = Start-StdEngine -Mode Run -Config $config -StepsPath $dir
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $events = New-Object 'System.Collections.Generic.List[object]'
                $null = Wait-TestEvent -Engine $engine -Into $events -Where { param($e) $e.Type -eq 'Log' -and $e.Message -like 'Uruchamiam:*' }
                Start-Sleep -Milliseconds 500
                Stop-StdEngine -Engine $engine -Force
                $result = Wait-StdEngine -Engine $engine -TimeoutSec 60
            } finally {
                Close-StdEngine -Engine $engine
            }
            $watch.Elapsed.TotalSeconds | Should -BeLessThan 60
            $result.Status | Should -Be 'Cancelled'
            (Get-TestStep $result 'Test.Proc').State | Should -Be 'Cancelled'
            (Get-TestStep $result 'Test.Proc').Detail | Should -Be 'Cancelled=True TimedOut=False'
            (Get-TestStep $result 'Test.After').State | Should -Be 'Cancelled'
        }

        It 'brak programu -> [STD-PROC-START]' {
            { Invoke-StdProcess -FilePath (Join-Path $TestDrive 'nie-ma-takiego-programu.exe') } | Should -Throw -ExpectedMessage '*`[STD-PROC-START`]*'
        }
    }

    Context 'Zasilanie i informacje o komputerze' {
        It 'Set-StdKeepAwake działa w Windows, a w innych systemach zwraca $false' {
            if ($script:IsWindowsOs) {
                Set-StdKeepAwake -Enabled $true | Should -BeTrue
                Set-StdKeepAwake -Enabled $false | Should -BeTrue
            } else {
                Set-StdKeepAwake -Enabled $true | Should -BeFalse
            }
        }

        It 'Get-StdComputerInfo zwraca wszystkie pola i nie rzuca wyjątków' {
            $info = Get-StdComputerInfo
            foreach ($key in @('ComputerName', 'Manufacturer', 'Model', 'SerialNumber', 'Bios', 'Os', 'OsBuild', 'Cpu', 'RamGB', 'Disks', 'Tpm', 'Network')) {
                $info.ContainsKey($key) | Should -BeTrue -Because "brak pola $key"
            }
            $info.ComputerName | Should -Not -BeNullOrEmpty
            if ($script:IsWindowsOs) {
                $info.Os | Should -Not -BeNullOrEmpty
                $info.RamGB | Should -BeGreaterThan 0
                @($info.Disks).Count | Should -BeGreaterThan 0
                foreach ($disk in @($info.Disks)) { $disk.ContainsKey('FreeGB') | Should -BeTrue }
            }
        }
    }
}
