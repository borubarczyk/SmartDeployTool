# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy Modules\Std.App.ps1 i plików startowych (SPEC 8): kopia narzędzia do App\<wersja>, pomijanie
# identycznej kopii, kopia po zmianie pliku, usuwanie starych wersji, argumenty ponownego uruchomienia
# (cudzysłowy, spacje, przełączniki, -OriginRoot). Pester 5, Windows PowerShell 5.1. Bez administratora
# i bez sieci; dane lokalne w TestDrive ($env:STD_DATA_ROOT). Kody błędów sprawdzamy przez
# -ExpectedMessage '*`[STD-...`]*' - nawiasy kwadratowe trzeba poprzedzić znakiem ` (w -like oznaczają zbiór znaków).

Describe 'Std.App' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $script:V4Root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
        $script:OldDate = New-Object System.DateTime(2001, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)

        function New-TestFolder {
            param([Parameter(Mandatory = $true)][string]$Name)
            $path = Join-Path $TestDrive ('{0}-{1}' -f $Name, [guid]::NewGuid().ToString('N').Substring(0, 8))
            [void][System.IO.Directory]::CreateDirectory($path)
            return $path
        }

        function Set-TestFile {
            param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)
            $folder = Split-Path -Parent $Path
            if (-not [System.IO.Directory]::Exists($folder)) { [void][System.IO.Directory]::CreateDirectory($folder) }
            [System.IO.File]::WriteAllText($Path, $Text)
        }

        function New-TestSource {
            # Udawany folder v4 (pendrive): pliki narzędzia i pliki, których nie kopiujemy.
            param([string]$Version = '4.0.2')
            $root = New-TestFolder 'pendrive'
            $files = [ordered]@{
                'STD.psd1'                   = "@{ ModuleVersion = '$Version' }"
                'STD.psm1'                   = '# moduł'
                'SmartToolforDeployment.ps1' = '# program'
                'Modules\Std.Core.ps1'       = 'core-aaaa'
                'GUI\Std.UI.ps1'             = 'okna'
                'GUI\Screens\Old.ps1'        = 'stary ekran'
                'Steps\Network.ps1'          = 'sieć'
                'Tests\X.Tests.ps1'          = 'test'
                'Tools\Y.ps1'                = 'narzędzie'
                'bootstrap.json'             = '{ "server": "https://srv/std/" }'
                'wifi\Firma.xml'             = '<WLANProfile/>'
                'offline\lizard\release.sdt' = '{}'
            }
            foreach ($key in $files.Keys) { Set-TestFile -Path (Join-Path $root $key) -Text $files[$key] }
            return $root
        }

        function Get-TestRelativeFiles {
            param([Parameter(Mandatory = $true)][string]$Root)
            $full = (Resolve-Path -LiteralPath $Root).ProviderPath.TrimEnd('\', '/')
            $list = @(Get-ChildItem -LiteralPath $full -Recurse -File | ForEach-Object { $_.FullName.Substring($full.Length + 1).Replace('/', '\') } | Sort-Object)
            return , $list
        }

        function Set-TestOldTimes {
            param([Parameter(Mandatory = $true)][string]$Root)
            Get-ChildItem -LiteralPath $Root -Recurse -File | ForEach-Object { $_.LastWriteTimeUtc = $script:OldDate }
        }

        function Get-TestAppVersions {
            $app = Get-StdPath -Kind App
            $names = [string[]]@(Get-ChildItem -LiteralPath $app -Directory | ForEach-Object { $_.Name })
            [System.Array]::Sort($names, [System.StringComparer]::Ordinal)
            return , $names
        }

        function ConvertFrom-TestCommandLine {
            # Podział wiersza poleceń według reguł Windows (CommandLineToArgvW / MSVCRT): tak powershell.exe
            # odczyta argumenty po ponownym uruchomieniu.
            param([Parameter(Mandatory = $true)][string]$Line)
            $result = New-Object 'System.Collections.Generic.List[string]'
            $current = New-Object System.Text.StringBuilder
            $inQuotes = $false
            $hasToken = $false
            $i = 0
            while ($i -lt $Line.Length) {
                $c = $Line[$i]
                if ($c -eq [char]'\') {
                    $count = 0
                    while ($i -lt $Line.Length -and $Line[$i] -eq [char]'\') { $count++; $i++ }
                    if ($i -lt $Line.Length -and $Line[$i] -eq [char]'"') {
                        [void]$current.Append([char]'\', [int][math]::Floor($count / 2))
                        if (($count % 2) -eq 1) {
                            [void]$current.Append([char]'"')
                            $i++
                        }
                    } else {
                        [void]$current.Append([char]'\', $count)
                    }
                    $hasToken = $true
                    continue
                }
                if ($c -eq [char]'"') {
                    $inQuotes = -not $inQuotes
                    $hasToken = $true
                    $i++
                    continue
                }
                if (-not $inQuotes -and [char]::IsWhiteSpace($c)) {
                    if ($hasToken) {
                        $result.Add($current.ToString())
                        [void]$current.Clear()
                        $hasToken = $false
                    }
                    $i++
                    continue
                }
                [void]$current.Append($c)
                $hasToken = $true
                $i++
            }
            if ($hasToken) { $result.Add($current.ToString()) }
            return , $result.ToArray()
        }

        function Get-TestBound {
            # Prawdziwy $PSBoundParameters (Dictionary, nie hashtable) z parametrami jak w programie.
            param([string]$Server, [string]$OriginRoot, [switch]$NoCopy, [switch]$PreviewMode, [string]$StdRelaunch)
            return $PSBoundParameters
        }

        function New-TestLink {
            # Dowiązanie do folderu: junction w Windows (mklink /J - bez uprawnień administratora), symlink w Linuksie.
            # $false, gdy nie udało się go utworzyć (test jest wtedy pomijany).
            param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Target)
            try {
                if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
                    $null = & cmd.exe /c mklink /J "$Path" "$Target" 2>&1
                } else {
                    $null = New-Item -ItemType SymbolicLink -Path $Path -Target $Target -ErrorAction Stop
                }
            } catch {
                return $false
            }
            if (-not (Test-Path -LiteralPath $Path)) { return $false }
            return (([System.IO.File]::GetAttributes($Path) -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        }

        function Remove-TestLink {
            # Usuwa samo dowiązanie, gdy test je zostawił (bez wchodzenia do folderu docelowego).
            param([Parameter(Mandatory = $true)][string]$Path)
            try {
                if (([System.IO.File]::GetAttributes($Path) -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { [System.IO.Directory]::Delete($Path, $false) }
            } catch { }
        }

        function New-TestOutside {
            # Folder poza App\ z plikiem, który nie może zniknąć.
            $outside = New-TestFolder 'poza'
            Set-TestFile -Path (Join-Path $outside 'cenny.txt') -Text 'nie usuwać'
            Set-TestFile -Path (Join-Path $outside 'Sub\drugi.txt') -Text 'też nie'
            return $outside
        }

        function Import-TestStartupFunction {
            # Funkcje pomocnicze programu (SmartToolforDeployment.ps1) bez uruchamiania go: definicje z drzewa składni.
            param([Parameter(Mandatory = $true)][string[]]$Name)
            $path = Join-Path $script:V4Root 'SmartToolforDeployment.ps1'
            $tokens = $null
            $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
            $found = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) | Where-Object { $Name -contains $_.Name })
            if ($found.Count -ne $Name.Count) { throw "Nie znaleziono wszystkich funkcji: $($Name -join ', ')" }
            return [scriptblock]::Create((($found | ForEach-Object { $_.Extent.Text }) -join [System.Environment]::NewLine))
        }

        function Start-TestMutexHolder {
            # Inny wątek (osobny runspace) trzyma mutex o podanej nazwie - jak drugi, działający proces STD.
            param([Parameter(Mandatory = $true)][string]$Name)
            $ready = New-Object System.Threading.ManualResetEventSlim($false)
            $stop = New-Object System.Threading.ManualResetEventSlim($false)
            $ps = [powershell]::Create()
            [void]$ps.AddScript({
                    param($MutexName, $Ready, $Stop)
                    $mutex = New-Object System.Threading.Mutex($false, $MutexName)
                    [void]$mutex.WaitOne()
                    $Ready.Set()
                    [void]$Stop.Wait(60000)
                    $mutex.ReleaseMutex()
                    $mutex.Dispose()
                }).AddArgument($Name).AddArgument($ready).AddArgument($stop)
            $handle = $ps.BeginInvoke()
            if (-not $ready.Wait(30000)) { throw 'Wątek testowy nie przejął mutexu.' }
            return @{ PowerShell = $ps; Handle = $handle; Stop = $stop }
        }

        function Stop-TestMutexHolder {
            param([Parameter(Mandatory = $true)]$Holder)
            $Holder.Stop.Set()
            [void]$Holder.PowerShell.EndInvoke($Holder.Handle)
            $Holder.PowerShell.Dispose()
        }

        function Test-TestMutexFree {
            # Czy inny wątek może teraz przejąć mutex (czyli nikt go nie trzyma).
            param([Parameter(Mandatory = $true)][string]$Name)
            $ps = [powershell]::Create()
            [void]$ps.AddScript({
                    param($MutexName)
                    $mutex = New-Object System.Threading.Mutex($false, $MutexName)
                    try {
                        $free = $mutex.WaitOne(0)
                    } catch [System.Threading.AbandonedMutexException] {
                        $free = $true
                    }
                    if ($free) { $mutex.ReleaseMutex() }
                    $mutex.Dispose()
                    $free
                }).AddArgument($Name)
            $result = @($ps.Invoke())
            $ps.Dispose()
            return [bool]$result[-1]
        }
    }

    BeforeEach {
        # Każdy test ma własny folder danych (App\ i prefs.json nie przechodzą między testami).
        $env:STD_DATA_ROOT = Join-Path $TestDrive ('data-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    }

    AfterAll {
        Remove-Item Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
        Remove-Variable -Name StdAppTestLine -Scope Global -ErrorAction SilentlyContinue
    }

    Context 'Get-StdAppFolder i Test-StdInAppFolder' {
        It 'Get-StdAppFolder zwraca App\(wersja modułu)' {
            Get-StdAppFolder | Should -Be (Join-Path (Get-StdPath -Kind App) (Get-StdVersion).ToString())
        }

        It 'Get-StdAppFolder -Version 4.1.0 zwraca App\4.1.0 i nie tworzy folderu' {
            $path = Get-StdAppFolder -Version '4.1.0'
            $path | Should -Be (Join-Path (Get-StdPath -Kind App) '4.1.0')
            Test-Path -LiteralPath $path | Should -BeFalse
        }

        It 'Test-StdInAppFolder: folder w App\ to $true, pendrive i podobna nazwa to $false' {
            $app = Get-StdPath -Kind App
            Test-StdInAppFolder -Path (Join-Path $app '4.0.0') | Should -BeTrue
            Test-StdInAppFolder -Path (Join-Path (Join-Path $app '4.0.0') 'GUI') | Should -BeTrue
            Test-StdInAppFolder -Path ($app + '2') | Should -BeFalse
            Test-StdInAppFolder -Path (New-TestSource) | Should -BeFalse
        }

        It 'Test-StdInAppFolder nie rozróżnia wielkości liter w Windows' -Skip:([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
            $app = Get-StdPath -Kind App
            Test-StdInAppFolder -Path ((Join-Path $app '4.0.0').ToUpperInvariant()) | Should -BeTrue
        }
    }

    Context 'Sync-StdAppFolder: kopia do App\(wersja)' {
        It 'kopiuje tylko pliki narzędzia do App\(wersja z STD.psd1) i zwraca ten folder' {
            $source = New-TestSource -Version '4.0.2'
            $target = Sync-StdAppFolder -Source $source
            $target | Should -Be (Join-Path (Get-StdPath -Kind App) '4.0.2')
            $files = Get-TestRelativeFiles -Root $target
            ($files -join '|') | Should -Be ((@('GUI\Screens\Old.ps1', 'GUI\Std.UI.ps1', 'Modules\Std.Core.ps1', 'SmartToolforDeployment.ps1', 'STD.psd1', 'STD.psm1', 'Steps\Network.ps1') | Sort-Object) -join '|')
            [System.IO.File]::ReadAllText((Join-Path $target 'Modules\Std.Core.ps1')) | Should -Be 'core-aaaa'
            Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
        }

        It 'nie kopiuje testów, narzędzi, bootstrap.json, wifi\ ani offline\' {
            $target = Sync-StdAppFolder -Source (New-TestSource)
            foreach ($name in @('Tests', 'Tools', 'bootstrap.json', 'wifi', 'offline')) {
                Test-Path -LiteralPath (Join-Path $target $name) | Should -BeFalse
            }
        }

        It 'pomija kopię, gdy pliki są identyczne (SHA-256)' {
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            Set-TestOldTimes -Root $target
            $again = Sync-StdAppFolder -Source $source
            $again | Should -Be $target
            @(Get-ChildItem -LiteralPath $target -Recurse -File | Where-Object { $_.LastWriteTimeUtc -ne $script:OldDate }) | Should -HaveCount 0
        }

        It 'zmieniony plik (ten sam rozmiar) oznacza nieaktualną kopię i jest kopiowany ponownie' {
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            Set-TestOldTimes -Root $target
            Set-TestFile -Path (Join-Path $source 'Modules\Std.Core.ps1') -Text 'core-bbbb'
            Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeFalse

            [void](Sync-StdAppFolder -Source $source)
            [System.IO.File]::ReadAllText((Join-Path $target 'Modules\Std.Core.ps1')) | Should -Be 'core-bbbb'
            # Niezmienione pliki nie są kopiowane drugi raz.
            (Get-Item -LiteralPath (Join-Path $target 'STD.psm1')).LastWriteTimeUtc | Should -Be $script:OldDate
            Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
        }

        It 'dopisuje nowe pliki, usuwa pliki, których w źródle już nie ma, i puste foldery' {
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            Set-TestFile -Path (Join-Path $source 'Steps\Power.ps1') -Text 'zasilanie'
            Remove-Item -LiteralPath (Join-Path $source 'GUI\Screens\Old.ps1')

            [void](Sync-StdAppFolder -Source $source)
            Test-Path -LiteralPath (Join-Path $target 'Steps\Power.ps1') | Should -BeTrue
            Test-Path -LiteralPath (Join-Path $target 'GUI\Screens\Old.ps1') | Should -BeFalse
            Test-Path -LiteralPath (Join-Path $target 'GUI\Screens') | Should -BeFalse
            Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
        }

        It 'nadpisuje plik z atrybutem tylko do odczytu' {
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            $file = Join-Path $target 'Modules\Std.Core.ps1'
            Set-ItemProperty -LiteralPath $file -Name IsReadOnly -Value $true
            Set-TestFile -Path (Join-Path $source 'Modules\Std.Core.ps1') -Text 'nowa wersja pliku'
            [void](Sync-StdAppFolder -Source $source)
            [System.IO.File]::ReadAllText($file) | Should -Be 'nowa wersja pliku'
        }

        It 'nie kopiuje, gdy źródło to już folder App\(wersja)' {
            $target = Sync-StdAppFolder -Source (New-TestSource)
            Set-TestOldTimes -Root $target
            Sync-StdAppFolder -Source $target | Should -Be $target
            @(Get-ChildItem -LiteralPath $target -Recurse -File | Where-Object { $_.LastWriteTimeUtc -ne $script:OldDate }) | Should -HaveCount 0
        }

        It '-Version nadpisuje wersję z STD.psd1' {
            $target = Sync-StdAppFolder -Source (New-TestSource -Version '4.0.2') -Version '4.3.0'
            $target | Should -Be (Join-Path (Get-StdPath -Kind App) '4.3.0')
            Test-Path -LiteralPath (Join-Path $target 'STD.psd1') | Should -BeTrue
        }

        It 'folder bez STD.psd1 to błąd [STD-APP-SOURCE]' {
            $empty = New-TestFolder 'pusty'
            { Sync-StdAppFolder -Source $empty } | Should -Throw -ExpectedMessage '*`[STD-APP-SOURCE`]*'
            { Test-StdAppCopyCurrent -Source $empty -Target $empty } | Should -Throw -ExpectedMessage '*`[STD-APP-SOURCE`]*'
        }

        It 'folder kopii wewnątrz źródła to błąd [STD-APP-COPY]' {
            $source = New-TestSource
            $env:STD_DATA_ROOT = Join-Path $source 'dane'
            { Sync-StdAppFolder -Source $source } | Should -Throw -ExpectedMessage '*`[STD-APP-COPY`]*'
        }

        It 'Test-StdAppCopyCurrent: brak folderu docelowego to $false' {
            Test-StdAppCopyCurrent -Source (New-TestSource) -Target (Join-Path $TestDrive 'nie-ma-takiego') | Should -BeFalse
        }
    }

    Context 'Sync-StdAppFolder: stare wersje' {
        BeforeEach {
            $script:App = Get-StdPath -Kind App
            foreach ($name in @('3.9.0', '4.0.9', '4.0.10', 'Notatki')) {
                Set-TestFile -Path (Join-Path (Join-Path $script:App $name) 'plik.txt') -Text $name
            }
        }

        It 'zostawia dwie najnowsze według [version] (4.0.11, 4.0.10), nie według tekstu' {
            [void](Sync-StdAppFolder -Source (New-TestSource -Version '4.0.11'))
            # Sortowanie tekstowe uznałoby 4.0.9 za nowszą niż 4.0.10 i 4.0.11.
            (Get-TestAppVersions) -join '|' | Should -Be '4.0.10|4.0.11|Notatki'
        }

        It 'nie rusza folderów, których nazwa nie jest wersją' {
            [void](Sync-StdAppFolder -Source (New-TestSource -Version '5.0.0'))
            Test-Path -LiteralPath (Join-Path (Join-Path $script:App 'Notatki') 'plik.txt') | Should -BeTrue
        }

        It 'zawsze zostawia wersję, z której startuje, także starszą niż dwie najnowsze' {
            [void](Sync-StdAppFolder -Source (New-TestSource -Version '4.0.2'))
            (Get-TestAppVersions) -join '|' | Should -Be '4.0.10|4.0.2|4.0.9|Notatki'
        }

        It 'usuwa starą wersję z plikiem tylko do odczytu' {
            Set-ItemProperty -LiteralPath (Join-Path (Join-Path $script:App '3.9.0') 'plik.txt') -Name IsReadOnly -Value $true
            [void](Sync-StdAppFolder -Source (New-TestSource -Version '4.1.0'))
            Test-Path -LiteralPath (Join-Path $script:App '3.9.0') | Should -BeFalse
        }

        It '-Keep 3 zostawia trzy najnowsze' {
            [void](Sync-StdAppFolder -Source (New-TestSource -Version '4.1.0') -Keep 3)
            (Get-TestAppVersions) -join '|' | Should -Be '4.0.10|4.0.9|4.1.0|Notatki'
        }
    }

    Context 'Sync-StdAppFolder: dowiązania (junction / symlink) i właściciel kopii' {
        It 'Test-StdAppCopyCurrent odrzuca dowiązanie w kopii błędem [STD-APP-COPY] i niczego nie usuwa' {
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            $outside = New-TestOutside
            $link = Join-Path $target 'Modules\Obcy'
            if (-not (New-TestLink -Path $link -Target $outside)) { Set-ItResult -Skipped -Because 'nie udało się utworzyć dowiązania'; return }
            try {
                { Test-StdAppCopyCurrent -Source $source -Target $target } | Should -Throw -ExpectedMessage '*`[STD-APP-COPY`]*'
                Test-Path -LiteralPath (Join-Path $outside 'cenny.txt') | Should -BeTrue
            } finally {
                Remove-TestLink -Path $link
            }
        }

        It 'kopia z dowiązaniem w środku jest kopiowana od nowa, a pliki, na które wskazywało, zostają' {
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            $outside = New-TestOutside
            $link = Join-Path $target 'Modules\Obcy'
            if (-not (New-TestLink -Path $link -Target $outside)) { Set-ItResult -Skipped -Because 'nie udało się utworzyć dowiązania'; return }
            try {
                Sync-StdAppFolder -Source $source | Should -Be $target
                Test-Path -LiteralPath $link | Should -BeFalse
                Test-Path -LiteralPath (Join-Path $outside 'cenny.txt') | Should -BeTrue
                Test-Path -LiteralPath (Join-Path $outside 'Sub\drugi.txt') | Should -BeTrue
                Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
            } finally {
                Remove-TestLink -Path $link
            }
        }

        It 'kopia, która sama jest dowiązaniem, jest zastępowana zwykłym folderem' {
            $source = New-TestSource -Version '4.0.2'
            $outside = New-TestOutside
            $target = Join-Path (Get-StdPath -Kind App) '4.0.2'
            if (-not (New-TestLink -Path $target -Target $outside)) { Set-ItResult -Skipped -Because 'nie udało się utworzyć dowiązania'; return }
            try {
                Sync-StdAppFolder -Source $source | Should -Be $target
                ([System.IO.File]::GetAttributes($target) -band [System.IO.FileAttributes]::ReparsePoint) | Should -Be 0
                Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
                @(Get-ChildItem -LiteralPath $outside -Recurse -File).Count | Should -Be 2
            } finally {
                Remove-TestLink -Path $target
            }
        }

        It 'stara wersja będąca dowiązaniem albo z dowiązaniem w środku znika bez plików, na które wskazuje' {
            $app = Get-StdPath -Kind App
            $outside = New-TestOutside
            Set-TestFile -Path (Join-Path $app '3.0.0\plik.txt') -Text 'stara'
            $inner = Join-Path $app '3.0.0\Modules'
            $whole = Join-Path $app '3.1.0'
            if (-not (New-TestLink -Path $inner -Target $outside) -or -not (New-TestLink -Path $whole -Target $outside)) {
                Remove-TestLink -Path $inner
                Set-ItResult -Skipped -Because 'nie udało się utworzyć dowiązania'
                return
            }
            try {
                [void](Sync-StdAppFolder -Source (New-TestSource -Version '4.1.0') -Keep 1)
                Test-Path -LiteralPath (Join-Path $app '3.0.0') | Should -BeFalse
                Test-Path -LiteralPath $whole | Should -BeFalse
                Test-Path -LiteralPath (Join-Path $outside 'cenny.txt') | Should -BeTrue
                Test-Path -LiteralPath (Join-Path $outside 'Sub\drugi.txt') | Should -BeTrue
            } finally {
                Remove-TestLink -Path $inner
                Remove-TestLink -Path $whole
            }
        }

        It 'folder App\ będący dowiązaniem jest zastępowany zwykłym folderem przed kopią' {
            $outside = New-TestOutside
            $app = Join-Path $env:STD_DATA_ROOT 'App'
            [void][System.IO.Directory]::CreateDirectory($env:STD_DATA_ROOT)
            if (-not (New-TestLink -Path $app -Target $outside)) { Set-ItResult -Skipped -Because 'nie udało się utworzyć dowiązania'; return }
            try {
                $target = Sync-StdAppFolder -Source (New-TestSource -Version '4.0.2')
                $target | Should -Be (Join-Path $app '4.0.2')
                ([System.IO.File]::GetAttributes($app) -band [System.IO.FileAttributes]::ReparsePoint) | Should -Be 0
                Test-Path -LiteralPath (Join-Path $outside '4.0.2') | Should -BeFalse
                @(Get-ChildItem -LiteralPath $outside -Recurse -File).Count | Should -Be 2
            } finally {
                Remove-TestLink -Path $app
            }
        }

        It 'kopia, której właścicielem nie są Administratorzy ani SYSTEM, jest kopiowana od nowa (administrator w Windows)' {
            Mock -ModuleName STD -CommandName Test-CoreDataProtectionNeeded -MockWith { $true }
            Mock -ModuleName STD -CommandName Protect-StdDataFolder -MockWith { }
            Mock -ModuleName STD -CommandName Test-AppTrustedOwner -MockWith { $false }
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            # Plik spoza plików narzędzia - zwykła kopia lustrzana by go nie usunęła.
            Set-TestFile -Path (Join-Path $target 'podrzucony.txt') -Text 'obcy'
            [void](Sync-StdAppFolder -Source $source)
            Test-Path -LiteralPath (Join-Path $target 'podrzucony.txt') | Should -BeFalse
            Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
            Should -Invoke -ModuleName STD -CommandName Protect-StdDataFolder -ParameterFilter { $Path -eq (Get-StdPath -Kind App) }
            # Kopia dostaje właściciela Administratorzy (także bez UAC), więc następny start jej nie odrzuci.
            Should -Invoke -ModuleName STD -CommandName Protect-StdDataFolder -ParameterFilter { $Path -eq $target }
        }

        It 'kopia Administratorów z uprawnieniami tylko dla Administratorów i SYSTEM zostaje (bez kopiowania od nowa)' {
            Mock -ModuleName STD -CommandName Test-CoreDataProtectionNeeded -MockWith { $true }
            Mock -ModuleName STD -CommandName Protect-StdDataFolder -MockWith { }
            Mock -ModuleName STD -CommandName Test-AppTrustedOwner -MockWith { $true }
            Mock -ModuleName STD -CommandName Test-CoreDataFolderProtected -MockWith { $true }
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            Set-TestFile -Path (Join-Path $target 'podrzucony.txt') -Text 'obcy'
            [void](Sync-StdAppFolder -Source $source)
            Test-Path -LiteralPath (Join-Path $target 'podrzucony.txt') | Should -BeTrue
        }

        It 'kopia sprzed zabezpieczenia uprawnień (dziedziczonych z ProgramData) jest kopiowana od nowa' {
            Mock -ModuleName STD -CommandName Test-CoreDataProtectionNeeded -MockWith { $true }
            Mock -ModuleName STD -CommandName Protect-StdDataFolder -MockWith { }
            Mock -ModuleName STD -CommandName Test-AppTrustedOwner -MockWith { $true }
            Mock -ModuleName STD -CommandName Test-CoreDataFolderProtected -MockWith { $false }
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            Set-TestFile -Path (Join-Path $target 'podrzucony.txt') -Text 'obcy'
            [void](Sync-StdAppFolder -Source $source)
            Test-Path -LiteralPath (Join-Path $target 'podrzucony.txt') | Should -BeFalse
            Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
        }

        It 'bez uprawnień administratora (testy, STD_DATA_ROOT) właściciel i uprawnienia nie są sprawdzane' {
            Mock -ModuleName STD -CommandName Protect-StdDataFolder -MockWith { }
            Mock -ModuleName STD -CommandName Test-AppTrustedOwner -MockWith { $false }
            Mock -ModuleName STD -CommandName Test-CoreDataFolderProtected -MockWith { $false }
            $source = New-TestSource
            $target = Sync-StdAppFolder -Source $source
            Set-TestFile -Path (Join-Path $target 'podrzucony.txt') -Text 'obcy'
            [void](Sync-StdAppFolder -Source $source)
            Test-Path -LiteralPath (Join-Path $target 'podrzucony.txt') | Should -BeTrue
            Should -Invoke -ModuleName STD -CommandName Protect-StdDataFolder -Times 0 -Exactly
            Should -Invoke -ModuleName STD -CommandName Test-AppTrustedOwner -Times 0 -Exactly
            Should -Invoke -ModuleName STD -CommandName Test-CoreDataFolderProtected -Times 0 -Exactly
        }
    }

    Context 'Enter-StdAppInstance: jedna instancja STD' {
        BeforeAll {
            # Własna nazwa mutexu: test nie może trafić w prawdziwy, działający STD.
            $script:TestMutexName = 'Global\STD4-Test-' + [guid]::NewGuid().ToString('N')
            $script:RealMutexName = & (Get-Module STD) { $script:StdAppMutexName }
            & (Get-Module STD) { param($Name) $script:StdAppMutexName = $Name } $script:TestMutexName
        }

        AfterAll {
            Exit-StdAppInstance
            & (Get-Module STD) { param($Name) $script:StdAppMutexName = $Name } $script:RealMutexName
        }

        AfterEach {
            Exit-StdAppInstance
        }

        It 'daje blokadę; drugie wejście w tym samym procesie też; Exit-StdAppInstance ją zwalnia' {
            Enter-StdAppInstance | Should -BeTrue
            Enter-StdAppInstance | Should -BeTrue
            Test-TestMutexFree -Name $script:TestMutexName | Should -BeFalse
            Exit-StdAppInstance
            Test-TestMutexFree -Name $script:TestMutexName | Should -BeTrue
            { Exit-StdAppInstance } | Should -Not -Throw
        }

        It 'gdy blokadę trzyma inny proces STD: $false (także po czekaniu -WaitSec)' {
            $holder = Start-TestMutexHolder -Name $script:TestMutexName
            try {
                Enter-StdAppInstance | Should -BeFalse
                Enter-StdAppInstance -WaitSec 1 | Should -BeFalse
            } finally {
                Stop-TestMutexHolder -Holder $holder
            }
            Enter-StdAppInstance | Should -BeTrue
        }

        It 'gdy działa inny proces STD, Sync-StdAppFolder kopiuje, ale nie usuwa starych wersji' {
            $app = Get-StdPath -Kind App
            foreach ($name in @('3.9.0', '4.0.9', '4.0.10')) { Set-TestFile -Path (Join-Path (Join-Path $app $name) 'plik.txt') -Text $name }
            $source = New-TestSource -Version '4.0.11'
            $holder = Start-TestMutexHolder -Name $script:TestMutexName
            try {
                $target = Sync-StdAppFolder -Source $source
                Test-StdAppCopyCurrent -Source $source -Target $target | Should -BeTrue
                (Get-TestAppVersions) -join '|' | Should -Be '3.9.0|4.0.10|4.0.11|4.0.9'
            } finally {
                Stop-TestMutexHolder -Holder $holder
            }
            [void](Sync-StdAppFolder -Source $source)
            (Get-TestAppVersions) -join '|' | Should -Be '4.0.10|4.0.11'
        }

        It 'blokada trzymana przez ten proces nie wstrzymuje usuwania starych wersji' {
            Enter-StdAppInstance | Should -BeTrue
            $app = Get-StdPath -Kind App
            foreach ($name in @('3.9.0', '4.0.9', '4.0.10')) { Set-TestFile -Path (Join-Path (Join-Path $app $name) 'plik.txt') -Text $name }
            [void](Sync-StdAppFolder -Source (New-TestSource -Version '4.0.11'))
            (Get-TestAppVersions) -join '|' | Should -Be '4.0.10|4.0.11'
            # Sprawdzenie w Sync-StdAppFolder nie zwolniło blokady tego procesu.
            Test-TestMutexFree -Name $script:TestMutexName | Should -BeFalse
        }
    }

    Context 'Get-StdNetworkPath: start z dysku sieciowego' {
        It 'ścieżka UNC zostaje bez zmian' {
            $result = Get-StdNetworkPath -Path '\\srv-files\STD$\std\SmartToolforDeployment.ps1'
            $result.IsNetwork | Should -BeTrue
            $result.UncPath | Should -Be '\\srv-files\STD$\std\SmartToolforDeployment.ps1'
        }

        It 'ścieżka \\?\UNC\ to udział sieciowy' {
            $result = Get-StdNetworkPath -Path '\\?\UNC\srv\udział\STD'
            $result.IsNetwork | Should -BeTrue
            $result.UncPath | Should -Be '\\srv\udział\STD'
        }

        It 'litera dysku zmapowanego zamienia się na \\serwer\udział' {
            Mock -ModuleName STD -CommandName Get-AppDriveType -MockWith { 'Network' }
            Mock -ModuleName STD -CommandName Get-AppDriveRemotePath -MockWith { '\\srv\Wdrożenia STD\' }
            $result = Get-StdNetworkPath -Path 'z:\STD 4\SmartToolforDeployment.ps1'
            $result.IsNetwork | Should -BeTrue
            $result.UncPath | Should -Be '\\srv\Wdrożenia STD\STD 4\SmartToolforDeployment.ps1'
            (Get-StdNetworkPath -Path 'Z:\').UncPath | Should -Be '\\srv\Wdrożenia STD\'
            Should -Invoke -ModuleName STD -CommandName Get-AppDriveRemotePath -ParameterFilter { $Drive -eq 'Z:' }
        }

        It 'dysk sieciowy bez znanego udziału: IsNetwork = $true, UncPath = $null' {
            Mock -ModuleName STD -CommandName Get-AppDriveType -MockWith { 'Network' }
            Mock -ModuleName STD -CommandName Get-AppDriveRemotePath -MockWith { $null }
            $result = Get-StdNetworkPath -Path 'Z:\STD\SmartToolforDeployment.ps1'
            $result.IsNetwork | Should -BeTrue
            $result.UncPath | Should -BeNullOrEmpty
        }

        It 'dysk lokalny, pendrive i ścieżka bez litery to nie sieć' {
            Mock -ModuleName STD -CommandName Get-AppDriveType -MockWith { 'Removable' }
            (Get-StdNetworkPath -Path 'E:\STD\SmartToolforDeployment.ps1').IsNetwork | Should -BeFalse
            (Get-StdNetworkPath -Path '/media/usb/STD').IsNetwork | Should -BeFalse
            (Get-StdNetworkPath -Path $TestDrive).IsNetwork | Should -BeFalse
        }
    }

    Context 'Program: plik startowy, dziennik okien, ponowne uruchomienie' {
        BeforeAll {
            $script:StartupFunctions = Import-TestStartupFunction -Name @('Protect-StartupText', 'Get-StartupBootstrap', 'Get-StartupEarlyExitCode')
        }

        It 'Get-StartupBootstrap: błędny bootstrap.json na pendrive''ie -> adres zapisany na komputerze i jeden komunikat' {
            . $script:StartupFunctions
            $messages = New-Object 'System.Collections.Generic.List[string]'
            function Show-StdMessage { param($Message, $Title, $Image) $messages.Add([string]$Message) }
            function Show-StartupMessage { param($Message, $Title, $Icon) $messages.Add([string]$Message) }
            Save-StdBootstrapLocal -Bootstrap @{ server = 'https://zapisany.firma.pl/std/' }
            $usb = New-TestFolder 'usb'
            Set-TestFile -Path (Join-Path $usb 'bootstrap.json') -Text '{ "server": "http://bez-https/std/" }'
            $result = Get-StartupBootstrap -ServerUrl $null -Root $usb
            $result.server | Should -Be 'https://zapisany.firma.pl/std/'
            $messages | Should -HaveCount 1
            $messages[0] | Should -Match 'https://zapisany\.firma\.pl/std/'
        }

        It 'Get-StartupBootstrap: błędny bootstrap.json bez zapisanego adresu -> $null i jeden komunikat' {
            . $script:StartupFunctions
            $messages = New-Object 'System.Collections.Generic.List[string]'
            function Show-StdMessage { param($Message, $Title, $Image) $messages.Add([string]$Message) }
            function Show-StartupMessage { param($Message, $Title, $Icon) $messages.Add([string]$Message) }
            $usb = New-TestFolder 'usb'
            Set-TestFile -Path (Join-Path $usb 'bootstrap.json') -Text '{ to nie jest JSON'
            Get-StartupBootstrap -ServerUrl $null -Root $usb | Should -BeNullOrEmpty
            $messages | Should -HaveCount 1
            $messages[0] | Should -Match 'oknie logowania'
        }

        It 'Get-StartupBootstrap: zły -Server -> poprawny bootstrap.json z pendrive''a' {
            . $script:StartupFunctions
            $messages = New-Object 'System.Collections.Generic.List[string]'
            function Show-StdMessage { param($Message, $Title, $Image) $messages.Add([string]$Message) }
            function Show-StartupMessage { param($Message, $Title, $Icon) $messages.Add([string]$Message) }
            $usb = New-TestFolder 'usb'
            Set-TestFile -Path (Join-Path $usb 'bootstrap.json') -Text '{ "server": "https://pendrive.firma.pl/std/" }'
            (Get-StartupBootstrap -ServerUrl 'ftp://zly' -Root $usb).server | Should -Be 'https://pendrive.firma.pl/std/'
            $messages | Should -HaveCount 1
        }

        It 'Get-StartupBootstrap: poprawny bootstrap.json -> bez komunikatu' {
            . $script:StartupFunctions
            $messages = New-Object 'System.Collections.Generic.List[string]'
            function Show-StdMessage { param($Message, $Title, $Image) $messages.Add([string]$Message) }
            function Show-StartupMessage { param($Message, $Title, $Icon) $messages.Add([string]$Message) }
            $usb = New-TestFolder 'usb'
            Set-TestFile -Path (Join-Path $usb 'bootstrap.json') -Text '{ "server": "https://pendrive.firma.pl/std/" }'
            (Get-StartupBootstrap -ServerUrl $null -Root $usb).server | Should -Be 'https://pendrive.firma.pl/std/'
            $messages | Should -HaveCount 0
        }

        It 'Get-StartupEarlyExitCode: kod procesu, który zakończył się od razu; $null, gdy proces nadal działa' {
            . $script:StartupFunctions
            $exe = (Get-Process -Id $PID).Path
            $quick = [System.Diagnostics.Process]::Start($exe, '-NoProfile -NonInteractive -Command exit 3')
            Get-StartupEarlyExitCode -Process $quick -Milliseconds 60000 | Should -Be 3
            $slow = [System.Diagnostics.Process]::Start($exe, '-NoProfile -NonInteractive -Command Start-Sleep -Seconds 60')
            $slowId = $slow.Id
            try {
                Get-StartupEarlyExitCode -Process $slow -Milliseconds 200 | Should -BeNullOrEmpty
            } finally {
                Stop-Process -Id $slowId -Force -ErrorAction SilentlyContinue
            }
        }

        It 'Start-StdElevated -PassThru zwraca obiekt procesu, bez obiektu $true, a przy odmowie UAC $false' {
            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { New-Object System.Diagnostics.Process }
            $result = Start-StdElevated -ScriptPath 'C:\STD\SmartToolforDeployment.ps1' -PassThru
            ($result -is [System.Diagnostics.Process]) | Should -BeTrue
            $result.Dispose()
            Should -Invoke -ModuleName STD -CommandName Start-CoreElevatedProcess -Times 1 -Exactly -ParameterFilter { $PassThru }
            # Bez -PassThru wynik to nadal [bool].
            Start-StdElevated -ScriptPath 'C:\STD\SmartToolforDeployment.ps1' | Should -BeExactly $true

            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { }
            Start-StdElevated -ScriptPath 'C:\STD\SmartToolforDeployment.ps1' -PassThru | Should -BeExactly $true
            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { throw (New-Object System.ComponentModel.Win32Exception(1223)) }
            Start-StdElevated -ScriptPath 'C:\STD\SmartToolforDeployment.ps1' -PassThru | Should -BeExactly $false
        }

        It 'program: ścieżki UNC przy starcie z sieci, sprawdzenie szybkiego końca, blokada przed kopią, dziennik okien przed pierwszym oknem' {
            $text = [System.IO.File]::ReadAllText((Join-Path $script:V4Root 'SmartToolforDeployment.ps1'))
            $text | Should -Match 'Get-StdNetworkPath -Path \$stdAppScriptPath'
            $text | Should -Match "Invoke-StartupRelaunch -ScriptPath \`$stdAppRelaunchPath -Stage 'admin'"
            $text | Should -Match 'Get-StartupEarlyExitCode -Process'
            $text | Should -Match 'Protect-StdDataFolder -Path \(Get-StdDataRoot\)'
            $text | Should -Match "Register-StdFileLogSink -Name 'StdUi'"
            $text | Should -Match '\$script:StdUiLogPath = \$stdAppUiLog'
            $text.IndexOf('Enter-StdAppInstance -WaitSec') | Should -BeGreaterThan 0
            $text.IndexOf('Enter-StdAppInstance -WaitSec') | Should -BeLessThan $text.IndexOf('Sync-StdAppFolder -Source')
            $text.IndexOf("Register-StdFileLogSink -Name 'StdUi'") | Should -BeLessThan $text.IndexOf('Show-StdConnectWindow -OriginRoot')
            # Okno główne dostaje plik startowy z logowania, bez ponownego czytania bootstrap.json.
            $text | Should -Match "\`$stdAppLogin\['Bootstrap'\]"
            ([regex]::Matches($text, 'Get-StartupBootstrap -ServerUrl')).Count | Should -Be 1
        }
    }

    Context 'Get-StdRelaunchArgument' {
        It 'przełączniki: $true to sama nazwa, $false jest pomijany' {
            $bound = Get-TestBound -NoCopy -PreviewMode:$false
            $result = Get-StdRelaunchArgument -BoundParameters $bound
            ($result -is [string[]]) | Should -BeTrue
            ($result -join '|') | Should -Be '-NoCopy'
        }

        It 'przyjmuje prawdziwy $PSBoundParameters i zwraca surowe (niezacytowane) wartości ze spacjami i cudzysłowami' {
            $bound = Get-TestBound -Server 'https://srv/std/' -OriginRoot 'E:\Pendrive STD\' -NoCopy
            $result = Get-StdRelaunchArgument -BoundParameters $bound -Add @{ Note = 'a "b" c' }
            $result | Should -HaveCount 7
            ($result -join '|') | Should -Be '-NoCopy|-Note|a "b" c|-OriginRoot|E:\Pendrive STD\|-Server|https://srv/std/'
        }

        It 'dopisuje -OriginRoot i marker ponownego uruchomienia z -Add' {
            $bound = Get-TestBound -Server 'https://srv/std/'
            $result = Get-StdRelaunchArgument -BoundParameters $bound -Add @{ OriginRoot = 'E:\'; StdRelaunch = 'admin,copy' }
            ($result -join '|') | Should -Be '-OriginRoot|E:\|-Server|https://srv/std/|-StdRelaunch|admin,copy'
        }

        It '-Add nadpisuje parametr, a $null / $false w -Add go usuwa' {
            $bound = Get-TestBound -Server 'https://stary/std/' -OriginRoot 'D:\' -NoCopy
            $result = Get-StdRelaunchArgument -BoundParameters $bound -Add @{ OriginRoot = 'E:\USB'; Server = $null; NoCopy = $false }
            ($result -join '|') | Should -Be '-OriginRoot|E:\USB'
        }

        It 'pomija $null i pusty tekst, łączy tablicę przecinkami, liczby bez ustawień regionalnych' {
            $result = Get-StdRelaunchArgument -BoundParameters @{ A = $null; B = ''; C = $true; D = $false; E = @('x', 'y'); F = 1.5 }
            ($result -join '|') | Should -Be '-C|-E|x,y|-F|1.5'
        }

        It 'puste parametry to pusta tablica' {
            $result = Get-StdRelaunchArgument -BoundParameters @{}
            ($result -is [string[]]) | Should -BeTrue
            $result | Should -HaveCount 0
        }

        It 'hasło (SecureString / PSCredential) to błąd [STD-APP-ARGS]' {
            $secure = New-Object System.Security.SecureString
            foreach ($ch in 'tajne'.ToCharArray()) { $secure.AppendChar($ch) }
            $credential = New-Object System.Management.Automation.PSCredential('admin', $secure)
            { Get-StdRelaunchArgument -BoundParameters @{ Password = $secure } } | Should -Throw -ExpectedMessage '*`[STD-APP-ARGS`]*'
            { Get-StdRelaunchArgument -BoundParameters @{ Credential = $credential } } | Should -Throw -ExpectedMessage '*`[STD-APP-ARGS`]*'
        }

        It 'nieprawidłowa nazwa parametru to błąd [STD-APP-ARGS]' {
            { Get-StdRelaunchArgument -BoundParameters @{ 'zła nazwa' = 'x' } } | Should -Throw -ExpectedMessage '*`[STD-APP-ARGS`]*'
        }

        It 'Start-StdElevated cytuje wynik poprawnie: spacje, cudzysłowy, ukośnik na końcu, przełączniki' {
            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { $global:StdAppTestLine = $ArgumentLine }
            $bound = Get-TestBound -Server 'https://srv/std/' -OriginRoot 'E:\Pendrive STD\' -NoCopy
            $arguments = Get-StdRelaunchArgument -BoundParameters $bound -Add @{ Note = 'a "b" c'; StdRelaunch = 'admin' }
            Start-StdElevated -ScriptPath 'C:\Program Files\STD\SmartToolforDeployment.ps1' -Arguments $arguments | Should -BeTrue
            $global:StdAppTestLine | Should -Be ('-NoProfile -ExecutionPolicy Bypass -STA -File "C:\Program Files\STD\SmartToolforDeployment.ps1" ' +
                '-NoCopy -Note "a \"b\" c" -OriginRoot "E:\Pendrive STD\\" -Server https://srv/std/ -StdRelaunch admin')
        }

        It 'po podziale wiersza poleceń (reguły Windows) powershell.exe dostanie dokładnie te same argumenty' {
            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { $global:StdAppTestLine = $ArgumentLine }
            $add = [ordered]@{
                OriginRoot  = 'E:\Wdrożenie STD\'
                Server      = 'https://srv.firma.pl/std/?a=1&b=2'
                Note        = 'cudzysłów " w środku i \" ukośnik przed nim'
                Share       = '\\srv-files\STD$\std\'
                NoCopy      = [switch]$true
                StdRelaunch = 'admin,copy'
            }
            $arguments = Get-StdRelaunchArgument -BoundParameters @{} -Add $add
            [void](Start-StdElevated -ScriptPath 'C:\Users\Jan Kowalski\STD 4\SmartToolforDeployment.ps1' -Arguments $arguments)
            $parsed = ConvertFrom-TestCommandLine -Line $global:StdAppTestLine
            $parsed[4] | Should -Be '-File'
            $parsed[5] | Should -Be 'C:\Users\Jan Kowalski\STD 4\SmartToolforDeployment.ps1'
            $rest = @($parsed | Select-Object -Skip 6)
            ($rest -join '|') | Should -Be ($arguments -join '|')
            $rest | Should -Contain 'E:\Wdrożenie STD\'
            $rest | Should -Contain 'cudzysłów " w środku i \" ukośnik przed nim'
        }
    }

    Context 'Test-StdNetworkAvailable' {
        It 'zwraca [bool] bez błędu' {
            Test-StdNetworkAvailable | Should -BeOfType [bool]
        }
    }

    Context 'Pliki startowe' {
        It 'Std.App.ps1 nie używa WPF' {
            $text = [System.IO.File]::ReadAllText((Join-Path $script:V4Root 'Modules\Std.App.ps1'))
            $text | Should -Not -Match 'PresentationFramework|System\.Windows'
        }

        It 'SmartToolforDeployment.ps1 parsuje się bez błędów i ma parametry z SPEC 8 oraz trap' {
            $path = Join-Path $script:V4Root 'SmartToolforDeployment.ps1'
            $tokens = $null
            $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
            $errors | Should -HaveCount 0
            $names = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
            foreach ($name in @('Server', 'OriginRoot', 'NoCopy', 'PreviewMode', 'StdRelaunch')) { $names | Should -Contain $name }
            $traps = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.TrapStatementAst] }, $false))
            $traps.Count | Should -BeGreaterThan 0
        }

        It 'SmartToolforDeployment.ps1 woła okna z nazwanymi parametrami z SPEC 7.3' {
            $text = [System.IO.File]::ReadAllText((Join-Path $script:V4Root 'SmartToolforDeployment.ps1'))
            $text | Should -Match 'Show-StdConnectWindow -OriginRoot '
            $text | Should -Match 'Show-StdLoginWindow -Bootstrap \S+ -OriginRoot '
            $text | Should -Match 'Show-StdMainWindow -Session \S+ -Index \S+ -IndexSource \S+'
        }

        It 'Start-STD.cmd ma końce linii CRLF i uruchamia program w trybie STA z argumentami' {
            $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:V4Root 'Start-STD.cmd'))
            $text = [System.Text.Encoding]::ASCII.GetString($bytes)
            $text | Should -Match "`r`n"
            ($text -replace "`r`n", '') | Should -Not -Match "`n"
            $text | Should -Match '-NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0SmartToolforDeployment\.ps1" %\*'
        }

        It 'tryb -PreviewMode wczytuje moduł i okna bez pokazywania okien' -Skip:([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
            $script = Join-Path $script:V4Root 'SmartToolforDeployment.ps1'
            $command = "& { . '$($script.Replace("'", "''"))' -PreviewMode; if (`$StdPreview -and (Get-Command -Name Initialize-StdTheme -ErrorAction SilentlyContinue)) { 'PREVIEW-OK' } }"
            $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $output = & $exe -NoProfile -ExecutionPolicy Bypass -STA -Command $command 2>&1
            $LASTEXITCODE | Should -Be 0
            ($output | Out-String) | Should -Match 'PREVIEW-OK'
        }
    }
}
