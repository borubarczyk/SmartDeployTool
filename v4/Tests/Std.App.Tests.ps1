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
        It 'Get-StdAppFolder zwraca App\<wersja modułu>' {
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

    Context 'Sync-StdAppFolder: kopia do App\<wersja>' {
        It 'kopiuje tylko pliki narzędzia do App\<wersja z STD.psd1> i zwraca ten folder' {
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

        It 'nie kopiuje, gdy źródło to już folder App\<wersja>' {
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
