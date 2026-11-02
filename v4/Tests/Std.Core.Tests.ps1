# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy Modules\Std.Core.ps1 (Pester 5, Windows PowerShell 5.1). Bez administratora i bez sieci.
# Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' - nawiasy kwadratowe trzeba poprzedzić
# znakiem ` (w -like oznaczają zbiór znaków, np. [STD-HTTP] to zakres liter, a nie tekst).

Describe 'Std.Core' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'
        $script:FixturePath = Join-Path $PSScriptRoot 'Fixtures\release.v4.json'
        $script:ManifestPath = Join-Path $PSScriptRoot '..\STD.psd1'
        $script:CoreFile = Join-Path $PSScriptRoot '..\Modules\Std.Core.ps1'
    }

    AfterAll {
        Clear-StdLogMasks
        Remove-Item Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
        Remove-Item Env:\STD_DEBUG -ErrorAction SilentlyContinue
    }

    Context 'Moduł i wersja' {
        It 'eksportuje tylko funkcje *-Std*' {
            $names = @(Get-Command -Module STD | Select-Object -ExpandProperty Name)
            $names | Should -Contain 'Write-StdLog'
            @($names | Where-Object { $_ -notlike '*-Std*' }) | Should -HaveCount 0
        }

        It 'Get-StdVersion zwraca [version] z manifestu' {
            $expected = [version](Import-PowerShellDataFile -Path $script:ManifestPath).ModuleVersion
            $version = Get-StdVersion
            $version | Should -BeOfType [version]
            $version | Should -Be $expected
        }

        It 'Get-StdVersion działa, gdy Std.Core.ps1 jest wczytany samodzielnie' {
            $coreFile = $script:CoreFile
            $expected = [version](Import-PowerShellDataFile -Path $script:ManifestPath).ModuleVersion
            $result = & {
                . $coreFile
                [pscustomobject]@{ Version = (Get-StdVersion); Module = (Get-Item -LiteralPath 'function:Get-StdVersion').Module }
            }
            $result.Module | Should -BeNullOrEmpty
            $result.Version | Should -BeOfType [version]
            $result.Version | Should -Be $expected
        }
    }

    Context 'Katalogi danych' {
        It 'Get-StdDataRoot zwraca STD_DATA_ROOT i tworzy katalog' {
            $root = Get-StdDataRoot
            $root | Should -Be $env:STD_DATA_ROOT
            Test-Path -LiteralPath $root -PathType Container | Should -BeTrue
        }

        It 'Get-StdPath <Kind> tworzy katalog <Name>' -TestCases @(
            @{ Kind = 'App'; Name = 'App' }
            @{ Kind = 'Cache'; Name = 'Cache' }
            @{ Kind = 'Logs'; Name = 'Logs' }
            @{ Kind = 'Outbox'; Name = 'Outbox' }
        ) {
            param($Kind, $Name)
            $path = Get-StdPath -Kind $Kind
            $path | Should -Be (Join-Path (Get-StdDataRoot) $Name)
            Test-Path -LiteralPath $path -PathType Container | Should -BeTrue
        }

        It 'Get-StdPath <Kind> zwraca plik <Name> i go nie tworzy' -TestCases @(
            @{ Kind = 'State'; Name = 'state.json' }
            @{ Kind = 'Revisions'; Name = 'revisions.json' }
            @{ Kind = 'Prefs'; Name = 'prefs.json' }
            @{ Kind = 'Bootstrap'; Name = 'bootstrap.local.json' }
        ) {
            param($Kind, $Name)
            $path = Get-StdPath -Kind $Kind
            $path | Should -Be (Join-Path (Get-StdDataRoot) $Name)
            Test-Path -LiteralPath $path | Should -BeFalse
        }

        It 'Get-StdPath Root zwraca katalog główny' {
            Get-StdPath -Kind Root | Should -Be (Get-StdDataRoot)
        }

        It 'Get-StdPath odrzuca nieznany rodzaj' {
            { Get-StdPath -Kind 'Nieznany' } | Should -Throw
        }

        It 'Get-StdDataRoot tworzy brakujący katalog' {
            $env:STD_DATA_ROOT = Join-Path $TestDrive 'nowy\podkatalog'
            try {
                $root = Get-StdDataRoot
                Test-Path -LiteralPath $root -PathType Container | Should -BeTrue
            } finally {
                $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'
            }
        }
    }

    Context 'ConvertTo-StdHashtable' {
        It 'zamienia PSCustomObject głęboko na [ordered] i zachowuje kolejność' {
            $obj = [pscustomobject]@{ z = 1; y = [pscustomobject]@{ x = 'a' }; w = @{ v = 2 } }
            $h = ConvertTo-StdHashtable -InputObject $obj
            $h | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
            (@($h.Keys) -join ',') | Should -Be 'z,y,w'
            $h.y | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
            $h.y.x | Should -Be 'a'
            $h.w | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
            $h.w.v | Should -Be 2
        }

        It 'zostawia tablice jedno-elementowe i puste jako object[]' {
            $obj = [pscustomobject]@{ one = @(5); none = @(); list = (New-Object 'System.Collections.Generic.List[string]') }
            $obj.list.Add('q')
            $h = ConvertTo-StdHashtable -InputObject $obj
            ($h.one -is [object[]]) | Should -BeTrue
            $h.one.Count | Should -Be 1
            ($h.none -is [object[]]) | Should -BeTrue
            $h.none.Count | Should -Be 0
            ($h.list -is [object[]]) | Should -BeTrue
            $h.list[0] | Should -Be 'q'
        }

        It 'zwraca tablicę z najwyższego poziomu bez rozwijania' {
            $single = ConvertTo-StdHashtable -InputObject @([pscustomobject]@{ a = 1 })
            ($single -is [object[]]) | Should -BeTrue
            $single.Count | Should -Be 1
            $single[0].a | Should -Be 1
            $empty = ConvertTo-StdHashtable -InputObject @()
            ($empty -is [object[]]) | Should -BeTrue
            $empty.Count | Should -Be 0
        }

        It 'zostawia inne wartości bez zmian' {
            ConvertTo-StdHashtable -InputObject $null | Should -BeNullOrEmpty
            ConvertTo-StdHashtable -InputObject 'tekst' | Should -Be 'tekst'
            ConvertTo-StdHashtable -InputObject 42 | Should -Be 42
            $bytes = ConvertTo-StdHashtable -InputObject ([pscustomobject]@{ b = [byte[]](1, 2) })
            ($bytes.b -is [byte[]]) | Should -BeTrue
        }

        It 'radzi sobie z kluczami Keys / Count / Values' {
            $h = ConvertTo-StdHashtable -InputObject @{ Keys = 'k'; Count = 'c'; Values = @(1) }
            $h['Keys'] | Should -Be 'k'
            $h['Count'] | Should -Be 'c'
            $h.get_Count() | Should -Be 3
        }
    }

    Context 'ConvertFrom-StdJson' {
        BeforeAll {
            $script:Release = ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText($script:FixturePath, (New-Object System.Text.UTF8Encoding($false))))
        }

        It 'wczytuje przykładową konfigurację jako [ordered]' {
            $script:Release | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
            $script:Release.client.id | Should -Be 'lizard'
            $script:Release.revision | Should -Be 42
        }

        It 'tablice z konfiguracji są object[], także jedno-elementowe' {
            ($script:Release.sites -is [object[]]) | Should -BeTrue
            $script:Release.sites.Count | Should -Be 2
            ($script:Release.catalog.AV.sourceOrder -is [object[]]) | Should -BeTrue
            $script:Release.catalog.AV.sourceOrder.Count | Should -Be 1
            $script:Release.catalog.AV.sourceOrder[0] | Should -Be 'web'
            ($script:Release.defaults.scripts -is [object[]]) | Should -BeTrue
            $script:Release.defaults.scripts.Count | Should -Be 1
            $script:Release.catalog.AV.successCodes[1] | Should -Be 3010
        }

        It 'zachowuje polskie znaki w kluczach i wartościach' {
            $script:Release.sites[0].name | Should -Be 'Kraków - biuro'
            $script:Release.sites[0].profiles.Contains('Księgowość') | Should -BeTrue
        }

        It 'daty zostają tekstem' {
            $script:Release.publishedAt | Should -BeOfType [string]
            $script:Release.expiresAt | Should -Be '2099-01-31T00:00:00Z'
        }

        It 'klucz "type" nie jest mylony z właściwością' {
            $script:Release.sources.lan.type | Should -Be 'unc'
            $script:Release.sources.lan.path | Should -Be '\\srv-files\STD$\std\files\lizard\'
        }

        It 'tablica na najwyższym poziomie: [1] i []' {
            $one = ConvertFrom-StdJson -Json '[1]'
            ($one -is [object[]]) | Should -BeTrue
            $one.Count | Should -Be 1
            $none = ConvertFrom-StdJson -Json '[]'
            ($none -is [object[]]) | Should -BeTrue
            $none.Count | Should -Be 0
        }

        It 'pusta tablica i null w obiekcie' {
            $h = ConvertFrom-StdJson -Json '{"a":[],"b":null,"c":[{"d":"ż"}]}'
            ($h.a -is [object[]]) | Should -BeTrue
            $h.a.Count | Should -Be 0
            $h.Contains('b') | Should -BeTrue
            $h.b | Should -BeNullOrEmpty
            $h.c[0] | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
            $h.c[0].d | Should -Be 'ż'
        }

        It 'pomija BOM na początku tekstu' {
            (ConvertFrom-StdJson -Json ([string][char]0xFEFF + '{"a":1}')).a | Should -Be 1
        }

        It 'przyjmuje linie z potoku' {
            $h = Get-Content -LiteralPath $script:FixturePath -Encoding UTF8 | ConvertFrom-StdJson
            $h.client.id | Should -Be 'lizard'
        }

        It 'zły JSON -> [STD-JSON]' {
            { ConvertFrom-StdJson -Json '{"a":' } | Should -Throw -ExpectedMessage '*`[STD-JSON`]*'
            { ConvertFrom-StdJson -Json '' } | Should -Throw -ExpectedMessage '*`[STD-JSON`]*'
            { ConvertFrom-StdJson -Json '   ' } | Should -Throw -ExpectedMessage '*`[STD-JSON`]*'
        }
    }

    Context 'ConvertTo-StdJson' {
        It 'zapisuje głęboko (więcej niż 2 poziomy)' {
            $deep = @{ l1 = @{ l2 = @{ l3 = @{ l4 = @{ l5 = @{ l6 = 'głęboko' } } } } } }
            (ConvertFrom-StdJson -Json (ConvertTo-StdJson -InputObject $deep)).l1.l2.l3.l4.l5.l6 | Should -Be 'głęboko'
        }

        It 'tablice zostają tablicami' {
            ConvertTo-StdJson -InputObject @{ a = @(1) } -Compress | Should -Be '{"a":[1]}'
            ConvertTo-StdJson -InputObject @{ a = @() } -Compress | Should -Be '{"a":[]}'
            ConvertTo-StdJson -InputObject @(1) -Compress | Should -Be '[1]'
        }

        It 'tablica opakowana w PSObject nie zamienia się w {"value":...,"Count":...}' {
            $wrapped = Write-Output -NoEnumerate @(1, 2)
            ConvertTo-StdJson -InputObject @{ a = $wrapped } -Compress | Should -Be '{"a":[1,2]}'
        }

        It 'daty jako ISO 8601, wersje jako tekst' {
            $date = [datetime]::SpecifyKind([datetime]'2026-11-01T20:00:00', [System.DateTimeKind]::Utc)
            ConvertTo-StdJson -InputObject @{ d = $date } -Compress | Should -Be '{"d":"2026-11-01T20:00:00.0000000Z"}'
            ConvertTo-StdJson -InputObject @{ v = [version]'4.0.0' } -Compress | Should -Be '{"v":"4.0.0"}'
        }

        It '$null -> null, -Compress bez nowych linii' {
            ConvertTo-StdJson -InputObject $null | Should -Be 'null'
            (ConvertTo-StdJson -InputObject @{ a = @{ b = 1 } } -Compress).Contains("`n") | Should -BeFalse
        }

        It 'konfiguracja przechodzi tam i z powrotem bez zmian' {
            $release = ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText($script:FixturePath))
            $again = ConvertFrom-StdJson -Json (ConvertTo-StdJson -InputObject $release)
            ConvertTo-StdJson -InputObject $again -Compress | Should -Be (ConvertTo-StdJson -InputObject $release -Compress)
        }
    }

    Context 'Write-StdJsonFile i Read-StdJsonFile' {
        BeforeAll {
            $script:Records = New-Object System.Collections.ArrayList
            $records = $script:Records
            Register-StdLogSink -Name 'test-json' -ScriptBlock { [void]$records.Add($args[0]) }.GetNewClosure()
        }
        AfterAll {
            Unregister-StdLogSink -Name 'test-json'
        }

        It 'zapisuje UTF-8 bez BOM, tworzy katalogi i nie zostawia .tmp' {
            $path = Join-Path $TestDrive 'json\a\b\test.json'
            Write-StdJsonFile -Path $path -InputObject ([ordered]@{ tekst = 'Zażółć gęślą jaźń'; lista = @(1) })
            $bytes = [System.IO.File]::ReadAllBytes($path)
            ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should -BeFalse
            Test-Path -LiteralPath "$path.tmp" | Should -BeFalse
            $back = Read-StdJsonFile -Path $path
            $back.tekst | Should -Be 'Zażółć gęślą jaźń'
            ($back.lista -is [object[]]) | Should -BeTrue
            $back.lista.Count | Should -Be 1
        }

        It 'nadpisuje istniejący plik' {
            $path = Join-Path $TestDrive 'json\nadpis.json'
            Write-StdJsonFile -Path $path -InputObject @{ wersja = 1 }
            Write-StdJsonFile -Path $path -InputObject @{ wersja = 2 }
            (Read-StdJsonFile -Path $path).wersja | Should -Be 2
            Test-Path -LiteralPath "$path.tmp" | Should -BeFalse
        }

        It 'działa ze ścieżką TestDrive:\' {
            Write-StdJsonFile -Path 'TestDrive:\json\drive.json' -InputObject @{ x = 1 }
            (Read-StdJsonFile -Path 'TestDrive:\json\drive.json').x | Should -Be 1
        }

        It 'brak pliku -> $null' {
            Read-StdJsonFile -Path (Join-Path $TestDrive 'json\brak.json') | Should -BeNullOrEmpty
        }

        It 'uszkodzony plik -> $null, nazwa .bad i ostrzeżenie w logu' {
            $path = Join-Path $TestDrive 'json\zepsuty.json'
            [System.IO.File]::WriteAllText($path, '{ to nie jest JSON')
            $script:Records.Clear()
            Read-StdJsonFile -Path $path | Should -BeNullOrEmpty
            Test-Path -LiteralPath $path | Should -BeFalse
            Test-Path -LiteralPath "$path.bad" | Should -BeTrue
            @($script:Records | Where-Object { $_.Level -eq 'Warn' -and $_.Message -like '*zepsuty.json*' }).Count | Should -Be 1
            # drugi raz: istniejący .bad jest zastępowany
            [System.IO.File]::WriteAllText($path, '')
            Read-StdJsonFile -Path $path | Should -BeNullOrEmpty
            Test-Path -LiteralPath $path | Should -BeFalse
        }
    }

    Context 'Write-StdLog i odbiorniki' {
        BeforeEach {
            $script:Records = New-Object System.Collections.ArrayList
            $records = $script:Records
            Register-StdLogSink -Name 'test' -ScriptBlock { [void]$records.Add($args[0]) }.GetNewClosure()
        }
        AfterEach {
            Unregister-StdLogSink -Name 'test'
            Remove-Item Env:\STD_DEBUG -ErrorAction SilentlyContinue
        }

        It 'przekazuje rekord @{ Time; Level; Message; Source }' {
            Write-StdLog -Message 'Ala ma kota' -Level Warn -Source 'Test'
            $script:Records.Count | Should -Be 1
            $record = $script:Records[0]
            $record.Level | Should -Be 'Warn'
            $record.Message | Should -Be 'Ala ma kota'
            $record.Source | Should -Be 'Test'
            $record.Time | Should -BeOfType [datetime]
        }

        It 'domyślny poziom to Info' {
            Write-StdLog -Message 'informacja'
            $script:Records[0].Level | Should -Be 'Info'
        }

        It 'Debug tylko przy STD_DEBUG=1' {
            Write-StdLog -Message 'ukryty' -Level Debug
            $script:Records.Count | Should -Be 0
            $env:STD_DEBUG = '1'
            Write-StdLog -Message 'widoczny' -Level Debug
            $script:Records.Count | Should -Be 1
            $script:Records[0].Level | Should -Be 'Debug'
        }

        It 'błąd odbiornika nie przerywa pracy i nie zatrzymuje innych' {
            $records = $script:Records
            Register-StdLogSink -Name 'test-blad' -ScriptBlock { throw 'odbiornik się wywrócił' }
            Register-StdLogSink -Name 'test-po' -ScriptBlock { [void]$records.Add('po') }.GetNewClosure()
            try {
                { Write-StdLog -Message 'dalej' } | Should -Not -Throw
                $script:Records.Count | Should -Be 2
                $script:Records[1] | Should -Be 'po'
            } finally {
                Unregister-StdLogSink -Name 'test-blad'
                Unregister-StdLogSink -Name 'test-po'
            }
        }

        It 'wynik odbiornika nie trafia na wyjście Write-StdLog' {
            Register-StdLogSink -Name 'test-gadatliwy' -ScriptBlock { 'tekst z odbiornika' }
            try {
                Write-StdLog -Message 'cisza' | Should -BeNullOrEmpty
            } finally {
                Unregister-StdLogSink -Name 'test-gadatliwy'
            }
        }

        It 'po wyrejestrowaniu odbiornik nic nie dostaje' {
            Unregister-StdLogSink -Name 'test'
            Write-StdLog -Message 'nikt nie słucha'
            $script:Records.Count | Should -Be 0
            { Unregister-StdLogSink -Name 'nie-ma-takiego' } | Should -Not -Throw
        }
    }

    Context 'Maski logu' {
        AfterEach { Clear-StdLogMasks }

        It 'ignoruje wartości krótsze niż 3 znaki' {
            Register-StdLogMask -Value 'ab'
            Protect-StdLogText -Text 'ab cd' | Should -Be 'ab cd'
        }

        It 'najdłuższe najpierw, bez względu na wielkość liter' {
            Register-StdLogMask -Value 'Tajne'
            Register-StdLogMask -Value 'TajneHaslo123'
            Protect-StdLogText -Text 'x TajneHaslo123 y TAJNE' | Should -Be 'x *** y ***'
        }

        It 'znaki specjalne w wartości nie są wyrażeniem regularnym' {
            Register-StdLogMask -Value 'a.b*c'
            Protect-StdLogText -Text 'a.b*c axbbc' | Should -Be '*** axbbc'
        }

        It 'Write-StdLog maskuje wiadomość i źródło' {
            $records = New-Object System.Collections.ArrayList
            Register-StdLogSink -Name 'test-maska' -ScriptBlock { [void]$records.Add($args[0]) }.GetNewClosure()
            try {
                Register-StdLogMask -Value 'Sekret!2026'
                Write-StdLog -Message 'Hasło: Sekret!2026' -Source 'sekret!2026'
                $records[0].Message | Should -Be 'Hasło: ***'
                $records[0].Source | Should -Be '***'
            } finally {
                Unregister-StdLogSink -Name 'test-maska'
            }
        }

        It 'nagłówek Authorization jest maskowany zawsze' {
            Protect-StdLogText -Text 'Authorization: Basic dGVjaDpwYXNz koniec' | Should -Be 'Authorization: Basic *** koniec'
        }

        It 'Clear-StdLogMasks usuwa maski' {
            Register-StdLogMask -Value 'TajneHaslo123'
            Clear-StdLogMasks
            Protect-StdLogText -Text 'TajneHaslo123' | Should -Be 'TajneHaslo123'
        }
    }

    Context 'Plik logu' {
        It 'zapisuje linie "yyyy-MM-dd HH:mm:ss [LEVEL] wiadomość" w UTF-8 bez BOM' {
            $path = Join-Path $TestDrive 'logi\deploy.log'
            Register-StdFileLogSink -Name 'test-plik' -Path $path
            try {
                Register-StdLogMask -Value 'sekret123'
                Write-StdLog -Message 'Zażółć gęślą jaźń'
                Write-StdLog -Message 'uwaga sekret123' -Level Warn
                Write-StdLog -Message 'błąd' -Level Error
                # plik da się czytać w trakcie zapisu
                $lines = @([System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($false))) -split "`r?`n" | Where-Object { $_ })
            } finally {
                Unregister-StdLogSink -Name 'test-plik'
                Clear-StdLogMasks
            }
            $lines.Count | Should -Be 3
            $lines[0] | Should -Match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[INFO\] Zażółć gęślą jaźń$'
            $lines[1] | Should -Match '\[WARN\] uwaga \*\*\*$'
            $lines[2] | Should -Match '\[ERROR\] błąd$'
            $bytes = [System.IO.File]::ReadAllBytes($path)
            $bytes[0] | Should -Not -Be 0xEF
        }

        It 'Unregister-StdLogSink zamyka plik' {
            $path = Join-Path $TestDrive 'logi\zamkniety.log'
            Register-StdFileLogSink -Name 'test-zamkniecie' -Path $path
            Write-StdLog -Message 'jedna linia'
            Unregister-StdLogSink -Name 'test-zamkniecie'
            $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Dispose()
            Remove-Item -LiteralPath $path
            Test-Path -LiteralPath $path | Should -BeFalse
        }

        It 'dopisuje do istniejącego pliku, także z dwóch odbiorników' {
            $path = Join-Path $TestDrive 'logi\dwa.log'
            Register-StdFileLogSink -Name 'test-a' -Path $path
            Register-StdFileLogSink -Name 'test-b' -Path $path
            try {
                Write-StdLog -Message 'raz'
            } finally {
                Unregister-StdLogSink -Name 'test-a'
                Unregister-StdLogSink -Name 'test-b'
            }
            Register-StdFileLogSink -Name 'test-a' -Path $path
            try { Write-StdLog -Message 'dwa' } finally { Unregister-StdLogSink -Name 'test-a' }
            $lines = @([System.IO.File]::ReadAllLines($path))
            $lines.Count | Should -Be 3
            @($lines | Where-Object { $_ -like '* raz' }).Count | Should -Be 2
            $lines[2] | Should -Match '\[INFO\] dwa$'
        }
    }

    Context 'Uprawnienia administratora' {
        It 'Test-StdIsAdmin zwraca [bool]' {
            Test-StdIsAdmin | Should -BeOfType [bool]
        }

        It 'Start-StdElevated uruchamia powershell.exe z poprawnie zacytowanymi argumentami' {
            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { }
            $result = Start-StdElevated -ScriptPath 'C:\Program Files\STD\Start-STD.ps1' -Arguments @('-Server', 'https://srv/std/', 'a "b" c', '', 'C:\katalog ze spacją\') -WorkingDirectory 'C:\Program Files\STD'
            $result | Should -BeTrue
            Should -Invoke -ModuleName STD -CommandName Start-CoreElevatedProcess -Times 1 -Exactly -ParameterFilter {
                $FilePath -like '*powershell.exe' -and
                $ArgumentLine -eq '-NoProfile -ExecutionPolicy Bypass -STA -File "C:\Program Files\STD\Start-STD.ps1" -Server https://srv/std/ "a \"b\" c" "" "C:\katalog ze spacją\\"' -and
                $WorkingDirectory -eq 'C:\Program Files\STD'
            }
        }

        It 'Start-StdElevated zwraca $false, gdy użytkownik odmówi w UAC (1223)' {
            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { throw (New-Object System.ComponentModel.Win32Exception(1223)) }
            Start-StdElevated -ScriptPath 'C:\STD\Start-STD.ps1' | Should -BeFalse
        }

        It 'Start-StdElevated zgłasza [STD-ELEVATE] przy innym błędzie' {
            Mock -ModuleName STD -CommandName Start-CoreElevatedProcess -MockWith { throw (New-Object System.ComponentModel.Win32Exception(2)) }
            { Start-StdElevated -ScriptPath 'C:\STD\Start-STD.ps1' } | Should -Throw -ExpectedMessage '*`[STD-ELEVATE`]*'
        }
    }
}
