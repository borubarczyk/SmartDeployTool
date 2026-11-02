# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy Modules\Std.Http.ps1 (Pester 5, Windows PowerShell 5.1). Bez sieci: tylko localhost i pliki.
# Test z prawdziwym serwerem HTTPS uruchamia się tylko, gdy ustawiono $env:STD_TEST_HTTPS_URL
# (np. https://localhost:8443/) i ten adres odpowiada; w CI jest pomijany.
# Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' (nawiasy w -like trzeba poprzedzić znakiem `).

BeforeDiscovery {
    $script:HttpsTestUrl = $env:STD_TEST_HTTPS_URL
}

Describe 'Std.Http' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'
        $script:Password = 'P@ss:wórd'
        $script:SecurePassword = ConvertTo-SecureString -String $script:Password -AsPlainText -Force
        $script:Session = New-StdSession -UserName 'tech' -Password $script:SecurePassword

        function Get-ClosedLocalPort {
            $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
            $listener.Start()
            $port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
            $listener.Stop()
            return $port
        }
    }

    AfterAll {
        Remove-Item Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
    }

    Context 'Join-StdLocation' {
        It 'łączy adres https przez "/" bez podwójnych ukośników' -TestCases @(
            @{ Base = 'https://srv/std/'; Child = 'clients/lizard/release.sdt'; Expected = 'https://srv/std/clients/lizard/release.sdt' }
            @{ Base = 'https://srv/std'; Child = 'clients.sdt'; Expected = 'https://srv/std/clients.sdt' }
            @{ Base = 'https://srv/std//'; Child = '/clients//x.sdt'; Expected = 'https://srv/std/clients/x.sdt' }
            @{ Base = 'https://srv/std/'; Child = 'files\lizard\a b.exe'; Expected = 'https://srv/std/files/lizard/a%20b.exe' }
            @{ Base = 'https://srv/std-admin/'; Child = 'clients/lizard/'; Expected = 'https://srv/std-admin/clients/lizard/' }
            @{ Base = 'https://srv/std/'; Child = ''; Expected = 'https://srv/std/' }
        ) {
            param($Base, $Child, $Expected)
            Join-StdLocation -Base $Base -Child $Child | Should -Be $Expected
        }

        It 'koduje segmenty adresu (EscapeDataString)' {
            Join-StdLocation -Base 'https://srv/std/' -Child 'files/Księgowość 1#.ps1' | Should -Be 'https://srv/std/files/Ksi%C4%99gowo%C5%9B%C4%87%201%23.ps1'
            Join-StdLocation -Base 'https://srv/std/' -Child '100%.txt' | Should -Be 'https://srv/std/100%25.txt'
        }

        It 'łączy ścieżki separatorem systemu' {
            $sep = [string][System.IO.Path]::DirectorySeparatorChar
            if ($sep -eq '\') {
                Join-StdLocation -Base '\\srv-files\STD$\std\' -Child 'files/lizard/7z.exe' | Should -Be '\\srv-files\STD$\std\files\lizard\7z.exe'
                Join-StdLocation -Base 'C:\STD' -Child 'Cache/lizard/release.sdt' | Should -Be 'C:\STD\Cache\lizard\release.sdt'
                Join-StdLocation -Base 'E:\' -Child '\offline\lizard' | Should -Be 'E:\offline\lizard'
                # dysk, którego nie ma, nie jest błędem (w odróżnieniu od Join-Path)
                Join-StdLocation -Base 'Q:\brak\dysku' -Child 'x' | Should -Be 'Q:\brak\dysku\x'
            } else {
                Join-StdLocation -Base '/srv/std' -Child 'clients\lizard\release.sdt' | Should -Be '/srv/std/clients/lizard/release.sdt'
                Join-StdLocation -Base '\\srv-files\STD$\std\' -Child 'files/a.exe' | Should -Be '//srv-files/STD$/std/files/a.exe'
            }
        }
    }

    Context 'New-StdSession i Get-StdBasicAuthHeader' {
        It 'sesja ma Mode, UserName i hasło jako SecureString' {
            $script:Session.Mode | Should -Be 'basic'
            $script:Session.UserName | Should -Be 'tech'
            $script:Session.Password | Should -BeOfType [System.Security.SecureString]
        }

        It 'jawne hasło nie występuje w sesji ani w jej opisie' {
            ($script:Session | Format-List * | Out-String) | Should -Not -BeLike "*$($script:Password)*"
            ($script:Session | ConvertTo-Json -Depth 5) | Should -Not -BeLike "*$($script:Password)*"
            foreach ($prop in $script:Session.PSObject.Properties) {
                [string]$prop.Value | Should -Not -Be $script:Password
            }
        }

        It 'nagłówek Basic to base64(UTF-8 "login:hasło")' {
            $expected = 'Basic ' + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('tech:' + $script:Password))
            Get-StdBasicAuthHeader -Session $script:Session | Should -Be $expected
            $expected | Should -Be 'Basic dGVjaDpQQHNzOnfDs3Jk'
        }

        It 'sesja z -Credential daje ten sam nagłówek' {
            $credential = New-Object System.Management.Automation.PSCredential('tech', $script:SecurePassword)
            $session = New-StdSession -Credential $credential
            Get-StdBasicAuthHeader -Session $session | Should -Be (Get-StdBasicAuthHeader -Session $script:Session)
        }

        It 'puste hasło daje "login:"' {
            $session = New-StdSession -UserName 'u' -Password (New-Object System.Security.SecureString)
            Get-StdBasicAuthHeader -Session $session | Should -Be ('Basic ' + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('u:')))
        }

        It 'odrzuca login z dwukropkiem i nieprawidłową sesję' {
            { New-StdSession -UserName 'a:b' -Password $script:SecurePassword } | Should -Throw -ExpectedMessage '*`[STD-SESSION`]*'
            { Get-StdBasicAuthHeader -Session @{ Mode = 'basic'; UserName = 'x'; Password = 'jawne' } } | Should -Throw -ExpectedMessage '*`[STD-SESSION`]*'
        }
    }

    Context 'Invoke-StdHttpRequest bez sieci' {
        It 'odrzuca adres <Uri>' -TestCases @(
            @{ Uri = 'http://srv/std/clients.sdt' }
            @{ Uri = 'ftp://srv/std/clients.sdt' }
            @{ Uri = 'std/clients.sdt' }
            @{ Uri = '\\srv\STD$\std\clients.sdt' }
        ) {
            param($Uri)
            { Invoke-StdHttpRequest -Uri $Uri } | Should -Throw -ExpectedMessage '*`[STD-HTTP-INSECURE`]*'
        }

        It 'odrzuca adres z loginem i hasłem i nie pokazuje hasła' {
            $message = $null
            try { Invoke-StdHttpRequest -Uri 'https://tech:Sekret99@srv/std/' } catch { $message = $_.Exception.Message }
            $message | Should -BeLike '*`[STD-HTTP-INSECURE`]*'
            $message | Should -Not -BeLike '*Sekret99*'
        }

        It 'zamknięty port -> [STD-HTTP-NET] bez hasła i nagłówka w komunikacie' {
            $port = Get-ClosedLocalPort
            $message = $null
            try {
                Invoke-StdHttpRequest -Uri "https://127.0.0.1:$port/std/clients.sdt" -Session $script:Session -TimeoutSec 10
            } catch {
                $message = $_.Exception.Message
            }
            $message | Should -BeLike '*`[STD-HTTP-NET`]*'
            $message | Should -BeLike "*GET https://127.0.0.1:$port/std/clients.sdt*"
            $message | Should -Not -BeLike "*$($script:Password)*"
            $message | Should -Not -BeLike '*Basic *'
            $message | Should -Not -BeLike '*dGVjaDpQ*'
        }
    }

    Context 'Get-StdBytes' {
        It 'czyta plik lokalny' {
            $path = Join-Path $TestDrive 'a.bin'
            [System.IO.File]::WriteAllBytes($path, [byte[]](0, 1, 2, 255))
            $bytes = Get-StdBytes -Location $path
            ($bytes -is [byte[]]) | Should -BeTrue
            $bytes.Length | Should -Be 4
            $bytes[3] | Should -Be 255
        }

        It 'pusty i jedno-bajtowy plik zwracają byte[]' {
            $empty = Join-Path $TestDrive 'pusty.bin'
            [System.IO.File]::WriteAllBytes($empty, (New-Object byte[] 0))
            $bytes = Get-StdBytes -Location $empty
            ($bytes -is [byte[]]) | Should -BeTrue
            $bytes.Length | Should -Be 0
            $one = Join-Path $TestDrive 'jeden.bin'
            [System.IO.File]::WriteAllBytes($one, [byte[]](7))
            $bytes = Get-StdBytes -Location $one
            ($bytes -is [byte[]]) | Should -BeTrue
            $bytes.Length | Should -Be 1
        }

        It 'działa ze ścieżką TestDrive:\' {
            [System.IO.File]::WriteAllBytes((Join-Path $TestDrive 'drive.bin'), [byte[]](1, 2, 3))
            (Get-StdBytes -Location 'TestDrive:\drive.bin').Length | Should -Be 3
        }

        It 'brak pliku -> [STD-SRC-NOTFOUND]' {
            { Get-StdBytes -Location (Join-Path $TestDrive 'nie-ma.bin') } | Should -Throw -ExpectedMessage '*`[STD-SRC-NOTFOUND`]*'
            { Get-StdBytes -Location $TestDrive } | Should -Throw -ExpectedMessage '*`[STD-SRC-NOTFOUND`]*'
        }

        It 'adres http:// -> [STD-HTTP-INSECURE]' {
            { Get-StdBytes -Location 'http://srv/std/clients.sdt' } | Should -Throw -ExpectedMessage '*`[STD-HTTP-INSECURE`]*'
        }

        It 'niedostępny serwer UNC -> [STD-SRC-NET] po teście portu 445' {
            # 192.0.2.1 to adres testowy (TEST-NET-1), nigdy nie odpowiada
            { Get-StdBytes -Location '\\192.0.2.1\STD$\std\clients.sdt' } | Should -Throw -ExpectedMessage '*`[STD-SRC-NET`]*'
        }
    }

    Context 'Test-StdTcpPort' {
        It 'zamknięty port na localhost -> $false' {
            $port = Get-ClosedLocalPort
            $result = Test-StdTcpPort -HostName '127.0.0.1' -Port $port -TimeoutMs 3000
            $result | Should -BeOfType [bool]
            $result | Should -BeFalse
        }

        It 'otwarty port na localhost -> $true' {
            $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
            $listener.Start()
            try {
                $port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
                Test-StdTcpPort -HostName '127.0.0.1' -Port $port | Should -BeTrue
            } finally {
                $listener.Stop()
            }
        }

        It 'nieznana nazwa serwera -> $false bez wyjątku' {
            Test-StdTcpPort -HostName 'nie-ma-takiego-serwera.invalid' -Port 445 -TimeoutMs 2000 | Should -BeFalse
        }
    }

    Context 'Initialize-StdNetwork' {
        BeforeAll {
            $script:OriginalProtocol = [System.Net.ServicePointManager]::SecurityProtocol
        }
        AfterAll {
            [System.Net.ServicePointManager]::SecurityProtocol = $script:OriginalProtocol
        }

        It 'ustawia TLS 1.2 i jest idempotentna' {
            Initialize-StdNetwork
            $first = [System.Net.ServicePointManager]::SecurityProtocol
            { Initialize-StdNetwork } | Should -Not -Throw
            $second = [System.Net.ServicePointManager]::SecurityProtocol
            $second | Should -Be $first
            ($first -band [System.Net.SecurityProtocolType]::Tls12) | Should -Be ([System.Net.SecurityProtocolType]::Tls12)
            ($first -band [System.Net.SecurityProtocolType]::Tls11) | Should -Be 0
        }

        It 'TLS 1.3 tylko na Windows build >= 20348' {
            Initialize-StdNetwork
            $protocol = [System.Net.ServicePointManager]::SecurityProtocol
            $names = [System.Enum]::GetNames([System.Net.SecurityProtocolType])
            if ($names -contains 'Tls13' -and [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
                $build = [int](Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name CurrentBuildNumber).CurrentBuildNumber
                $tls13 = [System.Net.SecurityProtocolType]'Tls13'
                if ($build -ge 20348) {
                    ($protocol -band $tls13) | Should -Be $tls13
                } else {
                    ($protocol -band $tls13) | Should -Be 0
                }
            }
        }
    }

    Context 'Serwer HTTPS na localhost (opcjonalnie)' {
        It 'GET przez STD_TEST_HTTPS_URL' -Skip:([string]::IsNullOrWhiteSpace($script:HttpsTestUrl)) {
            $uri = [System.Uri]$env:STD_TEST_HTTPS_URL
            if (-not (Test-StdTcpPort -HostName $uri.Host -Port $uri.Port -TimeoutMs 2000)) {
                Set-ItResult -Skipped -Because "serwer $($uri.Host):$($uri.Port) nie odpowiada"
                return
            }
            $result = Invoke-StdHttpRequest -Uri $env:STD_TEST_HTTPS_URL -Method GET -AllowStatus @(301, 302, 401, 403, 404) -TimeoutSec 15
            $result.StatusCode | Should -BeOfType [int]
            ($result.Bytes -is [byte[]]) | Should -BeTrue
            $result.Headers | Should -Not -BeNullOrEmpty
        }
    }
}
