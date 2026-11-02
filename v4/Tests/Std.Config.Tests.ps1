# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy Modules\Std.Config.ps1 (Pester 5, Windows PowerShell 5.1): zaufane certyfikaty, plik startowy
# bootstrap.json, widziane rewizje, Test-StdRelease, Merge-StdConfig, Resolve-StdEffectiveConfig,
# Get-StdRelease i Get-StdClientIndex. Bez administratora, sieci i magazynu certyfikatów: certyfikaty
# testowe powstają w pamięci (CertificateRequest), serwer to nieosiągalny adres https://127.0.0.1:9/std/
# (pierwsza próba od razu się nie udaje), kopia lustrzana, cache i pendrive to katalogi w TestDrive.
# Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' (nawiasy w -like trzeba poprzedzić znakiem `).

Describe 'Std.Config' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'

        $script:FixturePath = Join-Path (Join-Path $PSScriptRoot 'Fixtures') 'release.v4.json'
        $script:ConfigFixtures = Join-Path (Join-Path $PSScriptRoot 'Fixtures') 'config'
        $script:Utf8 = New-Object System.Text.UTF8Encoding($false)
        $script:CredentialText = 'Dane logowania do domeny i hasło lokalnego administratora wpisuje się ręcznie przy wdrożeniu - nie mogą być w konfiguracji.'
        $script:UnreachableServer = 'https://127.0.0.1:9/std/'

        # Klucze trzymamy do końca testów, żeby certyfikaty w pamięci nie straciły klucza prywatnego.
        $script:TestKeys = New-Object System.Collections.ArrayList

        function New-TestCertificate {
            param([string]$Subject)
            if ($PSVersionTable.PSEdition -eq 'Desktop') {
                # Windows PowerShell 5.1 (.NET Framework 4.8): klucz CNG tylko w pamięci.
                $key = New-Object System.Security.Cryptography.RSACng(2048)
            } else {
                $key = [System.Security.Cryptography.RSA]::Create(2048)
            }
            [void]$script:TestKeys.Add($key)
            $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
                $Subject, $key,
                [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $now = [System.DateTimeOffset]::UtcNow
            return $request.CreateSelfSigned($now.AddDays(-1), $now.AddDays(30))
        }

        function Get-TestPublicCertificate {
            # Kopia certyfikatu bez klucza prywatnego - tak jak wpis w Std.TrustedSigners.psd1.
            param($Certificate)
            return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([byte[]]$Certificate.RawData)
        }

        function Join-TestPath {
            # Join-Path z wieloma częściami (Windows PowerShell 5.1 nie ma -AdditionalChildPath).
            param([string]$Base, [string[]]$Parts)
            $path = $Base
            foreach ($part in $Parts) { $path = Join-Path $path $part }
            return $path
        }

        function Get-TestRelease {
            # Świeża kopia wspólnego przykładu konfiguracji (Tests\Fixtures\release.v4.json).
            return (ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText($script:FixturePath, $script:Utf8)))
        }

        function Get-TestIndex {
            param([int]$Revision = 7)
            return [ordered]@{
                schemaVersion = 4
                kind          = 'std.clients'
                revision      = $Revision
                publishedAt   = '2026-11-01T20:00:00Z'
                clients       = @([ordered]@{ id = 'lizard'; name = 'Lizard-Tech'; release = 'clients/lizard/release.sdt' })
            }
        }

        function Write-TestEnvelope {
            # Podpisuje dokument certyfikatem testowym i zapisuje kopertę *.sdt (UTF-8 bez BOM).
            param($Document, [string]$Path, $Certificate)
            if ($null -eq $Certificate) { $Certificate = $script:Cert }
            $payload = $script:Utf8.GetBytes((ConvertTo-StdJson -InputObject $Document))
            $envelope = New-StdSignedEnvelope -Payload $payload -Certificate $Certificate
            $dir = Split-Path -Parent $Path
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [System.IO.File]::WriteAllText($Path, $envelope, $script:Utf8)
        }

        function Write-TestTamperedEnvelope {
            # Koperta z podmienioną treścią i starym podpisem (jak plik zmieniony po publikacji).
            param($Document, $OtherDocument, [string]$Path)
            Write-TestEnvelope -Document $Document -Path $Path
            $envelope = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($Path, $script:Utf8))
            $envelope.payload = [System.Convert]::ToBase64String($script:Utf8.GetBytes((ConvertTo-StdJson -InputObject $OtherDocument)))
            [System.IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $envelope -Depth 30), $script:Utf8)
        }

        function Get-TestErrors {
            param($Document, [switch]$Unpublished)
            return @(Test-StdRelease -Release $Document -Unpublished:$Unpublished)
        }

        function Get-TestCaught {
            # Wywołuje blok i zwraca rekord błędu (albo $null).
            param([scriptblock]$Block)
            try {
                & $Block | Out-Null
            } catch {
                return $_
            }
            return $null
        }

        function Get-TestUtcText {
            param([datetime]$Value)
            return $Value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", [System.Globalization.CultureInfo]::InvariantCulture)
        }

        $script:Cert = New-TestCertificate -Subject 'CN=STD test - podpis konfiguracji'
        $script:OtherCert = New-TestCertificate -Subject 'CN=STD test - obcy podpis'
        $script:Trusted = Get-TestPublicCertificate $script:Cert
    }

    AfterAll {
        Remove-Item Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
        Clear-StdLogMasks
    }

    Context 'Get-StdTrustedCertificates' {
        It 'czyta listę z Modules\Std.TrustedSigners.psd1 bez błędu (w repozytorium może być pusta)' {
            $certificates = @(Get-StdTrustedCertificates)
            foreach ($certificate in $certificates) {
                $certificate | Should -BeOfType ([System.Security.Cryptography.X509Certificates.X509Certificate2])
            }
        }

        It 'pomija puste wpisy, błędne certyfikaty i wpisy z niezgodnym odciskiem' {
            $good = New-TestCertificate -Subject 'CN=STD test - zaufany'
            $base64 = [System.Convert]::ToBase64String($good.RawData)
            $thumbprint = $good.Thumbprint.ToUpperInvariant()
            $wrong = '0' * 40
            $content = @"
@{
    Signers = @(
        @{ Name = 'pusty'; Thumbprint = ''; Certificate = '' }
        @{ Name = 'dobry'; Thumbprint = '$($thumbprint.ToLowerInvariant())'; Certificate = '$base64' }
        @{ Name = 'zły odcisk'; Thumbprint = '$wrong'; Certificate = '$base64' }
        @{ Name = 'zły base64'; Thumbprint = ''; Certificate = 'to nie jest certyfikat' }
    )
}
"@
            $path = Join-Path $TestDrive 'signers.psd1'
            [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($true)))
            $result = @(InModuleScope -ModuleName STD -Parameters @{ SignersPath = $path } -ScriptBlock { param($SignersPath) foreach ($item in (Read-ConfigTrustedSigners -Path $SignersPath)) { $item } })
            $result.Count | Should -Be 1
            $result[0].Thumbprint.ToUpperInvariant() | Should -Be $thumbprint
        }

        It 'brak pliku z zaufanymi certyfikatami -> pusta lista' {
            $result = @(InModuleScope -ModuleName STD -Parameters @{ SignersPath = (Join-Path $TestDrive 'brak.psd1') } -ScriptBlock { param($SignersPath) foreach ($item in (Read-ConfigTrustedSigners -Path $SignersPath)) { $item } })
            $result.Count | Should -Be 0
        }
    }

    Context 'Get-StdBootstrap i Save-StdBootstrapLocal' {
        BeforeEach {
            $env:STD_DATA_ROOT = Join-Path $TestDrive ('data-' + [guid]::NewGuid().ToString('N'))
            $script:UsbDir = Join-Path $TestDrive ('usb-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $script:UsbDir | Out-Null
        }

        It '-Server: dopisuje ukośnik i uzupełnia pola domyślne' {
            $bootstrap = Get-StdBootstrap -Server 'https://pobieranie.firma.pl/std'
            $bootstrap.server | Should -Be 'https://pobieranie.firma.pl/std/'
            @($bootstrap.mirrors).Count | Should -Be 0
            $bootstrap.auth.mode | Should -Be 'basic'
            $bootstrap.proxy | Should -BeNullOrEmpty
        }

        It '-Server inny niż https:// -> [STD-CFG-BOOTSTRAP]' {
            { Get-StdBootstrap -Server 'http://pobieranie.firma.pl/std/' } | Should -Throw -ExpectedMessage '*`[STD-CFG-BOOTSTRAP`]*'
            { Get-StdBootstrap -Server 'pobieranie.firma.pl/std/' } | Should -Throw -ExpectedMessage '*`[STD-CFG-BOOTSTRAP`]*'
        }

        It 'adres z loginem i hasłem jest odrzucany, a hasło nie trafia do komunikatu' {
            $caught = Get-TestCaught { Get-StdBootstrap -Server 'https://technik:TajneHaslo123@pobieranie.firma.pl/std/' }
            $caught | Should -Not -BeNullOrEmpty
            $caught.Exception.Message | Should -BeLike '*`[STD-CFG-BOOTSTRAP`]*'
            $caught.Exception.Message | Should -Not -BeLike '*TajneHaslo123*'
        }

        It 'czyta bootstrap.json z katalogu skryptu (także z BOM)' {
            $json = '{ "server": "https://pobieranie.firma.pl/std/", "mirrors": [ "\\\\srv-files\\STD$\\std\\" ], "auth": { "mode": "basic" }, "proxy": null }'
            [System.IO.File]::WriteAllText((Join-Path $script:UsbDir 'bootstrap.json'), $json, (New-Object System.Text.UTF8Encoding($true)))
            $bootstrap = Get-StdBootstrap -ScriptRoot $script:UsbDir
            $bootstrap.server | Should -Be 'https://pobieranie.firma.pl/std/'
            @($bootstrap.mirrors).Count | Should -Be 1
            @($bootstrap.mirrors)[0] | Should -Be '\\srv-files\STD$\std\'
        }

        It '-Server ma pierwszeństwo przed bootstrap.json' {
            [System.IO.File]::WriteAllText((Join-Path $script:UsbDir 'bootstrap.json'), '{ "server": "https://a.firma.pl/std/" }', $script:Utf8)
            (Get-StdBootstrap -Server 'https://b.firma.pl/std/' -ScriptRoot $script:UsbDir).server | Should -Be 'https://b.firma.pl/std/'
        }

        It 'bootstrap.json z adresem http:// albo z błędnym JSON -> [STD-CFG-BOOTSTRAP]' {
            $path = Join-Path $script:UsbDir 'bootstrap.json'
            [System.IO.File]::WriteAllText($path, '{ "server": "http://pobieranie.firma.pl/std/" }', $script:Utf8)
            { Get-StdBootstrap -ScriptRoot $script:UsbDir } | Should -Throw -ExpectedMessage '*`[STD-CFG-BOOTSTRAP`]*'
            [System.IO.File]::WriteAllText($path, '{ "server": ', $script:Utf8)
            { Get-StdBootstrap -ScriptRoot $script:UsbDir } | Should -Throw -ExpectedMessage '*`[STD-CFG-BOOTSTRAP`]*'
        }

        It 'kopia lustrzana podana jako adres -> [STD-CFG-BOOTSTRAP]' {
            [System.IO.File]::WriteAllText((Join-Path $script:UsbDir 'bootstrap.json'), '{ "server": "https://a.firma.pl/std/", "mirrors": [ "http://kopia/std/" ] }', $script:Utf8)
            { Get-StdBootstrap -ScriptRoot $script:UsbDir } | Should -Throw -ExpectedMessage '*`[STD-CFG-BOOTSTRAP`]*'
        }

        It 'bez -Server i bootstrap.json bierze adres zapisany na komputerze; bez niczego zwraca $null' {
            Get-StdBootstrap -ScriptRoot $script:UsbDir | Should -BeNullOrEmpty
            Save-StdBootstrapLocal -Bootstrap @{ server = 'https://lokalny.firma.pl/std' }
            $bootstrap = Get-StdBootstrap -ScriptRoot $script:UsbDir
            $bootstrap.server | Should -Be 'https://lokalny.firma.pl/std/'
            (Get-StdBootstrap).server | Should -Be 'https://lokalny.firma.pl/std/'
        }

        It 'Save-StdBootstrapLocal zapisuje tylko znane pola (bez haseł) do bootstrap.local.json' {
            Save-StdBootstrapLocal -Bootstrap @{ server = 'https://lokalny.firma.pl/std/'; auth = @{ mode = 'basic'; password = 'NieZapisuj123' }; inne = 'x' }
            $path = Get-StdPath -Kind Bootstrap
            $path | Should -BeLike '*bootstrap.local.json'
            $text = [System.IO.File]::ReadAllText($path, $script:Utf8)
            $text | Should -Not -BeLike '*NieZapisuj123*'
            $saved = ConvertFrom-StdJson -Json $text
            @($saved.Keys) -join ',' | Should -Be 'server,mirrors,auth,proxy'
            @($saved.auth.Keys) -join ',' | Should -Be 'mode'
        }

        It 'Save-StdBootstrapLocal odrzuca adres inny niż https://' {
            { Save-StdBootstrapLocal -Bootstrap @{ server = 'http://lokalny.firma.pl/std/' } } | Should -Throw -ExpectedMessage '*`[STD-CFG-BOOTSTRAP`]*'
        }
    }

    Context 'Widziane rewizje (revisions.json)' {
        BeforeEach {
            $env:STD_DATA_ROOT = Join-Path $TestDrive ('data-' + [guid]::NewGuid().ToString('N'))
        }

        It 'brak wpisu -> 0' {
            Get-StdSeenRevision -Key 'lizard' | Should -Be 0
        }

        It 'zapisuje rewizje w formacie { "clients": 7, "lizard": 42 }' {
            Set-StdSeenRevision -Key 'clients' -Revision 7
            Set-StdSeenRevision -Key 'lizard' -Revision 42
            Get-StdSeenRevision -Key 'clients' | Should -Be 7
            Get-StdSeenRevision -Key 'lizard' | Should -Be 42
            $saved = ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText((Get-StdPath -Kind Revisions), $script:Utf8))
            $saved['clients'] | Should -Be 7
            $saved['lizard'] | Should -Be 42
        }

        It 'nigdy nie zmniejsza zapisanej rewizji' {
            Set-StdSeenRevision -Key 'lizard' -Revision 42
            Set-StdSeenRevision -Key 'lizard' -Revision 41
            Get-StdSeenRevision -Key 'lizard' | Should -Be 42
            Set-StdSeenRevision -Key 'lizard' -Revision 43
            Get-StdSeenRevision -Key 'lizard' | Should -Be 43
        }

        It 'uszkodzony revisions.json -> 0 (plik przeniesiony do .bad)' {
            $path = Get-StdPath -Kind Revisions
            [System.IO.File]::WriteAllText($path, '{ "lizard": ', $script:Utf8)
            Get-StdSeenRevision -Key 'lizard' | Should -Be 0
            Test-Path -LiteralPath ($path + '.bad') | Should -BeTrue
        }
    }

    Context 'Test-StdRelease - poprawne dokumenty' {
        It 'wspólny przykład (Tests\Fixtures\release.v4.json) nie ma błędów' {
            $errors = Get-TestErrors (Get-TestRelease)
            ($errors -join ' | ') | Should -BeNullOrEmpty
            $errors.Count | Should -Be 0
        }

        It 'przyjmuje też obiekt z ConvertFrom-Json (PSCustomObject)' {
            $object = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($script:FixturePath, $script:Utf8))
            (Get-TestErrors $object).Count | Should -Be 0
        }

        It '-Unpublished: revision, publishedAt i publishedBy nie są wymagane' {
            $doc = Get-TestRelease
            $doc.Remove('revision')
            $doc.Remove('publishedAt')
            $doc.Remove('publishedBy')
            (Get-TestErrors $doc -Unpublished).Count | Should -Be 0
            $errors = Get-TestErrors $doc
            @($errors | Where-Object { $_.StartsWith('revision:') }).Count | Should -Be 1
            @($errors | Where-Object { $_.StartsWith('publishedAt:') }).Count | Should -Be 1
            @($errors | Where-Object { $_.StartsWith('publishedBy:') }).Count | Should -Be 1
        }

        It 'localAdmin.username jest dozwolony, a puste pole o nazwie password nie jest błędem' {
            $doc = Get-TestRelease
            $doc.catalog.Chrome['password'] = ''
            (Get-TestErrors $doc).Count | Should -Be 0
        }

        It 'minimalny dokument z pliku Fixtures\config\minimal.v4.json nie ma błędów' {
            $doc = ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText((Join-Path $script:ConfigFixtures 'minimal.v4.json'), $script:Utf8))
            $errors = Get-TestErrors $doc -Unpublished
            ($errors -join ' | ') | Should -BeNullOrEmpty
        }

        It 'brak dokumentu ($null) -> jeden błąd' {
            @(Test-StdRelease -Release $null).Count | Should -Be 1
        }
    }

    Context 'Test-StdRelease - reguły (jeden błąd na regułę)' {
        It 'odrzuca: <Name>' -TestCases @(
            @{ Name = 'schemaVersion inny niż 4'; Expected = 'schemaVersion:'; Mutate = { param($d) $d['schemaVersion'] = 3 } }
            @{ Name = 'brak schemaVersion'; Expected = 'schemaVersion:'; Mutate = { param($d) $d.Remove('schemaVersion') } }
            @{ Name = 'kind inny niż std.release'; Expected = 'kind:'; Mutate = { param($d) $d['kind'] = 'std.clients' } }
            @{ Name = 'client.id niezgodny ze wzorcem'; Expected = 'client.id:'; Mutate = { param($d) $d.client['id'] = 'Lizard_Tech' } }
            @{ Name = 'client.id zarezerwowany (clients)'; Expected = 'client.id:'; Mutate = { param($d) $d.client['id'] = 'clients' } }
            @{ Name = 'pusty client.name'; Expected = 'client.name:'; Mutate = { param($d) $d.client['name'] = ' ' } }
            @{ Name = 'brak revision po publikacji'; Expected = 'revision:'; Mutate = { param($d) $d.Remove('revision') } }
            @{ Name = 'revision 0'; Expected = 'revision:'; Mutate = { param($d) $d['revision'] = 0 } }
            @{ Name = 'revision jako tekst'; Expected = 'revision:'; Mutate = { param($d) $d['revision'] = '42' } }
            @{ Name = 'publishedAt nie jest datą ISO 8601'; Expected = 'publishedAt:'; Mutate = { param($d) $d['publishedAt'] = '01.11.2026 20:00' } }
            @{ Name = 'pusty publishedBy'; Expected = 'publishedBy:'; Mutate = { param($d) $d['publishedBy'] = '' } }
            @{ Name = 'expiresAt bez strefy czasu'; Expected = 'expiresAt:'; Mutate = { param($d) $d['expiresAt'] = '2099-01-31T00:00:00' } }
            @{ Name = 'minClientVersion nie jest wersją'; Expected = 'minClientVersion:'; Mutate = { param($d) $d['minClientVersion'] = 'najnowsza' } }
            @{ Name = 'allowOffline jako tekst'; Expected = 'allowOffline:'; Mutate = { param($d) $d['allowOffline'] = 'true' } }
            @{ Name = 'requireSha256 jako liczba'; Expected = 'requireSha256:'; Mutate = { param($d) $d['requireSha256'] = 1 } }
            @{ Name = 'changeNote nie jest tekstem'; Expected = 'changeNote:'; Mutate = { param($d) $d['changeNote'] = 5 } }
            @{ Name = 'tool.url inny niż https://'; Expected = 'tool.url:'; Mutate = { param($d) $d['tool'] = [ordered]@{ version = '4.0.1'; url = 'http://serwer/std/tool/STD-4.0.1.zip' } } }
            @{ Name = 'brak sources'; Expected = 'sources:'; Mutate = { param($d) $d.Remove('sources') } }
            @{ Name = 'puste sources'; Expected = 'sources:'; Mutate = { param($d) $d['sources'] = [ordered]@{} } }
            @{ Name = 'sources.*.type spoza listy'; Expected = 'sources.lan.type:'; Mutate = { param($d) $d.sources.lan['type'] = 'ftp' } }
            @{ Name = 'źródło https z adresem http://'; Expected = 'sources.web.url:'; Mutate = { param($d) $d.sources.web['url'] = 'http://pobieranie.example.com/std/files/lizard/' } }
            @{ Name = 'źródło https z loginem w adresie'; Expected = 'sources.web.url:'; Mutate = { param($d) $d.sources.web['url'] = 'https://user:haslo@pobieranie.example.com/' } }
            @{ Name = 'źródło unc bez path'; Expected = 'sources.lan.path:'; Mutate = { param($d) $d.sources.lan.Remove('path') } }
            @{ Name = 'sources.*.auth spoza listy'; Expected = 'sources.lan.auth:'; Mutate = { param($d) $d.sources.lan['auth'] = 'basic' } }
            @{ Name = 'brak catalog'; Expected = 'catalog:'; Mutate = { param($d) $d.Remove('catalog') } }
            @{ Name = 'klucz katalogu niezgodny ze wzorcem'; Expected = 'catalog.7 zip:'; Mutate = { param($d) $d.catalog['7 zip'] = [ordered]@{ name = 'X'; winget = 'X.X' } } }
            @{ Name = 'pusta nazwa pozycji katalogu'; Expected = 'catalog.7zip.name:'; Mutate = { param($d) $d.catalog['7zip']['name'] = '' } }
            @{ Name = 'pozycja bez file, url i winget'; Expected = 'catalog.Git:'; Mutate = { param($d) $d.catalog.Git.Remove('winget') } }
            @{ Name = 'url pozycji katalogu inny niż https://'; Expected = 'catalog.Chrome.url:'; Mutate = { param($d) $d.catalog.Chrome['url'] = 'http://dl.google.com/chrome.msi' } }
            @{ Name = 'sha256 nie jest 64 znakami HEX'; Expected = 'catalog.7zip.sha256:'; Mutate = { param($d) $d.catalog['7zip']['sha256'] = 'abc' } }
            @{ Name = 'timeoutMin 0'; Expected = 'catalog.7zip.timeoutMin:'; Mutate = { param($d) $d.catalog['7zip']['timeoutMin'] = 0 } }
            @{ Name = 'timeoutMin 241'; Expected = 'catalog.7zip.timeoutMin:'; Mutate = { param($d) $d.catalog['7zip']['timeoutMin'] = 241 } }
            @{ Name = 'successCodes z tekstem'; Expected = 'catalog.AV.successCodes[1]:'; Mutate = { param($d) $d.catalog.AV['successCodes'] = @(0, 'ok') } }
            @{ Name = 'whenDetected spoza listy'; Expected = 'catalog.TeamViewer.whenDetected:'; Mutate = { param($d) $d.catalog.TeamViewer['whenDetected'] = 'always' } }
            @{ Name = 'sourceOrder pozycji wskazuje nieistniejące źródło'; Expected = 'catalog.AV.sourceOrder[0]:'; Mutate = { param($d) $d.catalog.AV['sourceOrder'] = @('av-web') } }
            @{ Name = 'requireSha256 bez sha256 w pozycji z file'; Expected = 'catalog.TeamViewer.sha256:'; Mutate = { param($d) $d['requireSha256'] = $true } }
            @{ Name = 'detect bez uninstallName i path'; Expected = 'catalog.7zip.detect:'; Mutate = { param($d) $d.catalog['7zip']['detect'] = [ordered]@{ minVersion = '24.08' } } }
            @{ Name = 'detect.minVersion nie jest wersją'; Expected = 'catalog.7zip.detect.minVersion:'; Mutate = { param($d) $d.catalog['7zip'].detect['minVersion'] = 'v24' } }
            @{ Name = 'nieznane zadanie w defaults.tasks'; Expected = 'defaults.tasks.JoinIntune:'; Mutate = { param($d) $d.defaults.tasks['JoinIntune'] = $true } }
            @{ Name = 'zadanie z wartością inną niż bool'; Expected = 'defaults.tasks.WaitForNetwork:'; Mutate = { param($d) $d.defaults.tasks['WaitForNetwork'] = 'tak' } }
            @{ Name = 'nieznane zadanie w profiles.*.tasks'; Expected = 'sites[0].profiles.Księgowość.tasks.DryRun:'; Mutate = { param($d) $d.sites[0].profiles['Księgowość'].tasks['DryRun'] = $true } }
            @{ Name = 'bloatware z pustym wzorcem'; Expected = 'defaults.bloatware[0]:'; Mutate = { param($d) $d.defaults['bloatware'] = @('') } }
            @{ Name = 'uninstall bez name'; Expected = 'defaults.uninstall[1].name:'; Mutate = { param($d) $d.defaults.uninstall[1].Remove('name') } }
            @{ Name = 'flaga systemSettings inna niż bool'; Expected = 'defaults.systemSettings.DisableTelemetry:'; Mutate = { param($d) $d.defaults.systemSettings['DisableTelemetry'] = 'tak' } }
            @{ Name = 'registry[].hive spoza listy'; Expected = 'defaults.systemSettings.registry[0].hive:'; Mutate = { param($d) $d.defaults.systemSettings.registry[0]['hive'] = 'HKEY_LOCAL_MACHINE' } }
            @{ Name = 'registry[].type spoza listy'; Expected = 'defaults.systemSettings.registry[0].type:'; Mutate = { param($d) $d.defaults.systemSettings.registry[0]['type'] = 'Integer' } }
            @{ Name = 'registry[] bez path'; Expected = 'defaults.systemSettings.registry[1].path:'; Mutate = { param($d) $d.defaults.systemSettings.registry[1].Remove('path') } }
            @{ Name = 'registry[] bez name'; Expected = 'defaults.systemSettings.registry[1].name:'; Mutate = { param($d) $d.defaults.systemSettings.registry[1].Remove('name') } }
            @{ Name = 'scripts[] bez file'; Expected = 'defaults.scripts[0].file:'; Mutate = { param($d) $d.defaults.scripts[0].Remove('file') } }
            @{ Name = 'brak sites'; Expected = 'sites:'; Mutate = { param($d) $d['sites'] = @() } }
            @{ Name = 'sites[].id niezgodny ze wzorcem'; Expected = 'sites[0].id:'; Mutate = { param($d) $d.sites[0]['id'] = 'Kraków' } }
            @{ Name = 'powtórzony sites[].id'; Expected = 'sites[1].id:'; Mutate = { param($d) $d.sites[1]['id'] = 'krakow' } }
            @{ Name = 'pusta nazwa lokalizacji'; Expected = 'sites[1].name:'; Mutate = { param($d) $d.sites[1]['name'] = '' } }
            @{ Name = 'sites[].sourceOrder wskazuje nieistniejące źródło'; Expected = 'sites[0].sourceOrder[1]:'; Mutate = { param($d) $d.sites[0]['sourceOrder'] = @('lan', 'ftp') } }
            @{ Name = 'computerName.maxLength powyżej 15'; Expected = 'sites[0].computerName.maxLength:'; Mutate = { param($d) $d.sites[0].computerName['maxLength'] = 16 } }
            @{ Name = 'computerName.pattern z niedozwoloną wstawką'; Expected = 'sites[0].computerName.pattern:'; Mutate = { param($d) $d.sites[0].computerName['pattern'] = 'KRK-{user}' } }
            @{ Name = 'computerName.pattern z nawiasem bez pary'; Expected = 'sites[0].computerName.pattern:'; Mutate = { param($d) $d.sites[0].computerName['pattern'] = 'KRK-{serial' } }
            @{ Name = 'wifi z auth spoza listy'; Expected = 'sites[0].wifi[0].auth:'; Mutate = { param($d) $d.sites[0].wifi[0]['auth'] = 'WEP' } }
            @{ Name = 'wifi z hasłem bez keyRef'; Expected = 'sites[0].wifi[0].keyRef:'; Mutate = { param($d) $d.sites[0].wifi[0].Remove('keyRef') } }
            @{ Name = 'bitLocker.escrow spoza listy'; Expected = 'sites[0].bitLocker.escrow[1]:'; Mutate = { param($d) $d.sites[0].bitLocker['escrow'] = @('ad', 'usb') } }
            @{ Name = 'domainJoin bez domain'; Expected = 'sites[1].domainJoin.domain:'; Mutate = { param($d) $d.sites[1].domainJoin.Remove('domain') } }
            @{ Name = 'localAdmin bez username'; Expected = 'sites[1].localAdmin.username:'; Mutate = { param($d) $d.sites[1].localAdmin.Remove('username') } }
            @{ Name = 'profiles.*.apps wskazuje nieistniejącą pozycję katalogu'; Expected = 'sites[1].profiles.Standard.apps[3]:'; Mutate = { param($d) $d.sites[1].profiles.Standard['apps'] = @('Chrome', '7zip', 'Git', 'Firefox') } }
            @{ Name = 'profiles.*.locked z nieznaną wartością'; Expected = 'sites[0].profiles.Standard.locked[1]:'; Mutate = { param($d) $d.sites[0].profiles.Standard['locked'] = @('JoinDomain', 'wszystko') } }
            @{ Name = 'lokalizacja bez profili'; Expected = 'sites[1].profiles:'; Mutate = { param($d) $d.sites[1]['profiles'] = [ordered]@{} } }
            @{ Name = 'secrets.mode inny niż sealed'; Expected = 'secrets.mode:'; Mutate = { param($d) $d.secrets['mode'] = 'plain' } }
            @{ Name = 'secrets.kdf.alg inny niż pbkdf2-sha256'; Expected = 'secrets.kdf.alg:'; Mutate = { param($d) $d.secrets.kdf['alg'] = 'md5' } }
            @{ Name = 'pozycja sekretu z błędnym mac'; Expected = 'secrets.items.wifi-firma.mac:'; Mutate = { param($d) $d.secrets.items['wifi-firma']['mac'] = 'AAAA' } }
        ) {
            param($Name, $Expected, $Mutate)
            $doc = Get-TestRelease
            & $Mutate $doc
            $errors = Get-TestErrors $doc
            $hits = @($errors | Where-Object { $_.Contains($Expected) })
            $hits.Count | Should -BeGreaterThan 0 -Because ("oczekiwano błędu '$Expected', są: " + ($errors -join ' | '))
        }
    }

    Context 'Test-StdRelease - dane logowania, jawne hasła i sekrety' {
        It 'odrzuca dane logowania w <Field>' -TestCases @(
            @{ Field = 'sites[0].domainJoin.password'; Mutate = { param($d) $d.sites[0].domainJoin['password'] = 'Haslo123!' } }
            @{ Field = 'sites[0].domainJoin.username'; Mutate = { param($d) $d.sites[0].domainJoin['username'] = 'FIRMA\dolaczanie' } }
            @{ Field = 'sites[0].domainJoin.User'; Mutate = { param($d) $d.sites[0].domainJoin['User'] = 'dolaczanie' } }
            @{ Field = 'sites[0].domainJoin.joinAccount'; Mutate = { param($d) $d.sites[0].domainJoin['joinAccount'] = 'dolaczanie' } }
            @{ Field = 'sites[0].domainJoin.login'; Mutate = { param($d) $d.sites[0].domainJoin['login'] = 'dolaczanie' } }
            @{ Field = 'sites[0].domainJoin.credential'; Mutate = { param($d) $d.sites[0].domainJoin['credential'] = [ordered]@{ name = 'x' } } }
            @{ Field = 'sites[0].domainJoin.pwd'; Mutate = { param($d) $d.sites[0].domainJoin['pwd'] = '' } }
            @{ Field = 'sites[0].domainJoin.secretRef'; Mutate = { param($d) $d.sites[0].domainJoin['secretRef'] = 'domena' } }
            @{ Field = 'sites[0].domainJoin.extra.haslo'; Mutate = { param($d) $d.sites[0].domainJoin['extra'] = [ordered]@{ haslo = 'x' } } }
            @{ Field = 'sites[0].localAdmin.password'; Mutate = { param($d) $d.sites[0].localAdmin['password'] = 'Haslo123!' } }
            @{ Field = 'sites[0].localAdmin.hasło'; Mutate = { param($d) $d.sites[0].localAdmin['hasło'] = 'Haslo123!' } }
            @{ Field = 'sites[0].localAdmin.passwordRef'; Mutate = { param($d) $d.sites[0].localAdmin['passwordRef'] = 'admin' } }
            @{ Field = 'sites[1].localAdmin.user'; Mutate = { param($d) $d.sites[1].localAdmin['user'] = 'drugi' } }
            @{ Field = 'defaults.domainJoin.password'; Mutate = { param($d) $d.defaults['domainJoin'] = [ordered]@{ domain = 'firma.local'; password = 'x' } } }
            @{ Field = 'sites[0].domainJoin.ou'; Mutate = { param($d) $d.sites[0].domainJoin['ou'] = '{{secret:wifi-firma}}' } }
        ) {
            param($Field, $Mutate)
            $doc = Get-TestRelease
            & $Mutate $doc
            $errors = Get-TestErrors $doc
            $expected = $Field + ': ' + $script:CredentialText
            $hits = @($errors | Where-Object { $_.StartsWith($expected) })
            $hits.Count | Should -BeGreaterThan 0 -Because ("oczekiwano '$expected', są: " + ($errors -join ' | '))
        }

        It 'odrzuca jawne hasło lub token w polu <Field>' -TestCases @(
            @{ Field = 'catalog.Chrome.token'; Mutate = { param($d) $d.catalog.Chrome['token'] = 'abc123' } }
            @{ Field = 'catalog.TeamViewer.ApiToken'; Mutate = { param($d) $d.catalog.TeamViewer['ApiToken'] = 'abc123' } }
            @{ Field = 'sources.web.password'; Mutate = { param($d) $d.sources.web['password'] = 'abc123' } }
            @{ Field = 'sites[0].wifi[0].secret'; Mutate = { param($d) $d.sites[0].wifi[0]['secret'] = 'KluczWifi1' } }
            @{ Field = 'defaults.systemSettings.registry[0].Passwd'; Mutate = { param($d) $d.defaults.systemSettings.registry[0]['Passwd'] = 'x' } }
            @{ Field = 'escrow.HASŁO'; Mutate = { param($d) $d['escrow'] = [ordered]@{ drop = 'drop/'; 'HASŁO' = 'x' } } }
        ) {
            param($Field, $Mutate)
            $doc = Get-TestRelease
            & $Mutate $doc
            $errors = Get-TestErrors $doc
            $hits = @($errors | Where-Object { $_.StartsWith($Field + ': ') -and $_.Contains('jawnego hasła ani tokenu') })
            $hits.Count | Should -BeGreaterThan 0 -Because ("oczekiwano błędu w $Field, są: " + ($errors -join ' | '))
        }

        It 'komunikat o jawnym haśle nie zawiera jego wartości' {
            $doc = Get-TestRelease
            $doc.sources.web['password'] = 'BardzoTajneHaslo987'
            $errors = Get-TestErrors $doc
            $errors.Count | Should -BeGreaterThan 0
            ($errors -join ' ') | Should -Not -BeLike '*BardzoTajneHaslo987*'
        }

        It 'sekret o nazwie wskazującej dane logowania do domeny jest odrzucany' {
            $doc = Get-TestRelease
            $doc.secrets.items['domain-join-password'] = $doc.secrets.items['wifi-firma']
            $errors = Get-TestErrors $doc
            @($errors | Where-Object { $_.StartsWith('secrets.items.domain-join-password: ' + $script:CredentialText) }).Count | Should -Be 1
        }

        It 'zgłasza brakujący sekret z <Field>' -TestCases @(
            @{ Field = 'catalog.TeamViewer.args'; Secret = 'teamviewer-apitoken'; Mutate = { param($d) $d.secrets.items.Remove('teamviewer-apitoken') } }
            @{ Field = 'catalog.Chrome.url'; Secret = 'chrome-token'; Mutate = { param($d) $d.catalog.Chrome['url'] = 'https://dl.example.com/chrome.msi?t={{secret:chrome-token}}' } }
            @{ Field = 'sites[0].profiles.Księgowość.scripts[0].args'; Secret = 'ksiegowosc-klucz'; Mutate = { param($d) $d.sites[0].profiles['Księgowość'].scripts[0]['args'] = '-Klucz {{secret:ksiegowosc-klucz}}' } }
            @{ Field = 'defaults.scripts[0].args'; Secret = 'drukarki'; Mutate = { param($d) $d.defaults.scripts[0]['args'] = '-Haslo {{secret:drukarki}}' } }
            @{ Field = 'sites[0].wifi[0].keyRef'; Secret = 'wifi-firma'; Mutate = { param($d) $d.secrets.items.Remove('wifi-firma') } }
        ) {
            param($Field, $Secret, $Mutate)
            $doc = Get-TestRelease
            & $Mutate $doc
            $errors = Get-TestErrors $doc
            $hits = @($errors | Where-Object { $_.StartsWith($Field + ': ') -and $_.Contains("'$Secret'") })
            $hits.Count | Should -BeGreaterThan 0 -Because ("oczekiwano braku sekretu '$Secret' w $Field, są: " + ($errors -join ' | '))
        }

        It 'bez bloku secrets zgłasza wszystkie wstawki i keyRef' {
            $doc = Get-TestRelease
            $doc.Remove('secrets')
            $errors = Get-TestErrors $doc
            foreach ($name in @('teamviewer-customconfigid', 'teamviewer-apitoken', 'wifi-firma')) {
                @($errors | Where-Object { $_.Contains("brak sekretu '$name'") }).Count | Should -BeGreaterThan 0 -Because ($errors -join ' | ')
            }
        }

        It 'odrzuca nieprawidłową wstawkę sekretu' {
            $doc = Get-TestRelease
            $doc.catalog.AV['args'] = '/qn KLUCZ={{secret:zła nazwa}}'
            $errors = Get-TestErrors $doc
            @($errors | Where-Object { $_.StartsWith('catalog.AV.args: ') -and $_.Contains('nieprawidłowa wstawka sekretu') }).Count | Should -Be 1
        }
    }

    Context 'Merge-StdConfig' {
        It 'łączy słowniki rekurencyjnie, zastępuje tablice i wartości proste, $null zostawia Base' {
            $base = [ordered]@{ a = [ordered]@{ x = 1; y = 2 }; list = @(1, 2, 3); text = 'base'; keep = 'k' }
            $override = @{ a = @{ y = 3; z = 4 }; list = @(9); text = $null; extra = 'e' }
            $result = Merge-StdConfig -Base $base -Override $override
            $result.a.x | Should -Be 1
            $result.a.y | Should -Be 3
            $result.a.z | Should -Be 4
            , $result.list | Should -BeOfType ([object[]])
            @($result.list).Count | Should -Be 1
            $result.list[0] | Should -Be 9
            $result.text | Should -Be 'base'
            $result.keep | Should -Be 'k'
            $result.extra | Should -Be 'e'
        }

        It 'zwraca nowy obiekt - wejścia się nie zmieniają' {
            $base = [ordered]@{ a = [ordered]@{ x = 1 }; list = @(1, 2) }
            $override = [ordered]@{ a = [ordered]@{ y = 2 } }
            $result = Merge-StdConfig -Base $base -Override $override
            $result.a['x'] = 100
            $result.list[0] = 100
            $base.a.x | Should -Be 1
            $base.a.Contains('y') | Should -BeFalse
            $base.list[0] | Should -Be 1
            $override.a.Contains('x') | Should -BeFalse
        }

        It 'słownik zastępuje wartość prostą i odwrotnie; $null jako Base albo Override' {
            (Merge-StdConfig -Base @{ a = 'tekst' } -Override @{ a = @{ b = 1 } }).a.b | Should -Be 1
            (Merge-StdConfig -Base @{ a = @{ b = 1 } } -Override @{ a = 'tekst' }).a | Should -Be 'tekst'
            (Merge-StdConfig -Base $null -Override @{ a = 1 }).a | Should -Be 1
            (Merge-StdConfig -Base @{ a = 1 } -Override $null).a | Should -Be 1
        }

        It 'defaults.tasks + profile.tasks z przykładu' {
            $doc = Get-TestRelease
            $tasks = Merge-StdConfig -Base $doc.defaults.tasks -Override $doc.sites[0].profiles['Księgowość'].tasks
            $tasks.EnableBitLocker | Should -BeTrue
            $tasks.JoinDomain | Should -BeTrue
            @($tasks.Keys).Count | Should -Be @($doc.defaults.tasks.Keys).Count
            $doc.defaults.tasks.EnableBitLocker | Should -BeFalse
        }
    }

    Context 'Resolve-StdEffectiveConfig' {
        BeforeAll {
            $script:Release = Get-TestRelease
        }

        It 'Kraków / Standard' {
            $config = Resolve-StdEffectiveConfig -Release $script:Release -SiteId 'krakow' -ProfileName 'Standard'
            $config.ClientId | Should -Be 'lizard'
            $config.ClientName | Should -Be 'Lizard-Tech'
            $config.Revision | Should -Be 42
            $config.ProfileName | Should -Be 'Standard'
            ($config.Apps -join ',') | Should -Be 'TeamViewer,AV,Chrome,7zip'
            ($config.Locked -join ',') | Should -Be 'JoinDomain'
            $config.Tasks.EnableBitLocker | Should -BeFalse
            $config.Tasks.JoinDomain | Should -BeTrue
            (@($config.Scripts | ForEach-Object { $_.file }) -join ',') | Should -Be 'Drukarki.ps1'
            $config.Site.id | Should -Be 'krakow'
            $config.Site.Contains('profiles') | Should -BeFalse
            $config.Site.domainJoin.domain | Should -Be 'firma.local'
            $config.Site.localAdmin.username | Should -Be 'LokalnyIT'
            ($config.Site.sourceOrder -join ',') | Should -Be 'lan,web'
            @($config.Bloatware).Count | Should -Be 3
            @($config.Uninstall).Count | Should -Be 2
            $config.SystemSettings.DisableFastStartup | Should -BeTrue
            @($config.SystemSettings.registry).Count | Should -Be 2
            (@($config.Sources.Keys) -join ',') | Should -Be 'lan,web'
            @($config.Catalog.Keys).Count | Should -Be 5
        }

        It 'Kraków / Księgowość: zadania łączone, skrypty doklejane' {
            $config = Resolve-StdEffectiveConfig -Release $script:Release -SiteId 'krakow' -ProfileName 'Księgowość'
            $config.ProfileName | Should -Be 'Księgowość'
            $config.Tasks.EnableBitLocker | Should -BeTrue
            $config.Tasks.WaitForNetwork | Should -BeTrue
            $config.Tasks.RebootAtEnd | Should -BeFalse
            @($config.Scripts).Count | Should -Be 2
            $config.Scripts[0].file | Should -Be 'Drukarki.ps1'
            $config.Scripts[1].file | Should -Be 'Ksiegowosc.ps1'
            $config.Scripts[1].args | Should -Be '-Tryb pelny'
            @($config.Locked).Count | Should -Be 0
            $config.Site.bitLocker.escrow -join ',' | Should -Be 'ad,server'
        }

        It 'Warszawa / Standard' {
            $config = Resolve-StdEffectiveConfig -Release $script:Release -SiteId 'warszawa' -ProfileName 'Standard'
            ($config.Apps -join ',') | Should -Be 'Chrome,7zip,Git'
            $config.Site.computerName.editable | Should -BeFalse
            , $config.Site.sourceOrder | Should -BeOfType ([object[]])
            ($config.Site.sourceOrder -join ',') | Should -Be 'web'
            $config.Site.Contains('wifi') | Should -BeFalse
            $config.Tasks.EnableBitLocker | Should -BeFalse
            @($config.Locked).Count | Should -Be 0
            (@($config.Scripts | ForEach-Object { $_.file }) -join ',') | Should -Be 'Drukarki.ps1'
        }

        It 'wynik jest kopią - zmiana nie dotyka konfiguracji' {
            $config = Resolve-StdEffectiveConfig -Release $script:Release -SiteId 'krakow' -ProfileName 'Standard'
            $config.Tasks['EnableBitLocker'] = $true
            $config.Site.domainJoin['domain'] = 'inna.local'
            $config.Catalog['7zip']['name'] = 'zmienione'
            $script:Release.defaults.tasks.EnableBitLocker | Should -BeFalse
            $script:Release.sites[0].domainJoin.domain | Should -Be 'firma.local'
            $script:Release.catalog['7zip'].name | Should -Be '7-Zip'
        }

        It 'ustawienia lokalizacji są między defaults a profilem' {
            $doc = Get-TestRelease
            $doc.sites[1]['tasks'] = [ordered]@{ RebootAtEnd = $true; EnableBitLocker = $true }
            $doc.sites[1]['bloatware'] = @('*Solitaire*')
            $doc.sites[1].profiles.Standard['tasks'] = [ordered]@{ EnableBitLocker = $false }
            (Get-TestErrors $doc).Count | Should -Be 0
            $config = Resolve-StdEffectiveConfig -Release $doc -SiteId 'warszawa' -ProfileName 'Standard'
            $config.Tasks.RebootAtEnd | Should -BeTrue
            $config.Tasks.EnableBitLocker | Should -BeFalse
            ($config.Bloatware -join ',') | Should -Be '*Solitaire*'
        }

        It 'nieznana lokalizacja -> [STD-CFG-SITE], nieznany profil -> [STD-CFG-PROFILE]' {
            { Resolve-StdEffectiveConfig -Release $script:Release -SiteId 'gdansk' -ProfileName 'Standard' } | Should -Throw -ExpectedMessage '*`[STD-CFG-SITE`]*'
            { Resolve-StdEffectiveConfig -Release $script:Release -SiteId 'warszawa' -ProfileName 'Księgowość' } | Should -Throw -ExpectedMessage '*`[STD-CFG-PROFILE`]*'
        }
    }

    Context 'Get-StdRelease' {
        BeforeEach {
            $script:Work = Join-Path $TestDrive ('rel-' + [guid]::NewGuid().ToString('N'))
            $env:STD_DATA_ROOT = Join-Path $script:Work 'data'
            $script:Mirror = Join-Path $script:Work 'mirror'
            $script:Usb = Join-Path $script:Work 'usb'
            New-Item -ItemType Directory -Path $script:Mirror, $script:Usb -Force | Out-Null
            $script:Bootstrap = @{ server = $script:UnreachableServer; mirrors = @($script:Mirror); auth = @{ mode = 'basic' }; proxy = $null }
            $script:MirrorFile = Join-TestPath $script:Mirror @('clients', 'lizard', 'release.sdt')
            $script:CacheFile = Join-TestPath $env:STD_DATA_ROOT @('Cache', 'lizard', 'release.sdt')
            $script:UsbFile = Join-TestPath $script:Usb @('offline', 'lizard', 'release.sdt')
        }

        It 'pierwsze poprawne źródło (kopia lustrzana) wygrywa i trafia do cache' {
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:MirrorFile
            $result = Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted -UsbRoot $script:Usb
            $result.Source | Should -Be 'mirror'
            $result.Revision | Should -Be 42
            $result.Kid | Should -Be $script:Cert.Thumbprint.ToUpperInvariant()
            $result.Release.client.id | Should -Be 'lizard'
            $result.Release | Should -BeOfType ([System.Collections.Specialized.OrderedDictionary])
            @($result.Attempts).Count | Should -Be 2
            $result.Attempts[0].Source | Should -Be 'server'
            $result.Attempts[0].Error | Should -BeLike '*`[STD-HTTP-NET`]*'
            $result.Attempts[1].Source | Should -Be 'mirror'
            $result.Attempts[1].Error | Should -BeNullOrEmpty
            Test-Path -LiteralPath $script:CacheFile | Should -BeTrue
            [System.IO.File]::ReadAllText($script:CacheFile) | Should -Be ([System.IO.File]::ReadAllText($script:MirrorFile))
            Get-StdSeenRevision -Key 'lizard' | Should -Be 42
        }

        It 'bez serwera i kopii lustrzanej używa poprawnej kopii z cache' {
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:CacheFile
            $result = Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted
            $result.Source | Should -Be 'cache'
            $result.Attempts[1].Error | Should -BeLike '*`[STD-SRC-NOTFOUND`]*'
        }

        It 'zmieniona kopia lustrzana odpada, wygrywa kopia z cache' {
            $good = Get-TestRelease
            $evil = Get-TestRelease
            $evil.catalog['7zip']['args'] = '/S /D=C:\Zle'
            Write-TestTamperedEnvelope -Document $good -OtherDocument $evil -Path $script:MirrorFile
            Write-TestEnvelope -Document $good -Path $script:CacheFile
            $result = Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted
            $result.Source | Should -Be 'cache'
            $result.Release.catalog['7zip'].args | Should -Be '/S'
            ($result.Attempts | Where-Object { $_.Source -eq 'mirror' }).Error | Should -BeLike '*`[STD-SIG-INVALID`]*'
        }

        It 'podpis certyfikatem spoza listy zaufanych odpada, wygrywa pendrive (i trafia do cache)' {
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:MirrorFile -Certificate $script:OtherCert
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:UsbFile
            $result = Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted -UsbRoot $script:Usb
            $result.Source | Should -Be 'usb'
            ($result.Attempts | Where-Object { $_.Source -eq 'mirror' }).Error | Should -BeLike '*`[STD-SIG-UNTRUSTED`]*'
            Test-Path -LiteralPath $script:CacheFile | Should -BeTrue
        }

        It 'starsza rewizja niż widziana -> [STD-CFG-ROLLBACK]; ta sama rewizja jest przyjmowana' {
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:MirrorFile
            Set-StdSeenRevision -Key 'lizard' -Revision 50
            { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-ROLLBACK`]*'
            Get-StdSeenRevision -Key 'lizard' | Should -Be 50
            $same = Get-TestRelease
            $same['revision'] = 50
            Write-TestEnvelope -Document $same -Path $script:MirrorFile
            (Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted).Revision | Should -Be 50
        }

        It 'nowsza rewizja z kopii lustrzanej podmienia starszą kopię w cache' {
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:CacheFile
            $newer = Get-TestRelease
            $newer['revision'] = 43
            Write-TestEnvelope -Document $newer -Path $script:MirrorFile
            (Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted).Revision | Should -Be 43
            Remove-Item -LiteralPath $script:MirrorFile
            $fromCache = Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted
            $fromCache.Source | Should -Be 'cache'
            $fromCache.Revision | Should -Be 43
        }

        It 'po terminie ważności -> [STD-CFG-EXPIRED] (z tolerancją 5 minut)' {
            $expired = Get-TestRelease
            $expired['expiresAt'] = Get-TestUtcText ([datetime]::UtcNow.AddMinutes(-10))
            Write-TestEnvelope -Document $expired -Path $script:MirrorFile
            { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-EXPIRED`]*'
            $grace = Get-TestRelease
            $grace['expiresAt'] = Get-TestUtcText ([datetime]::UtcNow.AddMinutes(-2))
            Write-TestEnvelope -Document $grace -Path $script:MirrorFile
            (Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted).Source | Should -Be 'mirror'
        }

        It 'bez expiresAt ważność to data publikacji + 90 dni' {
            $old = Get-TestRelease
            $old.Remove('expiresAt')
            $old['publishedAt'] = Get-TestUtcText ([datetime]::UtcNow.AddDays(-91))
            Write-TestEnvelope -Document $old -Path $script:MirrorFile
            { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-EXPIRED`]*'
        }

        It 'minClientVersion wyższy niż wersja STD -> [STD-CFG-CLIENT-OLD]; równy jest przyjmowany' {
            $tooNew = Get-TestRelease
            $tooNew['minClientVersion'] = '99.0.0'
            Write-TestEnvelope -Document $tooNew -Path $script:MirrorFile
            { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-CLIENT-OLD`]*'
            $equal = Get-TestRelease
            $equal['minClientVersion'] = (Get-StdVersion).ToString()
            Write-TestEnvelope -Document $equal -Path $script:MirrorFile
            (Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted).Source | Should -Be 'mirror'
        }

        It 'allowOffline = false blokuje cache i pendrive -> [STD-CFG-OFFLINE]' {
            $online = Get-TestRelease
            $online['allowOffline'] = $false
            Write-TestEnvelope -Document $online -Path $script:CacheFile
            Write-TestEnvelope -Document $online -Path $script:UsbFile
            $caught = Get-TestCaught { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted -UsbRoot $script:Usb }
            $caught | Should -Not -BeNullOrEmpty
            $caught.Exception.Message | Should -BeLike '*`[STD-CFG-UNAVAILABLE`]*'
            $caught.Exception.Message | Should -BeLike '*`[STD-CFG-OFFLINE`]*'
            $attempts = @($caught.TargetObject)
            ($attempts | ForEach-Object { $_.Source }) -join ',' | Should -Be 'server,mirror,cache,usb'
            ($attempts | Where-Object { $_.Source -eq 'cache' }).Error | Should -BeLike '*`[STD-CFG-OFFLINE`]*'
            ($attempts | Where-Object { $_.Source -eq 'usb' }).Error | Should -BeLike '*`[STD-CFG-OFFLINE`]*'
        }

        It 'allowOffline = false nie blokuje kopii lustrzanej' {
            $online = Get-TestRelease
            $online['allowOffline'] = $false
            Write-TestEnvelope -Document $online -Path $script:MirrorFile
            (Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted).Source | Should -Be 'mirror'
        }

        It 'konfiguracja innego klienta w miejscu klienta lizard jest odrzucana' {
            $other = Get-TestRelease
            $other.client['id'] = 'inny-klient'
            Write-TestEnvelope -Document $other -Path $script:MirrorFile
            $caught = Get-TestCaught { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted }
            $caught.Exception.Message | Should -BeLike '*`[STD-CFG-UNAVAILABLE`]*'
            ($caught.TargetObject | Where-Object { $_.Source -eq 'mirror' }).Error | Should -BeLike "*``[STD-CFG-INVALID``]*'inny-klient'*"
        }

        It 'dokument z błędami walidacji -> [STD-CFG-INVALID] z opisem błędu' {
            $broken = Get-TestRelease
            $broken.sites[1].profiles.Standard['apps'] = @('Chrome', 'Firefox')
            Write-TestEnvelope -Document $broken -Path $script:MirrorFile
            $caught = Get-TestCaught { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted }
            ($caught.TargetObject | Where-Object { $_.Source -eq 'mirror' }).Error | Should -BeLike '*`[STD-CFG-INVALID`]*sites`[1`].profiles.Standard.apps`[1`]*'
        }

        It 'brak jakiegokolwiek źródła -> [STD-CFG-UNAVAILABLE]' {
            { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted -UsbRoot $script:Usb } | Should -Throw -ExpectedMessage '*`[STD-CFG-UNAVAILABLE`]*'
        }

        It 'pusta lista zaufanych certyfikatów -> [STD-SIG-NOTRUST]' {
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:MirrorFile
            { Get-StdRelease -ClientId 'lizard' -Bootstrap $script:Bootstrap -TrustedCertificates @() } | Should -Throw -ExpectedMessage '*`[STD-SIG-NOTRUST`]*'
        }

        It 'nieprawidłowy identyfikator klienta -> [STD-CFG-ARGS]' {
            { Get-StdRelease -ClientId '..\lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-ARGS`]*'
            { Get-StdRelease -ClientId 'Lizard' -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-ARGS`]*'
        }
    }

    Context 'Get-StdClientIndex' {
        BeforeEach {
            $script:Work = Join-Path $TestDrive ('idx-' + [guid]::NewGuid().ToString('N'))
            $env:STD_DATA_ROOT = Join-Path $script:Work 'data'
            $script:Mirror = Join-Path $script:Work 'mirror'
            $script:Usb = Join-Path $script:Work 'usb'
            New-Item -ItemType Directory -Path $script:Mirror, $script:Usb -Force | Out-Null
            $script:Bootstrap = @{ server = $script:UnreachableServer; mirrors = @($script:Mirror); auth = @{ mode = 'basic' }; proxy = $null }
            $script:IndexMirrorFile = Join-Path $script:Mirror 'clients.sdt'
            $script:IndexCacheFile = Join-TestPath $env:STD_DATA_ROOT @('Cache', 'clients.sdt')
        }

        It 'pobiera indeks z kopii lustrzanej, zapisuje w cache i rewizję pod kluczem clients' {
            Write-TestEnvelope -Document (Get-TestIndex) -Path $script:IndexMirrorFile
            $result = Get-StdClientIndex -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted
            $result.Source | Should -Be 'mirror'
            $result.Revision | Should -Be 7
            $result.Index.clients[0].id | Should -Be 'lizard'
            $result.Index.clients[0].release | Should -Be 'clients/lizard/release.sdt'
            Test-Path -LiteralPath $script:IndexCacheFile | Should -BeTrue
            Get-StdSeenRevision -Key 'clients' | Should -Be 7
        }

        It 'kopia z cache i z pendrive''a jest dozwolona zawsze' {
            Write-TestEnvelope -Document (Get-TestIndex) -Path (Join-TestPath $script:Usb @('offline', 'clients.sdt'))
            (Get-StdClientIndex -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted -UsbRoot $script:Usb).Source | Should -Be 'usb'
            (Get-StdClientIndex -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted).Source | Should -Be 'cache'
        }

        It 'starsza rewizja indeksu -> [STD-CFG-ROLLBACK]' {
            Write-TestEnvelope -Document (Get-TestIndex -Revision 7) -Path $script:IndexMirrorFile
            Set-StdSeenRevision -Key 'clients' -Revision 8
            { Get-StdClientIndex -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-ROLLBACK`]*'
        }

        It 'konfiguracja klienta zamiast indeksu -> [STD-CFG-INVALID]' {
            Write-TestEnvelope -Document (Get-TestRelease) -Path $script:IndexMirrorFile
            { Get-StdClientIndex -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-CFG-INVALID`]*'
        }

        It 'zmieniony indeks -> [STD-SIG-INVALID]' {
            $other = Get-TestIndex
            $other.clients[0].release = 'clients/obcy/release.sdt'
            Write-TestTamperedEnvelope -Document (Get-TestIndex) -OtherDocument $other -Path $script:IndexMirrorFile
            { Get-StdClientIndex -Bootstrap $script:Bootstrap -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-INVALID`]*'
        }
    }
}
