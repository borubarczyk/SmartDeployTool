# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy narzędzi Tools\Publish-StdRelease.ps1 i Tools\New-StdSigningCertificate.ps1 (Pester 5, Windows PowerShell 5.1).
# Bez administratora, sieci i magazynu certyfikatów: certyfikaty testowe powstają w pamięci (CertificateRequest)
# i są podawane przez -Certificate, a miejsce docelowe to katalog w TestDrive (tak jak katalog std\ serwera
# albo kopia lustrzana UNC). Publikacji przez https:// (WebDAV) nie da się tu sprawdzić bez serwera.
# Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' (nawiasy w -like trzeba poprzedzić znakiem `).

Describe 'Publish-StdRelease' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'

        $script:ToolsDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'Tools'
        $script:Publish = Join-Path $script:ToolsDir 'Publish-StdRelease.ps1'
        $script:FixturePath = Join-Path (Join-Path $PSScriptRoot 'Fixtures') 'release.v4.json'
        $script:Utf8 = New-Object System.Text.UTF8Encoding($false)
        $script:CredentialText = 'Dane logowania do domeny i hasło lokalnego administratora wpisuje się ręcznie przy wdrożeniu - nie mogą być w konfiguracji.'

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

        function New-TestDestination {
            # Nowy (jeszcze nieistniejący) katalog drzewa std\ serwera w TestDrive.
            return (Join-TestPath $TestDrive @(('srv-' + [guid]::NewGuid().ToString('N').Substring(0, 8)), 'std'))
        }

        function Get-TestRelease {
            # Świeża kopia wspólnego przykładu konfiguracji (Tests\Fixtures\release.v4.json, revision 42).
            return (ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText($script:FixturePath, $script:Utf8)))
        }

        function Write-TestSource {
            # Zapisuje dokument jako plik źródłowy (JSON bez podpisu) i zwraca ścieżkę.
            param($Document)
            $path = Join-Path $TestDrive ('zrodlo-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.json')
            [System.IO.File]::WriteAllText($path, (ConvertTo-StdJson -InputObject $Document), $script:Utf8)
            return $path
        }

        function Invoke-TestPublish {
            # Publikacja bez ostrzeżeń (certyfikat testowy nie jest w Std.TrustedSigners.psd1) i bez Write-Host.
            param([hashtable]$Arguments)
            return (& $script:Publish @Arguments -WarningAction SilentlyContinue 6>$null)
        }

        function Read-TestEnvelope {
            # Sprawdza podpis pliku *.sdt certyfikatem testowym i zwraca @{ Envelope; Document }.
            param([string]$Path)
            $envelope = Read-StdSignedEnvelope -Bytes ([System.IO.File]::ReadAllBytes($Path)) -TrustedCertificates @($script:Trusted)
            return @{ Envelope = $envelope; Document = (ConvertFrom-StdJson -Json $envelope.PayloadText) }
        }

        function Get-TestError {
            param([scriptblock]$Block)
            try {
                & $Block | Out-Null
            } catch {
                return $_
            }
            return $null
        }

        function Get-TestTreeHash {
            # Odcisk całego drzewa (ścieżki i SHA-256 plików) - do sprawdzenia, że nic się nie zmieniło.
            param([string]$Root)
            if (-not (Test-Path -LiteralPath $Root)) { return '' }
            $lines = foreach ($file in @(Get-ChildItem -LiteralPath $Root -Recurse -File | Sort-Object FullName)) {
                '{0}={1}' -f $file.FullName, (Get-StdSha256 -Path $file.FullName)
            }
            return (@($lines) -join "`n")
        }

        $script:Cert = New-TestCertificate -Subject 'CN=STD test - publikacja'
        $script:OtherCert = New-TestCertificate -Subject 'CN=STD test - obcy podpis'
        $script:Trusted = Get-TestPublicCertificate $script:Cert
    }

    AfterAll {
        Remove-Item Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
    }

    Context 'Parametry' {
        It 'ma parametry z SPEC 4.2 i obsługuje -WhatIf' {
            $parameters = (Get-Command -Name $script:Publish).Parameters
            foreach ($name in @('SourcePath', 'Certificate', 'CertificateThumbprint', 'Destination', 'Credential', 'ChangeNote', 'ExpiresInDays', 'PublishedBy', 'WhatIf', 'Confirm')) {
                $parameters.ContainsKey($name) | Should -BeTrue -Because "parametr -$name"
            }
            $parameters['Certificate'].ParameterType | Should -Be ([System.Security.Cryptography.X509Certificates.X509Certificate2])
            $parameters['Credential'].ParameterType | Should -Be ([System.Management.Automation.PSCredential])
        }
    }

    Context 'Pierwsza publikacja do katalogu' {
        BeforeAll {
            $script:Dest = New-TestDestination
            $script:Result = Invoke-TestPublish @{
                SourcePath  = $script:FixturePath
                Certificate = $script:Cert
                Destination = $script:Dest
                ChangeNote  = 'Test publikacji'
                PublishedBy = 'tester'
            }
            $script:ReleaseFile = Join-TestPath $script:Dest @('clients', 'lizard', 'release.sdt')
            $script:HistoryFile = Join-TestPath $script:Dest @('clients', 'lizard', 'history', '43.sdt')
            $script:IndexFile = Join-Path $script:Dest 'clients.sdt'
        }

        It 'tworzy pliki w ścieżkach z SPEC 2.4 (release.sdt, history\<rev>.sdt, clients.sdt)' {
            Test-Path -LiteralPath $script:ReleaseFile -PathType Leaf | Should -BeTrue
            Test-Path -LiteralPath $script:HistoryFile -PathType Leaf | Should -BeTrue
            Test-Path -LiteralPath $script:IndexFile -PathType Leaf | Should -BeTrue
            @(Get-ChildItem -LiteralPath $script:Dest -Recurse -File).Count | Should -Be 3
        }

        It 'zwraca wynik @{ ClientId; Revision; ReleasePath; IndexRevision; Sha256 }' {
            $script:Result.ClientId | Should -Be 'lizard'
            $script:Result.Revision | Should -Be 43
            $script:Result.IndexRevision | Should -Be 1
            $script:Result.ReleasePath | Should -Be $script:ReleaseFile
            $script:Result.Sha256 | Should -Be (Get-StdSha256 -Path $script:ReleaseFile)
        }

        It 'konfiguracja ma poprawny podpis (Read-StdSignedEnvelope) i przechodzi Test-StdRelease' {
            $read = Read-TestEnvelope $script:ReleaseFile
            $read.Envelope.Kid | Should -Be $script:Cert.Thumbprint.ToUpperInvariant()
            @(Test-StdRelease -Release $read.Document).Count | Should -Be 0
            $read.Document['kind'] | Should -Be 'std.release'
            $read.Document['client']['id'] | Should -Be 'lizard'
        }

        It 'ustawia revision (z pliku źródłowego 42 + 1), publishedAt, publishedBy, expiresAt i changeNote' {
            $document = (Read-TestEnvelope $script:ReleaseFile).Document
            $document['revision'] | Should -Be 43
            $document['publishedBy'] | Should -Be 'tester'
            $document['changeNote'] | Should -Be 'Test publikacji'
            $document['publishedAt'] | Should -Match 'Z$'
            $published = [System.DateTimeOffset]::Parse($document['publishedAt'], [System.Globalization.CultureInfo]::InvariantCulture)
            $expires = [System.DateTimeOffset]::Parse($document['expiresAt'], [System.Globalization.CultureInfo]::InvariantCulture)
            [math]::Abs(([System.DateTimeOffset]::UtcNow - $published).TotalMinutes) | Should -BeLessThan 10
            [math]::Round(($expires - $published).TotalDays, 3) | Should -Be 90
        }

        It 'zachowuje resztę dokumentu bez zmian' {
            $document = (Read-TestEnvelope $script:ReleaseFile).Document
            $source = Get-TestRelease
            foreach ($key in @('sources', 'catalog', 'defaults', 'sites', 'secrets', 'allowOffline', 'minClientVersion')) {
                (ConvertTo-StdJson -InputObject $document[$key] -Compress) | Should -Be (ConvertTo-StdJson -InputObject $source[$key] -Compress) -Because "pole $key"
            }
        }

        It 'archiwum history\43.sdt ma te same bajty co release.sdt' {
            Get-StdSha256 -Path $script:HistoryFile | Should -Be (Get-StdSha256 -Path $script:ReleaseFile)
        }

        It 'lista klientów clients.sdt jest podpisana i zawiera klienta (rewizja 1)' {
            $read = Read-TestEnvelope $script:IndexFile
            $index = $read.Document
            $index['kind'] | Should -Be 'std.clients'
            $index['schemaVersion'] | Should -Be 4
            $index['revision'] | Should -Be 1
            @($index['clients']).Count | Should -Be 1
            $index['clients'][0]['id'] | Should -Be 'lizard'
            $index['clients'][0]['name'] | Should -Be 'Lizard-Tech'
            $index['clients'][0]['release'] | Should -Be 'clients/lizard/release.sdt'
            $read.Envelope.PayloadText | Should -Match '"clients":\s*\['
        }

        It 'pliki są w UTF-8 bez BOM i nie zostają pliki tymczasowe' {
            foreach ($path in @($script:ReleaseFile, $script:IndexFile)) {
                $bytes = [System.IO.File]::ReadAllBytes($path)
                ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should -BeFalse
            }
            @(Get-ChildItem -LiteralPath $script:Dest -Recurse -File -Filter '*.tmp').Count | Should -Be 0
        }
    }

    Context 'Kolejne publikacje' {
        It 'druga publikacja zwiększa rewizję, a archiwum zachowuje obie' {
            $dest = New-TestDestination
            $first = Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest }
            $second = Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest }
            $first.Revision | Should -Be 43
            $second.Revision | Should -Be 44
            $history = Join-TestPath $dest @('clients', 'lizard', 'history')
            (Read-TestEnvelope (Join-Path $history '43.sdt')).Document['revision'] | Should -Be 43
            (Read-TestEnvelope (Join-Path $history '44.sdt')).Document['revision'] | Should -Be 44
            (Read-TestEnvelope (Join-TestPath $dest @('clients', 'lizard', 'release.sdt'))).Document['revision'] | Should -Be 44
        }

        It 'rewizja listy klientów rośnie przy każdej publikacji, klient jest na liście raz' {
            $dest = New-TestDestination
            $first = Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest }
            $second = Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest }
            $first.IndexRevision | Should -Be 1
            $second.IndexRevision | Should -Be 2
            $index = (Read-TestEnvelope (Join-Path $dest 'clients.sdt')).Document
            $index['revision'] | Should -Be 2
            @($index['clients']).Count | Should -Be 1
            $index['clients'][0]['id'] | Should -Be 'lizard'
        }

        It 'drugi klient jest dopisany do listy, pierwszy zostaje (także po ponownej publikacji)' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } | Out-Null
            $acme = Get-TestRelease
            $acme['client']['id'] = 'acme'
            $acme['client']['name'] = 'ACME sp. z o.o.'
            $acme.Remove('revision')
            $acmeResult = Invoke-TestPublish @{ SourcePath = (Write-TestSource $acme); Certificate = $script:Cert; Destination = $dest }
            $acmeResult.Revision | Should -Be 1
            $acmeResult.IndexRevision | Should -Be 2
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } | Out-Null

            $index = (Read-TestEnvelope (Join-Path $dest 'clients.sdt')).Document
            $index['revision'] | Should -Be 3
            @($index['clients'] | ForEach-Object { $_['id'] }) -join ',' | Should -Be 'lizard,acme'
            ($index['clients'] | Where-Object { $_['id'] -eq 'acme' })['release'] | Should -Be 'clients/acme/release.sdt'
            Test-Path -LiteralPath (Join-TestPath $dest @('clients', 'acme', 'release.sdt')) | Should -BeTrue
        }

        It 'bez -ChangeNote zostaje opis z pliku źródłowego; -ExpiresInDays ustawia ważność' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest; ExpiresInDays = 10 } | Out-Null
            $document = (Read-TestEnvelope (Join-TestPath $dest @('clients', 'lizard', 'release.sdt'))).Document
            $document['changeNote'] | Should -Be 'Przykładowa konfiguracja do testów'
            $published = [System.DateTimeOffset]::Parse($document['publishedAt'], [System.Globalization.CultureInfo]::InvariantCulture)
            $expires = [System.DateTimeOffset]::Parse($document['expiresAt'], [System.Globalization.CultureInfo]::InvariantCulture)
            [math]::Round(($expires - $published).TotalDays, 3) | Should -Be 10
        }

        It 'numer rewizji zajęty w archiwum jest pomijany (np. po usunięciu release.sdt)' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } | Out-Null
            Remove-Item -LiteralPath (Join-TestPath $dest @('clients', 'lizard', 'release.sdt')) -Force
            $result = Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest }
            $result.Revision | Should -Be 44
            (Read-TestEnvelope (Join-TestPath $dest @('clients', 'lizard', 'history', '43.sdt'))).Document['revision'] | Should -Be 43
        }
    }

    Context 'Odczyt przez STD na komputerze' {
        It 'Get-StdRelease i Get-StdClientIndex czytają opublikowane drzewo jako kopię lustrzaną' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } | Out-Null
            $previousRoot = $env:STD_DATA_ROOT
            $env:STD_DATA_ROOT = Join-Path $TestDrive ('komputer-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            try {
                $bootstrap = @{ mirrors = @($dest) }
                $release = Get-StdRelease -ClientId 'lizard' -Bootstrap $bootstrap -TrustedCertificates @($script:Trusted)
                $release.Source | Should -Be 'mirror'
                $release.Revision | Should -Be 43
                $release.Kid | Should -Be $script:Cert.Thumbprint.ToUpperInvariant()
                $release.Release['publishedBy'] | Should -Not -BeNullOrEmpty

                $index = Get-StdClientIndex -Bootstrap $bootstrap -TrustedCertificates @($script:Trusted)
                $index.Source | Should -Be 'mirror'
                $index.Revision | Should -Be 1
                $index.Index['clients'][0]['id'] | Should -Be 'lizard'
            } finally {
                $env:STD_DATA_ROOT = $previousRoot
            }
        }
    }

    Context 'Błędy' {
        It 'nieprawidłowy plik źródłowy: przerwanie z listą błędów i nic nie jest zapisane' {
            $bad = Get-TestRelease
            $bad['sites'][0]['domainJoin']['password'] = 'Tajne-Haslo-123'
            $bad['catalog']['Zly'] = [ordered]@{ name = 'Bez instalatora' }
            $source = Write-TestSource $bad
            $dest = New-TestDestination

            $caught = Get-TestError { Invoke-TestPublish @{ SourcePath = $source; Certificate = $script:Cert; Destination = $dest } }
            $caught | Should -Not -BeNullOrEmpty
            $message = $caught.Exception.Message
            $message | Should -BeLike '*`[STD-PUB-INVALID`]*'
            $message | Should -BeLike "*$($script:CredentialText)*"
            $message | Should -BeLike '*catalog.Zly*'
            $message | Should -Not -BeLike '*Tajne-Haslo-123*'
            @($caught.TargetObject).Count | Should -BeGreaterThan 1
            Test-Path -LiteralPath (Split-Path -Parent $dest) | Should -BeFalse
        }

        It 'nieprawidłowy plik źródłowy nie zmienia opublikowanej konfiguracji' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } | Out-Null
            $before = Get-TestTreeHash $dest
            $bad = Get-TestRelease
            $bad['defaults']['tasks']['NieznaneZadanie'] = $true
            { Invoke-TestPublish @{ SourcePath = (Write-TestSource $bad); Certificate = $script:Cert; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-INVALID`]*'
            Get-TestTreeHash $dest | Should -Be $before
        }

        It 'brak pliku, błędny JSON albo podpisana koperta jako źródło -> [STD-PUB-SOURCE]' {
            $dest = New-TestDestination
            $broken = Join-Path $TestDrive 'zepsuty.json'
            [System.IO.File]::WriteAllText($broken, '{ "schemaVersion": 4, ', $script:Utf8)
            { Invoke-TestPublish @{ SourcePath = (Join-Path $TestDrive 'nie-ma.json'); Certificate = $script:Cert; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-SOURCE`]*'
            { Invoke-TestPublish @{ SourcePath = $broken; Certificate = $script:Cert; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-SOURCE`]*'

            $published = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $published } | Out-Null
            $envelopeFile = Join-TestPath $published @('clients', 'lizard', 'release.sdt')
            { Invoke-TestPublish @{ SourcePath = $envelopeFile; Certificate = $script:Cert; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-SOURCE`]*'
            Test-Path -LiteralPath (Split-Path -Parent $dest) | Should -BeFalse
        }

        It 'opublikowana konfiguracja z nieznanym podpisem -> [STD-PUB-PREVIOUS], nic nie jest zmienione' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:OtherCert; Destination = $dest } | Out-Null
            $before = Get-TestTreeHash $dest
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-PREVIOUS`]*`[STD-SIG-UNTRUSTED`]*'
            Get-TestTreeHash $dest | Should -Be $before
        }

        It 'zmieniona (uszkodzona) opublikowana konfiguracja -> [STD-PUB-PREVIOUS]' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } | Out-Null
            $releaseFile = Join-TestPath $dest @('clients', 'lizard', 'release.sdt')
            $envelope = ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText($releaseFile, $script:Utf8))
            $changed = $script:Utf8.GetString([System.Convert]::FromBase64String($envelope['payload'])).Replace('"Lizard-Tech"', '"Lizard-Teck"')
            $envelope['payload'] = [System.Convert]::ToBase64String($script:Utf8.GetBytes($changed))
            [System.IO.File]::WriteAllText($releaseFile, (ConvertTo-StdJson -InputObject $envelope), $script:Utf8)
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-PREVIOUS`]*`[STD-SIG-INVALID`]*'
        }

        It 'lista klientów z nieznanym podpisem -> [STD-PUB-INDEX], konfiguracja nie jest zapisana' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:OtherCert; Destination = $dest } | Out-Null
            Remove-Item -LiteralPath (Join-Path $dest 'clients') -Recurse -Force
            $before = Get-TestTreeHash $dest
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-INDEX`]*'
            Get-TestTreeHash $dest | Should -Be $before
            Test-Path -LiteralPath (Join-Path $dest 'clients') | Should -BeFalse
        }

        It 'certyfikat bez klucza prywatnego albo zły odcisk -> [STD-PUB-CERT]' {
            $dest = New-TestDestination
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Trusted; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-CERT`]*'
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; CertificateThumbprint = 'to-nie-odcisk'; Destination = $dest } } | Should -Throw -ExpectedMessage '*`[STD-PUB-CERT`]*'
            Test-Path -LiteralPath (Split-Path -Parent $dest) | Should -BeFalse
        }

        It 'katalog danych serwera zamiast podkatalogu std -> [STD-PUB-DEST]' {
            $dataPath = Join-Path $TestDrive 'STD-Server-data'
            New-Item -ItemType Directory -Path (Join-TestPath $dataPath @('std', 'clients')) -Force | Out-Null
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dataPath } } | Should -Throw -ExpectedMessage '*`[STD-PUB-DEST`]*'
            Test-Path -LiteralPath (Join-Path $dataPath 'clients.sdt') | Should -BeFalse
        }

        It 'adres http:// -> [STD-PUB-DEST]; https:// bez -Credential -> [STD-PUB-ARGS]' {
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = 'http://pobieranie.example.com/std-admin/' } } | Should -Throw -ExpectedMessage '*`[STD-PUB-DEST`]*'
            { Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = 'https://pobieranie.example.com/std-admin/' } } | Should -Throw -ExpectedMessage '*`[STD-PUB-ARGS`]*'
        }
    }

    Context '-WhatIf' {
        It 'nie tworzy katalogów ani plików i nic nie zwraca' {
            $dest = New-TestDestination
            $result = Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest; WhatIf = $true }
            $result | Should -BeNullOrEmpty
            Test-Path -LiteralPath (Split-Path -Parent $dest) | Should -BeFalse
        }

        It 'nie zmienia istniejącej publikacji' {
            $dest = New-TestDestination
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest } | Out-Null
            $before = Get-TestTreeHash $dest
            Invoke-TestPublish @{ SourcePath = $script:FixturePath; Certificate = $script:Cert; Destination = $dest; WhatIf = $true } | Out-Null
            Get-TestTreeHash $dest | Should -Be $before
        }
    }
}

Describe 'New-StdSigningCertificate' {
    BeforeAll {
        $script:NewCertificate = Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'Tools') 'New-StdSigningCertificate.ps1'
    }

    It 'ma parametry z SPEC 4.1 i ich wartości domyślne' {
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:NewCertificate, [ref]$tokens, [ref]$parseErrors)
        @($parseErrors).Count | Should -Be 0
        $parameters = @{}
        foreach ($parameter in $ast.ParamBlock.Parameters) { $parameters[$parameter.Name.VariablePath.UserPath] = $parameter }
        $parameters['Name'].DefaultValue.Value | Should -Be 'STD - podpis konfiguracji'
        $parameters['Years'].DefaultValue.Value | Should -Be 5
        $parameters['Exportable'].StaticType | Should -Be ([System.Management.Automation.SwitchParameter])
        $parameters['OutPath'].DefaultValue.Value | Should -Be '.'
    }

    It 'tworzy RSA 3072 / SHA-256 z EKU Document Signing w Cert:\CurrentUser\My, klucz domyślnie nieeksportowalny' {
        $text = [System.IO.File]::ReadAllText($script:NewCertificate)
        $text | Should -Match 'KeyLength\s*=\s*3072'
        $text | Should -Match "HashAlgorithm\s*=\s*'SHA256'"
        $text | Should -Match "KeyUsage\s*=\s*'DigitalSignature'"
        $text | Should -Match '1\.3\.6\.1\.4\.1\.311\.10\.3\.12'
        $text | Should -Match "Cert:\\CurrentUser\\My"
        $text | Should -Match "exportPolicy = 'NonExportable'"
    }

    It 'poza Windows kończy się czytelnym błędem [STD-CERT-OS]' -Skip:([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        { & $script:NewCertificate -OutPath (Join-Path $TestDrive 'cert') 6>$null } | Should -Throw -ExpectedMessage '*`[STD-CERT-OS`]*'
        Test-Path -LiteralPath (Join-Path $TestDrive 'cert') | Should -BeFalse
    }
}
