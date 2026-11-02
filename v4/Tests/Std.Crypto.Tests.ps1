# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy Modules\Std.Crypto.ps1 (Pester 5, Windows PowerShell 5.1): SHA-256, podpisana koperta *.sdt,
# zaszyfrowane sekrety i wstawki {{secret:nazwa}}. Bez administratora, sieci i magazynu certyfikatów:
# certyfikaty testowe powstają w pamięci (CertificateRequest).
# Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' (nawiasy w -like trzeba poprzedzić znakiem `).
# Wektory "niezależne" (koperta i magazyn sekretów) wygenerowano poza STD (Python, biblioteka cryptography),
# żeby sprawdzić zgodność formatu z SPEC 2.1 i 2.7, a nie tylko zgodność STD z samym sobą.

Describe 'Std.Crypto' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'

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

        function ConvertTo-TestSecure {
            param([string]$Text)
            return (ConvertTo-SecureString -String $Text -AsPlainText -Force)
        }

        function ConvertTo-TestUtf8 {
            param([string]$Text)
            return , ((New-Object System.Text.UTF8Encoding($false)).GetBytes($Text))
        }

        function Edit-TestEnvelope {
            # Zmienia pola koperty (JSON) i zwraca nowy JSON. $Change dostaje [ordered] z polami.
            param([string]$Json, [scriptblock]$Change)
            $object = ConvertFrom-Json -InputObject $Json
            $fields = [ordered]@{}
            foreach ($property in $object.PSObject.Properties) { $fields[$property.Name] = $property.Value }
            & $Change $fields
            return (ConvertTo-Json -InputObject $fields -Depth 30 -Compress)
        }

        $script:Emoji = [char]::ConvertFromUtf32(0x1F511)
    }

    AfterAll {
        Remove-Item Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
        Remove-Item Env:\STD_DEBUG -ErrorAction SilentlyContinue
        Clear-StdLogMasks
    }

    Context 'Get-StdSha256 i Test-StdFileHash' {
        BeforeAll {
            $script:AbcHash = 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'
            $script:AbcPath = Join-Path $TestDrive 'abc.txt'
            [System.IO.File]::WriteAllBytes($script:AbcPath, [byte[]](0x61, 0x62, 0x63))
        }

        It 'liczy znany wektor SHA-256 dla "abc"' {
            Get-StdSha256 -Bytes ([byte[]](0x61, 0x62, 0x63)) | Should -Be $script:AbcHash
        }

        It 'liczy SHA-256 pustych danych' {
            Get-StdSha256 -Bytes (New-Object byte[] 0) | Should -Be 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
        }

        It 'liczy SHA-256 pliku (ścieżka pełna i TestDrive:\)' {
            Get-StdSha256 -Path $script:AbcPath | Should -Be $script:AbcHash
            Get-StdSha256 -Path 'TestDrive:\abc.txt' | Should -Be $script:AbcHash
        }

        It 'zwraca HEX małymi literami' {
            $hash = Get-StdSha256 -Bytes (ConvertTo-TestUtf8 'Zażółć gęślą jaźń')
            $hash | Should -Match '^[0-9a-f]{64}$'
            $hash | Should -BeOfType [string]
        }

        It 'brak pliku -> [STD-IO]' {
            { Get-StdSha256 -Path (Join-Path $TestDrive 'brak.bin') } | Should -Throw -ExpectedMessage '*`[STD-IO`]*'
        }

        It 'Test-StdFileHash porównuje bez względu na wielkość liter' {
            Test-StdFileHash -Path $script:AbcPath -Expected $script:AbcHash | Should -BeTrue
            Test-StdFileHash -Path $script:AbcPath -Expected $script:AbcHash.ToUpperInvariant() | Should -BeTrue
            Test-StdFileHash -Path $script:AbcPath -Expected (' ' + $script:AbcHash + ' ') | Should -BeTrue
        }

        It 'Test-StdFileHash: inna suma, brak pliku albo zła wartość -> $false' {
            Test-StdFileHash -Path $script:AbcPath -Expected ('0' * 64) | Should -BeFalse
            Test-StdFileHash -Path (Join-Path $TestDrive 'brak.bin') -Expected $script:AbcHash | Should -BeFalse
            Test-StdFileHash -Path $script:AbcPath -Expected 'abc' | Should -BeFalse
            Test-StdFileHash -Path $script:AbcPath -Expected '' | Should -BeFalse
        }
    }

    Context 'Podpisana koperta (*.sdt)' {
        BeforeAll {
            $script:Cert = New-TestCertificate -Subject 'CN=STD test - podpis'
            $script:OtherCert = New-TestCertificate -Subject 'CN=STD test - inny podpis'
            $script:Trusted = Get-TestPublicCertificate $script:Cert
            $script:OtherTrusted = Get-TestPublicCertificate $script:OtherCert
            $script:PayloadText = '{"schemaVersion":4,"kind":"std.release","client":{"id":"lizard","name":"Kraków - zażółć gęślą jaźń"}}'
            $script:Payload = ConvertTo-TestUtf8 $script:PayloadText
            $script:EnvelopeJson = New-StdSignedEnvelope -Payload $script:Payload -Certificate $script:Cert

            # Wektor niezależny: certyfikat RSA 2048 i koperta podpisana poza STD (Python, PKCS#1 v1.5 + SHA-256).
            $script:ExternalCertBase64 = 'MIICxDCCAaygAwIBAgICA+gwDQYJKoZIhvcNAQELBQAwJDEiMCAGA1UEAwwZU1REIHRlc3QgLSB3ZWt0b3IgT3BlblNTTDAgFw0yNjAxMDEwMDAwMDBaGA8yMTI1MTIwODAwMDAwMFowJDEiMCAGA1UEAwwZU1REIHRlc3QgLSB3ZWt0b3IgT3BlblNTTDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAN510JX2pwfzRsEuo7tAYWiqTsv81BMalTLjhWLd/cpv1k7KmixyyDMMbe7UHEQJ5t18FhAcSVLmLTiWUdWeLHDZPGU/ju3PqnydTPSAhomwF841QkC94fcbDuWDPSMzTUbviHY5qk+48bpt++pg30cTx2oShuOmqAxgeX/rbYiJKhbFjoYv930iyDnyedz7WWGeW6ThYHIOLPS/+RHqYrBtpDNPzCrxcw1hqchpBJpqkkCS94sdxVAeEcZMTFur+YSDpmd6gQg2JiUrxMnB6BOBGI9bsqsTcoDN5O3Pic1G+iae9J+f1AkgH0oLgBAL2qiydyU7BOqUAEzBFmV5gRUCAwEAATANBgkqhkiG9w0BAQsFAAOCAQEAxViCI76ysGu+E8nQky0ArVpNtQg+//i9fk1WtIxBiRMZTVEWP0CHjTJA2cdMydUgZPbOHa5W1DFGmnHKG2eOC/SvfDjeEQMs+bGYcDtpdE4OYginoQesO1rYx/RxdV9rdohJSLFPD5cUCfxibPcmQ1mwm8yyaeMCrhRt+P427mkzZm4Gzjrl7201ObdA3h23IEiiGOjcj1+xmZOUfjVCNUXP8JZdIvY7NkjHF7ICLeB9h3SuDJ1GLfZ3wC9fmkLJCBy4AAssQYuv7ERhO/TG1bVC29QyJ85L3aziAwOVMj+SHdbP19QMMn1F+twDa0t8OQK3y/5hdaAurhHc4qrthg=='
            $script:ExternalKid = '7F13C218087E2E2174E713B175DF128E9EA00F8E'
            $script:ExternalEnvelope = '{"kind":"std.signed","alg":"RS256","kid":"7F13C218087E2E2174E713B175DF128E9EA00F8E","payload":"eyJraW5kIjoic3RkLnJlbGVhc2UiLCJjbGllbnQiOnsiaWQiOiJsaXphcmQiLCJuYW1lIjoixYHDs2TFuiAtIHphxbzDs8WCxIcifX0=","sig":"YR4ZKsMPl7uKh8/ZqdfgQHLpZ995SmKHjxKyUZ3ZQKud0Fz19qmMoXdN/7ebiHtak81TZPI25PDV4jc7DBuMz3Bd4L6YskH9How6qfkdMIf7qje+8wfQuPuMiAoxD3g2BGNjDZfM/kG22mhkh69FfYAgu6jm9S28cGyIrCbs51rxlmeYtbETtJr20ecQStwPGqIrsAD4Wiux3n/eRZauiLV+hRdBwDIl4QyssgRv1stnyMyBuqeN+xjXauAYO118DBoP4mE138AV09n4EdL60IrCmqgdxGj3ci0GZiFeNN6jsmgI4zXT9I4xaGkl+HPijm+5LvIFjzhlUp6iKHtx8w=="}'
            $script:ExternalPayloadText = '{"kind":"std.release","client":{"id":"lizard","name":"Łódź - zażółć"}}'
        }

        It 'tworzy kopertę z polami kind, alg, kid, payload, sig' {
            $envelope = ConvertFrom-Json -InputObject $script:EnvelopeJson
            $envelope.kind | Should -BeExactly 'std.signed'
            $envelope.alg | Should -BeExactly 'RS256'
            $envelope.kid | Should -BeExactly $script:Cert.Thumbprint.ToUpperInvariant()
            $envelope.kid | Should -Match '^[0-9A-F]{40}$'
            $envelope.payload | Should -BeExactly ([System.Convert]::ToBase64String($script:Payload))
            @($envelope.PSObject.Properties).Count | Should -Be 5
        }

        It 'sig to RSA PKCS#1 v1.5 + SHA-256 dokładnych bajtów payload' {
            $envelope = ConvertFrom-Json -InputObject $script:EnvelopeJson
            $signature = [System.Convert]::FromBase64String($envelope.sig)
            $publicKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($script:Trusted)
            $publicKey.VerifyData($script:Payload, $signature, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1) | Should -BeTrue
            # PKCS#1 v1.5 jest deterministyczny: ten sam klucz i te same bajty dają ten sam podpis.
            $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($script:Cert)
            $expected = $privateKey.SignData($script:Payload, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            [System.Convert]::ToBase64String($signature) | Should -BeExactly ([System.Convert]::ToBase64String($expected))
        }

        It 'Read-StdSignedEnvelope zwraca Kid, PayloadBytes, PayloadText i Signer' {
            $result = Read-StdSignedEnvelope -Json $script:EnvelopeJson -TrustedCertificates @($script:OtherTrusted, $script:Trusted)
            $result.Kid | Should -BeExactly $script:Cert.Thumbprint.ToUpperInvariant()
            [System.Convert]::ToBase64String($result.PayloadBytes) | Should -BeExactly ([System.Convert]::ToBase64String($script:Payload))
            $result.PayloadText | Should -BeExactly $script:PayloadText
            $result.Signer.Thumbprint | Should -Be $script:Trusted.Thumbprint
        }

        It 'Read-StdSignedEnvelope przyjmuje bajty pliku (z BOM i bez)' {
            $bytes = ConvertTo-TestUtf8 $script:EnvelopeJson
            (Read-StdSignedEnvelope -Bytes $bytes -TrustedCertificates $script:Trusted).PayloadText | Should -BeExactly $script:PayloadText
            $withBom = [byte[]](@([byte]0xEF, [byte]0xBB, [byte]0xBF) + $bytes)
            (Read-StdSignedEnvelope -Bytes $withBom -TrustedCertificates $script:Trusted).PayloadText | Should -BeExactly $script:PayloadText
        }

        It 'przyjmuje kopertę podpisaną niezależnie (poza STD)' {
            $external = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String($script:ExternalCertBase64))
            $external.Thumbprint | Should -Be $script:ExternalKid
            $result = Read-StdSignedEnvelope -Json $script:ExternalEnvelope -TrustedCertificates $external
            $result.Kid | Should -BeExactly $script:ExternalKid
            $result.PayloadText | Should -BeExactly $script:ExternalPayloadText
        }

        It 'kid małymi literami jest przyjmowany, a Kid w wyniku ma wielkie litery' {
            $json = Edit-TestEnvelope -Json $script:EnvelopeJson -Change { param($f) $f.kid = ([string]$f.kid).ToLowerInvariant() }
            (Read-StdSignedEnvelope -Json $json -TrustedCertificates $script:Trusted).Kid | Should -BeExactly $script:Cert.Thumbprint.ToUpperInvariant()
        }

        It 'jeden zmieniony bajt payload -> [STD-SIG-INVALID]' {
            $json = Edit-TestEnvelope -Json $script:EnvelopeJson -Change {
                param($f)
                $bytes = [System.Convert]::FromBase64String($f.payload)
                $bytes[10] = $bytes[10] -bxor 0x01
                $f.payload = [System.Convert]::ToBase64String($bytes)
            }
            { Read-StdSignedEnvelope -Json $json -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-INVALID`]*'
        }

        It 'jeden zmieniony bajt podpisu -> [STD-SIG-INVALID]' {
            $json = Edit-TestEnvelope -Json $script:EnvelopeJson -Change {
                param($f)
                $bytes = [System.Convert]::FromBase64String($f.sig)
                $bytes[0] = $bytes[0] -bxor 0x80
                $f.sig = [System.Convert]::ToBase64String($bytes)
            }
            { Read-StdSignedEnvelope -Json $json -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-INVALID`]*'
        }

        It 'podpis ucięty -> [STD-SIG-INVALID]' {
            $json = Edit-TestEnvelope -Json $script:EnvelopeJson -Change {
                param($f)
                $bytes = [System.Convert]::FromBase64String($f.sig)
                $f.sig = [System.Convert]::ToBase64String($bytes, 0, 128)
            }
            { Read-StdSignedEnvelope -Json $json -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-INVALID`]*'
        }

        It 'kid innego zaufanego certyfikatu przy podpisie tym certyfikatem -> [STD-SIG-INVALID]' {
            $otherKid = $script:OtherCert.Thumbprint.ToUpperInvariant()
            $json = Edit-TestEnvelope -Json $script:EnvelopeJson -Change ({ param($f) $f.kid = $otherKid }.GetNewClosure())
            { Read-StdSignedEnvelope -Json $json -TrustedCertificates @($script:Trusted, $script:OtherTrusted) } | Should -Throw -ExpectedMessage '*`[STD-SIG-INVALID`]*'
        }

        It 'nieznany certyfikat -> [STD-SIG-UNTRUSTED]' {
            { Read-StdSignedEnvelope -Json $script:EnvelopeJson -TrustedCertificates $script:OtherTrusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-UNTRUSTED`]*'
        }

        It 'pusta lista zaufanych -> [STD-SIG-NOTRUST]' {
            { Read-StdSignedEnvelope -Json $script:EnvelopeJson -TrustedCertificates @() } | Should -Throw -ExpectedMessage '*`[STD-SIG-NOTRUST`]*'
            { Read-StdSignedEnvelope -Json $script:EnvelopeJson -TrustedCertificates $null } | Should -Throw -ExpectedMessage '*`[STD-SIG-NOTRUST`]*'
        }

        It 'nieobsługiwany alg <Alg> -> [STD-SIG-ALG]' -TestCases @(
            @{ Alg = 'HS256' }
            @{ Alg = 'none' }
            @{ Alg = 'rs256' }
            @{ Alg = 'RS512' }
            @{ Alg = 'PS256' }
        ) {
            param($Alg)
            $json = Edit-TestEnvelope -Json $script:EnvelopeJson -Change ({ param($f) $f.alg = $Alg }.GetNewClosure())
            { Read-StdSignedEnvelope -Json $json -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-ALG`]*'
        }

        It 'zła koperta (<Case>) -> [STD-SIG-FORMAT]' -TestCases @(
            @{ Case = 'kind inny niż std.signed'; Field = 'kind'; Value = 'std.other'; Remove = $false }
            @{ Case = 'kind liczbą'; Field = 'kind'; Value = 5; Remove = $false }
            @{ Case = 'brak kind'; Field = 'kind'; Value = $null; Remove = $true }
            @{ Case = 'brak alg'; Field = 'alg'; Value = $null; Remove = $true }
            @{ Case = 'brak kid'; Field = 'kid'; Value = $null; Remove = $true }
            @{ Case = 'brak payload'; Field = 'payload'; Value = $null; Remove = $true }
            @{ Case = 'brak sig'; Field = 'sig'; Value = $null; Remove = $true }
            @{ Case = 'pusty sig'; Field = 'sig'; Value = ''; Remove = $false }
            @{ Case = 'payload nie jest base64'; Field = 'payload'; Value = '!!! to nie base64 !!!'; Remove = $false }
            @{ Case = 'sig nie jest base64'; Field = 'sig'; Value = 'ąę'; Remove = $false }
            @{ Case = 'kid nie jest odciskiem SHA-1'; Field = 'kid'; Value = 'ABC123'; Remove = $false }
        ) {
            param($Case, $Field, $Value, $Remove)
            $json = Edit-TestEnvelope -Json $script:EnvelopeJson -Change ({
                    param($f)
                    if ($Remove) { $f.Remove($Field) } else { $f[$Field] = $Value }
                }.GetNewClosure())
            { Read-StdSignedEnvelope -Json $json -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-FORMAT`]*'
        }

        It 'tekst, który nie jest obiektem JSON (<Case>) -> [STD-SIG-FORMAT]' -TestCases @(
            @{ Case = 'pusty'; Text = '' }
            @{ Case = 'spacje'; Text = '   ' }
            @{ Case = 'nie JSON'; Text = 'to nie jest JSON' }
            @{ Case = 'tablica'; Text = '[]' }
            @{ Case = 'null'; Text = 'null' }
            @{ Case = 'urwany'; Text = '{"kind":"std.signed",' }
        ) {
            param($Case, $Text)
            { Read-StdSignedEnvelope -Json $Text -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-FORMAT`]*'
        }

        It 'bajty, które nie są tekstem UTF-8 -> [STD-SIG-FORMAT]' {
            { Read-StdSignedEnvelope -Bytes ([byte[]](0x7B, 0xFF, 0xFE, 0x7D)) -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-FORMAT`]*'
            { Read-StdSignedEnvelope -Bytes (New-Object byte[] 0) -TrustedCertificates $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-FORMAT`]*'
        }

        It 'certyfikat bez klucza prywatnego -> [STD-SIG-KEY]' {
            { New-StdSignedEnvelope -Payload $script:Payload -Certificate $script:Trusted } | Should -Throw -ExpectedMessage '*`[STD-SIG-KEY`]*'
        }

        It 'pusty dokument -> [STD-SIG-FORMAT]' {
            { New-StdSignedEnvelope -Payload (New-Object byte[] 0) -Certificate $script:Cert } | Should -Throw -ExpectedMessage '*`[STD-SIG-FORMAT`]*'
        }

        It 'bez -TrustedCertificates używa Get-StdTrustedCertificates' {
            if ($null -eq (Get-Command -Name 'Get-StdTrustedCertificates' -Module 'STD' -ErrorAction SilentlyContinue)) {
                Set-ItResult -Skipped -Because 'Std.Config.ps1 (Get-StdTrustedCertificates) nie jest wczytany'
                return
            }
            $certBase64 = [System.Convert]::ToBase64String($script:Trusted.RawData)
            $mockBody = [scriptblock]::Create("[System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String('$certBase64'))")
            Mock -ModuleName 'STD' -CommandName 'Get-StdTrustedCertificates' -MockWith $mockBody
            (Read-StdSignedEnvelope -Json $script:EnvelopeJson).Kid | Should -BeExactly $script:Cert.Thumbprint.ToUpperInvariant()
            Should -Invoke -ModuleName 'STD' -CommandName 'Get-StdTrustedCertificates' -Times 1 -Exactly
        }

        It 'z podanym -TrustedCertificates nie sięga po Get-StdTrustedCertificates' {
            if ($null -eq (Get-Command -Name 'Get-StdTrustedCertificates' -Module 'STD' -ErrorAction SilentlyContinue)) {
                Set-ItResult -Skipped -Because 'Std.Config.ps1 (Get-StdTrustedCertificates) nie jest wczytany'
                return
            }
            $certBase64 = [System.Convert]::ToBase64String($script:Trusted.RawData)
            $mockBody = [scriptblock]::Create("[System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String('$certBase64'))")
            Mock -ModuleName 'STD' -CommandName 'Get-StdTrustedCertificates' -MockWith $mockBody
            # Pusta lista podana jawnie nie jest uzupełniana listą z pliku.
            { Read-StdSignedEnvelope -Json $script:EnvelopeJson -TrustedCertificates @() } | Should -Throw -ExpectedMessage '*`[STD-SIG-NOTRUST`]*'
            Should -Invoke -ModuleName 'STD' -CommandName 'Get-StdTrustedCertificates' -Times 0 -Exactly
        }

        It 'bez -TrustedCertificates i z pustą listą w pliku -> [STD-SIG-NOTRUST]' {
            if ($null -eq (Get-Command -Name 'Get-StdTrustedCertificates' -Module 'STD' -ErrorAction SilentlyContinue)) {
                Set-ItResult -Skipped -Because 'Std.Config.ps1 (Get-StdTrustedCertificates) nie jest wczytany'
                return
            }
            Mock -ModuleName 'STD' -CommandName 'Get-StdTrustedCertificates' -MockWith { }
            { Read-StdSignedEnvelope -Json $script:EnvelopeJson } | Should -Throw -ExpectedMessage '*`[STD-SIG-NOTRUST`]*'
        }
    }

    Context 'Sekrety (tryb sealed)' {
        BeforeAll {
            $script:Password = ConvertTo-TestSecure 'Hasło wdrożeniowe ąę!'
            $script:WrongPassword = ConvertTo-TestSecure 'Hasło wdrożeniowe ąę?'
            $script:Iterations = 1000

            # Wektor niezależny (Python: hashlib.pbkdf2_hmac + cryptography AES-256-CBC), iter = 1000.
            $script:ExternalPassword = ConvertTo-TestSecure 'Hasło-wdrożeniowe-ąę'
            $script:ExternalValue = 'Zażółć gęślą jaźń ' + $script:Emoji
            $script:ExternalJson = '{"mode": "sealed", "kdf": {"alg": "pbkdf2-sha256", "iter": 1000, "salt": "AAECAwQFBgcICQoLDA0ODw=="}, "check": "nRTbQw/4Rkulfv3yZ6BBYLxOq6da1iO5XnH15oCWEXc=", "items": {"tv-token": {"iv": "EBESExQVFhcYGRobHB0eHw==", "ct": "6d/IFjl5pqQdRAEGFMZ+zOO3aynRjPfDqdR2iVsb6js=", "mac": "Ffnu1qFmhhFC5IHGyMZBEdW4GeIoP9KgdOGDuS8nOfo="}, "short": {"iv": "ICEiIyQlJicoKSorLC0uLw==", "ct": "JY2VjXzhFbxujxheKANWOQ==", "mac": "fgoPSV5VG7JA9e8SLtC3+JgjpLp0MlLLjbA3/a63yqU="}}}'

            function New-TestStore {
                $store = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
                Set-StdSealedItem -Store $store -Name 'tv-token' -Value (ConvertTo-TestSecure 'tajny-token-123') -Password $script:Password
                return $store
            }

            function Edit-TestItemBytes {
                # Zmienia jeden bajt pola (iv / ct / mac) pozycji magazynu.
                param($Store, [string]$Name, [string]$Field)
                $bytes = [System.Convert]::FromBase64String($Store.items[$Name][$Field])
                $bytes[$bytes.Length - 1] = $bytes[$bytes.Length - 1] -bxor 0x01
                $Store.items[$Name][$Field] = [System.Convert]::ToBase64String($bytes)
            }
        }

        BeforeEach {
            Clear-StdLogMasks
        }

        It 'nowy magazyn ma budowę z SPEC 2.7' {
            $store = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            $store.mode | Should -BeExactly 'sealed'
            $store.kdf.alg | Should -BeExactly 'pbkdf2-sha256'
            $store.kdf.iter | Should -Be 1000
            ([System.Convert]::FromBase64String($store.kdf.salt)).Length | Should -Be 16
            ([System.Convert]::FromBase64String($store.check)).Length | Should -Be 32
            $store.items.Count | Should -Be 0
        }

        It 'domyślna liczba iteracji to 600000' {
            $command = Get-Command -Name 'New-StdSealedStore' -Module 'STD'
            $parameterAst = $command.ScriptBlock.Ast.Find({
                    param($node)
                    $node -is [System.Management.Automation.Language.ParameterAst] -and $node.Name.VariablePath.UserPath -eq 'Iterations'
                }, $true)
            $parameterAst.DefaultValue.Extent.Text | Should -Be '600000'
        }

        It 'każdy magazyn ma własną losową sól' {
            $first = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            $second = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            $first.kdf.salt | Should -Not -Be $second.kdf.salt
            $first.check | Should -Not -Be $second.check
        }

        It 'puste hasło -> [STD-SEC-PASSWORD]' {
            { New-StdSealedStore -Password (New-Object System.Security.SecureString) -Iterations $script:Iterations } | Should -Throw -ExpectedMessage '*`[STD-SEC-PASSWORD`]*'
        }

        It 'Test-StdSealedPassword: dobre hasło $true, złe $false' {
            $store = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            Test-StdSealedPassword -Store $store -Password $script:Password | Should -BeTrue
            Test-StdSealedPassword -Store $store -Password $script:WrongPassword | Should -BeFalse
            Test-StdSealedPassword -Store $store -Password (New-Object System.Security.SecureString) | Should -BeFalse
        }

        It 'zapisuje i odczytuje sekret: <Case>' -TestCases @(
            @{ Case = 'zwykły tekst'; Value = 'tajny-token-123' }
            @{ Case = 'polskie znaki'; Value = 'Zażółć gęślą jaźń' }
            @{ Case = 'cyrylica, CJK i emoji'; Value = 'ключ 鍵 <emoji>' }
            @{ Case = 'cudzysłowy i klamry'; Value = 'a "b" ''c'' {{secret:x}} \\srv\d$' }
            @{ Case = 'dokładnie 16 bajtów'; Value = '0123456789abcdef' }
        ) {
            param($Case, $Value)
            $Value = $Value.Replace('<emoji>', $script:Emoji)
            $store = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            Set-StdSealedItem -Store $store -Name 'wifi-firma' -Value (ConvertTo-TestSecure $Value) -Password $script:Password
            $result = Unprotect-StdSealedItem -Store $store -Name 'wifi-firma' -Password $script:Password
            $result | Should -BeOfType [string]
            $result | Should -BeExactly $Value
        }

        It 'odczytana wartość od razu trafia do maski logu' {
            $store = New-TestStore
            $value = Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password
            Protect-StdLogText -Text "APITOKEN=$value koniec" | Should -BeExactly 'APITOKEN=*** koniec'
        }

        It 'wartość sekretu nie trafia do logu' {
            $records = New-Object System.Collections.ArrayList
            Register-StdLogSink -Name 'crypto-test' -ScriptBlock ({ [void]$records.Add($args[0]) }.GetNewClosure())
            $env:STD_DEBUG = '1'
            try {
                $store = New-TestStore
                $value = Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password
                Write-StdLog -Message "Instalacja z tokenem $value"
            } finally {
                Remove-Item Env:\STD_DEBUG -ErrorAction SilentlyContinue
                Unregister-StdLogSink -Name 'crypto-test'
            }
            $records.Count | Should -BeGreaterThan 0
            foreach ($record in $records) {
                $record.Message | Should -Not -BeLike '*tajny-token-123*'
                $record.Message | Should -Not -BeLike '*Hasło wdrożeniowe*'
            }
        }

        It 'Set-StdSealedItem podmienia istniejącą pozycję' {
            $store = New-TestStore
            Set-StdSealedItem -Store $store -Name 'tv-token' -Value (ConvertTo-TestSecure 'nowy-token-456') -Password $script:Password
            $store.items.Count | Should -Be 1
            Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password | Should -BeExactly 'nowy-token-456'
        }

        It 'każdy zapis ma nowe losowe IV' {
            $store = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            Set-StdSealedItem -Store $store -Name 'a' -Value (ConvertTo-TestSecure 'ta sama wartość') -Password $script:Password
            Set-StdSealedItem -Store $store -Name 'b' -Value (ConvertTo-TestSecure 'ta sama wartość') -Password $script:Password
            $store.items['a'].iv | Should -Not -Be $store.items['b'].iv
            $store.items['a'].ct | Should -Not -Be $store.items['b'].ct
            ([System.Convert]::FromBase64String($store.items['a'].iv)).Length | Should -Be 16
            ([System.Convert]::FromBase64String($store.items['a'].mac)).Length | Should -Be 32
        }

        It 'magazyn w JSON nie zawiera wartości ani hasła' {
            $store = New-TestStore
            $json = ConvertTo-StdJson -InputObject $store
            $json | Should -Not -BeLike '*tajny-token-123*'
            $json | Should -Not -BeLike '*Hasło wdrożeniowe*'
            $json | Should -BeLike '*"tv-token"*'
        }

        It 'działa po zapisie do JSON i odczycie (ConvertFrom-StdJson)' {
            $store = New-TestStore
            $copy = ConvertFrom-StdJson -Json (ConvertTo-StdJson -InputObject $store)
            Test-StdSealedPassword -Store $copy -Password $script:Password | Should -BeTrue
            Unprotect-StdSealedItem -Store $copy -Name 'tv-token' -Password $script:Password | Should -BeExactly 'tajny-token-123'
            Set-StdSealedItem -Store $copy -Name 'wifi-firma' -Value (ConvertTo-TestSecure 'klucz-wifi') -Password $script:Password
            Unprotect-StdSealedItem -Store $copy -Name 'wifi-firma' -Password $script:Password | Should -BeExactly 'klucz-wifi'
        }

        It 'odczytuje magazyn zaszyfrowany niezależnie (poza STD)' {
            $store = ConvertFrom-StdJson -Json $script:ExternalJson
            Test-StdSealedPassword -Store $store -Password $script:ExternalPassword | Should -BeTrue
            Test-StdSealedPassword -Store $store -Password $script:WrongPassword | Should -BeFalse
            Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:ExternalPassword | Should -BeExactly $script:ExternalValue
            Unprotect-StdSealedItem -Store $store -Name 'short' -Password $script:ExternalPassword | Should -BeExactly 'ab'
        }

        It 'złe hasło przy odczycie -> [STD-SEC-PASSWORD]' {
            $store = New-TestStore
            { Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:WrongPassword } | Should -Throw -ExpectedMessage '*`[STD-SEC-PASSWORD`]*'
        }

        It 'złe hasło przy zapisie -> [STD-SEC-PASSWORD], magazyn bez zmian' {
            $store = New-TestStore
            $before = ConvertTo-StdJson -InputObject $store
            { Set-StdSealedItem -Store $store -Name 'tv-token' -Value (ConvertTo-TestSecure 'podmiana') -Password $script:WrongPassword } | Should -Throw -ExpectedMessage '*`[STD-SEC-PASSWORD`]*'
            { Set-StdSealedItem -Store $store -Name 'nowy' -Value (ConvertTo-TestSecure 'podmiana') -Password $script:WrongPassword } | Should -Throw -ExpectedMessage '*`[STD-SEC-PASSWORD`]*'
            ConvertTo-StdJson -InputObject $store | Should -BeExactly $before
        }

        It 'zmienione pole <Field> -> [STD-SEC-TAMPER]' -TestCases @(
            @{ Field = 'ct' }
            @{ Field = 'iv' }
            @{ Field = 'mac' }
        ) {
            param($Field)
            $store = New-TestStore
            Edit-TestItemBytes -Store $store -Name 'tv-token' -Field $Field
            { Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-TAMPER`]*'
        }

        It 'uszkodzona pozycja (<Case>) -> [STD-SEC-TAMPER]' -TestCases @(
            @{ Case = 'ct ucięty'; Field = 'ct'; Value = 'AAAA' }
            @{ Case = 'ct nie base64'; Field = 'ct'; Value = '###' }
            @{ Case = 'iv za krótki'; Field = 'iv'; Value = 'AAAAAAAAAAA=' }
            @{ Case = 'brak mac'; Field = 'mac'; Value = $null }
        ) {
            param($Case, $Field, $Value)
            $store = New-TestStore
            if ($null -eq $Value) { $store.items['tv-token'].Remove($Field) } else { $store.items['tv-token'][$Field] = $Value }
            { Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-TAMPER`]*'
        }

        It 'pozycja przeniesiona z innego magazynu -> [STD-SEC-TAMPER]' {
            $store = New-TestStore
            $other = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            Set-StdSealedItem -Store $other -Name 'tv-token' -Value (ConvertTo-TestSecure 'obcy') -Password $script:Password
            $store.items['tv-token'] = $other.items['tv-token']
            { Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-TAMPER`]*'
        }

        It 'brak pozycji -> [STD-SEC-MISSING]' {
            $store = New-TestStore
            { Unprotect-StdSealedItem -Store $store -Name 'nie-ma-takiego' -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-MISSING`]*'
            $empty = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            $empty.Remove('items')
            { Unprotect-StdSealedItem -Store $empty -Name 'tv-token' -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-MISSING`]*'
        }

        It 'Set-StdSealedItem dodaje items, gdy magazyn ich nie ma' {
            $store = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            $store.Remove('items')
            Set-StdSealedItem -Store $store -Name 'tv-token' -Value (ConvertTo-TestSecure 'abc-123') -Password $script:Password
            Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password | Should -BeExactly 'abc-123'
        }

        It 'zły układ magazynu (<Case>) -> [STD-SEC-FORMAT]' -TestCases @(
            @{ Case = 'mode inny niż sealed'; FieldPath = 'mode'; Value = 'plain' }
            @{ Case = 'inny algorytm kdf'; FieldPath = 'kdf.alg'; Value = 'pbkdf2-sha1' }
            @{ Case = 'iter = 0'; FieldPath = 'kdf.iter'; Value = 0 }
            @{ Case = 'iter tekstem'; FieldPath = 'kdf.iter'; Value = '1000' }
            @{ Case = 'salt za krótki'; FieldPath = 'kdf.salt'; Value = 'AAAA' }
            @{ Case = 'check nie base64'; FieldPath = 'check'; Value = '%%%' }
            @{ Case = 'brak check'; FieldPath = 'check'; Value = $null }
        ) {
            param($Case, $FieldPath, $Value)
            $store = New-TestStore
            $target = $store
            $parts = $FieldPath.Split('.')
            for ($i = 0; $i -lt $parts.Length - 1; $i++) { $target = $target[$parts[$i]] }
            if ($null -eq $Value) { $target.Remove($parts[-1]) } else { $target[$parts[-1]] = $Value }
            { Unprotect-StdSealedItem -Store $store -Name 'tv-token' -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-FORMAT`]*'
            { Test-StdSealedPassword -Store $store -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-FORMAT`]*'
        }

        It 'zła nazwa sekretu (<SecretName>) -> [STD-SEC-FORMAT]' -TestCases @(
            @{ SecretName = 'z spacją' }
            @{ SecretName = '-od-myślnika' }
            @{ SecretName = 'klamra}' }
            @{ SecretName = 'ąę' }
        ) {
            param($SecretName)
            $store = New-StdSealedStore -Password $script:Password -Iterations $script:Iterations
            { Set-StdSealedItem -Store $store -Name $SecretName -Value (ConvertTo-TestSecure 'x-y-z') -Password $script:Password } | Should -Throw -ExpectedMessage '*`[STD-SEC-FORMAT`]*'
        }
    }

    Context 'Wstawki {{secret:nazwa}}' {
        BeforeAll {
            $script:Password = ConvertTo-TestSecure 'Hasło wdrożeniowe ąę!'
        }

        BeforeEach {
            Clear-StdLogMasks
        }

        It 'Get-StdSecretReference zwraca unikalne nazwy w kolejności wystąpienia' {
            $text = '/qn CUSTOMCONFIGID={{secret:teamviewer-customconfigid}} APITOKEN={{secret:teamviewer-apitoken}} X={{secret:teamviewer-customconfigid}}'
            (@(Get-StdSecretReference -Text $text) -join ',') | Should -BeExactly 'teamviewer-customconfigid,teamviewer-apitoken'
        }

        It 'Get-StdSecretReference bez wstawek zwraca pustą listę' {
            @(Get-StdSecretReference -Text '/S /norestart').Count | Should -Be 0
            @(Get-StdSecretReference -Text '').Count | Should -Be 0
            @(Get-StdSecretReference -Text '{{secret:}} {{ secret:a }} {secret:b} {{secret:z spacją}}').Count | Should -Be 0
        }

        It 'Get-StdSecretReference: jedna nazwa, kropka i podkreślenie' {
            @(Get-StdSecretReference -Text 'KEY={{secret:wifi_firma.v2}}') | Should -Be 'wifi_firma.v2'
        }

        It 'Expand-StdSecretTemplate podmienia wstawki wynikiem resolvera (każda nazwa raz)' {
            $calls = New-Object System.Collections.ArrayList
            $resolver = { param($name) [void]$calls.Add($name); "wartosc-$name" }.GetNewClosure()
            $text = 'A={{secret:alfa}} B={{secret:beta}} A2={{secret:alfa}}'
            Expand-StdSecretTemplate -Text $text -Resolver $resolver | Should -BeExactly 'A=wartosc-alfa B=wartosc-beta A2=wartosc-alfa'
            ($calls -join ',') | Should -BeExactly 'alfa,beta'
        }

        It 'tekst bez wstawek wraca bez zmian, resolver nie jest wołany' {
            $calls = New-Object System.Collections.ArrayList
            $resolver = { param($name) [void]$calls.Add($name); 'x' }.GetNewClosure()
            Expand-StdSecretTemplate -Text '/S /D=C:\Program Files\7-Zip' -Resolver $resolver | Should -BeExactly '/S /D=C:\Program Files\7-Zip'
            Expand-StdSecretTemplate -Text '' -Resolver $resolver | Should -BeNullOrEmpty
            $calls.Count | Should -Be 0
        }

        It 'podstawione wartości są maskowane przez Protect-StdLogText' {
            $resolver = { param($name) if ($name -eq 'teamviewer-apitoken') { '12345-AbCdEf' } else { 'cfg-987' } }
            $expanded = Expand-StdSecretTemplate -Text '/qn CUSTOMCONFIGID={{secret:teamviewer-customconfigid}} APITOKEN={{secret:teamviewer-apitoken}}' -Resolver $resolver
            $expanded | Should -BeExactly '/qn CUSTOMCONFIGID=cfg-987 APITOKEN=12345-AbCdEf'
            Protect-StdLogText -Text "msiexec $expanded" | Should -BeExactly 'msiexec /qn CUSTOMCONFIGID=*** APITOKEN=***'
        }

        It 'działa z Unprotect-StdSealedItem jako resolverem' {
            $store = New-StdSealedStore -Password $script:Password -Iterations 1000
            Set-StdSealedItem -Store $store -Name 'wifi-firma' -Value (ConvertTo-TestSecure 'Klucz Wi-Fi ąę') -Password $script:Password
            $password = $script:Password
            $resolver = { param($name) Unprotect-StdSealedItem -Store $store -Name $name -Password $password }.GetNewClosure()
            $expanded = Expand-StdSecretTemplate -Text 'key={{secret:wifi-firma}}' -Resolver $resolver
            $expanded | Should -BeExactly 'key=Klucz Wi-Fi ąę'
            Protect-StdLogText -Text $expanded | Should -BeExactly 'key=***'
            { Expand-StdSecretTemplate -Text 'x={{secret:brak}}' -Resolver $resolver } | Should -Throw -ExpectedMessage '*`[STD-SEC-MISSING`]*'
        }

        It 'resolver może zwrócić SecureString' {
            $resolver = { param($name) ConvertTo-SecureString -String "ss-$name-wartosc" -AsPlainText -Force }
            $expanded = Expand-StdSecretTemplate -Text 'v={{secret:abc}}' -Resolver $resolver
            $expanded | Should -BeExactly 'v=ss-abc-wartosc'
            Protect-StdLogText -Text $expanded | Should -BeExactly 'v=***'
        }

        It 'resolver bez wartości -> [STD-SEC-MISSING]' {
            { Expand-StdSecretTemplate -Text 'v={{secret:abc}}' -Resolver { param($name) $null } } | Should -Throw -ExpectedMessage '*`[STD-SEC-MISSING`]*'
        }

        It 'wartość z wstawką nie jest rozwijana drugi raz' {
            $resolver = { param($name) if ($name -eq 'a') { '{{secret:b}}' } else { 'NIE' } }
            Expand-StdSecretTemplate -Text '[{{secret:a}}]' -Resolver $resolver | Should -BeExactly '[{{secret:b}}]'
        }
    }
}
