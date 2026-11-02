# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.

<#
.SYNOPSIS
Test dymny STD 4: cała droga konfiguracji od starego config.json do wartości sekretu w poleceniu instalacji.

.DESCRIPTION
Zwykły skrypt PowerShella (bez Pester), działa w Windows PowerShell 5.1 i w PowerShell 7 (także na Linuksie).
Nie wymaga administratora, sieci ani magazynu certyfikatów. Wszystko dzieje się w katalogu tymczasowym,
który na końcu jest usuwany (chyba że podano -KeepTemp).

Kroki:
  1. Wczytanie modułu STD (STD.psd1).
  2. Certyfikat testowy w pamięci (CertificateRequest, RSA 2048).
  3. Konwersja config.json z STD 3.1 (ConvertFrom-StdLegacyConfig) - bez haseł i loginów z 3.1.
  4. Uzupełnienie dokumentu: adresy https://, zaszyfrowane sekrety (New-StdSealedStore, Set-StdSealedItem),
     walidacja Test-StdRelease -Unpublished; dokument z loginem lub hasłem do domeny albo hasłem lokalnego
     administratora jest odrzucany.
  5. Publikacja do katalogu serwera (Tools\Publish-StdRelease.ps1), dwa razy: rewizje 1 i 2.
  6. Plik startowy bootstrap.json (Get-StdBootstrap): serwer nieosiągalny, kopia lustrzana = katalog serwera.
  7. Lista klientów z kopii lustrzanej (Get-StdClientIndex).
  8. Konfiguracja klienta z kopii lustrzanej (Get-StdRelease), kopia w cache, widziana rewizja.
  9. Bez serwera i kopii lustrzanej: kopia z cache odrzucona (allowOffline = false).
 10. Konfiguracja wynikowa lokalizacji i profilu (Resolve-StdEffectiveConfig).
 11. Odszyfrowanie sekretów (Unprotect-StdSealedItem) i podstawienie w argumentach TeamViewer
     (Expand-StdSecretTemplate).
 12. Log do pliku (Register-StdFileLogSink): wartości sekretów zamaskowane (***).

Wypisuje PASS / FAIL dla każdego kroku (po pierwszym błędzie kolejne kroki: SKIP) i kończy się kodem 1,
gdy którykolwiek krok się nie udał.

.PARAMETER LegacyConfigPath
Plik config.json z STD 3.1. Domyślnie config.example.json z katalogu głównego repozytorium.

.PARAMETER Iterations
Liczba iteracji PBKDF2 magazynu sekretów w teście (domyślnie 50000, żeby test był szybki;
w prawdziwej konfiguracji domyślnie 600000).

.PARAMETER KeepTemp
Nie usuwa katalogu tymczasowego (do obejrzenia plików po teście).

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\v4\Tests\Invoke-StdSmoke.ps1

.EXAMPLE
pwsh -NoProfile -File ./v4/Tests/Invoke-StdSmoke.ps1 -KeepTemp
#>
[CmdletBinding()]
param(
    [string]$LegacyConfigPath,
    [ValidateRange(1000, 100000000)]
    [int]$Iterations = 50000,
    [switch]$KeepTemp
)

$ErrorActionPreference = 'Stop'

$script:V4Root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($LegacyConfigPath)) {
    $LegacyConfigPath = Join-Path (Split-Path -Parent $script:V4Root) 'config.example.json'
}

# Wspólny stan kroków (każdy krok działa we własnym zakresie, więc wyniki trzymamy w słowniku).
$script:Smoke = @{}
# Klucze certyfikatów w pamięci trzymamy do końca testu, żeby certyfikaty nie straciły klucza prywatnego.
$script:SmokeKeys = New-Object System.Collections.ArrayList
$script:Results = New-Object System.Collections.ArrayList
$script:Failed = $false
$script:StepNumber = 0
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)

# Wartości testowe. Hasło wdrożeniowe i sekrety są tylko w pamięci tego procesu.
$script:SmokeClientId = 'lizard'
$script:SmokeClientName = 'Lizard-Tech'
$script:SmokeSiteId = 'default'
$script:SmokeProfile = 'Standard'
$script:SmokeServer = 'https://127.0.0.1:9/std/'
$smokeSuffix = [guid]::NewGuid().ToString('N').Substring(0, 12)
$script:SmokeSecretValues = [ordered]@{
    'teamviewer-apitoken'       = 'smoke-token-' + $smokeSuffix
    'teamviewer-customconfigid' = 'smoke-cfgid-' + $smokeSuffix
}
$script:SmokePassword = ConvertTo-SecureString -String ('Smoke-Wdrozenie-' + $smokeSuffix) -AsPlainText -Force

# ---------------------------------------------------------------------------------------------
# Pomocnicze
# ---------------------------------------------------------------------------------------------

function Assert-Smoke {
    param([bool]$Condition, [string]$Message)
    # Numer linii wywołania (nie tej funkcji), żeby od razu było widać, która kontrola się nie udała.
    if (-not $Condition) { throw "$Message (linia $($MyInvocation.ScriptLineNumber))" }
}

function Join-SmokePath {
    # Join-Path z wieloma częściami (Windows PowerShell 5.1 nie ma -AdditionalChildPath).
    param([string]$Base, [string[]]$Parts)
    $path = $Base
    foreach ($part in $Parts) { $path = Join-Path $path $part }
    return $path
}

function Get-SmokeErrorText {
    param($ErrorRecord)
    $ex = $ErrorRecord.Exception
    while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
    $text = ([string]$ex.Message -replace '\s+', ' ').Trim()
    if ($text -notmatch '\(linia \d+\)$' -and $null -ne $ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.ScriptLineNumber -gt 0) {
        $text += " (linia $($ErrorRecord.InvocationInfo.ScriptLineNumber))"
    }
    # Komunikat błędu też przechodzi przez maskę logu (gdyby zawierał wartość sekretu).
    if (Get-Command -Name 'Protect-StdLogText' -ErrorAction SilentlyContinue) {
        $text = Protect-StdLogText -Text $text
    }
    foreach ($value in $script:SmokeSecretValues.Values) { $text = $text.Replace($value, '***') }
    return $text
}

function Invoke-SmokeStep {
    param([Parameter(Mandatory = $true)][string]$Name, [Parameter(Mandatory = $true)][scriptblock]$Body)
    $script:StepNumber++
    $label = '{0,2}. {1}' -f $script:StepNumber, $Name
    if ($script:Failed) {
        Write-Host "[SKIP] $label"
        [void]$script:Results.Add([pscustomobject]@{ Step = $label; Result = 'SKIP'; Detail = '' })
        return
    }
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $detail = & $Body
        $watch.Stop()
        $text = "[PASS] $label ($($watch.ElapsedMilliseconds) ms)"
        if (-not [string]::IsNullOrEmpty([string]$detail)) { $text += " - $detail" }
        Write-Host $text -ForegroundColor Green
        [void]$script:Results.Add([pscustomobject]@{ Step = $label; Result = 'PASS'; Detail = [string]$detail })
    } catch {
        $watch.Stop()
        $script:Failed = $true
        $message = Get-SmokeErrorText $_
        Write-Host "[FAIL] $label - $message" -ForegroundColor Red
        [void]$script:Results.Add([pscustomobject]@{ Step = $label; Result = 'FAIL'; Detail = $message })
    }
}

function New-SmokeCertificate {
    # Certyfikat do podpisu tylko w pamięci (bez magazynu certyfikatów).
    param([string]$Subject)
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        # Windows PowerShell 5.1 (.NET Framework 4.8): klucz CNG tylko w pamięci.
        $key = New-Object System.Security.Cryptography.RSACng(2048)
    } else {
        $key = [System.Security.Cryptography.RSA]::Create(2048)
    }
    [void]$script:SmokeKeys.Add($key)
    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        $Subject, $key,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $now = [System.DateTimeOffset]::UtcNow
    return $request.CreateSelfSigned($now.AddDays(-1), $now.AddDays(30))
}

function Get-SmokeTextValues {
    # Wszystkie wartości tekstowe w drzewie (słowniki, listy) - do szukania wycieków.
    param($Node)
    $out = New-Object 'System.Collections.Generic.List[string]'
    $stack = New-Object System.Collections.Stack
    $stack.Push($Node)
    while ($stack.Count -gt 0) {
        $item = $stack.Pop()
        if ($null -eq $item) { continue }
        if ($item -is [string]) { $out.Add($item); continue }
        if ($item -is [System.Collections.IDictionary]) {
            foreach ($entry in @($item.GetEnumerator())) { $out.Add([string]$entry.Key); $stack.Push($entry.Value) }
            continue
        }
        if ($item -is [System.Collections.IEnumerable]) {
            foreach ($child in $item) { $stack.Push($child) }
        }
    }
    return , $out.ToArray()
}

# ---------------------------------------------------------------------------------------------
# Katalog tymczasowy i dane lokalne
# ---------------------------------------------------------------------------------------------

$script:TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('std-smoke-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
[void][System.IO.Directory]::CreateDirectory($script:TempRoot)
$script:PreviousDataRoot = $env:STD_DATA_ROOT
$env:STD_DATA_ROOT = Join-Path $script:TempRoot 'data'
$script:ServerStd = Join-SmokePath $script:TempRoot @('server', 'data', 'std')
$script:UsbRoot = Join-Path $script:TempRoot 'usb'
$script:WorkDir = Join-Path $script:TempRoot 'work'
$script:LogPath = Join-SmokePath $script:TempRoot @('logs', 'deploy.log')
foreach ($dir in @($script:ServerStd, $script:UsbRoot, $script:WorkDir)) { [void][System.IO.Directory]::CreateDirectory($dir) }

Write-Host "STD 4 - test dymny ($($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)), katalog: $($script:TempRoot)"

try {
    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Wczytanie modułu STD' {
        Get-Module -Name STD | Remove-Module -Force
        Import-Module (Join-Path $script:V4Root 'STD.psd1') -Force
        $version = Get-StdVersion
        Assert-Smoke ($version -is [version]) 'Get-StdVersion nie zwraca [version].'
        $manifest = Import-PowerShellDataFile -Path (Join-Path $script:V4Root 'STD.psd1')
        Assert-Smoke ($version -eq [version]$manifest.ModuleVersion) "Get-StdVersion = $version, a manifest ma $($manifest.ModuleVersion)."
        $root = Get-StdDataRoot
        Assert-Smoke ([System.IO.Path]::GetFullPath($root).TrimEnd('\', '/') -eq [System.IO.Path]::GetFullPath($env:STD_DATA_ROOT).TrimEnd('\', '/')) "Get-StdDataRoot = $root zamiast STD_DATA_ROOT."
        Clear-StdLogMasks
        "wersja $version"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Certyfikat testowy w pamięci' {
        $cert = New-SmokeCertificate -Subject 'CN=STD test dymny'
        Assert-Smoke $cert.HasPrivateKey 'Certyfikat nie ma klucza prywatnego.'
        $script:Smoke.Cert = $cert
        # Kopia bez klucza prywatnego - tak jak wpis w Std.TrustedSigners.psd1.
        $public = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([byte[]]$cert.RawData)
        Assert-Smoke (-not $public.HasPrivateKey) 'Kopia certyfikatu nie powinna mieć klucza prywatnego.'
        $script:Smoke.Trusted = [System.Security.Cryptography.X509Certificates.X509Certificate2[]]@($public)
        $script:Smoke.Kid = ([string]$cert.Thumbprint).ToUpperInvariant()
        "odcisk $($script:Smoke.Kid)"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Konwersja config.json z STD 3.1 (ConvertFrom-StdLegacyConfig)' {
        Assert-Smoke ([System.IO.File]::Exists($LegacyConfigPath)) "Brak pliku $LegacyConfigPath."
        $legacy = ConvertFrom-StdJson -Json ([System.IO.File]::ReadAllText($LegacyConfigPath, $script:Utf8))
        $converted = ConvertFrom-StdLegacyConfig -Path $LegacyConfigPath -ClientId $script:SmokeClientId -ClientName $script:SmokeClientName
        $doc = $converted.Document
        Assert-Smoke ($doc -is [System.Collections.IDictionary]) 'Document nie jest słownikiem.'
        Assert-Smoke ($doc['kind'] -ceq 'std.release' -and $doc['schemaVersion'] -eq 4) 'Document nie jest konfiguracją std.release w wersji 4.'
        Assert-Smoke (-not $doc.Contains('revision') -and -not $doc.Contains('publishedAt')) 'Dokument po konwersji nie powinien mieć revision ani publishedAt.'
        $secretsToSet = @($converted.SecretsToSet)
        foreach ($name in $script:SmokeSecretValues.Keys) {
            Assert-Smoke ($secretsToSet -contains $name) "SecretsToSet nie zawiera '$name'."
        }

        # Żadna wartość hasła, tokenu ani loginu z 3.1 nie może trafić do wyniku.
        $forbidden = New-Object 'System.Collections.Generic.List[string]'
        $tvArgs = ''
        if ($legacy['TeamViewer'] -is [System.Collections.IDictionary]) { $tvArgs = [string]$legacy['TeamViewer']['Arguments'] }
        foreach ($match in [regex]::Matches($tvArgs, '(?i)\b(APITOKEN|CUSTOMCONFIGID)=("[^"]*"|\S+)')) { $forbidden.Add($match.Groups[2].Value.Trim('"')) }
        foreach ($pair in @(@('WebAuth', 'Password'), @('DomainJoin', 'Username'))) {
            if ($legacy[$pair[0]] -is [System.Collections.IDictionary]) { $forbidden.Add([string]$legacy[$pair[0]][$pair[1]]) }
        }
        if ($legacy['AntyVirus'] -is [System.Collections.IDictionary] -and $legacy['AntyVirus']['Credentials'] -is [System.Collections.IDictionary]) {
            $forbidden.Add([string]$legacy['AntyVirus']['Credentials']['Password'])
        }
        $values = Get-SmokeTextValues $doc
        $checked = 0
        foreach ($secret in $forbidden) {
            if ([string]::IsNullOrEmpty($secret) -or $secret.Length -lt 4) { continue }
            $checked++
            foreach ($value in $values) {
                Assert-Smoke ($value.IndexOf($secret, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) 'W wyniku konwersji jest hasło, token albo login z konfiguracji 3.1.'
            }
        }
        foreach ($site in @($doc['sites'])) {
            if ($site['domainJoin'] -is [System.Collections.IDictionary]) {
                foreach ($key in @($site['domainJoin'].Keys)) { Assert-Smoke (@('domain', 'ou') -contains $key) "sites[].domainJoin ma pole '$key' (dozwolone tylko domain i ou)." }
            }
            if ($site['localAdmin'] -is [System.Collections.IDictionary]) {
                foreach ($key in @($site['localAdmin'].Keys)) { Assert-Smoke ($key -ceq 'username') "sites[].localAdmin ma pole '$key' (dozwolone tylko username)." }
            }
        }
        $script:Smoke.Document = $doc
        $script:Smoke.Converted = $converted
        "pozycji katalogu $($doc['catalog'].Count), ostrzeżeń $(@($converted.Warnings).Count), sekretów do ustawienia $($secretsToSet.Count), sprawdzonych wartości z 3.1: $checked"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Uzupełnienie dokumentu (https, sekrety) i Test-StdRelease -Unpublished' {
        $doc = $script:Smoke.Document
        $fixed = 0
        foreach ($entry in @($doc['sources'].GetEnumerator())) {
            $source = $entry.Value
            if ($source['type'] -eq 'https' -and ([string]$source['url']).StartsWith('http://', [System.StringComparison]::OrdinalIgnoreCase)) {
                $source['url'] = 'https://' + ([string]$source['url']).Substring(7)
                $fixed++
            }
        }
        foreach ($entry in @($doc['catalog'].GetEnumerator())) {
            $item = $entry.Value
            if ($item.Contains('url') -and ([string]$item['url']).StartsWith('http://', [System.StringComparison]::OrdinalIgnoreCase)) {
                $item['url'] = 'https://' + ([string]$item['url']).Substring(7)
                $fixed++
            }
        }

        $store = New-StdSealedStore -Password $script:SmokePassword -Iterations $Iterations
        foreach ($name in @($script:Smoke.Converted.SecretsToSet)) {
            $plain = 'smoke-' + $name
            if ($script:SmokeSecretValues.Contains($name)) { $plain = $script:SmokeSecretValues[$name] }
            $secure = ConvertTo-SecureString -String $plain -AsPlainText -Force
            Set-StdSealedItem -Store $store -Name $name -Value $secure -Password $script:SmokePassword
        }
        Assert-Smoke (Test-StdSealedPassword -Store $store -Password $script:SmokePassword) 'Test-StdSealedPassword nie przyjmuje hasła wdrożeniowego.'
        $wrong = ConvertTo-SecureString -String 'zle-haslo-wdrozeniowe' -AsPlainText -Force
        Assert-Smoke (-not (Test-StdSealedPassword -Store $store -Password $wrong)) 'Test-StdSealedPassword przyjmuje złe hasło.'
        $doc['secrets'] = $store

        $storeJson = ConvertTo-StdJson -InputObject $store
        foreach ($value in $script:SmokeSecretValues.Values) {
            Assert-Smoke ($storeJson.IndexOf($value, [System.StringComparison]::Ordinal) -lt 0) 'Magazyn sekretów zawiera jawną wartość sekretu.'
        }

        $errors = @(Test-StdRelease -Release $doc -Unpublished)
        Assert-Smoke ($errors.Count -eq 0) ('Test-StdRelease -Unpublished: ' + ($errors -join '; '))

        # Login i hasło do domeny oraz hasło lokalnego administratora nie mogą być w konfiguracji.
        foreach ($case in @(@('domainJoin', 'password'), @('domainJoin', 'username'), @('localAdmin', 'password'))) {
            $copy = ConvertFrom-StdJson -Json (ConvertTo-StdJson -InputObject $doc)
            $section = $copy['sites'][0][$case[0]]
            if (-not ($section -is [System.Collections.IDictionary])) { $section = [ordered]@{}; $copy['sites'][0][$case[0]] = $section }
            $section[$case[1]] = 'Smoke-Nie-Zapisuj-1'
            $rejected = @(Test-StdRelease -Release $copy -Unpublished | Where-Object { $_ -like "sites``[0``].$($case[0]).$($case[1]):*ręcznie*" })
            Assert-Smoke ($rejected.Count -ge 1) "Test-StdRelease nie odrzuca pola sites[0].$($case[0]).$($case[1])."
        }
        $script:Smoke.SourcePath = Join-Path $script:WorkDir 'lizard.release.v4.json'
        Write-StdJsonFile -Path $script:Smoke.SourcePath -InputObject $doc
        "poprawionych adresów http: $fixed, sekretów: $($store['items'].Count)"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Publikacja do katalogu serwera (Publish-StdRelease.ps1, rewizje 1 i 2)' {
        $publish = Join-SmokePath $script:V4Root @('Tools', 'Publish-StdRelease.ps1')
        $first = & $publish -SourcePath $script:Smoke.SourcePath -Certificate $script:Smoke.Cert -Destination $script:ServerStd `
            -ChangeNote 'Test dymny - pierwsza publikacja' -PublishedBy 'smoke' -WarningAction SilentlyContinue 6>$null
        Assert-Smoke ($null -ne $first) 'Publish-StdRelease.ps1 nie zwrócił wyniku.'
        Assert-Smoke ($first.ClientId -ceq $script:SmokeClientId -and $first.Revision -eq 1 -and $first.IndexRevision -eq 1) "Pierwsza publikacja: klient $($first.ClientId), rewizja $($first.Revision), lista $($first.IndexRevision) (oczekiwano 1 i 1)."
        $second = & $publish -SourcePath $script:Smoke.SourcePath -Certificate $script:Smoke.Cert -Destination $script:ServerStd `
            -ChangeNote 'Test dymny - druga publikacja' -PublishedBy 'smoke' -WarningAction SilentlyContinue 6>$null
        Assert-Smoke ($second.Revision -eq 2 -and $second.IndexRevision -eq 2) "Druga publikacja: rewizja $($second.Revision), lista $($second.IndexRevision) (oczekiwano 2 i 2)."

        $releaseFile = Join-SmokePath $script:ServerStd @('clients', $script:SmokeClientId, 'release.sdt')
        foreach ($relative in @(@('clients.sdt'), @('clients', $script:SmokeClientId, 'release.sdt'), @('clients', $script:SmokeClientId, 'history', '1.sdt'), @('clients', $script:SmokeClientId, 'history', '2.sdt'))) {
            $path = Join-SmokePath $script:ServerStd $relative
            Assert-Smoke ([System.IO.File]::Exists($path)) "Brak pliku $path."
        }
        Assert-Smoke ((Get-StdSha256 -Path $releaseFile) -ceq $second.Sha256) 'Sha256 z wyniku publikacji nie zgadza się z plikiem release.sdt.'
        Assert-Smoke (Test-StdFileHash -Path $releaseFile -Expected $second.Sha256.ToUpperInvariant()) 'Test-StdFileHash nie przyjmuje sumy (wielkie litery).'

        $envelope = Read-StdSignedEnvelope -Bytes ([System.IO.File]::ReadAllBytes($releaseFile)) -TrustedCertificates $script:Smoke.Trusted
        Assert-Smoke ($envelope.Kid -ceq $script:Smoke.Kid) 'kid koperty nie jest odciskiem certyfikatu testowego.'
        foreach ($value in $script:SmokeSecretValues.Values) {
            Assert-Smoke ($envelope.PayloadText.IndexOf($value, [System.StringComparison]::Ordinal) -lt 0) 'Opublikowana konfiguracja zawiera jawną wartość sekretu.'
        }
        $published = ConvertFrom-StdJson -Json $envelope.PayloadText
        Assert-Smoke ($published['publishedBy'] -ceq 'smoke' -and $published['changeNote'] -ceq 'Test dymny - druga publikacja') 'Pola publikacji (publishedBy, changeNote) nie zgadzają się.'
        Assert-Smoke ($published.Contains('expiresAt') -and $published.Contains('publishedAt')) 'Brak expiresAt albo publishedAt w opublikowanej konfiguracji.'

        $untrusted = $false
        try {
            $other = New-SmokeCertificate -Subject 'CN=STD obcy'
            $null = Read-StdSignedEnvelope -Bytes ([System.IO.File]::ReadAllBytes($releaseFile)) -TrustedCertificates @([System.Security.Cryptography.X509Certificates.X509Certificate2]::new([byte[]]$other.RawData))
        } catch {
            $untrusted = ([string]$_.Exception.Message).StartsWith('[STD-SIG-UNTRUSTED]')
        }
        Assert-Smoke $untrusted 'Koperta podpisana nieznanym certyfikatem nie została odrzucona z [STD-SIG-UNTRUSTED].'
        "rewizja $($second.Revision), lista klientów $($second.IndexRevision)"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Plik startowy bootstrap.json (Get-StdBootstrap)' {
        $bootstrapFile = [ordered]@{
            server  = $script:SmokeServer.TrimEnd('/')
            mirrors = @($script:ServerStd)
            auth    = [ordered]@{ mode = 'basic' }
            proxy   = $null
        }
        [System.IO.File]::WriteAllText((Join-Path $script:UsbRoot 'bootstrap.json'), (ConvertTo-StdJson -InputObject $bootstrapFile), $script:Utf8)
        $bootstrap = Get-StdBootstrap -ScriptRoot $script:UsbRoot
        Assert-Smoke ($null -ne $bootstrap) 'Get-StdBootstrap nie wczytał pliku bootstrap.json.'
        Assert-Smoke ($bootstrap['server'] -ceq $script:SmokeServer) "Adres serwera '$($bootstrap['server'])' (oczekiwano '$($script:SmokeServer)' z ukośnikiem na końcu)."
        Assert-Smoke (@($bootstrap['mirrors']).Count -eq 1) 'Brak kopii lustrzanej w pliku startowym.'
        $script:Smoke.Bootstrap = $bootstrap
        "serwer $($bootstrap['server']) (nieosiągalny), kopia lustrzana: katalog serwera"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Lista klientów z kopii lustrzanej (Get-StdClientIndex)' {
        $index = Get-StdClientIndex -Bootstrap $script:Smoke.Bootstrap -TrustedCertificates $script:Smoke.Trusted
        Assert-Smoke ($index.Source -ceq 'mirror') "Źródło listy klientów: $($index.Source) (oczekiwano mirror)."
        Assert-Smoke ($index.Revision -eq 2) "Rewizja listy klientów: $($index.Revision) (oczekiwano 2)."
        $attempts = @($index.Attempts)
        Assert-Smoke ($attempts.Count -ge 2 -and $attempts[0].Source -ceq 'server' -and -not [string]::IsNullOrEmpty([string]$attempts[0].Error)) 'Pierwsza próba (serwer) powinna się nie udać i być zapisana w Attempts.'
        Assert-Smoke ([string]$attempts[0].Error -like '`[STD-HTTP-*') "Błąd próby serwera bez kodu [STD-HTTP-...]: $($attempts[0].Error)"
        $clients = @($index.Index['clients'])
        Assert-Smoke ($clients.Count -eq 1 -and $clients[0]['id'] -ceq $script:SmokeClientId) 'Lista klientów nie zawiera klienta testowego.'
        Assert-Smoke ($clients[0]['release'] -ceq "clients/$($script:SmokeClientId)/release.sdt") "Ścieżka konfiguracji klienta: $($clients[0]['release'])."
        Assert-Smoke ((Get-StdSeenRevision -Key 'clients') -eq 2) 'Widziana rewizja listy klientów nie została zapisana.'
        "klient $($clients[0]['id']) ($($clients[0]['name'])), serwer: $($attempts[0].Error.Split(' ')[0])"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Konfiguracja klienta z kopii lustrzanej (Get-StdRelease)' {
        $release = Get-StdRelease -ClientId $script:SmokeClientId -Bootstrap $script:Smoke.Bootstrap -TrustedCertificates $script:Smoke.Trusted
        Assert-Smoke ($release.Source -ceq 'mirror') "Źródło konfiguracji: $($release.Source) (oczekiwano mirror)."
        Assert-Smoke ($release.Revision -eq 2) "Rewizja konfiguracji: $($release.Revision) (oczekiwano 2)."
        Assert-Smoke ($release.Kid -ceq $script:Smoke.Kid) 'Kid konfiguracji nie jest odciskiem certyfikatu testowego.'
        Assert-Smoke ($release.Release -is [System.Collections.IDictionary]) 'Release nie jest słownikiem.'
        Assert-Smoke ($release.Release['client']['id'] -ceq $script:SmokeClientId) 'Konfiguracja innego klienta.'
        $cacheFile = Join-SmokePath (Get-StdPath -Kind Cache) @($script:SmokeClientId, 'release.sdt')
        Assert-Smoke ([System.IO.File]::Exists($cacheFile)) "Brak kopii w cache: $cacheFile."
        Assert-Smoke ((Get-StdSeenRevision -Key $script:SmokeClientId) -eq 2) 'Widziana rewizja klienta nie została zapisana.'
        $script:Smoke.Release = $release
        "rewizja $($release.Revision), podpis $($release.Kid.Substring(0, 8))..., kopia w cache"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Bez serwera i kopii lustrzanej: cache odrzucony (allowOffline = false)' {
        $offlineBootstrap = @{ server = $script:SmokeServer; mirrors = @() }
        $code = $null
        $attempts = @()
        try {
            $null = Get-StdRelease -ClientId $script:SmokeClientId -Bootstrap $offlineBootstrap -TrustedCertificates $script:Smoke.Trusted
        } catch {
            $code = ([string]$_.Exception.Message).Split(' ')[0]
            $attempts = @($_.TargetObject)
        }
        Assert-Smoke ($code -ceq '[STD-CFG-UNAVAILABLE]') "Oczekiwano [STD-CFG-UNAVAILABLE], jest: $code."
        $cacheAttempt = @($attempts | Where-Object { $_.Source -eq 'cache' })
        Assert-Smoke ($cacheAttempt.Count -eq 1 -and ([string]$cacheAttempt[0].Error).StartsWith('[STD-CFG-OFFLINE]')) 'Kopia z cache powinna zostać odrzucona z [STD-CFG-OFFLINE].'
        'cache: [STD-CFG-OFFLINE]'
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep "Konfiguracja wynikowa (Resolve-StdEffectiveConfig, profil $($script:SmokeProfile))" {
        $effective = Resolve-StdEffectiveConfig -Release $script:Smoke.Release.Release -SiteId $script:SmokeSiteId -ProfileName $script:SmokeProfile
        Assert-Smoke ($effective.ClientId -ceq $script:SmokeClientId -and $effective.Revision -eq 2) 'ClientId albo Revision konfiguracji wynikowej się nie zgadza.'
        Assert-Smoke ($effective.ProfileName -ceq $script:SmokeProfile) "ProfileName = $($effective.ProfileName)."
        $apps = @($effective.Apps)
        foreach ($app in @('TeamViewer', 'Chrome')) { Assert-Smoke ($apps -contains $app) "Profil nie zawiera aplikacji $app." }
        foreach ($app in $apps) { Assert-Smoke ($effective.Catalog.Contains($app)) "Aplikacji $app nie ma w katalogu." }
        Assert-Smoke ($effective.Tasks -is [System.Collections.IDictionary] -and $effective.Tasks['JoinDomain'] -eq $true) 'Zadanie JoinDomain z defaults.tasks nie przeszło do konfiguracji wynikowej.'
        Assert-Smoke (-not $effective.Site.Contains('profiles')) 'Site w konfiguracji wynikowej nie powinien mieć profiles.'
        Assert-Smoke ($effective.Site['domainJoin']['domain'] -ceq 'firma.local') 'Brak nazwy domeny w lokalizacji.'
        $threw = $false
        try { $null = Resolve-StdEffectiveConfig -Release $script:Smoke.Release.Release -SiteId $script:SmokeSiteId -ProfileName 'nie-ma-takiego' } catch { $threw = ([string]$_.Exception.Message).StartsWith('[STD-CFG-PROFILE]') }
        Assert-Smoke $threw 'Nieznany profil nie dał błędu [STD-CFG-PROFILE].'
        $script:Smoke.Effective = $effective
        "aplikacji $($apps.Count), zadań $($effective.Tasks.Count)"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Sekrety: Unprotect-StdSealedItem i Expand-StdSecretTemplate (TeamViewer)' {
        $store = $script:Smoke.Release.Release['secrets']
        Assert-Smoke ($store -is [System.Collections.IDictionary]) 'Opublikowana konfiguracja nie ma magazynu sekretów.'
        $token = Unprotect-StdSealedItem -Store $store -Name 'teamviewer-apitoken' -Password $script:SmokePassword
        Assert-Smoke ($token -ceq $script:SmokeSecretValues['teamviewer-apitoken']) 'Odszyfrowany token TeamViewer nie zgadza się z zapisanym.'
        $code = $null
        try {
            $null = Unprotect-StdSealedItem -Store $store -Name 'teamviewer-apitoken' -Password (ConvertTo-SecureString -String 'zle-haslo-wdrozeniowe' -AsPlainText -Force)
        } catch { $code = ([string]$_.Exception.Message).Split(' ')[0] }
        Assert-Smoke ($code -ceq '[STD-SEC-PASSWORD]') "Złe hasło wdrożeniowe: oczekiwano [STD-SEC-PASSWORD], jest: $code."

        $template = [string]$script:Smoke.Effective.Catalog['TeamViewer']['args']
        $references = @(Get-StdSecretReference -Text $template)
        foreach ($name in $script:SmokeSecretValues.Keys) { Assert-Smoke ($references -contains $name) "Argumenty TeamViewer nie mają wstawki {{secret:$name}}." }
        # Resolver działa w zakresie tego skryptu: magazyn i hasło bierze ze stanu testu.
        $resolver = { param($name) Unprotect-StdSealedItem -Store $script:Smoke.Release.Release['secrets'] -Name $name -Password $script:SmokePassword }
        $expanded = Expand-StdSecretTemplate -Text $template -Resolver $resolver
        Assert-Smoke ($expanded.IndexOf('{{secret:', [System.StringComparison]::Ordinal) -lt 0) 'Po podstawieniu zostały wstawki {{secret:...}}.'
        foreach ($value in $script:SmokeSecretValues.Values) {
            Assert-Smoke ($expanded.IndexOf($value, [System.StringComparison]::Ordinal) -ge 0) 'Po podstawieniu brakuje wartości sekretu w argumentach.'
        }
        $script:Smoke.Expanded = $expanded
        "podstawiono $($references.Count) sekrety"
    }

    # -----------------------------------------------------------------------------------------
    Invoke-SmokeStep 'Log do pliku: wartości sekretów zamaskowane (Register-StdFileLogSink)' {
        Register-StdFileLogSink -Name 'smoke-file' -Path $script:LogPath
        try {
            Write-StdLog -Level Info -Source 'Invoke-StdSmoke' -Message ("Instalacja TeamViewer: msiexec /i TeamViewer_Host.msi " + $script:Smoke.Expanded)
            Write-StdLog -Level Warn -Source 'Invoke-StdSmoke' -Message ('Token wielkimi literami: ' + $script:SmokeSecretValues['teamviewer-apitoken'].ToUpperInvariant())
            Write-StdLog -Level Info -Source 'Invoke-StdSmoke' -Message 'Authorization: Basic c21va2U6c21va2U='
        } finally {
            Unregister-StdLogSink -Name 'smoke-file'
        }
        Assert-Smoke ([System.IO.File]::Exists($script:LogPath)) "Brak pliku logu $($script:LogPath)."
        $logBytes = [System.IO.File]::ReadAllBytes($script:LogPath)
        Assert-Smoke (-not ($logBytes.Length -ge 3 -and $logBytes[0] -eq 0xEF -and $logBytes[1] -eq 0xBB -and $logBytes[2] -eq 0xBF)) 'Plik logu ma BOM (oczekiwano UTF-8 bez BOM).'
        $log = $script:Utf8.GetString($logBytes)
        foreach ($value in $script:SmokeSecretValues.Values) {
            Assert-Smoke ($log.IndexOf($value, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) 'W pliku logu jest jawna wartość sekretu.'
        }
        Assert-Smoke ($log.IndexOf('c21va2U6c21va2U=', [System.StringComparison]::Ordinal) -lt 0) 'W pliku logu jest wartość nagłówka Authorization.'
        Assert-Smoke ($log -match 'APITOKEN=\*\*\*' -and $log -match 'CUSTOMCONFIGID=\*\*\*') 'W logu brakuje zamaskowanych wartości APITOKEN=*** i CUSTOMCONFIGID=***.'
        $lines = @($log -split "`n" | Where-Object { $_.Trim() -ne '' })
        Assert-Smoke ($lines.Count -eq 3) "Plik logu ma $($lines.Count) linii zamiast 3."
        Assert-Smoke ($lines[0] -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[INFO\] ') "Zły format linii logu: $($lines[0].Substring(0, [Math]::Min(40, $lines[0].Length)))"
        Assert-Smoke ($lines[1] -match '\[WARN\] ') 'Brak poziomu [WARN] w drugiej linii logu.'
        Assert-Smoke ((Protect-StdLogText -Text $script:Smoke.Expanded).IndexOf($script:SmokeSecretValues['teamviewer-customconfigid'], [System.StringComparison]::Ordinal) -lt 0) 'Protect-StdLogText nie maskuje wartości sekretu.'
        "linii w logu: $($lines.Count), sekrety jako ***"
    }
} finally {
    try { if (Get-Command -Name 'Unregister-StdLogSink' -ErrorAction SilentlyContinue) { Unregister-StdLogSink -Name 'smoke-file' } } catch { }
    try { if (Get-Command -Name 'Clear-StdLogMasks' -ErrorAction SilentlyContinue) { Clear-StdLogMasks } } catch { }
    foreach ($key in $script:SmokeKeys) { try { $key.Dispose() } catch { } }
    $env:STD_DATA_ROOT = $script:PreviousDataRoot
    if ($KeepTemp) {
        Write-Host "Katalog testu zostawiono: $($script:TempRoot)"
    } else {
        try { Remove-Item -LiteralPath $script:TempRoot -Recurse -Force -ErrorAction Stop } catch { Write-Host "Nie udało się usunąć katalogu $($script:TempRoot): $($_.Exception.Message)" }
    }
}

$passed = @($script:Results | Where-Object { $_.Result -eq 'PASS' }).Count
$failedCount = @($script:Results | Where-Object { $_.Result -eq 'FAIL' }).Count
$skipped = @($script:Results | Where-Object { $_.Result -eq 'SKIP' }).Count
if ($failedCount -gt 0 -or $passed -eq 0) {
    Write-Host "WYNIK: FAIL (PASS $passed, FAIL $failedCount, SKIP $skipped)" -ForegroundColor Red
    exit 1
}
Write-Host "WYNIK: PASS (kroków: $passed)" -ForegroundColor Green
exit 0
