# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.

<#
.SYNOPSIS
Publikacja podpisanej konfiguracji klienta STD 4 na serwerze STD (SPEC 4.2).

.DESCRIPTION
Kroki:
 1. Wczytuje plik konfiguracji (JSON bez podpisu) i sprawdza go (Test-StdRelease -Unpublished).
    Przy błędach przerywa z listą błędów i niczego nie zapisuje.
 2. Ustala numer rewizji: większy z numerów w opublikowanym już clients/<id>/release.sdt (sprawdzonym
    podpisem) i w pliku źródłowym, plus 1. Jeśli w archiwum (history) jest już plik z tym numerem,
    bierze następny wolny numer - komputery mogły już widzieć tamtą rewizję.
 3. Ustawia revision, publishedAt (UTC), publishedBy, expiresAt (publikacja + -ExpiresInDays dni)
    i changeNote (gdy podano -ChangeNote).
 4. Podpisuje dokument certyfikatem i zapisuje clients/<id>/history/<rev>.sdt oraz clients/<id>/release.sdt.
 5. Aktualizuje podpisaną listę klientów clients.sdt (dopisuje albo podmienia klienta, rewizja + 1).

Miejsce docelowe:
 - katalog drzewa /std/ serwera (lokalnie albo UNC), np. C:\STD-Server\data\std lub \\srv\STD$\std;
   brakujące katalogi są tworzone, pliki zapisywane atomowo (najpierw *.tmp, potem podmiana);
 - adres https://<serwer>/std-admin/ (WebDAV): MKCOL brakujących katalogów, potem PUT.
   Wymaga konta administratora STD (-Credential).

Najpierw powstają wszystkie podpisane pliki w pamięci, dopiero potem cokolwiek jest zapisywane.
Z -WhatIf skrypt sprawdza plik i serwer, pokazuje, co by opublikował, i niczego nie zapisuje.

W konfiguracji nie ma i nie może być loginu ani hasła do domeny ani hasła lokalnego administratora -
technik zawsze wpisuje je ręcznie przy wdrożeniu (walidacja odrzuca dokument, który je ma).
Hasło z -Credential służy tylko do logowania na serwer i nigdzie nie jest zapisywane.

.PARAMETER SourcePath
Plik konfiguracji klienta STD 4 (JSON bez podpisu), np. z edytora administratora
albo z Tools\Convert-StdLegacyConfig.ps1.

.PARAMETER Certificate
Certyfikat do podpisu (X509Certificate2 z kluczem prywatnym RSA).

.PARAMETER CertificateThumbprint
Odcisk certyfikatu do podpisu w magazynie Cert:\CurrentUser\My (tworzy go Tools\New-StdSigningCertificate.ps1).

.PARAMETER Destination
Katalog drzewa /std/ serwera (lokalny albo UNC, razem z podkatalogiem std) albo adres https://<serwer>/std-admin/.

.PARAMETER Credential
Konto administratora STD do publikacji przez https://. Dla ścieżek lokalnych i UNC nie jest używane.

.PARAMETER ChangeNote
Opis zmian w tej rewizji. Bez parametru zostaje opis z pliku źródłowego (jeśli jest).

.PARAMETER ExpiresInDays
Ważność konfiguracji w dniach od publikacji (1-3650, domyślnie 90). Po tym terminie komputery odrzucają
konfigurację, więc trzeba ją opublikować ponownie.

.PARAMETER PublishedBy
Login publikującego zapisywany w konfiguracji. Domyślnie bieżący użytkownik.

.EXAMPLE
.\Publish-StdRelease.ps1 -SourcePath .\lizard.release.v4.json -CertificateThumbprint AB12... -Destination https://pobieranie.firma.pl/std-admin/ -Credential (Get-Credential) -ChangeNote 'Nowa wersja 7-Zip'

.EXAMPLE
.\Publish-StdRelease.ps1 -SourcePath D:\STD4\lizard.json -CertificateThumbprint AB12... -Destination C:\STD-Server\data\std -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Thumbprint')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
    [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

    [Parameter(Mandatory = $true, ParameterSetName = 'Thumbprint')]
    [ValidateNotNullOrEmpty()]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,

    [System.Management.Automation.PSCredential]$Credential,

    [AllowEmptyString()]
    [string]$ChangeNote,

    [ValidateRange(1, 3650)]
    [int]$ExpiresInDays = 90,

    [string]$PublishedBy = $env:USERNAME
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Moduł STD (bez ponownego wczytania, jeśli ten sam moduł już działa - zostają odbiorniki logu wywołującego)
# ---------------------------------------------------------------------------------------------

$stdRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\', '/')
$stdManifest = Join-Path -Path $stdRoot -ChildPath 'STD.psd1'
$stdLoaded = $false
foreach ($loadedModule in @(Get-Module -Name STD)) {
    $loadedBase = [System.IO.Path]::GetFullPath([string]$loadedModule.ModuleBase).TrimEnd('\', '/')
    if ([string]::Equals($loadedBase, $stdRoot, [System.StringComparison]::OrdinalIgnoreCase)) { $stdLoaded = $true }
}
if (-not $stdLoaded) {
    Get-Module -Name STD | Remove-Module -Force -WhatIf:$false -Confirm:$false
    Import-Module -Name $stdManifest -Force -ErrorAction Stop
}

# ---------------------------------------------------------------------------------------------
# Pomocnicze
# ---------------------------------------------------------------------------------------------

$script:PubUtf8 = New-Object System.Text.UTF8Encoding($false)
$script:PubIdPattern = '^[a-z0-9][a-z0-9-]{1,39}$'
$script:PubMaxHistorySkip = 100

function Get-PubErrorText {
    # Krótki, zamaskowany tekst błędu do komunikatu.
    param($ErrorObject)
    $ex = $null
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ex = $ErrorObject.Exception }
    elseif ($ErrorObject -is [System.Exception]) { $ex = $ErrorObject }
    if ($null -eq $ex) { $text = [string]$ErrorObject } else {
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $text = [string]$ex.Message
    }
    if ($null -eq $text) { $text = '' }
    $text = ($text -replace '\s+', ' ').Trim()
    if ($text.Length -gt 400) { $text = $text.Substring(0, 400) + '...' }
    return (Protect-StdLogText -Text $text)
}

function Resolve-PubPath {
    # Ścieżka PowerShella (także względna i TestDrive:\) -> pełna ścieżka dla [IO.File]; UNC bez zmian.
    param([Parameter(Mandatory = $true)][string]$Path)
    if ($Path -match '^(\\\\|//)') { return $Path }
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Test-PubInteger {
    param($Value)
    if ($Value -is [int] -or $Value -is [long]) { return $true }
    if ($Value -is [double] -or $Value -is [decimal]) { return ([math]::Floor([double]$Value) -eq [double]$Value) }
    return $false
}

function Find-PubCertificate {
    # Certyfikat z kluczem prywatnym z magazynu CurrentUser\My po odcisku.
    param([Parameter(Mandatory = $true)][string]$Thumbprint)
    # Odcisk skopiowany z okna certyfikatu ma spacje i niewidoczne znaki - zostawiamy same cyfry HEX.
    $clean = ($Thumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    if ($clean.Length -ne 40) {
        throw "[STD-PUB-CERT] Nieprawidłowy odcisk certyfikatu '$Thumbprint' - podaj 40 znaków HEX (SHA-1), np. z wyniku Tools\New-StdSigningCertificate.ps1."
    }
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $found = $null
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $found = $store.Certificates.Find([System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $clean, $false)
    } catch {
        $found = $null
    } finally {
        $store.Close()
    }
    if ($null -eq $found -or $found.Count -eq 0) {
        throw "[STD-PUB-CERT] Nie ma certyfikatu $clean w magazynie Cert:\CurrentUser\My tego konta. Utwórz go skryptem Tools\New-StdSigningCertificate.ps1 albo podaj certyfikat w parametrze -Certificate."
    }
    return $found[0]
}

function Resolve-PubDestination {
    # Miejsce docelowe: @{ Kind = 'https' | 'path'; Base; Display }.
    param([Parameter(Mandatory = $true)][string]$Value)
    $text = $Value.Trim()
    if ($text -match '^[A-Za-z][A-Za-z0-9+.-]*://') {
        $uri = $null
        if (-not [System.Uri]::TryCreate($text, [System.UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') {
            throw "[STD-PUB-DEST] Publikacja przez sieć jest możliwa tylko przez https:// (np. https://pobieranie.firma.pl/std-admin/) - odrzucono: $(Protect-StdLogText -Text ($text -replace '(?<=://)[^/@\s]*@', '***@'))"
        }
        if (-not [string]::IsNullOrEmpty($uri.UserInfo)) {
            throw '[STD-PUB-DEST] Adres serwera nie może zawierać loginu ani hasła - podaj konto w parametrze -Credential.'
        }
        if (-not [string]::IsNullOrEmpty($uri.Query) -or -not [string]::IsNullOrEmpty($uri.Fragment)) {
            throw "[STD-PUB-DEST] Adres $text nie może mieć części ? ani # - podaj adres katalogu, np. https://pobieranie.firma.pl/std-admin/."
        }
        $base = $text
        if (-not $base.EndsWith('/')) { $base += '/' }
        return @{ Kind = 'https'; Base = $base; Display = $base }
    }

    $full = Resolve-PubPath -Path $text
    if ([System.IO.File]::Exists($full)) {
        throw "[STD-PUB-DEST] $full jest plikiem - podaj katalog drzewa /std/ serwera, np. C:\STD-Server\data\std."
    }
    # Częsta pomyłka: katalog danych serwera (DataPath) zamiast jego podkatalogu std.
    $stdChild = Join-StdLocation -Base $full -Child 'std'
    if ([System.IO.Directory]::Exists($stdChild) -and
        -not [System.IO.Directory]::Exists((Join-StdLocation -Base $full -Child 'clients')) -and
        -not [System.IO.File]::Exists((Join-StdLocation -Base $full -Child 'clients.sdt'))) {
        throw "[STD-PUB-DEST] Katalog $full wygląda na katalog danych serwera - konfiguracje leżą w jego podkatalogu std. Podaj -Destination $stdChild."
    }
    return @{ Kind = 'path'; Base = $full; Display = $full }
}

function Get-PubLocation {
    param([Parameter(Mandatory = $true)]$Target, [Parameter(Mandatory = $true)][string]$Relative)
    return (Join-StdLocation -Base $Target.Base -Child $Relative)
}

function Read-PubBytes {
    # Bajty pliku w miejscu docelowym albo $null, gdy pliku nie ma.
    param([Parameter(Mandatory = $true)]$Target, [Parameter(Mandatory = $true)][string]$Relative, $Session)
    $location = Get-PubLocation -Target $Target -Relative $Relative
    if ($Target.Kind -eq 'https') {
        $response = Invoke-StdHttpRequest -Uri $location -Method GET -Session $Session -AllowStatus @(404, 410)
        if ($response.StatusCode -eq 404 -or $response.StatusCode -eq 410) { return $null }
        $bytes = [byte[]]$response.Bytes
        return , $bytes
    }
    if ([System.IO.Directory]::Exists($location)) {
        throw "[STD-PUB-DEST] $location jest katalogiem, a powinien być plikiem."
    }
    if (-not [System.IO.File]::Exists($location)) { return $null }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($location)
    } catch {
        throw "[STD-PUB-DEST] Nie udało się odczytać pliku $location - $(Get-PubErrorText $_)"
    }
    return , $bytes
}

function Test-PubExists {
    # Czy plik istnieje w miejscu docelowym (https: HEAD).
    param([Parameter(Mandatory = $true)]$Target, [Parameter(Mandatory = $true)][string]$Relative, $Session)
    $location = Get-PubLocation -Target $Target -Relative $Relative
    if ($Target.Kind -eq 'https') {
        $response = Invoke-StdHttpRequest -Uri $location -Method HEAD -Session $Session -AllowStatus @(404, 410)
        return ($response.StatusCode -ge 200 -and $response.StatusCode -le 299)
    }
    return ([System.IO.File]::Exists($location) -or [System.IO.Directory]::Exists($location))
}

function Read-PubSignedDocument {
    # Sprawdza podpis koperty i zwraca dokument (ordered hashtable).
    param([Parameter(Mandatory = $true)][byte[]]$Bytes, [Parameter(Mandatory = $true)]$Trusted)
    $envelope = Read-StdSignedEnvelope -Bytes $Bytes -TrustedCertificates $Trusted
    $document = ConvertFrom-StdJson -Json $envelope.PayloadText
    if (-not ($document -is [System.Collections.IDictionary])) {
        throw '[STD-CFG-INVALID] Treść podpisanego pliku nie jest obiektem JSON.'
    }
    return @{ Document = $document; Kid = $envelope.Kid }
}

function Write-PubFile {
    # Atomowy zapis bajtów: <plik>.tmp, potem podmiana (Replace albo Move). Katalogi są tworzone.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][byte[]]$Bytes)
    $tmp = $Path + '.tmp'
    try {
        $dir = [System.IO.Path]::GetDirectoryName($Path)
        if (-not [string]::IsNullOrEmpty($dir) -and -not [System.IO.Directory]::Exists($dir)) {
            [void][System.IO.Directory]::CreateDirectory($dir)
        }
        [System.IO.File]::WriteAllBytes($tmp, $Bytes)
        if ([System.IO.File]::Exists($Path)) {
            try {
                [System.IO.File]::Replace($tmp, $Path, [NullString]::Value)
            } catch {
                # Replace nie działa np. na niektórych udziałach sieciowych - zwykła podmiana.
                [System.IO.File]::Copy($tmp, $Path, $true)
                [System.IO.File]::Delete($tmp)
            }
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
    } catch {
        try { if ([System.IO.File]::Exists($tmp)) { [System.IO.File]::Delete($tmp) } } catch { }
        throw "[STD-PUB-WRITE] Nie udało się zapisać pliku $Path - $(Get-PubErrorText $_)"
    }
}

function New-PubInvalidError {
    # Błąd [STD-PUB-INVALID] z listą błędów walidacji (lista także w TargetObject).
    param([Parameter(Mandatory = $true)][string]$Intro, [Parameter(Mandatory = $true)][string[]]$Errors)
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $Errors) { $lines.Add(' - ' + $item) }
    $message = "[STD-PUB-INVALID] $Intro ($($Errors.Count)) - nic nie zostało opublikowane:`n" + ($lines -join "`n")
    $exception = New-Object System.InvalidOperationException($message)
    return [System.Management.Automation.ErrorRecord]::new($exception, 'STD-PUB-INVALID', [System.Management.Automation.ErrorCategory]::InvalidData, $Errors)
}

# ---------------------------------------------------------------------------------------------
# 1. Plik źródłowy i walidacja
# ---------------------------------------------------------------------------------------------

$sourceFull = Resolve-PubPath -Path $SourcePath
if (-not [System.IO.File]::Exists($sourceFull)) {
    throw "[STD-PUB-SOURCE] Nie ma pliku konfiguracji $sourceFull."
}
try {
    $sourceText = [System.IO.File]::ReadAllText($sourceFull, $script:PubUtf8)
} catch {
    throw "[STD-PUB-SOURCE] Nie udało się odczytać pliku $sourceFull - $(Get-PubErrorText $_)"
}
try {
    $source = ConvertFrom-StdJson -Json $sourceText
} catch {
    throw "[STD-PUB-SOURCE] Plik $sourceFull nie jest poprawnym JSON - $(Get-PubErrorText $_)"
}
if (-not ($source -is [System.Collections.IDictionary])) {
    throw "[STD-PUB-SOURCE] Plik $sourceFull musi zawierać obiekt JSON z konfiguracją klienta."
}
if ($source['kind'] -is [string] -and $source['kind'] -ceq 'std.signed') {
    throw "[STD-PUB-SOURCE] Plik $sourceFull jest już podpisaną kopertą (*.sdt). Podaj plik konfiguracji JSON bez podpisu."
}

$sourceErrors = @(Test-StdRelease -Release $source -Unpublished)
if ($sourceErrors.Count -gt 0) {
    throw (New-PubInvalidError -Intro "Konfiguracja $sourceFull ma błędy" -Errors $sourceErrors)
}
$clientId = [string]$source['client']['id']
$clientName = [string]$source['client']['name']
$sourceRevision = 0
if ((Test-PubInteger $source['revision']) -and [double]$source['revision'] -ge 1) { $sourceRevision = [int][double]$source['revision'] }

# ---------------------------------------------------------------------------------------------
# Publikujący, miejsce docelowe, certyfikat
# ---------------------------------------------------------------------------------------------

$publisher = $PublishedBy
if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = $env:USERNAME }
if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = $env:USER }
if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = [System.Environment]::UserName }
if ([string]::IsNullOrWhiteSpace($publisher)) {
    throw '[STD-PUB-ARGS] Nie udało się ustalić loginu publikującego - podaj -PublishedBy.'
}
$publisher = $publisher.Trim()

$target = Resolve-PubDestination -Value $Destination
$session = $null
if ($target.Kind -eq 'https') {
    if ($null -eq $Credential) {
        throw '[STD-PUB-ARGS] Publikacja przez https:// wymaga konta administratora STD: dodaj -Credential (Get-Credential).'
    }
    $session = New-StdSession -Credential $Credential
    if ($target.Base -match '/std/$') {
        Write-Warning "Adres $($target.Base) to drzewo tylko do odczytu. Publikacja idzie zwykle przez /std-admin/ (tylko administratorzy)."
    }
} elseif ($null -ne $Credential) {
    Write-Warning "-Credential jest używany tylko dla adresu https:// - zapis do $($target.Display) idzie z uprawnieniami bieżącego konta."
}

if ($PSCmdlet.ParameterSetName -eq 'Thumbprint') {
    $Certificate = Find-PubCertificate -Thumbprint $CertificateThumbprint
}
if ($null -eq $Certificate) {
    throw '[STD-PUB-CERT] Podaj certyfikat do podpisu (-Certificate albo -CertificateThumbprint).'
}
$kid = ([string]$Certificate.Thumbprint).ToUpperInvariant()
if (-not $Certificate.HasPrivateKey) {
    throw "[STD-PUB-CERT] Certyfikat $kid nie ma klucza prywatnego - nie można nim podpisać konfiguracji. Użyj certyfikatu z Tools\New-StdSigningCertificate.ps1 na komputerze, na którym go utworzono."
}
if ($Certificate.NotAfter.ToUniversalTime() -lt [DateTime]::UtcNow) {
    Write-Warning "Certyfikat $kid wygasł $($Certificate.NotAfter.ToString('yyyy-MM-dd')). Podpis nadal zadziała, ale czas utworzyć nowy certyfikat (Tools\New-StdSigningCertificate.ps1)."
}

# Lista do sprawdzenia plików już opublikowanych: ten sam certyfikat (bez klucza prywatnego) oraz
# zaufane certyfikaty z Modules\Std.TrustedSigners.psd1 (po wymianie certyfikatu stare pliki są podpisane poprzednim).
$verifyCerts = New-Object 'System.Collections.Generic.List[System.Security.Cryptography.X509Certificates.X509Certificate2]'
$verifyCerts.Add([System.Security.Cryptography.X509Certificates.X509Certificate2]::new([byte[]]$Certificate.RawData))
$signerTrusted = $false
try {
    foreach ($trustedCert in @(Get-StdTrustedCertificates)) {
        if ($null -eq $trustedCert) { continue }
        if (([string]$trustedCert.Thumbprint).ToUpperInvariant() -ceq $kid) { $signerTrusted = $true } else { $verifyCerts.Add($trustedCert) }
    }
} catch {
    Write-Warning "Nie udało się wczytać listy zaufanych certyfikatów (Modules\Std.TrustedSigners.psd1): $(Get-PubErrorText $_)"
}
if (-not $signerTrusted) {
    Write-Warning "Certyfikatu $kid nie ma na liście zaufanych w Modules\Std.TrustedSigners.psd1. Komputery z tą wersją STD odrzucą tę konfigurację ([STD-SIG-UNTRUSTED]), dopóki nie dopiszesz wpisu (wypisuje go Tools\New-StdSigningCertificate.ps1)."
}
$verifyArray = $verifyCerts.ToArray()

# ---------------------------------------------------------------------------------------------
# 2. Numer rewizji
# ---------------------------------------------------------------------------------------------

$releaseRelative = "clients/$clientId/release.sdt"
$releaseLocation = Get-PubLocation -Target $target -Relative $releaseRelative
$existingRevision = 0
Write-Verbose "Sprawdzam opublikowaną konfigurację: $releaseLocation"
$existingBytes = Read-PubBytes -Target $target -Relative $releaseRelative -Session $session
if ($null -ne $existingBytes) {
    try {
        $existing = Read-PubSignedDocument -Bytes $existingBytes -Trusted $verifyArray
        $existingDoc = $existing.Document
        $existingClient = $null
        if ($existingDoc['client'] -is [System.Collections.IDictionary]) { $existingClient = $existingDoc['client']['id'] }
        if (-not ($existingDoc['kind'] -is [string]) -or $existingDoc['kind'] -cne 'std.release') {
            throw 'to nie jest konfiguracja klienta (kind różny od std.release).'
        }
        if (-not ($existingClient -is [string]) -or $existingClient -cne $clientId) {
            throw "to konfiguracja innego klienta ($existingClient)."
        }
        if (-not (Test-PubInteger $existingDoc['revision']) -or [double]$existingDoc['revision'] -lt 1 -or [double]$existingDoc['revision'] -gt [int]::MaxValue) {
            throw 'brak poprawnego numeru rewizji.'
        }
        $existingRevision = [int][double]$existingDoc['revision']
    } catch {
        throw "[STD-PUB-PREVIOUS] W miejscu docelowym jest już konfiguracja klienta '$clientId' ($releaseLocation), ale nie da się jej sprawdzić: $(Get-PubErrorText $_) Nie publikuję, bo nie wiadomo, którą rewizję widziały już komputery. Sprawdź ten plik (podpisany innym certyfikatem? uszkodzony?); jeśli jest zbędny, usuń go - numer rewizji zostanie wtedy ustalony także na podstawie archiwum history."
    }
}

$baseRevision = [math]::Max($existingRevision, $sourceRevision)
if ($baseRevision -ge [int]::MaxValue) {
    throw '[STD-PUB-PREVIOUS] Numer rewizji osiągnął wartość maksymalną - nie można opublikować kolejnej rewizji.'
}
$revision = $baseRevision + 1
$skipped = 0
while (Test-PubExists -Target $target -Relative "clients/$clientId/history/$revision.sdt" -Session $session) {
    Write-Warning "W archiwum jest już rewizja $revision klienta '$clientId' - komputery mogły ją widzieć, więc biorę następny numer."
    $skipped++
    if ($skipped -ge $script:PubMaxHistorySkip -or $revision -ge [int]::MaxValue) {
        throw "[STD-PUB-PREVIOUS] W archiwum clients/$clientId/history jest za dużo rewizji nowszych niż opublikowana konfiguracja. Sprawdź katalog i plik release.sdt klienta '$clientId'."
    }
    $revision++
}

# ---------------------------------------------------------------------------------------------
# 3. Dokument do podpisu
# ---------------------------------------------------------------------------------------------

$now = [DateTime]::UtcNow
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$publishedAt = $now.ToString('o', $invariant)
$expiresAt = $now.AddDays($ExpiresInDays).ToString('o', $invariant)

$noteSet = $PSBoundParameters.ContainsKey('ChangeNote')
$publishedKeys = @('schemaVersion', 'kind', 'client', 'revision', 'publishedAt', 'publishedBy', 'changeNote', 'expiresAt')
$document = [ordered]@{}
$document.Add('schemaVersion', $source['schemaVersion'])
$document.Add('kind', $source['kind'])
$document.Add('client', $source['client'])
$document.Add('revision', $revision)
$document.Add('publishedAt', $publishedAt)
$document.Add('publishedBy', $publisher)
if ($noteSet) {
    if (-not [string]::IsNullOrWhiteSpace($ChangeNote)) { $document.Add('changeNote', $ChangeNote.Trim()) }
} elseif ($source.Contains('changeNote') -and $null -ne $source['changeNote']) {
    $document.Add('changeNote', $source['changeNote'])
}
$document.Add('expiresAt', $expiresAt)
foreach ($entry in @($source.GetEnumerator())) {
    if ($publishedKeys -contains [string]$entry.Key) { continue }
    $document.Add($entry.Key, $entry.Value)
}

$payloadText = ConvertTo-StdJson -InputObject $document
$payloadBytes = $script:PubUtf8.GetBytes($payloadText)

# Sprawdzamy dokładnie to, co odczyta STD na komputerze (po ponownym wczytaniu JSON).
$finalErrors = @(Test-StdRelease -Release (ConvertFrom-StdJson -Json $payloadText))
if ($finalErrors.Count -gt 0) {
    throw (New-PubInvalidError -Intro "Konfiguracja po ustawieniu pól publikacji ma błędy" -Errors $finalErrors)
}

# ---------------------------------------------------------------------------------------------
# 4. Podpis konfiguracji
# ---------------------------------------------------------------------------------------------

$releaseEnvelope = New-StdSignedEnvelope -Payload $payloadBytes -Certificate $Certificate
$releaseBytes = $script:PubUtf8.GetBytes($releaseEnvelope)
$releaseSha256 = Get-StdSha256 -Bytes $releaseBytes

# ---------------------------------------------------------------------------------------------
# 5. Lista klientów
# ---------------------------------------------------------------------------------------------

$indexRelative = 'clients.sdt'
$indexLocation = Get-PubLocation -Target $target -Relative $indexRelative
$indexRevision = 0
$indexEntries = New-Object 'System.Collections.Generic.List[object]'
$clientEntry = [ordered]@{ id = $clientId; name = $clientName; release = $releaseRelative }
Write-Verbose "Sprawdzam listę klientów: $indexLocation"
$indexBytes = Read-PubBytes -Target $target -Relative $indexRelative -Session $session
if ($null -ne $indexBytes) {
    try {
        $existingIndex = (Read-PubSignedDocument -Bytes $indexBytes -Trusted $verifyArray).Document
        if (-not ($existingIndex['kind'] -is [string]) -or $existingIndex['kind'] -cne 'std.clients') {
            throw 'to nie jest lista klientów (kind różny od std.clients).'
        }
        if (-not (Test-PubInteger $existingIndex['revision']) -or [double]$existingIndex['revision'] -lt 1 -or [double]$existingIndex['revision'] -ge [int]::MaxValue) {
            throw 'brak poprawnego numeru rewizji.'
        }
        $existingClients = $existingIndex['clients']
        if ($null -ne $existingClients -and -not ($existingClients -is [System.Array])) {
            throw 'pole clients nie jest listą.'
        }
        $indexRevision = [int][double]$existingIndex['revision']
        $seenIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
        foreach ($item in @($existingClients)) {
            if ($null -eq $item) { continue }
            if (-not ($item -is [System.Collections.IDictionary]) -or -not ($item['id'] -is [string]) -or [string]::IsNullOrWhiteSpace($item['id'])) {
                throw 'lista ma wpis bez identyfikatora klienta.'
            }
            $itemId = [string]$item['id']
            if (-not $seenIds.Add($itemId)) { continue }
            if ($itemId -ceq $clientId) {
                # Ten sam klient: zostaje na swoim miejscu, z aktualną nazwą i ścieżką.
                $updated = [ordered]@{}
                foreach ($field in @($item.GetEnumerator())) { $updated[[string]$field.Key] = $field.Value }
                $updated['name'] = $clientName
                $updated['release'] = $releaseRelative
                $indexEntries.Add($updated)
            } else {
                $indexEntries.Add($item)
            }
        }
    } catch {
        throw "[STD-PUB-INDEX] W miejscu docelowym jest lista klientów ($indexLocation), ale nie da się jej sprawdzić: $(Get-PubErrorText $_) Nie publikuję, żeby nie podpisać listy, której nie da się zweryfikować. Sprawdź ten plik; jeśli jest zbędny, usuń go - zostanie utworzony od nowa (wtedy trzeba ponownie opublikować pozostałych klientów)."
    }
}
$hasClient = $false
foreach ($item in $indexEntries) { if ([string]$item['id'] -ceq $clientId) { $hasClient = $true } }
if (-not $hasClient) { $indexEntries.Add($clientEntry) }
$newIndexRevision = $indexRevision + 1

$index = [ordered]@{
    schemaVersion = 4
    kind          = 'std.clients'
    revision      = $newIndexRevision
    publishedAt   = $publishedAt
    clients       = $indexEntries.ToArray()
}
$indexPayload = $script:PubUtf8.GetBytes((ConvertTo-StdJson -InputObject $index))
$indexEnvelope = New-StdSignedEnvelope -Payload $indexPayload -Certificate $Certificate
$indexFileBytes = $script:PubUtf8.GetBytes($indexEnvelope)

# ---------------------------------------------------------------------------------------------
# 6. Zapis (najpierw archiwum, potem konfiguracja, na końcu lista klientów)
# ---------------------------------------------------------------------------------------------

$historyRelative = "clients/$clientId/history/$revision.sdt"
$files = @(
    @{ Relative = $historyRelative; Bytes = $releaseBytes }
    @{ Relative = $releaseRelative; Bytes = $releaseBytes }
    @{ Relative = $indexRelative; Bytes = $indexFileBytes }
)
$whatTarget = "$($target.Display) - klient $clientId ($clientName)"
$whatAction = "Publikacja konfiguracji: rewizja $revision, lista klientów rewizja $newIndexRevision, pliki: $(($files | ForEach-Object { $_.Relative }) -join ', ')"
if (-not $PSCmdlet.ShouldProcess($whatTarget, $whatAction)) {
    return
}

if ($target.Kind -eq 'https') {
    foreach ($dirRelative in @('clients/', "clients/$clientId/", "clients/$clientId/history/")) {
        # 405 = katalog już jest.
        $null = Invoke-StdHttpRequest -Uri (Get-PubLocation -Target $target -Relative $dirRelative) -Method MKCOL -Session $session -AllowStatus @(405)
    }
    foreach ($file in $files) {
        $null = Invoke-StdHttpRequest -Uri (Get-PubLocation -Target $target -Relative $file.Relative) -Method PUT -Session $session -Body $file.Bytes -ContentType 'application/json'
    }
} else {
    foreach ($file in $files) {
        Write-PubFile -Path (Get-PubLocation -Target $target -Relative $file.Relative) -Bytes $file.Bytes
    }
}

$message = "Opublikowano konfigurację klienta '$clientId' ($clientName): rewizja $revision, ważna do $($now.AddDays($ExpiresInDays).ToString('yyyy-MM-dd HH:mm', $invariant)) UTC, podpis $kid, $releaseLocation. Lista klientów: rewizja $newIndexRevision."
Write-StdLog -Level Info -Source 'Publish-StdRelease' -Message $message
Write-Host $message

return [ordered]@{
    ClientId      = $clientId
    Revision      = $revision
    ReleasePath   = $releaseLocation
    IndexRevision = $newIndexRevision
    Sha256        = $releaseSha256
}
