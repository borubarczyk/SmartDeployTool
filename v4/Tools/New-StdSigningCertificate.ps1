# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.

<#
.SYNOPSIS
Tworzy certyfikat do podpisu konfiguracji STD 4 (SPEC 4.1). Tylko Windows.

.DESCRIPTION
Tworzy certyfikat RSA 3072 / SHA-256 w magazynie Cert:\CurrentUser\My bieżącego konta
(użycie klucza: podpis cyfrowy, EKU "Document Signing" 1.3.6.1.4.1.311.10.3.12). Klucz prywatny
jest domyślnie nieeksportowalny: zostaje na tym koncie i tym komputerze.

Skrypt zapisuje certyfikat publiczny do pliku .cer i wypisuje gotowy wpis do
v4\Modules\Std.TrustedSigners.psd1. Dopiero po dopisaniu wpisu (i wydaniu tej wersji STD) komputery
przyjmują konfiguracje podpisane tym certyfikatem. Konfigurację podpisuje Tools\Publish-StdRelease.ps1
z parametrem -CertificateThumbprint <odcisk>.

Dobrze jest mieć na liście zaufanych dwa certyfikaty (główny i zapasowy, np. na dwóch komputerach
administratorów): gdy jeden przepadnie, drugi dalej działa.

Nie wymaga uprawnień administratora.

.PARAMETER Name
Nazwa certyfikatu (CN i nazwa przyjazna). Domyślnie 'STD - podpis konfiguracji'.

.PARAMETER Years
Ważność w latach (1-30, domyślnie 5).

.PARAMETER Exportable
Pozwala wyeksportować klucz prywatny (np. kopia zapasowa .pfx). Bez tego klucza nie da się skopiować.

.PARAMETER OutPath
Katalog na plik .cer (domyślnie bieżący katalog). Brakujący katalog jest tworzony.

.EXAMPLE
.\New-StdSigningCertificate.ps1

.EXAMPLE
.\New-StdSigningCertificate.ps1 -Name 'STD - podpis (zapasowy)' -Years 3 -OutPath D:\STD4\certyfikaty
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateNotNullOrEmpty()]
    [string]$Name = 'STD - podpis konfiguracji',

    [ValidateRange(1, 30)]
    [int]$Years = 5,

    [switch]$Exportable,

    [ValidateNotNullOrEmpty()]
    [string]$OutPath = '.'
)

$ErrorActionPreference = 'Stop'

$documentSigningOid = '1.3.6.1.4.1.311.10.3.12'
$storePath = 'Cert:\CurrentUser\My'

# --- Warunki ----------------------------------------------------------------------------------

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
    throw '[STD-CERT-OS] Certyfikat do podpisu konfiguracji tworzy się tylko w Windows (magazyn Cert:\CurrentUser\My). Uruchom ten skrypt w Windows PowerShell 5.1 na komputerze administratora STD.'
}
if ($null -eq (Get-Command -Name 'New-SelfSignedCertificate' -ErrorAction SilentlyContinue)) {
    throw '[STD-CERT-OS] Brak polecenia New-SelfSignedCertificate (moduł PKI). Potrzebny jest Windows 10, Windows Server 2016 albo nowszy.'
}
$Name = $Name.Trim()
if ($Name -match '[,=+<>#;"\\]') {
    throw "[STD-CERT-ARGS] Nazwa certyfikatu nie może zawierać znaków , = + < > # ; `" \ - podaj prostą nazwę, np. 'STD - podpis konfiguracji'."
}

$outDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutPath)
if ([System.IO.File]::Exists($outDir)) {
    throw "[STD-CERT-ARGS] $outDir jest plikiem - podaj katalog na plik .cer (-OutPath)."
}

# --- Ostrzeżenie o istniejących certyfikatach o tej samej nazwie -----------------------------

$subject = "CN=$Name"
$sameName = @(Get-ChildItem -Path $storePath | Where-Object { $_.Subject -eq $subject })
if ($sameName.Count -gt 0) {
    $list = ($sameName | ForEach-Object { '{0} (ważny do {1})' -f $_.Thumbprint, $_.NotAfter.ToString('yyyy-MM-dd') }) -join ', '
    Write-Warning "W $storePath są już certyfikaty o nazwie '$Name': $list. Nowy powstanie obok - stary usuń, gdy nie będzie potrzebny (i zdejmij go z Std.TrustedSigners.psd1)."
}

if (-not $PSCmdlet.ShouldProcess($storePath, "Utworzenie certyfikatu '$Name' (RSA 3072, SHA-256, ważny $Years lat)")) {
    return
}

# --- Certyfikat ------------------------------------------------------------------------------

$exportPolicy = 'NonExportable'
if ($Exportable) { $exportPolicy = 'Exportable' }
$certificateParams = @{
    Subject           = $subject
    FriendlyName      = $Name
    CertStoreLocation = $storePath
    Type              = 'Custom'
    Provider          = 'Microsoft Software Key Storage Provider'
    KeyAlgorithm      = 'RSA'
    KeyLength         = 3072
    HashAlgorithm     = 'SHA256'
    KeyUsage          = 'DigitalSignature'
    KeyExportPolicy   = $exportPolicy
    TextExtension     = @("2.5.29.37={text}$documentSigningOid")
    NotBefore         = (Get-Date).AddMinutes(-10)
    NotAfter          = (Get-Date).AddYears($Years)
}
try {
    $certificate = New-SelfSignedCertificate @certificateParams -WhatIf:$false -Confirm:$false
} catch {
    throw "[STD-CERT-CREATE] Nie udało się utworzyć certyfikatu: $($_.Exception.Message)"
}
$thumbprint = ([string]$certificate.Thumbprint).ToUpperInvariant()

# Próba podpisu: klucz musi działać tak, jak użyje go Publish-StdRelease (RSA PKCS#1 v1.5, SHA-256).
$testBytes = [System.Text.Encoding]::UTF8.GetBytes('std-signing-test')
$hashName = [System.Security.Cryptography.HashAlgorithmName]::SHA256
$padding = [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
$privateKey = $null
$publicKey = $null
try {
    $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
    $publicKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($certificate)
    $signature = $privateKey.SignData($testBytes, $hashName, $padding)
    if (-not $publicKey.VerifyData($testBytes, $signature, $hashName, $padding)) { throw 'podpis testowy się nie zgadza.' }
} catch {
    throw "[STD-CERT-CREATE] Certyfikat $thumbprint powstał w $storePath, ale próba podpisu się nie udała: $($_.Exception.Message) Usuń go (Remove-Item $storePath\$thumbprint) i uruchom skrypt ponownie."
} finally {
    if ($null -ne $privateKey) { $privateKey.Dispose() }
    if ($null -ne $publicKey) { $publicKey.Dispose() }
}

# --- Plik .cer i wpis do Std.TrustedSigners.psd1 ---------------------------------------------

if (-not [System.IO.Directory]::Exists($outDir)) {
    [void][System.IO.Directory]::CreateDirectory($outDir)
}
$fileBase = $Name
foreach ($invalid in [System.IO.Path]::GetInvalidFileNameChars()) { $fileBase = $fileBase.Replace([string]$invalid, '_') }
$fileBase = ($fileBase -replace '\s+', ' ').Trim()
if ([string]::IsNullOrEmpty($fileBase)) { $fileBase = 'STD - podpis' }
$cerPath = Join-Path -Path $outDir -ChildPath ('{0} {1}.cer' -f $fileBase, $thumbprint)
[System.IO.File]::WriteAllBytes($cerPath, $certificate.RawData)

$entryName = '{0} - {1}, {2}' -f $Name, $env:USERNAME, (Get-Date).ToString('yyyy-MM-dd')
$entry = "@{{ Name = '{0}'; Thumbprint = '{1}'; Certificate = '{2}' }}" -f $entryName.Replace("'", "''"), $thumbprint, [System.Convert]::ToBase64String($certificate.RawData)

Write-Host ''
Write-Host 'Utworzono certyfikat do podpisu konfiguracji STD.' -ForegroundColor Green
Write-Host ('  Nazwa:      {0}' -f $Name)
Write-Host ('  Odcisk:     {0}' -f $thumbprint)
Write-Host ('  Ważny do:   {0}' -f $certificate.NotAfter.ToString('yyyy-MM-dd'))
Write-Host ('  Magazyn:    {0}\{1}' -f $storePath, $thumbprint)
Write-Host ('  Plik .cer:  {0}' -f $cerPath)
Write-Host ''
Write-Host 'Dopisz ten wpis do v4\Modules\Std.TrustedSigners.psd1 (wewnątrz Signers = @( ... )):' -ForegroundColor Cyan
Write-Host ''
Write-Host ('        ' + $entry)
Write-Host ''
if ($Exportable) {
    Write-Host 'Klucz prywatny można wyeksportować. Zrób kopię zapasową (Export-PfxCertificate z mocnym hasłem),' -ForegroundColor Yellow
    Write-Host 'trzymaj ją w bezpiecznym miejscu i nie zostawiaj pliku .pfx na dysku.' -ForegroundColor Yellow
} else {
    Write-Host 'Klucza prywatnego nie da się wyeksportować: działa tylko na tym koncie i tym komputerze.'
    Write-Host 'Gdy przepadnie, utwórz nowy certyfikat i dopisz go do Std.TrustedSigners.psd1 (warto mieć też certyfikat zapasowy).'
}
Write-Host ''
Write-Host ('Publikacja: Tools\Publish-StdRelease.ps1 -SourcePath <plik.json> -CertificateThumbprint {0} -Destination <serwer>' -f $thumbprint)

[pscustomobject]@{
    Thumbprint         = $thumbprint
    Subject            = $certificate.Subject
    NotAfter           = $certificate.NotAfter
    StorePath          = '{0}\{1}' -f $storePath, $thumbprint
    CerPath            = $cerPath
    TrustedSignerEntry = $entry
}
