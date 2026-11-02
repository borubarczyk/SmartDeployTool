# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.

<#
.SYNOPSIS
Konwersja konfiguracji STD 3.1 (config.json) do konfiguracji klienta STD 4 (SPEC 4.3).

.DESCRIPTION
Wywołuje ConvertFrom-StdLegacyConfig, zapisuje dokument STD 4 (JSON bez podpisu, UTF-8 bez BOM)
i wypisuje ostrzeżenia, listę sekretów do ustawienia oraz błędy walidacji (Test-StdRelease -Unpublished),
które trzeba poprawić przed publikacją.

Hasła i tokeny z 3.1 nie trafiają do wyniku: dane logowania (WebAuth, AntyVirus.Credentials,
DomainJoin.Username) są pomijane, a APITOKEN / CUSTOMCONFIGID z argumentów TeamViewer są zamieniane
na wstawki {{secret:...}}. Login i hasło do domeny oraz hasło lokalnego administratora technik zawsze
wpisuje ręcznie przy wdrożeniu.

Następny krok: ustawić sekrety, poprawić ostrzeżenia i opublikować plik (Tools\Publish-StdRelease.ps1).

.PARAMETER Path
Plik config.json z STD 3.1.

.PARAMETER ClientId
Identyfikator klienta STD 4: małe litery, cyfry i myślnik (2-40 znaków), np. lizard.

.PARAMETER ClientName
Nazwa klienta wyświetlana technikom, np. 'Lizard-Tech'.

.PARAMETER OutPath
Plik wynikowy. Domyślnie <ClientId>.release.v4.json w katalogu pliku config.json.
Istniejący plik jest zastępowany. Plik wynikowy nie może być plikiem config.json.

.EXAMPLE
.\Convert-StdLegacyConfig.ps1 -Path .\config.json -ClientId lizard -ClientName 'Lizard-Tech'

.EXAMPLE
.\Convert-StdLegacyConfig.ps1 -Path D:\STD31\config.json -ClientId lizard -ClientName 'Lizard-Tech' -OutPath D:\STD4\lizard.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    [Parameter(Mandatory = $true)]
    [string]$ClientName,
    [string]$OutPath
)

$ErrorActionPreference = 'Stop'

# Moduł STD. Bez ponownego wczytania, jeśli ten sam moduł już działa: ponowny Import-Module -Force
# zamknąłby odbiorniki logu wywołującego i skasował maski sekretów (jak w Tools\Publish-StdRelease.ps1).
$stdRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\', '/')
$manifest = Join-Path -Path $stdRoot -ChildPath 'STD.psd1'
$stdLoaded = $false
foreach ($loadedModule in @(Get-Module -Name STD)) {
    $loadedBase = [System.IO.Path]::GetFullPath([string]$loadedModule.ModuleBase).TrimEnd('\', '/')
    if ([string]::Equals($loadedBase, $stdRoot, [System.StringComparison]::OrdinalIgnoreCase)) { $stdLoaded = $true }
}
if (-not $stdLoaded) {
    Get-Module -Name STD | Remove-Module -Force
    Import-Module -Name $manifest -Force -ErrorAction Stop
}

$sourceFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
$result = ConvertFrom-StdLegacyConfig -Path $sourceFull -ClientId $ClientId -ClientName $ClientName

if ([string]::IsNullOrWhiteSpace($OutPath)) {
    $outFull = Join-Path -Path ([System.IO.Path]::GetDirectoryName($sourceFull)) -ChildPath ($ClientId + '.release.v4.json')
} else {
    $outFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutPath)
}
if ([string]::Equals([System.IO.Path]::GetFullPath($outFull), [System.IO.Path]::GetFullPath($sourceFull), [System.StringComparison]::OrdinalIgnoreCase)) {
    throw '[STD-LEGACY-ARG] Plik wynikowy nie może zastąpić pliku config.json z STD 3.1. Podaj inny -OutPath.'
}

Write-StdJsonFile -Path $outFull -InputObject $result.Document

Write-Host ''
Write-Host ('Zapisano konfigurację STD 4 (bez podpisu): {0}' -f $outFull)
Write-Host ('Klient: {0} ({1}), pozycji katalogu: {2}.' -f $ClientName, $ClientId, $result.Document['catalog'].get_Count())

$warnings = @($result.Warnings)
Write-Host ''
if ($warnings.Count -gt 0) {
    Write-Host ('Ostrzeżenia ({0}):' -f $warnings.Count) -ForegroundColor Yellow
    foreach ($warning in $warnings) { Write-Host (' - ' + $warning) -ForegroundColor Yellow }
} else {
    Write-Host 'Brak ostrzeżeń.'
}

$secrets = @($result.SecretsToSet)
Write-Host ''
if ($secrets.Count -gt 0) {
    Write-Host ('Sekrety do ustawienia przed publikacją ({0}):' -f $secrets.Count) -ForegroundColor Cyan
    foreach ($secret in $secrets) { Write-Host (' - ' + $secret) -ForegroundColor Cyan }
    Write-Host 'Ustaw je w sekcji secrets (Set-StdSealedItem albo edytor administratora). Wartości wpisz z konsoli TeamViewer - nie zostały przeniesione z 3.1.'
} else {
    Write-Host 'Brak sekretów do ustawienia.'
}

if (Get-Command -Name 'Test-StdRelease' -ErrorAction SilentlyContinue) {
    $errors = @(Test-StdRelease -Release $result.Document -Unpublished)
    Write-Host ''
    if ($errors.Count -gt 0) {
        Write-Host ('Do poprawienia przed publikacją - błędy walidacji ({0}):' -f $errors.Count) -ForegroundColor Yellow
        foreach ($validationError in $errors) { Write-Host (' - ' + $validationError) -ForegroundColor Yellow }
    } else {
        Write-Host 'Konfiguracja przechodzi walidację (Test-StdRelease -Unpublished).' -ForegroundColor Green
    }
}

Write-Host ''
Write-Host 'Następny krok: popraw ostrzeżenia, ustaw sekrety i opublikuj plik (Tools\Publish-StdRelease.ps1).'
