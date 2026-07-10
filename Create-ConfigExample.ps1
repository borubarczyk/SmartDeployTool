<#
.SYNOPSIS
Skrypt generujący zanonimizowany plik config.example.json na bazie istniejącego config.json.
.DESCRIPTION
Odczytuje aktualną konfigurację, usuwa z niej wrażliwe hasła, loginy, tokeny i ścieżki
wewnętrzne, po czym zapisuje jako plik przykładowy, który bezpiecznie można wysłać na GitHuba.
#>

$ScriptDir = $PSScriptRoot
$sourcePath = Join-Path $ScriptDir "config.json"
$targetPath = Join-Path $ScriptDir "config.example.json"

if (-not (Test-Path $sourcePath)) {
    Write-Warning "Nie znaleziono pliku źródłowego: $sourcePath"
    exit
}

Write-Host "Czytanie pliku $sourcePath..." -ForegroundColor Cyan
$config = Get-Content $sourcePath -Raw -Encoding UTF8 | ConvertFrom-Json

# 1. Maskowanie ścieżek instalacyjnych
if ($config.InstallSourcePaths.network) { $config.InstallSourcePaths.network = "\\SERVER\share\" }
if ($config.InstallSourcePaths.web) { $config.InstallSourcePaths.web = "https://example.com/apps/" }
if ($config.CustomWebDataLocation.URL) { $config.CustomWebDataLocation.URL = "https://example.com/custom/" }

# 2. Maskowanie poświadczeń do domeny i zasobów
if ($config.DomainJoin) {
    $config.DomainJoin.DomainName = "firma.local"
    $config.DomainJoin.Username = "ad\administrator"
}
if ($config.LocalAdmin) { $config.LocalAdmin.Username = "LocalAdmin" }
if ($config.WebAuth) {
    $config.WebAuth.Username = "admin"
    $config.WebAuth.Password = "***REDACTED***"
}

# 3. Maskowanie specyficznych pakietów
if ($config.TeamViewer.Arguments) { 
    $config.TeamViewer.Arguments = "/qn CUSTOMCONFIGID=1234567 APITOKEN=1234567-example ASSIGNMENTOPTIONS=`"--alias %COMPUTERNAME% --grant-easy-access`"" 
}
if ($config.AntyVirus) {
    $config.AntyVirus.InstallSourcePaths.network = "\\SERVER\AV\"
    $config.AntyVirus.InstallSourcePaths.web = "https://example.com/av/"
    $config.AntyVirus.Credentials.Username = "admin"
    $config.AntyVirus.Credentials.Password = "***REDACTED***"
}
if ($config.Programs -and $config.Programs.LanSweeper_Agent) {
    $config.Programs.LanSweeper_Agent.SilentArgs = "--mode unattended --server 192.168.1.100"
}

$config | ConvertTo-Json -Depth 10 | Set-Content -Path $targetPath -Encoding UTF8
Write-Host "Gotowe! Wygenerowano bezpieczny plik $targetPath" -ForegroundColor Green