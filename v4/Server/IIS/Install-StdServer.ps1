<#
.SYNOPSIS
    Instalka serwera STD (Smart Tool for Deployment) dla Windows Server + IIS.

.DESCRIPTION
    Instaluje i konfiguruje IIS tak, żeby udostępniał drzewo /std/ (odczyt dla techników
    i administratorów, zapis raportów do /std/drop/) oraz /std-admin/ (publikacja, tylko
    administratorzy). Logowanie: Basic, wyłącznie przez HTTPS. Szczegóły: v4\SPEC.md, sekcje 5.1 i 5.2.

    Aktualizacja = ponowne uruchomienie tej instalki. Ustawienia z pierwszej instalacji są zapisane
    w C:\ProgramData\STD-Server\settings.json i przy kolejnym uruchomieniu są wartościami domyślnymi.
    Parametry podane w wierszu poleceń je nadpisują (i też są zapisywane).

    Instalka nigdy nie usuwa danych (<DataPath>\std), kont ani grup. Przed zmianą konfiguracji IIS
    robi kopię zapasową, a gdy test po zmianie się nie uda, przywraca ją.

    Skrypt nie przyjmuje i nie zapisuje żadnych haseł. Konta techników dodaje Add-StdServerUser.ps1.

.PARAMETER HostName
    Nazwa serwera w adresie (np. pobieranie.firma.pl). Domyślnie: pełna nazwa komputera.

.PARAMETER CertificateThumbprint
    Odcisk certyfikatu TLS w Cert:\LocalMachine\My. Domyślnie: najnowszy ważny certyfikat
    z kluczem prywatnym, który obejmuje HostName.

.PARAMETER Port
    Port HTTPS. Domyślnie 443.

.PARAMETER SiteName
    Nazwa witryny IIS. Domyślnie 'STD'.

.PARAMETER DataPath
    Katalog danych. Pliki STD leżą w <DataPath>\std. Domyślnie 'C:\STD-Server\data'.

.PARAMETER TechnicianGroup
    Lokalna grupa techników. Domyślnie 'STD-Technicy'.

.PARAMETER AdminGroup
    Lokalna grupa administratorów STD. Domyślnie 'STD-Administratorzy'.

.PARAMETER SkipFeatureInstall
    Nie instaluje funkcji Windows (IIS musi już być zainstalowany ze wszystkimi wymaganymi funkcjami).

.EXAMPLE
    .\Install-StdServer.ps1 -HostName pobieranie.firma.pl -CertificateThumbprint 0123456789ABCDEF0123456789ABCDEF01234567

.EXAMPLE
    .\Install-StdServer.ps1
    Aktualizacja: ponowne uruchomienie z ustawieniami zapisanymi przy instalacji.
#>
[CmdletBinding()]
param(
    [string]$HostName,
    [string]$CertificateThumbprint,
    [int]$Port,
    [string]$SiteName,
    [string]$DataPath,
    [string]$TechnicianGroup,
    [string]$AdminGroup,
    [switch]$SkipFeatureInstall,
    # Wewnętrzny: ustawiany przy samodzielnym ponownym uruchomieniu jako administrator (okno czeka na Enter).
    [Parameter(DontShow = $true)]
    [switch]$ElevatedRelaunch
)

# ---------------------------------------------------------------------------------------------
# Stałe
# ---------------------------------------------------------------------------------------------

$script:StdSrvDefaults = [ordered]@{
    HostName              = $null
    CertificateThumbprint = $null
    Port                  = 443
    SiteName              = 'STD'
    DataPath              = 'C:\STD-Server\data'
    TechnicianGroup       = 'STD-Technicy'
    AdminGroup            = 'STD-Administratorzy'
}
$script:StdSrvSettingKeys = @('HostName', 'CertificateThumbprint', 'Port', 'SiteName', 'DataPath', 'TechnicianGroup', 'AdminGroup')

# Funkcje Windows Server (Install-WindowsFeature). Web-Mgmt-Console i Web-Scripting-Tools dodajemy
# do listy z SPEC, bo dają konsolę IIS i moduł WebAdministration (Backup-WebConfiguration).
$script:StdSrvServerFeatures = @('Web-Server', 'Web-Static-Content', 'Web-Http-Redirect', 'Web-Basic-Auth', 'Web-Url-Auth',
    'Web-DAV-Publishing', 'Web-Filtering', 'Web-Mgmt-Tools', 'Web-Mgmt-Console', 'Web-Scripting-Tools')

# Odpowiedniki na Windows 10 / 11 (Enable-WindowsOptionalFeature).
$script:StdSrvClientFeatures = @('IIS-WebServerRole', 'IIS-WebServer', 'IIS-CommonHttpFeatures', 'IIS-StaticContent', 'IIS-HttpRedirect',
    'IIS-Security', 'IIS-BasicAuthentication', 'IIS-URLAuthorization', 'IIS-WebDAV', 'IIS-RequestFiltering',
    'IIS-WebServerManagementTools', 'IIS-ManagementConsole', 'IIS-ManagementScriptingTools')

# Typy MIME z SPEC 5.1. '.*' = każdy inny plik jako application/octet-stream (instalatory .msu, .7z, .appx...),
# bez tego IIS odpowiada 404.3 na nieznane rozszerzenia.
$script:StdSrvMimeTypes = [ordered]@{
    '.sdt'  = 'application/json'
    '.json' = 'application/json'
    '.msi'  = 'application/octet-stream'
    '.exe'  = 'application/octet-stream'
    '.msix' = 'application/octet-stream'
    '.zip'  = 'application/octet-stream'
    '.ps1'  = 'application/octet-stream'
    '.xml'  = 'application/octet-stream'
    '.cab'  = 'application/octet-stream'
    '.*'    = 'application/octet-stream'
}

$script:StdSrvMaxAdminUpload = [uint32]4294967295   # ok. 4 GB (największa wartość maxAllowedContentLength)
$script:StdSrvMaxDropUpload = [uint32]52428800      # 50 MB
$script:StdSrvIisAppId = '{4dc3e181-e14b-4a21-b022-59fc669b0914}'   # identyfikator aplikacji IIS w http.sys
$script:StdSrvLastBackup = $null

# ---------------------------------------------------------------------------------------------
# Funkcje pomocnicze (bez zależności od IIS - testowane lokalnie)
# ---------------------------------------------------------------------------------------------

function Write-StdSrvStep {
    param([string]$Message)
    Write-Host ''
    Write-Host ("==> " + $Message) -ForegroundColor Cyan
}

function Write-StdSrvInfo {
    param([string]$Message)
    Write-Host ("    " + $Message)
}

function Write-StdSrvWarn {
    param([string]$Message)
    Write-Host ("UWAGA: " + $Message) -ForegroundColor Yellow
}

function Get-StdSrvStateRoot {
    $base = $env:ProgramData
    if ([string]::IsNullOrEmpty($base)) { $base = [IO.Path]::GetTempPath() }
    return (Join-Path $base 'STD-Server')
}

function ConvertTo-StdSrvThumbprint {
    # Usuwa spacje i niewidoczne znaki (kopiowanie z certmgr.msc dodaje U+200E), zwraca HEX wielkimi literami.
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $clean = ($Value -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    if ($clean -notmatch '^[0-9A-F]{40}$') {
        throw "[STD-SRV-CERT] Odcisk certyfikatu '$Value' jest nieprawidłowy (oczekiwano 40 znaków HEX)."
    }
    return $clean
}

function Test-StdSrvHostName {
    # Nazwa DNS (SNI wymaga nazwy, nie adresu IP).
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name.Length -gt 253) { return $false }
    if ($Name -match '^\d{1,3}(\.\d{1,3}){3}$') { return $false }
    if ($Name.Contains(':')) { return $false }
    return ($Name -match '^(?=.{1,253}$)[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*$')
}

function Test-StdSrvGroupName {
    # Lokalna grupa (bez znaków zabronionych w Windows) albo istniejąca grupa domenowa 'DOMENA\grupa'.
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name.Length -gt 256) { return $false }
    $parts = $Name.Split('\')
    if ($parts.Count -gt 2) { return $false }
    foreach ($p in $parts) {
        if ($p.Length -eq 0 -or $p -match '["/\[\]:;|=,+*?<>@]' -or $p.Trim() -ne $p) { return $false }
    }
    return $true
}

function Test-StdSrvSiteName {
    param([string]$Name)
    return ($Name -match '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$')
}

function Test-StdSrvDictionaryKey {
    # Działa dla Hashtable, [ordered] i PSBoundParameters (Dictionary<string,object> nie ma publicznego Contains).
    param([System.Collections.IDictionary]$Dictionary, [string]$Key)
    if ($null -eq $Dictionary) { return $false }
    foreach ($k in $Dictionary.Keys) { if ([string]$k -eq $Key) { return $true } }
    return $false
}

function Get-StdSrvSavedValue {
    # Wartość z wczytanego settings.json (PSCustomObject albo IDictionary); brak = $null.
    param($Saved, [string]$Key)
    if ($null -eq $Saved) { return $null }
    if ($Saved -is [System.Collections.IDictionary]) {
        if (Test-StdSrvDictionaryKey -Dictionary $Saved -Key $Key) { return $Saved[$Key] }
        return $null
    }
    $prop = $Saved.PSObject.Properties[$Key]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function Resolve-StdSrvSettings {
    <#
        Łączy ustawienia: parametr z wiersza poleceń > zapisane settings.json > wartość domyślna.
        Zwraca @{ Settings = [ordered]; Sources = @{ klucz = 'parametr'|'zapisane'|'domyślne' } }.
    #>
    param(
        [System.Collections.IDictionary]$Bound,
        $Saved,
        [System.Collections.IDictionary]$Defaults
    )
    $result = [ordered]@{}
    $sources = [ordered]@{}
    foreach ($key in $script:StdSrvSettingKeys) {
        $value = $null
        $source = 'domyślne'
        if ((Test-StdSrvDictionaryKey -Dictionary $Bound -Key $key) -and $null -ne $Bound[$key] -and "$($Bound[$key])" -ne '' -and "$($Bound[$key])" -ne '0') {
            $value = $Bound[$key]
            $source = 'parametr'
        }
        else {
            $savedValue = Get-StdSrvSavedValue -Saved $Saved -Key $key
            if ($null -ne $savedValue -and "$savedValue" -ne '') {
                $value = $savedValue
                $source = 'zapisane'
            }
            elseif (Test-StdSrvDictionaryKey -Dictionary $Defaults -Key $key) {
                $value = $Defaults[$key]
            }
        }
        if ($key -eq 'Port' -and $null -ne $value) { $value = [int]$value }
        elseif ($null -ne $value) { $value = [string]$value }
        $result[$key] = $value
        $sources[$key] = $source
    }
    return @{ Settings = $result; Sources = $sources }
}

function Assert-StdSrvSettings {
    # Sprawdza ustawienia, które nie wymagają Windows. Rzuca [STD-SRV-PARAM] z opisem.
    param([System.Collections.IDictionary]$Settings)
    if (-not (Test-StdSrvHostName -Name $Settings.HostName)) {
        throw "[STD-SRV-PARAM] HostName '$($Settings.HostName)' jest nieprawidłowy. Podaj nazwę DNS serwera (np. pobieranie.firma.pl), nie adres IP."
    }
    if ($Settings.Port -lt 1 -or $Settings.Port -gt 65535) {
        throw "[STD-SRV-PARAM] Port $($Settings.Port) jest spoza zakresu 1-65535."
    }
    if ($Settings.Port -eq 80) {
        throw "[STD-SRV-PARAM] Port 80 jest zarezerwowany dla przekierowania HTTP na HTTPS. Wybierz inny port HTTPS (domyślnie 443)."
    }
    if (-not (Test-StdSrvSiteName -Name $Settings.SiteName)) {
        throw "[STD-SRV-PARAM] SiteName '$($Settings.SiteName)' jest nieprawidłowa (litery, cyfry, '.', '_', '-', do 64 znaków)."
    }
    if ([string]::IsNullOrWhiteSpace($Settings.DataPath) -or $Settings.DataPath -notmatch '^[A-Za-z]:\\') {
        throw "[STD-SRV-PARAM] DataPath '$($Settings.DataPath)' musi być pełną ścieżką lokalną, np. C:\STD-Server\data."
    }
    foreach ($g in 'TechnicianGroup', 'AdminGroup') {
        if (-not (Test-StdSrvGroupName -Name $Settings[$g])) {
            throw "[STD-SRV-PARAM] $g '$($Settings[$g])' jest nieprawidłowa nazwą grupy."
        }
    }
    if ($Settings.TechnicianGroup -eq $Settings.AdminGroup) {
        throw "[STD-SRV-PARAM] TechnicianGroup i AdminGroup muszą być różnymi grupami."
    }
}

function Get-StdSrvBaseUrl {
    # Adres https://<host>[:port]/ (port pomijany dla 443).
    param([string]$HostName, [int]$Port)
    $h = $HostName.ToLowerInvariant()
    if ($Port -eq 443) { return "https://$h/" }
    return "https://${h}:$Port/"
}

function Get-StdSrvBootstrapJson {
    # Gotowy bootstrap.json dla pendrive'a (SPEC 2.3). Nie jest sekretem.
    param([string]$HostName, [int]$Port)
    $doc = [ordered]@{
        server  = (Get-StdSrvBaseUrl -HostName $HostName -Port $Port) + 'std/'
        mirrors = @()
        auth    = [ordered]@{ mode = 'basic' }
        proxy   = $null
    }
    return (ConvertTo-Json -InputObject $doc -Depth 30 -Compress)
}

function Read-StdSrvSettingsFile {
    # Brak pliku -> $null. Uszkodzony plik -> błąd (bez zgadywania ustawień przy aktualizacji).
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $text = [IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))
    try {
        return (ConvertFrom-Json -InputObject $text -ErrorAction Stop)
    }
    catch {
        throw "[STD-SRV-SETTINGS] Plik ustawień '$Path' jest uszkodzony: $($_.Exception.Message). Popraw go albo przywróć z kopii w katalogu backup."
    }
}

function Write-StdSrvSettingsFile {
    # Zapis atomowy: <plik>.tmp (UTF-8 bez BOM), potem podmiana.
    param([string]$Path, [System.Collections.IDictionary]$Settings)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $json = ConvertTo-Json -InputObject $Settings -Depth 30
    $tmp = "$Path.tmp"
    [IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
    if (Test-Path -LiteralPath $Path) {
        # [NullString]::Value: PowerShell zamienia $null na '' w argumentach typu string, a Replace nie przyjmuje ''.
        [IO.File]::Replace($tmp, $Path, [System.Management.Automation.Language.NullString]::Value)
    }
    else {
        [IO.File]::Move($tmp, $Path)
    }
}

function New-StdSrvSettingsDocument {
    # Dokument settings.json: ustawienia + wersja + daty (bez haseł - instalka ich nie zna).
    param([System.Collections.IDictionary]$Settings, $Saved, [string]$Version, [datetime]$Now)
    $doc = [ordered]@{}
    foreach ($key in $script:StdSrvSettingKeys) { $doc[$key] = $Settings[$key] }
    $doc['InstalledVersion'] = $Version
    $installedAt = Get-StdSrvSavedValue -Saved $Saved -Key 'InstalledAt'
    if ([string]::IsNullOrEmpty("$installedAt")) { $installedAt = $Now.ToUniversalTime().ToString('o') }
    elseif ($installedAt -is [datetime]) { $installedAt = $installedAt.ToUniversalTime().ToString('o') }
    $doc['InstalledAt'] = [string]$installedAt
    $doc['UpdatedAt'] = $Now.ToUniversalTime().ToString('o')
    return $doc
}

function Format-StdSrvArgument {
    # Cytowanie argumentu dla powershell.exe -File (reguły CommandLineToArgvW).
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    if ($Value.Length -gt 0 -and $Value -notmatch '[\s"]') { return $Value }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    $backslashes = 0
    foreach ($ch in $Value.ToCharArray()) {
        if ($ch -eq [char]'\') { $backslashes++; continue }
        if ($ch -eq [char]'"') {
            [void]$sb.Append([char]'\', ($backslashes * 2 + 1))
            [void]$sb.Append('"')
        }
        else {
            if ($backslashes -gt 0) { [void]$sb.Append([char]'\', $backslashes) }
            [void]$sb.Append($ch)
        }
        $backslashes = 0
    }
    if ($backslashes -gt 0) { [void]$sb.Append([char]'\', ($backslashes * 2)) }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Get-StdSrvRelaunchArguments {
    # Argumenty do ponownego uruchomienia skryptu z tymi samymi parametrami (bez haseł - tu ich nie ma).
    param([string]$ScriptPath, [System.Collections.IDictionary]$Bound, [switch]$MarkElevated)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($a in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File')) { $list.Add($a) }
    $list.Add((Format-StdSrvArgument -Value $ScriptPath))
    foreach ($key in $Bound.Keys) {
        if ($key -eq 'ElevatedRelaunch') { continue }
        $value = $Bound[$key]
        if ($value -is [System.Management.Automation.SwitchParameter]) {
            if ($value.IsPresent) { $list.Add("-$key") }
            continue
        }
        if ($value -is [bool]) {
            if ($value) { $list.Add("-$key") }
            continue
        }
        $list.Add("-$key")
        $list.Add((Format-StdSrvArgument -Value ([string]$value)))
    }
    if ($MarkElevated) { $list.Add('-ElevatedRelaunch') }
    return , $list.ToArray()
}

function Join-StdSrvWinPath {
    # Łączenie ścieżek Windows bez sprawdzania dysku (Join-Path wymaga istniejącego dysku).
    param([string]$Base, [string]$Child)
    return ($Base.TrimEnd('\') + '\' + $Child.TrimStart('\'))
}

function Get-StdSrvAclPlan {
    <#
        Lista wywołań icacls (każde to tablica argumentów). SID-y zamiast nazw: nie zależą od języka systemu.

        <std> (ochrona dziedziczenia, jawne wpisy dziedziczone w dół):
          SYSTEM, Administrators     (OI)(CI)F   pełna kontrola
          grupa administratorów STD  (OI)(CI)M   modyfikacja (PUT, DELETE, MKCOL przez /std-admin/)
          grupa techników            (OI)(CI)RX  odczyt i listowanie (GET /std/...)
          IIS_IUSRS                  (OI)(CI)RX  proces IIS czyta katalogi (szukanie web.config) - treść
                                                 i tak jest czytana w imieniu zalogowanego użytkownika

        <std>\drop (własna lista, bez dziedziczenia z <std>, więc technicy nie mają tu odczytu):
          SYSTEM, Administrators     (OI)(CI)F
          grupa administratorów STD  (OI)(CI)M   odczyt i porządki raportów przez /std-admin/drop/
          grupa techników            (CI)(WD,X,RA,S)  tylko katalogi: WD = "utwórz pliki", X = przejście,
                                                 RA = odczyt atrybutów (IIS sprawdza, że katalog istnieje),
                                                 S = synchronizacja. Bez RD (listowanie), bez AD (podkatalogi),
                                                 bez prawa do istniejących plików (brak OI) - nie widzą, nie czytają
                                                 i nie nadpisują cudzych raportów.
          CREATOR OWNER              (OI)(IO)(WD,AD,WEA,WA,RA,S,DE)  tylko na nowych plikach i tylko dla ich
                                                 autora: zapis treści i atrybutów (IIS może otworzyć świeżo
                                                 utworzony plik ponownie), bez odczytu treści.
          IIS_IUSRS                  (CI)(RX)    tylko katalogi (web.config), bez odczytu plików

        Witryna (pusty katalog główny) i katalog ustawień STD-Server: tylko SYSTEM i Administrators
        (+ IIS_IUSRS odczyt katalogu głównego witryny), żeby nikt nie podłożył tam web.config.
    #>
    param(
        [string]$StdPath,
        [string]$SiteRoot,
        [string]$StateRoot,
        [string]$TechnicianSid,
        [string]$AdminSid
    )
    $sys = '*S-1-5-18'
    $adm = '*S-1-5-32-544'
    $iis = '*S-1-5-32-568'
    $creator = '*S-1-3-0'
    $tech = '*' + $TechnicianSid
    $stdAdm = '*' + $AdminSid
    $drop = Join-StdSrvWinPath -Base $StdPath -Child 'drop'
    $targets = New-Object System.Collections.Generic.List[object]
    $targets.Add(@($StdPath, "${sys}:(OI)(CI)F", "${adm}:(OI)(CI)F", "${stdAdm}:(OI)(CI)M", "${tech}:(OI)(CI)RX", "${iis}:(OI)(CI)RX"))
    $targets.Add(@($drop, "${sys}:(OI)(CI)F", "${adm}:(OI)(CI)F", "${stdAdm}:(OI)(CI)M", "${tech}:(CI)(WD,X,RA,S)", "${creator}:(OI)(IO)(WD,AD,WEA,WA,RA,S,DE)", "${iis}:(CI)(RX)"))
    if (-not [string]::IsNullOrEmpty($SiteRoot)) {
        $targets.Add(@($SiteRoot, "${sys}:(OI)(CI)F", "${adm}:(OI)(CI)F", "${iis}:(OI)(CI)RX"))
    }
    if (-not [string]::IsNullOrEmpty($StateRoot)) {
        $targets.Add(@($StateRoot, "${sys}:(OI)(CI)F", "${adm}:(OI)(CI)F"))
    }
    # Dla każdego katalogu dwa wywołania: najpierw jawne wpisy (/grant:r podmienia wpisy tych SID-ów),
    # potem /inheritance:r (usuwa odziedziczone, np. "Użytkownicy uwierzytelnieni: modyfikacja" z C:\).
    # W tej kolejności katalog ani przez chwilę nie zostaje bez uprawnień.
    $plan = New-Object System.Collections.Generic.List[object]
    foreach ($t in $targets) {
        $grant = @($t[0], '/grant:r') + @($t[1..($t.Count - 1)]) + @('/c', '/q')
        $plan.Add($grant)
        $plan.Add(@($t[0], '/inheritance:r', '/c', '/q'))
    }
    return , $plan.ToArray()
}

function Get-StdSrvLocationPlan {
    <#
        Docelowa konfiguracja IIS jako dane (testowane lokalnie, stosowane przez Set-StdSrvIisConfiguration).
        Każdy wpis: @{ Location; Authorization (lista reguł allow albo $null = nie ruszaj); AuthoringRules; MaxUpload }.
        Lista reguł pusta = nikt (domyślna reguła "allow *" z poziomu serwera jest czyszczona).
    #>
    param([string]$SiteName, [string]$TechnicianGroup, [string]$AdminGroup)
    $both = "$TechnicianGroup,$AdminGroup"
    return @(
        @{ Location = $SiteName; Authorization = @(); AuthoringRules = @(); MaxUpload = $null }
        @{ Location = "$SiteName/std"; Authorization = @(@{ Roles = $both; Verbs = 'GET,HEAD' }); AuthoringRules = @(); MaxUpload = $null }
        @{ Location = "$SiteName/std/drop"; Authorization = @(@{ Roles = $both; Verbs = 'PUT' }); AuthoringRules = @(@{ Roles = $both; Access = 'Write' }); MaxUpload = $script:StdSrvMaxDropUpload }
        @{ Location = "$SiteName/std-admin"; Authorization = @(@{ Roles = $AdminGroup; Verbs = '' }); AuthoringRules = @(@{ Roles = $AdminGroup; Access = 'Read, Write, Source' }); MaxUpload = $script:StdSrvMaxAdminUpload }
    )
}

# ---------------------------------------------------------------------------------------------
# Funkcje systemowe (Windows)
# ---------------------------------------------------------------------------------------------

function Invoke-StdSrvNative {
    # Uruchamia program (icacls, netsh) i zwraca @{ ExitCode; Output } bez przerywania na stderr.
    param([string]$FilePath, [string[]]$Arguments)
    $ErrorActionPreference = 'Continue'
    $output = & $FilePath @Arguments 2>&1 | ForEach-Object { "$_" }
    return @{ ExitCode = $LASTEXITCODE; Output = (@($output) -join [Environment]::NewLine) }
}

function Test-StdSrvIsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-StdSrvPowerShellPath {
    # 64-bitowy Windows PowerShell także z 32-bitowego procesu (sysnative), bo konfiguracja IIS jest 64-bitowa.
    $sysnative = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess -and (Test-Path -LiteralPath $sysnative)) {
        return $sysnative
    }
    return (Join-Path $env:windir 'System32\WindowsPowerShell\v1.0\powershell.exe')
}

function Invoke-StdSrvRelaunch {
    <#
        Uruchamia skrypt ponownie (jako administrator albo w 64-bitowym PowerShellu) i czeka na koniec.
        Zwraca kod wyjścia. Odmowa w oknie UAC -> [STD-SRV-ADMIN].
    #>
    param([string]$ScriptPath, [System.Collections.IDictionary]$Bound, [switch]$Elevate)
    $exe = Get-StdSrvPowerShellPath
    $argList = Get-StdSrvRelaunchArguments -ScriptPath $ScriptPath -Bound $Bound -MarkElevated:$Elevate
    try {
        if ($Elevate) {
            $p = Start-Process -FilePath $exe -ArgumentList $argList -Verb RunAs -Wait -PassThru -ErrorAction Stop
        }
        else {
            $p = Start-Process -FilePath $exe -ArgumentList $argList -NoNewWindow -Wait -PassThru -ErrorAction Stop
        }
    }
    catch {
        $inner = $_.Exception
        while ($null -ne $inner.InnerException) { $inner = $inner.InnerException }
        if ($inner -is [System.ComponentModel.Win32Exception] -and $inner.NativeErrorCode -eq 1223) {
            throw '[STD-SRV-ADMIN] Instalka wymaga uprawnień administratora, a uruchomienie jako administrator zostało anulowane.'
        }
        throw "[STD-SRV-ADMIN] Nie udało się uruchomić instalki jako administrator: $($inner.Message)"
    }
    return $p.ExitCode
}

function Install-StdSrvFeatures {
    # Instaluje brakujące funkcje IIS. Idempotentne: zainstalowane pomija.
    if (Get-Command -Name Install-WindowsFeature -ErrorAction SilentlyContinue) {
        $missing = @(Get-WindowsFeature -Name $script:StdSrvServerFeatures | Where-Object { -not $_.Installed } | ForEach-Object { $_.Name })
        $unknown = @($script:StdSrvServerFeatures | Where-Object { $null -eq (Get-WindowsFeature -Name $_) })
        if ($unknown.Count -gt 0) { Write-StdSrvWarn ("Ten system nie zna funkcji: " + ($unknown -join ', ')) }
        if ($missing.Count -eq 0) {
            Write-StdSrvInfo 'Wszystkie funkcje IIS są już zainstalowane.'
            return
        }
        Write-StdSrvInfo ('Instaluję: ' + ($missing -join ', ') + ' (to może potrwać kilka minut)')
        $result = Install-WindowsFeature -Name $missing -IncludeManagementTools -ErrorAction Stop
        if (-not $result.Success) {
            throw "[STD-SRV-FEATURE] Instalacja funkcji Windows nie powiodła się (wynik: $($result.ExitCode))."
        }
        if ("$($result.RestartNeeded)" -eq 'Yes') {
            Write-StdSrvWarn 'System wymaga ponownego uruchomienia po instalacji funkcji. Uruchom ponownie serwer i jeszcze raz tę instalkę.'
        }
        return
    }
    if (Get-Command -Name Enable-WindowsOptionalFeature -ErrorAction SilentlyContinue) {
        $restart = $false
        foreach ($name in $script:StdSrvClientFeatures) {
            $f = Get-WindowsOptionalFeature -Online -FeatureName $name -ErrorAction SilentlyContinue
            if ($null -eq $f) { Write-StdSrvWarn "Ten system nie zna funkcji $name."; continue }
            if ("$($f.State)" -eq 'Enabled') { continue }
            Write-StdSrvInfo "Włączam $name"
            $r = Enable-WindowsOptionalFeature -Online -FeatureName $name -All -NoRestart -ErrorAction Stop
            if ($r.RestartNeeded) { $restart = $true }
        }
        if ($restart) { Write-StdSrvWarn 'System wymaga ponownego uruchomienia po włączeniu funkcji IIS.' }
        return
    }
    throw '[STD-SRV-FEATURE] Brak Install-WindowsFeature i Enable-WindowsOptionalFeature. Zainstaluj IIS ręcznie i uruchom instalkę z -SkipFeatureInstall.'
}

function Get-StdSrvCertificateNames {
    # Nazwy DNS z certyfikatu (SAN przez DnsNameList z dostawcy Cert:, plus CN).
    param($Certificate)
    $names = New-Object System.Collections.Generic.List[string]
    $prop = $Certificate.PSObject.Properties['DnsNameList']
    if ($null -ne $prop -and $null -ne $prop.Value) {
        foreach ($n in $prop.Value) { if ($n) { $names.Add(("$($n.Unicode)").ToLowerInvariant()) } }
    }
    $cn = $Certificate.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::DnsName, $false)
    if ($cn) { $names.Add($cn.ToLowerInvariant()) }
    return , $names.ToArray()
}

function Test-StdSrvNameMatch {
    param([string[]]$Names, [string]$HostName)
    $h = $HostName.ToLowerInvariant()
    foreach ($n in $Names) {
        if ($n -eq $h) { return $true }
        if ($n.StartsWith('*.') -and $h.EndsWith($n.Substring(1)) -and $h.Substring(0, $h.Length - $n.Length + 1).IndexOf('.') -lt 0) { return $true }
    }
    return $false
}

function Resolve-StdSrvCertificate {
    # Zwraca certyfikat z Cert:\LocalMachine\My (po odcisku albo dobrany do HostName).
    param([string]$Thumbprint, [string]$HostName)
    $store = 'Cert:\LocalMachine\My'
    if ($Thumbprint) {
        $cert = Get-ChildItem -Path $store | Where-Object { $_.Thumbprint.ToUpperInvariant() -eq $Thumbprint } | Select-Object -First 1
        if ($null -eq $cert) {
            throw "[STD-SRV-CERT] Nie ma certyfikatu $Thumbprint w $store. Zaimportuj certyfikat (z kluczem prywatnym) do magazynu komputera."
        }
    }
    else {
        $now = Get-Date
        $candidates = @(Get-ChildItem -Path $store | Where-Object {
                $_.HasPrivateKey -and $_.NotBefore -le $now -and $_.NotAfter -gt $now -and (Test-StdSrvNameMatch -Names (Get-StdSrvCertificateNames -Certificate $_) -HostName $HostName)
            } | Sort-Object -Property NotAfter -Descending)
        if ($candidates.Count -eq 0) {
            throw "[STD-SRV-CERT] Nie znaleziono certyfikatu dla '$HostName' w $store. Zaimportuj certyfikat serwera i podaj -CertificateThumbprint <odcisk>."
        }
        $cert = $candidates[0]
        Write-StdSrvInfo "Wybrano certyfikat $($cert.Thumbprint) ($($cert.Subject), ważny do $($cert.NotAfter.ToString('yyyy-MM-dd')))."
    }
    if (-not $cert.HasPrivateKey) {
        throw "[STD-SRV-CERT] Certyfikat $($cert.Thumbprint) nie ma klucza prywatnego. Zaimportuj go z plikiem .pfx."
    }
    if ($cert.NotAfter -le (Get-Date)) {
        throw "[STD-SRV-CERT] Certyfikat $($cert.Thumbprint) wygasł $($cert.NotAfter.ToString('yyyy-MM-dd')). Zaimportuj nowy i podaj -CertificateThumbprint."
    }
    if ($cert.NotAfter -le (Get-Date).AddDays(30)) {
        Write-StdSrvWarn "Certyfikat $($cert.Thumbprint) wygasa $($cert.NotAfter.ToString('yyyy-MM-dd')). Wymień go: nowy certyfikat, potem instalka z -CertificateThumbprint."
    }
    if (-not (Test-StdSrvNameMatch -Names (Get-StdSrvCertificateNames -Certificate $cert) -HostName $HostName)) {
        Write-StdSrvWarn "Certyfikat $($cert.Thumbprint) nie obejmuje nazwy '$HostName'. STD odrzuci takie połączenie, jeśli certyfikat nie pasuje do adresu."
    }
    return $cert
}

function Get-StdSrvSid {
    param([string]$Name)
    try {
        $account = New-Object System.Security.Principal.NTAccount($Name)
        return $account.Translate([System.Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        return $null
    }
}

function Initialize-StdSrvGroup {
    # Tworzy lokalną grupę, jeśli jej nie ma (grupy domenowe 'DOMENA\grupa' muszą istnieć). Zwraca SID.
    param([string]$Name, [string]$Description)
    $local = $Name
    if ($Name.Contains('\')) {
        $prefix, $rest = $Name.Split([char[]]@([char]'\'), 2)
        if ($prefix -ne $env:COMPUTERNAME -and $prefix -ne '.') {
            $sid = Get-StdSrvSid -Name $Name
            if (-not $sid) { throw "[STD-SRV-GROUP] Grupa '$Name' nie istnieje. Grupy domenowe trzeba utworzyć w domenie." }
            Write-StdSrvInfo "Grupa $Name (domenowa) istnieje."
            return $sid
        }
        $local = $rest
    }
    $sid = Get-StdSrvSid -Name "$env:COMPUTERNAME\$local"
    if ($sid) {
        Write-StdSrvInfo "Grupa $local już istnieje."
        return $sid
    }
    if (Get-Command -Name New-LocalGroup -ErrorAction SilentlyContinue) {
        [void](New-LocalGroup -Name $local -Description $Description -ErrorAction Stop)
    }
    else {
        $computer = [ADSI]"WinNT://$env:COMPUTERNAME,computer"
        $group = $computer.Create('group', $local)
        [void]$group.Put('Description', $Description)
        [void]$group.SetInfo()
    }
    Write-StdSrvInfo "Utworzono grupę $local."
    $sid = Get-StdSrvSid -Name "$env:COMPUTERNAME\$local"
    if (-not $sid) { throw "[STD-SRV-GROUP] Nie udało się odczytać SID grupy $local." }
    return $sid
}

function Initialize-StdSrvFolders {
    # Tworzy katalogi danych. Nic nie usuwa.
    param([string]$StdPath, [string]$SiteRoot, [string]$StateRoot)
    foreach ($dir in @($StateRoot, (Join-Path $StateRoot 'backup'), (Join-Path $StateRoot 'logs'), $SiteRoot, $StdPath,
            (Join-Path $StdPath 'clients'), (Join-Path $StdPath 'files'), (Join-Path $StdPath 'files\common'),
            (Join-Path $StdPath 'tool'), (Join-Path $StdPath 'drop'))) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
            [void](New-Item -ItemType Directory -Path $dir -Force)
            Write-StdSrvInfo "Utworzono $dir"
        }
    }
}

function Set-StdSrvAcl {
    param([object[]]$Plan, [string[]]$RemoveSids, [string]$StdPath)
    foreach ($oldSid in @($RemoveSids | Where-Object { $_ })) {
        foreach ($p in @($StdPath, (Join-StdSrvWinPath -Base $StdPath -Child 'drop'))) {
            [void](Invoke-StdSrvNative -FilePath 'icacls.exe' -Arguments @($p, '/remove:g', "*$oldSid", '/c', '/q'))
        }
    }
    foreach ($call in $Plan) {
        $r = Invoke-StdSrvNative -FilePath 'icacls.exe' -Arguments $call
        if ($r.ExitCode -ne 0) {
            throw "[STD-SRV-ACL] icacls $($call[0]) zakończył się kodem $($r.ExitCode): $($r.Output)"
        }
        Write-StdSrvInfo "Uprawnienia NTFS: $($call[0])"
    }
}

function Get-StdSrvInetsrvConfigPath {
    return (Join-Path $env:windir 'System32\inetsrv\config')
}

function New-StdSrvBackup {
    <#
        Kopia przed zmianą: settings.json, pliki konfiguracji IIS, listy ACL danych, wiązania http.sys,
        plus Backup-WebConfiguration. Zwraca @{ Path; WebBackupName }.
    #>
    param([string]$StateRoot, [string]$StdPath, [string]$SettingsPath)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $dir = Join-Path (Join-Path $StateRoot 'backup') $stamp
    [void](New-Item -ItemType Directory -Path $dir -Force)
    if (Test-Path -LiteralPath $SettingsPath) { Copy-Item -LiteralPath $SettingsPath -Destination $dir -Force }
    $cfgDir = Get-StdSrvInetsrvConfigPath
    $iisCopy = Join-Path $dir 'inetsrv'
    [void](New-Item -ItemType Directory -Path $iisCopy -Force)
    foreach ($f in 'applicationHost.config', 'administration.config', 'redirection.config') {
        $src = Join-Path $cfgDir $f
        if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination $iisCopy -Force }
    }
    if (Test-Path -LiteralPath $StdPath) {
        [void](Invoke-StdSrvNative -FilePath 'icacls.exe' -Arguments @($StdPath, '/save', (Join-Path $dir 'acl-std.txt'), '/t', '/c', '/q'))
    }
    $ssl = Invoke-StdSrvNative -FilePath 'netsh.exe' -Arguments @('http', 'show', 'sslcert')
    [IO.File]::WriteAllText((Join-Path $dir 'http-sslcert.txt'), $ssl.Output, (New-Object System.Text.UTF8Encoding($false)))
    $webName = $null
    if (Get-Command -Name Backup-WebConfiguration -ErrorAction SilentlyContinue) {
        try {
            $webName = "STD-$stamp"
            [void](Backup-WebConfiguration -Name $webName -ErrorAction Stop)
        }
        catch {
            Write-StdSrvWarn "Backup-WebConfiguration nie zadziałał ($($_.Exception.Message)). Zostaje kopia plików w $iisCopy."
            $webName = $null
        }
    }
    Write-StdSrvInfo "Kopia zapasowa: $dir"
    return @{ Path = $dir; WebBackupName = $webName }
}

function Restore-StdSrvBackup {
    # Przywraca konfigurację IIS z kopii (Restore-WebConfiguration albo kopia applicationHost.config).
    param([hashtable]$Backup)
    if ($null -eq $Backup) { return }
    try {
        if ($Backup.WebBackupName -and (Get-Command -Name Restore-WebConfiguration -ErrorAction SilentlyContinue)) {
            Restore-WebConfiguration -Name $Backup.WebBackupName -ErrorAction Stop
        }
        else {
            $src = Join-Path (Join-Path $Backup.Path 'inetsrv') 'applicationHost.config'
            if (Test-Path -LiteralPath $src) {
                Copy-Item -LiteralPath $src -Destination (Get-StdSrvInetsrvConfigPath) -Force
            }
        }
        Write-StdSrvWarn "Przywrócono konfigurację IIS z kopii $($Backup.Path)."
    }
    catch {
        Write-StdSrvWarn "Nie udało się przywrócić konfiguracji IIS: $($_.Exception.Message). Kopia jest w $($Backup.Path)."
    }
}

# ---------------------------------------------------------------------------------------------
# IIS (Microsoft.Web.Administration - ta sama biblioteka, której używają moduł WebAdministration
# i appcmd). Cała konfiguracja trafia do applicationHost.config w znacznikach <location> i jest
# zapisywana jednym CommitChanges(). Sekcje zablokowane domyślnie (anonymousAuthentication,
# basicAuthentication, access, webdav) da się ustawiać w <location> w applicationHost.config, więc
# NIE odblokowujemy ich globalnie - plik web.config w danych nie może zmienić logowania.
# ---------------------------------------------------------------------------------------------

function Import-StdSrvIisApi {
    if (Get-Module -ListAvailable -Name WebAdministration) {
        Import-Module WebAdministration -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
    if (-not ('Microsoft.Web.Administration.ServerManager' -as [type])) {
        $dll = Join-Path $env:windir 'System32\inetsrv\Microsoft.Web.Administration.dll'
        if (-not (Test-Path -LiteralPath $dll)) {
            throw "[STD-SRV-IIS] Brak $dll - IIS nie jest zainstalowany. Uruchom instalkę bez -SkipFeatureInstall."
        }
        Add-Type -Path $dll
    }
}

function New-StdSrvServerManager {
    return (New-Object Microsoft.Web.Administration.ServerManager)
}

function Set-StdSrvCollection {
    # Czyści kolekcję na tym poziomie (w pliku: <clear />, czyli bez reguł dziedziczonych) i dodaje podane elementy.
    param($Collection, [string]$ElementName, [object[]]$Items)
    $Collection.Clear()
    foreach ($item in @($Items)) {
        if ($null -eq $item) { continue }
        $element = $Collection.CreateElement($ElementName)
        foreach ($key in @($item.Keys)) { $element[$key] = $item[$key] }
        [void]$Collection.Add($element)
    }
}

function Set-StdSrvMimeMaps {
    # Lista typów MIME witryny od zera (<clear /> + pełna lista STD). Witryna serwuje tylko pliki STD, a dopisywanie
    # do listy dziedziczonej z poziomu serwera kończyło się błędem "Cannot add duplicate collection entry" (np. '.zip').
    # Za każdym uruchomieniem wynik jest ten sam, więc aktualizacja jest bezpieczna.
    param($StaticContentSection, [System.Collections.IDictionary]$Types)
    $collection = $StaticContentSection.GetCollection()
    $collection.Clear()
    foreach ($ext in @($Types.Keys)) {
        $map = $collection.CreateElement('mimeMap')
        $map['fileExtension'] = $ext
        $map['mimeType'] = $Types[$ext]
        [void]$collection.Add($map)
    }
}

function Get-StdSrvOrNewSite {
    param($ServerManager, [string]$Name, [string]$Protocol, [string]$BindingInformation, [string]$PhysicalPath)
    $site = $ServerManager.Sites[$Name]
    if ($null -eq $site) {
        $site = $ServerManager.Sites.Add($Name, $Protocol, $BindingInformation, $PhysicalPath)
        if ($site.Id -le 0) {
            $max = 0
            foreach ($other in $ServerManager.Sites) { if ($other.Id -gt $max) { $max = $other.Id } }
            $site.Id = $max + 1
        }
        Write-StdSrvInfo "Utworzono witrynę $Name."
    }
    return $site
}

function Get-StdSrvBindingFlags {
    param($Binding)
    try { return [int]$Binding.GetAttributeValue('sslFlags') } catch { return 0 }
}

function Set-StdSrvHttpsBinding {
    <#
        Dokładnie jedno wiązanie witryny: https *:<port>:<host> z SNI (sslFlags 1). Zwraca $true, gdy coś zmieniono.
        Certyfikatu NIE ustawiamy tutaj: wiązanie certyfikatu żyje w http.sys (nie w applicationHost.config)
        i ustawia je zawsze Confirm-StdSrvSslBinding (netsh http add/update sslcert hostnameport=...).
        Jedno, przewidywalne miejsce - także przy wymianie certyfikatu i przy ponownym uruchomieniu po błędzie.
    #>
    param($Site, [string]$BindingInformation)
    $current = @($Site.Bindings)
    if ($current.Count -eq 1) {
        $b = $current[0]
        if ($b.Protocol -eq 'https' -and $b.BindingInformation -eq $BindingInformation -and ((Get-StdSrvBindingFlags -Binding $b) -band 1) -eq 1) {
            return $false
        }
    }
    foreach ($b in $current) { $Site.Bindings.Remove($b) }
    $added = $Site.Bindings.Add($BindingInformation, 'https')
    $added.SetAttributeValue('sslFlags', 1)
    return $true
}

function Set-StdSrvHttpBinding {
    param($Site, [string]$BindingInformation)
    $current = @($Site.Bindings)
    if ($current.Count -eq 1 -and $current[0].Protocol -eq 'http' -and $current[0].BindingInformation -eq $BindingInformation) { return $false }
    foreach ($b in $current) { $Site.Bindings.Remove($b) }
    [void]$Site.Bindings.Add($BindingInformation, 'http')
    return $true
}

function Set-StdSrvRootPath {
    # Aplikacja '/' witryny: pula i katalog główny. Zwraca aplikację.
    param($Site, [string]$PoolName, [string]$SiteRoot)
    $app = $Site.Applications['/']
    if ($null -eq $app) { $app = $Site.Applications.Add('/', $SiteRoot) }
    $app.ApplicationPoolName = $PoolName
    $root = $app.VirtualDirectories['/']
    if ($null -eq $root) { $root = $app.VirtualDirectories.Add('/', $SiteRoot) }
    $root.PhysicalPath = $SiteRoot
    return $app
}

function Get-StdSrvSitesWithBinding {
    # Nazwy innych witryn, które mają takie samo wiązanie (konflikt).
    param($ServerManager, [string]$Protocol, [string]$BindingInformation, [string[]]$Exclude)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($site in $ServerManager.Sites) {
        if ($Exclude -contains $site.Name) { continue }
        foreach ($b in $site.Bindings) {
            if ($b.Protocol -eq $Protocol -and $b.BindingInformation -eq $BindingInformation) { $names.Add($site.Name); break }
        }
    }
    return , $names.ToArray()
}

function Set-StdSrvIisSections {
    param($Config, [System.Collections.IDictionary]$Settings, [string]$RedirectSiteName)
    $site = $Settings.SiteName

    # Logowanie: tylko Basic (anonimowe wyłączone), tylko HTTPS. Ustawione dla całej witryny.
    $anon = $Config.GetSection('system.webServer/security/authentication/anonymousAuthentication', $site)
    $anon['enabled'] = $false
    $basic = $Config.GetSection('system.webServer/security/authentication/basicAuthentication', $site)
    $basic['enabled'] = $true
    $basic['realm'] = 'STD'
    # Login bez domeny = konto lokalne tego serwera (konto domenowe: DOMENA\login).
    $basic['defaultLogonDomain'] = $env:COMPUTERNAME
    $access = $Config.GetSection('system.webServer/security/access', $site)
    $access['sslFlags'] = 'Ssl'

    $browse = $Config.GetSection('system.webServer/directoryBrowse', $site)
    $browse['enabled'] = $false

    # '+' w nazwach plików (np. Notepad++.exe) - bez tego IIS odrzuca adres (404.11). Witryna serwuje tylko pliki statyczne.
    $filtering = $Config.GetSection('system.webServer/security/requestFiltering', $site)
    $filtering['allowDoubleEscaping'] = $true

    $static = $Config.GetSection('system.webServer/staticContent', $site)
    Set-StdSrvMimeMaps -StaticContentSection $static -Types $script:StdSrvMimeTypes
    # Cache-Control: no-cache dla wszystkich odpowiedzi (obejmuje .sdt z SPEC 5.1).
    $clientCache = $static.GetChildElement('clientCache')
    $clientCache['cacheControlMode'] = 'DisableCache'

    $protocol = $Config.GetSection('system.webServer/httpProtocol', $site)
    Set-StdSrvCollection -Collection $protocol.GetCollection('customHeaders') -ElementName 'add' -Items @(
        @{ name = 'Strict-Transport-Security'; value = 'max-age=31536000' },
        @{ name = 'X-Content-Type-Options'; value = 'nosniff' }
    )

    # WebDAV włączony dla witryny; kto i gdzie może zapisywać, decydują reguły niżej (i autoryzacja URL).
    $authoring = $Config.GetSection('system.webServer/webdav/authoring', $site)
    $authoring['enabled'] = $true
    $authoring['requireSsl'] = $true

    foreach ($loc in (Get-StdSrvLocationPlan -SiteName $site -TechnicianGroup $Settings.TechnicianGroup -AdminGroup $Settings.AdminGroup)) {
        $authorization = $Config.GetSection('system.webServer/security/authorization', $loc.Location)
        $authItems = @($loc.Authorization | ForEach-Object { @{ accessType = 'Allow'; roles = $_.Roles; verbs = $_.Verbs } })
        Set-StdSrvCollection -Collection $authorization.GetCollection() -ElementName 'add' -Items $authItems

        $rules = $Config.GetSection('system.webServer/webdav/authoringRules', $loc.Location)
        $ruleItems = @($loc.AuthoringRules | ForEach-Object { @{ roles = $_.Roles; path = '*'; access = $_.Access } })
        Set-StdSrvCollection -Collection $rules.GetCollection() -ElementName 'add' -Items $ruleItems

        if ($null -ne $loc.MaxUpload) {
            $locFiltering = $Config.GetSection('system.webServer/security/requestFiltering', $loc.Location)
            $limits = $locFiltering.GetChildElement('requestLimits')
            $limits['maxAllowedContentLength'] = [uint32]$loc.MaxUpload
        }
    }

    if ($RedirectSiteName) {
        # Osobna witryna na porcie 80 tylko przekierowuje (301) na HTTPS z zachowaniem ścieżki.
        $redirect = $Config.GetSection('system.webServer/httpRedirect', $RedirectSiteName)
        $redirect['enabled'] = $true
        $redirect['destination'] = (Get-StdSrvBaseUrl -HostName $Settings.HostName -Port $Settings.Port).TrimEnd('/')
        $redirect['exactDestination'] = $false
        $redirect['childOnly'] = $false
        $redirect['httpResponseStatus'] = 'Permanent'
        $redirectAnon = $Config.GetSection('system.webServer/security/authentication/anonymousAuthentication', $RedirectSiteName)
        $redirectAnon['enabled'] = $true
        # Konto anonimowe = tożsamość puli (ma odczyt pustego katalogu witryny), nie IUSR.
        $redirectAnon['userName'] = ''
        $redirectBasic = $Config.GetSection('system.webServer/security/authentication/basicAuthentication', $RedirectSiteName)
        $redirectBasic['enabled'] = $false
        $redirectBrowse = $Config.GetSection('system.webServer/directoryBrowse', $RedirectSiteName)
        $redirectBrowse['enabled'] = $false
    }
}

function Set-StdSrvIisConfiguration {
    # Pula, witryna HTTPS, katalogi wirtualne std i std-admin, witryna przekierowania i wszystkie sekcje. Jeden zapis.
    param($ServerManager, [System.Collections.IDictionary]$Settings, [string]$SiteRoot, [string]$StdPath,
        [string]$RedirectSiteName)
    $sm = $ServerManager
    $name = $Settings.SiteName

    $pool = $sm.ApplicationPools[$name]
    if ($null -eq $pool) {
        $pool = $sm.ApplicationPools.Add($name)
        Write-StdSrvInfo "Utworzono pulę aplikacji $name."
    }
    $pool.ManagedRuntimeVersion = ''
    $pool.ManagedPipelineMode = 'Integrated'
    $pool.ProcessModel.IdentityType = 'ApplicationPoolIdentity'
    $pool.AutoStart = $true

    $httpsInfo = '*:{0}:{1}' -f $Settings.Port, $Settings.HostName
    # Nowa witryna powstaje z tymczasowym wiązaniem http (tylko w pamięci, przed zapisem), które
    # Set-StdSrvHttpsBinding od razu podmienia na https z certyfikatem.
    $site = Get-StdSrvOrNewSite -ServerManager $sm -Name $name -Protocol 'http' -BindingInformation $httpsInfo -PhysicalPath $SiteRoot
    $site.ServerAutoStart = $true
    if (Set-StdSrvHttpsBinding -Site $site -BindingInformation $httpsInfo) {
        Write-StdSrvInfo "Wiązanie HTTPS: $httpsInfo (SNI)."
    }
    $app = Set-StdSrvRootPath -Site $site -PoolName $name -SiteRoot $SiteRoot
    foreach ($vpath in '/std', '/std-admin') {
        $vdir = $app.VirtualDirectories[$vpath]
        if ($null -eq $vdir) {
            [void]$app.VirtualDirectories.Add($vpath, $StdPath)
            Write-StdSrvInfo "Utworzono katalog wirtualny $vpath -> $StdPath"
        }
        else {
            $vdir.PhysicalPath = $StdPath
        }
    }

    if ($RedirectSiteName) {
        $httpInfo = '*:80:{0}' -f $Settings.HostName
        $redirectSite = Get-StdSrvOrNewSite -ServerManager $sm -Name $RedirectSiteName -Protocol 'http' -BindingInformation $httpInfo -PhysicalPath $SiteRoot
        $redirectSite.ServerAutoStart = $true
        [void](Set-StdSrvHttpBinding -Site $redirectSite -BindingInformation $httpInfo)
        [void](Set-StdSrvRootPath -Site $redirectSite -PoolName $name -SiteRoot $SiteRoot)
    }

    Set-StdSrvIisSections -Config $sm.GetApplicationHostConfiguration() -Settings $Settings -RedirectSiteName $RedirectSiteName
    $sm.CommitChanges()
    Write-StdSrvInfo 'Zapisano konfigurację IIS.'
}

function Confirm-StdSrvSslBinding {
    # Sprawdza wiązanie certyfikatu w http.sys (hostname:port, SNI); gdy go nie ma albo jest inny certyfikat - ustawia przez netsh.
    param([string]$HostName, [int]$Port, [string]$Thumbprint)
    $hp = '{0}:{1}' -f $HostName, $Port
    $show = Invoke-StdSrvNative -FilePath 'netsh.exe' -Arguments @('http', 'show', 'sslcert', "hostnameport=$hp")
    if ($show.ExitCode -eq 0 -and $show.Output.ToUpperInvariant().Contains($Thumbprint)) {
        Write-StdSrvInfo "http.sys: certyfikat dla $hp jest ustawiony."
        return
    }
    $add = @('http', 'add', 'sslcert', "hostnameport=$hp", "certhash=$Thumbprint", "appid=$script:StdSrvIisAppId", 'certstorename=MY')
    if ($show.ExitCode -eq 0) {
        # Jest wiązanie z innym certyfikatem (wymiana certyfikatu albo pozostałość): update, a gdy się nie uda - delete + add.
        $update = @('http', 'update', 'sslcert') + $add[3..6]
        $r = Invoke-StdSrvNative -FilePath 'netsh.exe' -Arguments $update
        if ($r.ExitCode -eq 0) {
            Write-StdSrvInfo "http.sys: zmieniono certyfikat dla $hp."
            return
        }
        [void](Invoke-StdSrvNative -FilePath 'netsh.exe' -Arguments @('http', 'delete', 'sslcert', "hostnameport=$hp"))
    }
    $r = Invoke-StdSrvNative -FilePath 'netsh.exe' -Arguments $add
    if ($r.ExitCode -ne 0) {
        throw "[STD-SRV-TLS] Nie udało się powiązać certyfikatu z $hp w http.sys: $($r.Output)"
    }
    Write-StdSrvInfo "http.sys: ustawiono certyfikat dla $hp."
}

function Remove-StdSrvSslBinding {
    # Usuwa stare wiązanie certyfikatu po zmianie nazwy albo portu (błędy pomijamy - mogło go już nie być).
    param([string]$HostName, [int]$Port)
    $hp = '{0}:{1}' -f $HostName, $Port
    $r = Invoke-StdSrvNative -FilePath 'netsh.exe' -Arguments @('http', 'delete', 'sslcert', "hostnameport=$hp")
    if ($r.ExitCode -eq 0) { Write-StdSrvInfo "http.sys: usunięto stare wiązanie $hp." }
}

function Start-StdSrvSites {
    # Uruchamia pulę i witryny (po zapisie konfiguracji usługa WAS potrzebuje chwili).
    param([string]$PoolName, [string[]]$SiteNames)
    $lastError = $null
    for ($i = 0; $i -lt 15; $i++) {
        $sm = New-StdSrvServerManager
        try {
            $pool = $sm.ApplicationPools[$PoolName]
            if ($null -ne $pool -and "$($pool.State)" -ne 'Started') { [void]$pool.Start() }
            $allStarted = $true
            foreach ($n in $SiteNames) {
                $site = $sm.Sites[$n]
                if ($null -eq $site) { continue }
                if ("$($site.State)" -ne 'Started') { [void]$site.Start() }
                if ("$($site.State)" -ne 'Started') { $allStarted = $false }
            }
            if ($allStarted) { return $true }
        }
        catch {
            $lastError = $_.Exception.Message
        }
        finally {
            $sm.Dispose()
        }
        Start-Sleep -Seconds 2
    }
    Write-StdSrvWarn "Nie wszystkie witryny wystartowały: $lastError"
    return $false
}

function Invoke-StdSrvProbe {
    <#
        Zapytanie HTTP(S) do 127.0.0.1 z podaną nazwą hosta (SNI i nagłówek Host), bez zależności od DNS.
        Zwraca @{ StatusCode; Location; Thumbprint }. Certyfikat nie jest tu weryfikowany - sprawdzamy odcisk.
    #>
    param([string]$HostName, [int]$Port, [string]$Path, [switch]$Tls, [int]$TimeoutMs = 10000)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect('127.0.0.1', $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs)) { throw "brak połączenia z 127.0.0.1:$Port" }
        $client.EndConnect($iar)
        $client.ReceiveTimeout = $TimeoutMs
        $client.SendTimeout = $TimeoutMs
        $stream = $client.GetStream()
        $thumb = $null
        if ($Tls) {
            $callback = [System.Net.Security.RemoteCertificateValidationCallback] { param($s, $c, $ch, $e) return $true }
            $ssl = New-Object System.Net.Security.SslStream($stream, $false, $callback)
            $ssl.AuthenticateAsClient($HostName, $null, [System.Security.Authentication.SslProtocols]::Tls12, $false)
            if ($null -ne $ssl.RemoteCertificate) {
                $thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)).Thumbprint.ToUpperInvariant()
            }
            $stream = $ssl
        }
        $hostHeader = $HostName
        if (($Tls -and $Port -ne 443) -or (-not $Tls -and $Port -ne 80)) { $hostHeader = "${HostName}:$Port" }
        $request = "GET $Path HTTP/1.1`r`nHost: $hostHeader`r`nUser-Agent: STD-Install`r`nConnection: close`r`n`r`n"
        $bytes = [Text.Encoding]::ASCII.GetBytes($request)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush()
        $reader = New-Object System.IO.StreamReader($stream, [Text.Encoding]::ASCII)
        $status = $reader.ReadLine()
        if ($status -notmatch '^HTTP/\d\.\d (\d{3})') { throw "nieoczekiwana odpowiedź: $status" }
        $code = [int]$Matches[1]
        $location = $null
        while ($true) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrEmpty($line)) { break }
            if ($line -match '^Location:\s*(.+)$') { $location = $Matches[1].Trim() }
        }
        return @{ StatusCode = $code; Location = $location; Thumbprint = $thumb }
    }
    finally {
        $client.Close()
    }
}

function Test-StdSrvServer {
    # Test po zmianie: HTTPS z właściwym certyfikatem i 401 dla /std/ bez logowania.
    param([string]$HostName, [int]$Port, [string]$Thumbprint, [int]$Attempts = 15)
    $last = ''
    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $r = Invoke-StdSrvProbe -HostName $HostName -Port $Port -Path '/std/' -Tls
            if ($r.Thumbprint -ne $Thumbprint) { $last = "serwer przedstawił inny certyfikat ($($r.Thumbprint))" }
            elseif ($r.StatusCode -eq 401) { return @{ Ok = $true; Message = 'HTTPS działa, /std/ wymaga logowania (401).' } }
            else { $last = "odpowiedź $($r.StatusCode) zamiast 401 dla /std/ bez logowania" }
        }
        catch {
            $last = $_.Exception.Message
        }
        Start-Sleep -Seconds 2
    }
    return @{ Ok = $false; Message = $last }
}

# ---------------------------------------------------------------------------------------------
# Instalacja / aktualizacja
# ---------------------------------------------------------------------------------------------

function Get-StdSrvDefaultHostName {
    $name = $null
    try { $name = [System.Net.Dns]::GetHostEntry([string]$env:COMPUTERNAME).HostName } catch { $name = $null }
    if ([string]::IsNullOrEmpty($name)) { $name = $env:COMPUTERNAME }
    return $name.ToLowerInvariant()
}

function Get-StdSrvToolVersion {
    # Wersja z manifestu modułu STD (v4\STD.psd1), nigdzie nie wpisana drugi raz.
    param([string]$ScriptRoot)
    $manifest = Join-Path $ScriptRoot '..\..\STD.psd1'
    if (Test-Path -LiteralPath $manifest) {
        try { return [string](Import-PowerShellDataFile -Path $manifest).ModuleVersion } catch { }
    }
    return 'nieznana'
}

function Get-StdSrvSiteRoot {
    # Pusty katalog główny witryny (poza danymi, żeby nikt nie podłożył web.config dla całej witryny).
    param([string]$SiteName)
    return (Join-Path $env:SystemDrive "inetpub\$SiteName-wwwroot")
}

function Write-StdSrvSummary {
    param([System.Collections.IDictionary]$Settings, [string]$StdPath, [string]$BootstrapPath, [bool]$Redirect)
    $base = Get-StdSrvBaseUrl -HostName $Settings.HostName -Port $Settings.Port
    Write-Host ''
    Write-Host 'Serwer STD jest gotowy.' -ForegroundColor Green
    Write-Host ''
    Write-Host "  Adres dla techników:       ${base}std/"
    Write-Host "  Publikacja (administrator): ${base}std-admin/"
    Write-Host "  Dane:                      $StdPath"
    if ($Redirect) { Write-Host "  HTTP (port 80):            przekierowanie 301 na HTTPS" }
    Write-Host ''
    Write-Host '  Role:'
    Write-Host "    technik        = grupa '$($Settings.TechnicianGroup)': odczyt /std/, zapis raportów do /std/drop/ (bez odczytu i listowania)"
    Write-Host "    administrator  = grupa '$($Settings.AdminGroup)': odczyt /std/, zapis do /std/drop/ i pełny dostęp do /std-admin/"
    Write-Host ''
    Write-Host '  Konta:'
    Write-Host '    .\Add-StdServerUser.ps1 -UserName jan.kowalski -Role Technician -Create'
    Write-Host '    .\Add-StdServerUser.ps1 -UserName anna.nowak -Role Admin -Create'
    Write-Host ''
    Write-Host "  Plik bootstrap.json na pendrive (obok Start-STD.cmd), kopia: $BootstrapPath"
    Write-Host ('    ' + (Get-StdSrvBootstrapJson -HostName $Settings.HostName -Port $Settings.Port))
    Write-Host ''
}

function Invoke-StdSrvInstall {
    param([System.Collections.IDictionary]$Bound, [string]$ScriptRoot)

    $stateRoot = Get-StdSrvStateRoot
    $settingsPath = Join-Path $stateRoot 'settings.json'

    Write-StdSrvStep 'Ustawienia'
    $saved = Read-StdSrvSettingsFile -Path $settingsPath
    if ($null -ne $saved) { Write-StdSrvInfo "Wczytano $settingsPath (aktualizacja)." } else { Write-StdSrvInfo 'Pierwsza instalacja.' }
    $defaults = [ordered]@{}
    foreach ($k in $script:StdSrvDefaults.Keys) { $defaults[$k] = $script:StdSrvDefaults[$k] }
    $defaults['HostName'] = Get-StdSrvDefaultHostName
    $resolved = Resolve-StdSrvSettings -Bound $Bound -Saved $saved -Defaults $defaults
    $s = $resolved.Settings
    $s['HostName'] = $s.HostName.Trim().TrimEnd('.').ToLowerInvariant()
    if ($s.DataPath.Length -gt 3) { $s['DataPath'] = $s.DataPath.TrimEnd('\') }
    if ($s.CertificateThumbprint) { $s['CertificateThumbprint'] = ConvertTo-StdSrvThumbprint -Value $s.CertificateThumbprint }
    Assert-StdSrvSettings -Settings $s
    foreach ($k in $script:StdSrvSettingKeys) {
        $shown = $s[$k]
        if ($null -eq $shown) { $shown = '(wybór automatyczny)' }
        Write-StdSrvInfo ('{0,-22} {1}  [{2}]' -f $k, $shown, $resolved.Sources[$k])
    }
    $previousSite = [string](Get-StdSrvSavedValue -Saved $saved -Key 'SiteName')
    $previousData = [string](Get-StdSrvSavedValue -Saved $saved -Key 'DataPath')

    Write-StdSrvStep 'Funkcje Windows (IIS)'
    if ((Test-StdSrvDictionaryKey -Dictionary $Bound -Key 'SkipFeatureInstall') -and $Bound['SkipFeatureInstall']) {
        Write-StdSrvInfo 'Pominięto (-SkipFeatureInstall).'
    }
    else {
        Install-StdSrvFeatures
    }
    Import-StdSrvIisApi

    Write-StdSrvStep 'Certyfikat'
    $cert = Resolve-StdSrvCertificate -Thumbprint $s.CertificateThumbprint -HostName $s.HostName
    $s['CertificateThumbprint'] = $cert.Thumbprint.ToUpperInvariant()
    Write-StdSrvInfo "Certyfikat: $($cert.Subject), ważny do $($cert.NotAfter.ToString('yyyy-MM-dd'))."

    $stdPath = Join-Path $s.DataPath 'std'
    $siteRoot = Get-StdSrvSiteRoot -SiteName $s.SiteName
    $redirectName = "$($s.SiteName)-HTTP"
    $httpsInfo = '*:{0}:{1}' -f $s.Port, $s.HostName

    Write-StdSrvStep 'Sprawdzenie IIS'
    $sm = New-StdSrvServerManager
    try {
        if ($previousSite -and $previousSite -ne $s.SiteName -and $null -ne $sm.Sites[$previousSite]) {
            throw "[STD-SRV-SITE] Zmiana nazwy witryny ($previousSite -> $($s.SiteName)) nie jest obsługiwana. Uruchom instalkę bez -SiteName albo usuń najpierw starą witrynę w Menedżerze IIS."
        }
        $existing = $sm.Sites[$s.SiteName]
        if ($null -ne $existing -and $null -eq $saved) {
            $existingRoot = ''
            try { $existingRoot = [Environment]::ExpandEnvironmentVariables($existing.Applications['/'].VirtualDirectories['/'].PhysicalPath).TrimEnd('\') } catch { }
            if ($existingRoot -ne $siteRoot) {
                throw "[STD-SRV-SITE] Witryna '$($s.SiteName)' już istnieje i nie jest witryną STD (katalog: $existingRoot). Podaj inną nazwę: -SiteName."
            }
        }
        $httpsConflicts = Get-StdSrvSitesWithBinding -ServerManager $sm -Protocol 'https' -BindingInformation $httpsInfo -Exclude @($s.SiteName)
        if ($httpsConflicts.Count -gt 0) {
            throw "[STD-SRV-SITE] Wiązanie $httpsInfo ma już witryna: $($httpsConflicts -join ', '). Zmień -HostName albo -Port."
        }
        $httpConflicts = Get-StdSrvSitesWithBinding -ServerManager $sm -Protocol 'http' -BindingInformation ('*:80:{0}' -f $s.HostName) -Exclude @($s.SiteName, $redirectName)
    }
    finally {
        $sm.Dispose()
    }
    $useRedirect = $true
    if ($httpConflicts.Count -gt 0) {
        $useRedirect = $false
        Write-StdSrvWarn "Port 80 dla $($s.HostName) zajmuje witryna: $($httpConflicts -join ', '). Pomijam przekierowanie HTTP -> HTTPS."
    }
    Write-StdSrvInfo 'OK'

    Write-StdSrvStep 'Kopia zapasowa'
    [void](New-Item -ItemType Directory -Path $stateRoot -Force)
    $backup = New-StdSrvBackup -StateRoot $stateRoot -StdPath $stdPath -SettingsPath $settingsPath
    $script:StdSrvLastBackup = $backup.Path

    Write-StdSrvStep 'Katalogi i grupy'
    Initialize-StdSrvFolders -StdPath $stdPath -SiteRoot $siteRoot -StateRoot $stateRoot
    if ($previousData -and $previousData.TrimEnd('\') -ne $s.DataPath) {
        Write-StdSrvWarn "Zmieniono katalog danych z $previousData na $($s.DataPath). Stare dane zostały na miejscu - przenieś je ręcznie (np. robocopy /E /COPY:DAT), uprawnienia ustawi ponowne uruchomienie instalki."
    }
    # Opis grupy: najwyżej 48 znaków (limit New-LocalGroup).
    $techSid = Initialize-StdSrvGroup -Name $s.TechnicianGroup -Description 'STD: technicy (odczyt /std/, zapis /std/drop/)'
    $adminSid = Initialize-StdSrvGroup -Name $s.AdminGroup -Description 'STD: administratorzy (pełny dostęp, /std-admin/)'
    $removeSids = New-Object System.Collections.Generic.List[string]
    foreach ($key in 'TechnicianGroup', 'AdminGroup') {
        $old = [string](Get-StdSrvSavedValue -Saved $saved -Key $key)
        if ($old -and $old -ne $s.TechnicianGroup -and $old -ne $s.AdminGroup) {
            $oldSid = Get-StdSrvSid -Name $old
            if ($oldSid) { $removeSids.Add($oldSid); Write-StdSrvInfo "Grupa $old nie ma już dostępu do danych (grupa zostaje w systemie)." }
        }
    }

    Write-StdSrvStep 'Uprawnienia NTFS'
    $aclPlan = Get-StdSrvAclPlan -StdPath $stdPath -SiteRoot $siteRoot -StateRoot $stateRoot -TechnicianSid $techSid -AdminSid $adminSid
    Set-StdSrvAcl -Plan $aclPlan -RemoveSids $removeSids.ToArray() -StdPath $stdPath

    Write-StdSrvStep 'Konfiguracja IIS'
    $redirectArg = $null
    if ($useRedirect) { $redirectArg = $redirectName }
    try {
        $sm = New-StdSrvServerManager
        try {
            Set-StdSrvIisConfiguration -ServerManager $sm -Settings $s -SiteRoot $siteRoot -StdPath $stdPath -RedirectSiteName $redirectArg
        }
        finally {
            $sm.Dispose()
        }
        Confirm-StdSrvSslBinding -HostName $s.HostName -Port $s.Port -Thumbprint $s.CertificateThumbprint
        $oldHost = [string](Get-StdSrvSavedValue -Saved $saved -Key 'HostName')
        $oldPort = Get-StdSrvSavedValue -Saved $saved -Key 'Port'
        if ($oldHost -and $null -ne $oldPort -and ($oldHost.ToLowerInvariant() -ne $s.HostName -or [int]$oldPort -ne $s.Port)) {
            Remove-StdSrvSslBinding -HostName $oldHost.ToLowerInvariant() -Port ([int]$oldPort)
        }
        $siteNames = @($s.SiteName)
        if ($useRedirect) { $siteNames += $redirectName }
        [void](Start-StdSrvSites -PoolName $s.SiteName -SiteNames $siteNames)

        Write-StdSrvStep 'Test serwera'
        $test = Test-StdSrvServer -HostName $s.HostName -Port $s.Port -Thumbprint $s.CertificateThumbprint
        if (-not $test.Ok) {
            throw "[STD-SRV-TEST] Serwer nie działa poprawnie po zmianie konfiguracji: $($test.Message)"
        }
        Write-StdSrvInfo $test.Message
    }
    catch {
        Restore-StdSrvBackup -Backup $backup
        throw
    }

    if ($useRedirect) {
        try {
            $r = Invoke-StdSrvProbe -HostName $s.HostName -Port 80 -Path '/std/'
            if ($r.StatusCode -eq 301 -and "$($r.Location)".StartsWith('https://')) { Write-StdSrvInfo "HTTP: 301 -> $($r.Location)" }
            else { Write-StdSrvWarn "HTTP na porcie 80 odpowiada $($r.StatusCode) zamiast 301." }
        }
        catch {
            Write-StdSrvWarn "Przekierowanie HTTP nie odpowiada: $($_.Exception.Message)"
        }
    }

    Write-StdSrvStep 'Zapis ustawień'
    $doc = New-StdSrvSettingsDocument -Settings $s -Saved $saved -Version (Get-StdSrvToolVersion -ScriptRoot $ScriptRoot) -Now (Get-Date)
    Write-StdSrvSettingsFile -Path $settingsPath -Settings $doc
    Write-StdSrvInfo $settingsPath
    $bootstrapPath = Join-Path $stateRoot 'bootstrap.json'
    [IO.File]::WriteAllText($bootstrapPath, (Get-StdSrvBootstrapJson -HostName $s.HostName -Port $s.Port), (New-Object System.Text.UTF8Encoding($false)))

    Write-StdSrvSummary -Settings $s -StdPath $stdPath -BootstrapPath $bootstrapPath -Redirect $useRedirect
}

# ---------------------------------------------------------------------------------------------
# Start (pomijany przy dot-source - testy ładują same funkcje)
# ---------------------------------------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    $stdExitCode = 0
    $stdTranscript = $false
    try {
        if ($env:OS -ne 'Windows_NT') {
            throw '[STD-SRV-OS] Ta instalka działa tylko w Windows (Windows Server z IIS). Na Linuksie użyj Server/Linux/install-std-server.sh.'
        }
        $stdIsAdmin = Test-StdSrvIsAdmin
        if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
            Write-Host 'Uruchamiam instalkę ponownie w 64-bitowym Windows PowerShell (konfiguracja IIS jest 64-bitowa)...'
            exit (Invoke-StdSrvRelaunch -ScriptPath $PSCommandPath -Bound $PSBoundParameters -Elevate:(-not $stdIsAdmin))
        }
        if (-not $stdIsAdmin) {
            Write-Host 'Instalka wymaga uprawnień administratora - uruchamiam ją ponownie jako administrator...'
            exit (Invoke-StdSrvRelaunch -ScriptPath $PSCommandPath -Bound $PSBoundParameters -Elevate)
        }
        $ErrorActionPreference = 'Stop'
        try {
            $stdLogDir = Join-Path (Get-StdSrvStateRoot) 'logs'
            [void](New-Item -ItemType Directory -Path $stdLogDir -Force)
            $stdLogPath = Join-Path $stdLogDir ('install-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
            [void](Start-Transcript -Path $stdLogPath)
            $stdTranscript = $true
        }
        catch { }
        Write-Host 'STD (Smart Tool for Deployment) - instalka serwera IIS'
        Invoke-StdSrvInstall -Bound $PSBoundParameters -ScriptRoot $PSScriptRoot
    }
    catch {
        $stdExitCode = 1
        Write-Host ''
        Write-Host ('BŁĄD: ' + $_.Exception.Message) -ForegroundColor Red
        if ($_.Exception.Message -notmatch '^\[STD-') {
            Write-Host $_.InvocationInfo.PositionMessage
            Write-Host $_.ScriptStackTrace
        }
        if ($script:StdSrvLastBackup) { Write-Host "Kopia zapasowa sprzed zmian: $script:StdSrvLastBackup" }
        Write-Host 'Pomoc: Server\IIS\README.md, rozdział "Problemy".'
    }
    finally {
        if ($stdTranscript) { try { [void](Stop-Transcript) } catch { } }
    }
    if ($ElevatedRelaunch) { [void](Read-Host 'Naciśnij Enter, aby zamknąć to okno') }
    exit $stdExitCode
}
