# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.Config.ps1 - konfiguracja STD (SPEC 1, 2.3, 2.5-2.8 i 3.4): zaufane certyfikaty, plik startowy
# bootstrap.json, najwyższe widziane rewizje (ochrona przed starszą wersją), pobranie podpisanego indeksu
# klientów i konfiguracji klienta (serwer -> kopie lustrzane -> cache -> pendrive), walidacja konfiguracji,
# łączenie ustawień (defaults -> lokalizacja -> profil) i konfiguracja wynikowa dla wdrożenia.
#
# Wymaga Std.Core.ps1, Std.Http.ps1 i Std.Crypto.ps1 - STD.psm1 wczytuje je wcześniej.
#
# Zasady:
# - STD na komputerze konfiguracji nigdy nie zmienia: tylko ją pobiera, sprawdza podpis i reguły, zapisuje
#   do cache dokładnie te bajty, które przyszły (podpisana koperta), i sprawdza ją ponownie przy każdym odczycie;
# - login i hasło do domeny oraz hasło lokalnego administratora nie mają pól w konfiguracji - Test-StdRelease
#   odrzuca dokument, który je ma (SPEC 0.1). Nic w tym pliku nie zapisuje haseł ani tokenów;
# - plik bootstrap.json nie jest sekretem ani źródłem zaufania: zaufanie daje tylko podpis.
#
# Funkcje pomocnicze bez "-Std" w nazwie nie są eksportowane z modułu (Export-ModuleMember *-Std*).

# ---------------------------------------------------------------------------------------------
# Stałe
# ---------------------------------------------------------------------------------------------

$script:StdConfigFile = $PSCommandPath

$script:StdConfigIdPattern = '^[a-z0-9][a-z0-9-]{1,39}$'
$script:StdConfigCatalogKeyPattern = '^[A-Za-z0-9][A-Za-z0-9._+-]{0,63}$'
$script:StdConfigSecretNamePattern = '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$'
$script:StdConfigIsoDatePattern = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:\d{2})$'
$script:StdConfigSha256Pattern = '^[0-9A-Fa-f]{64}$'

# Klucz rewizji indeksu klientów w revisions.json - dlatego "clients" nie może być identyfikatorem klienta.
$script:StdConfigIndexRevisionKey = 'clients'

$script:StdConfigTaskIds = @(
    'WaitForNetwork', 'SuspendSleep', 'ConnectWiFi', 'RemoveBloatware', 'UninstallPreinstalled', 'UninstallM365',
    'UninstallOneDrive', 'InstallApps', 'CreateLocalAdmin', 'RenameComputer', 'JoinDomain', 'SystemSettings',
    'RunScripts', 'HardwareAudit', 'EnableBitLocker', 'RebootAtEnd'
)
$script:StdConfigLockableExtra = @('computerName', 'apps')
$script:StdConfigSourceTypes = @('https', 'unc')
$script:StdConfigSourceAuth = @('session', 'none')
$script:StdConfigWhenDetected = @('skip', 'reinstall')
$script:StdConfigRegistryHives = @('HKLM', 'HKCU', 'DefaultUser')
$script:StdConfigRegistryTypes = @('String', 'ExpandString', 'DWord', 'QWord', 'MultiString', 'Binary')
$script:StdConfigWifiAuth = @('WPA2PSK', 'WPA3SAE', 'open')
$script:StdConfigBitLockerEscrow = @('ad', 'server')
$script:StdConfigNamePlaceholders = @('serial', 'site', 'random')
$script:StdConfigMaxComputerName = 15

# Ustawienia dziedziczone: defaults -> lokalizacja -> profil (SPEC 2.6). scripts są doklejane, reszta łączona.
$script:StdConfigOverrideKeys = @('tasks', 'bloatware', 'uninstall', 'systemSettings', 'scripts')

# Klucze, które w całym dokumencie (poza secrets) nie mogą mieć niepustej wartości tekstowej.
$script:StdConfigPlaintextKeys = @('password', 'passwd', 'pwd', 'haslo', 'hasło', 'apitoken', 'token', 'secret')

# Pola zabronione w domainJoin i localAdmin (poza localAdmin.username). SPEC podaje "hasl"; dopisane są
# też polskie formy z "ł" i "ż" (hasło, użytkownik) oraz "konto", bo tego samego dotyczy zakaz.
$script:StdConfigCredentialFieldPattern = 'pass|pwd|has[lł]|secret|credential|user(name)?|login|account|konto|u[zż]ytkownik'
$script:StdConfigCredentialMessage = 'Dane logowania do domeny i hasło lokalnego administratora wpisuje się ręcznie przy wdrożeniu - nie mogą być w konfiguracji.'
# Nazwa sekretu, która wygląda na dane logowania do domeny albo hasło lokalnego administratora (SPEC 2.7).
$script:StdConfigCredentialSecretTarget = 'domain|domen|join|admin'
$script:StdConfigCredentialSecretKind = 'pass|pwd|has[lł]|credential|login|user|u[zż]ytkownik|konto|account'

$script:StdConfigExpiryGraceMinutes = 5
$script:StdConfigDefaultValidityDays = 90
$script:StdConfigMaxReportedErrors = 5

# ---------------------------------------------------------------------------------------------
# Pomocnicze (prywatne)
# ---------------------------------------------------------------------------------------------

function Test-ConfigIsDictionary {
    param([AllowNull()]$Value)
    return ($Value -is [System.Collections.IDictionary])
}

function Test-ConfigIsList {
    param([AllowNull()]$Value)
    return (($Value -is [System.Collections.IList]) -and -not ($Value -is [System.Collections.IDictionary]) -and -not ($Value -is [string]))
}

function Test-ConfigIsText {
    # Niepusty tekst (nie same spacje).
    param([AllowNull()]$Value)
    return (($Value -is [string]) -and -not [string]::IsNullOrWhiteSpace($Value))
}

function Test-ConfigHasValue {
    # Czy słownik ma klucz z wartością różną od $null.
    param([AllowNull()]$Dictionary, [Parameter(Mandatory = $true)][string]$Key)
    if (-not ($Dictionary -is [System.Collections.IDictionary])) { return $false }
    return ($Dictionary.Contains($Key) -and $null -ne $Dictionary[$Key])
}

function Join-ConfigPath {
    param([AllowEmptyString()][string]$Path, [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Key)
    if ([string]::IsNullOrEmpty($Path)) { return $Key }
    return ('{0}.{1}' -f $Path, $Key)
}

function Join-ConfigIndexPath {
    param([AllowEmptyString()][string]$Path, [Parameter(Mandatory = $true)][int]$Index)
    return ('{0}[{1}]' -f $Path, $Index)
}

function Add-ConfigError {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()]$Errors, [AllowEmptyString()][string]$Path, [Parameter(Mandatory = $true)][string]$Message)
    if ([string]::IsNullOrEmpty($Path)) { $Path = '(dokument)' }
    $Errors.Add(('{0}: {1}' -f $Path, $Message))
}

function Format-ConfigValue {
    # Krótki, bezpieczny zapis wartości do komunikatu (nigdy nie używać dla pól z hasłami i tokenami).
    param([AllowNull()]$Value, [int]$MaxLength = 60)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [System.Collections.IDictionary]) { return '{...}' }
    if (Test-ConfigIsList $Value) { return '[...]' }
    $text = [string]$Value
    $text = $text -replace '[\x00-\x1F\x7F]', '?'
    if ($text.Length -gt $MaxLength) { $text = $text.Substring(0, $MaxLength) + '...' }
    return ("'{0}'" -f $text)
}

function Test-ConfigIsInteger {
    # Liczba całkowita z JSON (Int32 / Int64 / Decimal / Double bez części ułamkowej); bool i tekst to nie liczby.
    param([AllowNull()]$Value)
    if ($null -eq $Value -or $Value -is [bool] -or $Value -is [string]) { return $false }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte] -or $Value -is [uint32] -or $Value -is [uint64]) { return $true }
    if ($Value -is [decimal] -or $Value -is [double] -or $Value -is [single]) {
        $number = [double]$Value
        return (-not [double]::IsNaN($number) -and -not [double]::IsInfinity($number) -and [System.Math]::Floor($number) -eq $number)
    }
    return $false
}

function ConvertTo-ConfigDate {
    # Data ISO 8601 ze strefą czasu (np. 2026-11-01T20:00:00Z) -> [DateTimeOffset]; inna wartość -> $null.
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.DateTimeOffset]) { return $Value }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Utc) { return (New-Object System.DateTimeOffset($Value, [System.TimeSpan]::Zero)) }
        return (New-Object System.DateTimeOffset($Value))
    }
    if ($Value -isnot [string]) { return $null }
    $text = $Value.Trim()
    if ($text -cnotmatch $script:StdConfigIsoDatePattern) { return $null }
    $parsed = [System.DateTimeOffset]::MinValue
    if (-not [System.DateTimeOffset]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $null
    }
    return $parsed
}

function ConvertTo-ConfigVersion {
    # Tekst wersji (np. 4.0.0, 24.08, 24) -> [version] z czterema częściami (brakujące = 0); inna wartość -> $null.
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    $version = $null
    if ($Value -is [version]) {
        $version = $Value
    } elseif ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -notmatch '^\d{1,9}(\.\d{1,9}){0,3}$') { return $null }
        if ($text -notmatch '\.') { $text = $text + '.0' }
        $parsed = $null
        if (-not [version]::TryParse($text, [ref]$parsed)) { return $null }
        $version = $parsed
    } else {
        return $null
    }
    $build = $version.Build
    if ($build -lt 0) { $build = 0 }
    $revision = $version.Revision
    if ($revision -lt 0) { $revision = 0 }
    return (New-Object System.Version($version.Major, $version.Minor, $build, $revision))
}

function Test-ConfigHttpsUrl {
    # Adres https:// z nazwą serwera, bez loginu i hasła w adresie. Wstawki {{secret:...}} są dozwolone.
    param([AllowNull()]$Value)
    if (-not (Test-ConfigIsText $Value)) { return $false }
    if ($Value -notmatch '^https://') { return $false }
    $probe = $Value -replace '\{\{secret:[^}]*\}\}', 'x'
    $uri = $null
    if (-not [System.Uri]::TryCreate($probe, [System.UriKind]::Absolute, [ref]$uri)) { return $false }
    return ($uri.Scheme -eq 'https' -and -not [string]::IsNullOrEmpty($uri.Host) -and [string]::IsNullOrEmpty($uri.UserInfo))
}

function Test-ConfigBase64 {
    # Tekst base64; -Length: wymagana liczba bajtów (0 = dowolna, ale niepusta).
    param([AllowNull()]$Value, [int]$Length = 0)
    if (-not (Test-ConfigIsText $Value)) { return $false }
    try {
        $bytes = [System.Convert]::FromBase64String($Value)
    } catch {
        return $false
    }
    if ($bytes.Length -eq 0) { return $false }
    if ($Length -gt 0 -and $bytes.Length -ne $Length) { return $false }
    return $true
}

function Get-ConfigSubstringCount {
    param([AllowEmptyString()][string]$Text, [Parameter(Mandatory = $true)][string]$Value)
    $count = 0
    $index = 0
    while ($true) {
        $found = $Text.IndexOf($Value, $index, [System.StringComparison]::OrdinalIgnoreCase)
        if ($found -lt 0) { break }
        $count++
        $index = $found + $Value.Length
    }
    return $count
}

function Get-ConfigSourceLabel {
    param([string]$Source)
    switch ($Source) {
        'server' { return 'serwer' }
        'mirror' { return 'kopia lustrzana' }
        'cache' { return 'kopia na tym komputerze (cache)' }
        'usb' { return 'pendrive' }
    }
    return $Source
}

function Assert-ConfigClientId {
    param([AllowNull()][AllowEmptyString()][string]$ClientId)
    if ([string]::IsNullOrEmpty($ClientId) -or $ClientId -cnotmatch $script:StdConfigIdPattern) {
        throw "[STD-CFG-ARGS] Nieprawidłowy identyfikator klienta $(Format-ConfigValue $ClientId): dozwolone są małe litery, cyfry i '-' (2-40 znaków, na początku litera lub cyfra)."
    }
}

function Save-ConfigCacheFile {
    # Atomowy zapis bajtów (koperta *.sdt) do cache: najpierw <plik>.tmp, potem podmiana.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][byte[]]$Bytes)
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    if (-not [string]::IsNullOrEmpty($dir) -and -not [System.IO.Directory]::Exists($dir)) {
        [void][System.IO.Directory]::CreateDirectory($dir)
    }
    $tmp = $Path + '.tmp'
    try {
        [System.IO.File]::WriteAllBytes($tmp, $Bytes)
        if ([System.IO.File]::Exists($Path)) {
            try {
                [System.IO.File]::Replace($tmp, $Path, [NullString]::Value)
            } catch {
                [System.IO.File]::Copy($tmp, $Path, $true)
                [System.IO.File]::Delete($tmp)
            }
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
    } catch {
        try { if ([System.IO.File]::Exists($tmp)) { [System.IO.File]::Delete($tmp) } } catch { }
        throw
    }
}

# --- Zaufane certyfikaty ---------------------------------------------------------------------

function Get-ConfigTrustedSignersPath {
    $configFile = $script:StdConfigFile
    if ([string]::IsNullOrEmpty($configFile)) { $configFile = $MyInvocation.MyCommand.ScriptBlock.File }
    return (Join-Path (Split-Path -Parent $configFile) 'Std.TrustedSigners.psd1')
}

function Read-ConfigTrustedSigners {
    # Plik w formacie Std.TrustedSigners.psd1 -> lista X509Certificate2. Wpisy bez certyfikatu są pomijane;
    # wpis z błędnym certyfikatem albo z odciskiem niezgodnym z certyfikatem jest pomijany z ostrzeżeniem.
    param([Parameter(Mandatory = $true)][string]$Path)
    $list = New-Object 'System.Collections.Generic.List[System.Security.Cryptography.X509Certificates.X509Certificate2]'
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Write-StdLog -Level Warn -Source 'Get-StdTrustedCertificates' -Message "Brak pliku z zaufanymi certyfikatami $Path."
        return , $list
    }
    try {
        $data = Import-PowerShellDataFile -LiteralPath $Path -ErrorAction Stop
    } catch {
        throw "[STD-SIG-NOTRUST] Nie udało się odczytać listy zaufanych certyfikatów $Path - $(Format-CoreErrorText $_ 200)"
    }
    $signers = $null
    if ($data -is [System.Collections.IDictionary] -and $data.Contains('Signers')) { $signers = $data['Signers'] }
    $position = 0
    foreach ($entry in @($signers)) {
        $position++
        if ($null -eq $entry) { continue }
        if (-not ($entry -is [System.Collections.IDictionary])) {
            Write-StdLog -Level Warn -Source 'Get-StdTrustedCertificates' -Message "Pominięto wpis nr $position w $Path - wpis musi być tablicą @{ Name; Thumbprint; Certificate }."
            continue
        }
        $name = [string]$entry['Name']
        if ([string]::IsNullOrWhiteSpace($name)) { $name = "nr $position" }
        $certificateText = [string]$entry['Certificate']
        if ([string]::IsNullOrWhiteSpace($certificateText)) { continue }
        $certificate = $null
        try {
            $raw = [System.Convert]::FromBase64String(($certificateText -replace '\s', ''))
            $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([byte[]]$raw)
        } catch {
            Write-StdLog -Level Warn -Source 'Get-StdTrustedCertificates' -Message "Pominięto zaufany certyfikat '$name' - pole Certificate nie jest poprawnym certyfikatem (base64 DER, plik .cer)."
            continue
        }
        $expected = ([string]$entry['Thumbprint']) -replace '\s', ''
        $actual = ([string]$certificate.Thumbprint).ToUpperInvariant()
        if (-not [string]::IsNullOrEmpty($expected) -and $expected.ToUpperInvariant() -cne $actual) {
            Write-StdLog -Level Warn -Source 'Get-StdTrustedCertificates' -Message "Pominięto zaufany certyfikat '$name' - odcisk $($expected.ToUpperInvariant()) nie zgadza się z certyfikatem ($actual)."
            continue
        }
        $list.Add($certificate)
    }
    return , $list
}

function Resolve-ConfigTrustedList {
    # Lista zaufanych certyfikatów: podana w parametrze (także pusta) albo z Std.TrustedSigners.psd1.
    param([bool]$Given, [AllowNull()]$TrustedCertificates)
    $candidates = $null
    if ($Given) {
        $candidates = $TrustedCertificates
    } else {
        $candidates = @(Get-StdTrustedCertificates)
    }
    $list = New-Object 'System.Collections.Generic.List[System.Security.Cryptography.X509Certificates.X509Certificate2]'
    foreach ($candidate in @($candidates)) {
        if ($null -ne $candidate -and $candidate -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) { $list.Add($candidate) }
    }
    return , $list
}

# --- Plik startowy (bootstrap) ----------------------------------------------------------------

function ConvertTo-ConfigBootstrap {
    # Sprawdza i normalizuje plik startowy (SPEC 2.3). Wynik zawiera tylko znane pola.
    param([AllowNull()]$InputObject, [Parameter(Mandatory = $true)][string]$Origin)
    $prefix = "[STD-CFG-BOOTSTRAP] $Origin -"
    $doc = $InputObject
    if ($null -ne $doc -and -not ($doc -is [System.Collections.IDictionary])) { $doc = ConvertTo-StdHashtable -InputObject $doc }
    if (-not ($doc -is [System.Collections.IDictionary])) {
        throw "$prefix plik startowy musi być obiektem JSON, np. { `"server`": `"https://serwer/std/`" }."
    }

    $server = $doc['server']
    if (-not (Test-ConfigIsText $server)) {
        throw "$prefix brak adresu serwera STD (pole server, np. https://pobieranie.firma.pl/std/)."
    }
    $server = $server.Trim()
    if (-not $server.EndsWith('/')) { $server = $server + '/' }
    $uri = $null
    $valid = [System.Uri]::TryCreate($server, [System.UriKind]::Absolute, [ref]$uri)
    if (-not $valid -or $uri.Scheme -ne 'https' -or [string]::IsNullOrEmpty($uri.Host)) {
        throw "$prefix adres serwera musi zaczynać się od https:// (podano: $(Get-HttpSafeLocation $server))."
    }
    if (-not [string]::IsNullOrEmpty($uri.UserInfo)) {
        throw "$prefix adres serwera nie może zawierać loginu ani hasła (podano: $(Get-HttpSafeLocation $server))."
    }
    if (-not [string]::IsNullOrEmpty($uri.Query) -or -not [string]::IsNullOrEmpty($uri.Fragment)) {
        throw "$prefix adres serwera nie może zawierać znaków ? ani # (podano: $(Get-HttpSafeLocation $server))."
    }

    $mirrors = New-Object 'System.Collections.Generic.List[object]'
    $rawMirrors = $doc['mirrors']
    if ($null -ne $rawMirrors) {
        if ($rawMirrors -is [string]) { $rawMirrors = @($rawMirrors) }
        if (-not (Test-ConfigIsList $rawMirrors)) {
            throw "$prefix pole mirrors musi być listą ścieżek UNC albo lokalnych."
        }
        foreach ($mirror in $rawMirrors) {
            if ($null -eq $mirror) { continue }
            if (-not (Test-ConfigIsText $mirror)) {
                throw "$prefix każda kopia lustrzana (mirrors) musi być niepustą ścieżką UNC albo lokalną."
            }
            if ($mirror -match '^[A-Za-z][A-Za-z0-9+.-]*://') {
                throw "$prefix kopia lustrzana musi być ścieżką UNC albo lokalną, nie adresem (podano: $(Get-HttpSafeLocation $mirror))."
            }
            $mirrors.Add($mirror.Trim())
        }
    }

    $mode = 'basic'
    $auth = $doc['auth']
    if ($null -ne $auth) {
        if (-not ($auth -is [System.Collections.IDictionary])) {
            throw "$prefix pole auth musi być obiektem, np. { `"mode`": `"basic`" }."
        }
        if ($null -ne $auth['mode']) {
            if (-not ($auth['mode'] -is [string]) -or $auth['mode'] -ne 'basic') {
                throw "$prefix nieobsługiwany sposób logowania auth.mode $(Format-ConfigValue $auth['mode']) - STD 4 obsługuje tylko 'basic'."
            }
        }
    }

    $proxy = $doc['proxy']
    if ($null -ne $proxy) {
        if (-not (Test-ConfigIsText $proxy)) {
            throw "$prefix pole proxy musi być adresem serwera proxy albo null."
        }
        $proxy = $proxy.Trim()
        $proxyUri = $null
        if (-not [System.Uri]::TryCreate($proxy, [System.UriKind]::Absolute, [ref]$proxyUri) -or [string]::IsNullOrEmpty($proxyUri.Host)) {
            throw "$prefix pole proxy musi być pełnym adresem serwera proxy (np. http://proxy.firma.local:8080/)."
        }
        if (-not [string]::IsNullOrEmpty($proxyUri.UserInfo)) {
            throw "$prefix adres proxy nie może zawierać loginu ani hasła."
        }
    }

    return [ordered]@{
        server  = $server
        mirrors = $mirrors.ToArray()
        auth    = [ordered]@{ mode = $mode }
        proxy   = $proxy
    }
}

# --- Walidacja konfiguracji -------------------------------------------------------------------

function Test-ConfigOptionalText {
    # Pole opcjonalne: jeśli jest, musi być tekstem (-NonEmpty: niepustym).
    param($Container, [string]$Key, [string]$Path, $Errors, [switch]$NonEmpty)
    if (-not (Test-ConfigHasValue $Container $Key)) { return }
    $value = $Container[$Key]
    $fieldPath = Join-ConfigPath $Path $Key
    if ($NonEmpty) {
        if (-not (Test-ConfigIsText $value)) { Add-ConfigError $Errors $fieldPath 'musi być niepustym tekstem.' }
    } elseif (-not ($value -is [string])) {
        Add-ConfigError $Errors $fieldPath 'musi być tekstem.'
    }
}

function Test-ConfigOptionalBool {
    param($Container, [string]$Key, [string]$Path, $Errors)
    if (-not (Test-ConfigHasValue $Container $Key)) { return }
    if (-not ($Container[$Key] -is [bool])) { Add-ConfigError $Errors (Join-ConfigPath $Path $Key) 'musi mieć wartość true albo false.' }
}

function Test-ConfigOptionalDate {
    param($Container, [string]$Key, [string]$Path, $Errors)
    if (-not (Test-ConfigHasValue $Container $Key)) { return }
    if ($null -eq (ConvertTo-ConfigDate $Container[$Key])) {
        Add-ConfigError $Errors (Join-ConfigPath $Path $Key) 'musi być datą ISO 8601 ze strefą czasu, np. 2026-11-01T20:00:00Z.'
    }
}

function Test-ConfigSha256Field {
    param($Container, [string]$Key, [string]$Path, $Errors)
    if (-not (Test-ConfigHasValue $Container $Key)) { return }
    $value = $Container[$Key]
    if (-not ($value -is [string]) -or $value.Trim() -notmatch $script:StdConfigSha256Pattern) {
        Add-ConfigError $Errors (Join-ConfigPath $Path $Key) 'musi być sumą SHA-256: 64 znaki HEX (0-9, a-f).'
    }
}

function Test-ConfigNameList {
    # Lista nazw, które muszą być kluczami słownika $Known (np. sourceOrder -> sources, apps -> catalog).
    param($Value, [string]$Path, $Errors, $Known, [string]$KnownName)
    if (-not (Test-ConfigIsList $Value)) {
        Add-ConfigError $Errors $Path 'musi być listą nazw (tablicą tekstów).'
        return
    }
    for ($i = 0; $i -lt $Value.Count; $i++) {
        $item = $Value[$i]
        $itemPath = Join-ConfigIndexPath $Path $i
        if (-not (Test-ConfigIsText $item)) {
            Add-ConfigError $Errors $itemPath 'musi być niepustym tekstem.'
        } elseif (-not (($Known -is [System.Collections.IDictionary]) -and $Known.Contains($item))) {
            Add-ConfigError $Errors $itemPath ("wskazuje {0}, którego nie ma w {1}." -f (Format-ConfigValue $item), $KnownName)
        }
    }
}

function Test-ConfigTasks {
    param($Value, [string]$Path, $Errors)
    if (-not ($Value -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors $Path 'musi być obiektem { "IdentyfikatorZadania": true/false }.'
        return
    }
    foreach ($entry in @($Value.GetEnumerator())) {
        $taskId = [string]$entry.Key
        $taskPath = Join-ConfigPath $Path $taskId
        if (-not ($script:StdConfigTaskIds -contains $taskId)) {
            Add-ConfigError $Errors $taskPath ('nieznane zadanie. Dozwolone: {0}.' -f ($script:StdConfigTaskIds -join ', '))
        } elseif (-not ($entry.Value -is [bool])) {
            Add-ConfigError $Errors $taskPath 'musi mieć wartość true albo false.'
        }
    }
}

function Test-ConfigScripts {
    param($Value, [string]$Path, $Errors)
    if (-not (Test-ConfigIsList $Value)) {
        Add-ConfigError $Errors $Path 'musi być listą skryptów [ { "file": "Skrypt.ps1" } ].'
        return
    }
    for ($i = 0; $i -lt $Value.Count; $i++) {
        $item = $Value[$i]
        $itemPath = Join-ConfigIndexPath $Path $i
        if (-not ($item -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors $itemPath 'musi być obiektem { "file": "Skrypt.ps1", "sha256": "...", "args": "" }.'
            continue
        }
        if (-not (Test-ConfigIsText $item['file'])) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'file') 'brak nazwy pliku skryptu.' }
        Test-ConfigSha256Field $item 'sha256' $itemPath $Errors
        Test-ConfigOptionalText $item 'args' $itemPath $Errors
    }
}

function Test-ConfigSystemSettings {
    param($Value, [string]$Path, $Errors)
    if (-not ($Value -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors $Path 'musi być obiektem z flagami (true/false) i listą registry.'
        return
    }
    foreach ($entry in @($Value.GetEnumerator())) {
        $key = [string]$entry.Key
        $entryPath = Join-ConfigPath $Path $key
        if ($key -eq 'registry') {
            if ($null -eq $entry.Value) { continue }
            if (-not (Test-ConfigIsList $entry.Value)) {
                Add-ConfigError $Errors $entryPath 'musi być listą wpisów rejestru.'
                continue
            }
            $registry = $entry.Value
            for ($i = 0; $i -lt $registry.Count; $i++) {
                $item = $registry[$i]
                $itemPath = Join-ConfigIndexPath $entryPath $i
                if (-not ($item -is [System.Collections.IDictionary])) {
                    Add-ConfigError $Errors $itemPath 'musi być obiektem { "hive", "path", "name", "type", "value" }.'
                    continue
                }
                $hive = $item['hive']
                if (-not ($hive -is [string]) -or -not ($script:StdConfigRegistryHives -contains $hive)) {
                    Add-ConfigError $Errors (Join-ConfigPath $itemPath 'hive') ('musi mieć jedną z wartości: {0}.' -f ($script:StdConfigRegistryHives -join ', '))
                }
                if (-not (Test-ConfigIsText $item['path'])) {
                    Add-ConfigError $Errors (Join-ConfigPath $itemPath 'path') 'brak ścieżki klucza rejestru (np. Software\\Firma).'
                }
                if (-not ($item['name'] -is [string])) {
                    Add-ConfigError $Errors (Join-ConfigPath $itemPath 'name') 'brak nazwy wartości rejestru ("" oznacza wartość domyślną).'
                }
                $type = $item['type']
                if (-not ($type -is [string]) -or -not ($script:StdConfigRegistryTypes -contains $type)) {
                    Add-ConfigError $Errors (Join-ConfigPath $itemPath 'type') ('musi mieć jedną z wartości: {0}.' -f ($script:StdConfigRegistryTypes -join ', '))
                }
            }
        } elseif ($null -ne $entry.Value -and -not ($entry.Value -is [bool])) {
            Add-ConfigError $Errors $entryPath 'flaga ustawień systemu musi mieć wartość true albo false.'
        }
    }
}

function Test-ConfigOverrides {
    # Ustawienia dziedziczone (defaults, lokalizacja, profil): tasks, bloatware, uninstall, systemSettings, scripts.
    param($Container, [string]$Path, $Errors)
    if (Test-ConfigHasValue $Container 'tasks') { Test-ConfigTasks -Value $Container['tasks'] -Path (Join-ConfigPath $Path 'tasks') -Errors $Errors }
    if (Test-ConfigHasValue $Container 'bloatware') {
        $bloatPath = Join-ConfigPath $Path 'bloatware'
        $bloatware = $Container['bloatware']
        if (-not (Test-ConfigIsList $bloatware)) {
            Add-ConfigError $Errors $bloatPath 'musi być listą wzorców nazw aplikacji Sklepu, np. [ "*TikTok*" ].'
        } else {
            for ($i = 0; $i -lt $bloatware.Count; $i++) {
                if (-not (Test-ConfigIsText $bloatware[$i])) { Add-ConfigError $Errors (Join-ConfigIndexPath $bloatPath $i) 'musi być niepustym wzorcem nazwy.' }
            }
        }
    }
    if (Test-ConfigHasValue $Container 'uninstall') {
        $uninstallPath = Join-ConfigPath $Path 'uninstall'
        $uninstall = $Container['uninstall']
        if (-not (Test-ConfigIsList $uninstall)) {
            Add-ConfigError $Errors $uninstallPath 'musi być listą programów [ { "name": "McAfee*", "args": "" } ].'
        } else {
            for ($i = 0; $i -lt $uninstall.Count; $i++) {
                $item = $uninstall[$i]
                $itemPath = Join-ConfigIndexPath $uninstallPath $i
                if (-not ($item -is [System.Collections.IDictionary])) {
                    Add-ConfigError $Errors $itemPath 'musi być obiektem { "name": "McAfee*", "args": "" }.'
                    continue
                }
                if (-not (Test-ConfigIsText $item['name'])) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'name') 'brak nazwy (wzorca) programu do usunięcia.' }
                Test-ConfigOptionalText $item 'args' $itemPath $Errors
            }
        }
    }
    if (Test-ConfigHasValue $Container 'systemSettings') { Test-ConfigSystemSettings -Value $Container['systemSettings'] -Path (Join-ConfigPath $Path 'systemSettings') -Errors $Errors }
    if (Test-ConfigHasValue $Container 'scripts') { Test-ConfigScripts -Value $Container['scripts'] -Path (Join-ConfigPath $Path 'scripts') -Errors $Errors }
}

function Test-ConfigSources {
    param($Document, $Errors)
    if (-not (Test-ConfigHasValue $Document 'sources')) {
        Add-ConfigError $Errors 'sources' 'brak wymaganego pola: co najmniej jedno źródło instalatorów (https albo unc).'
        return
    }
    $sources = $Document['sources']
    if (-not ($sources -is [System.Collections.IDictionary]) -or $sources.get_Count() -eq 0) {
        Add-ConfigError $Errors 'sources' 'musi być obiektem z co najmniej jednym źródłem instalatorów.'
        return
    }
    foreach ($entry in @($sources.GetEnumerator())) {
        $sourcePath = Join-ConfigPath 'sources' ([string]$entry.Key)
        $source = $entry.Value
        if (-not ($source -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors $sourcePath 'musi być obiektem { "type", "url" albo "path", "auth" }.'
            continue
        }
        $type = $source['type']
        if (-not ($type -is [string]) -or -not ($script:StdConfigSourceTypes -contains $type)) {
            Add-ConfigError $Errors (Join-ConfigPath $sourcePath 'type') 'musi mieć wartość https albo unc.'
        } elseif ($type -eq 'https') {
            if (-not (Test-ConfigHttpsUrl $source['url'])) {
                Add-ConfigError $Errors (Join-ConfigPath $sourcePath 'url') 'musi być pełnym adresem zaczynającym się od https:// (bez loginu i hasła w adresie).'
            }
        } else {
            if (-not (Test-ConfigIsText $source['path'])) {
                Add-ConfigError $Errors (Join-ConfigPath $sourcePath 'path') 'brak ścieżki udziału (np. \\serwer\STD$\files\klient\).'
            }
        }
        $auth = $source['auth']
        if (-not ($auth -is [string]) -or -not ($script:StdConfigSourceAuth -contains $auth)) {
            Add-ConfigError $Errors (Join-ConfigPath $sourcePath 'auth') 'musi mieć wartość session albo none.'
        }
    }
}

function Test-ConfigCatalog {
    param($Document, $Errors)
    if (-not (Test-ConfigHasValue $Document 'catalog')) {
        Add-ConfigError $Errors 'catalog' 'brak wymaganego pola: katalog musi mieć co najmniej jedną pozycję.'
        return
    }
    $catalog = $Document['catalog']
    if (-not ($catalog -is [System.Collections.IDictionary]) -or $catalog.get_Count() -eq 0) {
        Add-ConfigError $Errors 'catalog' 'musi być obiektem z co najmniej jedną pozycją.'
        return
    }
    $sources = $Document['sources']
    $requireSha = ($Document['requireSha256'] -is [bool]) -and [bool]$Document['requireSha256']
    foreach ($entry in @($catalog.GetEnumerator())) {
        $key = [string]$entry.Key
        $itemPath = Join-ConfigPath 'catalog' $key
        if ($key -cnotmatch $script:StdConfigCatalogKeyPattern) {
            Add-ConfigError $Errors $itemPath "nieprawidłowy klucz katalogu: dozwolone litery, cyfry i znaki . _ + - (do 64 znaków, na początku litera lub cyfra)."
        }
        $item = $entry.Value
        if (-not ($item -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors $itemPath 'musi być obiektem (pozycja katalogu).'
            continue
        }
        if (-not (Test-ConfigIsText $item['name'])) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'name') 'brak nazwy wyświetlanej programu.' }

        $hasFile = Test-ConfigHasValue $item 'file'
        $hasUrl = Test-ConfigHasValue $item 'url'
        $hasWinget = Test-ConfigHasValue $item 'winget'
        if (-not ($hasFile -or $hasUrl -or $hasWinget)) {
            Add-ConfigError $Errors $itemPath 'musi mieć co najmniej jedno z pól: file, url, winget.'
        }
        if ($hasFile -and -not (Test-ConfigIsText $item['file'])) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'file') 'musi być niepustą nazwą pliku w źródle.' }
        if ($hasUrl -and -not (Test-ConfigHttpsUrl $item['url'])) {
            Add-ConfigError $Errors (Join-ConfigPath $itemPath 'url') 'musi być pełnym adresem zaczynającym się od https:// (bez loginu i hasła w adresie).'
        }
        if ($hasWinget -and -not (Test-ConfigIsText $item['winget'])) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'winget') 'musi być identyfikatorem pakietu winget (np. Google.Chrome).' }
        Test-ConfigOptionalText $item 'wingetArgs' $itemPath $Errors
        Test-ConfigOptionalText $item 'args' $itemPath $Errors
        Test-ConfigSha256Field $item 'sha256' $itemPath $Errors
        if ($requireSha -and ($hasFile -or $hasUrl) -and -not (Test-ConfigHasValue $item 'sha256')) {
            Add-ConfigError $Errors (Join-ConfigPath $itemPath 'sha256') 'konfiguracja ma requireSha256: true, więc pozycja z file / url musi mieć sumę SHA-256.'
        }

        if (Test-ConfigHasValue $item 'detect') {
            $detectPath = Join-ConfigPath $itemPath 'detect'
            $detect = $item['detect']
            if (-not ($detect -is [System.Collections.IDictionary])) {
                Add-ConfigError $Errors $detectPath 'musi być obiektem { "uninstallName": "...", "minVersion": "..." } albo { "path": "..." }.'
            } else {
                if (-not ((Test-ConfigHasValue $detect 'uninstallName') -or (Test-ConfigHasValue $detect 'path'))) {
                    Add-ConfigError $Errors $detectPath 'musi mieć pole uninstallName albo path.'
                }
                Test-ConfigOptionalText $detect 'uninstallName' $detectPath $Errors -NonEmpty
                Test-ConfigOptionalText $detect 'path' $detectPath $Errors -NonEmpty
                if ((Test-ConfigHasValue $detect 'minVersion') -and $null -eq (ConvertTo-ConfigVersion $detect['minVersion'])) {
                    Add-ConfigError $Errors (Join-ConfigPath $detectPath 'minVersion') 'musi być numerem wersji, np. 24.08 albo 1.2.3.'
                }
            }
        }
        if (Test-ConfigHasValue $item 'timeoutMin') {
            $timeout = $item['timeoutMin']
            if (-not (Test-ConfigIsInteger $timeout) -or [double]$timeout -lt 1 -or [double]$timeout -gt 240) {
                Add-ConfigError $Errors (Join-ConfigPath $itemPath 'timeoutMin') 'musi być liczbą całkowitą od 1 do 240 (minuty).'
            }
        }
        if (Test-ConfigHasValue $item 'successCodes') {
            $codesPath = Join-ConfigPath $itemPath 'successCodes'
            $codes = $item['successCodes']
            if (-not (Test-ConfigIsList $codes) -or $codes.Count -eq 0) {
                Add-ConfigError $Errors $codesPath 'musi być niepustą listą liczb całkowitych, np. [0, 1641, 3010].'
            } else {
                for ($i = 0; $i -lt $codes.Count; $i++) {
                    if (-not (Test-ConfigIsInteger $codes[$i])) { Add-ConfigError $Errors (Join-ConfigIndexPath $codesPath $i) 'musi być liczbą całkowitą.' }
                }
            }
        }
        if (Test-ConfigHasValue $item 'whenDetected') {
            $when = $item['whenDetected']
            if (-not ($when -is [string]) -or -not ($script:StdConfigWhenDetected -contains $when)) {
                Add-ConfigError $Errors (Join-ConfigPath $itemPath 'whenDetected') 'musi mieć wartość skip albo reinstall.'
            }
        }
        if (Test-ConfigHasValue $item 'sourceOrder') {
            Test-ConfigNameList -Value $item['sourceOrder'] -Path (Join-ConfigPath $itemPath 'sourceOrder') -Errors $Errors -Known $sources -KnownName 'sources'
        }
    }
}

function Test-ConfigComputerName {
    param($Value, [string]$Path, $Errors)
    if (-not ($Value -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors $Path 'musi być obiektem { "pattern": "KRK-{serial}", "maxLength": 15, "editable": true }.'
        return
    }
    $maxLength = $script:StdConfigMaxComputerName
    if (Test-ConfigHasValue $Value 'maxLength') {
        $raw = $Value['maxLength']
        if (-not (Test-ConfigIsInteger $raw) -or [double]$raw -lt 1 -or [double]$raw -gt $script:StdConfigMaxComputerName) {
            Add-ConfigError $Errors (Join-ConfigPath $Path 'maxLength') "musi być liczbą całkowitą od 1 do $($script:StdConfigMaxComputerName) (Windows dopuszcza najwyżej $($script:StdConfigMaxComputerName) znaków nazwy komputera)."
        }
    }
    Test-ConfigOptionalBool $Value 'editable' $Path $Errors
    $patternPath = Join-ConfigPath $Path 'pattern'
    $pattern = $Value['pattern']
    if (-not (Test-ConfigIsText $pattern)) {
        Add-ConfigError $Errors $patternPath 'brak wzorca nazwy komputera (np. KRK-{serial}).'
        return
    }
    foreach ($match in [regex]::Matches($pattern, '\{([^{}]*)\}')) {
        $placeholder = $match.Groups[1].Value
        if (-not ($script:StdConfigNamePlaceholders -contains $placeholder)) {
            Add-ConfigError $Errors $patternPath ('niedozwolona wstawka {0}. Dozwolone: {{serial}}, {{site}}, {{random}}.' -f (Format-ConfigValue ('{' + $placeholder + '}')))
        }
    }
    $rest = [regex]::Replace($pattern, '\{[^{}]*\}', '')
    if ($rest.Contains('{') -or $rest.Contains('}')) {
        Add-ConfigError $Errors $patternPath 'nawias { albo } bez pary - wstawki zapisuje się jako {serial}, {site}, {random}.'
    }
}

function Test-ConfigWifi {
    param($Value, [string]$Path, $Errors, $SecretNames)
    if (-not (Test-ConfigIsList $Value)) {
        Add-ConfigError $Errors $Path 'musi być listą sieci [ { "ssid", "auth", "keyRef" } ] albo [ { "file": "Profil.xml" } ].'
        return
    }
    for ($i = 0; $i -lt $Value.Count; $i++) {
        $item = $Value[$i]
        $itemPath = Join-ConfigIndexPath $Path $i
        if (-not ($item -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors $itemPath 'musi być obiektem { "ssid", "auth", "keyRef" } albo { "file" }.'
            continue
        }
        $hasFile = Test-ConfigHasValue $item 'file'
        $hasSsid = Test-ConfigHasValue $item 'ssid'
        if ($hasFile -and $hasSsid) {
            Add-ConfigError $Errors $itemPath 'podaj ssid (z auth i keyRef) albo file (dawny profil XML), nie oba.'
            continue
        }
        if ($hasFile) {
            if (-not (Test-ConfigIsText $item['file'])) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'file') 'musi być nazwą pliku profilu Wi-Fi (XML).' }
            continue
        }
        if (-not (Test-ConfigIsText $item['ssid'])) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'ssid') 'brak nazwy sieci (SSID).' }
        $auth = $item['auth']
        if (-not ($auth -is [string]) -or -not ($script:StdConfigWifiAuth -contains $auth)) {
            Add-ConfigError $Errors (Join-ConfigPath $itemPath 'auth') ('musi mieć jedną z wartości: {0}.' -f ($script:StdConfigWifiAuth -join ', '))
        }
        $keyRefPath = Join-ConfigPath $itemPath 'keyRef'
        $isOpen = ($auth -is [string]) -and ($auth -eq 'open')
        if (Test-ConfigHasValue $item 'keyRef') {
            $keyRef = $item['keyRef']
            if (-not ($keyRef -is [string]) -or $keyRef -cnotmatch $script:StdConfigSecretNamePattern) {
                Add-ConfigError $Errors $keyRefPath 'musi być nazwą sekretu z secrets.items (litery, cyfry, . _ -).'
            } elseif (-not $SecretNames.Contains($keyRef)) {
                Add-ConfigError $Errors $keyRefPath ("brak sekretu '{0}' w secrets.items - dodaj go w edytorze administratora." -f $keyRef)
            }
        } elseif (-not $isOpen) {
            Add-ConfigError $Errors $keyRefPath 'sieć z hasłem wymaga keyRef - nazwy sekretu z kluczem sieci w secrets.items.'
        }
    }
}

function Test-ConfigProfile {
    param($ProfileNode, [string]$Path, $Errors, $Catalog)
    if (-not ($profileNode -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors $Path 'musi być obiektem { "apps", "tasks", "scripts", "locked" }.'
        return
    }
    if (Test-ConfigHasValue $profileNode 'apps') {
        Test-ConfigNameList -Value $profileNode['apps'] -Path (Join-ConfigPath $Path 'apps') -Errors $Errors -Known $Catalog -KnownName 'catalog'
    }
    Test-ConfigOverrides -Container $profileNode -Path $Path -Errors $Errors
    if (Test-ConfigHasValue $profileNode 'locked') {
        $lockedPath = Join-ConfigPath $Path 'locked'
        $locked = $profileNode['locked']
        $allowed = @($script:StdConfigTaskIds) + @($script:StdConfigLockableExtra)
        if (-not (Test-ConfigIsList $locked)) {
            Add-ConfigError $Errors $lockedPath 'musi być listą, np. [ "JoinDomain", "computerName", "apps" ].'
        } else {
            for ($i = 0; $i -lt $locked.Count; $i++) {
                $value = $locked[$i]
                if (-not ($value -is [string]) -or -not ($allowed -contains $value)) {
                    Add-ConfigError $Errors (Join-ConfigIndexPath $lockedPath $i) ('niedozwolona wartość {0}. Dozwolone: identyfikatory zadań, computerName, apps.' -f (Format-ConfigValue $value))
                }
            }
        }
    }
}

function Test-ConfigSites {
    param($Document, $Errors, $SecretNames)
    if (-not (Test-ConfigHasValue $Document 'sites')) {
        Add-ConfigError $Errors 'sites' 'brak wymaganego pola: co najmniej jedna lokalizacja.'
        return
    }
    $sites = $Document['sites']
    if (-not (Test-ConfigIsList $sites) -or $sites.Count -eq 0) {
        Add-ConfigError $Errors 'sites' 'musi być niepustą listą lokalizacji.'
        return
    }
    $sources = $Document['sources']
    $catalog = $Document['catalog']
    $seenIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    for ($i = 0; $i -lt $sites.Count; $i++) {
        $site = $sites[$i]
        $sitePath = Join-ConfigIndexPath 'sites' $i
        if (-not ($site -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors $sitePath 'musi być obiektem (lokalizacja).'
            continue
        }
        $id = $site['id']
        if (-not ($id -is [string]) -or $id -cnotmatch $script:StdConfigIdPattern) {
            Add-ConfigError $Errors (Join-ConfigPath $sitePath 'id') ("nieprawidłowy identyfikator lokalizacji {0}: dozwolone małe litery, cyfry i '-' (2-40 znaków, na początku litera lub cyfra)." -f (Format-ConfigValue $id))
        } elseif (-not $seenIds.Add($id)) {
            Add-ConfigError $Errors (Join-ConfigPath $sitePath 'id') ("identyfikator lokalizacji '{0}' się powtarza - musi być unikalny." -f $id)
        }
        if (-not (Test-ConfigIsText $site['name'])) { Add-ConfigError $Errors (Join-ConfigPath $sitePath 'name') 'brak nazwy lokalizacji.' }
        if (Test-ConfigHasValue $site 'sourceOrder') {
            Test-ConfigNameList -Value $site['sourceOrder'] -Path (Join-ConfigPath $sitePath 'sourceOrder') -Errors $Errors -Known $sources -KnownName 'sources'
        }
        if (Test-ConfigHasValue $site 'wifi') { Test-ConfigWifi -Value $site['wifi'] -Path (Join-ConfigPath $sitePath 'wifi') -Errors $Errors -SecretNames $SecretNames }
        if (Test-ConfigHasValue $site 'computerName') { Test-ConfigComputerName -Value $site['computerName'] -Path (Join-ConfigPath $sitePath 'computerName') -Errors $Errors }
        if (Test-ConfigHasValue $site 'domainJoin') {
            $joinPath = Join-ConfigPath $sitePath 'domainJoin'
            $join = $site['domainJoin']
            if (-not ($join -is [System.Collections.IDictionary])) {
                Add-ConfigError $Errors $joinPath 'musi być obiektem { "domain": "firma.local", "ou": "OU=...,DC=firma,DC=local" }.'
            } else {
                if (-not (Test-ConfigIsText $join['domain'])) { Add-ConfigError $Errors (Join-ConfigPath $joinPath 'domain') 'brak nazwy domeny (np. firma.local).' }
                Test-ConfigOptionalText $join 'ou' $joinPath $Errors
            }
        }
        if (Test-ConfigHasValue $site 'localAdmin') {
            $adminPath = Join-ConfigPath $sitePath 'localAdmin'
            $admin = $site['localAdmin']
            if (-not ($admin -is [System.Collections.IDictionary])) {
                Add-ConfigError $Errors $adminPath 'musi być obiektem { "username": "LokalnyIT" }.'
            } elseif (-not (Test-ConfigIsText $admin['username'])) {
                Add-ConfigError $Errors (Join-ConfigPath $adminPath 'username') 'brak nazwy konta lokalnego administratora.'
            }
        }
        if (Test-ConfigHasValue $site 'bitLocker') {
            $bitLockerPath = Join-ConfigPath $sitePath 'bitLocker'
            $bitLocker = $site['bitLocker']
            if (-not ($bitLocker -is [System.Collections.IDictionary])) {
                Add-ConfigError $Errors $bitLockerPath 'musi być obiektem { "escrow": [ "ad", "server" ] }.'
            } elseif (Test-ConfigHasValue $bitLocker 'escrow') {
                $escrow = $bitLocker['escrow']
                $escrowPath = Join-ConfigPath $bitLockerPath 'escrow'
                if (-not (Test-ConfigIsList $escrow)) {
                    Add-ConfigError $Errors $escrowPath 'musi być listą: ad, server.'
                } else {
                    for ($j = 0; $j -lt $escrow.Count; $j++) {
                        if (-not ($escrow[$j] -is [string]) -or -not ($script:StdConfigBitLockerEscrow -contains $escrow[$j])) {
                            Add-ConfigError $Errors (Join-ConfigIndexPath $escrowPath $j) 'musi mieć wartość ad albo server.'
                        }
                    }
                }
            }
        }
        Test-ConfigOverrides -Container $site -Path $sitePath -Errors $Errors

        $profilesPath = Join-ConfigPath $sitePath 'profiles'
        $profiles = $site['profiles']
        if (-not ($profiles -is [System.Collections.IDictionary]) -or $profiles.get_Count() -eq 0) {
            Add-ConfigError $Errors $profilesPath 'lokalizacja musi mieć co najmniej jeden profil.'
            continue
        }
        foreach ($entry in @($profiles.GetEnumerator())) {
            $profileName = [string]$entry.Key
            $profilePath = Join-ConfigPath $profilesPath $profileName
            if ([string]::IsNullOrWhiteSpace($profileName)) {
                Add-ConfigError $Errors $profilePath 'nazwa profilu nie może być pusta.'
            }
            Test-ConfigProfile -ProfileNode $entry.Value -Path $profilePath -Errors $Errors -Catalog $catalog
        }
    }
}

function Test-ConfigSecretsBlock {
    # Budowa magazynu sekretów (SPEC 2.7), bez odszyfrowania. Zwraca zbiór nazw pozycji.
    param($Document, $Errors)
    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-ConfigHasValue $Document 'secrets')) { return , $names }
    $secrets = $Document['secrets']
    if (-not ($secrets -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors 'secrets' 'musi być obiektem (zaszyfrowane sekrety, tryb sealed).'
        return , $names
    }
    if (-not ($secrets['mode'] -is [string]) -or $secrets['mode'] -cne 'sealed') {
        Add-ConfigError $Errors 'secrets.mode' "musi mieć wartość 'sealed'."
    }
    $kdf = $secrets['kdf']
    if (-not ($kdf -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors 'secrets.kdf' 'musi być obiektem { "alg": "pbkdf2-sha256", "iter", "salt" }.'
    } else {
        if (-not ($kdf['alg'] -is [string]) -or $kdf['alg'] -cne 'pbkdf2-sha256') { Add-ConfigError $Errors 'secrets.kdf.alg' "musi mieć wartość 'pbkdf2-sha256'." }
        if (-not (Test-ConfigIsInteger $kdf['iter']) -or [double]$kdf['iter'] -lt 1000 -or [double]$kdf['iter'] -gt 100000000) {
            Add-ConfigError $Errors 'secrets.kdf.iter' 'musi być liczbą całkowitą od 1000 do 100000000.'
        }
        $saltOk = $false
        if (Test-ConfigBase64 $kdf['salt']) { $saltOk = ([System.Convert]::FromBase64String($kdf['salt']).Length -ge 8) }
        if (-not $saltOk) { Add-ConfigError $Errors 'secrets.kdf.salt' 'musi być tekstem base64 (co najmniej 8 bajtów).' }
    }
    if (-not (Test-ConfigBase64 $secrets['check'] 32)) { Add-ConfigError $Errors 'secrets.check' 'musi być 32 bajtami w base64.' }
    $items = $secrets['items']
    if ($null -eq $items) { return , $names }
    if (-not ($items -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors 'secrets.items' 'musi być obiektem { "nazwa": { "iv", "ct", "mac" } }.'
        return , $names
    }
    foreach ($entry in @($items.GetEnumerator())) {
        $name = [string]$entry.Key
        $itemPath = Join-ConfigPath 'secrets.items' $name
        if ($name -cnotmatch $script:StdConfigSecretNamePattern) {
            Add-ConfigError $Errors $itemPath 'nieprawidłowa nazwa sekretu: dozwolone litery, cyfry i znaki . _ - (do 64 znaków, na początku litera lub cyfra).'
        } else {
            [void]$names.Add($name)
        }
        if ($name -match $script:StdConfigCredentialSecretTarget -and $name -match $script:StdConfigCredentialSecretKind) {
            Add-ConfigError $Errors $itemPath $script:StdConfigCredentialMessage
        }
        $item = $entry.Value
        if (-not ($item -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors $itemPath 'musi być obiektem { "iv", "ct", "mac" }.'
            continue
        }
        if (-not (Test-ConfigBase64 $item['iv'] 16)) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'iv') 'musi być 16 bajtami w base64.' }
        $ctOk = $false
        if (Test-ConfigBase64 $item['ct']) { $ctOk = (([System.Convert]::FromBase64String($item['ct']).Length % 16) -eq 0) }
        if (-not $ctOk) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'ct') 'musi być szyfrogramem AES w base64 (wielokrotność 16 bajtów).' }
        if (-not (Test-ConfigBase64 $item['mac'] 32)) { Add-ConfigError $Errors (Join-ConfigPath $itemPath 'mac') 'musi być 32 bajtami w base64.' }
    }
    return , $names
}

function Test-ConfigCredentialSection {
    # Zawartość domainJoin / localAdmin: żadnego pola z danymi logowania i żadnej wstawki sekretu.
    param($Node, [string]$Path, [bool]$IsLocalAdmin, $Errors, [int]$Depth = 0)
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($entry in @($Node.GetEnumerator())) {
            $key = [string]$entry.Key
            $childPath = Join-ConfigPath $Path $key
            $allowed = $IsLocalAdmin -and $Depth -eq 0 -and $key -eq 'username'
            if (-not $allowed -and $key -match $script:StdConfigCredentialFieldPattern) {
                Add-ConfigError $Errors $childPath $script:StdConfigCredentialMessage
            } else {
                Test-ConfigCredentialSection -Node $entry.Value -Path $childPath -IsLocalAdmin $IsLocalAdmin -Errors $Errors -Depth ($Depth + 1)
            }
        }
    } elseif (Test-ConfigIsList $Node) {
        for ($i = 0; $i -lt $Node.Count; $i++) {
            Test-ConfigCredentialSection -Node $Node[$i] -Path (Join-ConfigIndexPath $Path $i) -IsLocalAdmin $IsLocalAdmin -Errors $Errors -Depth ($Depth + 1)
        }
    } elseif ($Node -is [string] -and $Node.IndexOf('{{secret:', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        Add-ConfigError $Errors $Path ($script:StdConfigCredentialMessage + ' Wstawka {{secret:...}} jest tu niedozwolona.')
    }
}

function Test-ConfigDocumentWalk {
    # Reguły dla całego dokumentu (poza secrets): sekcje domainJoin / localAdmin w dowolnym miejscu,
    # jawne hasła i tokeny w polach o "podejrzanych" nazwach, wstawki {{secret:nazwa}} wskazujące istniejące sekrety.
    param($Node, [string]$Path, $Errors, $SecretNames)
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($entry in @($Node.GetEnumerator())) {
            $key = [string]$entry.Key
            if ([string]::IsNullOrEmpty($Path) -and $key -eq 'secrets') { continue }
            $childPath = Join-ConfigPath $Path $key
            $value = $entry.Value
            if ($script:StdConfigPlaintextKeys -contains $key -and $value -is [string] -and $value.Length -gt 0) {
                # Wartości nie wypisujemy - to może być hasło albo token.
                Add-ConfigError $Errors $childPath 'pole o tej nazwie nie może zawierać jawnego hasła ani tokenu. Sekrety zapisuje się zaszyfrowane w secrets i wstawia przez {{secret:nazwa}}.'
            }
            if ($key -eq 'domainJoin' -or $key -eq 'localAdmin') {
                Test-ConfigCredentialSection -Node $value -Path $childPath -IsLocalAdmin ($key -eq 'localAdmin') -Errors $Errors
            }
            Test-ConfigDocumentWalk -Node $value -Path $childPath -Errors $Errors -SecretNames $SecretNames
        }
    } elseif (Test-ConfigIsList $Node) {
        for ($i = 0; $i -lt $Node.Count; $i++) {
            Test-ConfigDocumentWalk -Node $Node[$i] -Path (Join-ConfigIndexPath $Path $i) -Errors $Errors -SecretNames $SecretNames
        }
    } elseif ($Node -is [string] -and $Node.IndexOf('{{secret:', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $total = Get-ConfigSubstringCount -Text $Node -Value '{{secret:'
        $valid = 0
        if ($null -ne $script:StdCryptoSecretRefRegex) { $valid = $script:StdCryptoSecretRefRegex.Matches($Node).Count }
        if ($total -gt $valid) {
            Add-ConfigError $Errors $Path 'nieprawidłowa wstawka sekretu - zapis to {{secret:nazwa}} (nazwa: litery, cyfry, . _ -).'
        }
        foreach ($name in @(Get-StdSecretReference -Text $Node)) {
            if (-not $SecretNames.Contains($name)) {
                Add-ConfigError $Errors $Path ("brak sekretu '{0}' w secrets.items - dodaj go w edytorze administratora." -f $name)
            }
        }
    }
}

function Test-ConfigReleaseDocument {
    param([Parameter(Mandatory = $true)]$Document, [bool]$Unpublished, [Parameter(Mandatory = $true)][AllowEmptyCollection()]$Errors)
    $doc = $Document

    if (-not (Test-ConfigHasValue $doc 'schemaVersion')) {
        Add-ConfigError $Errors 'schemaVersion' 'brak wymaganego pola (wartość 4).'
    } elseif (-not (Test-ConfigIsInteger $doc['schemaVersion']) -or [double]$doc['schemaVersion'] -ne 4) {
        Add-ConfigError $Errors 'schemaVersion' ('musi mieć wartość 4 (jest {0}).' -f (Format-ConfigValue $doc['schemaVersion']))
    }
    if (-not ($doc['kind'] -is [string]) -or $doc['kind'] -cne 'std.release') {
        Add-ConfigError $Errors 'kind' ("musi mieć wartość 'std.release' (jest {0})." -f (Format-ConfigValue $doc['kind']))
    }

    $client = $doc['client']
    if (-not ($client -is [System.Collections.IDictionary])) {
        Add-ConfigError $Errors 'client' 'brak wymaganego obiektu { "id", "name" }.'
    } else {
        $clientId = $client['id']
        if (-not ($clientId -is [string]) -or $clientId -cnotmatch $script:StdConfigIdPattern) {
            Add-ConfigError $Errors 'client.id' ("nieprawidłowy identyfikator klienta {0}: dozwolone małe litery, cyfry i '-' (2-40 znaków, na początku litera lub cyfra)." -f (Format-ConfigValue $clientId))
        } elseif ($clientId -eq $script:StdConfigIndexRevisionKey) {
            Add-ConfigError $Errors 'client.id' "identyfikator '$($script:StdConfigIndexRevisionKey)' jest zarezerwowany dla listy klientów."
        }
        if (-not (Test-ConfigIsText $client['name'])) { Add-ConfigError $Errors 'client.name' 'brak nazwy klienta.' }
    }

    if (Test-ConfigHasValue $doc 'revision') {
        if (-not (Test-ConfigIsInteger $doc['revision']) -or [double]$doc['revision'] -lt 1 -or [double]$doc['revision'] -gt [int]::MaxValue) {
            Add-ConfigError $Errors 'revision' 'musi być liczbą całkowitą większą od zera.'
        }
    } elseif (-not $Unpublished) {
        Add-ConfigError $Errors 'revision' 'brak numeru rewizji (ustawia go Publish-StdRelease).'
    }
    if (Test-ConfigHasValue $doc 'publishedAt') {
        Test-ConfigOptionalDate $doc 'publishedAt' '' $Errors
    } elseif (-not $Unpublished) {
        Add-ConfigError $Errors 'publishedAt' 'brak daty publikacji (ustawia ją Publish-StdRelease).'
    }
    if (Test-ConfigHasValue $doc 'publishedBy') {
        Test-ConfigOptionalText $doc 'publishedBy' '' $Errors -NonEmpty
    } elseif (-not $Unpublished) {
        Add-ConfigError $Errors 'publishedBy' 'brak loginu publikującego (ustawia go Publish-StdRelease).'
    }
    Test-ConfigOptionalText $doc 'changeNote' '' $Errors
    Test-ConfigOptionalDate $doc 'expiresAt' '' $Errors
    if ((Test-ConfigHasValue $doc 'minClientVersion') -and $null -eq (ConvertTo-ConfigVersion $doc['minClientVersion'])) {
        Add-ConfigError $Errors 'minClientVersion' 'musi być numerem wersji STD, np. 4.0.0.'
    }
    Test-ConfigOptionalBool $doc 'allowOffline' '' $Errors
    Test-ConfigOptionalBool $doc 'requireSha256' '' $Errors

    if (Test-ConfigHasValue $doc 'tool') {
        $tool = $doc['tool']
        if (-not ($tool -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors 'tool' 'musi być obiektem { "version", "url", "sha256" }.'
        } else {
            if ((Test-ConfigHasValue $tool 'version') -and $null -eq (ConvertTo-ConfigVersion $tool['version'])) { Add-ConfigError $Errors 'tool.version' 'musi być numerem wersji, np. 4.0.1.' }
            if ((Test-ConfigHasValue $tool 'url') -and -not (Test-ConfigHttpsUrl $tool['url'])) { Add-ConfigError $Errors 'tool.url' 'musi być pełnym adresem zaczynającym się od https://.' }
            Test-ConfigSha256Field $tool 'sha256' 'tool' $Errors
        }
    }
    if (Test-ConfigHasValue $doc 'escrow') {
        $escrow = $doc['escrow']
        if (-not ($escrow -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors 'escrow' 'musi być obiektem { "drop": "drop/", "recipientCert": "<base64 DER>" }.'
        } else {
            Test-ConfigOptionalText $escrow 'drop' 'escrow' $Errors -NonEmpty
            if ((Test-ConfigHasValue $escrow 'recipientCert') -and -not (Test-ConfigBase64 $escrow['recipientCert'])) { Add-ConfigError $Errors 'escrow.recipientCert' 'musi być certyfikatem w base64 (DER).' }
        }
    }

    Test-ConfigSources -Document $doc -Errors $Errors
    Test-ConfigCatalog -Document $doc -Errors $Errors
    if (Test-ConfigHasValue $doc 'defaults') {
        $defaults = $doc['defaults']
        if (-not ($defaults -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors 'defaults' 'musi być obiektem { "tasks", "bloatware", "uninstall", "systemSettings", "scripts" }.'
        } else {
            Test-ConfigOverrides -Container $defaults -Path 'defaults' -Errors $Errors
        }
    }
    $secretNames = Test-ConfigSecretsBlock -Document $doc -Errors $Errors
    Test-ConfigSites -Document $doc -Errors $Errors -SecretNames $secretNames
    Test-ConfigDocumentWalk -Node $doc -Path '' -Errors $Errors -SecretNames $secretNames
}

function Test-ConfigClientIndexDocument {
    # Indeks klientów (payload clients.sdt, SPEC 2.5).
    param([Parameter(Mandatory = $true)]$Document, [Parameter(Mandatory = $true)][AllowEmptyCollection()]$Errors)
    $doc = $Document
    if (-not (Test-ConfigIsInteger $doc['schemaVersion']) -or [double]$doc['schemaVersion'] -ne 4) { Add-ConfigError $Errors 'schemaVersion' 'musi mieć wartość 4.' }
    if (-not ($doc['kind'] -is [string]) -or $doc['kind'] -cne 'std.clients') { Add-ConfigError $Errors 'kind' "musi mieć wartość 'std.clients'." }
    if (-not (Test-ConfigIsInteger $doc['revision']) -or [double]$doc['revision'] -lt 1 -or [double]$doc['revision'] -gt [int]::MaxValue) {
        Add-ConfigError $Errors 'revision' 'musi być liczbą całkowitą większą od zera.'
    }
    Test-ConfigOptionalDate $doc 'publishedAt' '' $Errors
    Test-ConfigOptionalDate $doc 'expiresAt' '' $Errors
    $clients = $doc['clients']
    if (-not (Test-ConfigIsList $clients)) {
        Add-ConfigError $Errors 'clients' 'musi być listą klientów [ { "id", "name", "release" } ].'
        return
    }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    for ($i = 0; $i -lt $clients.Count; $i++) {
        $entry = $clients[$i]
        $entryPath = Join-ConfigIndexPath 'clients' $i
        if (-not ($entry -is [System.Collections.IDictionary])) {
            Add-ConfigError $Errors $entryPath 'musi być obiektem { "id", "name", "release" }.'
            continue
        }
        $id = $entry['id']
        if (-not ($id -is [string]) -or $id -cnotmatch $script:StdConfigIdPattern -or $id -eq $script:StdConfigIndexRevisionKey) {
            Add-ConfigError $Errors (Join-ConfigPath $entryPath 'id') ('nieprawidłowy identyfikator klienta {0}.' -f (Format-ConfigValue $id))
        } elseif (-not $seen.Add($id)) {
            Add-ConfigError $Errors (Join-ConfigPath $entryPath 'id') ("identyfikator klienta '{0}' się powtarza." -f $id)
        }
        if (-not (Test-ConfigIsText $entry['name'])) { Add-ConfigError $Errors (Join-ConfigPath $entryPath 'name') 'brak nazwy klienta.' }
        if (-not (Test-ConfigIsText $entry['release'])) { Add-ConfigError $Errors (Join-ConfigPath $entryPath 'release') 'brak ścieżki konfiguracji klienta (np. clients/lizard/release.sdt).' }
    }
}

# --- Kontrole przy pobieraniu ----------------------------------------------------------------

function Assert-ConfigMinClientVersion {
    param($Document)
    if (-not (Test-ConfigHasValue $Document 'minClientVersion')) { return }
    $required = ConvertTo-ConfigVersion $Document['minClientVersion']
    if ($null -eq $required) {
        throw "[STD-CFG-INVALID] Pole minClientVersion ma nieprawidłową wartość $(Format-ConfigValue $Document['minClientVersion'])."
    }
    $current = ConvertTo-ConfigVersion (Get-StdVersion)
    if ($required -gt $current) {
        throw "[STD-CFG-CLIENT-OLD] Ta konfiguracja wymaga STD w wersji co najmniej $($Document['minClientVersion']), a uruchomiona jest wersja $(Get-StdVersion). Zaktualizuj STD."
    }
}

function Assert-ConfigNotExpired {
    # expiresAt + 5 minut >= teraz (UTC). Bez expiresAt: data publikacji + 90 dni (domyślna ważność, SPEC 2.6).
    param($Document)
    $expires = $null
    if (Test-ConfigHasValue $Document 'expiresAt') {
        $expires = ConvertTo-ConfigDate $Document['expiresAt']
        if ($null -eq $expires) { throw "[STD-CFG-INVALID] Pole expiresAt ma nieprawidłową wartość $(Format-ConfigValue $Document['expiresAt'])." }
    } elseif (Test-ConfigHasValue $Document 'publishedAt') {
        $published = ConvertTo-ConfigDate $Document['publishedAt']
        if ($null -eq $published) { throw "[STD-CFG-INVALID] Pole publishedAt ma nieprawidłową wartość $(Format-ConfigValue $Document['publishedAt'])." }
        $expires = $published.AddDays($script:StdConfigDefaultValidityDays)
    }
    if ($null -eq $expires) { return }
    $now = [System.DateTimeOffset]::UtcNow
    if ($expires.AddMinutes($script:StdConfigExpiryGraceMinutes) -lt $now) {
        $when = $expires.ToUniversalTime().ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        throw "[STD-CFG-EXPIRED] Ważność konfiguracji skończyła się $when UTC. Administrator musi opublikować ją ponownie."
    }
}

function Assert-ConfigRevision {
    param($Document, [int]$SeenRevision)
    $revision = 0
    if (Test-ConfigIsInteger $Document['revision']) { $revision = [int][double]$Document['revision'] }
    if ($revision -lt $SeenRevision) {
        throw "[STD-CFG-ROLLBACK] Ta kopia ma rewizję $revision, a ten komputer używał już rewizji $SeenRevision. Starsza kopia została odrzucona (ochrona przed podmianą na starszą wersję)."
    }
    return $revision
}

function Assert-ConfigReleaseChecks {
    # Kontrole 1-6 z SPEC 3.4 dla pobranej konfiguracji. Zwraca numer rewizji.
    param([Parameter(Mandatory = $true)]$Document, [Parameter(Mandatory = $true)][string]$ClientId, [int]$SeenRevision, [bool]$Offline)
    # 1. Rodzaj, wersja formatu, klient.
    if (-not ($Document['kind'] -is [string]) -or $Document['kind'] -cne 'std.release') {
        throw "[STD-CFG-INVALID] Pobrany plik nie jest konfiguracją klienta (kind = $(Format-ConfigValue $Document['kind']) zamiast 'std.release')."
    }
    if (-not (Test-ConfigIsInteger $Document['schemaVersion']) -or [double]$Document['schemaVersion'] -ne 4) {
        throw "[STD-CFG-INVALID] Nieobsługiwana wersja formatu konfiguracji (schemaVersion = $(Format-ConfigValue $Document['schemaVersion'])); STD 4 czyta tylko schemaVersion 4."
    }
    $documentClientId = $null
    if ($Document['client'] -is [System.Collections.IDictionary]) { $documentClientId = $Document['client']['id'] }
    if (-not ($documentClientId -is [string]) -or $documentClientId -cne $ClientId) {
        throw "[STD-CFG-INVALID] To konfiguracja klienta $(Format-ConfigValue $documentClientId), a pobierana była konfiguracja klienta '$ClientId'."
    }
    # 2. Minimalna wersja STD.
    Assert-ConfigMinClientVersion -Document $Document
    # 3. Ważność.
    Assert-ConfigNotExpired -Document $Document
    # 4. Ochrona przed starszą rewizją.
    $revision = Assert-ConfigRevision -Document $Document -SeenRevision $SeenRevision
    # 5. Kopia z cache / pendrive'a tylko przy allowOffline = true.
    if ($Offline -and -not (($Document['allowOffline'] -is [bool]) -and [bool]$Document['allowOffline'])) {
        throw '[STD-CFG-OFFLINE] Konfiguracja tego klienta nie pozwala na pracę bez serwera (allowOffline = false) - kopii z cache ani z pendrive''a nie można użyć. Połącz komputer z siecią.'
    }
    # 6. Reguły Test-StdRelease.
    $errors = @(Test-StdRelease -Release $Document)
    if ($errors.Count -gt 0) {
        $shown = @($errors | Select-Object -First $script:StdConfigMaxReportedErrors)
        $message = '[STD-CFG-INVALID] Konfiguracja jest nieprawidłowa: ' + ($shown -join '; ')
        if ($errors.Count -gt $shown.Count) { $message += " (i jeszcze $($errors.Count - $shown.Count))" }
        throw $message
    }
    return $revision
}

function Assert-ConfigIndexChecks {
    # Kontrole indeksu klientów: rodzaj i wersja formatu, minimalna wersja STD, ważność (jeśli podana),
    # rewizja (ochrona przed starszą wersją), budowa. Kopia offline jest dozwolona zawsze. Zwraca rewizję.
    param([Parameter(Mandatory = $true)]$Document, [int]$SeenRevision)
    if (-not ($Document['kind'] -is [string]) -or $Document['kind'] -cne 'std.clients') {
        throw "[STD-CFG-INVALID] Pobrany plik nie jest listą klientów (kind = $(Format-ConfigValue $Document['kind']) zamiast 'std.clients')."
    }
    if (-not (Test-ConfigIsInteger $Document['schemaVersion']) -or [double]$Document['schemaVersion'] -ne 4) {
        throw "[STD-CFG-INVALID] Nieobsługiwana wersja formatu listy klientów (schemaVersion = $(Format-ConfigValue $Document['schemaVersion']))."
    }
    Assert-ConfigMinClientVersion -Document $Document
    if (Test-ConfigHasValue $Document 'expiresAt') { Assert-ConfigNotExpired -Document $Document }
    $revision = Assert-ConfigRevision -Document $Document -SeenRevision $SeenRevision
    $errors = New-Object 'System.Collections.Generic.List[string]'
    Test-ConfigClientIndexDocument -Document $Document -Errors $errors
    if ($errors.Count -gt 0) {
        $shown = @($errors | Select-Object -First $script:StdConfigMaxReportedErrors)
        throw ('[STD-CFG-INVALID] Lista klientów jest nieprawidłowa: ' + ($shown -join '; '))
    }
    return $revision
}

function Invoke-ConfigSignedFetch {
    # Wspólna ścieżka Get-StdRelease i Get-StdClientIndex: źródła po kolei, pierwsze poprawne wygrywa.
    param(
        [Parameter(Mandatory = $true)][ValidateSet('release', 'index')][string]$Kind,
        [string]$ClientId,
        [AllowNull()]$Bootstrap,
        [AllowNull()]$Session,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()]$Trusted,
        [AllowNull()][AllowEmptyString()][string]$UsbRoot,
        [Parameter(Mandatory = $true)][string]$FunctionName
    )
    $cacheRoot = Get-StdPath -Kind Cache
    if ($Kind -eq 'release') {
        $relative = "clients/$ClientId/release.sdt"
        $usbRelative = "offline/$ClientId/release.sdt"
        $cachePath = Join-StdLocation -Base $cacheRoot -Child "$ClientId/release.sdt"
        $revisionKey = $ClientId
        $what = "konfiguracji klienta '$ClientId'"
    } else {
        $relative = 'clients.sdt'
        $usbRelative = 'offline/clients.sdt'
        $cachePath = Join-StdLocation -Base $cacheRoot -Child 'clients.sdt'
        $revisionKey = $script:StdConfigIndexRevisionKey
        $what = 'listy klientów'
    }

    $bootstrapDoc = $Bootstrap
    if ($null -ne $bootstrapDoc -and -not ($bootstrapDoc -is [System.Collections.IDictionary])) { $bootstrapDoc = ConvertTo-StdHashtable -InputObject $bootstrapDoc }
    $candidates = New-Object 'System.Collections.Generic.List[object]'
    if ($bootstrapDoc -is [System.Collections.IDictionary]) {
        if (Test-ConfigIsText $bootstrapDoc['server']) {
            $serverBase = ([string]$bootstrapDoc['server']).Trim()
            $candidates.Add(@{ Source = 'server'; Location = (Join-StdLocation -Base $serverBase -Child $relative); UseSession = $true })
        }
        foreach ($mirror in @($bootstrapDoc['mirrors'])) {
            if (Test-ConfigIsText $mirror) {
                $candidates.Add(@{ Source = 'mirror'; Location = (Join-StdLocation -Base $mirror.Trim() -Child $relative); UseSession = $false })
            }
        }
    }
    $candidates.Add(@{ Source = 'cache'; Location = $cachePath; UseSession = $false })
    if (-not [string]::IsNullOrWhiteSpace($UsbRoot)) {
        $usbBase = Resolve-CoreFullPath -Path $UsbRoot
        $candidates.Add(@{ Source = 'usb'; Location = (Join-StdLocation -Base $usbBase -Child $usbRelative); UseSession = $false })
    }

    $seen = Get-StdSeenRevision -Key $revisionKey
    $trustedArray = $Trusted.ToArray()
    $attempts = New-Object 'System.Collections.Generic.List[object]'
    foreach ($candidate in $candidates) {
        $source = [string]$candidate.Source
        $location = [string]$candidate.Location
        $safeLocation = Get-HttpSafeLocation -Location $location
        try {
            if ($candidate.UseSession) {
                $bytes = Get-StdBytes -Location $location -Session $Session
            } else {
                $bytes = Get-StdBytes -Location $location
            }
            $envelope = Read-StdSignedEnvelope -Bytes $bytes -TrustedCertificates $trustedArray
            $document = $null
            try {
                $document = ConvertFrom-StdJson -Json $envelope.PayloadText
            } catch {
                throw '[STD-CFG-INVALID] Treść podpisanego pliku nie jest poprawnym dokumentem JSON.'
            }
            if (-not ($document -is [System.Collections.IDictionary])) {
                throw '[STD-CFG-INVALID] Treść podpisanego pliku nie jest obiektem JSON.'
            }
            $offline = ($source -eq 'cache' -or $source -eq 'usb')
            if ($Kind -eq 'release') {
                $revision = Assert-ConfigReleaseChecks -Document $document -ClientId $ClientId -SeenRevision $seen -Offline $offline
            } else {
                $revision = Assert-ConfigIndexChecks -Document $document -SeenRevision $seen
            }
        } catch {
            $errorText = Format-CoreErrorText $_ 1500
            $attempts.Add(@{ Source = $source; Location = $safeLocation; Error = $errorText })
            Write-StdLog -Level Warn -Source $FunctionName -Message ("Odrzucono źródło {0} ({1}): {2}" -f (Get-ConfigSourceLabel $source), $safeLocation, $errorText)
            continue
        }

        $attempts.Add(@{ Source = $source; Location = $safeLocation; Error = $null })
        if ($source -ne 'cache') {
            try {
                Save-ConfigCacheFile -Path $cachePath -Bytes $bytes
            } catch {
                Write-StdLog -Level Warn -Source $FunctionName -Message "Nie udało się zapisać kopii w cache ($cachePath): $(Format-CoreErrorText $_ 200)"
            }
        }
        try {
            Set-StdSeenRevision -Key $revisionKey -Revision $revision
        } catch {
            Write-StdLog -Level Warn -Source $FunctionName -Message "Nie udało się zapisać widzianej rewizji: $(Format-CoreErrorText $_ 200)"
        }
        Write-StdLog -Level Info -Source $FunctionName -Message ("Wczytano {0} (rewizja {1}) - źródło: {2} ({3}), podpis {4}." -f $what, $revision, (Get-ConfigSourceLabel $source), $safeLocation, $envelope.Kid)
        return @{
            Document = $document
            Source   = $source
            Revision = $revision
            Kid      = $envelope.Kid
            Attempts = $attempts.ToArray()
        }
    }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($attempt in $attempts) {
        $lines.Add(('- {0} ({1}): {2}' -f (Get-ConfigSourceLabel $attempt.Source), $attempt.Location, $attempt.Error))
    }
    $message = "[STD-CFG-UNAVAILABLE] Nie udało się wczytać $what z żadnego źródła. Próby:`n" + ($lines -join "`n")
    Write-StdLog -Level Error -Source $FunctionName -Message $message
    $exception = New-Object System.InvalidOperationException($message)
    $record = [System.Management.Automation.ErrorRecord]::new($exception, 'STD-CFG-UNAVAILABLE', [System.Management.Automation.ErrorCategory]::ResourceUnavailable, $attempts.ToArray())
    throw $record
}

function Merge-ConfigNode {
    # Łączenie po skopiowaniu wejść (ConvertTo-StdHashtable): słowniki rekurencyjnie, reszta zastępowana,
    # $null w Override = zostaje Base.
    param([AllowNull()]$Base, [AllowNull()]$Override)
    if ($null -eq $Override) {
        if ($Base -is [array]) { return , $Base }
        return $Base
    }
    if (($Base -is [System.Collections.IDictionary]) -and ($Override -is [System.Collections.IDictionary])) {
        $result = [ordered]@{}
        foreach ($entry in @($Base.GetEnumerator())) { $result.Add($entry.Key, $entry.Value) }
        foreach ($entry in @($Override.GetEnumerator())) {
            if ($null -eq $entry.Value) { continue }
            if ($result.Contains($entry.Key)) {
                $merged = Merge-ConfigNode -Base $result[$entry.Key] -Override $entry.Value
                $result[$entry.Key] = $merged
            } else {
                $result.Add($entry.Key, $entry.Value)
            }
        }
        return $result
    }
    if ($Override -is [array]) { return , $Override }
    return $Override
}

function Get-ConfigListCopy {
    # Kopia listy jako object[] (pusta, gdy brak); pojedyncza wartość -> lista z jednym elementem.
    param([AllowNull()]$Value)
    $list = New-Object 'System.Collections.Generic.List[object]'
    if ($null -eq $Value) { return , $list.ToArray() }
    if (Test-ConfigIsList $Value) {
        foreach ($item in $Value) { $list.Add($item) }
    } else {
        $list.Add($Value)
    }
    return , $list.ToArray()
}

# ---------------------------------------------------------------------------------------------
# Funkcje publiczne
# ---------------------------------------------------------------------------------------------

function Get-StdTrustedCertificates {
    <#
    .SYNOPSIS
    Zaufane certyfikaty do sprawdzania podpisu konfiguracji (X509Certificate2) z Modules\Std.TrustedSigners.psd1.
    Puste wpisy są pomijane; wpis z błędnym certyfikatem albo niezgodnym odciskiem - pomijany z ostrzeżeniem w logu.
    Wynik: kolejne certyfikaty (pusta lista = brak wyniku). Nie ma innego miejsca, które dodaje zaufany certyfikat.
    #>
    [CmdletBinding()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param()
    $list = Read-ConfigTrustedSigners -Path (Get-ConfigTrustedSignersPath)
    foreach ($certificate in $list) { $certificate }
}

function Get-StdBootstrap {
    <#
    .SYNOPSIS
    Plik startowy STD (SPEC 2.3): adres serwera, kopie lustrzane, sposób logowania, proxy.
    Kolejność: -Server, <ScriptRoot>\bootstrap.json (pendrive, obok Start-STD.cmd), Get-StdPath Bootstrap
    (adres zapisany wcześniej na tym komputerze). Zwraca [ordered] @{ server; mirrors; auth; proxy } albo $null.
    Adres serwera musi zaczynać się od https:// (ukośnik na końcu jest dopisywany); błąd -> [STD-CFG-BOOTSTRAP].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Server,
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ScriptRoot
    )
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        return (ConvertTo-ConfigBootstrap -InputObject @{ server = $Server } -Origin 'parametr -Server')
    }

    if (-not [string]::IsNullOrWhiteSpace($ScriptRoot)) {
        $usbPath = Join-Path (Resolve-CoreFullPath -Path $ScriptRoot) 'bootstrap.json'
        if ([System.IO.File]::Exists($usbPath)) {
            try {
                $text = [System.IO.File]::ReadAllText($usbPath, (New-CoreUtf8NoBom))
            } catch {
                throw "[STD-CFG-BOOTSTRAP] Nie udało się odczytać pliku $usbPath - $(Format-CoreErrorText $_ 200)"
            }
            try {
                $doc = ConvertFrom-StdJson -Json $text
            } catch {
                throw "[STD-CFG-BOOTSTRAP] Plik $usbPath nie jest poprawnym JSON - popraw go albo podaj adres serwera ręcznie."
            }
            $result = ConvertTo-ConfigBootstrap -InputObject $doc -Origin "plik $usbPath"
            Write-StdLog -Level Debug -Source 'Get-StdBootstrap' -Message "Plik startowy: $usbPath"
            return $result
        }
    }

    $localPath = Get-StdPath -Kind Bootstrap
    $local = Read-StdJsonFile -Path $localPath
    if ($null -ne $local) {
        $result = ConvertTo-ConfigBootstrap -InputObject $local -Origin "plik $localPath"
        Write-StdLog -Level Debug -Source 'Get-StdBootstrap' -Message "Plik startowy: $localPath"
        return $result
    }
    return $null
}

function Save-StdBootstrapLocal {
    <#
    .SYNOPSIS
    Zapisuje plik startowy (adres serwera wpisany przez technika) do Get-StdPath Bootstrap, żeby przy
    następnym uruchomieniu i po aktualizacji STD nie trzeba było go wpisywać ponownie.
    Zapisywane są tylko pola server, mirrors, auth.mode i proxy - nigdy loginy ani hasła.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [object]$Bootstrap
    )
    $path = Get-StdPath -Kind Bootstrap
    $normalized = ConvertTo-ConfigBootstrap -InputObject $Bootstrap -Origin 'Save-StdBootstrapLocal'
    Write-StdJsonFile -Path $path -InputObject $normalized
    Write-StdLog -Level Info -Source 'Save-StdBootstrapLocal' -Message "Zapisano adres serwera STD $($normalized.server) w $path."
}

function Get-StdSeenRevision {
    <#
    .SYNOPSIS
    Najwyższa rewizja widziana na tym komputerze dla klucza (identyfikator klienta albo 'clients' dla listy
    klientów) z revisions.json. Brak wpisu -> 0.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Key
    )
    $data = Read-StdJsonFile -Path (Get-StdPath -Kind Revisions)
    if (-not ($data -is [System.Collections.IDictionary]) -or -not $data.Contains($Key)) { return 0 }
    $value = $data[$Key]
    if (-not (Test-ConfigIsInteger $value) -or [double]$value -lt 0 -or [double]$value -gt [int]::MaxValue) { return 0 }
    return [int][double]$value
}

function Set-StdSeenRevision {
    <#
    .SYNOPSIS
    Zapisuje widzianą rewizję w revisions.json. Nigdy nie zmniejsza zapisanej wartości.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateRange(0, 2147483647)]
        [int]$Revision
    )
    $path = Get-StdPath -Kind Revisions
    $data = Read-StdJsonFile -Path $path
    if (-not ($data -is [System.Collections.IDictionary])) { $data = [ordered]@{} }
    $current = 0
    if ($data.Contains($Key) -and (Test-ConfigIsInteger $data[$Key])) { $current = [double]$data[$Key] }
    if ($Revision -le $current) { return }
    if ($data.Contains($Key)) {
        $data[$Key] = $Revision
    } else {
        $data.Add($Key, $Revision)
    }
    Write-StdJsonFile -Path $path -InputObject $data
    Write-StdLog -Level Debug -Source 'Set-StdSeenRevision' -Message "Widziana rewizja '$Key': $Revision."
}

function Get-StdClientIndex {
    <#
    .SYNOPSIS
    Podpisany indeks klientów (clients.sdt): serwer, kopie lustrzane, cache (Cache\clients.sdt),
    pendrive (<UsbRoot>\offline\clients.sdt). Kopia offline jest dozwolona zawsze. Klucz rewizji: 'clients'.
    Wynik @{ Index; Source; Revision; Attempts }. Brak poprawnego źródła -> [STD-CFG-UNAVAILABLE] z listą prób.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [object]$Bootstrap,
        [object]$Session,
        [AllowNull()]
        [AllowEmptyCollection()]
        [System.Security.Cryptography.X509Certificates.X509Certificate2[]]$TrustedCertificates,
        [AllowNull()]
        [AllowEmptyString()]
        [string]$UsbRoot
    )
    $trusted = Resolve-ConfigTrustedList -Given $PSBoundParameters.ContainsKey('TrustedCertificates') -TrustedCertificates $TrustedCertificates
    if ($trusted.Count -eq 0) {
        throw '[STD-SIG-NOTRUST] Brak zaufanych certyfikatów do sprawdzenia podpisu listy klientów. Utwórz certyfikat skryptem Tools\New-StdSigningCertificate.ps1 i dopisz wypisany wpis do Modules\Std.TrustedSigners.psd1.'
    }
    $result = Invoke-ConfigSignedFetch -Kind index -Bootstrap $Bootstrap -Session $Session -Trusted $trusted -UsbRoot $UsbRoot -FunctionName 'Get-StdClientIndex'
    return @{
        Index    = $result.Document
        Source   = $result.Source
        Revision = $result.Revision
        Attempts = $result.Attempts
    }
}

function Get-StdRelease {
    <#
    .SYNOPSIS
    Podpisana konfiguracja klienta. Źródła po kolei: serwer (<server>clients/<id>/release.sdt), kopie lustrzane,
    cache (Cache\<id>\release.sdt), pendrive (<UsbRoot>\offline\<id>\release.sdt). Każde źródło: pobranie ->
    sprawdzenie podpisu -> kontrole (rodzaj i klient, minClientVersion, expiresAt, rewizja, allowOffline dla
    cache i pendrive'a, Test-StdRelease). Pierwsze poprawne wygrywa, trafia do cache i do revisions.json.
    Wynik @{ Release; Source; Revision; Kid; Attempts } (Attempts: każda próba @{ Source; Location; Error },
    także udana - z Error = $null). Brak poprawnego źródła -> [STD-CFG-UNAVAILABLE] z listą prób
    (lista jest też w TargetObject błędu).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$ClientId,
        [Parameter(Mandatory = $true, Position = 1)]
        [AllowNull()]
        [object]$Bootstrap,
        [object]$Session,
        [AllowNull()]
        [AllowEmptyCollection()]
        [System.Security.Cryptography.X509Certificates.X509Certificate2[]]$TrustedCertificates,
        [AllowNull()]
        [AllowEmptyString()]
        [string]$UsbRoot
    )
    Assert-ConfigClientId -ClientId $ClientId
    if ($ClientId -eq $script:StdConfigIndexRevisionKey) {
        throw "[STD-CFG-ARGS] Identyfikator '$($script:StdConfigIndexRevisionKey)' jest zarezerwowany dla listy klientów."
    }
    $trusted = Resolve-ConfigTrustedList -Given $PSBoundParameters.ContainsKey('TrustedCertificates') -TrustedCertificates $TrustedCertificates
    if ($trusted.Count -eq 0) {
        throw '[STD-SIG-NOTRUST] Brak zaufanych certyfikatów do sprawdzenia podpisu konfiguracji. Utwórz certyfikat skryptem Tools\New-StdSigningCertificate.ps1 i dopisz wypisany wpis do Modules\Std.TrustedSigners.psd1.'
    }
    $result = Invoke-ConfigSignedFetch -Kind release -ClientId $ClientId -Bootstrap $Bootstrap -Session $Session -Trusted $trusted -UsbRoot $UsbRoot -FunctionName 'Get-StdRelease'
    return @{
        Release  = $result.Document
        Source   = $result.Source
        Revision = $result.Revision
        Kid      = $result.Kid
        Attempts = $result.Attempts
    }
}

function Test-StdRelease {
    <#
    .SYNOPSIS
    Sprawdza konfigurację klienta (SPEC 2.6-2.8 i reguły 3.4). Wynik: kolejne komunikaty o błędach (tekst
    "pole: opis"); brak wyniku = konfiguracja poprawna. Używaj: $errors = @(Test-StdRelease -Release $doc).
    -Unpublished: nie wymaga pól ustawianych przy publikacji (revision, publishedAt, publishedBy).
    Odrzuca m.in. dane logowania w domainJoin / localAdmin, jawne hasła i tokeny oraz wstawki
    {{secret:nazwa}} bez sekretu w secrets.items.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [object]$Release,
        [switch]$Unpublished
    )
    $errors = New-Object 'System.Collections.Generic.List[string]'
    $doc = $null
    if ($null -ne $Release) { $doc = ConvertTo-StdHashtable -InputObject $Release }
    if (-not ($doc -is [System.Collections.IDictionary])) {
        Add-ConfigError $errors '' 'konfiguracja musi być obiektem JSON.'
    } else {
        Test-ConfigReleaseDocument -Document $doc -Unpublished ([bool]$Unpublished) -Errors $errors
    }
    foreach ($message in $errors) { $message }
}

function Merge-StdConfig {
    <#
    .SYNOPSIS
    Łączy dwa obiekty ustawień i zwraca nowy obiekt ([ordered]); wejścia nie są zmieniane.
    Słowniki łączone rekurencyjnie, tablice i wartości proste zastępowane, $null w -Override = zostaje -Base.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [object]$Base,
        [Parameter(Mandatory = $true, Position = 1)]
        [AllowNull()]
        [object]$Override
    )
    $baseCopy = $null
    $overrideCopy = $null
    if ($null -ne $Base) { $baseCopy = ConvertTo-StdHashtable -InputObject $Base }
    if ($null -ne $Override) { $overrideCopy = ConvertTo-StdHashtable -InputObject $Override }
    $result = Merge-ConfigNode -Base $baseCopy -Override $overrideCopy
    if ($result -is [array]) { return , $result }
    return $result
}

function Resolve-StdEffectiveConfig {
    <#
    .SYNOPSIS
    Konfiguracja wynikowa dla lokalizacji i profilu: @{ ClientId; ClientName; Revision; Sources; Catalog;
    Site (bez profiles); ProfileName; Apps; Tasks; Bloatware; Uninstall; SystemSettings; Scripts; Locked }.
    Dziedziczenie defaults -> lokalizacja -> profil (Merge-StdConfig); Scripts to defaults.scripts + scripts
    lokalizacji + scripts profilu (doklejane). Wynik jest kopią - zmiany nie dotykają -Release.
    Nieznana lokalizacja -> [STD-CFG-SITE], nieznany profil -> [STD-CFG-PROFILE].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [object]$Release,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$SiteId,
        [Parameter(Mandatory = $true, Position = 2)]
        [string]$ProfileName
    )
    $doc = ConvertTo-StdHashtable -InputObject $Release
    if (-not ($doc -is [System.Collections.IDictionary])) {
        throw '[STD-CFG-INVALID] Konfiguracja musi być obiektem (wynik Get-StdRelease).Release.'
    }
    $client = $doc['client']
    $clientId = $null
    $clientName = $null
    if ($client -is [System.Collections.IDictionary]) {
        $clientId = $client['id']
        $clientName = $client['name']
    }

    $site = $null
    $siteIds = New-Object 'System.Collections.Generic.List[string]'
    foreach ($candidate in @($doc['sites'])) {
        if (-not ($candidate -is [System.Collections.IDictionary])) { continue }
        $siteIds.Add([string]$candidate['id'])
        if ($null -eq $site -and [string]$candidate['id'] -eq $SiteId) { $site = $candidate }
    }
    if ($null -eq $site) {
        throw ("[STD-CFG-SITE] Konfiguracja klienta {0} nie ma lokalizacji {1}. Dostępne: {2}." -f (Format-ConfigValue $clientId), (Format-ConfigValue $SiteId), ($siteIds -join ', '))
    }

    $profileNode = $null
    $profileKey = $null
    $profileNames = New-Object 'System.Collections.Generic.List[string]'
    $profiles = $site['profiles']
    if ($profiles -is [System.Collections.IDictionary]) {
        foreach ($entry in @($profiles.GetEnumerator())) {
            $profileNames.Add([string]$entry.Key)
            if ($null -eq $profileKey -and [string]$entry.Key -eq $ProfileName) {
                $profileKey = [string]$entry.Key
                $profileNode = $entry.Value
            }
        }
    }
    if ($null -eq $profileKey) {
        throw ("[STD-CFG-PROFILE] Lokalizacja {0} nie ma profilu {1}. Dostępne: {2}." -f (Format-ConfigValue $SiteId), (Format-ConfigValue $ProfileName), ($profileNames -join ', '))
    }
    if (-not ($profileNode -is [System.Collections.IDictionary])) { $profileNode = [ordered]@{} }

    $defaults = $doc['defaults']
    if (-not ($defaults -is [System.Collections.IDictionary])) { $defaults = [ordered]@{} }

    $siteCopy = [ordered]@{}
    foreach ($entry in @($site.GetEnumerator())) {
        if ([string]$entry.Key -eq 'profiles') { continue }
        $siteCopy.Add($entry.Key, $entry.Value)
    }

    $tasks = Merge-ConfigNode -Base (Merge-ConfigNode -Base $defaults['tasks'] -Override $site['tasks']) -Override $profileNode['tasks']
    if (-not ($tasks -is [System.Collections.IDictionary])) { $tasks = [ordered]@{} }
    $systemSettings = Merge-ConfigNode -Base (Merge-ConfigNode -Base $defaults['systemSettings'] -Override $site['systemSettings']) -Override $profileNode['systemSettings']
    if (-not ($systemSettings -is [System.Collections.IDictionary])) { $systemSettings = [ordered]@{} }
    $bloatware = Merge-ConfigNode -Base (Merge-ConfigNode -Base $defaults['bloatware'] -Override $site['bloatware']) -Override $profileNode['bloatware']
    $uninstall = Merge-ConfigNode -Base (Merge-ConfigNode -Base $defaults['uninstall'] -Override $site['uninstall']) -Override $profileNode['uninstall']

    $scripts = New-Object 'System.Collections.Generic.List[object]'
    foreach ($part in @($defaults['scripts'], $site['scripts'], $profileNode['scripts'])) {
        foreach ($item in (Get-ConfigListCopy $part)) { $scripts.Add($item) }
    }

    $sources = $doc['sources']
    if (-not ($sources -is [System.Collections.IDictionary])) { $sources = [ordered]@{} }
    $catalog = $doc['catalog']
    if (-not ($catalog -is [System.Collections.IDictionary])) { $catalog = [ordered]@{} }
    $revision = $null
    if (Test-ConfigIsInteger $doc['revision']) { $revision = [int][double]$doc['revision'] }

    return @{
        ClientId       = $clientId
        ClientName     = $clientName
        Revision       = $revision
        Sources        = $sources
        Catalog        = $catalog
        Site           = $siteCopy
        ProfileName    = $profileKey
        Apps           = (Get-ConfigListCopy $profileNode['apps'])
        Tasks          = $tasks
        Bloatware      = (Get-ConfigListCopy $bloatware)
        Uninstall      = (Get-ConfigListCopy $uninstall)
        SystemSettings = $systemSettings
        Scripts        = $scripts.ToArray()
        Locked         = (Get-ConfigListCopy $profileNode['locked'])
    }
}
