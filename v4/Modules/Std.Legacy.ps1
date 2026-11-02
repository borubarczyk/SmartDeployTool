# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.Legacy.ps1 - konwersja konfiguracji STD 3.1 (config.json) do konfiguracji klienta STD 4 (SPEC 3.5).
#
# Wymaga Std.Core.ps1 (ConvertFrom-StdJson, ConvertTo-StdHashtable, Write-StdLog) - STD.psm1 wczytuje go wcześniej.
#
# Wynik konwersji to dokument v4 bez podpisu i bez revision / publishedAt: jedna lokalizacja "default"
# ("Domyślna"), katalog programów, profile i zadania. Logika wyboru źródła programu jest taka sama jak
# Resolve-AppInstallSource, Install-TeamViewer i Install-AV w STD 3.1.
#
# Bezpieczeństwo:
# - żadna wartość hasła ani tokenu z 3.1 nie trafia do wyniku (dokumentu, ostrzeżeń ani listy sekretów):
#   WebAuth, AntyVirus.Credentials, DomainJoin.Username i hasła są pomijane, a APITOKEN / CUSTOMCONFIGID
#   w argumentach zamieniamy na {{secret:...}};
# - na końcu sprawdzamy cały wynik: jeśli wartość hasła lub tokenu z 3.1 trafiła gdzieś indziej (np. ktoś
#   wpisał ją w argumenty programu), zastępujemy ją znakami *** i dodajemy ostrzeżenie;
# - ostrzeżenia zawierają tylko nazwy pól, programów i ścieżek - nigdy wartości haseł.
#
# Funkcje pomocnicze bez "-Std" w nazwie nie są eksportowane z modułu (Export-ModuleMember *-Std*).

# ---------------------------------------------------------------------------------------------
# Stałe
# ---------------------------------------------------------------------------------------------

$script:StdLegacyClientIdPattern = '^[a-z0-9][a-z0-9-]{1,39}$'
$script:StdLegacyCatalogKeyPattern = '^[A-Za-z0-9][A-Za-z0-9._+-]{0,63}$'
$script:StdLegacyCatalogKeyMaxLength = 64
$script:StdLegacySiteId = 'default'
$script:StdLegacySiteName = 'Domyślna'
$script:StdLegacyDefaultProfile = 'Domyślny'
$script:StdLegacyTeamViewerKey = 'TeamViewer'
$script:StdLegacyAntivirusKey = 'AV'
$script:StdLegacyRegistryTypes = @('String', 'ExpandString', 'DWord', 'QWord', 'MultiString', 'Binary')

# Lista wbudowana w Remove-Bloatware w STD 3.1 (tam nie była częścią config.json). W STD 4 wzorce są
# w konfiguracji (defaults.bloatware), więc przenosimy je, żeby zadanie RemoveBloatware działało tak samo.
$script:StdLegacyBloatware = @(
    '*BingNews*', '*BingWeather*', '*MicrosoftOfficeHub*', '*SkypeApp*',
    '*SolitaireCollection*', '*XboxApp*', '*XboxGamingOverlay*', '*XboxSpeechToTextOverlay*',
    '*YourPhone*', '*TikTok*', '*Spotify*', '*Facebook*', '*Instagram*', '*Twitter*',
    '*LinkedIn*', '*Netflix*', '*PandoraMedia*', '*CandyCrush*', '*Disney*'
)

# DefaultCheckboxes (3.1) -> defaults.tasks (SPEC 2.8).
$script:StdLegacyTaskMap = [ordered]@{
    WaitForNetwork        = 'WaitForNetwork'
    SuspendHibernation    = 'SuspendSleep'
    ImportWiFiProfile     = 'ConnectWiFi'
    RemoveBloatware       = 'RemoveBloatware'
    UninstallMicrosoft365 = 'UninstallM365'
    UninstallOneDrive     = 'UninstallOneDrive'
    InstallApplications   = 'InstallApps'
    CreateLocalAdmin      = 'CreateLocalAdmin'
    ChangeComputerName    = 'RenameComputer'
    JoinDomain            = 'JoinDomain'
    ChangeSystemSettings  = 'SystemSettings'
    RunPostInstallScripts = 'RunScripts'
    ExportHardwareAudit   = 'HardwareAudit'
    EnableBitLocker       = 'EnableBitLocker'
    AutoReboot            = 'RebootAtEnd'
}

# Kolejność zadań w dokumencie (SPEC 2.8).
$script:StdLegacyTaskOrder = @(
    'WaitForNetwork', 'SuspendSleep', 'ConnectWiFi', 'RemoveBloatware', 'UninstallPreinstalled', 'UninstallM365',
    'UninstallOneDrive', 'InstallApps', 'CreateLocalAdmin', 'RenameComputer', 'JoinDomain', 'SystemSettings',
    'RunScripts', 'HardwareAudit', 'EnableBitLocker', 'RebootAtEnd'
)

# Zadania 3.1, których nie ma w STD 4 (ostrzeżenie tylko przy wartości true).
$script:StdLegacyDroppedTasks = [ordered]@{
    DryRun             = 'Pominięto zadanie DryRun. W STD 4 zastępuje je przycisk „Sprawdź”, który pokazuje plan wdrożenia.'
    CreateRestorePoint = 'Pominięto zadanie CreateRestorePoint. STD 4 nie tworzy punktu przywracania.'
    RunWindowsUpdate   = 'Pominięto zadanie RunWindowsUpdate. STD 4 nie ma jeszcze zadania Windows Update.'
    JoinIntune         = 'Pominięto zadanie JoinIntune. STD 4 nie dołącza komputera do Intune.'
}

# Klucze najwyższego poziomu config.json w STD 3.1.
$script:StdLegacyKnownTopKeys = @(
    'DefaultInstallSource', 'InstallSourcePaths', 'Programs', 'Profiles', 'TeamViewer', 'AntyVirus',
    'DefaultCheckboxes', 'WebAuth', 'DomainJoin', 'LocalAdmin', 'WiFiProfile', 'SystemSettings',
    'PostInstallScripts', 'HardwareAudit', 'AutoUpdate', 'CustomWebDataLocation', 'DarkTheme'
)

# Pola programu w STD 3.1 (okno "Program" w Ustawieniach).
$script:StdLegacyProgramFields = @('Enabled', 'FileName', 'SilentArgs', 'DownloadUrl', 'WingetId', 'WingetArgs')

# Nazwy pól z wartościami, które traktujemy jak hasło lub token (do końcowej kontroli wyniku).
$script:StdLegacySecretKeyRegex = New-Object System.Text.RegularExpressions.Regex(
    '^(?:pwd|pass|apikey|api_key)$|(?:password|passwd|has(?:lo|ło)|token|secret)$',
    ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant))

# Nazwa wartości rejestru, która wygląda na hasło lub token (np. Winlogon\DefaultPassword).
$script:StdLegacySecretNameRegex = New-Object System.Text.RegularExpressions.Regex(
    '(?:pass|pwd|has(?:lo|ło)|token|secret)',
    ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant))

# APITOKEN= / CUSTOMCONFIGID= w argumentach instalatora TeamViewer (wartość w cudzysłowie albo do spacji).
$script:StdLegacyTeamViewerArgRegex = New-Object System.Text.RegularExpressions.Regex(
    '\b(?<prop>APITOKEN|CUSTOMCONFIGID)=(?:"(?<q>[^"]*)"|''(?<s>[^'']*)''|(?<v>[^\s"'']+))',
    ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant))

$script:StdLegacyPlaceholderRegex = New-Object System.Text.RegularExpressions.Regex(
    '(\{\{secret:[^}]*\}\})', [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)

# Wartości haseł krótsze niż 4 znaki pomijamy w końcowej kontroli (za dużo fałszywych trafień).
$script:StdLegacyMinSecretLength = 4

# ---------------------------------------------------------------------------------------------
# Pomocnicze (prywatne)
# ---------------------------------------------------------------------------------------------

function Add-LegacyWarning {
    param($Context, [string]$Message)
    $Context.Warnings.Add($Message)
}

function Add-LegacySecretValue {
    # Wartość hasła / tokenu z 3.1 - po konwersji sprawdzamy, czy nie trafiła do wyniku.
    param($Context, [string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return }
    $trimmed = $Value.Trim()
    # Wartości, które nie są hasłem (np. "UseToken": "true"), zamaskowałyby zwykłe argumenty.
    if (@('true', 'false', 'null', 'none', 'brak') -contains $trimmed) { return }
    foreach ($candidate in @($Value, $trimmed)) {
        if ($candidate.Length -ge $script:StdLegacyMinSecretLength -and -not $Context.InputSecrets.Contains($candidate)) {
            $Context.InputSecrets.Add($candidate)
        }
    }
}

function Get-LegacyKeyList {
    # Klucze słownika (kopia). Bez .Keys: klucz o nazwie "Keys" zasłoniłby właściwość.
    param($Dictionary)
    $list = New-Object 'System.Collections.Generic.List[string]'
    if ($Dictionary -is [System.Collections.IDictionary]) {
        foreach ($key in @($Dictionary.get_Keys())) { $list.Add([string]$key) }
    }
    return , $list.ToArray()
}

function Get-LegacyField {
    # Wartość pola słownika bez względu na wielkość liter albo $null (tablice zostają tablicami).
    param($Dictionary, [string]$Name)
    if ($Dictionary -isnot [System.Collections.IDictionary]) { return $null }
    if (-not $Dictionary.Contains($Name)) { return $null }
    return , $Dictionary[$Name]
}

function Get-LegacyText {
    # Pole tekstowe (przycięte) albo ''. Obiekty i tablice dają ''.
    param($Dictionary, [string]$Name)
    $value = Get-LegacyField -Dictionary $Dictionary -Name $Name
    if ($null -eq $value) { return '' }
    if ($value -is [System.Collections.IDictionary] -or $value -is [System.Collections.IList]) { return '' }
    return ([string]$value).Trim()
}

function ConvertTo-LegacyList {
    # Wartość -> object[]: $null -> pusta, tablica -> jej elementy, pojedyncza wartość -> jeden element.
    param($Value)
    $list = New-Object 'System.Collections.Generic.List[object]'
    if ($null -ne $Value) {
        if ($Value -is [System.Collections.IList]) {
            foreach ($item in $Value) { $list.Add($item) }
        } else {
            $list.Add($Value)
        }
    }
    return , $list.ToArray()
}

function ConvertTo-LegacyBool {
    # true / false (także "true" / "false" i 1 / 0) albo $null, gdy wartość nie jest logiczna.
    param($Value)
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -ieq 'true') { return $true }
        if ($text -ieq 'false') { return $false }
        return $null
    }
    if ($Value -is [int] -or $Value -is [long]) {
        if ($Value -eq 1) { return $true }
        if ($Value -eq 0) { return $false }
    }
    return $null
}

function Test-LegacyHasContent {
    # Czy w wartości jest coś poza pustymi tekstami, false i zerami (sekcja ustawiona, a nie domyślna).
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [string]) { return (-not [string]::IsNullOrWhiteSpace($Value)) }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in (Get-LegacyKeyList -Dictionary $Value)) {
            if (Test-LegacyHasContent -Value $Value[$key]) { return $true }
        }
        return $false
    }
    if ($Value -is [System.Collections.IList]) {
        foreach ($item in $Value) {
            if (Test-LegacyHasContent -Value $item) { return $true }
        }
        return $false
    }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return ($Value -ne 0) }
    return $true
}

function Add-LegacyUnknownFieldWarnings {
    # Ostrzeżenie dla każdego pola, którego STD 3.1 nie miało w tej sekcji (wartości nie wypisujemy).
    param($Dictionary, [string[]]$Known, [string]$Prefix, $Context)
    foreach ($key in (Get-LegacyKeyList -Dictionary $Dictionary)) {
        if ($Known -contains $key) { continue }
        Add-LegacyWarning $Context ('Pominięto pole „{0}.{1}”, którego STD 4 nie obsługuje.' -f $Prefix, $key)
    }
}

function Add-LegacyInputSecrets {
    # Zbiera z wejścia wartości pól o nazwach haseł i tokenów (Password, APITOKEN, ...) - tylko do kontroli wyniku.
    param($Node, $Context)
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($key in (Get-LegacyKeyList -Dictionary $Node)) {
            $value = $Node[$key]
            if ($value -is [string]) {
                if ($script:StdLegacySecretKeyRegex.IsMatch($key)) { Add-LegacySecretValue -Context $Context -Value $value }
            } else {
                Add-LegacyInputSecrets -Node $value -Context $Context
            }
        }
    } elseif ($Node -is [System.Collections.IList]) {
        foreach ($item in $Node) { Add-LegacyInputSecrets -Node $item -Context $Context }
    }
}

function Remove-LegacyUrlCredential {
    # 'https://login:haslo@host/...' -> 'https://host/...'. Login i hasło nigdy nie trafiają do wyniku.
    param([string]$Url, [string]$Field, $Context)
    # Login i hasło to wszystko do ostatniego '@' przed pierwszym '/' lub '?' (hasło z 3.1 bywa wpisane
    # bez kodowania, np. z '#' albo '@'); za '@' musi być nazwa hosta.
    $match = [regex]::Match($Url, '^(?<scheme>[A-Za-z][A-Za-z0-9+.-]*://)(?<info>[^/?]*)@(?<host>(?:\[[0-9A-Fa-f:.]+\]|[A-Za-z0-9._~%-]+)(?::[0-9]*)?)(?<rest>(?:[/?#].*)?)$')
    if (-not $match.Success) { return $Url }
    $info = $match.Groups['info'].Value
    Add-LegacySecretValue -Context $Context -Value $info
    $colon = $info.IndexOf(':')
    if ($colon -ge 0) {
        $password = $info.Substring($colon + 1)
        Add-LegacySecretValue -Context $Context -Value $password
        try { Add-LegacySecretValue -Context $Context -Value ([System.Uri]::UnescapeDataString($password)) } catch { }
    }
    Add-LegacyWarning $Context ('Adres w polu {0} zawierał login i hasło. Usunięto je - STD 4 pobiera pliki kontem technika.' -f $Field)
    return ($match.Groups['scheme'].Value + $match.Groups['host'].Value + $match.Groups['rest'].Value)
}

function Test-LegacyHttpsUrl {
    # Ostrzeżenie, gdy adres nie zaczyna się od https:// (walidacja v4 go odrzuci).
    param([string]$Url, [string]$What, $Context)
    if ($Url -match '^(?i)https://') { return }
    if ($Url -match '^(?i)http://') {
        Add-LegacyWarning $Context ('{0} ma adres http:// ({1}). STD 4 pobiera pliki tylko przez https:// - zmień adres przed publikacją.' -f $What, $Url)
    } else {
        Add-LegacyWarning $Context ('{0} ma adres „{1}”, który nie zaczyna się od https://. Popraw go przed publikacją.' -f $What, $Url)
    }
}

function New-LegacySource {
    # Źródło v4 z wpisu InstallSourcePaths: network -> unc, web -> https.
    param([string]$Kind, [string]$Value, [string]$Field, [string]$Name, $Context)
    if ($Kind -eq 'network') {
        if ($Value -notmatch '^\\\\[^\\]') {
            Add-LegacyWarning $Context ('Ścieżka źródła „{0}” ({1}) nie jest ścieżką UNC (\\serwer\udział\). Sprawdź, czy będzie dostępna na wdrażanym komputerze.' -f $Name, $Field)
        }
        return [ordered]@{ type = 'unc'; path = $Value; auth = 'session' }
    }
    $url = Remove-LegacyUrlCredential -Url $Value -Field $Field -Context $Context
    Test-LegacyHttpsUrl -Url $url -What ('Źródło „{0}” ({1})' -f $Name, $Field) -Context $Context
    return [ordered]@{ type = 'https'; url = $url; auth = 'session' }
}

function Get-LegacySourceKind {
    # DefaultInstallSource -> 'network' / 'web' / 'winget' / '' (brak) / '?' (nieznane).
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $text = $Value.Trim().ToLowerInvariant()
    if ($text -eq 'network' -or $text -eq 'web' -or $text -eq 'winget') { return $text }
    return '?'
}

function ConvertTo-LegacyCatalogKey {
    # Nazwa programu -> klucz katalogu ^[A-Za-z0-9][A-Za-z0-9._+-]{0,63}$ (polskie litery bez ogonków,
    # inne niedozwolone znaki -> '-'). Kolizje rozwiązuje New-LegacyUniqueKey.
    param([string]$Name)
    $text = $Name.Trim()
    # "ł" / "Ł" nie rozkładają się w Unicode (FormD) - zamieniamy je ręcznie.
    $text = $text.Replace([string][char]0x0142, 'l').Replace([string][char]0x0141, 'L')
    $decomposed = $text.Normalize([System.Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    foreach ($ch in $decomposed.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($ch)
        }
    }
    $text = $builder.ToString()
    $text = [regex]::Replace($text, '[^A-Za-z0-9._+-]+', '-')
    $text = [regex]::Replace($text, '-{2,}', '-')
    $text = [regex]::Replace($text, '^[^A-Za-z0-9]+', '')
    if ($text.Length -gt $script:StdLegacyCatalogKeyMaxLength) { $text = $text.Substring(0, $script:StdLegacyCatalogKeyMaxLength) }
    $text = $text.TrimEnd('-')
    if ($text.Length -eq 0) { $text = 'program' }
    return $text
}

function New-LegacyUniqueKey {
    # Klucz unikalny bez względu na wielkość liter: przy kolizji sufiks -2, -3, ...
    param([string]$Base, $Used)
    if ($Used.Add($Base)) { return $Base }
    $number = 2
    while ($true) {
        $suffix = '-' + $number
        $stem = $Base
        if ($stem.Length + $suffix.Length -gt $script:StdLegacyCatalogKeyMaxLength) {
            $stem = $stem.Substring(0, $script:StdLegacyCatalogKeyMaxLength - $suffix.Length).TrimEnd('-')
        }
        $candidate = $stem + $suffix
        if ($Used.Add($candidate)) { return $candidate }
        $number++
    }
}

function New-LegacyUniqueName {
    # Unikalna nazwa profilu: "Nazwa (3.1)", potem "Nazwa (3.1) 2", ...
    param([string]$Base, $Used)
    if ($Used.Add($Base)) { return $Base }
    $candidate = $Base + ' (3.1)'
    $number = 2
    while (-not $Used.Add($candidate)) {
        $candidate = $Base + ' (3.1) ' + $number
        $number++
    }
    return $candidate
}

function Protect-LegacyKnownSecretArgs {
    # APITOKEN=... / CUSTOMCONFIGID=... -> {{secret:teamviewer-apitoken}} / {{secret:teamviewer-customconfigid}}.
    # Wartość trafia tylko do listy kontrolnej (InputSecrets), nazwa sekretu do SecretsToSet.
    param([string]$Arguments, $Context)
    if ([string]::IsNullOrEmpty($Arguments)) { return $Arguments }
    $found = $script:StdLegacyTeamViewerArgRegex.Matches($Arguments)
    if ($found.Count -eq 0) { return $Arguments }
    $builder = New-Object System.Text.StringBuilder
    $position = 0
    foreach ($match in $found) {
        $group = $null
        foreach ($groupName in @('q', 's', 'v')) {
            if ($match.Groups[$groupName].Success) { $group = $match.Groups[$groupName]; break }
        }
        if ($null -eq $group -or $group.Length -eq 0) { continue }
        $value = $group.Value
        if ($value -match '^\{\{secret:[^}]*\}\}$') { continue }
        if ($match.Groups['prop'].Value -ieq 'APITOKEN') { $secretName = 'teamviewer-apitoken' } else { $secretName = 'teamviewer-customconfigid' }
        Add-LegacySecretValue -Context $Context -Value $value
        if (-not $Context.SecretsToSet.Contains($secretName)) { $Context.SecretsToSet.Add($secretName) }
        [void]$builder.Append($Arguments, $position, $group.Index - $position)
        [void]$builder.Append('{{secret:' + $secretName + '}}')
        $position = $group.Index + $group.Length
    }
    [void]$builder.Append($Arguments, $position, $Arguments.Length - $position)
    return $builder.ToString()
}

function ConvertTo-LegacyProgramEntry {
    # Program 3.1 -> pozycja katalogu v4 (bez klucza) albo $null, gdy nie ma czego instalować.
    # Reguły jak Resolve-AppInstallSource w 3.1 (SPEC 3.5): DownloadUrl -> url; WingetId -> winget (+ wingetArgs);
    # FileName -> file + args, ale przy źródle domyślnym winget i braku WingetId -> winget = FileName,
    # wingetArgs = WingetArgs albo SilentArgs.
    param([string]$DisplayName, $Program, [bool]$WingetDefault, $Context)
    $fileName = Get-LegacyText $Program 'FileName'
    $silentArgs = Protect-LegacyKnownSecretArgs -Arguments (Get-LegacyText $Program 'SilentArgs') -Context $Context
    $downloadUrl = Get-LegacyText $Program 'DownloadUrl'
    $wingetId = Get-LegacyText $Program 'WingetId'
    $wingetArgs = Protect-LegacyKnownSecretArgs -Arguments (Get-LegacyText $Program 'WingetArgs') -Context $Context

    $entry = [ordered]@{ name = $DisplayName }
    if ($downloadUrl.Length -gt 0) {
        $field = 'Programs.' + $DisplayName + '.DownloadUrl'
        $downloadUrl = Remove-LegacyUrlCredential -Url $downloadUrl -Field $field -Context $Context
        Test-LegacyHttpsUrl -Url $downloadUrl -What ('Program „{0}” (DownloadUrl)' -f $DisplayName) -Context $Context
        $entry['url'] = $downloadUrl
        if ($silentArgs.Length -gt 0) { $entry['args'] = $silentArgs }
    } elseif ($fileName.Length -gt 0) {
        if ($WingetDefault -and $wingetId.Length -eq 0) {
            # Dawny zapis 3.1: identyfikator winget w FileName.
            $entry['winget'] = $fileName
            if ($wingetArgs.Length -gt 0) { $entry['wingetArgs'] = $wingetArgs }
            elseif ($silentArgs.Length -gt 0) { $entry['wingetArgs'] = $silentArgs }
        } else {
            $entry['file'] = $fileName
            if ($silentArgs.Length -gt 0) { $entry['args'] = $silentArgs }
        }
    }
    if ($wingetId.Length -gt 0) {
        $entry['winget'] = $wingetId
        if ($wingetArgs.Length -gt 0) { $entry['wingetArgs'] = $wingetArgs }
    }
    if (-not ($entry.Contains('file') -or $entry.Contains('url') -or $entry.Contains('winget'))) { return $null }
    return $entry
}

function ConvertTo-LegacyRegistryValue {
    # Wartość DWord / QWord z tekstu ("1", "0x10") -> liczba; inne typy bez zmian. $null -> ''.
    param($Value, [string]$Type, [string]$Label, $Context)
    if ($null -eq $Value) { $Value = '' }
    if ($Type -ne 'DWord' -and $Type -ne 'QWord') { return , $Value }
    if ($Value -is [int] -or $Value -is [long]) { return $Value }
    $parsed = $null
    if ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -match '^(?i)0x[0-9a-f]{1,16}$') {
            try { $parsed = [System.Convert]::ToInt64($text.Substring(2), 16) } catch { $parsed = $null }
        } elseif ($text -match '^-?[0-9]{1,19}$') {
            $number = [long]0
            if ([long]::TryParse($text, [System.Globalization.NumberStyles]::AllowLeadingSign, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
                $parsed = $number
            }
        }
    } elseif ($Value -is [double] -or $Value -is [decimal]) {
        if ([math]::Floor($Value) -eq $Value -and $Value -ge [long]::MinValue -and $Value -le [long]::MaxValue) { $parsed = [long]$Value }
    }
    if ($null -eq $parsed) {
        Add-LegacyWarning $Context ('Wpis rejestru „{0}” ma typ {1}, ale jego wartość nie jest liczbą. Sprawdź ją przed publikacją.' -f $Label, $Type)
        return , $Value
    }
    if ($parsed -ge [int]::MinValue -and $parsed -le [int]::MaxValue) { return [int]$parsed }
    return $parsed
}

function ConvertTo-LegacyRegistryEntry {
    # Wpis SystemSettings.CustomRegistry (3.1) -> defaults.systemSettings.registry[] (v4) albo $null.
    param($Entry, [int]$Number, $Context)
    if ($Entry -isnot [System.Collections.IDictionary]) {
        Add-LegacyWarning $Context ('Pominięto wpis rejestru nr {0}: oczekiwano obiektu z polami Path, Name, Value i PropertyType.' -f $Number)
        return $null
    }
    Add-LegacyUnknownFieldWarnings -Dictionary $Entry -Known @('Path', 'Name', 'Value', 'PropertyType') -Prefix ('SystemSettings.CustomRegistry[{0}]' -f ($Number - 1)) -Context $Context
    $rawPath = Get-LegacyText $Entry 'Path'
    $name = Get-LegacyText $Entry 'Name'
    if ($rawPath.Length -eq 0 -or $name.Length -eq 0) {
        Add-LegacyWarning $Context ('Pominięto wpis rejestru nr {0}: brak ścieżki (Path) albo nazwy (Name).' -f $Number)
        return $null
    }
    $label = $rawPath.TrimEnd('\') + '\' + $name
    $value = Get-LegacyField -Dictionary $Entry -Name 'Value'

    if ($script:StdLegacySecretNameRegex.IsMatch($name) -and (Test-LegacyHasContent -Value $value)) {
        if ($value -is [string]) { Add-LegacySecretValue -Context $Context -Value $value }
        Add-LegacyWarning $Context ('Pominięto wpis rejestru „{0}”, bo wygląda na hasło lub token. Takich wartości nie przenosimy do konfiguracji.' -f $label)
        return $null
    }

    $path = $rawPath.Replace('/', '\')
    $path = [regex]::Replace($path, '^(?i)(?:Microsoft\.PowerShell\.Core\\)?Registry::', '')
    $match = [regex]::Match($path, '^(?<hive>[A-Za-z_]+)(?::\\*|\\+|$)(?<rest>.*)$')
    $hive = ''
    if ($match.Success) {
        switch ($match.Groups['hive'].Value.ToUpperInvariant()) {
            'HKLM' { $hive = 'HKLM' }
            'HKEY_LOCAL_MACHINE' { $hive = 'HKLM' }
            'HKCU' { $hive = 'HKCU' }
            'HKEY_CURRENT_USER' { $hive = 'HKCU' }
        }
    }
    if ($hive.Length -eq 0) {
        Add-LegacyWarning $Context ('Pominięto wpis rejestru „{0}”: STD 4 obsługuje tylko gałęzie HKLM i HKCU (oraz DefaultUser).' -f $label)
        return $null
    }
    $subPath = [regex]::Replace($match.Groups['rest'].Value, '\\{2,}', '\').Trim('\')
    if ($subPath.Length -eq 0) {
        Add-LegacyWarning $Context ('Pominięto wpis rejestru „{0}”: brak klucza pod gałęzią {1}.' -f $label, $hive)
        return $null
    }

    $rawType = Get-LegacyText $Entry 'PropertyType'
    $type = ''
    if ($rawType.Length -eq 0) {
        $type = 'String'
    } else {
        foreach ($known in $script:StdLegacyRegistryTypes) {
            if ($known -ieq $rawType) { $type = $known; break }
        }
    }
    if ($type.Length -eq 0) {
        Add-LegacyWarning $Context ('Pominięto wpis rejestru „{0}”: nieznany typ „{1}”. Dozwolone: {2}.' -f $label, $rawType, ($script:StdLegacyRegistryTypes -join ', '))
        return $null
    }
    if ($hive -eq 'HKCU') { $Context.HasHkcu = $true }
    $converted = ConvertTo-LegacyRegistryValue -Value $value -Type $type -Label $label -Context $Context
    return [ordered]@{ hive = $hive; path = $subPath; name = $name; type = $type; value = $converted }
}

function Hide-LegacySecretText {
    # Zastępuje w tekście wartości haseł i tokenów z 3.1 znakami *** (poza wstawkami {{secret:...}}).
    param([string]$Text, [string[]]$Secrets)
    if ([string]::IsNullOrEmpty($Text) -or $Secrets.Length -eq 0) { return $Text }
    $parts = $script:StdLegacyPlaceholderRegex.Split($Text)
    $builder = New-Object System.Text.StringBuilder
    foreach ($part in $parts) {
        $piece = $part
        if (-not $script:StdLegacyPlaceholderRegex.IsMatch($piece)) {
            foreach ($secret in $Secrets) {
                if ($piece.IndexOf($secret, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $piece = [regex]::Replace($piece, [regex]::Escape($secret), '***', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                }
            }
        }
        [void]$builder.Append($piece)
    }
    return $builder.ToString()
}

function Protect-LegacyNode {
    # Końcowa kontrola: w każdym tekście dokumentu szukamy wartości haseł i tokenów z 3.1.
    param($Node, [string]$Path, [string[]]$Secrets, $Context)
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($key in (Get-LegacyKeyList -Dictionary $Node)) {
            $value = $Node[$key]
            $childPath = $key
            if ($Path.Length -gt 0) { $childPath = $Path + '.' + $key }
            if ($value -is [string]) {
                $masked = Hide-LegacySecretText -Text $value -Secrets $Secrets
                if (-not [string]::Equals($masked, $value, [System.StringComparison]::Ordinal)) {
                    $Node[$key] = $masked
                    $Context.MaskedFields.Add($childPath)
                }
            } else {
                Protect-LegacyNode -Node $value -Path $childPath -Secrets $Secrets -Context $Context
            }
        }
    } elseif ($Node -is [System.Collections.IList]) {
        for ($i = 0; $i -lt $Node.Count; $i++) {
            $value = $Node[$i]
            $childPath = '{0}[{1}]' -f $Path, $i
            if ($value -is [string]) {
                $masked = Hide-LegacySecretText -Text $value -Secrets $Secrets
                if (-not [string]::Equals($masked, $value, [System.StringComparison]::Ordinal)) {
                    $Node[$i] = $masked
                    $Context.MaskedFields.Add($childPath)
                }
            } else {
                Protect-LegacyNode -Node $value -Path $childPath -Secrets $Secrets -Context $Context
            }
        }
    }
}

function ConvertTo-LegacyApps {
    # Lista nazw programów 3.1 -> klucze katalogu v4 (bez powtórzeń); nieznane pomijamy z ostrzeżeniem.
    param($Names, [string]$ProfileName, $KeyMap, $Context)
    $apps = New-Object 'System.Collections.Generic.List[object]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in (ConvertTo-LegacyList -Value $Names)) {
        if ($item -isnot [string]) {
            Add-LegacyWarning $Context ('Profil „{0}”: pominięto pozycję, która nie jest nazwą programu.' -f $ProfileName)
            continue
        }
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $key = $null
        if ($KeyMap.ContainsKey($item)) { $key = $KeyMap[$item] }
        elseif ($KeyMap.ContainsKey($item.Trim())) { $key = $KeyMap[$item.Trim()] }
        if ($null -eq $key) {
            Add-LegacyWarning $Context ('Profil „{0}”: pominięto program „{1}”, którego nie ma w katalogu.' -f $ProfileName, $item.Trim())
            continue
        }
        if ($seen.Add($key)) { $apps.Add($key) }
    }
    return , $apps
}

# ---------------------------------------------------------------------------------------------
# Funkcje publiczne
# ---------------------------------------------------------------------------------------------

function ConvertFrom-StdLegacyConfig {
    <#
    .SYNOPSIS
    Konwersja konfiguracji STD 3.1 (config.json) do konfiguracji klienta STD 4 (SPEC 3.5).

    .DESCRIPTION
    Zwraca @{ Document; Warnings; SecretsToSet }:
    - Document - [ordered] dokument v4 (schemaVersion 4, kind std.release) bez podpisu, bez revision i publishedAt,
      z jedną lokalizacją "default" ("Domyślna");
    - Warnings - string[] po polsku: co pominięto albo co trzeba poprawić przed publikacją;
    - SecretsToSet - string[] nazw sekretów ({{secret:nazwa}}), które trzeba ustawić przed publikacją.
    Dane logowania (WebAuth, AntyVirus.Credentials, DomainJoin.Username) i hasła są pomijane - żadna wartość
    hasła ani tokenu z 3.1 nie trafia do wyniku.
    Błędy: [STD-LEGACY-ARG] zły identyfikator lub pusta nazwa klienta, [STD-LEGACY-INPUT] brak pliku albo
    wejście nie jest obiektem JSON, [STD-JSON] niepoprawny JSON.

    .PARAMETER Path
    Ścieżka do config.json z STD 3.1 (UTF-8, z BOM albo bez).

    .PARAMETER InputObject
    Konfiguracja 3.1 już wczytana: wynik ConvertFrom-Json, hashtable albo tekst JSON.

    .PARAMETER ClientId
    Identyfikator klienta v4 (^[a-z0-9][a-z0-9-]{1,39}$).

    .PARAMETER ClientName
    Nazwa klienta wyświetlana technikom.

    .EXAMPLE
    $r = ConvertFrom-StdLegacyConfig -Path .\config.json -ClientId lizard -ClientName 'Lizard-Tech'
    $r.Warnings; $r.SecretsToSet
    #>
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Path', Position = 0)]
        [string]$Path,
        [Parameter(Mandatory = $true, ParameterSetName = 'InputObject')]
        [AllowNull()]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ClientName
    )

    if ($ClientId -cnotmatch $script:StdLegacyClientIdPattern) {
        throw ('[STD-LEGACY-ARG] Identyfikator klienta „{0}” jest nieprawidłowy. Dozwolone są małe litery, cyfry i myślnik (2-40 znaków, na początku litera lub cyfra).' -f $ClientId)
    }
    if ([string]::IsNullOrWhiteSpace($ClientName)) {
        throw '[STD-LEGACY-ARG] Nazwa klienta nie może być pusta.'
    }

    # --- Wczytanie wejścia -----------------------------------------------------------------
    if ($PSCmdlet.ParameterSetName -eq 'Path') {
        $full = Resolve-CoreFullPath -Path $Path
        if (-not [System.IO.File]::Exists($full)) {
            throw ('[STD-LEGACY-INPUT] Nie znaleziono pliku konfiguracji 3.1: {0}' -f $full)
        }
        try {
            $text = [System.IO.File]::ReadAllText($full, (New-Object System.Text.UTF8Encoding($false)))
        } catch {
            throw ('[STD-LEGACY-INPUT] Nie udało się odczytać pliku {0} - {1}' -f $full, (Format-CoreErrorText $_))
        }
        $config = ConvertFrom-StdJson -Json $text
    } elseif ($InputObject -is [string]) {
        $config = ConvertFrom-StdJson -Json $InputObject
    } else {
        $config = ConvertTo-StdHashtable -InputObject $InputObject
    }
    if ($config -isnot [System.Collections.IDictionary]) {
        throw '[STD-LEGACY-INPUT] Konfiguracja 3.1 musi być obiektem JSON (na przykład zawartością pliku config.json).'
    }

    $context = @{
        Warnings     = New-Object 'System.Collections.Generic.List[string]'
        SecretsToSet = New-Object 'System.Collections.Generic.List[string]'
        InputSecrets = New-Object 'System.Collections.Generic.List[string]'
        MaskedFields = New-Object 'System.Collections.Generic.List[string]'
        HasHkcu      = $false
    }
    Add-LegacyInputSecrets -Node $config -Context $context

    # --- Źródła ----------------------------------------------------------------------------
    $sources = [ordered]@{}
    $globalPaths = Get-LegacyField $config 'InstallSourcePaths'
    $globalNetwork = Get-LegacyText $globalPaths 'network'
    $globalWeb = Get-LegacyText $globalPaths 'web'
    if ($globalPaths -is [System.Collections.IDictionary]) {
        Add-LegacyUnknownFieldWarnings -Dictionary $globalPaths -Known @('network', 'web') -Prefix 'InstallSourcePaths' -Context $context
    }
    if ($globalNetwork.Length -gt 0) {
        $sources['lan'] = New-LegacySource -Kind 'network' -Value $globalNetwork -Field 'InstallSourcePaths.network' -Name 'lan' -Context $context
    }
    if ($globalWeb.Length -gt 0) {
        $sources['web'] = New-LegacySource -Kind 'web' -Value $globalWeb -Field 'InstallSourcePaths.web' -Name 'web' -Context $context
    }
    $globalSourceName = @{ network = 'lan'; web = 'web' }

    $defaultSourceText = Get-LegacyText $config 'DefaultInstallSource'
    $defaultKind = Get-LegacySourceKind -Value $defaultSourceText
    if ($defaultKind -eq '') {
        Add-LegacyWarning $context 'Brak źródła domyślnego (DefaultInstallSource). Przyjęto kolejność: najpierw udział sieciowy, potem web.'
        $defaultKind = 'network'
    } elseif ($defaultKind -eq '?') {
        Add-LegacyWarning $context ('Nieznane źródło domyślne „{0}” (DefaultInstallSource). Przyjęto kolejność: najpierw udział sieciowy, potem web.' -f $defaultSourceText)
        $defaultKind = 'network'
    }
    $wingetDefault = ($defaultKind -eq 'winget')
    if ($defaultKind -eq 'network') { $orderKinds = @('network', 'web') } else { $orderKinds = @('web', 'network') }
    $siteOrder = New-Object 'System.Collections.Generic.List[object]'
    foreach ($kind in $orderKinds) {
        $name = $globalSourceName[$kind]
        if ($sources.Contains($name)) { $siteOrder.Add($name) }
    }

    # --- TeamViewer i antywirus (klucze zarezerwowane przed programami) --------------------
    $teamViewer = Get-LegacyField $config 'TeamViewer'
    $teamViewerEntry = $null
    if ($null -ne $teamViewer) {
        if ($teamViewer -isnot [System.Collections.IDictionary]) {
            Add-LegacyWarning $context 'Sekcja TeamViewer ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        } else {
            Add-LegacyUnknownFieldWarnings -Dictionary $teamViewer -Known @('FileName', 'Arguments') -Prefix 'TeamViewer' -Context $context
            $tvFile = Get-LegacyText $teamViewer 'FileName'
            $tvArgs = Protect-LegacyKnownSecretArgs -Arguments (Get-LegacyText $teamViewer 'Arguments') -Context $context
            if ($tvFile.Length -eq 0) {
                if (Test-LegacyHasContent -Value $teamViewer) {
                    Add-LegacyWarning $context 'Sekcja TeamViewer nie ma pliku instalatora (FileName). TeamViewera nie przeniesiono do katalogu.'
                }
            } else {
                # Jak Install-TeamViewer w 3.1: przy źródle winget FileName jest identyfikatorem pakietu.
                $teamViewerEntry = [ordered]@{ name = 'TeamViewer' }
                if ($wingetDefault) {
                    $teamViewerEntry['winget'] = $tvFile
                    if ($tvArgs.Length -gt 0) {
                        $teamViewerEntry['wingetArgs'] = $tvArgs
                        Add-LegacyWarning $context 'TeamViewer będzie instalowany przez winget, więc jego argumenty z 3.1 trafiły do wingetArgs. Sprawdź je - argumenty instalatora MSI winget przyjmuje zwykle tylko z przełącznikiem --override.'
                    }
                } else {
                    $teamViewerEntry['file'] = $tvFile
                    if ($tvArgs.Length -gt 0) { $teamViewerEntry['args'] = $tvArgs }
                }
            }
        }
    }

    $antivirus = Get-LegacyField $config 'AntyVirus'
    $antivirusEntry = $null
    if ($null -ne $antivirus) {
        if ($antivirus -isnot [System.Collections.IDictionary]) {
            Add-LegacyWarning $context 'Sekcja AntyVirus ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        } else {
            Add-LegacyUnknownFieldWarnings -Dictionary $antivirus -Known @('DefaultInstallSource', 'InstallSourcePaths', 'Credentials', 'FileName') -Prefix 'AntyVirus' -Context $context
            if (Test-LegacyHasContent -Value (Get-LegacyField $antivirus 'Credentials')) {
                Add-LegacyWarning $context 'Pominięto AntyVirus.Credentials. W STD 4 pliki pobiera się kontem technika, a loginów i haseł nie zapisuje się w konfiguracji.'
            }
            $avFile = Get-LegacyText $antivirus 'FileName'
            if ($avFile.Length -eq 0) {
                # Domyślna konfiguracja 3.1 ma pustą sekcję z samym DefaultInstallSource - wtedy bez ostrzeżenia.
                $avRest = [ordered]@{}
                foreach ($key in (Get-LegacyKeyList -Dictionary $antivirus)) {
                    if ($key -ine 'Credentials' -and $key -ine 'DefaultInstallSource') { $avRest[$key] = $antivirus[$key] }
                }
                if (Test-LegacyHasContent -Value $avRest) {
                    Add-LegacyWarning $context 'Sekcja AntyVirus nie ma pliku instalatora (FileName). Antywirusa nie przeniesiono do katalogu.'
                }
            } else {
                $antivirusEntry = [ordered]@{ name = 'Antywirus' }
                # Jak Install-AV w 3.1: własne DefaultInstallSource antywirusa albo globalne.
                $avSourceText = Get-LegacyText $antivirus 'DefaultInstallSource'
                $avKind = Get-LegacySourceKind -Value $avSourceText
                if ($avKind -eq '?') {
                    Add-LegacyWarning $context ('Nieznane źródło antywirusa „{0}” (AntyVirus.DefaultInstallSource). Przyjęto kolejność: najpierw udział sieciowy, potem web.' -f $avSourceText)
                    $avKind = 'network'
                } elseif ($avKind -eq '') {
                    $avKind = $defaultKind
                }
                if ($avKind -eq 'winget') {
                    $antivirusEntry['winget'] = $avFile
                } else {
                    $antivirusEntry['file'] = $avFile
                    $avPaths = Get-LegacyField $antivirus 'InstallSourcePaths'
                    if ($avPaths -is [System.Collections.IDictionary]) {
                        Add-LegacyUnknownFieldWarnings -Dictionary $avPaths -Known @('network', 'web') -Prefix 'AntyVirus.InstallSourcePaths' -Context $context
                    }
                    # Własne ścieżki antywirusa -> av-lan / av-web; brak własnej ścieżki danego rodzaju -> globalna (jak w 3.1).
                    $avOwn = @{ network = (Get-LegacyText $avPaths 'network'); web = (Get-LegacyText $avPaths 'web') }
                    $avNames = @{ network = 'av-lan'; web = 'av-web' }
                    $avGlobal = @{ network = $globalNetwork; web = $globalWeb }
                    $avSourceFor = @{}
                    foreach ($kind in @('network', 'web')) {
                        $own = $avOwn[$kind]
                        if ($own.Length -gt 0) {
                            if ($avGlobal[$kind].Length -gt 0 -and $own.TrimEnd('\', '/') -ieq $avGlobal[$kind].TrimEnd('\', '/')) {
                                $avSourceFor[$kind] = $globalSourceName[$kind]
                            } else {
                                $sources[$avNames[$kind]] = New-LegacySource -Kind $kind -Value $own -Field ('AntyVirus.InstallSourcePaths.' + $kind) -Name $avNames[$kind] -Context $context
                                $avSourceFor[$kind] = $avNames[$kind]
                            }
                        } elseif ($sources.Contains($globalSourceName[$kind])) {
                            $avSourceFor[$kind] = $globalSourceName[$kind]
                        }
                    }
                    if ($avKind -eq 'web') { $avKinds = @('web', 'network') } else { $avKinds = @('network', 'web') }
                    $avOrder = New-Object 'System.Collections.Generic.List[object]'
                    foreach ($kind in $avKinds) {
                        if ($avSourceFor.ContainsKey($kind)) { $avOrder.Add($avSourceFor[$kind]) }
                    }
                    if ($avOrder.Count -gt 0 -and (($avOrder.ToArray() -join '|') -ne ($siteOrder.ToArray() -join '|'))) {
                        $antivirusEntry['sourceOrder'] = $avOrder.ToArray()
                    }
                }
            }
        }
    }

    if ($sources.get_Count() -eq 0) {
        Add-LegacyWarning $context 'Konfiguracja 3.1 nie ma żadnego źródła instalatorów (InstallSourcePaths). Dodaj źródło w sekcji sources przed publikacją.'
    }

    # --- Programy -> katalog ---------------------------------------------------------------
    $usedKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -ne $teamViewerEntry) { [void]$usedKeys.Add($script:StdLegacyTeamViewerKey) }
    if ($null -ne $antivirusEntry) { [void]$usedKeys.Add($script:StdLegacyAntivirusKey) }

    $programs = Get-LegacyField $config 'Programs'
    $programItems = New-Object 'System.Collections.Generic.List[object]'
    $wingetPreferredFile = New-Object 'System.Collections.Generic.List[string]'
    if ($null -ne $programs -and $programs -isnot [System.Collections.IDictionary]) {
        Add-LegacyWarning $context 'Sekcja Programs ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        $programs = $null
    }
    foreach ($rawName in (Get-LegacyKeyList -Dictionary $programs)) {
        $displayName = $rawName.Trim()
        $program = $programs[$rawName]
        if ($displayName.Length -eq 0) {
            Add-LegacyWarning $context 'Pominięto program bez nazwy.'
            continue
        }
        if ($program -isnot [System.Collections.IDictionary]) {
            Add-LegacyWarning $context ('Program „{0}” ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto go.' -f $displayName)
            continue
        }
        Add-LegacyUnknownFieldWarnings -Dictionary $program -Known $script:StdLegacyProgramFields -Prefix ('Programs.' + $displayName) -Context $context
        $entry = ConvertTo-LegacyProgramEntry -DisplayName $displayName -Program $program -WingetDefault $wingetDefault -Context $context
        if ($null -eq $entry) {
            Add-LegacyWarning $context ('Program „{0}” nie ma pliku (FileName), adresu (DownloadUrl) ani identyfikatora winget (WingetId). Pominięto go.' -f $displayName)
            continue
        }
        if ($wingetDefault -and $entry.Contains('file') -and $entry.Contains('winget')) { $wingetPreferredFile.Add($displayName) }
        $enabled = (ConvertTo-LegacyBool -Value (Get-LegacyField $program 'Enabled')) -eq $true
        $programItems.Add(@{ RawName = $rawName; Name = $displayName; Entry = $entry; Enabled = $enabled; Key = $null })
    }
    if ($wingetPreferredFile.Count -gt 0) {
        Add-LegacyWarning $context ('W 3.1 źródłem domyślnym był winget. W STD 4 programy, które mają i plik (FileName), i identyfikator winget, instalują się najpierw z pliku ze źródła, a winget jest zapasem. Dotyczy: {0}.' -f ($wingetPreferredFile.ToArray() -join ', '))
    }

    # Klucze: najpierw nazwy, które już pasują do wzorca (zostają bez zmian), potem oczyszczone.
    foreach ($item in $programItems) {
        if ($item.Name -cmatch $script:StdLegacyCatalogKeyPattern -and $usedKeys.Add($item.Name)) { $item.Key = $item.Name }
    }
    foreach ($item in $programItems) {
        if ($null -ne $item.Key) { continue }
        $base = ConvertTo-LegacyCatalogKey -Name $item.Name
        $item.Key = New-LegacyUniqueKey -Base $base -Used $usedKeys
        if ($item.Key -cne $base) {
            Add-LegacyWarning $context ('Program „{0}” ma w katalogu klucz „{1}”, bo klucz „{2}” był już zajęty.' -f $item.Name, $item.Key, $base)
        }
    }

    $catalog = [ordered]@{}
    $keyMap = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $enabledKeys = New-Object 'System.Collections.Generic.List[object]'
    foreach ($item in $programItems) {
        $catalog[$item.Key] = $item.Entry
        if (-not $keyMap.ContainsKey($item.RawName)) { $keyMap[$item.RawName] = $item.Key }
        if (-not $keyMap.ContainsKey($item.Name)) { $keyMap[$item.Name] = $item.Key }
        if ($item.Enabled) { $enabledKeys.Add($item.Key) }
    }
    if ($null -ne $teamViewerEntry) { $catalog[$script:StdLegacyTeamViewerKey] = $teamViewerEntry }
    if ($null -ne $antivirusEntry) { $catalog[$script:StdLegacyAntivirusKey] = $antivirusEntry }
    if ($catalog.get_Count() -eq 0) {
        Add-LegacyWarning $context 'Konfiguracja 3.1 nie ma żadnego programu do przeniesienia. Dodaj pozycje katalogu przed publikacją.'
    }

    # --- Zadania (DefaultCheckboxes) -------------------------------------------------------
    $taskValues = @{}
    $installTeamViewer = $false
    $installAntivirus = $false
    $checkboxes = Get-LegacyField $config 'DefaultCheckboxes'
    if ($null -ne $checkboxes -and $checkboxes -isnot [System.Collections.IDictionary]) {
        Add-LegacyWarning $context 'Sekcja DefaultCheckboxes ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        $checkboxes = $null
    }
    foreach ($key in (Get-LegacyKeyList -Dictionary $checkboxes)) {
        $flag = ConvertTo-LegacyBool -Value $checkboxes[$key]
        if ($script:StdLegacyTaskMap.Contains($key)) {
            if ($null -eq $flag) {
                Add-LegacyWarning $context ('Zadanie „{0}” w DefaultCheckboxes nie ma wartości true ani false. Pominięto je.' -f $key)
            } else {
                $taskValues[$script:StdLegacyTaskMap[$key]] = $flag
            }
        } elseif ($script:StdLegacyDroppedTasks.Contains($key)) {
            if ($flag -eq $true) { Add-LegacyWarning $context $script:StdLegacyDroppedTasks[$key] }
        } elseif ($key -ieq 'InstallTeamViewer') {
            $installTeamViewer = ($flag -eq $true)
            if ($installTeamViewer) {
                if ($null -ne $teamViewerEntry) {
                    Add-LegacyWarning $context 'Zadania InstallTeamViewer nie ma w STD 4. TeamViewer jest zwykłą pozycją katalogu i dopisano go do wszystkich profili.'
                } else {
                    Add-LegacyWarning $context 'Zadanie InstallTeamViewer było włączone, ale TeamViewera nie przeniesiono do katalogu, więc nie dopisano go do profili.'
                }
            }
        } elseif ($key -ieq 'InstallAV') {
            $installAntivirus = ($flag -eq $true)
            if ($installAntivirus) {
                if ($null -ne $antivirusEntry) {
                    Add-LegacyWarning $context 'Zadania InstallAV nie ma w STD 4. Antywirus jest zwykłą pozycją katalogu i dopisano go do wszystkich profili.'
                } else {
                    Add-LegacyWarning $context 'Zadanie InstallAV było włączone, ale antywirusa nie przeniesiono do katalogu, więc nie dopisano go do profili.'
                }
            }
        } else {
            Add-LegacyWarning $context ('Pominięto nieznane zadanie „{0}” z DefaultCheckboxes.' -f $key)
        }
    }
    $addTeamViewer = $installTeamViewer -and ($null -ne $teamViewerEntry)
    $addAntivirus = $installAntivirus -and ($null -ne $antivirusEntry)
    if (($addTeamViewer -or $addAntivirus) -and $taskValues.ContainsKey('InstallApps') -and $taskValues['InstallApps'] -eq $false) {
        Add-LegacyWarning $context 'W 3.1 TeamViewer i antywirus instalowały się niezależnie od zadania InstallApplications. W STD 4 są pozycjami katalogu i zainstalują się tylko przy włączonym zadaniu InstallApps.'
    }
    $tasks = [ordered]@{}
    foreach ($taskId in $script:StdLegacyTaskOrder) {
        if ($taskValues.ContainsKey($taskId)) { $tasks[$taskId] = [bool]$taskValues[$taskId] }
    }

    # --- Profile ---------------------------------------------------------------------------
    $profiles = [ordered]@{}
    $profileNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $legacyProfiles = Get-LegacyField $config 'Profiles'
    if ($null -ne $legacyProfiles -and $legacyProfiles -isnot [System.Collections.IDictionary]) {
        Add-LegacyWarning $context 'Sekcja Profiles ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        $legacyProfiles = $null
    }
    $legacyProfileKeys = Get-LegacyKeyList -Dictionary $legacyProfiles
    if ($enabledKeys.Count -gt 0 -or $legacyProfileKeys.Length -eq 0) {
        [void]$profileNames.Add($script:StdLegacyDefaultProfile)
        $profiles[$script:StdLegacyDefaultProfile] = [ordered]@{ apps = $enabledKeys }
    }
    foreach ($rawProfile in $legacyProfileKeys) {
        $profileName = $rawProfile.Trim()
        if ($profileName.Length -eq 0) {
            Add-LegacyWarning $context 'Pominięto profil bez nazwy.'
            continue
        }
        $names = $legacyProfiles[$rawProfile]
        if ($names -is [System.Collections.IDictionary]) {
            Add-LegacyWarning $context ('Profil „{0}” ma nieprawidłowy zapis (oczekiwano listy programów). Pominięto go.' -f $profileName)
            continue
        }
        $finalName = New-LegacyUniqueName -Base $profileName -Used $profileNames
        if ($finalName -cne $profileName) {
            Add-LegacyWarning $context ('Profil „{0}” z konfiguracji 3.1 nazwano „{1}”, bo ta nazwa była już zajęta.' -f $profileName, $finalName)
        }
        $apps = ConvertTo-LegacyApps -Names $names -ProfileName $profileName -KeyMap $keyMap -Context $context
        $profiles[$finalName] = [ordered]@{ apps = $apps }
    }
    # TeamViewer i antywirus (DefaultCheckboxes.InstallTeamViewer / InstallAV) - na początku listy,
    # bo w 3.1 instalowały się przed pozostałymi programami.
    foreach ($profileName in (Get-LegacyKeyList -Dictionary $profiles)) {
        $current = $profiles[$profileName]['apps']
        $apps = New-Object 'System.Collections.Generic.List[object]'
        if ($addTeamViewer) { $apps.Add($script:StdLegacyTeamViewerKey) }
        if ($addAntivirus) { $apps.Add($script:StdLegacyAntivirusKey) }
        foreach ($key in $current) {
            if (-not $apps.Contains($key)) { $apps.Add($key) }
        }
        $profiles[$profileName]['apps'] = $apps.ToArray()
    }

    # --- Lokalizacja "default" -------------------------------------------------------------
    $site = [ordered]@{
        id          = $script:StdLegacySiteId
        name        = $script:StdLegacySiteName
        sourceOrder = $siteOrder.ToArray()
    }

    $wifiSection = Get-LegacyField $config 'WiFiProfile'
    if ($null -ne $wifiSection) {
        if ($wifiSection -isnot [System.Collections.IDictionary]) {
            Add-LegacyWarning $context 'Sekcja WiFiProfile ma nieprawidłowy zapis (oczekiwano obiektu z polem FileName). Pominięto ją.'
        } else {
            Add-LegacyUnknownFieldWarnings -Dictionary $wifiSection -Known @('FileName') -Prefix 'WiFiProfile' -Context $context
            $wifi = New-Object 'System.Collections.Generic.List[object]'
            foreach ($file in (ConvertTo-LegacyList -Value (Get-LegacyField $wifiSection 'FileName'))) {
                if ($file -isnot [string]) {
                    Add-LegacyWarning $context 'Pominięto pozycję WiFiProfile.FileName, która nie jest nazwą pliku.'
                    continue
                }
                if ([string]::IsNullOrWhiteSpace($file)) { continue }
                $wifi.Add([ordered]@{ file = $file.Trim() })
                Add-LegacyWarning $context ('Profil Wi-Fi „{0}” przeniesiono jako plik XML (dawny format). Plik musi być w źródle instalatorów klienta i może zawierać hasło sieci jawnym tekstem - zalecany jest zapis ssid i keyRef z kluczem w sekretach.' -f $file.Trim())
            }
            if ($wifi.Count -gt 0) { $site['wifi'] = $wifi.ToArray() }
        }
    }

    $domainJoin = Get-LegacyField $config 'DomainJoin'
    $domainName = ''
    if ($null -ne $domainJoin) {
        if ($domainJoin -isnot [System.Collections.IDictionary]) {
            Add-LegacyWarning $context 'Sekcja DomainJoin ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        } else {
            $domainName = Get-LegacyText $domainJoin 'DomainName'
            foreach ($key in (Get-LegacyKeyList -Dictionary $domainJoin)) {
                if ($key -ieq 'DomainName') { continue }
                if ($key -ieq 'Username') {
                    if (Test-LegacyHasContent -Value $domainJoin[$key]) {
                        Add-LegacyWarning $context 'Pominięto DomainJoin.Username. Login i hasło do domeny technik wpisuje ręcznie przy każdym wdrożeniu.'
                    }
                    continue
                }
                Add-LegacyWarning $context ('Pominięto pole „DomainJoin.{0}”. W konfiguracji STD 4 jest tylko nazwa domeny i OU, a dane logowania wpisuje się ręcznie przy wdrożeniu.' -f $key)
            }
            if ($domainName.Length -gt 0) { $site['domainJoin'] = [ordered]@{ domain = $domainName } }
        }
    }
    if ($domainName.Length -eq 0 -and $tasks.Contains('JoinDomain') -and $tasks['JoinDomain']) {
        Add-LegacyWarning $context 'Zadanie JoinDomain jest włączone, ale konfiguracja 3.1 nie ma nazwy domeny (DomainJoin.DomainName). Uzupełnij ją przed publikacją.'
    }

    $localAdmin = Get-LegacyField $config 'LocalAdmin'
    $localAdminName = ''
    if ($null -ne $localAdmin) {
        if ($localAdmin -isnot [System.Collections.IDictionary]) {
            Add-LegacyWarning $context 'Sekcja LocalAdmin ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        } else {
            $localAdminName = Get-LegacyText $localAdmin 'Username'
            foreach ($key in (Get-LegacyKeyList -Dictionary $localAdmin)) {
                if ($key -ieq 'Username') { continue }
                Add-LegacyWarning $context ('Pominięto pole „LocalAdmin.{0}”. W konfiguracji STD 4 jest tylko nazwa konta, a hasło lokalnego administratora wpisuje się ręcznie przy wdrożeniu.' -f $key)
            }
            if ($localAdminName.Length -gt 0) { $site['localAdmin'] = [ordered]@{ username = $localAdminName } }
        }
    }
    if ($localAdminName.Length -eq 0 -and $tasks.Contains('CreateLocalAdmin') -and $tasks['CreateLocalAdmin']) {
        Add-LegacyWarning $context 'Zadanie CreateLocalAdmin jest włączone, ale konfiguracja 3.1 nie ma nazwy konta (LocalAdmin.Username). Uzupełnij ją przed publikacją.'
    }
    $site['profiles'] = $profiles

    # --- Ustawienia systemu i rejestr ------------------------------------------------------
    $systemSettings = [ordered]@{}
    $legacySettings = Get-LegacyField $config 'SystemSettings'
    if ($null -ne $legacySettings -and $legacySettings -isnot [System.Collections.IDictionary]) {
        Add-LegacyWarning $context 'Sekcja SystemSettings ma nieprawidłowy zapis (oczekiwano obiektu). Pominięto ją.'
        $legacySettings = $null
    }
    $registry = New-Object 'System.Collections.Generic.List[object]'
    foreach ($key in (Get-LegacyKeyList -Dictionary $legacySettings)) {
        $value = $legacySettings[$key]
        if ($key -ieq 'CustomRegistry') {
            $number = 0
            foreach ($regEntry in (ConvertTo-LegacyList -Value $value)) {
                $number++
                if ($null -eq $regEntry) { continue }
                $converted = ConvertTo-LegacyRegistryEntry -Entry $regEntry -Number $number -Context $context
                if ($null -ne $converted) { $registry.Add($converted) }
            }
            continue
        }
        if ($value -is [bool]) {
            $systemSettings[$key] = $value
        } else {
            Add-LegacyWarning $context ('Pominięto ustawienie systemu „{0}”, bo nie ma wartości true ani false.' -f $key)
        }
    }
    if ($context.HasHkcu) {
        Add-LegacyWarning $context 'Wpisy rejestru HKCU zmieniają ustawienia konta, na którym działa STD. Jeśli mają dotyczyć nowych użytkowników, zmień gałąź (hive) na DefaultUser.'
    }
    if ($registry.Count -gt 0) { $systemSettings['registry'] = $registry.ToArray() }

    # --- Skrypty poinstalacyjne ------------------------------------------------------------
    $scripts = New-Object 'System.Collections.Generic.List[object]'
    $number = 0
    foreach ($scriptEntry in (ConvertTo-LegacyList -Value (Get-LegacyField $config 'PostInstallScripts'))) {
        $number++
        if ($scriptEntry -isnot [string]) {
            Add-LegacyWarning $context ('Pominięto wpis nr {0} w PostInstallScripts: oczekiwano nazwy pliku skryptu.' -f $number)
            continue
        }
        $scriptText = $scriptEntry.Trim()
        if ($scriptText.Length -eq 0) { continue }
        if ($scriptText -match '^(?i)https?://') {
            $scriptUrl = Remove-LegacyUrlCredential -Url $scriptText -Field ('PostInstallScripts[{0}]' -f ($number - 1)) -Context $context
            $leaf = ''
            try { $leaf = [System.Uri]::UnescapeDataString(([System.Uri]$scriptUrl).AbsolutePath.TrimEnd('/').Split('/')[-1]) } catch { $leaf = '' }
            if ([string]::IsNullOrWhiteSpace($leaf)) {
                Add-LegacyWarning $context ('Pominięto wpis nr {0} w PostInstallScripts: z adresu „{1}” nie da się odczytać nazwy pliku.' -f $number, $scriptUrl)
                continue
            }
            Add-LegacyWarning $context ('Skrypt „{0}” był w 3.1 pobierany z adresu URL. W STD 4 skrypty są pobierane ze źródeł klienta - umieść plik „{1}” w źródle.' -f $scriptUrl, $leaf)
            $scriptText = $leaf
        }
        $scripts.Add([ordered]@{ file = $scriptText })
    }

    # --- Sekcje pominięte i nieznane klucze ------------------------------------------------
    if (Test-LegacyHasContent -Value (Get-LegacyField $config 'WebAuth')) {
        Add-LegacyWarning $context 'Pominięto WebAuth. W STD 4 pliki pobiera się kontem technika (logowanie do serwera STD), a loginów i haseł nie zapisuje się w konfiguracji.'
    }
    if (Test-LegacyHasContent -Value (Get-LegacyField $config 'HardwareAudit')) {
        Add-LegacyWarning $context 'Pominięto HardwareAudit.ExportPath. W STD 4 audyt sprzętu trafia do raportu wdrożenia, a klucza BitLocker nie zapisuje się do pliku.'
    }
    if (Test-LegacyHasContent -Value (Get-LegacyField $config 'AutoUpdate')) {
        Add-LegacyWarning $context 'Pominięto AutoUpdate. STD 4 pobiera nowe wersje z serwera STD.'
    }
    if (Test-LegacyHasContent -Value (Get-LegacyField $config 'CustomWebDataLocation')) {
        Add-LegacyWarning $context 'Pominięto CustomWebDataLocation. STD 4 nie używa tego ustawienia.'
    }
    foreach ($key in (Get-LegacyKeyList -Dictionary $config)) {
        if ($script:StdLegacyKnownTopKeys -contains $key) { continue }
        Add-LegacyWarning $context ('Pominięto klucz „{0}”, którego STD 4 nie obsługuje.' -f $key)
    }

    # --- Dokument v4 -----------------------------------------------------------------------
    $defaults = [ordered]@{}
    if ($tasks.get_Count() -gt 0) { $defaults['tasks'] = $tasks }
    $defaults['bloatware'] = [object[]]$script:StdLegacyBloatware.Clone()
    if ($systemSettings.get_Count() -gt 0) { $defaults['systemSettings'] = $systemSettings }
    if ($scripts.Count -gt 0) { $defaults['scripts'] = $scripts.ToArray() }

    $sites = New-Object 'System.Collections.Generic.List[object]'
    $sites.Add($site)

    $document = [ordered]@{
        schemaVersion = 4
        kind          = 'std.release'
        client        = [ordered]@{ id = $ClientId; name = $ClientName.Trim() }
        allowOffline  = $false
        requireSha256 = $false
        sources       = $sources
        catalog       = $catalog
        defaults      = $defaults
        sites         = $sites.ToArray()
    }

    # --- Kontrola: żadna wartość hasła ani tokenu z 3.1 nie może zostać w wyniku -----------
    $secretList = New-Object 'System.Collections.Generic.List[string]'
    foreach ($secret in $context.InputSecrets) { $secretList.Add($secret) }
    # Najdłuższe najpierw (żeby krótszy sekret nie zostawił końcówki dłuższego).
    $secretArray = [string[]]@($secretList.ToArray() | Sort-Object -Property Length -Descending)
    $warningArray = $context.Warnings.ToArray()
    if ($secretArray.Length -gt 0) {
        Protect-LegacyNode -Node $document -Path '' -Secrets $secretArray -Context $context
        for ($i = 0; $i -lt $warningArray.Length; $i++) {
            $warningArray[$i] = Hide-LegacySecretText -Text $warningArray[$i] -Secrets $secretArray
        }
        $extra = New-Object 'System.Collections.Generic.List[string]'
        foreach ($warning in $warningArray) { $extra.Add($warning) }
        foreach ($field in $context.MaskedFields) {
            $extra.Add((Hide-LegacySecretText -Text ('W polu {0} była wartość hasła lub tokenu z konfiguracji 3.1. Zastąpiono ją znakami *** - popraw to pole.' -f $field) -Secrets $secretArray))
        }
        $warningArray = $extra.ToArray()
    }

    Write-StdLog -Level Info -Source 'Std.Legacy' -Message ('Przekonwertowano konfigurację 3.1 klienta {0}: pozycji katalogu {1}, profili {2}, ostrzeżeń {3}.' -f $ClientId, $catalog.get_Count(), $profiles.get_Count(), $warningArray.Length)

    return @{
        Document     = $document
        Warnings     = [string[]]$warningArray
        SecretsToSet = [string[]]$context.SecretsToSet.ToArray()
    }
}
