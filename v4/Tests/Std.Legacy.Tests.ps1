# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy Modules\Std.Legacy.ps1 i Tools\Convert-StdLegacyConfig.ps1 (Pester 5, Windows PowerShell 5.1):
# konwersja config.json z STD 3.1 do konfiguracji klienta STD 4 (SPEC 3.5 i 4.3).
# Wejście: przykład z repozytorium (config.example.json, UTF-8 z BOM i CRLF) oraz małe konfiguracje 3.1
# z przypadkami brzegowymi. Najważniejsze: żadna wartość hasła ani tokenu z 3.1 nie trafia do wyniku.
# Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' (nawiasy w -like trzeba poprzedzić znakiem `).

Describe 'Std.Legacy' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\STD.psd1" -Force
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'

        $examplePath = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'config.example.json'
        $toolPath = Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path 'Tools' 'Convert-StdLegacyConfig.ps1')
        $catalogKeyPattern = '^[A-Za-z0-9][A-Za-z0-9._+-]{0,63}$'
        $taskIds = @('WaitForNetwork', 'SuspendSleep', 'ConnectWiFi', 'RemoveBloatware', 'UninstallPreinstalled', 'UninstallM365',
            'UninstallOneDrive', 'InstallApps', 'CreateLocalAdmin', 'RenameComputer', 'JoinDomain', 'SystemSettings',
            'RunScripts', 'HardwareAudit', 'EnableBitLocker', 'RebootAtEnd')

        function Convert-TestLegacy {
            # Konfiguracja 3.1 jako tekst JSON -> wynik konwersji (wejście jak z ConvertFrom-Json).
            param([string]$Json, [string]$ClientId = 'test')
            return (ConvertFrom-StdLegacyConfig -InputObject (ConvertFrom-Json -InputObject $Json) -ClientId $ClientId -ClientName 'Klient testowy')
        }

        function Get-TestAllStrings {
            # Wszystkie teksty wyniku (klucze i wartości, także ostrzeżenia i lista sekretów).
            param($Node)
            $list = New-Object System.Collections.Generic.List[string]
            $stack = New-Object System.Collections.Stack
            $stack.Push($Node)
            while ($stack.Count -gt 0) {
                $current = $stack.Pop()
                if ($null -eq $current) { continue }
                if ($current -is [string]) { $list.Add($current); continue }
                if ($current -is [System.Collections.IDictionary]) {
                    foreach ($entry in $current.GetEnumerator()) { $list.Add([string]$entry.Key); $stack.Push($entry.Value) }
                    continue
                }
                if ($current -is [System.Collections.IEnumerable]) {
                    foreach ($item in $current) { $stack.Push($item) }
                    continue
                }
                $list.Add([string]$current)
            }
            return , $list.ToArray()
        }

        function Get-TestSecretHits {
            # Teksty wyniku (także JSON całego wyniku), w których jest któraś z podanych wartości.
            param($Result, [string[]]$Secrets)
            $texts = New-Object System.Collections.Generic.List[string]
            foreach ($text in (Get-TestAllStrings -Node $Result)) { $texts.Add($text) }
            $texts.Add((ConvertTo-StdJson -InputObject ([ordered]@{ d = $Result.Document; w = $Result.Warnings; s = $Result.SecretsToSet }) -Compress))
            $hits = New-Object System.Collections.Generic.List[string]
            foreach ($secret in $Secrets) {
                foreach ($text in $texts) {
                    if ($text.IndexOf($secret, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $hits.Add($secret); break }
                }
            }
            return , $hits.ToArray()
        }

        function Test-TestPolishSentence {
            # Zdanie po polsku: wielka litera na początku, kropka na końcu, polskie znaki, bez typowych słów angielskich.
            param([string]$Text)
            return (($Text -cmatch '^[A-ZĄĆĘŁŃÓŚŹŻ]') -and ($Text -match '\.$') -and ($Text -match '[ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]') -and
                ($Text -notmatch '\b(the|skipped|skipping|warning|missing|invalid|unknown|not found|field)\b'))
        }

        function Get-TestUnexpectedReleaseErrors {
            # Błędy Test-StdRelease -Unpublished bez tych, które SPEC 3.5 dopuszcza po konwersji:
            # brakujące sekrety (z SecretsToSet) i opcjonalnie adresy http:// (z ostrzeżeń).
            param($Result, [switch]$AllowHttp)
            $errors = @(Test-StdRelease -Release $Result.Document -Unpublished)
            $left = New-Object System.Collections.Generic.List[string]
            foreach ($validationError in $errors) {
                $text = [string]$validationError
                $allowed = $text -match '(?i)secret|sekret'
                foreach ($name in @($Result.SecretsToSet)) { if ($text.IndexOf($name, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $allowed = $true } }
                if ($AllowHttp -and $text -match '(?i)http') { $allowed = $true }
                if (-not $allowed) { $left.Add($text) }
            }
            return , $left.ToArray()
        }
    }

    AfterAll {
        Remove-Item -Path Env:\STD_DATA_ROOT -ErrorAction SilentlyContinue
    }

    Context 'Przykład STD 3.1 (config.example.json)' {
        BeforeAll {
            $result = ConvertFrom-StdLegacyConfig -Path $examplePath -ClientId 'lizard' -ClientName 'Lizard-Tech'
            $doc = $result.Document
            $site = $doc['sites'][0]
        }

        It 'plik przykładu ma BOM i CRLF (przypadek brzegowy odczytu)' {
            $bytes = [System.IO.File]::ReadAllBytes($examplePath)
            ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should -BeTrue
            ([System.Text.Encoding]::UTF8.GetString($bytes)).Contains("`r`n") | Should -BeTrue
        }

        It 'zwraca dokument v4 bez revision i publishedAt' {
            $result.Keys | Sort-Object | Should -Be @('Document', 'SecretsToSet', 'Warnings')
            $doc['schemaVersion'] | Should -Be 4
            $doc['kind'] | Should -BeExactly 'std.release'
            $doc['client']['id'] | Should -BeExactly 'lizard'
            $doc['client']['name'] | Should -BeExactly 'Lizard-Tech'
            $doc.Contains('revision') | Should -BeFalse
            $doc.Contains('publishedAt') | Should -BeFalse
            $doc.Contains('secrets') | Should -BeFalse
            @($doc['sites']).Count | Should -Be 1
            $site['id'] | Should -BeExactly 'default'
            $site['name'] | Should -BeExactly 'Domyślna'
        }

        It 'zwraca Warnings i SecretsToSet jako string[]' {
            ($result.Warnings -is [string[]]) | Should -BeTrue
            ($result.SecretsToSet -is [string[]]) | Should -BeTrue
            $empty = Convert-TestLegacy -Json '{ "InstallSourcePaths": { "network": "\\\\srv\\x\\" }, "DefaultInstallSource": "network", "Programs": { "A": { "FileName": "a.exe" } } }'
            ($empty.Warnings -is [string[]]) | Should -BeTrue
            $empty.Warnings.Count | Should -Be 0
            ($empty.SecretsToSet -is [string[]]) | Should -BeTrue
            $empty.SecretsToSet.Count | Should -Be 0
        }

        It 'mapuje InstallSourcePaths na źródła lan i web oraz DefaultInstallSource na kolejność źródeł' {
            $doc['sources']['lan']['type'] | Should -BeExactly 'unc'
            $doc['sources']['lan']['path'] | Should -BeExactly '\\SERVER\share\'
            $doc['sources']['lan']['auth'] | Should -BeExactly 'session'
            $doc['sources']['web']['type'] | Should -BeExactly 'https'
            $doc['sources']['web']['url'] | Should -BeExactly 'https://example.com/apps/'
            $doc['sources']['web']['auth'] | Should -BeExactly 'session'
            $site['sourceOrder'] | Should -Be @('web', 'lan')
        }

        It 'przenosi programy do katalogu (plik, argumenty, winget)' {
            $chrome = $doc['catalog']['Chrome']
            $chrome['name'] | Should -BeExactly 'Chrome'
            $chrome['file'] | Should -BeExactly 'ChromeSetup.exe'
            $chrome['args'] | Should -BeExactly '/silent /install'
            $chrome['winget'] | Should -BeExactly 'Google.Chrome'
            $doc['catalog']['PowerShell_7']['wingetArgs'] | Should -BeExactly '--installer-type wix --scope machine'
            $doc['catalog']['LanSweeper_Agent'].Contains('winget') | Should -BeFalse
            $thunderbird = $doc['catalog']['Mozilla-Thunderbird']
            $thunderbird['name'] | Should -BeExactly 'Mozilla Thunderbird'
            $thunderbird['winget'] | Should -BeExactly 'Mozilla.Thunderbird.pl'
            $thunderbird.Contains('file') | Should -BeFalse
            $thunderbird.Contains('args') | Should -BeFalse
            $doc['catalog']['NET-Desktop-Runtime-10']['name'] | Should -BeExactly '.NET Desktop Runtime 10'
            $doc['catalog']['Visual-C++-Redistributable-x64']['winget'] | Should -BeExactly 'Microsoft.VCRedist.2015+.x64'
        }

        It 'każdy klucz katalogu pasuje do wzorca i każda pozycja ma file, url albo winget' {
            foreach ($entry in $doc['catalog'].GetEnumerator()) {
                $entry.Key | Should -Match $catalogKeyPattern
                ($entry.Value.Contains('file') -or $entry.Value.Contains('url') -or $entry.Value.Contains('winget')) | Should -BeTrue
                [string]::IsNullOrWhiteSpace($entry.Value['name']) | Should -BeFalse
            }
        }

        It 'antywirus: pozycja AV z własnymi źródłami av-lan i av-web' {
            $av = $doc['catalog']['AV']
            $av['name'] | Should -BeExactly 'Antywirus'
            $av['file'] | Should -BeExactly 'AntyVirus.exe'
            $av['sourceOrder'] | Should -Be @('av-lan', 'av-web')
            $doc['sources']['av-lan']['path'] | Should -BeExactly '\\SERVER\AV\'
            $doc['sources']['av-web']['url'] | Should -BeExactly 'https://example.com/av/'
            ($result.Warnings -like '*AntyVirus.Credentials*').Count | Should -Be 1
        }

        It 'TeamViewer: APITOKEN i CUSTOMCONFIGID zamienione na wstawki sekretów' {
            $tv = $doc['catalog']['TeamViewer']
            $tv['file'] | Should -BeExactly 'TeamViewer_Host.msi'
            $tv['args'] | Should -BeExactly '/qn CUSTOMCONFIGID={{secret:teamviewer-customconfigid}} APITOKEN={{secret:teamviewer-apitoken}} ASSIGNMENTOPTIONS="--alias %COMPUTERNAME% --grant-easy-access"'
            @($result.SecretsToSet | Sort-Object) | Should -Be @('teamviewer-apitoken', 'teamviewer-customconfigid')
        }

        It 'profil Domyślny ma programy z Enabled = true, a TeamViewer i AV są we wszystkich profilach' {
            $profiles = $site['profiles']
            @($profiles.Keys) | Should -Be @('Domyślny', 'Standard', 'Księgowość', 'Test')
            $default = @($profiles['Domyślny']['apps'])
            $default[0] | Should -BeExactly 'TeamViewer'
            $default[1] | Should -BeExactly 'AV'
            $default | Should -Contain 'Chrome'
            $default | Should -Contain 'Szafir_KIR'
            $default | Should -Not -Contain 'VLC'
            $default | Should -Not -Contain 'Mozilla-Thunderbird'
            @($profiles['Standard']['apps']) | Should -Be @('TeamViewer', 'AV', 'Chrome', '7zip', 'PowerToys')
            @($profiles['Księgowość']['apps']) | Should -Be @('TeamViewer', 'AV', 'Szafir_KIR', 'AdobeReader', '7zip', 'Chrome', 'DropBox')
            foreach ($profileEntry in $profiles.GetEnumerator()) {
                foreach ($app in @($profileEntry.Value['apps'])) { $doc['catalog'].Contains($app) | Should -BeTrue }
            }
        }

        It 'DefaultCheckboxes -> defaults.tasks tylko z identyfikatorów STD 4' {
            $tasks = $doc['defaults']['tasks']
            $tasks['SuspendSleep'] | Should -BeTrue
            $tasks['ConnectWiFi'] | Should -BeFalse
            $tasks['SystemSettings'] | Should -BeTrue
            $tasks['InstallApps'] | Should -BeTrue
            $tasks['RenameComputer'] | Should -BeFalse
            $tasks['UninstallM365'] | Should -BeFalse
            $tasks['JoinDomain'] | Should -BeTrue
            foreach ($key in @($tasks.Keys)) { $taskIds | Should -Contain $key }
            $tasks.Contains('RunWindowsUpdate') | Should -BeFalse
            $tasks.Contains('InstallTeamViewer') | Should -BeFalse
            ($result.Warnings -like '*RunWindowsUpdate*').Count | Should -Be 1
        }

        It 'domena i administrator lokalny bez loginu i hasła' {
            @($site['domainJoin'].Keys) | Should -Be @('domain')
            $site['domainJoin']['domain'] | Should -BeExactly 'firma.local'
            @($site['localAdmin'].Keys) | Should -Be @('username')
            $site['localAdmin']['username'] | Should -BeExactly 'LocalAdmin'
            ($result.Warnings -like '*DomainJoin.Username*').Count | Should -Be 1
            ($result.Warnings -like '*WebAuth*').Count | Should -Be 1
        }

        It 'Wi-Fi, ustawienia systemu i bloatware' {
            @($site['wifi']).Count | Should -Be 1
            $site['wifi'][0]['file'] | Should -BeExactly 'WiFiProfile.xml'
            ($result.Warnings -like '*WiFiProfile.xml*').Count | Should -Be 1
            $settings = $doc['defaults']['systemSettings']
            $settings['DisableFastStartup'] | Should -BeTrue
            $settings['EnableWin10StartMenu'] | Should -BeTrue
            $settings.Contains('CustomRegistry') | Should -BeFalse
            $settings.Contains('registry') | Should -BeFalse
            @($doc['defaults']['bloatware']) | Should -Contain '*TikTok*'
        }

        It 'żadna wartość hasła ani tokenu z 3.1 nie trafia do wyniku' {
            Get-TestSecretHits -Result $result -Secrets @('1234567', '1234567-example', 'ad\administrator', 'REDACTED') | Should -BeNullOrEmpty
        }

        It 'wszystkie ostrzeżenia są zdaniami po polsku' {
            $result.Warnings.Count | Should -BeGreaterThan 0
            foreach ($warning in $result.Warnings) { Test-TestPolishSentence -Text $warning | Should -BeTrue -Because $warning }
        }

        It 'przechodzi Test-StdRelease -Unpublished poza błędami o brakujących sekretach' {
            (Get-TestUnexpectedReleaseErrors -Result $result) -join ' | ' | Should -BeExactly ''
        }

        It 'dokument da się zapisać jako JSON i wczytać bez zmian' {
            $json = ConvertTo-StdJson -InputObject $doc
            $again = ConvertTo-StdJson -InputObject (ConvertFrom-StdJson -Json $json)
            $again | Should -BeExactly $json
        }

        It '-Path, -InputObject (obiekt z ConvertFrom-Json) i tekst JSON dają ten sam wynik' {
            $text = [System.IO.File]::ReadAllText($examplePath, [System.Text.Encoding]::UTF8)
            $fromObject = ConvertFrom-StdLegacyConfig -InputObject ($text.TrimStart([char]0xFEFF) | ConvertFrom-Json) -ClientId 'lizard' -ClientName 'Lizard-Tech'
            $fromText = ConvertFrom-StdLegacyConfig -InputObject $text -ClientId 'lizard' -ClientName 'Lizard-Tech'
            $expected = ConvertTo-StdJson -InputObject $doc -Compress
            ConvertTo-StdJson -InputObject $fromObject.Document -Compress | Should -BeExactly $expected
            ConvertTo-StdJson -InputObject $fromText.Document -Compress | Should -BeExactly $expected
            ($fromObject.Warnings -join '|') | Should -BeExactly ($result.Warnings -join '|')
        }
    }

    Context 'Źródło domyślne winget (identyfikator w FileName)' {
        BeforeAll {
            $json = @'
{
  "DefaultInstallSource": "winget",
  "InstallSourcePaths": { "network": "", "web": "https://pliki.example.com/std/" },
  "Programs": {
    "Stary": { "Enabled": true, "FileName": "Vendor.Stary", "SilentArgs": "--scope machine" },
    "ZArgumentami": { "Enabled": true, "FileName": "Vendor.Args", "SilentArgs": "/S", "WingetArgs": "--scope user" },
    "Oba": { "Enabled": false, "FileName": "oba.exe", "SilentArgs": "/S", "WingetId": "Vendor.Oba" },
    "TylkoWinget": { "Enabled": true, "FileName": "", "WingetId": "Vendor.Only" },
    "Pusty": { "Enabled": true, "FileName": "", "SilentArgs": "" }
  },
  "TeamViewer": { "FileName": "TeamViewer.TeamViewer.Host", "Arguments": "APITOKEN=tv-SEKRET-123 CUSTOMCONFIGID=cfg7777" },
  "AntyVirus": { "FileName": "ESET.Security", "InstallSourcePaths": { "network": "", "web": "" } },
  "DefaultCheckboxes": { "InstallTeamViewer": true, "InstallAV": true, "InstallApplications": true }
}
'@
            $result = Convert-TestLegacy -Json $json
            $catalog = $result.Document['catalog']
        }

        It 'FileName bez WingetId jest identyfikatorem winget, argumenty z WingetArgs albo SilentArgs' {
            $catalog['Stary']['winget'] | Should -BeExactly 'Vendor.Stary'
            $catalog['Stary']['wingetArgs'] | Should -BeExactly '--scope machine'
            $catalog['Stary'].Contains('file') | Should -BeFalse
            $catalog['Stary'].Contains('args') | Should -BeFalse
            $catalog['ZArgumentami']['winget'] | Should -BeExactly 'Vendor.Args'
            $catalog['ZArgumentami']['wingetArgs'] | Should -BeExactly '--scope user'
        }

        It 'program z FileName i WingetId ma plik i winget (z ostrzeżeniem o kolejności)' {
            $catalog['Oba']['file'] | Should -BeExactly 'oba.exe'
            $catalog['Oba']['args'] | Should -BeExactly '/S'
            $catalog['Oba']['winget'] | Should -BeExactly 'Vendor.Oba'
            ($result.Warnings -like '*winget*Dotyczy: Oba.').Count | Should -Be 1
            $catalog['TylkoWinget']['winget'] | Should -BeExactly 'Vendor.Only'
        }

        It 'program bez pliku i identyfikatora jest pominięty z ostrzeżeniem' {
            $catalog.Contains('Pusty') | Should -BeFalse
            ($result.Warnings -like '*„Pusty”*Pominięto go.').Count | Should -Be 1
        }

        It 'TeamViewer i antywirus przez winget, sekrety w wingetArgs' {
            $catalog['TeamViewer']['winget'] | Should -BeExactly 'TeamViewer.TeamViewer.Host'
            $catalog['TeamViewer']['wingetArgs'] | Should -BeExactly 'APITOKEN={{secret:teamviewer-apitoken}} CUSTOMCONFIGID={{secret:teamviewer-customconfigid}}'
            $catalog['TeamViewer'].Contains('file') | Should -BeFalse
            ($result.Warnings -like '*wingetArgs*').Count | Should -Be 1
            $catalog['AV']['winget'] | Should -BeExactly 'ESET.Security'
            $catalog['AV'].Contains('file') | Should -BeFalse
            $catalog['AV'].Contains('sourceOrder') | Should -BeFalse
            $result.Document['sources'].Contains('av-web') | Should -BeFalse
        }

        It 'kolejność źródeł zawiera tylko istniejące źródła' {
            @($result.Document['sources'].Keys) | Should -Be @('web')
            $result.Document['sites'][0]['sourceOrder'] | Should -Be @('web')
        }

        It 'tokeny TeamViewer nie trafiają do wyniku' {
            Get-TestSecretHits -Result $result -Secrets @('tv-SEKRET-123', 'cfg7777') | Should -BeNullOrEmpty
            @($result.SecretsToSet) | Should -Be @('teamviewer-apitoken', 'teamviewer-customconfigid')
        }
    }

    Context 'Własny adres programu (DownloadUrl)' {
        BeforeAll {
            $json = @'
{
  "DefaultInstallSource": "network",
  "InstallSourcePaths": { "network": "\\\\srv\\instalki\\" },
  "Programs": {
    "Pobierany": { "Enabled": true, "FileName": "lokalna-nazwa.exe", "SilentArgs": "/quiet", "DownloadUrl": "https://dl.example.com/app.exe", "WingetId": "Vendor.App" },
    "StaryAdres": { "Enabled": true, "DownloadUrl": "http://dl.example.com/stary.msi" },
    "ZHaslem": { "Enabled": true, "DownloadUrl": "https://jan:Password123!@dl.example.com/x.exe" },
    "ZHaslemHash": { "Enabled": true, "DownloadUrl": "https://jan:Ha#slo@99@dl.example.com/y.exe?v=1" }
  }
}
'@
            $result = Convert-TestLegacy -Json $json
            $catalog = $result.Document['catalog']
        }

        It 'DownloadUrl -> url (zamiast file) z argumentami, WingetId -> winget' {
            $catalog['Pobierany']['url'] | Should -BeExactly 'https://dl.example.com/app.exe'
            $catalog['Pobierany']['args'] | Should -BeExactly '/quiet'
            $catalog['Pobierany']['winget'] | Should -BeExactly 'Vendor.App'
            $catalog['Pobierany'].Contains('file') | Should -BeFalse
        }

        It 'adres http:// zostaje w url z ostrzeżeniem' {
            $catalog['StaryAdres']['url'] | Should -BeExactly 'http://dl.example.com/stary.msi'
            ($result.Warnings -like '*„StaryAdres”*http://*').Count | Should -Be 1
        }

        It 'login i hasło w adresie są usuwane' {
            $catalog['ZHaslem']['url'] | Should -BeExactly 'https://dl.example.com/x.exe'
            $catalog['ZHaslemHash']['url'] | Should -BeExactly 'https://dl.example.com/y.exe?v=1'
            ($result.Warnings -like '*Programs.ZHaslem.DownloadUrl*login i hasło*').Count | Should -Be 1
            Get-TestSecretHits -Result $result -Secrets @('Password123!', 'Ha#slo@99', 'jan:') | Should -BeNullOrEmpty
        }

        It 'Test-StdRelease odrzuca tylko adres http://' {
            $errors = @(Test-StdRelease -Release $result.Document -Unpublished)
            @($errors | Where-Object { $_ -match '(?i)http' }).Count | Should -BeGreaterThan 0
            (Get-TestUnexpectedReleaseErrors -Result $result -AllowHttp) -join ' | ' | Should -BeExactly ''
        }
    }

    Context 'Źródło web z adresem http://' {
        It 'adres zostaje w sources.web.url z ostrzeżeniem, a walidacja zgłasza tylko ten adres' {
            $result = Convert-TestLegacy -Json '{ "DefaultInstallSource": "web", "InstallSourcePaths": { "web": "http://pliki.example.com/" }, "Programs": { "A": { "Enabled": true, "FileName": "a.exe" } } }'
            $result.Document['sources']['web']['url'] | Should -BeExactly 'http://pliki.example.com/'
            ($result.Warnings -like '*InstallSourcePaths.web*http://*').Count | Should -Be 1
            @(Test-StdRelease -Release $result.Document -Unpublished).Count | Should -BeGreaterThan 0
            (Get-TestUnexpectedReleaseErrors -Result $result -AllowHttp) -join ' | ' | Should -BeExactly ''
        }
    }

    Context 'Antywirus z własnymi źródłami' {
        It 'własna ścieżka web antywirusa -> av-web, brak własnej ścieżki sieciowej -> globalna lan' {
            $result = Convert-TestLegacy -Json @'
{
  "DefaultInstallSource": "network",
  "InstallSourcePaths": { "network": "\\\\srv\\instalki\\", "web": "https://pliki.example.com/" },
  "AntyVirus": { "DefaultInstallSource": "web", "InstallSourcePaths": { "network": "", "web": "https://av.example.com/pakiety/" }, "FileName": "av.msi" }
}
'@
            $result.Document['sources']['av-web']['url'] | Should -BeExactly 'https://av.example.com/pakiety/'
            $result.Document['sources'].Contains('av-lan') | Should -BeFalse
            $result.Document['catalog']['AV']['sourceOrder'] | Should -Be @('av-web', 'lan')
            $result.Document['sites'][0]['sourceOrder'] | Should -Be @('lan', 'web')
        }

        It 'ta sama ścieżka co globalna -> bez nowego źródła i bez własnej kolejności' {
            $result = Convert-TestLegacy -Json @'
{
  "DefaultInstallSource": "network",
  "InstallSourcePaths": { "network": "\\\\srv\\instalki\\", "web": "https://pliki.example.com/" },
  "AntyVirus": { "InstallSourcePaths": { "network": "\\\\SRV\\instalki" }, "FileName": "av.exe" }
}
'@
            @($result.Document['sources'].Keys) | Should -Be @('lan', 'web')
            $result.Document['catalog']['AV'].Contains('sourceOrder') | Should -BeFalse
        }

        It 'Credentials antywirusa są pomijane, a hasło nie trafia do wyniku' {
            $result = Convert-TestLegacy -Json @'
{
  "DefaultInstallSource": "web",
  "InstallSourcePaths": { "web": "https://pliki.example.com/" },
  "AntyVirus": { "DefaultInstallSource": "network", "InstallSourcePaths": { "network": "\\\\av-srv\\AV$\\" }, "Credentials": { "Username": "avadmin", "Password": "Password123!" }, "FileName": "av.exe" }
}
'@
            $result.Document['catalog']['AV']['sourceOrder'] | Should -Be @('av-lan', 'web')
            ($result.Warnings -like '*AntyVirus.Credentials*').Count | Should -Be 1
            Get-TestSecretHits -Result $result -Secrets @('Password123!', 'avadmin') | Should -BeNullOrEmpty
        }
    }

    Context 'Rejestr (SystemSettings.CustomRegistry)' {
        BeforeAll {
            $json = @'
{
  "InstallSourcePaths": { "network": "\\\\srv\\instalki\\" },
  "DefaultInstallSource": "network",
  "Programs": { "A": { "Enabled": true, "FileName": "a.exe" } },
  "SystemSettings": {
    "DisableTelemetry": true,
    "Nieznane": "tak",
    "CustomRegistry": [
      { "Path": "HKLM:\\SOFTWARE\\Firma", "Name": "Wersja", "Value": "1", "PropertyType": "DWord" },
      { "Path": "HKCU:\\Software\\Firma\\Ustawienia\\", "Name": "Motyw", "Value": "ciemny", "PropertyType": "" },
      { "Path": "HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Firma", "Name": "Limit", "Value": "0x10", "PropertyType": "qword" },
      { "Path": "HKCR:\\.txt", "Name": "Opis", "Value": "x", "PropertyType": "String" },
      { "Path": "HKLM:\\SOFTWARE\\Firma", "Name": "", "Value": "x", "PropertyType": "String" },
      { "Path": "HKLM:\\SOFTWARE\\Firma", "Name": "Dziwny", "Value": "x", "PropertyType": "Unknown" },
      { "Path": "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon", "Name": "DefaultPassword", "Value": "Password123!", "PropertyType": "String" }
    ]
  }
}
'@
            $result = Convert-TestLegacy -Json $json
            $settings = $result.Document['defaults']['systemSettings']
            $registry = @($settings['registry'])
        }

        It 'hive z prefiksu, ścieżka bez prefiksu, typ i wartość' {
            $registry.Count | Should -Be 3
            $registry[0]['hive'] | Should -BeExactly 'HKLM'
            $registry[0]['path'] | Should -BeExactly 'SOFTWARE\Firma'
            $registry[0]['name'] | Should -BeExactly 'Wersja'
            $registry[0]['type'] | Should -BeExactly 'DWord'
            $registry[0]['value'] | Should -Be 1
            $registry[0]['value'] | Should -BeOfType [int]
            $registry[1]['hive'] | Should -BeExactly 'HKCU'
            $registry[1]['path'] | Should -BeExactly 'Software\Firma\Ustawienia'
            $registry[1]['type'] | Should -BeExactly 'String'
            $registry[1]['value'] | Should -BeExactly 'ciemny'
            $registry[2]['hive'] | Should -BeExactly 'HKLM'
            $registry[2]['path'] | Should -BeExactly 'SOFTWARE\Policies\Firma'
            $registry[2]['type'] | Should -BeExactly 'QWord'
            $registry[2]['value'] | Should -Be 16
        }

        It 'nieobsługiwana gałąź, brak nazwy, nieznany typ i wpis z hasłem są pominięte z ostrzeżeniem' {
            ($result.Warnings -like '*HKCR:\.txt\Opis*').Count | Should -Be 1
            ($result.Warnings -like '*nr 5*').Count | Should -Be 1
            ($result.Warnings -like '*Dziwny*Unknown*').Count | Should -Be 1
            ($result.Warnings -like '*DefaultPassword*hasło lub token*').Count | Should -Be 1
            ($result.Warnings -like 'Wpisy rejestru HKCU*DefaultUser*').Count | Should -Be 1
            Get-TestSecretHits -Result $result -Secrets @('Password123!') | Should -BeNullOrEmpty
        }

        It 'flagi logiczne trafiają do systemSettings, inne wartości nie' {
            $settings['DisableTelemetry'] | Should -BeTrue
            $settings.Contains('Nieznane') | Should -BeFalse
            ($result.Warnings -like '*„Nieznane”*').Count | Should -Be 1
            $settings.Contains('CustomRegistry') | Should -BeFalse
        }

        It 'wpisy rejestru przechodzą Test-StdRelease' {
            (Get-TestUnexpectedReleaseErrors -Result $result) -join ' | ' | Should -BeExactly ''
        }
    }

    Context 'Profil Wi-Fi (WiFiProfile.FileName)' {
        It 'tekst -> jedna pozycja { file }' {
            $result = Convert-TestLegacy -Json '{ "WiFiProfile": { "FileName": "Biuro.xml" } }'
            @($result.Document['sites'][0]['wifi']).Count | Should -Be 1
            $result.Document['sites'][0]['wifi'][0]['file'] | Should -BeExactly 'Biuro.xml'
            ($result.Warnings -like '*„Biuro.xml”*').Count | Should -Be 1
        }

        It 'tablica -> pozycja dla każdego pliku (puste pominięte)' {
            $result = Convert-TestLegacy -Json '{ "WiFiProfile": { "FileName": [ "Biuro.xml", "", "Goscie.xml" ] } }'
            $wifi = @($result.Document['sites'][0]['wifi'])
            $wifi.Count | Should -Be 2
            $wifi[0]['file'] | Should -BeExactly 'Biuro.xml'
            $wifi[1]['file'] | Should -BeExactly 'Goscie.xml'
            ($result.Warnings -like '*Profil Wi-Fi*').Count | Should -Be 2
        }

        It 'pusta nazwa pliku -> bez wifi i bez ostrzeżenia' {
            $result = Convert-TestLegacy -Json '{ "WiFiProfile": { "FileName": "" } }'
            $result.Document['sites'][0].Contains('wifi') | Should -BeFalse
            ($result.Warnings -like '*Wi-Fi*').Count | Should -Be 0
        }
    }

    Context 'Brakujące sekcje' {
        It 'pusta konfiguracja daje szkielet dokumentu z ostrzeżeniami' {
            $result = Convert-TestLegacy -Json '{}'
            $doc = $result.Document
            $doc['sources'].Count | Should -Be 0
            $doc['catalog'].Count | Should -Be 0
            $site = $doc['sites'][0]
            $site['id'] | Should -BeExactly 'default'
            @($site['sourceOrder']).Count | Should -Be 0
            @($site['profiles'].Keys) | Should -Be @('Domyślny')
            @($site['profiles']['Domyślny']['apps']).Count | Should -Be 0
            $site.Contains('domainJoin') | Should -BeFalse
            $site.Contains('localAdmin') | Should -BeFalse
            $doc['defaults'].Contains('tasks') | Should -BeFalse
            ($result.Warnings -like '*InstallSourcePaths*').Count | Should -Be 1
            ($result.Warnings -like '*żadnego programu*').Count | Should -Be 1
            ($result.Warnings -like '*DefaultInstallSource*').Count | Should -Be 1
            @($result.SecretsToSet).Count | Should -Be 0
            foreach ($warning in $result.Warnings) { Test-TestPolishSentence -Text $warning | Should -BeTrue -Because $warning }
        }

        It 'puste sekcje domyślnej konfiguracji 3.1 nie dają ostrzeżeń' {
            $result = Convert-TestLegacy -Json @'
{
  "DefaultInstallSource": "network",
  "InstallSourcePaths": { "network": "\\\\server\\share\\", "web": "" },
  "WebAuth": { "Username": "", "Password": "" },
  "TeamViewer": { "FileName": "", "Arguments": "" },
  "AntyVirus": { "DefaultInstallSource": "network", "InstallSourcePaths": { "network": "", "web": "" }, "Credentials": { "Username": "", "Password": "" }, "FileName": "" },
  "AutoUpdate": { "Enabled": false, "VersionCheckPath": "" },
  "CustomWebDataLocation": { "URL": "" },
  "DarkTheme": true,
  "Programs": { "A": { "Enabled": true, "FileName": "a.exe" } }
}
'@
            $result.Warnings.Count | Should -Be 0
            $result.Document['catalog'].Contains('TeamViewer') | Should -BeFalse
            $result.Document['catalog'].Contains('AV') | Should -BeFalse
        }

        It 'zadania włączone bez danych w konfiguracji dają ostrzeżenia' {
            $result = Convert-TestLegacy -Json '{ "DefaultCheckboxes": { "JoinDomain": true, "CreateLocalAdmin": true, "InstallTeamViewer": true } }'
            ($result.Warnings -like '*JoinDomain*DomainJoin.DomainName*').Count | Should -Be 1
            ($result.Warnings -like '*CreateLocalAdmin*LocalAdmin.Username*').Count | Should -Be 1
            ($result.Warnings -like '*InstallTeamViewer*nie przeniesiono*').Count | Should -Be 1
        }
    }

    Context 'Nazwy programów i kolizje kluczy' {
        BeforeAll {
            $longName = 'Bardzo długa nazwa programu ' + ('x' * 60)
            $json = @'
{
  "DefaultInstallSource": "network",
  "InstallSourcePaths": { "network": "\\\\srv\\instalki\\" },
  "Programs": {
    "My App": { "Enabled": true, "FileName": "myapp.exe" },
    "My-App": { "Enabled": false, "FileName": "myapp2.exe" },
    "Visual C++ (x64)": { "FileName": "vc1.exe" },
    "Visual C++ x64": { "FileName": "vc2.exe" },
    "Zażółć gęślą jaźń": { "FileName": "pl.exe" },
    ".NET 8": { "WingetId": "Microsoft.DotNet.Runtime.8" },
    "TeamViewer": { "FileName": "tv-program.exe" },
    "AV": { "FileName": "av-program.exe" },
    "(((": { "FileName": "nawiasy.exe" },
    "LONGNAME": { "FileName": "long.exe" }
  },
  "TeamViewer": { "FileName": "TeamViewer_Host.msi", "Arguments": "/qn" },
  "AntyVirus": { "FileName": "av.msi" },
  "Profiles": { "Biuro": [ "My App", "Visual C++ x64", "TeamViewer", "AV", "Nieistniejący", "My App" ], "Jeden": "Zażółć gęślą jaźń" }
}
'@
            $json = $json.Replace('LONGNAME', $longName)
            $result = Convert-TestLegacy -Json $json
            $catalog = $result.Document['catalog']
            $keys = @($catalog.Keys)
        }

        It 'klucze pasują do wzorca, nazwy zostają oryginalne' {
            foreach ($key in $keys) { $key | Should -Match $catalogKeyPattern }
            $catalog['My-App']['file'] | Should -BeExactly 'myapp2.exe'
            $catalog['My-App-2']['name'] | Should -BeExactly 'My App'
            $catalog['Visual-C++-x64']['name'] | Should -BeExactly 'Visual C++ (x64)'
            $catalog['Visual-C++-x64-2']['name'] | Should -BeExactly 'Visual C++ x64'
            $catalog['Zazolc-gesla-jazn']['name'] | Should -BeExactly 'Zażółć gęślą jaźń'
            $catalog['NET-8']['name'] | Should -BeExactly '.NET 8'
            $catalog['program']['name'] | Should -BeExactly '((('
        }

        It 'TeamViewer i AV z sekcji 3.1 mają klucze TeamViewer i AV, programy o tych nazwach dostają sufiks' {
            $catalog['TeamViewer']['file'] | Should -BeExactly 'TeamViewer_Host.msi'
            $catalog['TeamViewer-2']['file'] | Should -BeExactly 'tv-program.exe'
            $catalog['AV']['file'] | Should -BeExactly 'av.msi'
            $catalog['AV-2']['file'] | Should -BeExactly 'av-program.exe'
        }

        It 'długa nazwa jest skracana do 64 znaków' {
            $longKey = @($keys | Where-Object { $_ -like 'Bardzo-dluga*' })
            $longKey.Count | Should -Be 1
            $longKey[0].Length | Should -Be 64
        }

        It 'kolizje kluczy są w ostrzeżeniach' {
            ($result.Warnings -like '*„My App”*„My-App-2”*').Count | Should -Be 1
            ($result.Warnings -like '*„Visual C++ x64”*„Visual-C++-x64-2”*').Count | Should -Be 1
            ($result.Warnings -like '*„TeamViewer”*„TeamViewer-2”*').Count | Should -Be 1
        }

        It 'profile wskazują nowe klucze, nieznane programy są pominięte z ostrzeżeniem' {
            $profiles = $result.Document['sites'][0]['profiles']
            @($profiles['Biuro']['apps']) | Should -Be @('My-App-2', 'Visual-C++-x64-2', 'TeamViewer-2', 'AV-2')
            @($profiles['Jeden']['apps']) | Should -Be @('Zazolc-gesla-jazn')
            @($profiles['Domyślny']['apps']) | Should -Be @('My-App-2')
            ($result.Warnings -like '*„Biuro”*„Nieistniejący”*').Count | Should -Be 1
        }

        It 'wynik przechodzi Test-StdRelease' {
            (Get-TestUnexpectedReleaseErrors -Result $result) -join ' | ' | Should -BeExactly ''
        }
    }

    Context 'Hasła i tokeny nie trafiają do wyniku' {
        BeforeAll {
            $json = @'
{
  "DefaultInstallSource": "network",
  "InstallSourcePaths": { "network": "\\\\srv\\instalki\\" },
  "WebAuth": { "Username": "jan.kowalski", "Password": "Password123!" },
  "DomainJoin": { "DomainName": "firma.local", "Username": "FIRMA\\wdrozenia", "Password": "Password123!" },
  "LocalAdmin": { "Username": "LokalnyIT", "Password": "Password123!" },
  "TeamViewer": { "FileName": "TeamViewer_Host.msi", "Arguments": "/qn APITOKEN=\"12345678-AbCdEfGh\" CUSTOMCONFIGID=abc1234" },
  "AntyVirus": { "FileName": "av.msi", "Credentials": { "Username": "avadmin", "Password": "Password123!" } },
  "Programs": { "Narzędzie": { "Enabled": true, "FileName": "tool.exe", "SilentArgs": "/S /pass=Password123!" } },
  "DefaultCheckboxes": { "InstallTeamViewer": true, "JoinDomain": true, "CreateLocalAdmin": true }
}
'@
            $result = Convert-TestLegacy -Json $json
            $site = $result.Document['sites'][0]
        }

        It 'żadna wartość hasła ani tokenu nie występuje w dokumencie, ostrzeżeniach ani liście sekretów' {
            Get-TestSecretHits -Result $result -Secrets @('Password123!', '12345678-AbCdEfGh', 'abc1234', 'FIRMA\wdrozenia', 'wdrozenia', 'jan.kowalski', 'avadmin') | Should -BeNullOrEmpty
        }

        It 'sekrety TeamViewer jako wstawki i w SecretsToSet' {
            $result.Document['catalog']['TeamViewer']['args'] | Should -BeExactly '/qn APITOKEN="{{secret:teamviewer-apitoken}}" CUSTOMCONFIGID={{secret:teamviewer-customconfigid}}'
            @($result.SecretsToSet) | Should -Be @('teamviewer-apitoken', 'teamviewer-customconfigid')
        }

        It 'domainJoin i localAdmin bez pól z danymi logowania' {
            @($site['domainJoin'].Keys) | Should -Be @('domain')
            @($site['localAdmin'].Keys) | Should -Be @('username')
            ($result.Warnings -like '*DomainJoin.Password*').Count | Should -Be 1
            ($result.Warnings -like '*LocalAdmin.Password*').Count | Should -Be 1
            ($result.Warnings -like '*DomainJoin.Username*').Count | Should -Be 1
        }

        It 'hasło wpisane w argumenty programu jest zastąpione znakami *** z ostrzeżeniem' {
            $result.Document['catalog']['Narzedzie']['args'] | Should -BeExactly '/S /pass=***'
            ($result.Warnings -like '*catalog.Narzedzie.args*').Count | Should -Be 1
        }

        It 'Test-StdRelease nie zgłasza pól z danymi logowania' {
            (Get-TestUnexpectedReleaseErrors -Result $result) -join ' | ' | Should -BeExactly ''
        }

        It 'wejście nie jest zmieniane' {
            $source = ConvertFrom-Json -InputObject $json
            $null = ConvertFrom-StdLegacyConfig -InputObject $source -ClientId 'test' -ClientName 'Test'
            $source.TeamViewer.Arguments | Should -BeExactly '/qn APITOKEN="12345678-AbCdEfGh" CUSTOMCONFIGID=abc1234'
            $source.WebAuth.Password | Should -BeExactly 'Password123!'
        }
    }

    Context 'DefaultCheckboxes -> defaults.tasks' {
        BeforeAll {
            $json = @'
{
  "DefaultInstallSource": "network",
  "InstallSourcePaths": { "network": "\\\\srv\\instalki\\" },
  "TeamViewer": { "FileName": "TeamViewer_Host.msi" },
  "DefaultCheckboxes": {
    "WaitForNetwork": true, "SuspendHibernation": true, "ImportWiFiProfile": false, "RemoveBloatware": true,
    "UninstallMicrosoft365": true, "UninstallOneDrive": false, "InstallApplications": false, "CreateLocalAdmin": false,
    "ChangeComputerName": true, "JoinDomain": false, "ChangeSystemSettings": true, "RunPostInstallScripts": true,
    "ExportHardwareAudit": true, "EnableBitLocker": false, "AutoReboot": true,
    "DryRun": true, "CreateRestorePoint": true, "RunWindowsUpdate": false, "JoinIntune": true,
    "InstallTeamViewer": true, "InstallAV": false, "Nieznane": true
  }
}
'@
            $result = Convert-TestLegacy -Json $json
            $tasks = $result.Document['defaults']['tasks']
        }

        It 'mapuje wszystkie 15 zadań 3.1 na identyfikatory STD 4' {
            $tasks.Count | Should -Be 15
            $tasks['WaitForNetwork'] | Should -BeTrue
            $tasks['SuspendSleep'] | Should -BeTrue
            $tasks['ConnectWiFi'] | Should -BeFalse
            $tasks['RemoveBloatware'] | Should -BeTrue
            $tasks['UninstallM365'] | Should -BeTrue
            $tasks['UninstallOneDrive'] | Should -BeFalse
            $tasks['InstallApps'] | Should -BeFalse
            $tasks['CreateLocalAdmin'] | Should -BeFalse
            $tasks['RenameComputer'] | Should -BeTrue
            $tasks['JoinDomain'] | Should -BeFalse
            $tasks['SystemSettings'] | Should -BeTrue
            $tasks['RunScripts'] | Should -BeTrue
            $tasks['HardwareAudit'] | Should -BeTrue
            $tasks['EnableBitLocker'] | Should -BeFalse
            $tasks['RebootAtEnd'] | Should -BeTrue
            foreach ($key in @($tasks.Keys)) { $taskIds | Should -Contain $key }
            foreach ($value in @($tasks.Values)) { $value | Should -BeOfType [bool] }
        }

        It 'zadania usunięte z STD 4 dają ostrzeżenie tylko przy true' {
            ($result.Warnings -like '*DryRun*').Count | Should -Be 1
            ($result.Warnings -like '*CreateRestorePoint*').Count | Should -Be 1
            ($result.Warnings -like '*JoinIntune*').Count | Should -Be 1
            ($result.Warnings -like '*RunWindowsUpdate*').Count | Should -Be 0
            ($result.Warnings -like '*InstallAV*').Count | Should -Be 0
        }

        It 'nieznane zadanie daje ostrzeżenie' {
            ($result.Warnings -like '*„Nieznane”*').Count | Should -Be 1
        }

        It 'wartość "true" / "false" jako tekst jest przyjmowana, inna wartość daje ostrzeżenie' {
            $other = Convert-TestLegacy -Json '{ "DefaultCheckboxes": { "JoinDomain": "true", "AutoReboot": "False", "WaitForNetwork": "tak" } }'
            $otherTasks = $other.Document['defaults']['tasks']
            $otherTasks['JoinDomain'] | Should -BeTrue
            $otherTasks['RebootAtEnd'] | Should -BeFalse
            $otherTasks.Contains('WaitForNetwork') | Should -BeFalse
            ($other.Warnings -like '*„WaitForNetwork”*true ani false*').Count | Should -Be 1
        }

        It 'TeamViewer dopisany do profili i ostrzeżenie o wyłączonym InstallApps' {
            @($result.Document['sites'][0]['profiles']['Domyślny']['apps']) | Should -Be @('TeamViewer')
            ($result.Warnings -like '*InstallApplications*InstallApps*').Count | Should -Be 1
        }
    }

    Context 'Skrypty poinstalacyjne i profile' {
        It 'PostInstallScripts -> defaults.scripts = [ { file } ]' {
            $result = Convert-TestLegacy -Json '{ "PostInstallScripts": [ "Drukarki.ps1", "https://pliki.example.com/skrypty/Mapowanie%20dyskow.ps1", "", 5 ] }'
            $scripts = @($result.Document['defaults']['scripts'])
            $scripts.Count | Should -Be 2
            $scripts[0]['file'] | Should -BeExactly 'Drukarki.ps1'
            $scripts[1]['file'] | Should -BeExactly 'Mapowanie dyskow.ps1'
            ($result.Warnings -like '*„Mapowanie dyskow.ps1”*').Count | Should -Be 1
            ($result.Warnings -like '*nr 4 w PostInstallScripts*').Count | Should -Be 1
        }

        It 'profil 3.1 o nazwie Domyślny dostaje inną nazwę' {
            $result = Convert-TestLegacy -Json '{ "Programs": { "A": { "Enabled": true, "FileName": "a.exe" }, "B": { "FileName": "b.exe" } }, "Profiles": { "Domyślny": [ "B" ] } }'
            $profiles = $result.Document['sites'][0]['profiles']
            @($profiles.Keys) | Should -Be @('Domyślny', 'Domyślny (3.1)')
            @($profiles['Domyślny']['apps']) | Should -Be @('A')
            @($profiles['Domyślny (3.1)']['apps']) | Should -Be @('B')
            ($result.Warnings -like '*„Domyślny (3.1)”*').Count | Should -Be 1
        }

        It 'bez programów z Enabled = true zostają tylko profile 3.1' {
            $result = Convert-TestLegacy -Json '{ "Programs": { "A": { "Enabled": false, "FileName": "a.exe" } }, "Profiles": { "Standard": [ "A" ] } }'
            @($result.Document['sites'][0]['profiles'].Keys) | Should -Be @('Standard')
        }
    }

    Context 'Klucze nieznane i pominięte' {
        It 'nieznany klucz i nieznane pole programu dają ostrzeżenie (bez wartości)' {
            $result = Convert-TestLegacy -Json '{ "Nowosc": { "Haslo": "Password123!" }, "Programs": { "A": { "Enabled": true, "FileName": "a.exe", "Opis": "x" } }, "HardwareAudit": { "ExportPath": "C:\\Audit\\" }, "DarkTheme": false }'
            ($result.Warnings -like '*„Nowosc”*').Count | Should -Be 1
            ($result.Warnings -like '*„Programs.A.Opis”*').Count | Should -Be 1
            ($result.Warnings -like '*HardwareAudit*').Count | Should -Be 1
            ($result.Warnings -like '*DarkTheme*').Count | Should -Be 0
            Get-TestSecretHits -Result $result -Secrets @('Password123!') | Should -BeNullOrEmpty
        }
    }

    Context 'Błędy wejścia' {
        It 'nieprawidłowy identyfikator klienta -> [STD-LEGACY-ARG]' {
            { ConvertFrom-StdLegacyConfig -InputObject '{}' -ClientId 'Zły Klient' -ClientName 'X' } | Should -Throw -ExpectedMessage '*`[STD-LEGACY-ARG`]*'
            { ConvertFrom-StdLegacyConfig -InputObject '{}' -ClientId '' -ClientName 'X' } | Should -Throw -ExpectedMessage '*`[STD-LEGACY-ARG`]*'
        }

        It 'pusta nazwa klienta -> [STD-LEGACY-ARG]' {
            { ConvertFrom-StdLegacyConfig -InputObject '{}' -ClientId 'abc' -ClientName ' ' } | Should -Throw -ExpectedMessage '*`[STD-LEGACY-ARG`]*'
        }

        It 'brak pliku -> [STD-LEGACY-INPUT]' {
            { ConvertFrom-StdLegacyConfig -Path (Join-Path $TestDrive 'brak.json') -ClientId 'abc' -ClientName 'X' } | Should -Throw -ExpectedMessage '*`[STD-LEGACY-INPUT`]*'
        }

        It 'uszkodzony JSON -> [STD-JSON]' {
            $bad = Join-Path $TestDrive 'zly.json'
            [System.IO.File]::WriteAllText($bad, '{ "Programs": ', (New-Object System.Text.UTF8Encoding($false)))
            { ConvertFrom-StdLegacyConfig -Path $bad -ClientId 'abc' -ClientName 'X' } | Should -Throw -ExpectedMessage '*`[STD-JSON`]*'
        }

        It 'JSON, który nie jest obiektem -> [STD-LEGACY-INPUT]' {
            { ConvertFrom-StdLegacyConfig -InputObject '[ 1, 2 ]' -ClientId 'abc' -ClientName 'X' } | Should -Throw -ExpectedMessage '*`[STD-LEGACY-INPUT`]*'
            { ConvertFrom-StdLegacyConfig -InputObject $null -ClientId 'abc' -ClientName 'X' } | Should -Throw -ExpectedMessage '*`[STD-LEGACY-INPUT`]*'
        }

        It 'przyjmuje też hashtable' {
            $result = ConvertFrom-StdLegacyConfig -InputObject @{ Programs = @{ A = @{ Enabled = $true; FileName = 'a.exe' } } } -ClientId 'abc' -ClientName 'X'
            $result.Document['catalog']['A']['file'] | Should -BeExactly 'a.exe'
        }
    }

    Context 'Tools\Convert-StdLegacyConfig.ps1' {
        BeforeAll {
            $workDir = Join-Path $TestDrive 'narzedzie'
            [void][System.IO.Directory]::CreateDirectory($workDir)
            $configCopy = Join-Path $workDir 'config.json'
            [System.IO.File]::Copy($examplePath, $configCopy, $true)
        }

        It 'zapisuje dokument w -OutPath (UTF-8 bez BOM) bez haseł i tokenów' {
            $out = Join-Path $workDir 'wynik.json'
            & $toolPath -Path $configCopy -ClientId 'lizard' -ClientName 'Lizard-Tech' -OutPath $out 6> $null
            [System.IO.File]::Exists($out) | Should -BeTrue
            $bytes = [System.IO.File]::ReadAllBytes($out)
            ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should -BeFalse
            $text = [System.IO.File]::ReadAllText($out)
            $doc = ConvertFrom-StdJson -Json $text
            $doc['client']['id'] | Should -BeExactly 'lizard'
            $doc['catalog']['TeamViewer']['args'] | Should -BeLike '*{{secret:teamviewer-apitoken}}*'
            $text.Contains('1234567') | Should -BeFalse
            $text.Contains('administrator') | Should -BeFalse
        }

        It 'bez -OutPath zapisuje <ClientId>.release.v4.json obok config.json' {
            & $toolPath -Path $configCopy -ClientId 'lizard' -ClientName 'Lizard-Tech' 6> $null
            [System.IO.File]::Exists((Join-Path $workDir 'lizard.release.v4.json')) | Should -BeTrue
        }

        It 'nie zastępuje pliku config.json' {
            $before = [System.IO.File]::ReadAllBytes($configCopy)
            { & $toolPath -Path $configCopy -ClientId 'lizard' -ClientName 'Lizard-Tech' -OutPath $configCopy 6> $null } | Should -Throw -ExpectedMessage '*`[STD-LEGACY-ARG`]*'
            $after = [System.IO.File]::ReadAllBytes($configCopy)
            [System.Convert]::ToBase64String($after) | Should -BeExactly ([System.Convert]::ToBase64String($before))
        }
    }
}
