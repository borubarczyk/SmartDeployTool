#Requires -RunAsAdministrator
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13


$configPath = ".\config.json"
$script:ValidationWarnings = New-Object System.Collections.Generic.List[string]
$script:ValidationErrors   = New-Object System.Collections.Generic.List[string]
$SelectedApps = @{}
$CheckboxControls = @{}
$checkboxOptions = [ordered]@{
    "DisableHibernation"    = @{
        "Text"    = "Wyłącz hibernację"
        "Tooltip" = "Wyłącza funkcję hibernacji systemu Windows, na czas konfiguracji."
        "Enabled" = $true
    }
    "ImportWiFiProfile"     = @{
        "Text"    = "Importuj profil Wi-Fi"
        "Tooltip" = "Importuje zapisany profil Wi-Fi z pliku (Lizard-Tech)."
        "Enabled" = $false
    }
    "UninstallMicrosoft365" = @{
        "Text"    = "Odinstaluj preinstalowane produkty Microsoft 365"
        "Tooltip" = "Usuwa preinstalowane aplikacje Microsoft 365."
        "Enabled" = $false
    }
    "InstallTeamViewer"     = @{
        "Text"    = "Zainstaluj TeamViewer (jeśli jest przeinstaluje)"
        "Tooltip" = "Instaluje TeamViewer, jeśli jest już zainstalowany, odinstaluje i zainstaluje ponownie, jeśli jest uruchomiony QS zamknie proces i rozpocznie instalację."
        "Enabled" = $true
    }
    "InstallApplications"   = @{
        "Text"    = "Zainstaluj aplikacje (wybór)"
        "Tooltip" = "Instaluje wybrane aplikacje z listy."
        "Enabled" = $true
    }
    "CreateLocalAdmin"      = @{
        "Text"    = "Utwórz lokalnego administratora"
        "Tooltip" = "Tworzy lokalne konto administratora z niewygasającym hasłem, nazwa konta utworzy się na podstawie konfiguracji w pliku JSON."
        "Enabled" = $true
    }
    "InstallAV"             = @{
        "Text"    = "Zainstaluj Antywirusa"
        "Tooltip" = "Instaluje oprogramowanie antywirusowe (IN DEVELOPMENT)."
        "Enabled" = $true
    }
    "JoinDomain"            = @{
        "Text"    = "Dołącz do domeny"
        "Tooltip" = "Dołącza komputer do domeny - zgodnie z konfiguracją w pliku JSON - trzeba wpisać hasło do konta domenowego uprawnionego do tego."
        "Enabled" = $true
    }
    "ChangeSystemSettings"  = @{
        "Text"    = "Zmiany rejestru i ustawień systemowych"
        "Tooltip" = "Wprowadza zmiany w rejestrze i ustawieniach systemowych, zgodnie z konfiguracją w pliku JSON zmiana pobierania aktualizacji, ustawienia prywatności, wyłączenie Cortany, szybkiego uruchamiania, włącza stary widok menu kontekstowego itp."
        "Enabled" = $true
    }
    "RunWindowsUpdate"      = @{
        "Text"    = "Uruchom Windows Update"
        "Tooltip" = "Uruchamia usługę Windows Update po zakończeniu instalacji i sprawdza dostępność aktualizacji."
        "Enabled" = $true
    }
    "ChangeComputerName"    = @{
        "Text"    = "Zmień nazwę komputera"
        "Tooltip" = "Zmienia nazwę komputera na podstawie konfiguracji w pliku JSON i wprowadzonych danych wymaga ponownego uruchomienia."
        "Enabled" = $false
    }
}

function Test-ConfigurationFile {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red"
        return $false
    }
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        if ($null -eq $config.Programs) {
            Write-Log "Nie znaleziono sekcji 'Programs' w config.json" -Color "Red"
            return $false
        }
        return $true
    }
    catch {
        Write-Log "Błąd podczas odczytu config.json: $_" -Color "Red"
        return $false
    }
    
}

function Write-Log {
    param(
        [string]$Text,
        [string]$Color = "Black",
        [switch]$IsError
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    try {
        if ($null -ne $rtbLog -and $rtbLog -is [System.Windows.Forms.RichTextBox]) {
            $rtbLog.SelectionColor = $Color
            $rtbLog.AppendText("[$timestamp] $Text`r`n")
            $rtbLog.ScrollToCaret()
        }
    } catch { }
    try {
        Add-Content -Path "C:\deploy-log.txt" -Value "[$timestamp] $Text" -ErrorAction SilentlyContinue
        if ($IsError) {
            Add-Content -Path "C:\deploy-errors.txt" -Value "[$timestamp] $Text" -ErrorAction SilentlyContinue
        }
    } catch { }
    if ($IsError) { Write-Error $Text } else { Write-Host $Text }
}

function Test-UrlValid {
    param([Parameter(Mandatory)][string]$Url)
    try {
        $uri = $null
        if ([System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) {
            return ($uri.Scheme -in @('http','https'))
        }
        return $false
    } catch { return $false }
}

function Validate-BeforeRun {
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    try { $config = Get-Content $configPath -Raw | ConvertFrom-Json } catch { $config = $null }
    if ($null -eq $config) { $errors.Add("Brak lub niepoprawny config.json.") | Out-Null }

    if ($errors.Count -eq 0) {
        $source = $config.DefaultInstallSource
        $sourcePath = $config.InstallSourcePaths.$source

        switch ($source) {
            'network' {
                if ([string]::IsNullOrWhiteSpace($sourcePath)) { $errors.Add("Brak sciezki sieciowej (InstallSourcePaths.network).") | Out-Null }
                elseif ($sourcePath -notmatch '^\\\\') { $errors.Add("Sciezka sieciowa musi byc UNC (np. \\serwer\\udzial\\).") | Out-Null }
                elseif (-not (Test-Path -LiteralPath $sourcePath)) { $errors.Add("Nie znaleziono zasobu sieciowego: $sourcePath") | Out-Null }
            }
            'web' {
                if (-not (Test-UrlValid -Url $sourcePath)) { $errors.Add("Niepoprawny URL zrodla web: $sourcePath") | Out-Null }
            }
            default { $errors.Add("DefaultInstallSource musi byc 'network' lub 'web'.") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('InstallApplications') -and $CheckboxControls['InstallApplications'].Checked) {
            if ($SelectedApps.Count -eq 0) {
                $errors.Add("Brak wybranych aplikacji do instalacji.") | Out-Null
            } else {
                foreach ($appName in $SelectedApps.Keys) {
                    $app = $config.Programs.$appName
                    if ($null -eq $app) { $errors.Add("Brak konfiguracji dla aplikacji '$appName'.") | Out-Null; continue }
                    if ([string]::IsNullOrWhiteSpace($app.FileName)) { $errors.Add("Brak 'FileName' dla '$appName'.") | Out-Null; continue }
                    if ($source -eq 'network') {
                        $full = Join-Path $sourcePath $app.FileName
                        if (-not (Test-Path -LiteralPath $full)) { $errors.Add("Nie znaleziono pliku dla '$appName': $full") | Out-Null }
                    } elseif ($source -eq 'web') {
                        $url = if ($sourcePath[-1] -eq '/') { "$sourcePath$($app.FileName)" } else { "$sourcePath/$($app.FileName)" }
                        if (-not (Test-UrlValid -Url $url)) { $errors.Add("Niepoprawny URL dla '$appName': $url") | Out-Null }
                    }
                }
            }
        }

        if ($CheckboxControls.ContainsKey('InstallTeamViewer') -and $CheckboxControls['InstallTeamViewer'].Checked) {
            if (-not $config.TeamViewer -or [string]::IsNullOrWhiteSpace($config.TeamViewer.FileName)) {
                $errors.Add("TeamViewer nie ma poprawnie ustawionego FileName w config.json.") | Out-Null
            } elseif ($source -eq 'network') {
                $tvPath = Join-Path $sourcePath $config.TeamViewer.FileName
                if (-not (Test-Path -LiteralPath $tvPath)) { $errors.Add("Nie znaleziono instalatora TeamViewer: $tvPath") | Out-Null }
            } elseif ($source -eq 'web') {
                $tvUrl = if ($sourcePath[-1] -eq '/') { "$sourcePath$($config.TeamViewer.FileName)" } else { "$sourcePath/$($config.TeamViewer.FileName)" }
                if (-not (Test-UrlValid -Url $tvUrl)) { $errors.Add("Niepoprawny URL TeamViewer: $tvUrl") | Out-Null }
            }
        }

        if ($CheckboxControls.ContainsKey('InstallAV') -and $CheckboxControls['InstallAV'].Checked) {
            if (-not $config.AntyVirus -or [string]::IsNullOrWhiteSpace($config.AntyVirus.FileName)) { $errors.Add("Brak konfiguracji antywirusa lub FileName.") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('ImportWiFiProfile') -and $CheckboxControls['ImportWiFiProfile'].Checked) {
            $wifi = [string]$config.WiFiProfile.FileName
            if ([string]::IsNullOrWhiteSpace($wifi) -or -not (Test-Path -LiteralPath $wifi)) { $errors.Add("Brak pliku profilu Wi-Fi: $wifi") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('CreateLocalAdmin') -and $CheckboxControls['CreateLocalAdmin'].Checked) {
            if ([string]::IsNullOrWhiteSpace([string]$config.LocalAdmin.Username)) { $errors.Add("Brak LocalAdmin.Username w config.json.") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('JoinDomain') -and $CheckboxControls['JoinDomain'].Checked) {
            if (-not $config.DomainJoin -or [string]::IsNullOrWhiteSpace([string]$config.DomainJoin.DomainName)) { $errors.Add("Brak DomainJoin.DomainName w config.json.") | Out-Null }
            if ([string]::IsNullOrWhiteSpace([string]$config.DomainJoin.Username)) { $errors.Add("Brak DomainJoin.Username w config.json.") | Out-Null }
        }
    }

    foreach ($e in $errors) { Write-Log $e -Color 'Red' -IsError }
    foreach ($w in $warnings) { Write-Log $w -Color 'Yellow' }

    if ($errors.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(("Wykryto bledy walidacji:`r`n- " + ($errors -join "`r`n- ")), "Walidacja", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return $false
    }
    if ($warnings.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(("Uwaga:`r`n- " + ($warnings -join "`r`n- ")), "Walidacja", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
    return $true
}

function Install-SelectedApps {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red"
        return
    }

    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    $source = $config.DefaultInstallSource
    $sourcePath = $config.InstallSourcePaths.$source

    foreach ($appName in $SelectedApps.Keys) {
        $app = $config.Programs.$appName
        $fileName = $app.FileName
        $silentArgs = $app.SilentArgs
        $fullPath = if ($sourcePath -like "http*") {
            "$sourcePath$fileName"
        }
        else {
            Join-Path $sourcePath $fileName
        }

        $localPath = "$env:TEMP\$fileName"

        try {
            Write-Log "Pobieranie $appName"

            $iwrParams = @{
                Uri             = $fullPath
                OutFile         = $localPath
                UseBasicParsing = $true
                ErrorAction     = 'Stop'
            }

            if ($config.WebAuth.Username -and $config.WebAuth.Password) {
                $sec = ConvertTo-SecureString $config.WebAuth.Password -AsPlainText -Force
                $cred = [pscredential]::new($config.WebAuth.Username, $sec)
                $iwrParams.Credential = $cred
            }

            Invoke-WebRequest @iwrParams 

            Write-Log "Instalacja $appName..."

            if ($fileName -like "*.msi") {
                $args = "/i `"$localPath`" $silentArgs"
                Start-Process "msiexec.exe" -ArgumentList $args -Wait -NoNewWindow
            }
            else {
                Start-Process $localPath -ArgumentList $silentArgs -Wait -NoNewWindow
            }

            Write-Log "$appName zainstalowany." -Color "Green"
        }
        catch {
            Write-Log "Błąd przy $($appName): $_" -Color "Red"
        }

    }
}

function Disable-Hibernation {
    try {
        Write-Log "Wyłączanie hibernacji..."
        powercfg /h off
        Write-Log "Hibernacja wyłączona." -Color "Green"
    }
    catch { Write-Log "Błąd: $_" -Color "Red" }
}

function Install-TeamViewer {
    try {
        if (-not (Test-Path $configPath)) {
            Write-Log "Brak pliku config.json" -Color "Red" -IsError
            return
        }

        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        $source = $config.DefaultInstallSource
        $sourcePath = $config.InstallSourcePaths.$source
        $msiArgs = $config.TeamViewer.Arguments
        $fileName = $config.TeamViewer.FileName
        $localPath = "$env:TEMP\$fileName"

        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )


        $isInstalled = $false
        $uninstallString = $null

        foreach ($path in $regPaths) {
            $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                if ($item.DisplayName -like "*TeamViewer*") {
                    $isInstalled = $true
                    $uninstallString = $item.QuietUninstallString
                    if (-not $uninstallString) { $uninstallString = $item.UninstallString }

                    Write-Log "Znaleziono wpis TeamViewer: $($item.DisplayName)" -Color "Gray"
                    break
                }
            }
            if ($isInstalled) { break }
        }


        if ($isInstalled) {
            Write-Log "TeamViewer już zainstalowany, odinstalowuję..." -Color "Blue"
            if ($uninstallString) {
                $uninstallCommand = ($uninstallString -replace "/I", "/X") + " /qn"
                Get-Process -Name "*TeamViewer*" -ErrorAction SilentlyContinue | Stop-Process -Force
                Start-Sleep -Seconds 5
                Write-Log "Uruchamiam odinstalowanie: $uninstallCommand" -Color "Blue"
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallCommand" -Wait -ErrorAction Stop
                Write-Log "TeamViewer odinstalowany." -Color "Green"
            }
            else {
                Write-Log "Nie znaleziono polecenia odinstalowania TeamViewer." -Color "Red" -IsError
                return
            }
        }
        else {
            Write-Log "TeamViewer nie jest zainstalowany, przechodzę do instalacji." -Color "Blue"
            if ( [System.Windows.Forms.MessageBox]::Show("Nie wykryto instalacji TeamViewer. Czy chcesz kontynuować instalację? (Jeśli istnieje proces TeamViewera zostanie on ubity)", "Potwierdzenie", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes ) {
                Write-Log "Instalacja anulowana przez użytkownika." -Color "Red"
                return
            }
            Get-Process -Name "*TeamViewer*" -ErrorAction SilentlyContinue | Stop-Process -Force
        }

        if ($sourcePath -like "http*") {
            $DownloadPathOrUrl = "$sourcePath$fileName"
            Write-Log "Pobieranie TeamViewer z $DownloadPathOrUrl..." -Color "Blue"
            Invoke-WebRequest -Uri $DownloadPathOrUrl -OutFile $localPath -UseBasicParsing -ErrorAction Stop
            Write-Log "Pobrano TeamViewer do: $localPath" -Color "Green"
        }
        else {
            $localPath = Join-Path $sourcePath $fileName
            Write-Log "Instalacja TeamViewer z lokalnej ścieżki: $localPath" -Color "Blue"
        }

        Write-Log "Instalacja TeamViewer..." -Color "Blue"
        Write-Log "Używam argumentów MSI: $msiArgs" -Color "Blue"
        Start-Process "msiexec.exe" -ArgumentList "/i `"$localPath`" $msiArgs" -Wait -ErrorAction Stop
        Write-Log "TeamViewer zainstalowany." -Color "Green"
    }
    catch {
        Write-Log "Błąd podczas instalacji TeamViewer: $_" -Color "Red" -IsError
    }
    finally {
        if (Test-Path $localPath) {
            Remove-Item -Path $localPath -Force -ErrorAction SilentlyContinue
            Write-Log "Usunięto plik instalacyjny TeamViewer: $localPath" -Color "Green"
        }
    }
}

function Install-AV {
    try {
        if (-not (Test-Path $configPath)) {
            Write-Log "Brak pliku config.json" -Color "Red" -IsError
            return
        }

        $config = Get-Content $configPath -Raw | ConvertFrom-Json

        $avConfig = $config.AntyVirus
        $source = if ($avConfig.DefaultInstallSource) { 
            $avConfig.DefaultInstallSource 
        } else { 
            $config.DefaultInstallSource 
        }

        $sourcePath = $avConfig.InstallSourcePaths.$source
        if (-not $sourcePath) {
            $sourcePath = $config.InstallSourcePaths.$source
        }

        $fileName = $avConfig.FileName
        if (-not $fileName) {
            Write-Log "Brak nazwy pliku instalacyjnego antywirusa w konfiguracji" -Color "Red" -IsError
            return
        }

        if ($source -eq "web") {
            $avPath = Join-Path $env:TEMP $fileName
            Write-Log "Pobieranie antywirusa z $sourcePath$fileName..."

            $cred = $null
            if ($avConfig.Credentials.Username -and $avConfig.Credentials.Password) {
                $sec = ConvertTo-SecureString $avConfig.Credentials.Password -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential ($avConfig.Credentials.Username, $sec)
            }

            if ($cred) {
                Invoke-WebRequest -Uri ($sourcePath + $fileName) -OutFile $avPath -UseBasicParsing -Credential $cred
            }
            else {
                Invoke-WebRequest -Uri ($sourcePath + $fileName) -OutFile $avPath -UseBasicParsing
            }
        }
        else {
            $avPath = Join-Path $sourcePath $fileName
        }

        if (-not (Test-Path $avPath)) {
            Write-Log "Plik instalacyjny nie został znaleziony: $avPath" -Color "Red" -IsError
            return
        }

        Write-Log "Instalacja antywirusa z $avPath..."
        Start-Process -FilePath $avPath -Wait
        Write-Log "Instalacja antywirusa zakończona"
    }
    catch {
        Write-Log "Błąd podczas instalacji antywirusa: $_" -Color "Red" -IsError
    }
}

function Import-WiFiProfile {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red" -IsError
        return
    }
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    $wifiProfile = $config.WiFiProfile.FileName
    if (Test-Path $wifiProfile) {
        Write-Log "Import profilu Wi-Fi..."
        Start-Process netsh -ArgumentList "wlan add profile filename=`"$wifiProfile`"" -Wait
        Write-Log "Profil Wi-Fi zaimportowany." -Color "Green"
    }
    else {
        Write-Log "Brak pliku profilu Wi-Fi." -Color "Red"
    }
}

function New-LocalAdmin {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red" -IsError
        return
    }
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        $username = $config.LocalAdmin.Username
        $PasswordSecure = Read-Host "Wprowadź hasło dla konta $username" -AsSecureString
        if (-not (Get-LocalUser -Name $username -ErrorAction SilentlyContinue)) {
            if ($null -ne $PasswordSecure -and $PasswordSecure.Length -gt 8) {
                New-LocalUser -Name $username -Password $PasswordSecure -PasswordNeverExpires -AccountNeverExpires
                Add-LocalGroupMember -SID S-1-5-32-544 -Member $username
                Write-Log "Utworzono lokalne konto '$username' w grupie 'Administratorzy'." -Color "Green"
            }
            else {
                Write-Log "Nie podano hasła dla konta $username lub hasło jest zbyt krótkie." -Color "Red" -IsError
            }
        }
        else {
            Write-Log "Użytkownik '$username' już istnieje  pomijam." -Color "Red"
        }
    }
    catch {
        Write-Log "Błąd tworzenia konta: $_" -Color "Red"
    }
}

function Join-Domain {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red" -IsError
        return
    }
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        if (-not $config.DomainJoin -or -not $config.DomainJoin.DomainName) {
            Write-Log "Brak sekcji DomainJoin w config.json lub DomainName" -Color "Red" -IsError
            return
        }

        Add-Type -AssemblyName System.Windows.Forms

        $DomainName = ($config.DomainJoin.DomainName).Trim()
        $UserForJoin = $config.DomainJoin.Username
        $ComputerName = $env:COMPUTERNAME

        Write-Log "Dołączanie do domeny $DomainName jako $UserForJoin..." -Color "Blue"
        $Credential = Get-Credential -UserName $UserForJoin -Message "Podaj dane domenowe dla $UserForJoin"

        Add-Computer -DomainName $DomainName -Credential $Credential -Force -ErrorAction Stop
        Write-Log "Dołączono do domeny $DomainName z nazwą '$ComputerName'" -Color "Green"
        Write-Log "Wymagane jest ponowne uruchomienie komputera, aby zmiany zaczęły obowiązywać." -Color "Yellow"
    }
    catch {
        Write-Log "Błąd dołączania do domeny: $_" -Color "Red" -IsError
    }
}

function Set-NewComputerName {
    Read-Host "Podaj nową nazwę komputera (domyślnie 'PC-SERIAL_NUMBER'): " -OutVariable NewName
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        $NewName = "PC-$((Get-WmiObject -Class Win32_BIOS).SerialNumber)"
    }
    Rename-Computer -NewName $NewName -Force
    Write-Log "Zmieniono nazwę komputera na '$NewName'." -Color "Green"
}

function Set-RegistryDword {
    param(
        [string]$Path,
        [string]$Name,
        [int]$Value
    )
    try {
        # Sprawdź, czy wartość już istnieje
        $current = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        if ($null -eq $current) {
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
            Write-Log "Utworzono wartość $Name w $Path"
        }
        else {
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force
            Write-Log "Zmieniono wartość $Name w $Path"
        }
    }
    catch {
        Write-Log "Błąd rejestru dla $Name w $($Path): $_" -Color "Red"
    }
}

function Set-SystemTweaks {
    try {
        if (-not (Test-Path $configPath)) {
            Write-Log "Brak pliku config.json" -Color "Red"
            return
        }

        $settings = (Get-Content $configPath -Raw | ConvertFrom-Json).SystemSettings
        Write-Log "Zastosowanie ustawień systemowych..."

        if ($settings.DisableDeliveryOptimization) {
            Write-Log "Wyłączanie Delivery Optimization..."
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" -Force | Out-Null
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" -Name "DODownloadMode" -Value 0
        }

        if ($settings.EnableWin10StartMenu) {
            Write-Log "Włączanie klasycznego menu Start (Win10)..."
            Set-RegistryDword -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_ShowClassicMode" -Value 1
        }

        if ($settings.DisableTelemetry) {
            Write-Log "Wyłączanie telemetryki..."
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Force | Out-Null
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0
        }

        if ($settings.DisableCortana) {
            Write-Log "Wyłączanie Cortany..."
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Force | Out-Null
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0
        }

        if ($settings.DisableFastStartup) {
            Write-Log "Wyłączanie szybkiego uruchamiania..."
            Set-RegistryDword -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HiberbootEnabled" -Value 0
        }

        if ($settings.DisableNewsAndInterests) {
            Write-Log "Wyłączanie News and Interests..."
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Force | Out-Null
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Name "AllowNewsAndInterests" -Value 0
        }

        Write-Log "Zmiany systemowe zastosowane." -Color "Green"
    }
    catch {
        Write-Log "Błąd w Set-SystemTweaks: $_" -Color "Red"
    }
}

function Start-WindowsUpdate {
    try {
        Write-Log "Uruchamianie Windows Update..."
        Start-Process "control.exe" -ArgumentList "/name Microsoft.WindowsUpdate"
    }
    catch {
        Write-Log "Nie udało się uruchomić WU: $_" -Color "Red"
    }
}

function Uninstall-Microsoft365Apps {
    $OfficeUninstallStrings = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Microsoft 365*" } | Select-Object UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }    
}

function Show-AppSelectionWindow {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red"
        return
    }

    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    $programs = $config.Programs
    $popup = New-Object Windows.Forms.Form
    $popup.Text = "Wybór aplikacji"
    $popup.AutoScroll = $true
    $popup.Size = "400,600"
    $popup.StartPosition = "CenterScreen"
    $popup.FormBorderStyle = "FixedDialog"
    $popup.MaximizeBox = $false
    $popup.MinimizeBox = $false
    $popup.TopMost = $true

    $checkboxes = @{}
    $y = 20

    foreach ($name in $programs.PSObject.Properties.Name) {
        $cb = New-Object Windows.Forms.CheckBox
        $cb.Text = $name
        $cb.Checked = $programs.$name.Enabled
        $cb.Location = New-Object Drawing.Point(20, $y)
        $cb.Size = New-Object Drawing.Size(340, 20)
        $popup.Controls.Add($cb)
        $checkboxes[$name] = $cb
        $y += 25
    }

    $btnOK = New-Object Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object Drawing.Point(20, ($y + 10))
    $btnOK.Size = New-Object Drawing.Size(340, 30)
    $btnOK.Add_Click({
            $SelectedApps.Clear()
            foreach ($key in $checkboxes.Keys) {
                if ($checkboxes[$key].Checked) {
                    $SelectedApps[$key] = $true
                }
            }
            $popup.Close()
            if ($null -ne $btnChooseApps) {
                $btnChooseApps.Text = "Wybierz aplikacje ($($SelectedApps.Count))"
                $checkboxControls["InstallApplications"].Checked = $SelectedApps.Count -gt 0
            }
            
            if ($SelectedApps.Count -eq 0) {
                $CheckboxControls["InstallApplications"].Checked = $false
                Write-Log "Nie wybrano żadnych aplikacji do instalacji."
            }
            else {
                Write-Log "Wybrano aplikacje: $($SelectedApps.Keys -join ', ')" -Color "Green"
            }
        })

    $btnSelectAll = New-Object Windows.Forms.Button
    $btnSelectAll.Text = "Zaznacz wszystko"
    $btnSelectAll.Location = New-Object Drawing.Point(20, ($y + 50))
    $btnSelectAll.Size = New-Object Drawing.Size(160, 30)
    $btnSelectAll.Add_Click({
            foreach ($cb in $checkboxes.Values) {
                $cb.Checked = $true
            }
        })

    $btnDeselectAll = New-Object Windows.Forms.Button
    $btnDeselectAll.Text = "Odznacz wszystko"
    $btnDeselectAll.Location = New-Object Drawing.Point(200, ($y + 50))
    $btnDeselectAll.Size = New-Object Drawing.Size(160, 30)
    $btnDeselectAll.Add_Click({
            foreach ($cb in $checkboxes.Values) {
                $cb.Checked = $false
            }
        })
    
    $btnCancel = New-Object Windows.Forms.Button
    $btnCancel.Text = "Anuluj"
    $btnCancel.Location = New-Object Drawing.Point(20, ($y + 90))
    $btnCancel.Size = New-Object Drawing.Size(340, 30)
    $btnCancel.Add_Click({
            $popup.Close()
        })
    $popup.Controls.Add($btnCancel)

    $bottomBuffer = New-Object Windows.Forms.Label
    $bottomBuffer.Size = New-Object Drawing.Size(1, 40)
    $bottomBuffer.Location = New-Object Drawing.Point(1, ($y + 100))

    $popup.Controls.Add($bottomBuffer)
    $popup.Controls.Add($btnDeselectAll)
    $popup.Controls.Add($btnSelectAll)
    $popup.Controls.Add($btnOK)
    
    $popup.KeyPreview = $true

    $popup.Add_KeyDown({
            if ($_.KeyCode -eq "Enter") {
                $btnOK.PerformClick()
            }
        })

    $popup.Add_KeyDown({
            if ($_.KeyCode -eq "Escape") {
                $popup.Close()
            }
        })

    $popup.Add_Shown({ $popup.Activate() })
    $popup.ShowDialog()
}

function Start-Deployment {
    $btnStart.Enabled = $false
    $btnStart.BackColor = "Yellow"
    $progressBar.Value = 0

    if (-not (Validate-BeforeRun)) {
        Write-Log "Walidacja nie powiodla sie. Przerywam." -Color "Red" -IsError
        $btnStart.BackColor = "Red"
        $btnStart.Enabled = $true
        return
    }

    if ($CheckboxControls["DisableHibernation"].Checked) { Disable-Hibernation }
    $progressBar.Value = 10

    if ($CheckboxControls["InstallTeamViewer"].Checked) { Install-TeamViewer }
    $progressBar.Value = 20

    if ($CheckboxControls["UninstallMicrosoft365"].Checked) { Uninstall-Microsoft365Apps }
    $progressBar.Value = 25

    if ($CheckboxControls["InstallAV"].Checked) { Install-AV }
    $progressBar.Value = 30

    if ($CheckboxControls["ImportWiFiProfile"].Checked) { Import-WiFiProfile }
    $progressBar.Value = 40

    if ($CheckboxControls["CreateLocalAdmin"].Checked) { New-LocalAdmin }
    $progressBar.Value = 55

    if ($CheckboxControls["JoinDomain"].Checked) { Join-Domain }
    $progressBar.Value = 70

    if ($CheckboxControls["InstallApplications"].Checked -and $SelectedApps.Count -gt 0) {
        Install-SelectedApps
    }
    else {
        Write-Log "Instalacja aplikacji pominięta." 
    }
    $progressBar.Value = 85

    if ($CheckboxControls.ContainsKey("JoinIntune") -and $CheckboxControls["JoinIntune"].Checked) { Join-Intune }
    $progressBar.Value = 90

    if ($CheckboxControls["ChangeSystemSettings"].Checked) { Set-SystemTweaks }
    if ($CheckboxControls["RunWindowsUpdate"].Checked) { 
        Write-Log "Uruchamianie Windows Update..."
        Start-WindowsUpdate 
    }
    
    $progressBar.Value = 100
    Write-Log "Konfiguracja zakończona." -Color "Green"
    [System.Windows.Forms.MessageBox]::Show("Gotowe!", "Zakończono", "OK", "Information")
    $btnStart.BackColor = "Green"
    $btnStart.Enabled = $true
}

function Get-AppSelection {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red"
        return
    }

    $config = Get-Content $configPath -Raw | ConvertFrom-Json

    if ($null -eq $config.Programs) {
        Write-Log "Nie znaleziono sekcji 'Programs' w config.json" -Color "Red"
        return
    }

    $SelectedApps.Clear()

    foreach ($app in $config.Programs.PSObject.Properties.Name) {
        $appObj = $config.Programs.$app
        if ($appObj.Enabled -eq $true) {
            $SelectedApps[$app] = $true
        }
    }

    Write-Log "Załadowano domyślnie zaznaczone aplikacje: $($SelectedApps.Keys -join ', ')" -Color "Green"
    if ($null -ne $btnChooseApps) {
        $btnChooseApps.Text = "Wybierz aplikacje ($($SelectedApps.Count))"
    }
}

function Get-Config {
    if (-not (Test-Path $configPath)) {
        throw "Brak pliku config.json"
    }
    return (Get-Content $configPath -Raw | ConvertFrom-Json)
}

function Save-Config($config) {
    $json = $config | ConvertTo-Json -Depth 10
    $json | Set-Content -Path $configPath -Encoding UTF8
    Write-Log "Zapisano zmiany do config.json" -Color "Green"
}

function Show-ConfigEditor {
    try { $config = Get-Config } catch { Write-Log $_ -Color Red; return }

    $dlg = New-Object Windows.Forms.Form
    $dlg.Text = "Ustawienia – źródła, domena, konto lokalne"
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.TopMost = $true
    $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $dlg.ClientSize = New-Object Drawing.Size(720, 380)

    $grid = New-Object Windows.Forms.TableLayoutPanel
    $grid.Dock = 'Fill'
    $grid.Padding = New-Object Windows.Forms.Padding(16, 16, 16, 72) 
    $grid.ColumnCount = 2
    $grid.RowCount = 7
    $grid.AutoSize = $false
    $grid.GrowStyle = 'AddRows'
    $grid.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38)))
    $grid.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 62)))

    function AddRow([string]$labelText, [System.Windows.Forms.Control]$ctrl) {
        $row = $grid.RowCount
        $grid.RowStyles.Add((New-Object Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
        $lbl = New-Object Windows.Forms.Label
        $lbl.Text = $labelText
        $lbl.AutoSize = $true
        $lbl.Margin = New-Object Windows.Forms.Padding(0, 0, 10, 10)
        $ctrl.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 10)
        $ctrl.Dock = 'Fill'
        $grid.Controls.Add($lbl, 0, $row - 1)
        $grid.Controls.Add($ctrl, 1, $row - 1)
        $grid.RowCount++
    }

    $cmbSrc = New-Object Windows.Forms.ComboBox
    $cmbSrc.DropDownStyle = 'DropDownList'
    [void]$cmbSrc.Items.AddRange(@('network', 'web'))
    $src = [string]$config.DefaultInstallSource
    if ($cmbSrc.Items -contains $src) {
        $cmbSrc.SelectedItem = $src
    }
    else {
        $cmbSrc.SelectedItem = 'network'
    }
    AddRow "Domyślne źródło instalacji" $cmbSrc

    $txtNet = New-Object Windows.Forms.TextBox
    $txtNet.Text = [string]$config.InstallSourcePaths.network
    AddRow "Ścieżka network (UNC)" $txtNet

    $txtWeb = New-Object Windows.Forms.TextBox
    $txtWeb.Text = [string]$config.InstallSourcePaths.web
    AddRow "Ścieżka web (URL)" $txtWeb

    $txtCwd = New-Object Windows.Forms.TextBox
    $txtCwd.Text = [string]$config.CustomWebDataLocation.URL
    AddRow "CustomWebData URL" $txtCwd

    $txtDom = New-Object Windows.Forms.TextBox
    $txtDom.Text = [string]$config.DomainJoin.DomainName
    AddRow "Domena (DomainJoin.DomainName)" $txtDom

    $txtDomUser = New-Object Windows.Forms.TextBox
    $txtDomUser.Text = [string]$config.DomainJoin.Username
    AddRow "Użytkownik do join (domain\\user)" $txtDomUser

    $txtLoc = New-Object Windows.Forms.TextBox
    $txtLoc.Text = [string]$config.LocalAdmin.Username
    AddRow "Lokalny admin (LocalAdmin.Username)" $txtLoc

    $btnPanel = New-Object Windows.Forms.FlowLayoutPanel
    $btnPanel.Dock = 'Bottom'
    $btnPanel.FlowDirection = 'RightToLeft'
    $btnPanel.Padding = New-Object Windows.Forms.Padding(16, 12, 16, 16)
    $btnPanel.Height = 64
    $btnPanel.WrapContents = $false

    $btnCancel = New-Object Windows.Forms.Button
    $btnCancel.Text = "Anuluj"
    $btnCancel.Width = 120
    $btnCancel.Height = 34
    $btnCancel.Margin = New-Object Windows.Forms.Padding(8, 0, 0, 0)
    $btnCancel.Add_Click({ $dlg.Close() })

    $btnSave = New-Object Windows.Forms.Button
    $btnSave.Text = "Zapisz"
    $btnSave.Width = 120
    $btnSave.Height = 34
    $btnSave.Margin = New-Object Windows.Forms.Padding(8, 0, 0, 0)

    $btnPanel.Controls.Add($btnCancel)
    $btnPanel.Controls.Add($btnSave)

    $dlg.AcceptButton = $btnSave
    $dlg.CancelButton = $btnCancel

    $btnSave.Add_Click({
            try {
                $src = [string]$cmbSrc.SelectedItem
                if ([string]::IsNullOrWhiteSpace($src) -or ($src -notin @('network', 'web'))) {
                    [System.Windows.Forms.MessageBox]::Show("Wybierz poprawne źródło (network/web)."); return
                }
                $net = $txtNet.Text.Trim()
                $web = $txtWeb.Text.Trim()
                $cwd = $txtCwd.Text.Trim()
                $dom = $txtDom.Text.Trim()
                $domUser = $txtDomUser.Text.Trim()
                $locUser = $txtLoc.Text.Trim()

                if ($web -and $web[-1] -ne '/') { $web += '/' }
                if ($cwd -and $cwd[-1] -ne '/') { $cwd += '/' }
                if ($net -and $net -notmatch '^\\\\') {
                    [System.Windows.Forms.MessageBox]::Show("Ścieżka network musi być w formacie UNC (\\server\share\)."); return
                }

                $config.DefaultInstallSource = $src
                if (-not $config.InstallSourcePaths) { $config | Add-Member -NotePropertyName InstallSourcePaths -NotePropertyValue (@{}) -Force }
                $config.InstallSourcePaths.network = $net
                $config.InstallSourcePaths.web = $web

                if (-not $config.CustomWebDataLocation) { $config | Add-Member -NotePropertyName CustomWebDataLocation -NotePropertyValue (@{}) -Force }
                $config.CustomWebDataLocation.URL = $cwd

                if (-not $config.DomainJoin) { $config | Add-Member -NotePropertyName DomainJoin -NotePropertyValue (@{}) -Force }
                $config.DomainJoin.DomainName = $dom
                $config.DomainJoin.Username = $domUser

                if (-not $config.LocalAdmin) { $config | Add-Member -NotePropertyName LocalAdmin -NotePropertyValue (@{}) -Force }
                $config.LocalAdmin.Username = $locUser

                Save-Config $config
                [System.Windows.Forms.MessageBox]::Show("Zapisano konfigurację.")
                $dlg.Close()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Błąd zapisu: $($_.Exception.Message)")
            }
        })

    $dlg.Controls.Add($grid)
    $dlg.Controls.Add($btnPanel)
    $dlg.ShowDialog() | Out-Null
}

if (-not (Test-ConfigurationFile)) {
    Write-Log "Błąd konfiguracji. Sprawdź plik config.json." -Color "Red"
    exit 1
}

$form = New-Object Windows.Forms.Form
$form.Text = "Smart Tool for Deployment"
$form.Size = '600,630'
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$y = 20

$tooltip = New-Object Windows.Forms.ToolTip

foreach ($key in $checkboxOptions.Keys) {
    $cb = New-Object Windows.Forms.CheckBox
    $cb.Text = $checkboxOptions[$key].Text
    $cb.Name = $key
    $cb.Checked = $checkboxOptions[$key].Enabled
    $cb.Location = New-Object Drawing.Point(20, $y)
    $cb.Size = New-Object Drawing.Size(550, 20)
    $cb.AutoSize = $true
    $cb.Tag = $key

    # Utrwalenie lokalnej zmiennej do działania w bloku skryptu
    $localCb = $cb
    $tooltip.SetToolTip($localCb, $checkboxOptions[$key].Tooltip)

    $form.Controls.Add($localCb)
    $CheckboxControls[$key] = $localCb
    $y += 25
}

$btnChooseApps = New-Object Windows.Forms.Button
$btnChooseApps.Text = "Wybierz aplikacje"
$btnChooseApps.Location = New-Object Drawing.Point(20, $y)
$btnChooseApps.Size = New-Object Drawing.Size(150, 30)
$btnChooseApps.Add_Click({
        Show-AppSelectionWindow
    })
$form.Controls.Add($btnChooseApps)


$btnSelectAll = New-Object Windows.Forms.Button
$btnSelectAll.Text = "Zaznacz wszystko"
$btnSelectAll.Location = New-Object Drawing.Point(180, $y)
$btnSelectAll.Size = New-Object Drawing.Size(120, 30)
$btnSelectAll.Add_Click({
        foreach ($cb in $CheckboxControls.Values) {
            $cb.Checked = $true
        }
        Write-Log "Zaznaczono wszystkie opcje w głównym oknie." -Color "Green"
    })
$form.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object Windows.Forms.Button
$btnDeselectAll.Text = "Odznacz wszystko"
$btnDeselectAll.Location = New-Object Drawing.Point(310, $y)
$btnDeselectAll.Size = New-Object Drawing.Size(120, 30)
$btnDeselectAll.Add_Click({
        foreach ($cb in $CheckboxControls.Values) {
            $cb.Checked = $false
        }
        Write-Log "Odznaczono wszystkie opcje w głównym oknie." -Color "Green"
    })
$form.Controls.Add($btnDeselectAll)

$btnSettings = New-Object Windows.Forms.Button
$btnSettings.Text = "Ustawienia…"
$btnSettings.Location = New-Object Drawing.Point(440, $y)
$btnSettings.Size = New-Object Drawing.Size(130, 30)
$btnSettings.Add_Click({ Show-ConfigEditor })
$form.Controls.Add($btnSettings)
$y += 40

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Location = New-Object Drawing.Point(20, $y)
$progressBar.Size = New-Object Drawing.Size(550, 20)
$form.Controls.Add($progressBar)
$y += 30

$rtbLog = New-Object Windows.Forms.RichTextBox
$rtbLog.Location = New-Object Drawing.Point(20, $y)
$rtbLog.Size = New-Object Drawing.Size(550, 160)
$rtbLog.ReadOnly = $true
$form.Controls.Add($rtbLog)

$btnStart = New-Object Windows.Forms.Button
$btnStart.Text = "START"
$btnStart.BackColor = "Green"
$btnStart.Location = New-Object Drawing.Point(20, ($y + 170))
$btnStart.Size = New-Object Drawing.Size(550, 40)
$btnStart.Add_Click({ 
        Start-Deployment 
    })
$form.Controls.Add($btnStart)
$form.Add_Shown({ $form.Activate() })

$form.KeyPreview = $true
$form.Add_KeyDown({
        if ($_.KeyCode -eq "Enter") {
            if ($btnStart.Enabled) {
                if ([System.Windows.Forms.MessageBox]::Show("Czy na pewno chcesz rozpocząć konfigurację?", "Potwierdzenie", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $btnStart.PerformClick()
                }
            }
        }
    })

$form.Add_KeyDown({
        if ($_.KeyCode -eq "Escape") {
            if ([System.Windows.Forms.MessageBox]::Show("Czy na pewno chcesz zamknąć aplikację?", "Potwierdzenie", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes) {
                $form.Close()
            }
        }
    })

$CheckboxControls["InstallApplications"].Add_CheckedChanged({
        if ($CheckboxControls["InstallApplications"].Checked) {
            if ($btnChooseApps.Enabled -eq $false) {
                $btnChooseApps.Enabled = $true
            }
        }
        else {
            $btnChooseApps.Enabled = $false
        }
    })

Get-AppSelection

$form.Add_Shown({ $form.Activate() })
$form.ShowDialog() | Out-Null

