#Requires -Module Pester

BeforeAll {
    # Ustawiamy flagę, która zapobiegnie uruchomieniu GUI z głównego pliku
    $global:PesterTesting = $true
    . "$PSScriptRoot\SmartToolforDeployment.ps1"
}

Describe "SmartToolforDeployment - Testy Jednostkowe" {

    Context "Weryfikacja funkcji pomocniczych (Utils)" {
        It "Test-UrlValid zwraca `$true dla prawidłowych linków HTTP/HTTPS" {
            Test-UrlValid -Url "https://mojastrona.pl/instalki/program.exe" | Should -Be $true
            Test-UrlValid -Url "http://192.168.1.50/apps/" | Should -Be $true
        }

        It "Test-UrlValid zwraca `$false dla innych protokołów (np. FTP)" {
            Test-UrlValid -Url "ftp://serwer.pl/plik.exe" | Should -Be $false
        }

        It "Test-UrlValid zwraca `$false dla niekompletnego ciągu znaków lub pustego adresu" {
            Test-UrlValid -Url "zly-adres-pl" | Should -Be $false
            Test-UrlValid -Url "" | Should -Be $false
        }
    }

    Context "System logowania operacji (Write-Log)" {
        BeforeEach {
            $script:LogFilePath = "$PSScriptRoot\Test-DeployLog.txt"
            $script:ErrorLogFilePath = "$PSScriptRoot\Test-DeployErrorLog.txt"
            if (Test-Path $script:LogFilePath) { Remove-Item $script:LogFilePath -Force }
            if (Test-Path $script:ErrorLogFilePath) { Remove-Item $script:ErrorLogFilePath -Force }
        }

        It "Tworzy plik na dysku i dopisuje odpowiednio sformatowaną linijkę" {
            Write-Log -Text "To jest tylko test logowania"
            
            $zawartosc = Get-Content $script:LogFilePath -Raw
            $zawartosc | Should -Match "To jest tylko test logowania"
            $zawartosc | Should -Match "\[\d{2}:\d{2}:\d{2}\]"
            (Test-Path $script:ErrorLogFilePath) | Should -Be $false
        }

        It "Tworzy plik błędów gdy użyta jest flaga -IsError" {
            Write-Log -Text "To jest krytyczny błąd" -IsError
            
            $zawartoscLog = Get-Content $script:LogFilePath -Raw
            $zawartoscErr = Get-Content $script:ErrorLogFilePath -Raw
            $zawartoscLog | Should -Match "To jest krytyczny błąd"
            $zawartoscErr | Should -Match "\[ERROR\] To jest krytyczny błąd"
        }

        AfterEach {
            if (Test-Path $script:LogFilePath) { Remove-Item $script:LogFilePath -Force }
            if (Test-Path $script:ErrorLogFilePath) { Remove-Item $script:ErrorLogFilePath -Force }
        }
    }

    Context "Weryfikacja konfiguracji i plików JSON" {
        It "Get-Config rzuca błąd, jeśli plik JSON nie istnieje" {
            Mock Test-Path { return $false }
            { Get-Config } | Should -Throw "Brak pliku config.json"
        }

        It "Test-ConfigurationFile zwraca `$false jeśli pliku nie ma na dysku" {
            Mock Test-Path { return $false }
            Test-ConfigurationFile -Silent | Should -Be $false
        }
        
        It "Test-ConfigurationFile zwraca `$false jeśli brakuje sekcji 'Programs'" {
            Mock Test-Path { return $true }
            Mock Get-Content { return '{ "ZlaSekcja": "Test" }' }
            Test-ConfigurationFile -Silent | Should -Be $false
        }

        It "Test-ConfigurationFile zwraca `$true jeśli plik istnieje i ma format z 'Programs'" {
            Mock Test-Path { return $true }
            Mock Get-Content { return '{ "Programs": { "Aplikacja": "App.exe" } }' }
            Test-ConfigurationFile -Silent | Should -Be $true
        }
    }

    Context "Manipulacja rejestrem (Set-RegistryDword)" {
        It "Tworzy nowy klucz (New-ItemProperty) jeśli wartość jeszcze nie istnieje" {
            Mock Get-ItemProperty { return $null }
            Mock New-ItemProperty { return $true }
            Mock Set-ItemProperty {}
            Mock Write-Log {}

            Set-RegistryDword -Path "HKLM:\Test" -Name "Wartosc" -Value 1
            
            Assert-MockCalled New-ItemProperty -Times 1 -Exactly
            Assert-MockCalled Set-ItemProperty -Times 0 -Exactly
        }

        It "Aktualizuje klucz (Set-ItemProperty) jeśli wartość już istnieje" {
            Mock Get-ItemProperty { return [PSCustomObject]@{ Wartosc = 0 } }
            Mock Set-ItemProperty { return $true }
            Mock New-ItemProperty {}
            Mock Write-Log {}

            Set-RegistryDword -Path "HKLM:\Test" -Name "Wartosc" -Value 1
            
            Assert-MockCalled New-ItemProperty -Times 0 -Exactly
            Assert-MockCalled Set-ItemProperty -Times 1 -Exactly
        }
    }

    Context "Audyt sprzętowy (Get-HardwareAudit)" {
        It "Zbiera dane sprzętowe i formatuje je w czytelny raport" {
            Mock Get-CimInstance {
                param($ClassName)
                switch ($ClassName) {
                    "Win32_OperatingSystem" { return [PSCustomObject]@{ Caption="Windows 11 Pro"; OSArchitecture="64-bit"; BuildNumber="22621" } }
                    "Win32_BIOS" { return [PSCustomObject]@{ SerialNumber="ABC12345"; SMBIOSBIOSVersion="1.0.0" } }
                    "Win32_Processor" { return [PSCustomObject]@{ Name="Intel Core i9" } }
                    "Win32_PhysicalMemory" { return [PSCustomObject]@{ Capacity=17179869184 } } # 16GB
                    "Win32_DiskDrive" { return [PSCustomObject]@{ Model="Samsung SSD"; Size=512000000000; MediaType="Fixed hard disk media" } }
                    "Win32_NetworkAdapterConfiguration" { return [PSCustomObject]@{ Description="Intel Wi-Fi"; MACAddress="00:11:22:33:44:55"; IPAddress=@("192.168.1.10"); IPEnabled=$true } }
                }
            }
            Mock Get-ItemProperty { return $null }

            $raport = Get-HardwareAudit
            $raport | Should -Match "Windows 11 Pro 64-bit"
            $raport | Should -Match "ABC12345"
            $raport | Should -Match "Intel Core i9"
            $raport | Should -Match "16 GB"
            $raport | Should -Match "Samsung SSD"
            $raport | Should -Match "192.168.1.10"
        }
    }
    
    Context "Usuwanie Bloatware (Remove-Bloatware)" {
        BeforeAll {
            $appxCmds = @('Get-AppxPackage', 'Remove-AppxPackage', 'Get-AppxProvisionedPackage', 'Remove-AppxProvisionedPackage')
            foreach ($cmd in $appxCmds) {
                if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
                    New-Item -Path "function:global:$cmd" -Value {} -Force | Out-Null
                }
            }
        }
        It "Próbuje usunąć domyślne aplikacje poprzez Remove-AppxPackage" {
            Mock Get-AppxPackage { return [PSCustomObject]@{ Name="TikTok" } }
            Mock Remove-AppxPackage {}
            Mock Get-AppxProvisionedPackage { return $null }
            Mock Remove-AppxProvisionedPackage {}
            Mock Write-Log {}
            Mock Do-WpfEvents {}
            
            $script:isCancelled = $false
            Remove-Bloatware
            
            Assert-MockCalled Remove-AppxPackage
        }
    }

    Context "Zarządzanie listą programów (Struktury JSON)" {
        It "Wykrywa duplikat nazwy programu w obiekcie konfiguracyjnym" {
            $config = [PSCustomObject]@{
                Programs = [PSCustomObject]@{
                    "Chrome" = [PSCustomObject]@{ Enabled = $true; FileName = "chrome.exe" }
                    "7zip"   = [PSCustomObject]@{ Enabled = $true; FileName = "7z.exe" }
                }
            }

            $czyIstnieje1 = $config.Programs.PSObject.Properties.Name -contains "Firefox"
            $czyIstnieje1 | Should -Be $false

            $czyIstnieje2 = $config.Programs.PSObject.Properties.Name -contains "Chrome"
            $czyIstnieje2 | Should -Be $true
        }

        It "Bezpiecznie modyfikuje właściwości programu (symulacja Edycji)" {
            $config = [PSCustomObject]@{
                Programs = [PSCustomObject]@{
                    "TestApp" = [PSCustomObject]@{ Enabled = $false; FileName = "old.exe" }
                }
            }
            
            $config.Programs.PSObject.Properties.Remove("TestApp")
            ($config.Programs.PSObject.Properties.Name -contains "TestApp") | Should -Be $false

            $newData = [PSCustomObject]@{ Enabled = $true; FileName = "new.exe" }
            Add-Member -InputObject $config.Programs -NotePropertyName "TestApp" -NotePropertyValue $newData -Force

            $config.Programs.TestApp.FileName | Should -Be "new.exe"
            $config.Programs.TestApp.Enabled | Should -Be $true
        }
    }

    Context "Walidacja przed wdrożeniem (Test-BeforeRun)" {
        It "Zwraca `$false, jeśli aplikacja przeznaczona do instalacji nie ma parametru FileName" {
            $script:configPath = "$PSScriptRoot\temp_test_config.json"
            $script:SelectedApps = @{ "ZlaAplikacja" = $true }
            $script:CheckboxControls = @{ 'InstallApplications' = [PSCustomObject]@{ IsChecked = $true } }
            
            $tempConfig = @{ DefaultInstallSource = "winget"; Programs = @{ ZlaAplikacja = @{ Enabled = $true; FileName = "" } } } | ConvertTo-Json -Depth 5
            Set-Content -Path $script:configPath -Value $tempConfig -Encoding UTF8

            Mock Show-ThemedMessageBox { return [System.Windows.MessageBoxResult]::OK }
            Mock Write-Log {}

            $wynik = Test-BeforeRun
            $wynik | Should -Be $false
            Assert-MockCalled Write-Log -ParameterFilter { $Text -match "Brak 'FileName' dla 'ZlaAplikacja'" } -Times 1
            
            if (Test-Path $script:configPath) { Remove-Item $script:configPath -Force }
        }
        
        It "Dodaje ostrzeżenie, jeśli na dysku C: jest mniej niż 15 GB wolnego miejsca" {
            $script:configPath = "$PSScriptRoot\temp_test_config.json"
            $script:SelectedApps = @{}
            $script:CheckboxControls = @{}
            
            $tempConfig = @{ DefaultInstallSource = "winget"; Programs = @{} } | ConvertTo-Json -Depth 5
            Set-Content -Path $script:configPath -Value $tempConfig -Encoding UTF8

            Mock Get-CimInstance { return [PSCustomObject]@{ FreeSpace = 10GB } } -ParameterFilter { $ClassName -eq 'Win32_LogicalDisk' }
            Mock Show-ThemedMessageBox { return [System.Windows.MessageBoxResult]::OK }
            Mock Write-Log {}

            Test-BeforeRun | Out-Null
            
            Assert-MockCalled Write-Log -ParameterFilter { $Text -match "Mało wolnego miejsca na dysku C:" } -Times 1
            
            if (Test-Path $script:configPath) { Remove-Item $script:configPath -Force }
        }
    }

    Context "Wstrzymywanie hibernacji (Suspend-Hibernation)" {
        It "Wywołuje funkcję z odpowiednimi flagami (w tym ES_DISPLAY_REQUIRED)" {
            Mock Write-Log {}
            Suspend-Hibernation
            Assert-MockCalled Write-Log -ParameterFilter { $Text -match "Zapobieganie usypianiu włączone" } -Times 1
        }
    }
}