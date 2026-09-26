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

        It "Tworzy plik na dysku i dopisuje odpowiednio sformatowaną linijkę (data + godzina + poziom + kontekst)" {
            Write-Log -Text "To jest tylko test logowania"

            $zawartosc = Get-Content $script:LogFilePath -Raw
            $zawartosc | Should -Match "To jest tylko test logowania"
            $zawartosc | Should -Match "\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] \[INFO\] \[System\]"
            (Test-Path $script:ErrorLogFilePath) | Should -Be $false
        }

        It "Tworzy plik błędów gdy użyta jest flaga -IsError" {
            Write-Log -Text "To jest krytyczny błąd" -IsError

            $zawartoscLog = Get-Content $script:LogFilePath -Raw
            $zawartoscErr = Get-Content $script:ErrorLogFilePath -Raw
            $zawartoscLog | Should -Match "\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] \[ERROR\] \[System\] To jest krytyczny błąd"
            $zawartoscErr | Should -Match "\[ERROR\] \[System\] To jest krytyczny błąd"
        }

        It "Uzywa jawnie podanego -Context zamiast domyslnego 'System'" {
            Write-Log -Text "Krok wykonany automatycznie" -Context "Automat"

            $zawartosc = Get-Content $script:LogFilePath -Raw
            $zawartosc | Should -Match "\[Automat\] Krok wykonany automatycznie"
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

        It "Nie przerywa się (i nie rzuca dalej) w trybie Dry-Run, gdy Get-AppxPackage zawiedzie (np. brak modulu Appx na niektorych kompilacjach Windows)" {
            Mock Get-AppxPackage { throw "Operation is not supported on this platform." }
            Mock Write-Log {}
            Mock Do-WpfEvents {}

            $script:isCancelled = $false
            $script:DryRun = $true
            try {
                { Remove-Bloatware } | Should -Not -Throw
            } finally {
                $script:DryRun = $false
            }
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
            # UWAGA: CheckboxControls/SelectedApps przekazujemy jako PARAMETRY Test-BeforeRun (nie
            # przez nadpisanie zmiennych skryptowych) - dot-sourcing w BeforeAll i nadpisywanie
            # $script:x / $x z poziomu It okazało się zawodne (funkcja czasem domyka się nad innym
            # egzemplarzem scope'u niż ten, do którego pisze It). Jawne przekazanie parametrów jest
            # jednoznaczne niezależnie od scope'u. Zawartość config.json podajemy przez
            # Mock Get-Content/Test-Path z tego samego powodu.
            Mock Get-CimInstance {}
            Mock Test-Path { return $true }
            Mock Get-Content { return '{ "DefaultInstallSource": "winget", "Programs": { "ZlaAplikacja": { "Enabled": true, "FileName": "" } } }' }
            Mock Show-ThemedMessageBox { return [System.Windows.MessageBoxResult]::OK }
            Mock Write-Log {}

            $wynik = Test-BeforeRun -SelectedApps @{ "ZlaAplikacja" = $true } -CheckboxControls @{ 'InstallApplications' = [PSCustomObject]@{ IsChecked = $true } }
            $wynik | Should -Be $false
            Should -Invoke Write-Log -ParameterFilter { $Text -match "Brak 'FileName' dla 'ZlaAplikacja'" } -Times 1
        }

        It "Dodaje ostrzeżenie, jeśli na dysku C: jest mniej niż 15 GB wolnego miejsca" {
            Mock Test-Path { return $true }
            Mock Get-Content { return '{ "DefaultInstallSource": "winget", "Programs": {} }' }
            Mock Get-CimInstance { return [PSCustomObject]@{ FreeSpace = 10GB } } -ParameterFilter { $ClassName -eq 'Win32_LogicalDisk' }
            Mock Show-ThemedMessageBox { return [System.Windows.MessageBoxResult]::OK }
            Mock Write-Log {}

            Test-BeforeRun -SelectedApps @{} -CheckboxControls @{} | Out-Null

            Should -Invoke Write-Log -ParameterFilter { $Text -match "Mało wolnego miejsca na dysku C:" } -Times 1
        }
    }

    Context "Wstrzymywanie hibernacji (Suspend-Hibernation)" {
        It "Wywołuje funkcję z odpowiednimi flagami (w tym ES_DISPLAY_REQUIRED)" {
            Mock Write-Log {}
            Suspend-Hibernation
            Assert-MockCalled Write-Log -ParameterFilter { $Text -match "Zapobieganie usypianiu włączone" } -Times 1
        }
    }

    Context "Parsowanie deinstalatora Office Click-to-Run (Resolve-OfficeClickToRunUninstallInfo)" {
        It "Wyodrębnia ProductId (productstoremove=) i ścieżkę z cudzysłowu" {
            $cmd = '"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365HomePremRetail.16_pl-pl_x-none culture=pl-pl'
            $r = Resolve-OfficeClickToRunUninstallInfo -Cmd $cmd
            $r.ProductId | Should -Be "O365HomePremRetail.16_pl-pl_x-none"
            $r.ExePath | Should -Be "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
        }

        It "Wyodrębnia ProductId z alternatywnej pisowni ProductID=" {
            $cmd = '"C:\ClickToRun\OfficeClickToRun.exe" ProductID=O365ProPlusRetail'
            $r = Resolve-OfficeClickToRunUninstallInfo -Cmd $cmd
            $r.ProductId | Should -Be "O365ProPlusRetail"
        }

        It "Wyodrębnia samą ścieżkę .exe, gdy cmd nie jest w cudzysłowie" {
            $cmd = 'C:\ClickToRun\OfficeClickToRun.exe scenario=install'
            $r = Resolve-OfficeClickToRunUninstallInfo -Cmd $cmd
            $r.ExePath | Should -Be "C:\ClickToRun\OfficeClickToRun.exe"
        }

        It "Zwraca `$null dla ProductId i ExePath, gdy string nie pasuje do żadnego wzorca" {
            $r = Resolve-OfficeClickToRunUninstallInfo -Cmd "zupelnie niepowiazany tekst bez sciezki"
            $r.ProductId | Should -Be $null
            $r.ExePath | Should -Be $null
        }
    }

    Context "Uruchamianie deinstalatora z obsługą anulowania (Start-UninstallProcessWithCancel)" {
        BeforeEach {
            $script:isCancelledFromUninstall = $false
        }

        It "Zwraca -1 i loguje błąd, gdy Start-Process rzuci wyjątek" {
            Mock Start-Process { throw "Nie znaleziono pliku" }
            Mock Write-Log {}
            $exitCode = Start-UninstallProcessWithCancel -FilePath "brak.exe" -ArgumentList "" -LogContext "TestApp"
            $exitCode | Should -Be -1
            Assert-MockCalled Write-Log -ParameterFilter { $Text -match "Błąd uruchamiania deinstalatora \(TestApp\)" } -Times 1
        }

        It "Zwraca kod wyjścia procesu, gdy zakończy się od razu bez anulowania" {
            $fakeProc = [PSCustomObject]@{ HasExited = $true; ExitCode = 0; Id = 4242 }
            Mock Start-Process { return $fakeProc }
            Mock Do-WpfEvents {}
            $exitCode = Start-UninstallProcessWithCancel -FilePath "cicho.exe" -ArgumentList "/S" -LogContext "TestApp"
            $exitCode | Should -Be 0
        }
    }

    Context "Checkpoint wdrożenia (Import/Save/Clear-DeploymentCheckpoint, Test-StepDone)" {
        # Uwaga: celowo NIE odczytujemy/nadpisujemy $script:CompletedDeploymentSteps bezpośrednio
        # z poziomu It - tylko przez wywołania funkcji (Import-/Save-/Clear-DeploymentCheckpoint,
        # Test-StepDone), tak by test nie zależał od domykania zmiennych skryptowych przez
        # PowerShell/Pester pomiędzy blokiem BeforeAll a It (patrz komentarz w kontekście
        # "Walidacja przed wdrożeniem" powyżej).
        It "Import-DeploymentCheckpoint zwraca Found = `$false, gdy plik checkpointu nie istnieje" {
            Mock Test-Path { return $false }
            $result = Import-DeploymentCheckpoint
            $result.Found | Should -Be $false
        }

        It "Import-DeploymentCheckpoint wczytuje ukończone kroki, a Test-StepDone je widzi" {
            Mock Test-Path { return $true }
            Mock Get-Content { return '{ "Timestamp": "2026-01-01 10:00:00", "CompletedSteps": ["WaitForNetwork", "InstallAV"] }' }
            $result = Import-DeploymentCheckpoint
            $result.Found | Should -Be $true
            $result.StepCount | Should -Be 2
            Test-StepDone "InstallAV" | Should -Be $true
            Test-StepDone "JoinDomain" | Should -Be $false
        }

        It "Save-DeploymentCheckpointStep zapisuje krok, który Test-StepDone potem widzi" {
            Mock Set-Content {}
            Clear-DeploymentCheckpoint
            Save-DeploymentCheckpointStep "RemoveBloatware"
            Test-StepDone "RemoveBloatware" | Should -Be $true
            Test-StepDone "InnyKrok" | Should -Be $false
            Assert-MockCalled Set-Content -Times 1
        }

        It "Clear-DeploymentCheckpoint usuwa plik i czyści zapamiętane kroki" {
            Mock Set-Content {}
            Mock Test-Path { return $true }
            Mock Remove-Item {}
            Save-DeploymentCheckpointStep "JakisKrok"
            Clear-DeploymentCheckpoint
            Test-StepDone "JakisKrok" | Should -Be $false
            Assert-MockCalled Remove-Item -Times 1
        }
    }
}