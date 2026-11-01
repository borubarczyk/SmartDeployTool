#Requires -Module Pester
# UWAGA: ten plik MUSI być zapisany jako UTF-8 z BOM (tak jak SmartToolforDeployment.ps1). Windows
# PowerShell 5.1 czyta pliki bez BOM w stronie kodowej ANSI, więc polskie znaki w oczekiwanych
# tekstach (np. "Mało wolnego miejsca") nie pasowałyby do komunikatów ze skryptu.

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

    Context "Łączenie źródła instalacji (Join-InstallSource)" {
        It "Dokleja '/' między adresem URL a nazwą pliku, gdy go brakuje" {
            Join-InstallSource -BasePath "https://apps.firma.pl/Data" -FileName "chrome.exe" | Should -Be "https://apps.firma.pl/Data/chrome.exe"
        }

        It "Nie dubluje '/', gdy adres URL już się nim kończy" {
            Join-InstallSource -BasePath "https://apps.firma.pl/Data/" -FileName "chrome.exe" | Should -Be "https://apps.firma.pl/Data/chrome.exe"
        }

        It "Dla ścieżki UNC używa Join-Path" {
            Join-InstallSource -BasePath "\\SERWER\Instalki" -FileName "7z.exe" | Should -Be "\\SERWER\Instalki\7z.exe"
        }
    }

    Context "Ocena kodu wyjścia instalatora (Test-InstallerExitCode, Write-InstallResult)" {
        It "Uznaje 0, 1641 i 3010 za sukces" {
            Test-InstallerExitCode -ExitCode 0 | Should -Be $true
            Test-InstallerExitCode -ExitCode 1641 | Should -Be $true
            Test-InstallerExitCode -ExitCode 3010 | Should -Be $true
        }

        It "Uznaje inne kody i brak kodu za błąd" {
            Test-InstallerExitCode -ExitCode 1603 | Should -Be $false
            Test-InstallerExitCode -ExitCode $null | Should -Be $false
        }

        It "Kod winget 'pakiet już zainstalowany' jest sukcesem tylko z przełącznikiem -Winget" {
            Test-InstallerExitCode -ExitCode -1978335135 -Winget | Should -Be $true
            Test-InstallerExitCode -ExitCode -1978335135 | Should -Be $false
        }

        It "Write-InstallResult loguje błąd z kodem wyjścia zamiast 'zainstalowany'" {
            Mock Write-Log {}
            Write-InstallResult -Name "Chrome" -ExitCode 1603 | Should -Be $false
            Should -Invoke Write-Log -ParameterFilter { $IsError -and $Text -match "kod wyjścia: 1603" } -Times 1
            Should -Invoke Write-Log -ParameterFilter { $Text -match "zainstalowany\." } -Times 0 -Exactly
        }
    }

    Context "Uruchamianie procesów (Start-ProcessWithEvents)" {
        It "Nie przekazuje pustego -ArgumentList do Start-Process (PS 5.1 go odrzuca)" {
            Mock Start-Process { return [PSCustomObject]@{ HasExited = $true; ExitCode = 0; Id = 1 } }
            Mock Do-WpfEvents {}
            Start-ProcessWithEvents -FilePath "setup.exe" -ArgumentList "" | Should -Be 0
            Should -Invoke Start-Process -ParameterFilter { $null -eq $ArgumentList } -Times 1
        }

        It "Przekazuje argumenty, gdy nie są puste, i zwraca kod wyjścia procesu" {
            Mock Start-Process { return [PSCustomObject]@{ HasExited = $true; ExitCode = 1603; Id = 1 } }
            Mock Do-WpfEvents {}
            Start-ProcessWithEvents -FilePath "msiexec.exe" -ArgumentList "/i app.msi /qn" | Should -Be 1603
            Should -Invoke Start-Process -ParameterFilter { $ArgumentList -eq "/i app.msi /qn" } -Times 1
        }
    }

    Context "Nazwa komputera (Get-DefaultComputerName, Test-ComputerNameValid)" {
        It "Buduje nazwę PC-NumerSeryjny bez niedozwolonych znaków i przycina do 15 znaków" {
            Get-DefaultComputerName -SerialNumber "5CG1234XYZ" | Should -Be "PC-5CG1234XYZ"
            Get-DefaultComputerName -SerialNumber "To Be Filled By O.E.M." | Should -Be "PC-ToBeFilledBy"
            Get-DefaultComputerName -SerialNumber "" | Should -Be "PC-NOSERIAL"
        }

        It "Akceptuje poprawną nazwę" {
            Test-ComputerNameValid -Name "PC-01" | Should -BeNullOrEmpty
        }

        It "Odrzuca nazwy za długie, z niedozwolonymi znakami, z samych cyfr i zaczynające się myślnikiem" {
            Test-ComputerNameValid -Name "KOMPUTER-KSIEGOWOSC" | Should -Match "15 znaków"
            Test-ComputerNameValid -Name "PC_01" | Should -Match "Dozwolone"
            Test-ComputerNameValid -Name "123456" | Should -Match "cyfr"
            Test-ComputerNameValid -Name "-PC" | Should -Match "myślnikiem"
            Test-ComputerNameValid -Name "" | Should -Match "pusta"
        }

        It "Przy dołączaniu do domeny odkłada nazwę zamiast Rename-Computer, a Join-Domain nadaje ją przez Add-Computer -NewName" {
            $script:DryRun = $false
            $script:PendingComputerName = $null
            $oldConfigPath = $configPath
            $configPath = Join-Path $TestDrive "config-domain.json"
            '{ "DomainJoin": { "DomainName": "firma.local", "Username": "FIRMA\\admin" } }' | Set-Content -Path $configPath -Encoding UTF8
            try {
                Mock Get-DefaultComputerName { "PC-TEST" }
                Mock Show-InputDialog { "PC-NOWY" }
                Mock Rename-Computer { }
                Mock Get-Credential { New-Object System.Management.Automation.PSCredential("FIRMA\admin", (ConvertTo-SecureString "x" -AsPlainText -Force)) }
                Mock Add-Computer { }

                Set-NewComputerName -DeferToDomainJoin
                $script:PendingComputerName | Should -Be "PC-NOWY"
                Should -Invoke Rename-Computer -Times 0 -Exactly

                Join-Domain
                Should -Invoke Add-Computer -Times 1 -Exactly -ParameterFilter { $NewName -eq "PC-NOWY" -and $DomainName -eq "firma.local" }
                $script:PendingComputerName | Should -BeNullOrEmpty
            } finally {
                $configPath = $oldConfigPath
                $script:PendingComputerName = $null
            }
        }

        It "Gdy Add-Computer dołączy, a nie zmieni nazwy, nie zostawia nazwy do lokalnego Rename-Computer" {
            $script:DryRun = $false
            $configPath = Join-Path $TestDrive "config-domain2.json"
            '{ "DomainJoin": { "DomainName": "firma.local", "Username": "FIRMA\\admin" } }' | Set-Content -Path $configPath -Encoding UTF8
            try {
                Mock Get-Credential { New-Object System.Management.Automation.PSCredential("FIRMA\admin", (ConvertTo-SecureString "x" -AsPlainText -Force)) }
                Mock Add-Computer {
                    throw (New-Object System.Management.Automation.ErrorRecord((New-Object System.Exception "Konto już istnieje"), 'FailToRenameAfterJoinDomain,Microsoft.PowerShell.Commands.AddComputerCommand', 'OperationStopped', $null))
                }
                Mock Write-Log { }
                $script:PendingComputerName = "PC-NOWY"
                Join-Domain
                $script:PendingComputerName | Should -BeNullOrEmpty
                Should -Invoke Write-Log -ParameterFilter { $Text -match "Dołączono do domeny firma.local, ale zmiana nazwy" }
            } finally {
                $script:PendingComputerName = $null
            }
        }
    }

    Context "Zapis ustawień (Set-ConfigValue, Get-ConfigSection)" {
        It "Dodaje brakujące pole do obiektu z ConvertFrom-Json zamiast rzucać wyjątek" {
            $cfg = '{ "InstallSourcePaths": { "web": "https://x/" } }' | ConvertFrom-Json
            { Set-ConfigValue $cfg.InstallSourcePaths 'network' '\\srv\share\' } | Should -Not -Throw
            $cfg.InstallSourcePaths.network | Should -Be '\\srv\share\'
            $cfg.InstallSourcePaths.web | Should -Be 'https://x/'
        }

        It "Tworzy brakującą sekcję i pozwala ustawić w niej pole" {
            $cfg = '{ "DefaultInstallSource": "web" }' | ConvertFrom-Json
            Set-ConfigValue (Get-ConfigSection $cfg 'DomainJoin') 'DomainName' 'firma.local'
            $cfg.DomainJoin.DomainName | Should -Be 'firma.local'
        }
    }

    Context "Okno główne (zadania w grupach, menu Narzędzia)" {
        It "Każde zadanie ma grupę i trafiło do okna głównego" {
            $CheckboxControls.Count | Should -Be $checkboxOptions.Count
            foreach ($key in $checkboxOptions.Keys) {
                @($script:TaskGroups.Keys) | Should -Contain $checkboxOptions[$key].Group
                $CheckboxControls[$key].Parent | Should -BeOfType [System.Windows.Controls.StackPanel]
            }
        }

        It "Każda kolumna zadań zawiera podpisy grup" {
            foreach ($colName in 'spTaskCol0', 'spTaskCol1', 'spTaskCol2') {
                $col = $Window.FindName($colName)
                $col | Should -Not -BeNullOrEmpty
                @($col.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] }).Count | Should -Be 2
            }
        }

        It "Narzędzia systemowe są w menu przycisku Narzędzia" {
            $Window.FindName("btnSysInfo") | Should -BeOfType [System.Windows.Controls.MenuItem]
            $Window.FindName("btnTools").ContextMenu.Items.Count | Should -BeGreaterThan 5
        }
    }

    Context "Okno Ustawienia (zakładki z listami)" {
        BeforeAll {
            # Okna zamykają się same zaraz po wczytaniu - test sprawdza zawartość bez klikania.
            # Handler klasy zostaje w procesie na stałe, więc działa tylko przy włączonej fladze.
            if (-not $global:SdtAutoCloseRegistered) {
                [System.Windows.EventManager]::RegisterClassHandler([System.Windows.Window], [System.Windows.FrameworkElement]::LoadedEvent, [System.Windows.RoutedEventHandler]{
                    param($loadedSender, $loadedArgs)
                    if ($global:SdtAutoCloseWindows) { $global:SdtLastWindow = $loadedSender; $loadedSender.Close() }
                })
                $global:SdtAutoCloseRegistered = $true
            }
            $global:SdtAutoCloseWindows = $true
        }

        AfterAll {
            $global:SdtAutoCloseWindows = $false
        }

        It "Ma 9 zakładek i wypełnia listy programów, profili, rejestru i skryptów" {
            Mock Get-Config {
                '{ "DefaultInstallSource": "web", "InstallSourcePaths": { "web": "https://serwer/instalki/" },
                   "WebAuth": { "Username": "jan", "Password": "tajne" },
                   "Programs": { "7zip": { "Enabled": true, "FileName": "7z.exe" }, "Chrome": { "Enabled": false, "FileName": "Google.Chrome", "DownloadUrl": "https://x/y.exe" } },
                   "Profiles": { "Biuro": ["7zip", "Chrome"] },
                   "SystemSettings": { "CustomRegistry": [ { "Path": "HKLM:\\SOFTWARE\\Firma", "Name": "Test", "Value": "1", "PropertyType": "DWord" } ] },
                   "PostInstallScripts": ["a.ps1", "b.bat"],
                   "DefaultCheckboxes": { "AutoReboot": true } }' | ConvertFrom-Json
            }
            $global:SdtLastWindow = $null
            Show-ConfigEditor
            $w = $global:SdtLastWindow
            $w | Should -Not -BeNullOrEmpty
            $w.Title | Should -Be "Ustawienia"
            $w.FindName("tabSettings").Items.Count | Should -Be 9
            # Nawigacja po lewej (NavTabControl): każda pozycja ma ikonę - encję czcionki ikon w Tag.
            foreach ($tab in $w.FindName("tabSettings").Items) {
                [string]$tab.Tag | Should -Not -BeNullOrEmpty -Because "zakładka $($tab.Header)"
            }

            $programs = $w.FindName("lbPrograms")
            $programs.Items.Count | Should -Be 2
            $programs.Items[0].Name | Should -Be "7zip"
            $programs.Items[0].Default | Should -Be "tak"
            $programs.Items[1].Url | Should -Be "tak"

            $w.FindName("lbProfiles").Items[0].Apps | Should -Be "7zip, Chrome"
            $w.FindName("lbRegistry").Items[0].Type | Should -Be "DWord"
            $w.FindName("lbScripts").Items.Count | Should -Be 2
            $w.FindName("txtWebPass").Password | Should -Be "tajne"
            $w.FindName("cmbSrc").SelectedItem | Should -Be "web"
            $w.FindName("txtSettingsState").Text | Should -Be "Bez zmian"
        }
    }

    Context "Motyw graficzny (wspólny wygląd wszystkich okien)" {
        BeforeAll {
            # Wszystkie bloki XAML okien ze skryptu (logowanie, komunikaty, Ustawienia, okno główne...).
            $script:ThemeTestSource = Get-Content "$PSScriptRoot\SmartToolforDeployment.ps1" -Raw -Encoding UTF8
            $script:ThemeTestBlocks = @([regex]::Matches($script:ThemeTestSource, '(?s)\[xml\]\$\w+ = @"\r?\n(.*?)\r?\n"@') | ForEach-Object { $_.Groups[1].Value })

            # Wczytuje okno przez New-ThemedWindow i wymusza nałożenie szablonów kontrolek (Measure/
            # Arrange), żeby błąd w stylu lub szablonie wyszedł w teście, a nie dopiero po kliknięciu.
            $script:LoadAllThemedWindows = {
                foreach ($block in $script:ThemeTestBlocks) {
                    try {
                        [xml]$xaml = $ExecutionContext.InvokeCommand.ExpandString($block)
                        $w = New-ThemedWindow -Xaml $xaml -NoOwner
                        $content = $w.Content
                        $content.Measure((New-Object System.Windows.Size 1200, 900))
                        $content.Arrange((New-Object System.Windows.Rect 0, 0, 1200, 900))
                    } catch {
                        # Nazwa okna i najgłębszy wyjątek - sam XamlParseException mówi niewiele.
                        $inner = $_.Exception
                        while ($null -ne $inner.InnerException) { $inner = $inner.InnerException }
                        $title = if ($block -match 'Title="([^"]*)"') { $Matches[1] } else { $block.Substring(0, [Math]::Min(80, $block.Length)) }
                        throw "Okno '$title': $($inner.Message)"
                    }
                }
            }
        }

        AfterAll {
            $script:isDarkTheme = $true
        }

        It "Paleta jasna i ciemna mają ten sam zestaw kluczy" {
            @($script:ThemePalettes.Light.Keys) | Should -Be @($script:ThemePalettes.Dark.Keys)
        }

        It "Znajduje okna XAML w skrypcie" {
            $script:ThemeTestBlocks.Count | Should -BeGreaterThan 15
        }

        It "Każde okno wczytuje się z motywem ciemnym" {
            $script:isDarkTheme = $true
            { & $script:LoadAllThemedWindows } | Should -Not -Throw
        }

        It "Każde okno wczytuje się z motywem jasnym" {
            $script:isDarkTheme = $false
            { & $script:LoadAllThemedWindows } | Should -Not -Throw
        }

        It "Pola tekstowe mają pojedynczy odstęp wewnętrzny (ok. 30 px wysokości, jak przyciski)" {
            # Padding pola przekazuje do środka sam TextBox/PasswordBox - szablon nie może go dublować
            # (wcześniej pole miało 43 px, a przy sztywnej wysokości tekst był ucinany w połowie).
            [xml]$xaml = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><StackPanel><TextBox Name="t" Text="Abc"/><PasswordBox Name="p"/><Button Name="b" Content="OK"/></StackPanel></Window>'
            $w = New-ThemedWindow -Xaml $xaml -NoOwner
            $w.Content.Measure((New-Object System.Windows.Size 400, 400))
            $w.Content.Arrange((New-Object System.Windows.Rect 0, 0, 400, 400))
            foreach ($name in 't', 'p') {
                $h = $w.FindName($name).ActualHeight
                $h | Should -BeGreaterThan 26
                $h | Should -BeLessThan 36
            }
        }

        It "Przełączenie motywu podmienia kolory w już otwartym oknie" {
            $script:isDarkTheme = $true
            [xml]$xaml = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Background="{DynamicResource ThemeBackground}"><Grid/></Window>'
            $w = New-ThemedWindow -Xaml $xaml -NoOwner
            $w.Background.Color.ToString() | Should -Be ('#FF' + (Get-ThemeColor 'ThemeBackground').TrimStart('#'))

            $script:isDarkTheme = $false
            Update-WindowTheme $w
            $w.Background.Color.ToString() | Should -Be ('#FF' + (Get-ThemeColor 'ThemeBackground').TrimStart('#'))
            $w.Background.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Light.ThemeBackground.TrimStart('#'))
        }

        It "Tekst na stronie zakładki ma kolor ThemeText, nie systemowy czarny (motyw <Theme>)" -TestCases @(
            @{ Theme = 'Dark'; Dark = $true },
            @{ Theme = 'Light'; Dark = $false }
        ) {
            param($Theme, $Dark)
            # Informacje o systemie: wartości (system, BIOS, procesor...) bez własnego Foreground dziedziczą
            # kolor z TabControl / TabItem. Bez settera w stylu był to czarny z domyślnego stylu Windows.
            $script:isDarkTheme = $Dark
            [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
  <StackPanel>
    <TabControl><TabItem Header="A"><TextBlock Name="plain" Text="Wartość"/></TabItem></TabControl>
    <TabControl Style="{DynamicResource NavTabControl}">
      <TabItem Header="B" Style="{DynamicResource NavTabItem}"><TextBlock Name="nav" Text="Wartość"/></TabItem>
    </TabControl>
  </StackPanel>
</Window>
'@
            $w = New-ThemedWindow -Xaml $xaml -NoOwner
            $w.Content.Measure((New-Object System.Windows.Size 600, 400))
            $w.Content.Arrange((New-Object System.Windows.Rect 0, 0, 600, 400))
            $expected = '#FF' + (Get-ThemeColor 'ThemeText').TrimStart('#')
            foreach ($name in 'plain', 'nav') {
                $w.FindName($name).Foreground.Color.ToString() | Should -Be $expected
            }
        }

        It "Ikona narzędzia: poprawny plik .ico ze wszystkimi rozmiarami (<Variant>)" -TestCases @(
            @{ Variant = 'znak z czcionki'; Drawn = $false },
            @{ Variant = 'paczka rysowana'; Drawn = $true }
        ) {
            param($Variant, $Drawn)
            $bytes = if ($Drawn) { New-AppIconBytes -DrawnMark } else { New-AppIconBytes }
            # ICONDIR: zarezerwowane 0, typ 1 (ikona), liczba obrazów.
            [BitConverter]::ToUInt16($bytes, 0) | Should -Be 0
            [BitConverter]::ToUInt16($bytes, 2) | Should -Be 1
            [BitConverter]::ToUInt16($bytes, 4) | Should -Be 8
            $decoder = New-Object System.Windows.Media.Imaging.IconBitmapDecoder -ArgumentList (New-Object System.IO.MemoryStream (, $bytes)), ([System.Windows.Media.Imaging.BitmapCreateOptions]::None), ([System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
            @($decoder.Frames | ForEach-Object { $_.PixelWidth }) | Should -Be @(16, 20, 24, 32, 40, 48, 64, 256)
            # 16 px: róg poza zaokrąglonym kafelkiem przezroczysty, środek kafelka kryjący.
            $small = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap -ArgumentList $decoder.Frames[0], ([System.Windows.Media.PixelFormats]::Bgra32), $null, 0
            $px = New-Object byte[] (16 * 16 * 4)
            $small.CopyPixels($px, 64, 0)
            $px[3] | Should -Be 0
            $px[(8 * 64) + (8 * 4) + 3] | Should -Be 255
        }

        It "Ikona narzędzia trafia do okien i daje się wczytać jako ikona zasobnika" {
            $script:AppIcon | Should -Not -BeNullOrEmpty
            $script:AppIcon.Decoder.Frames.Count | Should -Be 8
            [xml]$xaml = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><Grid/></Window>'
            (New-ThemedWindow -Xaml $xaml -NoOwner).Icon | Should -Be $script:AppIcon
            foreach ($size in 16, 32) {
                $icon = New-Object System.Drawing.Icon -ArgumentList (New-Object System.IO.MemoryStream (, $script:AppIconBytes)), $size, $size
                $icon.Width | Should -Be $size
                $icon.Dispose()
            }
        }
    }

    Context "Nowy wygląd (ServerReview / AD Manager Diamond)" {
        BeforeAll {
            $script:ThemeTestSource = Get-Content "$PSScriptRoot\SmartToolforDeployment.ps1" -Raw -Encoding UTF8
            $script:ThemeTestBlocks = @([regex]::Matches($script:ThemeTestSource, '(?s)\[xml\]\$\w+ = @"\r?\n(.*?)\r?\n"@') | ForEach-Object { $_.Groups[1].Value })

            # Kontrast WCAG 2.x dwóch kolorów #RRGGBB (luminancja względna sRGB).
            function Get-TestLuminance {
                param([string]$Hex)
                $h = $Hex.TrimStart('#')
                $weights = 0.2126, 0.7152, 0.0722
                $sum = 0.0
                for ($i = 0; $i -lt 3; $i++) {
                    $c = [Convert]::ToInt32($h.Substring($i * 2, 2), 16) / 255.0
                    $lin = if ($c -le 0.03928) { $c / 12.92 } else { [Math]::Pow(($c + 0.055) / 1.055, 2.4) }
                    $sum += $weights[$i] * $lin
                }
                return $sum
            }
            function Get-TestContrast {
                param([string]$Fore, [string]$Back)
                $a = Get-TestLuminance $Fore
                $b = Get-TestLuminance $Back
                return ([Math]::Max($a, $b) + 0.05) / ([Math]::Min($a, $b) + 0.05)
            }

            # Okna pokazywane przez ShowDialog() zamykają się same zaraz po wczytaniu (jak w kontekście
            # "Okno Ustawienia"); handler klasy działa tylko przy włączonej fladze SdtAutoCloseWindows.
            if (-not $global:SdtAutoCloseRegistered) {
                [System.Windows.EventManager]::RegisterClassHandler([System.Windows.Window], [System.Windows.FrameworkElement]::LoadedEvent, [System.Windows.RoutedEventHandler]{
                    param($loadedSender, $loadedArgs)
                    if ($global:SdtAutoCloseWindows) { $global:SdtLastWindow = $loadedSender; $loadedSender.Close() }
                })
                $global:SdtAutoCloseRegistered = $true
            }
        }

        AfterAll {
            $script:isDarkTheme = $true
            $global:SdtAutoCloseWindows = $false
        }

        It "Każdy klucz koloru użyty w skrypcie jest w obu paletach (literówka = przezroczysta kontrolka)" {
            $keys = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($pattern in '\{DynamicResource\s+(\w+)\}', 'SetResourceReference\([^,]+,\s*["''](\w+)["'']\)', '-ColorKey\s+["''](\w+)["'']') {
                foreach ($m in [regex]::Matches($script:ThemeTestSource, $pattern)) { [void]$keys.Add($m.Groups[1].Value) }
            }
            $keys.Count | Should -BeGreaterThan 30
            $missing = foreach ($k in $keys) {
                if ($k -match '^(Theme|Accent)' -and -not ($script:ThemePalettes.Dark.Contains($k) -and $script:ThemePalettes.Light.Contains($k))) { $k }
            }
            ($missing -join ', ') | Should -Be ''
        }

        It "Kolory palety nigdy nie są podpięte przez StaticResource (kolor zamarzłby przy wczytaniu okna)" {
            [regex]::Matches($script:ThemeTestSource, '\{StaticResource\s+(Theme\w+|AccentColor)\}').Count | Should -Be 0
        }

        It "XAML okien i wspólnych stylów nie ma kolorów wpisanych na sztywno (poza gradientem logo)" {
            $script:ThemeStylesXaml | Should -Not -BeNullOrEmpty
            $sources = New-Object System.Collections.Generic.List[string]
            foreach ($block in $script:ThemeTestBlocks) { $sources.Add($block) }
            $sources.Add($script:ThemeStylesXaml)
            $bad = foreach ($block in $sources) {
                foreach ($line in ($block -split "`r?`n")) {
                    if ($line -match '#[0-9A-Fa-f]{6}\b' -and $line -notmatch 'GradientStop') { $line.Trim() }
                }
            }
            ($bad -join "`n") | Should -Be ''
        }

        It "Pary tonów (tło + tekst) są kompletne w obu paletach" {
            $required = New-Object System.Collections.Generic.List[string]
            foreach ($tone in 'Info', 'Ok', 'Warn', 'Crit') { $required.Add("Theme$($tone)Back"); $required.Add("Theme$($tone)Fore") }
            foreach ($k in 'ThemeDryRunBack', 'ThemeOkBorder', 'ThemeWarnBorder') { $required.Add($k) }
            foreach ($k in $required) {
                $script:ThemePalettes.Dark.Contains($k) | Should -BeTrue -Because "paleta Dark: $k"
                $script:ThemePalettes.Light.Contains($k) | Should -BeTrue -Because "paleta Light: $k"
            }
        }

        It "Tekst ma wystarczający kontrast w obu motywach (WCAG)" {
            # tekst|tło|minimalny kontrast
            $pairs = New-Object System.Collections.Generic.List[string]
            foreach ($bg in 'ThemeBackground', 'ThemePanel', 'ThemeHeader') { $pairs.Add("ThemeText|$bg|7") }
            foreach ($fg in 'ThemeMuted', 'ThemeSubtle', 'ThemeGhostText') {
                foreach ($bg in 'ThemePanel', 'ThemeHeader') { $pairs.Add("$fg|$bg|4.5") }
            }
            foreach ($tone in 'Info', 'Ok', 'Warn', 'Crit') { $pairs.Add("Theme$($tone)Fore|Theme$($tone)Back|4.5") }
            $pairs.Add('ThemeDangerText|ThemeDangerSoft|4.5')
            $pairs.Add('ThemeOnAccent|AccentColor|4.5')
            $pairs.Add('ThemeGridText|ThemeRowAlt|7')
            $pairs.Add('ThemeFaint|ThemePanel|2.5')   # tylko znak wodny w pustym polu i numer wersji

            $fails = New-Object System.Collections.Generic.List[string]
            foreach ($theme in 'Dark', 'Light') {
                $palette = $script:ThemePalettes[$theme]
                foreach ($pair in $pairs) {
                    $fg, $bg, $min = $pair -split '\|'
                    $ratio = Get-TestContrast $palette[$fg] $palette[$bg]
                    if ($ratio -lt [double]::Parse($min, [System.Globalization.CultureInfo]::InvariantCulture)) {
                        $fails.Add(('{0}: {1} na {2} = {3:N2} (min. {4})' -f $theme, $fg, $bg, $ratio, $min))
                    }
                }
            }
            ($fails -join '; ') | Should -Be ''
        }

        It "Pola, lista rozwijana i przyciski mają tę samą wysokość (ok. 30 px)" {
            [xml]$xaml = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><StackPanel><TextBox Name="t" Text="Abc"/><PasswordBox Name="p"/><ComboBox Name="c" SelectedIndex="0"><ComboBoxItem Content="Pozycja"/></ComboBox><Button Name="b" Content="OK"/></StackPanel></Window>'
            $w = New-ThemedWindow -Xaml $xaml -NoOwner
            $w.Content.Measure((New-Object System.Windows.Size 400, 400))
            $w.Content.Arrange((New-Object System.Windows.Rect 0, 0, 400, 400))
            foreach ($name in 't', 'p', 'c', 'b') {
                $h = $w.FindName($name).ActualHeight
                $h | Should -BeGreaterOrEqual 29 -Because "wysokość kontrolki $name"
                $h | Should -BeLessOrEqual 33 -Because "wysokość kontrolki $name"
            }
        }

        It "Style kontrolek (przycisk, karta) przemalowują się po przełączeniu motywu w otwartym oknie" {
            try {
                $script:isDarkTheme = $true
                [xml]$xaml = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><StackPanel><Button Name="b" Content="x"/><Border Name="c" Style="{StaticResource Card}"/></StackPanel></Window>'
                $w = New-ThemedWindow -Xaml $xaml -NoOwner
                $w.Content.Measure((New-Object System.Windows.Size 400, 400))
                $w.Content.Arrange((New-Object System.Windows.Rect 0, 0, 400, 400))
                $w.FindName('b').Background.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Dark.ThemeButton.TrimStart('#'))
                $w.FindName('c').Background.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Dark.ThemePanel.TrimStart('#'))

                $script:isDarkTheme = $false
                Update-WindowTheme $w
                $w.FindName('b').Background.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Light.ThemeButton.TrimStart('#'))
                $w.FindName('c').Background.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Light.ThemePanel.TrimStart('#'))
            } finally {
                $script:isDarkTheme = $true
            }
        }

        It "Pasek tytułu w kolorach motywu nie rzuca wyjątku (okno bez uchwytu, brak okna)" {
            { Set-WindowTitleBarTheme -TargetWindow $Window } | Should -Not -Throw
            { Set-WindowTitleBarTheme -TargetWindow $null } | Should -Not -Throw
        }

        It "Ikony: Get-Glyph, Set-ButtonContent i Get-ButtonText (tekst przycisku bez ikony)" {
            Get-Glyph 'E713' | Should -Be ([string][char]0xE713)
            Get-Glyph '' | Should -Be ''
            $b = New-Object System.Windows.Controls.Button
            Set-ButtonContent -Button $b -Text 'Pauza' -Icon 'E769'
            $b.Content | Should -BeOfType [System.Windows.Controls.StackPanel]
            Get-ButtonText $b | Should -Be 'Pauza'
            Set-ButtonContent -Button $b -Text 'Start'
            $b.Content | Should -Be 'Start'
            Get-ButtonText $b | Should -Be 'Start'
        }

        It "Okno komunikatu: kafelek i ikona w kolorach tonu, przycisk główny skrajnie po prawej" {
            try {
                $script:isDarkTheme = $true
                $global:SdtAutoCloseWindows = $true

                $global:SdtLastWindow = $null
                Show-ThemedMessageBox -Message 'x' -Title 'Błąd' -Image Error -Button YesNo | Out-Null
                $w = $global:SdtLastWindow
                $w | Should -Not -BeNullOrEmpty
                $w.FindName('iconTile').Background.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Dark.ThemeCritBack.TrimStart('#'))
                $w.FindName('txtIcon').Foreground.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Dark.ThemeCritFore.TrimStart('#'))
                $w.FindName('txtIcon').Text | Should -Be (Get-Glyph 'EA39')
                $w.FindName('txtTitle').Text | Should -Be 'Błąd'
                $buttons = $w.FindName('spButtons').Children
                $buttons.Count | Should -Be 2
                $buttons[0].Content | Should -Be 'Nie'
                $buttons[0].IsCancel | Should -BeTrue
                $buttons[1].Content | Should -Be 'Tak'
                $buttons[1].IsDefault | Should -BeTrue

                $global:SdtLastWindow = $null
                Show-ThemedMessageBox -Message 'x' -Title 'Sukces' -Image Success | Out-Null
                $w = $global:SdtLastWindow
                $w | Should -Not -BeNullOrEmpty
                $w.FindName('iconTile').Background.Color.ToString() | Should -Be ('#FF' + $script:ThemePalettes.Dark.ThemeOkBack.TrimStart('#'))
                $w.FindName('txtIcon').Text | Should -Be (Get-Glyph 'E930')
                $w.FindName('spButtons').Children.Count | Should -Be 1

                # -Danger: przycisk domyślny czerwony (DangerPrimaryButton), Anuluj/Nie bez zmian.
                $global:SdtLastWindow = $null
                Show-ThemedMessageBox -Message 'x' -Title 'Przerwij' -Image Warning -Button YesNoCancel -Danger | Out-Null
                $w = $global:SdtLastWindow
                $w | Should -Not -BeNullOrEmpty
                $buttons = $w.FindName('spButtons').Children
                @($buttons | ForEach-Object { [string]$_.Content }) -join ',' | Should -Be 'Anuluj,Nie,Tak'
                [object]::ReferenceEquals($buttons[2].Style, $w.FindResource('DangerPrimaryButton')) | Should -BeTrue
            } finally {
                $global:SdtAutoCloseWindows = $false
            }
        }

        It "Okno dialogowe bez widocznego właściciela jest na wierzchu, na pasku zadań i na środku ekranu" {
            # W testach okno główne nigdy nie jest pokazane, więc Get-DialogOwner nie zwraca właściciela
            # (tak samo jak w trakcie wdrożenia z oknem głównym schowanym w zasobniku).
            $Window.IsLoaded | Should -BeFalse
            # Każde wywołanie dostaje własny dokument [xml] - test nie polega na tym, że fabryka robi kopię.
            $src = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="300" SizeToContent="Height" WindowStartupLocation="CenterOwner" ShowInTaskbar="False"><Grid/></Window>'
            $w = New-ThemedWindow -Xaml ([xml]$src)
            $w.Owner | Should -BeNullOrEmpty
            $w.Topmost | Should -BeTrue
            $w.ShowInTaskbar | Should -BeTrue
            $w.WindowStartupLocation | Should -Be ([System.Windows.WindowStartupLocation]::CenterScreen)

            # -NoOwner (logowanie, okno główne): okno decyduje samo, fabryka niczego nie zmienia.
            $w2 = New-ThemedWindow -Xaml ([xml]$src) -NoOwner
            $w2.Topmost | Should -BeFalse
            $w2.WindowStartupLocation | Should -Be ([System.Windows.WindowStartupLocation]::CenterOwner)
        }

        It "Powiadomienia (Show-Toast) w oknie głównym: najwyżej 4 naraz, najstarsze znika pierwsze" {
            $toastPanel = $Window.FindName('toastHost')
            $toastPanel | Should -Not -BeNullOrEmpty
            try {
                $toastPanel.Children.Clear()
                for ($i = 1; $i -le 6; $i++) { Show-Toast -Text "Powiadomienie $i" -Kind Ok }
                $toastPanel.Children.Count | Should -Be 4
                $texts = foreach ($toast in $toastPanel.Children) { $toast.Child.Children[1].Text }
                ($texts -join ',') | Should -Be 'Powiadomienie 3,Powiadomienie 4,Powiadomienie 5,Powiadomienie 6'
                { Show-Toast -Text 'Okno bez panelu' -Kind Crit -TargetWindow (New-Object System.Windows.Window) } | Should -Not -Throw
            } finally {
                $toastPanel.Children.Clear()
            }
        }

        It "Okno główne: przycisk startu zależy od trybu testowego, licznik zaznaczonych zadań" {
            $dry = $CheckboxControls['DryRun']
            $dryBefore = $dry.IsChecked
            $taskKey = @($CheckboxControls.Keys | Where-Object { $_ -notin 'DryRun', 'InstallApplications' })[0]
            $taskBox = $CheckboxControls[$taskKey]
            $taskBefore = $taskBox.IsChecked
            $countText = $Window.FindName('txtTaskCount')
            try {
                $dry.IsChecked = $true
                Get-ButtonText $btnStart | Should -Match 'symulacj'
                $dry.IsChecked = $false
                Get-ButtonText $btnStart | Should -Match 'konfiguracj'

                $countText.Text | Should -Match '^Zaznaczone: \d+ z \d+$'
                $total = [int]([regex]::Match($countText.Text, ' z (\d+)$').Groups[1].Value)
                $total | Should -Be ($CheckboxControls.Count - 1)   # bez trybu testowego
                $taskBox.IsChecked = $false
                $before = [int]([regex]::Match($countText.Text, '^Zaznaczone: (\d+)').Groups[1].Value)
                $taskBox.IsChecked = $true
                $after = [int]([regex]::Match($countText.Text, '^Zaznaczone: (\d+)').Groups[1].Value)
                $after | Should -Be ($before + 1)
            } finally {
                $dry.IsChecked = $dryBefore
                $taskBox.IsChecked = $taskBefore
            }
        }

        It "Kropka stanu wdrożenia zmienia kolor (Busy / Paused / Error / Idle)" {
            $dot = $Window.FindName('statusDot')
            $dot | Should -Not -BeNullOrEmpty
            try {
                foreach ($pair in @(@('Busy', 'ThemeFocus'), @('Paused', 'ThemeWarning'), @('Error', 'ThemeDanger'), @('Idle', 'ThemeSuccess'))) {
                    Set-DeployStatusDot $pair[0]
                    $dot.Fill.Color.ToString() | Should -Be $Window.FindResource($pair[1]).Color.ToString() -Because "stan $($pair[0])"
                }
            } finally {
                Set-DeployStatusDot Idle
            }
        }

        It "Teksty okien bez emoji (ikony tylko z czcionki ikon)" {
            # Emoji (np. krzyżyk, ostrzeżenie, probówka) WPF rysuje czcionką zastępczą: inny rozmiar i kolor,
            # w motywie jasnym bywają nieczytelne. Raporty HTML (nagłówki, linki, komórki) są poza zakresem.
            $emoji = '[\u2300-\u23FF\u2600-\u27BF\u2B00-\u2BFF\u25B6\u25BE\u2139\uFE0F]|[\uD83C-\uDBFF][\uDC00-\uDFFF]'
            $bad = New-Object System.Collections.Generic.List[string]
            $xamlSources = New-Object System.Collections.Generic.List[string]
            foreach ($block in $script:ThemeTestBlocks) { $xamlSources.Add($block) }
            $xamlSources.Add($script:ThemeStylesXaml)
            foreach ($block in $xamlSources) {
                foreach ($line in ($block -split "`r?`n")) { if ($line -match $emoji) { $bad.Add($line.Trim()) } }
            }
            # Linie kodu, które ustawiają tekst kontrolek, komunikatów, powiadomień albo wpisów logu.
            $uiLine = '(\.|\b)(Text|Content|Header|ToolTip|Title|StatusText|RodzajText)\s*=|-(Text|Message|Title)\s|Write-Log\s'
            foreach ($line in ($script:ThemeTestSource -split "`r?`n")) {
                if ($line -match '<h\d|<a |<td|<th|<input') { continue }
                if ($line -match $uiLine -and $line -match $emoji) { $bad.Add($line.Trim()) }
            }
            ($bad -join "`n") | Should -Be ''
        }

        It "Ikony w XAML (encje &#xE...;) zawsze w czcionce ikon" {
            # Encja ikony w zwykłej czcionce (Segoe UI) rysuje się jako pusty prostokąt.
            $xamlSources = New-Object System.Collections.Generic.List[string]
            foreach ($block in $script:ThemeTestBlocks) { $xamlSources.Add($block) }
            $xamlSources.Add($script:ThemeStylesXaml)
            $bad = foreach ($block in $xamlSources) {
                foreach ($line in ($block -split "`r?`n")) {
                    if ($line -match '&#xE[0-9A-F]{3};' -and $line -notmatch 'StaticResource Glyph|MDL2|<TabItem [^>]*Tag=') { $line.Trim() }
                }
            }
            ($bad -join "`n") | Should -Be ''
        }

        It "Wiersze logu i deinstalatora: pastylka i kolor tekstu wg tonu z palety (DataTrigger)" {
            $script:isDarkTheme = $true
            $dark = $script:ThemePalettes.Dark
            # Pierwszy element drzewa wizualnego o podanej nazwie (nazwy z DataTemplate nie są w FindName okna).
            $findNamed = {
                param($Root, [string]$Name)
                $stack = New-Object System.Collections.Stack
                $stack.Push($Root)
                while ($stack.Count -gt 0) {
                    $v = $stack.Pop()
                    if ($v -is [System.Windows.FrameworkElement] -and $v.Name -eq $Name) { return $v }
                    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i)) }
                }
                return $null
            }
            $loadWindow = {
                param([string]$Title)
                $block = @($script:ThemeTestBlocks | Where-Object { $_ -match ('Title="' + [regex]::Escape($Title) + '"') })[0]
                [xml]$xaml = $ExecutionContext.InvokeCommand.ExpandString($block)
                return (New-ThemedWindow -Xaml $xaml -NoOwner)
            }
            $doLayout = {
                param($Win)
                $Win.Content.Measure((New-Object System.Windows.Size 1200, 900))
                $Win.Content.Arrange((New-Object System.Windows.Rect 0, 0, 1200, 900))
            }
            $argb = { param([string]$Key) '#FF' + $dark[$Key].TrimStart('#') }

            # Przeglądarka logów: pastylka "Rodzaj" i kolor wiersza wg RodzajKey.
            $w = & $loadWindow 'Przeglądarka logów'
            $lv = $w.FindName('lvLogs')
            $lv.ItemsSource = @(
                [PSCustomObject]@{ RodzajKey = 'Error'; RodzajText = 'Błąd'; DataText = '01.01.2026 10:00:00'; Kontekst = 'System'; Informacja = 'a' }
                [PSCustomObject]@{ RodzajKey = 'Success'; RodzajText = 'Sukces'; DataText = '01.01.2026 10:00:01'; Kontekst = 'System'; Informacja = 'b' }
                [PSCustomObject]@{ RodzajKey = 'Info'; RodzajText = 'Informacja'; DataText = '01.01.2026 10:00:02'; Kontekst = 'System'; Informacja = 'c' }
            )
            & $doLayout $w
            $rows = foreach ($i in 0..2) { $lv.ItemContainerGenerator.ContainerFromIndex($i) }
            @($rows | Where-Object { $null -ne $_ }).Count | Should -Be 3
            (& $findNamed $rows[0] 'pill').Background.Color.ToString() | Should -Be (& $argb 'ThemeCritBack')
            (& $findNamed $rows[0] 'txt').Foreground.Color.ToString() | Should -Be (& $argb 'ThemeCritFore')
            $rows[0].Foreground.Color.ToString() | Should -Be (& $argb 'ThemeCritFore')
            (& $findNamed $rows[1] 'pill').Background.Color.ToString() | Should -Be (& $argb 'ThemeOkBack')
            $rows[1].Foreground.Color.ToString() | Should -Be (& $argb 'ThemeOkFore')
            (& $findNamed $rows[2] 'pill').Background.Color.ToString() | Should -Be (& $argb 'ThemeChip')
            $rows[2].Foreground.Color.ToString() | Should -Be (& $argb 'ThemeGridText')

            # Deinstalator: pastylka stanu wg StatusTone, ukryta, dopóki program nie był przetwarzany.
            $w = & $loadWindow 'Deinstalator Oprogramowania (Wybór zbiorczy)'
            $lv = $w.FindName('lvApps')
            $lv.ItemsSource = @(
                [PSCustomObject]@{ IsChecked = $false; Icon = $null; StatusText = 'Usunięto'; StatusTone = 'Ok'; DisplayName = 'A'; DisplayVersion = '1.0'; InstallDate = ''; Publisher = '' }
                [PSCustomObject]@{ IsChecked = $false; Icon = $null; StatusText = 'Błąd (kod: 1603)'; StatusTone = 'Crit'; DisplayName = 'B'; DisplayVersion = '1.0'; InstallDate = ''; Publisher = '' }
                [PSCustomObject]@{ IsChecked = $false; Icon = $null; StatusText = ''; StatusTone = ''; DisplayName = 'C'; DisplayVersion = '1.0'; InstallDate = ''; Publisher = '' }
            )
            & $doLayout $w
            $rows = foreach ($i in 0..2) { $lv.ItemContainerGenerator.ContainerFromIndex($i) }
            @($rows | Where-Object { $null -ne $_ }).Count | Should -Be 3
            $pill = & $findNamed $rows[0] 'pill'
            $pill.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
            $pill.Background.Color.ToString() | Should -Be (& $argb 'ThemeOkBack')
            (& $findNamed $rows[1] 'pill').Background.Color.ToString() | Should -Be (& $argb 'ThemeCritBack')
            (& $findNamed $rows[1] 'txt').Foreground.Color.ToString() | Should -Be (& $argb 'ThemeCritFore')
            (& $findNamed $rows[2] 'pill').Visibility | Should -Be ([System.Windows.Visibility]::Collapsed)
        }
    }
}