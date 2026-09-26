
# ------------------------------------------
# Automatyczna elewacja uprawnień (UAC)
# ------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and $null -eq $global:PesterTesting) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not [string]::IsNullOrWhiteSpace($scriptPath)) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    }
    exit
}

# ------------------------------------------
# Siatka bezpieczeństwa na start: bez tego, nieobsłużony błąd gdziekolwiek w reszcie skryptu
# (np. zanim zdąży pokazać się okno GUI) po prostu ubijał cały proces - a że konsola PowerShell
# uruchomiona przez UAC/"Uruchom jako administrator" (powershell.exe -File ...) zamyka się
# natychmiast po zakończeniu skryptu (sukces czy błąd), użytkownik widział tylko znikające okno
# bez żadnego komunikatu. Teraz taki błąd jest pokazany, zapisany do pliku i konsola czeka na
# Enter zamiast się zamykać w milczeniu.
trap {
    # W testach Pester ten trap NIE MOŻE wołać exit - to ubiłoby cały proces Invoke-Pester, nie
    # tylko pojedynczy test (dokładnie ten sam rodzaj błędu, co niegdyś ubijało WSZYSTKIE testy
    # przez elewację UAC na starcie). Błędy oczekiwane przez testy (np. Should -Throw) i tak nie
    # dotrą tutaj - Pester łapie je lokalnie, zanim zdążą "uciec" aż do tego trapa. Jeśli mimo to
    # coś tu trafi podczas testów, po prostu wracamy do normalnego biegu (continue), zamiast
    # zamykać proces.
    if ($null -ne $global:PesterTesting) { continue }

    $errText = "Nieoczekiwany błąd uniemożliwił uruchomienie narzędzia:`n`n$($_.Exception.Message)"
    try {
        $logPath = Join-Path $env:TEMP "SmartToolforDeployment_crash.log"
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $($_ | Out-String) | $($_.ScriptStackTrace)" | Out-File -FilePath $logPath -Append -Encoding UTF8
        $errText += "`n`nSzczegóły zapisano w: $logPath"
    } catch {}
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show($errText, "Smart Tool for Deployment - Błąd krytyczny", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } catch {
        Write-Host $errText -ForegroundColor Red
    }
    try { Read-Host "Naciśnij Enter, aby zamknąć to okno" | Out-Null } catch {}
    exit 1
}

# ------------------------------------------
# Funkcja pokazująca okienko powitalne
# ------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

function Apply-ThemeToWindow($dlg) {
    if ($Window -is [System.Windows.Window]) {
        $dlg.Owner = $Window
        $keys = @("ThemeBackground", "ThemePanel", "ThemeText", "ThemeButton", "ThemeButtonText", "ThemeBorder", "ThemeTextBoxBg")
        foreach ($key in $keys) {
            $dlg.Resources[$key] = $Window.Resources[$key]
        }
    }
}

function global:Show-ThemedMessageBox {
    param(
        [string]$Message,
        [string]$Title = "Informacja",
            [string]$Button = "OK",
            [string]$Image = "Information"
    )
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Width="450" SizeToContent="Height" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="15,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Margin" Value="10,0,0,0"/>
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="MinWidth" Value="85"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,25" MaxWidth="390">
            <TextBlock Name="txtIcon" FontSize="36" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <TextBlock Name="txtMessage" FontSize="14" TextWrapping="Wrap" VerticalAlignment="Center" Width="330"/>
        </StackPanel>
        <StackPanel Name="spButtons" Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right"/>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    $dlg.Title = $Title
    try { Apply-ThemeToWindow $dlg } catch {}
    if ($null -eq $dlg.Resources["ThemeBackground"]) {
        $dlg.Resources["ThemeBackground"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF202225")
        $dlg.Resources["ThemeText"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFDCDDDE")
        $dlg.Resources["ThemeButton"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF4F545C")
        $dlg.Resources["ThemeButtonText"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("White")
    }
    $txtMessage = $dlg.FindName("txtMessage")
    $txtMessage.Text = $Message
    $txtIcon = $dlg.FindName("txtIcon")
    
    if ($Image -match 'Error') { $imgStr = 'Error' }
    elseif ($Image -match 'Warning') { $imgStr = 'Warning' }
    elseif ($Image -match 'Question') { $imgStr = 'Question' }
    else { $imgStr = 'Information' }
    
    switch ($imgStr) {
        "Error" { $txtIcon.Text = "❌"; $txtIcon.Foreground = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFC50F1F") }
        "Warning" { $txtIcon.Text = "⚠️"; $txtIcon.Foreground = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFE6A100") }
        "Question" { $txtIcon.Text = "❓"; $txtIcon.Foreground = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF0078D7") }
        default { $txtIcon.Text = "ℹ️"; $txtIcon.Foreground = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF107C10") }
    }
    $spButtons = $dlg.FindName("spButtons")
    $script:msgBoxResult = [System.Windows.MessageBoxResult]::None
    $AddBtn = {
        param($content, $resVal, $isDef, $isCanc, $bgHex)
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = $content
        $btn.IsDefault = $isDef
        $btn.IsCancel = $isCanc
        $btn.Tag = $resVal
        if ($bgHex) { $btn.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString($bgHex); $btn.Foreground = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("White") }
        $btn.Add_Click({ 
            $script:msgBoxResult = $this.Tag
            $dlg.Close() 
        })
        $spButtons.Children.Add($btn) | Out-Null
    }
    
    if ($Button -match 'YesNoCancel') { $btnStr = 'YesNoCancel' }
    elseif ($Button -match 'OKCancel') { $btnStr = 'OKCancel' }
    elseif ($Button -match 'YesNo') { $btnStr = 'YesNo' }
    else { $btnStr = 'OK' }
    
    # UWAGA: wartości enum MUSZĄ być w nawiasach. Przy wywołaniu polecenia (& $AddBtn ...) PowerShell
    # działa w "trybie argumentów", w którym [System.Windows.MessageBoxResult]::Yes bez nawiasów NIE
    # jest wyliczane, tylko przekazywane jako zwykły napis "[System.Windows.MessageBoxResult]::Yes".
    # Porównanie takiego napisu z enumem (-eq [System.Windows.MessageBoxResult]::Yes) zawsze dawało
    # $false, więc każde "Tak" działało jak "Nie". Nawias wymusza wyliczenie wyrażenia.
    switch ($btnStr) {
        "OKCancel" { & $AddBtn "OK" ([System.Windows.MessageBoxResult]::OK) $true $false "#FF0078D7"; & $AddBtn "Anuluj" ([System.Windows.MessageBoxResult]::Cancel) $false $true $null }
        "YesNo" { & $AddBtn "Tak" ([System.Windows.MessageBoxResult]::Yes) $true $false "#FF0078D7"; & $AddBtn "Nie" ([System.Windows.MessageBoxResult]::No) $false $true $null }
        "YesNoCancel" { & $AddBtn "Tak" ([System.Windows.MessageBoxResult]::Yes) $true $false "#FF0078D7"; & $AddBtn "Nie" ([System.Windows.MessageBoxResult]::No) $false $false $null; & $AddBtn "Anuluj" ([System.Windows.MessageBoxResult]::Cancel) $false $true $null }
        default { & $AddBtn "OK" ([System.Windows.MessageBoxResult]::OK) $true $false "#FF0078D7" }
    }
    if ($null -ne $Window -and $Window.IsLoaded) { $dlg.Owner = $Window; $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner }
    $dlg.ShowDialog() | Out-Null
    return $script:msgBoxResult
}

# ---------- Ekran autoryzacji (PIN) ----------
[xml]$authXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Autoryzacja STD" Width="350" SizeToContent="Height" WindowStartupLocation="CenterScreen"
        Background="#FF202225" Foreground="#FFDCDDDE" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#FF4F545C"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#FF1E1E1E"/>
            <Setter Property="Foreground" Value="#FFDCDDDE"/>
            <Setter Property="BorderBrush" Value="#FF4F545C"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,5"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="#FF1E1E1E"/>
            <Setter Property="Foreground" Value="#FFDCDDDE"/>
            <Setter Property="BorderBrush" Value="#FF4F545C"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,5"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Wprowadź dane dostępowe" FontSize="18" FontWeight="Bold" Margin="0,0,0,15" HorizontalAlignment="Center"/>
        
        <TextBlock Text="Twój identyfikator / Login:" Grid.Row="1" Margin="0,0,0,5"/>
        <TextBox Name="txtLogin" Grid.Row="2" Margin="0,0,0,15"/>
        
        <TextBlock Text="Wprowadź PIN:" Grid.Row="3" Margin="0,0,0,5"/>
        <PasswordBox Name="txtPin" Grid.Row="4" Margin="0,0,0,20"/>
        
        <Button Name="btnLogin" Content="Odblokuj narzędzie" Grid.Row="5" Height="35" Background="#FF107C10" FontWeight="SemiBold" IsDefault="True"/>
    </Grid>
</Window>
"@

$readerAuth = New-Object System.Xml.XmlNodeReader $authXaml
$authWindow = [Windows.Markup.XamlReader]::Load($readerAuth)
$txtLogin = $authWindow.FindName("txtLogin")
$txtPin = $authWindow.FindName("txtPin")
$btnLoginAuth = $authWindow.FindName("btnLogin")
$script:authSuccess = $false
$script:failedAttempts = 0

$btnLoginAuth.Add_Click({
    if ($txtPin.Password -eq "2137" -and -not [string]::IsNullOrWhiteSpace($txtLogin.Text)) {
        $script:authSuccess = $true
        $script:OperatorLogin = $txtLogin.Text
        $authWindow.Close()
    } else {
        $script:failedAttempts++
        if ($script:failedAttempts -ge 3) {
            Show-ThemedMessageBox -Message "Przekroczono limit błędnych prób (3). Aplikacja zostanie zamknięta." -Title "Blokada" -Button "OK" -Image "Error" | Out-Null
            $authWindow.Close()
        } else {
            $pozostalo = 3 - $script:failedAttempts
            Show-ThemedMessageBox -Message "Nieprawidłowy login lub PIN.`nPozostało prób: $pozostalo" -Title "Błąd autoryzacji" -Button "OK" -Image "Error" | Out-Null
            $txtPin.Clear()
        }
    }
})

$authWindow.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Escape) {
        $authWindow.Close()
    }
})

$authWindow.ShowDialog() | Out-Null

if (-not $script:authSuccess) {
    exit
}

# Zapisz login od razu do logów po pomyślnej autoryzacji
$authLogLine = "[$((Get-Date).ToString('HH:mm:ss'))] Zalogowano operatora narzędzia STD: $($script:OperatorLogin)"
Add-Content -Path "C:\deploy-log.txt" -Value $authLogLine -Encoding UTF8 -ErrorAction SilentlyContinue


# ---------- Utworzenie formularza (WPF) ----------
[xml]$welcomeXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Potwierdzenie" Width="650" SizeToContent="Height" WindowStartupLocation="CenterScreen"
        Background="#FF202225" Foreground="#FFDCDDDE" FontFamily="Segoe UI" ResizeMode="NoResize"
        WindowStyle="ToolWindow" Topmost="True">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#FF4F545C"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}" CornerRadius="2">
                            <GridViewRowPresenter Content="{TemplateBinding Content}" Columns="{TemplateBinding GridView.ColumnCollection}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Border Background="#FF2F3136" CornerRadius="8" Padding="20">
            <TextBlock TextWrapping="Wrap" FontSize="16" LineHeight="26">
                <Run FontWeight="Bold" FontSize="22" Foreground="#FF0078D7" Text="Smart Tool for Deployment&#10;&#10;"/>
                <Run Text="Celem niniejszego skryptu jest wsparcie przy szybkiej i skutecznej konfiguracji komputera.&#10;"/>
                <Run Text="Skrypt nie zastępuje decyzji ani uwag inżynierów IT.&#10;&#10;"/>
                <Run FontWeight="SemiBold" Text="Czy chcesz kontynuować?"/>
            </TextBlock>
        </Border>
        
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
            <Button Name="btnOk" Content="OK (3)" Width="120" Height="40" Margin="0,0,15,0" Background="#FF107C10" IsEnabled="False"/>
            <Button Name="btnCancel" Content="Anuluj" Width="120" Height="40" IsEnabled="False" Background="#FFD83B01" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $welcomeXaml
$welcomeWindow = [Windows.Markup.XamlReader]::Load($reader)

$btnOk = $welcomeWindow.FindName("btnOk")
$btnCancel = $welcomeWindow.FindName("btnCancel")

$script:countdown = 3
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)

$timer.Add_Tick({
    $script:countdown--
    if ($script:countdown -gt 0) {
        $btnOk.Content = "OK ($script:countdown)"
    } else {
        $btnOk.Content = "OK"
        $btnOk.IsEnabled = $true
        $btnCancel.IsEnabled = $true
        $timer.Stop()
    }
})
$timer.Start()

$btnOk.Add_Click({
    $welcomeWindow.DialogResult = $true
    $welcomeWindow.Close()
})

$btnCancel.Add_Click({
    $welcomeWindow.DialogResult = $false
    $welcomeWindow.Close()
})

$result = $welcomeWindow.ShowDialog()

if ($result -ne $true) {
    exit
}

$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    if ($MyInvocation.MyCommand.Path) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
}
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    try { $ScriptDir = [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch {}
}
if ([string]::IsNullOrWhiteSpace($ScriptDir)) { $ScriptDir = $PWD.Path }

$script:ScriptVersion = "3.1.0"
$configPath = Join-Path $ScriptDir "config.json"
$script:LogFilePath = "C:\deploy-log.txt"
$script:ErrorLogFilePath = "C:\deploy-error-log.txt"
$script:ValidationWarnings = New-Object System.Collections.Generic.List[string]
$script:ValidationErrors   = New-Object System.Collections.Generic.List[string]
$script:SelectedApps = @{}
$CheckboxControls = @{}
$checkboxOptions = [ordered]@{
    "DryRun"                = @{
        "Text"    = "🧪 Tryb testowy (Dry-Run - bez rzeczywistych zmian)"
        "Tooltip" = "Symuluje wdrożenie: loguje co zostałoby zrobione (instalacje, deinstalacje, zmiany rejestru, dołączenie do domeny, BitLocker itd.), ale nie wykonuje żadnej z tych operacji naprawdę. Przydatne do sprawdzenia poprawności konfiguracji przed wdrożeniem na realnej stacji."
        "Enabled" = $false
    }
    "CreateRestorePoint"    = @{
        "Text"    = "Utwórz punkt przywracania systemu przed startem"
        "Tooltip" = "Tworzy punkt przywracania systemu Windows tuż przed rozpoczęciem wdrożenia, aby w razie problemu można było łatwo cofnąć zmiany w systemie (System Restore). Nie obejmuje plików użytkownika ani zainstalowanych aplikacji spoza mechanizmu Restore."
        "Enabled" = $false
    }
    "WaitForNetwork"        = @{
        "Text"    = "Czekaj na połączenie z siecią przed startem"
        "Tooltip" = "Wstrzymuje konfigurację do momentu podłączenia kabla sieciowego lub Wi-Fi i uzyskania adresu IP."
        "Enabled" = $true
    }
    "SuspendHibernation"    = @{
        "Text"    = "Wstrzymaj usypianie i hibernację"
        "Tooltip" = "Tymczasowo zapobiega usypianiu i hibernacji komputera podczas działania skryptu."
        "Enabled" = $true
    }
    "ImportWiFiProfile"     = @{
        "Text"    = "Importuj profil Wi-Fi"
        "Tooltip" = "Importuje zapisany profil Wi-Fi z pliku (Lizard-Tech)."
        "Enabled" = $false
    }
    "RunPostInstallScripts" = @{
        "Text"    = "Uruchom skrypty poinstalacyjne"
        "Tooltip" = "Wykonuje dodatkowe skrypty (.ps1, .bat) zdefiniowane w pliku JSON w sekcji PostInstallScripts."
        "Enabled" = $false
    }
    "ExportHardwareAudit"   = @{
        "Text"    = "Eksportuj audyt sprzętowy"
        "Tooltip" = "Zapisuje pełny raport sprzętowy maszyny w zdefiniowanej ścieżce."
        "Enabled" = $false
    }
    "UninstallMicrosoft365" = @{
        "Text"    = "Odinstaluj preinstalowane produkty Microsoft 365"
        "Tooltip" = "Usuwa preinstalowane aplikacje Microsoft 365."
        "Enabled" = $false
    }
    "UninstallOneDrive"     = @{
        "Text"    = "Odinstaluj OneDrive"
        "Tooltip" = "Zatrzymuje proces i usuwa klienta OneDrive z systemu."
        "Enabled" = $false
    }
    "RemoveBloatware"       = @{
        "Text"    = "Usuń preinstalowane aplikacje (Bloatware)"
        "Tooltip" = "Usuwa zbędne aplikacje Appx (Xbox, TikTok, Solitaire itp.) z systemu."
        "Enabled" = $true
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
    "EnableBitLocker"       = @{
        "Text"    = "Zaszyfruj dysk systemowy (BitLocker TPM)"
        "Tooltip" = "Włącza po cichu szyfrowanie BitLocker na dysku C: i eksportuje klucz odzyskiwania na serwer."
        "Enabled" = $false
    }
    "AutoReboot"            = @{
        "Text"    = "Uruchom ponownie po zakończeniu"
        "Tooltip" = "Automatycznie uruchamia komputer ponownie po wdrożeniu (wymagane m.in. po zmianie nazwy i domeny)."
        "Enabled" = $false
    }
}

# --- Core Logic Functions ---
function global:Do-WpfEvents {
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [System.Action]{ $frame.Continue = $false }
    ) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        
        while ($script:isPaused -and -not $script:isCancelled) {
            Start-Sleep -Milliseconds 100
            $frame2 = New-Object System.Windows.Threading.DispatcherFrame
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [System.Action]{ $frame2.Continue = $false }
            ) | Out-Null
            [System.Windows.Threading.Dispatcher]::PushFrame($frame2)
        }
}

function global:Set-ProgressText {
    param([string]$Text)
    if ($null -ne $txtProgressInfo) {
        if ($txtProgressInfo.Dispatcher.CheckAccess()) {
            $txtProgressInfo.Text = $Text
        } else {
            $txtProgressInfo.Dispatcher.Invoke([Action]{ $txtProgressInfo.Text = $Text })
        }
    }
}

function global:Step-DeploymentProgress {
    if ($null -ne $script:TotalDeploymentSteps -and $script:TotalDeploymentSteps -gt 0) {
        $script:CurrentDeploymentStep++
        $target = [int](($script:CurrentDeploymentStep / $script:TotalDeploymentSteps) * 100)
        if ($target -gt 100) { $target = 100 }
        if ($null -ne $progressBar) {
            $progressBar.Value = $target
            # Sztuczka WPF z nadpisaniem wartości, aby zniwelować wolną systemową animację
            if ($target -lt 100) {
                $progressBar.Value = $target + 1
                $progressBar.Value = $target
                } else {
                    $progressBar.Value = 100
            }
        }
        Do-WpfEvents
    }
}

function Remove-Bloatware {
    Write-Log "Rozpoczynam usuwanie preinstalowanego Bloatware..."
    $bloatwareApps = @(
        "*BingNews*", "*BingWeather*", "*MicrosoftOfficeHub*", "*SkypeApp*",
        "*SolitaireCollection*", "*XboxApp*", "*XboxGamingOverlay*", "*XboxSpeechToTextOverlay*",
        "*YourPhone*", "*TikTok*", "*Spotify*", "*Facebook*", "*Instagram*", "*Twitter*",
        "*LinkedIn*", "*Netflix*", "*PandoraMedia*", "*CandyCrush*", "*Disney*"
    )
    foreach ($app in $bloatwareApps) {
        if ($script:isCancelled) { break }
        Do-WpfEvents
        if ($script:DryRun) {
            # -ErrorAction SilentlyContinue NIE tłumi błędu ładowania modułu Appx (np. na
            # niektórych kompilacjach Windows 11: "Operation is not supported on this platform") -
            # to osobny mechanizm od błędów samego cmdletu, więc bez try/catch przerywał całą
            # pętlę (i uniemożliwiał Przerwanie/Pauzę do końca jej trwania).
            try {
                $found = Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue
                if ($found) { Write-Log "[DRY-RUN] Usunięto by: $app" }
            } catch {
                Write-Log "[DRY-RUN] Nie udało się sprawdzić pakietu ${app}: $_" -IsError
            }
            continue
        }
        try {
            Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
            Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like $app } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
        } catch {
            Write-Log "Błąd podczas usuwania pakietu ${app}: $_" -IsError
        }
    }
    Write-Log "Zakończono usuwanie Bloatware."
}

function Wait-ForNetwork {
    Write-Log "Sprawdzanie dostępności sieci..."
    $attempts = 0
    
    while ($true) {
        if ($script:isCancelled) { break }
        $isAvailable = [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
        $validIp = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notmatch "^169\.254\." -and $_.IPAddress -ne "127.0.0.1" }
        
        if ($isAvailable -and $validIp) {
            Write-Log "Wykryto połączenie z siecią ($($validIp[0].IPAddress))."
            break
        }

        if ($attempts % 5 -eq 0) {
            Write-Log "Brak poprawnego IP. Oczekiwanie na sieć..."
        }
        for ($i = 0; $i -lt 10; $i++) {
            Do-WpfEvents
            Start-Sleep -Milliseconds 200
        }
        $attempts++
    }
}



function Test-ConfigurationFile {
    param([switch]$Silent)
    if (-not (Test-Path $configPath)) {
        if (-not $Silent) { Write-Log "Brak pliku: $configPath" -IsError }
        return $false
    }
    try {
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $config.Programs) {
            if (-not $Silent) { Write-Log "Nie znaleziono sekcji 'Programs' w $configPath" -IsError }
            return $false
        }
        return $true
    }
    catch {
        if (-not $Silent) { Write-Log "Błąd podczas odczytu ${configPath}: $_" -IsError }
        return $false
    }
    
}

function Ensure-Configuration {
    while (-not (Test-ConfigurationFile -Silent)) {
        if (-not (Test-Path -LiteralPath $configPath)) {
            $msg = "Nie znaleziono pliku konfiguracyjnego. Zostanie utworzony domyślny szablon w lokalizacji:`n$configPath`n`nMożesz go później edytować w Ustawieniach."
            Show-ThemedMessageBox -Message $msg -Title "Tworzenie konfiguracji" -Button "OK" -Image "Information" | Out-Null
            
            $defaultConfig = [ordered]@{
                DefaultInstallSource = "network"
                InstallSourcePaths = @{ network = "\\server\share\"; web = "https://example.com/apps/" }
                CustomWebDataLocation = @{ URL = "" }
                DomainJoin = @{ DomainName = ""; Username = "" }
                LocalAdmin = @{ Username = "Admin" }
                SystemSettings = @{ DisableDeliveryOptimization=$true; EnableWin10StartMenu=$false; DisableTelemetry=$true; DisableCortana=$true; DisableFastStartup=$true; DisableNewsAndInterests=$true; CustomRegistry=@() }
                WebAuth = @{ Username = ""; Password = "" }
                TeamViewer = @{ FileName = ""; Arguments = "" }
                AntyVirus = @{ DefaultInstallSource = "network"; InstallSourcePaths = @{ network = ""; web = "" }; Credentials = @{ Username = ""; Password = "" }; FileName = "" }
                WiFiProfile = @{ FileName = "" }
                Programs = @{}
                Profiles = @{
                    "Standard" = @("Chrome", "7zip", "PowerToys")
                    "Księgowość" = @("Chrome", "7zip", "Szafir_KIR", "AdobeReader")
                }
                PostInstallScripts = @()
                HardwareAudit = @{ ExportPath = "C:\Audit\" }
                DefaultCheckboxes = @{
                    DryRun = $false; CreateRestorePoint = $false
                    WaitForNetwork = $true; SuspendHibernation = $true; ImportWiFiProfile = $false; UninstallMicrosoft365 = $false
                    UninstallOneDrive = $false; RemoveBloatware = $true; InstallTeamViewer = $true; InstallApplications = $true
                    CreateLocalAdmin = $true; InstallAV = $true; JoinDomain = $true; ChangeSystemSettings = $true
                    RunWindowsUpdate = $true; ChangeComputerName = $false; RunPostInstallScripts = $false; ExportHardwareAudit = $false
                    EnableBitLocker = $false; AutoReboot = $false; JoinIntune = $false
                }
                AutoUpdate = @{ Enabled = $false; VersionCheckPath = "" }
                DarkTheme = $true
            }
            try {
                $defaultConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8
            } catch {
                Show-ThemedMessageBox -Message "Nie udało się utworzyć pliku: $_" -Title "Błąd krytyczny" -Button "OK" -Image "Error" | Out-Null
                exit 1
            }
        } else {
            Show-ThemedMessageBox -Message "Plik konfiguracyjny jest uszkodzony. Popraw go ręcznie lub usuń, aby utworzyć nowy." -Title "Błąd krytyczny" -Button "OK" -Image "Error" | Out-Null
            exit 1 # Wyjście, jeśli plik istnieje, ale jest niepoprawny
        }
    }
}


function Write-Log {
    param(
        [string]$Text,
        [switch]$IsError,
        # System = komunikat samego narzędzia (walidacja, zapis configu, cykl życia okna),
        # Automat = krok wykonywany automatycznie w ramach sekwencji wdrożenia (Start-Deployment),
        # Użytkownik = bezpośredni skutek kliknięcia/decyzji operatora.
        # Gdy nieustawiony, bierzemy $script:CurrentLogContext (patrz Start-Deployment), a jak i
        # tego brak - domyślnie "System".
        [ValidateSet("System", "Automat", "Użytkownik")]
        [string]$Context
    )
    if ([string]::IsNullOrWhiteSpace($Context)) {
        $Context = if (-not [string]::IsNullOrWhiteSpace($script:CurrentLogContext)) { $script:CurrentLogContext } else { "System" }
    }
    # Pełna data+godzina (nie tylko HH:mm:ss), poziom i kontekst wpisane wprost w linii - dzięki
    # temu C:\deploy-log.txt da się jednoznacznie sparsować (data + poziom + kontekst + treść)
    # zarówno przy przeglądaniu logów z wielu dni, jak i w Show-LogWindow (sortowalne kolumny,
    # filtrowanie, eksport raportu).
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $level = if ($IsError) { "ERROR" } else { "INFO" }
    $line = "[$timestamp] [$level] [$Context] $Text"

    if ($IsError) {
        try { [System.Media.SystemSounds]::Hand.Play() } catch { }
    }

    # Log to GUI
    try {
        if ($null -ne $rtbLog) {
            if ($rtbLog.Dispatcher.CheckAccess()) {
                $rtbLog.AppendText("$line`r`n")
                try { $rtbLog.ScrollToEnd() } catch { }
            } else {
                $rtbLog.Dispatcher.Invoke([Action]{
                    $rtbLog.AppendText("$line`r`n")
                    try { $rtbLog.ScrollToEnd() } catch { }
                })
            }
        }
    } catch { }

    # Log to File
    try {
        $line | Add-Content -Path $script:LogFilePath -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($IsError) {
            $line | Add-Content -Path $script:ErrorLogFilePath -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch { }
}

function Start-ProcessWithEvents {
    param (
        [string]$FilePath,
        [string]$ArgumentList
    )
    if ($script:DryRun) {
        Write-Log "[DRY-RUN] Pominięto uruchomienie: `"$FilePath`" $ArgumentList"
        return 0
    }
    try {
        # -ArgumentList przekazujemy tylko, gdy nie jest pusty: w Windows PowerShell 5.1 parametr ma
        # [ValidateNotNullOrEmpty], więc -ArgumentList "" (np. antywirus bez argumentów) kończył się
        # błędem "Cannot validate argument on parameter 'ArgumentList'".
        $startParams = @{ FilePath = $FilePath; PassThru = $true; NoNewWindow = $true; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($ArgumentList)) { $startParams.ArgumentList = $ArgumentList }
        $proc = Start-Process @startParams
        if ($null -ne $proc) {
            # Odczyt Handle zaraz po starcie "przypina" uchwyt procesu - bez tego w PS 5.1 ExitCode
            # po zakończeniu procesu bywa pusty ($null).
            $null = $proc.Handle
            while (-not $proc.HasExited) {
                Do-WpfEvents
                if ($script:isCancelled) {
                    try { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } catch {}
                    # Czekamy chwilę na faktyczne zakończenie, inaczej odczyt ExitCode rzuca wyjątek.
                    try { $proc.WaitForExit(3000) | Out-Null } catch {}
                    Write-Log "Proces przerwany." -IsError
                    break
                }
                Start-Sleep -Milliseconds 100
            }
            if ($proc.HasExited) { return $proc.ExitCode }
            return $null
        }
    } catch {
        Write-Log "Błąd uruchamiania procesu: $_" -IsError
    }
    return $null
}

function Start-UninstallProcessWithCancel {
    param (
        [string]$FilePath,
        [string]$ArgumentList,
        [string]$LogContext
    )
    try {
        $startParams = @{ FilePath = $FilePath; PassThru = $true; WindowStyle = 'Hidden'; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($ArgumentList)) { $startParams.ArgumentList = $ArgumentList }
        $proc = Start-Process @startParams
    } catch {
        Write-Log "Błąd uruchamiania deinstalatora ($LogContext): $_" -IsError
        return -1
    }
    # Bez tego odczytu ExitCode procesu uruchomionego z -WindowStyle Hidden w PS 5.1 często jest
    # pusty, a wtedy poprawna deinstalacja była pokazywana jako "BŁĄD (kod: -1)".
    try { $null = $proc.Handle } catch {}
    $script:uninstProc = $proc
    while (-not $proc.HasExited) {
        Do-WpfEvents
        if ($script:isCancelledFromUninstall) {
            try { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } catch {}
            try { $proc.WaitForExit(3000) } catch {}
            Write-Log "Proces deinstalacji przerwany ($LogContext)." -IsError
            break
        }
        Start-Sleep -Milliseconds 100
    }
    $exitCode = -1
    try { if ($proc.HasExited -and $null -ne $proc.ExitCode) { $exitCode = $proc.ExitCode } } catch {}
    return $exitCode
}

function Resolve-OfficeClickToRunUninstallInfo {
    param([string]$Cmd)
    $productId = $null
    if ($Cmd -match 'productstoremove="?([^"\s]+)"?') {
        $productId = $matches[1]
    } elseif ($Cmd -match 'ProductID="?([^"\s]+)"?') {
        $productId = $matches[1]
    }

    $exePath = $null
    if ($Cmd -match '^\s*"([^"]+)"') {
        $exePath = $matches[1]
    } elseif ($Cmd -match '^\s*(\S+\.exe)') {
        $exePath = $matches[1]
    }

    return [PSCustomObject]@{ ProductId = $productId; ExePath = $exePath }
}

function Invoke-DownloadFile {
    param (
        [string]$Uri,
        [string]$OutFile,
        [pscredential]$Credential
    )
    
    $progressBarDownload.Value = 0
    
    $webClient = New-Object System.Net.WebClient
    if ($null -ne $Credential) {
        $webClient.Credentials = $Credential.GetNetworkCredential()
    }

    $syncHash = [hashtable]::Synchronized(@{
        IsDone = $false
        Error = $null
    })

    $onProgress = {
        param($eventSender, $e)
        $pct = $e.ProgressPercentage
        if ($null -ne $progressBarDownload -and $progressBarDownload.Dispatcher) {
            $progressBarDownload.Dispatcher.Invoke([Action]{ $progressBarDownload.Value = $pct })
        }
    }
    
    $onComplete = {
        param($eventSender, $e)
        if ($e.Error) { $syncHash.Error = $e.Error }
        $syncHash.IsDone = $true
    }

    $webClient.add_DownloadProgressChanged($onProgress)
    $webClient.add_DownloadFileCompleted($onComplete)

    try {
        $webClient.DownloadFileAsync([uri]$Uri, $OutFile)
        while (-not $syncHash.IsDone) {
            Do-WpfEvents
            if ($script:isCancelled) {
                $webClient.CancelAsync()
                Write-Log "Pobieranie przerwane." -IsError
                # Rzucamy wyjątek, żeby wywołujący NIE uruchomił niepełnego pliku instalatora
                # (wcześniej funkcja po prostu wracała i kod szedł dalej, jakby pobieranie się udało).
                throw "Pobieranie przerwane przez użytkownika."
            }
            Start-Sleep -Milliseconds 50
        }
        if ($syncHash.Error) { throw $syncHash.Error }
    }
    finally {
        $webClient.remove_DownloadProgressChanged($onProgress)
        $webClient.remove_DownloadFileCompleted($onComplete)
        $webClient.Dispose()
        $progressBarDownload.Value = 0
    }
}

function Test-UrlValid {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    try {
        $uri = $null
        if ([System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) {
            return ($uri.Scheme -in @('http','https'))
        }
        return $false
    } catch { return $false }
}

# Łączy źródło instalacji (URL albo ścieżka UNC/lokalna) z nazwą pliku. Dla URL pilnuje dokładnie
# jednego "/" między częściami (wcześniej "https://x/Data" + "a.exe" dawało "https://x/Dataa.exe"),
# dla ścieżek plikowych używa Join-Path.
function Join-InstallSource {
    param([string]$BasePath, [string]$FileName)
    if ($BasePath -match '^https?://') {
        return "$($BasePath.TrimEnd('/'))/$($FileName.TrimStart('/'))"
    }
    return (Join-Path $BasePath $FileName)
}

# Kody wyjścia instalatorów uznawane za sukces:
#   0    - sukces
#   1641 - sukces, instalator MSI sam zainicjował restart
#   3010 - sukces, wymagany restart (MSI)
# Dla winget dodatkowo: 0x8A15002B (brak nowszej wersji) i 0x8A150061 (pakiet już zainstalowany).
$script:InstallerSuccessExitCodes = @(0, 1641, 3010)
$script:WingetAlreadyInstalledExitCodes = @(-1978335189, -1978335135)

function Test-InstallerExitCode {
    param($ExitCode, [switch]$Winget)
    if ($null -eq $ExitCode) { return $false }
    if ($script:InstallerSuccessExitCodes -contains [int]$ExitCode) { return $true }
    if ($Winget -and $script:WingetAlreadyInstalledExitCodes -contains [int]$ExitCode) { return $true }
    return $false
}

# Loguje wynik instalacji na podstawie kodu wyjścia (wcześniej "zainstalowany" było pisane zawsze,
# nawet gdy instalator zwrócił błąd). Zwraca $true przy sukcesie.
function Write-InstallResult {
    param([string]$Name, $ExitCode, [switch]$Winget)
    if (Test-InstallerExitCode -ExitCode $ExitCode -Winget:$Winget) {
        $suffix = ""
        if ([int]$ExitCode -in 1641, 3010) { $suffix = " Wymagany restart komputera." }
        elseif ($Winget -and $script:WingetAlreadyInstalledExitCodes -contains [int]$ExitCode) { $suffix = " Był już zainstalowany wcześniej." }
        Write-Log "$Name zainstalowany.$suffix"
        return $true
    }
    $codeText = if ($null -eq $ExitCode) { "brak - proces nie wystartował lub został przerwany" } else { "$ExitCode" }
    Write-Log "Instalator $Name zakończył się błędem (kod wyjścia: $codeText)." -IsError
    return $false
}

function Test-BeforeRun {
    # CheckboxControls/SelectedApps jako parametry z wartością domyślną = biezący stan globalny -
    # zachowanie identyczne jak wcześniej dla realnego GUI (wywołanie bez argumentów), ale pozwala
    # jednoznacznie wstrzyknąć stan w testach jednostkowych bez polegania na domykaniu zmiennych
    # skryptowych przez PowerShell/Pester (patrz SmartToolforDeployment.Tests.ps1).
    param(
        $CheckboxControls = $CheckboxControls,
        $SelectedApps = $script:SelectedApps
    )
    Write-Log "Rozpoczęto walidację konfiguracji i plików przed wdrożeniem..."
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    try {
        $diskC = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
        if ($null -ne $diskC) {
            $freeGB = [math]::Round($diskC.FreeSpace / 1GB, 2)
            if ($freeGB -lt 15) {
                $warnings.Add("Mało wolnego miejsca na dysku C: ($freeGB GB). Instalacja dużych programów może zakończyć się błędem.") | Out-Null
            }
        }
    } catch {}

    try { $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $config = $null }
    if ($null -eq $config) { $errors.Add("Brak lub niepoprawny config.json.") | Out-Null }

    if ($errors.Count -eq 0) {
        $source = $config.DefaultInstallSource
        $sourcePath = $config.InstallSourcePaths.$source

        switch ($source) {
            'network' {
                if ([string]::IsNullOrWhiteSpace($sourcePath)) { $errors.Add("Brak sciezki sieciowej (InstallSourcePaths.network).") | Out-Null }
                elseif ($sourcePath -notmatch '^\\\\') { $errors.Add("Sciezka sieciowa musi byc UNC (np. \\serwer\\udzial\\).") | Out-Null }
                else {
                    try {
                        if (-not (Test-Path -LiteralPath $sourcePath -ErrorAction Stop)) { $errors.Add("Nie znaleziono zasobu sieciowego: $sourcePath") | Out-Null }
                    } catch {
                        $errors.Add("Błąd dostępu do zasobu sieciowego: $sourcePath`n$($_.Exception.Message)`n(Możliwy problem z poświadczeniami na koncie niedomenowym).") | Out-Null
                    }
                }
            }
            'web' {
                if (-not (Test-UrlValid -Url $sourcePath)) { $errors.Add("Niepoprawny URL zrodla web: $sourcePath") | Out-Null }
            }
            'winget' {
                # Zrodlo winget nie wymaga sciezki sieciowej/url
            }
            default { $errors.Add("DefaultInstallSource musi byc 'network', 'web' lub 'winget'.") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('InstallApplications') -and $CheckboxControls['InstallApplications'].IsChecked -eq $true) {
            if ($SelectedApps.Count -eq 0) {
                $errors.Add("Brak wybranych aplikacji do instalacji.") | Out-Null
            } else {
                foreach ($appName in $SelectedApps.Keys) {
                    $app = $config.Programs.$appName
                    if ($null -eq $app) { $errors.Add("Brak konfiguracji dla aplikacji '$appName'.") | Out-Null; continue }
                    # Własny adres URL programu (DownloadUrl) nadpisuje globalne źródło - wtedy sprawdzamy tylko ten adres.
                    if (-not [string]::IsNullOrWhiteSpace([string]$app.DownloadUrl)) {
                        if (-not (Test-UrlValid -Url ([string]$app.DownloadUrl))) { $errors.Add("Niepoprawny własny URL (DownloadUrl) dla '$appName': $($app.DownloadUrl)") | Out-Null }
                        continue
                    }
                    if ([string]::IsNullOrWhiteSpace($app.FileName)) { $errors.Add("Brak 'FileName' dla '$appName'.") | Out-Null; continue }
                    if ($source -eq 'network') {
                        $full = Join-Path $sourcePath $app.FileName
                        if (-not (Test-Path -LiteralPath $full)) { $errors.Add("Nie znaleziono pliku dla '$appName': ${full}") | Out-Null }
                    } elseif ($source -eq 'web') {
                        $url = if ($sourcePath[-1] -eq '/') { "$sourcePath$($app.FileName)" } else { "$sourcePath/$($app.FileName)" }
                        if (-not (Test-UrlValid -Url $url)) { $errors.Add("Niepoprawny URL dla '$appName': $url") | Out-Null }
                    }
                }
            }
        }

        if ($CheckboxControls.ContainsKey('InstallTeamViewer') -and $CheckboxControls['InstallTeamViewer'].IsChecked -eq $true) {
            if (-not $config.TeamViewer -or [string]::IsNullOrWhiteSpace($config.TeamViewer.FileName)) {
                $errors.Add("TeamViewer nie ma poprawnie ustawionego FileName w config.json.") | Out-Null
            } elseif ($source -eq 'network') {
                $tvPath = Join-Path $sourcePath $config.TeamViewer.FileName
                if (-not (Test-Path -LiteralPath $tvPath)) { $errors.Add("Nie znaleziono instalatora TeamViewer: ${tvPath}") | Out-Null }
            } elseif ($source -eq 'web') {
                $tvUrl = if ($sourcePath[-1] -eq '/') { "$sourcePath$($config.TeamViewer.FileName)" } else { "$sourcePath/$($config.TeamViewer.FileName)" }
                if (-not (Test-UrlValid -Url $tvUrl)) { $errors.Add("Niepoprawny URL TeamViewer: $tvUrl") | Out-Null }
            }
        }

        if ($CheckboxControls.ContainsKey('InstallAV') -and $CheckboxControls['InstallAV'].IsChecked -eq $true) {
            if (-not $config.AntyVirus -or [string]::IsNullOrWhiteSpace($config.AntyVirus.FileName)) { $errors.Add("Brak konfiguracji antywirusa lub FileName.") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('ImportWiFiProfile') -and $CheckboxControls['ImportWiFiProfile'].IsChecked -eq $true) {
            $wifiProfiles = $config.WiFiProfile.FileName
            if ($null -ne $wifiProfiles) {
                foreach ($wifi in @($wifiProfiles)) {
                    if ([string]::IsNullOrWhiteSpace($wifi) -or -not (Test-Path -LiteralPath $wifi)) { $errors.Add("Brak pliku profilu Wi-Fi: $wifi") | Out-Null }
                }
            } else { $errors.Add("Brak konfiguracji WiFiProfile.FileName.") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('CreateLocalAdmin') -and $CheckboxControls['CreateLocalAdmin'].IsChecked -eq $true) {
            if ([string]::IsNullOrWhiteSpace([string]$config.LocalAdmin.Username)) { $errors.Add("Brak LocalAdmin.Username w config.json.") | Out-Null }
        }

        if ($CheckboxControls.ContainsKey('JoinDomain') -and $CheckboxControls['JoinDomain'].IsChecked -eq $true) {
            if (-not $config.DomainJoin -or [string]::IsNullOrWhiteSpace([string]$config.DomainJoin.DomainName)) { $errors.Add("Brak DomainJoin.DomainName w config.json.") | Out-Null }
            if ([string]::IsNullOrWhiteSpace([string]$config.DomainJoin.Username)) { $errors.Add("Brak DomainJoin.Username w config.json.") | Out-Null }
        }
    }

    foreach ($e in $errors) { Write-Log $e -IsError }
    foreach ($w in $warnings) { Write-Log $w }

    if ($errors.Count -gt 0) {
        Show-ThemedMessageBox -Message ("Wykryto bledy walidacji:`r`n- " + ($errors -join "`r`n- ")) -Title "Walidacja" -Button "OK" -Image "Error" | Out-Null
        return $false
    }
    if ($warnings.Count -gt 0) {
        Show-ThemedMessageBox -Message ("Uwaga:`r`n- " + ($warnings -join "`r`n- ")) -Title "Walidacja" -Button "OK" -Image "Warning" | Out-Null
    }
    return $true
}

function Install-SelectedApps {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -IsError
        return
    }

    $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $source = $config.DefaultInstallSource
    $sourcePath = $config.InstallSourcePaths.$source
    $totalApps = $script:SelectedApps.Count
    $currentAppIdx = 0

    foreach ($appName in $script:SelectedApps.Keys) {
        if ($script:isCancelled) { break }
        $currentAppIdx++
        Set-ProgressText "Instalacja aplikacji $currentAppIdx/$($totalApps): $appName..."
        
        $app = $config.Programs.${appName}
        $fileName = [string]$app.FileName
        $silentArgs = [string]$app.SilentArgs
        $downloadUrl = [string]$app.DownloadUrl
        $localPath = $null

        try {
            # Własny adres URL programu (pole "Zawsze pobieraj z niestandardowego adresu URL" w edytorze)
            # ma pierwszeństwo przed globalnym źródłem. Wcześniej był zapisywany, ale nigdy nieużywany.
            $appSource = if (-not [string]::IsNullOrWhiteSpace($downloadUrl)) { 'url' } else { $source }

            if ($appSource -eq 'winget') {
                $cmdArgs = "install --id `"$fileName`" -e --silent --accept-package-agreements --accept-source-agreements $silentArgs"
                if ($script:DryRun) {
                    Write-Log "[DRY-RUN] Zainstalowano by $appName przez Winget: winget $cmdArgs"
                } else {
                    Write-Log "Instalacja $appName (Winget)..."
                    $exitCode = Start-ProcessWithEvents -FilePath "winget.exe" -ArgumentList $cmdArgs
                    Write-InstallResult -Name $appName -ExitCode $exitCode -Winget | Out-Null
                }
            } else {
                $fullPath = if ($appSource -eq 'url') { $downloadUrl } else { Join-InstallSource -BasePath $sourcePath -FileName $fileName }
                # Nazwa pliku lokalnego: z FileName, a gdy jest pusty (sam DownloadUrl) - z końcówki adresu URL.
                $localName = if (-not [string]::IsNullOrWhiteSpace($fileName)) { Split-Path $fileName -Leaf } else { [System.IO.Path]::GetFileName(([uri]$downloadUrl).AbsolutePath) }

                if ($script:DryRun) {
                    # W trybie testowym nic nie pobieramy - wcześniej Dry-Run ściągał wszystkie instalatory.
                    Write-Log "[DRY-RUN] Pobrano by $appName z $fullPath i uruchomiono instalator $localName (argumenty: $silentArgs)."
                } else {
                    $localPath = Join-Path $env:TEMP $localName
                    Write-Log "Pobieranie $appName z $fullPath..."

                    $cred = $null
                    if ($config.WebAuth.Username -and $config.WebAuth.Password) {
                        $sec = ConvertTo-SecureString $config.WebAuth.Password -AsPlainText -Force
                        $cred = [pscredential]::new($config.WebAuth.Username, $sec)
                    }

                    Invoke-DownloadFile -Uri $fullPath -OutFile $localPath -Credential $cred

                    Write-Log "Instalacja $appName..."
                    if ($localName -like "*.msi") {
                        $exitCode = Start-ProcessWithEvents -FilePath "msiexec.exe" -ArgumentList "/i `"$localPath`" $silentArgs"
                    } else {
                        $exitCode = Start-ProcessWithEvents -FilePath $localPath -ArgumentList $silentArgs
                    }
                    Write-InstallResult -Name $appName -ExitCode $exitCode | Out-Null
                }
            }
        }
        catch {
            Write-Log "Błąd przy $($appName): $_" -IsError
        }
        finally {
            if ($source -ne 'winget' -and $null -ne $localPath -and (Test-Path -LiteralPath $localPath)) {
                try { Remove-Item -LiteralPath $localPath -Force -ErrorAction SilentlyContinue } catch {}
                Write-Log "Usunięto plik instalatora: $fileName"
            }
        }
        
        Step-DeploymentProgress
    }
}

function Suspend-Hibernation {
    try {
        Write-Log "Wstrzymywanie usypiania i hibernacji..."
        $code = @"
using System;
using System.Runtime.InteropServices;
public class PowerManagement {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
    public const uint ES_CONTINUOUS = 0x80000000;
    public const uint ES_SYSTEM_REQUIRED = 0x00000001;
    public const uint ES_DISPLAY_REQUIRED = 0x00000002;
}
"@
        if (-not ("PowerManagement" -as [type])) {
            Add-Type -TypeDefinition $code -Language CSharp
        }
        [PowerManagement]::SetThreadExecutionState([PowerManagement]::ES_CONTINUOUS -bor [PowerManagement]::ES_SYSTEM_REQUIRED -bor [PowerManagement]::ES_DISPLAY_REQUIRED) | Out-Null
        Write-Log "Zapobieganie usypianiu włączone na czas działania skryptu."
    }
    catch { Write-Log "Błąd podczas wstrzymywania hibernacji: $_" -IsError }
}

function Resume-Hibernation {
    try {
        if ("PowerManagement" -as [type]) {
            [PowerManagement]::SetThreadExecutionState(0x80000000) | Out-Null
            Write-Log "Przywrócono domyślne zachowanie zasilania."
        }
    } catch {}
}

function Install-TeamViewer {
    # Ścieżka pliku pobranego przez nas do %TEMP%. TYLKO ten plik wolno usunąć w bloku finally.
    # Wcześniej używana była jedna zmienna $localPath, która przy źródle "network" wskazywała
    # instalator na udziale (\\serwer\...\TeamViewer_Host.msi) - i finally kasował go z serwera.
    $downloadedFile = $null
    try {
        if (-not (Test-Path $configPath)) {
            Write-Log "Brak pliku config.json" -IsError
            return
        }

        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $source = $config.DefaultInstallSource
        $sourcePath = $config.InstallSourcePaths.$source
        $msiArgs = $config.TeamViewer.Arguments
        $fileName = $config.TeamViewer.FileName

        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )


        $isInstalled = $false
        $uninstallString = $null
        $isQuietUninstall = $false

        foreach ($path in $regPaths) {
            $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                if ($item.DisplayName -like "*TeamViewer*") {
                    $isInstalled = $true
                    $uninstallString = $item.QuietUninstallString
                    $isQuietUninstall = [bool]$uninstallString
                    if (-not $uninstallString) { $uninstallString = $item.UninstallString }

                    Write-Log "Znaleziono wpis TeamViewer: $($item.DisplayName)"
                    break
                }
            }
            if ($isInstalled) { break }
        }

        if ($script:DryRun) {
            if ($isInstalled) { Write-Log "[DRY-RUN] Zamknięto by procesy TeamViewer i odinstalowano go poleceniem: $uninstallString" }
            Write-Log "[DRY-RUN] Zainstalowano by TeamViewer ze źródła '$source' (plik/ID: $fileName, argumenty: $msiArgs)."
            return
        }

        if ($isInstalled) {
            Write-Log "TeamViewer już zainstalowany, odinstalowuję..."
            if ($uninstallString) {
                $uninstallCommand = $uninstallString
                if ($uninstallString -match '(?i)msiexec') {
                    # Zamieniamy tylko przełącznik instalacji "/I{GUID}" na deinstalację "/X{GUID}".
                    # Stare -replace "/I" podmieniało KAŻDE "/i" (bez rozróżniania wielkości liter).
                    $uninstallCommand = $uninstallString -replace '(?i)/I(?=\{)', '/X'
                    if ($uninstallCommand -notmatch '(?i)/qn') { $uninstallCommand += " /qn /norestart" }
                } elseif (-not $isQuietUninstall) {
                    # Deinstalator TeamViewera w wersji .exe (NSIS) ma tryb cichy "/S" - "/qn" to przełącznik msiexec.
                    $uninstallCommand += " /S"
                }
                Get-Process -Name "*TeamViewer*" -ErrorAction SilentlyContinue | Stop-Process -Force
                Start-Sleep -Seconds 5
                Write-Log "Uruchamiam odinstalowanie: $uninstallCommand"
                $uninstallExit = Start-ProcessWithEvents -FilePath "cmd.exe" -ArgumentList "/c $uninstallCommand"
                if (Test-InstallerExitCode -ExitCode $uninstallExit) {
                    Write-Log "TeamViewer odinstalowany."
                } else {
                    Write-Log "Deinstalacja TeamViewer zwróciła kod $uninstallExit - próbuję mimo to zainstalować ponownie." -IsError
                }
            }
            else {
                Write-Log "Nie znaleziono polecenia odinstalowania TeamViewer." -IsError
                return
            }
        }
        else {
            Write-Log "TeamViewer nie jest zainstalowany, przechodzę do instalacji."
        if ( (Show-ThemedMessageBox -Message "Nie wykryto instalacji TeamViewer. Czy chcesz kontynuować instalację? (Jeśli istnieje proces TeamViewera zostanie on ubity)" -Title "Potwierdzenie" -Button "YesNo" -Image "Question") -ne [System.Windows.MessageBoxResult]::Yes ) {
                Write-Log "Instalacja anulowana przez użytkownika." -IsError
                return
            }
            Get-Process -Name "*TeamViewer*" -ErrorAction SilentlyContinue | Stop-Process -Force
        }

        if ($source -eq 'winget') {
            Write-Log "Instalacja TeamViewer przez Winget..."
            $exitCode = Start-ProcessWithEvents -FilePath "winget.exe" -ArgumentList "install --id `"$fileName`" -e --silent --accept-package-agreements --accept-source-agreements $msiArgs"
            Write-InstallResult -Name "TeamViewer" -ExitCode $exitCode -Winget | Out-Null
        } else {
            $sourceFile = Join-InstallSource -BasePath $sourcePath -FileName $fileName
            if ($sourceFile -match '^https?://') {
                $downloadedFile = Join-Path $env:TEMP (Split-Path $fileName -Leaf)
                Write-Log "Pobieranie TeamViewer z $sourceFile..."
                Invoke-DownloadFile -Uri $sourceFile -OutFile $downloadedFile
                Write-Log "Pobrano TeamViewer do: $downloadedFile"
                $installerPath = $downloadedFile
            }
            else {
                # Instalujemy bezpośrednio z udziału/ścieżki lokalnej - tego pliku NIE usuwamy.
                $installerPath = $sourceFile
                Write-Log "Instalacja TeamViewer ze ścieżki: $installerPath"
            }

            Write-Log "Instalacja TeamViewer..."
            Write-Log "Używam argumentów MSI: $msiArgs"
            $exitCode = Start-ProcessWithEvents -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" $msiArgs"
            Write-InstallResult -Name "TeamViewer" -ExitCode $exitCode | Out-Null
        }
    }
    catch {
        Write-Log "Błąd podczas instalacji TeamViewer: $_" -IsError
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($downloadedFile) -and (Test-Path -LiteralPath $downloadedFile)) {
            Remove-Item -LiteralPath $downloadedFile -Force -ErrorAction SilentlyContinue
            Write-Log "Usunięto pobrany plik instalacyjny TeamViewer: $downloadedFile"
        }
    }
}

function Install-AV {
    try {
        if (-not (Test-Path $configPath)) {
            Write-Log "Brak pliku config.json" -IsError
            return
        }

        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

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
            Write-Log "Brak nazwy pliku instalacyjnego antywirusa w konfiguracji" -IsError
            return
        }

        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Zainstalowano by antywirusa ze źródła '$source' (plik/ID: $fileName, ścieżka: $sourcePath)."
            return
        }

        if ($source -eq 'winget') {
            Write-Log "Instalacja antywirusa przez Winget..."
            $exitCode = Start-ProcessWithEvents -FilePath "winget.exe" -ArgumentList "install --id `"$fileName`" -e --silent --accept-package-agreements --accept-source-agreements"
            Write-InstallResult -Name "Antywirus" -ExitCode $exitCode -Winget | Out-Null
            return
        }

        if ($source -eq "web") {
            # Używamy krótkiej nazwy docelowej, aby uniknąć problemów z długimi nazwami plików (MAX_PATH),
            # ale zachowujemy rozszerzenie oryginału (wcześniej nawet .msi było zapisywane jako .exe).
            $avExtension = [System.IO.Path]::GetExtension($fileName)
            if ([string]::IsNullOrWhiteSpace($avExtension)) { $avExtension = ".exe" }
            $avPath = Join-Path $env:TEMP "setup_av_temp$avExtension"
            $avUrl = Join-InstallSource -BasePath $sourcePath -FileName $fileName
            Write-Log "Pobieranie antywirusa z $avUrl..."

            $cred = $null
            if ($avConfig.Credentials.Username -and $avConfig.Credentials.Password) {
                $sec = ConvertTo-SecureString $avConfig.Credentials.Password -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential ($avConfig.Credentials.Username, $sec)
            }

            Invoke-DownloadFile -Uri $avUrl -OutFile $avPath -Credential $cred
        }
        else {
            $avPath = Join-Path $sourcePath $fileName
        }

        if (-not (Test-Path $avPath)) {
            Write-Log "Plik instalacyjny nie został znaleziony: $avPath" -IsError
            return
        }

        Write-Log "Instalacja antywirusa z $avPath..."
        if ($avPath -like "*.msi") {
            $exitCode = Start-ProcessWithEvents -FilePath "msiexec.exe" -ArgumentList "/i `"$avPath`""
        } else {
            $exitCode = Start-ProcessWithEvents -FilePath $avPath
        }
        Write-InstallResult -Name "Antywirus" -ExitCode $exitCode | Out-Null
    }
    catch {
        Write-Log "Błąd podczas instalacji antywirusa: $_" -IsError
    }
    finally {
        if ($source -eq 'web' -and $null -ne $avPath -and (Test-Path -LiteralPath $avPath)) {
            try { Remove-Item -LiteralPath $avPath -Force -ErrorAction SilentlyContinue } catch {}
            Write-Log "Usunięto plik instalatora antywirusa."
        }
    }
}

function Import-WiFiProfile {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -IsError
        return
    }
    $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $wifiProfiles = $config.WiFiProfile.FileName
    
    if ($null -ne $wifiProfiles) {
        foreach ($wifiProfile in @($wifiProfiles)) {
            if (Test-Path $wifiProfile) {
                Write-Log "Import profilu Wi-Fi z pliku: $wifiProfile..."
                Start-ProcessWithEvents -FilePath "netsh" -ArgumentList "wlan add profile filename=`"$wifiProfile`"" | Out-Null
                Write-Log "Profil Wi-Fi '$wifiProfile' zaimportowany."
            }
            else {
                Write-Log "Brak pliku profilu Wi-Fi: $wifiProfile" -IsError
            }
        }
    } else {
        Write-Log "Brak definicji profili Wi-Fi w konfiguracji." -IsError
    }
}

function New-LocalAdmin {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -IsError
        return
    }
    try {
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $username = $config.LocalAdmin.Username
        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Utworzono by lokalne konto administratora: '$username'."
            return
        }
        $PasswordSecure = Read-Host "Wprowadź hasło dla konta $username" -AsSecureString
        if (-not (Get-LocalUser -Name $username -ErrorAction SilentlyContinue)) {
            if ($null -ne $PasswordSecure -and $PasswordSecure.Length -ge 8) {
                New-LocalUser -Name $username -Password $PasswordSecure -PasswordNeverExpires -AccountNeverExpires
                Add-LocalGroupMember -SID S-1-5-32-544 -Member $username
                Write-Log "Utworzono lokalne konto '$username' w grupie 'Administratorzy'."
            }
            else {
                Write-Log "Nie podano hasła dla konta $username lub hasło jest zbyt krótkie." -IsError
            }
        }
        else {
            Write-Log "Użytkownik '$username' już istnieje  pomijam." -IsError
        }
    }
    catch {
        Write-Log "Błąd tworzenia konta: $_" -IsError
    }
}

function Join-Domain {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -IsError
        return
    }
    try {
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $config.DomainJoin -or -not $config.DomainJoin.DomainName) {
            Write-Log "Brak sekcji DomainJoin w config.json lub DomainName" -IsError
            return
        }

        $DomainName = ($config.DomainJoin.DomainName).Trim()
        $UserForJoin = $config.DomainJoin.Username
        $ComputerName = $env:COMPUTERNAME

        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Dołączono by do domeny '$DomainName' jako '$UserForJoin' (komputer: $ComputerName)."
            return
        }

        Write-Log "Dołączanie do domeny $DomainName jako $UserForJoin..."
        $Credential = Get-Credential -UserName $UserForJoin -Message "Podaj dane domenowe dla $UserForJoin"
        if ($null -eq $Credential) {
            Write-Log "Anulowano podawanie poświadczeń. Pominięto dołączanie do domeny."
            return
        }

        Add-Computer -DomainName $DomainName -Credential $Credential -Force -ErrorAction Stop
        Write-Log "Dołączono do domeny $DomainName z nazwą '$ComputerName'"
        Write-Log "Wymagane jest ponowne uruchomienie komputera, aby zmiany zaczęły obowiązywać."
    }
    catch {
        Write-Log "Błąd dołączania do domeny: $_" -IsError
    }
}

function Set-NewComputerName {
    if ($script:DryRun) {
        $previewName = "PC-$((Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber)"
        Write-Log "[DRY-RUN] Zmieniono by nazwę komputera (domyślnie '$previewName')."
        return
    }
    Read-Host "Podaj nową nazwę komputera (domyślnie 'PC-SERIAL_NUMBER'): " -OutVariable NewName
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        $NewName = "PC-$((Get-CimInstance -ClassName Win32_BIOS).SerialNumber)"
    }
    Rename-Computer -NewName $NewName -Force
    Write-Log "Zmieniono nazwę komputera na '$NewName'."
}

function Set-RegistryDword {
    param(
        [string]$Path,
        [string]$Name,
        [int]$Value
    )
    if ($script:DryRun) {
        Write-Log "[DRY-RUN] Ustawiono by rejestr: $Path\$Name = $Value"
        return
    }
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
        Write-Log "Błąd rejestru dla $Name w $($Path): $_" -IsError
    }
}

# ---------- Punkt przywracania i wznawianie przerwanego wdrożenia ----------
$script:CheckpointFilePath = "C:\deploy-checkpoint.json"

function New-DeploymentRestorePoint {
    if ($script:DryRun) {
        Write-Log "[DRY-RUN] Utworzono by punkt przywracania systemu."
        return
    }
    try {
        Write-Log "Tworzenie punktu przywracania systemu..."
        $svc = Get-Service -Name "swprv" -ErrorAction SilentlyContinue
        if ($null -ne $svc -and $svc.Status -ne 'Running') {
            try { Start-Service -Name "swprv" -ErrorAction Stop } catch {}
        }
        Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
        Checkpoint-Computer -Description "STD-PrzedWdrozeniem-$(Get-Date -Format 'yyyyMMdd-HHmmss')" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        Write-Log "Punkt przywracania systemu utworzony pomyślnie."
    } catch {
        Write-Log "Nie udało się utworzyć punktu przywracania (kontynuuję mimo to): $_" -IsError
    }
}

function Import-DeploymentCheckpoint {
    # Zwraca obiekt informujący, czy istnieje zapis przerwanego wdrożenia, i wczytuje ukończone
    # kroki do $script:CompletedDeploymentSteps, aby Test-StepDone mogło je pomijać przy wznowieniu.
    $script:CompletedDeploymentSteps = @{}
    if (-not (Test-Path -LiteralPath $script:CheckpointFilePath)) {
        return [PSCustomObject]@{ Found = $false; Timestamp = $null; StepCount = 0 }
    }
    try {
        $data = Get-Content -LiteralPath $script:CheckpointFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($step in @($data.CompletedSteps)) { $script:CompletedDeploymentSteps[$step] = $true }
        return [PSCustomObject]@{ Found = $true; Timestamp = $data.Timestamp; StepCount = @($data.CompletedSteps).Count }
    } catch {
        Write-Log "Nie udało się odczytać pliku checkpointu ($script:CheckpointFilePath): $_" -IsError
        return [PSCustomObject]@{ Found = $false; Timestamp = $null; StepCount = 0 }
    }
}

function Test-StepDone {
    param([string]$StepName)
    return ($null -ne $script:CompletedDeploymentSteps -and $script:CompletedDeploymentSteps.ContainsKey($StepName))
}

function Save-DeploymentCheckpointStep {
    param([string]$StepName)
    if ($null -eq $script:CompletedDeploymentSteps) { $script:CompletedDeploymentSteps = @{} }
    $script:CompletedDeploymentSteps[$StepName] = $true
    try {
        $data = [ordered]@{
            Timestamp      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            CompletedSteps = @($script:CompletedDeploymentSteps.Keys)
        }
        $data | ConvertTo-Json | Set-Content -LiteralPath $script:CheckpointFilePath -Encoding UTF8
    } catch {
        Write-Log "Nie udało się zapisać checkpointu wdrożenia: $_" -IsError
    }
}

function Clear-DeploymentCheckpoint {
    $script:CompletedDeploymentSteps = @{}
    if (Test-Path -LiteralPath $script:CheckpointFilePath) {
        try { Remove-Item -LiteralPath $script:CheckpointFilePath -Force -ErrorAction SilentlyContinue } catch {}
    }
}

# ---------- Aktualizacja narzędzia ----------
# Oczekiwany format pliku wersji (std_version.json) publikowanego pod AutoUpdate.VersionCheckPath:
# { "Version": "3.2.0", "FileName": "SmartToolforDeployment.ps1", "Notes": "Opis zmian..." }
# Aktualizuje plik .ps1 wskazywany przez $PSCommandPath. Jeśli w praktyce uruchamiany jest
# skompilowany SmartToolforDeployment_v3.exe, ten mechanizm NIE podmienia tego pliku exe -
# potrzebny byłby analogiczny, osobny mechanizm dla binarki.
function Test-ForAppUpdate {
    param([switch]$Silent)
    try {
        $cfg = Get-Config
        if ($null -eq $cfg.AutoUpdate -or $cfg.AutoUpdate.Enabled -ne $true -or [string]::IsNullOrWhiteSpace([string]$cfg.AutoUpdate.VersionCheckPath)) {
            if (-not $Silent) {
                Show-ThemedMessageBox -Message "Automatyczne aktualizacje nie są skonfigurowane.`nUstaw AutoUpdate.Enabled=true i AutoUpdate.VersionCheckPath w config.json (ścieżka UNC lub URL do katalogu z plikiem std_version.json)." -Title "Aktualizacje" -Button "OK" -Image "Information" | Out-Null
            }
            return
        }

        $basePath = [string]$cfg.AutoUpdate.VersionCheckPath
        $isWeb = $basePath -match '^https?://'
        $infoUri = if ($isWeb) { "$($basePath.TrimEnd('/'))/std_version.json" } else { Join-Path $basePath "std_version.json" }

        Write-Log "Sprawdzanie dostępności aktualizacji ($infoUri)..."
        $raw = if ($isWeb) {
            (Invoke-WebRequest -Uri $infoUri -UseBasicParsing -ErrorAction Stop).Content
        } else {
            Get-Content -LiteralPath $infoUri -Raw -Encoding UTF8 -ErrorAction Stop
        }
        $remote = $raw | ConvertFrom-Json

        $localVer  = [version]$script:ScriptVersion
        $remoteVer = [version]$remote.Version

        if ($remoteVer -le $localVer) {
            Write-Log "Aktualna wersja ($script:ScriptVersion) jest najnowsza."
            if (-not $Silent) { Show-ThemedMessageBox -Message "Masz już najnowszą wersję ($script:ScriptVersion)." -Title "Aktualizacje" -Button "OK" -Image "Information" | Out-Null }
            return
        }

        $msg = "Dostępna jest nowa wersja: $($remote.Version) (obecna: $script:ScriptVersion).`n`n$($remote.Notes)`n`nCzy pobrać i zainstalować aktualizację teraz? Aplikacja zostanie zamknięta i uruchomiona ponownie."
        $ans = Show-ThemedMessageBox -Message $msg -Title "Dostępna aktualizacja $($remote.Version)" -Button "YesNo" -Image "Question"
        if ($ans -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-Log "Użytkownik odrzucił aktualizację do wersji $($remote.Version)." -Context "Użytkownik"
            return
        }

        $remoteFileName = [string]$remote.FileName
        if ([string]::IsNullOrWhiteSpace($remoteFileName)) { throw "Brak nazwy pliku (FileName) w informacji o wersji." }
        $downloadSource = if ($isWeb) { "$($basePath.TrimEnd('/'))/$remoteFileName" } else { Join-Path $basePath $remoteFileName }
        $tempNewFile = Join-Path $env:TEMP $remoteFileName

        Write-Log "Pobieranie wersji $($remote.Version)..."
        if ($isWeb) {
            Invoke-WebRequest -Uri $downloadSource -OutFile $tempNewFile -UseBasicParsing -ErrorAction Stop
        } else {
            Copy-Item -LiteralPath $downloadSource -Destination $tempNewFile -Force -ErrorAction Stop
        }

        $currentFile = $PSCommandPath
        if ([string]::IsNullOrWhiteSpace($currentFile)) { throw "Nie udało się ustalić ścieżki bieżącego pliku (`$PSCommandPath) do podmiany." }

        # Uruchamiamy odrębny, krótkotrwały proces PowerShell, który poczeka aż bieżący proces się
        # zamknie, podmieni plik i uruchomi narzędzie ponownie - nie da się nadpisać pliku, który
        # jest w danym momencie wykonywany przez BIEŻĄCY proces.
        $updaterScript = "Start-Sleep -Seconds 2; Copy-Item -LiteralPath '$tempNewFile' -Destination '$currentFile' -Force; Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File `"$currentFile`"'"
        Start-Process powershell.exe -ArgumentList @("-NoProfile", "-WindowStyle", "Hidden", "-Command", $updaterScript) -WindowStyle Hidden

        Write-Log "Aktualizacja pobrana. Zamykanie aplikacji w celu dokończenia instalacji wersji $($remote.Version)..."
        if ($null -ne $Window) { $Window.Close() }
    } catch {
        Write-Log "Błąd podczas sprawdzania/pobierania aktualizacji: $_" -IsError
        if (-not $Silent) { Show-ThemedMessageBox -Message "Nie udało się sprawdzić/pobrać aktualizacji:`n$_" -Title "Błąd aktualizacji" -Button "OK" -Image "Error" | Out-Null }
    }
}

function Get-HardwareAudit {
    param([switch]$AsHtml, [switch]$AsObject)
    $os = Get-CimInstance Win32_OperatingSystem
    $bios = Get-CimInstance Win32_BIOS
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $ram = Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
    $ramGb = if ($ram.Sum) { [math]::Round($ram.Sum / 1GB, 2) } else { 0 }
    $disks = Get-CimInstance Win32_DiskDrive | Where-Object { $_.MediaType -eq "Fixed hard disk media" }
    $nets = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

    $installedApps = $null
    $appError = $null
    try {
        $paths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        $installedAppsList = New-Object System.Collections.Generic.List[PSCustomObject]
        foreach ($basePath in $paths) {
            if (Test-Path -LiteralPath $basePath) {
                foreach ($key in (Get-ChildItem -LiteralPath $basePath -ErrorAction SilentlyContinue)) {
                    $displayName = $key.GetValue("DisplayName")
                    $systemComponent = $key.GetValue("SystemComponent")
                    $parentKeyName = $key.GetValue("ParentKeyName")
                    if ($displayName -and (-not $systemComponent) -and ($null -eq $parentKeyName) -and ($displayName -notmatch '^KB\d+')) {
                        $installedAppsList.Add([PSCustomObject]@{ DisplayName = $displayName; DisplayVersion = $key.GetValue("DisplayVersion") })
                    }
                }
            }
        }
        $installedApps = $installedAppsList | Sort-Object DisplayName -Unique
    } catch {
        $appError = $_.Exception.Message
    }

    if ($AsObject) {
        return [PSCustomObject]@{
            ComputerName  = $env:COMPUTERNAME
            UserName      = $env:USERNAME
            GeneratedAt   = Get-Date
            OsSummary     = "$($os.Caption) $($os.OSArchitecture) (Build: $($os.BuildNumber))"
            BiosSerial    = $bios.SerialNumber
            BiosVersion   = $bios.SMBIOSBIOSVersion
            CpuName       = $cpu.Name
            RamGb         = $ramGb
            Disks         = @($disks | ForEach-Object { [PSCustomObject]@{ Model = $_.Model; SizeGb = [math]::Round($_.Size / 1GB, 2); SerialNumber = ([string]$_.SerialNumber).Trim() } })
            Networks      = @($nets | ForEach-Object { [PSCustomObject]@{ Description = $_.Description; Mac = $_.MACAddress; Ip = ($_.IPAddress -join ', ') } })
            InstalledApps = @($installedApps | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; DisplayVersion = $_.DisplayVersion } })
            AppError      = $appError
        }
    }

    if ($AsHtml) {
        $html = @"
<!DOCTYPE html>
<html lang='pl'>
<head>
    <meta charset='utf-8'>
    <title>Raport Systemowy: $($env:COMPUTERNAME)</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f9; color: #333; padding: 20px; }
        h1 { color: #0078D7; border-bottom: 2px solid #0078D7; padding-bottom: 5px; }
        h2 { color: #0078D7; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 0; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; background-color: white; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #0078D7; color: white; position: sticky; top: 0; }
        tr:hover { background-color: #f1f1f1; }
        .meta { font-size: 0.9em; color: #666; margin-bottom: 20px; }
        .toc { background: white; padding: 15px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.2); display: inline-block; margin-bottom: 20px; min-width: 250px; }
        .toc a { text-decoration: none; color: #0078D7; font-weight: bold; display: block; margin: 5px 0; }
        .toc a:hover { text-decoration: underline; }
        #search { display: block; padding: 10px; width: 100%; max-width: 400px; margin-bottom: 20px; border: 1px solid #ccc; border-radius: 4px; font-size: 15px; }
        .section { margin-bottom: 30px; background: white; padding: 20px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
        ul.list { list-style-type: none; padding: 0; margin: 0; }
        ul.list li { padding: 8px; border-bottom: 1px solid #eee; }
        ul.list li:hover { background-color: #f9f9f9; }
    </style>
    <script>
        function searchReport() {
            let input = document.getElementById('search').value.toLowerCase();
            let sections = document.querySelectorAll('.section');
            sections.forEach(sec => {
                let items = sec.querySelectorAll('tr:not(.header-row), ul.list li');
                items.forEach(item => {
                    if (item.textContent.toLowerCase().includes(input)) {
                        item.style.display = '';
                    } else {
                        item.style.display = 'none';
                    }
                });
            });
        }
    </script>
</head>
<body>
    <h1>Raport Systemowy: $($env:COMPUTERNAME)</h1>
    <div class='meta'>Wygenerowano: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>Użytkownik: $($env:USERNAME)</div>
    
    <input type="text" id="search" onkeyup="searchReport()" placeholder="🔍 Szukaj w raporcie (sprzęt, aplikacje)...">
    
    <div class="toc">
        <h2 style="border:none; margin-bottom:10px;">Spis treści</h2>
        <a href="#os">💻 System Operacyjny i BIOS</a>
        <a href="#cpu">⚙️ Procesor i RAM</a>
        <a href="#disks">💾 Dyski twarde</a>
        <a href="#network">🌐 Karty sieciowe</a>
        <a href="#apps">📦 Zainstalowane oprogramowanie</a>
    </div>

    <div class="section" id="os">
        <h2>💻 System Operacyjny i BIOS</h2>
        <ul class="list">
            <li><strong>OS:</strong> $($os.Caption) $($os.OSArchitecture) (Build: $($os.BuildNumber))</li>
            <li><strong>BIOS SN:</strong> $($bios.SerialNumber)</li>
            <li><strong>BIOS Wersja:</strong> $($bios.SMBIOSBIOSVersion)</li>
        </ul>
    </div>

    <div class="section" id="cpu">
        <h2>⚙️ Procesor i RAM</h2>
        <ul class="list">
            <li><strong>Procesor:</strong> $($cpu.Name)</li>
            <li><strong>Pamięć RAM:</strong> $ramGb GB</li>
        </ul>
    </div>

    <div class="section" id="disks">
        <h2>💾 Dyski twarde</h2>
        <ul class="list">
"@
        if ($disks) {
            foreach ($d in $disks) {
                $dSize = [math]::Round($d.Size / 1GB, 2)
                $html += "            <li><strong>$($d.Model)</strong> - $dSize GB</li>`r`n"
            }
        } else {
            $html += "            <li>Brak danych</li>`r`n"
        }
        $html += @"
        </ul>
    </div>

    <div class="section" id="network">
        <h2>🌐 Karty sieciowe</h2>
        <table>
            <tr class="header-row"><th>Opis</th><th>MAC Adres</th><th>Adresy IP</th></tr>
"@
        if ($nets) {
            foreach ($n in $nets) {
                $ip = $n.IPAddress -join ', '
                $html += "            <tr><td>$($n.Description)</td><td>$($n.MACAddress)</td><td>$ip</td></tr>`r`n"
            }
        } else {
            $html += "            <tr><td colspan='3'>Brak aktywnych kart sieciowych</td></tr>`r`n"
        }
        $html += @"
        </table>
    </div>

    <div class="section" id="apps">
        <h2>📦 Zainstalowane oprogramowanie</h2>
        <table>
            <tr class="header-row"><th>Nazwa programu</th><th>Wersja</th></tr>
"@
        if ($installedApps) {
            foreach ($app in $installedApps) {
                $name = [System.Security.SecurityElement]::Escape([string]$app.DisplayName)
                $ver = [System.Security.SecurityElement]::Escape([string]$app.DisplayVersion)
                $html += "            <tr><td>$name</td><td>$ver</td></tr>`r`n"
            }
        } else {
            $html += "            <tr><td colspan='2'>Brak zainstalowanych programów.</td></tr>`r`n"
        }
        $html += @"
        </table>
    </div>
</body>
</html>
"@
        return $html
    }

    $info = "================ AUDYT SPRZĘTOWY ================`r`n"
    $info += "Nazwa komputera: $($env:COMPUTERNAME)`r`n"
    $info += "Użytkownik: $($env:USERNAME)`r`n"
    $info += "Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n"
    $info += "-------------------------------------------------`r`n"
    $info += "OS: $($os.Caption) $($os.OSArchitecture) (Build: $($os.BuildNumber))`r`n"
    $info += "BIOS / SN: $($bios.SerialNumber) (Wersja: $($bios.SMBIOSBIOSVersion))`r`n"
    $info += "Procesor: $($cpu.Name)`r`n"
    $info += "Pamięć RAM: $ramGb GB`r`n"
    $info += "Dyski:`r`n"
    if ($disks) {
        foreach ($d in $disks) {
            $dSize = [math]::Round($d.Size / 1GB, 2)
            $info += " - $($d.Model) ($dSize GB)`r`n"
        }
    }
    $info += "Karty sieciowe:`r`n"
    if ($nets) {
        foreach ($n in $nets) {
            $ip = $n.IPAddress -join ', '
            $info += " - $($n.Description)`r`n    MAC: $($n.MACAddress)`r`n    IP: $ip`r`n"
        }
    }
    $info += "=================================================`r`n"
    $info += "`r`n============= ZAINSTALOWANE OPROGRAMOWANIE =============`r`n"
    if ($appError) {
        $info += " Błąd pobierania listy oprogramowania: $appError`r`n"
    } elseif ($installedApps) {
            foreach ($app in $installedApps) {
                $version = if ($app.DisplayVersion) { " (v. $($app.DisplayVersion))" } else { "" }
                $info += " - $($app.DisplayName)$version`r`n"
            }
        } else {
            $info += " Brak zainstalowanych programów.`r`n"
        }
    $info += "=========================================================="
    return $info
}

function Invoke-PostInstallScripts {
    try {
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $config.PostInstallScripts -or $config.PostInstallScripts.Count -eq 0) {
            Write-Log "Brak zdefiniowanych skryptów poinstalacyjnych."
            return
        }

        Write-Log "Rozpoczynam uruchamianie skryptów poinstalacyjnych..."
        $source = $config.DefaultInstallSource
        $sourcePath = $config.InstallSourcePaths.$source

        # Zmienna pętli nazywała się wcześniej $script - to poprawna, ale myląca nazwa, bo łatwo ją
        # pomylić z zakresem $script: (np. $script:isCancelled w tej samej pętli).
        foreach ($scriptEntry in $config.PostInstallScripts) {
            if ($script:isCancelled) { break }
            Write-Log "Wykonywanie skryptu: $scriptEntry..."
            $scriptToRun = $scriptEntry
            $localPath = ""
            $isRemote = ($source -eq 'web' -or $scriptEntry -match "^https?://")
            $scriptUrl = if ($scriptEntry -match "^https?://") { $scriptEntry } elseif ($isRemote) { Join-InstallSource -BasePath $sourcePath -FileName $scriptEntry } else { $null }
            if (-not $isRemote -and $source -eq 'network') { $scriptToRun = Join-Path $sourcePath $scriptEntry }

            if ($script:DryRun) {
                $what = if ($isRemote) { "pobrano by $scriptUrl i uruchomiono" } else { "uruchomiono by $scriptToRun" }
                Write-Log "[DRY-RUN] Skrypt poinstalacyjny: $what."
                continue
            }

            try {
                if ($isRemote) {
                    $localPath = Join-Path $env:TEMP (Split-Path $scriptUrl -Leaf)
                    Invoke-DownloadFile -Uri $scriptUrl -OutFile $localPath
                    $scriptToRun = $localPath
                }

                if (Test-Path $scriptToRun) {
                    if ($scriptToRun -match "\.ps1$") {
                        $exitCode = Start-ProcessWithEvents -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptToRun`""
                    } elseif ($scriptToRun -match "\.(bat|cmd)$") {
                        $exitCode = Start-ProcessWithEvents -FilePath "cmd.exe" -ArgumentList "/c `"$scriptToRun`""
                    } else {
                        $exitCode = Start-ProcessWithEvents -FilePath $scriptToRun
                    }
                    if ($exitCode -eq 0) {
                        Write-Log "Zakończono wykonywanie skryptu: $scriptEntry"
                    } else {
                        Write-Log "Skrypt $scriptEntry zakończył się kodem $exitCode." -IsError
                    }
                } else {
                    Write-Log "Nie znaleziono pliku skryptu: $scriptToRun" -IsError
                }
            } catch {
                # Błąd jednego skryptu (np. nieudane pobranie) nie przerywa kolejnych.
                Write-Log "Błąd skryptu poinstalacyjnego $($scriptEntry): $_" -IsError
            } finally {
                if ($localPath -and (Test-Path $localPath)) { Remove-Item $localPath -Force -ErrorAction SilentlyContinue }
            }
        }
    } catch {
        Write-Log "Błąd podczas wykonywania skryptów poinstalacyjnych: $_" -IsError
    }
}

function Export-HardwareAuditTask {
    try {
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $exportDir = "C:\Audit\"
        if ($null -ne $config.HardwareAudit -and -not [string]::IsNullOrWhiteSpace($config.HardwareAudit.ExportPath)) { 
            $exportDir = [System.Environment]::ExpandEnvironmentVariables($config.HardwareAudit.ExportPath)
        }
        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Wyeksportowano by audyt sprzętowy do katalogu: $exportDir"
            return
        }
        if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
        $auditData = Get-HardwareAudit -AsHtml
        $fileName = "Audit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        $fullPath = Join-Path $exportDir $fileName
        $auditData | Out-File -FilePath $fullPath -Encoding UTF8 -Force
        Write-Log "Audyt sprzętowy wyeksportowano do: $fullPath"
    } catch { Write-Log "Błąd podczas eksportu audytu sprzętowego: $_" -IsError }
}

function Enable-BitLockerEncryption {
    try {
        Write-Log "Sprawdzanie statusu modułu TPM i BitLockera..."
        $tpm = Get-Tpm -ErrorAction SilentlyContinue
        if (-not $tpm -or -not $tpm.TpmPresent -or -not $tpm.TpmReady) {
            Write-Log "Moduł TPM nie jest dostępny lub gotowy! Pominięto szyfrowanie." -IsError
            return
        }
        $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($bl.VolumeStatus -eq "FullyEncrypted" -or $bl.VolumeStatus -eq "EncryptionInProgress") {
            Write-Log "Dysk C: jest już zaszyfrowany lub proces jest w toku."
            return
        }
        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Zaszyfrowano by dysk C: (BitLocker XTS-AES 256) i wyeksportowano klucz odzyskiwania."
            return
        }
        Write-Log "Generowanie klucza odzyskiwania..."
        Add-BitLockerKeyProtector -MountPoint "C:" -TpmProtector -ErrorAction Stop | Out-Null
        $recovery = Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector -ErrorAction Stop
        $recPassword = ($recovery.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }).RecoveryPassword
        Write-Log "Rozpoczynanie szyfrowania dysku C: (XTS-AES 256)..."
        Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAes256 -UsedSpaceOnly -SkipHardwareTest -ErrorAction Stop | Out-Null
        Write-Log "Szyfrowanie zostało zainicjowane pomyślnie."
        
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $exportDir = "C:\Audit\"
        if ($null -ne $config.HardwareAudit -and -not [string]::IsNullOrWhiteSpace($config.HardwareAudit.ExportPath)) { 
            $exportDir = [System.Environment]::ExpandEnvironmentVariables($config.HardwareAudit.ExportPath)
        }
        if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
        $fileName = "BitLocker_Recovery_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $fullPath = Join-Path $exportDir $fileName
        $info = "Komputer: $($env:COMPUTERNAME)`r`nData: $(Get-Date)`r`nKlucz odzyskiwania: $recPassword"
        $info | Out-File -FilePath $fullPath -Encoding UTF8 -Force
        Write-Log "Klucz odzyskiwania zapisano w: $fullPath"
    } catch { Write-Log "Wystąpił błąd podczas aktywacji BitLockera: $_" -IsError }
}

# Tworzy klucz rejestru tylko, gdy jeszcze nie istnieje. "New-Item -Force" na ISTNIEJĄCYM kluczu
# rejestru tworzy go od nowa i kasuje wszystkie jego wartości (np. inne polityki ustawione przez
# GPO w tym samym kluczu), dlatego najpierw sprawdzamy Test-Path. W trybie Dry-Run nic nie robi -
# samą zmianę zaloguje Set-RegistryDword.
function New-RegistryKeyIfMissing {
    param([string]$Path)
    if ($script:DryRun) { return }
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
}

function Set-SystemTweaks {
    try {
        if (-not (Test-Path $configPath)) {
            Write-Log "Brak pliku config.json" -IsError
            return
        }

        $settings = (Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json).SystemSettings
        Write-Log "Zastosowanie ustawień systemowych..."

        if ($settings.DisableDeliveryOptimization) {
            Write-Log "Wyłączanie Delivery Optimization..."
            New-RegistryKeyIfMissing -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization"
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" -Name "DODownloadMode" -Value 0
        }

        if ($settings.EnableWin10StartMenu) {
            Write-Log "Włączanie klasycznego menu Start (Win10)..."
            Set-RegistryDword -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_ShowClassicMode" -Value 1
        }

        if ($settings.DisableTelemetry) {
            Write-Log "Wyłączanie telemetryki..."
            New-RegistryKeyIfMissing -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0
        }

        if ($settings.DisableCortana) {
            Write-Log "Wyłączanie Cortany..."
            New-RegistryKeyIfMissing -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0
        }

        if ($settings.DisableFastStartup) {
            Write-Log "Wyłączanie szybkiego uruchamiania..."
            Set-RegistryDword -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HiberbootEnabled" -Value 0
        }

        if ($settings.DisableNewsAndInterests) {
            Write-Log "Wyłączanie News and Interests..."
            New-RegistryKeyIfMissing -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
            Set-RegistryDword -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Name "AllowNewsAndInterests" -Value 0
        }

        Write-Log "Zmiany systemowe zastosowane."

        if ($null -ne $settings.CustomRegistry) {
            Write-Log "Wprowadzanie niestandardowych kluczy rejestru..."
            foreach ($reg in $settings.CustomRegistry) {
                if ($script:isCancelled) { break }
                Do-WpfEvents
                try {
                    if ([string]::IsNullOrWhiteSpace($reg.Path) -or [string]::IsNullOrWhiteSpace($reg.Name)) {
                        Write-Log "Pominięto wpis rejestru (brak Path lub Name)."
                        continue
                    }
                    $propType = if ([string]::IsNullOrWhiteSpace($reg.PropertyType)) { "String" } else { $reg.PropertyType }
                    if ($script:DryRun) {
                        Write-Log "[DRY-RUN] Ustawiono by klucz: $($reg.Path)\$($reg.Name) = $($reg.Value) [$propType]"
                        continue
                    }
                    if (-not (Test-Path $reg.Path)) {
                        New-Item -Path $reg.Path -Force | Out-Null
                        Write-Log "Utworzono nową ścieżkę: $($reg.Path)"
                    }
                    Set-ItemProperty -Path $reg.Path -Name $reg.Name -Value $reg.Value -Type $propType -Force
                    Write-Log "Ustawiono klucz: $($reg.Path)\$($reg.Name) = $($reg.Value) [$propType]"
                }
                catch {
                    Write-Log "Błąd przy ustawianiu klucza $($reg.Path)\$($reg.Name): $_" -IsError
                }
            }
        }
    }
    catch {
        Write-Log "Błąd w Set-SystemTweaks: $_" -IsError
    }
}

function Start-WindowsUpdate {
    try {
        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Otwarto by panel Windows Update."
            return
        }
        Write-Log "Uruchamianie Windows Update..."
        Start-Process "control.exe" -ArgumentList "/name Microsoft.WindowsUpdate"
    }
    catch {
        Write-Log "Nie udało się uruchomić WU: $_" -IsError
    }
}

function Uninstall-Microsoft365Apps {
    try {
        Write-Log "Rozpoczynam dezinstalację Microsoft 365, pakietów Office i OneNote..."
        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        $OfficeUninstallStrings = (Get-ItemProperty $regPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match "(?i)Microsoft 365|Microsoft Office|OneNote" } | Select-Object -ExpandProperty UninstallString)
        if ($script:DryRun) {
            foreach ($UninstallString in @($OfficeUninstallStrings)) { Write-Log "[DRY-RUN] Uruchomiono by deinstalator: $UninstallString DisplayLevel=False" }
            Write-Log "[DRY-RUN] Usunięto by pakiety Appx: MicrosoftOfficeHub, OneNote, Microsoft.Office.Desktop."
            return
        }
        if ($OfficeUninstallStrings) {
            ForEach ($UninstallString in $OfficeUninstallStrings) {
                if ($script:isCancelled) { break }
                Do-WpfEvents
                if ($UninstallString -match '^"(.*?)"\s*(.*)$') {
                    $UninstallEXE = $matches[1]
                    $UninstallArg = $matches[2] + " DisplayLevel=False"
                } elseif ($UninstallString -match '^([^"]*?\.exe)\s*(.*)$') {
                    $UninstallEXE = $matches[1]
                    $UninstallArg = $matches[2] + " DisplayLevel=False"
                } else {
                    $UninstallEXE = ($UninstallString -split '"')[1]
                    $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
                }
                Write-Log "Wykonywanie: $UninstallEXE $UninstallArg"
                Start-ProcessWithEvents -FilePath $UninstallEXE -ArgumentList $UninstallArg | Out-Null
            }
            Write-Log "Odinstalowano aplikacje Microsoft 365 / Office (desktop)."
        } else {
            Write-Log "Nie znaleziono preinstalowanego pakietu Microsoft 365/Office (desktop)."
        }
        
        Write-Log "Usuwanie aplikacji UWP (Appx) dla Office/OneNote..."
        $uwpApps = @("*MicrosoftOfficeHub*", "*OneNote*", "*Microsoft.Office.Desktop*")
        foreach ($app in $uwpApps) {
            if ($script:isCancelled) { break }
            Do-WpfEvents
            Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
            Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like $app } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
        }
        Write-Log "Zakończono usuwanie Microsoft 365/Office."
    } catch {
        Write-Log "Błąd deinstalacji M365: $_" -IsError
    }
}

function Uninstall-OneDrive {
    Write-Log "Rozpoczynam odinstalowywanie OneDrive..."
    if ($script:DryRun) {
        Write-Log "[DRY-RUN] Zamknięto by proces OneDrive i uruchomiono OneDriveSetup.exe /uninstall."
        return
    }
    try {
        Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        $oneDriveSetup64 = Join-Path $env:SystemRoot "SysWOW64\OneDriveSetup.exe"
        $oneDriveSetup32 = Join-Path $env:SystemRoot "System32\OneDriveSetup.exe"

        if (Test-Path $oneDriveSetup64) {
            Start-ProcessWithEvents -FilePath $oneDriveSetup64 -ArgumentList "/uninstall" | Out-Null
        } elseif (Test-Path $oneDriveSetup32) {
            Start-ProcessWithEvents -FilePath $oneDriveSetup32 -ArgumentList "/uninstall" | Out-Null
        } else {
            Write-Log "Nie znaleziono instalatora OneDriveSetup.exe. Być może został już odinstalowany."
        }
        Write-Log "Zakończono deinstalację OneDrive."
    } catch {
        Write-Log "Błąd podczas usuwania OneDrive: $_" -IsError
    }
}

function Join-Intune {
    Write-Log "Funkcja dołączania do Intune nie została jeszcze zaimplementowana."
}

function Show-AppSelectionWindow {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -Color "Red"
        return
    }

    $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $programs = $config.Programs
    
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Wybór aplikacji" Width="550" Height="650" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="Margin" Value="5,6"/>
            <Setter Property="Padding" Value="10,0,0,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Foreground="{TemplateBinding Foreground}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                            <TextBlock Text="▼" Foreground="{TemplateBinding Foreground}" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="10"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="8,0,25,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="250" CornerRadius="4">
                                    <ScrollViewer SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True"><Setter Property="Background" Value="{DynamicResource ThemeButton}"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Grid Grid.Row="0" Margin="0,0,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="🔍 Szukaj:" VerticalAlignment="Center" FontSize="15" Margin="0,0,10,0"/>
            <TextBox Name="txtSearch" Grid.Column="1" Height="32"/>
        </Grid>
        
        <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="8" Padding="15" Grid.Row="1">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="spApps"/>
            </ScrollViewer>
        </Border>
        
        <Grid Grid.Row="2" Margin="0,15,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Name="btnSelectAll" Content="Zaznacz wszystko" Grid.Column="0" Margin="0,0,5,0" Height="36"/>
            <Button Name="btnDeselectAll" Content="Odznacz wszystko" Grid.Column="1" Margin="5,0,5,0" Height="36"/>
            <Button Name="btnInvertSelection" Content="Odwróć zaznaczenie" Grid.Column="2" Margin="5,0,0,0" Height="36"/>
        </Grid>
        
        <Grid Grid.Row="3">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Name="btnOk" Content="OK" Margin="0,0,5,0" Height="40" Background="#FF0E639C" Foreground="White" FontWeight="SemiBold" IsDefault="True"/>
            <Button Name="btnCancel" Content="Anuluj" Grid.Column="1" Margin="5,0,0,0" Height="40" IsCancel="True" Foreground="White" Background="#FFD83B01"/>
        </Grid>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $popup = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $popup

    $spApps = $popup.FindName("spApps")
    $btnSelectAll = $popup.FindName("btnSelectAll")
    $btnDeselectAll = $popup.FindName("btnDeselectAll")
    $btnInvertSelection = $popup.FindName("btnInvertSelection")
    $btnOk = $popup.FindName("btnOk")
    $btnCancel = $popup.FindName("btnCancel")
    $txtSearch = $popup.FindName("txtSearch")

    $checkboxes = @{}

    foreach ($name in $programs.PSObject.Properties.Name | Sort-Object) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $name
        
        # Pamiętaj wybór w trakcie sesji za pomocą $script:SelectedApps
        if ($script:SelectedApps.ContainsKey($name)) {
            $cb.IsChecked = $true
        } else {
            $cb.IsChecked = $false
        }

        $spApps.Children.Add($cb) | Out-Null
        $checkboxes[$name] = $cb
    }

    $txtSearch.Add_TextChanged({
        $filter = $txtSearch.Text.ToLower()
        foreach ($cb in $checkboxes.Values) {
            if ($cb.Content.ToLower().Contains($filter)) {
                $cb.Visibility = [System.Windows.Visibility]::Visible
            } else {
                $cb.Visibility = [System.Windows.Visibility]::Collapsed
            }
        }
    })

    $btnSelectAll.Add_Click({
        foreach ($cb in $checkboxes.Values) { 
            if ($cb.Visibility -eq [System.Windows.Visibility]::Visible) { $cb.IsChecked = $true }
        }
    })

    $btnDeselectAll.Add_Click({
        foreach ($cb in $checkboxes.Values) { 
            if ($cb.Visibility -eq [System.Windows.Visibility]::Visible) { $cb.IsChecked = $false }
        }
    })

    $btnInvertSelection.Add_Click({
        foreach ($cb in $checkboxes.Values) { 
            if ($cb.Visibility -eq [System.Windows.Visibility]::Visible) { $cb.IsChecked = -not $cb.IsChecked }
        }
    })

    $btnOk.Add_Click({
        $script:SelectedApps.Clear()
        foreach ($key in $checkboxes.Keys) {
            if ($checkboxes[$key].IsChecked) {
                $script:SelectedApps[$key] = $true
            }
        }
        $popup.DialogResult = $true
        $popup.Close()
        
        $script:ignoreProfileChange = $true
        if ($null -ne $cmbProfiles -and $cmbProfiles.SelectedIndex -ne 0) {
            $cmbProfiles.SelectedIndex = 0
        }
        $script:ignoreProfileChange = $false
        
        if ($null -ne $btnChooseApps) {
            $btnChooseApps.Content = "Wybierz aplikacje ($($script:SelectedApps.Count))"
            $checkboxControls["InstallApplications"].IsChecked = $script:SelectedApps.Count -gt 0
        }
        
        if ($script:SelectedApps.Count -eq 0) {
            $CheckboxControls["InstallApplications"].IsChecked = $false
            Write-Log "Nie wybrano żadnych aplikacji do instalacji."
        }
        else {
            Write-Log "Wybrano aplikacje: $($script:SelectedApps.Keys -join ', ')"
        }
    })

    $btnCancel.Add_Click({
        $popup.DialogResult = $false
        $popup.Close()
    })

    $popup.ShowDialog() | Out-Null
}

function Start-Deployment {
    Write-Log "Użytkownik rozpoczął wdrożenie (przycisk 'ROZPOCZNIJ KONFIGURACJĘ')." -Context "Użytkownik"
    $btnStart.IsEnabled = $false
    $btnPause.IsEnabled = $true
    $btnCancelDeploy.IsEnabled = $true
    $btnStart.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFD83B01")
    $progressBar.Value = 0
    $script:isCancelled = $false
    $script:isPaused = $false
    $script:DryRun = ($CheckboxControls.ContainsKey("DryRun") -and $CheckboxControls["DryRun"].IsChecked -eq $true)
    if ($script:DryRun) { Write-Log "=== TRYB TESTOWY (DRY-RUN) WŁĄCZONY: żadne rzeczywiste zmiany nie zostaną wprowadzone ===" }

    if (-not (Test-BeforeRun)) {
        Write-Log "Walidacja nie powiodla sie. Przerywam." -IsError
        $btnStart.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFDC143C")
        $btnStart.IsEnabled = $true
        $btnPause.IsEnabled = $false
        $btnCancelDeploy.IsEnabled = $false
        return
    }

    # Checkpoint z poprzedniego, niedokończonego uruchomienia - służy tylko jako ślad diagnostyczny
    # (co zdążyło się wykonać przed przerwaniem); każde nowe wdrożenie zawsze startuje od zera.
    $resumeInfo = Import-DeploymentCheckpoint
    if ($resumeInfo.Found) {
        Write-Log "Wykryto ślad poprzedniego, niedokończonego wdrożenia z $($resumeInfo.Timestamp) ($($resumeInfo.StepCount) ukończonych kroków). Rozpoczynam nowe wdrożenie od początku."
    }
    Clear-DeploymentCheckpoint

    if ($CheckboxControls.ContainsKey("CreateRestorePoint") -and $CheckboxControls["CreateRestorePoint"].IsChecked -eq $true) {
        Set-ProgressText "Tworzenie punktu przywracania systemu..."
        New-DeploymentRestorePoint
    }

    $txtStopwatch.Visibility = [System.Windows.Visibility]::Visible
    $script:stopwatchStartTime = Get-Date
    $script:stopwatchAccumulated = [TimeSpan]::Zero
    $txtStopwatch.Text = "⏱ 00:00:00"
    $script:stopwatchTimer.Start()

    # Od tego miejsca wpisy Write-Log bez jawnego -Context oznaczane są jako "Automat" (kroki
    # wykonywane samodzielnie w ramach sekwencji wdrożenia), aż do zakończenia w bloku finally.
    $script:CurrentLogContext = "Automat"
    try {
        Set-ProgressText "Przygotowywanie do wdrożenia..."
        
        $script:TotalDeploymentSteps = 1
        if ($CheckboxControls.ContainsKey("WaitForNetwork") -and $CheckboxControls["WaitForNetwork"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls.ContainsKey("SuspendHibernation") -and $CheckboxControls["SuspendHibernation"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls["InstallTeamViewer"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls["UninstallMicrosoft365"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls.ContainsKey("UninstallOneDrive") -and $CheckboxControls["UninstallOneDrive"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls.ContainsKey("RemoveBloatware") -and $CheckboxControls["RemoveBloatware"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls["InstallAV"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls["ImportWiFiProfile"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls["CreateLocalAdmin"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls["JoinDomain"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls["InstallApplications"].IsChecked -eq $true -and $script:SelectedApps.Count -gt 0) { $script:TotalDeploymentSteps += $script:SelectedApps.Count }
        if ($CheckboxControls.ContainsKey("ChangeComputerName") -and $CheckboxControls["ChangeComputerName"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls.ContainsKey("JoinIntune") -and $CheckboxControls["JoinIntune"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls.ContainsKey("RunPostInstallScripts") -and $CheckboxControls["RunPostInstallScripts"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls.ContainsKey("ExportHardwareAudit") -and $CheckboxControls["ExportHardwareAudit"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
        if ($CheckboxControls.ContainsKey("EnableBitLocker") -and $CheckboxControls["EnableBitLocker"].IsChecked -eq $true) { $script:TotalDeploymentSteps++ }
    
        $script:CurrentDeploymentStep = 0
    
        if ($CheckboxControls.ContainsKey("WaitForNetwork") -and $CheckboxControls["WaitForNetwork"].IsChecked -eq $true) { Set-ProgressText "Oczekiwanie na sieć..."; Wait-ForNetwork; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "WaitForNetwork" }

        if ($CheckboxControls.ContainsKey("SuspendHibernation") -and $CheckboxControls["SuspendHibernation"].IsChecked -eq $true) { Set-ProgressText "Wstrzymywanie hibernacji..."; Suspend-Hibernation; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "SuspendHibernation" }

        if ($CheckboxControls["InstallTeamViewer"].IsChecked -eq $true) { Set-ProgressText "Instalacja: TeamViewer..."; Install-TeamViewer; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "InstallTeamViewer" }

        if ($CheckboxControls["UninstallMicrosoft365"].IsChecked -eq $true) { Set-ProgressText "Deinstalacja: Microsoft 365..."; Uninstall-Microsoft365Apps; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "UninstallMicrosoft365" }

        if ($CheckboxControls.ContainsKey("UninstallOneDrive") -and $CheckboxControls["UninstallOneDrive"].IsChecked -eq $true) { Set-ProgressText "Deinstalacja: OneDrive..."; Uninstall-OneDrive; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "UninstallOneDrive" }

        if ($CheckboxControls.ContainsKey("RemoveBloatware") -and $CheckboxControls["RemoveBloatware"].IsChecked -eq $true) { Set-ProgressText "Usuwanie Bloatware..."; Remove-Bloatware; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "RemoveBloatware" }

        if ($CheckboxControls["InstallAV"].IsChecked -eq $true) { Set-ProgressText "Instalacja: AntyVirus..."; Install-AV; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "InstallAV" }

        if ($CheckboxControls["ImportWiFiProfile"].IsChecked -eq $true) { Set-ProgressText "Importowanie profilu Wi-Fi..."; Import-WiFiProfile; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "ImportWiFiProfile" }

        if ($CheckboxControls["CreateLocalAdmin"].IsChecked -eq $true) { Set-ProgressText "Tworzenie konta lokalnego administratora..."; New-LocalAdmin; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "CreateLocalAdmin" }

        if ($CheckboxControls["JoinDomain"].IsChecked -eq $true) { Set-ProgressText "Dołączanie do domeny..."; Join-Domain; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "JoinDomain" }

        if ($CheckboxControls["InstallApplications"].IsChecked -eq $true -and $script:SelectedApps.Count -gt 0) {
            Install-SelectedApps
            if ($script:isCancelled) { return }
            Save-DeploymentCheckpointStep "InstallApplications"
        }
        else {
            Write-Log "Instalacja aplikacji pominięta."
        }

        if ($CheckboxControls.ContainsKey("ChangeComputerName") -and $CheckboxControls["ChangeComputerName"].IsChecked -eq $true) { Set-ProgressText "Zmiana nazwy komputera..."; Set-NewComputerName; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "ChangeComputerName" }

        if ($CheckboxControls.ContainsKey("JoinIntune") -and $CheckboxControls["JoinIntune"].IsChecked -eq $true) { Set-ProgressText "Dołączanie do Intune..."; Join-Intune; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "JoinIntune" }

        if ($CheckboxControls.ContainsKey("RunPostInstallScripts") -and $CheckboxControls["RunPostInstallScripts"].IsChecked -eq $true) { Set-ProgressText "Skrypty poinstalacyjne..."; Invoke-PostInstallScripts; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "RunPostInstallScripts" }

        if ($CheckboxControls.ContainsKey("ExportHardwareAudit") -and $CheckboxControls["ExportHardwareAudit"].IsChecked -eq $true) { Set-ProgressText "Eksport audytu sprzętowego..."; Export-HardwareAuditTask; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "ExportHardwareAudit" }

        if ($CheckboxControls.ContainsKey("EnableBitLocker") -and $CheckboxControls["EnableBitLocker"].IsChecked -eq $true) { Set-ProgressText "Włączanie szyfrowania BitLocker..."; Enable-BitLockerEncryption; if ($script:isCancelled) { return }; Step-DeploymentProgress; Save-DeploymentCheckpointStep "EnableBitLocker" }

        if ($CheckboxControls["ChangeSystemSettings"].IsChecked -eq $true) { Set-ProgressText "Aplikowanie modyfikacji systemu..."; Set-SystemTweaks; Save-DeploymentCheckpointStep "ChangeSystemSettings" }
        if ($script:isCancelled) { return }

        if ($CheckboxControls["RunWindowsUpdate"].IsChecked -eq $true) {
            Set-ProgressText "Uruchamianie Windows Update..."
            Write-Log "Uruchamianie Windows Update..."
            Start-WindowsUpdate
            Save-DeploymentCheckpointStep "RunWindowsUpdate"
        }

        if ($CheckboxControls.ContainsKey("SuspendHibernation") -and $CheckboxControls["SuspendHibernation"].IsChecked -eq $true) { Set-ProgressText "Przywracanie ustawień hibernacji..."; Resume-Hibernation }

        Step-DeploymentProgress
        if ($progressBar.Value -lt 100 -and -not $script:isCancelled) { $progressBar.Value = 100 }

        if (-not $script:isCancelled) {
            if ($null -ne $script:stopwatchTimer) { $script:stopwatchTimer.Stop() }
            Set-ProgressText "Konfiguracja zakończona pomyślnie!"
            Write-Log "Konfiguracja zakończona."
            Clear-DeploymentCheckpoint
            
            if ($null -ne $notifyIcon -and $Window.WindowState -eq [System.Windows.WindowState]::Minimized) {
                $notifyIcon.ShowBalloonTip(5000, "Instalacja zakończona", "Wszystkie zadania zostały pomyślnie wykonane. Możesz przywrócić okno.", [System.Windows.Forms.ToolTipIcon]::Info)
            }
    
            if ($CheckboxControls.ContainsKey("AutoReboot") -and $CheckboxControls["AutoReboot"].IsChecked -eq $true -and $script:DryRun) {
                # Wcześniej Dry-Run naprawdę restartował komputer (Restart-Computer -Force).
                Write-Log "[DRY-RUN] Uruchomiono by ponownie komputer (AutoReboot)."
                Show-ThemedMessageBox -Message "Symulacja (Dry-Run) zakończona. W prawdziwym wdrożeniu komputer zostałby teraz uruchomiony ponownie." -Title "Zakończono" -Button "OK" -Image "Information" | Out-Null
            } elseif ($CheckboxControls.ContainsKey("AutoReboot") -and $CheckboxControls["AutoReboot"].IsChecked -eq $true) {
                Show-ThemedMessageBox -Message "Konfiguracja zakończona! Komputer uruchomi się ponownie po zamknięciu tego okna." -Title "Zakończono" -Button "OK" -Image "Information" | Out-Null
                Write-Log "Wymuszono ponowne uruchomienie systemu..."
                Restart-Computer -Force
            } else {
                Show-ThemedMessageBox -Message "Gotowe!" -Title "Zakończono" -Button "OK" -Image "Information" | Out-Null
            }
        }
    } catch {
        # Zabezpieczenie przed sytuacją, w której NIEOCZEKIWANY wyjątek w dowolnym miejscu
        # sekwencji wdrożenia (np. brak/niedostępność modułu Appx na niektórych kompilacjach
        # Windows, błąd sieciowy itp.) uciekał poza tę funkcję nieobsłużony aż do handlera
        # kliknięcia $btnStart, co mogło ubić całą aplikację (i uniemożliwić Pauzę/Przerwanie,
        # bo w praktyce proces już nie żył). Teraz taki błąd jest logowany i pokazywany
        # użytkownikowi, a przyciski/stan interfejsu i tak wracają do normy w bloku finally.
        Write-Log "Nieoczekiwany błąd podczas wdrożenia - przerwano: $_" -IsError
        Show-ThemedMessageBox -Message "Wystąpił nieoczekiwany błąd i wdrożenie zostało przerwane:`n`n$($_.Exception.Message)`n`nSzczegóły w logu." -Title "Błąd wdrożenia" -Button "OK" -Image "Error" | Out-Null
        $script:isCancelled = $true
    } finally {
        if ($null -ne $script:stopwatchTimer) { $script:stopwatchTimer.Stop() }
        $btnStart.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF107C10")
        $btnStart.IsEnabled = $true
        $btnPause.IsEnabled = $false
        $btnCancelDeploy.IsEnabled = $false
        $wasCancelled = $script:isCancelled
        $script:isPaused = $false
        $script:isCancelled = $false
        $script:DryRun = $false
        $script:CurrentLogContext = "System"
        $btnPause.Content = "Pauza"
        $btnPause.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFE6A100")

        if ($wasCancelled) {
            Set-ProgressText "Wdrożenie zostało przerwane."
        }
    }
}

function Get-AppSelection {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -IsError
        return
    }

    $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

    if ($null -eq $config.Programs) {
        Write-Log "Nie znaleziono sekcji 'Programs' w config.json" -IsError
        return
    }

    $script:SelectedApps.Clear()

    if ($null -ne $config.Programs.PSObject) {
        foreach ($app in $config.Programs.PSObject.Properties.Name) {
            $appObj = $config.Programs.$app
            if ($appObj.Enabled -eq $true -or $appObj.Enabled -match "true") {
                $script:SelectedApps[$app] = $true
            }
        }
    }

    $count = $script:SelectedApps.Count
    $appsList = if ($count -gt 0) { $script:SelectedApps.Keys -join ', ' } else { "Brak" }
    Write-Log "Załadowano domyślnie zaznaczone aplikacje ($count): $appsList"
    if ($null -ne $btnChooseApps) {
        $btnChooseApps.Content = "Wybierz aplikacje ($count)"
        $CheckboxControls["InstallApplications"].IsChecked = ($count -gt 0)
    }
}

function Load-Profiles {
    if ($null -eq $cmbProfiles) { return }
    $cmbProfiles.Items.Clear()
    [void]$cmbProfiles.Items.Add("--- Niestandardowy wybór ---")
    try {
        $cfg = Get-Config
        if ($null -ne $cfg.Profiles) {
            foreach ($prof in $cfg.Profiles.PSObject.Properties.Name | Sort-Object) {
                [void]$cmbProfiles.Items.Add($prof)
            }
        }
    } catch {}
    
    $script:ignoreProfileChange = $true
    $cmbProfiles.SelectedIndex = 0
    $script:ignoreProfileChange = $false
}

function Get-Config {
    if (-not (Test-Path $configPath)) {
        throw "Brak pliku config.json"
    }
    return (Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Save-Config($config) {
    $json = $config | ConvertTo-Json -Depth 10
    $json | Set-Content -Path $configPath -Encoding UTF8
    Write-Log "Zapisano zmiany do config.json" -Context "Użytkownik"
}

# Generuje czytelny, samodzielny raport HTML z (przefiltrowanych/posortowanych) wpisów logu -
# do wysłania klientowi/dołączenia do ticketu, zamiast surowego pliku tekstowego.
function Export-LogReportHtml {
    param(
        [array]$Entries,
        [string]$FiltersText,
        [string]$OutFile
    )

    $rowBg = @{
        'Error'   = '#FDECEA'
        'Warning' = '#FFF4E0'
        'Success' = '#E7F6E9'
        'DryRun'  = '#F2E9F7'
        'Info'    = '#FFFFFF'
    }
    $rowFg = @{
        'Error'   = '#B91C1C'
        'Warning' = '#92600A'
        'Success' = '#0F6B1F'
        'DryRun'  = '#6A3691'
        'Info'    = '#1F2328'
    }

    $errorCount = @($Entries | Where-Object { $_.RodzajKey -eq 'Error' }).Count
    $warnCount = @($Entries | Where-Object { $_.RodzajKey -eq 'Warning' }).Count
    $successCount = @($Entries | Where-Object { $_.RodzajKey -eq 'Success' }).Count

    $rowsHtml = New-Object System.Text.StringBuilder
    foreach ($e in $Entries) {
        $bg = $rowBg[$e.RodzajKey]; if (-not $bg) { $bg = '#FFFFFF' }
        $fg = $rowFg[$e.RodzajKey]; if (-not $fg) { $fg = '#1F2328' }
        [void]$rowsHtml.AppendLine("<tr style='background:$bg;color:$fg;'>")
        [void]$rowsHtml.AppendLine("<td class='c-rodzaj'>$([System.Net.WebUtility]::HtmlEncode($e.RodzajText))</td>")
        [void]$rowsHtml.AppendLine("<td class='c-data'>$([System.Net.WebUtility]::HtmlEncode($e.DataText))</td>")
        [void]$rowsHtml.AppendLine("<td class='c-kontekst'>$([System.Net.WebUtility]::HtmlEncode($e.Kontekst))</td>")
        [void]$rowsHtml.AppendLine("<td class='c-info'>$([System.Net.WebUtility]::HtmlEncode($e.Informacja))</td>")
        [void]$rowsHtml.AppendLine("</tr>")
    }

    $html = @"
<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<title>Raport wdrożenia - $([System.Net.WebUtility]::HtmlEncode($env:COMPUTERNAME))</title>
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #F4F5F7; color: #1F2328; margin: 0; padding: 32px; }
    .card { background: #FFFFFF; border-radius: 10px; padding: 28px 32px; max-width: 1200px; margin: 0 auto; box-shadow: 0 1px 3px rgba(0,0,0,0.12); }
    h1 { font-size: 22px; margin: 0 0 4px 0; }
    .subtitle { color: #6B7280; font-size: 13px; margin-bottom: 20px; }
    .stats { display: flex; gap: 14px; margin-bottom: 20px; flex-wrap: wrap; }
    .stat { border-radius: 8px; padding: 10px 16px; font-size: 13px; font-weight: 600; }
    .stat.total { background: #EEF1F5; color: #374151; }
    .stat.error { background: #FDECEA; color: #B91C1C; }
    .stat.warn { background: #FFF4E0; color: #92600A; }
    .stat.success { background: #E7F6E9; color: #0F6B1F; }
    .filters { font-size: 12px; color: #6B7280; margin-bottom: 20px; }
    table { border-collapse: collapse; width: 100%; font-size: 13px; }
    th { text-align: left; background: #1F2328; color: #FFFFFF; padding: 9px 10px; position: sticky; top: 0; }
    td { padding: 7px 10px; border-bottom: 1px solid #E5E7EB; vertical-align: top; }
    .c-rodzaj { white-space: nowrap; font-weight: 600; }
    .c-data { white-space: nowrap; font-family: Consolas, monospace; }
    .c-kontekst { white-space: nowrap; }
    .c-info { word-break: break-word; }
    .footer { margin-top: 20px; font-size: 11px; color: #9CA3AF; }
    @media print { body { background: #FFFFFF; padding: 0; } .card { box-shadow: none; } }
</style>
</head>
<body>
<div class="card">
    <h1>Raport wdrożenia - Smart Tool for Deployment</h1>
    <div class="subtitle">Stacja: $([System.Net.WebUtility]::HtmlEncode($env:COMPUTERNAME)) &nbsp;·&nbsp; Wygenerowano: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</div>
    <div class="stats">
        <div class="stat total">$($Entries.Count) wpisów</div>
        <div class="stat error">$errorCount błędów</div>
        <div class="stat warn">$warnCount ostrzeżeń</div>
        <div class="stat success">$successCount sukcesów</div>
    </div>
    <div class="filters">Zastosowane filtry: $([System.Net.WebUtility]::HtmlEncode($FiltersText))</div>
    <table>
        <thead><tr><th>Rodzaj</th><th>Data zdarzenia</th><th>Kontekst</th><th>Informacja</th></tr></thead>
        <tbody>
$($rowsHtml.ToString())
        </tbody>
    </table>
    <div class="footer">Wygenerowano przez Smart Tool for Deployment (STD) v$script:ScriptVersion</div>
</div>
</body>
</html>
"@

    $html | Set-Content -Path $OutFile -Encoding UTF8
}

# Podgląd pojedynczego wpisu logu na osobnym, większym oknie - przydatne gdy treść Informacji jest
# długa (np. pełny stack trace błędu) i nie mieści się w jednej linii tabeli. Tekst jest w polu
# tylko do odczytu, ale zaznaczalnym/kopiowalnym.
function Show-LogEntryDetail {
    param($Entry)
    if ($null -eq $Entry) { return }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Szczegóły wpisu logu" Width="620" Height="440" MinWidth="420" MinHeight="280" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="15,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row="0" Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="{DynamicResource ThemeTextBoxBg}" CornerRadius="4" Padding="10,5" Margin="0,0,10,0">
                <TextBlock Name="txtRodzaj" FontWeight="SemiBold"/>
            </Border>
            <Border Grid.Column="1" Background="{DynamicResource ThemeTextBoxBg}" CornerRadius="4" Padding="10,5" Margin="0,0,10,0">
                <TextBlock Name="txtData" FontFamily="Consolas"/>
            </Border>
            <Border Grid.Column="2" Background="{DynamicResource ThemeTextBoxBg}" CornerRadius="4" Padding="10,5">
                <TextBlock Name="txtKontekst"/>
            </Border>
        </Grid>
        <Border Grid.Row="1" Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6">
            <TextBox Name="txtInformacja" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" BorderThickness="0" Background="Transparent"
                     Foreground="{DynamicResource ThemeText}" FontFamily="Consolas" FontSize="13" Padding="12"
                     VerticalScrollBarVisibility="Auto"/>
        </Border>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Name="btnCopy" Content="Kopiuj" Width="100" Margin="0,0,10,0"/>
            <Button Name="btnClose" Content="Zamknij" Width="100" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg

    $dlg.FindName("txtRodzaj").Text = $Entry.RodzajText
    $dlg.FindName("txtData").Text = $Entry.DataText
    $dlg.FindName("txtKontekst").Text = "Kontekst: $($Entry.Kontekst)"
    $txtInformacja = $dlg.FindName("txtInformacja")
    $txtInformacja.Text = $Entry.Informacja

    $dlg.FindName("btnCopy").Add_Click({
        try { [System.Windows.Clipboard]::SetText($Entry.Informacja) } catch {}
    })
    $dlg.FindName("btnClose").Add_Click({ $dlg.Close() })

    $dlg.Add_Loaded({ $txtInformacja.Focus() | Out-Null; $txtInformacja.SelectAll() })
    if ($null -ne $script:ActiveLogWindow -and $script:ActiveLogWindow.IsLoaded) {
        $dlg.Owner = $script:ActiveLogWindow
    }
    $dlg.ShowDialog() | Out-Null
}

function Show-LogWindow {
    if ($null -ne $script:ActiveLogWindow -and $script:ActiveLogWindow.IsLoaded) {
        if ($script:ActiveLogWindow.WindowState -eq [System.Windows.WindowState]::Minimized) {
            $script:ActiveLogWindow.WindowState = [System.Windows.WindowState]::Normal
        }
        $script:ActiveLogWindow.Activate()
        return
    }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Przeglądarka Logów" Height="700" Width="1150" MinHeight="420" MinWidth="820" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="15,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center" TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Foreground="{TemplateBinding Foreground}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                            <TextBlock Text="▼" Foreground="{TemplateBinding Foreground}" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="10"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="8,0,25,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="250" CornerRadius="4">
                                    <ScrollViewer SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8,7"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBackground}"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{Binding RowColorHex}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBackground}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
                            <GridViewRowPresenter Content="{TemplateBinding Content}" Columns="{TemplateBinding GridView.ColumnCollection}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <!-- Bez własnego szablonu domyślny motyw Windows rysowałby tu jasne
                                 tło zaznaczenia/najechania niezależnie od naszych kolorów tekstu
                                 (RowColorHex), przez co wiersz stawał się nieczytelny. -->
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Name="txtSearch" Grid.Column="0" Margin="0,0,10,0" ToolTip="Szukaj w treści wpisów..."/>
            <ComboBox Name="cmbRodzajFilter" Grid.Column="1" Margin="0,0,10,0" SelectedIndex="0">
                <ComboBoxItem Content="Rodzaj: wszystkie"/>
                <ComboBoxItem Content="❌ Błędy"/>
                <ComboBoxItem Content="⚠️ Ostrzeżenia"/>
                <ComboBoxItem Content="✔️ Sukcesy"/>
                <ComboBoxItem Content="🧪 Dry-Run"/>
                <ComboBoxItem Content="ℹ️ Informacyjne"/>
            </ComboBox>
            <ComboBox Name="cmbKontekstFilter" Grid.Column="2" Margin="0,0,10,0" SelectedIndex="0">
                <ComboBoxItem Content="Kontekst: wszystkie"/>
                <ComboBoxItem Content="System"/>
                <ComboBoxItem Content="Automat"/>
                <ComboBoxItem Content="Użytkownik"/>
            </ComboBox>
            <CheckBox Name="chkAutoRefresh" Grid.Column="3" Content="Auto-odświeżanie" IsChecked="True" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="{DynamicResource ThemeText}"/>
            <Button Name="btnRefresh" Grid.Column="4" Content="Odśwież" Width="100"/>
        </Grid>

        <TextBlock Name="txtStats" Grid.Row="1" Margin="2,0,0,8" FontSize="12" Opacity="0.75"/>

        <Border Grid.Row="2" Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6">
            <ListView Name="lvLogs" BorderThickness="0" Background="Transparent" Foreground="{DynamicResource ThemeText}"
                      FontFamily="Consolas" FontSize="13" ScrollViewer.HorizontalScrollBarVisibility="Auto">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Rodzaj" Width="130" DisplayMemberBinding="{Binding RodzajText}"/>
                        <GridViewColumn Header="Data zdarzenia" Width="160" DisplayMemberBinding="{Binding DataText}"/>
                        <GridViewColumn Header="Kontekst" Width="110" DisplayMemberBinding="{Binding Kontekst}"/>
                        <GridViewColumn Header="Informacja" Width="560" DisplayMemberBinding="{Binding Informacja}"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </Border>

        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,12,0,0">
            <Button Name="btnExportHtml" Content="Eksportuj raport (HTML)" Width="180" Margin="0,0,10,0" Foreground="White" Background="#FF7A3E9D" FontWeight="SemiBold"/>
            <Button Name="btnSave" Content="Zapisz jako..." Width="120" Margin="0,0,10,0"/>
            <Button Name="btnClearLogs" Content="Wyczyść" Width="90" Margin="0,0,10,0" Foreground="White" Background="#FFC50F1F" FontWeight="SemiBold"/>
            <Button Name="btnOpenLog" Content="Otwórz plik" Width="110" Margin="0,0,10,0"/>
            <Button Name="btnOpenDir" Content="Otwórz folder" Width="120" Margin="0,0,10,0"/>
            <Button Name="btnZipLogs" Content="Spakuj do ZIP" Width="130"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $logWindow = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $logWindow

    $script:txtSearch = $logWindow.FindName("txtSearch")
    $script:cmbRodzajFilter = $logWindow.FindName("cmbRodzajFilter")
    $script:cmbKontekstFilter = $logWindow.FindName("cmbKontekstFilter")
    $script:txtStats = $logWindow.FindName("txtStats")
    $script:lvLogs = $logWindow.FindName("lvLogs")
    $btnRefresh = $logWindow.FindName("btnRefresh")
    $script:chkAutoRefresh = $logWindow.FindName("chkAutoRefresh")
    $btnExportHtml = $logWindow.FindName("btnExportHtml")
    $btnSave = $logWindow.FindName("btnSave")
    $btnClearLogs = $logWindow.FindName("btnClearLogs")
    $btnOpenLog = $logWindow.FindName("btnOpenLog")
    $btnOpenDir = $logWindow.FindName("btnOpenDir")
    $btnZipLogs = $logWindow.FindName("btnZipLogs")

    $script:rawLogText = ""
    $script:LastRenderedEntries = @()
    $script:LogSortColumn = "Data zdarzenia"
    $script:LogSortAscending = $false   # najnowsze na górze domyślnie

    # Kolor tekstu "zwykłej" informacji dopasowany do aktualnie aktywnego motywu (jasny/ciemny) -
    # pobrany raz z zasobów okna (zsynchronizowanych przez Apply-ThemeToWindow) i skonwertowany na
    # HEX, bo wiązanie Foreground="{Binding RowColorHex}" oczekuje stringa.
    $script:GetThemeTextHex = {
        try {
            $c = $logWindow.Resources["ThemeText"].Color
            return "#{0:X2}{1:X2}{2:X2}{3:X2}" -f $c.A, $c.R, $c.G, $c.B
        } catch { return "#FFDCDDDE" }
    }

    # Parsuje jedną linię pliku logu do obiektu z polami do wyświetlenia w tabeli. Rozumie trzy
    # warianty formatu (najnowszy jest zapisywany od teraz przez Write-Log, starsze to wpisy
    # sprzed kolejnych zmian formatu logowania):
    #   [yyyy-MM-dd HH:mm:ss] [INFO|ERROR] [Kontekst] tekst   <- aktualny
    #   [yyyy-MM-dd HH:mm:ss] [INFO|ERROR] tekst               <- bez kontekstu
    #   [HH:mm:ss] tekst                                       <- bez pełnej daty i kontekstu
    $script:ParseLogLine = {
        param([string]$Line, [int]$OrderIndex)

        $fmtFull = '^\[(?<date>\d{4}-\d{2}-\d{2}) (?<time>\d{2}:\d{2}:\d{2})\] \[(?<level>INFO|ERROR)\] \[(?<ctx>[^\]]+)\] (?<text>.*)$'
        $fmtNoCtx = '^\[(?<date>\d{4}-\d{2}-\d{2}) (?<time>\d{2}:\d{2}:\d{2})\] \[(?<level>INFO|ERROR)\] (?<text>.*)$'
        $fmtLegacy = '^\[(?<time>\d{2}:\d{2}:\d{2})\] (?<text>.*)$'

        $date = $null; $time = $null; $level = 'INFO'; $ctx = $null; $text = $Line
        if ($Line -match $fmtFull) {
            $date = $Matches['date']; $time = $Matches['time']; $level = $Matches['level']; $ctx = $Matches['ctx']; $text = $Matches['text']
        } elseif ($Line -match $fmtNoCtx) {
            $date = $Matches['date']; $time = $Matches['time']; $level = $Matches['level']; $text = $Matches['text']
        } elseif ($Line -match $fmtLegacy) {
            $time = $Matches['time']; $text = $Matches['text']
            if ($text -match '^\[ERROR\]\s*') { $level = 'ERROR'; $text = $text -replace '^\[ERROR\]\s*', '' }
        }

        $sortTs = [datetime]::MinValue
        $dataText = "—"
        if ($date) {
            try {
                $dt = [datetime]::ParseExact("$date $time", "yyyy-MM-dd HH:mm:ss", $null)
                $sortTs = $dt
                $dataText = $dt.ToString("dd.MM.yyyy HH:mm:ss")
            } catch { $dataText = "$date $time" }
        } elseif ($time) {
            $dataText = "— (godz. $time)"
        }

        $rodzajKey = 'Info'; $rodzajText = 'ℹ️ Informacja'; $colorHex = & $script:GetThemeTextHex
        if ($level -eq 'ERROR') { $rodzajKey = 'Error'; $rodzajText = '❌ Błąd'; $colorHex = '#FFC50F1F' }
        elseif ($text -match '\[DRY-RUN\]') { $rodzajKey = 'DryRun'; $rodzajText = '🧪 Dry-Run'; $colorHex = '#FF9B59B6' }
        elseif ($text -match 'Mało wolnego|Nie udało się|Pominięto|Błąd dostępu|BRAK POLECENIA') { $rodzajKey = 'Warning'; $rodzajText = '⚠️ Ostrzeżenie'; $colorHex = '#FFE6A100' }
        elseif ($text -match 'zakończon[ao] pomyślnie|Konfiguracja zakończona|Utworzono|zainstalowany\.|utworzony pomyślnie|Dołączono do domeny') { $rodzajKey = 'Success'; $rodzajText = '✔️ Sukces'; $colorHex = '#FF107C10' }

        [PSCustomObject]@{
            RodzajKey    = $rodzajKey
            RodzajText   = $rodzajText
            DataText     = $dataText
            SortTimestamp = $sortTs
            OrderIndex   = $OrderIndex
            Kontekst     = if ($ctx) { $ctx } else { "—" }
            Informacja   = $text
            RowColorHex  = $colorHex
            Level        = $level
            Raw          = $Line
        }
    }

    # Buduje tabelę logu (ListView/GridView) na podstawie surowego tekstu pliku oraz aktualnych
    # filtrów (wyszukiwarka, rodzaj, kontekst) i wybranego sortowania kolumny.
    $script:RenderLogView = {
        $content = $script:rawLogText
        $searchTerm = $script:txtSearch.Text
        $rodzajIdx = $script:cmbRodzajFilter.SelectedIndex
        $kontekstIdx = $script:cmbKontekstFilter.SelectedIndex

        if ([string]::IsNullOrWhiteSpace($content)) {
            $script:lvLogs.ItemsSource = $null
            $script:LastRenderedEntries = @()
            $script:txtStats.Text = "0 wpisów"
            return
        }

        $lines = @($content -split "`r?`n" | Where-Object { $_ -ne "" })
        $entries = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $entries.Add((& $script:ParseLogLine -Line $lines[$i] -OrderIndex $i))
        }

        $totalCount = $entries.Count
        $filtered = [System.Collections.Generic.List[object]]$entries

        $rodzajMap = @{ 1 = 'Error'; 2 = 'Warning'; 3 = 'Success'; 4 = 'DryRun'; 5 = 'Info' }
        if ($rodzajMap.ContainsKey($rodzajIdx)) {
            $wanted = $rodzajMap[$rodzajIdx]
            $filtered = [System.Collections.Generic.List[object]]@($filtered | Where-Object { $_.RodzajKey -eq $wanted })
        }

        $kontekstMap = @{ 1 = 'System'; 2 = 'Automat'; 3 = 'Użytkownik' }
        if ($kontekstMap.ContainsKey($kontekstIdx)) {
            $wanted = $kontekstMap[$kontekstIdx]
            $filtered = [System.Collections.Generic.List[object]]@($filtered | Where-Object { $_.Kontekst -eq $wanted })
        }

        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
            $filtered = [System.Collections.Generic.List[object]]@($filtered | Where-Object { $_.Raw -like "*$searchTerm*" })
        }

        $shownCount = $filtered.Count
        $errorCount = @($filtered | Where-Object { $_.RodzajKey -eq 'Error' }).Count

        $sortProp = switch ($script:LogSortColumn) {
            'Rodzaj'         { 'RodzajKey' }
            'Data zdarzenia' { 'SortTimestamp' }
            'Kontekst'       { 'Kontekst' }
            'Informacja'     { 'Informacja' }
            default          { 'SortTimestamp' }
        }
        $sorted = $filtered | Sort-Object -Property @{Expression = $sortProp; Descending = (-not $script:LogSortAscending)}, @{Expression = 'OrderIndex'; Descending = $false}

        $maxShow = 3000
        $truncated = $false
        $sortedArr = @($sorted)
        if ($sortedArr.Count -gt $maxShow) {
            $sortedArr = $sortedArr[0..($maxShow - 1)]
            $truncated = $true
        }

        $script:LastRenderedEntries = $sortedArr
        $script:lvLogs.ItemsSource = $sortedArr

        # Strzałka sortowania w nagłówku aktywnej kolumny (zgodnie ze wzorcem już użytym w
        # deinstalatorze zbiorczym - patrz Show-SoftwareUninstaller).
        try {
            $arrow = if ($script:LogSortAscending) { " ▲" } else { " ▼" }
            foreach ($col in $script:lvLogs.View.Columns) {
                $clean = [string]$col.Header -replace ' [▲▼]', ''
                $col.Header = if ($clean -eq $script:LogSortColumn) { "$clean$arrow" } else { $clean }
            }
        } catch {}

        $statsText = "$shownCount z $totalCount wpisów"
        if ($errorCount -gt 0) { $statsText += " · $errorCount błędów w widoku" }
        if ($truncated) { $statsText += " · pokazano pierwsze $maxShow" }
        $script:txtStats.Text = $statsText
    }

    # Sortowanie po kliknięciu nagłówka kolumny - ten sam sprawdzony wzorzec (ButtonBase.ClickEvent
    # + AddHandler) co w liście programów w Show-SoftwareUninstaller, dla spójności.
    $script:lvLogs.AddHandler(
        [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
        [System.Windows.RoutedEventHandler]{
            param($senderObj, $e)
            if ($e.OriginalSource -isnot [System.Windows.Controls.GridViewColumnHeader]) { return }
            $header = $e.OriginalSource
            if ($header.Role -eq [System.Windows.Controls.GridViewColumnHeaderRole]::Padding) { return }
            if ($null -eq $header.Column) { return }
            $colName = ([string]$header.Column.Header) -replace ' [▲▼]', ''
            if ($script:LogSortColumn -eq $colName) {
                $script:LogSortAscending = -not $script:LogSortAscending
            } else {
                $script:LogSortColumn = $colName
                $script:LogSortAscending = $true
            }
            & $script:RenderLogView
        }
    )

    # Podwójne kliknięcie na wiersz - podgląd pełnej treści (przydatne przy długich wpisach).
    $script:lvLogs.Add_MouseDoubleClick({
        $item = $script:lvLogs.SelectedItem
        if ($null -ne $item) { Show-LogEntryDetail -Entry $item }
    })

    $script:LoadLogs = {
        if (Test-Path "C:\deploy-log.txt") {
            try {
                $fs = New-Object System.IO.FileStream("C:\deploy-log.txt", [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                $content = $sr.ReadToEnd()
                $sr.Close()
                $fs.Close()
                if ($script:rawLogText -ne $content) {
                    $script:rawLogText = $content
                    & $script:RenderLogView
                }
            } catch {}
        } else {
            $script:rawLogText = ""
            & $script:RenderLogView
        }
    }

    $btnRefresh.Add_Click($script:LoadLogs)
    $script:txtSearch.Add_TextChanged({ & $script:RenderLogView })
    $script:cmbRodzajFilter.Add_SelectionChanged({ & $script:RenderLogView })
    $script:cmbKontekstFilter.Add_SelectionChanged({ & $script:RenderLogView })

    $script:logTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:logTimer.Interval = [TimeSpan]::FromSeconds(1)
    $script:logTimer.Add_Tick({
        if ($script:chkAutoRefresh.IsChecked -eq $true) {
            & $script:LoadLogs
        }
    })

    $btnClearLogs.Add_Click({
        if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz bezpowrotnie usunąć wszystkie wpisy z plików logów (C:\deploy-log.txt)?" -Title "Potwierdzenie czyszczenia" -Button "YesNo" -Image "Warning") -eq [System.Windows.MessageBoxResult]::Yes) {
            try { Clear-Content -Path "C:\deploy-log.txt" -ErrorAction SilentlyContinue } catch {}
            try { Clear-Content -Path "C:\deploy-error-log.txt" -ErrorAction SilentlyContinue } catch {}
            $script:rawLogText = ""
            & $script:RenderLogView
            Write-Log "Utworzono nowy dziennik operacji." -Context "Użytkownik"
        }
    })

    $btnExportHtml.Add_Click({
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Strona HTML (*.html)|*.html|Wszystkie pliki (*.*)|*.*"
        $sfd.FileName = "raport-wdrozenia_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                $filtersSummary = @()
                if (-not [string]::IsNullOrWhiteSpace($script:txtSearch.Text)) { $filtersSummary += "wyszukiwanie: `"$($script:txtSearch.Text)`"" }
                if ($script:cmbRodzajFilter.SelectedIndex -ne 0) { $filtersSummary += "rodzaj: $($script:cmbRodzajFilter.Text)" }
                if ($script:cmbKontekstFilter.SelectedIndex -ne 0) { $filtersSummary += "kontekst: $($script:cmbKontekstFilter.Text)" }
                $filtersText = if ($filtersSummary.Count -gt 0) { $filtersSummary -join ", " } else { "brak (pokazano wszystkie wpisy)" }

                Export-LogReportHtml -Entries $script:LastRenderedEntries -FiltersText $filtersText -OutFile $sfd.FileName
                Show-ThemedMessageBox -Message "Wyeksportowano raport do $($sfd.FileName)" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
                Start-Process $sfd.FileName
            } catch {
                Show-ThemedMessageBox -Message "Błąd podczas eksportu raportu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $btnSave.Add_Click({
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Pliki tekstowe (*.txt)|*.txt|Wszystkie pliki (*.*)|*.*"
        $sfd.FileName = "deploy-log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                $script:rawLogText | Set-Content -Path $sfd.FileName -Encoding UTF8
                Show-ThemedMessageBox -Message "Zapisano logi do $($sfd.FileName)" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
            } catch {
                Show-ThemedMessageBox -Message "Błąd podczas zapisywania: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $btnOpenLog.Add_Click({
        if (Test-Path "C:\deploy-log.txt") {
            Start-Process "notepad.exe" -ArgumentList "C:\deploy-log.txt"
        } else {
            Show-ThemedMessageBox -Message "Plik C:\deploy-log.txt nie istnieje." -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
        }
    })

    $btnOpenDir.Add_Click({
        if (Test-Path "C:\deploy-log.txt") {
            Start-Process "explorer.exe" -ArgumentList "/select,`"C:\deploy-log.txt`""
        } else {
            Start-Process "explorer.exe" -ArgumentList "C:\"
        }
    })

    $btnZipLogs.Add_Click({
        $logFiles = @("C:\deploy-log.txt", "C:\deploy-error-log.txt") | Where-Object { Test-Path $_ }
        if ($logFiles.Count -eq 0) {
            Show-ThemedMessageBox -Message "Brak plików logów do spakowania." -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
            return
        }
        
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Archiwum ZIP (*.zip)|*.zip|Wszystkie pliki (*.*)|*.*"
        $sfd.FileName = "deploy-logs_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                Compress-Archive -Path $logFiles -DestinationPath $sfd.FileName -Force
                Show-ThemedMessageBox -Message "Spakowano logi do $($sfd.FileName)" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
            } catch {
                Show-ThemedMessageBox -Message "Błąd podczas pakowania: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $logWindow.Add_Loaded({
        & $script:LoadLogs
        $script:logTimer.Start()
    })
    $logWindow.Add_KeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Escape) {
            $script:ActiveLogWindow.Close()
        }
    })
    $logWindow.Add_Closed({
        if ($null -ne $script:logTimer) {
            $script:logTimer.Stop()
            $script:logTimer = $null
        }
        $script:ActiveLogWindow = $null
    })
    
    $script:ActiveLogWindow = $logWindow
    $logWindow.Show()
}

function Show-UninstallConfirmDialog {
    param($AppCount, $AppListText)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Potwierdzenie deinstalacji" Width="550" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="UWAGA: Próba cichej deinstalacji $AppCount programów:" FontSize="15" FontWeight="SemiBold" Foreground="#FFD83B01" Margin="0,0,0,10"/>
        <Border Grid.Row="1" Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="4" Padding="10" MaxHeight="250" Margin="0,0,0,15">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TextBlock Name="txtAppList" TextWrapping="Wrap" FontFamily="Consolas" FontSize="13"/>
            </ScrollViewer>
        </Border>
        <TextBlock Grid.Row="2" Text="Czy na pewno chcesz kontynuować? Ta operacja usunie wybrane aplikacje." FontSize="14" Margin="0,0,0,20"/>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnYes" Content="Tak, odinstaluj" Width="140" Height="35" Background="#FFC50F1F" Foreground="White" FontWeight="SemiBold" Margin="0,0,10,0" IsDefault="True"/>
            <Button Name="btnNo" Content="Anuluj" Width="100" Height="35" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $txtAppList = $dlg.FindName("txtAppList")
    $btnYes = $dlg.FindName("btnYes")
    $btnNo = $dlg.FindName("btnNo")
    
    $txtAppList.Text = $AppListText
    
    $script:confirmUninstallResult = $false
    
    $btnYes.Add_Click({
        $script:confirmUninstallResult = $true
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    
    $btnNo.Add_Click({
        $script:confirmUninstallResult = $false
        $dlg.DialogResult = $false
        $dlg.Close()
    })
    
    $dlg.ShowDialog() | Out-Null
    return $script:confirmUninstallResult
}

function Show-CustomInfoDialog {
    param($Title, $Message, [switch]$ShowCopy, [string]$HtmlData)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$Title" Width="450" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="15,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ScrollViewer Grid.Row="0" MaxHeight="500" VerticalScrollBarVisibility="Auto" Margin="0,0,0,20" Padding="0,0,10,0">
            <TextBlock Name="txtMessage" FontSize="15" TextWrapping="Wrap" />
        </ScrollViewer>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnExportHTML" Content="Zapisz HTML" Width="110" Height="35" Margin="0,0,10,0" Background="#FF7A3E9D" Foreground="White" FontWeight="SemiBold" Visibility="Collapsed"/>
            <Button Name="btnCopy" Content="Kopiuj do schowka" Width="140" Height="35" Margin="0,0,10,0" Background="#FF107C10" Foreground="White" FontWeight="SemiBold" Visibility="Collapsed"/>
            <Button Name="btnOk" Content="OK" Width="100" Height="35" Background="#FF0078D7" Foreground="White" FontWeight="SemiBold" IsDefault="True" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $txtMessage = $dlg.FindName("txtMessage")
    $txtMessage.Text = $Message
    
    $btnCopy = $dlg.FindName("btnCopy")
    if ($ShowCopy) {
        $btnCopy.Visibility = [System.Windows.Visibility]::Visible
        $btnCopy.Add_Click({
            Set-Clipboard -Value $Message
            Show-ThemedMessageBox -Message "Skopiowano do schowka!" -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
        })
    }
    
    $btnExportHTML = $dlg.FindName("btnExportHTML")
    if (-not [string]::IsNullOrWhiteSpace($HtmlData)) {
        $btnExportHTML.Visibility = [System.Windows.Visibility]::Visible
        $btnExportHTML.Add_Click({
            $sfd = New-Object Microsoft.Win32.SaveFileDialog
            $sfd.Filter = "Pliki HTML (*.html)|*.html|Wszystkie pliki (*.*)|*.*"
            $sfd.FileName = "RaportSystemowy_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
            if ($sfd.ShowDialog() -eq $true) {
                try {
                    $HtmlData | Set-Content -Path $sfd.FileName -Encoding UTF8
                    Show-ThemedMessageBox -Message "Zapisano raport do $($sfd.FileName)" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
                } catch {
                    Show-ThemedMessageBox -Message "Błąd podczas zapisywania: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
                }
            }
        })
    }

    $btnOk = $dlg.FindName("btnOk")
    $btnOk.Add_Click({ $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
}

# Uporządkowany widok "Informacje o systemie" - zamiast jednej ściany tekstu w wąskim, nierozwijalnym
# oknie (Show-CustomInfoDialog): sekcje (System/BIOS, CPU/RAM, Dyski, Sieć) jako czytelne "chipy" +
# lista zainstalowanego oprogramowania jako sortowalna/przeszukiwalna tabela (to zwykle najdłuższa,
# najmniej czytelna część raportu).
function Show-SystemInfoWindow {
    $audit = Get-HardwareAudit -AsObject

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Informacje o systemie" Width="800" Height="700" MinWidth="620" MinHeight="450" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="15,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8,7"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBackground}"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBackground}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
                            <GridViewRowPresenter Content="{TemplateBinding Content}" Columns="{TemplateBinding GridView.ColumnCollection}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ChipBorder" TargetType="Border">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="6"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Margin" Value="0,0,10,10"/>
        </Style>
        <Style x:Key="ChipBorderClickable" TargetType="Border" BasedOn="{StaticResource ChipBorder}">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="ToolTip" Value="Kliknij, aby skopiować"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="BorderBrush" Value="{DynamicResource ThemeButton}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="CopyableLine" TargetType="TextBlock">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="ToolTip" Value="Kliknij, aby skopiować"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="TextDecorations" Value="Underline"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,14">
            <TextBlock Name="txtHeader" FontSize="18" FontWeight="Bold"/>
            <TextBlock Name="txtMeta" FontSize="12" Opacity="0.7" Margin="0,3,0,0"/>
        </StackPanel>

        <WrapPanel Grid.Row="1">
            <Border Name="chipOs" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="💻 System operacyjny  📋" FontSize="11" Opacity="0.65"/>
                    <TextBlock Name="txtOs" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipBiosSn" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="🔧 BIOS - numer seryjny (SN)  📋" FontSize="11" Opacity="0.65"/>
                    <TextBlock Name="txtBiosSn" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipBiosVer" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="🔧 BIOS - wersja  📋" FontSize="11" Opacity="0.65"/>
                    <TextBlock Name="txtBiosVer" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipCpu" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="⚙️ Procesor  📋" FontSize="11" Opacity="0.65"/>
                    <TextBlock Name="txtCpu" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipRam" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="🧠 Pamięć RAM  📋" FontSize="11" Opacity="0.65"/>
                    <TextBlock Name="txtRam" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
        </WrapPanel>

        <Grid Grid.Row="2" Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Style="{StaticResource ChipBorder}" Margin="0,0,10,0">
                <StackPanel>
                    <TextBlock Text="💾 Dyski twarde" FontSize="11" Opacity="0.65" Margin="0,0,0,4"/>
                    <ItemsControl Name="icDisks">
                        <ItemsControl.ItemTemplate>
                            <DataTemplate>
                                <TextBlock Text="{Binding}" Style="{StaticResource CopyableLine}" FontWeight="SemiBold" Margin="0,1"/>
                            </DataTemplate>
                        </ItemsControl.ItemTemplate>
                    </ItemsControl>
                </StackPanel>
            </Border>
            <Border Grid.Column="1" Style="{StaticResource ChipBorder}" Margin="0">
                <StackPanel>
                    <TextBlock Text="🌐 Karty sieciowe" FontSize="11" Opacity="0.65" Margin="0,0,0,4"/>
                    <ItemsControl Name="icNets">
                        <ItemsControl.ItemTemplate>
                            <DataTemplate>
                                <TextBlock Text="{Binding}" Style="{StaticResource CopyableLine}" FontWeight="SemiBold" Margin="0,1" TextWrapping="Wrap"/>
                            </DataTemplate>
                        </ItemsControl.ItemTemplate>
                    </ItemsControl>
                </StackPanel>
            </Border>
        </Grid>

        <Grid Grid.Row="3" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="260"/>
            </Grid.ColumnDefinitions>
            <TextBlock Name="txtAppsHeader" Grid.Column="0" FontSize="14" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <TextBox Name="txtSearch" Grid.Column="1" ToolTip="Szukaj po nazwie programu..."/>
        </Grid>

        <Border Grid.Row="4" Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6">
            <ListView Name="lvApps" BorderThickness="0" Background="Transparent" Foreground="{DynamicResource ThemeText}">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Nazwa programu" Width="480" DisplayMemberBinding="{Binding DisplayName}"/>
                        <GridViewColumn Header="Wersja" Width="180" DisplayMemberBinding="{Binding DisplayVersion}"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </Border>

        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button Name="btnCopy" Content="Kopiuj do schowka" Width="150" Margin="0,0,10,0" Background="#FF107C10" Foreground="White" FontWeight="SemiBold"/>
            <Button Name="btnExportHtml" Content="Zapisz jako HTML" Width="150" Margin="0,0,10,0" Background="#FF7A3E9D" Foreground="White" FontWeight="SemiBold"/>
            <Button Name="btnClose" Content="Zamknij" Width="100" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg

    $dlg.FindName("txtHeader").Text = "🖥️ $($audit.ComputerName)"
    $dlg.FindName("txtMeta").Text = "Wygenerowano: $($audit.GeneratedAt.ToString('dd.MM.yyyy HH:mm:ss'))  ·  Użytkownik: $($audit.UserName)"
    $dlg.FindName("txtOs").Text = $audit.OsSummary
    $dlg.FindName("txtBiosSn").Text = $audit.BiosSerial
    $dlg.FindName("txtBiosVer").Text = $audit.BiosVersion
    $dlg.FindName("txtCpu").Text = $audit.CpuName
    $dlg.FindName("txtRam").Text = "$($audit.RamGb) GB"

    # Kopiowanie pojedynczych wartości (np. numeru seryjnego BIOS) do schowka jednym kliknięciem -
    # z krótkim wizualnym potwierdzeniem ("✅ Skopiowano!") zamiast modalnego okienka.
    # UWAGA: Tick NIE zatrzymuje się przez $script:CopyFeedbackTimer (ten zasięg rozwiązuje się na
    # żywo, a nie jest "zamrożony" przez .GetNewClosure()) - jeśli okno zostanie zamknięte i
    # otwarte ponownie zanim poprzedni timer wygaśnie, zmienna skryptowa zostaje nadpisana/
    # wyzerowana i osierocony timer wywołuje Stop() na $null. Zamiast tego używamy parametru
    # "sender" samego zdarzenia Tick - to zawsze dokładnie TEN konkretny timer, bez dwuznaczności.
    $script:CopyFeedbackTimer = $null
    $copyTextWithFeedback = {
        param([System.Windows.Controls.TextBlock]$TextBlock, [string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return }
        try { Set-Clipboard -Value $Value } catch { return }

        if ($null -ne $script:CopyFeedbackTimer) { try { $script:CopyFeedbackTimer.Stop() } catch {} }
        $originalText = $TextBlock.Text
        $TextBlock.Text = "✅ Skopiowano!"
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(900)
        $timer.Add_Tick({
            param($senderObj, $tickArgs)
            $TextBlock.Text = $originalText
            $senderObj.Stop()
        }.GetNewClosure())
        $script:CopyFeedbackTimer = $timer
        $timer.Start()
    }

    foreach ($pair in @(
        @{ Chip = "chipOs"; Value = "txtOs" }
        @{ Chip = "chipBiosSn"; Value = "txtBiosSn" }
        @{ Chip = "chipBiosVer"; Value = "txtBiosVer" }
        @{ Chip = "chipCpu"; Value = "txtCpu" }
        @{ Chip = "chipRam"; Value = "txtRam" }
    )) {
        $chip = $dlg.FindName($pair.Chip)
        $valueBlock = $dlg.FindName($pair.Value)
        $chip.Add_MouseLeftButtonUp({ & $copyTextWithFeedback -TextBlock $valueBlock -Value $valueBlock.Text }.GetNewClosure())
    }

    $icDisks = $dlg.FindName("icDisks")
    $diskLines = if ($audit.Disks.Count -gt 0) {
        @($audit.Disks | ForEach-Object {
            $snPart = if ($_.SerialNumber) { " — SN: $($_.SerialNumber)" } else { "" }
            "• $($_.Model) — $($_.SizeGb) GB$snPart"
        })
    } else { @("Brak danych") }
    $icDisks.ItemsSource = $diskLines

    $icNets = $dlg.FindName("icNets")
    $netLines = if ($audit.Networks.Count -gt 0) { @($audit.Networks | ForEach-Object { "• $($_.Description) — $($_.Mac) — $($_.Ip)" }) } else { @("Brak aktywnych kart sieciowych") }
    $icNets.ItemsSource = $netLines

    # Klik na dowolną linię dysku/karty sieciowej kopiuje jej pełną treść (m.in. numer seryjny
    # dysku, adres MAC) - jeden wspólny handler na ItemsControl zamiast po jednym na wpis.
    foreach ($ic in @($icDisks, $icNets)) {
        $ic.AddHandler(
            [System.Windows.UIElement]::MouseLeftButtonUpEvent,
            [System.Windows.Input.MouseButtonEventHandler]{
                param($senderObj, $e)
                $tb = $e.OriginalSource -as [System.Windows.Controls.TextBlock]
                if ($null -ne $tb -and $tb.Text -ne "Brak danych" -and $tb.Text -ne "Brak aktywnych kart sieciowych") {
                    & $copyTextWithFeedback -TextBlock $tb -Value ($tb.Text -replace '^•\s*', '')
                }
            }.GetNewClosure()
        )
    }

    $txtAppsHeader = $dlg.FindName("txtAppsHeader")
    $txtSearch = $dlg.FindName("txtSearch")
    $lvApps = $dlg.FindName("lvApps")

    $allApps = @($audit.InstalledApps)
    $script:SysInfoSortColumn = "Nazwa programu"
    $script:SysInfoSortAscending = $true

    $renderApps = {
        $term = $txtSearch.Text
        $filtered = if ([string]::IsNullOrWhiteSpace($term)) { $allApps } else { @($allApps | Where-Object { $_.DisplayName -like "*$term*" }) }

        $sortProp = if ($script:SysInfoSortColumn -eq "Wersja") { "DisplayVersion" } else { "DisplayName" }
        $sorted = @($filtered | Sort-Object -Property @{Expression = $sortProp; Descending = (-not $script:SysInfoSortAscending)})

        $lvApps.ItemsSource = $sorted
        $txtAppsHeader.Text = if ($audit.AppError) { "📦 Zainstalowane oprogramowanie - błąd odczytu" } else { "📦 Zainstalowane oprogramowanie ($($sorted.Count) z $($allApps.Count))" }

        try {
            $arrow = if ($script:SysInfoSortAscending) { " ▲" } else { " ▼" }
            foreach ($col in $lvApps.View.Columns) {
                $clean = [string]$col.Header -replace ' [▲▼]', ''
                $col.Header = if ($clean -eq $script:SysInfoSortColumn) { "$clean$arrow" } else { $clean }
            }
        } catch {}
    }

    $txtSearch.Add_TextChanged({ & $renderApps })

    $lvApps.AddHandler(
        [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
        [System.Windows.RoutedEventHandler]{
            param($senderObj, $e)
            if ($e.OriginalSource -isnot [System.Windows.Controls.GridViewColumnHeader]) { return }
            $header = $e.OriginalSource
            if ($header.Role -eq [System.Windows.Controls.GridViewColumnHeaderRole]::Padding) { return }
            if ($null -eq $header.Column) { return }
            $colName = ([string]$header.Column.Header) -replace ' [▲▼]', ''
            if ($script:SysInfoSortColumn -eq $colName) {
                $script:SysInfoSortAscending = -not $script:SysInfoSortAscending
            } else {
                $script:SysInfoSortColumn = $colName
                $script:SysInfoSortAscending = $true
            }
            & $renderApps
        }
    )

    & $renderApps

    $dlg.FindName("btnCopy").Add_Click({
        try {
            Set-Clipboard -Value (Get-HardwareAudit)
            Show-ThemedMessageBox -Message "Skopiowano do schowka!" -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
        } catch {}
    })

    $dlg.FindName("btnExportHtml").Add_Click({
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Pliki HTML (*.html)|*.html|Wszystkie pliki (*.*)|*.*"
        $sfd.FileName = "RaportSystemowy_$($audit.ComputerName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                Get-HardwareAudit -AsHtml | Set-Content -Path $sfd.FileName -Encoding UTF8
                Show-ThemedMessageBox -Message "Zapisano raport do $($sfd.FileName)" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
                Start-Process $sfd.FileName
            } catch {
                Show-ThemedMessageBox -Message "Błąd podczas zapisywania: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $dlg.FindName("btnClose").Add_Click({ $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
}

function Show-SoftwareUninstaller {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Deinstalator Oprogramowania (Wybór zbiorczy)" Height="600" Width="980" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="15,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="0,2" CornerRadius="2">
                            <GridViewRowPresenter Content="{TemplateBinding Content}" Columns="{TemplateBinding GridView.ColumnCollection}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeBorder}"/></Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/>
                                <Setter Property="TextElement.Foreground" Value="{DynamicResource ThemeButtonText}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="🔍 Szukaj:" VerticalAlignment="Center" FontSize="15" Margin="0,0,10,0"/>
            <TextBox Name="txtSearch" Grid.Column="1" Height="30" Margin="0,0,10,0" FontSize="14"/>
            <Button Name="btnSelectAllApps" Content="Zaznacz widoczne" Grid.Column="2" Width="130" Height="30" Margin="0,0,5,0"/>
            <Button Name="btnDeselectAllApps" Content="Odznacz widoczne" Grid.Column="3" Width="130" Height="30" Margin="0,0,5,0"/>
            <Button Name="btnExportCSV" Content="Eksportuj CSV" Grid.Column="4" Width="110" Height="30" Margin="0,0,10,0"/>
            <Button Name="btnExportHTML" Content="Eksportuj HTML" Grid.Column="5" Width="120" Height="30" Margin="0,0,10,0"/>
            <Button Name="btnRefresh" Content="Odśwież listę" Grid.Column="6" Width="110" Height="30"/>
        </Grid>
        <ListView Name="lvApps" Grid.Row="1" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}">
            <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Name="miOpenLocation" Header="Pokaż w eksploratorze"/>
                </ContextMenu>
            </ListView.ContextMenu>
            <ListView.View>
                <GridView>
                    <GridViewColumn Width="40">
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <CheckBox IsChecked="{Binding IsChecked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn Header="Nazwa programu ▲" Width="360">
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <StackPanel Orientation="Horizontal">
                                    <Image Source="{Binding Icon}" Width="16" Height="16" Margin="0,0,5,0"/>
                                <TextBlock Text="{Binding StatusText}" Foreground="{Binding StatusColor}" FontWeight="Bold" Margin="0,0,5,0" VerticalAlignment="Center"/>
                                    <TextBlock Text="{Binding DisplayName}" VerticalAlignment="Center"/>
                                </StackPanel>
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn Header="Wersja" Width="100" DisplayMemberBinding="{Binding DisplayVersion}"/>
                    <GridViewColumn Header="Data instalacji" Width="110" DisplayMemberBinding="{Binding InstallDate}"/>
                    <GridViewColumn Header="Wydawca" Width="220" DisplayMemberBinding="{Binding Publisher}"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Grid Grid.Row="2" Margin="0,15,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Name="lblSelectedCount" Grid.Column="0" Text="Wybrano: 0" VerticalAlignment="Center" Margin="0,0,15,0" FontWeight="SemiBold" FontSize="14"/>
            <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="0,0,15,0">
                <TextBlock Name="lblUninstallStatus" Text="" FontSize="11" Foreground="{DynamicResource ThemeText}" Margin="0,0,0,3" Visibility="Hidden" TextTrimming="CharacterEllipsis" Opacity="0.8"/>
                <ProgressBar Name="pbUninstall" Height="15" Minimum="0" Maximum="100" BorderThickness="0" Foreground="#FFC50F1F" Visibility="Hidden"/>
            </StackPanel>
            <StackPanel Grid.Column="2" Orientation="Horizontal">
                <Button Name="btnPauseUninstall" Content="Pauza" Width="90" Height="35" Background="#FFE6A100" Foreground="White" Margin="0,0,10,0" FontWeight="SemiBold" IsEnabled="False"/>
                <Button Name="btnKill" Content="Zabij proces" Width="120" Height="35" Background="#FFE6A100" Foreground="White" Margin="0,0,10,0" FontWeight="SemiBold" IsEnabled="False" ToolTip="Wymuś zamknięcie zawieszonego deinstalatora"/>
                <Button Name="btnUninstall" Content="Odinstaluj (Cicho)" Width="190" Height="35" Background="#FFC50F1F" Foreground="White" Margin="0,0,10,0" FontWeight="SemiBold"/>
                <Button Name="btnClose" Content="Zamknij" Width="90" Height="35" IsCancel="True"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $uninstWindow = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $uninstWindow

    $txtSearch = $uninstWindow.FindName("txtSearch")
    $btnRefresh = $uninstWindow.FindName("btnRefresh")
    $btnSelectAllApps = $uninstWindow.FindName("btnSelectAllApps")
    $btnDeselectAllApps = $uninstWindow.FindName("btnDeselectAllApps")
    $btnExportCSV = $uninstWindow.FindName("btnExportCSV")
    $btnExportHTML = $uninstWindow.FindName("btnExportHTML")
    $lblSelectedCount = $uninstWindow.FindName("lblSelectedCount")
    $lblUninstallStatus = $uninstWindow.FindName("lblUninstallStatus")
    $pbUninstall = $uninstWindow.FindName("pbUninstall")
    $lvApps = $uninstWindow.FindName("lvApps")
    $miOpenLocation = $uninstWindow.FindName("miOpenLocation")
    $btnUninstall = $uninstWindow.FindName("btnUninstall")
    $btnPauseUninstall = $uninstWindow.FindName("btnPauseUninstall")
    $btnKill = $uninstWindow.FindName("btnKill")
    $btnClose = $uninstWindow.FindName("btnClose")
    $script:uninstProc = $null
    $script:uninstSortCol = "DisplayName"
    $script:uninstSortDir = "Ascending"
    $script:isUninstallPaused = $false

    $script:uninstAllApps = @()
    $script:iconCache = @{}

    $GetAppIcon = {
        param($iconPath)
        if ([string]::IsNullOrWhiteSpace($iconPath)) { return $null }
        $cleanPath = $iconPath.Trim()
        if ($cleanPath -match '^(.*),-?\d+$') { $cleanPath = $matches[1] }
        $cleanPath = $cleanPath -replace '"', ''
        $cleanPath = [System.Environment]::ExpandEnvironmentVariables($cleanPath)

        if ($script:iconCache.ContainsKey($cleanPath)) {
            return $script:iconCache[$cleanPath]
        }

        try {
            if (-not (Test-Path -LiteralPath $cleanPath -ErrorAction Stop)) { $script:iconCache[$cleanPath] = $null; return $null }
        } catch { $script:iconCache[$cleanPath] = $null; return $null }

        try {
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($cleanPath)
            if ($null -ne $icon) {
                $bmpSrc = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon($icon.Handle, [System.Windows.Int32Rect]::Empty, [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
                $bmpSrc.Freeze()
                $icon.Dispose()
                $script:iconCache[$cleanPath] = $bmpSrc
                return $bmpSrc
            }
        } catch {}
        $script:iconCache[$cleanPath] = $null
        return $null
    }

    $LoadApps = {
        $uninstWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $paths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            $appList = New-Object System.Collections.Generic.List[PSCustomObject]
            foreach ($basePath in $paths) {
                if (Test-Path -LiteralPath $basePath) {
                    foreach ($key in (Get-ChildItem -LiteralPath $basePath -ErrorAction SilentlyContinue)) {
                        $displayName = $key.GetValue("DisplayName")
                        $systemComponent = $key.GetValue("SystemComponent")
                        $parentKeyName = $key.GetValue("ParentKeyName")
                        if ($displayName -and (-not $systemComponent) -and ($null -eq $parentKeyName) -and ($displayName -notmatch '^KB\d+')) {
                            $dateStr = $key.GetValue("InstallDate")
                            if ($dateStr -match '^\d{8}$') {
                                $dateStr = "$($dateStr.Substring(0,4))-$($dateStr.Substring(4,2))-$($dateStr.Substring(6,2))"
                            }
                            $iconSrc = & $GetAppIcon ($key.GetValue("DisplayIcon"))
                            $appList.Add([PSCustomObject]@{
                                IsChecked = $false
                                Icon = $iconSrc
                                StatusText = ""
                                StatusColor = "Transparent"
                                DisplayName = $displayName
                                DisplayVersion = $key.GetValue("DisplayVersion")
                                InstallDate = $dateStr
                                Publisher = $key.GetValue("Publisher")
                                UninstallString = $key.GetValue("UninstallString")
                                QuietUninstallString = $key.GetValue("QuietUninstallString")
                                InstallLocation = $key.GetValue("InstallLocation")
                                DisplayIcon = $key.GetValue("DisplayIcon")
                            })
                        }
                    }
                }
            }
            $script:uninstAllApps = $appList | Sort-Object DisplayName -Unique
            
            & $UpdateList
            Write-Log "Odświeżono listę zainstalowanych programów w Deinstalatorze ($($script:uninstAllApps.Count) pozycji)."
        } finally {
            $uninstWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
            $uninstWindow.Dispatcher.Invoke([Action]{ 
                $count = @($script:uninstAllApps | Where-Object { $_.IsChecked }).Count
                $lblSelectedCount.Text = "Wybrano: $count" 
            }) | Out-Null
        }
    }

    $UpdateList = {
        $filter = $txtSearch.Text.ToLower()
        $filtered = $script:uninstAllApps | Where-Object {
            [string]::IsNullOrWhiteSpace($filter) -or $_.DisplayName.ToLower().Contains($filter) -or ($_.Publisher -and $_.Publisher.ToLower().Contains($filter))
        }
        if ($script:uninstSortDir -eq "Ascending") { $filtered = $filtered | Sort-Object $script:uninstSortCol }
        else { $filtered = $filtered | Sort-Object $script:uninstSortCol -Descending }
        $lvApps.Items.Clear()
        foreach ($app in $filtered) {
            [void]$lvApps.Items.Add($app)
        }
    }

    $btnRefresh.Add_Click($LoadApps)
    $txtSearch.Add_TextChanged($UpdateList)
    
    $btnSelectAllApps.Add_Click({
        foreach ($app in $lvApps.Items) { $app.IsChecked = $true }
        $lvApps.Items.Refresh()
        $uninstWindow.Dispatcher.Invoke([Action]{ $lblSelectedCount.Text = "Wybrano: $(@($script:uninstAllApps | Where-Object { $_.IsChecked }).Count)" }) | Out-Null
    })
    $btnDeselectAllApps.Add_Click({
        foreach ($app in $lvApps.Items) { $app.IsChecked = $false }
        $lvApps.Items.Refresh()
        $uninstWindow.Dispatcher.Invoke([Action]{ $lblSelectedCount.Text = "Wybrano: $(@($script:uninstAllApps | Where-Object { $_.IsChecked }).Count)" }) | Out-Null
    })

    $btnExportCSV.Add_Click({
        $itemsToExport = @($lvApps.Items)
        if ($itemsToExport.Count -eq 0) {
            Show-ThemedMessageBox -Message "Brak programów do wyeksportowania na widocznej liście." -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
            return
        }
        
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Pliki CSV (*.csv)|*.csv|Wszystkie pliki (*.*)|*.*"
        $sfd.FileName = "ZainstalowaneProgramy_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                $itemsToExport | Select-Object DisplayName, DisplayVersion, InstallDate, Publisher | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                Show-ThemedMessageBox -Message "Wyeksportowano pomyślnie do $($sfd.FileName)" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
            } catch {
                Show-ThemedMessageBox -Message "Błąd eksportu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $btnExportHTML.Add_Click({
        $itemsToExport = @($lvApps.Items)
        if ($itemsToExport.Count -eq 0) {
            Show-ThemedMessageBox -Message "Brak programów do wyeksportowania na widocznej liście." -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
            return
        }
        
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Pliki HTML (*.html)|*.html|Wszystkie pliki (*.*)|*.*"
        $sfd.FileName = "ZainstalowaneProgramy_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                $html = @"
<!DOCTYPE html>
<html lang='pl'>
<head>
    <meta charset='utf-8'>
    <title>Raport: Zainstalowane Oprogramowanie</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f9; color: #333; padding: 20px; }
        h1 { color: #0078D7; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; background-color: white; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
        th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #0078D7; color: white; position: sticky; top: 0; }
        tr:hover { background-color: #f1f1f1; }
        .meta { font-size: 0.9em; color: #666; margin-bottom: 20px; }
    </style>
</head>
<body>
    <h1>Zainstalowane Oprogramowanie</h1>
    <div class='meta'>Wygenerowano: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>Komputer: $env:COMPUTERNAME</div>
    <table>
        <tr><th>Nazwa programu</th><th>Wersja</th><th>Data instalacji</th><th>Wydawca</th></tr>
"@
                foreach ($app in $itemsToExport) {
                    $name = [System.Security.SecurityElement]::Escape([string]$app.DisplayName)
                    $ver = [System.Security.SecurityElement]::Escape([string]$app.DisplayVersion)
                    $date = [System.Security.SecurityElement]::Escape([string]$app.InstallDate)
                    $pub = [System.Security.SecurityElement]::Escape([string]$app.Publisher)
                    $html += "        <tr><td>$name</td><td>$ver</td><td>$date</td><td>$pub</td></tr>`r`n"
                }
                $html += @"
    </table>
</body>
</html>
"@
                $html | Set-Content -Path $sfd.FileName -Encoding UTF8
                Show-ThemedMessageBox -Message "Wyeksportowano pomyślnie do $($sfd.FileName)" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
            } catch {
                Show-ThemedMessageBox -Message "Błąd eksportu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $btnClose.Add_Click({ $uninstWindow.Close() })

    $lvApps.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        if ($e.OriginalSource -is [System.Windows.Controls.CheckBox]) {
            $chk = $e.OriginalSource
            if ($null -ne $chk.DataContext) {
                $chk.DataContext.IsChecked = ($true -eq $chk.IsChecked)
            }
            $uninstWindow.Dispatcher.Invoke([Action]{
                $count = @($script:uninstAllApps | Where-Object { $_.IsChecked }).Count
                $lblSelectedCount.Text = "Wybrano: $count"
            }) | Out-Null
        }
        if ($e.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            $header = $e.OriginalSource
            if ($header.Role -ne [System.Windows.Controls.GridViewColumnHeaderRole]::Padding) {
                $cleanHeader = $header.Column.Header -replace ' [▲▼]', ''
                $colName = switch ($cleanHeader) {
                    "Nazwa programu" { "DisplayName" }
                    "Wersja" { "DisplayVersion" }
                    "Data instalacji" { "InstallDate" }
                    "Wydawca" { "Publisher" }
                    default { "DisplayName" }
                }
                if ($script:uninstSortCol -eq $colName) {
                    $script:uninstSortDir = if ($script:uninstSortDir -eq "Ascending") { "Descending" } else { "Ascending" }
                } else {
                    $script:uninstSortCol = $colName
                    $script:uninstSortDir = "Ascending"
                }
                foreach ($col in $lvApps.View.Columns) { $col.Header = $col.Header -replace ' [▲▼]', '' }
                $arrow = if ($script:uninstSortDir -eq "Ascending") { " ▲" } else { " ▼" }
                $header.Column.Header = "$cleanHeader$arrow"
                & $UpdateList
            }
        }
    })

    $miOpenLocation.Add_Click({
        if ($lvApps.SelectedItem) {
            $app = $lvApps.SelectedItem
            $path = $app.InstallLocation

            if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) {
                if (-not [string]::IsNullOrWhiteSpace($app.DisplayIcon)) {
                    $p = $app.DisplayIcon -replace '"', ''
                    if ($p -match '^(.*?),-?\d+$') { $p = $matches[1] }
                    $p = [System.Environment]::ExpandEnvironmentVariables($p)
                    try {
                        if (Test-Path -LiteralPath $p) {
                            if ((Get-Item -LiteralPath $p).PSIsContainer) { $path = $p }
                            else { $path = Split-Path -LiteralPath $p -Parent }
                        }
                    } catch {}
                }
            }

            if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) {
                if (-not [string]::IsNullOrWhiteSpace($app.UninstallString)) {
                    $uStr = $app.UninstallString
                    $p = if ($uStr -match '^"([^"]+)"') { $matches[1] } else { ($uStr -split ' ')[0] }
                    $p = [System.Environment]::ExpandEnvironmentVariables($p)
                    try { if (Test-Path -LiteralPath $p) { $path = Split-Path -LiteralPath $p -Parent } } catch {}
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
                Start-Process "explorer.exe" -ArgumentList "`"$path`""
            } else {
            Show-ThemedMessageBox -Message "Nie udało się automatycznie ustalić ścieżki instalacji dla tego programu." -Title "Brak ścieżki" -Button "OK" -Image "Warning" | Out-Null
            }
        }
    })

    $lvApps.Add_MouseDoubleClick({
        if ($lvApps.SelectedItem) {
            $app = $lvApps.SelectedItem
            $cmd = $app.UninstallString
            if ([string]::IsNullOrWhiteSpace($cmd)) {
                $cmd = $app.QuietUninstallString
            }

            if (-not [string]::IsNullOrWhiteSpace($cmd)) {
                if ($cmd -match "(?i)msiexec") {
                    $cmd = $cmd -replace "(?i)/I", "/X"
                }
                Write-Log "Uruchamianie interaktywnego deinstalatora dla $($app.DisplayName)..."
                try {
                    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd"
                } catch {
                Show-ThemedMessageBox -Message "Błąd uruchamiania deinstalatora: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
                }
            } else {
            Show-ThemedMessageBox -Message "Brak ścieżki deinstalatora w rejestrze." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
            }
        }
    })

    $uninstWindow.Add_Closing({
        param($sender, $e)
        # Ustaw flagę, aby przerwać pętlę deinstalacji, jeśli okno jest zamykane
        $script:isCancelledFromUninstall = $true
    })

    $btnKill.Add_Click({
        if ($null -ne $script:uninstProc -and -not $script:uninstProc.HasExited) {
            if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz wymusić zamknięcie procesu deinstalatora?" -Title "Zabij proces" -Button "YesNo" -Image "Warning") -eq [System.Windows.MessageBoxResult]::Yes) {
                try {
                    Start-Process -FilePath "taskkill.exe" -ArgumentList "/PID $($script:uninstProc.Id) /T /F" -WindowStyle Hidden -Wait
                    Write-Log "Wymuszono zamknięcie procesu deinstalatora (drzewo procesów)." -Context "Użytkownik"
                } catch {
                    Show-ThemedMessageBox -Message "Błąd podczas zamykania procesu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
                }
            }
        }
    })

    $btnPauseUninstall.Add_Click({
        if ($script:isUninstallPaused) {
            $script:isUninstallPaused = $false
            $btnPauseUninstall.Content = "Pauza"
            $btnPauseUninstall.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFE6A100")
            Write-Log "Wznowiono deinstalację zbiorczą."
        } else {
            $script:isUninstallPaused = $true
            $btnPauseUninstall.Content = "Wznów"
            $btnPauseUninstall.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF107C10")
            Write-Log "Deinstalacja zbiorcza wstrzymana. Oczekiwanie na interakcję..."
        }
    })

    $btnUninstall.Add_Click({
        $selectedApps = @($script:uninstAllApps | Where-Object { $_.IsChecked })
        if ($selectedApps.Count -eq 0) { 
            Show-ThemedMessageBox -Message "Wybierz co najmniej jeden program z listy (zaznacz pole)." -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
            return 
        }

        $displayNames = $selectedApps | Select-Object -ExpandProperty DisplayName
        $names = $displayNames -join "`n- "
        
        if (Show-UninstallConfirmDialog -AppCount $selectedApps.Count -AppListText "- $names") {
            try {
                $btnUninstall.IsEnabled = $false
                $btnPauseUninstall.IsEnabled = $true
                $btnSelectAllApps.IsEnabled = $false
                $btnDeselectAllApps.IsEnabled = $false
                $btnKill.IsEnabled = $true
                $lblUninstallStatus.Visibility = [System.Windows.Visibility]::Visible
                $pbUninstall.Visibility = [System.Windows.Visibility]::Visible
                $pbUninstall.Maximum = $selectedApps.Count
                $pbUninstall.Value = 0
                $uninstWindow.Cursor = [System.Windows.Input.Cursors]::AppStarting
                $script:isCancelledFromUninstall = $false
                
                $currentAppIndex = 0
                foreach ($app in $selectedApps) {
                    while ($script:isUninstallPaused -and -not $script:isCancelledFromUninstall) {
                        Do-WpfEvents
                        Start-Sleep -Milliseconds 100
                    }
                    if ($script:isCancelledFromUninstall) { break }

                    $currentAppIndex++
                    $uninstWindow.Dispatcher.Invoke([Action]{ 
                        $pbUninstall.Value = $currentAppIndex 
                    }) | Out-Null

                    $cmd = $app.QuietUninstallString
                    $isOfficeClickToRun = $false

                    if ([string]::IsNullOrWhiteSpace($cmd)) {
                        $cmd = $app.UninstallString
                        if (-not [string]::IsNullOrWhiteSpace($cmd)) {
                            if ($cmd -match "(?i)msiexec") {
                                $cmd = ($cmd -replace "(?i)/I", "/X") + " /qn /norestart"
                            } elseif ($cmd -match 'OfficeClickToRun\.exe') {
                                $isOfficeClickToRun = $true
                            } else {
                                $cmd = "$cmd /S /quiet /silent /norestart"
                            }
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($cmd)) {
                        Write-Log "Pominięto $($app.DisplayName) - brak polecenia deinstalacji w rejestrze." -IsError
                        $uninstWindow.Dispatcher.Invoke([Action]{
                            $app.StatusText = "❌ BRAK POLECENIA"
                            $app.StatusColor = "#FFE6A100"
                            $app.IsChecked = $false
                            $lvApps.Items.Refresh()
                        }) | Out-Null
                        Do-WpfEvents
                        Start-Sleep -Milliseconds 800
                        continue
                    }

                    Write-Log "Uruchamianie deinstalatora dla $($app.DisplayName)..."
                    $exitCode = -1
                    try {
                        if ($isOfficeClickToRun) {
                            $c2rInfo = Resolve-OfficeClickToRunUninstallInfo -Cmd $cmd
                            $productId = $c2rInfo.ProductId
                            $c2rPath = $c2rInfo.ExePath

                            if ($productId -and $c2rPath -and (Test-Path -LiteralPath $c2rPath)) {
                                $xmlPath = Join-Path $env:TEMP "uninstall_office_config.xml"
                                $xmlContent = "<Configuration><Remove><Product ID=`"$productId`" /></Remove><Display Level=`"None`" AcceptEULA=`"True`" /></Configuration>"
                                $xmlContent | Set-Content -Path $xmlPath -Encoding UTF8
                                $arguments = "/configure `"$xmlPath`""
                                $finalCmdForLog = "`"$c2rPath`" $arguments"
                                $uninstWindow.Dispatcher.Invoke([Action]{ $lblUninstallStatus.Text = "Odinstalowywanie: $($app.DisplayName) | Cmd: $finalCmdForLog" }) | Out-Null
                                Write-Log "Uruchamianie deinstalatora Office: `"$c2rPath`" z argumentami: $arguments"
                                $exitCode = Start-UninstallProcessWithCancel -FilePath $c2rPath -ArgumentList $arguments -LogContext $app.DisplayName
                                $script:uninstProc = $null
                                if (Test-Path $xmlPath) { Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue }
                            } else {
                                Write-Log "Nie udało się wyodrębnić ProductID lub ścieżki OfficeClickToRun.exe dla $($app.DisplayName)" -IsError
                                $exitCode = -1
                            }
                        } else {
                            $uninstWindow.Dispatcher.Invoke([Action]{ $lblUninstallStatus.Text = "Odinstalowywanie: $($app.DisplayName) | Cmd: $cmd" }) | Out-Null
                            $exitCode = Start-UninstallProcessWithCancel -FilePath "cmd.exe" -ArgumentList "/c $cmd" -LogContext $app.DisplayName
                            $script:uninstProc = $null
                        }
                    } catch {
                        Write-Log "Błąd wykonania deinstalatora dla $($app.DisplayName): $_" -IsError
                        $exitCode = -1
                    }

                    $uninstWindow.Dispatcher.Invoke([Action]{
                        if ($exitCode -eq 0 -or $exitCode -eq 3010) {
                            $app.StatusText = "✔️ SUKCES"
                            $app.StatusColor = "#FF107C10"
                        } else {
                            $app.StatusText = "❌ BŁĄD (kod: $exitCode)"
                            $app.StatusColor = "#FFC50F1F"
                        }
                        $lvApps.Items.Refresh()
                    }) | Out-Null
                    
                    Do-WpfEvents
                    Start-Sleep -Milliseconds 800
                    
                    if ($exitCode -eq 0 -or $exitCode -eq 3010) {
                        $uninstWindow.Dispatcher.Invoke([Action]{
                            $lvApps.Items.Remove($app)
                            $script:uninstAllApps = @($script:uninstAllApps | Where-Object { $_ -ne $app })
                        }) | Out-Null
                    } else {
                        $uninstWindow.Dispatcher.Invoke([Action]{
                            $app.IsChecked = $false
                            $lvApps.Items.Refresh()
                        }) | Out-Null
                    }
                }
                
                if ($null -ne $notifyIcon) {
                    $notifyIcon.ShowBalloonTip(3000, "Deinstalacja zakończona", "Przetworzono wybrane programy.", [System.Windows.Forms.ToolTipIcon]::Info)
                }
                Show-CustomInfoDialog -Title "Zakończono" -Message "Przetwarzanie deinstalacji zakończone.`n`nPoprawnie odinstalowane aplikacje zostały automatycznie usunięte z listy.`nJeśli instalator zwrócił błąd, aplikacja pozostała na liście oznaczona krzyżykiem (❌)."
            } catch {
            Show-ThemedMessageBox -Message "Wystąpił błąd podczas uruchamiania deinstalatora:`n$($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            } finally {
                $script:uninstProc = $null
                $btnUninstall.IsEnabled = $true
                $btnPauseUninstall.IsEnabled = $false
                $script:isUninstallPaused = $false
                $btnPauseUninstall.Content = "Pauza"
                $btnPauseUninstall.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFE6A100")
                $btnSelectAllApps.IsEnabled = $true
                $btnDeselectAllApps.IsEnabled = $true
                $btnKill.IsEnabled = $false
                $lblUninstallStatus.Visibility = [System.Windows.Visibility]::Hidden
                $lblUninstallStatus.Text = ""
                $pbUninstall.Visibility = [System.Windows.Visibility]::Hidden
                $uninstWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
                $uninstWindow.Dispatcher.Invoke([Action]{ 
                    $count = @($script:uninstAllApps | Where-Object { $_.IsChecked }).Count
                    $lblSelectedCount.Text = "Wybrano: $count" 
                }) | Out-Null
            }
        }
    })

    $uninstWindow.Add_Loaded($LoadApps)
    $uninstWindow.ShowDialog() | Out-Null
}

function Show-ProgramEditDialog {
    param($IsNew, $ProgramName, $ProgramData)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Edytuj program" Width="480" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True"><Setter Property="Background" Value="{DynamicResource ThemeButton}"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Identyfikator (np. Chrome):" Margin="0,0,0,5"/>
        <TextBox Name="txtName" Grid.Row="1" Margin="0,0,0,15"/>
        <TextBlock Text="Nazwa pliku / Winget ID:" Grid.Row="2" Margin="0,0,0,5"/>
        <TextBox Name="txtFile" Grid.Row="3" Margin="0,0,0,15"/>
        <TextBlock Text="Argumenty cichej instalacji:" Grid.Row="4" Margin="0,0,0,5"/>
        <TextBox Name="txtArgs" Grid.Row="5" Margin="0,0,0,15"/>
        <CheckBox Name="chkForceUrl" Grid.Row="6" Content="Zawsze pobieraj z niestandardowego adresu URL" Margin="0,0,0,5" Foreground="{DynamicResource ThemeText}" FontSize="14" ToolTip="Nadpisuje globalne źródło instalacji dla tego konkretnego programu."/>
        <TextBox Name="txtUrl" Grid.Row="7" Margin="0,0,0,15" IsEnabled="False" ToolTip="Pełny bezpośredni adres URL do instalatora (np. https://.../plik.exe)"/>
        <CheckBox Name="chkEnabled" Grid.Row="8" Content="Domyślnie zaznaczone do instalacji" Margin="0,0,0,20" Foreground="{DynamicResource ThemeText}" FontSize="14"/>
        <StackPanel Grid.Row="9" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,10,0" Background="#FF0E639C" Foreground="White" IsDefault="True"/>
            <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $dlg.Title = if ($IsNew) { "Dodaj program" } else { "Edytuj program" }
    $txtName = $dlg.FindName("txtName")
    $txtFile = $dlg.FindName("txtFile")
    $txtArgs = $dlg.FindName("txtArgs")
    $chkForceUrl = $dlg.FindName("chkForceUrl")
    $txtUrl = $dlg.FindName("txtUrl")
    $chkEnabled = $dlg.FindName("chkEnabled")
    $btnSave = $dlg.FindName("btnSave")
    $btnCancel = $dlg.FindName("btnCancel")
    
    $txtName.Text = $ProgramName
    if (-not $IsNew) { $txtName.IsReadOnly = $true; $txtName.Opacity = 0.6 }
    if ($ProgramData) { $txtFile.Text = $ProgramData.FileName }
    if ($ProgramData) { $txtArgs.Text = $ProgramData.SilentArgs }
    if ($ProgramData) { $chkEnabled.IsChecked = [bool]$ProgramData.Enabled } else { $chkEnabled.IsChecked = $true }
    if ($ProgramData -and -not [string]::IsNullOrWhiteSpace($ProgramData.DownloadUrl)) {
        $chkForceUrl.IsChecked = $true
        $txtUrl.IsEnabled = $true
        $txtUrl.Text = $ProgramData.DownloadUrl
    }
    
    $chkForceUrl.Add_Click({
        $txtUrl.IsEnabled = $chkForceUrl.IsChecked -eq $true
    })
    
    $script:progEditResult = $null
    
    $btnSave.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
            Show-ThemedMessageBox -Message "Identyfikator nie może być pusty." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
            return
        }
        $script:progEditResult = @{
            Name = $txtName.Text.Trim()
            Data = [ordered]@{
                Enabled = $chkEnabled.IsChecked -eq $true
                FileName = $txtFile.Text.Trim()
                SilentArgs = $txtArgs.Text.Trim()
                DownloadUrl = if ($chkForceUrl.IsChecked -eq $true) { $txtUrl.Text.Trim() } else { "" }
            }
        }
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    
    $btnCancel.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    })
    
    if ($dlg.ShowDialog() -eq $true) {
        return $script:progEditResult
    }
    return $null
}

function Show-ProgramsManager {
    param($config)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Zarządzaj programami" Height="450" Width="480" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="130"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbPrograms" Grid.Column="0" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}" Margin="0,0,15,0" Padding="5" FontSize="14"/>
        <StackPanel Grid.Column="1">
            <Button Name="btnAdd" Content="Dodaj" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FF0E639C"/>
            <Button Name="btnEdit" Content="Edytuj" Height="32" Margin="0,0,0,10"/>
            <Button Name="btnClone" Content="Powiel" Height="32" Margin="0,0,0,10"/>
            <Button Name="btnRemove" Content="Usuń" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FFC50F1F"/>
            <Button Name="btnClose" Content="Zamknij" Height="32" Margin="0,20,0,0" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $lbPrograms = $dlg.FindName("lbPrograms")
    $btnAdd = $dlg.FindName("btnAdd")
    $btnEdit = $dlg.FindName("btnEdit")
    $btnClone = $dlg.FindName("btnClone")
    $btnRemove = $dlg.FindName("btnRemove")
    $btnClose = $dlg.FindName("btnClose")
    
    $RefreshList = {
        $lbPrograms.Items.Clear()
        if ($config.Programs) {
            foreach ($p in $config.Programs.PSObject.Properties.Name | Sort-Object) {
                [void]$lbPrograms.Items.Add($p)
            }
        }
    }
    & $RefreshList

    $btnAdd.Add_Click({
        $res = Show-ProgramEditDialog -IsNew $true -ProgramName "" -ProgramData $null
        if ($res) {
            if ($null -eq $config.Programs) {
                $config | Add-Member -NotePropertyName Programs -NotePropertyValue (New-Object PSObject) -Force
            }
            if ($config.Programs.PSObject.Properties.Name -contains $res.Name) {
                Show-ThemedMessageBox -Message "Program o tym identyfikatorze już istnieje." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
                return
            }
            Add-Member -InputObject $config.Programs -NotePropertyName $res.Name -NotePropertyValue $res.Data -Force
            & $RefreshList
        }
    })

    $btnEdit.Add_Click({
        if ($lbPrograms.SelectedItem) {
            $pName = $lbPrograms.SelectedItem
            $pData = $config.Programs.$pName
            $res = Show-ProgramEditDialog -IsNew $false -ProgramName $pName -ProgramData $pData
            if ($res) {
                if ($null -eq $config.Programs) {
                    $config | Add-Member -NotePropertyName Programs -NotePropertyValue (New-Object PSObject) -Force
                }
                $config.Programs.PSObject.Properties.Remove($pName)
                Add-Member -InputObject $config.Programs -NotePropertyName $res.Name -NotePropertyValue $res.Data -Force
                & $RefreshList
            }
        }
    })

    $btnClone.Add_Click({
        if ($lbPrograms.SelectedItem) {
            $pName = $lbPrograms.SelectedItem
            $pData = $config.Programs.$pName
            $res = Show-ProgramEditDialog -IsNew $true -ProgramName "$pName-Kopia" -ProgramData $pData
            if ($res) {
                if ($null -eq $config.Programs) {
                    $config | Add-Member -NotePropertyName Programs -NotePropertyValue (New-Object PSObject) -Force
                }
                if ($config.Programs.PSObject.Properties.Name -contains $res.Name) {
                Show-ThemedMessageBox -Message "Program o tym identyfikatorze już istnieje." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
                    return
                }
                Add-Member -InputObject $config.Programs -NotePropertyName $res.Name -NotePropertyValue $res.Data -Force
                & $RefreshList
            }
        }
    })

    $btnRemove.Add_Click({
        if ($lbPrograms.SelectedItem) {
            $pName = $lbPrograms.SelectedItem
            if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz usunąć program $pName?" -Title "Potwierdzenie" -Button "YesNo" -Image "Warning") -eq [System.Windows.MessageBoxResult]::Yes) {
                $config.Programs.PSObject.Properties.Remove($pName)
                & $RefreshList
            }
        }
    })

    $btnClose.Add_Click({ $dlg.Close() })

    $dlg.ShowDialog() | Out-Null
}

function Show-RegistryEditDialog {
    param($IsNew, $RegData)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Edytuj wpis rejestru" Width="450" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid TextElement.Foreground="{TemplateBinding Foreground}">
                            <ToggleButton x:Name="ToggleButton" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Foreground="{TemplateBinding Foreground}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                            <TextBlock Text="▼" Foreground="{TemplateBinding Foreground}" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="10"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="8,0,25,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="250" CornerRadius="4">
                                    <ScrollViewer SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" CornerRadius="2">
                            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Klucz główny i subklucz (np. Software\MójKlucz):" Margin="0,0,0,5"/>
        <Grid Grid.Row="1" Margin="0,0,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="85"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <ComboBox Name="cmbRoot" Grid.Column="0" Margin="0,0,10,0" Height="28" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}">
                <ComboBoxItem Content="HKLM:\"/>
                <ComboBoxItem Content="HKCU:\"/>
            </ComboBox>
            <TextBox Name="txtSubKey" Grid.Column="1"/>
        </Grid>
        
        <TextBlock Text="Nazwa (Name):" Grid.Row="2" Margin="0,0,0,5"/>
        <TextBox Name="txtName" Grid.Row="3" Margin="0,0,0,15"/>
        
        <TextBlock Text="Wartość (Value):" Grid.Row="4" Margin="0,0,0,5"/>
        <TextBox Name="txtValue" Grid.Row="5" Margin="0,0,0,15"/>
        
        <TextBlock Text="Typ danych (PropertyType):" Grid.Row="6" Margin="0,0,0,5"/>
        <ComboBox Name="cmbType" Grid.Row="7" Margin="0,0,0,20" Height="28" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}">
            <ComboBoxItem Content="String"/>
            <ComboBoxItem Content="DWord"/>
            <ComboBoxItem Content="QWord"/>
            <ComboBoxItem Content="ExpandString"/>
            <ComboBoxItem Content="Binary"/>
            <ComboBoxItem Content="MultiString"/>
        </ComboBox>
        
        <StackPanel Grid.Row="8" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,10,0" Background="#FF0E639C" Foreground="White" IsDefault="True"/>
            <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $dlg.Title = if ($IsNew) { "Dodaj wpis rejestru" } else { "Edytuj wpis rejestru" }
    $cmbRoot = $dlg.FindName("cmbRoot")
    $txtSubKey = $dlg.FindName("txtSubKey")
    $txtName = $dlg.FindName("txtName")
    $txtValue = $dlg.FindName("txtValue")
    $cmbType = $dlg.FindName("cmbType")
    $btnSave = $dlg.FindName("btnSave")
    $btnCancel = $dlg.FindName("btnCancel")
    
    if ($RegData) {
        $path = $RegData.Path
        if ($path -match "^(?i)HKLM:\\\\?(.*)") {
            $cmbRoot.SelectedIndex = 0
            $txtSubKey.Text = $matches[1]
        } elseif ($path -match "^(?i)HKCU:\\\\?(.*)") {
            $cmbRoot.SelectedIndex = 1
            $txtSubKey.Text = $matches[1]
        } else {
            $cmbRoot.SelectedIndex = 0
            $txtSubKey.Text = $path
        }
        $txtName.Text = $RegData.Name
        $txtValue.Text = $RegData.Value
        foreach ($item in $cmbType.Items) {
            if ($item.Content -eq $RegData.PropertyType) {
                $cmbType.SelectedItem = $item
                break
            }
        }
    } else {
        $cmbRoot.SelectedIndex = 0
        $cmbType.SelectedIndex = 0
    }
    
    $script:regEditResult = $null
    
    $btnSave.Add_Click({
        $subKey = $txtSubKey.Text.Trim() -replace '^[\\/]+', ''
        if ([string]::IsNullOrWhiteSpace($subKey) -or [string]::IsNullOrWhiteSpace($txtName.Text)) {
            Show-ThemedMessageBox -Message "Ścieżka i nazwa nie mogą być puste." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
            return
        }
        if ($subKey -match "^(?i)HK(LM|CU|CR|U|CC)") {
            Show-ThemedMessageBox -Message "Wpisz tylko ścieżkę podrzędną (np. Software\MójKlucz). Główny klucz wybierasz z listy." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
            return
        }
        
        $root = if ($cmbRoot.SelectedItem) { $cmbRoot.SelectedItem.Content } else { "HKLM:\" }
        $fullPath = "$root$subKey"
        
        $typeStr = if ($cmbType.SelectedItem) { $cmbType.SelectedItem.Content } else { "String" }
        $script:regEditResult = [ordered]@{
            Path = $fullPath
            Name = $txtName.Text.Trim()
            Value = $txtValue.Text.Trim()
            PropertyType = $typeStr
        }
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    
    $btnCancel.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    })
    
    if ($dlg.ShowDialog() -eq $true) {
        return $script:regEditResult
    }
    return $null
}

function Show-RegistryManager {
    param($config)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Zarządzaj kluczami rejestru" Height="450" Width="600" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="130"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbRegistry" Grid.Column="0" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}" Margin="0,0,15,0" Padding="5" FontSize="14"/>
        <StackPanel Grid.Column="1">
            <Button Name="btnAdd" Content="Dodaj" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FF0E639C"/>
            <Button Name="btnEdit" Content="Edytuj" Height="32" Margin="0,0,0,10"/>
            <Button Name="btnRemove" Content="Usuń" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FFC50F1F"/>
            <Button Name="btnClose" Content="Zamknij" Height="32" Margin="0,40,0,0" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $lbRegistry = $dlg.FindName("lbRegistry")
    $btnAdd = $dlg.FindName("btnAdd")
    $btnEdit = $dlg.FindName("btnEdit")
    $btnRemove = $dlg.FindName("btnRemove")
    $btnClose = $dlg.FindName("btnClose")
    
    if (-not $config.SystemSettings) {
        $config | Add-Member -NotePropertyName SystemSettings -NotePropertyValue (New-Object PSObject) -Force
    }
    if ($null -eq $config.SystemSettings.CustomRegistry) {
        $config.SystemSettings | Add-Member -NotePropertyName CustomRegistry -NotePropertyValue @() -Force
    }
    $regList = [System.Collections.ArrayList]@($config.SystemSettings.CustomRegistry)
    
    $RefreshList = {
        $lbRegistry.Items.Clear()
        foreach ($r in $regList) {
            [void]$lbRegistry.Items.Add("$($r.Path)\$($r.Name) = $($r.Value)")
        }
        $config.SystemSettings | Add-Member -NotePropertyName CustomRegistry -NotePropertyValue ($regList.ToArray()) -Force
    }
    & $RefreshList
    
    $btnAdd.Add_Click({
        $res = Show-RegistryEditDialog -IsNew $true -RegData $null
        if ($res) {
            [void]$regList.Add([PSCustomObject]$res)
            & $RefreshList
        }
    })
    
    $btnEdit.Add_Click({
        $idx = $lbRegistry.SelectedIndex
        if ($idx -ge 0) {
            $res = Show-RegistryEditDialog -IsNew $false -RegData $regList[$idx]
            if ($res) {
                $regList[$idx] = [PSCustomObject]$res
                & $RefreshList
            }
        }
    })
    
    $btnRemove.Add_Click({
        $idx = $lbRegistry.SelectedIndex
        if ($idx -ge 0) {
            if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz usunąć ten wpis rejestru?" -Title "Potwierdzenie" -Button "YesNo" -Image "Warning") -eq [System.Windows.MessageBoxResult]::Yes) {
                $regList.RemoveAt($idx)
                & $RefreshList
            }
        }
    })
    
    $btnClose.Add_Click({ $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
}

function Show-DefaultTasksEditor {
    param($config)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Domyślne zadania startowe" Height="500" Width="450" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Zaznacz opcje, które mają być domyślnie włączone przy starcie aplikacji:" TextWrapping="Wrap" Margin="0,0,0,15" FontSize="14"/>
        <Border Grid.Row="1" Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Padding="10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="spDefaults"/>
            </ScrollViewer>
        </Border>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
            <Button Name="btnSave" Content="Zapisz" Width="100" Height="35" Margin="0,0,10,0" Background="#FF0E639C" Foreground="White" IsDefault="True"/>
            <Button Name="btnCancel" Content="Anuluj" Width="100" Height="35" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $spDefaults = $dlg.FindName("spDefaults")
    $btnSave = $dlg.FindName("btnSave")
    $btnCancel = $dlg.FindName("btnCancel")
    
    if (-not $config.DefaultCheckboxes) {
        $config | Add-Member -NotePropertyName DefaultCheckboxes -NotePropertyValue (New-Object PSObject) -Force
    }
    
    $chkList = @{}
    foreach ($key in $checkboxOptions.Keys) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $checkboxOptions[$key].Text
        if ($null -ne $config.DefaultCheckboxes.$key) { $cb.IsChecked = [bool]$config.DefaultCheckboxes.$key } else { $cb.IsChecked = $checkboxOptions[$key].Enabled }
        $cb.Margin = "0,4,0,4"
        $cb.Foreground = $dlg.Resources["ThemeText"]
        $spDefaults.Children.Add($cb) | Out-Null
        $chkList[$key] = $cb
    }
    
    $btnSave.Add_Click({
        foreach ($key in $chkList.Keys) {
            Add-Member -InputObject $config.DefaultCheckboxes -NotePropertyName $key -NotePropertyValue ($chkList[$key].IsChecked -eq $true) -Force
        }
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    if ($dlg.ShowDialog() -eq $true) {
        foreach ($key in $chkList.Keys) {
            if ($null -ne $CheckboxControls[$key]) {
                $CheckboxControls[$key].IsChecked = $config.DefaultCheckboxes.$key
                $checkboxOptions[$key].Enabled = $config.DefaultCheckboxes.$key
            }
        }
    }
}

function Show-ScriptEditDialog {
    param($IsNew, $ScriptPath)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Edytuj skrypt" Width="420" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Ścieżka do skryptu (nazwa pliku, UNC lub URL):" Margin="0,0,0,5"/>
        <TextBox Name="txtPath" Grid.Row="1" Margin="0,0,0,20"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,10,0" Background="#FF0E639C" Foreground="White" IsDefault="True"/>
            <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $dlg.Title = if ($IsNew) { "Dodaj skrypt Post-Install" } else { "Edytuj skrypt Post-Install" }
    $txtPath = $dlg.FindName("txtPath")
    $btnSave = $dlg.FindName("btnSave")
    $btnCancel = $dlg.FindName("btnCancel")
    
    $txtPath.Text = $ScriptPath
    $script:scriptEditResult = $null
    
    $btnSave.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtPath.Text)) {
            Show-ThemedMessageBox -Message "Ścieżka skryptu nie może być pusta." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
            return
        }
        $script:scriptEditResult = $txtPath.Text.Trim()
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    if ($dlg.ShowDialog() -eq $true) { return $script:scriptEditResult }
    return $null
}

function Show-PostInstallScriptsManager {
    param($config)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Zarządzaj skryptami Post-Install" Height="400" Width="520" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="130"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbScripts" Grid.Column="0" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}" Margin="0,0,15,0" Padding="5" FontSize="14"/>
        <StackPanel Grid.Column="1">
            <Button Name="btnAdd" Content="Dodaj" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FF0E639C"/>
            <Button Name="btnEdit" Content="Edytuj" Height="32" Margin="0,0,0,10"/>
            <Button Name="btnRemove" Content="Usuń" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FFC50F1F"/>
            <Button Name="btnClose" Content="Zamknij" Height="32" Margin="0,40,0,0" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $lbScripts = $dlg.FindName("lbScripts")
    $btnAdd = $dlg.FindName("btnAdd")
    $btnEdit = $dlg.FindName("btnEdit")
    $btnRemove = $dlg.FindName("btnRemove")
    $btnClose = $dlg.FindName("btnClose")
    
    if ($null -eq $config.PostInstallScripts) {
        $config | Add-Member -NotePropertyName PostInstallScripts -NotePropertyValue @() -Force
    }
    $scriptList = [System.Collections.ArrayList]@($config.PostInstallScripts)
    
    $RefreshList = {
        $lbScripts.Items.Clear()
        foreach ($s in $scriptList) { [void]$lbScripts.Items.Add($s) }
        $config | Add-Member -NotePropertyName PostInstallScripts -NotePropertyValue ($scriptList.ToArray()) -Force
    }
    & $RefreshList
    
    $btnAdd.Add_Click({
        $res = Show-ScriptEditDialog -IsNew $true -ScriptPath ""
        if ($res) { [void]$scriptList.Add($res); & $RefreshList }
    })
    
    $btnEdit.Add_Click({
        $idx = $lbScripts.SelectedIndex
        if ($idx -ge 0) {
            $res = Show-ScriptEditDialog -IsNew $false -ScriptPath $scriptList[$idx]
            if ($res) { $scriptList[$idx] = $res; & $RefreshList }
        }
    })
    
    $btnRemove.Add_Click({
        $idx = $lbScripts.SelectedIndex
        if ($idx -ge 0) {
            if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz usunąć ten skrypt z listy?" -Title "Potwierdzenie" -Button "YesNo" -Image "Warning") -eq [System.Windows.MessageBoxResult]::Yes) {
                $scriptList.RemoveAt($idx)
                & $RefreshList
            }
        }
    })
    
    $btnClose.Add_Click({ $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
}

function Show-ProfileEditDialog {
    param($IsNew, $ProfileName, $ProfileApps, $config)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Edytuj profil wdrożeniowy" Width="450" Height="550" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="5,4"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Nazwa profilu (np. Księgowość):" Margin="0,0,0,5"/>
        <TextBox Name="txtName" Grid.Row="1" Margin="0,0,0,15" Height="28"/>
        <TextBlock Text="Wybierz aplikacje przypisane do profilu:" Grid.Row="2" Margin="0,0,0,5"/>
        <Border Grid.Row="3" Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,15">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="spApps"/>
            </ScrollViewer>
        </Border>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,10,0" Background="#FF0E639C" Foreground="White" IsDefault="True"/>
            <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $dlg.Title = if ($IsNew) { "Dodaj profil wdrożeniowy" } else { "Edytuj profil wdrożeniowy" }
    $txtName = $dlg.FindName("txtName")
    $spApps = $dlg.FindName("spApps")
    $btnSave = $dlg.FindName("btnSave")
    $btnCancel = $dlg.FindName("btnCancel")
    
    $txtName.Text = $ProfileName
    if (-not $IsNew) { $txtName.IsReadOnly = $true; $txtName.Opacity = 0.6 }
    
    $checkboxes = @{}
    if ($config.Programs) {
        foreach ($p in $config.Programs.PSObject.Properties.Name | Sort-Object) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $p
            if ($ProfileApps -and ($ProfileApps -contains $p)) {
                $cb.IsChecked = $true
            }
            $spApps.Children.Add($cb) | Out-Null
            $checkboxes[$p] = $cb
        }
    }
    
    $script:profileEditResult = $null
    
    $btnSave.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
            Show-ThemedMessageBox -Message "Nazwa profilu nie może być pusta." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
            return
        }
        $selectedApps = @()
        foreach ($key in $checkboxes.Keys) {
            if ($checkboxes[$key].IsChecked -eq $true) {
                $selectedApps += $key
            }
        }
        $script:profileEditResult = @{
            Name = $txtName.Text.Trim()
            Apps = $selectedApps
        }
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    if ($dlg.ShowDialog() -eq $true) { return $script:profileEditResult }
    return $null
}

function Show-ProfilesManager {
    param($config)
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Zarządzaj profilami wdrożeniowymi" Height="450" Width="480" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="130"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbProfiles" Grid.Column="0" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}" Margin="0,0,15,0" Padding="5" FontSize="14"/>
        <StackPanel Grid.Column="1">
            <Button Name="btnAdd" Content="Dodaj" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FF0E639C"/>
            <Button Name="btnEdit" Content="Edytuj" Height="32" Margin="0,0,0,10"/>
            <Button Name="btnRemove" Content="Usuń" Height="32" Margin="0,0,0,10" Foreground="White" Background="#FFC50F1F"/>
            <Button Name="btnClose" Content="Zamknij" Height="32" Margin="0,20,0,0" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $lbProfiles = $dlg.FindName("lbProfiles")
    $btnAdd = $dlg.FindName("btnAdd")
    $btnEdit = $dlg.FindName("btnEdit")
    $btnRemove = $dlg.FindName("btnRemove")
    $btnClose = $dlg.FindName("btnClose")
    
    if ($null -eq $config.Profiles) {
        $config | Add-Member -NotePropertyName Profiles -NotePropertyValue (New-Object PSObject) -Force
    }
    
    $RefreshList = {
        $lbProfiles.Items.Clear()
        if ($config.Profiles) {
            foreach ($p in $config.Profiles.PSObject.Properties.Name | Sort-Object) {
                [void]$lbProfiles.Items.Add($p)
            }
        }
    }
    & $RefreshList
    
    $btnAdd.Add_Click({
        $res = Show-ProfileEditDialog -IsNew $true -ProfileName "" -ProfileApps @() -config $config
        if ($res) {
            if ($null -eq $config.Profiles) {
                $config | Add-Member -NotePropertyName Profiles -NotePropertyValue (New-Object PSObject) -Force
            }
            if ($config.Profiles.PSObject.Properties.Name -contains $res.Name) {
                Show-ThemedMessageBox -Message "Profil o tej nazwie już istnieje." -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
                return
            }
            Add-Member -InputObject $config.Profiles -NotePropertyName $res.Name -NotePropertyValue $res.Apps -Force
            & $RefreshList
        }
    })
    
    $btnEdit.Add_Click({
        if ($lbProfiles.SelectedItem) {
            $pName = $lbProfiles.SelectedItem
            $pApps = $config.Profiles.$pName
            $res = Show-ProfileEditDialog -IsNew $false -ProfileName $pName -ProfileApps $pApps -config $config
            if ($res) {
                if ($null -eq $config.Profiles) {
                    $config | Add-Member -NotePropertyName Profiles -NotePropertyValue (New-Object PSObject) -Force
                }
                $config.Profiles.PSObject.Properties.Remove($pName)
                Add-Member -InputObject $config.Profiles -NotePropertyName $res.Name -NotePropertyValue $res.Apps -Force
                & $RefreshList
            }
        }
    })
    
    $btnRemove.Add_Click({
        if ($lbProfiles.SelectedItem) {
            $pName = $lbProfiles.SelectedItem
            if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz usunąć profil '$pName'?" -Title "Potwierdzenie" -Button "YesNo" -Image "Warning") -eq [System.Windows.MessageBoxResult]::Yes) {
                $config.Profiles.PSObject.Properties.Remove($pName)
                & $RefreshList
            }
        }
    })
    
    $btnClose.Add_Click({ $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
}

function Show-PinPrompt {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Wymagana autoryzacja" Width="300" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,5"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Wprowadź PIN, aby edytować ustawienia:" Margin="0,0,0,10" FontSize="13" FontWeight="SemiBold"/>
        <PasswordBox Name="txtPin" Grid.Row="1" Margin="0,0,0,20"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnOk" Content="Odblokuj" Width="90" Height="32" Margin="0,0,10,0" Background="#FF107C10" Foreground="White" IsDefault="True"/>
            <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $txtPin = $dlg.FindName("txtPin")
    $btnOk = $dlg.FindName("btnOk")
    $btnCancel = $dlg.FindName("btnCancel")
    
    $script:pinResult = $false
    
    $btnOk.Add_Click({
        if ($txtPin.Password -eq "2137") {
            $script:pinResult = $true
            Write-Log "[Autoryzacja] Poprawny PIN. Odblokowano dostęp do ustawień konfiguracyjnych."
            $dlg.DialogResult = $true
            $dlg.Close()
        } else {
            Write-Log "[Autoryzacja] Błędny PIN przy próbie wejścia w ustawienia!" -IsError
            Show-ThemedMessageBox -Message "Nieprawidłowy PIN." -Title "Błąd autoryzacji" -Button "OK" -Image "Error" | Out-Null
            $txtPin.Clear()
        }
    })
    
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
    return $script:pinResult
}

function Show-ConfigEditor {
    try { $config = Get-Config } catch { Write-Log $_ -Color Red; return }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ustawienia – źródła, domena, konto lokalne" Height="650" Width="680" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.8"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Margin" Value="0,0,0,5"/>
            <Setter Property="VerticalAlignment" Value="Bottom"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid TextElement.Foreground="{TemplateBinding Foreground}">
                            <ToggleButton x:Name="ToggleButton" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Foreground="{TemplateBinding Foreground}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                            <TextBlock Text="▼" Foreground="{TemplateBinding Foreground}" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="10"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="8,0,25,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="250" CornerRadius="4">
                                    <ScrollViewer SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" CornerRadius="2">
                            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeButton}"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,0,0,10" Padding="0,0,10,0">
            <StackPanel>
                <Border Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Margin="0,0,0,12">
                    <StackPanel>
                        <Button Name="btnToggleSrc" HorizontalContentAlignment="Stretch" Height="42" Background="#FF0078D7" Foreground="White">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="📡 Źródła instalacji i uwierzytelnianie sieciowe" FontWeight="Bold" FontSize="15"/>
                                <TextBlock Name="chevronSrc" Grid.Column="1" Text="▾" FontSize="14"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelSrc" Margin="14,14,14,16" Visibility="Visible">
                            <TextBlock Text="Domyślne źródło (DefaultInstallSource):"/>
                            <ComboBox Name="cmbSrc" Height="28" Margin="0,0,0,10" Padding="5,2" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}"/>
                            <TextBlock Text="Ścieżka sieciowa [network] (UNC):"/>
                            <TextBox Name="txtNet" Margin="0,0,0,10"/>
                            <TextBlock Text="Ścieżka sieciowa [web] (URL):"/>
                            <Grid Margin="0,0,0,15">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Name="txtWeb" Grid.Column="0" Margin="0,0,10,0"/>
                                <Button Name="btnTestWeb" Content="Testuj" Grid.Column="1" Width="70" Height="28" Foreground="White" Background="#FF0078D7"/>
                            </Grid>
                            <TextBlock Text="Niestandardowe dane (CustomWebDataLocation URL):"/>
                            <TextBox Name="txtCwd" Margin="0,0,0,15"/>
                            <TextBlock Text="Konto logowania po HTTP/HTTPS (WebAuth Username):"/>
                            <TextBox Name="txtWebUser" Margin="0,0,0,10"/>
                            <TextBlock Text="Hasło do konta po HTTP/HTTPS (WebAuth Password):"/>
                            <TextBox Name="txtWebPass"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Margin="0,0,0,12">
                    <StackPanel>
                        <Button Name="btnToggleDom" HorizontalContentAlignment="Stretch" Height="42" Background="#FF2E5AAC" Foreground="White">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="🏢 Domena i konto lokalne" FontWeight="Bold" FontSize="15"/>
                                <TextBlock Name="chevronDom" Grid.Column="1" Text="▾" FontSize="14"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelDom" Margin="14,14,14,16" Visibility="Visible">
                            <TextBlock Text="Nazwa domeny (DomainName):"/>
                            <TextBox Name="txtDom" Margin="0,0,0,10"/>
                            <TextBlock Text="Konto uprawnione do podłączenia (Username):"/>
                            <TextBox Name="txtDomUser" Margin="0,0,0,10"/>
                            <TextBlock Text="Nazwa domyślnego konta lokalnego (LocalAdmin Username):"/>
                            <TextBox Name="txtLoc"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Margin="0,0,0,12">
                    <StackPanel>
                        <Button Name="btnToggleTv" HorizontalContentAlignment="Stretch" Height="42" Background="#FF5C2D91" Foreground="White">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="📺 TeamViewer" FontWeight="Bold" FontSize="15"/>
                                <TextBlock Name="chevronTv" Grid.Column="1" Text="▸" FontSize="14"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelTv" Margin="14,14,14,16" Visibility="Collapsed">
                            <TextBlock Text="Nazwa pliku / Winget ID:"/>
                            <TextBox Name="txtTvFile" Margin="0,0,0,10"/>
                            <TextBlock Text="Argumenty instalacji:"/>
                            <TextBox Name="txtTvArgs"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Margin="0,0,0,12">
                    <StackPanel>
                        <Button Name="btnToggleAv" HorizontalContentAlignment="Stretch" Height="42" Background="#FF0E7490" Foreground="White">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="🛡️ Antywirus" FontWeight="Bold" FontSize="15"/>
                                <TextBlock Name="chevronAv" Grid.Column="1" Text="▸" FontSize="14"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelAv" Margin="14,14,14,16" Visibility="Collapsed">
                            <TextBlock Text="Nazwa pliku / Winget ID:"/>
                            <TextBox Name="txtAvFile" Margin="0,0,0,10"/>
                            <TextBlock Text="Domyślne źródło instalacji:"/>
                            <ComboBox Name="cmbAvSrc" Height="28" Margin="0,0,0,10" Padding="5,2" Background="{DynamicResource ThemeTextBoxBg}" Foreground="{DynamicResource ThemeText}" BorderBrush="{DynamicResource ThemeBorder}"/>
                            <TextBlock Text="Ścieżka sieciowa [network] (UNC):"/>
                            <TextBox Name="txtAvNet" Margin="0,0,0,10"/>
                            <TextBlock Text="Ścieżka sieciowa [web] (URL):"/>
                            <TextBox Name="txtAvWeb" Margin="0,0,0,10"/>
                            <TextBlock Text="Użytkownik WebAuth:"/>
                            <TextBox Name="txtAvUser" Margin="0,0,0,10"/>
                            <TextBlock Text="Hasło WebAuth:"/>
                            <TextBox Name="txtAvPass"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Margin="0,0,0,12">
                    <StackPanel>
                        <Button Name="btnToggleWifi" HorizontalContentAlignment="Stretch" Height="42" Background="#FF3D5A80" Foreground="White">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="📶 Profil Wi-Fi" FontWeight="Bold" FontSize="15"/>
                                <TextBlock Name="chevronWifi" Grid.Column="1" Text="▸" FontSize="14"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelWifi" Margin="14,14,14,16" Visibility="Collapsed">
                            <TextBlock Text="Nazwy plików (po przecinku, jeśli kilka):"/>
                            <TextBox Name="txtWifiFile"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Padding="14" Margin="0,4,0,12">
                    <StackPanel>
                        <TextBlock Text="🛠️ Zarządzanie" FontWeight="Bold" FontSize="14" Foreground="#FF4A9EFF" Margin="0,0,0,10"/>
                        <UniformGrid Columns="2" Rows="3">
                            <Button Name="btnManageApps" Content="📦 Programy" Height="38" Margin="4" Foreground="White" Background="#FF0078D7" FontWeight="SemiBold"/>
                            <Button Name="btnManageProfiles" Content="🗂️ Profile wdrożeniowe" Height="38" Margin="4" Foreground="White" Background="#FF2E5AAC" FontWeight="SemiBold"/>
                            <Button Name="btnManageRegistry" Content="🧩 Rejestr niestandardowy" Height="38" Margin="4" Foreground="White" Background="#FF5C2D91" FontWeight="SemiBold"/>
                            <Button Name="btnManageDefaults" Content="☑️ Domyślne zadania" Height="38" Margin="4" Foreground="White" Background="#FF0E7490" FontWeight="SemiBold"/>
                            <Button Name="btnManageScripts" Content="📜 Skrypty Post-Install" Height="38" Margin="4" Foreground="White" Background="#FF3D5A80" FontWeight="SemiBold"/>
                            <Button Name="btnCheckUpdate" Content="🔄 Aktualizacje narzędzia" Height="38" Margin="4" Foreground="White" Background="#FF7A3E9D" FontWeight="SemiBold"/>
                        </UniformGrid>
                    </StackPanel>
                </Border>

                <Border Background="{DynamicResource ThemeTextBoxBg}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="6" Padding="14" Margin="0,0,0,10">
                    <StackPanel>
                        <TextBlock Text="💾 Kopia zapasowa konfiguracji" FontWeight="Bold" FontSize="14" Foreground="#FF4A9EFF" Margin="0,0,0,10"/>
                        <StackPanel Orientation="Horizontal">
                            <Button Name="btnExportConfig" Content="⬆️ Eksportuj..." Height="36" Width="150" Foreground="White" Background="#FF0E639C" FontWeight="SemiBold" Margin="0,0,10,0"/>
                            <Button Name="btnImportConfig" Content="⬇️ Importuj..." Height="36" Width="150" Foreground="White" Background="#FF4A5568" FontWeight="SemiBold"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="1" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Margin="-20,0,-20,0" Padding="20,15,20,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnSave" Content="Zapisz" Width="110" Height="35" Margin="0,0,15,0" Foreground="White" Background="#FF0E639C" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="110" Height="35" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Apply-ThemeToWindow $dlg
    
    $cmbSrc = $dlg.FindName("cmbSrc")
    $txtNet = $dlg.FindName("txtNet")
    $txtWeb = $dlg.FindName("txtWeb")
    $btnTestWeb = $dlg.FindName("btnTestWeb")
    $txtCwd = $dlg.FindName("txtCwd")
    $txtDom = $dlg.FindName("txtDom")
    $txtDomUser = $dlg.FindName("txtDomUser")
    $txtLoc = $dlg.FindName("txtLoc")
    $txtWebUser = $dlg.FindName("txtWebUser")
    $txtWebPass = $dlg.FindName("txtWebPass")
    $txtTvFile = $dlg.FindName("txtTvFile")
    $txtTvArgs = $dlg.FindName("txtTvArgs")
    $txtAvFile = $dlg.FindName("txtAvFile")
    $cmbAvSrc = $dlg.FindName("cmbAvSrc")
    $txtAvNet = $dlg.FindName("txtAvNet")
    $txtAvWeb = $dlg.FindName("txtAvWeb")
    $txtAvUser = $dlg.FindName("txtAvUser")
    $txtAvPass = $dlg.FindName("txtAvPass")
    $txtWifiFile = $dlg.FindName("txtWifiFile")
    $btnManageApps = $dlg.FindName("btnManageApps")
    $btnManageProfiles = $dlg.FindName("btnManageProfiles")
    $btnManageRegistry = $dlg.FindName("btnManageRegistry")
    $btnManageDefaults = $dlg.FindName("btnManageDefaults")
    $btnManageScripts = $dlg.FindName("btnManageScripts")
    $btnCheckUpdate = $dlg.FindName("btnCheckUpdate")
    $btnExportConfig = $dlg.FindName("btnExportConfig")
    $btnImportConfig = $dlg.FindName("btnImportConfig")
    $btnSave = $dlg.FindName("btnSave")
    $btnCancel = $dlg.FindName("btnCancel")

    # Zwijanie/rozwijanie sekcji - zwykłe przyciski (sprawdzony, już wszędzie indziej działający
    # wzorzec) zamiast natywnego Expandera, którego nagłówek w tym oknie nie łapał poprawnie
    # motywu (mały, czarny tekst w trybie ciemnym).
    foreach ($section in @(
        @{ Button = "btnToggleSrc"; Panel = "panelSrc"; Chevron = "chevronSrc" }
        @{ Button = "btnToggleDom"; Panel = "panelDom"; Chevron = "chevronDom" }
        @{ Button = "btnToggleTv"; Panel = "panelTv"; Chevron = "chevronTv" }
        @{ Button = "btnToggleAv"; Panel = "panelAv"; Chevron = "chevronAv" }
        @{ Button = "btnToggleWifi"; Panel = "panelWifi"; Chevron = "chevronWifi" }
    )) {
        $btn = $dlg.FindName($section.Button)
        $panel = $dlg.FindName($section.Panel)
        $chevron = $dlg.FindName($section.Chevron)
        $btn.Add_Click({
            if ($panel.Visibility -eq [System.Windows.Visibility]::Visible) {
                $panel.Visibility = [System.Windows.Visibility]::Collapsed
                $chevron.Text = "▸"
            } else {
                $panel.Visibility = [System.Windows.Visibility]::Visible
                $chevron.Text = "▾"
            }
        }.GetNewClosure())
    }

    [void]$cmbAvSrc.Items.Add("network")
    [void]$cmbAvSrc.Items.Add("web")
    [void]$cmbAvSrc.Items.Add("winget")

    [void]$cmbSrc.Items.Add("network")
    [void]$cmbSrc.Items.Add("web")
    [void]$cmbSrc.Items.Add("winget")
    
    $UpdateUIFields = {
        param($cfg)
        $src = [string]$cfg.DefaultInstallSource
        if ($cmbSrc.Items -contains $src) { $cmbSrc.SelectedItem = $src } else { $cmbSrc.SelectedItem = 'network' }
        
        $txtNet.Text = [string]$cfg.InstallSourcePaths.network
        $txtWeb.Text = [string]$cfg.InstallSourcePaths.web
        $txtCwd.Text = [string]$cfg.CustomWebDataLocation.URL
        $txtDom.Text = [string]$cfg.DomainJoin.DomainName
        $txtDomUser.Text = [string]$cfg.DomainJoin.Username
        $txtLoc.Text = [string]$cfg.LocalAdmin.Username
        $txtWebUser.Text = [string]$cfg.WebAuth.Username
        $txtWebPass.Text = [string]$cfg.WebAuth.Password
        $txtTvFile.Text = [string]$cfg.TeamViewer.FileName
        $txtTvArgs.Text = [string]$cfg.TeamViewer.Arguments
        $txtAvFile.Text = [string]$cfg.AntyVirus.FileName
        $avSrc = [string]$cfg.AntyVirus.DefaultInstallSource
        if ($cmbAvSrc.Items -contains $avSrc) { $cmbAvSrc.SelectedItem = $avSrc } else { $cmbAvSrc.SelectedItem = 'network' }
        $txtAvNet.Text = [string]$cfg.AntyVirus.InstallSourcePaths.network
        $txtAvWeb.Text = [string]$cfg.AntyVirus.InstallSourcePaths.web
        $txtAvUser.Text = [string]$cfg.AntyVirus.Credentials.Username
        $txtAvPass.Text = [string]$cfg.AntyVirus.Credentials.Password
        if ($cfg.WiFiProfile.FileName -is [array]) {
            $txtWifiFile.Text = $cfg.WiFiProfile.FileName -join ","
        } else {
            $txtWifiFile.Text = [string]$cfg.WiFiProfile.FileName
        }
    }
    
    & $UpdateUIFields $config
    
    $src = [string]$config.DefaultInstallSource
    if ($cmbSrc.Items -contains $src) {
        $cmbSrc.SelectedItem = $src
    }
    else {
        $cmbSrc.SelectedItem = 'network'
    }
    
    $btnManageApps.Add_Click({ Show-ProgramsManager -config $config })
    $btnManageProfiles.Add_Click({ Show-ProfilesManager -config $config })
    $btnManageRegistry.Add_Click({ Show-RegistryManager -config $config })
    $btnManageDefaults.Add_Click({ Show-DefaultTasksEditor -config $config })
    $btnManageScripts.Add_Click({ Show-PostInstallScriptsManager -config $config })
    $btnCheckUpdate.Add_Click({ Test-ForAppUpdate })

    $btnExportConfig.Add_Click({
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Pliki JSON (*.json)|*.json|Wszystkie pliki (*.*)|*.*"
        $sfd.FileName = "config_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                $config | ConvertTo-Json -Depth 10 | Set-Content -Path $sfd.FileName -Encoding UTF8
                Show-ThemedMessageBox -Message "Eksport zakończony pomyślnie." -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
            } catch {
                Show-ThemedMessageBox -Message "Błąd eksportu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $btnImportConfig.Add_Click({
        $ofd = New-Object Microsoft.Win32.OpenFileDialog
        $ofd.Filter = "Pliki JSON (*.json)|*.json|Wszystkie pliki (*.*)|*.*"
        if ($ofd.ShowDialog() -eq $true) {
            try {
                $importedConfig = Get-Content $ofd.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($null -ne $importedConfig) {
                    foreach ($prop in $importedConfig.PSObject.Properties) {
                        $config | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
                    }
                    & $UpdateUIFields $config
                    Show-ThemedMessageBox -Message "Konfiguracja została zaimportowana. Kliknij 'Zapisz', aby ją trwale zachować w aplikacji." -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
                }
            } catch {
                Show-ThemedMessageBox -Message "Błąd importu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
            }
        }
    })

    $btnTestWeb.Add_Click({
        $url = $txtWeb.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($url)) {
            Show-ThemedMessageBox -Message "Proszę wpisać adres URL do przetestowania." -Title "Informacja" -Button "OK" -Image "Information" | Out-Null
            return
        }
        if (-not (Test-UrlValid -Url $url)) {
            Show-ThemedMessageBox -Message "Niepoprawny format adresu URL. Pamiętaj o dodaniu http:// lub https://" -Title "Ostrzeżenie" -Button "OK" -Image "Warning" | Out-Null
            return
        }
        $dlg.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
            $statusMsg = if ($null -ne $response.StatusCode) { "$($response.StatusCode) $($response.StatusDescription)" } else { "OK" }
            Write-Log "[Ustawienia] Test połączenia z URL '$url' zakończony sukcesem: $statusMsg"
            Show-ThemedMessageBox -Message "Host odpowiada poprawnie!`n`nKod statusu: $statusMsg" -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
        }
        catch {
            Write-Log "[Ustawienia] Test połączenia z URL '$url' zakończony błędem: $($_.Exception.Message)" -IsError
            Show-ThemedMessageBox -Message "Host nie odpowiada lub wystąpił błąd komunikacji:`n`n$($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
        }
        finally {
            $dlg.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })
    $btnCancel.Add_Click({ $dlg.Close() })

    $btnSave.Add_Click({
        try {
            $src = [string]$cmbSrc.SelectedItem
            if ([string]::IsNullOrWhiteSpace($src) -or ($src -notin @('network', 'web', 'winget'))) {
                Show-ThemedMessageBox -Message "Wybierz poprawne źródło (network/web/winget)." -Title "Ostrzeżenie" -Button "OK" -Image "Warning" | Out-Null
                return
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
                Show-ThemedMessageBox -Message "Ścieżka network musi być w formacie UNC (\\server\share\)." -Title "Ostrzeżenie" -Button "OK" -Image "Warning" | Out-Null
                return
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

            if (-not $config.WebAuth) { $config | Add-Member -NotePropertyName WebAuth -NotePropertyValue (@{}) -Force }
            $config.WebAuth.Username = $txtWebUser.Text.Trim()
            $config.WebAuth.Password = $txtWebPass.Text.Trim()

            if (-not $config.TeamViewer) { $config | Add-Member -NotePropertyName TeamViewer -NotePropertyValue (@{}) -Force }
            $config.TeamViewer.FileName = $txtTvFile.Text.Trim()
            $config.TeamViewer.Arguments = $txtTvArgs.Text.Trim()

            if (-not $config.AntyVirus) { $config | Add-Member -NotePropertyName AntyVirus -NotePropertyValue (@{}) -Force }
            $config.AntyVirus.FileName = $txtAvFile.Text.Trim()
            $avSrcVal = [string]$cmbAvSrc.SelectedItem
            if ([string]::IsNullOrWhiteSpace($avSrcVal)) { $avSrcVal = "network" }
            $config.AntyVirus.DefaultInstallSource = $avSrcVal

            if (-not $config.AntyVirus.InstallSourcePaths) { $config.AntyVirus | Add-Member -NotePropertyName InstallSourcePaths -NotePropertyValue (@{}) -Force }
            $config.AntyVirus.InstallSourcePaths.network = $txtAvNet.Text.Trim()
            $config.AntyVirus.InstallSourcePaths.web = $txtAvWeb.Text.Trim()

            if (-not $config.AntyVirus.Credentials) { $config.AntyVirus | Add-Member -NotePropertyName Credentials -NotePropertyValue (@{}) -Force }
            $config.AntyVirus.Credentials.Username = $txtAvUser.Text.Trim()
            $config.AntyVirus.Credentials.Password = $txtAvPass.Text.Trim()

            if (-not $config.WiFiProfile) { $config | Add-Member -NotePropertyName WiFiProfile -NotePropertyValue (@{}) -Force }
            $wifiStr = $txtWifiFile.Text.Trim()
            if ($wifiStr -match ",") {
                $config.WiFiProfile.FileName = @($wifiStr -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
            } else { $config.WiFiProfile.FileName = $wifiStr }

            Save-Config $config
            Show-ThemedMessageBox -Message "Zapisano konfigurację." -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
            Get-AppSelection
            $dlg.Close()
        }
        catch {
            Show-ThemedMessageBox -Message "Błąd zapisu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
        }
    })

    $dlg.ShowDialog() | Out-Null
}

Ensure-Configuration

[xml]$mainXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart Tool for Deployment | Short: STD" Height="780" Width="900" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Window.Resources>
        <SolidColorBrush x:Key="ThemeBackground" Color="#FF202225"/>
        <SolidColorBrush x:Key="ThemePanel" Color="#FF2F3136"/>
        <SolidColorBrush x:Key="ThemeText" Color="#FFDCDDDE"/>
        <SolidColorBrush x:Key="ThemeButton" Color="#FF4F545C"/>
        <SolidColorBrush x:Key="ThemeButtonText" Color="White"/>
        <SolidColorBrush x:Key="ThemeBorder" Color="#FF4F545C"/>
        <SolidColorBrush x:Key="ThemeTextBoxBg" Color="#FF1E1E1E"/>
        <SolidColorBrush x:Key="AccentColor" Color="#FF0078D7"/>
        <SolidColorBrush x:Key="StartColor" Color="#FF107C10"/>
        
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Opacity" Value="0.8"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.5"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid TextElement.Foreground="{TemplateBinding Foreground}">
                            <ToggleButton x:Name="ToggleButton" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Foreground="{TemplateBinding Foreground}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                            <TextBlock Text="▼" Foreground="{TemplateBinding Foreground}" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="10"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="8,0,25,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="250" CornerRadius="4">
                                    <ScrollViewer SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True"><Setter Property="Background" Value="{DynamicResource ThemeButton}"/></Trigger>
            </Style.Triggers>
        </Style>
        
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="5,8"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Grid Grid.Row="0" Margin="0,0,0,20">
            <TextBlock Text="Smart Tool for Deployment" FontSize="26" FontWeight="Bold"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,150,0">
                <TextBlock Name="txtStopwatch" Text="⏱ 00:00:00" VerticalAlignment="Center" Foreground="{DynamicResource ThemeText}" FontWeight="Bold" FontSize="16" Margin="0,0,20,0" Visibility="Hidden"/>
                <Ellipse Name="shpNetworkStatus" Width="12" Height="12" Fill="#FFE6A100" Margin="0,0,8,0">
                    <Ellipse.Triggers>
                        <EventTrigger RoutedEvent="FrameworkElement.Loaded">
                            <BeginStoryboard>
                                <Storyboard RepeatBehavior="Forever">
                                    <DoubleAnimation Storyboard.TargetProperty="Opacity" From="1.0" To="0.3" Duration="0:0:1" AutoReverse="True"/>
                                </Storyboard>
                            </BeginStoryboard>
                        </EventTrigger>
                    </Ellipse.Triggers>
                </Ellipse>
                <TextBlock Name="txtNetworkStatus" Text="Sprawdzanie sieci..." VerticalAlignment="Center" Foreground="{DynamicResource ThemeText}" FontWeight="SemiBold" FontSize="14"/>
            </StackPanel>
            <ToggleButton Name="btnThemeToggle" Content="☀️ Jasny Motyw" HorizontalAlignment="Right" VerticalAlignment="Center" Background="{DynamicResource ThemeButton}" Foreground="{DynamicResource ThemeButtonText}" Padding="10,5" BorderThickness="0" Cursor="Hand"/>
        </Grid>
        
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="5*"/>
                <ColumnDefinition Width="4*"/>
            </Grid.ColumnDefinitions>
            
            <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="8" Padding="15" Margin="0,0,15,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                        <Button Name="btnSelectAll" Content="Zaznacz wszystko" Margin="0,0,10,0"/>
                        <Button Name="btnDeselectAll" Content="Odznacz wszystko" Margin="0,0,10,0"/>
                        <Button Name="btnInvertSelection" Content="Odwróć zaznaczenie"/>
                    </StackPanel>
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel Name="spCheckboxes"/>
                    </ScrollViewer>
                </Grid>
            </Border>
            
            <Grid Grid.Column="1">
                
                <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="8" Padding="15">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,0,5,0">
                        <StackPanel>
                        <TextBlock Text="Profil wdrożenia (Rola):" Margin="0,0,0,5" Foreground="{DynamicResource ThemeText}" FontWeight="SemiBold"/>
                        <ComboBox Name="cmbProfiles" Height="32" Margin="0,0,0,15" Padding="5,4"/>
                        <Button Name="btnChooseApps" Content="Wybierz aplikacje" Margin="0,0,0,10" Height="38" Foreground="White" Background="{DynamicResource AccentColor}" FontWeight="SemiBold"/>
                        <Button Name="btnSettings" Content="Ustawienia..." Margin="0,0,0,10" Height="35"/>
                        <Button Name="btnEditConfig" Content="Edytuj config.json w Notatniku" Margin="0,0,0,10" Height="35"/>
                        <Button Name="btnReloadConfig" Content="Przeładuj config.json (zaktualizuj GUI)" Margin="0,0,0,10" Height="35"/>
                        <Button Name="btnLogs" Content="Przeglądaj / Zapisz logi" Height="35" Margin="0,0,0,15"/>
                        <TextBlock Text="Szybkie narzędzia:" Margin="0,0,0,5" Foreground="{DynamicResource ThemeText}" FontWeight="SemiBold"/>
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Button Name="btnSysProps" Content="SysProperties" Grid.Row="0" Grid.Column="0" Margin="0,0,5,5" Height="28" ToolTip="Zaawansowane ustawienia systemu (Właściwości systemu)"/>
                            <Button Name="btnCompMgmt" Content="Zarządzanie" Grid.Row="0" Grid.Column="1" Margin="5,0,0,5" Height="28" ToolTip="Zarządzanie komputerem (compmgmt.msc)"/>
                            <Button Name="btnRegEdit" Content="RegEdit" Grid.Row="1" Grid.Column="0" Margin="0,0,5,5" Height="28" ToolTip="Edytor rejestru (regedit)"/>
                            <Button Name="btnPrinters" Content="Drukarki" Grid.Row="1" Grid.Column="1" Margin="5,0,0,5" Height="28" ToolTip="Klasyczny widok urządzeń i drukarek"/>
                            <Button Name="btnSysInfo" Content="Informacje o systemie" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,0,0,5" Height="28" ToolTip="Podstawowe informacje o sprzęcie i systemie" Background="#FF0078D7" Foreground="White" FontWeight="SemiBold"/>
                            <Button Name="btnUninstaller" Content="Odinstaluj programy" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,5,0,0" Height="30" Background="#FFC50F1F" Foreground="White" FontWeight="SemiBold" ToolTip="Moduł do wymuszania cichej deinstalacji oprogramowania"/>
                        </Grid>
                        </StackPanel>
                    </ScrollViewer>
                </Border>
            </Grid>
        </Grid>
        
        <StackPanel Grid.Row="2" Margin="0,20,0,15">
            <TextBlock Name="txtProgressInfo" Text="Oczekiwanie na rozpoczęcie..." Margin="0,0,0,5" FontWeight="SemiBold" Foreground="{DynamicResource ThemeText}"/>
            <ProgressBar Name="progressBar" Height="20" Minimum="0" Maximum="100" Margin="0,0,0,5" BorderThickness="0" Foreground="{DynamicResource AccentColor}"/>
            <ProgressBar Name="progressBarDownload" Height="8" Minimum="0" Maximum="100" BorderThickness="0" Foreground="#FF107C10"/>
        </StackPanel>
        
        <Grid Grid.Row="3">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="110"/>
                <ColumnDefinition Width="110"/>
            </Grid.ColumnDefinitions>
            <Button Name="btnStart" Grid.Column="0" Content="ROZPOCZNIJ KONFIGURACJĘ" Height="55" FontSize="18" FontWeight="Bold" Foreground="White" Background="{DynamicResource StartColor}" Margin="0,0,10,0"/>
            <Button Name="btnPause" Grid.Column="1" Content="Pauza" Height="55" FontSize="18" FontWeight="Bold" Foreground="White" Background="#FFE6A100" Margin="0,0,10,0" IsEnabled="False"/>
            <Button Name="btnCancelDeploy" Grid.Column="2" Content="Przerwij" Height="55" FontSize="18" FontWeight="Bold" Foreground="White" Background="#FFC50F1F" IsEnabled="False"/>
        </Grid>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $mainXaml
$Window = [System.Windows.Markup.XamlReader]::Load($reader)

$btnThemeToggle      = $Window.FindName("btnThemeToggle")
$btnSelectAll        = $Window.FindName("btnSelectAll")
$btnDeselectAll      = $Window.FindName("btnDeselectAll")
$btnInvertSelection  = $Window.FindName("btnInvertSelection")
$spCheckboxes        = $Window.FindName("spCheckboxes")
$btnChooseApps       = $Window.FindName("btnChooseApps")
$btnSettings         = $Window.FindName("btnSettings")
$btnEditConfig       = $Window.FindName("btnEditConfig")
$btnReloadConfig     = $Window.FindName("btnReloadConfig")
$btnLogs             = $Window.FindName("btnLogs")
$btnSysProps         = $Window.FindName("btnSysProps")
$btnCompMgmt         = $Window.FindName("btnCompMgmt")
$btnRegEdit         = $Window.FindName("btnRegEdit")
$btnPrinters         = $Window.FindName("btnPrinters")
$btnSysInfo          = $Window.FindName("btnSysInfo")
$btnUninstaller      = $Window.FindName("btnUninstaller")
$rtbLog              = $Window.FindName("rtbLog")
$progressBar         = $Window.FindName("progressBar")
$progressBarDownload = $Window.FindName("progressBarDownload")
$txtProgressInfo     = $Window.FindName("txtProgressInfo")
$shpNetworkStatus    = $Window.FindName("shpNetworkStatus")
$txtNetworkStatus    = $Window.FindName("txtNetworkStatus")
$txtStopwatch        = $Window.FindName("txtStopwatch")
$btnStart            = $Window.FindName("btnStart")
$btnPause            = $Window.FindName("btnPause")
$btnCancelDeploy     = $Window.FindName("btnCancelDeploy")
$cmbProfiles         = $Window.FindName("cmbProfiles")
$script:isPaused     = $false
$script:isCancelled  = $false
$script:ignoreProfileChange = $false

$script:stopwatchTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:stopwatchTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:stopwatchAccumulated = [TimeSpan]::Zero
$script:stopwatchTimer.Add_Tick({
    if ($null -ne $script:stopwatchStartTime) {
        $elapsed = (Get-Date) - $script:stopwatchStartTime + $script:stopwatchAccumulated
        $txtStopwatch.Text = "⏱ $($elapsed.ToString('hh\:mm\:ss'))"
    }
})

try {
    $cfgCheck = Get-Config
    if ($null -ne $cfgCheck.DefaultCheckboxes) {
        foreach ($key in $checkboxOptions.Keys) {
            if ($null -ne $cfgCheck.DefaultCheckboxes.$key) {
                $checkboxOptions[$key].Enabled = [bool]$cfgCheck.DefaultCheckboxes.$key
            }
        }
    }
} catch { }

foreach ($key in $checkboxOptions.Keys) {
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content = $checkboxOptions[$key].Text
    $cb.Name = $key
    $cb.IsChecked = $checkboxOptions[$key].Enabled
    $cb.ToolTip = $checkboxOptions[$key].Tooltip
    
    $spCheckboxes.Children.Add($cb) | Out-Null
    $CheckboxControls[$key] = $cb
}

try {
    $cfg = Get-Config
    if ($null -ne $cfg.DarkTheme) {
        $script:isDarkTheme = [bool]$cfg.DarkTheme
    } else {
        $script:isDarkTheme = $true
    }
} catch { $script:isDarkTheme = $true }

function Set-AppTheme {
    if ($script:isDarkTheme) {
        $Window.Resources["ThemeBackground"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF202225")
        $Window.Resources["ThemePanel"]      = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF2F3136")
        $Window.Resources["ThemeText"]       = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFDCDDDE")
        $Window.Resources["ThemeButton"]     = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF4F545C")
        $Window.Resources["ThemeButtonText"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("White")
        $Window.Resources["ThemeBorder"]     = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF4F545C")
        $Window.Resources["ThemeTextBoxBg"]  = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF1E1E1E")
        $btnThemeToggle.Content = "☀️ Jasny Motyw"
    } else {
        $Window.Resources["ThemeBackground"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFF3F3F3")
        $Window.Resources["ThemePanel"]      = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFFFFFFF")
        $Window.Resources["ThemeText"]       = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF333333")
        $Window.Resources["ThemeButton"]     = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFE1E1E1")
        $Window.Resources["ThemeButtonText"] = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF000000")
        $Window.Resources["ThemeBorder"]     = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFCCCCCC")
        $Window.Resources["ThemeTextBoxBg"]  = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFFFFFFF")
        $btnThemeToggle.Content = "🌙 Ciemny Motyw"
    }
}
Set-AppTheme

$btnThemeToggle.Add_Click({
    $script:isDarkTheme = -not $script:isDarkTheme
    Set-AppTheme
    try {
        $cfg = Get-Config
        if ($null -eq $cfg.DarkTheme) {
            $cfg | Add-Member -NotePropertyName DarkTheme -NotePropertyValue $script:isDarkTheme -Force
        } else {
            $cfg.DarkTheme = $script:isDarkTheme
        }
        $cfg | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8
    } catch { }
})

$btnSelectAll.Add_Click({ foreach ($cb in $spCheckboxes.Children) { $cb.IsChecked = $true } })
$btnDeselectAll.Add_Click({ foreach ($cb in $spCheckboxes.Children) { $cb.IsChecked = $false } })
$btnInvertSelection.Add_Click({ foreach ($cb in $spCheckboxes.Children) { $cb.IsChecked = -not $cb.IsChecked } })

$cmbProfiles.Add_SelectionChanged({
    if ($script:ignoreProfileChange) { return }
    if ($cmbProfiles.SelectedIndex -gt 0) {
        $profName = $cmbProfiles.SelectedItem
        if ($profName -is [System.Windows.Controls.ComboBoxItem]) { $profName = $profName.Content }
        $profName = [string]$profName

        try {
            $cfg = Get-Config
            $apps = $cfg.Profiles.$profName
            $script:SelectedApps.Clear()
            if ($null -ne $apps) {
                foreach ($app in $apps) {
                    $matchedKey = $cfg.Programs.PSObject.Properties.Name | Where-Object { $_ -ieq $app } | Select-Object -First 1
                    if ($null -ne $matchedKey) {
                        $script:SelectedApps[$matchedKey] = $true
                    } else {
                        Write-Log "Nie znaleziono programu '$app' (profil '$profName') w konfiguracji." -Color "Yellow"
                    }
                }
            }
            $count = $script:SelectedApps.Count
            if ($null -ne $btnChooseApps) {
                $btnChooseApps.Content = "Wybierz aplikacje ($count)"
            }
            if ($CheckboxControls.ContainsKey("InstallApplications")) {
                $CheckboxControls["InstallApplications"].IsChecked = ($count -gt 0)
                $btnChooseApps.IsEnabled = ($count -gt 0)
            }
            Write-Log "Zastosowano profil wdrożenia: $profName (Wybrano programów: $count)" -Color "Green"
        } catch {
            Write-Log "Błąd podczas ładowania profilu: $_" -Color "Red" -IsError
        }
    } elseif ($cmbProfiles.SelectedIndex -eq 0) {
        # Gdy użytkownik celowo kliknie powrót na wybór niestandardowy - ładujemy domyślne z config.json
        Write-Log "Przełączono na wybór niestandardowy - wczytywanie domyślnych aplikacji z pliku config." -Color "Blue"
        Get-AppSelection
    }
})

$btnChooseApps.Add_Click({ Show-AppSelectionWindow })
$btnSettings.Add_Click({ 
    if (Show-PinPrompt) {
        Show-ConfigEditor 
    }
})
$btnEditConfig.Add_Click({
    if (Show-PinPrompt) {
        if (Test-Path $configPath) {
            Start-Process "notepad.exe" -ArgumentList "`"$configPath`""
        } else {
            Show-ThemedMessageBox -Message "Plik config.json nie istnieje pod ścieżką: $configPath" -Title "Błąd" -Button "OK" -Image "Warning" | Out-Null
        }
    }
})
$btnReloadConfig.Add_Click({
    try {
        $cfg = Get-Config
        if ($null -ne $cfg.DefaultCheckboxes) {
            foreach ($key in $checkboxOptions.Keys) {
                if ($null -ne $cfg.DefaultCheckboxes.$key) {
                    $val = [bool]$cfg.DefaultCheckboxes.$key
                    $checkboxOptions[$key].Enabled = $val
                    if ($CheckboxControls.ContainsKey($key)) {
                        $CheckboxControls[$key].IsChecked = $val
                    }
                }
            }
        }
        if ($null -ne $cfg.DarkTheme) {
            $script:isDarkTheme = [bool]$cfg.DarkTheme
            Set-AppTheme
        }
        Get-AppSelection
    Load-Profiles
        if ($null -ne $btnChooseApps -and $CheckboxControls.ContainsKey("InstallApplications")) {
            $btnChooseApps.IsEnabled = ($CheckboxControls["InstallApplications"].IsChecked -eq $true)
        }
        Write-Log "Pomyślnie przeładowano plik config.json i zaktualizowano GUI."
    } catch {
        Write-Log "Błąd przeładowania config.json: $_" -IsError
        Show-ThemedMessageBox -Message "Nie udało się przeładować config.json.`nSprawdź poprawność składni JSON w pliku." -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
    }
})
$btnLogs.Add_Click({ Show-LogWindow })
$btnSysProps.Add_Click({ Start-Process "systempropertiesadvanced" })
$btnCompMgmt.Add_Click({ Start-Process "compmgmt.msc" })
$btnRegEdit.Add_Click({ Start-Process "regedit" })
$btnPrinters.Add_Click({ Start-Process "explorer.exe" -ArgumentList "shell:::{2227A280-3AEA-1069-A2DE-08002B30309D}" })
$btnSysInfo.Add_Click({ Show-SystemInfoWindow })
$btnUninstaller.Add_Click({ Show-SoftwareUninstaller })
$btnStart.Add_Click({ Start-Deployment })

$btnPause.Add_Click({
    if ($script:isPaused) {
        $script:isPaused = $false
        $btnPause.Content = "Pauza"
        $btnPause.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFE6A100")
        $script:stopwatchStartTime = Get-Date
        $script:stopwatchTimer.Start()
        Write-Log "Wznowiono wdrożenie." -Context "Użytkownik"
    } else {
        $script:isPaused = $true
        $btnPause.Content = "Wznów"
        $btnPause.Background = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FF107C10")
        $script:stopwatchTimer.Stop()
        $script:stopwatchAccumulated += (Get-Date) - $script:stopwatchStartTime
        Write-Log "Wdrożenie wstrzymane (Pauza). Oczekiwanie na interakcję..." -Context "Użytkownik"
    }
})

$btnCancelDeploy.Add_Click({
    if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz przerwać wdrożenie?" -Title "Przerwij" -Button "YesNo" -Image "Warning") -eq [System.Windows.MessageBoxResult]::Yes) {
        $script:isCancelled = $true
        $script:isPaused = $false
        Write-Log "Wdrożenie przerwane przez użytkownika!" -IsError -Context "Użytkownik"
        $btnCancelDeploy.IsEnabled = $false
    }
})

$CheckboxControls["InstallApplications"].Add_Click({
    $btnChooseApps.IsEnabled = ($CheckboxControls["InstallApplications"].IsChecked -eq $true)
})

$Window.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
        if ($btnStart.IsEnabled) {
            if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz rozpocząć konfigurację?" -Title "Potwierdzenie" -Button "YesNo" -Image "Question") -eq [System.Windows.MessageBoxResult]::Yes) {
                Start-Deployment
            }
        }
    }
    elseif ($_.Key -eq [System.Windows.Input.Key]::Escape) {
        if ((Show-ThemedMessageBox -Message "Czy na pewno chcesz zamknąć aplikację?" -Title "Potwierdzenie" -Button "YesNo" -Image "Question") -eq [System.Windows.MessageBoxResult]::Yes) {
            $Window.Close()
        }
    }
})

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
try {
    $procPath = (Get-Process -Id $pid).Path
    if ($procPath -and (Test-Path $procPath)) {
        $notifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($procPath)
    } else {
        $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
    }
} catch {
    $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
}
$notifyIcon.Text = "Smart Tool for Deployment"

$notifyIcon.add_DoubleClick({
    $Window.ShowInTaskbar = $true
    $Window.WindowState = [System.Windows.WindowState]::Normal
    $Window.Activate()
    $notifyIcon.Visible = $false
})

$Window.Add_StateChanged({
    if ($Window.WindowState -eq [System.Windows.WindowState]::Minimized) {
        $Window.ShowInTaskbar = $false
        $notifyIcon.Visible = $true
        $notifyIcon.ShowBalloonTip(3000, "Smart Tool for Deployment", "Aplikacja została zminimalizowana i działa w tle.", [System.Windows.Forms.ToolTipIcon]::Info)
    }
})

$Window.Add_Closed({
    if ($null -ne $script:networkCheckTimer) {
        $script:networkCheckTimer.Stop()
        $script:networkCheckTimer = $null
    }
    if ($null -ne $script:stopwatchTimer) {
        $script:stopwatchTimer.Stop()
        $script:stopwatchTimer = $null
    }
    if ($null -ne $notifyIcon) {
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
    }
})

$script:networkCheckTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:networkCheckTimer.Interval = [TimeSpan]::FromSeconds(3)
$script:lastNetworkError = $null
$script:networkCheckTimer.Add_Tick({
    if (-not $btnStart.IsEnabled) { return } # Wstrzymaj sprawdzanie podczas trwającego wdrożenia

    $isNet = [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
    $validIp = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notmatch "^169\.254\." -and $_.IPAddress -ne "127.0.0.1" }
    
    if ($isNet -and $validIp) {
        $ipStr = $validIp[0].IPAddress
        $statusText = "Sieć: $ipStr"
        $statusColor = "#FF107C10"

        try {
            $cfg = Get-Config
            if ($cfg.DefaultInstallSource -eq 'web' -and -not [string]::IsNullOrWhiteSpace($cfg.InstallSourcePaths.web)) {
                $url = $cfg.InstallSourcePaths.web
                $req = [System.Net.WebRequest]::Create($url)
                $req.Timeout = 1500
                $req.Method = "HEAD"
                $req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                try {
                    $res = $req.GetResponse()
                    $statusText = "Sieć: $ipStr | Web: OK"
                    $res.Close()
                    if ($script:lastNetworkError) {
                        Write-Log "[Status Sieci] Połączenie ze źródłem Web przywrócone."
                        $script:lastNetworkError = $null
                    }
                } catch {
                    $webResp = $null
                    if ($_.Exception -is [System.Net.WebException]) { $webResp = $_.Exception.Response }
                    elseif ($_.Exception.InnerException -is [System.Net.WebException]) { $webResp = $_.Exception.InnerException.Response }

                    if ($null -ne $webResp) {
                        $statusCode = [int]$webResp.StatusCode
                        $statusText = "Sieć: $ipStr | Web: OK ($statusCode)"
                        $statusColor = "#FF107C10"
                        $webResp.Close()
                        if ($script:lastNetworkError) {
                            Write-Log "[Status Sieci] Źródło Web odpowiedziało statusem $statusCode."
                            $script:lastNetworkError = $null
                        }
                    } else {
                        $statusText = "Sieć: $ipStr | Web: Brak odp."
                        $statusColor = "#FFE6A100"
                        $errMsg = $_.Exception.Message
                        if ($script:lastNetworkError -ne $errMsg) {
                            Write-Log "[Status Sieci] Błąd weryfikacji źródła Web ($url): $errMsg" -IsError
                            $script:lastNetworkError = $errMsg
                        }
                    }
                }
            }
        } catch {}

        $shpNetworkStatus.Fill = (New-Object System.Windows.Media.BrushConverter).ConvertFromString($statusColor)
        $txtNetworkStatus.Text = $statusText
    } else {
        $shpNetworkStatus.Fill = (New-Object System.Windows.Media.BrushConverter).ConvertFromString("#FFC50F1F")
        $txtNetworkStatus.Text = "Brak połączenia sieciowego"
    }
})
$script:networkCheckTimer.Start()

if ($null -eq $global:PesterTesting) {
    Get-AppSelection
    Load-Profiles
    
    Write-Log "[System] Inicjalizacja środowiska graficznego (GUI) zakończona pomyślnie."
    $Window.ShowDialog() | Out-Null
}