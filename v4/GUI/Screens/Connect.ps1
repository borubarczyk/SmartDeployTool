# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Screens\Connect.ps1 - ekran „Połącz z siecią” (SPEC 7.3 pkt 1). Program pokazuje go tylko wtedy, gdy
# komputer nie ma sieci (Test-StdNetworkAvailable). Technik może:
# - podłączyć kabel - okno co 2 s sprawdza połączenie i samo przechodzi dalej, gdy sieć jest;
# - wybrać sieć Wi-Fi z listy (netsh wlan show networks, w tle) i wpisać hasło - profil sieci powstaje
#   z szablonu XML w %TEMP%, jest importowany przez netsh i od razu usuwany z dysku;
# - zaimportować profile z pendrive'a (<OriginRoot>\wifi\*.xml);
# - pominąć ten ekran.
# Wynik Show-StdConnectWindow: 'Connected', 'Skipped' albo $null (okno zamknięte krzyżykiem).
#
# Hasło sieci Wi-Fi nie trafia do logu ani do żadnego pliku STD (tylko do profilu sieci w Windows).

$script:StdConnectXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Połącz z siecią - STD" Width="600" Height="640" MinWidth="460" MinHeight="460"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Nagłówek: kafelek z ikoną Wi-Fi, tytuł i opis -->
        <Border Grid.Row="0" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeChromeBorder}" BorderThickness="0,0,0,1" Padding="20,14">
            <DockPanel>
                <Border DockPanel.Dock="Left" Style="{StaticResource IconTile}" VerticalAlignment="Center">
                    <TextBlock Style="{StaticResource Glyph}" Text="&#xE701;" FontSize="17" Foreground="{DynamicResource ThemeInfoFore}" HorizontalAlignment="Center"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Połącz z siecią" Style="{StaticResource DialogTitle}"/>
                    <TextBlock Style="{StaticResource DialogSubtitle}" Text="Ten komputer nie ma połączenia z siecią. Podłącz kabel sieciowy albo wybierz sieć Wi-Fi."/>
                </StackPanel>
            </DockPanel>
        </Border>

        <Grid Grid.Row="1" Margin="20,16,20,12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*" MinHeight="50"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Stan połączenia -->
            <Border Grid.Row="0" Style="{StaticResource Card}" Padding="14,10" Margin="0,0,0,12">
                <DockPanel>
                    <Ellipse x:Name="dotStatus" DockPanel.Dock="Left" Width="10" Height="10" Margin="0,0,12,0" VerticalAlignment="Center" Fill="{DynamicResource ThemeWarning}"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock x:Name="txtStatus" FontWeight="SemiBold" TextWrapping="Wrap" Foreground="{DynamicResource ThemeText}" Text="Brak połączenia z siecią"/>
                        <TextBlock x:Name="txtStatusDetail" Style="{StaticResource HintText}" Margin="0,2,0,0" Text="STD sprawdza połączenie co kilka sekund."/>
                    </StackPanel>
                </DockPanel>
            </Border>

            <!-- Lista sieci Wi-Fi -->
            <Border Grid.Row="1" Style="{StaticResource Card}" Padding="14,10,14,10">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <DockPanel Margin="0,0,0,6">
                        <Button x:Name="btnRefresh" DockPanel.Dock="Right" Style="{StaticResource GhostButton}"/>
                        <TextBlock Text="Sieci Wi-Fi" Style="{StaticResource CardTitle}" Margin="0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                    </DockPanel>
                    <ListBox x:Name="lstNetworks" Grid.Row="1" BorderThickness="0" Padding="0" Background="Transparent">
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Style="{StaticResource Glyph}" Text="&#xE701;" FontSize="15" Margin="0,0,12,0" Foreground="{DynamicResource ThemeAccentText}"/>
                                    <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                        <TextBlock Text="{Binding Ssid, Mode=OneWay}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource ThemeText}"/>
                                        <TextBlock Text="{Binding AuthLabel, Mode=OneWay}" Style="{StaticResource HintText}" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="12,0,0,0">
                                        <TextBlock x:Name="lockGlyph" Style="{StaticResource Glyph}" Text="&#xE72E;" FontSize="12" Margin="0,0,6,0" Foreground="{DynamicResource ThemeMuted}"/>
                                        <TextBlock Text="{Binding SignalText, Mode=OneWay}" Foreground="{DynamicResource ThemeMuted}" FontSize="12" MinWidth="34" TextAlignment="Right" VerticalAlignment="Center"/>
                                    </StackPanel>
                                </Grid>
                                <DataTemplate.Triggers>
                                    <DataTrigger Binding="{Binding Auth}" Value="open">
                                        <Setter TargetName="lockGlyph" Property="Visibility" Value="Collapsed"/>
                                    </DataTrigger>
                                </DataTemplate.Triggers>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>
                    <TextBlock x:Name="txtListHint" Grid.Row="1" Style="{StaticResource MutedText}" Margin="2,6,2,0" VerticalAlignment="Top"/>
                </Grid>
            </Border>

            <!-- Hasło wybranej sieci -->
            <Border x:Name="panelKey" Grid.Row="2" Style="{StaticResource Card}" Padding="14,12" Margin="0,12,0,0" Visibility="Collapsed">
                <StackPanel>
                    <TextBlock x:Name="txtKeyCaption" Style="{StaticResource Caption}"/>
                    <DockPanel>
                        <Button x:Name="btnConnect" DockPanel.Dock="Right" Style="{StaticResource PrimaryButton}" Margin="8,0,0,0" MinWidth="110" IsDefault="True"/>
                        <PasswordBox x:Name="pwdKey"/>
                    </DockPanel>
                    <TextBlock x:Name="txtKeyHint" Style="{StaticResource HintText}" Margin="0,6,0,0"/>
                </StackPanel>
            </Border>

            <!-- Komunikat (informacja, sukces, ostrzeżenie, błąd) -->
            <Border x:Name="boxMessage" Grid.Row="3" CornerRadius="8" Padding="12,8" Margin="0,12,0,0" Background="{DynamicResource ThemeInfoBack}" Visibility="Collapsed">
                <DockPanel>
                    <TextBlock x:Name="txtMessageIcon" DockPanel.Dock="Left" Style="{StaticResource Glyph}" FontSize="14" Margin="0,1,10,0" VerticalAlignment="Top" Foreground="{DynamicResource ThemeInfoFore}"/>
                    <TextBlock x:Name="txtMessage" TextWrapping="Wrap" Foreground="{DynamicResource ThemeInfoFore}"/>
                </DockPanel>
            </Border>
        </Grid>

        <!-- Pasek przycisków: import z pendrive'a po lewej, Pomiń / Dalej po prawej -->
        <Border Grid.Row="2" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeChromeBorder}" BorderThickness="0,1,0,0" Padding="20,10">
            <DockPanel>
                <Button x:Name="btnSkip" DockPanel.Dock="Right" MinWidth="110" Margin="8,0,0,0"/>
                <Button x:Name="btnImport" DockPanel.Dock="Left" Style="{StaticResource GhostButton}"/>
                <TextBlock x:Name="txtImportHint" Style="{StaticResource HintText}" VerticalAlignment="Center" Margin="10,0,0,0" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
            </DockPanel>
        </Border>

        <StackPanel x:Name="toastHost" Grid.RowSpan="3" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,16,64"/>
    </Grid>
</Window>
'@

$script:StdConnectControlNames = @(
    'dotStatus', 'txtStatus', 'txtStatusDetail', 'btnRefresh', 'lstNetworks', 'txtListHint', 'panelKey', 'txtKeyCaption',
    'btnConnect', 'pwdKey', 'txtKeyHint', 'boxMessage', 'txtMessageIcon', 'txtMessage', 'btnSkip', 'btnImport', 'txtImportHint'
)

# Funkcje wklejane do skryptów zadań w tle (runspace zadania ma tylko moduł STD).
$script:StdConnectTaskHelpers = @('Invoke-StdUiNetsh', 'New-StdWifiProfileXml')

$script:StdConnectScanScript = @'
param($Arguments)
. ([scriptblock]::Create($Arguments.Helpers))
$scan = Invoke-StdUiNetsh -ArgumentLine 'wlan show networks mode=bssid' -TimeoutSec 30
@{ ExitCode = $scan.ExitCode; Output = $scan.Output }
'@

$script:StdConnectWifiScript = @'
param($Arguments)
. ([scriptblock]::Create($Arguments.Helpers))
$ssid = [string]$Arguments.Ssid
$plain = $null
if ($null -ne $Arguments.Key) {
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Arguments.Key)
    try { $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) } finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    Register-StdLogMask -Value $plain
}
$xml = New-StdWifiProfileXml -Ssid $ssid -Auth ([string]$Arguments.Auth) -Key $plain
$plain = $null
$path = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ('std-wifi-' + [guid]::NewGuid().ToString('N') + '.xml'))
$add = $null
try {
    [System.IO.File]::WriteAllText($path, $xml, (New-Object System.Text.UTF8Encoding($false)))
    $xml = $null
    $add = Invoke-StdUiNetsh -ArgumentLine ('wlan add profile filename="{0}" user=all' -f $path)
} finally {
    # Plik z hasłem znika zaraz po imporcie (także po błędzie).
    try { if ([System.IO.File]::Exists($path)) { [System.IO.File]::Delete($path) } } catch { }
}
if ($add.ExitCode -ne 0) { throw ('[STD-UI-WIFI] Nie udało się dodać profilu sieci {0}: {1}' -f $ssid, $add.Output) }
$name = $ssid -replace '"', ''
$connect = Invoke-StdUiNetsh -ArgumentLine ('wlan connect name="{0}" ssid="{0}"' -f $name)
if ($connect.ExitCode -ne 0) { throw ('[STD-UI-WIFI] Nie udało się połączyć z siecią {0}: {1}' -f $ssid, $connect.Output) }
Write-StdLog -Source 'Połącz z siecią' -Message ('Dodano profil sieci Wi-Fi {0} i wysłano żądanie połączenia.' -f $ssid)
@{ Ssid = $ssid; Output = $connect.Output }
'@

$script:StdConnectImportScript = @'
param($Arguments)
. ([scriptblock]::Create($Arguments.Helpers))
$added = New-Object 'System.Collections.Generic.List[string]'
$failed = New-Object 'System.Collections.Generic.List[string]'
foreach ($file in @($Arguments.Files)) {
    $name = $null
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.Load([string]$file)
        $name = [string]$doc.DocumentElement.Item('name').InnerText
    } catch { $name = $null }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = [System.IO.Path]::GetFileNameWithoutExtension([string]$file) }
    $result = Invoke-StdUiNetsh -ArgumentLine ('wlan add profile filename="{0}" user=all' -f $file)
    if ($result.ExitCode -eq 0) {
        $added.Add($name)
    } else {
        $failed.Add(('{0}: {1}' -f [System.IO.Path]::GetFileName([string]$file), $result.Output))
    }
}
$target = $null
foreach ($candidate in $added) { if (@($Arguments.Visible) -contains $candidate) { $target = $candidate; break } }
if ($null -eq $target -and $added.Count -gt 0) { $target = $added[0] }
$connectOutput = ''
if ($null -ne $target) {
    $connect = Invoke-StdUiNetsh -ArgumentLine ('wlan connect name="{0}"' -f ($target -replace '"', ''))
    $connectOutput = $connect.Output
}
Write-StdLog -Source 'Połącz z siecią' -Message ('Import profili Wi-Fi z pendrive''a: dodano {0}, błędy {1}.' -f $added.Count, $failed.Count)
@{ Added = $added.ToArray(); Failed = $failed.ToArray(); Target = $target; ConnectOutput = $connectOutput }
'@

# ---------------------------------------------------------------------------------------------
# Wygląd: stan połączenia, komunikat, zajętość
# ---------------------------------------------------------------------------------------------

function Set-StdConnectStatus {
    # Stan połączenia: offline (brak sieci), connecting (łączenie), connected (jest sieć).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [ValidateSet('offline', 'connecting', 'connected')][string]$State, [string]$Detail)
    $c = $Ui.C
    switch ($State) {
        'connected' {
            Set-StdUiColor $c.dotStatus 'Fill' 'ThemeSuccess'
            $c.txtStatus.Text = 'Połączono z siecią'
            if ([string]::IsNullOrEmpty($Detail)) { $Detail = 'Komputer ma adres w sieci. Za chwilę STD przejdzie do logowania.' }
        }
        'connecting' {
            Set-StdUiColor $c.dotStatus 'Fill' 'ThemeInfo'
            $c.txtStatus.Text = 'Łączenie z siecią...'
            if ([string]::IsNullOrEmpty($Detail)) { $Detail = 'Czekam, aż komputer dostanie adres w sieci.' }
        }
        default {
            Set-StdUiColor $c.dotStatus 'Fill' 'ThemeWarning'
            $c.txtStatus.Text = 'Brak połączenia z siecią'
            if ([string]::IsNullOrEmpty($Detail)) { $Detail = 'STD sprawdza połączenie co kilka sekund - po podłączeniu kabla przejdzie dalej samo.' }
        }
    }
    $c.txtStatusDetail.Text = $Detail
}

function Set-StdConnectMessage {
    # Komunikat pod listą: -Kind Info / Ok / Warn / Crit (para kolorów Theme<Ton>Back / Fore). Pusty tekst ukrywa ramkę.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowEmptyString()][string]$Text, [ValidateSet('Info', 'Ok', 'Warn', 'Crit')][string]$Kind = 'Info')
    $c = $Ui.C
    if ([string]::IsNullOrWhiteSpace($Text)) {
        Set-StdUiVisible $c.boxMessage $false
        return
    }
    $glyph = @{ Info = 'E946'; Ok = 'E930'; Warn = 'E7BA'; Crit = 'EA39' }[$Kind]
    Set-StdUiColor $c.boxMessage 'Background' ('Theme{0}Back' -f $Kind)
    Set-StdUiColor $c.txtMessage 'Foreground' ('Theme{0}Fore' -f $Kind)
    Set-StdUiColor $c.txtMessageIcon 'Foreground' ('Theme{0}Fore' -f $Kind)
    $c.txtMessageIcon.Text = Get-StdGlyph -Code $glyph
    $c.txtMessage.Text = $Text
    Set-StdUiVisible $c.boxMessage $true
}

function Update-StdConnectButtons {
    # Przyciski zależnie od zajętości (skanowanie, łączenie, import) i stanu połączenia.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $busy = -not [string]::IsNullOrEmpty($Ui.Busy)
    $c.btnRefresh.IsEnabled = -not $busy
    $c.lstNetworks.IsEnabled = -not $busy
    $c.btnImport.IsEnabled = (-not $busy) -and ($Ui.UsbFiles.Count -gt 0)
    $selected = $c.lstNetworks.SelectedItem
    $canConnect = (-not $busy) -and ($null -ne $selected) -and ($selected.Auth -ne 'enterprise') -and (-not $Ui.Connected)
    $c.btnConnect.IsEnabled = $canConnect
    $c.pwdKey.IsEnabled = $canConnect
    if ($Ui.Connected) {
        Set-StdUiButton -Button $c.btnSkip -Text 'Dalej' -Icon 'E72A' -Style 'PrimaryButton' -ToolTip 'Przejdź do logowania'
    } else {
        Set-StdUiButton -Button $c.btnSkip -Text 'Pomiń' -Icon 'E72A' -Style '' -ToolTip 'Przejdź dalej bez sieci (np. praca z kopią konfiguracji z pendrive''a)'
    }
}

function Set-StdConnectBusy {
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowEmptyString()][string]$Busy)
    $Ui.Busy = $Busy
    Update-StdConnectButtons -Ui $Ui
}

function Set-StdConnectNetworks {
    # Lista sieci (wynik ConvertFrom-StdNetshNetworks) z podpowiedzią, gdy pusta.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowNull()][AllowEmptyCollection()][object[]]$Networks, [string]$EmptyText)
    $c = $Ui.C
    $list = @()
    if ($null -ne $Networks) { $list = @($Networks | Where-Object { $null -ne $_ }) }
    $Ui.Networks = $list
    $c.lstNetworks.ItemsSource = $list
    if ($list.Count -eq 0) {
        if ([string]::IsNullOrEmpty($EmptyText)) { $EmptyText = 'Nie widać żadnej sieci Wi-Fi. Podłącz kabel sieciowy albo kliknij „Odśwież”.' }
        $c.txtListHint.Text = $EmptyText
        Set-StdUiVisible $c.txtListHint $true
    } else {
        Set-StdUiVisible $c.txtListHint $false
    }
    Update-StdConnectSelection -Ui $Ui
}

function Update-StdConnectSelection {
    # Wybrana sieć: pole hasła (sieć zabezpieczona), sam przycisk (otwarta), informacja (802.1X).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $selected = $c.lstNetworks.SelectedItem
    if ($null -eq $selected) {
        Set-StdUiVisible $c.panelKey $false
        Update-StdConnectButtons -Ui $Ui
        return
    }
    Set-StdUiVisible $c.panelKey $true
    Set-StdUiButton -Button $c.btnConnect -Text 'Połącz' -Icon 'E701'
    switch ([string]$selected.Auth) {
        'open' {
            $c.txtKeyCaption.Text = 'Sieć {0} jest otwarta (bez hasła).' -f $selected.Ssid
            Set-StdUiVisible $c.pwdKey $false
            $c.txtKeyHint.Text = 'W sieci otwartej dane nie są szyfrowane. Używaj jej tylko wtedy, gdy nie ma innej.'
        }
        'enterprise' {
            $c.txtKeyCaption.Text = 'Sieć {0} wymaga logowania firmowego (802.1X).' -f $selected.Ssid
            Set-StdUiVisible $c.pwdKey $false
            $c.txtKeyHint.Text = 'Do takiej sieci STD nie łączy się samo. Zaimportuj profil z pendrive''a albo podłącz kabel sieciowy.'
        }
        default {
            $c.txtKeyCaption.Text = 'Hasło sieci {0}' -f $selected.Ssid
            Set-StdUiVisible $c.pwdKey $true
            $c.txtKeyHint.Text = 'Hasło trafia tylko do profilu sieci w systemie Windows. STD go nigdzie nie zapisuje.'
        }
    }
    Update-StdConnectButtons -Ui $Ui
}

# ---------------------------------------------------------------------------------------------
# Akcje: skanowanie, łączenie, import, sprawdzanie sieci
# ---------------------------------------------------------------------------------------------

function Start-StdConnectScan {
    # Lista sieci Wi-Fi w tle (netsh wlan show networks mode=bssid).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.Preview) {
        Show-StdToast -Window $Ui.Window -Message 'Podgląd okna: lista sieci jest przykładowa.' -Kind Info
        return
    }
    if (-not [string]::IsNullOrEmpty($Ui.Busy)) { return }
    Set-StdConnectBusy -Ui $Ui -Busy 'scan'
    $Ui.C.txtListHint.Text = 'Wyszukiwanie sieci Wi-Fi...'
    Set-StdUiVisible $Ui.C.txtListHint ($Ui.Networks.Count -eq 0)
    $null = Start-StdUiTask -Name 'Lista sieci Wi-Fi' -Owner $Ui.Window -ScriptText $script:StdConnectScanScript `
        -Arguments @{ Helpers = (Get-StdUiFunctionText -Name $script:StdConnectTaskHelpers) } `
        -OnComplete (New-StdUiCallback -Ui $Ui -Handler { param($Ui, $Result) Complete-StdConnectScan -Ui $Ui -Result $Result }) `
        -OnError (New-StdUiCallback -Ui $Ui -Handler {
            param($Ui, $ErrorRecord)
            Set-StdConnectBusy -Ui $Ui -Busy ''
            Set-StdConnectNetworks -Ui $Ui -Networks @() -EmptyText ('Nie udało się odczytać listy sieci Wi-Fi: {0}' -f (Get-StdUiErrorMessage $ErrorRecord -MaxLength 400))
        })
}

function Complete-StdConnectScan {
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowNull()]$Result)
    Set-StdConnectBusy -Ui $Ui -Busy ''
    $output = [string](Get-StdUiValue $Result 'Output')
    $networks = ConvertFrom-StdNetshNetworks -Text $output
    $empty = $null
    if (@($networks).Count -eq 0) {
        # netsh bez sieci: brak karty Wi-Fi albo wyłączona usługa WLAN - pokazujemy pierwszą linię opisu.
        $firstLine = (@(($output -replace "`r", '') -split "`n" | Where-Object { $_.Trim().Length -gt 0 }) | Select-Object -First 1)
        $empty = 'Nie widać żadnej sieci Wi-Fi. Podłącz kabel sieciowy albo kliknij „Odśwież”.'
        if (-not [string]::IsNullOrWhiteSpace([string]$firstLine) -and [int](Get-StdUiValue $Result 'ExitCode') -ne 0) {
            $empty = '{0}{1}{1}netsh: {2}' -f $empty, "`n", ([string]$firstLine).Trim()
        }
    }
    Set-StdConnectNetworks -Ui $Ui -Networks $networks -EmptyText $empty
}

function Start-StdConnectWifi {
    # Profil sieci z hasłem (XML w %TEMP%, usuwany zaraz po imporcie) i połączenie - w tle.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $selected = $c.lstNetworks.SelectedItem
    if ($null -eq $selected -or -not [string]::IsNullOrEmpty($Ui.Busy)) { return }
    $auth = [string]$selected.Auth
    if ($auth -eq 'enterprise') { return }
    if ($auth -eq 'unknown') { $auth = 'WPA2PSK' }
    $key = $null
    if ($auth -ne 'open') {
        $length = $c.pwdKey.SecurePassword.Length
        if ($length -lt 8 -or $length -gt 64) {
            Set-StdConnectMessage -Ui $Ui -Kind Warn -Text 'Hasło sieci Wi-Fi musi mieć od 8 do 63 znaków.'
            [void]$c.pwdKey.Focus()
            return
        }
        $key = $c.pwdKey.SecurePassword
    }
    if ($Ui.Preview) {
        Set-StdConnectMessage -Ui $Ui -Kind Info -Text ('Podgląd okna: połączenie z siecią {0} nie jest wykonywane.' -f $selected.Ssid)
        return
    }
    Set-StdConnectBusy -Ui $Ui -Busy 'connect'
    Set-StdConnectStatus -Ui $Ui -State connecting -Detail ('Dodawanie profilu sieci {0}...' -f $selected.Ssid)
    Set-StdConnectMessage -Ui $Ui -Text ''
    $null = Start-StdUiTask -Name 'Połączenie z Wi-Fi' -Owner $Ui.Window -ScriptText $script:StdConnectWifiScript `
        -Arguments @{ Helpers = (Get-StdUiFunctionText -Name $script:StdConnectTaskHelpers); Ssid = [string]$selected.Ssid; Auth = $auth; Key = $key } `
        -OnComplete (New-StdUiCallback -Ui $Ui -Handler {
            param($Ui, $Result)
            Set-StdConnectBusy -Ui $Ui -Busy ''
            $Ui.C.pwdKey.Clear()
            $ssid = [string](Get-StdUiValue $Result 'Ssid')
            $Ui.ConnectSsid = $ssid
            $Ui.ConnectStartedAt = Get-Date
            Set-StdConnectStatus -Ui $Ui -State connecting -Detail ('Wysłano żądanie połączenia z siecią {0}. Czekam na adres w sieci...' -f $ssid)
        }) `
        -OnError (New-StdUiCallback -Ui $Ui -Handler {
            param($Ui, $ErrorRecord)
            Set-StdConnectBusy -Ui $Ui -Busy ''
            Set-StdConnectStatus -Ui $Ui -State offline
            Set-StdConnectMessage -Ui $Ui -Kind Crit -Text (Get-StdUiErrorMessage $ErrorRecord -MaxLength 600)
        })
}

function Start-StdConnectImport {
    # Profile Wi-Fi z pendrive'a (<OriginRoot>\wifi\*.xml): import przez netsh i połączenie - w tle.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.UsbFiles.Count -eq 0 -or -not [string]::IsNullOrEmpty($Ui.Busy)) { return }
    if ($Ui.Preview) {
        Set-StdConnectMessage -Ui $Ui -Kind Info -Text 'Podgląd okna: import profili nie jest wykonywany.'
        return
    }
    $visible = @($Ui.Networks | ForEach-Object { [string]$_.Ssid })
    Set-StdConnectBusy -Ui $Ui -Busy 'import'
    Set-StdConnectStatus -Ui $Ui -State connecting -Detail 'Importowanie profili Wi-Fi z pendrive''a...'
    Set-StdConnectMessage -Ui $Ui -Text ''
    $null = Start-StdUiTask -Name 'Import profili Wi-Fi' -Owner $Ui.Window -ScriptText $script:StdConnectImportScript `
        -Arguments @{ Helpers = (Get-StdUiFunctionText -Name $script:StdConnectTaskHelpers); Files = @($Ui.UsbFiles); Visible = $visible } `
        -OnComplete (New-StdUiCallback -Ui $Ui -Handler { param($Ui, $Result) Complete-StdConnectImport -Ui $Ui -Result $Result }) `
        -OnError (New-StdUiCallback -Ui $Ui -Handler {
            param($Ui, $ErrorRecord)
            Set-StdConnectBusy -Ui $Ui -Busy ''
            Set-StdConnectStatus -Ui $Ui -State offline
            Set-StdConnectMessage -Ui $Ui -Kind Crit -Text ('Nie udało się zaimportować profili Wi-Fi: {0}' -f (Get-StdUiErrorMessage $ErrorRecord -MaxLength 600))
        })
}

function Complete-StdConnectImport {
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowNull()]$Result)
    Set-StdConnectBusy -Ui $Ui -Busy ''
    $added = @(Get-StdUiValue $Result 'Added') | Where-Object { -not [string]::IsNullOrEmpty([string]$_) }
    $failed = @(Get-StdUiValue $Result 'Failed') | Where-Object { -not [string]::IsNullOrEmpty([string]$_) }
    $target = [string](Get-StdUiValue $Result 'Target')
    $lines = New-Object 'System.Collections.Generic.List[string]'
    if (@($added).Count -gt 0) { $lines.Add(('Zaimportowane profile: {0}.' -f (@($added) -join ', '))) }
    foreach ($line in $failed) { $lines.Add(('Błąd importu {0}' -f $line)) }
    if (@($added).Count -eq 0) {
        Set-StdConnectStatus -Ui $Ui -State offline
        if ($lines.Count -eq 0) { $lines.Add('Nie udało się zaimportować żadnego profilu Wi-Fi.') }
        Set-StdConnectMessage -Ui $Ui -Kind Crit -Text ($lines -join "`n")
        return
    }
    $kind = 'Ok'
    if (@($failed).Count -gt 0) { $kind = 'Warn' }
    Set-StdConnectMessage -Ui $Ui -Kind $kind -Text ($lines -join "`n")
    if (-not [string]::IsNullOrEmpty($target)) {
        $Ui.ConnectSsid = $target
        $Ui.ConnectStartedAt = Get-Date
        Set-StdConnectStatus -Ui $Ui -State connecting -Detail ('Łączenie z siecią {0}. Czekam na adres w sieci...' -f $target)
    }
}

function Invoke-StdConnectCheck {
    # Co 2 s: czy komputer ma już sieć (kabel, Wi-Fi). Tak -> zielony stan i przejście dalej po chwili.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.Connected) { return }
    if (Test-StdNetworkAvailable) {
        $Ui.Connected = $true
        $Ui.ConnectStartedAt = $null
        Set-StdConnectStatus -Ui $Ui -State connected
        Set-StdConnectMessage -Ui $Ui -Text ''
        Update-StdConnectButtons -Ui $Ui
        Write-StdLog -Source 'Połącz z siecią' -Message 'Komputer ma połączenie z siecią.'
        if ($null -ne $Ui.CloseTimer) { $Ui.CloseTimer.Start() }
        return
    }
    if ($null -ne $Ui.ConnectStartedAt -and ((Get-Date) - $Ui.ConnectStartedAt).TotalSeconds -ge 25) {
        $ssid = $Ui.ConnectSsid
        $Ui.ConnectStartedAt = $null
        Set-StdConnectStatus -Ui $Ui -State offline
        Set-StdConnectMessage -Ui $Ui -Kind Warn -Text ('Nie udało się połączyć z siecią {0} w ciągu 25 s. Sprawdź hasło i spróbuj ponownie albo podłącz kabel sieciowy.' -f $ssid)
    }
}

function Close-StdConnectWindow {
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [Parameter(Mandatory = $true)][string]$Result)
    $Ui.Result = $Result
    $Ui.Window.Close()
}

# ---------------------------------------------------------------------------------------------
# Okno
# ---------------------------------------------------------------------------------------------

function New-StdConnectWindow {
    <#
    .SYNOPSIS
    Buduje okno „Połącz z siecią” (bez pokazywania). Stan okna: Window.Tag ($ui). Praca w tle (lista sieci,
    sprawdzanie połączenia) startuje dopiero po pokazaniu okna (Loaded).
    .PARAMETER PreviewData
    Podgląd okna bez sieci i netsh: Networks (lista sieci jak z ConvertFrom-StdNetshNetworks), SelectedSsid,
    Status ('offline'|'connecting'|'connected'), Message, MessageKind ('Info'|'Ok'|'Warn'|'Crit'), UsbProfiles (liczba).
    Brakujące klucze biorą wartości przykładowe (Get-StdUiPreviewSample).
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$OriginRoot,
        [AllowNull()][hashtable]$PreviewData
    )
    $win = New-StdWindow -Xaml $script:StdConnectXaml
    $c = Get-StdUiControls -Root $win -Names $script:StdConnectControlNames
    $preview = ($null -ne $PreviewData)
    $ui = @{
        Kind             = 'Connect'
        Window           = $win
        C                = $c
        OriginRoot       = $OriginRoot
        Preview          = $preview
        PreviewData      = $null
        Result           = $null
        Networks         = @()
        Busy             = ''
        Connected        = $false
        ConnectSsid      = $null
        ConnectStartedAt = $null
        UsbFiles         = @()
        CheckTimer       = $null
        CloseTimer       = $null
    }
    $win.Tag = $ui
    if ($preview) { $ui.PreviewData = Merge-StdUiPreviewData -PreviewData $PreviewData }

    Set-StdUiButton -Button $c.btnRefresh -Text 'Odśwież' -Icon 'E72C' -ToolTip 'Wyszukaj sieci Wi-Fi jeszcze raz'
    Set-StdUiButton -Button $c.btnImport -Text 'Importuj z pendrive''a' -Icon 'E896' -ToolTip 'Zaimportuj profile Wi-Fi z folderu wifi na pendrivie'
    Set-StdUiButton -Button $c.btnConnect -Text 'Połącz' -Icon 'E701'
    Set-StdConnectStatus -Ui $ui -State offline

    # Profile Wi-Fi na pendrivie (<OriginRoot>\wifi\*.xml).
    if ($preview) {
        $count = [int](Get-StdUiValue $ui.PreviewData 'UsbProfiles')
        $files = New-Object 'System.Collections.Generic.List[string]'
        for ($i = 1; $i -le $count; $i++) { $files.Add(('wifi-{0}.xml' -f $i)) }
        $ui.UsbFiles = $files.ToArray()
    } elseif (-not [string]::IsNullOrWhiteSpace($OriginRoot)) {
        try {
            $wifiDir = Join-Path $OriginRoot 'wifi'
            if ([System.IO.Directory]::Exists($wifiDir)) {
                $ui.UsbFiles = @(Get-ChildItem -LiteralPath $wifiDir -Filter '*.xml' -File -ErrorAction Stop | Where-Object { $_.Extension -eq '.xml' } | Sort-Object -Property Name | ForEach-Object { $_.FullName })
            }
        } catch {
            Write-StdLog -Level Warn -Source 'Połącz z siecią' -Message ('Nie udało się odczytać folderu wifi na pendrivie: {0}' -f $_.Exception.Message)
        }
    }
    if ($ui.UsbFiles.Count -gt 0) {
        $c.txtImportHint.Text = 'Profile Wi-Fi na pendrivie: {0}' -f $ui.UsbFiles.Count
    } else {
        $c.txtImportHint.Text = 'Brak profili Wi-Fi na pendrivie (folder wifi).'
    }
    Update-StdConnectButtons -Ui $ui

    # Timery: sprawdzanie sieci co 2 s i jednorazowe przejście dalej po wykryciu sieci.
    if (-not $preview) {
        $ui.CheckTimer = New-StdUiTimer -Ui $ui -IntervalMs 2000 -OnTick { param($Ui, $Timer) Invoke-StdConnectCheck -Ui $Ui }
        $ui.CloseTimer = New-StdUiTimer -Ui $ui -IntervalMs 1200 -OnTick {
            param($Ui, $Timer)
            $Timer.Stop()
            Close-StdConnectWindow -Ui $Ui -Result 'Connected'
        }
    }

    # Handlery.
    $win.Add_Loaded({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                if ($Ui.Preview) { return }
                $Ui.CheckTimer.Start()
                Invoke-StdConnectCheck -Ui $Ui
                if (-not $Ui.Connected) { Start-StdConnectScan -Ui $Ui }
            }
        })
    $win.Add_Closed({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                foreach ($timer in @($Ui.CheckTimer, $Ui.CloseTimer)) { if ($null -ne $timer) { $timer.Stop() } }
                try { $Ui.C.pwdKey.Clear() } catch { }
            }
        })
    $c.lstNetworks.Add_SelectionChanged({
            param($s, $e)
            Invoke-StdUiHandler $s { param($Ui) Set-StdConnectMessage -Ui $Ui -Text ''; Update-StdConnectSelection -Ui $Ui }
        })
    $c.lstNetworks.Add_MouseDoubleClick({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                $item = $Ui.C.lstNetworks.SelectedItem
                if ($null -ne $item -and [string]$item.Auth -eq 'open') { Start-StdConnectWifi -Ui $Ui }
                elseif ($null -ne $item -and $Ui.C.pwdKey.IsVisible) { [void]$Ui.C.pwdKey.Focus() }
            }
        })
    $c.btnConnect.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-StdConnectWifi -Ui $Ui } })
    $c.btnRefresh.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-StdConnectScan -Ui $Ui } })
    $c.btnImport.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-StdConnectImport -Ui $Ui } })
    $c.btnSkip.Add_Click({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                if ($Ui.Connected) {
                    Close-StdConnectWindow -Ui $Ui -Result 'Connected'
                } else {
                    Write-StdLog -Source 'Połącz z siecią' -Message 'Technik pominął ekran „Połącz z siecią”.'
                    Close-StdConnectWindow -Ui $Ui -Result 'Skipped'
                }
            }
        })

    if ($preview) { Set-StdConnectPreview -Ui $ui }
    return $win
}

function Set-StdConnectPreview {
    # Okno z danymi podglądu (bez netsh i bez sprawdzania sieci).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $data = $Ui.PreviewData
    $networks = @(Get-StdUiValue $data 'Networks') | Where-Object { $null -ne $_ }
    Set-StdConnectNetworks -Ui $Ui -Networks @($networks)
    $selectedSsid = [string](Get-StdUiValue $data 'SelectedSsid')
    if ([string]::IsNullOrEmpty($selectedSsid) -and @($networks).Count -gt 0) { $selectedSsid = [string]@($networks)[0].Ssid }
    foreach ($item in @($networks)) {
        if ([string]$item.Ssid -eq $selectedSsid) { $Ui.C.lstNetworks.SelectedItem = $item; break }
    }
    $status = [string](Get-StdUiValue $data 'Status')
    if ($status -eq 'connected') { $Ui.Connected = $true }
    if ($status -notin @('offline', 'connecting', 'connected')) { $status = 'offline' }
    Set-StdConnectStatus -Ui $Ui -State $status
    $message = [string](Get-StdUiValue $data 'Message')
    $kind = [string](Get-StdUiValue $data 'MessageKind')
    if ($kind -notin @('Info', 'Ok', 'Warn', 'Crit')) { $kind = 'Info' }
    Set-StdConnectMessage -Ui $Ui -Text $message -Kind $kind
    Update-StdConnectSelection -Ui $Ui
}

function Show-StdConnectWindow {
    <#
    .SYNOPSIS
    Ekran „Połącz z siecią”. Wynik: 'Connected' (komputer ma sieć), 'Skipped' (technik pominął ekran)
    albo $null (okno zamknięte). -PreviewData: podgląd z przykładowymi danymi (bez netsh), jak w New-StdConnectWindow.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$OriginRoot,
        [AllowNull()][hashtable]$PreviewData
    )
    $win = Select-StdUiWindow -Output @(New-StdConnectWindow -OriginRoot $OriginRoot -PreviewData $PreviewData)
    [void]$win.ShowDialog()
    $result = $win.Tag.Result
    if ($result -eq 'Connected' -or $result -eq 'Skipped') { return [string]$result }
    return $null
}
