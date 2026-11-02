# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Screens\Login.ps1 - ekran „Logowanie” (SPEC 7.3 pkt 2): adres serwera STD (z bootstrap.json; gdy go nie ma -
# pole do wpisania, adres zapisywany przez Save-StdBootstrapLocal po udanym logowaniu), login i hasło konta
# technika, „Zaloguj” (w tle: New-StdSession + Get-StdClientIndex), błąd pod polami i „Pracuj offline”, gdy jest
# kopia listy klientów (cache albo pendrive).
#
# Logowanie jest udane tylko wtedy, gdy listę klientów dał serwer (źródło 'server') - wtedy login i hasło są
# sprawdzone. Gdy serwer odmówił albo nie odpowiada, a lista przyszła z kopii, okno pokazuje błąd serwera
# i proponuje pracę offline z tą kopią (bez sesji).
# Wynik Show-StdLoginWindow: @{ Session; Index; Source; Bootstrap } albo $null (okno zamknięte). Bootstrap to plik
# startowy dla okna głównego: po udanym logowaniu ten, którego użyło logowanie (także adres wpisany w oknie
# i zapisany), przy pracy offline - plik startowy podany oknu. Program nie musi więc czytać bootstrap.json drugi raz.
# Hasło żyje tylko w pamięci (SecureString w sesji) - nie trafia do logu ani do pliku.

$script:StdLoginXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Logowanie - STD" Width="420" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeHeader}" Foreground="{DynamicResource ThemeText}">
    <Grid>
        <!-- Przewijanie: przy 1366x768 i 150% okno może być niższe niż jego zawartość. -->
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
            <StackPanel Margin="28,24,28,22">
                <Border Width="48" Height="48" CornerRadius="13" HorizontalAlignment="Center">
                    <Border.Background>
                        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                            <GradientStop Color="#4C7DF0" Offset="0"/>
                            <GradientStop Color="#E0913E" Offset="1"/>
                        </LinearGradientBrush>
                    </Border.Background>
                    <TextBlock Style="{StaticResource Glyph}" Text="&#xE7B8;" FontSize="21" Foreground="{DynamicResource ThemeOnAccent}" HorizontalAlignment="Center"/>
                </Border>
                <TextBlock Text="Smart Tool for Deployment" FontSize="18" FontWeight="SemiBold" Foreground="{DynamicResource ThemeTitle}" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="0,12,0,0"/>
                <TextBlock x:Name="txtVersion" Foreground="{DynamicResource ThemeMuted}" FontSize="11.5" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="0,2,0,0"/>
                <TextBlock Text="Zaloguj się kontem technika serwera STD." Foreground="{DynamicResource ThemeMuted}" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="0,8,0,18"/>

                <TextBlock Text="Serwer STD" Style="{StaticResource Caption}"/>
                <Border x:Name="boxServer" Background="{DynamicResource ThemeChip}" CornerRadius="7" Padding="10,7" Margin="0,0,0,12">
                    <DockPanel>
                        <TextBlock DockPanel.Dock="Left" Style="{StaticResource Glyph}" Text="&#xE72E;" FontSize="12" Margin="0,0,8,0" Foreground="{DynamicResource ThemeMuted}"/>
                        <TextBlock x:Name="txtServerShown" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource ThemeText}" VerticalAlignment="Center"/>
                    </DockPanel>
                </Border>
                <TextBox x:Name="txtServer" Tag="https://serwer.firma.pl/std/" Visibility="Collapsed"/>
                <TextBlock x:Name="txtServerHint" Style="{StaticResource HintText}" Margin="0,5,0,12" Visibility="Collapsed"
                           Text="Po udanym logowaniu adres zostanie zapamiętany na tym komputerze."/>

                <TextBlock Text="Login" Style="{StaticResource Caption}"/>
                <TextBox x:Name="txtUser" Margin="0,0,0,12"/>
                <TextBlock Text="Hasło" Style="{StaticResource Caption}"/>
                <PasswordBox x:Name="pwdPassword"/>

                <Border x:Name="boxError" CornerRadius="8" Padding="12,8" Margin="0,14,0,0" Background="{DynamicResource ThemeCritBack}" Visibility="Collapsed">
                    <DockPanel>
                        <TextBlock DockPanel.Dock="Left" Style="{StaticResource Glyph}" Text="&#xEA39;" FontSize="14" Margin="0,1,10,0" VerticalAlignment="Top" Foreground="{DynamicResource ThemeCritFore}"/>
                        <TextBlock x:Name="txtError" TextWrapping="Wrap" Foreground="{DynamicResource ThemeCritFore}"/>
                    </DockPanel>
                </Border>
                <Border x:Name="boxBusy" CornerRadius="8" Padding="12,8" Margin="0,14,0,0" Background="{DynamicResource ThemeInfoBack}" Visibility="Collapsed">
                    <DockPanel>
                        <TextBlock DockPanel.Dock="Left" Style="{StaticResource Glyph}" Text="&#xE72C;" FontSize="13" Margin="0,1,10,0" VerticalAlignment="Top" Foreground="{DynamicResource ThemeInfoFore}"/>
                        <TextBlock x:Name="txtBusy" TextWrapping="Wrap" Foreground="{DynamicResource ThemeInfoFore}"/>
                    </DockPanel>
                </Border>

                <Button x:Name="btnLogin" Style="{StaticResource PrimaryButton}" IsDefault="True" MinHeight="34" Margin="0,18,0,0"/>
                <Button x:Name="btnOffline" MinHeight="32" Margin="0,8,0,0" Visibility="Collapsed"/>
                <TextBlock Style="{StaticResource HintText}" Margin="0,14,0,0" TextAlignment="Center"
                           Text="Konto technika służy tylko do pobrania konfiguracji i instalatorów z serwera STD. STD nie zapisuje hasła."/>
            </StackPanel>
        </ScrollViewer>
        <StackPanel x:Name="toastHost" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,16,16"/>
    </Grid>
</Window>
'@

$script:StdLoginControlNames = @(
    'txtVersion', 'boxServer', 'txtServerShown', 'txtServer', 'txtServerHint', 'txtUser', 'pwdPassword',
    'boxError', 'txtError', 'boxBusy', 'txtBusy', 'btnLogin', 'btnOffline'
)

$script:StdLoginScript = @'
param($Arguments)
$session = New-StdSession -UserName ([string]$Arguments.UserName) -Password $Arguments.Password
$index = Get-StdClientIndex -Bootstrap $Arguments.Bootstrap -Session $session -UsbRoot ([string]$Arguments.UsbRoot)
@{ Session = $session; Index = $index }
'@

$script:StdLoginOfflineScript = @'
param($Arguments)
Get-StdClientIndex -Bootstrap $Arguments.Bootstrap -UsbRoot ([string]$Arguments.UsbRoot)
'@

# ---------------------------------------------------------------------------------------------
# Wygląd
# ---------------------------------------------------------------------------------------------

function Set-StdLoginError {
    # Błąd pod polami (pusty tekst ukrywa ramkę).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowEmptyString()][string]$Text)
    $Ui.C.txtError.Text = $Text
    Set-StdUiVisible $Ui.C.boxError (-not [string]::IsNullOrWhiteSpace($Text))
}

function Set-StdLoginBusy {
    # Logowanie w toku: pola i przyciski zablokowane, informacja w niebieskiej ramce.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowEmptyString()][string]$Text)
    $c = $Ui.C
    $busy = -not [string]::IsNullOrWhiteSpace($Text)
    $Ui.Busy = $busy
    $c.txtBusy.Text = $Text
    Set-StdUiVisible $c.boxBusy $busy
    $c.txtServer.IsEnabled = -not $busy
    $c.txtUser.IsEnabled = -not $busy
    $c.pwdPassword.IsEnabled = -not $busy
    $c.btnLogin.IsEnabled = -not $busy
    $c.btnOffline.IsEnabled = -not $busy
}

function Update-StdLoginOffline {
    # „Pracuj offline”: widoczny, gdy jest kopia listy klientów (z nieudanego logowania albo w cache / na pendrivie).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $available = ($null -ne $Ui.OfflineResult) -or $Ui.OfflineIndexExists
    Set-StdUiVisible $c.btnOffline $available
    if (-not $available) { return }
    $label = 'Pracuj offline'
    $tip = 'Praca z kopią listy klientów zapisaną na tym komputerze albo na pendrivie (bez logowania do serwera).'
    if ($null -ne $Ui.OfflineResult) {
        $label = 'Pracuj offline ({0})' -f (Get-StdUiSourceLabel ([string]$Ui.OfflineResult.Source))
        $tip = 'Praca z kopią listy klientów, która przyszła ze źródła: {0}.' -f (Get-StdUiSourceLabel ([string]$Ui.OfflineResult.Source))
    }
    Set-StdUiButton -Button $c.btnOffline -Text $label -Icon 'E8B7' -ToolTip $tip
}

# ---------------------------------------------------------------------------------------------
# Akcje
# ---------------------------------------------------------------------------------------------

function Get-StdLoginBootstrap {
    # Plik startowy dla logowania: z bootstrap.json, a gdy go nie ma - z adresu wpisanego w oknie
    # (Get-StdBootstrap -Server sprawdza https:// i dopisuje ukośnik). Błąd adresu: [STD-CFG-BOOTSTRAP].
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($null -ne $Ui.Bootstrap -and -not $Ui.ServerEditable) { return $Ui.Bootstrap }
    $typed = ([string]$Ui.C.txtServer.Text).Trim()
    if ($typed.Length -eq 0) { throw '[STD-CFG-BOOTSTRAP] Wpisz adres serwera STD, np. https://serwer.firma.pl/std/.' }
    if ($typed -notmatch '^(?i)https://') {
        if ($typed -match '^(?i)http://') { throw '[STD-CFG-BOOTSTRAP] Adres serwera musi zaczynać się od https:// (połączenie szyfrowane).' }
        $typed = 'https://' + $typed
    }
    return (Get-StdBootstrap -Server $typed)
}

function Start-StdLoginAttempt {
    # „Zaloguj”: sprawdzenie pól, potem w tle New-StdSession + Get-StdClientIndex.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.Busy) { return }
    $c = $Ui.C
    Set-StdLoginError -Ui $Ui -Text ''
    $bootstrap = $null
    try {
        $bootstrap = Get-StdLoginBootstrap -Ui $Ui
    } catch {
        Set-StdLoginError -Ui $Ui -Text ((Get-StdUiErrorMessage $_) -replace '^\s*\[[A-Z0-9-]+\]\s*(parametr -Server -\s*)?', '')
        [void]$c.txtServer.Focus()
        return
    }
    $user = ([string]$c.txtUser.Text).Trim()
    if ($user.Length -eq 0) {
        Set-StdLoginError -Ui $Ui -Text 'Wpisz login konta technika.'
        [void]$c.txtUser.Focus()
        return
    }
    if ($user.Contains(':')) {
        Set-StdLoginError -Ui $Ui -Text 'Login nie może zawierać dwukropka (:).'
        [void]$c.txtUser.Focus()
        return
    }
    if ($c.pwdPassword.SecurePassword.Length -eq 0) {
        Set-StdLoginError -Ui $Ui -Text 'Wpisz hasło.'
        [void]$c.pwdPassword.Focus()
        return
    }
    if ($Ui.Preview) {
        $data = $Ui.PreviewData
        $Ui.Result = @{ Session = $null; Index = (Get-StdUiValue $data 'Index'); Source = 'preview'; Bootstrap = $Ui.Bootstrap }
        $Ui.Window.Close()
        return
    }
    $Ui.LoginBootstrap = $bootstrap
    Set-StdLoginBusy -Ui $Ui -Text ('Logowanie do serwera {0}...' -f [string]$bootstrap['server'])
    $null = Start-StdUiTask -Name 'Logowanie' -Owner $Ui.Window -ScriptText $script:StdLoginScript `
        -Arguments @{ UserName = $user; Password = $c.pwdPassword.SecurePassword; Bootstrap = (ConvertTo-StdHashtable -InputObject $bootstrap); UsbRoot = [string]$Ui.OriginRoot } `
        -OnComplete (New-StdUiCallback -Ui $Ui -Handler { param($Ui, $Result) Complete-StdLoginAttempt -Ui $Ui -Result $Result }) `
        -OnError (New-StdUiCallback -Ui $Ui -Handler {
            param($Ui, $ErrorRecord)
            Set-StdLoginBusy -Ui $Ui -Text ''
            $failure = Get-StdUiLoginFailure -ErrorObject $ErrorRecord
            Set-StdLoginError -Ui $Ui -Text $failure.Message
            $Ui.OfflineIndexExists = Test-StdUiOfflineIndex -OriginRoot $Ui.OriginRoot
            Update-StdLoginOffline -Ui $Ui
            if ($failure.Message -like 'Nieprawidłowy login lub hasło*') { $Ui.C.pwdPassword.Clear(); [void]$Ui.C.pwdPassword.Focus() }
        })
}

function Complete-StdLoginAttempt {
    # Wynik logowania: lista klientów z serwera = sukces; z kopii = błąd serwera i propozycja pracy offline.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowNull()]$Result)
    Set-StdLoginBusy -Ui $Ui -Text ''
    $index = Get-StdUiValue $Result 'Index'
    $source = [string](Get-StdUiValue $index 'Source')
    if ($source -eq 'server') {
        if ($Ui.ServerEditable) {
            try {
                Save-StdBootstrapLocal -Bootstrap $Ui.LoginBootstrap
            } catch {
                Write-StdLog -Level Warn -Source 'Logowanie' -Message ('Nie udało się zapisać adresu serwera: {0}' -f $_.Exception.Message)
            }
        }
        $session = Get-StdUiValue $Result 'Session'
        Write-StdLog -Source 'Logowanie' -Message ('Zalogowano do serwera STD jako {0}. Lista klientów: rewizja {1}.' -f [string]$session.UserName, (Get-StdUiValue $index 'Revision'))
        $Ui.Result = @{ Session = $session; Index = (Get-StdUiValue $index 'Index'); Source = 'server'; Bootstrap = $Ui.LoginBootstrap }
        $Ui.Window.Close()
        return
    }
    $failure = Get-StdUiLoginFailure -IndexResult $index
    $Ui.OfflineResult = @{ Session = $null; Index = (Get-StdUiValue $index 'Index'); Source = $source; Bootstrap = $Ui.Bootstrap }
    Set-StdLoginError -Ui $Ui -Text ('{0}{1}{1}Dostępna jest kopia listy klientów ({2}) - możesz pracować offline.' -f $failure.Message, "`n", (Get-StdUiSourceLabel $source))
    Update-StdLoginOffline -Ui $Ui
    if ($failure.Message -like 'Nieprawidłowy login lub hasło*') { $Ui.C.pwdPassword.Clear(); [void]$Ui.C.pwdPassword.Focus() }
}

function Start-StdLoginOffline {
    # „Pracuj offline”: kopia z nieudanego logowania albo w tle Get-StdClientIndex bez serwera
    # (kopie lustrzane, cache, pendrive).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.Busy) { return }
    if ($null -ne $Ui.OfflineResult) {
        Write-StdLog -Level Warn -Source 'Logowanie' -Message ('Praca offline z kopią listy klientów ({0}).' -f (Get-StdUiSourceLabel ([string]$Ui.OfflineResult.Source)))
        $Ui.Result = $Ui.OfflineResult
        $Ui.Window.Close()
        return
    }
    if ($Ui.Preview) {
        $Ui.Result = @{ Session = $null; Index = (Get-StdUiValue $Ui.PreviewData 'Index'); Source = 'cache'; Bootstrap = $Ui.Bootstrap }
        $Ui.Window.Close()
        return
    }
    $offlineBootstrap = Get-StdUiOfflineBootstrap -Bootstrap $Ui.Bootstrap
    Set-StdLoginError -Ui $Ui -Text ''
    Set-StdLoginBusy -Ui $Ui -Text 'Wczytywanie kopii listy klientów...'
    $null = Start-StdUiTask -Name 'Praca offline' -Owner $Ui.Window -ScriptText $script:StdLoginOfflineScript `
        -Arguments @{ Bootstrap = $offlineBootstrap; UsbRoot = [string]$Ui.OriginRoot } `
        -OnComplete (New-StdUiCallback -Ui $Ui -Handler {
            param($Ui, $Result)
            Set-StdLoginBusy -Ui $Ui -Text ''
            $source = [string](Get-StdUiValue $Result 'Source')
            Write-StdLog -Level Warn -Source 'Logowanie' -Message ('Praca offline z kopią listy klientów ({0}).' -f (Get-StdUiSourceLabel $source))
            $Ui.Result = @{ Session = $null; Index = (Get-StdUiValue $Result 'Index'); Source = $source; Bootstrap = $Ui.Bootstrap }
            $Ui.Window.Close()
        }) `
        -OnError (New-StdUiCallback -Ui $Ui -Handler {
            param($Ui, $ErrorRecord)
            Set-StdLoginBusy -Ui $Ui -Text ''
            Set-StdLoginError -Ui $Ui -Text ('Nie udało się wczytać kopii listy klientów: {0}' -f (Get-StdUiErrorMessage $ErrorRecord -MaxLength 800))
        })
}

# ---------------------------------------------------------------------------------------------
# Okno
# ---------------------------------------------------------------------------------------------

function New-StdLoginWindow {
    <#
    .SYNOPSIS
    Buduje okno logowania (bez pokazywania). Stan okna: Window.Tag ($ui).
    .PARAMETER PreviewData
    Podgląd bez serwera: Server (adres; pusty = pole do wpisania), UserName, Error (tekst błędu), Busy (tekst
    "w toku"), Offline ($true = przycisk „Pracuj offline”), Index (lista klientów zwracana po „Zaloguj”).
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$Bootstrap,
        [AllowNull()][AllowEmptyString()][string]$OriginRoot,
        [AllowNull()][hashtable]$PreviewData
    )
    $win = New-StdWindow -Xaml $script:StdLoginXaml
    $c = Get-StdUiControls -Root $win -Names $script:StdLoginControlNames
    $preview = ($null -ne $PreviewData)
    $ui = @{
        Kind               = 'Login'
        Window             = $win
        C                  = $c
        Bootstrap          = $Bootstrap
        OriginRoot         = $OriginRoot
        Preview            = $preview
        PreviewData        = $null
        ServerEditable     = $false
        LoginBootstrap     = $null
        Busy               = $false
        OfflineResult      = $null
        OfflineIndexExists = $false
        Result             = $null
    }
    $win.Tag = $ui
    $c.txtVersion.Text = 'wersja {0}' -f [string](Get-StdVersion)
    Set-StdUiButton -Button $c.btnLogin -Text 'Zaloguj' -Icon 'E72E'

    $server = $null
    if ($preview) {
        $ui.PreviewData = Merge-StdUiPreviewData -PreviewData $PreviewData
        if ($PreviewData.ContainsKey('Server')) { $server = [string]$PreviewData['Server'] } else { $server = [string](Get-StdUiValue $ui.PreviewData 'Server') }
    } else {
        $server = [string](Get-StdUiValue $Bootstrap 'server')
    }
    if ([string]::IsNullOrWhiteSpace($server)) {
        $ui.ServerEditable = $true
        Set-StdUiVisible $c.boxServer $false
        Set-StdUiVisible $c.txtServer $true
        Set-StdUiVisible $c.txtServerHint $true
    } else {
        $c.txtServerShown.Text = $server
        $c.boxServer.ToolTip = 'Adres z pliku startowego bootstrap.json: {0}' -f $server
    }

    if ($preview) {
        $data = $ui.PreviewData
        $c.txtUser.Text = [string](Get-StdUiValue $data 'UserName')
        Set-StdLoginError -Ui $ui -Text ([string](Get-StdUiValue $PreviewData 'Error'))
        $busyText = [string](Get-StdUiValue $PreviewData 'Busy')
        if (-not [string]::IsNullOrWhiteSpace($busyText)) { Set-StdLoginBusy -Ui $ui -Text $busyText }
        $ui.OfflineIndexExists = ((Get-StdUiValue $PreviewData 'Offline') -eq $true)
    } else {
        $ui.OfflineIndexExists = Test-StdUiOfflineIndex -OriginRoot $OriginRoot
    }
    Update-StdLoginOffline -Ui $ui

    $win.Add_Loaded({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                if ($Ui.ServerEditable -and [string]::IsNullOrWhiteSpace($Ui.C.txtServer.Text)) { [void]$Ui.C.txtServer.Focus() }
                elseif ([string]::IsNullOrWhiteSpace($Ui.C.txtUser.Text)) { [void]$Ui.C.txtUser.Focus() }
                else { [void]$Ui.C.pwdPassword.Focus() }
            }
        })
    $win.Add_KeyDown({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                if ($e.Key -eq [System.Windows.Input.Key]::Escape -and -not $Ui.Busy) { $Ui.Window.Close() }
            }
        })
    $win.Add_Closed({
            param($s, $e)
            Invoke-StdUiHandler $s { param($Ui) try { $Ui.C.pwdPassword.Clear() } catch { } }
        })
    $c.btnLogin.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-StdLoginAttempt -Ui $Ui } })
    $c.btnOffline.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-StdLoginOffline -Ui $Ui } })
    return $win
}

function Show-StdLoginWindow {
    <#
    .SYNOPSIS
    Ekran logowania do serwera STD. Wynik: @{ Session; Index; Source; Bootstrap } (Source = 'server', albo źródło
    kopii przy pracy offline - wtedy Session = $null) albo $null, gdy technik zamknął okno. Bootstrap: plik startowy
    użyty do logowania (także adres wpisany w oknie i zapisany), offline - plik startowy podany oknu (może być $null).
    -PreviewData: podgląd z przykładowymi danymi (bez serwera), jak w New-StdLoginWindow.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$Bootstrap,
        [AllowNull()][AllowEmptyString()][string]$OriginRoot,
        [AllowNull()][hashtable]$PreviewData
    )
    $win = Select-StdUiWindow -Output @(New-StdLoginWindow -Bootstrap $Bootstrap -OriginRoot $OriginRoot -PreviewData $PreviewData)
    [void]$win.ShowDialog()
    $result = $win.Tag.Result
    if ($result -is [hashtable] -and $result.ContainsKey('Index')) { return $result }
    return $null
}
