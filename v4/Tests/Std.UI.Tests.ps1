# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Testy okien STD 4 (SPEC 7 i 9): motyw (palety, New-StdWindow, Set-StdTheme w otwartym oknie, prefs.json),
# każde okno New-Std*Window w obu motywach (zbudowane i rozmieszczone przez Measure/Arrange, bez pokazywania),
# okno główne w każdym stanie podglądu, okno komunikatu (Show-StdMessage z kliknięciem przycisku), powiadomienie
# (Show-StdToast) i zadania w tle (Start-StdUiTask: OnComplete, OnError, log zadania, porzucenie po zamknięciu okna).
# Pester 5, Windows PowerShell 5.1 w wątku STA. Bez administratora i bez sieci; dane lokalne w TestDrive
# ($env:STD_DATA_ROOT). Kody błędów sprawdzamy przez -ExpectedMessage '*`[STD-...`]*' - nawiasy kwadratowe trzeba
# poprzedzić znakiem ` (w -like oznaczają zbiór znaków).
#
# Bloki z oknami WPF działają tylko w Windows i w wątku STA (poza nimi są pomijane). Funkcje ekranów bez WPF
# (teksty, prefs.json, netsh, podsumowanie, dane podglądu) są sprawdzane wszędzie.
#
# Test (nie kod okna) może sam obsłużyć kolejkę zdarzeń okna: Wait-StdTestDispatcher używa DispatcherFrame
# i timera, żeby doczekać wyniku zadania w tle. W kodzie okien to zabronione (SPEC 7.2).

BeforeDiscovery {
    $script:StdUiNoWpf = ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) -or
        ([System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA)

    # Każde okno STD 4 w obu motywach: funkcja (kontrola pokrycia), opis, budowanie, klucz koloru tła okna.
    $script:StdUiWindowSpecs = @(
        @{ Function = 'New-StdWindow'; Case = 'okno bazowe z wspólnymi stylami'; Back = 'ThemeBackground'; Build = { New-StdWindow -Xaml $script:StdUiTestXaml } }
        @{ Function = 'New-StdConnectWindow'; Case = 'Połącz z siecią'; Back = 'ThemeBackground'; Build = { New-StdConnectWindow -OriginRoot $TestDrive } }
        @{ Function = 'New-StdConnectWindow'; Case = 'Połącz z siecią (podgląd)'; Back = 'ThemeBackground'; Build = { New-StdConnectWindow -PreviewData @{ UsbProfiles = 2; Message = 'Test'; MessageKind = 'Warn' } } }
        @{ Function = 'New-StdLoginWindow'; Case = 'Logowanie (adres z bootstrap.json)'; Back = 'ThemeHeader'; Build = { New-StdLoginWindow -Bootstrap ([ordered]@{ server = 'https://std.example.com/std/'; mirrors = @() }) -OriginRoot $TestDrive } }
        @{ Function = 'New-StdLoginWindow'; Case = 'Logowanie (adres do wpisania)'; Back = 'ThemeHeader'; Build = { New-StdLoginWindow -Bootstrap $null -OriginRoot $TestDrive } }
        @{ Function = 'New-StdLoginWindow'; Case = 'Logowanie (podgląd z błędem)'; Back = 'ThemeHeader'; Build = { New-StdLoginWindow -PreviewData @{ Error = 'Nieprawidłowy login lub hasło.'; Offline = $true } } }
        @{ Function = 'New-StdMainWindow'; Case = 'okno główne (lista klientów, bez wyboru)'; Back = 'ThemeBackground'; Build = { New-StdMainWindow -Session $null -Index (Get-StdUiPreviewSample).Index -IndexSource 'cache' -Bootstrap $null -OriginRoot $TestDrive } }
        @{ Function = 'New-StdMainWindow'; Case = 'okno główne (podgląd: Select)'; Back = 'ThemeBackground'; Build = { New-StdMainWindow -PreviewData @{} -PreviewState Select } }
        @{ Function = 'New-StdMainWindow'; Case = 'okno główne (podgląd: Plan)'; Back = 'ThemeBackground'; Build = { New-StdMainWindow -PreviewData @{} -PreviewState Plan } }
        @{ Function = 'New-StdMainWindow'; Case = 'okno główne (podgląd: Running)'; Back = 'ThemeBackground'; Build = { New-StdMainWindow -PreviewData @{} -PreviewState Running } }
        @{ Function = 'New-StdMainWindow'; Case = 'okno główne (podgląd: Paused)'; Back = 'ThemeBackground'; Build = { New-StdMainWindow -PreviewData @{} -PreviewState Paused } }
        @{ Function = 'New-StdMainWindow'; Case = 'okno główne (podgląd: Error)'; Back = 'ThemeBackground'; Build = { New-StdMainWindow -PreviewData @{} -PreviewState Error } }
        @{ Function = 'New-StdMainWindow'; Case = 'okno główne (podgląd: Finished)'; Back = 'ThemeBackground'; Build = { New-StdMainWindow -PreviewData @{} -PreviewState Finished } }
        @{ Function = 'New-StdMessageWindow'; Case = 'komunikat OK, Information'; Back = 'ThemeHeader'; Build = { New-StdMessageWindow -Message 'Treść & "cudzysłów" <znaczniki>' -Title 'Tytuł' -Button OK -Image Information } }
        @{ Function = 'New-StdMessageWindow'; Case = 'komunikat OKCancel, Question'; Back = 'ThemeHeader'; Build = { New-StdMessageWindow -Message 'Pytanie?' -Title 'Pytanie' -Button OKCancel -Image Question } }
        @{ Function = 'New-StdMessageWindow'; Case = 'komunikat YesNo, Warning, Danger'; Back = 'ThemeHeader'; Build = { New-StdMessageWindow -Message 'Przerwać?' -Title 'Ostrzeżenie' -Button YesNo -Image Warning -Danger } }
        @{ Function = 'New-StdMessageWindow'; Case = 'komunikat YesNoCancel, Error'; Back = 'ThemeHeader'; Build = { New-StdMessageWindow -Message ("Błąd`n`n" + ('długi tekst ' * 400)) -Title 'Błąd' -Button YesNoCancel -Image Error } }
        @{ Function = 'New-StdMessageWindow'; Case = 'komunikat OK, Success'; Back = 'ThemeHeader'; Build = { New-StdMessageWindow -Message 'Gotowe.' -Title 'Sukces' -Button OK -Image Success } }
    )
    $script:StdUiWindowCases = foreach ($theme in 'Dark', 'Light') {
        foreach ($spec in $script:StdUiWindowSpecs) { @{ Theme = $theme; Case = $spec.Case; Back = $spec.Back; Build = $spec.Build } }
    }
    $script:StdUiCoveredFunctions = @($script:StdUiWindowSpecs | ForEach-Object { $_.Function } | Sort-Object -Unique)
    $script:StdUiMainStates = foreach ($theme in 'Dark', 'Light') {
        foreach ($state in 'Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished') { @{ Theme = $theme; State = $state } }
    }
}

Describe 'Std.UI' {
    BeforeAll {
        $script:StdUiOldDataRoot = $env:STD_DATA_ROOT
        $env:STD_DATA_ROOT = Join-Path $TestDrive 'data'
        $script:StdUiV4Root = Split-Path -Parent $PSScriptRoot
        Import-Module (Join-Path $script:StdUiV4Root 'STD.psd1') -Force
        $script:StdUiIsWindows = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
        $script:StdUiHasWpf = $script:StdUiIsWindows -and ([System.Threading.Thread]::CurrentThread.ApartmentState -eq [System.Threading.ApartmentState]::STA)
        if ($script:StdUiIsWindows) {
            . (Join-Path (Join-Path $script:StdUiV4Root 'GUI') 'Std.UI.ps1')
        } else {
            # Bez WPF: tylko funkcje bez okien (motyw: palety i ikony z czcionki; ekrany: teksty, dane, prefs.json).
            . (Join-Path (Join-Path $script:StdUiV4Root 'GUI') 'Std.Theme.ps1')
            . (Join-Path (Join-Path (Join-Path $script:StdUiV4Root 'GUI') 'Screens') 'Common.ps1')
        }
        $script:Sample = Get-StdUiPreviewSample

        $script:StdUiTestXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Test" Width="5000" Height="5000" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}">
    <Window.Resources>
        <Style x:Key="LocalCard" TargetType="Border" BasedOn="{StaticResource Card}">
            <Setter Property="Margin" Value="4"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Border Name="card" Style="{StaticResource LocalCard}">
            <StackPanel>
                <TextBlock Name="cap" Text="Podpis" Style="{StaticResource Caption}"/>
                <Button Name="btn" Content="OK" Style="{StaticResource PrimaryButton}"/>
                <Button Name="bar" Content="Edytuj" Style="{StaticResource BarButton}"/>
                <TabControl Name="tabs" Style="{StaticResource NavTabControl}">
                    <TabItem Header="A" Tag="&#xE713;"/>
                </TabControl>
            </StackPanel>
        </Border>
    </Grid>
</Window>
'@

        function ConvertTo-StdTestColor {
            # Kolor z palety (#RRGGBB) w zapisie Color.ToString() (#AARRGGBB).
            param([Parameter(Mandatory = $true)][string]$Key, [ValidateSet('Dark', 'Light')][string]$Theme)
            $hex = Get-StdThemeColor -Key $Key -Theme $Theme
            return ('#FF' + $hex.Substring(1)).ToUpperInvariant()
        }

        function Invoke-StdTestLayout {
            # Układ okna bez pokazywania: Measure/Arrange zawartości w rozmiarze okna. Przy tym powstają szablony
            # kontrolek i wierszy list, więc wychodzą błędy stylów i szablonów (np. brakujący StaticResource).
            param([Parameter(Mandatory = $true)][System.Windows.Window]$Window)
            $content = $Window.Content
            if ($content -isnot [System.Windows.FrameworkElement]) { throw 'Okno nie ma zawartości (Content).' }
            $width = $Window.Width
            if ([double]::IsNaN($width) -or $width -le 0) { $width = 1000 }
            $height = $Window.Height
            $measureHeight = $height
            if ([double]::IsNaN($height) -or $height -le 0) { $measureHeight = [double]::PositiveInfinity }
            [void]$Window.ApplyTemplate()
            $content.Measure((New-Object System.Windows.Size -ArgumentList $width, $measureHeight))
            $arrangeHeight = $height
            if ([double]::IsNaN($arrangeHeight) -or $arrangeHeight -le 0) { $arrangeHeight = $content.DesiredSize.Height }
            $content.Arrange((New-Object System.Windows.Rect -ArgumentList 0, 0, $width, $arrangeHeight))
            $content.UpdateLayout()
            return $content.DesiredSize
        }

        function Wait-StdTestDispatcher {
            # Obsługuje kolejkę zdarzeń okna (tylko w teście), aż $Until da $true albo minie limit czasu.
            param([Parameter(Mandatory = $true)][scriptblock]$Until, [int]$TimeoutSec = 60)
            $frame = New-Object System.Windows.Threading.DispatcherFrame
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromMilliseconds(50)
            $timer.Tag = @{ Frame = $frame; Until = $Until; Deadline = (Get-Date).AddSeconds($TimeoutSec) }
            $timer.Add_Tick({
                    param($s, $e)
                    $t = $s.Tag
                    $done = $true
                    try { $done = [bool](& $t.Until) -or (Get-Date) -gt $t.Deadline } catch { $done = $true }
                    if ($done) { $s.Stop(); $t.Frame.Continue = $false }
                })
            $timer.Start()
            [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        }

        function Show-StdTestMessage {
            # Show-StdMessage z "kliknięciem": timer znajduje okno komunikatu po tytule i klika przycisk z wynikiem
            # -Click (bez -Click: zamyka okno krzyżykiem). Po 30 s zamyka każde widoczne okno - test się nie zawiesi.
            param([Parameter(Mandatory = $true)][hashtable]$Params, $Click)
            $closer = New-Object System.Windows.Threading.DispatcherTimer
            $closer.Interval = [TimeSpan]::FromMilliseconds(100)
            $closer.Tag = @{ Title = [string]$Params.Title; Click = $Click; Ticks = 0; Seen = $false; Buttons = @() }
            $closer.Add_Tick({
                    param($s, $e)
                    $t = $s.Tag
                    $t.Ticks++
                    try {
                        foreach ($w in (Get-StdOpenWindows)) {
                            if (-not $w.IsVisible -or $w.Title -ne $t.Title) { continue }
                            $s.Stop()
                            $t.Seen = $true
                            $panel = $w.FindName('spButtons')
                            $t.Buttons = @(foreach ($b in $panel.Children) { [string]$b.Content })
                            if ($null -ne $t.Click) {
                                foreach ($b in $panel.Children) {
                                    if ($b.Tag -eq $t.Click) {
                                        $b.RaiseEvent((New-Object System.Windows.RoutedEventArgs -ArgumentList ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
                                        return
                                    }
                                }
                            }
                            $w.Close()
                            return
                        }
                        if ($t.Ticks -gt 300) { throw 'Okno komunikatu się nie pokazało.' }
                    } catch {
                        $s.Stop()
                        try { foreach ($w in [System.Windows.Application]::Current.Windows) { if ($w.IsVisible) { $w.Close() } } } catch { }
                    }
                })
            $closer.Start()
            try {
                $result = Show-StdMessage @Params
            } finally {
                $closer.Stop()
            }
            return @{ Result = $result; Seen = $closer.Tag.Seen; Buttons = $closer.Tag.Buttons }
        }
    }

    AfterAll {
        if ($script:StdUiHasWpf) { try { Set-StdTheme -Dark $true -NoSave } catch { } }
        $env:STD_DATA_ROOT = $script:StdUiOldDataRoot
    }

    # ---------------------------------------------------------------------------------------------
    # Bez WPF (każdy system)
    # ---------------------------------------------------------------------------------------------

    Context 'Okna STD - funkcje bez WPF' {
        It 'palety Dark i Light mają te same klucze w tej samej kolejności' {
            (@($script:StdThemePalettes.Dark.Keys) -join ',') | Should -Be (@($script:StdThemePalettes.Light.Keys) -join ',')
            @($script:StdThemePalettes.Dark.Keys).Count | Should -BeGreaterThan 40
            foreach ($tone in 'Info', 'Ok', 'Warn', 'Crit') {
                $script:StdThemePalettes.Dark.Contains("Theme$($tone)Back") | Should -BeTrue
                $script:StdThemePalettes.Dark.Contains("Theme$($tone)Fore") | Should -BeTrue
            }
        }

        It 'Get-StdThemeColor: kolor z palety i błąd dla nieznanego klucza' {
            Get-StdThemeColor -Key 'ThemeBackground' -Theme Light | Should -Be '#F4F6F9'
            { Get-StdThemeColor -Key 'BrakTakiegoKoloru' } | Should -Throw -ExpectedMessage '*`[STD-UI-THEME`]*'
        }

        It 'Get-StdGlyph: znak z kodu i błąd dla złego kodu' {
            Get-StdGlyph -Code 'E7B8' | Should -Be ([string][char]0xE7B8)
            Get-StdGlyph -Code '' | Should -Be ''
            { Get-StdGlyph -Code 'xyz' } | Should -Throw -ExpectedMessage '*`[STD-UI-GLYPH`]*'
        }

        It 'stany kroków i planu mają ikonę, napis i kolor z palety' {
            $palette = Get-StdThemePalette -Theme Dark
            foreach ($state in 'Pending', 'Running', 'Ok', 'Skipped', 'Error', 'Cancelled', 'Done', 'Todo', 'Blocked', 'Disabled') {
                $info = Get-StdUiStateInfo -State $state
                $info.State | Should -Be $state
                $info.Glyph | Should -Match '^E[0-9A-F]{3}$'
                $palette.Contains($info.ColorKey) | Should -BeTrue
            }
            (Get-StdUiStateInfo -State 'Nieznany').State | Should -Be 'Pending'
        }

        It 'wiersz dziennika jest zamaskowany i skrócony' {
            Register-StdLogMask -Value 'TajneHaslo987'
            $item = New-StdUiLogItem -Level 'warn' -Source 'Std.Engine' -Message 'hasło TajneHaslo987'
            $item.Message | Should -Be 'hasło ***'
            $item.Level | Should -Be 'Warn'
            (New-StdUiLogItem -Level Info -Message ('x' * 5000)).Message.Length | Should -Be 4003
        }

        It 'ostatnio używane: kolejność, limit i zachowanie motywu w prefs.json' {
            Write-StdJsonFile -Path (Get-StdPath -Kind Prefs) -InputObject ([ordered]@{ theme = 'Light' })
            Save-StdUiRecent -ClientId 'lizard' -SiteId 'krakow' -ProfileName 'Standard'
            Save-StdUiRecent -ClientId 'inny' -SiteId 'x' -ProfileName 'P'
            (Get-StdUiRecentIds -Kind clients) -join ',' | Should -Be 'inny,lizard'
            (Get-StdUiRecentIds -Kind profiles -ClientId lizard -SiteId krakow) -join ',' | Should -Be 'Standard'
            (Get-StdUiPrefs)['theme'] | Should -Be 'Light'
            $items = @((New-StdUiPickItem -Id 'a' -Name 'A'), (New-StdUiPickItem -Id 'lizard' -Name 'Lizard'))
            $sorted = Sort-StdUiByRecent -Items $items -Recent (Get-StdUiRecentIds -Kind clients)
            $sorted[0].Id | Should -Be 'lizard'
            $sorted[0].Hint | Should -Be 'ostatnio'
            Remove-Item -LiteralPath (Get-StdPath -Kind Prefs) -Force
        }

        It 'netsh: polska lista sieci (najsilniejszy sygnał pierwszy, ukryte pominięte)' {
            $text = "SSID 1 : Dom`r`n    Uwierzytelnianie : WPA2-Personal`r`n         Sygnał : 70%`r`nSSID 2 :`r`n    Uwierzytelnianie : WPA2-Personal`r`nSSID 3 : Kawiarnia`r`n    Uwierzytelnianie : Otwarte`r`n         Sygnał : 90%"
            $networks = ConvertFrom-StdNetshNetworks -Text $text
            $networks.Count | Should -Be 2
            $networks[0].Ssid | Should -Be 'Kawiarnia'
            $networks[0].Auth | Should -Be 'open'
            $networks[1].Auth | Should -Be 'WPA2PSK'
        }

        It 'profil Wi-Fi: poprawny XML i odrzucenie za krótkiego hasła' {
            $doc = New-Object System.Xml.XmlDocument
            $doc.LoadXml((New-StdWifiProfileXml -Ssid 'A&B' -Auth WPA2PSK -Key 'haslo&1234'))
            $doc.DocumentElement.Item('name').InnerText | Should -Be 'A&B'
            { New-StdWifiProfileXml -Ssid 'X' -Auth WPA2PSK -Key 'krotkie' } | Should -Throw -ExpectedMessage '*`[STD-UI-WIFI`]*'
        }

        It 'błąd logowania: 401 z serwera i kopia offline' {
            $result = @{ Source = 'cache'; Attempts = @(@{ Source = 'server'; Error = '[STD-HTTP-401] GET https://x/std/clients.sdt' }, @{ Source = 'cache'; Error = $null }) }
            $failure = Get-StdUiLoginFailure -IndexResult $result
            $failure.Message | Should -Be 'Nieprawidłowy login lub hasło.'
            $failure.OfflineSource | Should -Be 'cache'
        }

        It 'podsumowanie wdrożenia: wynik, kroki, bez sekretów' {
            Register-StdLogMask -Value 'SekretX12345'
            $steps = @((New-StdUiStepItem -Id 'a.a' -Title 'Krok A' -State Ok), (New-StdUiStepItem -Id 'a.b' -Title 'Krok B' -State Error -Detail 'SekretX12345'))
            $text = Format-StdUiRunSummary -Data @{ Version = '4.0.0'; DeploymentId = 'd1'; Status = 'Finished'; Steps = $steps }
            $text | Should -Match 'Wynik: Wdrożenie zakończone z błędami'
            $text | Should -Match '\[Błąd\] Krok B - \*\*\*'
            $text | Should -Not -Match 'SekretX12345'
            (Get-StdUiFailedStepIds -Steps $steps) -join ',' | Should -Be 'a.b'
        }

        It 'przykładowe dane podglądu dają konfigurację wynikową' {
            $eff = Resolve-StdEffectiveConfig -Release $script:Sample.Release -SiteId $script:Sample.SiteId -ProfileName $script:Sample.ProfileName
            (Get-StdUiSummaryRows -Effective $eff).Count | Should -BeGreaterThan 5
            (Get-StdUiAppChips -Effective $eff).Count | Should -Be 5
            (Get-StdUiTaskChips -Effective $eff).Count | Should -Be 16
            (Get-StdUiComputerRows -Info $script:Sample.ComputerInfo).Count | Should -Be 10
        }

        It 'dane podglądu dla stanu <_>: kroki, postęp i dziennik' -ForEach @('Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished') {
            $data = Merge-StdUiPreviewData -PreviewData @{ UserName = 'tester' } -State $_
            $data.UserName | Should -Be 'tester'
            $data.ClientId | Should -Be 'lizard'
            @($data.Plan).Count | Should -Be 5
            @($data.Log).Count | Should -BeGreaterThan 2
            if ($_ -in @('Running', 'Paused', 'Error', 'Finished')) { @($data.Steps).Count | Should -Be 4 } else { @($data.Steps).Count | Should -Be 0 }
            if ($_ -in @('Error', 'Finished')) { $data.Result | Should -Not -BeNullOrEmpty }
        }

        It 'podgląd okien przekazuje dane przez -PreviewData / -PreviewState (narzędzie i ekrany)' {
            $preview = Join-Path (Join-Path $script:StdUiV4Root 'Tools') 'Render-StdUiPreview.ps1'
            Test-Path -LiteralPath $preview -PathType Leaf | Should -BeTrue
            $text = [System.IO.File]::ReadAllText($preview)
            foreach ($screen in 'Show-StdConnectWindow', 'Show-StdLoginWindow', 'Show-StdMainWindow', 'Show-StdMessage', 'Show-StdToast') { $text | Should -Match ([regex]::Escape($screen)) }
            foreach ($state in 'Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished') { $text | Should -Match ('-PreviewState ' + $state) }
            foreach ($check in 'CLIP', 'STALE', 'GLYPH', 'EMOJI', 'CONTRAST') { $text | Should -Match $check }
        }
    }

    # ---------------------------------------------------------------------------------------------
    # Motyw (WPF)
    # ---------------------------------------------------------------------------------------------

    Context 'Std.Theme - okna w motywie' -Skip:$script:StdUiNoWpf {
        AfterEach { Set-StdTheme -Dark $true -NoSave }

        It 'buduje okno w motywie <_> (style, pędzle, czcionka, rozmiar)' -ForEach @('Dark', 'Light') {
            Set-StdTheme -Dark ($_ -eq 'Dark') -NoSave
            $w = New-StdWindow -Xaml $script:StdUiTestXaml
            try {
                $w | Should -BeOfType ([System.Windows.Window])
                $w.FindName('btn').Style | Should -Be ($w.FindResource('PrimaryButton'))
                $w.Resources['ThemeBackground'].Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeBackground' -Theme $_)
                $w.Background.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeBackground' -Theme $_)
                $w.FontSize | Should -Be 13
                $w.FontFamily.Source | Should -Be 'Segoe UI'
                $w.UseLayoutRounding | Should -BeTrue
                $w.Width | Should -BeLessOrEqual ([System.Windows.SystemParameters]::WorkArea.Width)
                $w.Height | Should -BeLessOrEqual ([System.Windows.SystemParameters]::WorkArea.Height)
                $w.WindowStartupLocation | Should -Be ([System.Windows.WindowStartupLocation]::CenterScreen)
                $w.Icon | Should -Not -BeNullOrEmpty
                (Get-StdOpenWindows).Contains($w) | Should -BeTrue
            } finally { $w.Close() }
            (Get-StdOpenWindows).Contains($w) | Should -BeFalse
        }

        It 'Set-StdTheme przemalowuje otwarte okno (także kolory z kodu) i zapisuje motyw w prefs.json' {
            Set-StdTheme -Dark $true -NoSave
            $w = New-StdMainWindow -PreviewData @{} -PreviewState Running
            $w.ShowActivated = $false
            $w.ShowInTaskbar = $false
            $w.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
            $w.Left = -20000
            $w.Top = -20000
            try {
                $w.Show()
                $w.UpdateLayout()
                $c = $w.Tag.C
                $w.Background.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeBackground' -Theme Dark)
                $c.txtStatus.Foreground.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeText' -Theme Dark)
                $c.dotStatus.Fill.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeInfo' -Theme Dark)

                Set-StdTheme -Dark $false
                $w.UpdateLayout()
                Test-StdThemeDark | Should -BeFalse
                $w.Resources['ThemeBackground'].Color.ToString() | Should -Be '#FFF4F6F9'
                $w.Background.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeBackground' -Theme Light)
                # Kolor z XAML ({DynamicResource}) i kolor ustawiony w kodzie (Set-StdUiColor = SetResourceReference).
                $c.txtStatus.Foreground.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeText' -Theme Light)
                $c.dotStatus.Fill.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeInfo' -Theme Light)
                (Read-StdJsonFile -Path (Get-StdPath -Kind Prefs)).theme | Should -Be 'Light'
                Initialize-StdTheme | Should -BeFalse

                Set-StdTheme -Dark $true
                $w.UpdateLayout()
                $c.txtStatus.Foreground.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeText' -Theme Dark)
                (Read-StdJsonFile -Path (Get-StdPath -Kind Prefs)).theme | Should -Be 'Dark'
            } finally {
                $w.Close()
                Set-StdTheme -Dark $true -NoSave
            }
        }

        It 'okno komunikatu: przyciski w kolejności, główny po prawej, Esc = Anuluj (<_>)' -ForEach @('OK', 'OKCancel', 'YesNo', 'YesNoCancel') {
            $d = New-StdMessageWindow -Message 'Treść & "cudzysłów"' -Title 'Tytuł' -Button $_ -Image Warning
            try {
                $d.FindName('txtMessage').Text | Should -Be 'Treść & "cudzysłów"'
                $d.Title | Should -Be 'Tytuł'
                $buttons = @($d.FindName('spButtons').Children)
                $buttons.Count | Should -Be (@{ OK = 1; OKCancel = 2; YesNo = 2; YesNoCancel = 3 }[$_])
                $buttons[$buttons.Count - 1].IsDefault | Should -BeTrue
                @($buttons | Where-Object { $_.IsCancel }).Count | Should -Be 1
                $buttons[$buttons.Count - 1].Style | Should -Be ($d.FindResource('PrimaryButton'))
            } finally { $d.Close() }
        }

        It 'okno komunikatu -Danger: główny przycisk czerwony' {
            $d = New-StdMessageWindow -Message 'Przerwać?' -Title 'Przerwać' -Button YesNo -Image Warning -Danger
            try {
                $buttons = @($d.FindName('spButtons').Children)
                $buttons[$buttons.Count - 1].Style | Should -Be ($d.FindResource('DangerPrimaryButton'))
            } finally { $d.Close() }
        }

        It 'Show-StdMessage pokazuje okno i zwraca kliknięty przycisk albo domyślny po zamknięciu' {
            $yes = Show-StdTestMessage -Params @{ Message = 'Rozpocząć wdrożenie?'; Title = 'Test Show-StdMessage 1'; Button = 'YesNo'; Image = 'Question' } -Click ([System.Windows.MessageBoxResult]::Yes)
            $yes.Seen | Should -BeTrue
            $yes.Result | Should -Be ([System.Windows.MessageBoxResult]::Yes)
            ($yes.Buttons -join ',') | Should -Be 'Nie,Tak'

            $closed = Show-StdTestMessage -Params @{ Message = 'Zamknięcie krzyżykiem'; Title = 'Test Show-StdMessage 2'; Button = 'YesNoCancel'; Image = 'Warning' }
            $closed.Seen | Should -BeTrue
            $closed.Result | Should -Be ([System.Windows.MessageBoxResult]::Cancel)
            ($closed.Buttons -join ',') | Should -Be 'Anuluj,Nie,Tak'

            $ok = Show-StdTestMessage -Params @{ Message = 'Gotowe.'; Title = 'Test Show-StdMessage 3'; Image = 'Success' } -Click ([System.Windows.MessageBoxResult]::OK)
            $ok.Seen | Should -BeTrue
            $ok.Result | Should -Be ([System.Windows.MessageBoxResult]::OK)
        }

        It 'Show-StdToast dodaje powiadomienie w kolorach motywu' {
            $w = New-StdMainWindow -PreviewData @{} -PreviewState Select
            try {
                $toastHost = $w.FindName('toastHost')
                $before = $toastHost.Children.Count
                Show-StdToast -Window $w -Message 'Skopiowano do schowka.' -Kind Success -Seconds 2
                Show-StdToast -Window $w -Message 'Błąd' -Kind Error -Seconds 2
                $toastHost.Children.Count | Should -Be ($before + 2)
                $toast = $toastHost.Children[$toastHost.Children.Count - 1]
                $toast.Background.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeToast' -Theme Dark)
            } finally { $w.Close() }
        }

        It 'ikona STD: plik .ico z wieloma rozmiarami i BitmapFrame' {
            $bytes = New-StdAppIconBytes
            ($bytes[0..3] -join ',') | Should -Be '0,0,1,0'
            ([int]$bytes[4] + 256 * [int]$bytes[5]) | Should -BeGreaterThan 1
            (New-StdAppIconBytes -DrawnMark)[2] | Should -Be 1
            New-StdAppIcon | Should -BeOfType ([System.Windows.Media.Imaging.BitmapFrame])
        }

        It 'zły XAML daje STD-UI-XAML' {
            { New-StdWindow -Xaml '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><Border Style="{StaticResource BrakStylu}"/></Window>' } |
                Should -Throw -ExpectedMessage '*`[STD-UI-XAML`]*'
            { New-StdWindow -Xaml '<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"/>' } |
                Should -Throw -ExpectedMessage '*`[STD-UI-XAML`]*'
        }
    }

    # ---------------------------------------------------------------------------------------------
    # Wszystkie okna STD 4 w obu motywach (WPF)
    # ---------------------------------------------------------------------------------------------

    Context 'Okna STD 4 - budowanie i układ' -Skip:$script:StdUiNoWpf {
        AfterEach { Set-StdTheme -Dark $true -NoSave }

        # Lista z BeforeDiscovery trafia do testu przez -ForEach: zmienne z fazy odkrywania nie istnieją w fazie uruchamiania.
        It 'każda funkcja New-Std*Window ma przypadek w tym teście' -ForEach @(@{ Covered = $script:StdUiCoveredFunctions }) {
            $Covered.Count | Should -BeGreaterThan 4
            $functions = @(Get-Command -Name 'New-Std*Window' -CommandType Function | ForEach-Object { $_.Name } | Sort-Object -Unique)
            $functions.Count | Should -BeGreaterThan 4
            $missing = @($functions | Where-Object { $Covered -notcontains $_ })
            ($missing -join ', ') | Should -Be ''
        }

        It 'okno: <Case> - buduje się w motywie <Theme> (Measure/Arrange)' -ForEach $script:StdUiWindowCases {
            Set-StdTheme -Dark ($Theme -eq 'Dark') -NoSave
            $win = Select-StdUiWindow -Output @(& $Build)
            try {
                $size = Invoke-StdTestLayout -Window $win
                $size.Width | Should -BeGreaterThan 0
                $size.Height | Should -BeGreaterThan 0
                $win.Resources['ThemeBackground'].Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key 'ThemeBackground' -Theme $Theme)
                $win.Background.Color.ToString() | Should -Be (ConvertTo-StdTestColor -Key $Back -Theme $Theme)
                if ($win.Tag -is [hashtable] -and $win.Tag.ContainsKey('Window')) {
                    [object]::ReferenceEquals($win.Tag.Window, $win) | Should -BeTrue
                    $win.Tag.C.Count | Should -BeGreaterThan 3
                }
            } finally { $win.Close() }
        }

        It 'Połącz z siecią: sieci z podglądu, import z pendrive''a, komunikat' {
            $win = New-StdConnectWindow -PreviewData @{ UsbProfiles = 2; Message = 'Test'; MessageKind = 'Warn' }
            try {
                $win.Tag.Kind | Should -Be 'Connect'
                $win.Tag.Result | Should -BeNullOrEmpty
                $win.Tag.C.lstNetworks.Items.Count | Should -Be 4
                $win.Tag.C.btnImport.IsEnabled | Should -BeTrue
                $win.Tag.C.boxMessage.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
            } finally { $win.Close() }
            $real = New-StdConnectWindow -OriginRoot $TestDrive
            try {
                $real.Tag.Preview | Should -BeFalse
                $real.Tag.C.btnImport.IsEnabled | Should -BeFalse
                $null -ne $real.Tag.CheckTimer | Should -BeTrue
                $real.Tag.CheckTimer.IsEnabled | Should -BeFalse
            } finally { $real.Close() }
        }

        It 'Logowanie: adres z bootstrap.json albo pole do wpisania, błąd i praca offline' {
            $withServer = New-StdLoginWindow -Bootstrap ([ordered]@{ server = 'https://std.example.com/std/'; mirrors = @() }) -OriginRoot $TestDrive
            try {
                $withServer.Tag.ServerEditable | Should -BeFalse
                $withServer.Tag.C.txtServerShown.Text | Should -Be 'https://std.example.com/std/'
            } finally { $withServer.Close() }
            $noServer = New-StdLoginWindow -Bootstrap $null -OriginRoot $TestDrive
            try {
                $noServer.Tag.ServerEditable | Should -BeTrue
                $noServer.Tag.C.txtServer.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
            } finally { $noServer.Close() }
            $preview = New-StdLoginWindow -PreviewData @{ Error = 'Nieprawidłowy login lub hasło.'; Offline = $true }
            try {
                $preview.Tag.C.boxError.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
                $preview.Tag.C.btnOffline.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
            } finally { $preview.Close() }
        }

        It 'okno główne: lista klientów bez sieci i bez silnika' {
            $win = New-StdMainWindow -Session $null -Index $script:Sample.Index -IndexSource 'cache' -Bootstrap $null -OriginRoot $TestDrive
            try {
                $ui = $win.Tag
                $ui.State | Should -Be 'Select'
                @($ui.ClientItems).Count | Should -Be 3
                $ui.C.btnDeploy.IsEnabled | Should -BeFalse
                $ui.C.txtUser.Text | Should -Be 'offline'
                $ui.Engine | Should -BeNullOrEmpty
                $ui.Timer.IsEnabled | Should -BeFalse
            } finally { $win.Close() }
        }

        It 'okno główne: podgląd stanu <State> w motywie <Theme>' -ForEach $script:StdUiMainStates {
            Set-StdTheme -Dark ($Theme -eq 'Dark') -NoSave
            $win = New-StdMainWindow -PreviewData @{} -PreviewState $State
            try {
                $ui = $win.Tag
                $ui.State | Should -Be $State
                $ui.Effective | Should -Not -BeNullOrEmpty
                $run = $State -in @('Running', 'Paused', 'Error', 'Finished')
                ($ui.C.viewRun.Visibility -eq [System.Windows.Visibility]::Visible) | Should -Be $run
                ($ui.C.viewSelect.Visibility -eq [System.Windows.Visibility]::Visible) | Should -Be (-not $run)
                ($ui.C.btnPause.Visibility -eq [System.Windows.Visibility]::Visible) | Should -Be ($State -in @('Running', 'Paused'))
                ($ui.C.btnRetry.Visibility -eq [System.Windows.Visibility]::Visible) | Should -Be ($State -eq 'Error')
                ($ui.C.btnCopy.Visibility -eq [System.Windows.Visibility]::Visible) | Should -Be ($State -in @('Error', 'Finished'))
                if ($run) { $ui.StepItems.Count | Should -Be 4 }
                if ($State -eq 'Plan') { $ui.C.tabSelect.SelectedIndex | Should -Be 1; $ui.PlanItems.Count | Should -Be 4 }
                if ($State -eq 'Running') { $ui.C.btnTheme.IsEnabled | Should -BeFalse }
                $ui.LogItems.Count | Should -BeGreaterThan 0
                [void](Invoke-StdTestLayout -Window $win)
            } finally { $win.Close() }
        }

        It 'stan okna znajduje się po kontrolce (handlery)' {
            $win = New-StdMainWindow -PreviewData @{}
            try {
                (Get-StdUiState -EventSource $win.Tag.C.btnCheck).Kind | Should -Be 'Main'
                (Get-StdUiState -EventSource $win).Kind | Should -Be 'Main'
            } finally { $win.Close() }
        }
    }

    # ---------------------------------------------------------------------------------------------
    # Zadania w tle (WPF: DispatcherTimer w wątku okna)
    # ---------------------------------------------------------------------------------------------

    Context 'Start-StdUiTask' -Skip:$script:StdUiNoWpf {
        It 'woła OnComplete z wynikiem w wątku okna (i przekazuje log zadania)' {
            $stdTestResult = @{ Value = $null; Thread = $null; Calls = 0 }
            $stdTestLog = New-Object 'System.Collections.Generic.List[string]'
            $stdTestSink = { $stdTestLog.Add([string]$args[0].Message) }.GetNewClosure()
            Register-StdLogSink -Name 'StdUiTest' -ScriptBlock $stdTestSink
            try {
                $uiThread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
                $task = Start-StdUiTask -Name 'test' -Arguments @{ A = 20 } -ScriptText 'Write-StdLog -Message "z zadania w tle"; $Arguments.A + [int](Get-StdVersion).Major' -OnComplete {
                    param($Result)
                    $stdTestResult.Value = $Result
                    $stdTestResult.Thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
                    $stdTestResult.Calls++
                }
                $task | Should -Not -BeNullOrEmpty
                Wait-StdTestDispatcher -Until { $null -ne $stdTestResult.Value }
                $stdTestResult.Value | Should -Be (20 + [int](Get-StdVersion).Major)
                $stdTestResult.Thread | Should -Be $uiThread
                $stdTestResult.Calls | Should -Be 1
                $task.Status | Should -Be 'Completed'
                $task.Finished | Should -BeTrue
                $stdTestLog | Should -Contain 'z zadania w tle'
            } finally {
                Unregister-StdLogSink -Name 'StdUiTest'
            }
        }

        It 'woła OnError z ErrorRecord' {
            $stdTestResult = @{ Error = $null; Completed = $false }
            [void](Start-StdUiTask -ScriptText 'throw "[STD-TEST] zle"' -OnComplete { $stdTestResult.Completed = $true } -OnError { param($ErrorRecord) $stdTestResult.Error = $ErrorRecord })
            Wait-StdTestDispatcher -Until { $null -ne $stdTestResult.Error }
            $stdTestResult.Error | Should -BeOfType ([System.Management.Automation.ErrorRecord])
            $stdTestResult.Error.Exception.Message | Should -Be '[STD-TEST] zle'
            $stdTestResult.Completed | Should -BeFalse
        }

        It 'skrypt z param($Arguments) dostaje argumenty' {
            $stdTestResult = @{ Value = $null }
            [void](Start-StdUiTask -ScriptText 'param($Arguments) "{0}-{1}" -f $Arguments.X, $Arguments.Y' -Arguments @{ X = 'a'; Y = 2 } -OnComplete { param($Result) $stdTestResult.Value = $Result })
            Wait-StdTestDispatcher -Until { $null -ne $stdTestResult.Value }
            $stdTestResult.Value | Should -Be 'a-2'
        }

        It 'po zamknięciu okna właściciela wynik jest pomijany' {
            $stdTestResult = @{ Completed = $false; Failed = $false }
            $owner = New-StdWindow -Xaml $script:StdUiTestXaml
            $task = Start-StdUiTask -Name 'porzucone' -Owner $owner -ScriptText 'Start-Sleep -Milliseconds 1500; "wynik"' `
                -OnComplete { $stdTestResult.Completed = $true } -OnError { $stdTestResult.Failed = $true }
            $owner.Close()
            $task.OwnerClosed | Should -BeTrue
            Wait-StdTestDispatcher -Until { $task.Finished }
            $task.Finished | Should -BeTrue
            $task.Abandoned | Should -BeTrue
            $stdTestResult.Completed | Should -BeFalse
            $stdTestResult.Failed | Should -BeFalse
        }
    }
}
