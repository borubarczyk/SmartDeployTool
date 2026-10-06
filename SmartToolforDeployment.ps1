
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

# ------------------------------------------
# Motyw graficzny (ciemny / jasny)
# ------------------------------------------
# Jedna paleta i jeden zestaw stylów dla WSZYSTKICH okien narzędzia - ta sama rodzina kolorów
# i kontrolek co w ServerReview / NPS Event Viewer (grafitowo-granatowe tło, niebieski akcent,
# zaokrąglone pola, ciemne paski przewijania). Wcześniej każde okno miało własną, nieco inną
# kopię stylów, a paski przewijania, zaznaczenie na listach, CheckBoxy, podpowiedzi (ToolTip),
# menu kontekstowe i nagłówki tabel zostawały w jasnym stylu Windows także w trybie ciemnym
# (np. jasnoniebieskie zaznaczenie z białym tekstem - nieczytelne).
$script:ThemePalettes = @{
    Dark  = [ordered]@{
        ThemeBackground  = '#0F1318'   # tło okna
        ThemeHeader      = '#12171D'   # pasek nagłówka/stopki, nagłówki tabel
        ThemePanel       = '#161B22'   # karty / panele
        ThemeBorder      = '#242B36'   # obramowanie kart
        ThemeRowLine     = '#1F2530'   # linie między wierszami tabel
        ThemeTextBoxBg   = '#1B212A'   # pola edycyjne, "chipy"
        ThemeFieldBorder = '#2A323F'   # obramowanie pól i przycisków
        ThemeButton      = '#1B212A'
        ThemeButtonText  = '#E4E8EF'
        ThemeText        = '#E4E8EF'
        ThemeMuted       = '#8791A5'   # podpisy, opisy
        ThemeFaint       = '#5E6779'
        ThemeHover       = '#1D2430'
        ThemeSelection   = '#233354'
        ThemeScrollThumb = '#39414F'
        ThemeTrack       = '#262D39'   # tło pasków postępu
        ThemeOverlay     = '#FFFFFF'   # rozjaśnienie przycisku po najechaniu
        ThemeOnAccent    = '#FFFFFF'   # tekst na kolorowych przyciskach
        AccentColor      = '#3E6FE0'
        ThemeFocus       = '#4C7DF0'
        ThemeAccentText  = '#7DA6FF'
        StartColor       = '#168A5E'
        ThemeSuccessFill = '#168A5E'
        ThemeWarningFill = '#B7770B'
        ThemeDangerFill  = '#D63B4B'
        ThemeSuccess     = '#4FD1A1'   # kolory statusów (tekst, kropki)
        ThemeWarning     = '#FFC46B'
        ThemeDanger      = '#FF7A86'
        ThemeInfo        = '#8CC0FF'
        ThemeDryRun      = '#C3A6FF'
    }
    Light = [ordered]@{
        ThemeBackground  = '#F4F6F9'
        ThemeHeader      = '#FFFFFF'
        ThemePanel       = '#FFFFFF'
        ThemeBorder      = '#E2E6EE'
        ThemeRowLine     = '#EEF1F5'
        ThemeTextBoxBg   = '#FFFFFF'
        ThemeFieldBorder = '#CDD3DE'
        ThemeButton      = '#F7F8FA'
        ThemeButtonText  = '#1A1F29'
        ThemeText        = '#1A1F29'
        ThemeMuted       = '#667085'
        ThemeFaint       = '#98A2B3'
        ThemeHover       = '#EEF2F8'
        ThemeSelection   = '#DCE6FB'
        ThemeScrollThumb = '#C3CAD6'
        ThemeTrack       = '#E2E6EE'
        ThemeOverlay     = '#000000'
        ThemeOnAccent    = '#FFFFFF'
        AccentColor      = '#3E6FE0'
        ThemeFocus       = '#3E6FE0'
        ThemeAccentText  = '#2F5FD0'
        StartColor       = '#138A5E'
        ThemeSuccessFill = '#138A5E'
        ThemeWarningFill = '#B86E00'
        ThemeDangerFill  = '#D92D3F'
        ThemeSuccess     = '#138A5E'
        ThemeWarning     = '#B86E00'
        ThemeDanger      = '#D92D3F'
        ThemeInfo        = '#2463C9'
        ThemeDryRun      = '#7A3EC8'
    }
}
# Systemowe pędzle Windows, z których korzystają domyślne szablony kontrolek (np. róg między
# paskami przewijania w ScrollViewerze był jasnoszarym kwadratem w trybie ciemnym).
# UWAGA: bez InactiveSelectionHighlight(Text)BrushKey - w .NET Framework są to aliasy
# ControlBrushKey/ControlTextBrushKey, więc w słowniku pojawiał się zdublowany klucz
# ("Item has already been added. Key in dictionary: 'ControlBrush'") i żadne okno się nie wczytywało.
$script:ThemeSystemBrushKeys = [ordered]@{
    ControlBrushKey       = 'ThemePanel'
    WindowBrushKey        = 'ThemeTextBoxBg'
    ControlTextBrushKey   = 'ThemeText'
    WindowTextBrushKey    = 'ThemeText'
    GrayTextBrushKey      = 'ThemeFaint'
    HighlightBrushKey     = 'AccentColor'
    HighlightTextBrushKey = 'ThemeOnAccent'
}
$script:isDarkTheme = $true

function global:Get-ThemePalette {
    if ($script:isDarkTheme) { return $script:ThemePalettes.Dark }
    return $script:ThemePalettes.Light
}

# Kolor (HEX) z aktualnej palety - dla miejsc, w których kolor trafia do danych (np. kolor
# wiersza logu albo statusu deinstalacji), a nie do XAML.
function global:Get-ThemeColor {
    param([string]$Name)
    return (Get-ThemePalette)[$Name]
}

function global:New-ThemeBrush {
    param([string]$Color)
    $brush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}

$script:ThemeStylesXaml = @'
    <!-- ===== Teksty i karty ===== -->
    <Style x:Key="Caption" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="FontSize" Value="11.5"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Margin" Value="0,0,0,5"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="MutedText" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="CardTitle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="FontSize" Value="14"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>
    <Style x:Key="Card" TargetType="Border">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="CornerRadius" Value="10"/>
        <Setter Property="Padding" Value="16,14"/>
    </Style>
    <Style x:Key="Chip" TargetType="Border">
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="CornerRadius" Value="7"/>
        <Setter Property="Padding" Value="10,5"/>
    </Style>

    <!-- ===== Paski przewijania ===== -->
    <Style x:Key="ScrollThumbStyle" TargetType="Thumb">
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Thumb">
                    <Border x:Name="t" CornerRadius="4" Background="{DynamicResource ThemeScrollThumb}"/>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="t" Property="Background" Value="{DynamicResource ThemeFaint}"/></Trigger>
                        <Trigger Property="IsDragging" Value="True"><Setter TargetName="t" Property="Background" Value="{DynamicResource ThemeMuted}"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style x:Key="ScrollPageButton" TargetType="RepeatButton">
        <Setter Property="Focusable" Value="False"/>
        <Setter Property="IsTabStop" Value="False"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="RepeatButton"><Border Background="Transparent"/></ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ScrollBar">
        <Setter Property="Width" Value="10"/>
        <Setter Property="MinWidth" Value="10"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ScrollBar">
                    <Grid Background="Transparent">
                        <Track x:Name="PART_Track" IsDirectionReversed="True">
                            <Track.DecreaseRepeatButton><RepeatButton Command="ScrollBar.PageUpCommand" Style="{StaticResource ScrollPageButton}"/></Track.DecreaseRepeatButton>
                            <Track.Thumb><Thumb Style="{StaticResource ScrollThumbStyle}" Margin="2"/></Track.Thumb>
                            <Track.IncreaseRepeatButton><RepeatButton Command="ScrollBar.PageDownCommand" Style="{StaticResource ScrollPageButton}"/></Track.IncreaseRepeatButton>
                        </Track>
                    </Grid>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
        <Style.Triggers>
            <Trigger Property="Orientation" Value="Horizontal">
                <Setter Property="Width" Value="Auto"/>
                <Setter Property="MinWidth" Value="0"/>
                <Setter Property="Height" Value="10"/>
                <Setter Property="MinHeight" Value="10"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ScrollBar">
                            <Grid Background="Transparent">
                                <Track x:Name="PART_Track" IsDirectionReversed="False">
                                    <Track.DecreaseRepeatButton><RepeatButton Command="ScrollBar.PageLeftCommand" Style="{StaticResource ScrollPageButton}"/></Track.DecreaseRepeatButton>
                                    <Track.Thumb><Thumb Style="{StaticResource ScrollThumbStyle}" Margin="2"/></Track.Thumb>
                                    <Track.IncreaseRepeatButton><RepeatButton Command="ScrollBar.PageRightCommand" Style="{StaticResource ScrollPageButton}"/></Track.IncreaseRepeatButton>
                                </Track>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Trigger>
        </Style.Triggers>
    </Style>

    <!-- ===== Przyciski ===== -->
    <!-- Domyślny przycisk = drugorzędny (ciemne tło z obramowaniem, jak w ServerReview).
         Warianty kolorowe: PrimaryButton (akcent), SuccessButton, WarningButton, DangerButton.
         Najechanie/kliknięcie to półprzezroczysta nakładka, więc działa także na przyciskach,
         którym kolor tła zmienia kod (np. Pauza/Wznów). -->
    <Style TargetType="Button">
        <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="14,6"/>
        <Setter Property="HorizontalContentAlignment" Value="Center"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="SnapsToDevicePixels" Value="True"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Grid>
                        <Border x:Name="bg" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7"/>
                        <Border x:Name="overlay" Background="{DynamicResource ThemeOverlay}" CornerRadius="7" Opacity="0"/>
                        <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" RecognizesAccessKey="True"/>
                    </Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="overlay" Property="Opacity" Value="0.07"/></Trigger>
                        <Trigger Property="IsPressed" Value="True"><Setter TargetName="overlay" Property="Opacity" Value="0.15"/></Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
            <Trigger Property="IsKeyboardFocused" Value="True"><Setter Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
        </Style.Triggers>
    </Style>
    <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource AccentColor}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeOnAccent}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource ThemeSuccessFill}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeOnAccent}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="WarningButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource ThemeWarningFill}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeOnAccent}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource ThemeDangerFill}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeOnAccent}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <!-- Nagłówek zwijanej sekcji (okno Ustawień) - wygląda jak tytuł karty, nie jak przycisk. -->
    <Style x:Key="SectionHeaderButton" TargetType="Button">
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="Padding" Value="16,12"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="10" Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource ThemeHover}"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Pola tekstowe ===== -->
    <Style TargetType="TextBox">
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="8,5"/>
        <Setter Property="CaretBrush" Value="{DynamicResource ThemeText}"/>
        <Setter Property="SelectionBrush" Value="{DynamicResource AccentColor}"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="TextBox">
                    <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" SnapsToDevicePixels="True">
                        <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFaint}"/></Trigger>
                        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="PasswordBox">
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="8,5"/>
        <Setter Property="CaretBrush" Value="{DynamicResource ThemeText}"/>
        <Setter Property="SelectionBrush" Value="{DynamicResource AccentColor}"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="PasswordBox">
                    <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" SnapsToDevicePixels="True">
                        <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFaint}"/></Trigger>
                        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Lista rozwijana ===== -->
    <Style TargetType="ComboBox">
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="MinHeight" Value="28"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ComboBox">
                    <Grid>
                        <ToggleButton Focusable="False" ClickMode="Press"
                                      IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                      Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                            <ToggleButton.Template>
                                <ControlTemplate TargetType="ToggleButton">
                                    <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7">
                                        <Path HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,11,0" Data="M 0 0 L 4 4 L 8 0" Stroke="{DynamicResource ThemeMuted}" StrokeThickness="1.6"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                                        <Trigger Property="IsChecked" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </ToggleButton.Template>
                        </ToggleButton>
                        <ContentPresenter IsHitTestVisible="False" Margin="10,0,28,0" VerticalAlignment="Center" HorizontalAlignment="Left"
                                          Content="{TemplateBinding SelectionBoxItem}"
                                          ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                          ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"/>
                        <Popup x:Name="PART_Popup" IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                            <Border Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeFieldBorder}" BorderThickness="1" CornerRadius="7" Margin="0,3,0,0" Padding="3"
                                    MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                <ScrollViewer SnapsToDevicePixels="True">
                                    <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                </ScrollViewer>
                            </Border>
                        </Popup>
                    </Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Padding" Value="9,6"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ComboBoxItem">
                    <Border x:Name="bd" Background="Transparent" CornerRadius="5" Padding="{TemplateBinding Padding}">
                        <ContentPresenter VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeSelection}"/></Trigger>
                        <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeHover}"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== CheckBox ===== -->
    <Style TargetType="CheckBox">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeScrollThumb}"/>
        <Setter Property="Padding" Value="8,0,0,0"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="CheckBox">
                    <Grid Background="Transparent">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Border x:Name="box" Width="16" Height="16" CornerRadius="4" BorderThickness="1"
                                Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}">
                            <Grid>
                                <Path x:Name="mark" Data="M 3,7.5 L 6,10.5 L 11.5,4" Stroke="{DynamicResource ThemeOnAccent}" StrokeThickness="2"
                                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
                                <Rectangle x:Name="dash" Width="8" Height="2" Fill="{DynamicResource ThemeOnAccent}" Visibility="Collapsed"/>
                            </Grid>
                        </Border>
                        <ContentPresenter x:Name="content" Grid.Column="1" Margin="{TemplateBinding Padding}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" RecognizesAccessKey="True"/>
                    </Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property="HasContent" Value="False"><Setter TargetName="content" Property="Visibility" Value="Collapsed"/></Trigger>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="box" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="box" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                        <Trigger Property="IsChecked" Value="True">
                            <Setter TargetName="box" Property="Background" Value="{DynamicResource AccentColor}"/>
                            <Setter TargetName="box" Property="BorderBrush" Value="{DynamicResource AccentColor}"/>
                            <Setter TargetName="mark" Property="Visibility" Value="Visible"/>
                        </Trigger>
                        <Trigger Property="IsChecked" Value="{x:Null}">
                            <Setter TargetName="box" Property="Background" Value="{DynamicResource AccentColor}"/>
                            <Setter TargetName="box" Property="BorderBrush" Value="{DynamicResource AccentColor}"/>
                            <Setter TargetName="dash" Property="Visibility" Value="Visible"/>
                        </Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Listy (ListBox) ===== -->
    <Style TargetType="ListBox">
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="4"/>
        <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ListBox">
                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" SnapsToDevicePixels="True">
                        <ScrollViewer Padding="{TemplateBinding Padding}" Focusable="False">
                            <ItemsPresenter/>
                        </ScrollViewer>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ListBoxItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Padding" Value="9,6"/>
        <Setter Property="Margin" Value="0,1"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ListBoxItem">
                    <Border x:Name="bd" Background="Transparent" CornerRadius="6" Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeHover}"/></Trigger>
                        <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeSelection}"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Tabele (ListView + GridView) ===== -->
    <Style TargetType="ListView">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
        <Setter Property="BorderThickness" Value="0"/>
    </Style>
    <Style TargetType="GridViewColumnHeader">
        <Setter Property="Background" Value="{DynamicResource ThemeHeader}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
        <Setter Property="FontFamily" Value="Segoe UI"/>
        <Setter Property="FontSize" Value="12"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Padding" Value="10,8"/>
        <Setter Property="HorizontalContentAlignment" Value="Left"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="GridViewColumnHeader">
                    <Grid>
                        <Border x:Name="hb" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="0,0,1,1" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <Thumb x:Name="PART_HeaderGripper" HorizontalAlignment="Right" Width="8" Margin="0,0,-4,0" Cursor="SizeWE">
                            <Thumb.Template>
                                <ControlTemplate TargetType="Thumb"><Border Background="Transparent"/></ControlTemplate>
                            </Thumb.Template>
                        </Thumb>
                    </Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="{DynamicResource ThemeText}"/></Trigger>
                        <Trigger Property="Role" Value="Padding">
                            <Setter TargetName="PART_HeaderGripper" Property="Visibility" Value="Collapsed"/>
                            <Setter TargetName="hb" Property="BorderThickness" Value="0,0,0,1"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ListViewItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeRowLine}"/>
        <Setter Property="BorderThickness" Value="0,0,0,1"/>
        <Setter Property="Padding" Value="2,5"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ListViewItem">
                    <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                        <GridViewRowPresenter Content="{TemplateBinding Content}" Columns="{TemplateBinding GridView.ColumnCollection}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeHover}"/></Trigger>
                        <Trigger Property="IsSelected" Value="True"><Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeSelection}"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Pasek postępu ===== -->
    <Style TargetType="ProgressBar">
        <Setter Property="Background" Value="{DynamicResource ThemeTrack}"/>
        <Setter Property="Foreground" Value="{DynamicResource AccentColor}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Height" Value="8"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ProgressBar">
                    <Grid>
                        <Border x:Name="PART_Track" CornerRadius="4" Background="{TemplateBinding Background}"/>
                        <Border x:Name="PART_Indicator" CornerRadius="4" Background="{TemplateBinding Foreground}" HorizontalAlignment="Left"/>
                    </Grid>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Podpowiedzi i menu kontekstowe ===== -->
    <Style TargetType="ToolTip">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="Padding" Value="9,6"/>
        <Setter Property="FontFamily" Value="Segoe UI"/>
        <Setter Property="FontSize" Value="12"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ToolTip">
                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}" MaxWidth="440">
                        <ContentPresenter>
                            <ContentPresenter.Resources>
                                <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/></Style>
                            </ContentPresenter.Resources>
                        </ContentPresenter>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ContextMenu">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="FontFamily" Value="Segoe UI"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ContextMenu">
                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="7" Padding="4">
                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Cycle"/>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="MenuItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Padding" Value="10,6"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="MenuItem">
                    <Border x:Name="bd" Background="Transparent" CornerRadius="5" Padding="{TemplateBinding Padding}">
                        <ContentPresenter ContentSource="Header" RecognizesAccessKey="True" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeHover}"/></Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
'@

# Pełny słownik zasobów motywu (pędzle bieżącej palety + style) jako tekst XAML.
function global:Get-ThemeResourcesXaml {
    $palette = Get-ThemePalette
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">')
    foreach ($key in $palette.Keys) {
        [void]$sb.AppendLine("    <SolidColorBrush x:Key=`"$key`" Color=`"$($palette[$key])`"/>")
    }
    foreach ($sysKey in $script:ThemeSystemBrushKeys.Keys) {
        [void]$sb.AppendLine("    <SolidColorBrush x:Key=`"{x:Static SystemColors.$sysKey}`" Color=`"$($palette[$script:ThemeSystemBrushKeys[$sysKey]])`"/>")
    }
    [void]$sb.AppendLine($script:ThemeStylesXaml)
    [void]$sb.AppendLine('</ResourceDictionary>')
    return $sb.ToString()
}

# Podmienia pędzle motywu w już otwartym oknie (przełączenie jasny/ciemny w locie). Wszystkie
# kolory w XAML są podpięte przez {DynamicResource ...}, więc okno przemalowuje się samo.
function global:Update-WindowTheme {
    param([System.Windows.Window]$TargetWindow)
    if ($null -eq $TargetWindow) { return }
    $palette = Get-ThemePalette
    foreach ($key in $palette.Keys) {
        $TargetWindow.Resources[$key] = New-ThemeBrush $palette[$key]
    }
    foreach ($sysKey in $script:ThemeSystemBrushKeys.Keys) {
        $TargetWindow.Resources[[System.Windows.SystemColors]::$sysKey] = New-ThemeBrush $palette[$script:ThemeSystemBrushKeys[$sysKey]]
    }
}

# Okno główne i wszystkie okna mu podległe (także zagnieżdżone, np. edytor programu otwarty
# z menedżera programów otwartego z Ustawień).
function global:Get-OpenAppWindows {
    $result = New-Object System.Collections.Generic.List[System.Windows.Window]
    if ($Window -isnot [System.Windows.Window]) { return ,$result }
    $stack = New-Object System.Collections.Stack
    $stack.Push($Window)
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        $result.Add($current)
        foreach ($owned in $current.OwnedWindows) { $stack.Push($owned) }
    }
    return ,$result
}

# Właściciel nowego okna dialogowego = okno, w którym użytkownik właśnie kliknął. Wcześniej
# każde okno było przypinane do okna głównego, więc np. edytor programu otwarty z (modalnego)
# menedżera programów był jego "rodzeństwem" - mógł schować się pod menedżerem, a komunikat
# z przeglądarki logów wyśrodkowywał się na oknie głównym zamiast na przeglądarce.
function global:Get-DialogOwner {
    if ($Window -isnot [System.Windows.Window] -or -not $Window.IsLoaded) { return $null }
    foreach ($w in (Get-OpenAppWindows)) {
        if ($w.IsActive -and $w.IsVisible) { return $w }
    }
    # Okno główne zminimalizowane do zasobnika: okno podległe zminimalizowanemu właścicielowi jest
    # niewidoczne, więc np. pytanie o hasło konta lokalnego w trakcie wdrożenia "wisiało" ukryte,
    # a narzędzie wyglądało na zawieszone. Wtedy okno dialogowe pokazujemy jako samodzielne.
    if ($Window.WindowState -eq [System.Windows.WindowState]::Minimized -or -not $Window.IsVisible) { return $null }
    return $Window
}

# Wczytuje okno z XAML, dokładając do jego zasobów wspólny słownik motywu (jako MergedDictionary,
# więc lokalne style okna mogą z niego dziedziczyć: BasedOn="{StaticResource ...}").
function global:New-ThemedWindow {
    param(
        [Parameter(Mandatory)][xml]$Xaml,
        [switch]$NoOwner
    )
    $ns = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    $root = $Xaml.DocumentElement
    $resNode = $null
    foreach ($child in $root.ChildNodes) {
        if ($child.LocalName -eq 'Window.Resources') { $resNode = $child; break }
    }
    if ($null -eq $resNode) {
        $resNode = $Xaml.CreateElement('Window.Resources', $ns)
        [void]$root.PrependChild($resNode)
    }
    $dict = $Xaml.CreateElement('ResourceDictionary', $ns)
    $merged = $Xaml.CreateElement('ResourceDictionary.MergedDictionaries', $ns)
    $themeDoc = New-Object System.Xml.XmlDocument
    $themeDoc.LoadXml((Get-ThemeResourcesXaml))
    [void]$merged.AppendChild($Xaml.ImportNode($themeDoc.DocumentElement, $true))
    [void]$dict.AppendChild($merged)
    while ($resNode.HasChildNodes) { [void]$dict.AppendChild($resNode.FirstChild) }
    [void]$resNode.AppendChild($dict)

    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Xaml))
    $win.UseLayoutRounding = $true
    if (-not $NoOwner) {
        $owner = Get-DialogOwner
        if ($null -ne $owner -and $owner -ne $win) {
            try { $win.Owner = $owner } catch {}
        }
    }
    return $win
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
        Width="460" SizeToContent="Height" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid Margin="22,22,22,20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid Width="38" Height="38" VerticalAlignment="Top" Margin="0,0,16,0">
                <Ellipse Name="iconBg" Opacity="0.18"/>
                <TextBlock Name="txtIcon" FontSize="18" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <ScrollViewer Grid.Column="1" MaxHeight="420" VerticalScrollBarVisibility="Auto" VerticalAlignment="Center">
                <TextBlock Name="txtMessage" FontSize="13.5" TextWrapping="Wrap"/>
            </ScrollViewer>
        </Grid>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Name="spButtons" Orientation="Horizontal" HorizontalAlignment="Right"/>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    # Tytuł i treść ustawiamy po wczytaniu XAML - znak & albo " w tekście nie zepsuje XML-a.
    $dlg.Title = $Title
    $dlg.FindName("txtMessage").Text = $Message

    if ($Image -match 'Error') { $imgStr = 'Error' }
    elseif ($Image -match 'Warning') { $imgStr = 'Warning' }
    elseif ($Image -match 'Question') { $imgStr = 'Question' }
    else { $imgStr = 'Information' }

    # Ikona = znak na tle koła w kolorze statusu z palety motywu (wcześniej emoji z kolorami
    # dobranymi pod jasne tło - np. ciemnoczerwony ❌ na ciemnym tle był słabo widoczny).
    $iconSpec = switch ($imgStr) {
        "Error"    { @("✕", "ThemeDanger") }
        "Warning"  { @("!", "ThemeWarning") }
        "Question" { @("?", "ThemeAccentText") }
        default    { @("i", "ThemeInfo") }
    }
    $txtIcon = $dlg.FindName("txtIcon")
    $txtIcon.Text = $iconSpec[0]
    $txtIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $iconSpec[1])
    $dlg.FindName("iconBg").SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $iconSpec[1])

    $spButtons = $dlg.FindName("spButtons")
    $script:msgBoxResult = [System.Windows.MessageBoxResult]::None
    $AddBtn = {
        param($content, $resVal, $isDef, $isCanc)
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = $content
        $btn.IsDefault = $isDef
        $btn.IsCancel = $isCanc
        $btn.Tag = $resVal
        $btn.MinWidth = 90
        $btn.MinHeight = 32
        $btn.Margin = "8,0,0,0"
        if ($isDef) { $btn.Style = $dlg.FindResource("PrimaryButton") }
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
        "OKCancel" { & $AddBtn "OK" ([System.Windows.MessageBoxResult]::OK) $true $false; & $AddBtn "Anuluj" ([System.Windows.MessageBoxResult]::Cancel) $false $true }
        "YesNo" { & $AddBtn "Tak" ([System.Windows.MessageBoxResult]::Yes) $true $false; & $AddBtn "Nie" ([System.Windows.MessageBoxResult]::No) $false $true }
        "YesNoCancel" { & $AddBtn "Tak" ([System.Windows.MessageBoxResult]::Yes) $true $false; & $AddBtn "Nie" ([System.Windows.MessageBoxResult]::No) $false $false; & $AddBtn "Anuluj" ([System.Windows.MessageBoxResult]::Cancel) $false $true }
        default { & $AddBtn "OK" ([System.Windows.MessageBoxResult]::OK) $true $false }
    }
    if ($null -ne $dlg.Owner) { $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner }
    $dlg.ShowDialog() | Out-Null
    return $script:msgBoxResult
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

# Motyw wybrany wcześniej przez użytkownika (config.json -> DarkTheme) czytamy już tutaj, żeby
# okno logowania i powitania też go respektowały - wcześniej zawsze były ciemne, nawet gdy
# użytkownik wybrał motyw jasny.
try {
    if (Test-Path -LiteralPath $configPath) {
        $earlyCfg = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -ne $earlyCfg.DarkTheme) { $script:isDarkTheme = [bool]$earlyCfg.DarkTheme }
    }
} catch {}

# ---------- Ekran autoryzacji (PIN) ----------
[xml]$authXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Autoryzacja STD" Width="380" SizeToContent="Height" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <StackPanel Margin="22">
        <TextBlock Text="Smart Tool for Deployment" FontSize="18" FontWeight="SemiBold"/>
        <TextBlock Text="Wprowadź dane dostępowe, aby odblokować narzędzie" Style="{StaticResource MutedText}" FontSize="12" Margin="0,3,0,18"/>
        <Border Style="{StaticResource Card}">
            <StackPanel>
                <TextBlock Text="IDENTYFIKATOR / LOGIN" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogin" Height="34" FontSize="14" Margin="0,0,0,14"/>
                <TextBlock Text="PIN" Style="{StaticResource Caption}"/>
                <PasswordBox Name="txtPin" Height="34" FontSize="14"/>
            </StackPanel>
        </Border>
        <Button Name="btnLogin" Content="Odblokuj narzędzie" Height="38" FontSize="14" Margin="0,18,0,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
    </StackPanel>
</Window>
"@

$authWindow = New-ThemedWindow -Xaml $authXaml -NoOwner
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
            $txtPin.Focus() | Out-Null
        }
    }
})

$authWindow.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Escape) {
        $authWindow.Close()
    }
})
# Kursor od razu w polu loginu - wcześniej trzeba było najpierw kliknąć w pole.
$authWindow.Add_Loaded({ $txtLogin.Focus() | Out-Null })

# W testach Pester (skrypt jest wczytywany przez dot-sourcing) okna logowania i powitania NIE mogą
# się pokazać - ShowDialog() czekałby w nieskończoność na kliknięcie, a brak logowania kończył się
# "exit", który zamykał cały proces Invoke-Pester.
if ($null -eq $global:PesterTesting) {
    $authWindow.ShowDialog() | Out-Null

    if (-not $script:authSuccess) {
        exit
    }

    # Zapisz login od razu do logów po pomyślnej autoryzacji - w tym samym formacie co Write-Log
    # (data, poziom, kontekst), żeby przeglądarka logów pokazała datę i kontekst tego wpisu.
    $authLogLine = "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] [Użytkownik] Zalogowano operatora narzędzia STD: $($script:OperatorLogin)"
    Add-Content -Path $script:LogFilePath -Value $authLogLine -Encoding UTF8 -ErrorAction SilentlyContinue
} else {
    $script:authSuccess = $true
    $script:OperatorLogin = "Pester"
}


# ---------- Utworzenie formularza (WPF) ----------
[xml]$welcomeXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Potwierdzenie" Width="620" SizeToContent="Height" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize"
        WindowStyle="ToolWindow" Topmost="True">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Border Style="{StaticResource Card}" Margin="22,22,22,20" Padding="22,20">
            <StackPanel>
                <TextBlock Text="Smart Tool for Deployment" FontSize="22" FontWeight="SemiBold" Foreground="{DynamicResource ThemeAccentText}"/>
                <TextBlock Name="txtWelcomeVersion" Style="{StaticResource MutedText}" FontSize="12" Margin="0,2,0,16"/>
                <TextBlock TextWrapping="Wrap" FontSize="14.5" LineHeight="23">
                    <Run Text="Celem niniejszego skryptu jest wsparcie przy szybkiej i skutecznej konfiguracji komputera."/>
                    <LineBreak/>
                    <Run Text="Skrypt nie zastępuje decyzji ani uwag inżynierów IT." Foreground="{DynamicResource ThemeMuted}"/>
                </TextBlock>
                <TextBlock Text="Czy chcesz kontynuować?" FontSize="14.5" FontWeight="SemiBold" Margin="0,16,0,0"/>
            </StackPanel>
        </Border>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="22,14">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnOk" Content="OK (3)" Width="120" Height="38" Margin="0,0,10,0" Style="{StaticResource PrimaryButton}" IsEnabled="False"/>
                <Button Name="btnCancel" Content="Anuluj" Width="120" Height="38" IsEnabled="False" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

$welcomeWindow = New-ThemedWindow -Xaml $welcomeXaml -NoOwner
$welcomeWindow.FindName("txtWelcomeVersion").Text = "Wersja $script:ScriptVersion  ·  operator: $($script:OperatorLogin)"

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

if ($null -eq $global:PesterTesting) {
    $result = $welcomeWindow.ShowDialog()

    if ($result -ne $true) {
        exit
    }
} else {
    $timer.Stop()
}

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
                elseif ($sourcePath -notmatch '^\\\\') { $errors.Add("Sciezka sieciowa musi byc UNC (np. \\serwer\udzial\).") | Out-Null }
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
                    $wifiPath = Resolve-ConfigFilePath -Path $wifi
                    if ([string]::IsNullOrWhiteSpace($wifi) -or -not (Test-Path -LiteralPath $wifiPath)) { $errors.Add("Brak pliku profilu Wi-Fi: $wifiPath") | Out-Null }
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
            # $localPath jest ustawiany tylko dla pliku pobranego do %TEMP%, więc to on decyduje o
            # sprzątaniu - nie globalne źródło. Wcześniej warunek "$source -ne 'winget'" sprawiał, że
            # przy źródle winget instalator programu z własnym adresem URL (DownloadUrl) zostawał
            # w %TEMP%, a log pokazywał pustą nazwę pliku, gdy FileName nie był ustawiony.
            if ($null -ne $localPath -and (Test-Path -LiteralPath $localPath)) {
                try { Remove-Item -LiteralPath $localPath -Force -ErrorAction SilentlyContinue } catch {}
                Write-Log "Usunięto plik instalatora: $(Split-Path $localPath -Leaf)"
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

# Pliki z config.json (np. "WiFiProfile.xml") podane bez pełnej ścieżki szukamy obok skryptu, a nie
# w bieżącym katalogu procesu - po uruchomieniu "jako administrator" jest nim zwykle C:\Windows\System32.
function Resolve-ConfigFilePath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or [System.IO.Path]::IsPathRooted($Path)) { return $Path }
    return (Join-Path $ScriptDir $Path)
}

function Import-WiFiProfile {
    if (-not (Test-Path $configPath)) {
        Write-Log "Brak pliku config.json" -IsError
        return
    }
    $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $wifiProfiles = $config.WiFiProfile.FileName

    if ($null -ne $wifiProfiles) {
        foreach ($wifiEntry in @($wifiProfiles)) {
            $wifiProfile = Resolve-ConfigFilePath -Path $wifiEntry
            if (-not [string]::IsNullOrWhiteSpace($wifiProfile) -and (Test-Path -LiteralPath $wifiProfile)) {
                if ($script:DryRun) {
                    Write-Log "[DRY-RUN] Zaimportowano by profil Wi-Fi z pliku: $wifiProfile"
                    continue
                }
                Write-Log "Import profilu Wi-Fi z pliku: $wifiProfile..."
                $exitCode = Start-ProcessWithEvents -FilePath "netsh" -ArgumentList "wlan add profile filename=`"$wifiProfile`""
                if ($exitCode -eq 0) {
                    Write-Log "Profil Wi-Fi '$wifiProfile' zaimportowany."
                } else {
                    Write-Log "Nie udało się zaimportować profilu Wi-Fi '$wifiProfile' (netsh, kod: $exitCode)." -IsError
                }
            }
            else {
                Write-Log "Brak pliku profilu Wi-Fi: $wifiProfile" -IsError
            }
        }
    } else {
        Write-Log "Brak definicji profili Wi-Fi w konfiguracji." -IsError
    }
}

# Proste, ostylowane okno do wpisania tekstu albo hasła. Zastępuje Read-Host, który w aplikacji
# okienkowej pyta w oknie konsoli (zwykle schowanym pod GUI albo niewidocznym w wersji .exe) -
# interfejs wyglądał wtedy na zawieszony.
# -Validate: scriptblock dostający wpisaną wartość; zwraca tekst błędu albo $null, gdy jest OK.
# Zwraca: string, SecureString (z -Password) albo $null po anulowaniu.
function Show-InputDialog {
    param(
        [string]$Title,
        [string]$Message,
        [string]$DefaultText = "",
        [switch]$Password,
        [scriptblock]$Validate
    )
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Width="440" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Margin="22,20,22,18">
            <TextBlock Name="txtMessage" TextWrapping="Wrap" FontSize="13" Margin="0,0,0,12"/>
            <TextBox Name="txtInput" Height="32"/>
            <PasswordBox Name="pwdInput" Height="32" Visibility="Collapsed"/>
            <TextBlock Name="lblConfirm" Text="POWTÓRZ HASŁO" Style="{StaticResource Caption}" Margin="0,12,0,5" Visibility="Collapsed"/>
            <PasswordBox Name="pwdConfirm" Height="32" Visibility="Collapsed"/>
            <TextBlock Name="txtError" Foreground="{DynamicResource ThemeDanger}" TextWrapping="Wrap" Margin="0,10,0,0" Visibility="Collapsed"/>
        </StackPanel>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnOk" Content="OK" Width="90" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    # Tytuł i treść ustawiamy po wczytaniu XAML, a nie przez wklejenie do XAML - znak & albo "
    # w tekście nie zepsuje wtedy XML-a.
    $dlg.Title = $Title
    $dlg.FindName("txtMessage").Text = $Message
    $txtInput = $dlg.FindName("txtInput")
    $pwdInput = $dlg.FindName("pwdInput")
    $pwdConfirm = $dlg.FindName("pwdConfirm")
    $txtError = $dlg.FindName("txtError")

    if ($Password) {
        $txtInput.Visibility = [System.Windows.Visibility]::Collapsed
        $pwdInput.Visibility = [System.Windows.Visibility]::Visible
        $pwdConfirm.Visibility = [System.Windows.Visibility]::Visible
        $dlg.FindName("lblConfirm").Visibility = [System.Windows.Visibility]::Visible
    } else {
        $txtInput.Text = $DefaultText
    }

    $script:inputDialogResult = $null
    $dlg.FindName("btnOk").Add_Click({
        $value = if ($Password) { $pwdInput.Password } else { $txtInput.Text.Trim() }
        $err = $null
        if ($Password -and $pwdInput.Password -ne $pwdConfirm.Password) { $err = "Hasła nie są identyczne." }
        elseif ($Validate) { $err = & $Validate $value }
        if ($err) {
            $txtError.Text = $err
            $txtError.Visibility = [System.Windows.Visibility]::Visible
            return
        }
        $script:inputDialogResult = if ($Password) { $pwdInput.SecurePassword } else { $value }
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    $dlg.FindName("btnCancel").Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    $dlg.Add_Loaded({ if ($Password) { $pwdInput.Focus() | Out-Null } else { $txtInput.Focus() | Out-Null; $txtInput.SelectAll() } })

    if ($dlg.ShowDialog() -eq $true) { return $script:inputDialogResult }
    return $null
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
        # Najpierw sprawdzamy, czy konto istnieje - wcześniej hasło było wpisywane niepotrzebnie.
        if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) {
            Write-Log "Pominięto tworzenie konta: użytkownik '$username' już istnieje."
            return
        }
        $PasswordSecure = Show-InputDialog -Title "Konto lokalnego administratora" -Message "Podaj hasło dla nowego konta '$username' (co najmniej 8 znaków):" -Password -Validate {
            param($value)
            if ($value.Length -lt 8) { "Hasło musi mieć co najmniej 8 znaków." }
        }
        if ($null -eq $PasswordSecure) {
            Write-Log "Pominięto tworzenie konta '$username' - anulowano wpisywanie hasła."
            return
        }
        New-LocalUser -Name $username -Password $PasswordSecure -PasswordNeverExpires -AccountNeverExpires -ErrorAction Stop | Out-Null
        Add-LocalGroupMember -SID S-1-5-32-544 -Member $username -ErrorAction Stop
        Write-Log "Utworzono lokalne konto '$username' w grupie 'Administratorzy'."
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

# Domyślna nazwa komputera "PC-<numer seryjny>" przycięta do zasad nazw NetBIOS: tylko litery
# łacińskie, cyfry i myślnik, maksymalnie 15 znaków. Numer seryjny potrafi zawierać spacje, kropki
# albo być dłuższy (np. "To Be Filled By O.E.M.") i wtedy Rename-Computer odrzucał nazwę.
function Get-DefaultComputerName {
    param([string]$SerialNumber = [string](Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber)
    $clean = $SerialNumber -replace '[^A-Za-z0-9-]', ''
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "NOSERIAL" }
    $name = "PC-$clean"
    if ($name.Length -gt 15) { $name = $name.Substring(0, 15) }
    return $name.TrimEnd('-')
}

# Zwraca opis błędu, gdy nazwa komputera jest niepoprawna, albo $null, gdy jest OK.
function Test-ComputerNameValid {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "Nazwa nie może być pusta." }
    if ($Name.Length -gt 15) { return "Nazwa może mieć maksymalnie 15 znaków (ograniczenie NetBIOS)." }
    if ($Name -notmatch '^[A-Za-z0-9-]+$') { return "Dozwolone są tylko litery bez polskich znaków, cyfry i myślnik." }
    if ($Name -match '^\d+$') { return "Nazwa nie może składać się wyłącznie z cyfr." }
    if ($Name.StartsWith('-') -or $Name.EndsWith('-')) { return "Nazwa nie może zaczynać się ani kończyć myślnikiem." }
    return $null
}

function Set-NewComputerName {
    $defaultName = Get-DefaultComputerName
    if ($script:DryRun) {
        Write-Log "[DRY-RUN] Zmieniono by nazwę komputera (domyślnie '$defaultName')."
        return
    }
    try {
        # Wcześniej: Read-Host -OutVariable NewName - OutVariable zapisuje wynik jako KOLEKCJĘ, a nie
        # tekst, i w dodatku Read-Host pytał w oknie konsoli, a nie w GUI.
        $newName = Show-InputDialog -Title "Zmiana nazwy komputera" -Message "Podaj nową nazwę komputera (maks. 15 znaków). Obecna nazwa: $env:COMPUTERNAME" -DefaultText $defaultName -Validate {
            param($value)
            Test-ComputerNameValid -Name $value
        }
        if ($null -eq $newName) {
            Write-Log "Pominięto zmianę nazwy komputera (anulowano)."
            return
        }
        if ($newName -eq $env:COMPUTERNAME) {
            Write-Log "Pominięto zmianę nazwy - komputer już nazywa się '$newName'."
            return
        }
        Rename-Computer -NewName $newName -Force -ErrorAction Stop
        Write-Log "Zmieniono nazwę komputera na '$newName'. Zmiana zadziała po ponownym uruchomieniu."
    } catch {
        Write-Log "Błąd zmiany nazwy komputera: $_" -IsError
    }
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
# { "Version": "3.2.0", "FileName": "SmartToolforDeployment.ps1", "Notes": "Opis zmian...", "Sha256": "<suma SHA-256 pliku>" }
# Pole Sha256 jest opcjonalne, ale zalecane: sumę liczy się poleceniem
#   (Get-FileHash .\SmartToolforDeployment.ps1 -Algorithm SHA256).Hash
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
        # Split-Path -Leaf: nazwa z pliku na serwerze nie może wskazać innego katalogu (np. "..\..\x.ps1").
        $tempNewFile = Join-Path $env:TEMP (Split-Path $remoteFileName -Leaf)

        Write-Log "Pobieranie wersji $($remote.Version)..."
        if ($isWeb) {
            Invoke-WebRequest -Uri $downloadSource -OutFile $tempNewFile -UseBasicParsing -ErrorAction Stop
        } else {
            Copy-Item -LiteralPath $downloadSource -Destination $tempNewFile -Force -ErrorAction Stop
        }

        # Pobrany plik zastąpi narzędzie uruchamiane jako administrator, więc zanim go podmienimy:
        # 1) sprawdzamy sumę SHA-256 z std_version.json (jeśli została podana),
        $expectedHash = ([string]$remote.Sha256).Trim()
        if (-not [string]::IsNullOrWhiteSpace($expectedHash)) {
            $actualHash = (Get-FileHash -LiteralPath $tempNewFile -Algorithm SHA256).Hash
            if ($actualHash -ne $expectedHash) {
                Remove-Item -LiteralPath $tempNewFile -Force -ErrorAction SilentlyContinue
                throw "Suma kontrolna pobranego pliku się nie zgadza (oczekiwano $expectedHash, jest $actualHash). Aktualizacja przerwana."
            }
            Write-Log "Suma kontrolna SHA-256 pobranej wersji jest zgodna."
        } else {
            Write-Log "Uwaga: std_version.json nie zawiera pola Sha256 - pobrany plik nie został zweryfikowany sumą kontrolną." -IsError
        }
        # 2) sprawdzamy, czy to w ogóle poprawny skrypt PowerShell (np. zamiast strony błędu HTML
        #    albo uciętego pliku) - inaczej po podmianie narzędzie przestałoby się uruchamiać.
        if ($tempNewFile -like "*.ps1") {
            $parseErrors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($tempNewFile, [ref]$null, [ref]$parseErrors) | Out-Null
            if ($parseErrors.Count -gt 0) {
                Remove-Item -LiteralPath $tempNewFile -Force -ErrorAction SilentlyContinue
                throw "Pobrany plik zawiera błędy składni PowerShell ($($parseErrors.Count)), np.: $($parseErrors[0].Message). Aktualizacja przerwana."
            }
        }

        $currentFile = $PSCommandPath
        if ([string]::IsNullOrWhiteSpace($currentFile)) { throw "Nie udało się ustalić ścieżki bieżącego pliku (`$PSCommandPath) do podmiany." }

        # Uruchamiamy odrębny, krótkotrwały proces PowerShell, który poczeka aż bieżący proces się
        # zamknie, podmieni plik i uruchomi narzędzie ponownie - nie da się nadpisać pliku, który
        # jest w danym momencie wykonywany przez BIEŻĄCY proces.
        # Apostrof w ścieżce (np. C:\Users\O'Brien\...) zamknąłby napis w pojedynczych cudzysłowach
        # i zepsuł polecenie - w PowerShell apostrof wewnątrz '...' zapisuje się jako ''.
        $tempQ = $tempNewFile.Replace("'", "''")
        $currentQ = $currentFile.Replace("'", "''")
        $updaterScript = "Start-Sleep -Seconds 2; Copy-Item -LiteralPath '$tempQ' -Destination '$currentQ' -Force; Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File `"$currentQ`"'"
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
        # Wartości z WMI/rejestru (model dysku, opis karty sieciowej itd.) mogą zawierać znaki <, >, &
        # - bez kodowania psuły układ raportu HTML.
        $enc = { param($v) [System.Net.WebUtility]::HtmlEncode([string]$v) }
        $html = @"
<!DOCTYPE html>
<html lang='pl'>
<head>
    <meta charset='utf-8'>
    <title>Raport Systemowy: $(& $enc $env:COMPUTERNAME)</title>
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
    <h1>Raport Systemowy: $(& $enc $env:COMPUTERNAME)</h1>
    <div class='meta'>Wygenerowano: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>Użytkownik: $(& $enc $env:USERNAME)</div>
    
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
            <li><strong>OS:</strong> $(& $enc "$($os.Caption) $($os.OSArchitecture) (Build: $($os.BuildNumber))")</li>
            <li><strong>BIOS SN:</strong> $(& $enc $bios.SerialNumber)</li>
            <li><strong>BIOS Wersja:</strong> $(& $enc $bios.SMBIOSBIOSVersion)</li>
        </ul>
    </div>

    <div class="section" id="cpu">
        <h2>⚙️ Procesor i RAM</h2>
        <ul class="list">
            <li><strong>Procesor:</strong> $(& $enc $cpu.Name)</li>
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
                $html += "            <li><strong>$(& $enc $d.Model)</strong> - $dSize GB</li>`r`n"
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
                $html += "            <tr><td>$(& $enc $n.Description)</td><td>$(& $enc $n.MACAddress)</td><td>$(& $enc $ip)</td></tr>`r`n"
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
        $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction Stop
        if ($bl.VolumeStatus -eq "FullyEncrypted" -or $bl.VolumeStatus -eq "EncryptionInProgress") {
            Write-Log "Dysk C: jest już zaszyfrowany lub proces jest w toku."
            return
        }
        if ($bl.VolumeStatus -ne "FullyDecrypted") {
            # Np. EncryptionPaused / DecryptionInProgress - dysk jest częściowo zaszyfrowany, więc nie
            # ruszamy automatycznie jego protektorów. To wymaga decyzji administratora.
            Write-Log "Dysk C: jest w stanie '$($bl.VolumeStatus)' - pomijam automatyczne szyfrowanie, sprawdź stan ręcznie (manage-bde -status C:)." -IsError
            return
        }
        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Zaszyfrowano by dysk C: (BitLocker XTS-AES 256) i wyeksportowano klucz odzyskiwania."
            return
        }

        # 1. Ustalamy katalog na klucz odzyskiwania ZANIM cokolwiek zmienimy na dysku.
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $exportDir = "C:\Audit\"
        if ($null -ne $config.HardwareAudit -and -not [string]::IsNullOrWhiteSpace($config.HardwareAudit.ExportPath)) {
            $exportDir = [System.Environment]::ExpandEnvironmentVariables($config.HardwareAudit.ExportPath)
        }
        if ($exportDir -notmatch '^\\\\') {
            Write-Log "Uwaga: klucz odzyskiwania zostanie zapisany lokalnie ($exportDir), czyli na szyfrowanym dysku. Zalecana jest ścieżka sieciowa (UNC) w HardwareAudit.ExportPath." -IsError
        }
        if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force -ErrorAction Stop | Out-Null }

        # 2. Hasło odzyskiwania: używamy istniejącego (np. z wcześniejszej nieudanej próby), inaczej dodajemy nowe.
        Write-Log "Generowanie klucza odzyskiwania..."
        $recProtector = @($bl.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }) | Select-Object -First 1
        if ($null -eq $recProtector) {
            Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector -ErrorAction Stop | Out-Null
            $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction Stop
            $recProtector = @($bl.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }) | Select-Object -First 1
        }
        $recPassword = $recProtector.RecoveryPassword
        if ([string]::IsNullOrWhiteSpace($recPassword)) { throw "Nie udało się odczytać hasła odzyskiwania BitLocker." }

        # 3. Zapis klucza PRZED włączeniem szyfrowania. Jeśli zapis się nie uda (-ErrorAction Stop),
        #    szyfrowanie w ogóle nie wystartuje - nie zostaniemy z zaszyfrowanym dyskiem bez kopii klucza.
        $fileName = "BitLocker_Recovery_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $fullPath = Join-Path $exportDir $fileName
        $info = "Komputer: $($env:COMPUTERNAME)`r`nData: $(Get-Date)`r`nID protektora: $($recProtector.KeyProtectorId)`r`nKlucz odzyskiwania: $recPassword"
        $info | Out-File -FilePath $fullPath -Encoding UTF8 -Force -ErrorAction Stop
        Write-Log "Klucz odzyskiwania zapisano w: $fullPath"

        # 4. Protektor TPM pozostawiony przez wcześniejszą wersję narzędzia (dodawała go, a potem
        #    Enable-BitLocker padał) usuwamy - Enable-BitLocker -TpmProtector doda własny. Dysk jest
        #    w stanie FullyDecrypted (sprawdzone wyżej), więc usunięcie protektora niczego nie odsłania.
        foreach ($tpmProtector in @($bl.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'Tpm' })) {
            Remove-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $tpmProtector.KeyProtectorId -ErrorAction Stop | Out-Null
        }

        # 5. Enable-BitLocker MUSI dostać protektor (-TpmProtector, -RecoveryPasswordProtector itd.).
        #    Wcześniej wywołanie bez niego zawsze kończyło się błędem "Parameter set cannot be resolved".
        Write-Log "Rozpoczynanie szyfrowania dysku C: (XTS-AES 256)..."
        Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAes256 -UsedSpaceOnly -SkipHardwareTest -TpmProtector -ErrorAction Stop | Out-Null
        Write-Log "Szyfrowanie zostało zainicjowane pomyślnie."
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
        Write-Log "Brak pliku config.json" -IsError
        return
    }

    $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $programs = $config.Programs
    
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Wybór aplikacji" Width="560" Height="660" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="4,5"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="20,18,20,12">
            <TextBlock Text="SZUKAJ APLIKACJI" Style="{StaticResource Caption}"/>
            <TextBox Name="txtSearch" Height="34" FontSize="14"/>
        </StackPanel>

        <Border Grid.Row="1" Style="{StaticResource Card}" Margin="20,0,20,0" Padding="12,10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="spApps"/>
            </ScrollViewer>
        </Border>

        <Grid Grid.Row="2" Margin="20,12,20,16">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Name="btnSelectAll" Content="Zaznacz wszystko" Grid.Column="0" Margin="0,0,5,0" Height="34"/>
            <Button Name="btnDeselectAll" Content="Odznacz wszystko" Grid.Column="1" Margin="5,0,5,0" Height="34"/>
            <Button Name="btnInvertSelection" Content="Odwróć zaznaczenie" Grid.Column="2" Margin="5,0,0,0" Height="34"/>
        </Grid>

        <Border Grid.Row="3" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="20,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnOk" Content="OK" Width="120" Height="36" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="120" Height="36" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $popup = New-ThemedWindow -Xaml $xaml

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
    # Kolory przycisków zmieniamy przez style motywu, a nie stałe HEX - przypisanie Background na
    # sztywno odcinało przycisk od motywu (po przełączeniu jasny/ciemny zostawał w starym kolorze).
    $btnStart.Style = $Window.FindResource("WarningButton")
    $progressBar.Value = 0
    $script:isCancelled = $false
    $script:isPaused = $false
    $script:DryRun = ($CheckboxControls.ContainsKey("DryRun") -and $CheckboxControls["DryRun"].IsChecked -eq $true)
    if ($script:DryRun) { Write-Log "=== TRYB TESTOWY (DRY-RUN) WŁĄCZONY: żadne rzeczywiste zmiany nie zostaną wprowadzone ===" }

    if (-not (Test-BeforeRun)) {
        Write-Log "Walidacja nie powiodla sie. Przerywam." -IsError
        $btnStart.Style = $Window.FindResource("DangerButton")
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
        $btnStart.Style = $Window.FindResource("SuccessButton")
        $btnStart.IsEnabled = $true
        $btnPause.IsEnabled = $false
        $btnCancelDeploy.IsEnabled = $false
        $wasCancelled = $script:isCancelled
        # Przy przerwaniu (return w środku try) krok "Przywracanie ustawień hibernacji" był pomijany,
        # więc blokada usypiania zostawała włączona aż do zamknięcia aplikacji.
        if ($wasCancelled -and $CheckboxControls.ContainsKey("SuspendHibernation") -and $CheckboxControls["SuspendHibernation"].IsChecked -eq $true) {
            Resume-Hibernation
        }
        $script:isPaused = $false
        $script:isCancelled = $false
        $script:DryRun = $false
        $script:CurrentLogContext = "System"
        $btnPause.Content = "Pauza"
        $btnPause.Style = $Window.FindResource("WarningButton")

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

# Ustawia pole w obiekcie konfiguracji niezależnie od tego, czy już istnieje. Obiekt z ConvertFrom-Json
# to PSCustomObject: przypisanie $obj.Pole = ... do NIEISTNIEJĄCEGO pola rzuca wyjątek "The property
# 'Pole' cannot be found on this object", dlatego brakujące pole dodajemy przez Add-Member.
function Set-ConfigValue {
    param($Object, [string]$Name, $Value)
    if ($Object -is [System.Collections.IDictionary]) { $Object[$Name] = $Value; return }
    if ($null -ne $Object.PSObject.Properties[$Name]) { $Object.$Name = $Value }
    else { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
}

# Zwraca sekcję konfiguracji (np. DomainJoin), a gdy jej brakuje - tworzy pustą i dodaje do obiektu.
function Get-ConfigSection {
    param($Object, [string]$Name)
    $section = if ($Object -is [System.Collections.IDictionary]) { $Object[$Name] } else { $Object.$Name }
    if ($null -eq $section) {
        $section = [PSCustomObject]@{}
        Set-ConfigValue -Object $Object -Name $Name -Value $section
    }
    return $section
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
        Title="Szczegóły wpisu logu" Width="640" Height="460" MinWidth="420" MinHeight="280" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <WrapPanel Grid.Row="0" Margin="20,18,20,12">
            <Border Style="{StaticResource Chip}" Margin="0,0,8,0">
                <TextBlock Name="txtRodzaj" FontWeight="SemiBold"/>
            </Border>
            <Border Style="{StaticResource Chip}" Margin="0,0,8,0">
                <TextBlock Name="txtData" FontFamily="Consolas"/>
            </Border>
            <Border Style="{StaticResource Chip}">
                <TextBlock Name="txtKontekst" Foreground="{DynamicResource ThemeMuted}"/>
            </Border>
        </WrapPanel>
        <Border Grid.Row="1" Style="{StaticResource Card}" Margin="20,0,20,0" Padding="2">
            <TextBox Name="txtInformacja" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" BorderThickness="0" Background="Transparent"
                     FontFamily="Consolas" FontSize="13" Padding="12" VerticalContentAlignment="Stretch"
                     VerticalScrollBarVisibility="Auto"/>
        </Border>
        <Border Grid.Row="2" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="20,12" Margin="0,16,0,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnCopy" Content="Kopiuj" Width="100" Height="32" Margin="0,0,8,0"/>
                <Button Name="btnClose" Content="Zamknij" Width="100" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml

    $txtRodzaj = $dlg.FindName("txtRodzaj")
    $txtRodzaj.Text = $Entry.RodzajText
    try { $txtRodzaj.Foreground = New-ThemeBrush $Entry.RowColorHex } catch {}
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
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,0,0,1" Padding="20,12">
            <StackPanel>
                <TextBlock Text="Przeglądarka logów" FontSize="18" FontWeight="SemiBold"/>
                <TextBlock Name="txtStats" Style="{StaticResource MutedText}" FontSize="11.5"/>
            </StackPanel>
        </Border>

        <Grid Grid.Row="1" Margin="20,14,20,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="170"/>
                <ColumnDefinition Width="170"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Name="txtSearch" Grid.Column="0" Height="32" Margin="0,0,10,0" ToolTip="Szukaj w treści wpisów..."/>
            <ComboBox Name="cmbRodzajFilter" Grid.Column="1" Height="32" Margin="0,0,10,0" SelectedIndex="0">
                <ComboBoxItem Content="Rodzaj: wszystkie"/>
                <ComboBoxItem Content="❌ Błędy"/>
                <ComboBoxItem Content="⚠️ Ostrzeżenia"/>
                <ComboBoxItem Content="✔️ Sukcesy"/>
                <ComboBoxItem Content="🧪 Dry-Run"/>
                <ComboBoxItem Content="ℹ️ Informacyjne"/>
            </ComboBox>
            <ComboBox Name="cmbKontekstFilter" Grid.Column="2" Height="32" Margin="0,0,10,0" SelectedIndex="0">
                <ComboBoxItem Content="Kontekst: wszystkie"/>
                <ComboBoxItem Content="System"/>
                <ComboBoxItem Content="Automat"/>
                <ComboBoxItem Content="Użytkownik"/>
            </ComboBox>
            <CheckBox Name="chkAutoRefresh" Grid.Column="3" Content="Auto-odświeżanie" IsChecked="True" VerticalAlignment="Center" Margin="4,0,16,0"/>
            <Button Name="btnRefresh" Grid.Column="4" Content="Odśwież" Width="100" Height="32"/>
        </Grid>

        <Border Grid.Row="2" Style="{StaticResource Card}" Margin="20,0,20,0" Padding="1">
            <ListView Name="lvLogs" FontFamily="Consolas" FontSize="13" ScrollViewer.HorizontalScrollBarVisibility="Auto">
                <ListView.ItemContainerStyle>
                    <!-- Kolor tekstu wiersza zależy od rodzaju wpisu (błąd, ostrzeżenie, sukces...). -->
                    <Style TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
                        <Setter Property="Foreground" Value="{Binding RowColorHex}"/>
                    </Style>
                </ListView.ItemContainerStyle>
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Rodzaj" Width="130" DisplayMemberBinding="{Binding RodzajText}"/>
                        <GridViewColumn Header="Data zdarzenia" Width="160" DisplayMemberBinding="{Binding DataText}"/>
                        <GridViewColumn Header="Kontekst" Width="110" DisplayMemberBinding="{Binding Kontekst}"/>
                        <GridViewColumn Header="Informacja" Width="660" DisplayMemberBinding="{Binding Informacja}"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </Border>

        <Border Grid.Row="3" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="20,12" Margin="0,14,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal">
                    <Button Name="btnExportHtml" Content="Eksportuj raport (HTML)" Width="190" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}"/>
                    <Button Name="btnSave" Content="Zapisz jako..." Width="120" Height="32" Margin="0,0,8,0"/>
                    <Button Name="btnOpenLog" Content="Otwórz plik" Width="110" Height="32" Margin="0,0,8,0"/>
                    <Button Name="btnOpenDir" Content="Otwórz folder" Width="120" Height="32" Margin="0,0,8,0"/>
                    <Button Name="btnZipLogs" Content="Spakuj do ZIP" Width="130" Height="32"/>
                </StackPanel>
                <Button Name="btnClearLogs" Grid.Column="1" Content="Wyczyść logi" Width="120" Height="32" Style="{StaticResource DangerButton}"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@
    $logWindow = New-ThemedWindow -Xaml $xaml

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

    # Kolory wierszy (tekst HEX, bo wiązanie Foreground="{Binding RowColorHex}" oczekuje stringa)
    # brane z palety AKTUALNEGO motywu. Wcześniej były stałe, dobrane pod jasne tło - np. ciemna
    # czerwień błędów (#C50F1F) i ciemna zieleń sukcesów (#107C10) były słabo czytelne na ciemnym.

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

        $rodzajKey = 'Info'; $rodzajText = 'ℹ️ Informacja'; $colorHex = $script:LogRowPalette.ThemeText
        if ($level -eq 'ERROR') { $rodzajKey = 'Error'; $rodzajText = '❌ Błąd'; $colorHex = $script:LogRowPalette.ThemeDanger }
        elseif ($text -match '\[DRY-RUN\]') { $rodzajKey = 'DryRun'; $rodzajText = '🧪 Dry-Run'; $colorHex = $script:LogRowPalette.ThemeDryRun }
        elseif ($text -match 'Mało wolnego|Nie udało się|Pominięto|Błąd dostępu|BRAK POLECENIA') { $rodzajKey = 'Warning'; $rodzajText = '⚠️ Ostrzeżenie'; $colorHex = $script:LogRowPalette.ThemeWarning }
        elseif ($text -match 'zakończon[ao] pomyślnie|Konfiguracja zakończona|Utworzono|zainstalowany\.|utworzony pomyślnie|Dołączono do domeny') { $rodzajKey = 'Success'; $rodzajText = '✔️ Sukces'; $colorHex = $script:LogRowPalette.ThemeSuccess }

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
        # Paleta pobierana raz na całe renderowanie (ParseLogLine woła się dla każdej linii).
        $script:LogRowPalette = Get-ThemePalette
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
        Title="Potwierdzenie deinstalacji" Width="560" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Margin="22,20,22,18">
            <TextBlock Name="txtHeader" FontSize="15" FontWeight="SemiBold" Foreground="{DynamicResource ThemeDanger}" TextWrapping="Wrap" Margin="0,0,0,12"/>
            <Border Style="{StaticResource Card}" Padding="12,10" MaxHeight="260" Margin="0,0,0,14">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock Name="txtAppList" TextWrapping="Wrap" FontFamily="Consolas" FontSize="13"/>
                </ScrollViewer>
            </Border>
            <TextBlock Text="Czy na pewno chcesz kontynuować? Ta operacja usunie wybrane aplikacje." FontSize="13.5" TextWrapping="Wrap"/>
        </StackPanel>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnYes" Content="Tak, odinstaluj" Width="140" Height="34" Margin="0,0,8,0" Style="{StaticResource DangerButton}" IsDefault="True"/>
                <Button Name="btnNo" Content="Anuluj" Width="100" Height="34" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    $dlg.FindName("txtHeader").Text = "UWAGA: Próba cichej deinstalacji $AppCount programów:"
    
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
        Width="480" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ScrollViewer Grid.Row="0" MaxHeight="500" VerticalScrollBarVisibility="Auto" Margin="22,20,12,18" Padding="0,0,10,0">
            <TextBlock Name="txtMessage" FontSize="14" TextWrapping="Wrap"/>
        </ScrollViewer>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnExportHTML" Content="Zapisz HTML" Width="110" Height="32" Margin="0,0,8,0" Visibility="Collapsed"/>
                <Button Name="btnCopy" Content="Kopiuj do schowka" Width="140" Height="32" Margin="0,0,8,0" Visibility="Collapsed"/>
                <Button Name="btnOk" Content="OK" Width="100" Height="32" Style="{StaticResource PrimaryButton}" IsDefault="True" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    # Tytuł ustawiamy po wczytaniu XAML - wklejony wprost do XAML (Title="$Title") tekst ze znakiem
    # & albo " psuł XML i okno w ogóle się nie otwierało.
    $dlg.Title = $Title
    
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
        Title="Informacje o systemie" Width="820" Height="720" MinWidth="620" MinHeight="450" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Window.Resources>
        <Style x:Key="ChipBorder" TargetType="Border" BasedOn="{StaticResource Card}">
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Margin" Value="0,0,10,10"/>
        </Style>
        <Style x:Key="ChipBorderClickable" TargetType="Border" BasedOn="{StaticResource ChipBorder}">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="ToolTip" Value="Kliknij, aby skopiować"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="ChipLabel" TargetType="TextBlock" BasedOn="{StaticResource Caption}">
            <Setter Property="Margin" Value="0"/>
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
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,0,0,1" Padding="20,12">
            <StackPanel>
                <TextBlock Name="txtHeader" FontSize="18" FontWeight="SemiBold"/>
                <TextBlock Name="txtMeta" Style="{StaticResource MutedText}" FontSize="11.5" Margin="0,2,0,0"/>
            </StackPanel>
        </Border>

        <WrapPanel Grid.Row="1" Margin="20,16,10,0">
            <Border Name="chipOs" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="💻 System operacyjny  📋" Style="{StaticResource ChipLabel}"/>
                    <TextBlock Name="txtOs" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipBiosSn" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="🔧 BIOS - numer seryjny (SN)  📋" Style="{StaticResource ChipLabel}"/>
                    <TextBlock Name="txtBiosSn" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipBiosVer" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="🔧 BIOS - wersja  📋" Style="{StaticResource ChipLabel}"/>
                    <TextBlock Name="txtBiosVer" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipCpu" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="⚙️ Procesor  📋" Style="{StaticResource ChipLabel}"/>
                    <TextBlock Name="txtCpu" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
            <Border Name="chipRam" Style="{StaticResource ChipBorderClickable}">
                <StackPanel>
                    <TextBlock Text="🧠 Pamięć RAM  📋" Style="{StaticResource ChipLabel}"/>
                    <TextBlock Name="txtRam" FontWeight="SemiBold" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
        </WrapPanel>

        <Grid Grid.Row="2" Margin="20,0,20,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Style="{StaticResource ChipBorder}" Margin="0,0,10,0">
                <StackPanel>
                    <TextBlock Text="💾 Dyski twarde" Style="{StaticResource ChipLabel}" Margin="0,0,0,4"/>
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
                    <TextBlock Text="🌐 Karty sieciowe" Style="{StaticResource ChipLabel}" Margin="0,0,0,4"/>
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

        <Grid Grid.Row="3" Margin="20,0,20,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="260"/>
            </Grid.ColumnDefinitions>
            <TextBlock Name="txtAppsHeader" Grid.Column="0" FontSize="14" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <TextBox Name="txtSearch" Grid.Column="1" Height="32" ToolTip="Szukaj po nazwie programu..."/>
        </Grid>

        <Border Grid.Row="4" Style="{StaticResource Card}" Margin="20,0,20,0" Padding="1">
            <ListView Name="lvApps">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Nazwa programu" Width="500" DisplayMemberBinding="{Binding DisplayName}"/>
                        <GridViewColumn Header="Wersja" Width="200" DisplayMemberBinding="{Binding DisplayVersion}"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </Border>

        <Border Grid.Row="5" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="20,12" Margin="0,14,0,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnCopy" Content="Kopiuj do schowka" Width="150" Height="32" Margin="0,0,8,0"/>
                <Button Name="btnExportHtml" Content="Zapisz jako HTML" Width="150" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}"/>
                <Button Name="btnClose" Content="Zamknij" Width="100" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml

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
        $filtered = if ([string]::IsNullOrWhiteSpace($term)) { $allApps } else { @($allApps | Where-Object { ([string]$_.DisplayName).IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }) }

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
        Title="Deinstalator Oprogramowania (Wybór zbiorczy)" Height="640" Width="1010" MinHeight="420" MinWidth="820" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,0,0,1" Padding="20,12">
            <StackPanel>
                <TextBlock Text="Deinstalator oprogramowania" FontSize="18" FontWeight="SemiBold"/>
                <TextBlock Text="Zaznacz programy do cichej deinstalacji. Dwuklik - interaktywny deinstalator, prawy przycisk - pokaż w eksploratorze." Style="{StaticResource MutedText}" FontSize="11.5"/>
            </StackPanel>
        </Border>

        <Grid Grid.Row="1" Margin="20,14,20,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Name="txtSearch" Grid.Column="0" Height="32" Margin="0,0,10,0" FontSize="13.5" ToolTip="Szukaj po nazwie programu lub wydawcy..."/>
            <Button Name="btnSelectAllApps" Content="Zaznacz widoczne" Grid.Column="1" Width="130" Height="32" Margin="0,0,6,0"/>
            <Button Name="btnDeselectAllApps" Content="Odznacz widoczne" Grid.Column="2" Width="130" Height="32" Margin="0,0,10,0"/>
            <Button Name="btnExportCSV" Content="Eksportuj CSV" Grid.Column="3" Width="110" Height="32" Margin="0,0,6,0"/>
            <Button Name="btnExportHTML" Content="Eksportuj HTML" Grid.Column="4" Width="120" Height="32" Margin="0,0,10,0"/>
            <Button Name="btnRefresh" Content="Odśwież listę" Grid.Column="5" Width="110" Height="32"/>
        </Grid>

        <Border Grid.Row="2" Style="{StaticResource Card}" Margin="20,0,20,0" Padding="1">
            <ListView Name="lvApps">
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
                        <GridViewColumn Header="Nazwa programu ▲" Width="380">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <StackPanel Orientation="Horizontal">
                                        <Image Source="{Binding Icon}" Width="16" Height="16" Margin="0,0,6,0"/>
                                        <TextBlock Text="{Binding StatusText}" Foreground="{Binding StatusColor}" FontWeight="Bold" Margin="0,0,6,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="{Binding DisplayName}" VerticalAlignment="Center"/>
                                    </StackPanel>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Wersja" Width="110" DisplayMemberBinding="{Binding DisplayVersion}"/>
                        <GridViewColumn Header="Data instalacji" Width="110" DisplayMemberBinding="{Binding InstallDate}"/>
                        <GridViewColumn Header="Wydawca" Width="250" DisplayMemberBinding="{Binding Publisher}"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </Border>

        <Border Grid.Row="3" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="20,12" Margin="0,14,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="lblSelectedCount" Grid.Column="0" Text="Wybrano: 0" VerticalAlignment="Center" Margin="0,0,16,0" FontWeight="SemiBold" FontSize="14"/>
                <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="0,0,16,0">
                    <TextBlock Name="lblUninstallStatus" Text="" FontSize="11" Style="{StaticResource MutedText}" TextWrapping="NoWrap" Margin="0,0,0,4" Visibility="Hidden" TextTrimming="CharacterEllipsis"/>
                    <ProgressBar Name="pbUninstall" Height="8" Minimum="0" Maximum="100" Foreground="{DynamicResource ThemeDangerFill}" Visibility="Hidden"/>
                </StackPanel>
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button Name="btnPauseUninstall" Content="Pauza" Width="90" Height="34" Margin="0,0,8,0" Style="{StaticResource WarningButton}" IsEnabled="False"/>
                    <Button Name="btnKill" Content="Zabij proces" Width="120" Height="34" Margin="0,0,8,0" Style="{StaticResource WarningButton}" IsEnabled="False" ToolTip="Wymuś zamknięcie zawieszonego deinstalatora"/>
                    <Button Name="btnUninstall" Content="Odinstaluj (Cicho)" Width="170" Height="34" Margin="0,0,8,0" Style="{StaticResource DangerButton}"/>
                    <Button Name="btnClose" Content="Zamknij" Width="90" Height="34" IsCancel="True"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@
    $uninstWindow = New-ThemedWindow -Xaml $xaml

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
                    $cmd = $cmd -replace '(?i)/I(?=\{)', '/X'
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
            $btnPauseUninstall.Style = $uninstWindow.FindResource("WarningButton")
            Write-Log "Wznowiono deinstalację zbiorczą."
        } else {
            $script:isUninstallPaused = $true
            $btnPauseUninstall.Content = "Wznów"
            $btnPauseUninstall.Style = $uninstWindow.FindResource("SuccessButton")
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
                                $cmd = ($cmd -replace '(?i)/I(?=\{)', '/X') + " /qn /norestart"
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
                            $app.StatusColor = Get-ThemeColor "ThemeWarning"
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
                            $app.StatusColor = Get-ThemeColor "ThemeSuccess"
                        } else {
                            $app.StatusText = "❌ BŁĄD (kod: $exitCode)"
                            $app.StatusColor = Get-ThemeColor "ThemeDanger"
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
                $btnPauseUninstall.Style = $uninstWindow.FindResource("WarningButton")
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
        Title="Edytuj program" Width="500" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Margin="22,20,22,18">
            <TextBlock Text="IDENTYFIKATOR (NP. CHROME)" Style="{StaticResource Caption}"/>
            <TextBox Name="txtName" Height="32" Margin="0,0,0,14"/>
            <TextBlock Text="NAZWA PLIKU / WINGET ID" Style="{StaticResource Caption}"/>
            <TextBox Name="txtFile" Height="32" Margin="0,0,0,14"/>
            <TextBlock Text="ARGUMENTY CICHEJ INSTALACJI" Style="{StaticResource Caption}"/>
            <TextBox Name="txtArgs" Height="32" Margin="0,0,0,16"/>
            <CheckBox Name="chkForceUrl" Content="Zawsze pobieraj z niestandardowego adresu URL" Margin="0,0,0,8" FontSize="13.5" ToolTip="Nadpisuje globalne źródło instalacji dla tego konkretnego programu."/>
            <TextBox Name="txtUrl" Height="32" Margin="0,0,0,16" IsEnabled="False" ToolTip="Pełny bezpośredni adres URL do instalatora (np. https://.../plik.exe)"/>
            <CheckBox Name="chkEnabled" Content="Domyślnie zaznaczone do instalacji" FontSize="13.5"/>
        </StackPanel>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        Title="Zarządzaj programami" Height="450" Width="500" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="140"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbPrograms" Grid.Column="0" Margin="0,0,14,0" FontSize="13.5"/>
        <DockPanel Grid.Column="1" LastChildFill="False">
            <StackPanel DockPanel.Dock="Top">
                <Button Name="btnAdd" Content="Dodaj" Height="34" Margin="0,0,0,8" Style="{StaticResource PrimaryButton}"/>
                <Button Name="btnEdit" Content="Edytuj" Height="34" Margin="0,0,0,8"/>
                <Button Name="btnClone" Content="Powiel" Height="34" Margin="0,0,0,8"/>
                <Button Name="btnRemove" Content="Usuń" Height="34" Margin="0,0,0,8" Style="{StaticResource DangerButton}"/>
            </StackPanel>
            <Button Name="btnClose" DockPanel.Dock="Bottom" Content="Zamknij" Height="34" IsCancel="True"/>
        </DockPanel>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        Title="Edytuj wpis rejestru" Width="480" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Margin="22,20,22,18">
            <TextBlock Text="KLUCZ GŁÓWNY I SUBKLUCZ (NP. SOFTWARE\MÓJKLUCZ)" Style="{StaticResource Caption}"/>
            <Grid Margin="0,0,0,14">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="96"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <ComboBox Name="cmbRoot" Grid.Column="0" Margin="0,0,8,0" Height="32">
                    <ComboBoxItem Content="HKLM:\"/>
                    <ComboBoxItem Content="HKCU:\"/>
                </ComboBox>
                <TextBox Name="txtSubKey" Grid.Column="1" Height="32"/>
            </Grid>
            <TextBlock Text="NAZWA (NAME)" Style="{StaticResource Caption}"/>
            <TextBox Name="txtName" Height="32" Margin="0,0,0,14"/>
            <TextBlock Text="WARTOŚĆ (VALUE)" Style="{StaticResource Caption}"/>
            <TextBox Name="txtValue" Height="32" Margin="0,0,0,14"/>
            <TextBlock Text="TYP DANYCH (PROPERTYTYPE)" Style="{StaticResource Caption}"/>
            <ComboBox Name="cmbType" Height="32">
                <ComboBoxItem Content="String"/>
                <ComboBoxItem Content="DWord"/>
                <ComboBoxItem Content="QWord"/>
                <ComboBoxItem Content="ExpandString"/>
                <ComboBoxItem Content="Binary"/>
                <ComboBoxItem Content="MultiString"/>
            </ComboBox>
        </StackPanel>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        Title="Zarządzaj kluczami rejestru" Height="450" Width="620" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="140"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbRegistry" Grid.Column="0" Margin="0,0,14,0" FontSize="13.5"/>
        <DockPanel Grid.Column="1" LastChildFill="False">
            <StackPanel DockPanel.Dock="Top">
                <Button Name="btnAdd" Content="Dodaj" Height="34" Margin="0,0,0,8" Style="{StaticResource PrimaryButton}"/>
                <Button Name="btnEdit" Content="Edytuj" Height="34" Margin="0,0,0,8"/>
                <Button Name="btnRemove" Content="Usuń" Height="34" Margin="0,0,0,8" Style="{StaticResource DangerButton}"/>
            </StackPanel>
            <Button Name="btnClose" DockPanel.Dock="Bottom" Content="Zamknij" Height="34" IsCancel="True"/>
        </DockPanel>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Domyślne zadania startowe" Height="540" Width="480" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
            <Setter Property="Margin" Value="2,5"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Zaznacz opcje, które mają być domyślnie włączone przy starcie aplikacji:" TextWrapping="Wrap" Margin="22,18,22,12" FontSize="13.5"/>
        <Border Grid.Row="1" Style="{StaticResource Card}" Margin="22,0,22,16" Padding="12,8">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="spDefaults"/>
            </ScrollViewer>
        </Border>
        <Border Grid.Row="2" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnSave" Content="Zapisz" Width="100" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="100" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        Title="Edytuj skrypt" Width="460" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Margin="22,20,22,18">
            <TextBlock Text="ŚCIEŻKA DO SKRYPTU (NAZWA PLIKU, UNC LUB URL)" Style="{StaticResource Caption}"/>
            <TextBox Name="txtPath" Height="32"/>
        </StackPanel>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        Title="Zarządzaj skryptami Post-Install" Height="400" Width="540" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="140"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbScripts" Grid.Column="0" Margin="0,0,14,0" FontSize="13.5"/>
        <DockPanel Grid.Column="1" LastChildFill="False">
            <StackPanel DockPanel.Dock="Top">
                <Button Name="btnAdd" Content="Dodaj" Height="34" Margin="0,0,0,8" Style="{StaticResource PrimaryButton}"/>
                <Button Name="btnEdit" Content="Edytuj" Height="34" Margin="0,0,0,8"/>
                <Button Name="btnRemove" Content="Usuń" Height="34" Margin="0,0,0,8" Style="{StaticResource DangerButton}"/>
            </StackPanel>
            <Button Name="btnClose" DockPanel.Dock="Bottom" Content="Zamknij" Height="34" IsCancel="True"/>
        </DockPanel>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Edytuj profil wdrożeniowy" Width="470" Height="580" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
            <Setter Property="FontSize" Value="13.5"/>
            <Setter Property="Margin" Value="2,4"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Margin="22,20,22,0">
            <TextBlock Text="NAZWA PROFILU (NP. KSIĘGOWOŚĆ)" Style="{StaticResource Caption}"/>
            <TextBox Name="txtName" Height="32" Margin="0,0,0,14"/>
        </StackPanel>
        <TextBlock Grid.Row="1" Text="APLIKACJE PRZYPISANE DO PROFILU" Style="{StaticResource Caption}" Margin="22,0,22,5"/>
        <Border Grid.Row="2" Style="{StaticResource Card}" Margin="22,0,22,16" Padding="12,8">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="spApps"/>
            </ScrollViewer>
        </Border>
        <Border Grid.Row="3" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnSave" Content="Zapisz" Width="90" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        Title="Zarządzaj profilami wdrożeniowymi" Height="450" Width="500" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="140"/>
        </Grid.ColumnDefinitions>
        <ListBox Name="lbProfiles" Grid.Column="0" Margin="0,0,14,0" FontSize="13.5"/>
        <DockPanel Grid.Column="1" LastChildFill="False">
            <StackPanel DockPanel.Dock="Top">
                <Button Name="btnAdd" Content="Dodaj" Height="34" Margin="0,0,0,8" Style="{StaticResource PrimaryButton}"/>
                <Button Name="btnEdit" Content="Edytuj" Height="34" Margin="0,0,0,8"/>
                <Button Name="btnRemove" Content="Usuń" Height="34" Margin="0,0,0,8" Style="{StaticResource DangerButton}"/>
            </StackPanel>
            <Button Name="btnClose" DockPanel.Dock="Bottom" Content="Zamknij" Height="34" IsCancel="True"/>
        </DockPanel>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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
        Title="Wymagana autoryzacja" Width="330" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Margin="22,20,22,18">
            <TextBlock Text="Wprowadź PIN, aby edytować ustawienia" FontSize="13.5" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,12"/>
            <PasswordBox Name="txtPin" Height="34" FontSize="14"/>
        </StackPanel>
        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnOk" Content="Odblokuj" Width="96" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="90" Height="32" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    # Kursor od razu w polu PIN - wcześniej trzeba było najpierw w nie kliknąć.
    $dlg.Add_Loaded({ $dlg.FindName("txtPin").Focus() | Out-Null })
    
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
    try { $config = Get-Config } catch { Write-Log "Nie udało się wczytać config.json: $_" -IsError; return }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ustawienia – źródła, domena, konto lokalne" Height="680" Width="700" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <!-- Podpis nad polem -->
        <Style x:Key="FieldLabel" TargetType="TextBlock" BasedOn="{StaticResource Caption}">
            <Setter Property="Margin" Value="0,0,0,4"/>
        </Style>
        <Style TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
            <Setter Property="Height" Value="30"/>
        </Style>
        <Style TargetType="ComboBox" BasedOn="{StaticResource {x:Type ComboBox}}">
            <Setter Property="Height" Value="30"/>
        </Style>
        <!-- Karta zwijanej sekcji -->
        <Style x:Key="SectionCard" TargetType="Border" BasedOn="{StaticResource Card}">
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Margin" Value="0,0,0,12"/>
        </Style>
        <Style x:Key="SectionTitle" TargetType="TextBlock">
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="14.5"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        </Style>
        <Style x:Key="Chevron" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        </Style>
        <Style x:Key="ToolButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Height" Value="38"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Padding" Value="14,6"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="20,18,12,8">
            <StackPanel Margin="0,0,8,0">
                <Border Style="{StaticResource SectionCard}">
                    <StackPanel>
                        <Button Name="btnToggleSrc" Style="{StaticResource SectionHeaderButton}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="📡 Źródła instalacji i uwierzytelnianie sieciowe" Style="{StaticResource SectionTitle}"/>
                                <TextBlock Name="chevronSrc" Grid.Column="1" Text="▾" Style="{StaticResource Chevron}"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelSrc" Margin="16,2,16,16" Visibility="Visible">
                            <TextBlock Text="Domyślne źródło (DefaultInstallSource)" Style="{StaticResource FieldLabel}"/>
                            <ComboBox Name="cmbSrc" Margin="0,0,0,12"/>
                            <TextBlock Text="Ścieżka sieciowa [network] (UNC)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtNet" Margin="0,0,0,12"/>
                            <TextBlock Text="Ścieżka sieciowa [web] (URL)" Style="{StaticResource FieldLabel}"/>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Name="txtWeb" Grid.Column="0" Margin="0,0,8,0"/>
                                <Button Name="btnTestWeb" Content="Testuj" Grid.Column="1" Width="80" Height="30" Style="{StaticResource PrimaryButton}"/>
                            </Grid>
                            <TextBlock Text="Niestandardowe dane (CustomWebDataLocation URL)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtCwd" Margin="0,0,0,12"/>
                            <TextBlock Text="Konto logowania po HTTP/HTTPS (WebAuth Username)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtWebUser" Margin="0,0,0,12"/>
                            <TextBlock Text="Hasło do konta po HTTP/HTTPS (WebAuth Password)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtWebPass"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource SectionCard}">
                    <StackPanel>
                        <Button Name="btnToggleDom" Style="{StaticResource SectionHeaderButton}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="🏢 Domena i konto lokalne" Style="{StaticResource SectionTitle}"/>
                                <TextBlock Name="chevronDom" Grid.Column="1" Text="▾" Style="{StaticResource Chevron}"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelDom" Margin="16,2,16,16" Visibility="Visible">
                            <TextBlock Text="Nazwa domeny (DomainName)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtDom" Margin="0,0,0,12"/>
                            <TextBlock Text="Konto uprawnione do podłączenia (Username)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtDomUser" Margin="0,0,0,12"/>
                            <TextBlock Text="Nazwa domyślnego konta lokalnego (LocalAdmin Username)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtLoc"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource SectionCard}">
                    <StackPanel>
                        <Button Name="btnToggleTv" Style="{StaticResource SectionHeaderButton}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="📺 TeamViewer" Style="{StaticResource SectionTitle}"/>
                                <TextBlock Name="chevronTv" Grid.Column="1" Text="▸" Style="{StaticResource Chevron}"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelTv" Margin="16,2,16,16" Visibility="Collapsed">
                            <TextBlock Text="Nazwa pliku / Winget ID" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtTvFile" Margin="0,0,0,12"/>
                            <TextBlock Text="Argumenty instalacji" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtTvArgs"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource SectionCard}">
                    <StackPanel>
                        <Button Name="btnToggleAv" Style="{StaticResource SectionHeaderButton}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="🛡️ Antywirus" Style="{StaticResource SectionTitle}"/>
                                <TextBlock Name="chevronAv" Grid.Column="1" Text="▸" Style="{StaticResource Chevron}"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelAv" Margin="16,2,16,16" Visibility="Collapsed">
                            <TextBlock Text="Nazwa pliku / Winget ID" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtAvFile" Margin="0,0,0,12"/>
                            <TextBlock Text="Domyślne źródło instalacji" Style="{StaticResource FieldLabel}"/>
                            <ComboBox Name="cmbAvSrc" Margin="0,0,0,12"/>
                            <TextBlock Text="Ścieżka sieciowa [network] (UNC)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtAvNet" Margin="0,0,0,12"/>
                            <TextBlock Text="Ścieżka sieciowa [web] (URL)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtAvWeb" Margin="0,0,0,12"/>
                            <TextBlock Text="Użytkownik WebAuth" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtAvUser" Margin="0,0,0,12"/>
                            <TextBlock Text="Hasło WebAuth" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtAvPass"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource SectionCard}">
                    <StackPanel>
                        <Button Name="btnToggleWifi" Style="{StaticResource SectionHeaderButton}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="📶 Profil Wi-Fi" Style="{StaticResource SectionTitle}"/>
                                <TextBlock Name="chevronWifi" Grid.Column="1" Text="▸" Style="{StaticResource Chevron}"/>
                            </Grid>
                        </Button>
                        <StackPanel Name="panelWifi" Margin="16,2,16,16" Visibility="Collapsed">
                            <TextBlock Text="Nazwy plików (po przecinku, jeśli kilka)" Style="{StaticResource FieldLabel}"/>
                            <TextBox Name="txtWifiFile"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="🛠️ Zarządzanie" Style="{StaticResource CardTitle}" Margin="4,0,0,6"/>
                        <UniformGrid Columns="2" Rows="3">
                            <Button Name="btnManageApps" Content="📦 Programy" Style="{StaticResource ToolButton}"/>
                            <Button Name="btnManageProfiles" Content="🗂️ Profile wdrożeniowe" Style="{StaticResource ToolButton}"/>
                            <Button Name="btnManageRegistry" Content="🧩 Rejestr niestandardowy" Style="{StaticResource ToolButton}"/>
                            <Button Name="btnManageDefaults" Content="☑️ Domyślne zadania" Style="{StaticResource ToolButton}"/>
                            <Button Name="btnManageScripts" Content="📜 Skrypty Post-Install" Style="{StaticResource ToolButton}"/>
                            <Button Name="btnCheckUpdate" Content="🔄 Aktualizacje narzędzia" Style="{StaticResource ToolButton}"/>
                        </UniformGrid>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}" Margin="0,0,0,4">
                    <StackPanel>
                        <TextBlock Text="💾 Kopia zapasowa konfiguracji" Style="{StaticResource CardTitle}"/>
                        <StackPanel Orientation="Horizontal">
                            <Button Name="btnExportConfig" Content="⬆️ Eksportuj..." Height="34" Width="150" Margin="0,0,8,0"/>
                            <Button Name="btnImportConfig" Content="⬇️ Importuj..." Height="34" Width="150"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="1" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="20,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnSave" Content="Zapisz" Width="110" Height="34" Margin="0,0,8,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
                <Button Name="btnCancel" Content="Anuluj" Width="110" Height="34" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    $dlg = New-ThemedWindow -Xaml $xaml
    
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

            # Set-ConfigValue / Get-ConfigSection zamiast "$config.Sekcja.Pole = ...": przypisanie do pola,
            # którego nie ma w obiekcie z ConvertFrom-Json, rzuca wyjątek i cały zapis się nie udawał
            # (np. gdy w config.json była sekcja InstallSourcePaths tylko z "web", bez "network").
            Set-ConfigValue $config 'DefaultInstallSource' $src
            $paths = Get-ConfigSection $config 'InstallSourcePaths'
            Set-ConfigValue $paths 'network' $net
            Set-ConfigValue $paths 'web' $web

            Set-ConfigValue (Get-ConfigSection $config 'CustomWebDataLocation') 'URL' $cwd

            $domainSection = Get-ConfigSection $config 'DomainJoin'
            Set-ConfigValue $domainSection 'DomainName' $dom
            Set-ConfigValue $domainSection 'Username' $domUser

            Set-ConfigValue (Get-ConfigSection $config 'LocalAdmin') 'Username' $locUser

            $webAuth = Get-ConfigSection $config 'WebAuth'
            Set-ConfigValue $webAuth 'Username' $txtWebUser.Text.Trim()
            Set-ConfigValue $webAuth 'Password' $txtWebPass.Text.Trim()

            $tv = Get-ConfigSection $config 'TeamViewer'
            Set-ConfigValue $tv 'FileName' $txtTvFile.Text.Trim()
            Set-ConfigValue $tv 'Arguments' $txtTvArgs.Text.Trim()

            $av = Get-ConfigSection $config 'AntyVirus'
            Set-ConfigValue $av 'FileName' $txtAvFile.Text.Trim()
            $avSrcVal = [string]$cmbAvSrc.SelectedItem
            if ([string]::IsNullOrWhiteSpace($avSrcVal)) { $avSrcVal = "network" }
            Set-ConfigValue $av 'DefaultInstallSource' $avSrcVal
            $avPaths = Get-ConfigSection $av 'InstallSourcePaths'
            Set-ConfigValue $avPaths 'network' $txtAvNet.Text.Trim()
            Set-ConfigValue $avPaths 'web' $txtAvWeb.Text.Trim()
            $avCred = Get-ConfigSection $av 'Credentials'
            Set-ConfigValue $avCred 'Username' $txtAvUser.Text.Trim()
            Set-ConfigValue $avCred 'Password' $txtAvPass.Text.Trim()

            $wifiStr = $txtWifiFile.Text.Trim()
            $wifiValue = if ($wifiStr -match ",") { @($wifiStr -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }) } else { $wifiStr }
            Set-ConfigValue (Get-ConfigSection $config 'WiFiProfile') 'FileName' $wifiValue

            Save-Config $config
            Show-ThemedMessageBox -Message "Zapisano konfigurację." -Title "Sukces" -Button "OK" -Image "Information" | Out-Null
            Get-AppSelection
            # Profile dodane/zmienione w "Profile wdrożeniowe" od razu na liście w oknie głównym -
            # wcześniej pojawiały się dopiero po ręcznym "Przeładuj config.json".
            Load-Profiles
            $dlg.Close()
        }
        catch {
            Show-ThemedMessageBox -Message "Błąd zapisu: $($_.Exception.Message)" -Title "Błąd" -Button "OK" -Image "Error" | Out-Null
        }
    })

    $dlg.ShowDialog() | Out-Null
}

# W testach nie tworzymy/nie sprawdzamy config.json interaktywnie (okno komunikatu zablokowałoby testy).
if ($null -eq $global:PesterTesting) {
    Ensure-Configuration
}

[xml]$mainXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart Tool for Deployment | Short: STD" Height="800" Width="940" MinHeight="680" MinWidth="820" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
            <Setter Property="FontSize" Value="13.5"/>
            <Setter Property="Margin" Value="2,6"/>
        </Style>
        <Style x:Key="SideButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Height" Value="34"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
        </Style>
        <Style x:Key="QuickButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Height" Value="30"/>
            <Setter Property="Padding" Value="8,4"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Pasek nagłówka -->
        <Border Grid.Row="0" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,0,0,1" Padding="20,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Smart Tool for Deployment" FontSize="20" FontWeight="SemiBold"/>
                    <TextBlock Text="Automatyczna konfiguracja stacji roboczych Windows  ·  v$script:ScriptVersion" Style="{StaticResource MutedText}" FontSize="11.5"/>
                </StackPanel>
                <TextBlock Name="txtStopwatch" Grid.Column="1" Text="⏱ 00:00:00" VerticalAlignment="Center" FontWeight="SemiBold" FontSize="15" Margin="0,0,16,0" Visibility="Hidden"/>
                <Border Grid.Column="2" Background="{DynamicResource ThemePanel}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="1" CornerRadius="14" Padding="12,6" VerticalAlignment="Center">
                    <StackPanel Orientation="Horizontal">
                        <Ellipse Name="shpNetworkStatus" Width="9" Height="9" Fill="{DynamicResource ThemeWarning}" Margin="0,0,8,0" VerticalAlignment="Center">
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
                        <TextBlock Name="txtNetworkStatus" Text="Sprawdzanie sieci..." VerticalAlignment="Center" FontSize="12.5"/>
                    </StackPanel>
                </Border>
                <Button Name="btnThemeToggle" Grid.Column="3" Content="☀️ Jasny motyw" Padding="12,5" Height="32" Margin="12,0,0,0" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- Zadania + panel boczny -->
        <Grid Grid.Row="1" Margin="20,16,20,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="5*"/>
                <ColumnDefinition Width="4*"/>
            </Grid.ColumnDefinitions>

            <Border Style="{StaticResource Card}" Margin="0,0,14,0">
                <DockPanel>
                    <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Zadania wdrożenia" Style="{StaticResource CardTitle}" Margin="0" VerticalAlignment="Center"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <Button Name="btnSelectAll" Content="Zaznacz wszystko" Style="{StaticResource QuickButton}" Margin="0,0,6,0"/>
                            <Button Name="btnDeselectAll" Content="Odznacz" Style="{StaticResource QuickButton}" Margin="0,0,6,0" ToolTip="Odznacz wszystko"/>
                            <Button Name="btnInvertSelection" Content="Odwróć" Style="{StaticResource QuickButton}" ToolTip="Odwróć zaznaczenie"/>
                        </StackPanel>
                    </Grid>
                    <Border DockPanel.Dock="Top" Height="1" Background="{DynamicResource ThemeBorder}" Margin="0,0,0,6"/>
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Name="spCheckboxes"/>
                    </ScrollViewer>
                </DockPanel>
            </Border>

            <Border Grid.Column="1" Style="{StaticResource Card}">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,0,4,0">
                    <StackPanel>
                        <TextBlock Text="PROFIL WDROŻENIA (ROLA)" Style="{StaticResource Caption}"/>
                        <ComboBox Name="cmbProfiles" Height="32" Margin="0,0,0,10"/>
                        <Button Name="btnChooseApps" Content="Wybierz aplikacje" Height="38" Margin="0,0,0,16" Style="{StaticResource PrimaryButton}"/>

                        <TextBlock Text="KONFIGURACJA I LOGI" Style="{StaticResource Caption}"/>
                        <Button Name="btnSettings" Content="Ustawienia..." Style="{StaticResource SideButton}"/>
                        <Button Name="btnEditConfig" Content="Edytuj config.json w Notatniku" Style="{StaticResource SideButton}"/>
                        <Button Name="btnReloadConfig" Content="Przeładuj config.json (zaktualizuj GUI)" Style="{StaticResource SideButton}"/>
                        <Button Name="btnLogs" Content="Przeglądaj / Zapisz logi" Style="{StaticResource SideButton}" Margin="0,0,0,16"/>

                        <TextBlock Text="SZYBKIE NARZĘDZIA" Style="{StaticResource Caption}"/>
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
                            <Button Name="btnSysProps" Content="SysProperties" Grid.Row="0" Grid.Column="0" Margin="0,0,4,8" Style="{StaticResource QuickButton}" ToolTip="Zaawansowane ustawienia systemu (Właściwości systemu)"/>
                            <Button Name="btnCompMgmt" Content="Zarządzanie" Grid.Row="0" Grid.Column="1" Margin="4,0,0,8" Style="{StaticResource QuickButton}" ToolTip="Zarządzanie komputerem (compmgmt.msc)"/>
                            <Button Name="btnRegEdit" Content="RegEdit" Grid.Row="1" Grid.Column="0" Margin="0,0,4,8" Style="{StaticResource QuickButton}" ToolTip="Edytor rejestru (regedit)"/>
                            <Button Name="btnPrinters" Content="Drukarki" Grid.Row="1" Grid.Column="1" Margin="4,0,0,8" Style="{StaticResource QuickButton}" ToolTip="Klasyczny widok urządzeń i drukarek"/>
                            <Button Name="btnSysInfo" Content="Informacje o systemie" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,0,0,8" Height="32" ToolTip="Podstawowe informacje o sprzęcie i systemie"/>
                            <Button Name="btnUninstaller" Content="Odinstaluj programy" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Height="32" Style="{StaticResource DangerButton}" ToolTip="Moduł do wymuszania cichej deinstalacji oprogramowania"/>
                        </Grid>
                    </StackPanel>
                </ScrollViewer>
            </Border>
        </Grid>

        <!-- Postęp -->
        <StackPanel Grid.Row="2" Margin="20,16,20,14">
            <TextBlock Name="txtProgressInfo" Text="Oczekiwanie na rozpoczęcie..." Margin="0,0,0,8" FontWeight="SemiBold"/>
            <ProgressBar Name="progressBar" Height="10" Minimum="0" Maximum="100" Margin="0,0,0,6"/>
            <ProgressBar Name="progressBarDownload" Height="5" Minimum="0" Maximum="100" Foreground="{DynamicResource ThemeSuccess}"/>
        </StackPanel>

        <!-- Akcje -->
        <Border Grid.Row="3" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeBorder}" BorderThickness="0,1,0,0" Padding="20,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="120"/>
                </Grid.ColumnDefinitions>
                <Button Name="btnStart" Grid.Column="0" Content="ROZPOCZNIJ KONFIGURACJĘ" Height="50" FontSize="17" FontWeight="Bold" Style="{StaticResource SuccessButton}" Margin="0,0,10,0"/>
                <Button Name="btnPause" Grid.Column="1" Content="Pauza" Height="50" FontSize="16" FontWeight="Bold" Style="{StaticResource WarningButton}" Margin="0,0,10,0" IsEnabled="False"/>
                <Button Name="btnCancelDeploy" Grid.Column="2" Content="Przerwij" Height="50" FontSize="16" FontWeight="Bold" Style="{StaticResource DangerButton}" IsEnabled="False"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

$Window = New-ThemedWindow -Xaml $mainXaml -NoOwner

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

# Przełącza motyw we WSZYSTKICH otwartych oknach (główne + np. niemodalna przeglądarka logów).
# Wcześniej zmieniane były tylko zasoby okna głównego - otwarta przeglądarka logów zostawała
# w starym motywie aż do ponownego otwarcia.
function Set-AppTheme {
    foreach ($w in (Get-OpenAppWindows)) { Update-WindowTheme $w }
    if ($null -ne $btnThemeToggle) {
        $btnThemeToggle.Content = if ($script:isDarkTheme) { "☀️ Jasny motyw" } else { "🌙 Ciemny motyw" }
    }
    # Kolory wierszy logu są zapisane w danych wierszy (RowColorHex) - przebudowujemy tabelę,
    # żeby wzięła kolory z nowej palety.
    if ($null -ne $script:ActiveLogWindow -and $script:ActiveLogWindow.IsLoaded -and $null -ne $script:RenderLogView) {
        try { & $script:RenderLogView } catch {}
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
                        Write-Log "Nie znaleziono programu '$app' (profil '$profName') w konfiguracji." -IsError
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
            Write-Log "Zastosowano profil wdrożenia: $profName (Wybrano programów: $count)" -Context "Użytkownik"
        } catch {
            Write-Log "Błąd podczas ładowania profilu: $_" -IsError
        }
    } elseif ($cmbProfiles.SelectedIndex -eq 0) {
        # Gdy użytkownik celowo kliknie powrót na wybór niestandardowy - ładujemy domyślne z config.json
        Write-Log "Przełączono na wybór niestandardowy - wczytywanie domyślnych aplikacji z pliku config." -Context "Użytkownik"
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
        $btnPause.Style = $Window.FindResource("WarningButton")
        $script:stopwatchStartTime = Get-Date
        $script:stopwatchTimer.Start()
        Write-Log "Wznowiono wdrożenie." -Context "Użytkownik"
    } else {
        $script:isPaused = $true
        $btnPause.Content = "Wznów"
        $btnPause.Style = $Window.FindResource("SuccessButton")
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
    if ($null -ne $script:WebProbePs) { try { $script:WebProbePs.Dispose() } catch {} }
    if ($null -ne $script:WebProbeRunspace) { try { $script:WebProbeRunspace.Dispose() } catch {} }
    if ($null -ne $notifyIcon) {
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
    }
})

$script:networkCheckTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:networkCheckTimer.Interval = [TimeSpan]::FromSeconds(3)
$script:lastNetworkError = $null

# Sprawdzenie źródła Web (żądanie HTTP HEAD) działa w osobnym runspace. Wcześniej wykonywało się
# synchronicznie w wątku interfejsu co 3 sekundy - przy niedostępnym serwerze (timeout 1,5 s)
# okno co chwilę przestawało reagować na kliknięcia, mimo że status sieci miał działać "w tle".
$script:WebProbe = [hashtable]::Synchronized(@{ State = 'Idle'; Url = $null; Result = $null; Code = $null; Error = $null })
$script:WebProbePs = $null
$script:WebProbeRunspace = $null
$script:LastWebStatus = $null
$script:WebProbeScript = @'
$url = $probe.Url
try {
    $req = [System.Net.WebRequest]::Create($url)
    $req.Timeout = 1500
    $req.Method = "HEAD"
    $req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    $res = $req.GetResponse()
    $probe.Code = [int]$res.StatusCode
    $res.Close()
    $probe.Result = 'OK'
} catch {
    $webResp = $null
    if ($_.Exception -is [System.Net.WebException]) { $webResp = $_.Exception.Response }
    elseif ($_.Exception.InnerException -is [System.Net.WebException]) { $webResp = $_.Exception.InnerException.Response }
    if ($null -ne $webResp) {
        $probe.Code = [int]$webResp.StatusCode
        $webResp.Close()
        $probe.Result = 'HttpStatus'
    } else {
        $probe.Error = $_.Exception.Message
        $probe.Result = 'NoResponse'
    }
}
$probe.State = 'Done'
'@

function Start-WebSourceProbe {
    param([string]$Url)
    if ($null -eq $script:WebProbeRunspace) {
        $script:WebProbeRunspace = [runspacefactory]::CreateRunspace()
        $script:WebProbeRunspace.Open()
        $script:WebProbeRunspace.SessionStateProxy.SetVariable('probe', $script:WebProbe)
    }
    $script:WebProbe.Url = $Url
    $script:WebProbe.Result = $null
    $script:WebProbe.Code = $null
    $script:WebProbe.Error = $null
    $script:WebProbe.State = 'Running'
    $script:WebProbePs = [powershell]::Create()
    $script:WebProbePs.Runspace = $script:WebProbeRunspace
    [void]$script:WebProbePs.AddScript($script:WebProbeScript)
    [void]$script:WebProbePs.BeginInvoke()
}

$script:networkCheckTimer.Add_Tick({
    if (-not $btnStart.IsEnabled) { return } # Wstrzymaj sprawdzanie podczas trwającego wdrożenia

    $isNet = [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
    $validIp = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notmatch "^169\.254\." -and $_.IPAddress -ne "127.0.0.1" }

    if ($isNet -and $validIp) {
        $ipStr = @($validIp)[0].IPAddress
        $statusText = "Sieć: $ipStr"
        $statusColor = "ThemeSuccess"

        try {
            $cfg = Get-Config
            if ($cfg.DefaultInstallSource -eq 'web' -and -not [string]::IsNullOrWhiteSpace($cfg.InstallSourcePaths.web)) {
                $url = [string]$cfg.InstallSourcePaths.web
                if ($script:WebProbe.State -eq 'Done') {
                    switch ($script:WebProbe.Result) {
                        'OK' {
                            $script:LastWebStatus = @{ Text = "Web: OK"; Color = "ThemeSuccess" }
                            if ($script:lastNetworkError) {
                                Write-Log "[Status Sieci] Połączenie ze źródłem Web przywrócone."
                                $script:lastNetworkError = $null
                            }
                        }
                        'HttpStatus' {
                            $statusCode = $script:WebProbe.Code
                            $script:LastWebStatus = @{ Text = "Web: OK ($statusCode)"; Color = "ThemeSuccess" }
                            if ($script:lastNetworkError) {
                                Write-Log "[Status Sieci] Źródło Web odpowiedziało statusem $statusCode."
                                $script:lastNetworkError = $null
                            }
                        }
                        default {
                            $script:LastWebStatus = @{ Text = "Web: Brak odp."; Color = "ThemeWarning" }
                            $errMsg = $script:WebProbe.Error
                            if ($script:lastNetworkError -ne $errMsg) {
                                Write-Log "[Status Sieci] Błąd weryfikacji źródła Web ($($script:WebProbe.Url)): $errMsg" -IsError
                                $script:lastNetworkError = $errMsg
                            }
                        }
                    }
                    if ($null -ne $script:WebProbePs) { try { $script:WebProbePs.Dispose() } catch {}; $script:WebProbePs = $null }
                    $script:WebProbe.State = 'Idle'
                }
                if ($script:WebProbe.State -eq 'Idle') { Start-WebSourceProbe -Url $url }
                # Do czasu pierwszej odpowiedzi pokazujemy sam adres IP; potem ostatni znany wynik.
                if ($null -ne $script:LastWebStatus) {
                    $statusText = "Sieć: $ipStr | $($script:LastWebStatus.Text)"
                    $statusColor = $script:LastWebStatus.Color
                }
            } else {
                $script:LastWebStatus = $null
            }
        } catch {}

        $shpNetworkStatus.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $statusColor)
        $txtNetworkStatus.Text = $statusText
    } else {
        $shpNetworkStatus.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, "ThemeDanger")
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