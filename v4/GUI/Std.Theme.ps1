# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.Theme.ps1 - motyw okien STD 4 (SPEC 7.1): paleta ciemna i jasna, wspólne style kontrolek,
# budowanie okna z XAML (New-StdWindow), przełączanie motywu we wszystkich otwartych oknach,
# okno komunikatu, krótkie powiadomienia, ikony z czcionki i ikona STD. Kolory i style są
# przeniesione z STD 3.1 bez zmian wyglądu, nazwy funkcji mają prefiks Std.
#
# Plik wczytuje GUI\Std.UI.ps1 (dot-source w zakresie programu). Funkcje NIE są globalne:
# handlery zdarzeń okien widzą je, bo okno jest pokazywane przez ShowDialog() z kodu programu,
# który w tym czasie wciąż działa. Stan motywu jest w jednym obiekcie $script:StdThemeState.
# Funkcje zapisu i logu (Get-StdPath, Read-StdJsonFile, Write-StdJsonFile, Write-StdLog) pochodzą
# z modułu STD, który program importuje przed oknami.
#
# Zasady motywu:
# - obie palety mają te same klucze w tej samej kolejności;
# - w XAML kolor z palety zawsze przez {DynamicResource ...}, w kodzie przez SetResourceReference -
#   wtedy przełączenie motywu w locie przemalowuje także otwarte okna;
# - style zawsze przez {StaticResource ...}; słownik stylów jest wczytywany z XAML tylko raz
#   i dołączany do każdego okna (New-StdWindow);
# - pary tonów Theme<Ton>Back/Fore (Info, Ok, Warn, Crit) służą do tekstu na zabarwionym tle.
#   W motywie jasnym są ciemniejsze niż ThemeSuccess/ThemeWarning/ThemeDanger (te są dla kropek
#   stanu), żeby tekst miał kontrast co najmniej 4,5:1;
# - ikony to znaki z czcionki Segoe Fluent Icons / Segoe MDL2 Assets, nie emoji. W kodzie podajemy
#   kod szesnastkowy (Get-StdGlyph 'E713'), w XAML encję (&#xE713;) - nigdy wklejony znak.

# ---------------------------------------------------------------------------------------------
# Paleta (ciemna i jasna)
# ---------------------------------------------------------------------------------------------
$script:StdThemePalettes = [ordered]@{
    Dark  = [ordered]@{
        ThemeBackground       = '#0F1318'   # tło okna
        ThemeHeader           = '#12171D'   # pasek nagłówka/wdrożenia, tło okien dialogowych, pasek tytułu (DWM)
        ThemeChromeBorder     = '#1E242E'   # linia oddzielająca pasek nagłówka/wdrożenia
        ThemePanel            = '#161B22'   # karty / panele
        ThemeBorder           = '#242B36'   # obramowanie kart
        ThemeRowLine          = '#1F2530'   # linie między wierszami tabel
        ThemeRowAlt           = '#181E26'   # co drugi wiersz tabeli
        ThemeGridHeader       = '#1B212A'   # nagłówki kolumn tabel
        ThemeGridHeaderBorder = '#252C38'
        ThemeTextBoxBg        = '#1B212A'   # pola edycyjne
        ThemeFieldBorder      = '#2A323F'   # obramowanie pól i przycisków, separator w menu
        ThemeFieldHover       = '#39445A'   # obramowanie pola po najechaniu
        ThemeButton           = '#1B212A'   # przycisk domyślny (drugorzędny)
        ThemeButtonText       = '#E4E8EF'
        ThemeGhostText        = '#AEB6C4'   # przyciski "ghost", tekst stanu
        ThemeGhostHover       = '#1E252F'
        ThemeChip             = '#1E252F'   # neutralny chip / pigułka
        ThemePopup            = '#1B212A'   # menu kontekstowe, lista rozwijana
        ThemePopupBorder      = '#2F3846'   # obramowanie menu, podpowiedzi i powiadomień
        ThemeTooltip          = '#232A35'   # podpowiedź (ToolTip)
        ThemeToast            = '#1E2530'   # powiadomienie (toast)
        ThemeTitle            = '#FFFFFF'   # tytuły okien i okien dialogowych
        ThemeText             = '#E4E8EF'   # tekst podstawowy
        ThemeBodyText         = '#C9D0DC'   # treść okien dialogowych
        ThemeGridText         = '#D5DAE3'   # tekst w tabelach i listach
        ThemeMuted            = '#8791A5'   # podpisy pól, opisy
        ThemeSubtle           = '#7B8496'   # podtytuły, podpowiedzi
        ThemeFaint            = '#5E6779'   # znak wodny w pustym polu, wersja
        ThemeHover            = '#1D2430'   # wiersz tabeli pod kursorem
        ThemeListHover        = '#1A2029'   # pozycja listy / nawigacji pod kursorem
        ThemeSelection        = '#233354'   # zaznaczony wiersz tabeli
        ThemeSelected         = '#1A2640'   # zaznaczona pozycja listy / nawigacji
        ThemeSelectedBorder   = '#2F4478'   # obramowanie zaznaczenia, kafelek pod kursorem
        ThemeSegmentChecked   = '#28303D'   # aktywny segment (zakładki)
        ThemeMenuHighlight    = '#24304A'   # podświetlona pozycja menu / listy rozwijanej
        ThemeScrollThumb      = '#39414F'   # suwak paska przewijania, ramka pola wyboru
        ThemeTrack            = '#262D39'   # tło pasków postępu
        ThemeOnAccent         = '#FFFFFF'   # tekst na kolorowych przyciskach
        AccentColor           = '#3E6FE0'   # akcent: przycisk główny, zaznaczenie
        ThemeAccentHover      = '#5582EC'   # przycisk główny pod kursorem
        ThemeFocus            = '#4C7DF0'   # fokus / najechanie, wypełnienie paska postępu
        ThemeAccentText       = '#8CB0FF'   # ikony w kolorze akcentu
        ThemeDangerFill       = '#D9475A'   # przycisk DangerPrimary (nieodwracalna akcja)
        ThemeDangerFillHover  = '#E85C6E'
        ThemeDangerText       = '#FF8A95'   # miękki przycisk Danger: tekst
        ThemeDangerSoft       = '#21181C'   # miękki przycisk Danger: tło
        ThemeDangerSoftHover  = '#3A1F26'
        ThemeDangerBorder     = '#4A2A32'
        ThemeSuccess          = '#5EE3AE'   # kropki stanu (ok / ostrzeżenie / błąd)
        ThemeWarning          = '#FFC46B'
        ThemeDanger           = '#FF7A86'
        ThemeInfo             = '#8CC0FF'   # tekst informacyjny
        ThemeDryRun           = '#C3A6FF'   # tryb testowy (para tonów zachowana z 3.1)
        ThemeDryRunBack       = '#251E38'
        ThemeInfoBack         = '#1A2640'   # pary tonów: tło kafelka / pigułki + tekst na nim
        ThemeInfoFore         = '#8CB0FF'
        ThemeOkBack           = '#15291F'
        ThemeOkFore           = '#5EE3AE'
        ThemeOkBorder         = '#24503C'   # miękki przycisk Success (Wznów)
        ThemeWarnBack         = '#2E2616'
        ThemeWarnFore         = '#FFC46B'
        ThemeWarnBorder       = '#4A3B1E'   # miękki przycisk Warning (Pauza)
        ThemeCritBack         = '#2E1A1E'
        ThemeCritFore         = '#FF7A86'
    }
    Light = [ordered]@{
        ThemeBackground       = '#F4F6F9'
        ThemeHeader           = '#FFFFFF'
        ThemeChromeBorder     = '#E2E6EE'
        ThemePanel            = '#FFFFFF'
        ThemeBorder           = '#E2E6EE'
        ThemeRowLine          = '#EEF1F5'
        ThemeRowAlt           = '#FAFBFC'
        ThemeGridHeader       = '#F7F8FA'
        ThemeGridHeaderBorder = '#E2E6EE'
        ThemeTextBoxBg        = '#FFFFFF'
        ThemeFieldBorder      = '#CDD3DE'
        ThemeFieldHover       = '#AEB6C4'
        ThemeButton           = '#F7F8FA'
        ThemeButtonText       = '#1A1F29'
        ThemeGhostText        = '#475467'
        ThemeGhostHover       = '#EEF1F5'
        ThemeChip             = '#EEF1F5'
        ThemePopup            = '#FFFFFF'
        ThemePopupBorder      = '#D5DAE3'
        ThemeTooltip          = '#FFFFFF'
        ThemeToast            = '#FFFFFF'
        ThemeTitle            = '#1A1F29'
        ThemeText             = '#1A1F29'
        ThemeBodyText         = '#344054'
        ThemeGridText         = '#1F2633'
        ThemeMuted            = '#667085'
        ThemeSubtle           = '#667085'
        ThemeFaint            = '#98A2B3'
        ThemeHover            = '#EEF2F8'
        ThemeListHover        = '#F1F4F9'
        ThemeSelection        = '#DCE6FB'
        ThemeSelected         = '#E8EFFC'
        ThemeSelectedBorder   = '#B7CBF5'
        ThemeSegmentChecked   = '#E4E9F2'
        ThemeMenuHighlight    = '#E8EFFC'
        ThemeScrollThumb      = '#C3CAD6'
        ThemeTrack            = '#E2E6EE'
        ThemeOnAccent         = '#FFFFFF'
        AccentColor           = '#3E6FE0'
        ThemeAccentHover      = '#3461CC'
        ThemeFocus            = '#3E6FE0'
        ThemeAccentText       = '#2F5FD0'
        ThemeDangerFill       = '#D92D3F'
        ThemeDangerFillHover  = '#C4283A'
        ThemeDangerText       = '#C4283A'
        ThemeDangerSoft       = '#FDF0F1'
        ThemeDangerSoftHover  = '#FBE1E4'
        ThemeDangerBorder     = '#F3C5CB'
        ThemeSuccess          = '#138A5E'
        ThemeWarning          = '#B86E00'
        ThemeDanger           = '#D92D3F'
        ThemeInfo             = '#2463C9'
        ThemeDryRun           = '#7A3EC8'
        ThemeDryRunBack       = '#F2E9F7'
        ThemeInfoBack         = '#E8EFFC'
        ThemeInfoFore         = '#2463C9'
        ThemeOkBack           = '#E3F4EC'
        ThemeOkFore           = '#0E6E4B'
        ThemeOkBorder         = '#B5E0C9'
        ThemeWarnBack         = '#FDF1DC'
        ThemeWarnFore         = '#915A00'
        ThemeWarnBorder       = '#F0D9A8'
        ThemeCritBack         = '#FCE8EA'
        ThemeCritFore         = '#B4232F'
    }
}

# Systemowe pędzle Windows, z których korzystają domyślne szablony kontrolek (np. róg między
# paskami przewijania w ScrollViewerze był jasnoszarym kwadratem w trybie ciemnym).
# UWAGA: bez InactiveSelectionHighlight(Text)BrushKey - w .NET Framework są to aliasy
# ControlBrushKey/ControlTextBrushKey, więc w słowniku pojawiał się zdublowany klucz
# ("Item has already been added. Key in dictionary: 'ControlBrush'") i żadne okno się nie wczytywało.
$script:StdThemeSystemBrushKeys = [ordered]@{
    ControlBrushKey       = 'ThemePanel'
    WindowBrushKey        = 'ThemeTextBoxBg'
    ControlTextBrushKey   = 'ThemeText'
    WindowTextBrushKey    = 'ThemeText'
    GrayTextBrushKey      = 'ThemeFaint'
    HighlightBrushKey     = 'AccentColor'
    HighlightTextBrushKey = 'ThemeOnAccent'
}

# Czcionka ikon (pierwsza istniejąca w systemie: Windows 11, potem Windows 10).
$script:StdIconFontFamily = 'Segoe Fluent Icons, Segoe MDL2 Assets'

# Stan motywu. Przy ponownym wczytaniu pliku (testy) zostaje, żeby nie zgubić listy otwartych okien.
if (-not (Get-Variable -Name 'StdThemeState' -Scope Script -ErrorAction SilentlyContinue) -or $null -eq $script:StdThemeState) {
    $script:StdThemeState = @{
        Dark          = $true      # bieżący motyw
        Initialized   = $false     # czy motyw został już odczytany z prefs.json albo ustawiony
        Styles        = $null      # wczytany raz słownik stylów (ResourceDictionary)
        StylesHash    = $null      # SHA-256 tekstu stylów, z którego powstał słownik
        StylesSource  = $null      # tekst stylów, dla którego policzono StylesHash
        StylesInline  = $false     # $true: brak obiektu Application, style dokładane do XAML każdego okna
        Windows       = New-Object 'System.Collections.Generic.List[object]'   # WeakReference do okien STD
        DwmReady      = $null      # $true/$false po pierwszej próbie Add-Type dla DWM
        AppIcon       = $null      # ikona okien (BitmapFrame), tworzona przy pierwszym oknie
        AppIconTried  = $false
    }
}

# ---------------------------------------------------------------------------------------------
# Style kontrolek (jeden słownik dla wszystkich okien)
# ---------------------------------------------------------------------------------------------
# Same style, bez pędzli palety: pędzle są dokładane do zasobów każdego okna (Update-StdWindowTheme),
# a style odwołują się do nich przez {DynamicResource} (przełączanie motywu w locie). Style bazowe
# przez {StaticResource} - kolejność ma znaczenie, styl bazowy musi być zdefiniowany wcześniej.
# Pierwsza grupa: ikony i teksty.
$script:StdThemeStylesXaml = @'
    <Style x:Key="Glyph" TargetType="TextBlock">
        <Setter Property="FontFamily" Value="Segoe Fluent Icons, Segoe MDL2 Assets"/>
        <Setter Property="FontSize" Value="14"/>
        <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <!-- Podpis pola: zwykłą czcionką i zwykłą wielkością liter ("Identyfikator / login"). -->
    <Style x:Key="Caption" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="FontSize" Value="12"/>
        <Setter Property="Margin" Value="0,0,0,5"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <!-- Nagłówek sekcji WIELKIMI LITERAMI (grupy zadań, "DYSKI"). Kolor Muted, nie Faint: w motywie
         jasnym Faint ma za mały kontrast, a to główna struktura okna. -->
    <Style x:Key="SectionCaps" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="FontSize" Value="11"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Margin" Value="0,0,0,8"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="MutedText" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="HintText" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeSubtle}"/>
        <Setter Property="FontSize" Value="11.5"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="CardTitle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="FontSize" Value="15"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>
    <Style x:Key="WindowTitle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeTitle}"/>
        <Setter Property="FontSize" Value="18"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="WindowSubtitle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="FontSize" Value="11.5"/>
        <Setter Property="Margin" Value="0,2,0,0"/>
        <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="DialogTitle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeTitle}"/>
        <Setter Property="FontSize" Value="16"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="DialogSubtitle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="Margin" Value="0,3,0,0"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="DialogBody" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource ThemeBodyText}"/>
        <Setter Property="LineHeight" Value="19"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>

    <!-- ===== Powierzchnie: karta, kafelek, przełącznik segmentowy, chip, kafelek ikony ===== -->
    <Style x:Key="Card" TargetType="Border">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="CornerRadius" Value="10"/>
        <Setter Property="Padding" Value="16,14"/>
    </Style>
    <Style x:Key="Tile" TargetType="Border" BasedOn="{StaticResource Card}">
        <Setter Property="Cursor" Value="Hand"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="BorderBrush" Value="{DynamicResource ThemeSelectedBorder}"/></Trigger>
        </Style.Triggers>
    </Style>
    <Style x:Key="SegmentHost" TargetType="Border">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="CornerRadius" Value="8"/>
        <Setter Property="Padding" Value="2"/>
    </Style>
    <Style x:Key="Chip" TargetType="Border">
        <Setter Property="Background" Value="{DynamicResource ThemeChip}"/>
        <Setter Property="CornerRadius" Value="9"/>
        <Setter Property="Padding" Value="9,2"/>
        <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <!-- Kafelek ikony w nagłówku okna dialogowego (tło = ton komunikatu, np. ThemeCritBack). -->
    <Style x:Key="IconTile" TargetType="Border">
        <Setter Property="Width" Value="38"/>
        <Setter Property="Height" Value="38"/>
        <Setter Property="CornerRadius" Value="10"/>
        <Setter Property="Background" Value="{DynamicResource ThemeInfoBack}"/>
        <Setter Property="VerticalAlignment" Value="Top"/>
        <Setter Property="Margin" Value="0,0,14,0"/>
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
    <!-- Domyślny przycisk = drugorzędny (30 px wysokości, odstęp 12,5).
         Najechanie i fokus = niebieskie obramowanie, wciśnięcie = przygaszenie. Warianty mają to samo
         obramowanie 1 px, więc wszystkie przyciski są równej wysokości:
         PrimaryButton (jeden na okno), DangerPrimaryButton (nieodwracalna akcja), miękkie DangerButton,
         SuccessButton i WarningButton (kod podmienia je przy Pauza/Wznów) oraz GhostButton (pasek narzędzi). -->
    <Style TargetType="Button">
        <Setter Property="Background" Value="{DynamicResource ThemeButton}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeButtonText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="12,5"/>
        <Setter Property="MinHeight" Value="30"/>
        <Setter Property="HorizontalContentAlignment" Value="Center"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="SnapsToDevicePixels" Value="True"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                        <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" RecognizesAccessKey="True"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                        <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.8"/></Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.4"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource AccentColor}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource AccentColor}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeOnAccent}"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{DynamicResource ThemeAccentHover}"/></Trigger>
        </Style.Triggers>
    </Style>
    <!-- Miękkie warianty tonów; klucze zostają, bo kod podmienia styl Pauza/Wznów (FindResource). -->
    <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource ThemeOkBack}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeOkBorder}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeOkFore}"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="WarningButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource ThemeWarnBack}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeWarnBorder}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeWarnFore}"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource ThemeDangerSoft}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeDangerBorder}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeDangerText}"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{DynamicResource ThemeDangerSoftHover}"/></Trigger>
        </Style.Triggers>
    </Style>
    <Style x:Key="DangerPrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="{DynamicResource ThemeDangerFill}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeDangerFill}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeOnAccent}"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{DynamicResource ThemeDangerFillHover}"/></Trigger>
        </Style.Triggers>
    </Style>
    <Style x:Key="GhostButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderBrush" Value="Transparent"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeGhostText}"/>
        <Setter Property="Padding" Value="8,4"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeGhostHover}"/>
                <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
            </Trigger>
        </Style.Triggers>
    </Style>

    <!-- ===== Pola tekstowe ===== -->
    <Style TargetType="TextBox">
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="6,4"/>
        <Setter Property="MinHeight" Value="30"/>
        <Setter Property="CaretBrush" Value="{DynamicResource ThemeText}"/>
        <Setter Property="SelectionBrush" Value="{DynamicResource AccentColor}"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="TextBox">
                    <!-- Padding NIE jest tu nakładany na PART_ContentHost: TextBox sam przekazuje swój Padding
                         do wewnętrznego ScrollViewera. Wcześniej Margin="{TemplateBinding Padding}" dublował
                         odstęp (pole 43 px zamiast 31 px), a przy sztywnej wysokości tekst był ucinany w połowie. -->
                    <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" SnapsToDevicePixels="True">
                        <Grid>
                            <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" Focusable="False"
                                          HorizontalScrollBarVisibility="{TemplateBinding ScrollViewer.HorizontalScrollBarVisibility}" VerticalScrollBarVisibility="{TemplateBinding ScrollViewer.VerticalScrollBarVisibility}"/>
                            <!-- Podpowiedź w pustym polu: tekst z Tag (np. Tag="Szukaj..."), wyrównana do tekstu pola -->
                            <Border Padding="{TemplateBinding Padding}" IsHitTestVisible="False">
                                <TextBlock x:Name="wm" Text="{TemplateBinding Tag}" Foreground="{DynamicResource ThemeFaint}" Margin="2,0,0,0" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" Visibility="Collapsed"/>
                            </Border>
                        </Grid>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="Text" Value=""><Setter TargetName="wm" Property="Visibility" Value="Visible"/></Trigger>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFieldHover}"/></Trigger>
                        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFocus}"/></Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <!-- Pole wielowierszowe (np. lista wpisów) i tekst tylko do odczytu (szczegóły wpisu logu). -->
    <Style x:Key="MultiText" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
        <Setter Property="AcceptsReturn" Value="True"/>
        <Setter Property="AcceptsTab" Value="True"/>
        <Setter Property="TextWrapping" Value="NoWrap"/>
        <Setter Property="VerticalContentAlignment" Value="Top"/>
        <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
        <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
        <Setter Property="FontFamily" Value="Consolas"/>
        <Setter Property="FontSize" Value="12.5"/>
    </Style>
    <Style x:Key="ReadOnlyText" TargetType="TextBox" BasedOn="{StaticResource MultiText}">
        <Setter Property="IsReadOnly" Value="True"/>
        <Setter Property="AcceptsTab" Value="False"/>
        <Setter Property="TextWrapping" Value="Wrap"/>
        <Setter Property="HorizontalScrollBarVisibility" Value="Disabled"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderThickness" Value="0"/>
    </Style>
    <Style TargetType="PasswordBox">
        <Setter Property="Background" Value="{DynamicResource ThemeTextBoxBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeFieldBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="6,4"/>
        <Setter Property="MinHeight" Value="30"/>
        <Setter Property="CaretBrush" Value="{DynamicResource ThemeText}"/>
        <Setter Property="SelectionBrush" Value="{DynamicResource AccentColor}"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="PasswordBox">
                    <!-- Jak w TextBox: Padding przekazuje sam PasswordBox, bez dublowania przez Margin. -->
                    <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" SnapsToDevicePixels="True">
                        <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource ThemeFieldHover}"/></Trigger>
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
        <Setter Property="MinHeight" Value="30"/>
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
                        <!-- Wysokość daje MinHeight (30 px), więc bez pionowego marginesu. -->
                        <ContentPresenter IsHitTestVisible="False" Margin="9,0,28,0" VerticalAlignment="Center" HorizontalAlignment="Left"
                                          Content="{TemplateBinding SelectionBoxItem}"
                                          ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                          ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"/>
                        <Popup x:Name="PART_Popup" IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True" Focusable="False" PopupAnimation="Fade">
                            <Border Background="{DynamicResource ThemePopup}" BorderBrush="{DynamicResource ThemePopupBorder}" BorderThickness="1" CornerRadius="7" Margin="0,3,0,0" Padding="3"
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
        <Setter Property="Padding" Value="10,6"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ComboBoxItem">
                    <Border x:Name="bd" Background="Transparent" CornerRadius="5" Padding="{TemplateBinding Padding}">
                        <ContentPresenter VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeSelected}"/></Trigger>
                        <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeMenuHighlight}"/></Trigger>
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
                        <!-- 16 px + Margin 0,1: całe piksele przy skalowaniu 125/150% (bez tego dół ramki bywał ucinany). -->
                        <Border x:Name="box" Width="16" Height="16" Margin="0,1" CornerRadius="4" BorderThickness="1"
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

    <!-- ===== Listy (ListBox): pozycje jak karty ===== -->
    <Style TargetType="ListBox">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Padding" Value="6"/>
        <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ListBox">
                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="10" SnapsToDevicePixels="True">
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
        <Setter Property="Padding" Value="8,6"/>
        <Setter Property="Margin" Value="0,1"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ListBoxItem">
                    <Border x:Name="bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="1" CornerRadius="7" Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeListHover}"/></Trigger>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeSelected}"/>
                            <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource ThemeSelectedBorder}"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Tabele (ListView + GridView) ===== -->
    <Style TargetType="ListView">
        <Setter Property="Background" Value="{DynamicResource ThemePanel}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeGridText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontSize" Value="12.5"/>
        <Setter Property="AlternationCount" Value="2"/>
    </Style>
    <Style TargetType="GridViewColumnHeader">
        <Setter Property="Background" Value="{DynamicResource ThemeGridHeader}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeMuted}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeGridHeaderBorder}"/>
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
    <!-- Wiersz: 28 px, co drugi wiersz ThemeRowAlt (wyzwalacz stylu), najechanie i zaznaczenie w szablonie.
         Bez VirtualizationMode=Recycling - kolumny z polami wyboru (np. w Deinstalatorze) polegają na swoich kontenerach. -->
    <Style TargetType="ListViewItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeGridText}"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemeRowLine}"/>
        <Setter Property="BorderThickness" Value="0,0,0,1"/>
        <Setter Property="Padding" Value="4,4"/>
        <Setter Property="MinHeight" Value="28"/>
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
                        <Trigger Property="IsSelected" Value="True">
                            <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeSelection}"/>
                            <Setter Property="Foreground" Value="{DynamicResource ThemeTitle}"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
        <Style.Triggers>
            <Trigger Property="ItemsControl.AlternationIndex" Value="1"><Setter Property="Background" Value="{DynamicResource ThemeRowAlt}"/></Trigger>
        </Style.Triggers>
    </Style>

    <!-- ===== Zakładki: domyślnie przełącznik segmentowy ===== -->
    <!-- Kolor nagłówka przez TextElement.Foreground prezentera, NIGDY przez TabItem.Foreground: WPF
         przekazuje Foreground zakładki do jej zawartości, więc zabarwiłby całą stronę. -->
    <!-- Foreground tekstu strony (zawartości zakładki) ustawia TabControl i TabItem: bez tego treść
         dziedziczyła systemowy czarny z domyślnego stylu Windows (czarny tekst na ciemnym tle w
         Informacjach o systemie). Kolor nagłówka zakładki idzie osobno, przez TextElement.Foreground. -->
    <Style TargetType="TabItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Padding" Value="14,5"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="TabItem">
                    <Border x:Name="bd" Background="Transparent" CornerRadius="6" Padding="{TemplateBinding Padding}">
                        <ContentPresenter x:Name="hdr" ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True" TextElement.Foreground="{DynamicResource ThemeMuted}"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="hdr" Property="TextElement.Foreground" Value="{DynamicResource ThemeText}"/></Trigger>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeSegmentChecked}"/>
                            <Setter TargetName="hdr" Property="TextElement.Foreground" Value="{DynamicResource ThemeText}"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="TabControl">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Padding" Value="0"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="TabControl">
                    <Grid>
                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                        <Border Style="{StaticResource SegmentHost}" HorizontalAlignment="Left" Margin="0,0,0,14">
                            <StackPanel Orientation="Horizontal" IsItemsHost="True" KeyboardNavigation.TabNavigation="Once" KeyboardNavigation.DirectionalNavigation="Cycle"/>
                        </Border>
                        <ContentPresenter x:Name="PART_SelectedContentHost" Grid.Row="1" ContentSource="SelectedContent"/>
                    </Grid>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <!-- Nawigacja po lewej (NavTabControl): ikona z TabItem.Tag (encja XAML), pasek akcentu 3 px
         przy zaznaczonej pozycji. -->
    <Style x:Key="NavTabItem" TargetType="TabItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Margin" Value="0,1"/>
        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
        <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="TabItem">
                    <Grid>
                        <Border x:Name="b" Background="Transparent" CornerRadius="7" Padding="10,6">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock x:Name="ic" Text="{TemplateBinding Tag}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="14" Width="18" Margin="2,0,11,0" VerticalAlignment="Center" Foreground="{DynamicResource ThemeMuted}"/>
                                <ContentPresenter x:Name="hdr" ContentSource="Header" VerticalAlignment="Center" RecognizesAccessKey="True" TextElement.Foreground="{DynamicResource ThemeGhostText}"/>
                            </StackPanel>
                        </Border>
                        <Border x:Name="bar" Width="3" CornerRadius="2" Background="{DynamicResource ThemeFocus}" HorizontalAlignment="Left" Margin="0,8" Visibility="Hidden"/>
                    </Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="b" Property="Background" Value="{DynamicResource ThemeListHover}"/>
                            <Setter TargetName="hdr" Property="TextElement.Foreground" Value="{DynamicResource ThemeText}"/>
                        </Trigger>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter TargetName="b" Property="Background" Value="{DynamicResource ThemeSelected}"/>
                            <Setter TargetName="bar" Property="Visibility" Value="Visible"/>
                            <Setter TargetName="hdr" Property="TextElement.Foreground" Value="{DynamicResource ThemeTitle}"/>
                            <Setter TargetName="ic" Property="Foreground" Value="{DynamicResource ThemeAccentText}"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style x:Key="NavTabControl" TargetType="TabControl">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Padding" Value="0"/>
        <Setter Property="TabStripPlacement" Value="Left"/>
        <Setter Property="ItemContainerStyle" Value="{StaticResource NavTabItem}"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="TabControl">
                    <Grid>
                        <Grid.ColumnDefinitions><ColumnDefinition Width="196"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                        <!-- Przewijana lista: przy 1366x768 i 150% (472 DIP wysokości) 9 pozycji ledwo się mieści. -->
                        <Border Style="{StaticResource Card}" Padding="6" Margin="0,0,16,0" VerticalAlignment="Top">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
                                <StackPanel IsItemsHost="True" KeyboardNavigation.TabNavigation="Once" KeyboardNavigation.DirectionalNavigation="Cycle"/>
                            </ScrollViewer>
                        </Border>
                        <ContentPresenter x:Name="PART_SelectedContentHost" Grid.Column="1" ContentSource="SelectedContent"/>
                    </Grid>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <!-- Podpis + pole: StackPanel z Caption i kontrolką, układany w WrapPanel. -->
    <Style x:Key="FieldBox" TargetType="StackPanel">
        <Setter Property="Margin" Value="0,0,14,12"/>
    </Style>

    <!-- ===== Pasek postępu (6 px, wypełnienie ThemeFocus) ===== -->
    <Style TargetType="ProgressBar">
        <Setter Property="Background" Value="{DynamicResource ThemeTrack}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeFocus}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Height" Value="6"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ProgressBar">
                    <Grid>
                        <Border x:Name="PART_Track" CornerRadius="3" Background="{TemplateBinding Background}"/>
                        <Border x:Name="PART_Indicator" CornerRadius="3" Background="{TemplateBinding Foreground}" HorizontalAlignment="Left"/>
                    </Grid>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Podpowiedzi i menu kontekstowe ===== -->
    <Style TargetType="ToolTip">
        <Setter Property="Background" Value="{DynamicResource ThemeTooltip}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemePopupBorder}"/>
        <Setter Property="Padding" Value="9,6"/>
        <Setter Property="FontFamily" Value="Segoe UI"/>
        <Setter Property="FontSize" Value="12"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ToolTip">
                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}" MaxWidth="460">
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
        <Setter Property="Background" Value="{DynamicResource ThemePopup}"/>
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource ThemePopupBorder}"/>
        <Setter Property="FontFamily" Value="Segoe UI"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ContextMenu">
                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="4">
                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Cycle"/>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <!-- Pozycja menu z kolumną ikony 26 px (MenuItem.Icon = TextBlock ze stylem Glyph). Bez podmenu. -->
    <Style TargetType="MenuItem">
        <Setter Property="Foreground" Value="{DynamicResource ThemeText}"/>
        <Setter Property="Padding" Value="8,6"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="MenuItem">
                    <Border x:Name="bd" Background="Transparent" CornerRadius="6" Padding="{TemplateBinding Padding}" MinWidth="200">
                        <Grid>
                            <Grid.ColumnDefinitions><ColumnDefinition Width="26"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <ContentPresenter ContentSource="Icon" VerticalAlignment="Center" TextElement.Foreground="{DynamicResource ThemeMuted}"/>
                            <ContentPresenter Grid.Column="1" ContentSource="Header" RecognizesAccessKey="True" VerticalAlignment="Center"/>
                        </Grid>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeMenuHighlight}"/></Trigger>
                        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <!-- Separator w menu: WPF szuka go pod kluczem MenuItem.SeparatorStyleKey, nie jako stylu domyślnego. -->
    <Style x:Key="{x:Static MenuItem.SeparatorStyleKey}" TargetType="Separator">
        <Setter Property="Margin" Value="6,4"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Separator">
                    <Border Height="1" Background="{DynamicResource ThemeFieldBorder}"/>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- ===== Style z okien STD 3.1, wspólne w STD 4 (te same klucze) ===== -->
    <!-- Przycisk na pasku nad listą (Edytuj / Usuń): zwykły przycisk z odstępem z lewej. -->
    <Style x:Key="BarButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
        <Setter Property="Margin" Value="6,0,0,0"/>
    </Style>
    <!-- Podpowiedź pod listą. -->
    <Style x:Key="ListHint" TargetType="TextBlock" BasedOn="{StaticResource MutedText}">
        <Setter Property="FontSize" Value="11.5"/>
        <Setter Property="Margin" Value="0,2,0,0"/>
    </Style>
    <!-- Karta na stronie zakładki i przewijana strona zakładki. -->
    <Style x:Key="TabCard" TargetType="Border" BasedOn="{StaticResource Card}">
        <Setter Property="Padding" Value="16,14,2,2"/>
        <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="TabScroll" TargetType="ScrollViewer">
        <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
        <Setter Property="HorizontalScrollBarVisibility" Value="Disabled"/>
    </Style>
    <!-- Kafelek z wartością (np. karta Komputer): zwykły i klikalny (kopiowanie do schowka). -->
    <Style x:Key="ChipBorder" TargetType="Border" BasedOn="{StaticResource Card}">
        <Setter Property="Padding" Value="12,8"/>
        <Setter Property="Margin" Value="0,0,10,10"/>
    </Style>
    <Style x:Key="ChipBorderClickable" TargetType="Border" BasedOn="{StaticResource Tile}">
        <Setter Property="Padding" Value="12,8"/>
        <Setter Property="Margin" Value="0,0,10,10"/>
        <Setter Property="ToolTip" Value="Kliknij, aby skopiować"/>
    </Style>
    <!-- Bez zawijania: w WrapPanel ramka ma dokładnie szerokość napisu, a po zaokrągleniu układu
         ostatnie słowo spadało do drugiej, uciętej linii. -->
    <Style x:Key="ChipLabel" TargetType="TextBlock" BasedOn="{StaticResource SectionCaps}">
        <Setter Property="Margin" Value="0"/>
        <Setter Property="TextWrapping" Value="NoWrap"/>
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
'@

# ---------------------------------------------------------------------------------------------
# Paleta: odczyt
# ---------------------------------------------------------------------------------------------

function Get-StdThemePalette {
    <#
    .SYNOPSIS
    Paleta kolorów ([ordered] klucz -> HEX) bieżącego motywu albo motywu podanego w -Theme.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Dark', 'Light')]
        [string]$Theme
    )
    if ([string]::IsNullOrEmpty($Theme)) {
        if ($script:StdThemeState.Dark) { $Theme = 'Dark' } else { $Theme = 'Light' }
    }
    # Przecinek: słownik ma wyjść z funkcji jako jeden obiekt.
    return , $script:StdThemePalettes[$Theme]
}

function Get-StdThemeColor {
    <#
    .SYNOPSIS
    Kolor (HEX, np. '#3E6FE0') z palety - dla miejsc, w których kolor trafia do danych, a nie do XAML.
    Nieznany klucz to błąd w kodzie okna: [STD-UI-THEME].
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Key,
        [ValidateSet('Dark', 'Light')]
        [string]$Theme
    )
    if ([string]::IsNullOrEmpty($Theme)) { $palette = Get-StdThemePalette } else { $palette = Get-StdThemePalette -Theme $Theme }
    if (-not $palette.Contains($Key)) {
        throw "[STD-UI-THEME] W palecie motywu nie ma koloru '$Key'."
    }
    return [string]$palette[$Key]
}

function Test-StdThemeDark {
    <#
    .SYNOPSIS
    $true, gdy bieżący motyw jest ciemny.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    return [bool]$script:StdThemeState.Dark
}

function New-StdThemeBrush {
    <#
    .SYNOPSIS
    Zamrożony SolidColorBrush z koloru HEX. Zwraca "surowy" obiekt .NET (BaseObject), nie PSObject:
    opakowanie PSObject w ResourceDictionary dawało błąd "'#FF...' is not a valid value for property".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Color
    )
    $brush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush.PSObject.BaseObject
}

# ---------------------------------------------------------------------------------------------
# Słownik stylów: wczytany raz
# ---------------------------------------------------------------------------------------------

function Get-StdThemeStylesHash {
    # SHA-256 tekstu stylów (HEX). Po zmianie tekstu (ponowne wczytanie pliku) słownik powstaje od nowa.
    $text = $script:StdThemeStylesXaml
    $state = $script:StdThemeState
    if ($null -ne $state.StylesHash -and [object]::ReferenceEquals($state.StylesSource, $text)) { return $state.StylesHash }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash((New-Object System.Text.UTF8Encoding($false)).GetBytes($text))
    } finally {
        $sha.Dispose()
    }
    return ([System.BitConverter]::ToString($bytes) -replace '-', '')
}

function Get-StdThemeApplication {
    # Obiekt Application, w którego zasobach leży słownik stylów. Dzięki niemu {StaticResource Card}
    # w XAML okna znajduje styl już w chwili wczytywania XAML, bez ponownego parsowania stylów.
    # Program STD nie ma własnego Application (okna pokazuje ShowDialog), więc tworzymy go raz,
    # z ShutdownMode = OnExplicitShutdown: zamknięcie ostatniego okna (np. ekranu logowania przed
    # oknem głównym) nie może zatrzymać obsługi okien w tym wątku.
    # Cudzego Application (np. edytor PowerShell ISE) nie ruszamy - wtedy $null i style trafiają
    # do XAML każdego okna (jak w STD 3.1).
    $app = [System.Windows.Application]::Current
    if ($null -ne $app) {
        if (-not $app.Dispatcher.CheckAccess()) { return $null }
        if ($app.Resources.Contains('StdThemeApplication')) { return $app }
        return $null
    }
    try {
        $app = New-Object System.Windows.Application
        $app.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
        $app.Resources['StdThemeApplication'] = 'STD'
        return $app
    } catch {
        Write-StdLog -Level Warn -Source 'Std.Theme' -Message "Nie udało się utworzyć obiektu Application dla stylów okien ($($_.Exception.Message)). Style będą dołączane do każdego okna osobno."
        return $null
    }
}

function Get-StdThemeStyles {
    <#
    .SYNOPSIS
    Słownik stylów STD (ResourceDictionary) wczytany z $script:StdThemeStylesXaml tylko raz.
    Zwraca $null, gdy w procesie nie da się użyć obiektu Application - wtedy New-StdWindow
    dokłada style do XAML każdego okna.
    #>
    [CmdletBinding()]
    param()
    $state = $script:StdThemeState
    $hash = Get-StdThemeStylesHash
    if ($state.StylesHash -eq $hash) {
        if ($state.StylesInline) { return $null }
        # Przecinek: ResourceDictionary jest kolekcją i nie może się rozsypać na elementy.
        if ($null -ne $state.Styles) { return , $state.Styles }
    }
    $state.Styles = $null
    $state.StylesInline = $false
    $state.StylesHash = $hash
    $state.StylesSource = $script:StdThemeStylesXaml

    $app = Get-StdThemeApplication
    if ($null -eq $app) {
        $state.StylesInline = $true
        return $null
    }
    # Słownik mógł już powstać w tym procesie (np. drugi plik testów wczytał okna ponownie):
    # ten sam tekst = ten sam słownik, starszy tekst = słownik do usunięcia.
    $dict = $null
    foreach ($merged in @($app.Resources.MergedDictionaries)) {
        if ($merged.Contains('StdThemeStylesHash')) {
            if ($null -eq $dict -and [string]$merged['StdThemeStylesHash'] -eq $hash) {
                $dict = $merged
            } else {
                [void]$app.Resources.MergedDictionaries.Remove($merged)
            }
        }
    }
    if ($null -eq $dict) {
        $text = '<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">' +
            "`n" + $script:StdThemeStylesXaml + "`n" + '</ResourceDictionary>'
        try {
            $dict = [System.Windows.Markup.XamlReader]::Parse($text)
        } catch {
            throw "[STD-UI-XAML] Nie udało się wczytać stylów okien: $($_.Exception.Message)"
        }
        $dict['StdThemeStylesHash'] = $hash
        $app.Resources.MergedDictionaries.Add($dict)
    }
    $state.Styles = $dict
    return , $dict
}

function Get-StdThemeResourcesXaml {
    <#
    .SYNOPSIS
    Pełny słownik motywu jako tekst XAML: pędzle bieżącej palety, pędzle systemowe i style.
    Używany, gdy style nie mogą leżeć w obiekcie Application (wtedy trafiają do XAML okna).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $palette = Get-StdThemePalette
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">')
    foreach ($key in $palette.Keys) {
        [void]$sb.AppendLine("    <SolidColorBrush x:Key=`"$key`" Color=`"$($palette[$key])`"/>")
    }
    foreach ($sysKey in $script:StdThemeSystemBrushKeys.Keys) {
        [void]$sb.AppendLine("    <SolidColorBrush x:Key=`"{x:Static SystemColors.$sysKey}`" Color=`"$($palette[$script:StdThemeSystemBrushKeys[$sysKey]])`"/>")
    }
    [void]$sb.AppendLine($script:StdThemeStylesXaml)
    [void]$sb.AppendLine('</ResourceDictionary>')
    return $sb.ToString()
}

function Add-StdThemeXmlResources {
    # Dokłada pełny słownik motywu do <Window.Resources> dokumentu XAML (jako MergedDictionary, więc
    # lokalne style okna mogą z niego dziedziczyć: BasedOn="{StaticResource ...}"). Działa na kopii
    # dokumentu przekazanej przez New-StdWindow. Tylko dla przypadku bez obiektu Application.
    param([Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document)
    $ns = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    $root = $Document.DocumentElement
    $resNode = $null
    foreach ($child in $root.ChildNodes) {
        if ($child.LocalName -eq 'Window.Resources') { $resNode = $child; break }
    }
    if ($null -eq $resNode) {
        $resNode = $Document.CreateElement('Window.Resources', $ns)
        [void]$root.PrependChild($resNode)
    }
    # Okno może już mieć jawny <ResourceDictionary> w Window.Resources - wtedy dokładamy do niego.
    $dict = $null
    $elements = @($resNode.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
    if ($elements.Count -eq 1 -and $elements[0].LocalName -eq 'ResourceDictionary') {
        $dict = $elements[0]
    } else {
        $dict = $Document.CreateElement('ResourceDictionary', $ns)
        while ($resNode.HasChildNodes) { [void]$dict.AppendChild($resNode.FirstChild) }
        [void]$resNode.AppendChild($dict)
    }
    $merged = $null
    foreach ($child in $dict.ChildNodes) {
        if ($child.LocalName -eq 'ResourceDictionary.MergedDictionaries') { $merged = $child; break }
    }
    if ($null -eq $merged) {
        $merged = $Document.CreateElement('ResourceDictionary.MergedDictionaries', $ns)
        [void]$dict.PrependChild($merged)
    }
    $themeDoc = New-Object System.Xml.XmlDocument
    $themeDoc.LoadXml((Get-StdThemeResourcesXaml))
    [void]$merged.PrependChild($Document.ImportNode($themeDoc.DocumentElement, $true))
}

# ---------------------------------------------------------------------------------------------
# Okna STD: lista otwartych, motyw okna, pasek tytułu
# ---------------------------------------------------------------------------------------------

function Get-StdOpenWindows {
    <#
    .SYNOPSIS
    Okna utworzone przez New-StdWindow, które jeszcze nie zostały zamknięte (najstarsze pierwsze).
    #>
    [CmdletBinding()]
    param()
    $result = New-Object 'System.Collections.Generic.List[object]'
    $list = $script:StdThemeState.Windows
    for ($i = $list.Count - 1; $i -ge 0; $i--) {
        $target = $list[$i].Target
        if ($null -eq $target) { $list.RemoveAt($i) }
    }
    foreach ($ref in $list) {
        $target = $ref.Target
        if ($null -ne $target) { $result.Add($target) }
    }
    return , $result
}

function Get-StdActiveWindow {
    <#
    .SYNOPSIS
    Okno STD, w którym użytkownik właśnie pracuje (aktywne i widoczne). Gdy żadne nie jest aktywne
    (technik przełączył się np. na przeglądarkę): ostatnie widoczne, niezminimalizowane okno STD.
    Brak takiego okna: $null.
    #>
    [CmdletBinding()]
    param()
    $windows = Get-StdOpenWindows
    foreach ($w in $windows) {
        if ($w.IsActive -and $w.IsVisible) { return $w }
    }
    for ($i = $windows.Count - 1; $i -ge 0; $i--) {
        $w = $windows[$i]
        if ($w.IsVisible -and $w.WindowState -ne [System.Windows.WindowState]::Minimized) { return $w }
    }
    return $null
}

function Update-StdWindowTheme {
    <#
    .SYNOPSIS
    Wstawia pędzle bieżącej palety (i pędzle systemowe) do zasobów okna. Kolory w XAML są podpięte
    przez {DynamicResource ...}, więc okno przemalowuje się samo, także już otwarte.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Windows.FrameworkElement]$Window
    )
    $palette = Get-StdThemePalette
    $resources = $Window.Resources
    foreach ($key in $palette.Keys) {
        $resources[$key] = (New-StdThemeBrush -Color $palette[$key]).PSObject.BaseObject
    }
    foreach ($sysKey in $script:StdThemeSystemBrushKeys.Keys) {
        $resources[[System.Windows.SystemColors]::$sysKey] = (New-StdThemeBrush -Color $palette[$script:StdThemeSystemBrushKeys[$sysKey]]).PSObject.BaseObject
    }
}

function Set-StdWindowTitleBarTheme {
    <#
    .SYNOPSIS
    Pasek tytułu okna w kolorach motywu: tryb ciemny DWM (atrybut 20, w starszych Windows 10 - 19)
    i kolor paska = ThemeHeader (atrybut 35, tylko Windows 11). Bez uchwytu okna (okno jeszcze
    niepokazane) albo na starszym systemie nic nie robi - nigdy nie rzuca wyjątku.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Windows.Window]$Window
    )
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper $Window).Handle
        if ($hwnd -eq [IntPtr]::Zero) { return }
        $state = $script:StdThemeState
        if ($null -eq $state.DwmReady) {
            $state.DwmReady = $false
            # Typ może już istnieć: testy i podgląd okien wczytują plik kilka razy w tym samym procesie.
            if (-not ('StdUi.Dwm' -as [type])) {
                Add-Type -Namespace StdUi -Name Dwm -ErrorAction Stop -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("dwmapi.dll")]
public static extern int DwmSetWindowAttribute(System.IntPtr hwnd, int attribute, ref int value, int size);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetWindowPos(System.IntPtr hWnd, System.IntPtr after, int x, int y, int cx, int cy, uint flags);
'@
            }
            $state.DwmReady = $true
        }
        if (-not $state.DwmReady) { return }
        $dark = 0
        if ($state.Dark) { $dark = 1 }
        if ([StdUi.Dwm]::DwmSetWindowAttribute($hwnd, 20, [ref]$dark, 4) -ne 0) { [void][StdUi.Dwm]::DwmSetWindowAttribute($hwnd, 19, [ref]$dark, 4) }
        $c = [System.Windows.Media.ColorConverter]::ConvertFromString((Get-StdThemeColor -Key 'ThemeHeader'))
        # COLORREF 0x00BBGGRR liczony arytmetycznie (literał 0xFFFFFFFF to w PS 5.1 liczba -1).
        $caption = [int]$c.R + ([int]$c.G * 256) + ([int]$c.B * 65536)
        [void][StdUi.Dwm]::DwmSetWindowAttribute($hwnd, 35, [ref]$caption, 4)
        # SWP_NOSIZE|NOMOVE|NOZORDER|NOACTIVATE|FRAMECHANGED: Windows 10 przemalowuje pasek dopiero po zmianie ramki.
        [void][StdUi.Dwm]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, 0, 0, [uint32]0x37)
    } catch { }
}

# ---------------------------------------------------------------------------------------------
# Wybór motywu (prefs.json)
# ---------------------------------------------------------------------------------------------

function Initialize-StdTheme {
    <#
    .SYNOPSIS
    Odczytuje motyw z prefs.json (pole "theme": "Dark" albo "Light") i ustawia go dla kolejnych okien.
    Brak pliku albo nieznana wartość: motyw ciemny. Zwraca $true dla motywu ciemnego.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $dark = $true
    try {
        $prefs = Read-StdJsonFile -Path (Get-StdPath -Kind Prefs)
        if ($prefs -is [System.Collections.IDictionary] -and $prefs.Contains('theme')) {
            $value = $prefs['theme']
            if ($value -is [bool]) {
                $dark = $value
            } elseif ([string]$value -eq 'Light') {
                $dark = $false
            } elseif ([string]$value -ne 'Dark') {
                Write-StdLog -Level Warn -Source 'Std.Theme' -Message "Nieznany motyw w prefs.json: '$value'. Używam motywu ciemnego."
            }
        }
    } catch {
        Write-StdLog -Level Warn -Source 'Std.Theme' -Message "Nie udało się odczytać motywu z prefs.json: $($_.Exception.Message)"
    }
    $script:StdThemeState.Dark = $dark
    $script:StdThemeState.Initialized = $true
    return $dark
}

function Set-StdTheme {
    <#
    .SYNOPSIS
    Zmienia motyw we wszystkich otwartych oknach STD (pędzle i pasek tytułu) i zapisuje go w prefs.json
    (pole "theme"; pozostałe pola pliku zostają). -NoSave: bez zapisu (podgląd okien, testy).
    Błąd zapisu nie przerywa zmiany motywu - trafia do logu jako ostrzeżenie.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [bool]$Dark,
        [switch]$NoSave
    )
    $script:StdThemeState.Dark = $Dark
    $script:StdThemeState.Initialized = $true
    foreach ($w in (Get-StdOpenWindows)) {
        try {
            Update-StdWindowTheme -Window $w
            Set-StdWindowTitleBarTheme -Window $w
        } catch {
            Write-StdLog -Level Warn -Source 'Std.Theme' -Message "Nie udało się zmienić motywu okna: $($_.Exception.Message)"
        }
    }
    if ($NoSave) { return }
    $themeName = 'Light'
    if ($Dark) { $themeName = 'Dark' }
    try {
        $path = Get-StdPath -Kind Prefs
        $prefs = Read-StdJsonFile -Path $path
        if ($prefs -isnot [System.Collections.IDictionary]) { $prefs = [ordered]@{} }
        $prefs['theme'] = $themeName
        Write-StdJsonFile -Path $path -InputObject $prefs
    } catch {
        Write-StdLog -Level Warn -Source 'Std.Theme' -Message "Nie udało się zapisać motywu w prefs.json: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------------------------
# Okno z XAML
# ---------------------------------------------------------------------------------------------

function Get-StdAppIconCached {
    # Ikona okien tworzona raz (przy pierwszym oknie). Błąd = domyślna ikona Windows, bez ponawiania.
    $state = $script:StdThemeState
    if ($null -eq $state.AppIcon -and -not $state.AppIconTried) {
        $state.AppIconTried = $true
        try {
            $state.AppIcon = New-StdAppIcon
        } catch {
            Write-StdLog -Level Warn -Source 'Std.Theme' -Message "Nie udało się narysować ikony STD: $($_.Exception.Message)"
        }
    }
    return $state.AppIcon
}

function New-StdWindowClosedHandler {
    # Handler Closed: usuwa okno z listy otwartych okien. Domknięcie (GetNewClosure) zabiera tylko
    # listę, więc handler nie zależy od zakresu, w którym okno zostanie zamknięte.
    param([Parameter(Mandatory = $true)][System.Collections.Generic.List[object]]$WindowList)
    return {
        param($s, $e)
        try {
            for ($i = $WindowList.Count - 1; $i -ge 0; $i--) {
                $target = $WindowList[$i].Target
                if ($null -eq $target -or [object]::ReferenceEquals($target, $s)) { $WindowList.RemoveAt($i) }
            }
        } catch { }
    }.GetNewClosure()
}

function New-StdWindow {
    <#
    .SYNOPSIS
    Buduje okno WPF z XAML w motywie STD i zwraca je (bez pokazywania).
    .DESCRIPTION
    - słownik stylów jest wczytany raz (Get-StdThemeStyles) i dołączany do zasobów okna, a pędzle
      bieżącej palety trafiają do zasobów okna (Update-StdWindowTheme); XAML okna używa stylów przez
      {StaticResource ...} i kolorów przez {DynamicResource ...};
    - ostre krawędzie i tekst przy skalowaniu 125-150%: UseLayoutRounding, TextFormattingMode Display,
      domyślnie czcionka Segoe UI 13 (chyba że XAML okna podaje własną);
    - ikona STD, pasek tytułu w kolorach motywu (po utworzeniu uchwytu okna);
    - okno nigdy większe niż obszar roboczy ekranu (świeże komputery bez sterownika grafiki 1024x768,
      laptopy ze skalowaniem 150%: wcześniej dół okna z przyciskami wychodził poza ekran);
    - -Owner: okno podległe (wyśrodkowane na właścicielu, zawsze nad nim). Gdy właściciel nie jest
      aktywny (technik pracuje w innym programie), okno jest na wierzchu, żeby nie zginęło pod innymi.
    Błąd w XAML: [STD-UI-XAML].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [xml]$Xaml,
        [System.Windows.Window]$Owner
    )
    $state = $script:StdThemeState
    if (-not $state.Initialized) { [void](Initialize-StdTheme) }
    $styles = Get-StdThemeStyles
    $doc = $Xaml
    if ($null -eq $styles) {
        # Kopia dokumentu: [xml] przychodzi przez referencję, a dokładamy do niego Window.Resources.
        $doc = $Xaml.Clone()
        Add-StdThemeXmlResources -Document $doc
    }
    try {
        $win = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $doc))
    } catch {
        $ex = $_.Exception
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        throw "[STD-UI-XAML] Nie udało się zbudować okna z XAML: $($ex.Message)"
    }
    if ($win -isnot [System.Windows.Window]) {
        throw '[STD-UI-XAML] Element główny XAML okna musi być typu Window.'
    }
    if ($null -ne $styles) { $win.Resources.MergedDictionaries.Add($styles) }
    Update-StdWindowTheme -Window $win

    $win.UseLayoutRounding = $true
    $win.SnapsToDevicePixels = $true
    [System.Windows.Media.TextOptions]::SetTextFormattingMode($win, [System.Windows.Media.TextFormattingMode]::Display)
    if ($win.ReadLocalValue([System.Windows.Controls.Control]::FontFamilyProperty) -eq [System.Windows.DependencyProperty]::UnsetValue) {
        $win.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe UI'
    }
    if ($win.ReadLocalValue([System.Windows.Controls.Control]::FontSizeProperty) -eq [System.Windows.DependencyProperty]::UnsetValue) {
        $win.FontSize = 13
    }
    # Pasek tytułu w kolorach motywu, gdy tylko okno dostanie uchwyt (HWND).
    $win.Add_SourceInitialized({
        param($s, $e)
        try { Set-StdWindowTitleBarTheme -Window $s } catch { }
    })
    $icon = Get-StdAppIconCached
    if ($null -ne $icon) { try { $win.Icon = $icon } catch { } }

    try {
        $work = [System.Windows.SystemParameters]::WorkArea
        if ($win.MinWidth -gt $work.Width) { $win.MinWidth = $work.Width }
        if ($win.MinHeight -gt $work.Height) { $win.MinHeight = $work.Height }
        if (-not [double]::IsNaN($win.Width) -and $win.Width -gt $work.Width) { $win.Width = $work.Width }
        if (-not [double]::IsNaN($win.Height) -and $win.Height -gt $work.Height) { $win.Height = $work.Height }
        if ($win.SizeToContent -ne [System.Windows.SizeToContent]::Manual) { $win.MaxHeight = $work.Height; $win.MaxWidth = $work.Width }
    } catch { }

    if ($null -ne $Owner -and -not [object]::ReferenceEquals($Owner, $win)) {
        # Właściciel, który nie był jeszcze pokazany, daje wyjątek - wtedy okno zostaje samodzielne.
        try { $win.Owner = $Owner } catch { }
    }
    if ($null -ne $win.Owner) {
        if (-not $win.Owner.IsActive) { $win.Topmost = $true }
    } elseif ($win.WindowStartupLocation -eq [System.Windows.WindowStartupLocation]::CenterOwner) {
        # CenterOwner bez właściciela ląduje w lewym górnym rogu ekranu.
        $win.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
    }

    # Lista otwartych okien (Set-StdTheme). WeakReference: okno zbudowane, ale nigdy niepokazane
    # (testy, podgląd) nie zostaje w pamięci.
    $state.Windows.Add((New-Object System.WeakReference -ArgumentList $win))
    $win.Add_Closed((New-StdWindowClosedHandler -WindowList $state.Windows))
    return $win
}

# ---------------------------------------------------------------------------------------------
# Okno komunikatu
# ---------------------------------------------------------------------------------------------

$script:StdMessageXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Width="480" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" ShowInTaskbar="False"
        Background="{DynamicResource ThemeHeader}" Foreground="{DynamicResource ThemeText}" FontFamily="Segoe UI">
    <Grid Margin="22,20,22,18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Name="iconTile" Style="{StaticResource IconTile}">
                <TextBlock Name="txtIcon" Style="{StaticResource Glyph}" FontSize="17" HorizontalAlignment="Center"/>
            </Border>
            <TextBlock Name="txtTitle" Grid.Column="1" Style="{StaticResource DialogTitle}" VerticalAlignment="Center"/>
        </Grid>
        <ScrollViewer Grid.Row="1" MaxHeight="420" VerticalScrollBarVisibility="Auto">
            <TextBlock Name="txtMessage" Style="{StaticResource DialogBody}"/>
        </ScrollViewer>
        <StackPanel Name="spButtons" Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,18,0,0"/>
    </Grid>
</Window>
'@

function New-StdMessageWindow {
    <#
    .SYNOPSIS
    Buduje (bez pokazywania) okno komunikatu z Show-StdMessage: kafelek z ikoną w kolorze tonu, tytuł,
    treść i przyciski od prawej (drugorzędne najpierw, główny - Enter - na końcu). Wynik kliknięcia
    trafia do Tag okna ([System.Windows.MessageBoxResult]). Osobna funkcja dla podglądu okien i testów.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyString()]
        [string]$Message,
        [Parameter(Position = 1)]
        [string]$Title = 'STD',
        [ValidateSet('OK', 'OKCancel', 'YesNo', 'YesNoCancel')]
        [string]$Button = 'OK',
        [ValidateSet('Information', 'Warning', 'Error', 'Question', 'Success')]
        [string]$Image = 'Information',
        [System.Windows.Window]$Owner,
        [switch]$Danger
    )
    if ($null -eq $Owner) { $Owner = Get-StdActiveWindow }
    $dlg = New-StdWindow -Xaml $script:StdMessageXaml -Owner $Owner
    if ($null -eq $dlg.Owner) {
        # Brak widocznego właściciela: komunikat nie może zginąć pod innymi oknami.
        $dlg.Topmost = $true
        $dlg.ShowInTaskbar = $true
        $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
    }
    # Tytuł i treść ustawiamy po wczytaniu XAML - znak & albo " w tekście nie zepsuje XML-a.
    $dlg.Title = $Title
    $dlg.FindName('txtTitle').Text = $Title
    $dlg.FindName('txtMessage').Text = $Message

    # Ton (para kolorów z palety Theme<Ton>Back / Theme<Ton>Fore) i znak ikony.
    $iconSpec = switch ($Image) {
        'Error'    { @{ Tone = 'Crit'; Glyph = 'EA39' } }
        'Warning'  { @{ Tone = 'Warn'; Glyph = 'E7BA' } }
        'Question' { @{ Tone = 'Info'; Glyph = 'E897' } }
        'Success'  { @{ Tone = 'Ok'; Glyph = 'E930' } }
        default    { @{ Tone = 'Info'; Glyph = 'E946' } }
    }
    $dlg.FindName('iconTile').SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, "Theme$($iconSpec.Tone)Back")
    $txtIcon = $dlg.FindName('txtIcon')
    $txtIcon.Text = Get-StdGlyph -Code $iconSpec.Glyph
    $txtIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "Theme$($iconSpec.Tone)Fore")

    # Kolejność: drugorzędne najpierw, główny (Enter) skrajnie po prawej; Esc = przycisk "Cancel".
    $ok = [System.Windows.MessageBoxResult]::OK
    $cancel = [System.Windows.MessageBoxResult]::Cancel
    $yes = [System.Windows.MessageBoxResult]::Yes
    $no = [System.Windows.MessageBoxResult]::No
    $specs = @(switch ($Button) {
        'OKCancel' {
            @{ Text = 'Anuluj'; Result = $cancel; IsDefault = $false; IsCancel = $true }
            @{ Text = 'OK'; Result = $ok; IsDefault = $true; IsCancel = $false }
        }
        'YesNo' {
            @{ Text = 'Nie'; Result = $no; IsDefault = $false; IsCancel = $true }
            @{ Text = 'Tak'; Result = $yes; IsDefault = $true; IsCancel = $false }
        }
        'YesNoCancel' {
            @{ Text = 'Anuluj'; Result = $cancel; IsDefault = $false; IsCancel = $true }
            @{ Text = 'Nie'; Result = $no; IsDefault = $false; IsCancel = $false }
            @{ Text = 'Tak'; Result = $yes; IsDefault = $true; IsCancel = $false }
        }
        default {
            @{ Text = 'OK'; Result = $ok; IsDefault = $true; IsCancel = $true }
        }
    })
    $spButtons = $dlg.FindName('spButtons')
    $primaryStyle = 'PrimaryButton'
    if ($Danger) { $primaryStyle = 'DangerPrimaryButton' }
    foreach ($spec in $specs) {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = $spec.Text
        $btn.IsDefault = $spec.IsDefault
        $btn.IsCancel = $spec.IsCancel
        $btn.Tag = $spec.Result
        $btn.MinWidth = 96
        if ($spButtons.Children.Count -gt 0) { $btn.Margin = New-Object System.Windows.Thickness(8, 0, 0, 0) }
        if ($spec.IsDefault) { $btn.Style = $dlg.FindResource($primaryStyle) }
        # Handler bez zmiennych z zewnątrz: okno bierze z przycisku, wynik zapisuje w Tag okna.
        $btn.Add_Click({
            param($s, $e)
            try {
                $w = [System.Windows.Window]::GetWindow($s)
                if ($null -ne $w) {
                    $w.Tag = $s.Tag
                    $w.Close()
                }
            } catch { }
        })
        [void]$spButtons.Children.Add($btn)
    }
    return $dlg
}

function Show-StdMessage {
    <#
    .SYNOPSIS
    Okno komunikatu w motywie STD (zamiast MessageBox). Zwraca [System.Windows.MessageBoxResult].
    Zamknięcie krzyżykiem = przycisk Anuluj / Nie (dla samego OK: OK).
    .PARAMETER Image
    Information (domyślnie), Question, Warning, Error albo Success (zielony "ok").
    .PARAMETER Danger
    Główny przycisk czerwony (DangerPrimaryButton) - dla nieodwracalnych operacji.
    .PARAMETER Owner
    Okno nadrzędne. Bez niego: okno STD, w którym użytkownik właśnie pracuje.
    #>
    [CmdletBinding()]
    [OutputType('System.Windows.MessageBoxResult')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyString()]
        [string]$Message,
        [Parameter(Position = 1)]
        [string]$Title = 'STD',
        [ValidateSet('OK', 'OKCancel', 'YesNo', 'YesNoCancel')]
        [string]$Button = 'OK',
        [ValidateSet('Information', 'Warning', 'Error', 'Question', 'Success')]
        [string]$Image = 'Information',
        [System.Windows.Window]$Owner,
        [switch]$Danger
    )
    # Zmienne tej funkcji są prywatne: w trakcie ShowDialog() działają handlery i timery innych okien,
    # a ich bloki skryptu widzą zmienne z całego stosu wywołań. Bez tego np. $Title albo $Message
    # z tej funkcji zasłoniłyby zmienne o tej samej nazwie w kodzie ekranu.
    $private:stdMsgDialog = New-StdMessageWindow -Message $Message -Title $Title -Button $Button -Image $Image -Owner $Owner -Danger:$Danger
    $private:stdMsgDefault = [System.Windows.MessageBoxResult]::OK
    if ($Button -eq 'OKCancel' -or $Button -eq 'YesNoCancel') { $private:stdMsgDefault = [System.Windows.MessageBoxResult]::Cancel }
    elseif ($Button -eq 'YesNo') { $private:stdMsgDefault = [System.Windows.MessageBoxResult]::No }
    Remove-Variable -Name Message, Title, Button, Image, Owner, Danger -Scope Local -ErrorAction SilentlyContinue
    [void]$private:stdMsgDialog.ShowDialog()
    if ($private:stdMsgDialog.Tag -is [System.Windows.MessageBoxResult]) { return $private:stdMsgDialog.Tag }
    return $private:stdMsgDefault
}

# ---------------------------------------------------------------------------------------------
# Powiadomienie (toast)
# ---------------------------------------------------------------------------------------------

function Get-StdToastHost {
    # Panel "toastHost" okna. Gdy okno go nie ma, dokładamy go w prawym dolnym rogu: do głównej siatki
    # (Grid) okna na wszystkie wiersze i kolumny, a gdy zawartość nie jest siatką - owijamy ją w siatkę.
    param([Parameter(Mandatory = $true)][System.Windows.Window]$Window)
    $toastHost = $Window.FindName('toastHost')
    if ($toastHost -is [System.Windows.Controls.Panel]) { return $toastHost }
    $content = $Window.Content
    if ($content -isnot [System.Windows.UIElement]) { return $null }
    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $panel.VerticalAlignment = [System.Windows.VerticalAlignment]::Bottom
    $panel.Margin = New-Object System.Windows.Thickness(0, 0, 16, 16)
    [System.Windows.Controls.Panel]::SetZIndex($panel, 1000)
    if ($content -is [System.Windows.Controls.Grid]) {
        [System.Windows.Controls.Grid]::SetRowSpan($panel, [Math]::Max(1, $content.RowDefinitions.Count))
        [System.Windows.Controls.Grid]::SetColumnSpan($panel, [Math]::Max(1, $content.ColumnDefinitions.Count))
        [void]$content.Children.Add($panel)
    } else {
        $grid = New-Object System.Windows.Controls.Grid
        $Window.Content = $null
        [void]$grid.Children.Add($content)
        [void]$grid.Children.Add($panel)
        $Window.Content = $grid
    }
    try {
        if ($null -eq [System.Windows.NameScope]::GetNameScope($Window)) {
            [System.Windows.NameScope]::SetNameScope($Window, (New-Object System.Windows.NameScope))
        }
        $Window.RegisterName('toastHost', $panel)
    } catch { }
    return $panel
}

function Show-StdToast {
    <#
    .SYNOPSIS
    Krótkie powiadomienie w prawym dolnym rogu okna, znika samo (pojawia się w 220 ms, najwyżej
    4 naraz - najstarsze znika pierwsze). Bez -Window: okno STD, w którym użytkownik pracuje.
    Kolory przez SetResourceReference, więc powiadomienie przemalowuje się razem z motywem.
    Błąd (np. zamknięte okno) nie jest zgłaszany - powiadomienie jest tylko dodatkiem.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [System.Windows.Window]$Window,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Kind = 'Info',
        [ValidateRange(2, 60)]
        [int]$Seconds = 4
    )
    try {
        if ($null -eq $Window) { $Window = Get-StdActiveWindow }
        if ($null -eq $Window) { return }
        $toastHost = Get-StdToastHost -Window $Window
        if ($null -eq $toastHost) { return }
        $tone = @{ Info = 'Info'; Success = 'Ok'; Warning = 'Warn'; Error = 'Crit' }[$Kind]
        $glyph = @{ Info = 'E946'; Ok = 'E930'; Warn = 'E7BA'; Crit = 'EA39' }[$tone]

        $b = New-Object System.Windows.Controls.Border
        $b.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeToast')
        $b.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemePopupBorder')
        $b.BorderThickness = New-Object System.Windows.Thickness(1)
        $b.CornerRadius = New-Object System.Windows.CornerRadius(9)
        $b.Padding = New-Object System.Windows.Thickness(14, 10, 14, 10)
        $b.Margin = New-Object System.Windows.Thickness(0, 8, 0, 0)
        $b.MaxWidth = 420
        $b.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
        # Cień to efekt, nie pędzel - jego krycie dobieramy do motywu w chwili tworzenia powiadomienia.
        $fx = New-Object System.Windows.Media.Effects.DropShadowEffect
        $fx.BlurRadius = 18
        $fx.ShadowDepth = 3
        $fx.Color = [System.Windows.Media.Colors]::Black
        $fx.Opacity = 0.18
        if ($script:StdThemeState.Dark) { $fx.Opacity = 0.45 }
        $b.Effect = $fx

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $ic = New-StdGlyphBlock -Code $glyph -Size 15 -ColorKey "Theme$($tone)Fore"
        $ic.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
        $tx = New-Object System.Windows.Controls.TextBlock
        $tx.Text = $Message
        $tx.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $tx.MaxWidth = 360
        $tx.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $tx.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeText')
        [void]$sp.Children.Add($ic)
        [void]$sp.Children.Add($tx)
        $b.Child = $sp

        while ($toastHost.Children.Count -ge 4) { $toastHost.Children.RemoveAt(0) }
        [void]$toastHost.Children.Add($b)
        # Podgląd okien (STD_UI_PREVIEW=1) renderuje okno od razu - animacja zostawiłaby na zrzucie krycie 0.
        if ($env:STD_UI_PREVIEW -ne '1') {
            try {
                $fade = New-Object System.Windows.Media.Animation.DoubleAnimation -ArgumentList 0.0, 1.0, (New-Object System.Windows.Duration ([TimeSpan]::FromMilliseconds(220)))
                $b.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fade)
            } catch { }
        }
        # Stan timera w jego Tag: zmienne lokalne tej funkcji znikną, zanim timer zadziała.
        $toastTimer = New-Object System.Windows.Threading.DispatcherTimer
        $toastTimer.Interval = [TimeSpan]::FromSeconds($Seconds)
        $toastTimer.Tag = @{ Panel = $toastHost; Toast = $b }
        $toastTimer.Add_Tick({
            param($s, $e)
            try {
                $s.Stop()
                $t = $s.Tag
                if ($t.Panel.Children.Contains($t.Toast)) { $t.Panel.Children.Remove($t.Toast) }
            } catch { }
        })
        $toastTimer.Start()
    } catch { }
}

# ---------------------------------------------------------------------------------------------
# Ikony z czcionki
# ---------------------------------------------------------------------------------------------

function Get-StdGlyph {
    <#
    .SYNOPSIS
    Znak ikony z kodu szesnastkowego (np. 'E7B8') dla czcionki Segoe Fluent Icons / Segoe MDL2 Assets.
    Pusty kod -> pusty tekst. Zły kod: [STD-UI-GLYPH].
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Code
    )
    if ([string]::IsNullOrWhiteSpace($Code)) { return '' }
    $match = [regex]::Match($Code.Trim(), '^(?:0x|U\+)?([0-9A-Fa-f]{4,6})$')
    if (-not $match.Success) {
        throw "[STD-UI-GLYPH] Nieprawidłowy kod ikony '$Code' (oczekiwany kod szesnastkowy, np. E7B8)."
    }
    $value = [Convert]::ToInt32($match.Groups[1].Value, 16)
    if ($value -gt 0x10FFFF -or ($value -ge 0xD800 -and $value -le 0xDFFF)) {
        throw "[STD-UI-GLYPH] Nieprawidłowy kod ikony '$Code'."
    }
    return [char]::ConvertFromUtf32($value)
}

function New-StdGlyphBlock {
    <#
    .SYNOPSIS
    TextBlock z ikoną. -ColorKey = klucz koloru z palety (np. 'ThemeOkFore'), podpięty przez
    SetResourceReference, więc ikona zmienia kolor razem z motywem.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Code,
        [double]$Size = 13,
        [string]$ColorKey = ''
    )
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = Get-StdGlyph -Code $Code
    $t.FontFamily = New-Object System.Windows.Media.FontFamily $script:StdIconFontFamily
    $t.FontSize = $Size
    $t.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    if (-not [string]::IsNullOrEmpty($ColorKey)) { $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $ColorKey) }
    return $t
}

function New-StdIconContent {
    <#
    .SYNOPSIS
    Zawartość przycisku: ikona + tekst (albo sam tekst, gdy bez ikony).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [AllowEmptyString()]
        [string]$Text = '',
        [Parameter(Position = 1)]
        [AllowEmptyString()]
        [string]$Icon = '',
        [double]$IconSize = 13
    )
    if ([string]::IsNullOrWhiteSpace($Icon)) { return $Text }
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $g = New-StdGlyphBlock -Code $Icon -Size $IconSize
    if (-not [string]::IsNullOrEmpty($Text)) { $g.Margin = New-Object System.Windows.Thickness(0, 0, 8, 0) }
    [void]$sp.Children.Add($g)
    if (-not [string]::IsNullOrEmpty($Text)) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = $Text
        $tb.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$sp.Children.Add($tb)
    }
    return $sp
}

function Set-StdButtonContent {
    <#
    .SYNOPSIS
    Ustawia zawartość przycisku: ikona (kod szesnastkowy, np. 'E72C') i tekst.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [System.Windows.Controls.ContentControl]$Button,
        [Parameter(Position = 1)]
        [AllowEmptyString()]
        [string]$Text = '',
        [Parameter(Position = 2)]
        [AllowEmptyString()]
        [string]$Icon = '',
        [double]$IconSize = 12
    )
    if ($null -eq $Button) { return }
    $Button.Content = New-StdIconContent -Text $Text -Icon $Icon -IconSize $IconSize
}

function Get-StdButtonText {
    <#
    .SYNOPSIS
    Tekst przycisku bez ikon (Content bywa napisem albo panelem ikona + tekst z New-StdIconContent).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        $Button
    )
    if ($null -eq $Button) { return '' }
    $c = $Button.Content
    if ($c -is [string]) { return $c }
    if ($c -is [System.Windows.Controls.TextBlock]) {
        if ($c.FontFamily.Source -match 'MDL2|Fluent Icons') { return '' }
        return $c.Text
    }
    if ($c -is [System.Windows.Controls.Panel]) {
        $parts = foreach ($ch in $c.Children) {
            if ($ch -is [System.Windows.Controls.TextBlock] -and $ch.FontFamily.Source -notmatch 'MDL2|Fluent Icons') { $ch.Text }
        }
        return (@($parts) -join ' ')
    }
    return [string]$c
}

# ---------------------------------------------------------------------------------------------
# Ikona STD
# ---------------------------------------------------------------------------------------------

function New-StdAppIconBytes {
    <#
    .SYNOPSIS
    Ikona STD rysowana w kodzie jako bajty pliku .ico: kafelek z gradientem #4C7DF0 -> #E0913E,
    zaokrąglone rogi, biała paczka (znak E7B8 z Segoe Fluent Icons / Segoe MDL2 Assets).
    .DESCRIPTION
    Nie zależy od pliku obok skryptu. Kilka rozmiarów: małe jako 32-bitowy DIB (czyta go każdy
    Windows i System.Drawing.Icon), 256 px jako PNG - tak zapisuje go sam Windows.
    -DrawnMark: paczka narysowana kreską zamiast znaku z czcionki (zapas, gdy czcionki z ikonami
    nie ma, np. Windows Server bez Segoe MDL2 Assets).
    #>
    [CmdletBinding()]
    param(
        [int[]]$Sizes = @(16, 20, 24, 32, 40, 48, 64, 256),
        [switch]$DrawnMark
    )
    $glyph = $null
    $glyphIndex = 0
    if (-not $DrawnMark) {
        foreach ($familyName in 'Segoe Fluent Icons', 'Segoe MDL2 Assets') {
            $typeface = New-Object System.Windows.Media.Typeface -ArgumentList (New-Object System.Windows.Media.FontFamily $familyName), ([System.Windows.FontStyles]::Normal), ([System.Windows.FontWeights]::Normal), ([System.Windows.FontStretches]::Normal)
            $candidate = $null
            # Brak czcionki = WPF podstawia inną (np. Segoe UI), która nie ma znaku E7B8 - szukamy dalej.
            if ($typeface.TryGetGlyphTypeface([ref]$candidate) -and $candidate.CharacterToGlyphMap.ContainsKey(0xE7B8)) {
                $glyph = $candidate
                $glyphIndex = $candidate.CharacterToGlyphMap[0xE7B8]
                break
            }
        }
    }
    # Zapas bez czcionki: ta sama paczka, w jednostkach 0-1.
    $drawnGeometry = $null
    if ($null -eq $glyph) { $drawnGeometry = [System.Windows.Media.Geometry]::Parse('M0.5,0.08 L0.92,0.29 L0.92,0.71 L0.5,0.92 L0.08,0.71 L0.08,0.29 Z M0.08,0.29 L0.5,0.5 L0.92,0.29 M0.5,0.5 L0.5,0.92') }
    $from = [System.Windows.Media.Color]::FromRgb(0x4C, 0x7D, 0xF0)
    $to = [System.Windows.Media.Color]::FromRgb(0xE0, 0x91, 0x3E)
    $images = New-Object 'System.Collections.Generic.List[object]'
    foreach ($size in $Sizes) {
        $dv = New-Object System.Windows.Media.DrawingVisual
        $dc = $dv.RenderOpen()
        $tile = New-Object System.Windows.Media.LinearGradientBrush -ArgumentList $from, $to, (New-Object System.Windows.Point 0, 0), (New-Object System.Windows.Point 1, 1)
        $radius = $size * 0.22
        $dc.DrawRoundedRectangle($tile, $null, (New-Object System.Windows.Rect 0, 0, $size, $size), $radius, $radius)
        # W najmniejszych rozmiarach (pasek tytułu) znak większy, żeby był czytelny.
        $mark = $size * 0.6
        if ($size -le 24) { $mark = $size * 0.7 }
        if ($null -ne $glyph) {
            $geo = $glyph.GetGlyphOutline($glyphIndex, $mark, $mark)
            $bounds = $geo.Bounds
            $dc.PushTransform((New-Object System.Windows.Media.TranslateTransform -ArgumentList ((($size - $bounds.Width) / 2) - $bounds.X), ((($size - $bounds.Height) / 2) - $bounds.Y)))
            $dc.DrawGeometry([System.Windows.Media.Brushes]::White, $null, $geo)
            $dc.Pop()
        } else {
            $offset = ($size - $mark) / 2
            $dc.PushTransform((New-Object System.Windows.Media.MatrixTransform -ArgumentList $mark, 0, 0, $mark, $offset, $offset))
            $pen = New-Object System.Windows.Media.Pen -ArgumentList ([System.Windows.Media.Brushes]::White), 0.09
            $pen.LineJoin = [System.Windows.Media.PenLineJoin]::Round
            $dc.DrawGeometry($null, $pen, $drawnGeometry)
            $dc.Pop()
        }
        $dc.Close()
        $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap -ArgumentList $size, $size, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
        $rtb.Render($dv)
        $ms = New-Object System.IO.MemoryStream
        if ($size -ge 256) {
            $png = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
            $png.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
            $png.Save($ms)
        } else {
            # DIB w .ico: BITMAPINFOHEADER z podwójną wysokością (obraz + maska AND), wiersze od dołu,
            # kolory bez przemnożenia przez alfę (Bgra32, nie Pbgra32 z RenderTargetBitmap).
            $bgra = [System.Windows.Media.Imaging.FormatConvertedBitmap]::new($rtb, [System.Windows.Media.PixelFormats]::Bgra32, $null, 0)
            $stride = $size * 4
            $pixels = New-Object byte[] ($stride * $size)
            $bgra.CopyPixels($pixels, $stride, 0)
            $maskStride = [int]([Math]::Ceiling($size / 32.0) * 4)
            $mask = New-Object byte[] ($maskStride * $size)
            for ($y = 0; $y -lt $size; $y++) {
                $row = ($size - 1 - $y) * $maskStride
                for ($x = 0; $x -lt $size; $x++) {
                    # Bit 1 w masce AND = piksel w pełni przezroczysty (dla programów, które nie czytają alfy).
                    if ($pixels[($y * $stride) + ($x * 4) + 3] -eq 0) {
                        $i = $row + ($x -shr 3)
                        $mask[$i] = $mask[$i] -bor (0x80 -shr ($x % 8))
                    }
                }
            }
            $bw = New-Object System.IO.BinaryWriter $ms
            $bw.Write([int32]40); $bw.Write([int32]$size); $bw.Write([int32]($size * 2))
            $bw.Write([int16]1); $bw.Write([int16]32); $bw.Write([int32]0)
            $bw.Write([int32]($pixels.Length + $mask.Length))
            $bw.Write([int32]0); $bw.Write([int32]0); $bw.Write([int32]0); $bw.Write([int32]0)
            for ($y = $size - 1; $y -ge 0; $y--) { $bw.Write($pixels, $y * $stride, $stride) }
            $bw.Write($mask)
            $bw.Flush()
        }
        $images.Add([pscustomobject]@{ Size = $size; Data = $ms.ToArray() })
    }
    # Plik .ico: ICONDIR (6 B), po jednym ICONDIRENTRY (16 B) na rozmiar, potem dane obrazów.
    $out = New-Object System.IO.MemoryStream
    $w = New-Object System.IO.BinaryWriter $out
    $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$images.Count)
    $dataOffset = 6 + (16 * $images.Count)
    foreach ($img in $images) {
        $dim = $img.Size
        if ($img.Size -ge 256) { $dim = 0 }
        $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
        $w.Write([uint16]1); $w.Write([uint16]32)
        $w.Write([uint32]$img.Data.Length); $w.Write([uint32]$dataOffset)
        $dataOffset += $img.Data.Length
    }
    foreach ($img in $images) { $w.Write([byte[]]$img.Data) }
    $w.Flush()
    return , $out.ToArray()
}

function New-StdAppIcon {
    <#
    .SYNOPSIS
    Ikona STD (New-StdAppIconBytes) jako zamrożony BitmapFrame z całego pliku .ico: WPF sam wybiera
    z niego rozmiar na pasek tytułu (16-24 px) i pasek zadań (32-48 px).
    #>
    [CmdletBinding()]
    param(
        [switch]$DrawnMark
    )
    $bytes = New-StdAppIconBytes -DrawnMark:$DrawnMark
    $stream = New-Object System.IO.MemoryStream (, $bytes)
    $frame = [System.Windows.Media.Imaging.BitmapFrame]::Create($stream, [System.Windows.Media.Imaging.BitmapCreateOptions]::None, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    if ($frame.CanFreeze) { $frame.Freeze() }
    return $frame
}
