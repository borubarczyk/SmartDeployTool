# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Screens\Main.ps1 - okno główne STD 4 (SPEC 7.3 pkt 3 i 4):
# - nagłówek: logo, „Smart Tool for Deployment 4.0.0”, klient · lokalizacja · rewizja, pastylka źródła
#   konfiguracji (serwer / kopia lustrzana / cache / pendrive), konto technika, przełącznik motywu;
# - widok wyboru: Klient -> Lokalizacja -> Profil (ostatnio używane na górze list, prefs.json), konfiguracja
#   pobierana w tle (Get-StdRelease), karta „Komputer” (Get-StdComputerInfo w tle), zakładki „Konfiguracja”
#   (podsumowanie konfiguracji wynikowej, tylko do odczytu) i „Plan” (wynik „Sprawdź”: Krok | Stan | Szczegóły);
# - widok wdrożenia: stan i liczniki (OK / Pominięto / Błąd / Przerwano), lista kroków z ikoną i kolorem stanu,
#   dziennik na żywo (ostatnie 3000 wpisów, filtr poziomu), Pauza / Wznów / Przerwij (drugie potwierdzenie dla
#   „Przerwij natychmiast”), na końcu „Ponów nieudane” (-OnlySteps), „Kopiuj podsumowanie” i „Wróć do wyboru”.
#
# „Sprawdź” i „Wdróż” uruchamiają silnik (Start-StdEngine -Mode Plan / Run) w osobnym runspace. Okno co 200 ms
# odbiera zdarzenia (Receive-StdEngineEvent) w DispatcherTimer, więc nigdy się nie zawiesza. W trakcie
# wdrożenia zablokowane jest wszystko poza Pauza / Wznów / Przerwij i dziennikiem, a zamknięcie okna wymaga
# potwierdzenia (i przerywa wdrożenie).
#
# Stan okna to jeden [hashtable] ($ui) w Window.Tag. Stany widoku (Set-StdMainState): Select, Plan (widok
# wyboru), Running, Paused, Error, Finished (widok wdrożenia). -PreviewData / -PreviewState pokazują okno
# z przykładowymi danymi bez serwera i bez silnika (podgląd okien, testy).

$script:StdMainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart Tool for Deployment" Width="1100" Height="720" MinWidth="900" MinHeight="560" WindowStartupLocation="CenterScreen"
        Background="{DynamicResource ThemeBackground}" Foreground="{DynamicResource ThemeText}">
    <Window.Resources>
        <!-- Ikona i napis stanu kroku / pozycji planu. Kolory jak tony w Common.ps1 ($script:StdUiStateTable). -->
        <Style x:Key="StdStateGlyph" TargetType="TextBlock" BasedOn="{StaticResource Glyph}">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeSubtle}"/>
            <Style.Triggers>
                <DataTrigger Binding="{Binding State}" Value="Running"><Setter Property="Foreground" Value="{DynamicResource ThemeInfoFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Ok"><Setter Property="Foreground" Value="{DynamicResource ThemeOkFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Error"><Setter Property="Foreground" Value="{DynamicResource ThemeCritFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Cancelled"><Setter Property="Foreground" Value="{DynamicResource ThemeWarnFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Done"><Setter Property="Foreground" Value="{DynamicResource ThemeOkFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Todo"><Setter Property="Foreground" Value="{DynamicResource ThemeInfoFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Blocked"><Setter Property="Foreground" Value="{DynamicResource ThemeWarnFore}"/></DataTrigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="StdStateText" TargetType="TextBlock">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeSubtle}"/>
            <Style.Triggers>
                <DataTrigger Binding="{Binding State}" Value="Running">
                    <Setter Property="Foreground" Value="{DynamicResource ThemeInfoFore}"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Ok"><Setter Property="Foreground" Value="{DynamicResource ThemeOkFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Error">
                    <Setter Property="Foreground" Value="{DynamicResource ThemeCritFore}"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Cancelled"><Setter Property="Foreground" Value="{DynamicResource ThemeWarnFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Done"><Setter Property="Foreground" Value="{DynamicResource ThemeOkFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Todo"><Setter Property="Foreground" Value="{DynamicResource ThemeInfoFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Blocked">
                    <Setter Property="Foreground" Value="{DynamicResource ThemeWarnFore}"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
        <!-- Wiersz listy kroków: bieżący krok podświetlony. -->
        <Style x:Key="StdStepRow" TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
            <Style.Triggers>
                <DataTrigger Binding="{Binding State}" Value="Running"><Setter Property="Background" Value="{DynamicResource ThemeInfoBack}"/></DataTrigger>
            </Style.Triggers>
        </Style>
        <!-- Wpis dziennika: kolor wg poziomu. -->
        <Style x:Key="StdLevelText" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource ThemeGridText}"/>
            <Style.Triggers>
                <DataTrigger Binding="{Binding Level}" Value="Error"><Setter Property="Foreground" Value="{DynamicResource ThemeCritFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding Level}" Value="Warn"><Setter Property="Foreground" Value="{DynamicResource ThemeWarnFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding Level}" Value="Debug"><Setter Property="Foreground" Value="{DynamicResource ThemeSubtle}"/></DataTrigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="StdLevelGlyph" TargetType="TextBlock" BasedOn="{StaticResource Glyph}">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="{DynamicResource ThemeSubtle}"/>
            <Style.Triggers>
                <DataTrigger Binding="{Binding Level}" Value="Error"><Setter Property="Foreground" Value="{DynamicResource ThemeCritFore}"/></DataTrigger>
                <DataTrigger Binding="{Binding Level}" Value="Warn"><Setter Property="Foreground" Value="{DynamicResource ThemeWarnFore}"/></DataTrigger>
            </Style.Triggers>
        </Style>
        <!-- Kafelki liczników w widoku wdrożenia. -->
        <Style x:Key="StdCountTile" TargetType="Border">
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="10,4"/>
            <Setter Property="Margin" Value="6,2,0,2"/>
            <Setter Property="MinWidth" Value="64"/>
        </Style>
        <Style x:Key="StdCountValue" TargetType="TextBlock">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
        </Style>
        <Style x:Key="StdCountLabel" TargetType="TextBlock">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
            <Setter Property="TextWrapping" Value="NoWrap"/>
        </Style>

        <!-- Komórki tabel (ListView + GridView). Pola wierszy: Common.ps1 (New-StdUiStepItem, New-StdUiLogItem). -->
        <DataTemplate x:Key="StdGlyphCell">
            <TextBlock Text="{Binding Glyph, Mode=OneWay}" Style="{StaticResource StdStateGlyph}"/>
        </DataTemplate>
        <DataTemplate x:Key="StdTitleCell">
            <TextBlock Text="{Binding Title, Mode=OneWay}" TextWrapping="Wrap"/>
        </DataTemplate>
        <DataTemplate x:Key="StdStateCell">
            <TextBlock Text="{Binding StateLabel, Mode=OneWay}" Style="{StaticResource StdStateText}"/>
        </DataTemplate>
        <DataTemplate x:Key="StdDetailCell">
            <TextBlock Text="{Binding Detail, Mode=OneWay}" TextWrapping="Wrap" FontSize="12" Foreground="{DynamicResource ThemeMuted}"/>
        </DataTemplate>
        <DataTemplate x:Key="StdStepCell">
            <StackPanel>
                <TextBlock Text="{Binding Title, Mode=OneWay}" TextWrapping="Wrap"/>
                <TextBlock x:Name="detail" Text="{Binding Detail, Mode=OneWay}" TextWrapping="Wrap" FontSize="11.5" Margin="0,2,0,0" Foreground="{DynamicResource ThemeMuted}"/>
            </StackPanel>
            <DataTemplate.Triggers>
                <DataTrigger Binding="{Binding HasDetail}" Value="no">
                    <Setter TargetName="detail" Property="Visibility" Value="Collapsed"/>
                </DataTrigger>
            </DataTemplate.Triggers>
        </DataTemplate>
        <DataTemplate x:Key="StdLogTimeCell">
            <TextBlock Text="{Binding Time, Mode=OneWay}" FontSize="11.5" TextWrapping="NoWrap" Foreground="{DynamicResource ThemeSubtle}"/>
        </DataTemplate>
        <DataTemplate x:Key="StdLogGlyphCell">
            <TextBlock Text="{Binding Glyph, Mode=OneWay}" Style="{StaticResource StdLevelGlyph}"/>
        </DataTemplate>
        <DataTemplate x:Key="StdLogMessageCell">
            <TextBlock TextWrapping="Wrap" Style="{StaticResource StdLevelText}"><Run Text="{Binding Source, Mode=OneWay}" Foreground="{DynamicResource ThemeSubtle}"/><Run Text=" "/><Run Text="{Binding Message, Mode=OneWay}"/></TextBlock>
        </DataTemplate>

        <!-- Pozycja list wyboru (klient / lokalizacja / profil): nazwa i dopisek "ostatnio". -->
        <DataTemplate x:Key="StdPickItem">
            <DockPanel>
                <TextBlock DockPanel.Dock="Right" Text="{Binding Hint, Mode=OneWay}" FontSize="11.5" Margin="10,0,0,0" VerticalAlignment="Center" TextWrapping="NoWrap" Foreground="{DynamicResource ThemeSubtle}"/>
                <TextBlock Text="{Binding Name, Mode=OneWay}" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
            </DockPanel>
        </DataTemplate>
        <!-- „Chipy” aplikacji i zadań w podsumowaniu konfiguracji. -->
        <DataTemplate x:Key="StdAppChip">
            <Border Background="{DynamicResource ThemeChip}" CornerRadius="9" Padding="9,3" Margin="0,0,6,6" MaxWidth="300" ToolTip="{Binding Tip, Mode=OneWay}">
                <TextBlock Text="{Binding Name, Mode=OneWay}" FontSize="12" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource ThemeText}"/>
            </Border>
        </DataTemplate>
        <DataTemplate x:Key="StdTaskChip">
            <Border x:Name="bd" Background="{DynamicResource ThemeChip}" CornerRadius="9" Padding="8,3" Margin="0,0,6,6" MaxWidth="320" ToolTip="{Binding Tip, Mode=OneWay}">
                <DockPanel>
                    <TextBlock x:Name="ic" DockPanel.Dock="Left" Style="{StaticResource Glyph}" FontSize="11" Text="&#xE738;" Margin="0,0,6,0" Foreground="{DynamicResource ThemeSubtle}"/>
                    <TextBlock x:Name="tx" Text="{Binding Name, Mode=OneWay}" FontSize="12" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis" VerticalAlignment="Center" Foreground="{DynamicResource ThemeSubtle}"/>
                </DockPanel>
            </Border>
            <DataTemplate.Triggers>
                <DataTrigger Binding="{Binding Mark}" Value="on">
                    <Setter TargetName="bd" Property="Background" Value="{DynamicResource ThemeOkBack}"/>
                    <Setter TargetName="ic" Property="Text" Value="&#xE73E;"/>
                    <Setter TargetName="ic" Property="Foreground" Value="{DynamicResource ThemeOkFore}"/>
                    <Setter TargetName="tx" Property="Foreground" Value="{DynamicResource ThemeOkFore}"/>
                </DataTrigger>
            </DataTemplate.Triggers>
        </DataTemplate>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Nagłówek: logo, nazwa i wersja, klient · lokalizacja · rewizja; źródło, konto, motyw -->
        <Border Grid.Row="0" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeChromeBorder}" BorderThickness="0,0,0,1" Padding="16,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <DockPanel VerticalAlignment="Center">
                    <Border DockPanel.Dock="Left" Width="36" Height="36" CornerRadius="10">
                        <Border.Background>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                                <GradientStop Color="#4C7DF0" Offset="0"/>
                                <GradientStop Color="#E0913E" Offset="1"/>
                            </LinearGradientBrush>
                        </Border.Background>
                        <TextBlock Style="{StaticResource Glyph}" Text="&#xE7B8;" FontSize="17" Foreground="{DynamicResource ThemeOnAccent}" HorizontalAlignment="Center"/>
                    </Border>
                    <StackPanel Margin="11,0,0,0" VerticalAlignment="Center">
                        <DockPanel HorizontalAlignment="Left">
                            <TextBlock x:Name="txtVersion" DockPanel.Dock="Right" FontSize="11" Margin="7,3,0,0" Foreground="{DynamicResource ThemeFaint}"/>
                            <TextBlock Text="Smart Tool for Deployment" FontSize="15" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource ThemeTitle}"/>
                        </DockPanel>
                        <TextBlock x:Name="txtSubtitle" FontSize="11.5" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource ThemeSubtle}"/>
                    </StackPanel>
                </DockPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="16,0,0,0">
                    <Border x:Name="pillSource" Style="{StaticResource Chip}" Padding="10,4" MaxWidth="230" Margin="0,0,8,0">
                        <DockPanel>
                            <Ellipse x:Name="dotSource" DockPanel.Dock="Left" Width="8" Height="8" Margin="0,0,8,0" VerticalAlignment="Center" Fill="{DynamicResource ThemeSuccess}"/>
                            <TextBlock x:Name="txtSource" FontSize="12" TextTrimming="CharacterEllipsis" VerticalAlignment="Center" Foreground="{DynamicResource ThemeGhostText}"/>
                        </DockPanel>
                    </Border>
                    <Border x:Name="pillUser" Style="{StaticResource Chip}" Padding="10,4" MaxWidth="200" Margin="0,0,6,0">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Left" Style="{StaticResource Glyph}" Text="&#xE77B;" FontSize="12" Margin="0,0,7,0" Foreground="{DynamicResource ThemeGhostText}"/>
                            <TextBlock x:Name="txtUser" FontSize="12" TextTrimming="CharacterEllipsis" VerticalAlignment="Center" Foreground="{DynamicResource ThemeGhostText}"/>
                        </DockPanel>
                    </Border>
                    <Button x:Name="btnTheme" Style="{StaticResource GhostButton}" Padding="10,7"/>
                </StackPanel>
            </Grid>
        </Border>

        <Grid Grid.Row="1">
            <!-- ===== Widok wyboru ===== -->
            <Grid x:Name="viewSelect" Margin="20,16,20,12">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" MinWidth="250" MaxWidth="340"/>
                    <ColumnDefinition Width="14"/>
                    <ColumnDefinition Width="2*"/>
                </Grid.ColumnDefinitions>

                <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
                    <StackPanel>
                        <Border Style="{StaticResource Card}" Margin="0,0,0,12">
                            <StackPanel>
                                <TextBlock Text="Wybór" Style="{StaticResource CardTitle}"/>
                                <TextBlock Text="Klient" Style="{StaticResource Caption}"/>
                                <ComboBox x:Name="cmbClient" ItemTemplate="{StaticResource StdPickItem}" Margin="0,0,0,12"/>
                                <TextBlock Text="Lokalizacja" Style="{StaticResource Caption}"/>
                                <ComboBox x:Name="cmbSite" ItemTemplate="{StaticResource StdPickItem}" Margin="0,0,0,12" IsEnabled="False"/>
                                <TextBlock Text="Profil" Style="{StaticResource Caption}"/>
                                <ComboBox x:Name="cmbProfile" ItemTemplate="{StaticResource StdPickItem}" IsEnabled="False"/>
                                <TextBlock x:Name="txtPickHint" Style="{StaticResource HintText}" Margin="0,10,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource Card}">
                            <StackPanel>
                                <DockPanel Margin="0,0,0,10">
                                    <TextBlock DockPanel.Dock="Left" Style="{StaticResource Glyph}" Text="&#xE977;" FontSize="14" Margin="0,0,8,0" Foreground="{DynamicResource ThemeAccentText}"/>
                                    <TextBlock Text="Komputer" Style="{StaticResource CardTitle}" Margin="0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                                </DockPanel>
                                <TextBlock x:Name="txtComputerHint" Style="{StaticResource MutedText}" Text="Odczytywanie informacji o komputerze..."/>
                                <Grid x:Name="gridComputer" Visibility="Collapsed">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>

                <TabControl x:Name="tabSelect" Grid.Column="2">
                    <TabItem Header="Konfiguracja">
                        <Border Style="{StaticResource Card}" Padding="0">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
                                <StackPanel Margin="16,14,16,14">
                                    <!-- Brak wyboru, pobieranie, błąd albo pusta lista klientów -->
                                    <DockPanel x:Name="panelSummaryInfo">
                                        <Border x:Name="tileSummaryInfo" DockPanel.Dock="Left" Style="{StaticResource IconTile}">
                                            <TextBlock x:Name="txtSummaryInfoIcon" Style="{StaticResource Glyph}" FontSize="17" HorizontalAlignment="Center" Foreground="{DynamicResource ThemeInfoFore}"/>
                                        </Border>
                                        <StackPanel VerticalAlignment="Center">
                                            <TextBlock x:Name="txtSummaryInfoTitle" Style="{StaticResource DialogTitle}" FontSize="15"/>
                                            <TextBlock x:Name="txtSummaryInfoText" Style="{StaticResource MutedText}" Margin="0,4,0,0"/>
                                            <Button x:Name="btnRetryRelease" HorizontalAlignment="Left" Margin="0,12,0,0" Visibility="Collapsed"/>
                                        </StackPanel>
                                    </DockPanel>
                                    <StackPanel x:Name="panelSummary" Visibility="Collapsed">
                                        <TextBlock Text="KONFIGURACJA WYNIKOWA" Style="{StaticResource SectionCaps}"/>
                                        <Grid x:Name="gridSummary">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="Auto"/>
                                                <ColumnDefinition Width="*"/>
                                            </Grid.ColumnDefinitions>
                                        </Grid>
                                        <TextBlock x:Name="txtAppsCaption" Style="{StaticResource SectionCaps}" Margin="0,18,0,8"/>
                                        <ItemsControl x:Name="icApps" ItemTemplate="{StaticResource StdAppChip}" Focusable="False">
                                            <ItemsControl.ItemsPanel>
                                                <ItemsPanelTemplate><WrapPanel/></ItemsPanelTemplate>
                                            </ItemsControl.ItemsPanel>
                                        </ItemsControl>
                                        <TextBlock x:Name="txtAppsEmpty" Style="{StaticResource MutedText}" Margin="0,0,0,6" Visibility="Collapsed" Text="Profil nie ma aplikacji do instalacji."/>
                                        <TextBlock x:Name="txtTasksCaption" Style="{StaticResource SectionCaps}" Margin="0,12,0,8"/>
                                        <ItemsControl x:Name="icTasks" ItemTemplate="{StaticResource StdTaskChip}" Focusable="False">
                                            <ItemsControl.ItemsPanel>
                                                <ItemsPanelTemplate><WrapPanel/></ItemsPanelTemplate>
                                            </ItemsControl.ItemsPanel>
                                        </ItemsControl>
                                        <TextBlock Style="{StaticResource HintText}" Margin="0,10,0,0" Text="Podsumowanie jest tylko do odczytu. Konfigurację zmienia administrator i publikuje ją na serwerze STD."/>
                                    </StackPanel>
                                </StackPanel>
                            </ScrollViewer>
                        </Border>
                    </TabItem>
                    <TabItem Header="Plan">
                        <Border Style="{StaticResource Card}" Padding="0">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <DockPanel Margin="16,12,16,10">
                                    <CheckBox x:Name="chkShowDisabled" DockPanel.Dock="Right" Margin="12,0,0,0" VerticalAlignment="Center" ToolTip="Pokaż też kroki, których zadania są wyłączone w profilu">
                                        <TextBlock Text="Pokaż wyłączone" TextWrapping="NoWrap"/>
                                    </CheckBox>
                                    <TextBlock x:Name="txtPlanCounts" Style="{StaticResource MutedText}" VerticalAlignment="Center"/>
                                </DockPanel>
                                <ListView x:Name="lvPlan" Grid.Row="1" Margin="1,0,1,8" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                                    <ListView.View>
                                        <GridView>
                                            <GridViewColumn Width="36" CellTemplate="{StaticResource StdGlyphCell}"/>
                                            <GridViewColumn Header="Krok" Width="200" CellTemplate="{StaticResource StdTitleCell}"/>
                                            <GridViewColumn Header="Stan" Width="116" CellTemplate="{StaticResource StdStateCell}"/>
                                            <GridViewColumn Header="Szczegóły" Width="260" CellTemplate="{StaticResource StdDetailCell}"/>
                                        </GridView>
                                    </ListView.View>
                                </ListView>
                                <TextBlock x:Name="txtPlanEmpty" Grid.Row="1" Style="{StaticResource MutedText}" Margin="16,4,16,16" VerticalAlignment="Top"/>
                            </Grid>
                        </Border>
                    </TabItem>
                </TabControl>
            </Grid>

            <!-- ===== Widok wdrożenia ===== -->
            <Grid x:Name="viewRun" Margin="20,16,20,12" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Border Grid.Row="0" Style="{StaticResource Card}" Padding="14,12" Margin="0,0,0,12">
                    <DockPanel>
                        <WrapPanel DockPanel.Dock="Right" VerticalAlignment="Center" Margin="12,0,0,0">
                            <Border Style="{StaticResource StdCountTile}" Background="{DynamicResource ThemeOkBack}" ToolTip="Kroki zakończone poprawnie">
                                <StackPanel>
                                    <TextBlock x:Name="txtCountOk" Style="{StaticResource StdCountValue}" Text="0" Foreground="{DynamicResource ThemeOkFore}"/>
                                    <TextBlock Style="{StaticResource StdCountLabel}" Text="OK" Foreground="{DynamicResource ThemeOkFore}"/>
                                </StackPanel>
                            </Border>
                            <Border Style="{StaticResource StdCountTile}" Background="{DynamicResource ThemeChip}" ToolTip="Kroki pominięte (już zrobione albo pominięte po błędzie kroku krytycznego)">
                                <StackPanel>
                                    <TextBlock x:Name="txtCountSkipped" Style="{StaticResource StdCountValue}" Text="0" Foreground="{DynamicResource ThemeGhostText}"/>
                                    <TextBlock Style="{StaticResource StdCountLabel}" Text="Pominięto" Foreground="{DynamicResource ThemeGhostText}"/>
                                </StackPanel>
                            </Border>
                            <Border Style="{StaticResource StdCountTile}" Background="{DynamicResource ThemeCritBack}" ToolTip="Kroki zakończone błędem">
                                <StackPanel>
                                    <TextBlock x:Name="txtCountError" Style="{StaticResource StdCountValue}" Text="0" Foreground="{DynamicResource ThemeCritFore}"/>
                                    <TextBlock Style="{StaticResource StdCountLabel}" Text="Błąd" Foreground="{DynamicResource ThemeCritFore}"/>
                                </StackPanel>
                            </Border>
                            <Border Style="{StaticResource StdCountTile}" Background="{DynamicResource ThemeWarnBack}" ToolTip="Kroki przerwane">
                                <StackPanel>
                                    <TextBlock x:Name="txtCountCancelled" Style="{StaticResource StdCountValue}" Text="0" Foreground="{DynamicResource ThemeWarnFore}"/>
                                    <TextBlock Style="{StaticResource StdCountLabel}" Text="Przerwano" Foreground="{DynamicResource ThemeWarnFore}"/>
                                </StackPanel>
                            </Border>
                        </WrapPanel>
                        <Border x:Name="tileRun" DockPanel.Dock="Left" Style="{StaticResource IconTile}" VerticalAlignment="Center">
                            <TextBlock x:Name="txtRunIcon" Style="{StaticResource Glyph}" FontSize="17" HorizontalAlignment="Center" Foreground="{DynamicResource ThemeInfoFore}"/>
                        </Border>
                        <StackPanel VerticalAlignment="Center">
                            <TextBlock x:Name="txtRunTitle" Style="{StaticResource DialogTitle}" FontSize="15"/>
                            <TextBlock x:Name="txtRunDetail" Style="{StaticResource MutedText}" Margin="0,3,0,0"/>
                            <TextBlock x:Name="txtRunLog" Style="{StaticResource HintText}" Margin="0,3,0,0" Visibility="Collapsed"/>
                        </StackPanel>
                    </DockPanel>
                </Border>
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="5*" MinWidth="250"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="6*" MinWidth="260"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" Style="{StaticResource Card}" Padding="0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <DockPanel Margin="16,12,16,10">
                                <TextBlock x:Name="txtStepsCount" DockPanel.Dock="Right" Style="{StaticResource HintText}" TextWrapping="NoWrap" VerticalAlignment="Center" Margin="10,0,0,0"/>
                                <TextBlock Text="Kroki" Style="{StaticResource CardTitle}" Margin="0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                            </DockPanel>
                            <ListView x:Name="lvSteps" Grid.Row="1" Margin="1,0,1,8" ScrollViewer.HorizontalScrollBarVisibility="Disabled" ItemContainerStyle="{StaticResource StdStepRow}">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Width="36" CellTemplate="{StaticResource StdGlyphCell}"/>
                                        <GridViewColumn Header="Krok" Width="220" CellTemplate="{StaticResource StdStepCell}"/>
                                        <GridViewColumn Header="Stan" Width="100" CellTemplate="{StaticResource StdStateCell}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </Border>
                    <Border Grid.Column="2" Style="{StaticResource Card}" Padding="0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <DockPanel Margin="16,10,12,8">
                                <ComboBox x:Name="cmbLogLevel" DockPanel.Dock="Right" Width="150" Margin="10,0,0,0" ToolTip="Które wpisy pokazać w dzienniku"/>
                                <CheckBox x:Name="chkFollow" DockPanel.Dock="Right" IsChecked="True" VerticalAlignment="Center" Margin="10,0,0,0" ToolTip="Przewijaj dziennik do najnowszego wpisu">
                                    <TextBlock Text="Na bieżąco" TextWrapping="NoWrap"/>
                                </CheckBox>
                                <TextBlock Text="Dziennik" Style="{StaticResource CardTitle}" Margin="0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" ToolTip="Dwuklik na wpisie kopiuje go do schowka"/>
                            </DockPanel>
                            <ListView x:Name="lvLog" Grid.Row="1" Margin="1,0,1,8" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Czas" Width="70" CellTemplate="{StaticResource StdLogTimeCell}"/>
                                        <GridViewColumn Width="30" CellTemplate="{StaticResource StdLogGlyphCell}"/>
                                        <GridViewColumn Header="Wpis" Width="300" CellTemplate="{StaticResource StdLogMessageCell}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </Border>
                </Grid>
            </Grid>

            <StackPanel x:Name="toastHost" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,26,22"/>
        </Grid>

        <!-- Pasek stanu i przyciski: Sprawdź / Wdróż; w trakcie Pauza / Przerwij; na końcu Wróć / Kopiuj / Ponów -->
        <Border Grid.Row="2" Background="{DynamicResource ThemeHeader}" BorderBrush="{DynamicResource ThemeChromeBorder}" BorderThickness="0,1,0,0" Padding="16,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Ellipse x:Name="dotStatus" Width="8" Height="8" VerticalAlignment="Top" Margin="0,6,10,0" Fill="{DynamicResource ThemeSuccess}"/>
                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock x:Name="txtStatus" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Margin="0,0,0,6" Foreground="{DynamicResource ThemeText}"/>
                    <ProgressBar x:Name="barProgress" Minimum="0" Maximum="100" Value="0"/>
                </StackPanel>
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="16,0,0,0">
                    <Button x:Name="btnBack" Style="{StaticResource GhostButton}" Margin="0,0,8,0" Visibility="Collapsed"/>
                    <Button x:Name="btnCopy" MinHeight="34" Margin="0,0,8,0" Visibility="Collapsed"/>
                    <Button x:Name="btnRetry" Style="{StaticResource PrimaryButton}" MinHeight="34" Visibility="Collapsed"/>
                    <Button x:Name="btnPause" Style="{StaticResource WarningButton}" MinWidth="104" MinHeight="34" Margin="0,0,8,0" Visibility="Collapsed"/>
                    <Button x:Name="btnCancel" Style="{StaticResource DangerButton}" MinWidth="104" MinHeight="34" Visibility="Collapsed"/>
                    <Button x:Name="btnCheck" MinWidth="110" MinHeight="34" Margin="0,0,8,0"/>
                    <Button x:Name="btnDeploy" Style="{StaticResource PrimaryButton}" MinWidth="130" MinHeight="34"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

$script:StdMainControlNames = @(
    'txtVersion', 'txtSubtitle', 'pillSource', 'dotSource', 'txtSource', 'pillUser', 'txtUser', 'btnTheme',
    'viewSelect', 'cmbClient', 'cmbSite', 'cmbProfile', 'txtPickHint', 'txtComputerHint', 'gridComputer',
    'tabSelect', 'panelSummaryInfo', 'tileSummaryInfo', 'txtSummaryInfoIcon', 'txtSummaryInfoTitle', 'txtSummaryInfoText',
    'btnRetryRelease', 'panelSummary', 'gridSummary', 'txtAppsCaption', 'icApps', 'txtAppsEmpty', 'txtTasksCaption', 'icTasks',
    'chkShowDisabled', 'txtPlanCounts', 'lvPlan', 'txtPlanEmpty',
    'viewRun', 'txtCountOk', 'txtCountSkipped', 'txtCountError', 'txtCountCancelled', 'tileRun', 'txtRunIcon', 'txtRunTitle',
    'txtRunDetail', 'txtRunLog', 'txtStepsCount', 'lvSteps', 'cmbLogLevel', 'chkFollow', 'lvLog',
    'dotStatus', 'txtStatus', 'barProgress', 'btnBack', 'btnCopy', 'btnRetry', 'btnPause', 'btnCancel', 'btnCheck', 'btnDeploy'
)

$script:StdMainReleaseScript = @'
param($Arguments)
$fetch = @{ ClientId = [string]$Arguments.ClientId; Bootstrap = $Arguments.Bootstrap; UsbRoot = [string]$Arguments.UsbRoot }
if ($null -ne $Arguments.Session) { $fetch.Session = $Arguments.Session }
Get-StdRelease @fetch
'@

$script:StdMainComputerScript = @'
Get-StdComputerInfo
'@

# ---------------------------------------------------------------------------------------------
# Nagłówek
# ---------------------------------------------------------------------------------------------

function Update-StdMainThemeButton {
    # Ikona pokazuje motyw, na który przełącza przycisk: słońce = jasny, półksiężyc = ciemny.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $button = $Ui.C.btnTheme
    if (Test-StdThemeDark) {
        Set-StdUiButton -Button $button -Text '' -Icon 'E706' -ToolTip 'Przełącz na jasny motyw'
    } else {
        Set-StdUiButton -Button $button -Text '' -Icon 'E708' -ToolTip 'Przełącz na ciemny motyw'
    }
    $button.Content.Children[0].FontSize = 15
}

function Update-StdMainHeader {
    # Podtytuł (klient · lokalizacja · rewizja), pastylka źródła konfiguracji i konto technika.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $eff = $Ui.Effective
    $parts = New-Object 'System.Collections.Generic.List[string]'
    $clientItem = $c.cmbClient.SelectedItem
    if ($null -ne $eff) {
        $clientName = [string]$eff['ClientName']
        if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = [string]$eff['ClientId'] }
        $parts.Add($clientName)
        $siteName = Get-StdUiText $eff['Site'] 'name'
        if ($null -eq $siteName) { $siteName = Get-StdUiText $eff['Site'] 'id' }
        if ($null -ne $siteName) { $parts.Add($siteName) }
        $parts.Add([string]$eff['ProfileName'])
        if ($null -ne $eff['Revision']) { $parts.Add(('rev {0}' -f $eff['Revision'])) }
    } elseif ($null -ne $clientItem) {
        $parts.Add([string]$clientItem.Name)
    }
    if ($parts.Count -gt 0) {
        $c.txtSubtitle.Text = $parts -join '  ·  '
    } else {
        $c.txtSubtitle.Text = 'Przygotowanie komputera z Windows - wybierz klienta, lokalizację i profil'
    }
    $c.txtSubtitle.ToolTip = $c.txtSubtitle.Text

    $source = [string]$Ui.IndexSource
    $tip = 'Lista klientów: {0}' -f (Get-StdUiSourceLabel $Ui.IndexSource)
    if ($null -ne $Ui.IndexRevision) { $tip += (', rewizja {0}' -f $Ui.IndexRevision) }
    $label = 'Lista klientów: {0}' -f (Get-StdUiSourceLabel $source)
    if ($null -ne $Ui.Release) {
        $source = [string]$Ui.Release['Source']
        $label = 'Konfiguracja: {0}' -f (Get-StdUiSourceLabel $source)
        $tip += ('. Konfiguracja klienta: {0}, rewizja {1}' -f (Get-StdUiSourceLabel $source), $Ui.Release['Revision'])
        if (-not [string]::IsNullOrEmpty([string]$Ui.Release['Kid'])) { $tip += (', podpis {0}' -f $Ui.Release['Kid']) }
    }
    if ($source -eq 'cache' -or $source -eq 'usb') { $tip += '. Kopia offline - może nie być najnowsza.' }
    $c.txtSource.Text = $label
    $c.pillSource.ToolTip = $tip + '.'
    Set-StdUiColor $c.dotSource 'Fill' (Get-StdUiSourceTone $source)

    if ($null -ne $Ui.Session) {
        $c.txtUser.Text = [string]$Ui.Session.UserName
        $c.pillUser.ToolTip = 'Zalogowano do serwera STD jako {0}.' -f $Ui.Session.UserName
    } elseif (-not [string]::IsNullOrEmpty($Ui.UserName)) {
        $c.txtUser.Text = [string]$Ui.UserName
        $c.pillUser.ToolTip = 'Konto technika (podgląd okna).'
    } else {
        $c.txtUser.Text = 'offline'
        $c.pillUser.ToolTip = 'Praca offline - bez logowania do serwera STD. Instalatory z serwera mogą być niedostępne.'
    }
}

# ---------------------------------------------------------------------------------------------
# Wybór: klient, lokalizacja, profil
# ---------------------------------------------------------------------------------------------

function Set-StdMainComboItems {
    # Pozycje listy bez wywoływania handlera wyboru (flaga Suppress) i z zaznaczeniem -SelectId.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [Parameter(Mandatory = $true)]$Combo, [AllowEmptyCollection()][object[]]$Items, [string]$SelectId)
    $Ui.Suppress = $true
    try {
        $Combo.ItemsSource = $Items
        $Combo.SelectedItem = $null
        if (-not [string]::IsNullOrEmpty($SelectId)) {
            foreach ($item in $Items) {
                if ([string]::Equals([string]$item.Id, $SelectId, [System.StringComparison]::OrdinalIgnoreCase)) { $Combo.SelectedItem = $item; break }
            }
        }
    } finally {
        $Ui.Suppress = $false
    }
}

function Set-StdMainClients {
    # Lista klientów z indeksu, ostatnio używani na górze.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($client in @(Get-StdUiValue $Ui.Index 'clients')) {
        $id = Get-StdUiText $client 'id'
        if ($null -eq $id) { continue }
        $items.Add((New-StdUiPickItem -Id $id -Name (Get-StdUiText $client 'name')))
    }
    $sorted = Sort-StdUiByRecent -Items $items.ToArray() -Recent (Get-StdUiRecentIds -Kind clients -Prefs $Ui.Prefs)
    $Ui.ClientItems = $sorted
    Set-StdMainComboItems -Ui $Ui -Combo $Ui.C.cmbClient -Items $sorted
}

function Get-StdMainDefaultPick {
    # Pozycja zaznaczana od razu: ostatnio używana (pierwsza na liście po Sort-StdUiByRecent) albo jedyna.
    param([AllowEmptyCollection()][object[]]$Items)
    if ($null -eq $Items -or $Items.Count -eq 0) { return $null }
    if ([string]$Items[0].Hint -eq 'ostatnio' -or $Items.Count -eq 1) { return [string]$Items[0].Id }
    return $null
}

function Get-StdMainSite {
    # Lokalizacja z wczytanej konfiguracji klienta.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [string]$SiteId)
    if ($null -eq $Ui.Release) { return $null }
    foreach ($site in @(Get-StdUiValue $Ui.Release['Release'] 'sites')) {
        if ([string](Get-StdUiValue $site 'id') -eq $SiteId) { return $site }
    }
    return $null
}

function Select-StdMainClient {
    # Wybór klienta: konfiguracja z pamięci okna albo pobieranie w tle (Get-StdRelease).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [Parameter(Mandatory = $true)][string]$ClientId)
    if ($ClientId -eq $Ui.ClientId -and ($null -ne $Ui.Release -or $null -ne $Ui.ReleaseTask)) { return }
    $Ui.ClientId = $ClientId
    $Ui.Release = $null
    $Ui.Effective = $null
    Clear-StdMainPlan -Ui $Ui
    Set-StdMainComboItems -Ui $Ui -Combo $Ui.C.cmbSite -Items @()
    Set-StdMainComboItems -Ui $Ui -Combo $Ui.C.cmbProfile -Items @()
    if ($null -ne $Ui.ReleaseTask) {
        try { Stop-StdUiTask -Task $Ui.ReleaseTask } catch { }
        $Ui.ReleaseTask = $null
    }
    if ($Ui.ReleaseCache.ContainsKey($ClientId)) {
        Complete-StdMainRelease -Ui $Ui -ClientId $ClientId -Result $Ui.ReleaseCache[$ClientId]
        return
    }
    $name = $ClientId
    $item = $Ui.C.cmbClient.SelectedItem
    if ($null -ne $item) { $name = [string]$item.Name }
    if ($Ui.Preview) {
        Set-StdMainSummary -Ui $Ui -Mode Error -Title ('Brak przykładowej konfiguracji klienta {0}' -f $name) -Text 'Podgląd okna ma przykładową konfigurację tylko dla jednego klienta.'
        Update-StdMainSelection -Ui $Ui
        return
    }
    Set-StdMainSummary -Ui $Ui -Mode Busy -Title ('Pobieranie konfiguracji klienta {0}...' -f $name) -Text 'STD pobiera konfigurację i sprawdza jej podpis. To zwykle trwa kilka sekund.'
    $bootstrap = $Ui.Bootstrap
    if ($null -eq $Ui.Session -and $null -ne $bootstrap) {
        # Praca offline: bez serwera (bez sesji tylko kopie lustrzane, cache i pendrive).
        $mirrors = @(Get-StdUiValue $bootstrap 'mirrors') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
        $bootstrap = $null
        if (@($mirrors).Count -gt 0) { $bootstrap = @{ mirrors = @($mirrors) } }
    } elseif ($null -ne $bootstrap) {
        $bootstrap = ConvertTo-StdHashtable -InputObject $bootstrap
    }
    $handler = { param($Ui, $Result) Complete-StdMainReleaseTask -Ui $Ui -Result $Result }
    $errorHandler = { param($Ui, $ErrorRecord) Complete-StdMainReleaseTask -Ui $Ui -ErrorRecord $ErrorRecord }
    $Ui.ReleaseTaskClient = $ClientId
    $Ui.ReleaseTask = Start-StdUiTask -Name 'Pobieranie konfiguracji klienta' -Owner $Ui.Window -ScriptText $script:StdMainReleaseScript `
        -Arguments @{ ClientId = $ClientId; Bootstrap = $bootstrap; Session = $Ui.Session; UsbRoot = [string]$Ui.OriginRoot } `
        -OnComplete (New-StdUiCallback -Ui $Ui -Handler $handler) -OnError (New-StdUiCallback -Ui $Ui -Handler $errorHandler)
    Update-StdMainSelection -Ui $Ui
}

function Complete-StdMainReleaseTask {
    # Wynik zadania Get-StdRelease (sukces albo błąd). Wynik dla innego klienta niż wybrany jest pomijany.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowNull()]$Result, [AllowNull()]$ErrorRecord)
    $clientId = [string]$Ui.ReleaseTaskClient
    $Ui.ReleaseTask = $null
    if ($clientId -ne $Ui.ClientId) { Update-StdMainSelection -Ui $Ui; return }
    if ($null -ne $ErrorRecord) {
        $name = $clientId
        $item = $Ui.C.cmbClient.SelectedItem
        if ($null -ne $item) { $name = [string]$item.Name }
        $text = Get-StdUiErrorMessage $ErrorRecord -MaxLength 2500
        Set-StdMainSummary -Ui $Ui -Mode Error -Title ('Nie udało się pobrać konfiguracji klienta {0}' -f $name) -Text ($text -replace '^\s*\[[A-Z0-9-]+\]\s*', '')
        Show-StdToast -Window $Ui.Window -Message 'Nie udało się pobrać konfiguracji klienta.' -Kind Error
        Update-StdMainSelection -Ui $Ui
        return
    }
    $Ui.ReleaseCache[$clientId] = $Result
    Complete-StdMainRelease -Ui $Ui -ClientId $clientId -Result $Result
}

function Complete-StdMainRelease {
    # Konfiguracja klienta wczytana: lista lokalizacji (ostatnio używane na górze, zaznaczona ostatnia albo jedyna).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [Parameter(Mandatory = $true)][string]$ClientId, [Parameter(Mandatory = $true)]$Result)
    $Ui.Release = $Result
    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($site in @(Get-StdUiValue $Result['Release'] 'sites')) {
        $id = Get-StdUiText $site 'id'
        if ($null -ne $id) { $items.Add((New-StdUiPickItem -Id $id -Name (Get-StdUiText $site 'name'))) }
    }
    $recent = Get-StdUiRecentIds -Kind sites -ClientId $ClientId -Prefs $Ui.Prefs
    $sorted = Sort-StdUiByRecent -Items $items.ToArray() -Recent $recent
    $select = $null
    if ($null -ne $Ui.PendingSite) { $select = $Ui.PendingSite; $Ui.PendingSite = $null }
    else { $select = Get-StdMainDefaultPick -Items $sorted }
    Set-StdMainComboItems -Ui $Ui -Combo $Ui.C.cmbSite -Items $sorted -SelectId $select
    if ($sorted.Count -eq 0) {
        Set-StdMainSummary -Ui $Ui -Mode Error -Title 'Konfiguracja nie ma lokalizacji' -Text 'Poproś administratora o dodanie lokalizacji w edytorze konfiguracji.'
    }
    Set-StdMainProfiles -Ui $Ui
}

function Set-StdMainProfiles {
    # Lista profili wybranej lokalizacji (ostatnio używane na górze).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $Ui.Effective = $null
    Clear-StdMainPlan -Ui $Ui
    $siteItem = $Ui.C.cmbSite.SelectedItem
    if ($null -eq $siteItem) {
        Set-StdMainComboItems -Ui $Ui -Combo $Ui.C.cmbProfile -Items @()
        if ($null -ne $Ui.Release) { Set-StdMainSummary -Ui $Ui -Mode Empty -Title 'Wybierz lokalizację i profil' -Text 'Po wyborze pojawi się tu podsumowanie konfiguracji: aplikacje, zadania, domena i nazwa komputera.' }
        Update-StdMainSelection -Ui $Ui
        return
    }
    $site = Get-StdMainSite -Ui $Ui -SiteId ([string]$siteItem.Id)
    $items = New-Object 'System.Collections.Generic.List[object]'
    $profiles = Get-StdUiValue $site 'profiles'
    if ($profiles -is [System.Collections.IDictionary]) {
        foreach ($key in @($profiles.Keys)) { $items.Add((New-StdUiPickItem -Id ([string]$key) -Name ([string]$key))) }
    }
    $recent = Get-StdUiRecentIds -Kind profiles -ClientId $Ui.ClientId -SiteId ([string]$siteItem.Id) -Prefs $Ui.Prefs
    $sorted = Sort-StdUiByRecent -Items $items.ToArray() -Recent $recent
    $select = $null
    if ($null -ne $Ui.PendingProfile) { $select = $Ui.PendingProfile; $Ui.PendingProfile = $null }
    else { $select = Get-StdMainDefaultPick -Items $sorted }
    Set-StdMainComboItems -Ui $Ui -Combo $Ui.C.cmbProfile -Items $sorted -SelectId $select
    Update-StdMainEffective -Ui $Ui
}

function Update-StdMainEffective {
    # Konfiguracja wynikowa (Resolve-StdEffectiveConfig) dla wybranej lokalizacji i profilu + podsumowanie.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $Ui.Effective = $null
    Clear-StdMainPlan -Ui $Ui
    $siteItem = $Ui.C.cmbSite.SelectedItem
    $profileItem = $Ui.C.cmbProfile.SelectedItem
    if ($null -eq $Ui.Release -or $null -eq $siteItem -or $null -eq $profileItem) {
        if ($null -ne $Ui.Release -and $null -ne $siteItem) {
            Set-StdMainSummary -Ui $Ui -Mode Empty -Title 'Wybierz profil' -Text 'Profil decyduje, które aplikacje i zadania zostaną wykonane.'
        }
        Update-StdMainSelection -Ui $Ui
        return
    }
    try {
        $Ui.Effective = Resolve-StdEffectiveConfig -Release $Ui.Release['Release'] -SiteId ([string]$siteItem.Id) -ProfileName ([string]$profileItem.Id)
    } catch {
        Set-StdMainSummary -Ui $Ui -Mode Error -Title 'Nie udało się ustalić konfiguracji wynikowej' -Text ((Get-StdUiErrorMessage $_) -replace '^\s*\[[A-Z0-9-]+\]\s*', '')
        Update-StdMainSelection -Ui $Ui
        return
    }
    Show-StdMainSummary -Ui $Ui
    Update-StdMainSelection -Ui $Ui
}

function Update-StdMainSelection {
    # Po każdej zmianie wyboru: nagłówek, podpowiedź pod listami, dostępność list i przycisków, pasek stanu.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $hint = ''
    if (@($Ui.ClientItems).Count -eq 0) { $hint = 'Lista klientów jest pusta.' }
    elseif ($null -eq $c.cmbClient.SelectedItem) { $hint = 'Zacznij od wyboru klienta.' }
    elseif ($null -ne $Ui.ReleaseTask) { $hint = 'Pobieranie konfiguracji...' }
    elseif ($null -ne $Ui.Effective) { $hint = 'Kliknij „Sprawdź”, aby zobaczyć plan, albo „Wdróż”, aby zacząć.' }
    $c.txtPickHint.Text = $hint
    Set-StdUiVisible $c.txtPickHint ($hint.Length -gt 0)
    Update-StdMainHeader -Ui $Ui
    if ($Ui.State -eq 'Select' -or $Ui.State -eq 'Plan') { Update-StdMainSelectControls -Ui $Ui; Update-StdMainStatus -Ui $Ui }
}

function Update-StdMainSelectControls {
    # Dostępność list i przycisków widoku wyboru (blokada w trakcie pobierania i „Sprawdź”).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $engineBusy = Test-StdMainEngineBusy -Ui $Ui
    $releaseBusy = ($null -ne $Ui.ReleaseTask)
    # Uwaga: @($null).Count to 1 - liczymy pozycje tylko wtedy, gdy lista ma źródło.
    $siteCount = 0
    if ($null -ne $c.cmbSite.ItemsSource) { $siteCount = @($c.cmbSite.ItemsSource).Count }
    $profileCount = 0
    if ($null -ne $c.cmbProfile.ItemsSource) { $profileCount = @($c.cmbProfile.ItemsSource).Count }
    $c.cmbClient.IsEnabled = -not $engineBusy
    $c.cmbSite.IsEnabled = (-not $engineBusy) -and (-not $releaseBusy) -and ($siteCount -gt 0)
    $c.cmbProfile.IsEnabled = (-not $engineBusy) -and (-not $releaseBusy) -and ($profileCount -gt 0)
    $ready = ($null -ne $Ui.Effective) -and (-not $engineBusy) -and (-not $releaseBusy)
    $c.btnCheck.IsEnabled = $ready
    $c.btnDeploy.IsEnabled = $ready
}

# ---------------------------------------------------------------------------------------------
# Podsumowanie konfiguracji i karta „Komputer”
# ---------------------------------------------------------------------------------------------

function Set-StdMainSummary {
    # Zakładka „Konfiguracja”: -Mode Empty / Busy / Error / Ready.
    param(
        [Parameter(Mandatory = $true)][hashtable]$Ui,
        [Parameter(Mandatory = $true)][ValidateSet('Empty', 'Busy', 'Error', 'Ready')][string]$Mode,
        [string]$Title,
        [AllowEmptyString()][string]$Text
    )
    $c = $Ui.C
    $Ui.SummaryMode = $Mode
    Set-StdUiVisible $c.panelSummary ($Mode -eq 'Ready')
    Set-StdUiVisible $c.panelSummaryInfo ($Mode -ne 'Ready')
    Set-StdUiVisible $c.btnRetryRelease ($Mode -eq 'Error' -and -not [string]::IsNullOrEmpty($Ui.ClientId) -and -not $Ui.Preview)
    if ($Mode -eq 'Ready') { return }
    $spec = @{
        Empty = @{ Tone = 'Info'; Glyph = 'E71D' }
        Busy  = @{ Tone = 'Info'; Glyph = 'E72C' }
        Error = @{ Tone = 'Crit'; Glyph = 'EA39' }
    }[$Mode]
    Set-StdUiColor $c.tileSummaryInfo 'Background' ('Theme{0}Back' -f $spec.Tone)
    Set-StdUiColor $c.txtSummaryInfoIcon 'Foreground' ('Theme{0}Fore' -f $spec.Tone)
    $c.txtSummaryInfoIcon.Text = Get-StdGlyph -Code $spec.Glyph
    $c.txtSummaryInfoTitle.Text = $Title
    $c.txtSummaryInfoText.Text = $Text
    Set-StdUiVisible $c.txtSummaryInfoText (-not [string]::IsNullOrWhiteSpace($Text))
}

function Show-StdMainSummary {
    # Podsumowanie konfiguracji wynikowej: wiersze, aplikacje i zadania (tylko do odczytu).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $eff = $Ui.Effective
    Set-StdUiInfoRows -Grid $c.gridSummary -Rows (Get-StdUiSummaryRows -Effective $eff)
    $apps = Get-StdUiAppChips -Effective $eff
    $c.icApps.ItemsSource = $apps
    $c.txtAppsCaption.Text = 'APLIKACJE ({0})' -f $apps.Count
    Set-StdUiVisible $c.txtAppsEmpty ($apps.Count -eq 0)
    $tasks = Get-StdUiTaskChips -Effective $eff
    $c.icTasks.ItemsSource = $tasks
    $enabled = @($tasks | Where-Object { $_.Mark -eq 'on' }).Count
    $c.txtTasksCaption.Text = 'ZADANIA (WŁĄCZONE: {0} Z {1})' -f $enabled, $tasks.Count
    Set-StdMainSummary -Ui $Ui -Mode Ready
}

function Set-StdMainComputer {
    # Karta „Komputer” z wyniku Get-StdComputerInfo (kliknięcie wartości kopiuje ją do schowka).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [AllowNull()]$Info, [string]$ErrorText)
    $c = $Ui.C
    $Ui.ComputerInfo = $Info
    if ($null -eq $Info) {
        $text = 'Nie udało się odczytać informacji o komputerze.'
        if (-not [string]::IsNullOrWhiteSpace($ErrorText)) { $text = '{0} {1}' -f $text, $ErrorText }
        $c.txtComputerHint.Text = $text
        Set-StdUiVisible $c.txtComputerHint $true
        Set-StdUiVisible $c.gridComputer $false
        return
    }
    Set-StdUiInfoRows -Grid $c.gridComputer -Rows (Get-StdUiComputerRows -Info $Info) -Copyable
    Set-StdUiVisible $c.txtComputerHint $false
    Set-StdUiVisible $c.gridComputer $true
}

# ---------------------------------------------------------------------------------------------
# Plan („Sprawdź”)
# ---------------------------------------------------------------------------------------------

function Clear-StdMainPlan {
    # Plan przestaje pasować po zmianie wyboru - czyścimy tabelę i wracamy do zakładki „Konfiguracja”.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($null -eq $Ui.Plan -and $Ui.PlanItems.Count -eq 0 -and [string]::IsNullOrEmpty($Ui.PlanError)) { return }
    $Ui.Plan = $null
    $Ui.PlanError = $null
    $Ui.PlanItems.Clear()
    Update-StdMainPlanView -Ui $Ui
    if ($Ui.C.tabSelect.SelectedIndex -ne 0 -and -not (Test-StdMainEngineBusy -Ui $Ui)) { $Ui.C.tabSelect.SelectedIndex = 0 }
}

function Update-StdMainPlanView {
    # Tabela planu (Krok | Stan | Szczegóły), liczniki i podpowiedź; wyłączone kroki tylko z „Pokaż wyłączone”.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $Ui.PlanItems.Clear()
    $counts = @{ Todo = 0; Done = 0; Blocked = 0; Disabled = 0 }
    $showDisabled = ($c.chkShowDisabled.IsChecked -eq $true)
    foreach ($entry in @($Ui.Plan)) {
        if ($null -eq $entry) { continue }
        $state = [string](Get-StdUiValue $entry 'State')
        if ($counts.ContainsKey($state)) { $counts[$state] = $counts[$state] + 1 }
        if ($state -eq 'Disabled' -and -not $showDisabled) { continue }
        $Ui.PlanItems.Add((New-StdUiStepItem -Id ([string](Get-StdUiValue $entry 'Id')) -Title ([string](Get-StdUiValue $entry 'Title')) -State $state -Detail ([string](Get-StdUiValue $entry 'Detail')) -Phase (Get-StdUiValue $entry 'Phase')))
    }
    $empty = ''
    if (-not [string]::IsNullOrEmpty($Ui.PlanError)) {
        $empty = $Ui.PlanError
    } elseif (Test-StdMainEngineBusy -Ui $Ui) {
        if ($Ui.PlanItems.Count -eq 0) { $empty = 'Sprawdzanie, co zostanie zrobione...' }
    } elseif ($null -eq $Ui.Plan) {
        $empty = 'Kliknij „Sprawdź”, aby zobaczyć, co zostanie zrobione, co już jest zrobione i gdzie jest problem. Sprawdzenie niczego nie zmienia na komputerze.'
    } elseif ($Ui.PlanItems.Count -eq 0) {
        $empty = 'Brak kroków do pokazania. Zaznacz „Pokaż wyłączone”, aby zobaczyć wszystkie kroki.'
    }
    $c.txtPlanEmpty.Text = $empty
    Set-StdUiVisible $c.txtPlanEmpty ($empty.Length -gt 0)
    if ($null -ne $Ui.Plan) {
        $c.txtPlanCounts.Text = 'Do zrobienia: {0}  ·  Już zrobione: {1}  ·  Problem: {2}  ·  Wyłączone: {3}' -f $counts.Todo, $counts.Done, $counts.Blocked, $counts.Disabled
    } else {
        $c.txtPlanCounts.Text = 'Plan pokazuje wynik „Sprawdź”.'
    }
    $Ui.PlanCounts = $counts
}

function Start-StdMainPlan {
    # „Sprawdź”: Start-StdEngine -Mode Plan. -ThenDeploy: po planie potwierdzenie wdrożenia („Wdróż” bez planu).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [switch]$ThenDeploy)
    if ($null -eq $Ui.Effective -or (Test-StdMainEngineBusy -Ui $Ui)) { return }
    $Ui.C.tabSelect.SelectedIndex = 1
    if ($Ui.Preview) {
        $Ui.Plan = @(Get-StdUiValue $Ui.PreviewData 'Plan')
        $Ui.PlanError = $null
        Set-StdMainState -Ui $Ui -State Plan
        if ($ThenDeploy) { Confirm-StdMainDeploy -Ui $Ui }
        return
    }
    Save-StdMainRecent -Ui $Ui
    $Ui.Plan = $null
    $Ui.PlanError = $null
    $Ui.DeployAfterPlan = [bool]$ThenDeploy
    $config = New-StdMainEngineConfig -Ui $Ui
    $engine = Start-StdEngine -Mode Plan -Config $config
    Set-StdMainEngine -Ui $Ui -Engine $engine -Mode Plan
    Update-StdMainPlanView -Ui $Ui
    Set-StdMainState -Ui $Ui -State Select
}

function New-StdMainEngineConfig {
    # Konfiguracja dla silnika: wynik Resolve-StdEffectiveConfig + Answers (bez haseł) + plik startowy (adres serwera).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $eff = $Ui.Effective
    $config = @{}
    foreach ($key in @($eff.Keys)) { $config[[string]$key] = $eff[$key] }
    # Odpowiedzi technika (etap 3: ekran pytań). Hasła nigdy tutaj nie trafiają.
    $config['Answers'] = @{ computerName = $null; apps = @($eff['Apps']); tasks = $eff['Tasks'] }
    if ($null -ne $Ui.Bootstrap) {
        $config['Bootstrap'] = ConvertTo-StdHashtable -InputObject $Ui.Bootstrap
        $server = Get-StdUiText $Ui.Bootstrap 'server'
        if ($null -ne $server) { $config['Server'] = $server }
    }
    if (-not [string]::IsNullOrWhiteSpace($Ui.OriginRoot)) { $config['OriginRoot'] = [string]$Ui.OriginRoot }
    return $config
}

function Save-StdMainRecent {
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $siteItem = $Ui.C.cmbSite.SelectedItem
    $profileItem = $Ui.C.cmbProfile.SelectedItem
    if ([string]::IsNullOrEmpty($Ui.ClientId) -or $null -eq $siteItem -or $null -eq $profileItem) { return }
    Save-StdUiRecent -ClientId $Ui.ClientId -SiteId ([string]$siteItem.Id) -ProfileName ([string]$profileItem.Id)
    $Ui.Prefs = Get-StdUiPrefs
}

# ---------------------------------------------------------------------------------------------
# Wdrożenie („Wdróż”, pauza, przerwanie, ponowienie)
# ---------------------------------------------------------------------------------------------

function Confirm-StdMainDeploy {
    # „Wdróż”: bez planu najpierw „Sprawdź” (lista kroków), potem potwierdzenie z listą kroków.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($null -eq $Ui.Effective -or (Test-StdMainEngineBusy -Ui $Ui)) { return }
    if ($null -eq $Ui.Plan) {
        Start-StdMainPlan -Ui $Ui -ThenDeploy
        return
    }
    $enabled = @($Ui.Plan | Where-Object { $null -ne $_ -and (Get-StdUiValue $_ 'Enabled') -eq $true })
    if ($enabled.Count -eq 0) {
        [void](Show-StdMessage -Owner $Ui.Window -Title 'Brak kroków do wykonania' -Image Information -Message 'Dla tego profilu nie ma żadnego włączonego kroku do wykonania. Sprawdź profil albo wybierz inny.')
        return
    }
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $blocked = 0
    $max = 16
    foreach ($entry in $enabled) {
        $state = [string](Get-StdUiValue $entry 'State')
        if ($state -eq 'Blocked') { $blocked++ }
        if ($lines.Count -ge $max) { continue }
        $line = '•  {0}' -f [string](Get-StdUiValue $entry 'Title')
        if ($state -eq 'Done') { $line += ' (już zrobione - zostanie pominięty)' }
        elseif ($state -eq 'Blocked') { $line += ' (problem - zakończy się błędem)' }
        $lines.Add($line)
    }
    if ($enabled.Count -gt $max) { $lines.Add(('... i jeszcze {0}' -f ($enabled.Count - $max))) }
    $message = 'Kroki do wykonania ({0}):{1}{1}{2}{1}{1}W trakcie wdrożenia STD o nic nie pyta. Pauza i Przerwij są dostępne w każdej chwili.' -f $enabled.Count, "`n", ($lines -join "`n")
    if ($blocked -gt 0) { $message += ('{0}{0}Uwaga: kroki z problemem ({1}) zakończą się błędem - szczegóły są w zakładce „Plan”.' -f "`n", $blocked) }
    $answer = Show-StdMessage -Owner $Ui.Window -Title 'Rozpocząć wdrożenie?' -Image Question -Button YesNo -Message $message
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
    if ($Ui.Closed -or $null -eq $Ui.Effective) { return }
    Start-StdMainRun -Ui $Ui
}

function Start-StdMainRun {
    # Start-StdEngine -Mode Run. -OnlySteps: „Ponów nieudane” w tym samym wdrożeniu (ten sam identyfikator).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [string[]]$OnlySteps)
    if (Test-StdMainEngineBusy -Ui $Ui) { return }
    $retry = ($null -ne $OnlySteps -and $OnlySteps.Count -gt 0)
    if ($Ui.Preview) {
        Set-StdMainPreview -Ui $Ui -State Running
        return
    }
    $config = New-StdMainEngineConfig -Ui $Ui
    $params = @{ Mode = 'Run'; Config = $config }
    if ($retry) {
        $params.OnlySteps = $OnlySteps
        $params.DeploymentId = $Ui.DeploymentId
        Add-StdMainLog -Ui $Ui -Items @(New-StdUiLogItem -Level Info -Source 'STD' -Message ('Ponowienie nieudanych kroków: {0}.' -f $OnlySteps.Count))
    } else {
        Save-StdMainRecent -Ui $Ui
        $Ui.StepItems.Clear()
        $Ui.StepIndex = @{}
        $Ui.RunStartedAt = Get-Date
        if ($Ui.LogAll.Count -gt 0) { Add-StdMainLog -Ui $Ui -Items @(New-StdUiLogItem -Level Info -Source 'STD' -Message 'Start wdrożenia.') }
    }
    $engine = Start-StdEngine @params
    Set-StdMainEngine -Ui $Ui -Engine $engine -Mode Run
    $Ui.DeploymentId = [string]$engine.Id
    $Ui.LogPath = Join-Path (Join-Path (Get-StdPath -Kind Logs) $Ui.DeploymentId) 'deploy.log'
    $Ui.RunStatus = 'Running'
    $Ui.RunError = $null
    $Ui.RebootRequired = $false
    $Ui.RunFinishedAt = $null
    $Ui.Progress = 0
    $Ui.ProgressText = 'Start wdrożenia...'
    $Ui.CurrentStep = $null
    Write-StdLog -Source 'UI' -Message ('Start wdrożenia {0} (klient {1}).' -f $Ui.DeploymentId, $Ui.ClientId)
    Set-StdMainState -Ui $Ui -State Running
}

function Request-StdMainPause {
    # Pauza (po bieżącym kroku) albo wznowienie.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.Preview) {
        if ($Ui.State -eq 'Paused') { Set-StdMainPreview -Ui $Ui -State Running } else { Set-StdMainPreview -Ui $Ui -State Paused }
        return
    }
    if (-not (Test-StdMainEngineBusy -Ui $Ui) -or $Ui.EngineMode -ne 'Run') { return }
    if ($Ui.State -eq 'Paused' -or $Ui.PauseRequested) {
        Resume-StdEngine -Engine $Ui.Engine
        $Ui.PauseRequested = $false
        Write-StdLog -Source 'UI' -Message 'Technik wznowił wdrożenie.'
        if ($Ui.State -eq 'Paused') { $Ui.EnginePaused = $false; Set-StdMainState -Ui $Ui -State Running } else { Set-StdMainState -Ui $Ui -State Running }
        return
    }
    Suspend-StdEngine -Engine $Ui.Engine
    $Ui.PauseRequested = $true
    Write-StdLog -Source 'UI' -Message 'Technik wstrzymał wdrożenie (pauza po bieżącym kroku).'
    Set-StdMainState -Ui $Ui -State Running
}

function Request-StdMainCancel {
    # Przerwij (po bieżącym kroku), a po pierwszym żądaniu: Przerwij natychmiast (drugie potwierdzenie).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.Preview) { Set-StdMainPreview -Ui $Ui -State Error; return }
    if (-not (Test-StdMainEngineBusy -Ui $Ui) -or $Ui.EngineMode -ne 'Run') { return }
    if (-not $Ui.CancelRequested) {
        $answer = Show-StdMessage -Owner $Ui.Window -Title 'Przerwać wdrożenie?' -Image Warning -Button YesNo -Danger `
            -Message ('Wdrożenie zostanie przerwane po zakończeniu bieżącego kroku. Pozostałe kroki dostaną stan „Przerwano” i będzie można je uruchomić przyciskiem „Ponów nieudane”.{0}{0}Przerwać wdrożenie?' -f "`n")
        if ($answer -ne [System.Windows.MessageBoxResult]::Yes -or -not (Test-StdMainEngineBusy -Ui $Ui)) { return }
        Stop-StdEngine -Engine $Ui.Engine
        $Ui.CancelRequested = $true
        Write-StdLog -Level Warn -Source 'UI' -Message 'Technik przerwał wdrożenie (po bieżącym kroku).'
    } else {
        $answer = Show-StdMessage -Owner $Ui.Window -Title 'Przerwać natychmiast?' -Image Warning -Button YesNo -Danger `
            -Message "Bieżący krok zostanie zatrzymany od razu, razem z uruchomionym instalatorem. Program może zostać zainstalowany tylko częściowo.`n`nPrzerwać natychmiast?"
        if ($answer -ne [System.Windows.MessageBoxResult]::Yes -or -not (Test-StdMainEngineBusy -Ui $Ui)) { return }
        Stop-StdEngine -Engine $Ui.Engine -Force
        $Ui.ForceRequested = $true
        Write-StdLog -Level Warn -Source 'UI' -Message 'Technik przerwał wdrożenie natychmiast (z zatrzymaniem bieżącego kroku).'
    }
    if ($Ui.State -eq 'Paused') { Set-StdMainState -Ui $Ui -State Paused } else { Set-StdMainState -Ui $Ui -State Running }
}

function Invoke-StdMainRetry {
    # „Ponów nieudane”: kroki z błędem i przerwane, w tym samym wdrożeniu.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $failed = Get-StdUiFailedStepIds -Steps @($Ui.StepItems)
    if ($failed.Count -eq 0) { return }
    if ($Ui.Preview) { Set-StdMainPreview -Ui $Ui -State Finished; return }
    $answer = Show-StdMessage -Owner $Ui.Window -Title 'Ponowić nieudane kroki?' -Image Question -Button YesNo `
        -Message ('Kroki z błędem i przerwane ({0}) zostaną uruchomione jeszcze raz. Kroki zakończone poprawnie nie będą powtarzane.' -f $failed.Count)
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
    Start-StdMainRun -Ui $Ui -OnlySteps $failed
}

function Copy-StdMainSummary {
    # „Kopiuj podsumowanie”: tekst wdrożenia do schowka (bez haseł).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $eff = $Ui.Effective
    $data = @{
        Version        = [string](Get-StdVersion)
        DeploymentId   = $Ui.DeploymentId
        ClientName     = $null
        ClientId       = $Ui.ClientId
        Revision       = $null
        SiteName       = $null
        SiteId         = $null
        ProfileName    = $null
        Computer       = $Ui.ComputerInfo
        Status         = $Ui.RunStatus
        RebootRequired = [bool]$Ui.RebootRequired
        Error          = $Ui.RunError
        StartedAt      = $Ui.RunStartedAt
        FinishedAt     = $Ui.RunFinishedAt
        LogPath        = $Ui.LogPath
        Steps          = @($Ui.StepItems)
    }
    if ($null -ne $eff) {
        $data.ClientName = $eff['ClientName']
        $data.ClientId = $eff['ClientId']
        $data.Revision = $eff['Revision']
        $data.SiteName = Get-StdUiText $eff['Site'] 'name'
        $data.SiteId = Get-StdUiText $eff['Site'] 'id'
        $data.ProfileName = $eff['ProfileName']
    }
    Copy-StdUiText -Text (Format-StdUiRunSummary -Data $data) -Window $Ui.Window -Message 'Podsumowanie wdrożenia skopiowane do schowka.'
}

function Reset-StdMainToSelect {
    # „Wróć do wyboru” po zakończonym wdrożeniu (plan z poprzedniego sprawdzenia jest już nieaktualny).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if (Test-StdMainEngineBusy -Ui $Ui) { return }
    $Ui.Plan = $null
    $Ui.PlanError = $null
    Update-StdMainPlanView -Ui $Ui
    $Ui.C.tabSelect.SelectedIndex = 0
    $Ui.Progress = 0
    Set-StdMainState -Ui $Ui -State Select
}

# ---------------------------------------------------------------------------------------------
# Silnik: uchwyt, zdarzenia, koniec pracy
# ---------------------------------------------------------------------------------------------

function Test-StdMainEngineBusy {
    # $true, gdy silnik (plan albo wdrożenie) jeszcze pracuje.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    return ($null -ne $Ui.Engine -and -not $Ui.EngineDone)
}

function Set-StdMainEngine {
    # Nowy uchwyt silnika; poprzedni (zakończony, jeszcze niezwolniony) trafia do listy do zwolnienia.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [Parameter(Mandatory = $true)]$Engine, [Parameter(Mandatory = $true)][string]$Mode)
    if ($null -ne $Ui.Engine) { $Ui.Retired.Add($Ui.Engine) }
    $Ui.Engine = $Engine
    $Ui.EngineMode = $Mode
    $Ui.EngineDone = $false
    $Ui.EnginePaused = $false
    $Ui.PauseRequested = $false
    $Ui.CancelRequested = $false
    $Ui.ForceRequested = $false
    $Ui.Progress = 0
    $Ui.ProgressText = ''
}

function Invoke-StdMainTick {
    # Co 200 ms (DispatcherTimer okna): wpisy logu okna, zdarzenia silnika, zwalnianie zakończonych silników,
    # zamykanie okna po przerwaniu wdrożenia. Bez ponownego wejścia (np. gdy wynik otwiera okno komunikatu).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.InTick) { return }
    $Ui.InTick = $true
    try {
        Receive-StdMainUiLog -Ui $Ui
        Receive-StdMainEngine -Ui $Ui
        for ($i = $Ui.Retired.Count - 1; $i -ge 0; $i--) {
            $old = $Ui.Retired[$i]
            if ($null -eq $old.AsyncResult -or $old.AsyncResult.IsCompleted) {
                try { Close-StdEngine -Engine $old } catch { }
                $Ui.Retired.RemoveAt($i)
            }
        }
        if ($Ui.CloseRequested -and -not $Ui.ForceClose) {
            $elapsed = ((Get-Date) - $Ui.CloseRequestedAt).TotalSeconds
            if (-not (Test-StdMainEngineBusy -Ui $Ui) -or $elapsed -ge 20) {
                $Ui.ForceClose = $true
                $Ui.Window.Close()
            }
        }
    } finally {
        $Ui.InTick = $false
    }
}

function Receive-StdMainUiLog {
    # Wpisy logu z wątku okna (np. pobieranie konfiguracji) do dziennika na żywo.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $queue = $Ui.UiLogQueue
    if ($null -eq $queue -or $queue.IsEmpty) { return }
    $items = New-Object 'System.Collections.Generic.List[object]'
    $record = $null
    while ($items.Count -lt 500 -and $queue.TryDequeue([ref]$record)) {
        $items.Add((New-StdUiLogItem -Time (Get-StdUiValue $record 'Time') -Level ([string](Get-StdUiValue $record 'Level')) -Source ([string](Get-StdUiValue $record 'Source')) -Message ([string](Get-StdUiValue $record 'Message'))))
    }
    if ($items.Count -gt 0) { Add-StdMainLog -Ui $Ui -Items $items.ToArray() }
}

function Receive-StdMainEngine {
    # Zdarzenia silnika (paczka do 500): Log, Plan, StepState, Progress, Paused, Resumed, Finished.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $engine = $Ui.Engine
    if ($null -eq $engine) { return }
    $events = @(Receive-StdEngineEvent -Engine $engine -Max 500)
    $logItems = New-Object 'System.Collections.Generic.List[object]'
    $finished = $null
    $stepsChanged = $false
    $stateChanged = $false
    foreach ($engineEvent in $events) {
        if ($null -eq $engineEvent) { continue }
        $type = [string]$engineEvent.Type
        if ($type -eq 'Log') {
            $logItems.Add((New-StdUiLogItem -Time $engineEvent.Time -Level ([string]$engineEvent.Level) -Source ([string]$engineEvent.Source) -Message ([string]$engineEvent.Message)))
        } elseif ($type -eq 'Plan') {
            $Ui.Plan = @($engineEvent.Steps)
        } elseif ($type -eq 'StepState') {
            Set-StdMainStep -Ui $Ui -Id ([string]$engineEvent.StepId) -Title ([string]$engineEvent.Title) -State ([string]$engineEvent.State) -Detail ([string]$engineEvent.Detail)
            $stepsChanged = $true
        } elseif ($type -eq 'Progress') {
            $Ui.Progress = [int]$engineEvent.Percent
            $Ui.ProgressText = [string]$engineEvent.Text
        } elseif ($type -eq 'Paused') {
            $Ui.EnginePaused = $true
            $stateChanged = $true
        } elseif ($type -eq 'Resumed') {
            $Ui.EnginePaused = $false
            $Ui.PauseRequested = $false
            $stateChanged = $true
        } elseif ($type -eq 'Finished') {
            $finished = $engineEvent
        }
    }
    if ($logItems.Count -gt 0) { Add-StdMainLog -Ui $Ui -Items $logItems.ToArray() }
    if ($Ui.EngineMode -eq 'Run' -and -not $Ui.EngineDone) {
        if ($stateChanged) {
            if ($Ui.EnginePaused) { Set-StdMainState -Ui $Ui -State Paused } else { Set-StdMainState -Ui $Ui -State Running }
        } elseif ($stepsChanged) {
            Update-StdMainRunBanner -Ui $Ui
        }
    }
    if ($events.Count -gt 0 -and -not $Ui.EngineDone) {
        if ($Ui.EngineMode -eq 'Plan') { Update-StdMainPlanView -Ui $Ui }
        Update-StdMainStatus -Ui $Ui
    }

    $completed = ($null -ne $engine.AsyncResult -and $engine.AsyncResult.IsCompleted)
    if ($null -ne $finished -and -not $Ui.EngineDone) {
        Complete-StdMainEngine -Ui $Ui -Status ([string]$finished.Status) -ErrorText ([string]$finished.Error) -RebootRequired ([bool]$finished.RebootRequired)
    } elseif ($completed -and -not $Ui.EngineDone -and $engine.Shared.Events.IsEmpty) {
        # Wątek silnika skończył się bez zdarzenia Finished (awaria) - stan z Get-StdEngineResult.
        $result = Get-StdEngineResult -Engine $engine
        Complete-StdMainEngine -Ui $Ui -Status ([string]$result.Status) -ErrorText ([string]$result.Error) -RebootRequired ([bool]$result.RebootRequired)
    }
    if ($completed -and $Ui.EngineDone -and $engine.Shared.Events.IsEmpty -and [object]::ReferenceEquals($Ui.Engine, $engine)) {
        try { Close-StdEngine -Engine $engine } catch { }
        $Ui.Engine = $null
    }
}

function Complete-StdMainEngine {
    # Koniec pracy silnika: plan gotowy (i ewentualnie potwierdzenie wdrożenia) albo podsumowanie wdrożenia.
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [string]$Status, [string]$ErrorText, [bool]$RebootRequired)
    $Ui.EngineDone = $true
    if ($Ui.EngineMode -eq 'Plan') {
        $deploy = $Ui.DeployAfterPlan
        $Ui.DeployAfterPlan = $false
        if ($Status -eq 'Finished' -and $null -ne $Ui.Plan) {
            $Ui.PlanError = $null
            Set-StdMainState -Ui $Ui -State Plan
            if ($deploy -and -not $Ui.CloseRequested -and -not $Ui.Closed) { Confirm-StdMainDeploy -Ui $Ui }
            return
        }
        if ($Status -eq 'Cancelled') {
            $Ui.PlanError = 'Sprawdzanie przerwano.'
        } else {
            $text = $ErrorText
            if ([string]::IsNullOrWhiteSpace($text)) { $text = 'brak szczegółów (zobacz dziennik)' }
            $Ui.PlanError = 'Nie udało się sprawdzić planu: {0}' -f $text
            Show-StdToast -Window $Ui.Window -Message 'Sprawdzanie zakończyło się błędem.' -Kind Error
        }
        Set-StdMainState -Ui $Ui -State Select
        Update-StdMainPlanView -Ui $Ui
        return
    }
    $Ui.RunStatus = $Status
    $Ui.RunError = $ErrorText
    $Ui.RebootRequired = $RebootRequired
    $Ui.RunFinishedAt = Get-Date
    $Ui.Progress = 100
    $counts = Get-StdUiStepCounts -Steps @($Ui.StepItems)
    $outcome = Get-StdUiRunOutcome -Status $Status -Counts $counts -RebootRequired $RebootRequired -ErrorText $ErrorText
    $Ui.Outcome = @{ DeploymentId = $Ui.DeploymentId; Status = $Status; Counts = $counts; RebootRequired = $RebootRequired }
    Set-StdMainState -Ui $Ui -State $outcome.View
    $kind = @{ Ok = 'Success'; Warn = 'Warning'; Crit = 'Error' }[$outcome.Tone]
    if ($null -eq $kind) { $kind = 'Info' }
    Show-StdToast -Window $Ui.Window -Message $outcome.Title -Kind $kind -Seconds 6
}

function Set-StdMainStep {
    # Wiersz listy kroków ze zdarzenia StepState (nowy albo podmieniony w miejscu).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [string]$Id, [string]$Title, [string]$State, [string]$Detail)
    if ([string]::IsNullOrEmpty($Id)) { return }
    $item = New-StdUiStepItem -Id $Id -Title $Title -State $State -Detail $Detail
    if ($Ui.StepIndex.ContainsKey($Id)) {
        $index = [int]$Ui.StepIndex[$Id]
        $Ui.StepItems[$index] = $item
    } else {
        $Ui.StepIndex[$Id] = $Ui.StepItems.Count
        $Ui.StepItems.Add($item)
    }
    if ($item.State -eq 'Running') {
        $Ui.CurrentStep = $item.Title
        try { $Ui.C.lvSteps.ScrollIntoView($item) } catch { }
    }
}

function Add-StdMainLog {
    # Wpisy do dziennika na żywo: ostatnie 3000, filtr poziomu, przewijanie do najnowszego („Na bieżąco”).
    param([Parameter(Mandatory = $true)][hashtable]$Ui, [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Items)
    if ($Items.Count -eq 0) { return }
    $max = $script:StdUiLogMax
    foreach ($item in $Items) { $Ui.LogAll.Add($item) }
    $over = $Ui.LogAll.Count - $max
    if ($over -gt 0) { $Ui.LogAll.RemoveRange(0, $over) }
    $added = $null
    foreach ($item in $Items) {
        if ([int]$item.Rank -ge [int]$Ui.LogMinRank) { $Ui.LogItems.Add($item); $added = $item }
    }
    while ($Ui.LogItems.Count -gt $max) { $Ui.LogItems.RemoveAt(0) }
    if ($null -ne $added -and $Ui.C.chkFollow.IsChecked -eq $true -and $Ui.C.lvLog.IsVisible) {
        try { $Ui.C.lvLog.ScrollIntoView($added) } catch { }
    }
}

function Update-StdMainLogFilter {
    # Filtr poziomu dziennika: lista od nowa z zapamiętanych wpisów.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $index = $Ui.C.cmbLogLevel.SelectedIndex
    if ($index -lt 0 -or $index -ge $script:StdUiLogFilters.Count) { $index = 0 }
    $Ui.LogMinRank = [int]$script:StdUiLogFilters[$index].MinRank
    $Ui.LogItems.Clear()
    $last = $null
    foreach ($item in $Ui.LogAll) {
        if ([int]$item.Rank -ge $Ui.LogMinRank) { $Ui.LogItems.Add($item); $last = $item }
    }
    if ($null -ne $last -and $Ui.C.lvLog.IsVisible) { try { $Ui.C.lvLog.ScrollIntoView($last) } catch { } }
}

# ---------------------------------------------------------------------------------------------
# Stan widoku
# ---------------------------------------------------------------------------------------------

function Set-StdMainState {
    <#
    .SYNOPSIS
    Stan okna głównego: Select / Plan (widok wyboru, Plan = zakładka planu), Running / Paused (wdrożenie trwa -
    zablokowane wszystko poza Pauza / Wznów / Przerwij i dziennikiem), Error / Finished (koniec wdrożenia).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Ui,
        [Parameter(Mandatory = $true)][ValidateSet('Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished')][string]$State
    )
    $c = $Ui.C
    $previous = $Ui.State
    $Ui.State = $State
    $isRun = ($State -eq 'Running' -or $State -eq 'Paused' -or $State -eq 'Error' -or $State -eq 'Finished')
    $active = ($State -eq 'Running' -or $State -eq 'Paused')
    $done = ($State -eq 'Error' -or $State -eq 'Finished')
    Set-StdUiVisible $c.viewSelect (-not $isRun)
    Set-StdUiVisible $c.viewRun $isRun
    if ($State -eq 'Plan' -and $previous -ne 'Plan') { $c.tabSelect.SelectedIndex = 1 }
    if ($State -eq 'Plan' -or $State -eq 'Select') { Update-StdMainPlanView -Ui $Ui }

    # Przyciski paska.
    Set-StdUiVisible $c.btnCheck (-not $isRun)
    Set-StdUiVisible $c.btnDeploy (-not $isRun)
    Set-StdUiVisible $c.btnPause $active
    Set-StdUiVisible $c.btnCancel $active
    Set-StdUiVisible $c.btnBack $done
    Set-StdUiVisible $c.btnCopy $done
    $failed = Get-StdUiFailedStepIds -Steps @($Ui.StepItems)
    Set-StdUiVisible $c.btnRetry ($done -and $failed.Count -gt 0)
    if ($active) {
        if ($State -eq 'Paused') {
            Set-StdUiButton -Button $c.btnPause -Text 'Wznów' -Icon 'E768' -Style 'SuccessButton' -ToolTip 'Wznów wdrożenie'
            $c.btnPause.IsEnabled = -not $Ui.CancelRequested
        } elseif ($Ui.PauseRequested) {
            Set-StdUiButton -Button $c.btnPause -Text 'Wznów' -Icon 'E768' -Style 'SuccessButton' -ToolTip 'Pauza nastąpi po bieżącym kroku - kliknij, aby jej nie robić'
            $c.btnPause.IsEnabled = -not $Ui.CancelRequested
        } else {
            Set-StdUiButton -Button $c.btnPause -Text 'Pauza' -Icon 'E769' -Style 'WarningButton' -ToolTip 'Wstrzymaj wdrożenie po bieżącym kroku'
            $c.btnPause.IsEnabled = -not $Ui.CancelRequested
        }
        if ($Ui.CancelRequested) {
            Set-StdUiButton -Button $c.btnCancel -Text 'Przerwij natychmiast' -Icon 'E71A' -Style 'DangerPrimaryButton' -ToolTip 'Zatrzymaj od razu także bieżący krok'
            $c.btnCancel.IsEnabled = -not $Ui.ForceRequested
        } else {
            Set-StdUiButton -Button $c.btnCancel -Text 'Przerwij' -Icon 'E71A' -Style 'DangerButton' -ToolTip 'Przerwij wdrożenie po bieżącym kroku'
            $c.btnCancel.IsEnabled = $true
        }
    }
    # W trakcie wdrożenia także motyw jest zablokowany (SPEC 7.3: zablokowane wszystko poza Pauza / Wznów / Przerwij i dziennikiem).
    $c.btnTheme.IsEnabled = -not $active
    if (-not $isRun) { Update-StdMainSelectControls -Ui $Ui }
    if ($isRun) { Update-StdMainRunBanner -Ui $Ui }
    Update-StdMainStatus -Ui $Ui
    Update-StdMainColumns -Ui $Ui
}

function Update-StdMainRunBanner {
    # Karta stanu wdrożenia: ikona i kolor, tytuł, bieżący krok / wynik, plik logu, liczniki.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $counts = Get-StdUiStepCounts -Steps @($Ui.StepItems)
    $c.txtCountOk.Text = [string]$counts.Ok
    $c.txtCountSkipped.Text = [string]$counts.Skipped
    $c.txtCountError.Text = [string]$counts.Error
    $c.txtCountCancelled.Text = [string]$counts.Cancelled
    $doneSteps = $counts.Ok + $counts.Skipped + $counts.Error + $counts.Cancelled
    $c.txtStepsCount.Text = 'zakończone {0} z {1}' -f $doneSteps, $counts.Total

    $tone = 'Info'
    $glyph = 'E768'
    $title = 'Wdrażanie'
    $detail = 'Przygotowanie wdrożenia...'
    switch ($Ui.State) {
        'Running' {
            if (-not [string]::IsNullOrEmpty($Ui.CurrentStep)) { $detail = 'Bieżący krok: {0}' -f $Ui.CurrentStep }
            if ($Ui.ForceRequested) { $tone = 'Warn'; $glyph = 'E71A'; $title = 'Przerywanie natychmiast...' }
            elseif ($Ui.CancelRequested) { $tone = 'Warn'; $glyph = 'E71A'; $title = 'Przerywanie po bieżącym kroku...' }
            elseif ($Ui.PauseRequested) { $tone = 'Warn'; $glyph = 'E769'; $title = 'Pauza po bieżącym kroku...' }
            if ($Ui.CloseRequested) { $detail = 'Okno zamknie się po zatrzymaniu bieżącego kroku.' }
        }
        'Paused' {
            $tone = 'Warn'
            $glyph = 'E769'
            $title = 'Pauza'
            $detail = 'Wdrożenie czeka na wznowienie. Kliknij „Wznów”, aby kontynuować.'
            if ($Ui.CancelRequested) { $title = 'Przerywanie...'; $detail = 'Wdrożenie zostanie przerwane.' }
        }
        default {
            $outcome = Get-StdUiRunOutcome -Status ([string]$Ui.RunStatus) -Counts $counts -RebootRequired ([bool]$Ui.RebootRequired) -ErrorText ([string]$Ui.RunError)
            $tone = $outcome.Tone
            $glyph = $outcome.Glyph
            $title = $outcome.Title
            $detail = $outcome.Detail
        }
    }
    Set-StdUiColor $c.tileRun 'Background' ('Theme{0}Back' -f $tone)
    Set-StdUiColor $c.txtRunIcon 'Foreground' ('Theme{0}Fore' -f $tone)
    $c.txtRunIcon.Text = Get-StdGlyph -Code $glyph
    $c.txtRunTitle.Text = $title
    $c.txtRunDetail.Text = $detail
    if (-not [string]::IsNullOrEmpty($Ui.LogPath)) {
        $c.txtRunLog.Text = 'Wdrożenie {0} - log: {1}' -f $Ui.DeploymentId, $Ui.LogPath
        $c.txtRunLog.ToolTip = $c.txtRunLog.Text
        Set-StdUiVisible $c.txtRunLog $true
    } else {
        Set-StdUiVisible $c.txtRunLog $false
    }
}

function Update-StdMainStatus {
    # Pasek stanu na dole: kropka, tekst, pasek postępu (także na pasku zadań Windows).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    $dot = 'ThemeSuccess'
    $text = ''
    $percent = 0
    $taskbar = [System.Windows.Shell.TaskbarItemProgressState]::None
    $engineBusy = Test-StdMainEngineBusy -Ui $Ui
    switch ($Ui.State) {
        'Running' {
            $dot = 'ThemeInfo'
            $text = $Ui.ProgressText
            if ([string]::IsNullOrEmpty($text)) { $text = 'Wdrażanie...' }
            if ($Ui.PauseRequested) { $text = '{0} - pauza po bieżącym kroku' -f $text; $dot = 'ThemeWarning' }
            if ($Ui.CancelRequested) { $text = '{0} - przerywanie' -f $text; $dot = 'ThemeWarning' }
            $percent = $Ui.Progress
            $taskbar = [System.Windows.Shell.TaskbarItemProgressState]::Normal
        }
        'Paused' {
            $dot = 'ThemeWarning'
            $text = 'Pauza - wdrożenie czeka na wznowienie'
            $percent = $Ui.Progress
            $taskbar = [System.Windows.Shell.TaskbarItemProgressState]::Paused
        }
        'Error' {
            $dot = 'ThemeDanger'
            $text = (Get-StdUiRunOutcome -Status ([string]$Ui.RunStatus) -Counts (Get-StdUiStepCounts -Steps @($Ui.StepItems)) -RebootRequired ([bool]$Ui.RebootRequired) -ErrorText ([string]$Ui.RunError)).Title
            $percent = 100
            $taskbar = [System.Windows.Shell.TaskbarItemProgressState]::Error
        }
        'Finished' {
            $outcome = Get-StdUiRunOutcome -Status ([string]$Ui.RunStatus) -Counts (Get-StdUiStepCounts -Steps @($Ui.StepItems)) -RebootRequired ([bool]$Ui.RebootRequired) -ErrorText ([string]$Ui.RunError)
            $text = $outcome.Title
            if ($outcome.Tone -eq 'Warn') { $dot = 'ThemeWarning' }
            $percent = 100
        }
        default {
            if ($engineBusy) {
                $dot = 'ThemeInfo'
                $text = $Ui.ProgressText
                if ([string]::IsNullOrEmpty($text)) { $text = 'Sprawdzanie...' }
                $percent = $Ui.Progress
                $taskbar = [System.Windows.Shell.TaskbarItemProgressState]::Normal
            } elseif ($null -ne $Ui.ReleaseTask) {
                $dot = 'ThemeInfo'
                $text = 'Pobieranie konfiguracji klienta...'
            } elseif ($Ui.SummaryMode -eq 'Error') {
                $dot = 'ThemeDanger'
                $text = 'Konfiguracja klienta jest niedostępna'
            } elseif ($null -ne $Ui.Effective) {
                if ($null -ne $Ui.Plan) {
                    $counts = $Ui.PlanCounts
                    $text = 'Plan gotowy: do zrobienia {0}, już zrobione {1}, problem {2}.' -f $counts.Todo, $counts.Done, $counts.Blocked
                    if ($counts.Blocked -gt 0) { $dot = 'ThemeWarning' }
                } else {
                    $text = 'Gotowe do sprawdzenia i wdrożenia.'
                }
            } else {
                $dot = 'ThemeWarning'
                $text = 'Wybierz klienta, lokalizację i profil.'
            }
        }
    }
    Set-StdUiColor $c.dotStatus 'Fill' $dot
    $c.txtStatus.Text = $text
    $c.txtStatus.ToolTip = $text
    if ($percent -lt 0) { $percent = 0 }
    if ($percent -gt 100) { $percent = 100 }
    $c.barProgress.Value = $percent
    $info = $Ui.Window.TaskbarItemInfo
    if ($null -ne $info) {
        $info.ProgressState = $taskbar
        $info.ProgressValue = $percent / 100.0
    }
}

function Update-StdMainColumns {
    # Kolumny tabel dopasowane do szerokości (tekst się zawija, bez poziomego przewijania).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $c = $Ui.C
    Update-StdUiGridColumns -ListView $c.lvPlan -Flex @{ 1 = 0.4; 3 = 0.6 }
    Update-StdUiGridColumns -ListView $c.lvSteps -Flex @{ 1 = 1.0 } -MinWidth 120
    Update-StdUiGridColumns -ListView $c.lvLog -Flex @{ 2 = 1.0 } -MinWidth 120
}

# ---------------------------------------------------------------------------------------------
# Start pracy okna i sprzątanie
# ---------------------------------------------------------------------------------------------

function Start-StdMainWork {
    # Po pokazaniu okna: timer zdarzeń, log okna w dzienniku, informacje o komputerze, ostatni klient.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    Update-StdMainColumns -Ui $Ui
    if ($Ui.Preview -or $Ui.Started) { return }
    $Ui.Started = $true
    $queue = $Ui.UiLogQueue
    $sink = { $queue.Enqueue($args[0]) }.GetNewClosure()
    Register-StdLogSink -Name $Ui.LogSinkName -ScriptBlock $sink
    $Ui.Timer.Start()
    $null = Start-StdUiTask -Name 'Informacje o komputerze' -Owner $Ui.Window -ScriptText $script:StdMainComputerScript `
        -OnComplete (New-StdUiCallback -Ui $Ui -Handler { param($Ui, $Result) Set-StdMainComputer -Ui $Ui -Info $Result }) `
        -OnError (New-StdUiCallback -Ui $Ui -Handler { param($Ui, $ErrorRecord) Set-StdMainComputer -Ui $Ui -Info $null -ErrorText (Get-StdUiErrorMessage $ErrorRecord -MaxLength 300) })

    # Ostatnio używany klient (prefs.json) albo jedyny na liście.
    $items = @($Ui.ClientItems)
    $select = Get-StdMainDefaultPick -Items $items
    if ($null -ne $select) {
        foreach ($item in $items) {
            if ([string]::Equals([string]$item.Id, $select, [System.StringComparison]::OrdinalIgnoreCase)) {
                $Ui.Suppress = $true
                try { $Ui.C.cmbClient.SelectedItem = $item } finally { $Ui.Suppress = $false }
                Select-StdMainClient -Ui $Ui -ClientId ([string]$item.Id)
                break
            }
        }
    }
}

function Stop-StdMainWork {
    # Zamknięcie okna: timer, odbiornik logu, silnik (zakończony - zwolniony; pracujący - zatrzymanie bez czekania).
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    $Ui.Closed = $true
    if ($null -ne $Ui.Timer) { try { $Ui.Timer.Stop() } catch { } }
    if ($Ui.Started) { try { Unregister-StdLogSink -Name $Ui.LogSinkName } catch { } }
    $engines = New-Object 'System.Collections.Generic.List[object]'
    foreach ($old in $Ui.Retired) { $engines.Add($old) }
    if ($null -ne $Ui.Engine) { $engines.Add($Ui.Engine) }
    foreach ($engine in $engines) {
        try {
            if ($null -eq $engine.AsyncResult -or $engine.AsyncResult.IsCompleted) {
                Close-StdEngine -Engine $engine
            } else {
                Stop-StdEngine -Engine $engine -Force
                [void]$engine.PowerShell.BeginStop($null, $null)
                Write-StdLog -Level Warn -Source 'UI' -Message ('Okno zamknięte w trakcie pracy silnika ({0}) - silnik zatrzymany.' -f $engine.Id)
            }
        } catch { }
    }
    $Ui.Retired.Clear()
    $Ui.Engine = $null
}

function Confirm-StdMainClosing {
    # Zamknięcie okna w trakcie wdrożenia: potwierdzenie, przerwanie natychmiast, zamknięcie po zatrzymaniu kroku.
    # Zwraca $true, gdy okno może się zamknąć od razu.
    param([Parameter(Mandatory = $true)][hashtable]$Ui)
    if ($Ui.Preview -or $Ui.ForceClose) { return $true }
    if (-not (Test-StdMainEngineBusy -Ui $Ui)) { return $true }
    if ($Ui.EngineMode -ne 'Run') {
        # „Sprawdź” niczego nie zmienia - zatrzymujemy je bez pytania.
        try { Stop-StdEngine -Engine $Ui.Engine -Force } catch { }
        return $true
    }
    if ($Ui.CloseRequested) {
        Show-StdToast -Window $Ui.Window -Message 'Czekam na zatrzymanie bieżącego kroku - okno zamknie się samo.' -Kind Warning
        return $false
    }
    $answer = Show-StdMessage -Owner $Ui.Window -Title 'Przerwać wdrożenie i zamknąć okno?' -Image Warning -Button YesNo -Danger `
        -Message "Wdrożenie wciąż trwa. Zamknięcie okna przerwie je natychmiast, razem z bieżącym krokiem (instalator może zostać przerwany w połowie).`n`nZamknąć okno?"
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return $false }
    if (-not (Test-StdMainEngineBusy -Ui $Ui)) { return $true }
    Stop-StdEngine -Engine $Ui.Engine -Force
    $Ui.CancelRequested = $true
    $Ui.ForceRequested = $true
    $Ui.CloseRequested = $true
    $Ui.CloseRequestedAt = Get-Date
    Write-StdLog -Level Warn -Source 'UI' -Message 'Technik zamknął okno w trakcie wdrożenia - przerwanie natychmiast.'
    Set-StdMainState -Ui $Ui -State $Ui.State
    return $false
}

# ---------------------------------------------------------------------------------------------
# Podgląd okna (bez serwera i bez silnika)
# ---------------------------------------------------------------------------------------------

function Set-StdMainPreview {
    <#
    .SYNOPSIS
    Wypełnia okno danymi podglądu ($Ui.PreviewData, Merge-StdUiPreviewData) i pokazuje stan -State:
    Select, Plan, Running, Paused, Error, Finished.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Ui,
        [ValidateSet('Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished')][string]$State = 'Select'
    )
    $data = Merge-StdUiPreviewData -PreviewData $Ui.PreviewSource -State $State
    $Ui.PreviewData = $data
    $c = $Ui.C
    if (-not $Ui.PreviewFilled) {
        $Ui.PreviewFilled = $true
        $Ui.UserName = [string](Get-StdUiValue $data 'UserName')
        $Ui.Index = Get-StdUiValue $data 'Index'
        $Ui.IndexSource = [string](Get-StdUiValue $data 'IndexSource')
        $Ui.IndexRevision = Get-StdUiValue $Ui.Index 'revision'
        $release = Get-StdUiValue $data 'Release'
        $clientId = [string](Get-StdUiValue $data 'ClientId')
        if ($null -ne $release -and -not [string]::IsNullOrEmpty($clientId)) {
            $Ui.ReleaseCache[$clientId] = @{ Release = $release; Source = [string](Get-StdUiValue $data 'ReleaseSource'); Revision = (Get-StdUiValue $release 'revision'); Kid = 'PODGLAD'; Attempts = @() }
        }
        Set-StdMainClients -Ui $Ui
        $Ui.PendingSite = [string](Get-StdUiValue $data 'SiteId')
        $Ui.PendingProfile = [string](Get-StdUiValue $data 'ProfileName')
        foreach ($item in @($Ui.ClientItems)) {
            if ([string]$item.Id -eq $clientId) {
                Set-StdMainComboItems -Ui $Ui -Combo $c.cmbClient -Items @($Ui.ClientItems) -SelectId $clientId
                Select-StdMainClient -Ui $Ui -ClientId $clientId
                break
            }
        }
        $Ui.PendingSite = $null
        $Ui.PendingProfile = $null
        $releaseError = [string](Get-StdUiValue $data 'ReleaseError')
        if (-not [string]::IsNullOrWhiteSpace($releaseError)) {
            $Ui.Effective = $null
            Set-StdMainSummary -Ui $Ui -Mode Error -Title 'Nie udało się pobrać konfiguracji klienta' -Text $releaseError
        }
        Set-StdMainComputer -Ui $Ui -Info (Get-StdUiValue $data 'ComputerInfo')
        $Ui.LogAll.Clear()
        $Ui.LogItems.Clear()
        $logItems = foreach ($entry in @(Get-StdUiValue $data 'Log')) {
            if ($null -ne $entry) { New-StdUiLogItem -Time (Get-StdUiValue $entry 'Time') -Level ([string](Get-StdUiValue $entry 'Level')) -Source ([string](Get-StdUiValue $entry 'Source')) -Message ([string](Get-StdUiValue $entry 'Message')) }
        }
        Add-StdMainLog -Ui $Ui -Items @($logItems)
    }

    $Ui.Plan = $null
    $Ui.PlanError = $null
    $Ui.PauseRequested = $false
    $Ui.CancelRequested = $false
    $Ui.ForceRequested = $false
    if ($State -eq 'Plan') { $Ui.Plan = @(Get-StdUiValue $data 'Plan') }
    if ($State -eq 'Select' -or $State -eq 'Plan') {
        Set-StdMainState -Ui $Ui -State $State
        if ($State -eq 'Select') { $c.tabSelect.SelectedIndex = 0 }
        return
    }
    $Ui.EngineMode = 'Run'
    $Ui.StepItems.Clear()
    $Ui.StepIndex = @{}
    $Ui.CurrentStep = $null
    foreach ($step in @(Get-StdUiValue $data 'Steps')) {
        if ($null -eq $step) { continue }
        Set-StdMainStep -Ui $Ui -Id ([string](Get-StdUiValue $step 'Id')) -Title ([string](Get-StdUiValue $step 'Title')) -State ([string](Get-StdUiValue $step 'State')) -Detail ([string](Get-StdUiValue $step 'Detail'))
    }
    $progress = Get-StdUiValue $data 'Progress'
    $Ui.Progress = [int](Get-StdUiValue $progress 'Percent')
    $Ui.ProgressText = [string](Get-StdUiValue $progress 'Text')
    $Ui.DeploymentId = [string](Get-StdUiValue $data 'DeploymentId')
    $Ui.LogPath = 'C:\ProgramData\STD\Logs\{0}\deploy.log' -f $Ui.DeploymentId
    $Ui.RunStartedAt = (Get-Date).Date.AddHours(10).AddMinutes(15)
    $result = Get-StdUiValue $data 'Result'
    $Ui.RunStatus = 'Running'
    $Ui.RunError = $null
    $Ui.RebootRequired = $false
    $Ui.EnginePaused = ($State -eq 'Paused')
    if ($State -eq 'Error' -or $State -eq 'Finished') {
        $Ui.RunStatus = [string](Get-StdUiValue $result 'Status')
        if ([string]::IsNullOrEmpty($Ui.RunStatus)) { $Ui.RunStatus = 'Finished' }
        $Ui.RunError = [string](Get-StdUiValue $result 'Error')
        $Ui.RebootRequired = ((Get-StdUiValue $result 'RebootRequired') -eq $true)
        $Ui.RunFinishedAt = $Ui.RunStartedAt.AddMinutes(12)
        $Ui.Progress = 100
    }
    Set-StdMainState -Ui $Ui -State $State
}

# ---------------------------------------------------------------------------------------------
# Okno
# ---------------------------------------------------------------------------------------------

function New-StdMainWindow {
    <#
    .SYNOPSIS
    Buduje okno główne (bez pokazywania). Stan okna: Window.Tag ($ui). Praca w tle (informacje o komputerze,
    konfiguracja ostatniego klienta, timer zdarzeń) startuje dopiero po pokazaniu okna (Loaded).
    .PARAMETER PreviewData
    Podgląd bez serwera i bez silnika. Klucze (brakujące = przykładowe z Get-StdUiPreviewSample): UserName, Index,
    IndexSource, Release, ReleaseSource, ReleaseError, ClientId, SiteId, ProfileName, ComputerInfo, Plan
    (lista @{ Id; Title; Phase; Enabled; State; Detail }), Steps (lista @{ Id; Title; State; Detail }), Log (lista
    @{ Time; Level; Source; Message }), Progress (@{ Percent; Text }), Result (@{ Status; RebootRequired; Error }),
    DeploymentId.
    .PARAMETER PreviewState
    Stan pokazywany w podglądzie: Select, Plan, Running, Paused, Error, Finished (domyślnie Select).
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$Session,
        [AllowNull()]$Index,
        [AllowNull()][AllowEmptyString()][string]$IndexSource,
        [AllowNull()]$Bootstrap,
        [AllowNull()][AllowEmptyString()][string]$OriginRoot,
        [AllowNull()][hashtable]$PreviewData,
        [ValidateSet('Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished')][string]$PreviewState = 'Select'
    )
    $preview = ($null -ne $PreviewData) -or $PSBoundParameters.ContainsKey('PreviewState')
    $win = New-StdWindow -Xaml $script:StdMainXaml
    $c = Get-StdUiControls -Root $win -Names $script:StdMainControlNames
    $ui = @{
        Kind              = 'Main'
        Window            = $win
        C                 = $c
        Preview           = $preview
        PreviewSource     = $PreviewData
        PreviewData       = $null
        PreviewFilled     = $false
        Session           = $Session
        UserName          = $null
        Index             = $Index
        IndexSource       = $IndexSource
        IndexRevision     = (Get-StdUiValue $Index 'revision')
        Bootstrap         = $Bootstrap
        OriginRoot        = $OriginRoot
        Prefs             = $null
        ClientItems       = @()
        ClientId          = $null
        Release           = $null
        ReleaseCache      = @{}
        ReleaseTask       = $null
        ReleaseTaskClient = $null
        PendingSite       = $null
        PendingProfile    = $null
        Effective         = $null
        SummaryMode       = 'Empty'
        ComputerInfo      = $null
        Plan              = $null
        PlanError         = $null
        PlanCounts        = @{ Todo = 0; Done = 0; Blocked = 0; Disabled = 0 }
        PlanItems         = (New-Object 'System.Collections.ObjectModel.ObservableCollection[object]')
        StepItems         = (New-Object 'System.Collections.ObjectModel.ObservableCollection[object]')
        StepIndex         = @{}
        LogAll            = (New-Object 'System.Collections.Generic.List[object]')
        LogItems          = (New-Object 'System.Collections.ObjectModel.ObservableCollection[object]')
        LogMinRank        = 0
        UiLogQueue        = (New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]')
        LogSinkName       = 'StdUi.MainWindow'
        Engine            = $null
        EngineMode        = $null
        EngineDone        = $false
        EnginePaused      = $false
        Retired           = (New-Object 'System.Collections.Generic.List[object]')
        DeployAfterPlan   = $false
        PauseRequested    = $false
        CancelRequested   = $false
        ForceRequested    = $false
        DeploymentId      = $null
        LogPath           = $null
        RunStatus         = $null
        RunError          = $null
        RebootRequired    = $false
        RunStartedAt      = $null
        RunFinishedAt     = $null
        CurrentStep       = $null
        Progress          = 0
        ProgressText      = ''
        Outcome           = $null
        State             = 'Select'
        Suppress          = $false
        InTick            = $false
        Started           = $false
        Closed            = $false
        CloseRequested    = $false
        CloseRequestedAt  = $null
        ForceClose        = $false
        Timer             = $null
    }
    $win.Tag = $ui
    $ui.Prefs = Get-StdUiPrefs
    try { $win.TaskbarItemInfo = New-Object System.Windows.Shell.TaskbarItemInfo } catch { }

    # Treść stała.
    $c.txtVersion.Text = [string](Get-StdVersion)
    $c.lvPlan.ItemsSource = $ui.PlanItems
    $c.lvSteps.ItemsSource = $ui.StepItems
    $c.lvLog.ItemsSource = $ui.LogItems
    $c.cmbLogLevel.ItemsSource = @($script:StdUiLogFilters | ForEach-Object { [string]$_.Label })
    $c.cmbLogLevel.SelectedIndex = 0
    Set-StdUiButton -Button $c.btnCheck -Text 'Sprawdź' -Icon 'E9D5' -ToolTip 'Pokaż, co zostanie zrobione, co już jest i gdzie jest problem (niczego nie zmienia)'
    Set-StdUiButton -Button $c.btnDeploy -Text 'Wdróż' -Icon 'E768' -ToolTip 'Rozpocznij wdrożenie (najpierw potwierdzenie z listą kroków)'
    Set-StdUiButton -Button $c.btnBack -Text 'Wróć do wyboru' -Icon 'E72B'
    Set-StdUiButton -Button $c.btnCopy -Text 'Kopiuj podsumowanie' -Icon 'E8C8' -ToolTip 'Skopiuj wynik wdrożenia do schowka (np. do zgłoszenia)'
    Set-StdUiButton -Button $c.btnRetry -Text 'Ponów nieudane' -Icon 'E72C' -ToolTip 'Uruchom jeszcze raz kroki z błędem i przerwane'
    Set-StdUiButton -Button $c.btnRetryRelease -Text 'Spróbuj ponownie' -Icon 'E72C'
    Update-StdMainThemeButton -Ui $ui
    Set-StdMainSummary -Ui $ui -Mode Empty -Title 'Wybierz klienta, lokalizację i profil' -Text 'Po wyborze pojawi się tu podsumowanie konfiguracji: aplikacje, zadania, domena i nazwa komputera.'

    if (-not $preview) {
        Set-StdMainClients -Ui $ui
        if (@($ui.ClientItems).Count -eq 0) {
            Set-StdMainSummary -Ui $ui -Mode Error -Title 'Lista klientów jest pusta' -Text 'Na serwerze STD nie ma jeszcze opublikowanej konfiguracji żadnego klienta. Poproś administratora o publikację.'
        }
        $ui.Timer = New-StdUiTimer -Ui $ui -IntervalMs 200 -OnTick { param($Ui, $Timer) Invoke-StdMainTick -Ui $Ui }
    }

    # Handlery.
    $win.Add_Loaded({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-StdMainWork -Ui $Ui } })
    $win.Add_Closing({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                if (-not (Confirm-StdMainClosing -Ui $Ui)) { $e.Cancel = $true }
            }
        })
    $win.Add_Closed({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Stop-StdMainWork -Ui $Ui } })
    $c.btnTheme.Add_Click({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                Set-StdTheme -Dark (-not (Test-StdThemeDark)) -NoSave:$Ui.Preview
                Update-StdMainThemeButton -Ui $Ui
            }
        })
    $c.cmbClient.Add_SelectionChanged({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                if ($Ui.Suppress) { return }
                $item = $Ui.C.cmbClient.SelectedItem
                if ($null -ne $item) { Select-StdMainClient -Ui $Ui -ClientId ([string]$item.Id) }
            }
        })
    $c.cmbSite.Add_SelectionChanged({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) if (-not $Ui.Suppress) { Set-StdMainProfiles -Ui $Ui } } })
    $c.cmbProfile.Add_SelectionChanged({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) if (-not $Ui.Suppress) { Update-StdMainEffective -Ui $Ui } } })
    $c.btnRetryRelease.Add_Click({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                $clientId = $Ui.ClientId
                if ([string]::IsNullOrEmpty($clientId)) { return }
                $Ui.ReleaseCache.Remove($clientId)
                $Ui.ClientId = $null
                Select-StdMainClient -Ui $Ui -ClientId $clientId
            }
        })
    $c.chkShowDisabled.Add_Checked({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Update-StdMainPlanView -Ui $Ui } })
    $c.chkShowDisabled.Add_Unchecked({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Update-StdMainPlanView -Ui $Ui } })
    $c.cmbLogLevel.Add_SelectionChanged({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Update-StdMainLogFilter -Ui $Ui } })
    $c.lvLog.Add_MouseDoubleClick({
            param($s, $e)
            Invoke-StdUiHandler $s {
                param($Ui)
                $item = $Ui.C.lvLog.SelectedItem
                if ($null -eq $item) { return }
                $text = '{0} [{1}] {2} {3}' -f $item.Time, $item.Level, $item.Source, $item.Message
                Copy-StdUiText -Text $text -Window $Ui.Window -Message 'Wpis dziennika skopiowany do schowka.'
            }
        })
    foreach ($list in @($c.lvPlan, $c.lvSteps, $c.lvLog)) {
        $list.Add_SizeChanged({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Update-StdMainColumns -Ui $Ui } })
    }
    $c.btnCheck.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-StdMainPlan -Ui $Ui } -ErrorContext 'Nie udało się uruchomić sprawdzania.' })
    $c.btnDeploy.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Confirm-StdMainDeploy -Ui $Ui } -ErrorContext 'Nie udało się uruchomić wdrożenia.' })
    $c.btnPause.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Request-StdMainPause -Ui $Ui } })
    $c.btnCancel.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Request-StdMainCancel -Ui $Ui } })
    $c.btnRetry.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Invoke-StdMainRetry -Ui $Ui } -ErrorContext 'Nie udało się ponowić kroków.' })
    $c.btnCopy.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Copy-StdMainSummary -Ui $Ui } })
    $c.btnBack.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Reset-StdMainToSelect -Ui $Ui } })

    if ($preview) {
        Set-StdMainPreview -Ui $ui -State $PreviewState
    } else {
        Update-StdMainSelection -Ui $ui
        Set-StdMainState -Ui $ui -State Select
    }
    return $win
}

function Show-StdMainWindow {
    <#
    .SYNOPSIS
    Okno główne STD (wybór klienta, „Sprawdź”, „Wdróż”, widok wdrożenia). Wynik: @{ DeploymentId; Status; Counts;
    RebootRequired } ostatniego wdrożenia albo $null. -PreviewData / -PreviewState: podgląd jak w New-StdMainWindow.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$Session,
        [AllowNull()]$Index,
        [AllowNull()][AllowEmptyString()][string]$IndexSource,
        [AllowNull()]$Bootstrap,
        [AllowNull()][AllowEmptyString()][string]$OriginRoot,
        [AllowNull()][hashtable]$PreviewData,
        [ValidateSet('Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished')][string]$PreviewState = 'Select'
    )
    $windowParams = @{ Session = $Session; Index = $Index; IndexSource = $IndexSource; Bootstrap = $Bootstrap; OriginRoot = $OriginRoot; PreviewData = $PreviewData }
    if ($PSBoundParameters.ContainsKey('PreviewState')) { $windowParams.PreviewState = $PreviewState }
    $win = Select-StdUiWindow -Output @(New-StdMainWindow @windowParams)
    [void]$win.ShowDialog()
    return $win.Tag.Outcome
}
