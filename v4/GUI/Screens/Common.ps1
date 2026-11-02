# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Screens\Common.ps1 - wspólne funkcje ekranów STD 4 (SPEC 7.3): stan okna w handlerach, pomocnicze
# funkcje WPF (kolory z palety, wiersze "etykieta: wartość", schowek), teksty po polsku (stany kroków,
# nazwy zadań, źródła konfiguracji), ostatnio używane pozycje (prefs.json), Wi-Fi (netsh, profil XML),
# komunikaty błędów logowania, podsumowanie wdrożenia i przykładowe dane dla podglądu okien.
#
# Plik wczytuje GUI\Std.UI.ps1 (razem z innymi ekranami). Funkcje bez WPF (opisane "bez WPF") działają
# też w PowerShell 7 na Linuksie - testy wczytują ten plik osobno.
#
# Stan okna i handlery
# --------------------
# Każde okno ma jeden obiekt stanu [hashtable] ($ui) w Window.Tag (klucz Window = okno). Handlery
# zdarzeń dostają go jawnie: Invoke-StdUiHandler szuka stanu po nadawcy zdarzenia (okno, kontrolka
# w oknie, DispatcherTimer z Tag = stan), a wynik zadania w tle trafia do funkcji przez
# New-StdUiCallback. Dzięki temu handlery nie zależą od zmiennych funkcji, która wywołała ShowDialog().

# ---------------------------------------------------------------------------------------------
# Teksty i stałe (bez WPF)
# ---------------------------------------------------------------------------------------------

# Zadania z SPEC 2.8 w kolejności wdrożenia, z nazwą dla technika.
$script:StdUiTaskTitles = [ordered]@{
    WaitForNetwork        = 'Czekanie na sieć'
    SuspendSleep          = 'Wstrzymanie usypiania'
    ConnectWiFi           = 'Połączenie z Wi-Fi'
    RemoveBloatware       = 'Usunięcie zbędnych aplikacji Sklepu'
    UninstallPreinstalled = 'Usunięcie preinstalowanych programów'
    UninstallM365         = 'Odinstalowanie Microsoft 365'
    UninstallOneDrive     = 'Odinstalowanie OneDrive'
    InstallApps           = 'Instalacja aplikacji'
    CreateLocalAdmin      = 'Konto lokalnego administratora'
    RenameComputer        = 'Zmiana nazwy komputera'
    JoinDomain            = 'Dołączenie do domeny'
    SystemSettings        = 'Ustawienia systemu'
    RunScripts            = 'Skrypty poinstalacyjne'
    HardwareAudit         = 'Audyt sprzętu'
    EnableBitLocker       = 'Włączenie BitLockera'
    RebootAtEnd           = 'Restart na końcu'
}

# Stany kroków (silnik: Pending, Running, Ok, Skipped, Error, Cancelled) i planu (Done, Todo, Blocked,
# Disabled): ikona (kod znaku Segoe Fluent Icons / MDL2), napis i ton koloru. Kolory w XAML ekranów
# (DataTrigger po State) muszą odpowiadać tonom z tej tabeli.
$script:StdUiStateTable = [ordered]@{
    Pending   = @{ Glyph = 'E916'; Label = 'Czeka'; Tone = 'Subtle' }
    Running   = @{ Glyph = 'E768'; Label = 'W toku'; Tone = 'Info' }
    Ok        = @{ Glyph = 'E73E'; Label = 'OK'; Tone = 'Ok' }
    Skipped   = @{ Glyph = 'E72A'; Label = 'Pominięto'; Tone = 'Subtle' }
    Error     = @{ Glyph = 'EA39'; Label = 'Błąd'; Tone = 'Crit' }
    Cancelled = @{ Glyph = 'E711'; Label = 'Przerwano'; Tone = 'Warn' }
    Done      = @{ Glyph = 'E930'; Label = 'Już zrobione'; Tone = 'Ok' }
    Todo      = @{ Glyph = 'E710'; Label = 'Do zrobienia'; Tone = 'Info' }
    Blocked   = @{ Glyph = 'E7BA'; Label = 'Problem'; Tone = 'Warn' }
    Disabled  = @{ Glyph = 'E738'; Label = 'Wyłączony'; Tone = 'Subtle' }
}

# Poziomy logu: ranga (filtr dziennika), ikona i napis.
$script:StdUiLogLevels = @{
    Debug = @{ Rank = 0; Glyph = 'E90F'; Label = 'Debug' }
    Info  = @{ Rank = 1; Glyph = ''; Label = 'Info' }
    Warn  = @{ Rank = 2; Glyph = 'E7BA'; Label = 'Uwaga' }
    Error = @{ Rank = 3; Glyph = 'EA39'; Label = 'Błąd' }
}

# Filtr dziennika: napis na liście i najniższa pokazywana ranga.
$script:StdUiLogFilters = @(
    @{ Label = 'Wszystkie wpisy'; MinRank = 0 }
    @{ Label = 'Informacje i ważniejsze'; MinRank = 1 }
    @{ Label = 'Ostrzeżenia i błędy'; MinRank = 2 }
    @{ Label = 'Tylko błędy'; MinRank = 3 }
)

$script:StdUiLogMax = 3000          # dziennik na żywo: ostatnie wpisy (SPEC 7.3)
$script:StdUiRecentMax = 8          # ostatnio używane pozycje na listę

function Get-StdUiSourceLabel {
    <#
    .SYNOPSIS
    Napis pastylki źródła konfiguracji (bez WPF): server -> serwer, mirror -> kopia lustrzana, cache, usb -> pendrive.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()][AllowEmptyString()][string]$Source)
    switch ([string]$Source) {
        'server' { return 'serwer' }
        'mirror' { return 'kopia lustrzana' }
        'cache' { return 'cache' }
        'usb' { return 'pendrive' }
        'preview' { return 'podgląd' }
    }
    if ([string]::IsNullOrEmpty($Source)) { return 'brak' }
    return [string]$Source
}

function Get-StdUiSourceTone {
    # Kolor kropki pastylki źródła: serwer = zielony, kopia lustrzana = niebieski, kopia offline = pomarańczowy.
    param([AllowNull()][AllowEmptyString()][string]$Source)
    switch ([string]$Source) {
        'server' { return 'ThemeSuccess' }
        'mirror' { return 'ThemeInfo' }
        'preview' { return 'ThemeInfo' }
    }
    return 'ThemeWarning'
}

function Get-StdUiTaskTitle {
    <#
    .SYNOPSIS
    Nazwa zadania (SPEC 2.8) dla technika (bez WPF). Nieznany identyfikator wraca bez zmian.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true, Position = 0)][string]$TaskId)
    if ($script:StdUiTaskTitles.Contains($TaskId)) { return [string]$script:StdUiTaskTitles[$TaskId] }
    return $TaskId
}

function Get-StdUiStateInfo {
    <#
    .SYNOPSIS
    Ikona, napis i ton stanu kroku albo pozycji planu (bez WPF): @{ State; Glyph; Label; Tone; ColorKey }.
    Nieznany stan = Pending (z napisem równym stanowi).
    #>
    [CmdletBinding()]
    param([Parameter(Position = 0)][AllowNull()][AllowEmptyString()][string]$State)
    $known = $null
    foreach ($key in $script:StdUiStateTable.Keys) {
        if ([string]::Equals($key, $State, [System.StringComparison]::OrdinalIgnoreCase)) { $known = $key; break }
    }
    if ($null -eq $known) {
        $entry = $script:StdUiStateTable['Pending']
        $label = $entry.Label
        if (-not [string]::IsNullOrEmpty($State)) { $label = $State }
        return @{ State = 'Pending'; Glyph = $entry.Glyph; Label = $label; Tone = $entry.Tone; ColorKey = (Get-StdUiToneColorKey $entry.Tone) }
    }
    $entry = $script:StdUiStateTable[$known]
    return @{ State = $known; Glyph = $entry.Glyph; Label = $entry.Label; Tone = $entry.Tone; ColorKey = (Get-StdUiToneColorKey $entry.Tone) }
}

function Get-StdUiToneColorKey {
    # Ton -> klucz koloru tekstu z palety.
    param([string]$Tone)
    switch ($Tone) {
        'Ok' { return 'ThemeOkFore' }
        'Info' { return 'ThemeInfoFore' }
        'Warn' { return 'ThemeWarnFore' }
        'Crit' { return 'ThemeCritFore' }
    }
    return 'ThemeSubtle'
}

function Get-StdUiLevelInfo {
    # Poziom logu -> @{ Level; Rank; Glyph; Label } (nieznany = Info).
    param([AllowNull()][AllowEmptyString()][string]$Level)
    foreach ($key in @('Debug', 'Info', 'Warn', 'Error')) {
        if ([string]::Equals($key, $Level, [System.StringComparison]::OrdinalIgnoreCase)) {
            $entry = $script:StdUiLogLevels[$key]
            return @{ Level = $key; Rank = $entry.Rank; Glyph = $entry.Glyph; Label = $entry.Label }
        }
    }
    $info = $script:StdUiLogLevels['Info']
    return @{ Level = 'Info'; Rank = $info.Rank; Glyph = $info.Glyph; Label = $info.Label }
}

function Get-StdUiGlyphText {
    # Znak ikony z kodu (Get-StdGlyph z motywu); pusty kod = pusty tekst.
    param([AllowNull()][AllowEmptyString()][string]$Code)
    if ([string]::IsNullOrEmpty($Code)) { return '' }
    return (Get-StdGlyph -Code $Code)
}

function Get-StdUiNumberText {
    # Liczba z jednym miejscem po przecinku w zapisie polskim (15,8), bez zbędnego ",0".
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    $culture = $null
    try { $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL') } catch { $culture = [System.Globalization.CultureInfo]::InvariantCulture }
    try { return ([double]$Value).ToString('0.#', $culture) } catch { return [string]$Value }
}

function Get-StdUiTimeText {
    # Czas wpisu dziennika HH:mm:ss (data z innego wątku też jest [datetime]).
    param([AllowNull()]$Time)
    if ($Time -is [datetime]) { return $Time.ToString('HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($null -ne $Time) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse([string]$Time, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
            return $parsed.ToLocalTime().ToString('HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return (Get-Date).ToString('HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-StdUiValue {
    # Pole słownika albo obiektu (bez błędu, gdy go nie ma).
    param([AllowNull()]$Object, [Parameter(Mandatory = $true)][string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Get-StdUiText {
    # Tekst pola albo $null, gdy pusty.
    param([AllowNull()]$Object, [Parameter(Mandatory = $true)][string]$Name)
    $value = Get-StdUiValue -Object $Object -Name $Name
    if ($null -eq $value) { return $null }
    $text = ([string]$value).Trim()
    if ($text.Length -eq 0) { return $null }
    return $text
}

# ---------------------------------------------------------------------------------------------
# Pozycje list, kroki i dziennik (bez WPF)
# ---------------------------------------------------------------------------------------------

function New-StdUiPickItem {
    <#
    .SYNOPSIS
    Pozycja listy wyboru (klient / lokalizacja / profil): [pscustomobject]@{ Id; Name; Hint }.
    Hint = "ostatnio" dla pozycji z listy ostatnio używanych.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [AllowNull()][AllowEmptyString()][string]$Name,
        [AllowNull()][AllowEmptyString()][string]$Hint
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = $Id }
    if ($null -eq $Hint) { $Hint = '' }
    return [pscustomobject]@{ Id = [string]$Id; Name = [string]$Name; Hint = [string]$Hint }
}

function Sort-StdUiByRecent {
    <#
    .SYNOPSIS
    Pozycje listy: najpierw ostatnio używane (w kolejności z -Recent, z Hint "ostatnio"), potem reszta
    w pierwotnej kolejności (bez WPF). Porównanie Id bez względu na wielkość liter.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Items,
        [AllowNull()][AllowEmptyCollection()][string[]]$Recent
    )
    $result = New-Object 'System.Collections.Generic.List[object]'
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -ne $Recent) {
        foreach ($id in $Recent) {
            if ([string]::IsNullOrEmpty($id) -or $used.Contains($id)) { continue }
            foreach ($item in $Items) {
                if ([string]::Equals([string]$item.Id, $id, [System.StringComparison]::OrdinalIgnoreCase)) {
                    [void]$used.Add([string]$item.Id)
                    $result.Add((New-StdUiPickItem -Id $item.Id -Name $item.Name -Hint 'ostatnio'))
                    break
                }
            }
        }
    }
    foreach ($item in $Items) {
        if ($used.Contains([string]$item.Id)) { continue }
        [void]$used.Add([string]$item.Id)
        $result.Add($item)
    }
    return , $result.ToArray()
}

function New-StdUiStepItem {
    <#
    .SYNOPSIS
    Wiersz listy kroków / planu (bez WPF): [pscustomobject]@{ Id; Title; State; StateLabel; Glyph; Detail; HasDetail; Phase }.
    Wszystkie pola są tekstem (wyzwalacze kolorów w XAML porównują State z napisem).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Id,
        [AllowNull()][AllowEmptyString()][string]$Title,
        [AllowNull()][AllowEmptyString()][string]$State,
        [AllowNull()][AllowEmptyString()][string]$Detail,
        [AllowNull()]$Phase
    )
    $info = Get-StdUiStateInfo -State $State
    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = $Id }
    if ($null -eq $Detail) { $Detail = '' }
    $detailText = ([string]$Detail).Trim()
    $hasDetail = 'no'
    if ($detailText.Length -gt 0) { $hasDetail = 'yes' }
    $phaseText = ''
    if ($null -ne $Phase) { $phaseText = [string]$Phase }
    return [pscustomobject]@{
        Id         = [string]$Id
        Title      = [string]$Title
        State      = [string]$info.State
        StateLabel = [string]$info.Label
        Glyph      = [string](Get-StdUiGlyphText $info.Glyph)
        Detail     = [string]$detailText
        HasDetail  = [string]$hasDetail
        Phase      = [string]$phaseText
    }
}

function New-StdUiLogItem {
    <#
    .SYNOPSIS
    Wpis dziennika na żywo (bez WPF): [pscustomobject]@{ Time; Level; Rank; Glyph; Source; Message }.
    Wiadomość jest jeszcze raz maskowana (Protect-StdLogText) i skracana do 4000 znaków.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$Time,
        [AllowNull()][AllowEmptyString()][string]$Level,
        [AllowNull()][AllowEmptyString()][string]$Source,
        [AllowNull()][AllowEmptyString()][string]$Message
    )
    $info = Get-StdUiLevelInfo -Level $Level
    if ($null -eq $Message) { $Message = '' }
    $text = Protect-StdLogText -Text ($Message.TrimEnd())
    if ($null -eq $text) { $text = '' }
    if ($text.Length -gt 4000) { $text = $text.Substring(0, 4000) + '...' }
    $sourceText = ''
    if (-not [string]::IsNullOrWhiteSpace($Source)) { $sourceText = $Source.Trim() }
    return [pscustomobject]@{
        Time    = [string](Get-StdUiTimeText $Time)
        Level   = [string]$info.Level
        Rank    = [int]$info.Rank
        Glyph   = [string](Get-StdUiGlyphText $info.Glyph)
        Source  = [string]$sourceText
        Message = [string]$text
    }
}

function Get-StdUiStepCounts {
    <#
    .SYNOPSIS
    Liczniki wierszy kroków (bez WPF): @{ Ok; Skipped; Error; Cancelled; Pending; Running; Total }.
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyCollection()][object[]]$Steps)
    $counts = @{ Ok = 0; Skipped = 0; Error = 0; Cancelled = 0; Pending = 0; Running = 0; Total = 0 }
    if ($null -eq $Steps) { return $counts }
    foreach ($step in $Steps) {
        if ($null -eq $step) { continue }
        $counts.Total = $counts.Total + 1
        $state = [string](Get-StdUiValue -Object $step -Name 'State')
        if ($counts.ContainsKey($state) -and $state -ne 'Total') { $counts[$state] = $counts[$state] + 1 }
    }
    return $counts
}

function Get-StdUiFailedStepIds {
    <#
    .SYNOPSIS
    Identyfikatory kroków do „Ponów nieudane” (bez WPF): stan Error albo Cancelled, bez powtórzeń.
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyCollection()][object[]]$Steps)
    $result = New-Object 'System.Collections.Generic.List[string]'
    if ($null -eq $Steps) { return , $result.ToArray() }
    foreach ($step in $Steps) {
        $state = [string](Get-StdUiValue -Object $step -Name 'State')
        $id = [string](Get-StdUiValue -Object $step -Name 'Id')
        if (($state -eq 'Error' -or $state -eq 'Cancelled') -and $id.Length -gt 0 -and -not $result.Contains($id)) { $result.Add($id) }
    }
    return , $result.ToArray()
}

function Get-StdUiRunOutcome {
    <#
    .SYNOPSIS
    Jak pokazać koniec wdrożenia (bez WPF): @{ View = 'Finished'|'Error'; Title; Detail; Tone; Glyph }.
    View = 'Error', gdy silnik zakończył się błędem albo któryś krok ma błąd.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$Status,
        [AllowNull()][hashtable]$Counts,
        [bool]$RebootRequired,
        [AllowNull()][AllowEmptyString()][string]$ErrorText
    )
    if ($null -eq $Counts) { $Counts = @{ Ok = 0; Skipped = 0; Error = 0; Cancelled = 0 } }
    $errors = [int]$Counts['Error']
    $cancelled = [int]$Counts['Cancelled']
    $outcome = @{ View = 'Finished'; Title = 'Wdrożenie zakończone'; Detail = 'Wszystkie kroki zakończyły się poprawnie albo zostały pominięte.'; Tone = 'Ok'; Glyph = 'E930' }
    if ($Status -eq 'Failed') {
        $outcome = @{ View = 'Error'; Title = 'Wdrożenie przerwane przez błąd'; Detail = 'Silnik wdrożenia zakończył pracę z błędem.'; Tone = 'Crit'; Glyph = 'EA39' }
        if (-not [string]::IsNullOrWhiteSpace($ErrorText)) { $outcome.Detail = 'Silnik wdrożenia zakończył pracę z błędem: ' + $ErrorText }
    } elseif ($errors -gt 0) {
        $word = 'kroków'
        if ($errors -eq 1) { $word = 'krok' } elseif ($errors -lt 5) { $word = 'kroki' }
        $outcome = @{ View = 'Error'; Title = 'Wdrożenie zakończone z błędami'; Detail = ('{0} {1} z błędem. Sprawdź szczegóły i użyj „Ponów nieudane”.' -f $errors, $word); Tone = 'Crit'; Glyph = 'EA39' }
    } elseif ($Status -eq 'Cancelled' -or $cancelled -gt 0) {
        $outcome = @{ View = 'Finished'; Title = 'Wdrożenie przerwane'; Detail = 'Wdrożenie przerwano na żądanie. Niewykonane kroki można uruchomić przyciskiem „Ponów nieudane”.'; Tone = 'Warn'; Glyph = 'E711' }
    }
    if ($RebootRequired) {
        $outcome.Detail = $outcome.Detail + ' Komputer wymaga ponownego uruchomienia - kroki drugiej fazy czekają na restart.'
    }
    return $outcome
}

# ---------------------------------------------------------------------------------------------
# Wiersze kart "Konfiguracja" i "Komputer" (bez WPF)
# ---------------------------------------------------------------------------------------------

function Get-StdUiComputerRows {
    <#
    .SYNOPSIS
    Wiersze karty „Komputer” z wyniku Get-StdComputerInfo (bez WPF): lista @{ Label; Value }.
    Brak danych = "brak danych".
    #>
    [CmdletBinding()]
    param([AllowNull()]$Info)
    $none = 'brak danych'
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $name = Get-StdUiText $Info 'ComputerName'
    $maker = Get-StdUiText $Info 'Manufacturer'
    $model = Get-StdUiText $Info 'Model'
    $modelText = (@($maker, $model) | Where-Object { $null -ne $_ }) -join ' '
    $serial = Get-StdUiText $Info 'SerialNumber'
    $os = Get-StdUiText $Info 'Os'
    $build = Get-StdUiText $Info 'OsBuild'
    if ($null -ne $os -and $null -ne $build) { $os = '{0} (kompilacja {1})' -f $os, $build } elseif ($null -eq $os -and $null -ne $build) { $os = 'kompilacja ' + $build }
    $cpu = Get-StdUiText $Info 'Cpu'
    $ram = $null
    $ramValue = Get-StdUiValue $Info 'RamGB'
    if ($null -ne $ramValue) { $ram = '{0} GB' -f (Get-StdUiNumberText $ramValue) }

    $disks = New-Object 'System.Collections.Generic.List[string]'
    foreach ($disk in @(Get-StdUiValue $Info 'Disks')) {
        if ($null -eq $disk) { continue }
        $parts = New-Object 'System.Collections.Generic.List[string]'
        $diskName = Get-StdUiText $disk 'Name'
        $size = Get-StdUiValue $disk 'SizeGB'
        $free = Get-StdUiValue $disk 'FreeGB'
        $type = Get-StdUiText $disk 'Type'
        if ($null -ne $size) { $parts.Add(('{0} GB' -f (Get-StdUiNumberText $size))) }
        if ($null -ne $type) { $parts.Add($type) }
        if ($null -ne $free) { $parts.Add(('wolne {0} GB' -f (Get-StdUiNumberText $free))) }
        $line = $parts -join ', '
        if ($null -ne $diskName) { $line = ('{0}: {1}' -f $diskName, $line).TrimEnd(' ', ':') }
        if ($line.Length -gt 0) { $disks.Add($line) }
    }

    $tpm = Get-StdUiValue $Info 'Tpm'
    $tpmText = $null
    if ($null -ne $tpm) {
        $present = Get-StdUiValue $tpm 'Present'
        $ready = Get-StdUiValue $tpm 'Ready'
        $version = Get-StdUiText $tpm 'Version'
        if ($present -eq $false) {
            $tpmText = 'brak modułu TPM'
        } elseif ($present -eq $true) {
            $tpmText = 'TPM'
            if ($null -ne $version) { $tpmText = 'TPM ' + $version }
            if ($ready -eq $true) { $tpmText += ', gotowy' } elseif ($ready -eq $false) { $tpmText += ', nie jest gotowy' }
        }
    }

    $networks = New-Object 'System.Collections.Generic.List[string]'
    foreach ($nic in @(Get-StdUiValue $Info 'Network')) {
        if ($null -eq $nic) { continue }
        $nicName = Get-StdUiText $nic 'Name'
        $ip = Get-StdUiValue $nic 'Ipv4'
        $ipText = $null
        if ($null -ne $ip) { $ipText = (@($ip) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join ', ' }
        $mac = Get-StdUiText $nic 'Mac'
        $line = $nicName
        if (-not [string]::IsNullOrEmpty($ipText)) { if ($null -ne $line) { $line = '{0}: {1}' -f $line, $ipText } else { $line = $ipText } }
        if ($null -ne $mac) { if ($null -ne $line) { $line = '{0} ({1})' -f $line, $mac } else { $line = $mac } }
        if (-not [string]::IsNullOrEmpty($line)) { $networks.Add($line) }
    }

    foreach ($pair in @(
            @('Nazwa', $name),
            @('Model', $modelText),
            @('Numer seryjny', $serial),
            @('Windows', $os),
            @('Procesor', $cpu),
            @('Pamięć RAM', $ram),
            @('Dysk', ($disks -join "`n")),
            @('TPM', $tpmText),
            @('BIOS', (Get-StdUiText $Info 'Bios')),
            @('Sieć', ($networks -join "`n"))
        )) {
        $value = [string]$pair[1]
        if ([string]::IsNullOrWhiteSpace($value)) { $value = $none }
        $rows.Add(@{ Label = $pair[0]; Value = $value })
    }
    return , $rows.ToArray()
}

function Get-StdUiSummaryRows {
    <#
    .SYNOPSIS
    Wiersze podsumowania konfiguracji wynikowej (bez WPF) z wyniku Resolve-StdEffectiveConfig: lista @{ Label; Value }.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Effective)
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $site = $Effective['Site']
    $client = [string]$Effective['ClientName']
    if ([string]::IsNullOrWhiteSpace($client)) { $client = [string]$Effective['ClientId'] }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$Effective['ClientId'])) { $client = '{0} ({1})' -f $client, $Effective['ClientId'] }
    if ($null -ne $Effective['Revision']) { $client = '{0}, rewizja {1}' -f $client, $Effective['Revision'] }
    $rows.Add(@{ Label = 'Klient'; Value = $client })

    $siteText = Get-StdUiText $site 'name'
    $siteId = Get-StdUiText $site 'id'
    if ($null -eq $siteText) { $siteText = $siteId } elseif ($null -ne $siteId) { $siteText = '{0} ({1})' -f $siteText, $siteId }
    $rows.Add(@{ Label = 'Lokalizacja'; Value = [string]$siteText })
    $rows.Add(@{ Label = 'Profil'; Value = [string]$Effective['ProfileName'] })

    $nameRule = Get-StdUiValue $site 'computerName'
    $nameText = 'bez zmiany nazwy'
    $pattern = Get-StdUiText $nameRule 'pattern'
    if ($null -ne $pattern) {
        $nameText = $pattern
        $editable = Get-StdUiValue $nameRule 'editable'
        if ($editable -eq $false) { $nameText += ' (nazwa stała)' } else { $nameText += ' (można ją zmienić przed wdrożeniem)' }
    }
    $rows.Add(@{ Label = 'Nazwa komputera'; Value = $nameText })

    $domain = Get-StdUiValue $site 'domainJoin'
    $domainName = Get-StdUiText $domain 'domain'
    $domainText = 'bez dołączania do domeny'
    if ($null -ne $domainName) {
        $domainText = $domainName
        $ou = Get-StdUiText $domain 'ou'
        if ($null -ne $ou) { $domainText = "{0}`nOU: {1}" -f $domainName, $ou }
        $domainText += "`nLogin i hasło do domeny wpisuje się ręcznie przy wdrożeniu."
    }
    $rows.Add(@{ Label = 'Domena'; Value = $domainText })

    $adminName = Get-StdUiText (Get-StdUiValue $site 'localAdmin') 'username'
    $adminText = 'nie dotyczy'
    if ($null -ne $adminName) { $adminText = '{0} (hasło wpisuje się ręcznie przy wdrożeniu)' -f $adminName }
    $rows.Add(@{ Label = 'Lokalny administrator'; Value = $adminText })

    $wifi = New-Object 'System.Collections.Generic.List[string]'
    foreach ($entry in @(Get-StdUiValue $site 'wifi')) {
        $ssid = Get-StdUiText $entry 'ssid'
        $file = Get-StdUiText $entry 'file'
        if ($null -ne $ssid) { $wifi.Add($ssid) } elseif ($null -ne $file) { $wifi.Add(('profil {0}' -f $file)) }
    }
    if ($wifi.Count -gt 0) { $rows.Add(@{ Label = 'Wi-Fi'; Value = ($wifi -join ', ') }) }

    $escrow = @(Get-StdUiValue (Get-StdUiValue $site 'bitLocker') 'escrow') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    if (@($escrow).Count -gt 0) {
        $targets = foreach ($item in $escrow) {
            switch ([string]$item) { 'ad' { 'Active Directory' } 'server' { 'serwer STD (zaszyfrowany)' } default { [string]$item } }
        }
        $rows.Add(@{ Label = 'Klucz BitLockera'; Value = (@($targets) -join ', ') })
    }

    $order = @(Get-StdUiValue $site 'sourceOrder') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    if (@($order).Count -gt 0) { $rows.Add(@{ Label = 'Źródła instalatorów'; Value = (@($order) -join ' -> ') }) }

    $scripts = @($Effective['Scripts']) | Where-Object { $null -ne $_ }
    if (@($scripts).Count -gt 0) {
        $names = foreach ($item in $scripts) { $file = Get-StdUiText $item 'file'; if ($null -ne $file) { $file } }
        $rows.Add(@{ Label = 'Skrypty'; Value = (@($names) -join ', ') })
    }

    $locked = @($Effective['Locked']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    if (@($locked).Count -gt 0) {
        $lockedNames = foreach ($item in $locked) {
            switch ([string]$item) { 'computerName' { 'nazwa komputera' } 'apps' { 'aplikacje' } default { Get-StdUiTaskTitle ([string]$item) } }
        }
        $rows.Add(@{ Label = 'Zablokowane w profilu'; Value = (@($lockedNames) -join ', ') })
    }
    return , $rows.ToArray()
}

function Get-StdUiAppChips {
    <#
    .SYNOPSIS
    Aplikacje profilu jako pozycje „chipów” (bez WPF): lista [pscustomobject]@{ Name; Tip }.
    Nazwa z katalogu (catalog.<klucz>.name), podpowiedź: skąd instalator.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Effective)
    $result = New-Object 'System.Collections.Generic.List[object]'
    $catalog = $Effective['Catalog']
    foreach ($key in @($Effective['Apps'])) {
        if ([string]::IsNullOrWhiteSpace([string]$key)) { continue }
        $entry = Get-StdUiValue $catalog ([string]$key)
        $name = Get-StdUiText $entry 'name'
        if ($null -eq $name) { $name = [string]$key }
        $tips = New-Object 'System.Collections.Generic.List[string]'
        $file = Get-StdUiText $entry 'file'
        $winget = Get-StdUiText $entry 'winget'
        $url = Get-StdUiText $entry 'url'
        if ($null -ne $file) { $tips.Add(('plik {0}' -f $file)) }
        if ($null -ne $url) { $tips.Add('pobieranie z adresu https') }
        if ($null -ne $winget) { $tips.Add(('winget {0}' -f $winget)) }
        $tip = '{0} ({1})' -f $name, $key
        if ($tips.Count -gt 0) { $tip = '{0}: {1}' -f $tip, ($tips -join ', ') }
        $result.Add([pscustomobject]@{ Name = [string]$name; Tip = [string]$tip })
    }
    return , $result.ToArray()
}

function Get-StdUiTaskChips {
    <#
    .SYNOPSIS
    Wszystkie zadania z SPEC 2.8 jako „chipy” (bez WPF): [pscustomobject]@{ Name; Mark ('on'|'off'); Tip }.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Effective)
    $result = New-Object 'System.Collections.Generic.List[object]'
    $tasks = $Effective['Tasks']
    foreach ($id in $script:StdUiTaskTitles.Keys) {
        $value = Get-StdUiValue $tasks $id
        $mark = 'off'
        $tip = 'Wyłączone w profilu'
        if ($value -is [bool] -and $value) { $mark = 'on'; $tip = 'Włączone w profilu' }
        $result.Add([pscustomobject]@{ Name = [string]$script:StdUiTaskTitles[$id]; Mark = $mark; Tip = ('{0} ({1})' -f $tip, $id) })
    }
    return , $result.ToArray()
}

# ---------------------------------------------------------------------------------------------
# Ostatnio używane (prefs.json, bez WPF)
# ---------------------------------------------------------------------------------------------

function Get-StdUiPrefs {
    <#
    .SYNOPSIS
    Zawartość prefs.json jako [ordered] słownik (brak albo uszkodzony plik = pusty słownik).
    #>
    [CmdletBinding()]
    param()
    $prefs = $null
    try { $prefs = Read-StdJsonFile -Path (Get-StdPath -Kind Prefs) } catch { $prefs = $null }
    if ($prefs -isnot [System.Collections.IDictionary]) { $prefs = [ordered]@{} }
    return $prefs
}

function Get-StdUiRecentIds {
    <#
    .SYNOPSIS
    Ostatnio używane identyfikatory (bez WPF): -Kind clients | sites (-ClientId) | profiles (-ClientId, -SiteId).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('clients', 'sites', 'profiles')][string]$Kind,
        [string]$ClientId,
        [string]$SiteId,
        [System.Collections.IDictionary]$Prefs
    )
    if ($null -eq $Prefs) { $Prefs = Get-StdUiPrefs }
    $recent = Get-StdUiValue $Prefs 'recent'
    $list = $null
    switch ($Kind) {
        'clients' { $list = Get-StdUiValue $recent 'clients' }
        'sites' { $list = Get-StdUiValue (Get-StdUiValue $recent 'sites') ([string]$ClientId) }
        'profiles' { $list = Get-StdUiValue (Get-StdUiValue $recent 'profiles') ('{0}/{1}' -f $ClientId, $SiteId) }
    }
    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in @($list)) {
        if ($item -is [string] -and $item.Length -gt 0 -and -not $result.Contains($item)) { $result.Add($item) }
    }
    return , $result.ToArray()
}

function Add-StdUiRecentValue {
    # Wartość na początek listy (bez powtórzeń), najwyżej $script:StdUiRecentMax pozycji.
    param([AllowNull()]$List, [Parameter(Mandatory = $true)][string]$Value)
    $result = New-Object 'System.Collections.Generic.List[string]'
    $result.Add($Value)
    foreach ($item in @($List)) {
        if ($result.Count -ge $script:StdUiRecentMax) { break }
        if ($item -is [string] -and $item.Length -gt 0 -and -not [string]::Equals($item, $Value, [System.StringComparison]::OrdinalIgnoreCase)) { $result.Add($item) }
    }
    return , $result.ToArray()
}

function Save-StdUiRecent {
    <#
    .SYNOPSIS
    Zapisuje wybór technika w prefs.json (bez WPF): lastClient, lastSite, lastProfile i listy „recent”.
    Pozostałe pola pliku (np. motyw) zostają. Błąd zapisu trafia tylko do logu.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ClientId,
        [string]$SiteId,
        [string]$ProfileName
    )
    try {
        $path = Get-StdPath -Kind Prefs
        $prefs = Get-StdUiPrefs
        $recent = Get-StdUiValue $prefs 'recent'
        if ($recent -isnot [System.Collections.IDictionary]) { $recent = [ordered]@{} }
        $recent['clients'] = Add-StdUiRecentValue -List (Get-StdUiValue $recent 'clients') -Value $ClientId
        $prefs['lastClient'] = $ClientId
        if (-not [string]::IsNullOrEmpty($SiteId)) {
            $sites = Get-StdUiValue $recent 'sites'
            if ($sites -isnot [System.Collections.IDictionary]) { $sites = [ordered]@{} }
            $sites[$ClientId] = Add-StdUiRecentValue -List (Get-StdUiValue $sites $ClientId) -Value $SiteId
            $recent['sites'] = $sites
            $prefs['lastSite'] = $SiteId
            if (-not [string]::IsNullOrEmpty($ProfileName)) {
                $profiles = Get-StdUiValue $recent 'profiles'
                if ($profiles -isnot [System.Collections.IDictionary]) { $profiles = [ordered]@{} }
                $key = '{0}/{1}' -f $ClientId, $SiteId
                $profiles[$key] = Add-StdUiRecentValue -List (Get-StdUiValue $profiles $key) -Value $ProfileName
                $recent['profiles'] = $profiles
                $prefs['lastProfile'] = $ProfileName
            }
        }
        $prefs['recent'] = $recent
        Write-StdJsonFile -Path $path -InputObject $prefs
    } catch {
        Write-StdLog -Level Warn -Source 'UI' -Message ('Nie udało się zapisać ostatniego wyboru w prefs.json: {0}' -f $_.Exception.Message)
    }
}

function Test-StdUiOfflineIndex {
    <#
    .SYNOPSIS
    $true, gdy jest kopia listy klientów do pracy offline: Cache\clients.sdt albo <OriginRoot>\offline\clients.sdt (bez WPF).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$OriginRoot)
    try {
        if ([System.IO.File]::Exists((Join-Path (Get-StdPath -Kind Cache) 'clients.sdt'))) { return $true }
    } catch { }
    if (-not [string]::IsNullOrWhiteSpace($OriginRoot)) {
        try {
            if ([System.IO.File]::Exists((Join-Path (Join-Path $OriginRoot 'offline') 'clients.sdt'))) { return $true }
        } catch { }
    }
    return $false
}

function Get-StdUiOfflineBootstrap {
    <#
    .SYNOPSIS
    Plik startowy dla pracy offline (bez sesji): tylko kopie lustrzane, bez adresu serwera STD, który nie
    odpowiadał. Zwraca @{ mirrors } albo $null, gdy kopii nie ma (bez WPF).
    #>
    [CmdletBinding()]
    param([AllowNull()]$Bootstrap)
    if ($null -eq $Bootstrap) { return $null }
    $mirrors = @(@(Get-StdUiValue $Bootstrap 'mirrors') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($mirrors.Count -eq 0) { return $null }
    return @{ mirrors = $mirrors }
}

# ---------------------------------------------------------------------------------------------
# Błędy (bez WPF)
# ---------------------------------------------------------------------------------------------

function Get-StdUiErrorMessage {
    <#
    .SYNOPSIS
    Tekst błędu do pokazania w oknie (bez WPF): komunikat wyjątku bez opakowania PowerShella, z zachowanymi
    końcami linii (lista prób pobrania), zamaskowany i skrócony.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$ErrorObject, [int]$MaxLength = 2000)
    $text = ''
    $ex = $null
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ex = $ErrorObject.Exception }
    elseif ($ErrorObject -is [System.Exception]) { $ex = $ErrorObject }
    if ($null -ne $ex) {
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $text = [string]$ex.Message
    } elseif ($null -ne $ErrorObject) {
        $text = [string]$ErrorObject
    }
    $lines = foreach ($line in ($text -replace "`r", '') -split "`n") {
        $clean = ($line -replace '[ \t]+', ' ').TrimEnd()
        if ($clean.Trim().Length -gt 0) { $clean }
    }
    $text = (@($lines) -join "`n").Trim()
    if ($text.Length -gt $MaxLength) { $text = $text.Substring(0, $MaxLength) + '...' }
    return (Protect-StdLogText -Text $text)
}

function Get-StdUiServerErrorText {
    <#
    .SYNOPSIS
    Zrozumiały opis błędu połączenia z serwerem STD (bez WPF) na podstawie kodu [STD-...] w tekście próby.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return 'Nie udało się połączyć z serwerem STD.' }
    $detail = ($Text -replace '^\s*\[[A-Z0-9-]+\]\s*', '').Trim()
    if ($Text -match '\[STD-HTTP-401\]') { return 'Nieprawidłowy login lub hasło.' }
    if ($Text -match '\[STD-HTTP-403\]') { return 'To konto nie ma dostępu do serwera STD. Poproś administratora o dodanie konta do grupy techników.' }
    if ($Text -match '\[STD-HTTP-404\]') { return 'Serwer nie ma listy klientów STD. Sprawdź adres serwera (powinien kończyć się na /std/).' }
    if ($Text -match '\[STD-HTTP-(3\d\d)\]') { return 'Serwer przekierowuje żądanie. Sprawdź adres serwera STD: ' + $detail }
    if ($Text -match '\[STD-HTTP-5\d\d\]') { return 'Błąd po stronie serwera STD. Spróbuj ponownie za chwilę albo zgłoś to administratorowi.' }
    if ($Text -match '\[STD-HTTP-NET\]') { return 'Brak połączenia z serwerem STD: ' + $detail }
    if ($Text -match '\[STD-HTTP-INSECURE\]') { return 'Adres serwera musi zaczynać się od https://.' }
    if ($Text -match '\[STD-SIG-[A-Z]+\]') { return 'Lista klientów z serwera ma nieprawidłowy podpis: ' + $detail }
    if ($Text -match '\[STD-CFG-[A-Z-]+\]') { return 'Lista klientów z serwera jest nieprawidłowa: ' + $detail }
    return $detail
}

function Get-StdUiLoginFailure {
    <#
    .SYNOPSIS
    Komunikat o nieudanym logowaniu (bez WPF): z listy prób Get-StdClientIndex (wynik albo TargetObject
    błędu [STD-CFG-UNAVAILABLE]) bierze próbę serwera. Zwraca @{ Message; OfflineSource }; OfflineSource =
    źródło kopii (mirror / cache / usb), z której można pracować offline, albo $null.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$IndexResult,
        [AllowNull()]$ErrorObject
    )
    $attempts = @()
    $offline = $null
    if ($null -ne $IndexResult) {
        $attempts = @(Get-StdUiValue $IndexResult 'Attempts')
        $source = [string](Get-StdUiValue $IndexResult 'Source')
        if ($source -ne 'server' -and $source.Length -gt 0) { $offline = $source }
    } elseif ($ErrorObject -is [System.Management.Automation.ErrorRecord]) {
        $target = $ErrorObject.TargetObject
        if ($null -ne $target) { $attempts = @($target) }
    }
    $serverAttempt = $null
    foreach ($attempt in $attempts) {
        if ([string](Get-StdUiValue $attempt 'Source') -eq 'server') { $serverAttempt = $attempt; break }
    }
    if ($null -ne $serverAttempt) {
        $message = Get-StdUiServerErrorText -Text ([string](Get-StdUiValue $serverAttempt 'Error'))
    } elseif ($null -ne $ErrorObject) {
        $raw = Get-StdUiErrorMessage -ErrorObject $ErrorObject -MaxLength 600
        if ($raw -match '\[STD-CFG-UNAVAILABLE\]') {
            $message = 'Nie udało się pobrać listy klientów z żadnego źródła.'
        } else {
            $message = ($raw -replace '^\s*\[[A-Z0-9-]+\]\s*', '').Trim()
        }
    } else {
        $message = 'Brak adresu serwera STD - nie można się zalogować.'
    }
    return @{ Message = Protect-StdLogText -Text $message; OfflineSource = $offline }
}

# ---------------------------------------------------------------------------------------------
# Wi-Fi: lista sieci z netsh, profil XML, uruchamianie netsh (bez WPF; działa w zadaniu w tle)
# ---------------------------------------------------------------------------------------------

function ConvertFrom-StdNetshNetworks {
    <#
    .SYNOPSIS
    Sieci Wi-Fi z wyniku "netsh wlan show networks mode=bssid" (bez WPF; angielski, polski i niemiecki Windows):
    lista [pscustomobject]@{ Ssid; Auth ('WPA2PSK'|'WPA3SAE'|'open'|'enterprise'|'unknown'); AuthText; AuthLabel;
    Encryption; Signal (0-100 albo -1); SignalText; Secured }, posortowana od najsilniejszego sygnału.
    Ukryte sieci (pusty SSID) są pomijane, powtórzenia SSID łączone.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][AllowEmptyString()][AllowNull()][string]$Text)
    $networks = New-Object 'System.Collections.Generic.List[hashtable]'
    if ([string]::IsNullOrEmpty($Text)) { return , @() }
    $current = $null
    foreach ($raw in ($Text -replace "`r", '') -split "`n") {
        $ssidMatch = [regex]::Match($raw, '^\s*SSID\s+\d+\s*:\s?(.*)$')
        if ($ssidMatch.Success) {
            $current = @{ Ssid = $ssidMatch.Groups[1].Value.TrimEnd(); AuthText = ''; Encryption = ''; Signal = -1 }
            $networks.Add($current)
            continue
        }
        if ($null -eq $current) { continue }
        if ($raw -match '^\s*BSSID\s+\d+') { continue }
        $signalMatch = [regex]::Match($raw, ':\s*(\d{1,3})\s*%\s*$')
        if ($signalMatch.Success) {
            $value = [int]$signalMatch.Groups[1].Value
            if ($value -gt 100) { $value = 100 }
            if ($value -gt $current.Signal) { $current.Signal = $value }
            continue
        }
        $pair = [regex]::Match($raw, '^\s*([^:]+?)\s*:\s*(.+?)\s*$')
        if (-not $pair.Success) { continue }
        $label = $pair.Groups[1].Value
        $value = $pair.Groups[2].Value
        if ($label -match '(?i)authenti|uwierzyteln|authentifiz|authentification') { $current.AuthText = $value }
        elseif ($label -match '(?i)encrypt|szyfrow|verschl|chiffr') { $current.Encryption = $value }
    }
    $byName = [ordered]@{}
    foreach ($net in $networks) {
        if ([string]::IsNullOrWhiteSpace($net.Ssid)) { continue }
        $key = $net.Ssid
        if ($byName.Contains($key)) {
            if ($net.Signal -gt $byName[$key].Signal) { $byName[$key].Signal = $net.Signal }
            continue
        }
        $byName[$key] = $net
    }
    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach ($net in $byName.Values) {
        $authText = [string]$net.AuthText
        $auth = 'unknown'
        if ($authText -match '(?i)enterprise|802\.1x') { $auth = 'enterprise' }
        elseif ($authText -match '(?i)WPA3') { $auth = 'WPA3SAE' }
        elseif ($authText -match '(?i)WPA2?-?Personal|WPA2?PSK|WPA2?-PSK') { $auth = 'WPA2PSK' }
        elseif ($authText -match '(?i)^\s*(open|otwart|offen|ouvert|abiert)') { $auth = 'open' }
        $authLabel = switch ($auth) {
            'WPA2PSK' { 'WPA2' }
            'WPA3SAE' { 'WPA3' }
            'open' { 'otwarta' }
            'enterprise' { 'firmowa (802.1X)' }
            default { if ($authText.Length -gt 0) { $authText } else { 'nieznane zabezpieczenia' } }
        }
        $signalText = ''
        if ($net.Signal -ge 0) { $signalText = '{0}%' -f $net.Signal }
        $result.Add([pscustomobject]@{
                Ssid       = [string]$net.Ssid
                Auth       = [string]$auth
                AuthText   = [string]$authText
                AuthLabel  = [string]$authLabel
                Encryption = [string]$net.Encryption
                Signal     = [int]$net.Signal
                SignalText = [string]$signalText
                Secured    = [bool]($auth -ne 'open')
            })
    }
    $sorted = @($result.ToArray() | Sort-Object -Property @{ Expression = { $_.Signal }; Descending = $true }, @{ Expression = { $_.Ssid }; Descending = $false })
    return , $sorted
}

function New-StdWifiProfileXml {
    <#
    .SYNOPSIS
    Profil sieci Wi-Fi (WLANProfile v1) dla "netsh wlan add profile" (bez WPF). -Auth: WPA2PSK, WPA3SAE albo open.
    Hasło (8-63 znaków albo 64 znaki HEX) tylko dla WPA2PSK / WPA3SAE. Błąd: [STD-UI-WIFI].
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][string]$Ssid,
        [Parameter(Mandatory = $true)][ValidateSet('WPA2PSK', 'WPA3SAE', 'open')][string]$Auth,
        [AllowNull()][AllowEmptyString()][string]$Key
    )
    if ([string]::IsNullOrEmpty($Ssid) -or [System.Text.Encoding]::UTF8.GetByteCount($Ssid) -gt 32) {
        throw '[STD-UI-WIFI] Nazwa sieci Wi-Fi (SSID) musi mieć od 1 do 32 bajtów.'
    }
    $secured = ($Auth -ne 'open')
    if ($secured) {
        if ([string]::IsNullOrEmpty($Key)) { throw '[STD-UI-WIFI] Podaj hasło sieci Wi-Fi.' }
        $isHex = ($Key.Length -eq 64 -and $Key -match '^[0-9A-Fa-f]{64}$')
        if (-not $isHex -and ($Key.Length -lt 8 -or $Key.Length -gt 63)) { throw '[STD-UI-WIFI] Hasło sieci Wi-Fi musi mieć od 8 do 63 znaków.' }
    }
    $hex = ([System.BitConverter]::ToString([System.Text.Encoding]::UTF8.GetBytes($Ssid)) -replace '-', '')
    $name = [System.Security.SecurityElement]::Escape($Ssid)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0"?>')
    [void]$sb.AppendLine('<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">')
    [void]$sb.AppendLine(('    <name>{0}</name>' -f $name))
    [void]$sb.AppendLine(('    <SSIDConfig><SSID><hex>{0}</hex><name>{1}</name></SSID></SSIDConfig>' -f $hex, $name))
    [void]$sb.AppendLine('    <connectionType>ESS</connectionType>')
    [void]$sb.AppendLine('    <connectionMode>auto</connectionMode>')
    [void]$sb.AppendLine('    <MSM>')
    [void]$sb.AppendLine('        <security>')
    if ($secured) {
        [void]$sb.AppendLine(('            <authEncryption><authentication>{0}</authentication><encryption>AES</encryption><useOneX>false</useOneX></authEncryption>' -f $Auth))
        $keyType = 'passPhrase'
        if ($Key.Length -eq 64 -and $Key -match '^[0-9A-Fa-f]{64}$') { $keyType = 'networkKey' }
        [void]$sb.AppendLine(('            <sharedKey><keyType>{0}</keyType><protected>false</protected><keyMaterial>{1}</keyMaterial></sharedKey>' -f $keyType, [System.Security.SecurityElement]::Escape($Key)))
    } else {
        [void]$sb.AppendLine('            <authEncryption><authentication>open</authentication><encryption>none</encryption><useOneX>false</useOneX></authEncryption>')
    }
    [void]$sb.AppendLine('        </security>')
    [void]$sb.AppendLine('    </MSM>')
    [void]$sb.AppendLine('</WLANProfile>')
    return $sb.ToString()
}

function Invoke-StdUiNetsh {
    <#
    .SYNOPSIS
    Uruchamia netsh.exe z podanym wierszem argumentów (bez okna) i zwraca @{ ExitCode; Output }.
    Wyjście czytane w stronie kodowej OEM konsoli (polskie znaki). Limit czasu: [STD-UI-NETSH].
    Tylko Windows; wołana w zadaniu w tle (Start-StdUiTask), nigdy w wątku okna.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ArgumentLine,
        [ValidateRange(1, 300)][int]$TimeoutSec = 30
    )
    $exe = 'netsh.exe'
    if (-not [string]::IsNullOrEmpty($env:SystemRoot)) {
        $candidate = Join-Path $env:SystemRoot 'System32\netsh.exe'
        if ([System.IO.File]::Exists($candidate)) { $exe = $candidate }
    }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe
    $psi.Arguments = $ArgumentLine
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    try {
        $oem = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
        $psi.StandardOutputEncoding = $oem
        $psi.StandardErrorEncoding = $oem
    } catch { }
    $process = $null
    try {
        $process = [System.Diagnostics.Process]::Start($psi)
    } catch {
        throw ('[STD-UI-NETSH] Nie udało się uruchomić netsh: {0}' -f $_.Exception.Message)
    }
    try {
        $outTask = $process.StandardOutput.ReadToEndAsync()
        $errTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSec * 1000)) {
            try { $process.Kill() } catch { }
            throw ('[STD-UI-NETSH] netsh nie odpowiedział w ciągu {0} s.' -f $TimeoutSec)
        }
        $process.WaitForExit()
        $output = ([string]$outTask.Result + [string]$errTask.Result).Trim()
        return @{ ExitCode = [int]$process.ExitCode; Output = $output }
    } finally {
        try { $process.Dispose() } catch { }
    }
}

function Get-StdUiFunctionText {
    <#
    .SYNOPSIS
    Tekst definicji podanych funkcji (function Nazwa { ... }) do wklejenia w skrypt zadania w tle -
    runspace zadania ma tylko moduł STD, bez funkcji okien.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)][string[]]$Name)
    $sb = New-Object System.Text.StringBuilder
    foreach ($item in $Name) {
        $command = Get-Command -Name $item -CommandType Function -ErrorAction Stop | Select-Object -First 1
        [void]$sb.AppendLine(('function {0} {{' -f $item))
        [void]$sb.AppendLine($command.ScriptBlock.ToString())
        [void]$sb.AppendLine('}')
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------------------------
# Podsumowanie wdrożenia (bez WPF)
# ---------------------------------------------------------------------------------------------

function Format-StdUiRunSummary {
    <#
    .SYNOPSIS
    Tekst „Kopiuj podsumowanie” (bez WPF). -Data: @{ Version; DeploymentId; ClientName; ClientId; Revision;
    SiteName; SiteId; ProfileName; Computer (wynik Get-StdComputerInfo); Status; RebootRequired; Error;
    StartedAt; FinishedAt; LogPath; Steps (wiersze z Title, State, StateLabel, Detail) }. Tekst jest zamaskowany.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Data)
    $nl = [System.Environment]::NewLine
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $version = [string]$Data['Version']
    $lines.Add(('STD {0} - podsumowanie wdrożenia' -f $version).Replace('STD  -', 'STD -'))
    if (-not [string]::IsNullOrEmpty([string]$Data['DeploymentId'])) { $lines.Add(('Wdrożenie: {0}' -f $Data['DeploymentId'])) }
    $client = [string]$Data['ClientName']
    if ([string]::IsNullOrEmpty($client)) { $client = [string]$Data['ClientId'] } elseif (-not [string]::IsNullOrEmpty([string]$Data['ClientId'])) { $client = '{0} ({1})' -f $client, $Data['ClientId'] }
    if ($null -ne $Data['Revision']) { $client = '{0}, rewizja {1}' -f $client, $Data['Revision'] }
    if (-not [string]::IsNullOrEmpty($client)) { $lines.Add(('Klient: {0}' -f $client)) }
    $site = [string]$Data['SiteName']
    if ([string]::IsNullOrEmpty($site)) { $site = [string]$Data['SiteId'] } elseif (-not [string]::IsNullOrEmpty([string]$Data['SiteId'])) { $site = '{0} ({1})' -f $site, $Data['SiteId'] }
    if (-not [string]::IsNullOrEmpty($site)) { $lines.Add(('Lokalizacja: {0}' -f $site)) }
    if (-not [string]::IsNullOrEmpty([string]$Data['ProfileName'])) { $lines.Add(('Profil: {0}' -f $Data['ProfileName'])) }
    $computer = $Data['Computer']
    if ($null -ne $computer) {
        $parts = New-Object 'System.Collections.Generic.List[string]'
        foreach ($field in @('ComputerName', 'Manufacturer', 'Model')) { $v = Get-StdUiText $computer $field; if ($null -ne $v) { $parts.Add($v) } }
        $serial = Get-StdUiText $computer 'SerialNumber'
        if ($null -ne $serial) { $parts.Add(('nr seryjny {0}' -f $serial)) }
        if ($parts.Count -gt 0) { $lines.Add(('Komputer: {0}' -f ($parts -join ', '))) }
    }
    foreach ($pair in @(@('Start', 'StartedAt'), @('Koniec', 'FinishedAt'))) {
        $time = $Data[$pair[1]]
        if ($time -is [datetime]) { $lines.Add(('{0}: {1}' -f $pair[0], $time.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture))) }
    }
    $steps = @($Data['Steps']) | Where-Object { $null -ne $_ }
    $counts = Get-StdUiStepCounts -Steps $steps
    $outcome = Get-StdUiRunOutcome -Status ([string]$Data['Status']) -Counts $counts -RebootRequired ([bool]$Data['RebootRequired']) -ErrorText ([string]$Data['Error'])
    $lines.Add(('Wynik: {0}. OK: {1}, pominięto: {2}, błąd: {3}, przerwano: {4}.' -f $outcome.Title, $counts.Ok, $counts.Skipped, $counts.Error, $counts.Cancelled))
    if ([bool]$Data['RebootRequired']) { $lines.Add('Wymagany restart komputera.') }
    if (-not [string]::IsNullOrWhiteSpace([string]$Data['Error'])) { $lines.Add(('Błąd silnika: {0}' -f $Data['Error'])) }
    if (-not [string]::IsNullOrEmpty([string]$Data['LogPath'])) { $lines.Add(('Log: {0}' -f $Data['LogPath'])) }
    if (@($steps).Count -gt 0) {
        $lines.Add('')
        $lines.Add('Kroki:')
        foreach ($step in $steps) {
            $label = Get-StdUiText $step 'StateLabel'
            if ($null -eq $label) { $label = (Get-StdUiStateInfo -State ([string](Get-StdUiValue $step 'State'))).Label }
            $title = Get-StdUiText $step 'Title'
            if ($null -eq $title) { $title = [string](Get-StdUiValue $step 'Id') }
            $line = '[{0}] {1}' -f $label, $title
            $detail = Get-StdUiText $step 'Detail'
            if ($null -ne $detail) { $line = '{0} - {1}' -f $line, (($detail -replace '\s+', ' ').Trim()) }
            $lines.Add($line)
        }
    }
    return (Protect-StdLogText -Text (($lines -join $nl) + $nl))
}

# ---------------------------------------------------------------------------------------------
# Przykładowe dane dla podglądu okien (bez WPF)
# ---------------------------------------------------------------------------------------------

function Get-StdUiPreviewSample {
    <#
    .SYNOPSIS
    Przykładowe dane okien (bez WPF, bez serwera) dla -PreviewData: lista klientów, konfiguracja klienta,
    informacje o komputerze, plan, kroki i dziennik dla stanu -State, sieci Wi-Fi i dane logowania.
    Klucze, które podgląd poda sam w -PreviewData, mają pierwszeństwo (Merge-StdUiPreviewData).
    #>
    [CmdletBinding()]
    param([ValidateSet('Select', 'Plan', 'Running', 'Paused', 'Error', 'Finished')][string]$State = 'Select')
    $release = [ordered]@{
        schemaVersion = 4; kind = 'std.release'; revision = 42
        client        = [ordered]@{ id = 'lizard'; name = 'Lizard-Tech' }
        sources       = [ordered]@{
            lan = [ordered]@{ type = 'unc'; path = '\\srv-files\STD$\std\files\lizard\'; auth = 'session' }
            web = [ordered]@{ type = 'https'; url = 'https://pobieranie.example.com/std/files/lizard/'; auth = 'session' }
        }
        catalog       = [ordered]@{
            '7zip'     = [ordered]@{ name = '7-Zip'; file = '7z2408-x64.exe'; args = '/S' }
            Chrome     = [ordered]@{ name = 'Google Chrome'; winget = 'Google.Chrome' }
            TeamViewer = [ordered]@{ name = 'TeamViewer Host'; file = 'TeamViewer_Host.msi' }
            AV         = [ordered]@{ name = 'Antywirus'; file = 'eset.msi' }
            Office     = [ordered]@{ name = 'Microsoft 365 Apps dla firm (wersja 64-bitowa, kanał bieżący)'; file = 'setup.exe' }
        }
        defaults      = [ordered]@{
            tasks     = [ordered]@{ WaitForNetwork = $true; SuspendSleep = $true; RemoveBloatware = $true; InstallApps = $true; CreateLocalAdmin = $true; RenameComputer = $true; JoinDomain = $true; SystemSettings = $true }
            bloatware = @('*TikTok*', '*Xbox*')
            scripts   = @([ordered]@{ file = 'Drukarki.ps1' })
        }
        sites         = @(
            [ordered]@{
                id = 'krakow'; name = 'Kraków - biuro przy ulicy Długiej'; sourceOrder = @('lan', 'web')
                wifi = @([ordered]@{ ssid = 'Firma'; auth = 'WPA2PSK'; keyRef = 'wifi-firma' })
                computerName = [ordered]@{ pattern = 'KRK-{serial}'; maxLength = 15; editable = $true }
                domainJoin = [ordered]@{ domain = 'firma.local'; ou = 'OU=Komputery,OU=Krakow,DC=firma,DC=local' }
                localAdmin = [ordered]@{ username = 'LokalnyIT' }
                bitLocker = [ordered]@{ escrow = @('ad', 'server') }
                profiles = [ordered]@{
                    Standard    = [ordered]@{ apps = @('TeamViewer', 'AV', 'Chrome', '7zip', 'Office'); locked = @('JoinDomain') }
                    'Księgowość' = [ordered]@{ apps = @('TeamViewer', 'AV', 'Chrome'); tasks = [ordered]@{ EnableBitLocker = $true } }
                }
            },
            [ordered]@{
                id = 'warszawa'; name = 'Warszawa'; sourceOrder = @('web')
                computerName = [ordered]@{ pattern = 'WAW-{serial}'; maxLength = 15; editable = $false }
                domainJoin = [ordered]@{ domain = 'firma.local' }
                localAdmin = [ordered]@{ username = 'LokalnyIT' }
                profiles = [ordered]@{ Standard = [ordered]@{ apps = @('Chrome', '7zip') } }
            }
        )
    }
    $index = [ordered]@{
        schemaVersion = 4; kind = 'std.clients'; revision = 7
        clients       = @(
            [ordered]@{ id = 'lizard'; name = 'Lizard-Tech'; release = 'clients/lizard/release.sdt' },
            [ordered]@{ id = 'kowalski-i-wspolnicy'; name = 'Kancelaria Kowalski i Wspólnicy sp.k.'; release = 'clients/kowalski-i-wspolnicy/release.sdt' },
            [ordered]@{ id = 'piekarnia'; name = 'Piekarnia Złoty Kłos'; release = 'clients/piekarnia/release.sdt' }
        )
    }
    $computer = @{
        ComputerName = 'DESKTOP-7Q2M4KD'; Manufacturer = 'Dell Inc.'; Model = 'Latitude 5440'; SerialNumber = '8H2KXY3'
        Bios = 'Dell Inc. 1.14.1 (2026-03-12)'; Os = 'Microsoft Windows 11 Pro 24H2'; OsBuild = '26100.2314'
        Cpu = '13th Gen Intel(R) Core(TM) i5-1345U'; RamGB = 15.7
        Disks = @(@{ Name = 'C:'; SizeGB = 476.3; FreeGB = 431.9; Type = 'SSD NVMe' })
        Tpm = @{ Present = $true; Ready = $true; Version = '2.0' }
        Network = @(@{ Name = 'Ethernet'; Ipv4 = @('192.168.10.57'); Mac = '00-1A-2B-3C-4D-5E' })
    }
    $plan = @(
        @{ Id = 'Network.WaitForNetwork'; Title = 'Czekanie na sieć'; Phase = 1; Enabled = $true; State = 'Done'; Detail = 'Adres 192.168.10.57, serwer pobieranie.example.com odpowiada na porcie 443.' },
        @{ Id = 'Power.SuspendSleep'; Title = 'Wstrzymanie usypiania'; Phase = 1; Enabled = $true; State = 'Todo'; Detail = '' },
        @{ Id = 'Apps.Install'; Title = 'Instalacja aplikacji'; Phase = 1; Enabled = $true; State = 'Todo'; Detail = '5 aplikacji: TeamViewer Host, Antywirus, Google Chrome, 7-Zip, Microsoft 365 Apps.' },
        @{ Id = 'Domain.Join'; Title = 'Dołączenie do domeny'; Phase = 1; Enabled = $true; State = 'Blocked'; Detail = 'Kontroler domeny firma.local nie odpowiada (sprawdź połączenie z siecią firmową).' },
        @{ Id = 'Security.BitLocker'; Title = 'Włączenie BitLockera'; Phase = 2; Enabled = $false; State = 'Disabled'; Detail = 'Zadanie EnableBitLocker jest wyłączone.' }
    )
    $steps = @()
    $progress = @{ Percent = 0; Text = '' }
    $result = $null
    $log = New-Object 'System.Collections.Generic.List[object]'
    $base = (Get-Date).Date.AddHours(10).AddMinutes(15)
    $log.Add(@{ Time = $base; Level = 'Info'; Source = 'Std.Engine'; Message = 'STD 4.0.0: wdrożenie 20261102-101500-ab12 - klient lizard, lokalizacja krakow, profil Standard, rewizja 42.' })
    $log.Add(@{ Time = $base.AddSeconds(1); Level = 'Info'; Source = 'Std.Engine'; Message = 'Kroki do wykonania: 4.' })
    $log.Add(@{ Time = $base.AddSeconds(2); Level = 'Info'; Source = 'Network.WaitForNetwork'; Message = 'Wynik kroku Czekanie na sieć: Skipped - Już zrobione: adres 192.168.10.57.' })
    $log.Add(@{ Time = $base.AddSeconds(3); Level = 'Info'; Source = 'Power.SuspendSleep'; Message = 'Wynik kroku Wstrzymanie usypiania: Ok - usypianie wstrzymane do końca wdrożenia.' })
    switch ($State) {
        'Running' {
            $steps = @(
                @{ Id = 'Network.WaitForNetwork'; Title = 'Czekanie na sieć'; State = 'Skipped'; Detail = 'Już zrobione: adres 192.168.10.57.' },
                @{ Id = 'Power.SuspendSleep'; Title = 'Wstrzymanie usypiania'; State = 'Ok'; Detail = 'Usypianie wstrzymane do końca wdrożenia.' },
                @{ Id = 'Apps.Install'; Title = 'Instalacja aplikacji'; State = 'Running'; Detail = '' },
                @{ Id = 'Domain.Join'; Title = 'Dołączenie do domeny'; State = 'Pending'; Detail = '' }
            )
            $progress = @{ Percent = 50; Text = 'Krok 3 z 4: Instalacja aplikacji' }
            $log.Add(@{ Time = $base.AddSeconds(4); Level = 'Info'; Source = 'Std.Engine'; Message = 'Krok 3 z 4: Instalacja aplikacji (Apps.Install).' })
            $log.Add(@{ Time = $base.AddSeconds(9); Level = 'Warn'; Source = 'Invoke-StdProcess'; Message = 'TeamViewer_Host.msi: kod wyjścia 3010 po 41,2 s (wymaga restartu).' })
        }
        'Paused' {
            $steps = @(
                @{ Id = 'Network.WaitForNetwork'; Title = 'Czekanie na sieć'; State = 'Skipped'; Detail = 'Już zrobione: adres 192.168.10.57.' },
                @{ Id = 'Power.SuspendSleep'; Title = 'Wstrzymanie usypiania'; State = 'Ok'; Detail = 'Usypianie wstrzymane do końca wdrożenia.' },
                @{ Id = 'Apps.Install'; Title = 'Instalacja aplikacji'; State = 'Pending'; Detail = '' },
                @{ Id = 'Domain.Join'; Title = 'Dołączenie do domeny'; State = 'Pending'; Detail = '' }
            )
            $progress = @{ Percent = 50; Text = 'Pauza po kroku 2 z 4' }
            $log.Add(@{ Time = $base.AddSeconds(4); Level = 'Info'; Source = 'Std.Engine'; Message = 'Pauza: wdrożenie czeka na wznowienie.' })
        }
        'Error' {
            $steps = @(
                @{ Id = 'Network.WaitForNetwork'; Title = 'Czekanie na sieć'; State = 'Skipped'; Detail = 'Już zrobione: adres 192.168.10.57.' },
                @{ Id = 'Power.SuspendSleep'; Title = 'Wstrzymanie usypiania'; State = 'Ok'; Detail = 'Usypianie wstrzymane do końca wdrożenia.' },
                @{ Id = 'Apps.Install'; Title = 'Instalacja aplikacji'; State = 'Error'; Detail = 'Antywirus (eset.msi): kod wyjścia 1603 - instalator zgłosił błąd krytyczny. Pozostałe aplikacje zainstalowane.' },
                @{ Id = 'Domain.Join'; Title = 'Dołączenie do domeny'; State = 'Cancelled'; Detail = 'Przerwano przed rozpoczęciem kroku.' }
            )
            $progress = @{ Percent = 100; Text = 'Zakończono' }
            $result = @{ Status = 'Cancelled'; RebootRequired = $true; Error = $null }
            $log.Add(@{ Time = $base.AddSeconds(60); Level = 'Error'; Source = 'Apps.Install'; Message = 'Wynik kroku Instalacja aplikacji: Error - Antywirus (eset.msi): kod wyjścia 1603.' })
            $log.Add(@{ Time = $base.AddSeconds(61); Level = 'Warn'; Source = 'Std.Engine'; Message = 'Przerwano wdrożenie na żądanie. Niewykonane kroki: 1.' })
        }
        'Finished' {
            $steps = @(
                @{ Id = 'Network.WaitForNetwork'; Title = 'Czekanie na sieć'; State = 'Skipped'; Detail = 'Już zrobione: adres 192.168.10.57.' },
                @{ Id = 'Power.SuspendSleep'; Title = 'Wstrzymanie usypiania'; State = 'Ok'; Detail = 'Usypianie wstrzymane do końca wdrożenia.' },
                @{ Id = 'Apps.Install'; Title = 'Instalacja aplikacji'; State = 'Ok'; Detail = 'Zainstalowano 5 aplikacji.' },
                @{ Id = 'Domain.Join'; Title = 'Dołączenie do domeny'; State = 'Ok'; Detail = 'Komputer dołączony do firma.local.' }
            )
            $progress = @{ Percent = 100; Text = 'Zakończono' }
            $result = @{ Status = 'Finished'; RebootRequired = $false; Error = $null }
            $log.Add(@{ Time = $base.AddSeconds(300); Level = 'Info'; Source = 'Std.Engine'; Message = 'Koniec wdrożenia 20261102-101500-ab12: zakończono. OK: 3, pominięto: 1, błąd: 0, przerwano: 0.' })
        }
    }
    return @{
        UserName      = 'jkowalski'
        Server        = 'https://pobieranie.example.com/std/'
        Index         = $index
        IndexSource   = 'server'
        Release       = $release
        ReleaseSource = 'server'
        ClientId      = 'lizard'
        SiteId        = 'krakow'
        ProfileName   = 'Standard'
        ComputerInfo  = $computer
        Plan          = $plan
        Steps         = $steps
        Log           = $log.ToArray()
        Progress      = $progress
        Result        = $result
        DeploymentId  = '20261102-101500-ab12'
        Networks      = @(
            [pscustomobject]@{ Ssid = 'Firma'; Auth = 'WPA2PSK'; AuthText = 'WPA2-Personal'; AuthLabel = 'WPA2'; Encryption = 'CCMP'; Signal = 86; SignalText = '86%'; Secured = $true },
            [pscustomobject]@{ Ssid = 'Firma-Goscie z bardzo długą nazwą sieci do sprawdzenia'; Auth = 'open'; AuthText = 'Open'; AuthLabel = 'otwarta'; Encryption = 'None'; Signal = 64; SignalText = '64%'; Secured = $false },
            [pscustomobject]@{ Ssid = 'Biuro-5G'; Auth = 'WPA3SAE'; AuthText = 'WPA3-Personal'; AuthLabel = 'WPA3'; Encryption = 'CCMP'; Signal = 41; SignalText = '41%'; Secured = $true },
            [pscustomobject]@{ Ssid = 'Firma-802.1X'; Auth = 'enterprise'; AuthText = 'WPA2-Enterprise'; AuthLabel = 'firmowa (802.1X)'; Encryption = 'CCMP'; Signal = 30; SignalText = '30%'; Secured = $true }
        )
    }
}

function Merge-StdUiPreviewData {
    <#
    .SYNOPSIS
    Dane podglądu: przykładowe (Get-StdUiPreviewSample -State) nadpisane kluczami z -PreviewData (bez WPF).
    #>
    [CmdletBinding()]
    param([AllowNull()][hashtable]$PreviewData, [string]$State = 'Select')
    $data = Get-StdUiPreviewSample -State $State
    if ($null -ne $PreviewData) {
        foreach ($key in @($PreviewData.Keys)) { $data[$key] = $PreviewData[$key] }
    }
    return $data
}

# ---------------------------------------------------------------------------------------------
# Stan okna w handlerach (WPF)
# ---------------------------------------------------------------------------------------------

function Get-StdUiState {
    <#
    .SYNOPSIS
    Obiekt stanu okna ([hashtable] z kluczem Window) dla nadawcy zdarzenia: sam stan, DispatcherTimer
    albo inny obiekt z takim Tag, okno albo kontrolka w oknie (szukanie okna w górę drzewa). Brak: $null.
    #>
    [CmdletBinding()]
    param([Parameter(Position = 0)][AllowNull()]$EventSource)
    if ($null -eq $EventSource) { return $null }
    if ($EventSource -is [hashtable] -and $EventSource.ContainsKey('Window')) { return $EventSource }
    $tag = $null
    try { $tag = $EventSource.Tag } catch { $tag = $null }
    if ($tag -is [hashtable] -and $tag.ContainsKey('Window')) { return $tag }
    if ($EventSource -is [System.Windows.DependencyObject]) {
        $window = $null
        try { $window = [System.Windows.Window]::GetWindow($EventSource) } catch { $window = $null }
        $node = $EventSource
        $guard = 0
        while ($null -eq $window -and $null -ne $node -and $guard -lt 200) {
            $guard++
            if ($node -is [System.Windows.Window]) { $window = $node; break }
            $parent = $null
            if ($node -is [System.Windows.Media.Visual] -or $node -is [System.Windows.Media.Media3D.Visual3D]) {
                try { $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($node) } catch { $parent = $null }
            }
            if ($null -eq $parent) {
                try { $parent = [System.Windows.LogicalTreeHelper]::GetParent($node) } catch { $parent = $null }
            }
            $node = $parent
        }
        if ($null -ne $window -and $window.Tag -is [hashtable] -and $window.Tag.ContainsKey('Window')) { return $window.Tag }
    }
    return $null
}

function Invoke-StdUiHandler {
    <#
    .SYNOPSIS
    Handler zdarzenia okna STD: znajduje stan okna po nadawcy (Get-StdUiState) i uruchamia -Action
    z parametrem $Ui w Invoke-StdUiAction (błąd = log + okno komunikatu, okno działa dalej).

        $btn.Add_Click({ param($s, $e) Invoke-StdUiHandler $s { param($Ui) Start-Cos -Ui $Ui } })

    Akcja widzi też zmienne handlera ($s, $e).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][AllowNull()]$EventSource,
        [Parameter(Mandatory = $true, Position = 1)][scriptblock]$Action,
        [string]$ErrorContext = 'Nie udało się wykonać tej czynności.'
    )
    # Nazwy z prefiksem: akcja działa w zakresie potomnym i nie może trafić na zwykłe nazwy tej funkcji.
    $stdUiHandlerAction = $Action
    $stdUiHandlerUi = Get-StdUiState -EventSource $EventSource
    $stdUiHandlerContext = $ErrorContext
    $stdUiHandlerOwner = $null
    if ($null -ne $stdUiHandlerUi -and $stdUiHandlerUi['Window'] -is [System.Windows.Window]) { $stdUiHandlerOwner = $stdUiHandlerUi['Window'] }
    Remove-Variable -Name EventSource, Action, ErrorContext -Scope Local -ErrorAction SilentlyContinue
    Invoke-StdUiAction -Owner $stdUiHandlerOwner -ErrorContext $stdUiHandlerContext -Action {
        if ($null -eq $stdUiHandlerUi) { throw '[STD-UI-STATE] Nie znaleziono stanu okna dla tego zdarzenia.' }
        [void](& $stdUiHandlerAction $stdUiHandlerUi)
    }
}

function New-StdUiCallback {
    <#
    .SYNOPSIS
    Blok dla -OnComplete / -OnError zadania w tle: woła -Handler z parametrami ($Ui, $Result).
    Stan okna jest zapamiętany w bloku, więc wynik trafia do właściwego okna bez zmiennych z ShowDialog().
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Ui,
        [Parameter(Mandatory = $true)][scriptblock]$Handler
    )
    $stdCbUi = $Ui
    $stdCbHandler = $Handler
    return { param($StdCbResult) & $stdCbHandler $stdCbUi $StdCbResult }.GetNewClosure()
}

function New-StdUiTimer {
    <#
    .SYNOPSIS
    DispatcherTimer okna. -OnTick: blok param($Ui, $Timer) wołany przez Invoke-StdUiHandler; błąd
    zatrzymuje timer (bez serii okien z błędem) i jest pokazany raz. Timer nie jest uruchamiany.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Ui,
        [Parameter(Mandatory = $true)][int]$IntervalMs,
        [Parameter(Mandatory = $true)][scriptblock]$OnTick
    )
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds($IntervalMs)
    # Tag: stan okna i blok timera (zmienne tej funkcji znikną, zanim timer zadziała).
    $timer.Tag = @{ State = $Ui; OnTick = $OnTick }
    $timer.Add_Tick({
            param($s, $e)
            $stdTimerInfo = $s.Tag
            Invoke-StdUiHandler $stdTimerInfo.State {
                param($Ui)
                try {
                    & $stdTimerInfo.OnTick $Ui $s
                } catch {
                    try { $s.Stop() } catch { }
                    throw
                }
            }
        })
    return $timer
}

# ---------------------------------------------------------------------------------------------
# Pomocnicze WPF: kontrolki, kolory, wiersze, schowek
# ---------------------------------------------------------------------------------------------

function Get-StdUiControls {
    <#
    .SYNOPSIS
    Kontrolki okna po nazwach (FindName) jako [hashtable]. Brak kontrolki = błąd w XAML: [STD-UI-XAML].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.FrameworkElement]$Root,
        [Parameter(Mandatory = $true)][string[]]$Names
    )
    $controls = @{}
    foreach ($name in $Names) {
        $control = $Root.FindName($name)
        if ($null -eq $control) { throw "[STD-UI-XAML] W oknie brakuje kontrolki '$name'." }
        $controls[$name] = $control
    }
    return $controls
}

function Set-StdUiColor {
    <#
    .SYNOPSIS
    Kolor właściwości (np. Foreground, Background, Fill, BorderBrush) z palety motywu przez
    SetResourceReference - przemalowuje się razem z motywem. Nieznany klucz: [STD-UI-THEME].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][System.Windows.FrameworkElement]$Element,
        [Parameter(Mandatory = $true, Position = 1)][string]$Property,
        [Parameter(Mandatory = $true, Position = 2)][string]$Key
    )
    $palette = Get-StdThemePalette
    if (-not $palette.Contains($Key)) { throw "[STD-UI-THEME] W palecie motywu nie ma koloru '$Key'." }
    $flags = [System.Reflection.BindingFlags]'Public, Static, FlattenHierarchy'
    $field = $Element.GetType().GetField($Property + 'Property', $flags)
    if ($null -eq $field) { throw "[STD-UI-THEME] Kontrolka $($Element.GetType().Name) nie ma właściwości $Property." }
    $Element.SetResourceReference($field.GetValue($null), $Key)
}

function Set-StdUiVisible {
    # Pokazuje (Visible) albo zwija (Collapsed) element.
    param([Parameter(Mandatory = $true, Position = 0)][System.Windows.UIElement]$Element, [Parameter(Mandatory = $true, Position = 1)][bool]$Visible)
    if ($Visible) { $Element.Visibility = [System.Windows.Visibility]::Visible } else { $Element.Visibility = [System.Windows.Visibility]::Collapsed }
}

function Set-StdUiButton {
    # Przycisk: tekst z ikoną, styl z zasobów okna (np. PrimaryButton; pusty = domyślny), podpowiedź.
    param(
        [Parameter(Mandatory = $true, Position = 0)][System.Windows.Controls.Button]$Button,
        # Pusty tekst = sam przycisk z ikoną (np. przełącznik motywu).
        [Parameter(Mandatory = $true, Position = 1)][AllowEmptyString()][string]$Text,
        [string]$Icon = '',
        [AllowNull()][AllowEmptyString()][string]$Style,
        [AllowNull()][AllowEmptyString()][string]$ToolTip
    )
    Set-StdButtonContent -Button $Button -Text $Text -Icon $Icon
    if ($PSBoundParameters.ContainsKey('Style')) {
        if ([string]::IsNullOrEmpty($Style)) {
            $Button.ClearValue([System.Windows.FrameworkElement]::StyleProperty)
        } else {
            $Button.Style = $Button.FindResource($Style)
        }
    }
    if ($PSBoundParameters.ContainsKey('ToolTip')) {
        if ([string]::IsNullOrEmpty($ToolTip)) { $Button.ToolTip = $null } else { $Button.ToolTip = $ToolTip }
    }
}

function Set-StdUiInfoRows {
    <#
    .SYNOPSIS
    Wypełnia siatkę (Grid z dwiema kolumnami: Auto, *) wierszami "etykieta / wartość". Tekst zawija się,
    kolory z palety. -Copyable: kliknięcie wartości kopiuje ją do schowka.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Controls.Grid]$Grid,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [switch]$Copyable
    )
    $Grid.Children.Clear()
    $Grid.RowDefinitions.Clear()
    $index = 0
    foreach ($row in $Rows) {
        $definition = New-Object System.Windows.Controls.RowDefinition
        $definition.Height = [System.Windows.GridLength]::Auto
        $Grid.RowDefinitions.Add($definition)
        $bottom = 7
        if ($index -eq $Rows.Count - 1) { $bottom = 0 }

        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = [string]$row.Label
        $label.FontSize = 12
        $label.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $label.MaxWidth = 150
        $label.Margin = New-Object System.Windows.Thickness(0, 0, 14, $bottom)
        Set-StdUiColor $label 'Foreground' 'ThemeMuted'
        [System.Windows.Controls.Grid]::SetRow($label, $index)
        [System.Windows.Controls.Grid]::SetColumn($label, 0)
        [void]$Grid.Children.Add($label)

        $value = New-Object System.Windows.Controls.TextBlock
        $value.Text = [string]$row.Value
        $value.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $value.Margin = New-Object System.Windows.Thickness(0, 0, 0, $bottom)
        Set-StdUiColor $value 'Foreground' 'ThemeText'
        if ($Copyable -and [string]$row.Value -ne 'brak danych') {
            $value.Style = $Grid.FindResource('CopyableLine')
            $value.Add_MouseLeftButtonUp({
                    param($s, $e)
                    Invoke-StdUiHandler $s {
                        param($Ui)
                        Copy-StdUiText -Text $s.Text -Window $Ui.Window
                    }
                })
        }
        [System.Windows.Controls.Grid]::SetRow($value, $index)
        [System.Windows.Controls.Grid]::SetColumn($value, 1)
        [void]$Grid.Children.Add($value)
        $index++
    }
}

function Copy-StdUiText {
    <#
    .SYNOPSIS
    Kopiuje tekst do schowka (z ponowieniem, gdy schowek jest zajęty) i pokazuje krótkie powiadomienie.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [System.Windows.Window]$Window,
        [string]$Message = 'Skopiowano do schowka.'
    )
    $done = $false
    $lastError = $null
    for ($attempt = 0; $attempt -lt 3 -and -not $done; $attempt++) {
        try {
            [System.Windows.Clipboard]::SetText((Protect-StdLogText -Text $Text))
            $done = $true
        } catch {
            $lastError = $_
        }
    }
    if (-not $done) {
        Show-StdToast -Window $Window -Message 'Nie udało się skopiować do schowka - spróbuj ponownie.' -Kind Warning
        if ($null -ne $lastError) { Write-StdLog -Level Warn -Source 'UI' -Message ('Schowek: {0}' -f $lastError.Exception.Message) }
        return
    }
    Show-StdToast -Window $Window -Message $Message -Kind Success -Seconds 2
}

function Update-StdUiGridColumns {
    <#
    .SYNOPSIS
    Szerokość „rozciągliwych” kolumn tabeli (ListView + GridView): -Flex @{ <nr kolumny> = <udział> } dzieli
    miejsce, które zostaje po kolumnach o stałej szerokości. Tekst w tych kolumnach się zawija, więc tabela
    nie potrzebuje poziomego paska przewijania.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Controls.ListView]$ListView,
        [Parameter(Mandatory = $true)][hashtable]$Flex,
        [double]$MinWidth = 60,
        [double]$Reserve = 30
    )
    $view = $ListView.View
    if ($view -isnot [System.Windows.Controls.GridView]) { return }
    $width = $ListView.ActualWidth
    if ($width -le 0 -or -not $ListView.IsVisible) { return }
    $fixed = 0.0
    for ($i = 0; $i -lt $view.Columns.Count; $i++) {
        if ($Flex.ContainsKey($i)) { continue }
        $column = $view.Columns[$i]
        if ([double]::IsNaN($column.Width)) { $fixed += $column.ActualWidth } else { $fixed += $column.Width }
    }
    $free = $width - $fixed - $Reserve
    foreach ($key in @($Flex.Keys)) {
        $target = [Math]::Floor($free * [double]$Flex[$key])
        if ($target -lt $MinWidth) { $target = $MinWidth }
        $column = $view.Columns[[int]$key]
        if ([double]::IsNaN($column.Width) -or [Math]::Abs($column.Width - $target) -ge 1) { $column.Width = $target }
    }
}

function Select-StdUiWindow {
    # Okno z wyniku funkcji New-Std*Window (ostatni obiekt typu Window).
    param([AllowNull()][AllowEmptyCollection()][object[]]$Output)
    $window = $null
    foreach ($item in @($Output)) { if ($item -is [System.Windows.Window]) { $window = $item } }
    if ($null -eq $window) { throw '[STD-UI-XAML] Funkcja okna nie zwróciła okna.' }
    return $window
}
