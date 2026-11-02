# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.App.ps1 - uruchamianie STD (SPEC 8): kopia narzędzia z pendrive'a do %ProgramData%\STD\App\<wersja>,
# porównanie kopii po SHA-256, usuwanie starych wersji z App\ oraz argumenty do ponownego uruchomienia
# programu (jako administrator, w trybie STA, z kopii w App\<wersja>). Do tego szybkie sprawdzenie,
# czy komputer ma sieć (program pokazuje wtedy ekran „Połącz z siecią” albo od razu logowanie).
#
# Bez WPF - działa bez okna i na Linuksie (testy). Program (SmartToolforDeployment.ps1) woła te funkcje
# w krokach 1-3, zanim wczyta okna.
#
# Jak działa kopia:
# - kopiujemy tylko pliki narzędzia ($script:StdAppItems): STD.psd1, STD.psm1, SmartToolforDeployment.ps1
#   i foldery Modules, GUI, Steps. Testy, narzędzia administratora, bootstrap.json, offline\ i wifi\
#   zostają na pendrive'ie (program dostaje jego folder w -OriginRoot);
# - kopia jest aktualna, gdy ma dokładnie te same pliki (te same ścieżki i te same sumy SHA-256).
#   Wtedy niczego nie kopiujemy. Inaczej kopiujemy tylko pliki zmienione i nowe, a usuwamy pliki,
#   których w źródle już nie ma (żeby stary krok z Steps\ nie został wczytany);
# - nowa wersja trafia do nowego folderu App\<wersja>. Reszta %ProgramData%\STD (ustawienia, cache)
#   zostaje bez zmian. W App\ zostają dwie najnowsze wersje (według [version]) i zawsze ta,
#   z której właśnie startujemy; foldery o nazwach, które nie są wersją, nie są ruszane.
#
# Funkcje pomocnicze bez "-Std" w nazwie nie są eksportowane z modułu (Export-ModuleMember *-Std*).

# Pliki i foldery narzędzia kopiowane do App\<wersja> (względem folderu v4).
$script:StdAppItems = @('STD.psd1', 'STD.psm1', 'SmartToolforDeployment.ps1', 'Modules', 'GUI', 'Steps')

# ---------------------------------------------------------------------------------------------
# Pomocnicze (prywatne)
# ---------------------------------------------------------------------------------------------

function Get-AppPathComparison {
    # Ścieżki w Windows porównujemy bez względu na wielkość liter, na Linuksie dokładnie.
    if (Test-CoreIsWindows) { return [System.StringComparison]::OrdinalIgnoreCase }
    return [System.StringComparison]::Ordinal
}

function Get-AppPathComparer {
    if (Test-CoreIsWindows) { return [System.StringComparer]::OrdinalIgnoreCase }
    return [System.StringComparer]::Ordinal
}

function Get-AppNormalPath {
    # Pełna ścieżka systemu plików bez ukośnika na końcu (poza katalogiem głównym, np. E:\ albo /).
    param([Parameter(Mandatory = $true)][string]$Path)
    $full = [System.IO.Path]::GetFullPath((Resolve-CoreFullPath -Path $Path))
    $root = [System.IO.Path]::GetPathRoot($full)
    if ($null -eq $root) { $root = '' }
    if ($full.Length -gt $root.Length) {
        $full = $full.TrimEnd([char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar))
        if ($full.Length -lt $root.Length) { $full = $root }
    }
    return $full
}

function Test-AppPathEqual {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Other)
    return [string]::Equals((Get-AppNormalPath -Path $Path), (Get-AppNormalPath -Path $Other), (Get-AppPathComparison))
}

function Test-AppPathInside {
    # Czy $Path to $Parent albo leży w nim (na dowolnej głębokości).
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Parent)
    $child = Get-AppNormalPath -Path $Path
    $root = Get-AppNormalPath -Path $Parent
    $comparison = Get-AppPathComparison
    if ([string]::Equals($child, $root, $comparison)) { return $true }
    $separator = [string][System.IO.Path]::DirectorySeparatorChar
    $prefix = $root
    if (-not $prefix.EndsWith($separator)) { $prefix = $prefix + $separator }
    return $child.StartsWith($prefix, $comparison)
}

function Get-AppSourceVersion {
    # Wersja z <Source>\STD.psd1 (ModuleVersion). Nie da się odczytać -> [STD-APP-SOURCE].
    param([Parameter(Mandatory = $true)][string]$SourceRoot)
    $manifest = [System.IO.Path]::Combine($SourceRoot, 'STD.psd1')
    try {
        $data = Import-PowerShellDataFile -LiteralPath $manifest -ErrorAction Stop
        return [version][string]$data['ModuleVersion']
    } catch {
        throw "[STD-APP-SOURCE] Nie udało się odczytać wersji STD z pliku $manifest - $(Format-CoreErrorText $_)"
    }
}

function Assert-AppSource {
    # Folder źródłowy musi istnieć i mieć STD.psd1 (inaczej to nie jest folder STD).
    param([Parameter(Mandatory = $true)][string]$SourceRoot)
    if (-not [System.IO.Directory]::Exists($SourceRoot)) {
        throw "[STD-APP-SOURCE] Nie znaleziono folderu STD $SourceRoot."
    }
    if (-not [System.IO.File]::Exists([System.IO.Path]::Combine($SourceRoot, 'STD.psd1'))) {
        throw "[STD-APP-SOURCE] W folderze $SourceRoot nie ma pliku STD.psd1 - to nie jest folder STD."
    }
}

function Get-AppFileTable {
    # Pliki narzędzia w folderze: słownik ścieżka względna -> pełna ścieżka (tylko elementy $script:StdAppItems).
    param([Parameter(Mandatory = $true)][string]$Root)
    $table = New-Object 'System.Collections.Generic.Dictionary[string,string]' -ArgumentList (Get-AppPathComparer)
    if (-not [System.IO.Directory]::Exists($Root)) { return , $table }
    $trim = [char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    foreach ($item in $script:StdAppItems) {
        $path = [System.IO.Path]::Combine($Root, $item)
        if ([System.IO.File]::Exists($path)) {
            $table[$item] = $path
            continue
        }
        if (-not [System.IO.Directory]::Exists($path)) { continue }
        foreach ($file in [System.IO.Directory]::GetFiles($path, '*', [System.IO.SearchOption]::AllDirectories)) {
            $relative = $file.Substring($Root.Length).TrimStart($trim)
            $table[$relative] = $file
        }
    }
    return , $table
}

function Get-AppCopyDifference {
    # Co trzeba zrobić, żeby cel był taki sam jak źródło: @{ Copy = [ścieżki względne]; Remove = [ścieżki względne] }.
    # Najpierw rozmiar (szybko), potem SHA-256.
    param(
        [Parameter(Mandatory = $true)]$SourceTable,
        [Parameter(Mandatory = $true)]$TargetTable
    )
    $copy = New-Object 'System.Collections.Generic.List[string]'
    $remove = New-Object 'System.Collections.Generic.List[string]'
    foreach ($relative in $SourceTable.Keys) {
        if (-not $TargetTable.ContainsKey($relative)) {
            $copy.Add($relative)
            continue
        }
        $sourceFile = $SourceTable[$relative]
        $targetFile = $TargetTable[$relative]
        if ((New-Object System.IO.FileInfo($sourceFile)).Length -ne (New-Object System.IO.FileInfo($targetFile)).Length) {
            $copy.Add($relative)
            continue
        }
        if ((Get-StdSha256 -Path $sourceFile) -ne (Get-StdSha256 -Path $targetFile)) { $copy.Add($relative) }
    }
    foreach ($relative in $TargetTable.Keys) {
        if (-not $SourceTable.ContainsKey($relative)) { $remove.Add($relative) }
    }
    return @{ Copy = $copy.ToArray(); Remove = $remove.ToArray() }
}

function Clear-AppReadOnly {
    # Pliki z pendrive'a / płyty mogą mieć atrybut „tylko do odczytu” - wtedy nie dałoby się ich nadpisać ani usunąć.
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not [System.IO.File]::Exists($Path)) { return }
    $attributes = [System.IO.File]::GetAttributes($Path)
    if (($attributes -band [System.IO.FileAttributes]::ReadOnly) -ne 0) {
        [System.IO.File]::SetAttributes($Path, ($attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly)))
    }
}

function Remove-AppEmptyFolders {
    # Puste podfoldery (np. po usunięciu pliku z GUI\Screens) w folderach narzędzia w celu.
    param([Parameter(Mandatory = $true)][string]$TargetRoot)
    foreach ($item in $script:StdAppItems) {
        $path = [System.IO.Path]::Combine($TargetRoot, $item)
        if (-not [System.IO.Directory]::Exists($path)) { continue }
        $dirs = @([System.IO.Directory]::GetDirectories($path, '*', [System.IO.SearchOption]::AllDirectories) | Sort-Object -Property Length -Descending)
        foreach ($dir in $dirs) {
            if (@([System.IO.Directory]::GetFileSystemEntries($dir)).Count -eq 0) {
                [System.IO.Directory]::Delete($dir, $false)
            }
        }
    }
}

function Copy-AppFiles {
    # Kopia lustrzana plików narzędzia: kopiuje zmienione i nowe pliki, usuwa zbędne. Wynik: @{ Copied; Removed }.
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$TargetRoot
    )
    $sourceTable = Get-AppFileTable -Root $SourceRoot
    $targetTable = Get-AppFileTable -Root $TargetRoot
    $diff = Get-AppCopyDifference -SourceTable $sourceTable -TargetTable $targetTable
    if (-not [System.IO.Directory]::Exists($TargetRoot)) { [void][System.IO.Directory]::CreateDirectory($TargetRoot) }
    foreach ($relative in $diff.Copy) {
        $destination = [System.IO.Path]::Combine($TargetRoot, $relative)
        $folder = [System.IO.Path]::GetDirectoryName($destination)
        if (-not [System.IO.Directory]::Exists($folder)) { [void][System.IO.Directory]::CreateDirectory($folder) }
        Clear-AppReadOnly -Path $destination
        [System.IO.File]::Copy($sourceTable[$relative], $destination, $true)
        Clear-AppReadOnly -Path $destination
    }
    foreach ($relative in $diff.Remove) {
        $file = $targetTable[$relative]
        Clear-AppReadOnly -Path $file
        [System.IO.File]::Delete($file)
    }
    Remove-AppEmptyFolders -TargetRoot $TargetRoot
    return @{ Copied = $diff.Copy.Count; Removed = $diff.Remove.Count }
}

function Remove-AppFolder {
    # Usuwa folder z całą zawartością (najpierw zdejmuje atrybut „tylko do odczytu” z plików).
    param([Parameter(Mandatory = $true)][string]$Path)
    foreach ($file in [System.IO.Directory]::GetFiles($Path, '*', [System.IO.SearchOption]::AllDirectories)) {
        Clear-AppReadOnly -Path $file
    }
    [System.IO.Directory]::Delete($Path, $true)
}

function Remove-AppOldVersions {
    # W App\ zostawia $Keep najnowszych wersji (według [version]) i foldery z $Protect; usuwa resztę folderów-wersji.
    # Foldery, których nazwa nie jest wersją, zostają. Błąd usuwania (plik w użyciu) = ostrzeżenie w logu.
    # Zwraca nazwy usuniętych folderów.
    param(
        [Parameter(Mandatory = $true)][string]$AppRoot,
        [Parameter(Mandatory = $true)][int]$Keep,
        [string[]]$Protect = @()
    )
    $removed = New-Object 'System.Collections.Generic.List[string]'
    if (-not [System.IO.Directory]::Exists($AppRoot)) { return , $removed.ToArray() }
    $versions = New-Object 'System.Collections.Generic.List[object]'
    foreach ($dir in [System.IO.Directory]::GetDirectories($AppRoot)) {
        $name = [System.IO.Path]::GetFileName($dir)
        $parsed = $null
        if ([version]::TryParse($name, [ref]$parsed)) {
            $versions.Add([pscustomobject]@{ Path = $dir; Name = $name; Version = $parsed })
        }
    }
    $sorted = @($versions | Sort-Object -Property Version -Descending)
    for ($i = $Keep; $i -lt $sorted.Count; $i++) {
        $entry = $sorted[$i]
        $isProtected = $false
        foreach ($keepPath in $Protect) {
            if (-not [string]::IsNullOrEmpty($keepPath) -and (Test-AppPathEqual -Path $entry.Path -Other $keepPath)) { $isProtected = $true }
        }
        if ($isProtected) { continue }
        try {
            Remove-AppFolder -Path $entry.Path
            $removed.Add($entry.Name)
            Write-StdLog -Level Info -Source 'Sync-StdAppFolder' -Message "Usunięto starą wersję STD: $($entry.Path)"
        } catch {
            Write-StdLog -Level Warn -Source 'Sync-StdAppFolder' -Message "Nie udało się usunąć starej wersji STD $($entry.Path) - $(Format-CoreErrorText $_). Spróbuję przy następnym uruchomieniu."
        }
    }
    return , $removed.ToArray()
}

function ConvertTo-AppArgumentText {
    # Wartość parametru jako tekst dla wiersza poleceń (liczby i daty w formacie niezależnym od języka systemu).
    # Hasła nigdy nie trafiają do wiersza poleceń (widać go np. w Menedżerze zadań).
    param([Parameter(Mandatory = $true)][string]$Name, [AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Security.SecureString] -or $Value -is [System.Management.Automation.PSCredential]) {
        throw "[STD-APP-ARGS] Parametr -$Name zawiera hasło - hasła nie mogą być przekazywane w wierszu poleceń."
    }
    if ($Value -is [System.Collections.IDictionary]) {
        throw "[STD-APP-ARGS] Parametr -${Name}: słownika nie da się przekazać w wierszu poleceń."
    }
    if ($Value -is [datetime]) { return $Value.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture) }
    return [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------------------------------------
# Funkcje publiczne
# ---------------------------------------------------------------------------------------------

function Get-StdAppFolder {
    <#
    .SYNOPSIS
    Folder kopii narzędzia dla wersji: <Get-StdPath App>\<wersja> (domyślnie wersja bieżącego modułu).
    Tworzy tylko App\ (jak Get-StdPath), folderu wersji nie.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 0)]
        [version]$Version
    )
    if ($null -eq $Version) { $Version = Get-StdVersion }
    return [System.IO.Path]::Combine((Get-StdPath -Kind App), $Version.ToString())
}

function Test-StdInAppFolder {
    <#
    .SYNOPSIS
    $true, gdy ścieżka leży w <Get-StdPath App> (czyli program już działa z kopii App\<wersja>).
    Program nie kopiuje się wtedy drugi raz.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    try {
        return [bool](Test-AppPathInside -Path $Path -Parent (Get-StdPath -Kind App))
    } catch {
        Write-StdLog -Level Warn -Source 'Test-StdInAppFolder' -Message "Nie udało się porównać ścieżki $Path z folderem App - $(Format-CoreErrorText $_)"
        return $false
    }
}

function Test-StdAppCopyCurrent {
    <#
    .SYNOPSIS
    $true, gdy folder -Target ma dokładnie te same pliki narzędzia co -Source (te same ścieżki względne
    i te same sumy SHA-256; liczą się STD.psd1, STD.psm1, SmartToolforDeployment.ps1, Modules, GUI, Steps).
    Brak folderu docelowego -> $false. Źródło bez STD.psd1 -> [STD-APP-SOURCE].
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Source,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Target
    )
    $sourceRoot = Get-AppNormalPath -Path $Source
    $targetRoot = Get-AppNormalPath -Path $Target
    Assert-AppSource -SourceRoot $sourceRoot
    if (-not [System.IO.Directory]::Exists($targetRoot)) { return $false }
    if ([string]::Equals($sourceRoot, $targetRoot, (Get-AppPathComparison))) { return $true }
    $diff = Get-AppCopyDifference -SourceTable (Get-AppFileTable -Root $sourceRoot) -TargetTable (Get-AppFileTable -Root $targetRoot)
    return ($diff.Copy.Count -eq 0 -and $diff.Remove.Count -eq 0)
}

function Sync-StdAppFolder {
    <#
    .SYNOPSIS
    Kopiuje narzędzie z folderu -Source (v4, np. pendrive) do <Get-StdPath App>\<wersja> i zwraca ten folder.
    Kopia jest pomijana, gdy pliki są identyczne (SHA-256); przy różnicach kopiowane są tylko zmienione
    i nowe pliki, a zbędne usuwane. Potem usuwa stare wersje z App\: zostają -Keep najnowsze (według
    [version]) i zawsze wersja docelowa. -Version domyślnie z <Source>\STD.psd1.
    Gdy -Source to już folder docelowy, niczego nie kopiuje. Błędy: [STD-APP-SOURCE], [STD-APP-COPY].
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Source,
        [Parameter(Position = 1)]
        [version]$Version,
        [ValidateRange(1, 1000)]
        [int]$Keep = 2
    )
    $sourceRoot = Get-AppNormalPath -Path $Source
    Assert-AppSource -SourceRoot $sourceRoot
    if ($null -eq $Version) { $Version = Get-AppSourceVersion -SourceRoot $sourceRoot }
    $appRoot = Get-StdPath -Kind App
    $targetRoot = Get-AppNormalPath -Path (Get-StdAppFolder -Version $Version)

    if (Test-AppPathEqual -Path $sourceRoot -Other $targetRoot) {
        Write-StdLog -Level Debug -Source 'Sync-StdAppFolder' -Message "STD już działa z kopii $targetRoot - bez kopiowania."
    } else {
        if (Test-AppPathInside -Path $targetRoot -Parent $sourceRoot) {
            throw "[STD-APP-COPY] Folder kopii $targetRoot leży wewnątrz folderu źródłowego $sourceRoot - ustaw inny folder danych STD."
        }
        if (Test-StdAppCopyCurrent -Source $sourceRoot -Target $targetRoot) {
            Write-StdLog -Level Info -Source 'Sync-StdAppFolder' -Message "Kopia STD $Version w $targetRoot jest aktualna - bez kopiowania."
        } else {
            try {
                $result = Copy-AppFiles -SourceRoot $sourceRoot -TargetRoot $targetRoot
            } catch {
                throw "[STD-APP-COPY] Nie udało się skopiować STD do $targetRoot - $(Format-CoreErrorText $_)"
            }
            if (-not (Test-StdAppCopyCurrent -Source $sourceRoot -Target $targetRoot)) {
                throw "[STD-APP-COPY] Kopia STD w $targetRoot różni się od plików w $sourceRoot po skopiowaniu (plik w użyciu albo uszkodzony nośnik)."
            }
            Write-StdLog -Level Info -Source 'Sync-StdAppFolder' -Message ("Skopiowano STD {0} do {1} (skopiowane pliki: {2}, usunięte: {3})." -f $Version, $targetRoot, $result.Copied, $result.Removed)
        }
    }

    [void](Remove-AppOldVersions -AppRoot $appRoot -Keep $Keep -Protect @($targetRoot, $sourceRoot))
    return $targetRoot
}

function Get-StdRelaunchArgument {
    <#
    .SYNOPSIS
    Argumenty do ponownego uruchomienia programu (dla Start-StdElevated -Arguments): parametry z
    -BoundParameters (np. $PSBoundParameters), nadpisane i uzupełnione przez -Add.
    Wynik to lista surowych elementów, np. @('-OriginRoot', 'E:\Pendrive STD', '-NoCopy') - cudzysłowy
    i ukośniki dodaje Start-StdElevated (każdy element osobno, reguły wiersza poleceń Windows).
    Zasady: przełącznik ([switch] / [bool]) $true -> sama nazwa, $false -> pomijany; $null i pusty tekst
    pomijane (w -Add: usuwa parametr); tablica -> elementy połączone przecinkiem; parametry w kolejności
    alfabetycznej. Hasło (SecureString / PSCredential) albo nieprawidłowa nazwa -> [STD-APP-ARGS].
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [System.Collections.IDictionary]$BoundParameters,
        [Parameter(Position = 1)]
        [AllowNull()]
        [System.Collections.IDictionary]$Add
    )
    $merged = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($source in @($BoundParameters, $Add)) {
        if ($null -eq $source) { continue }
        foreach ($key in @($source.Keys)) { $merged[[string]$key] = $source[$key] }
    }
    $names = [string[]]@($merged.Keys)
    [System.Array]::Sort($names, [System.StringComparer]::OrdinalIgnoreCase)

    $tokens = New-Object 'System.Collections.Generic.List[string]'
    foreach ($name in $names) {
        if ($name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
            throw "[STD-APP-ARGS] Nieprawidłowa nazwa parametru: '$name'."
        }
        $value = $merged[$name]
        if ($null -eq $value) { continue }
        if ($value -is [System.Management.Automation.SwitchParameter]) {
            if ($value.IsPresent) { $tokens.Add('-' + $name) }
            continue
        }
        if ($value -is [bool]) {
            if ($value) { $tokens.Add('-' + $name) }
            continue
        }
        if ($value -is [string]) {
            $text = $value
        } elseif ($value -is [System.Collections.IEnumerable] -and -not ($value -is [System.Collections.IDictionary])) {
            $parts = New-Object 'System.Collections.Generic.List[string]'
            foreach ($element in $value) {
                $elementText = ConvertTo-AppArgumentText -Name $name -Value $element
                if (-not [string]::IsNullOrEmpty($elementText)) { $parts.Add($elementText) }
            }
            $text = $parts -join ','
        } else {
            $text = ConvertTo-AppArgumentText -Name $name -Value $value
        }
        if ([string]::IsNullOrEmpty($text)) { continue }
        $tokens.Add('-' + $name)
        $tokens.Add($text)
    }
    return , $tokens.ToArray()
}

function Test-StdNetworkAvailable {
    <#
    .SYNOPSIS
    $true, gdy działająca karta sieciowa ma adres IPv4 spoza 169.254.x.x (brak DHCP) i 127.x.x.x.
    Program pokazuje ekran „Połącz z siecią” tylko wtedy, gdy wynik to $false. Serwera tu nie sprawdzamy
    (to robi logowanie), więc wynik jest natychmiastowy.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    try {
        foreach ($nic in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
            if ($nic.NetworkInterfaceType -eq [System.Net.NetworkInformation.NetworkInterfaceType]::Loopback) { continue }
            # Unknown: część kart wirtualnych nie podaje stanu, choć działa.
            $status = $nic.OperationalStatus
            if ($status -ne [System.Net.NetworkInformation.OperationalStatus]::Up -and $status -ne [System.Net.NetworkInformation.OperationalStatus]::Unknown) { continue }
            foreach ($unicast in $nic.GetIPProperties().UnicastAddresses) {
                $address = $unicast.Address
                if ($address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
                $text = $address.ToString()
                if ($text.StartsWith('169.254.') -or $text.StartsWith('127.')) { continue }
                return $true
            }
        }
    } catch {
        Write-StdLog -Level Warn -Source 'Test-StdNetworkAvailable' -Message "Nie udało się odczytać stanu kart sieciowych - $(Format-CoreErrorText $_)"
    }
    return $false
}
