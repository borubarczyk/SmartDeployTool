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
#   z której właśnie startujemy; foldery o nazwach, które nie są wersją, nie są ruszane;
# - kod z App\<wersja> działa jako administrator, więc App\ ma uprawnienia tylko dla Administratorów
#   i SYSTEM (Protect-StdDataFolder). Kopii, która jest dowiązaniem (junction / symlink), ma dowiązanie
#   w środku, nie należy do Administratorów / SYSTEM albo nie ma jeszcze tych uprawnień (kopia sprzed tego
#   zabezpieczenia), nie ufamy: jest usuwana i kopiowana od nowa. Sprawdzona kopia dostaje te uprawnienia.
#   Foldery przeglądamy sami (bez GetFiles(AllDirectories), które w .NET 4.8 wchodzi w junction) i nigdy
#   nie wchodzimy w dowiązania - dowiązanie usuwamy jako dowiązanie, a nie pliki, na które wskazuje;
# - jedna instancja STD na komputerze: mutex Global\STD4-SmartToolForDeployment (Enter-StdAppInstance).
#   Gdy trzyma go inny proces STD, Sync-StdAppFolder nie usuwa starych wersji.
#
# Funkcje pomocnicze bez "-Std" w nazwie nie są eksportowane z modułu (Export-ModuleMember *-Std*).

# Pliki i foldery narzędzia kopiowane do App\<wersja> (względem folderu v4).
$script:StdAppItems = @('STD.psd1', 'STD.psm1', 'SmartToolforDeployment.ps1', 'Modules', 'GUI', 'Steps')

# Blokada jednej instancji STD (wszystkie sesje Windows) i mutex trzymany przez ten proces.
$script:StdAppMutexName = 'Global\STD4-SmartToolForDeployment'
if (-not (Get-Variable -Name 'StdAppInstanceMutex' -Scope Script -ErrorAction SilentlyContinue)) {
    $script:StdAppInstanceMutex = $null
}

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

function Get-AppTree {
    # Zawartość folderu na każdej głębokości: @{ Files; Directories; Links } (pełne ścieżki). Własny przegląd
    # ze stosem zamiast GetFiles/GetDirectories(AllDirectories), które w .NET 4.8 wchodzą w junction i zwróciłyby
    # pliki spoza folderu (junction może założyć każdy użytkownik z prawem zapisu w folderze). Dowiązania
    # (junction, symlink) trafiają do Links i nie są przeglądane; folder, który sam jest dowiązaniem, też.
    param([Parameter(Mandatory = $true)][string]$Path)
    $files = New-Object 'System.Collections.Generic.List[string]'
    $directories = New-Object 'System.Collections.Generic.List[string]'
    $links = New-Object 'System.Collections.Generic.List[string]'
    $pending = New-Object 'System.Collections.Generic.Stack[string]'
    if (Test-CoreIsLink -Path $Path) {
        $links.Add($Path)
    } else {
        $pending.Push($Path)
    }
    while ($pending.Count -gt 0) {
        $current = $pending.Pop()
        foreach ($entry in [System.IO.Directory]::GetFileSystemEntries($current)) {
            $attributes = [System.IO.File]::GetAttributes($entry)
            if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $links.Add($entry)
            } elseif (($attributes -band [System.IO.FileAttributes]::Directory) -ne 0) {
                $directories.Add($entry)
                $pending.Push($entry)
            } else {
                $files.Add($entry)
            }
        }
    }
    return @{ Files = $files.ToArray(); Directories = $directories.ToArray(); Links = $links.ToArray() }
}

function Get-AppLinkError {
    # Komunikat [STD-APP-COPY] dla dowiązania w folderze narzędzia.
    param([Parameter(Mandatory = $true)][string]$Link)
    return "[STD-APP-COPY] $Link to dowiązanie (junction / symlink) - STD nie kopiuje ani nie usuwa plików przez dowiązania."
}

function Get-AppFileTable {
    # Pliki narzędzia w folderze: słownik ścieżka względna -> pełna ścieżka (tylko elementy $script:StdAppItems).
    # Dowiązanie wśród nich -> [STD-APP-COPY] (pliki, na które wskazuje, nie mogą trafić do kopii ani do usunięcia).
    param([Parameter(Mandatory = $true)][string]$Root)
    $table = New-Object 'System.Collections.Generic.Dictionary[string,string]' -ArgumentList (Get-AppPathComparer)
    if (-not [System.IO.Directory]::Exists($Root)) { return , $table }
    $trim = [char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    foreach ($item in $script:StdAppItems) {
        $path = [System.IO.Path]::Combine($Root, $item)
        if (Test-CoreIsLink -Path $path) { throw (Get-AppLinkError -Link $path) }
        if ([System.IO.File]::Exists($path)) {
            $table[$item] = $path
            continue
        }
        if (-not [System.IO.Directory]::Exists($path)) { continue }
        $tree = Get-AppTree -Path $path
        if ($tree.Links.Count -gt 0) { throw (Get-AppLinkError -Link $tree.Links[0]) }
        foreach ($file in $tree.Files) {
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
    # Puste podfoldery (np. po usunięciu pliku z GUI\Screens) w folderach narzędzia w celu. Bez wchodzenia w dowiązania.
    param([Parameter(Mandatory = $true)][string]$TargetRoot)
    foreach ($item in $script:StdAppItems) {
        $path = [System.IO.Path]::Combine($TargetRoot, $item)
        if ((Test-CoreIsLink -Path $path) -or -not [System.IO.Directory]::Exists($path)) { continue }
        $dirs = @((Get-AppTree -Path $path).Directories | Sort-Object -Property Length -Descending)
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
    if (Test-CoreIsLink -Path $TargetRoot) { throw (Get-AppLinkError -Link $TargetRoot) }
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
    # Usuwa folder z całą zawartością (najpierw zdejmuje atrybut „tylko do odczytu” z plików). Dowiązania
    # (junction, symlink; także sam folder, gdy jest dowiązaniem) usuwa jako dowiązania - nigdy nie wchodzi
    # do folderu, na który wskazują, i nie usuwa niczego poza $Path.
    param([Parameter(Mandatory = $true)][string]$Path)
    if (Test-CoreIsLink -Path $Path) {
        Remove-CoreLink -Path $Path
        return
    }
    $tree = Get-AppTree -Path $Path
    foreach ($link in $tree.Links) { Remove-CoreLink -Path $link }
    foreach ($file in $tree.Files) {
        Clear-AppReadOnly -Path $file
        [System.IO.File]::Delete($file)
    }
    foreach ($dir in @($tree.Directories | Sort-Object -Property Length -Descending)) {
        [System.IO.Directory]::Delete($dir, $false)
    }
    [System.IO.Directory]::Delete($Path, $false)
}

function Test-AppTrustedOwner {
    # Czy właścicielem folderu są Administratorzy albo SYSTEM. Folder założony wcześniej przez zwykłego
    # użytkownika należy do niego: może on zmieniać jego uprawnienia i podmieniać pliki po sprawdzeniu SHA-256.
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        $owner = (Get-CoreDirectorySecurity -Path $Path).GetOwner([System.Security.Principal.SecurityIdentifier])
    } catch {
        return $false
    }
    if ($null -eq $owner) { return $false }
    return ($owner.Value -eq 'S-1-5-32-544' -or $owner.Value -eq 'S-1-5-18')
}

function Reset-AppUntrustedFolder {
    # Przed sprawdzeniem i kopią App\<wersja>: kopia, która jest dowiązaniem albo ma dowiązanie w środku, jest
    # usuwana - Copy-AppFiles założy ją od nowa z uprawnieniami App\. -CheckTrust (administrator, Windows):
    # usuwana jest też kopia, której właścicielem nie są Administratorzy / SYSTEM, i kopia bez uprawnień
    # z Protect-StdDataFolder (STD nadaje je po każdej sprawdzonej kopii) - to kopia sprzed tego zabezpieczenia,
    # do której zwykły użytkownik mógł dopisać swoje pliki. Usuwane są tylko same dowiązania, nigdy pliki,
    # na które wskazują.
    param([Parameter(Mandatory = $true)][string]$Path, [bool]$CheckTrust)
    $reason = $null
    if (Test-CoreIsLink -Path $Path) {
        $reason = 'folder jest dowiązaniem (junction / symlink)'
    } elseif (-not [System.IO.Directory]::Exists($Path)) {
        return
    } elseif ((Get-AppTree -Path $Path).Links.Count -gt 0) {
        $reason = 'w folderze jest dowiązanie (junction / symlink)'
    } elseif ($CheckTrust) {
        if (-not (Test-AppTrustedOwner -Path $Path)) {
            $reason = 'właścicielem folderu nie są Administratorzy ani SYSTEM'
        } else {
            $sealed = $false
            try { $sealed = [bool](Test-CoreDataFolderProtected -Path $Path) } catch { }
            if (-not $sealed) { $reason = 'folder nie ma jeszcze uprawnień tylko dla Administratorów i SYSTEM' }
        }
    }
    if ($null -eq $reason) { return }
    Remove-AppFolder -Path $Path
    Write-StdLog -Level Warn -Source 'Sync-StdAppFolder' -Message "Usunięto kopię STD $Path ($reason) - STD skopiuje ją od nowa."
}

function Test-AppOtherInstance {
    # Czy mutex jednej instancji trzyma inny proces (albo wątek) STD. Mutex trzymany przez ten wątek się nie liczy
    # (WaitOne właściciela zawsze się udaje). Nie da się sprawdzić -> $true (ostrożnie: bez usuwania starych wersji).
    $mutex = $null
    try {
        if (-not [System.Threading.Mutex]::TryOpenExisting($script:StdAppMutexName, [ref]$mutex)) { return $false }
        $acquired = $false
        try {
            $acquired = $mutex.WaitOne(0)
        } catch [System.Threading.AbandonedMutexException] {
            # Poprzedni proces STD zakończył się bez zwolnienia mutexu - nikt go już nie trzyma.
            $acquired = $true
        }
        if (-not $acquired) { return $true }
        $mutex.ReleaseMutex()
        return $false
    } catch {
        Write-StdLog -Level Warn -Source 'Sync-StdAppFolder' -Message "Nie udało się sprawdzić, czy działa inny proces STD - $(Format-CoreErrorText $_)"
        return $true
    } finally {
        if ($null -ne $mutex) { $mutex.Dispose() }
    }
}

function Get-AppDriveType {
    # Typ dysku dla litery (np. 'Z:'): 'Network', 'Fixed', 'Removable'... Poza Windows: ''.
    # Osobna funkcja, żeby testy mogły ją podmienić (Mock).
    param([Parameter(Mandatory = $true)][string]$Drive)
    if (-not (Test-CoreIsWindows)) { return '' }
    try {
        return [string](New-Object System.IO.DriveInfo($Drive)).DriveType
    } catch {
        return ''
    }
}

function Get-AppDriveRemotePath {
    # Udział sieciowy (\\serwer\udział) dla litery zmapowanego dysku (np. 'Z:'): Win32_LogicalDisk.ProviderName,
    # a gdy go brak - HKCU:\Network\<litera>\RemotePath (dysk mapowany na stałe). Nie wiadomo -> $null.
    param([Parameter(Mandatory = $true)][string]$Drive)
    $letter = $Drive.Substring(0, 1).ToUpperInvariant()
    try {
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}:'" -f $letter) -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $disk -and -not [string]::IsNullOrWhiteSpace([string]$disk.ProviderName)) { return [string]$disk.ProviderName }
    } catch { }
    try {
        $remote = [string](Get-ItemProperty -LiteralPath ('HKCU:\Network\' + $letter) -Name 'RemotePath' -ErrorAction Stop).RemotePath
        if (-not [string]::IsNullOrWhiteSpace($remote)) { return $remote }
    } catch { }
    return $null
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
    Brak folderu docelowego -> $false. Źródło bez STD.psd1 -> [STD-APP-SOURCE]. Dowiązanie (junction / symlink)
    wśród plików narzędzia -> [STD-APP-COPY].
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
    [version]) i zawsze wersja docelowa; gdy działa inny proces STD (Enter-StdAppInstance), stare wersje zostają.
    -Version domyślnie z <Source>\STD.psd1. Gdy -Source to już folder docelowy, niczego nie kopiuje.
    Bezpieczeństwo (kod z App\ działa jako administrator): App\ jako dowiązanie jest zastępowany zwykłym folderem;
    w Windows w procesie administratora App\ i App\<wersja> dostają uprawnienia z Protect-StdDataFolder. Kopia, która jest
    dowiązaniem, ma dowiązanie w środku albo (administrator, Windows) nie należy do Administratorów / SYSTEM lub nie ma
    jeszcze tych uprawnień, jest usuwana i kopiowana od nowa. Nic poza App\<wersja> nie jest usuwane.
    Błędy: [STD-APP-SOURCE], [STD-APP-COPY].
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
    $protect = Test-CoreDataProtectionNeeded
    try {
        if (Test-CoreIsLink -Path $appRoot) {
            # Junction App\ -> dowolny folder: kopia i usuwanie starych wersji trafiłyby tam.
            Remove-CoreLink -Path $appRoot
            [void][System.IO.Directory]::CreateDirectory($appRoot)
            Write-StdLog -Level Warn -Source 'Sync-StdAppFolder' -Message "Folder $appRoot był dowiązaniem (junction / symlink) - usunięto samo dowiązanie."
        }
        if ($protect) { Protect-StdDataFolder -Path $appRoot }
    } catch {
        throw "[STD-APP-COPY] Nie udało się zabezpieczyć folderu kopii STD $appRoot - $(Format-CoreErrorText $_)"
    }
    $targetRoot = Get-AppNormalPath -Path (Get-StdAppFolder -Version $Version)

    if (Test-AppPathEqual -Path $sourceRoot -Other $targetRoot) {
        Write-StdLog -Level Debug -Source 'Sync-StdAppFolder' -Message "STD już działa z kopii $targetRoot - bez kopiowania."
    } else {
        if (Test-AppPathInside -Path $targetRoot -Parent $sourceRoot) {
            throw "[STD-APP-COPY] Folder kopii $targetRoot leży wewnątrz folderu źródłowego $sourceRoot - ustaw inny folder danych STD."
        }
        try {
            Reset-AppUntrustedFolder -Path $targetRoot -CheckTrust $protect
        } catch {
            throw "[STD-APP-COPY] Nie udało się usunąć niezaufanej kopii STD $targetRoot - $(Format-CoreErrorText $_)"
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
        if ($protect) {
            # Sprawdzona kopia dostaje uprawnienia tylko dla Administratorów i SYSTEM i właściciela Administratorzy (także
            # gdy folder założył proces bez UAC, np. wbudowane konto Administrator): następny start jej zaufa.
            try {
                Protect-StdDataFolder -Path $targetRoot
            } catch {
                Write-StdLog -Level Warn -Source 'Sync-StdAppFolder' -Message (Format-CoreErrorText $_ 600)
            }
        }
    }

    if (Test-AppOtherInstance) {
        # Inny STD może właśnie działać z jednej ze starszych wersji - usunięcie folderu spod niego by go zepsuło.
        Write-StdLog -Level Info -Source 'Sync-StdAppFolder' -Message 'Działa inny proces STD - stare wersje z App\ usunę przy następnym uruchomieniu.'
    } else {
        [void](Remove-AppOldVersions -AppRoot $appRoot -Keep $Keep -Protect @($targetRoot, $sourceRoot))
    }
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

function Get-StdNetworkPath {
    <#
    .SYNOPSIS
    Czy ścieżka leży na dysku sieciowym i jaka jest jej postać UNC: @{ IsNetwork; UncPath }.
    Proces uruchomiony jako administrator nie widzi liter dysków zmapowanych przez zwykłego użytkownika
    (osobna sesja logowania), więc program przekazuje mu ścieżkę UNC. Ścieżka UNC (\\serwer\udział\...)
    zostaje bez zmian; litera dysku zmapowanego (np. Z:\STD) jest zamieniana na \\serwer\udział\STD.
    UncPath = $null, gdy dysk jest sieciowy, ale nie wiadomo, na jaki udział wskazuje. Dysk lokalny,
    pendrive i Linux: IsNetwork = $false.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    if ($Path -match '^[\\/]{2}[?.][\\/]') {
        # Ścieżki urządzeń: \\?\UNC\serwer\udział to udział sieciowy, \\?\C:\... - zwykła ścieżka z literą dysku.
        if ($Path -match '^[\\/]{2}[?.][\\/]UNC[\\/](.+)$') { return @{ IsNetwork = $true; UncPath = ('\\' + $Matches[1]) } }
        $Path = $Path.Substring(4)
    } elseif ($Path -match '^[\\/]{2}[^\\/]') {
        return @{ IsNetwork = $true; UncPath = $Path }
    }
    if ($Path -notmatch '^([A-Za-z]):') { return @{ IsNetwork = $false; UncPath = $null } }
    $drive = $Matches[1].ToUpperInvariant() + ':'
    if ((Get-AppDriveType -Drive $drive) -ne 'Network') { return @{ IsNetwork = $false; UncPath = $null } }
    $remote = Get-AppDriveRemotePath -Drive $drive
    if ([string]::IsNullOrWhiteSpace($remote)) {
        Write-StdLog -Level Warn -Source 'Get-StdNetworkPath' -Message "Nie udało się ustalić udziału sieciowego dysku $drive."
        return @{ IsNetwork = $true; UncPath = $null }
    }
    $rest = $Path.Substring(2)
    if ($rest.Length -gt 0 -and $rest[0] -ne [char]'\' -and $rest[0] -ne [char]'/') { $rest = '\' + $rest }
    return @{ IsNetwork = $true; UncPath = ($remote.TrimEnd('\', '/') + $rest) }
}

function Enter-StdAppInstance {
    <#
    .SYNOPSIS
    Blokada jednej instancji STD na komputerze: mutex Global\STD4-SmartToolForDeployment (wszystkie sesje
    Windows), trzymany przez wątek wywołujący do końca procesu albo do Exit-StdAppInstance.
    $true: ten proces ma blokadę (także gdy już ją miał, albo gdy blokady nie da się utworzyć - wtedy ostrzeżenie
    w logu i praca bez niej); $false: działa już inny proces STD. -WaitSec: ile sekund czekać na zwolnienie
    blokady (proces z kopii w App\<wersja> czeka, bo proces z pendrive'a oddaje ją chwilę po jego uruchomieniu).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [ValidateRange(0, 600)]
        [int]$WaitSec = 0
    )
    if ($null -ne $script:StdAppInstanceMutex) { return $true }
    $mutex = $null
    $acquired = $false
    try {
        $mutex = New-Object System.Threading.Mutex($false, $script:StdAppMutexName)
        try {
            $acquired = $mutex.WaitOne($WaitSec * 1000)
        } catch [System.Threading.AbandonedMutexException] {
            # Poprzedni proces STD zakończył się bez zwolnienia blokady - teraz należy do tego wątku.
            $acquired = $true
        }
    } catch {
        if ($null -ne $mutex) { $mutex.Dispose() }
        Write-StdLog -Level Warn -Source 'Enter-StdAppInstance' -Message "Nie udało się utworzyć blokady jednej instancji STD ($($script:StdAppMutexName)) - $(Format-CoreErrorText $_). STD działa bez niej."
        return $true
    }
    if (-not $acquired) {
        $mutex.Dispose()
        Write-StdLog -Level Warn -Source 'Enter-StdAppInstance' -Message 'STD już działa na tym komputerze (blokada jednej instancji jest zajęta).'
        return $false
    }
    $script:StdAppInstanceMutex = $mutex
    return $true
}

function Exit-StdAppInstance {
    <#
    .SYNOPSIS
    Zwalnia blokadę z Enter-StdAppInstance (np. przed przekazaniem pracy do kopii w App\<wersja>).
    Bez blokady nic nie robi. Wołać z tego samego wątku co Enter-StdAppInstance.
    #>
    [CmdletBinding()]
    param()
    $mutex = $script:StdAppInstanceMutex
    if ($null -eq $mutex) { return }
    $script:StdAppInstanceMutex = $null
    try {
        $mutex.ReleaseMutex()
    } catch {
        Write-StdLog -Level Warn -Source 'Exit-StdAppInstance' -Message "Nie udało się zwolnić blokady jednej instancji STD - $(Format-CoreErrorText $_)"
    } finally {
        $mutex.Dispose()
    }
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
