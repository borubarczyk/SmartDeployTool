# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.Core.ps1 - podstawy STD (SPEC 1 i 3.1): wersja, lokalne katalogi danych, JSON, log z maskowaniem
# sekretów, sprawdzenie uprawnień administratora i ponowne uruchomienie jako administrator.
#
# Plik jest dot-source'owany przez STD.psm1, ale działa też wczytany samodzielnie
# (. .\Modules\Std.Core.ps1) - wtedy wersję czyta z ..\STD.psd1.
#
# Funkcje pomocnicze bez "-Std" w nazwie nie są eksportowane z modułu (Export-ModuleMember *-Std*).

# ---------------------------------------------------------------------------------------------
# Stan wspólny (zakres modułu). Przy ponownym dot-source tego samego pliku nie kasujemy
# zarejestrowanych odbiorników ani masek.
# ---------------------------------------------------------------------------------------------
if (-not (Get-Variable -Name 'StdLogSinks' -Scope Script -ErrorAction SilentlyContinue) -or $null -eq $script:StdLogSinks) {
    $script:StdLogSinks = [ordered]@{}
}
if (-not (Get-Variable -Name 'StdLogFileWriters' -Scope Script -ErrorAction SilentlyContinue) -or $null -eq $script:StdLogFileWriters) {
    $script:StdLogFileWriters = @{}
}
if (-not (Get-Variable -Name 'StdLogMasks' -Scope Script -ErrorAction SilentlyContinue) -or $null -eq $script:StdLogMasks) {
    $script:StdLogMasks = New-Object 'System.Collections.Generic.List[string]'
    $script:StdLogMaskRegex = $null
}
if (-not (Get-Variable -Name 'StdVersionCache' -Scope Script -ErrorAction SilentlyContinue)) {
    $script:StdVersionCache = $null
}
$script:StdCoreFile = $PSCommandPath

# Nagłówki Authorization maskujemy zawsze, także gdy wartość nie została zarejestrowana.
$script:StdAuthHeaderRegex = New-Object System.Text.RegularExpressions.Regex(
    '(?<p>\bAuthorization\s*[:=]\s*(?:Basic|Bearer|Negotiate|NTLM|Digest)?\s*)[^\s"'',;]+',
    ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant))

# ---------------------------------------------------------------------------------------------
# Pomocnicze (prywatne)
# ---------------------------------------------------------------------------------------------

function Test-CoreIsWindows {
    # Bez $IsWindows (nie ma go w Windows PowerShell 5.1).
    return ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
}

function Resolve-CoreFullPath {
    # Ścieżka PowerShella (także względna, TestDrive:\...) -> pełna ścieżka systemu plików dla [IO.File].
    param([Parameter(Mandatory = $true)][string]$Path)
    if ($Path -match '^(\\\\|//)') { return $Path }
    try {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    } catch {
        return [System.IO.Path]::GetFullPath($Path)
    }
}

function New-CoreUtf8NoBom {
    return (New-Object System.Text.UTF8Encoding($false))
}

function Get-CoreManifestVersion {
    # Wersja z ..\STD.psd1 (gdy Std.Core.ps1 jest wczytany samodzielnie albo moduł nie ma wersji).
    $coreFile = $script:StdCoreFile
    if ([string]::IsNullOrEmpty($coreFile)) { $coreFile = $MyInvocation.MyCommand.ScriptBlock.File }
    $candidates = New-Object 'System.Collections.Generic.List[string]'
    if (-not [string]::IsNullOrEmpty($coreFile)) {
        $modulesDir = Split-Path -Parent $coreFile
        $candidates.Add((Join-Path (Split-Path -Parent $modulesDir) 'STD.psd1'))
        $candidates.Add((Join-Path $modulesDir 'STD.psd1'))
    }
    foreach ($manifest in $candidates) {
        if (Test-Path -LiteralPath $manifest -PathType Leaf) {
            $data = Import-PowerShellDataFile -LiteralPath $manifest
            if ($data -and $data.ContainsKey('ModuleVersion') -and -not [string]::IsNullOrWhiteSpace([string]$data['ModuleVersion'])) {
                return [version][string]$data['ModuleVersion']
            }
        }
    }
    return $null
}

function ConvertTo-CoreHashtableNode {
    # Rekurencja ConvertTo-StdHashtable (zwykła funkcja - szybsza niż zaawansowana).
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [byte[]]) { return , $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        # GetEnumerator zamiast .Keys: klucz o nazwie "Keys" / "Count" zasłoniłby właściwość.
        # Add zamiast indeksu: [ordered] ma też indeks liczbowy (pozycja), więc klucz 1 trafiłby w pozycję.
        $result = [ordered]@{}
        foreach ($entry in @($Value.GetEnumerator())) {
            $converted = ConvertTo-CoreHashtableNode -Value $entry.Value
            if ($result.Contains($entry.Key)) { $result.Remove($entry.Key) }
            $result.Add($entry.Key, $converted)
        }
        return $result
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $result = [ordered]@{}
        foreach ($prop in $Value.PSObject.Properties) {
            $result[$prop.Name] = ConvertTo-CoreHashtableNode -Value $prop.Value
        }
        return $result
    }
    if ($Value -is [System.Collections.IList] -or $Value -is [System.Collections.ICollection]) {
        # Bez @($Value): dla List[T] w 5.1 daje błąd "Argument types do not match".
        $buffer = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Value) { $buffer.Add((ConvertTo-CoreHashtableNode -Value $item)) }
        return , $buffer.ToArray()
    }
    if ($Value -is [System.Management.Automation.PSObject]) {
        # Opakowany obiekt bez własnych pól (np. tekst z Get-Content) - zwracamy obiekt bazowy.
        return $Value.PSObject.BaseObject
    }
    return $Value
}

function ConvertTo-CoreJsonNode {
    # Przygotowanie danych do ConvertTo-Json, żeby wynik był taki sam w 5.1 i 7:
    # - PSCustomObject / IDictionary -> [ordered], tablice -> object[] (bez opakowania PSObject,
    #   które w 5.1 dawało {"value":[...],"Count":n}),
    # - DateTime / DateTimeOffset -> tekst ISO 8601 ('o'), [version] -> tekst.
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Management.Automation.PSObject] -and -not ($Value -is [System.Management.Automation.PSCustomObject])) {
        $Value = $Value.PSObject.BaseObject
    }
    if ($Value -is [string]) { return [string]$Value }
    if ($Value -is [datetime]) { return $Value.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [System.DateTimeOffset]) { return $Value.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [version]) { return $Value.ToString() }
    if ($Value -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($entry in @($Value.GetEnumerator())) {
            $result[[string]$entry.Key] = ConvertTo-CoreJsonNode -Value $entry.Value
        }
        return $result
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $result = [ordered]@{}
        foreach ($prop in $Value.PSObject.Properties) {
            $result[$prop.Name] = ConvertTo-CoreJsonNode -Value $prop.Value
        }
        return $result
    }
    if ($Value -is [System.Collections.IList] -or $Value -is [System.Collections.ICollection]) {
        # Bez @($Value): dla List[T] w 5.1 daje błąd "Argument types do not match".
        $buffer = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Value) { $buffer.Add((ConvertTo-CoreJsonNode -Value $item)) }
        return , $buffer.ToArray()
    }
    return $Value
}

function Get-CoreNewtonsoftReaderType {
    # PowerShell 7: parser Newtonsoft bez zamiany dat na DateTime (ConvertFrom-Json w 7 zamienia
    # "2026-11-01T20:00:00Z" na [datetime], a w 5.1 zostawia tekst - chcemy zawsze tekst).
    if ($PSVersionTable.PSEdition -ne 'Core') { return $null }
    $type = 'Newtonsoft.Json.JsonTextReader' -as [type]
    if ($null -eq $type) {
        try { $null = ConvertFrom-Json -InputObject '{}' } catch { }
        $type = 'Newtonsoft.Json.JsonTextReader' -as [type]
    }
    return $type
}

function ConvertFrom-CoreJToken {
    # Uwaga: na JObject nie używamy .Type ani .Count - PowerShell traktuje JObject jak słownik
    # i np. klucz "type" z JSON zasłania właściwość. Dlatego rozpoznajemy typ operatorem -is.
    param($Token)
    if ($null -eq $Token) { return $null }
    if ($Token -is [Newtonsoft.Json.Linq.JObject]) {
        $result = [ordered]@{}
        foreach ($prop in $Token.Properties()) {
            $result[$prop.Name] = ConvertFrom-CoreJToken -Token $prop.Value
        }
        return $result
    }
    if ($Token -is [Newtonsoft.Json.Linq.JArray]) {
        $children = @($Token.Children())
        $array = New-Object object[] $children.Count
        for ($i = 0; $i -lt $children.Count; $i++) {
            $array[$i] = ConvertFrom-CoreJToken -Token $children[$i]
        }
        return , $array
    }
    if ($Token -is [Newtonsoft.Json.Linq.JValue]) {
        $v = $Token.Value
        if ($null -eq $v) { return $null }
        if ($v -is [long] -and $v -ge [int]::MinValue -and $v -le [int]::MaxValue) { return [int]$v }
        return $v
    }
    return $null
}

function ConvertFrom-CoreJsonText {
    param([string]$Json)
    $readerType = Get-CoreNewtonsoftReaderType
    if ($null -ne $readerType) {
        $stringReader = New-Object System.IO.StringReader($Json)
        $reader = New-Object Newtonsoft.Json.JsonTextReader($stringReader)
        try {
            $reader.DateParseHandling = [Newtonsoft.Json.DateParseHandling]::None
            $reader.FloatParseHandling = [Newtonsoft.Json.FloatParseHandling]::Double
            $reader.MaxDepth = 1024
            $token = [Newtonsoft.Json.Linq.JToken]::ReadFrom($reader)
            while ($reader.Read()) {
                if ($reader.TokenType -ne [Newtonsoft.Json.JsonToken]::Comment) {
                    throw 'Za zakończonym dokumentem JSON jest dodatkowy tekst.'
                }
            }
        } finally {
            $reader.Close()
            $stringReader.Dispose()
        }
        $converted = ConvertFrom-CoreJToken -Token $token
        return , $converted
    }
    # Windows PowerShell 5.1: ConvertFrom-Json nie rozwija tablicy z najwyższego poziomu i nie zmienia dat.
    $parsed = ConvertFrom-Json -InputObject $Json -ErrorAction Stop
    $converted = ConvertTo-CoreHashtableNode -Value $parsed
    return , $converted
}

function Format-CoreErrorText {
    # Krótki tekst wyjątku do komunikatu (bez długich śladów), zamaskowany.
    param($ErrorObject, [int]$MaxLength = 300)
    $text = ''
    $ex = $null
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ex = $ErrorObject.Exception }
    elseif ($ErrorObject -is [System.Exception]) { $ex = $ErrorObject }
    if ($null -ne $ex) {
        # "Exception calling ... with N argument(s)" to opakowanie PowerShella - bierzemy wyjątek ze środka
        # i dopisujemy najgłębszą przyczynę (np. błąd certyfikatu), jeśli jej jeszcze nie ma w tekście.
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $text = [string]$ex.Message
        $innermost = $ex
        while ($null -ne $innermost.InnerException) { $innermost = $innermost.InnerException }
        if (-not [object]::ReferenceEquals($innermost, $ex) -and -not [string]::IsNullOrEmpty($innermost.Message) -and -not $text.Contains($innermost.Message)) {
            $text = $text + ' ' + $innermost.Message
        }
    } else {
        $text = [string]$ErrorObject
    }
    if ($null -eq $text) { $text = '' }
    $text = ($text -replace '\s+', ' ').Trim()
    if ($text.Length -gt $MaxLength) { $text = $text.Substring(0, $MaxLength) + '...' }
    return (Protect-StdLogText -Text $text)
}

function ConvertTo-CoreCommandLineArgument {
    # Cytowanie argumentu dla wiersza poleceń Windows (reguły CommandLineToArgvW / MSVCRT).
    param([AllowEmptyString()][AllowNull()][string]$Argument)
    if ($null -eq $Argument) { $Argument = '' }
    if ($Argument.Length -gt 0 -and $Argument -notmatch '[\s"]') { return $Argument }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    $backslashes = 0
    foreach ($ch in $Argument.ToCharArray()) {
        if ($ch -eq [char]'\') {
            $backslashes++
            continue
        }
        if ($ch -eq [char]'"') {
            [void]$sb.Append([char]'\', ($backslashes * 2 + 1))
            [void]$sb.Append('"')
        } else {
            if ($backslashes -gt 0) { [void]$sb.Append([char]'\', $backslashes) }
            [void]$sb.Append($ch)
        }
        $backslashes = 0
    }
    if ($backslashes -gt 0) { [void]$sb.Append([char]'\', ($backslashes * 2)) }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Get-CoreElevationArgumentLine {
    param([Parameter(Mandatory = $true)][string]$ScriptPath, [string[]]$Arguments)
    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($fixed in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File')) { $parts.Add($fixed) }
    # Ścieżkę skryptu zawsze w cudzysłowie (SPEC 3.1).
    $parts.Add('"' + ((ConvertTo-CoreCommandLineArgument -Argument $ScriptPath).Trim('"')) + '"')
    if ($null -ne $Arguments) {
        foreach ($arg in $Arguments) { $parts.Add((ConvertTo-CoreCommandLineArgument -Argument $arg)) }
    }
    return ($parts -join ' ')
}

function Get-CorePowerShellExePath {
    if (-not [string]::IsNullOrEmpty($env:SystemRoot)) {
        $candidate = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return 'powershell.exe'
}

function Start-CoreElevatedProcess {
    # Start procesu z czasownikiem RunAs (okno UAC). Osobna funkcja, żeby testy mogły ją podmienić (Mock).
    # Process.Start zamiast Start-Process: Start-Process w 5.1 gubi kod błędu 1223 (odmowa UAC)
    # i zostawia tylko przetłumaczony komunikat.
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ArgumentLine,
        [string]$WorkingDirectory
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $ArgumentLine
    $psi.Verb = 'RunAs'
    $psi.UseShellExecute = $true
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $psi.WorkingDirectory = $WorkingDirectory }
    $process = [System.Diagnostics.Process]::Start($psi)
    if ($null -ne $process) { $process.Dispose() }
}

function Test-CoreIsUacCancel {
    # Czy wyjątek (lub któryś wewnętrzny) to anulowanie UAC (ERROR_CANCELLED = 1223).
    param($Exception)
    $ex = $Exception
    while ($null -ne $ex) {
        if ($ex -is [System.ComponentModel.Win32Exception] -and $ex.NativeErrorCode -eq 1223) { return $true }
        if ($ex.HResult -eq 1223 -or $ex.HResult -eq -2147023673) { return $true }
        $ex = $ex.InnerException
    }
    return $false
}

# ---------------------------------------------------------------------------------------------
# Wersja i katalogi danych
# ---------------------------------------------------------------------------------------------

function Get-StdVersion {
    <#
    .SYNOPSIS
    Zwraca wersję STD ([version]) z manifestu modułu STD.psd1.
    #>
    [CmdletBinding()]
    [OutputType([version])]
    param()
    if ($null -ne $script:StdVersionCache) { return $script:StdVersionCache }
    $module = $MyInvocation.MyCommand.Module
    if ($null -ne $module -and $module.Name -eq 'STD' -and $null -ne $module.Version -and $module.Version -gt [version]'0.0') {
        $script:StdVersionCache = [version]$module.Version
        return $script:StdVersionCache
    }
    $fromManifest = Get-CoreManifestVersion
    if ($null -ne $fromManifest) {
        $script:StdVersionCache = $fromManifest
        return $script:StdVersionCache
    }
    throw '[STD-VERSION] Nie udało się odczytać wersji STD: brak pliku STD.psd1 obok katalogu Modules.'
}

function Get-StdDataRoot {
    <#
    .SYNOPSIS
    Katalog lokalnych danych STD: $env:STD_DATA_ROOT, inaczej $env:ProgramData\STD, inaczej katalog tymczasowy\STD.
    Tworzy katalog, jeśli go nie ma.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()
    if (-not [string]::IsNullOrWhiteSpace($env:STD_DATA_ROOT)) {
        $root = $env:STD_DATA_ROOT
    } elseif (-not [string]::IsNullOrWhiteSpace($env:ProgramData)) {
        $root = Join-Path $env:ProgramData 'STD'
    } else {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) 'STD'
    }
    $root = Resolve-CoreFullPath -Path $root
    if (-not [System.IO.Directory]::Exists($root)) {
        try {
            [void][System.IO.Directory]::CreateDirectory($root)
        } catch {
            throw "[STD-IO] Nie udało się utworzyć katalogu danych STD $root - $(Format-CoreErrorText $_)"
        }
    }
    return $root
}

function Get-StdPath {
    <#
    .SYNOPSIS
    Pełna ścieżka lokalnych danych STD (SPEC 1). Katalogi tworzy, plików nie.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('Root', 'App', 'Cache', 'Logs', 'Outbox', 'State', 'Revisions', 'Prefs', 'Bootstrap')]
        [string]$Kind
    )
    $root = Get-StdDataRoot
    $dirs = @{ App = 'App'; Cache = 'Cache'; Logs = 'Logs'; Outbox = 'Outbox' }
    $files = @{ State = 'state.json'; Revisions = 'revisions.json'; Prefs = 'prefs.json'; Bootstrap = 'bootstrap.local.json' }
    if ($Kind -eq 'Root') { return $root }
    if ($dirs.ContainsKey($Kind)) {
        $path = Join-Path $root $dirs[$Kind]
        if (-not [System.IO.Directory]::Exists($path)) {
            try {
                [void][System.IO.Directory]::CreateDirectory($path)
            } catch {
                throw "[STD-IO] Nie udało się utworzyć katalogu $path - $(Format-CoreErrorText $_)"
            }
        }
        return $path
    }
    return (Join-Path $root $files[$Kind])
}

# ---------------------------------------------------------------------------------------------
# JSON
# ---------------------------------------------------------------------------------------------

function ConvertTo-StdHashtable {
    <#
    .SYNOPSIS
    Głęboka zamiana PSCustomObject / IDictionary na [ordered]@{} i tablic na object[] (także jedno-
    elementowych i pustych). Pozostałe wartości bez zmian (byte[] zostaje byte[]).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object]$InputObject
    )
    process {
        $result = ConvertTo-CoreHashtableNode -Value $InputObject
        if ($result -is [array]) { return , $result }
        return $result
    }
}

function ConvertFrom-StdJson {
    <#
    .SYNOPSIS
    Tekst JSON -> [ordered]@{} / object[] / wartości proste. Daty zostają tekstem.
    Błąd składni -> [STD-JSON].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Json
    )
    begin { $lines = New-Object 'System.Collections.Generic.List[string]' }
    process { if ($null -ne $Json) { $lines.Add($Json) } }
    end {
        $text = ($lines -join "`n")
        if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
        if ([string]::IsNullOrWhiteSpace($text)) {
            throw '[STD-JSON] Nieprawidłowy JSON: pusty tekst.'
        }
        try {
            $result = ConvertFrom-CoreJsonText -Json $text
        } catch {
            throw "[STD-JSON] Nieprawidłowy JSON: $(Format-CoreErrorText $_ 200)"
        }
        if ($result -is [array]) { return , $result }
        return $result
    }
}

function ConvertTo-StdJson {
    <#
    .SYNOPSIS
    Obiekt -> tekst JSON (ConvertTo-Json -Depth 30). Daty jako ISO 8601, tablice zawsze jako tablice.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object]$InputObject,
        [switch]$Compress
    )
    if ($null -eq $InputObject) { return 'null' }
    $normalized = ConvertTo-CoreJsonNode -Value $InputObject
    if ($Compress) {
        return (ConvertTo-Json -InputObject $normalized -Depth 30 -Compress)
    }
    return (ConvertTo-Json -InputObject $normalized -Depth 30)
}

function Write-StdJsonFile {
    <#
    .SYNOPSIS
    Atomowy zapis JSON (UTF-8 bez BOM): najpierw <path>.tmp, potem podmiana pliku.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path,
        [Parameter(Mandatory = $true, Position = 1)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object]$InputObject
    )
    $full = Resolve-CoreFullPath -Path $Path
    $json = ConvertTo-StdJson -InputObject $InputObject
    $tmp = $full + '.tmp'
    try {
        $dir = [System.IO.Path]::GetDirectoryName($full)
        if (-not [string]::IsNullOrEmpty($dir) -and -not [System.IO.Directory]::Exists($dir)) {
            [void][System.IO.Directory]::CreateDirectory($dir)
        }
        [System.IO.File]::WriteAllText($tmp, $json, (New-CoreUtf8NoBom))
        if ([System.IO.File]::Exists($full)) {
            try {
                [System.IO.File]::Replace($tmp, $full, [NullString]::Value)
            } catch {
                # Replace nie działa np. na niektórych udziałach sieciowych - zwykła podmiana.
                [System.IO.File]::Copy($tmp, $full, $true)
                [System.IO.File]::Delete($tmp)
            }
        } else {
            [System.IO.File]::Move($tmp, $full)
        }
    } catch {
        try { if ([System.IO.File]::Exists($tmp)) { [System.IO.File]::Delete($tmp) } } catch { }
        throw "[STD-IO] Nie udało się zapisać pliku $full - $(Format-CoreErrorText $_)"
    }
}

function Read-StdJsonFile {
    <#
    .SYNOPSIS
    Odczyt pliku JSON. Brak pliku -> $null. Uszkodzony JSON -> plik przemianowany na <path>.bad,
    ostrzeżenie w logu, wynik $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path
    )
    $full = Resolve-CoreFullPath -Path $Path
    if (-not [System.IO.File]::Exists($full)) { return $null }
    try {
        $text = [System.IO.File]::ReadAllText($full, (New-CoreUtf8NoBom))
    } catch {
        throw "[STD-IO] Nie udało się odczytać pliku $full - $(Format-CoreErrorText $_)"
    }
    try {
        $result = ConvertFrom-StdJson -Json $text
    } catch {
        $bad = $full + '.bad'
        $reason = Format-CoreErrorText $_ 200
        try {
            if ([System.IO.File]::Exists($bad)) { [System.IO.File]::Delete($bad) }
            [System.IO.File]::Move($full, $bad)
            Write-StdLog -Level Warn -Source 'Read-StdJsonFile' -Message "Uszkodzony plik JSON $full przeniesiono do $bad ($reason)."
        } catch {
            Write-StdLog -Level Warn -Source 'Read-StdJsonFile' -Message "Uszkodzony plik JSON $full ($reason). Nie udało się zmienić jego nazwy: $(Format-CoreErrorText $_)"
        }
        return $null
    }
    if ($result -is [array]) { return , $result }
    return $result
}

# ---------------------------------------------------------------------------------------------
# Log i maski
# ---------------------------------------------------------------------------------------------

function Register-StdLogMask {
    <#
    .SYNOPSIS
    Dodaje wartość (np. sekret) do maski logu. Wartości krótsze niż 3 znaki są ignorowane.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )
    process {
        if ([string]::IsNullOrEmpty($Value) -or $Value.Length -lt 3) { return }
        foreach ($existing in $script:StdLogMasks) {
            if ([string]::Equals($existing, $Value, [System.StringComparison]::OrdinalIgnoreCase)) { return }
        }
        $script:StdLogMasks.Add($Value)
        # Najdłuższe najpierw - alternatywa w wyrażeniu regularnym jest sprawdzana po kolei.
        $sorted = @($script:StdLogMasks | Sort-Object -Property Length -Descending)
        $pattern = (@($sorted | ForEach-Object { [regex]::Escape($_) }) -join '|')
        $script:StdLogMaskRegex = New-Object System.Text.RegularExpressions.Regex(
            $pattern,
            ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant))
    }
}

function Clear-StdLogMasks {
    <#
    .SYNOPSIS
    Usuwa wszystkie wartości z maski logu.
    #>
    [CmdletBinding()]
    param()
    $script:StdLogMasks.Clear()
    $script:StdLogMaskRegex = $null
}

function Protect-StdLogText {
    <#
    .SYNOPSIS
    Zamienia każdą zarejestrowaną wartość (najdłuższe najpierw, bez względu na wielkość liter) na ***.
    Wartość nagłówka Authorization jest maskowana zawsze.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Text
    )
    process {
        if ([string]::IsNullOrEmpty($Text)) { return $Text }
        $out = $Text
        if ($null -ne $script:StdLogMaskRegex) {
            $out = $script:StdLogMaskRegex.Replace($out, '***')
        }
        $out = $script:StdAuthHeaderRegex.Replace($out, '${p}***')
        return $out
    }
}

function Register-StdLogSink {
    <#
    .SYNOPSIS
    Rejestruje odbiornik logu. Odbiornik dostaje rekord @{ Time; Level; Message; Source } jako $args[0].
    Ta sama nazwa podmienia poprzedni odbiornik.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory = $true, Position = 1)]
        [scriptblock]$ScriptBlock
    )
    if ($script:StdLogSinks.Contains($Name)) { Unregister-StdLogSink -Name $Name }
    $script:StdLogSinks[$Name] = $ScriptBlock
}

function Unregister-StdLogSink {
    <#
    .SYNOPSIS
    Usuwa odbiornik logu; odbiornik plikowy zamyka plik. Nieznana nazwa nie jest błędem.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Name
    )
    if ($script:StdLogSinks.Contains($Name)) { $script:StdLogSinks.Remove($Name) }
    if ($script:StdLogFileWriters.ContainsKey($Name)) {
        $writer = $script:StdLogFileWriters[$Name]
        $script:StdLogFileWriters.Remove($Name)
        try {
            try { $writer.Flush() } finally { $writer.Dispose() }
        } catch { }
    }
}

function Register-StdFileLogSink {
    <#
    .SYNOPSIS
    Odbiornik zapisujący log do pliku (UTF-8 bez BOM, dopisywanie, AutoFlush), linia:
    yyyy-MM-dd HH:mm:ss [LEVEL] wiadomość
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    $full = Resolve-CoreFullPath -Path $Path
    if ($script:StdLogSinks.Contains($Name) -or $script:StdLogFileWriters.ContainsKey($Name)) { Unregister-StdLogSink -Name $Name }
    try {
        $dir = [System.IO.Path]::GetDirectoryName($full)
        if (-not [string]::IsNullOrEmpty($dir) -and -not [System.IO.Directory]::Exists($dir)) {
            [void][System.IO.Directory]::CreateDirectory($dir)
        }
        # Udostępnienie ReadWrite: okno dziennika może czytać plik na bieżąco.
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($full, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, $share)
        $writer = New-Object System.IO.StreamWriter($stream, (New-CoreUtf8NoBom))
        $writer.AutoFlush = $true
    } catch {
        throw "[STD-IO] Nie udało się otworzyć pliku logu $full - $(Format-CoreErrorText $_)"
    }
    $script:StdLogFileWriters[$Name] = $writer
    $fileWriter = $writer
    $fileStream = $stream
    $sink = {
        $record = $args[0]
        $time = $record.Time
        if ($time -isnot [datetime]) { $time = Get-Date }
        $line = '{0} [{1}] {2}' -f $time.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture), ([string]$record.Level).ToUpperInvariant(), $record.Message
        # Przed każdym zapisem na koniec pliku: ten sam plik może mieć kilku piszących
        # (np. drugi wątek wdrożenia), a FileMode.Append ustawia pozycję tylko przy otwarciu.
        [void]$fileStream.Seek(0, [System.IO.SeekOrigin]::End)
        $fileWriter.WriteLine($line)
    }.GetNewClosure()
    $script:StdLogSinks[$Name] = $sink
}

function Write-StdLog {
    <#
    .SYNOPSIS
    Zapis do logu: maskuje tekst i przekazuje rekord @{ Time; Level; Message; Source } do każdego
    odbiornika. Błąd odbiornika nie przerywa pracy. Debug tylko gdy $env:STD_DEBUG = '1'.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Message,
        [ValidateSet('Info', 'Warn', 'Error', 'Debug')]
        [string]$Level = 'Info',
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Source = ''
    )
    switch ($Level) {
        'Info' { $Level = 'Info' }
        'Warn' { $Level = 'Warn' }
        'Error' { $Level = 'Error' }
        'Debug' { $Level = 'Debug' }
    }
    if ($Level -eq 'Debug' -and $env:STD_DEBUG -ne '1') { return }
    if ($null -eq $Message) { $Message = '' }
    if ($null -eq $Source) { $Source = '' }
    $record = @{
        Time    = Get-Date
        Level   = $Level
        Message = (Protect-StdLogText -Text $Message)
        Source  = (Protect-StdLogText -Text $Source)
    }
    Write-Verbose ('[{0}] {1}' -f $Level.ToUpperInvariant(), $record.Message)
    foreach ($sink in @($script:StdLogSinks.get_Values())) {
        try {
            $null = & $sink $record
        } catch {
            # Odbiornik nie może przerwać pracy narzędzia.
        }
    }
}

# ---------------------------------------------------------------------------------------------
# Uprawnienia administratora
# ---------------------------------------------------------------------------------------------

function Test-StdIsAdmin {
    <#
    .SYNOPSIS
    Czy bieżący proces ma uprawnienia administratora. Na Linuksie: id -u = 0.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    if (Test-CoreIsWindows) {
        $identity = $null
        try {
            $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
            return [bool]$principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
        } catch {
            return $false
        } finally {
            if ($null -ne $identity) { $identity.Dispose() }
        }
    }
    try {
        $uid = (& id -u 2>$null | Select-Object -First 1)
        return ([string]$uid).Trim() -eq '0'
    } catch {
        return $false
    }
}

function Start-StdElevated {
    <#
    .SYNOPSIS
    Uruchamia skrypt ponownie jako administrator:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "<ScriptPath>" <argumenty> (z -Verb RunAs).
    Zwraca $true, gdy proces wystartował, $false, gdy użytkownik odmówił w oknie UAC.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptPath,
        [Parameter(Position = 1)]
        [AllowEmptyCollection()]
        [string[]]$Arguments = @(),
        [string]$WorkingDirectory
    )
    $argumentLine = Get-CoreElevationArgumentLine -ScriptPath $ScriptPath -Arguments $Arguments
    $startParams = @{
        FilePath     = (Get-CorePowerShellExePath)
        ArgumentLine = $argumentLine
    }
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $startParams.WorkingDirectory = $WorkingDirectory }
    try {
        Start-CoreElevatedProcess @startParams
    } catch {
        $err = $_
        if ((Test-CoreIsUacCancel -Exception $err.Exception)) {
            Write-StdLog -Level Warn -Source 'Start-StdElevated' -Message 'Użytkownik nie zgodził się na uruchomienie jako administrator (UAC).'
            return $false
        }
        throw "[STD-ELEVATE] Nie udało się uruchomić STD jako administrator: $(Format-CoreErrorText $err)"
    }
    Write-StdLog -Level Info -Source 'Start-StdElevated' -Message 'Uruchomiono STD ponownie z uprawnieniami administratora.'
    return $true
}
