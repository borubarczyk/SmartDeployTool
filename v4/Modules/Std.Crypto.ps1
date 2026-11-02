# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.Crypto.ps1 - kryptografia STD (SPEC 3.3): suma SHA-256, podpisana koperta *.sdt (SPEC 2.1),
# zaszyfrowane sekrety konfiguracji (SPEC 2.7) i podstawianie {{secret:nazwa}}.
#
# Wymaga Std.Core.ps1 (Write-StdLog, Register-StdLogMask) - STD.psm1 wczytuje go wcześniej.
# Get-StdTrustedCertificates (Std.Config.ps1) jest wołana tylko wtedy, gdy nie podano -TrustedCertificates.
#
# Koperta podpisu (*.sdt):
#   { "kind": "std.signed", "alg": "RS256", "kid": "<odcisk SHA-1 certyfikatu, HEX wielkimi literami>",
#     "payload": "<base64 bajtów dokumentu>", "sig": "<base64 podpisu RSA PKCS#1 v1.5 SHA-256 bajtów payload>" }
#   Podpisujemy dokładne bajty payload (bez kanonizacji JSON). Koperta jest zaufana tylko wtedy, gdy kid jest
#   na liście zaufanych certyfikatów I podpis zgadza się z kluczem publicznym tego certyfikatu.
#
# Sekrety (tryb sealed):
#   klucz = PBKDF2-SHA256(UTF-8 hasła wdrożeniowego, salt, iter) -> 64 bajty: encKey = 0-31, macKey = 32-63;
#   check = HMAC-SHA256(macKey, 'std-check'); pozycja = AES-256-CBC (PKCS7, losowe IV) + mac = HMAC-SHA256(macKey, iv||ct).
#   Najpierw sprawdzamy mac (porównanie w stałym czasie), dopiero potem odszyfrowujemy.
#   W sekretach nie ma i nie będzie loginu ani hasła do domeny ani hasła lokalnego administratora.
#
# Zasady:
# - jawny tekst z SecureString wyłącznie przez Marshal.SecureStringToBSTR / ZeroFreeBSTR;
# - tablice bajtów z hasłem, kluczami i odszyfrowaną wartością zerujemy zaraz po użyciu;
# - żadna funkcja nie zapisuje haseł ani wartości sekretów do logu; odszyfrowana albo podstawiona
#   wartość od razu trafia do maski logu (Register-StdLogMask).
#
# Funkcje pomocnicze bez "-Std" w nazwie nie są eksportowane z modułu (Export-ModuleMember *-Std*).

# ---------------------------------------------------------------------------------------------
# Stałe
# ---------------------------------------------------------------------------------------------

$script:StdCryptoEnvelopeKind = 'std.signed'
$script:StdCryptoEnvelopeAlg = 'RS256'
$script:StdCryptoSealedMode = 'sealed'
$script:StdCryptoKdfAlg = 'pbkdf2-sha256'
$script:StdCryptoCheckText = 'std-check'
$script:StdCryptoMinSigningKeySize = 2048

# Nazwa sekretu: litera lub cyfra na początku, potem litery, cyfry, '.', '_', '-' (razem do 64 znaków).
$script:StdCryptoSecretNamePattern = '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$'
$script:StdCryptoSecretRefRegex = New-Object System.Text.RegularExpressions.Regex(
    '\{\{secret:(?<name>[A-Za-z0-9][A-Za-z0-9._-]{0,63})\}\}',
    [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)

# ---------------------------------------------------------------------------------------------
# Pomocnicze (prywatne)
# ---------------------------------------------------------------------------------------------

function Resolve-CryptoFullPath {
    # Ścieżka PowerShella (także względna, TestDrive:\...) -> pełna ścieżka systemu plików dla [IO.File].
    param([Parameter(Mandatory = $true)][string]$Path)
    if ($Path -match '^(\\\\|//)') { return $Path }
    try {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    } catch {
        return [System.IO.Path]::GetFullPath($Path)
    }
}

function ConvertTo-CryptoHex {
    # Bajty -> HEX małymi literami.
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)
    if ($Bytes.Length -eq 0) { return '' }
    return ([System.BitConverter]::ToString($Bytes).Replace('-', '').ToLowerInvariant())
}

function New-CryptoSha256 {
    # Przy wymuszonym trybie FIPS (zasady grupy) SHA256Managed z .NET Framework rzuca wyjątek,
    # więc wtedy bierzemy implementację systemową (CryptoServiceProvider).
    if ([System.Security.Cryptography.CryptoConfig]::AllowOnlyFipsAlgorithms) {
        return (New-Object System.Security.Cryptography.SHA256CryptoServiceProvider)
    }
    return [System.Security.Cryptography.SHA256]::Create()
}

function New-CryptoRandomBytes {
    param([Parameter(Mandatory = $true)][int]$Count)
    $buffer = New-Object byte[] $Count
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($buffer)
    } finally {
        $rng.Dispose()
    }
    return , $buffer
}

function Clear-CryptoBytes {
    # Zeruje tablicę bajtów (lub znaków); $null jest pomijany.
    param([AllowNull()]$Buffer)
    if ($null -ne $Buffer -and $Buffer -is [array] -and $Buffer.Length -gt 0) {
        [System.Array]::Clear($Buffer, 0, $Buffer.Length)
    }
}

function Join-CryptoBytes {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$First,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Second
    )
    $result = New-Object byte[] ($First.Length + $Second.Length)
    [System.Buffer]::BlockCopy($First, 0, $result, 0, $First.Length)
    [System.Buffer]::BlockCopy($Second, 0, $result, $First.Length, $Second.Length)
    return , $result
}

function Test-CryptoBytesEqual {
    # Porównanie w stałym czasie (bez wczesnego wyjścia przy pierwszej różnicy). Długość MAC nie jest
    # tajna (zawsze 32 bajty), więc różna długość od razu daje $false, ale pętla i tak przechodzi całość.
    param([AllowNull()][byte[]]$Expected, [AllowNull()][byte[]]$Actual)
    if ($null -eq $Expected -or $null -eq $Actual) { return $false }
    $diff = $Expected.Length -bxor $Actual.Length
    $count = [System.Math]::Min($Expected.Length, $Actual.Length)
    for ($i = 0; $i -lt $count; $i++) {
        $diff = $diff -bor ($Expected[$i] -bxor $Actual[$i])
    }
    return ($diff -eq 0)
}

function ConvertTo-CryptoUtf8Bytes {
    # SecureString -> bajty UTF-8. Jawny tekst tylko w buforze BSTR (zerowanym przez ZeroFreeBSTR)
    # i w tablicy znaków (zerowanej tutaj); bez tworzenia obiektu [string] z hasłem.
    # Wynik (byte[]) wywołujący musi wyzerować po użyciu.
    param([AllowNull()][System.Security.SecureString]$SecureString)
    if ($null -eq $SecureString -or $SecureString.Length -eq 0) {
        return , (New-Object byte[] 0)
    }
    $bstr = [System.IntPtr]::Zero
    $chars = $null
    try {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        $chars = New-Object char[] $SecureString.Length
        [System.Runtime.InteropServices.Marshal]::Copy($bstr, $chars, 0, $chars.Length)
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($chars)
        return , $bytes
    } finally {
        Clear-CryptoBytes $chars
        if ($bstr -ne [System.IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function ConvertFrom-CryptoSecureString {
    # SecureString -> [string] (tylko przez BSTR, bufor zerowany po odczycie).
    param([Parameter(Mandatory = $true)][System.Security.SecureString]$SecureString)
    $bstr = [System.IntPtr]::Zero
    try {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        if ($bstr -ne [System.IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function ConvertFrom-CryptoUtf8 {
    # Bajty UTF-8 -> tekst, bez znacznika BOM na początku. -Strict: błędne bajty UTF-8 -> wyjątek.
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes,
        [switch]$Strict
    )
    $encoding = New-Object System.Text.UTF8Encoding($false, [bool]$Strict)
    $offset = 0
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { $offset = 3 }
    return $encoding.GetString($Bytes, $offset, $Bytes.Length - $offset)
}

function Get-CryptoSafeText {
    # Tekst z niezaufanej koperty do komunikatu: krótki, tylko bezpieczne znaki.
    param([AllowNull()][AllowEmptyString()][string]$Text, [int]$MaxLength = 32)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $safe = $Text -replace '[^A-Za-z0-9._+/=-]', '?'
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength) + '...' }
    return $safe
}

function Get-CryptoMember {
    # Odczyt pola z IDictionary (hashtable, [ordered]) albo PSCustomObject; nazwa bez względu na wielkość liter.
    param([AllowNull()]$InputObject, [Parameter(Mandatory = $true)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in @($InputObject.Keys)) {
            if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $InputObject[$key]
            }
        }
        return $null
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Find-CryptoDictionaryKey {
    # Istniejący klucz słownika (bez względu na wielkość liter) albo $null.
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.IDictionary]$Dictionary, [Parameter(Mandatory = $true)][string]$Name)
    foreach ($key in @($Dictionary.Keys)) {
        if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) { return $key }
    }
    return $null
}

function ConvertFrom-CryptoBase64Field {
    # Pole base64 -> bajty; błąd -> wyjątek z podanym kodem. -Length: wymagana liczba bajtów (-1 = dowolna, > 0).
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message,
        [int]$Length = -1
    )
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { throw "[$Code] $Message" }
    try {
        $bytes = [System.Convert]::FromBase64String($Value)
    } catch {
        throw "[$Code] $Message"
    }
    if ($bytes.Length -eq 0) { throw "[$Code] $Message" }
    if ($Length -gt 0 -and $bytes.Length -ne $Length) { throw "[$Code] $Message" }
    return , $bytes
}

function Get-CryptoIntegerValue {
    # Liczba całkowita z JSON (Int32 / Int64 / Double bez części ułamkowej) albo $null.
    param([AllowNull()]$Value)
    if ($null -eq $Value -or $Value -is [bool] -or $Value -is [string]) { return $null }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte] -or
        $Value -is [uint32] -or $Value -is [uint64] -or $Value -is [decimal] -or $Value -is [double] -or $Value -is [single]) {
        $number = [double]$Value
        if ([System.Math]::Floor($number) -eq $number -and $number -ge [int]::MinValue -and $number -le [int]::MaxValue) {
            return [int]$number
        }
    }
    return $null
}

function Get-CryptoRsaKeySize {
    # Długość klucza RSA w bitach (0 = nie udało się ustalić). Odczyt przez właściwość klasy bazowej,
    # bo w .NET 8 niektóre implementacje (RSAOpenSsl) nadpisują tylko setter KeySize i PowerShell
    # widzi ją jako "tylko do zapisu".
    param([Parameter(Mandatory = $true)]$Rsa)
    try {
        $property = [System.Security.Cryptography.AsymmetricAlgorithm].GetProperty('KeySize')
        $size = [int]$property.GetValue($Rsa, $null)
        if ($size -gt 0) { return $size }
    } catch { }
    try {
        return ($Rsa.ExportParameters($false).Modulus.Length * 8)
    } catch {
        return 0
    }
}

function Get-CryptoHmacSha256 {
    param(
        [Parameter(Mandatory = $true)][byte[]]$Key,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Data
    )
    $hmac = [System.Security.Cryptography.HMACSHA256]::new($Key)
    try {
        return , ($hmac.ComputeHash($Data))
    } finally {
        $hmac.Dispose()
    }
}

function Invoke-CryptoAesCbc {
    # AES-256-CBC z dopełnieniem PKCS7. -Decrypt: odszyfrowanie (zły klucz / dopełnienie -> wyjątek).
    param(
        [Parameter(Mandatory = $true)][byte[]]$Key,
        [Parameter(Mandatory = $true)][byte[]]$InitVector,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Data,
        [switch]$Decrypt
    )
    # Uwaga: parametr nie może nazywać się -IV (to alias wspólnego parametru -InformationVariable).
    $aes = [System.Security.Cryptography.Aes]::Create()
    try {
        $aes.KeySize = 256
        $aes.BlockSize = 128
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        if ($Decrypt) {
            $transform = $aes.CreateDecryptor($Key, $InitVector)
        } else {
            $transform = $aes.CreateEncryptor($Key, $InitVector)
        }
        try {
            $output = $transform.TransformFinalBlock($Data, 0, $Data.Length)
        } finally {
            $transform.Dispose()
        }
        return , $output
    } finally {
        $aes.Dispose()
    }
}

function Get-CryptoSealedParameters {
    # Sprawdza budowę magazynu (mode, kdf, check) i zwraca @{ Salt; Iterations; Check }.
    param([Parameter(Mandatory = $true)][AllowNull()]$Store)
    if ($null -eq $Store) { throw '[STD-SEC-FORMAT] Brak magazynu sekretów.' }
    $mode = Get-CryptoMember $Store 'mode'
    if ($mode -isnot [string] -or $mode -cne $script:StdCryptoSealedMode) {
        throw "[STD-SEC-FORMAT] Nieprawidłowy magazyn sekretów: pole mode musi mieć wartość '$($script:StdCryptoSealedMode)'."
    }
    $kdf = Get-CryptoMember $Store 'kdf'
    if ($null -eq $kdf) { throw '[STD-SEC-FORMAT] Nieprawidłowy magazyn sekretów: brak pola kdf.' }
    $alg = Get-CryptoMember $kdf 'alg'
    if ($alg -isnot [string] -or $alg -cne $script:StdCryptoKdfAlg) {
        throw "[STD-SEC-FORMAT] Nieprawidłowy magazyn sekretów: kdf.alg musi mieć wartość '$($script:StdCryptoKdfAlg)'."
    }
    $iterations = Get-CryptoIntegerValue (Get-CryptoMember $kdf 'iter')
    if ($null -eq $iterations -or $iterations -lt 1) {
        throw '[STD-SEC-FORMAT] Nieprawidłowy magazyn sekretów: kdf.iter musi być liczbą całkowitą większą od zera.'
    }
    $salt = ConvertFrom-CryptoBase64Field -Value (Get-CryptoMember $kdf 'salt') -Code 'STD-SEC-FORMAT' -Message 'Nieprawidłowy magazyn sekretów: kdf.salt musi być tekstem base64.'
    if ($salt.Length -lt 8) { throw '[STD-SEC-FORMAT] Nieprawidłowy magazyn sekretów: kdf.salt jest za krótki (minimum 8 bajtów).' }
    $check = ConvertFrom-CryptoBase64Field -Value (Get-CryptoMember $Store 'check') -Code 'STD-SEC-FORMAT' -Message 'Nieprawidłowy magazyn sekretów: pole check musi być 32 bajtami w base64.' -Length 32
    $items = Get-CryptoMember $Store 'items'
    if ($null -ne $items -and $items -isnot [System.Collections.IDictionary] -and $items -isnot [System.Management.Automation.PSCustomObject]) {
        throw '[STD-SEC-FORMAT] Nieprawidłowy magazyn sekretów: pole items musi być obiektem.'
    }
    return @{ Salt = $salt; Iterations = $iterations; Check = $check }
}

function New-CryptoSealedKeys {
    # PBKDF2-SHA256(UTF-8 hasła, salt, iter) -> 64 bajty: EncKey = 0-31, MacKey = 32-63.
    # Wywołujący zeruje klucze przez Clear-CryptoSealedKeys.
    param(
        [Parameter(Mandatory = $true)][System.Security.SecureString]$Password,
        [Parameter(Mandatory = $true)][byte[]]$Salt,
        [Parameter(Mandatory = $true)][int]$Iterations
    )
    $passwordBytes = ConvertTo-CryptoUtf8Bytes -SecureString $Password
    $kdf = $null
    $all = $null
    try {
        $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($passwordBytes, $Salt, $Iterations, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
        $all = $kdf.GetBytes(64)
        $encKey = New-Object byte[] 32
        $macKey = New-Object byte[] 32
        [System.Buffer]::BlockCopy($all, 0, $encKey, 0, 32)
        [System.Buffer]::BlockCopy($all, 32, $macKey, 0, 32)
        return @{ EncKey = $encKey; MacKey = $macKey }
    } finally {
        Clear-CryptoBytes $all
        Clear-CryptoBytes $passwordBytes
        if ($null -ne $kdf) { $kdf.Dispose() }
    }
}

function Clear-CryptoSealedKeys {
    param([AllowNull()]$Keys)
    if ($null -eq $Keys) { return }
    Clear-CryptoBytes $Keys.EncKey
    Clear-CryptoBytes $Keys.MacKey
}

function Test-CryptoSealedCheck {
    # Czy klucze pasują do pola check (HMAC-SHA256(macKey, 'std-check')), porównanie w stałym czasie.
    param([Parameter(Mandatory = $true)]$Keys, [Parameter(Mandatory = $true)][byte[]]$Check)
    $data = [System.Text.Encoding]::UTF8.GetBytes($script:StdCryptoCheckText)
    $actual = Get-CryptoHmacSha256 -Key $Keys.MacKey -Data $data
    return (Test-CryptoBytesEqual -Expected $actual -Actual $Check)
}

function Assert-CryptoSecretName {
    param([AllowNull()][AllowEmptyString()][string]$Name)
    if ([string]::IsNullOrEmpty($Name) -or $Name -notmatch $script:StdCryptoSecretNamePattern) {
        throw "[STD-SEC-FORMAT] Nieprawidłowa nazwa sekretu '$(Get-CryptoSafeText $Name 64)': dozwolone są litery, cyfry, '.', '_' i '-' (do 64 znaków, na początku litera lub cyfra)."
    }
}

function Get-CryptoTrustedCandidates {
    # Lista zaufanych certyfikatów: podana w parametrze (także pusta - wtedy nie sięgamy do pliku)
    # albo z Get-StdTrustedCertificates (Std.Config.ps1, plik Modules\Std.TrustedSigners.psd1).
    param([bool]$Given, [AllowNull()]$TrustedCertificates)
    $candidates = $null
    if ($Given) {
        $candidates = $TrustedCertificates
    } elseif ($null -ne (Get-Command -Name 'Get-StdTrustedCertificates' -ErrorAction SilentlyContinue)) {
        $candidates = Get-StdTrustedCertificates
    }
    $list = New-Object 'System.Collections.Generic.List[System.Security.Cryptography.X509Certificates.X509Certificate2]'
    foreach ($candidate in @($candidates)) {
        if ($null -ne $candidate -and $candidate -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) {
            $list.Add($candidate)
        }
    }
    return , $list
}

function Resolve-CryptoSecretValue {
    # Wynik resolvera dla jednej nazwy: ostatnia wartość różna od $null; SecureString zamieniany na tekst.
    param([Parameter(Mandatory = $true)][string]$Name, [Parameter(Mandatory = $true)][scriptblock]$Resolver)
    $output = @(& $Resolver $Name)
    $value = $null
    for ($i = $output.Count - 1; $i -ge 0; $i--) {
        if ($null -ne $output[$i]) { $value = $output[$i]; break }
    }
    if ($null -eq $value) {
        throw "[STD-SEC-MISSING] Brak wartości sekretu '$Name'."
    }
    if ($value -is [System.Security.SecureString]) {
        return (ConvertFrom-CryptoSecureString -SecureString $value)
    }
    return [string]$value
}

# ---------------------------------------------------------------------------------------------
# SHA-256
# ---------------------------------------------------------------------------------------------

function Get-StdSha256 {
    <#
    .SYNOPSIS
    Suma SHA-256 pliku (-Path) albo bajtów (-Bytes) jako HEX małymi literami.
    Brak pliku albo błąd odczytu -> [STD-IO].
    #>
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        [Parameter(Mandatory = $true, ParameterSetName = 'Bytes')]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )
    $sha = New-CryptoSha256
    try {
        if ($PSCmdlet.ParameterSetName -eq 'Bytes') {
            $hash = $sha.ComputeHash($Bytes)
        } else {
            $full = Resolve-CryptoFullPath -Path $Path
            if (-not [System.IO.File]::Exists($full)) {
                throw "[STD-IO] Nie znaleziono pliku $full - nie można policzyć sumy SHA-256."
            }
            try {
                $stream = New-Object System.IO.FileStream($full, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read, 1048576)
                try {
                    $hash = $sha.ComputeHash($stream)
                } finally {
                    $stream.Dispose()
                }
            } catch {
                throw "[STD-IO] Nie udało się odczytać pliku $full - $($_.Exception.Message)"
            }
        }
    } finally {
        $sha.Dispose()
    }
    return (ConvertTo-CryptoHex -Bytes $hash)
}

function Test-StdFileHash {
    <#
    .SYNOPSIS
    Czy suma SHA-256 pliku jest równa oczekiwanej (bez względu na wielkość liter). Brak pliku albo
    oczekiwana wartość, która nie jest 64 znakami HEX -> $false (z ostrzeżeniem w logu).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        [Parameter(Mandatory = $true, Position = 1)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Expected
    )
    $expectedHash = ''
    if ($null -ne $Expected) { $expectedHash = $Expected.Trim() }
    $full = Resolve-CryptoFullPath -Path $Path
    if ($expectedHash -notmatch '^[0-9A-Fa-f]{64}$') {
        Write-StdLog -Level Warn -Source 'Std.Crypto' -Message "Oczekiwana suma SHA-256 dla pliku $full jest nieprawidłowa (potrzebne 64 znaki HEX)."
        return $false
    }
    if (-not [System.IO.File]::Exists($full)) {
        Write-StdLog -Level Warn -Source 'Std.Crypto' -Message "Nie znaleziono pliku $full - nie można sprawdzić sumy SHA-256."
        return $false
    }
    $actual = Get-StdSha256 -Path $full
    if ([string]::Equals($actual, $expectedHash, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    Write-StdLog -Level Warn -Source 'Std.Crypto' -Message ('Suma SHA-256 pliku {0} się nie zgadza: oczekiwano {1}, jest {2}.' -f $full, $expectedHash.ToLowerInvariant(), $actual)
    return $false
}

# ---------------------------------------------------------------------------------------------
# Podpisana koperta (*.sdt)
# ---------------------------------------------------------------------------------------------

function New-StdSignedEnvelope {
    <#
    .SYNOPSIS
    Podpisuje bajty dokumentu (RS256: RSA PKCS#1 v1.5 + SHA-256) i zwraca JSON koperty *.sdt (tekst).
    Certyfikat musi mieć klucz prywatny RSA (co najmniej 2048 bitów). kid = odcisk certyfikatu (HEX wielkimi literami).
    Błędy: [STD-SIG-FORMAT] pusta treść, [STD-SIG-KEY] brak klucza prywatnego RSA albo nieudany podpis.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyCollection()]
        [byte[]]$Payload,
        [Parameter(Mandatory = $true, Position = 1)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    if ($Payload.Length -eq 0) {
        throw '[STD-SIG-FORMAT] Nie można podpisać pustego dokumentu.'
    }
    $kid = ([string]$Certificate.Thumbprint).ToUpperInvariant()
    if (-not $Certificate.HasPrivateKey) {
        throw "[STD-SIG-KEY] Certyfikat $kid nie ma klucza prywatnego - nie można nim podpisać konfiguracji."
    }
    $hashName = [System.Security.Cryptography.HashAlgorithmName]::SHA256
    $padding = [System.Security.Cryptography.RSASignaturePadding]::Pkcs1

    $privateKey = $null
    try {
        $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    } catch {
        throw "[STD-SIG-KEY] Nie udało się otworzyć klucza prywatnego certyfikatu $kid - $($_.Exception.Message)"
    }
    if ($null -eq $privateKey) {
        throw "[STD-SIG-KEY] Certyfikat $kid nie ma klucza RSA - STD podpisuje tylko kluczem RSA (RS256)."
    }
    try {
        $keySize = Get-CryptoRsaKeySize -Rsa $privateKey
        if ($keySize -gt 0 -and $keySize -lt $script:StdCryptoMinSigningKeySize) {
            throw "[STD-SIG-KEY] Klucz certyfikatu $kid jest za krótki ($keySize bitów, minimum $($script:StdCryptoMinSigningKeySize))."
        }
        try {
            $signature = $privateKey.SignData($Payload, $hashName, $padding)
        } catch {
            throw "[STD-SIG-KEY] Nie udało się podpisać dokumentu certyfikatem $kid - $($_.Exception.Message)"
        }
    } finally {
        $privateKey.Dispose()
    }

    # Kontrola: świeży podpis musi się zgadzać z kluczem publicznym certyfikatu, zanim trafi na serwer.
    $publicKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($Certificate)
    $selfCheck = $false
    try {
        $selfCheck = $publicKey.VerifyData($Payload, $signature, $hashName, $padding)
    } catch {
        $selfCheck = $false
    } finally {
        if ($null -ne $publicKey) { $publicKey.Dispose() }
    }
    if (-not $selfCheck) {
        throw "[STD-SIG-KEY] Podpis utworzony certyfikatem $kid nie zgadza się z jego kluczem publicznym."
    }

    # Stały układ tekstu (bez ConvertTo-Json): ten sam wynik w Windows PowerShell 5.1 i PowerShell 7.
    # Wartości to tylko HEX i base64, więc nie wymagają znaków ucieczki.
    $lines = @(
        '{'
        ('  "kind": "{0}",' -f $script:StdCryptoEnvelopeKind)
        ('  "alg": "{0}",' -f $script:StdCryptoEnvelopeAlg)
        ('  "kid": "{0}",' -f $kid)
        ('  "payload": "{0}",' -f [System.Convert]::ToBase64String($Payload))
        ('  "sig": "{0}"' -f [System.Convert]::ToBase64String($signature))
        '}'
    )
    Write-StdLog -Level Debug -Source 'Std.Crypto' -Message "Podpisano dokument ($($Payload.Length) B) certyfikatem $kid."
    return ($lines -join "`n")
}

function Read-StdSignedEnvelope {
    <#
    .SYNOPSIS
    Sprawdza podpisaną kopertę *.sdt (tekst -Json albo bajty pliku -Bytes) i zwraca
    @{ Kid; PayloadBytes; PayloadText; Signer }.
    Zaufanie: kid musi być na liście -TrustedCertificates (bez parametru: Get-StdTrustedCertificates)
    i podpis musi się zgadzać z kluczem publicznym tego certyfikatu.
    Błędy: [STD-SIG-FORMAT], [STD-SIG-ALG], [STD-SIG-NOTRUST], [STD-SIG-UNTRUSTED], [STD-SIG-INVALID].
    #>
    [CmdletBinding(DefaultParameterSetName = 'Json')]
    param(
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Json')]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Json,
        [Parameter(Mandatory = $true, ParameterSetName = 'Bytes')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [byte[]]$Bytes,
        [AllowNull()]
        [AllowEmptyCollection()]
        [System.Security.Cryptography.X509Certificates.X509Certificate2[]]$TrustedCertificates
    )
    $formatPrefix = '[STD-SIG-FORMAT] Nieprawidłowy plik podpisanej konfiguracji (*.sdt):'

    # 1. Tekst i JSON koperty.
    if ($PSCmdlet.ParameterSetName -eq 'Bytes') {
        if ($null -eq $Bytes -or $Bytes.Length -eq 0) { throw "$formatPrefix plik jest pusty." }
        try {
            $text = ConvertFrom-CryptoUtf8 -Bytes $Bytes -Strict
        } catch {
            throw "$formatPrefix to nie jest tekst UTF-8."
        }
    } else {
        $text = $Json
    }
    if ($null -eq $text) { $text = '' }
    $text = $text.TrimStart([char]0xFEFF)
    if ([string]::IsNullOrWhiteSpace($text)) { throw "$formatPrefix plik jest pusty." }
    if (-not $text.TrimStart().StartsWith('{')) { throw "$formatPrefix to nie jest obiekt JSON." }
    $envelope = $null
    try {
        $envelope = ConvertFrom-Json -InputObject $text -ErrorAction Stop
    } catch {
        $envelope = $null
    }
    if ($null -eq $envelope -or $envelope -isnot [System.Management.Automation.PSCustomObject]) {
        throw "$formatPrefix to nie jest poprawny obiekt JSON."
    }

    # 2. Pola koperty.
    $fields = @{}
    foreach ($fieldName in @('kind', 'alg', 'kid', 'payload', 'sig')) {
        $property = $envelope.PSObject.Properties[$fieldName]
        if ($null -eq $property -or $property.Value -isnot [string] -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            throw "$formatPrefix brak pola '$fieldName'."
        }
        $fields[$fieldName] = [string]$property.Value
    }
    if ($fields['kind'] -cne $script:StdCryptoEnvelopeKind) {
        throw "$formatPrefix pole kind ma wartość '$(Get-CryptoSafeText $fields['kind'])' zamiast '$($script:StdCryptoEnvelopeKind)'."
    }
    if ($fields['alg'] -cne $script:StdCryptoEnvelopeAlg) {
        throw "[STD-SIG-ALG] Nieobsługiwany algorytm podpisu '$(Get-CryptoSafeText $fields['alg'])' - STD przyjmuje tylko $($script:StdCryptoEnvelopeAlg)."
    }
    if ($fields['kid'] -notmatch '^[0-9A-Fa-f]{40}$') {
        throw "$formatPrefix pole kid musi być odciskiem SHA-1 certyfikatu (40 znaków HEX)."
    }
    $kid = $fields['kid'].ToUpperInvariant()
    $payloadBytes = ConvertFrom-CryptoBase64Field -Value $fields['payload'] -Code 'STD-SIG-FORMAT' -Message 'Nieprawidłowy plik podpisanej konfiguracji (*.sdt): pole payload nie jest poprawnym base64.'
    $signature = ConvertFrom-CryptoBase64Field -Value $fields['sig'] -Code 'STD-SIG-FORMAT' -Message 'Nieprawidłowy plik podpisanej konfiguracji (*.sdt): pole sig nie jest poprawnym base64.'

    # 3. Zaufane certyfikaty.
    $trusted = Get-CryptoTrustedCandidates -Given $PSBoundParameters.ContainsKey('TrustedCertificates') -TrustedCertificates $TrustedCertificates
    if ($trusted.Count -eq 0) {
        throw '[STD-SIG-NOTRUST] Brak zaufanych certyfikatów do sprawdzenia podpisu konfiguracji. Utwórz certyfikat skryptem Tools\New-StdSigningCertificate.ps1 i dopisz wypisany wpis do Modules\Std.TrustedSigners.psd1.'
    }
    $signer = $null
    foreach ($candidate in $trusted) {
        $thumbprint = [string]$candidate.Thumbprint
        if (-not [string]::IsNullOrEmpty($thumbprint) -and $thumbprint.ToUpperInvariant() -ceq $kid) {
            $signer = $candidate
            break
        }
    }
    if ($null -eq $signer) {
        throw "[STD-SIG-UNTRUSTED] Konfiguracja jest podpisana certyfikatem $kid, którego nie ma na liście zaufanych (Modules\Std.TrustedSigners.psd1)."
    }

    # 4. Podpis kluczem publicznym zaufanego certyfikatu.
    $valid = $false
    $publicKey = $null
    try {
        $publicKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($signer)
        if ($null -ne $publicKey) {
            $valid = $publicKey.VerifyData($payloadBytes, $signature, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        }
    } catch {
        $valid = $false
    } finally {
        if ($null -ne $publicKey) { $publicKey.Dispose() }
    }
    if (-not $valid) {
        throw "[STD-SIG-INVALID] Podpis konfiguracji jest nieprawidłowy (certyfikat $kid) - plik mógł zostać zmieniony albo uszkodzony."
    }

    Write-StdLog -Level Debug -Source 'Std.Crypto' -Message "Podpis poprawny: certyfikat $kid ($($signer.Subject)), $($payloadBytes.Length) B."
    return @{
        Kid          = $kid
        PayloadBytes = $payloadBytes
        PayloadText  = (ConvertFrom-CryptoUtf8 -Bytes $payloadBytes)
        Signer       = $signer
    }
}

# ---------------------------------------------------------------------------------------------
# Sekrety (tryb sealed)
# ---------------------------------------------------------------------------------------------

function New-StdSealedStore {
    <#
    .SYNOPSIS
    Tworzy pusty magazyn sekretów (SPEC 2.7) dla hasła wdrożeniowego klienta:
    [ordered]@{ mode; kdf = @{ alg; iter; salt }; check; items = @{} }. Hasło nie trafia do magazynu.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Security.SecureString]$Password,
        [ValidateRange(1000, 100000000)]
        [int]$Iterations = 600000
    )
    if ($Password.Length -eq 0) {
        throw '[STD-SEC-PASSWORD] Hasło wdrożeniowe klienta nie może być puste.'
    }
    $salt = New-CryptoRandomBytes -Count 16
    $keys = New-CryptoSealedKeys -Password $Password -Salt $salt -Iterations $Iterations
    try {
        $check = Get-CryptoHmacSha256 -Key $keys.MacKey -Data ([System.Text.Encoding]::UTF8.GetBytes($script:StdCryptoCheckText))
    } finally {
        Clear-CryptoSealedKeys $keys
    }
    return [ordered]@{
        mode  = $script:StdCryptoSealedMode
        kdf   = [ordered]@{
            alg  = $script:StdCryptoKdfAlg
            iter = $Iterations
            salt = [System.Convert]::ToBase64String($salt)
        }
        check = [System.Convert]::ToBase64String($check)
        items = [ordered]@{}
    }
}

function Test-StdSealedPassword {
    <#
    .SYNOPSIS
    Czy hasło pasuje do magazynu sekretów (pole check). Zły układ magazynu -> [STD-SEC-FORMAT].
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Collections.IDictionary]$Store,
        [Parameter(Mandatory = $true, Position = 1)]
        [System.Security.SecureString]$Password
    )
    $parameters = Get-CryptoSealedParameters -Store $Store
    if ($Password.Length -eq 0) { return $false }
    $keys = New-CryptoSealedKeys -Password $Password -Salt $parameters.Salt -Iterations $parameters.Iterations
    try {
        return [bool](Test-CryptoSealedCheck -Keys $keys -Check $parameters.Check)
    } finally {
        Clear-CryptoSealedKeys $keys
    }
}

function Set-StdSealedItem {
    <#
    .SYNOPSIS
    Szyfruje wartość i dodaje albo podmienia pozycję -Name w magazynie (zmienia przekazany obiekt).
    Najpierw sprawdza hasło (jak Test-StdSealedPassword): złe hasło -> [STD-SEC-PASSWORD], magazyn bez zmian.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Collections.IDictionary]$Store,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$Name,
        [Parameter(Mandatory = $true, Position = 2)]
        [System.Security.SecureString]$Value,
        [Parameter(Mandatory = $true, Position = 3)]
        [System.Security.SecureString]$Password
    )
    Assert-CryptoSecretName -Name $Name
    $parameters = Get-CryptoSealedParameters -Store $Store
    $items = Get-CryptoMember $Store 'items'
    if ($null -ne $items -and $items -isnot [System.Collections.IDictionary]) {
        throw '[STD-SEC-FORMAT] Nieprawidłowy magazyn sekretów: pole items musi być słownikiem (hashtable), żeby dało się je zmienić.'
    }
    if ($Password.Length -eq 0) {
        throw '[STD-SEC-PASSWORD] Nieprawidłowe hasło wdrożeniowe klienta.'
    }
    # Jedno wyprowadzenie klucza: sprawdzenie hasła (pole check) i szyfrowanie tym samym kluczem.
    $keys = New-CryptoSealedKeys -Password $Password -Salt $parameters.Salt -Iterations $parameters.Iterations
    try {
        if (-not (Test-CryptoSealedCheck -Keys $keys -Check $parameters.Check)) {
            throw '[STD-SEC-PASSWORD] Nieprawidłowe hasło wdrożeniowe klienta - sekret nie został zapisany.'
        }
        $plain = ConvertTo-CryptoUtf8Bytes -SecureString $Value
        try {
            $iv = New-CryptoRandomBytes -Count 16
            $cipherText = Invoke-CryptoAesCbc -Key $keys.EncKey -InitVector $iv -Data $plain
        } finally {
            Clear-CryptoBytes $plain
        }
        $mac = Get-CryptoHmacSha256 -Key $keys.MacKey -Data (Join-CryptoBytes -First $iv -Second $cipherText)
    } finally {
        Clear-CryptoSealedKeys $keys
    }
    if ($null -eq $items) {
        $items = [ordered]@{}
        $itemsKey = Find-CryptoDictionaryKey -Dictionary $Store -Name 'items'
        if ($null -eq $itemsKey) { $itemsKey = 'items' }
        $Store[$itemsKey] = $items
    }
    $key = Find-CryptoDictionaryKey -Dictionary $items -Name $Name
    if ($null -eq $key) { $key = $Name }
    $items[$key] = [ordered]@{
        iv  = [System.Convert]::ToBase64String($iv)
        ct  = [System.Convert]::ToBase64String($cipherText)
        mac = [System.Convert]::ToBase64String($mac)
    }
    Write-StdLog -Level Info -Source 'Std.Crypto' -Message "Zapisano zaszyfrowany sekret '$Name'."
}

function Unprotect-StdSealedItem {
    <#
    .SYNOPSIS
    Odszyfrowuje pozycję -Name i zwraca ją jako [string]; wartość od razu trafia do maski logu.
    Kolejność: brak pozycji -> [STD-SEC-MISSING], złe hasło -> [STD-SEC-PASSWORD],
    zmieniona / uszkodzona pozycja (mac, iv, ct) -> [STD-SEC-TAMPER]. Zły układ magazynu -> [STD-SEC-FORMAT].
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Collections.IDictionary]$Store,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$Name,
        [Parameter(Mandatory = $true, Position = 2)]
        [System.Security.SecureString]$Password
    )
    $parameters = Get-CryptoSealedParameters -Store $Store
    $item = Get-CryptoMember (Get-CryptoMember $Store 'items') $Name
    if ($null -eq $item) {
        throw "[STD-SEC-MISSING] Brak sekretu '$(Get-CryptoSafeText $Name 64)' w konfiguracji."
    }
    if ($Password.Length -eq 0) {
        throw '[STD-SEC-PASSWORD] Nieprawidłowe hasło wdrożeniowe klienta.'
    }
    $tamperMessage = "Sekret '$(Get-CryptoSafeText $Name 64)' jest uszkodzony albo został zmieniony (suma kontrolna się nie zgadza)."
    $tamper = "[STD-SEC-TAMPER] $tamperMessage"
    $value = $null
    $keys = New-CryptoSealedKeys -Password $Password -Salt $parameters.Salt -Iterations $parameters.Iterations
    try {
        if (-not (Test-CryptoSealedCheck -Keys $keys -Check $parameters.Check)) {
            throw '[STD-SEC-PASSWORD] Nieprawidłowe hasło wdrożeniowe klienta - nie można odczytać sekretów.'
        }
        $iv = ConvertFrom-CryptoBase64Field -Value (Get-CryptoMember $item 'iv') -Code 'STD-SEC-TAMPER' -Message $tamperMessage -Length 16
        $cipherText = ConvertFrom-CryptoBase64Field -Value (Get-CryptoMember $item 'ct') -Code 'STD-SEC-TAMPER' -Message $tamperMessage
        $mac = ConvertFrom-CryptoBase64Field -Value (Get-CryptoMember $item 'mac') -Code 'STD-SEC-TAMPER' -Message $tamperMessage -Length 32
        if (($cipherText.Length % 16) -ne 0) { throw $tamper }
        $expectedMac = Get-CryptoHmacSha256 -Key $keys.MacKey -Data (Join-CryptoBytes -First $iv -Second $cipherText)
        if (-not (Test-CryptoBytesEqual -Expected $expectedMac -Actual $mac)) { throw $tamper }
        $plain = $null
        try {
            $plain = Invoke-CryptoAesCbc -Key $keys.EncKey -InitVector $iv -Data $cipherText -Decrypt
        } catch {
            throw $tamper
        }
        try {
            $value = (New-Object System.Text.UTF8Encoding($false)).GetString($plain)
        } finally {
            Clear-CryptoBytes $plain
        }
    } finally {
        Clear-CryptoSealedKeys $keys
    }
    Register-StdLogMask -Value $value
    Write-StdLog -Level Debug -Source 'Std.Crypto' -Message "Odczytano sekret '$Name'."
    return $value
}

# ---------------------------------------------------------------------------------------------
# Wstawki {{secret:nazwa}}
# ---------------------------------------------------------------------------------------------

function Get-StdSecretReference {
    <#
    .SYNOPSIS
    Nazwy sekretów z wstawek {{secret:nazwa}} w tekście: unikalne (bez względu na wielkość liter),
    w kolejności pierwszego wystąpienia. Brak wstawek -> brak wyniku (pusta lista przy @(...)).
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Text
    )
    if ([string]::IsNullOrEmpty($Text)) { return }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $names = New-Object 'System.Collections.Generic.List[string]'
    foreach ($match in $script:StdCryptoSecretRefRegex.Matches($Text)) {
        $name = $match.Groups['name'].Value
        if ($seen.Add($name)) { $names.Add($name) }
    }
    foreach ($name in $names) { $name }
}

function Expand-StdSecretTemplate {
    <#
    .SYNOPSIS
    Podmienia każde {{secret:nazwa}} wynikiem & $Resolver nazwa (każda nazwa jest rozwiązywana raz)
    i każdą wartość dodaje do maski logu. Resolver może zwrócić [string] albo SecureString;
    brak wartości ($null) -> [STD-SEC-MISSING]. Wynik jest jednorazowy - nie zapisuj go nigdzie.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Text,
        [Parameter(Mandatory = $true, Position = 1)]
        [scriptblock]$Resolver
    )
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $found = $script:StdCryptoSecretRefRegex.Matches($Text)
    if ($found.Count -eq 0) { return $Text }

    $values = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @(Get-StdSecretReference -Text $Text)) {
        $value = Resolve-CryptoSecretValue -Name $name -Resolver $Resolver
        Register-StdLogMask -Value $value
        $values[$name] = $value
    }

    $builder = New-Object System.Text.StringBuilder
    $position = 0
    foreach ($match in $found) {
        [void]$builder.Append($Text, $position, $match.Index - $position)
        [void]$builder.Append($values[$match.Groups['name'].Value])
        $position = $match.Index + $match.Length
    }
    [void]$builder.Append($Text, $position, $Text.Length - $position)
    Write-StdLog -Level Debug -Source 'Std.Crypto' -Message ('Podstawiono sekrety: {0}.' -f (@($values.Keys) -join ', '))
    return $builder.ToString()
}
