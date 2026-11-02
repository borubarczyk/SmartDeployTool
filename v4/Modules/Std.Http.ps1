# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Std.Http.ps1 - sieć (SPEC 3.2): TLS, sesja technika (Basic), żądania HTTPS, odczyt plików z serwera,
# udziału sieciowego i dysku, łączenie adresów i ścieżek, test portu TCP.
#
# Wymaga Std.Core.ps1 (Write-StdLog, Protect-StdLogText, Get-StdVersion) - STD.psm1 wczytuje go wcześniej.
#
# Zasady bezpieczeństwa:
# - tylko https:// (inaczej [STD-HTTP-INSECURE]), bez wyłączania sprawdzania certyfikatu serwera;
# - przekierowania nie są wykonywane (AllowAutoRedirect = $false), więc nagłówek Authorization
#   nie trafi na inny adres;
# - hasło jest w sesji tylko jako SecureString; jawny tekst istnieje tylko na czas składania nagłówka;
# - nagłówka Authorization i hasła nie ma w żadnym komunikacie ani logu.

# ---------------------------------------------------------------------------------------------
# Pomocnicze (prywatne)
# ---------------------------------------------------------------------------------------------

function Get-HttpWindowsBuild {
    if (-not (Test-CoreIsWindows)) { return 0 }
    try {
        $value = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'CurrentBuildNumber' -ErrorAction Stop).CurrentBuildNumber
        return [int]$value
    } catch {
        return [System.Environment]::OSVersion.Version.Build
    }
}

function Get-HttpSafeLocation {
    # Adres / ścieżka do komunikatu: bez loginu i hasła w adresie, z zamaskowanymi sekretami.
    param([AllowNull()][AllowEmptyString()][string]$Location)
    if ([string]::IsNullOrEmpty($Location)) { return '' }
    $safe = $Location -replace '(?<=://)[^/@\s]*@', '***@'
    return (Protect-StdLogText -Text $safe)
}

function Assert-HttpSession {
    param($Session)
    $ok = $false
    if ($null -ne $Session) {
        try {
            $ok = ([string]$Session.Mode -eq 'basic') -and
                  (-not [string]::IsNullOrEmpty([string]$Session.UserName)) -and
                  ($Session.Password -is [System.Security.SecureString])
        } catch {
            $ok = $false
        }
    }
    if (-not $ok) {
        throw '[STD-SESSION] Nieprawidłowa sesja STD - utwórz ją przez New-StdSession.'
    }
}

function Get-HttpWebException {
    # PowerShell opakowuje wyjątki metod .NET (MethodInvocationException) - szukamy WebException w środku.
    param($Exception)
    $ex = $Exception
    while ($null -ne $ex) {
        if ($ex -is [System.Net.WebException]) { return $ex }
        $ex = $ex.InnerException
    }
    return $null
}

function Get-HttpStatusHint {
    param([int]$StatusCode, [string]$RedirectLocation)
    switch ($StatusCode) {
        400 { return 'nieprawidłowe żądanie' }
        401 { return 'brak lub błędne dane logowania' }
        403 { return 'brak uprawnień' }
        404 { return 'nie znaleziono' }
        405 { return 'metoda niedozwolona' }
        409 { return 'konflikt (np. brak katalogu nadrzędnego)' }
        413 { return 'plik jest za duży dla serwera' }
        429 { return 'za dużo żądań - spróbuj później' }
    }
    if ($StatusCode -ge 300 -and $StatusCode -le 399) {
        if (-not [string]::IsNullOrEmpty($RedirectLocation)) {
            return "przekierowanie na $(Get-HttpSafeLocation $RedirectLocation) (STD nie wykonuje przekierowań)"
        }
        return 'przekierowanie (STD nie wykonuje przekierowań)'
    }
    if ($StatusCode -ge 500) { return 'błąd serwera' }
    return ''
}

function Get-HttpNetHint {
    param($WebException, [int]$TimeoutSec)
    if ($null -eq $WebException) { return '' }
    switch ([string]$WebException.Status) {
        'Timeout' { return "przekroczono limit czasu ($TimeoutSec s)" }
        'NameResolutionFailure' { return 'nie można odnaleźć serwera (DNS)' }
        'ConnectFailure' { return 'brak połączenia z serwerem' }
        'TrustFailure' { return 'certyfikat serwera nie jest zaufany' }
        'SecureChannelFailure' { return 'nie udało się zestawić połączenia TLS (certyfikat lub wersja TLS)' }
        'ProxyNameResolutionFailure' { return 'nie można odnaleźć serwera proxy' }
        'ConnectionClosed' { return 'serwer zamknął połączenie' }
        'ReceiveFailure' { return 'błąd odbioru danych' }
        'SendFailure' { return 'błąd wysyłania danych' }
    }
    # PowerShell 7 (.NET Core) zgłasza zwykle UnknownError - rozpoznajemy przyczynę po wyjątku w środku.
    $inner = $WebException.InnerException
    while ($null -ne $inner) {
        if ($inner -is [System.Security.Authentication.AuthenticationException]) {
            return 'nie udało się zestawić połączenia TLS (certyfikat serwera nie jest zaufany albo wersja TLS)'
        }
        if ($inner -is [System.TimeoutException]) { return "przekroczono limit czasu ($TimeoutSec s)" }
        if ($inner -is [System.Net.Sockets.SocketException]) {
            switch ([string]$inner.SocketErrorCode) {
                'HostNotFound' { return 'nie można odnaleźć serwera (DNS)' }
                'NoData' { return 'nie można odnaleźć serwera (DNS)' }
                'TryAgain' { return 'nie można odnaleźć serwera (DNS)' }
                'TimedOut' { return "przekroczono limit czasu ($TimeoutSec s)" }
            }
            return 'brak połączenia z serwerem'
        }
        $inner = $inner.InnerException
    }
    return ''
}

function Get-HttpUncHost {
    # \\host\udział\... albo //host/udział/... -> host; ścieżki \\?\ i \\.\ oraz WebDAV (host@SSL) -> $null.
    param([string]$Path)
    if ($Path -match '^[\\/]{2}([^\\/]+)[\\/]') {
        $hostPart = $Matches[1]
        if ($hostPart -eq '?' -or $hostPart -eq '.' -or $hostPart.Contains('@')) { return $null }
        return $hostPart
    }
    return $null
}

function Save-HttpResponseStream {
    # Strumień odpowiedzi -> plik <OutFile>.part, potem podmiana na <OutFile>.
    param([System.IO.Stream]$Source, [string]$Destination)
    $part = $Destination + '.part'
    $dir = [System.IO.Path]::GetDirectoryName($Destination)
    if (-not [string]::IsNullOrEmpty($dir) -and -not [System.IO.Directory]::Exists($dir)) {
        [void][System.IO.Directory]::CreateDirectory($dir)
    }
    $target = $null
    try {
        $target = New-Object System.IO.FileStream($part, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $Source.CopyTo($target, 81920)
        $target.Dispose()
        $target = $null
        if ([System.IO.File]::Exists($Destination)) {
            try {
                [System.IO.File]::Replace($part, $Destination, [NullString]::Value)
            } catch {
                [System.IO.File]::Copy($part, $Destination, $true)
                [System.IO.File]::Delete($part)
            }
        } else {
            [System.IO.File]::Move($part, $Destination)
        }
    } catch {
        if ($null -ne $target) { try { $target.Dispose() } catch { } }
        try { if ([System.IO.File]::Exists($part)) { [System.IO.File]::Delete($part) } } catch { }
        throw
    }
}

# ---------------------------------------------------------------------------------------------
# Funkcje publiczne
# ---------------------------------------------------------------------------------------------

function Initialize-StdNetwork {
    <#
    .SYNOPSIS
    Ustawia protokoły TLS: TLS 1.2, a TLS 1.3 tylko gdy .NET go zna i system to Windows build >= 20348
    (Windows 11 / Server 2022). Na Linuksie TLS 1.3 jest dozwolony, gdy .NET go zna. Można wołać wiele razy.
    #>
    [CmdletBinding()]
    param()
    $protocol = [System.Net.SecurityProtocolType]::Tls12
    $names = [System.Enum]::GetNames([System.Net.SecurityProtocolType])
    if ($names -contains 'Tls13') {
        $allowTls13 = $true
        if (Test-CoreIsWindows) { $allowTls13 = ((Get-HttpWindowsBuild) -ge 20348) }
        if ($allowTls13) {
            # Tekst zamiast ::Tls13 - starszy .NET nie ma tej wartości i nie może się na niej wywrócić.
            $protocol = $protocol -bor ([System.Net.SecurityProtocolType]'Tls13')
        }
    }
    if ([System.Net.ServicePointManager]::SecurityProtocol -ne $protocol) {
        [System.Net.ServicePointManager]::SecurityProtocol = $protocol
        Write-StdLog -Level Debug -Source 'Initialize-StdNetwork' -Message "Protokoły TLS: $protocol"
    }
}

function New-StdSession {
    <#
    .SYNOPSIS
    Sesja technika do serwera STD: [pscustomobject]@{ Mode = 'basic'; UserName; Password (SecureString) }.
    Hasło nigdy nie jest trzymane jawnym tekstem.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Credential')]
    param(
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Credential')]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true, ParameterSetName = 'UserName')]
        [string]$UserName,
        [Parameter(Mandatory = $true, ParameterSetName = 'UserName')]
        [System.Security.SecureString]$Password
    )
    if ($PSCmdlet.ParameterSetName -eq 'Credential') {
        if ($null -eq $Credential -or $null -eq $Credential.Password) {
            throw '[STD-SESSION] Brak danych logowania do serwera STD.'
        }
        $UserName = $Credential.UserName
        $Password = $Credential.Password
    }
    if ([string]::IsNullOrWhiteSpace($UserName)) {
        throw '[STD-SESSION] Podaj login konta technika.'
    }
    if ($UserName.Contains(':')) {
        throw '[STD-SESSION] Login nie może zawierać dwukropka (:).'
    }
    $passwordCopy = $Password.Copy()
    $passwordCopy.MakeReadOnly()
    $session = [pscustomobject]@{
        Mode     = 'basic'
        UserName = $UserName.Trim()
        Password = $passwordCopy
    }
    $session.PSObject.TypeNames.Insert(0, 'STD.Session')
    return $session
}

function Get-StdBasicAuthHeader {
    <#
    .SYNOPSIS
    Wartość nagłówka Authorization: "Basic <base64(UTF-8 login:hasło)>". Jawne hasło istnieje tylko
    na czas składania i jest zaraz zerowane.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [object]$Session
    )
    Assert-HttpSession -Session $Session
    $secure = [System.Security.SecureString]$Session.Password
    $ptr = [System.IntPtr]::Zero
    $chars = $null
    $bytes = $null
    try {
        $userChars = ([string]$Session.UserName + ':').ToCharArray()
        $length = $secure.Length
        $chars = New-Object char[] ($userChars.Length + $length)
        [System.Array]::Copy($userChars, 0, $chars, 0, $userChars.Length)
        if ($length -gt 0) {
            $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($secure)
            [System.Runtime.InteropServices.Marshal]::Copy($ptr, $chars, $userChars.Length, $length)
        }
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($chars)
        return ('Basic ' + [System.Convert]::ToBase64String($bytes))
    } finally {
        if ($ptr -ne [System.IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr) }
        if ($null -ne $chars) { [System.Array]::Clear($chars, 0, $chars.Length) }
        if ($null -ne $bytes) { [System.Array]::Clear($bytes, 0, $bytes.Length) }
    }
}

function Invoke-StdHttpRequest {
    <#
    .SYNOPSIS
    Żądanie HTTPS (HttpWebRequest). Wynik @{ StatusCode; Bytes; Headers } (Bytes puste przy -OutFile i HEAD).
    Status spoza 2xx i -AllowStatus -> [STD-HTTP-<kod>] <metoda> <adres>; brak połączenia / limit czasu
    -> [STD-HTTP-NET]; adres inny niż https:// -> [STD-HTTP-INSECURE].
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Uri,
        [ValidateSet('GET', 'HEAD', 'PUT', 'MKCOL', 'DELETE')]
        [string]$Method = 'GET',
        [object]$Session,
        [byte[]]$Body,
        [string]$ContentType,
        [string]$OutFile,
        [ValidateRange(1, 86400)]
        [int]$TimeoutSec = 30,
        [int[]]$AllowStatus = @()
    )
    $Method = $Method.ToUpperInvariant()
    $safeUri = Get-HttpSafeLocation -Location $Uri

    $parsed = $null
    if (-not [System.Uri]::TryCreate($Uri, [System.UriKind]::Absolute, [ref]$parsed) -or $parsed.Scheme -ne 'https') {
        throw "[STD-HTTP-INSECURE] Dozwolone są tylko adresy https:// - odrzucono: $safeUri"
    }
    if (-not [string]::IsNullOrEmpty($parsed.UserInfo)) {
        throw "[STD-HTTP-INSECURE] Adres nie może zawierać loginu ani hasła - odrzucono: $safeUri"
    }
    if ($null -ne $Body -and ($Method -eq 'GET' -or $Method -eq 'HEAD')) {
        throw "[STD-HTTP-ARGS] Żądanie $Method nie może mieć treści (-Body)."
    }
    if (-not [string]::IsNullOrEmpty($OutFile)) { $OutFile = Resolve-CoreFullPath -Path $OutFile }

    Initialize-StdNetwork

    $request = [System.Net.HttpWebRequest]([System.Net.WebRequest]::Create($parsed))
    $request.Method = $Method
    $request.AllowAutoRedirect = $false
    $request.UserAgent = "STD/$(Get-StdVersion)"
    $request.Timeout = $TimeoutSec * 1000
    $request.ReadWriteTimeout = $TimeoutSec * 1000
    try { $request.ServicePoint.Expect100Continue = $false } catch { }
    if ($null -ne $Session) {
        Assert-HttpSession -Session $Session
        # Nagłówek od razu, bez czekania na 401.
        $request.Headers.Add('Authorization', (Get-StdBasicAuthHeader -Session $Session))
    }

    $response = $null
    try {
        try {
            if ($Method -ne 'GET' -and $Method -ne 'HEAD') {
                if ($null -ne $Body -and $Body.Length -gt 0) {
                    if ([string]::IsNullOrEmpty($ContentType)) { $ContentType = 'application/octet-stream' }
                    $request.ContentType = $ContentType
                    $request.ContentLength = $Body.Length
                    $request.AllowWriteStreamBuffering = $false
                    $requestStream = $request.GetRequestStream()
                    try { $requestStream.Write($Body, 0, $Body.Length) } finally { $requestStream.Dispose() }
                } else {
                    if (-not [string]::IsNullOrEmpty($ContentType)) { $request.ContentType = $ContentType }
                    $request.ContentLength = 0
                }
            }
            $response = [System.Net.HttpWebResponse]$request.GetResponse()
        } catch {
            $webEx = Get-HttpWebException -Exception $_.Exception
            if ($null -ne $webEx -and $null -ne $webEx.Response) {
                $response = [System.Net.HttpWebResponse]$webEx.Response
            } else {
                $hint = Get-HttpNetHint -WebException $webEx -TimeoutSec $TimeoutSec
                $detail = Format-CoreErrorText $_ 200
                if ([string]::IsNullOrEmpty($hint)) { $hint = $detail } else { $hint = "$hint ($detail)" }
                throw "[STD-HTTP-NET] $Method $safeUri - $hint"
            }
        }

        $statusCode = [int]$response.StatusCode
        $headers = @{}
        foreach ($key in $response.Headers.AllKeys) { $headers[$key] = $response.Headers[$key] }

        $isSuccess = ($statusCode -ge 200 -and $statusCode -le 299)
        if (-not $isSuccess -and -not ($AllowStatus -contains $statusCode)) {
            $hint = Get-HttpStatusHint -StatusCode $statusCode -RedirectLocation ([string]$headers['Location'])
            $message = "[STD-HTTP-$statusCode] $Method $safeUri"
            if (-not [string]::IsNullOrEmpty($hint)) { $message = "$message - $hint" }
            throw $message
        }

        $bytes = New-Object byte[] 0
        if ($Method -ne 'HEAD') {
            $responseStream = $response.GetResponseStream()
            try {
                if (-not [string]::IsNullOrEmpty($OutFile)) {
                    if ($isSuccess) { Save-HttpResponseStream -Source $responseStream -Destination $OutFile }
                } else {
                    $memory = New-Object System.IO.MemoryStream
                    try {
                        $responseStream.CopyTo($memory, 81920)
                        $bytes = $memory.ToArray()
                    } finally {
                        $memory.Dispose()
                    }
                }
            } catch {
                $webEx = Get-HttpWebException -Exception $_.Exception
                $hint = Get-HttpNetHint -WebException $webEx -TimeoutSec $TimeoutSec
                $detail = Format-CoreErrorText $_ 200
                if ([string]::IsNullOrEmpty($hint)) { $hint = "przerwany odbiór danych ($detail)" } else { $hint = "$hint ($detail)" }
                throw "[STD-HTTP-NET] $Method $safeUri - $hint"
            } finally {
                if ($null -ne $responseStream) { $responseStream.Dispose() }
            }
        }

        Write-StdLog -Level Debug -Source 'Invoke-StdHttpRequest' -Message "$Method $safeUri -> $statusCode"
        return @{
            StatusCode = $statusCode
            Bytes      = $bytes
            Headers    = $headers
        }
    } finally {
        if ($null -ne $response) { try { $response.Close() } catch { } }
    }
}

function Get-StdBytes {
    <#
    .SYNOPSIS
    Bajty pliku z adresu https://, ścieżki UNC (najpierw test portu 445, 3 s) albo ścieżki lokalnej.
    Brak pliku -> [STD-SRC-NOTFOUND] <lokalizacja>.
    #>
    [CmdletBinding()]
    [OutputType([byte[]])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Location,
        [object]$Session,
        [ValidateRange(1, 86400)]
        [int]$TimeoutSec = 30
    )
    $safeLocation = Get-HttpSafeLocation -Location $Location

    if ($Location -match '^[A-Za-z][A-Za-z0-9+.-]*://') {
        $result = Invoke-StdHttpRequest -Uri $Location -Method GET -Session $Session -TimeoutSec $TimeoutSec -AllowStatus @(404, 410)
        if ($result.StatusCode -eq 404 -or $result.StatusCode -eq 410) {
            throw "[STD-SRC-NOTFOUND] $safeLocation"
        }
        $bytes = [byte[]]$result.Bytes
        return , $bytes
    }

    $uncHost = Get-HttpUncHost -Path $Location
    if ($null -ne $uncHost) {
        if (-not (Test-StdTcpPort -HostName $uncHost -Port 445 -TimeoutMs 3000)) {
            throw "[STD-SRC-NET] $safeLocation - serwer $uncHost nie odpowiada na porcie 445 (udział sieciowy)."
        }
        $fullPath = $Location
    } else {
        $fullPath = Resolve-CoreFullPath -Path $Location
    }

    # Bez wstępnego File.Exists: przy braku uprawnień zwraca $false i "brak dostępu" wyglądałby jak "brak pliku".
    if ([System.IO.Directory]::Exists($fullPath)) {
        throw "[STD-SRC-NOTFOUND] $safeLocation"
    }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($fullPath)
    } catch {
        $inner = $_.Exception
        while ($null -ne $inner.InnerException -and -not ($inner -is [System.IO.IOException]) -and -not ($inner -is [System.UnauthorizedAccessException])) {
            $inner = $inner.InnerException
        }
        if ($inner -is [System.IO.FileNotFoundException] -or $inner -is [System.IO.DirectoryNotFoundException]) {
            throw "[STD-SRC-NOTFOUND] $safeLocation"
        }
        if ($inner -is [System.UnauthorizedAccessException]) {
            throw "[STD-SRC-READ] $safeLocation - brak dostępu do pliku."
        }
        throw "[STD-SRC-READ] $safeLocation - $(Format-CoreErrorText $_ 200)"
    }
    return , $bytes
}

function Join-StdLocation {
    <#
    .SYNOPSIS
    Łączy adres lub ścieżkę z dalszą częścią. Adres (https://...): przez "/", bez podwójnych ukośników,
    każdy segment kodowany EscapeDataString. Ścieżka: separator systemu (Windows "\", Linux "/").
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Base,
        [Parameter(Mandatory = $true, Position = 1)]
        [AllowEmptyString()]
        [string]$Child
    )
    $trailing = ($Child.Length -gt 0 -and ($Child.EndsWith('/') -or $Child.EndsWith('\')))
    $segments = @($Child -split '[\\/]' | Where-Object { $_ -ne '' })

    if ($Base -match '^[A-Za-z][A-Za-z0-9+.-]*://') {
        if ($segments.Count -eq 0) { return $Base }
        $encoded = @($segments | ForEach-Object { [System.Uri]::EscapeDataString($_) })
        $result = $Base.TrimEnd('/') + '/' + ($encoded -join '/')
        if ($trailing) { $result += '/' }
        return $result
    }

    $sep = [string][System.IO.Path]::DirectorySeparatorChar
    if ($sep -eq '\') { $other = '/' } else { $other = '\' }
    $basePath = $Base.Replace($other, $sep)
    if ($segments.Count -eq 0) { return $basePath }
    $result = $basePath.TrimEnd([char]$sep) + $sep + ($segments -join $sep)
    if ($trailing) { $result += $sep }
    return $result
}

function Test-StdTcpPort {
    <#
    .SYNOPSIS
    Czy port TCP odpowiada (TcpClient.BeginConnect + WaitOne). Zwraca [bool], nie rzuca wyjątków.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$HostName,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateRange(1, 65535)]
        [int]$Port,
        [ValidateRange(1, 600000)]
        [int]$TimeoutMs = 3000
    )
    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs)) { return $false }
        try {
            $client.EndConnect($async)
        } catch {
            return $false
        }
        return [bool]$client.Connected
    } catch {
        return $false
    } finally {
        if ($null -ne $client) { try { $client.Close() } catch { } }
    }
}
