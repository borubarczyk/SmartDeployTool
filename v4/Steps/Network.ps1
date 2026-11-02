# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Steps\Network.ps1 - krok „Czekanie na sieć” (SPEC 9).
#
# Plik jest dot-source'owany w wątku silnika (Start-StdEngine). Bloki Detect / Run silnik wywołuje później,
# dlatego funkcje pomocnicze są globalne (widoczne przez cały czas działania silnika).
#
# Sieć jest gotowa, gdy komputer ma adres IPv4 spoza 169.254.x.x i 127.x.x.x, a serwer STD (albo pierwsze
# źródło https z konfiguracji) odpowiada na porcie 443 (albo porcie z adresu). Gdy konfiguracja nie podaje
# żadnego serwera, wystarcza adres IPv4. Adres serwera: Config.Server, Config.Bootstrap.server, a potem
# pierwsze źródło typu https w kolejności Site.sourceOrder.

function global:Get-NetworkStepTarget {
    # Serwer do sprawdzenia: @{ Host; Port; Label } albo $null.
    param([AllowNull()]$Context)
    $config = $null
    if ($Context -is [System.Collections.IDictionary]) { $config = $Context['Config'] }
    if (-not ($config -is [System.Collections.IDictionary])) { return $null }

    $urls = New-Object 'System.Collections.Generic.List[string]'
    if ($config['Server'] -is [string]) { $urls.Add($config['Server']) }
    $bootstrap = $config['Bootstrap']
    if ($bootstrap -is [System.Collections.IDictionary] -and $bootstrap['server'] -is [string]) { $urls.Add($bootstrap['server']) }

    $sources = $config['Sources']
    if ($sources -is [System.Collections.IDictionary]) {
        $names = New-Object 'System.Collections.Generic.List[string]'
        $site = $config['Site']
        if ($site -is [System.Collections.IDictionary] -and $null -ne $site['sourceOrder']) {
            foreach ($name in @($site['sourceOrder'])) { if ($name -is [string] -and -not $names.Contains($name)) { $names.Add($name) } }
        }
        foreach ($name in @($sources.Keys)) { if (-not $names.Contains([string]$name)) { $names.Add([string]$name) } }
        foreach ($name in $names) {
            $source = $sources[$name]
            if ($source -is [System.Collections.IDictionary] -and [string]$source['type'] -eq 'https' -and $source['url'] -is [string]) {
                $urls.Add($source['url'])
            }
        }
    }

    foreach ($url in $urls) {
        $uri = $null
        if (-not [System.Uri]::TryCreate($url, [System.UriKind]::Absolute, [ref]$uri)) { continue }
        if ($uri.Scheme -ne 'https' -or [string]::IsNullOrEmpty($uri.Host)) { continue }
        return @{ Host = $uri.Host; Port = $uri.Port; Label = ('{0}:{1}' -f $uri.Host, $uri.Port) }
    }
    return $null
}

function global:Get-NetworkStepAddresses {
    # Adresy IPv4 działających kart sieciowych, bez 169.254.x.x (brak DHCP) i 127.x.x.x.
    $addresses = New-Object 'System.Collections.Generic.List[string]'
    try {
        foreach ($nic in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
            if ($nic.NetworkInterfaceType -eq [System.Net.NetworkInformation.NetworkInterfaceType]::Loopback) { continue }
            # Unknown: część kart wirtualnych (np. na Linuksie) nie podaje stanu, choć działa.
            $status = $nic.OperationalStatus
            if ($status -ne [System.Net.NetworkInformation.OperationalStatus]::Up -and $status -ne [System.Net.NetworkInformation.OperationalStatus]::Unknown) { continue }
            foreach ($unicast in $nic.GetIPProperties().UnicastAddresses) {
                $address = $unicast.Address
                if ($address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
                $text = $address.ToString()
                if ($text.StartsWith('169.254.') -or $text.StartsWith('127.')) { continue }
                if (-not $addresses.Contains($text)) { $addresses.Add($text) }
            }
        }
    } catch {
        # Brak danych o kartach = brak sieci (Detect zwróci Todo).
    }
    return , $addresses.ToArray()
}

function global:Get-NetworkStepStatus {
    # Stan sieci: @{ Ok = [bool]; Detail = '...' }.
    param([AllowNull()]$Context)
    $addresses = Get-NetworkStepAddresses
    if ($addresses.Count -eq 0) {
        return @{ Ok = $false; Detail = 'Brak adresu IPv4 - podłącz kabel sieciowy albo połącz się z Wi-Fi.' }
    }
    $shown = @($addresses | Select-Object -First 3) -join ', '
    $target = Get-NetworkStepTarget -Context $Context
    if ($null -eq $target) {
        return @{ Ok = $true; Detail = ('Adres IPv4: {0} (konfiguracja nie podaje serwera do sprawdzenia).' -f $shown) }
    }
    if (Test-StdTcpPort -HostName $target.Host -Port $target.Port -TimeoutMs 3000) {
        return @{ Ok = $true; Detail = ('Adres IPv4: {0}, serwer {1} odpowiada.' -f $shown, $target.Label) }
    }
    return @{ Ok = $false; Detail = ('Adres IPv4: {0}, ale serwer {1} nie odpowiada.' -f $shown, $target.Label) }
}

$networkStep = @{
    Id         = 'Network.WaitForNetwork'
    Title      = 'Czekanie na sieć'
    Phase      = 1
    Order      = 10
    TaskId     = 'WaitForNetwork'
    TimeoutMin = 10
    Detect     = {
        param($Context)
        $status = Get-NetworkStepStatus -Context $Context
        if ($status.Ok) { return @{ State = 'Done'; Detail = $status.Detail } }
        return @{ State = 'Todo'; Detail = $status.Detail }
    }
    Run        = {
        param($Context)
        $timeoutMin = 10.0
        if ($Context -is [System.Collections.IDictionary] -and $Context['Step'] -is [System.Collections.IDictionary] -and $null -ne $Context['Step']['TimeoutMin']) {
            $timeoutMin = [double]$Context['Step']['TimeoutMin']
        }
        $timeoutText = $timeoutMin.ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture)
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $lastDetail = $null
        while ($true) {
            $attemptStartMs = $watch.ElapsedMilliseconds
            $status = Get-NetworkStepStatus -Context $Context
            if ($status.Ok) {
                return @{ Result = 'Ok'; Detail = ('{0} Czekanie: {1} s.' -f $status.Detail, [int]$watch.Elapsed.TotalSeconds) }
            }
            if ($status.Detail -ne $lastDetail) {
                Write-StdLog -Source 'Network.WaitForNetwork' -Message ('Czekam na sieć: {0}' -f $status.Detail)
                $lastDetail = $status.Detail
            }
            if ($watch.Elapsed.TotalMinutes -ge $timeoutMin) {
                return @{ Result = 'Error'; Detail = ('Brak sieci po {0} min. {1}' -f $timeoutText, $status.Detail) }
            }
            # Następna próba 3 s po początku poprzedniej; w tym czasie co 250 ms sprawdzamy przerwanie.
            while (($watch.ElapsedMilliseconds - $attemptStartMs) -lt 3000) {
                if (Test-StdEngineCancel -Context $Context) {
                    return @{ Result = 'Cancelled'; Detail = 'Przerwano czekanie na sieć.' }
                }
                Start-Sleep -Milliseconds 250
            }
        }
    }
}
Register-StdStep @networkStep
Remove-Variable -Name networkStep -ErrorAction SilentlyContinue
