<#
.SYNOPSIS
    Dodaje konto do roli STD na serwerze IIS (technik albo administrator).

.DESCRIPTION
    Dodaje lokalne (albo domenowe) konto do grupy techników lub administratorów STD.
    Nazwy grup bierze z C:\ProgramData\STD-Server\settings.json (zapisanego przez Install-StdServer.ps1).
    Z -Create tworzy lokalne konto, jeśli go nie ma. Hasło podaje się jako SecureString albo wpisuje
    w oknie (dwa razy). Hasło nie jest nigdzie zapisywane ani wypisywane.

    Istniejącego konta skrypt nie zmienia (także hasła). Zmiana hasła:
    Set-LocalUser -Name <login> -Password (Read-Host -AsSecureString 'Nowe hasło')

.PARAMETER UserName
    Login. Konto lokalne: 'jan.kowalski'. Konto domenowe (tylko dodanie do grupy): 'FIRMA\jan.kowalski'.

.PARAMETER Role
    Technician (odczyt /std/, zapis raportów do /std/drop/) albo Admin (także /std-admin/).

.PARAMETER Create
    Utwórz lokalne konto, jeśli nie istnieje.

.PARAMETER Password
    Hasło nowego konta (SecureString). Bez tego parametru skrypt pyta o hasło.

.EXAMPLE
    .\Add-StdServerUser.ps1 -UserName jan.kowalski -Role Technician -Create

.EXAMPLE
    .\Add-StdServerUser.ps1 -UserName 'FIRMA\anna.nowak' -Role Admin
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [ValidateSet('Technician', 'Admin')]
    [string]$Role,
    [switch]$Create,
    [System.Security.SecureString]$Password,
    # Wewnętrzny: ustawiany przy samodzielnym ponownym uruchomieniu jako administrator (okno czeka na Enter).
    [Parameter(DontShow = $true)]
    [switch]$ElevatedRelaunch
)

$script:StdSrvDefaultGroups = @{ Technician = 'STD-Technicy'; Admin = 'STD-Administratorzy' }

function Get-StdSrvStateRoot {
    $base = $env:ProgramData
    if ([string]::IsNullOrEmpty($base)) { $base = [IO.Path]::GetTempPath() }
    return (Join-Path $base 'STD-Server')
}

function Split-StdSrvUserName {
    <#
        Rozbija login na @{ Domain; Name; IsLocal }. Bez domeny albo z nazwą tego komputera / '.' = konto lokalne.
        Zły login -> [STD-SRV-USER].
    #>
    param([string]$Value, [string]$ComputerName)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw '[STD-SRV-USER] Podaj login (-UserName).' }
    $parts = $Value.Split([char[]]@([char]'\'))
    if ($parts.Count -gt 2) { throw "[STD-SRV-USER] Login '$Value' jest nieprawidłowy." }
    $domain = $null
    $name = $parts[-1]
    if ($parts.Count -eq 2) { $domain = $parts[0] }
    $isLocal = ($null -eq $domain -or $domain -eq '.' -or $domain -eq $ComputerName)
    if ($name.Length -eq 0 -or $name -match '["/\[\]:;|=,+*?<>@]' -or $name.Trim() -ne $name -or $name.TrimEnd('.') -ne $name) {
        throw "[STD-SRV-USER] Login '$Value' zawiera niedozwolone znaki (\`" / \ [ ] : ; | = , + * ? < > @, spacja lub kropka na końcu)."
    }
    if ($isLocal -and $name.Length -gt 20) {
        throw "[STD-SRV-USER] Login konta lokalnego może mieć najwyżej 20 znaków ('$name' ma $($name.Length))."
    }
    if ($null -ne $domain -and $domain.Length -eq 0) { throw "[STD-SRV-USER] Login '$Value' jest nieprawidłowy." }
    return @{ Domain = $domain; Name = $name; IsLocal = $isLocal }
}

function Get-StdSrvRoleGroup {
    # Nazwa grupy dla roli: z settings.json instalki, a gdy go nie ma - domyślna.
    param([string]$Role, [string]$SettingsPath)
    $key = 'TechnicianGroup'
    if ($Role -eq 'Admin') { $key = 'AdminGroup' }
    if ($SettingsPath -and (Test-Path -LiteralPath $SettingsPath -PathType Leaf)) {
        $text = [IO.File]::ReadAllText($SettingsPath, (New-Object System.Text.UTF8Encoding($false)))
        try { $saved = ConvertFrom-Json -InputObject $text -ErrorAction Stop }
        catch { throw "[STD-SRV-SETTINGS] Plik ustawień '$SettingsPath' jest uszkodzony: $($_.Exception.Message)" }
        $prop = $saved.PSObject.Properties[$key]
        if ($null -ne $prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return [string]$prop.Value }
    }
    return $script:StdSrvDefaultGroups[$Role]
}

function Format-StdSrvArgument {
    # Cytowanie argumentu dla powershell.exe -File (reguły CommandLineToArgvW).
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    if ($Value.Length -gt 0 -and $Value -notmatch '[\s"]') { return $Value }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    $backslashes = 0
    foreach ($ch in $Value.ToCharArray()) {
        if ($ch -eq [char]'\') { $backslashes++; continue }
        if ($ch -eq [char]'"') {
            [void]$sb.Append([char]'\', ($backslashes * 2 + 1))
            [void]$sb.Append('"')
        }
        else {
            if ($backslashes -gt 0) { [void]$sb.Append([char]'\', $backslashes) }
            [void]$sb.Append($ch)
        }
        $backslashes = 0
    }
    if ($backslashes -gt 0) { [void]$sb.Append([char]'\', ($backslashes * 2)) }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Get-StdSrvUserRelaunchArguments {
    # Argumenty do ponownego uruchomienia jako administrator. Hasła NIE przekazujemy (okno zapyta o nie samo).
    param([string]$ScriptPath, [string]$UserName, [string]$Role, [bool]$Create, [switch]$MarkElevated)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($a in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File')) { $list.Add($a) }
    $list.Add((Format-StdSrvArgument -Value $ScriptPath))
    $list.Add('-UserName'); $list.Add((Format-StdSrvArgument -Value $UserName))
    $list.Add('-Role'); $list.Add($Role)
    if ($Create) { $list.Add('-Create') }
    if ($MarkElevated) { $list.Add('-ElevatedRelaunch') }
    return , $list.ToArray()
}

function Test-StdSrvIsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-StdSrvSecureStringEqual {
    # Porównanie dwóch haseł; jawny tekst tylko na chwilę, potem zerowany.
    param([System.Security.SecureString]$A, [System.Security.SecureString]$B)
    $pa = [IntPtr]::Zero
    $pb = [IntPtr]::Zero
    try {
        $pa = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($A)
        $pb = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($B)
        $la = [Runtime.InteropServices.Marshal]::ReadInt32($pa, -4)
        $lb = [Runtime.InteropServices.Marshal]::ReadInt32($pb, -4)
        if ($la -ne $lb) { return $false }
        $diff = 0
        for ($i = 0; $i -lt $la; $i++) {
            $diff = $diff -bor ([Runtime.InteropServices.Marshal]::ReadByte($pa, $i) -bxor [Runtime.InteropServices.Marshal]::ReadByte($pb, $i))
        }
        return ($diff -eq 0)
    }
    finally {
        if ($pa -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pa) }
        if ($pb -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pb) }
    }
}

function Read-StdSrvNewPassword {
    # Pyta o hasło dwa razy. W trybie bez okna (np. CI) wymaga -Password.
    param([string]$Name)
    if (-not [Environment]::UserInteractive) {
        throw "[STD-SRV-PASSWORD] Brak hasła dla nowego konta $Name. Podaj -Password (SecureString)."
    }
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $first = Read-Host -AsSecureString "Hasło dla $Name"
        $second = Read-Host -AsSecureString 'Powtórz hasło'
        if ($first.Length -eq 0) { Write-Host 'Hasło nie może być puste.' -ForegroundColor Yellow; continue }
        if (Test-StdSrvSecureStringEqual -A $first -B $second) { $second.Dispose(); return $first }
        Write-Host 'Hasła się różnią. Spróbuj jeszcze raz.' -ForegroundColor Yellow
    }
    throw '[STD-SRV-PASSWORD] Hasła się różniły - konto nie zostało utworzone.'
}

function Add-StdSrvGroupMember {
    # Dodaje członka do lokalnej grupy. Zwraca $true (dodano) albo $false (już był w grupie).
    param([string]$Group, [string]$Member)
    try {
        Add-LocalGroupMember -Group $Group -Member $Member -ErrorAction Stop
        return $true
    }
    catch {
        if ("$($_.FullyQualifiedErrorId)" -like 'MemberExists*' -or $_.Exception.GetType().Name -eq 'MemberExistsException') { return $false }
        throw "[STD-SRV-GROUP] Nie udało się dodać $Member do grupy ${Group}: $($_.Exception.Message)"
    }
}

function Invoke-StdSrvAddUser {
    param([string]$UserName, [string]$Role, [bool]$Create, [System.Security.SecureString]$Password)
    if (-not (Get-Command -Name Get-LocalUser -ErrorAction SilentlyContinue)) {
        throw '[STD-SRV-USER] Brak modułu Microsoft.PowerShell.LocalAccounts. Uruchom skrypt w 64-bitowym Windows PowerShell 5.1.'
    }
    $settingsPath = Join-Path (Get-StdSrvStateRoot) 'settings.json'
    $group = Get-StdSrvRoleGroup -Role $Role -SettingsPath $settingsPath
    $groupObj = Get-LocalGroup -Name $group -ErrorAction SilentlyContinue
    if ($null -eq $groupObj) {
        throw "[STD-SRV-GROUP] Grupa '$group' nie istnieje. Najpierw uruchom Install-StdServer.ps1."
    }

    $who = Split-StdSrvUserName -Value $UserName -ComputerName $env:COMPUTERNAME
    $created = $false
    if ($who.IsLocal) {
        $member = $who.Name
        $user = Get-LocalUser -Name $who.Name -ErrorAction SilentlyContinue
        if ($null -eq $user) {
            if (-not $Create) {
                throw "[STD-SRV-USER] Konto lokalne '$($who.Name)' nie istnieje. Dodaj -Create, aby je utworzyć."
            }
            if ($null -eq $Password -or $Password.Length -eq 0) { $Password = Read-StdSrvNewPassword -Name $who.Name }
            $description = 'STD - technik'
            if ($Role -eq 'Admin') { $description = 'STD - administrator' }
            try {
                # Hasło bez wygasania: logowanie Basic nie ma ekranu zmiany hasła, wygasłe hasło = nagła odmowa dostępu.
                [void](New-LocalUser -Name $who.Name -Password $Password -Description $description -PasswordNeverExpires -AccountNeverExpires -ErrorAction Stop)
            }
            catch {
                throw "[STD-SRV-USER] Nie udało się utworzyć konta '$($who.Name)': $($_.Exception.Message) (sprawdź wymagania co do długości i złożoności hasła)."
            }
            $created = $true
            Write-Host "Utworzono konto lokalne $($who.Name)."
        }
        else {
            if ($Create) { Write-Host "Konto $($who.Name) już istnieje - nie zmieniam go (także hasła)." }
            if (-not $user.Enabled) { Write-Host "UWAGA: konto $($who.Name) jest wyłączone. Włącz je: Enable-LocalUser -Name $($who.Name)" -ForegroundColor Yellow }
        }
    }
    else {
        if ($Create) {
            throw "[STD-SRV-USER] Konta domenowego '$UserName' nie da się utworzyć tym skryptem. Utwórz je w domenie i uruchom skrypt bez -Create."
        }
        $member = $UserName
    }

    $added = Add-StdSrvGroupMember -Group $group -Member $member
    if ($added) { Write-Host "Dodano $member do grupy $group." }
    else { Write-Host "$member już jest w grupie $group." }

    $roleText = 'technik: odczyt /std/, zapis raportów do /std/drop/'
    if ($Role -eq 'Admin') { $roleText = 'administrator: odczyt /std/, zapis do /std/drop/, pełny dostęp do /std-admin/' }
    Write-Host "Rola: $roleText"
    if ($who.IsLocal) { Write-Host "Logowanie w STD: login '$($who.Name)' (konto lokalne serwera)." }
    else { Write-Host "Logowanie w STD: login '$UserName'." }

    return [pscustomobject]@{ UserName = $member; Role = $Role; Group = $group; Created = $created; AddedToGroup = $added }
}

if ($MyInvocation.InvocationName -ne '.') {
    $stdExitCode = 0
    try {
        if ($env:OS -ne 'Windows_NT') { throw '[STD-SRV-OS] Ten skrypt działa tylko w Windows. Na Linuksie użyj std-server add-tech / add-admin.' }
        $stdIsAdmin = Test-StdSrvIsAdmin
        $needs64 = ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)
        if ($needs64 -or -not $stdIsAdmin) {
            if ($null -ne $Password) {
                Write-Host 'Hasło nie jest przekazywane do nowego okna - zostaniesz o nie zapytany ponownie.' -ForegroundColor Yellow
            }
            $exe = Join-Path $env:windir 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $sysnative = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
            if ($needs64 -and (Test-Path -LiteralPath $sysnative)) { $exe = $sysnative }
            $argList = Get-StdSrvUserRelaunchArguments -ScriptPath $PSCommandPath -UserName $UserName -Role $Role -Create ([bool]$Create) -MarkElevated:(-not $stdIsAdmin)
            try {
                if ($stdIsAdmin) { $p = Start-Process -FilePath $exe -ArgumentList $argList -NoNewWindow -Wait -PassThru -ErrorAction Stop }
                else {
                    Write-Host 'Skrypt wymaga uprawnień administratora - uruchamiam go ponownie jako administrator...'
                    $p = Start-Process -FilePath $exe -ArgumentList $argList -Verb RunAs -Wait -PassThru -ErrorAction Stop
                }
            }
            catch {
                throw '[STD-SRV-ADMIN] Skrypt wymaga uprawnień administratora, a uruchomienie jako administrator zostało anulowane albo się nie udało.'
            }
            exit $p.ExitCode
        }
        $ErrorActionPreference = 'Stop'
        Invoke-StdSrvAddUser -UserName $UserName -Role $Role -Create ([bool]$Create) -Password $Password
    }
    catch {
        $stdExitCode = 1
        Write-Host ('BŁĄD: ' + $_.Exception.Message) -ForegroundColor Red
    }
    if ($ElevatedRelaunch) { [void](Read-Host 'Naciśnij Enter, aby zamknąć to okno') }
    exit $stdExitCode
}
