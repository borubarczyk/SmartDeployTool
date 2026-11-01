# STD 4 - specyfikacja techniczna (etap 1 i serwer)

Ten dokument jest umową między częściami STD 4. Nazwy funkcji, parametrów, pól JSON i kodów błędów
muszą się z nim zgadzać. Plan i decyzje: [PLAN.md](PLAN.md).

## 0. Zasady wspólne

### 0.1 Zgodność z Windows PowerShell 5.1

Kod modułów i narzędzi działa w **Windows PowerShell 5.1** (.NET Framework 4.8) i musi też działać
w PowerShell 7 (testy lokalne). Zabronione, bo nie działa w 5.1:

- `??`, `??=`, `?.`, `?[]`, operator trójargumentowy `a ? b : c`, `&&` / `||` w potokach;
- `ConvertFrom-Json -AsHashtable`, `ConvertFrom-Json -Depth`, `ForEach-Object -Parallel`, `Join-String`,
  `Get-Error`, `$PSStyle`, `-Encoding utf8BOM` / `utf8NoBOM`;
- `[Convert]::ToHexString`, `[SHA256]::HashData`, `[RandomNumberGenerator]::Fill` / `GetBytes(int)`,
  `CryptographicOperations.FixedTimeEquals`, `Rfc2898DeriveBytes::Pbkdf2`, `X509Certificate2.CreateFromPem`,
  `Convert.FromHexString`.

Dozwolone (są w .NET 4.8 / PS 5.1): `[type]::new()`, klasy `System.Security.Cryptography.*` z konstruktorami
(`Rfc2898DeriveBytes(byte[] password, byte[] salt, int iter, HashAlgorithmName)`), `RSACertificateExtensions`,
`CertificateRequest` (.NET 4.7.2+), `Import-PowerShellDataFile`.

Inne zasady:

- `ConvertTo-Json` zawsze z `-Depth 30` (domyślnie 2 i dane po cichu przepadają).
- Pliki zapisujemy przez `[System.IO.File]::WriteAllText/WriteAllBytes` i `New-Object System.Text.UTF8Encoding($false)`
  (JSON bez BOM). Nie używamy `Set-Content -Encoding UTF8`: w 5.1 dopisuje BOM, a w 7 go nie dopisuje.
- Lista generyczna zwracana z funkcji: `return ,$list` (bez `@()` wokół `List[T]`; w 5.1 to błąd „Argument types do not match”).
- Porównania tekstu: `-eq` / `-contains` nie rozróżniają wielkości liter. Gdzie ma to znaczenie (np. odcisk certyfikatu),
  porównujemy po `ToUpperInvariant()`.
- Żadnego WPF w `Modules\` (WPF będzie w osobnym `GUI\` od etapu 2). Moduły muszą działać bez okna i w osobnym wątku.
- Komunikaty dla użytkownika po polsku. Błędy rzucamy przez `throw` z komunikatem zaczynającym się od kodu w nawiasie
  kwadratowym, np. `throw "[STD-SIG-INVALID] Podpis konfiguracji jest nieprawidłowy."`. Testy sprawdzają kod.
- Żadna funkcja nie pisze haseł, tokenów ani nagłówków `Authorization` do logu ani do plików.
- **Dane logowania do domeny i hasło lokalnego administratora nie mogą trafić do żadnego pliku** (konfiguracja,
  `state.json`, log, raport, cache). W konfiguracji v4 nie ma na nie pól; walidacja odrzuca dokument, który je ma.

### 0.2 Pliki

- `.ps1`, `.psm1`, `.psd1`: UTF-8 **z BOM**, końce linii LF. Bez BOM Windows PowerShell 5.1 czyta je jako ANSI i polskie znaki się psują.
- `.sh`, `.md`, `.json`, `.yml`, szablony nginx: UTF-8 bez BOM, LF. Skrypty `.sh` i `std-server` z bitem wykonywalnym (`git update-index --chmod=+x`).
- Testy Pester 5 (`v4\Tests\*.Tests.ps1`) uruchamia istniejący workflow `pester.yml` (Windows PowerShell 5.1, Pester 5.6.1,
  wszystkie `*.Tests.ps1` w repo). Testy importują moduł przez `Import-Module "$PSScriptRoot\..\STD.psd1" -Force`,
  dane lokalne kierują do `TestDrive:` przez `$env:STD_DATA_ROOT` i nie wymagają administratora, sieci ani magazynu certyfikatów
  (certyfikaty testowe tworzymy w pamięci przez `CertificateRequest`).

### 0.3 Moduł STD

- `v4\STD.psd1`: `RootModule = 'STD.psm1'`, `ModuleVersion = '4.0.0'`, `PowerShellVersion = '5.1'`,
  `CompatiblePSEditions = @('Desktop','Core')`, `FunctionsToExport = '*'`, `GUID` stały.
- `v4\STD.psm1` dot-source'uje w tej kolejności: `Modules\Std.Core.ps1`, `Std.Http.ps1`, `Std.Crypto.ps1`,
  `Std.Config.ps1`, `Std.Legacy.ps1`. Wszystkie funkcje i zmienne `$script:` są więc w jednym zakresie modułu (wspólne
  odbiorniki logu i maski). Na końcu `Export-ModuleMember -Function *-Std*`.
- Wersję zwraca `Get-StdVersion` (z manifestu modułu), nigdzie nie wpisujemy jej drugi raz.

## 1. Lokalne dane na komputerze

Katalog główny: `Get-StdDataRoot`. Jest to `$env:STD_DATA_ROOT`, jeśli ustawiona (testy); w przeciwnym razie
`$env:ProgramData\STD`. Gdy `ProgramData` nie ma (Linux), używamy `[IO.Path]::GetTempPath()` + `STD`.

| `Get-StdPath -Kind` | Ścieżka | Zawartość |
| --- | --- | --- |
| `Root` | `<root>` | |
| `App` | `<root>\App` | kopie narzędzia (`App\<wersja>\`), etap 2 |
| `Cache` | `<root>\Cache` | `clients.sdt`, `<klient>\release.sdt` (podpisane koperty, sprawdzane przy każdym odczycie) |
| `Logs` | `<root>\Logs` | `<deploymentId>\deploy.log`, `events.jsonl`, etap 2 |
| `Outbox` | `<root>\Outbox` | raporty czekające na wysłanie, etap 4 |
| `State` | `<root>\state.json` | stan bieżącego wdrożenia (bez haseł), etap 2 |
| `Revisions` | `<root>\revisions.json` | najwyższa widziana rewizja: `{ "clients": 7, "lizard": 42 }` |
| `Prefs` | `<root>\prefs.json` | motyw, ostatni klient/lokalizacja/profil |
| `Bootstrap` | `<root>\bootstrap.local.json` | adres serwera wpisany przez technika |

Aktualizacja STD (nowy `App\<wersja>\`) niczego poza `App\` nie rusza, więc ustawienia lokalne przechodzą do nowej wersji.

## 2. Formaty

### 2.1 Koperta podpisu (`*.sdt`)

```json
{ "kind": "std.signed", "alg": "RS256", "kid": "<odcisk SHA-1 certyfikatu, HEX wielkimi literami>",
  "payload": "<base64 bajtów UTF-8 dokumentu>", "sig": "<base64 podpisu RSA PKCS#1 v1.5 SHA-256 bajtów payload>" }
```

Podpisujemy dokładne bajty `payload` (bez kanonizacji JSON). Weryfikacja: `kid` musi należeć do zaufanych certyfikatów
**i** podpis musi się zgadzać z kluczem publicznym tego certyfikatu.

### 2.2 Zaufane certyfikaty

`v4\Modules\Std.TrustedSigners.psd1` (część kodu, podpisywana razem z nim):

```powershell
@{ Signers = @(
    # @{ Name = 'Borys Kaleta - główny'; Thumbprint = 'AB12...'; Certificate = 'MIIC...base64 DER...' }
) }
```

Pusta lista jest dozwolona w repozytorium; wtedy pobranie konfiguracji kończy się błędem `[STD-SIG-NOTRUST]` z instrukcją
uruchomienia `Tools\New-StdSigningCertificate.ps1`. Wszystkie funkcje weryfikujące mają parametr `-TrustedCertificates`
(testy go podają); bez niego biorą listę z tego pliku. Nie ma zmiennej środowiskowej ani pliku obok skryptu, które mogłyby
dodać zaufany certyfikat.

### 2.3 Plik startowy `bootstrap.json` (pendrive, obok `Start-STD.cmd`)

```json
{ "server": "https://pobieranie.firma.pl/std/", "mirrors": [ "\\\\srv-files\\STD$\\std\\" ], "auth": { "mode": "basic" }, "proxy": null }
```

Nie jest sekretem ani źródłem zaufania. `server` musi zaczynać się od `https://` i kończyć `/` (normalizujemy dopisując `/`).
`mirrors`: ścieżki UNC albo lokalne do kopii tego samego drzewa co `/std/` na serwerze.

### 2.4 Drzewo na serwerze (to samo w IIS, nginx i kopii lustrzanej)

```text
/std/                          odczyt: technicy i administratorzy (GET, HEAD)
  clients.sdt                  podpisany indeks klientów
  clients/<id>/release.sdt     podpisana konfiguracja klienta
  clients/<id>/history/<rev>.sdt   archiwum opublikowanych rewizji
  files/<id>/...               instalatory klienta (sprawdzane po SHA-256)
  files/common/...             instalatory wspólne
  tool/STD-<wersja>.zip        wydania narzędzia
  drop/                        tylko zapis (PUT) dla techników: raporty i depozyty; bez odczytu i listowania
/std-admin/                    to samo drzewo, tylko administratorzy: GET, PUT, DELETE, MKCOL (publikacja)
```

### 2.5 Indeks klientów (payload `clients.sdt`)

```json
{ "schemaVersion": 4, "kind": "std.clients", "revision": 7, "publishedAt": "2026-11-01T20:00:00Z",
  "clients": [ { "id": "lizard", "name": "Lizard-Tech", "release": "clients/lizard/release.sdt" } ] }
```

### 2.6 Konfiguracja klienta (payload `release.sdt`, `schemaVersion: 4`)

Pełny przykład: `v4\Tests\Fixtures\release.v4.json`. Pola:

| Pole | Typ | Wymagane | Opis |
| --- | --- | --- | --- |
| `schemaVersion` | int | tak | `4` |
| `kind` | string | tak | `std.release` |
| `client.id` | string | tak | `^[a-z0-9][a-z0-9-]{1,39}$` |
| `client.name` | string | tak | niepusty |
| `revision` | int ≥ 1 | po publikacji | ustawia `Publish-StdRelease` |
| `publishedAt`, `publishedBy` | string | po publikacji | ISO 8601 UTC, login |
| `changeNote` | string | nie | |
| `expiresAt` | string | nie | ISO 8601 UTC; domyślnie publikacja + 90 dni |
| `minClientVersion` | string | nie | `[version]`, np. `4.0.0` |
| `allowOffline` | bool | nie (false) | czy wolno użyć kopii z cache / pendrive'a |
| `requireSha256` | bool | nie (false) | czy każda pozycja katalogu z `file` / `url` musi mieć `sha256` |
| `tool` | obiekt | nie | `{ version, url, sha256 }` - wydanie narzędzia (etap 4) |
| `sources.<nazwa>` | obiekt | tak (≥1) | `{ "type": "https", "url": "https://.../", "auth": "session" }` albo `{ "type": "unc", "path": "\\\\srv\\STD$\\files\\lizard\\", "auth": "session" }`; `auth`: `session` / `none` |
| `catalog.<klucz>` | obiekt | tak (≥1) | klucz `^[A-Za-z0-9][A-Za-z0-9._+-]{0,63}$`; pola niżej |
| `defaults` | obiekt | nie | `tasks`, `bloatware`, `uninstall`, `systemSettings`, `scripts` |
| `sites[]` | tablica | tak (≥1) | lokalizacje, pola niżej |
| `secrets` | obiekt | nie | zaszyfrowane sekrety (2.7) |
| `escrow` | obiekt | nie | `{ "drop": "drop/", "recipientCert": "<base64 DER>" }` (etap 4) |

Pozycja katalogu `catalog.<klucz>`:

| Pole | Opis |
| --- | --- |
| `name` | nazwa wyświetlana (wymagana) |
| `file` | nazwa pliku w źródle (np. `7z2408-x64.exe`), szukana w kolejnych źródłach z `site.sourceOrder` |
| `url` | pełny adres `https://` instalatora (zamiast `file`) |
| `winget`, `wingetArgs` | identyfikator pakietu winget i dodatkowe argumenty (np. `--scope machine`) |
| `args` | argumenty cichej instalacji (`file` / `url`); mogą zawierać `{{secret:nazwa}}` |
| `sha256` | 64 znaki HEX; wymagany przy `file` / `url`, gdy konfiguracja ma `requireSha256: true` |
| `detect` | `{ "uninstallName": "7-Zip*", "minVersion": "24.08" }` albo `{ "path": "C:\\Program Files\\X\\x.exe" }` |
| `timeoutMin` | 1-240, domyślnie 30 |
| `successCodes` | tablica int, domyślnie `[0, 1641, 3010]` |
| `whenDetected` | `skip` (domyślnie) albo `reinstall` |
| `sourceOrder` | własna kolejność źródeł dla tej pozycji (np. antywirus z osobnego serwera); domyślnie `site.sourceOrder` |

Musi być co najmniej jedno z: `file`, `url`, `winget`. Gdy jest `file` i `winget`, instalator (etap 3) bierze plik ze źródeł,
a winget jest zapasem.

`defaults`:

| Pole | Opis |
| --- | --- |
| `tasks` | `{ "<TaskId>": bool }`, identyfikatory z listy 2.8 |
| `bloatware` | wzorce nazw aplikacji Sklepu, np. `"*TikTok*"` |
| `uninstall` | preinstalowane programy do usunięcia w kroku „Deinstalator”: `[ { "name": "McAfee*", "args": "" } ]` |
| `systemSettings` | flagi `{ "DisableFastStartup": true, ... }` i `registry: [ { "hive": "HKLM"\|"HKCU"\|"DefaultUser", "path": "Software\\...", "name": "...", "type": "String"\|"ExpandString"\|"DWord"\|"QWord"\|"MultiString"\|"Binary", "value": ... } ]` |
| `scripts` | `[ { "file": "Drukarki.ps1", "sha256": "...", "args": "" } ]` |

`sites[]`:

| Pole | Opis |
| --- | --- |
| `id`, `name` | `id` jak `client.id`, unikalne |
| `sourceOrder` | nazwy z `sources`, np. `["lan","web"]` |
| `wifi` | `[ { "ssid": "Firma", "auth": "WPA2PSK"\|"WPA3SAE"\|"open", "keyRef": "wifi-firma" } ]` (klucz w `secrets`) albo `[ { "file": "Firma.xml" } ]` (dawny format) |
| `computerName` | `{ "pattern": "KRK-{serial}", "maxLength": 15, "editable": true }`; dozwolone wstawki `{serial}`, `{site}`, `{random}` |
| `domainJoin` | `{ "domain": "firma.local", "ou": "OU=...,DC=firma,DC=local" }` - **bez loginu i hasła** |
| `localAdmin` | `{ "username": "LokalnyIT" }` - **bez hasła** |
| `bitLocker` | `{ "escrow": [ "ad", "server" ] }` |
| `profiles.<nazwa>` | `{ "apps": [klucze katalogu], "tasks": { ... }, "scripts": [ ... ], "locked": [ "JoinDomain", "computerName", "apps" ] }` |

Dziedziczenie: `defaults` → lokalizacja → profil. Obiekty łączone głęboko, tablice zastępowane, `null` = dziedzicz.

### 2.7 Sekrety (`secrets`, tryb `sealed`)

```json
"secrets": { "mode": "sealed",
  "kdf": { "alg": "pbkdf2-sha256", "iter": 600000, "salt": "<base64 16 B>" },
  "check": "<base64 HMAC-SHA256(macKey, 'std-check')>",
  "items": { "tv-token": { "iv": "<base64 16 B>", "ct": "<base64>", "mac": "<base64 HMAC-SHA256(macKey, iv||ct)>" } } }
```

- Klucz: PBKDF2-SHA256(UTF-8 hasła wdrożeniowego klienta, salt, iter) → 64 bajty: `encKey` = 0-31, `macKey` = 32-63.
- Szyfr: AES-256-CBC, PKCS7, losowe IV; najpierw sprawdzenie `mac` (porównanie w stałym czasie), potem odszyfrowanie.
- W tekstach konfiguracji sekret wstawiamy jako `{{secret:nazwa}}`; podstawienie tuż przed użyciem, wartość od razu idzie do maski logu.
- **W sekretach nie ma i nie będzie loginu ani hasła do domeny ani hasła lokalnego administratora.**

### 2.8 Identyfikatory zadań (`tasks`)

`WaitForNetwork`, `SuspendSleep`, `ConnectWiFi`, `RemoveBloatware`, `UninstallPreinstalled`, `UninstallM365`, `UninstallOneDrive`,
`InstallApps`, `CreateLocalAdmin`, `RenameComputer`, `JoinDomain`, `SystemSettings`, `RunScripts`, `HardwareAudit`,
`EnableBitLocker`, `RebootAtEnd`. Inne klucze w `tasks` to błąd walidacji.

## 3. Funkcje modułów (etap 1)

Parametry w nawiasach kwadratowych są opcjonalne. Wszystkie funkcje mają `[CmdletBinding()]`.

### 3.1 `Modules\Std.Core.ps1`

| Funkcja | Opis |
| --- | --- |
| `Get-StdVersion` | `[version]` modułu |
| `Get-StdDataRoot` | jak w sekcji 1; tworzy katalog, jeśli go nie ma |
| `Get-StdPath -Kind <Root\|App\|Cache\|Logs\|Outbox\|State\|Revisions\|Prefs\|Bootstrap>` | pełna ścieżka; katalogi tworzy, plików nie |
| `ConvertTo-StdHashtable -InputObject <obj>` | głęboko: `PSCustomObject` / `IDictionary` → `[ordered]@{}`, tablice → `object[]` (także jednoelementowe i puste), reszta bez zmian |
| `ConvertFrom-StdJson -Json <string>` | `ConvertFrom-Json` + `ConvertTo-StdHashtable`; błąd → `[STD-JSON] ...` |
| `ConvertTo-StdJson -InputObject <obj> [-Compress]` | `ConvertTo-Json -Depth 30` |
| `Write-StdJsonFile -Path <string> -InputObject <obj>` | atomowo: zapis do `<path>.tmp` (UTF-8 bez BOM), potem podmiana (`[IO.File]::Replace` albo `Move`) |
| `Read-StdJsonFile -Path <string>` | brak pliku → `$null`; uszkodzony JSON → zmiana nazwy na `<path>.bad`, `Write-StdLog -Level Warn`, zwraca `$null` |
| `Write-StdLog -Message <string> [-Level Info\|Warn\|Error\|Debug] [-Source <string>]` | maskuje, tworzy rekord `@{ Time; Level; Message; Source }` i przekazuje do każdego odbiornika; wyjątek odbiornika nie przerywa pracy. `Debug` tylko gdy `$env:STD_DEBUG -eq '1'` |
| `Register-StdLogSink -Name <string> -ScriptBlock <sb>` / `Unregister-StdLogSink -Name` | odbiornik dostaje rekord jako `$args[0]` |
| `Register-StdFileLogSink -Name <string> -Path <string>` | `StreamWriter` UTF-8 bez BOM, `AutoFlush`, linia `yyyy-MM-dd HH:mm:ss [LEVEL] wiadomość`; `Unregister-StdLogSink` zamyka plik |
| `Register-StdLogMask -Value <string>` | wartości krótsze niż 3 znaki ignorowane |
| `Clear-StdLogMasks` | |
| `Protect-StdLogText -Text <string>` | podmienia każdą zarejestrowaną wartość (najdłuższe najpierw, bez względu na wielkość liter) na `***` |
| `Test-StdIsAdmin` | `[bool]`; na Linuksie: `id -u` = 0 |
| `Start-StdElevated -ScriptPath <string> [-Arguments <string[]>] [-WorkingDirectory <string>]` | uruchamia `powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "<ScriptPath>" <args>` z `-Verb RunAs`; zwraca `$true`, gdy proces wystartował, `$false` gdy użytkownik odmówił UAC (błąd 1223). Argumenty z cudzysłowami są poprawnie cytowane. |

### 3.2 `Modules\Std.Http.ps1`

| Funkcja | Opis |
| --- | --- |
| `Initialize-StdNetwork` | `SecurityProtocol` = TLS 1.2, plus TLS 1.3 tylko gdy wartość `Tls13` istnieje w `[Net.SecurityProtocolType]` **i** system to build ≥ 20348 (Windows 11 / Server 2022). Bez wyjątków na starszym .NET. Idempotentna |
| `New-StdSession -Credential <pscredential>` albo `-UserName <string> -Password <SecureString>` | `[pscustomobject]@{ Mode = 'basic'; UserName; Password (SecureString) }`; hasło nigdy jawnie w obiekcie |
| `Get-StdBasicAuthHeader -Session <obj>` | `"Basic <base64(user:pass)>"`; jawne hasło tylko na czas składania, potem zerowane |
| `Invoke-StdHttpRequest -Uri <string> [-Method GET\|HEAD\|PUT\|MKCOL\|DELETE] [-Session <obj>] [-Body <byte[]>] [-ContentType <string>] [-OutFile <string>] [-TimeoutSec 30] [-AllowStatus <int[]>]` | `HttpWebRequest`; tylko `https://` (inaczej `[STD-HTTP-INSECURE]`); nagłówek Basic od razu (bez czekania na 401); `AllowAutoRedirect = $false`; `UserAgent = "STD/<wersja>"`. Wynik `@{ StatusCode; Bytes; Headers }` (`Bytes` puste, gdy `-OutFile`). Status spoza 2xx i `-AllowStatus` → `[STD-HTTP-<kod>] <metoda> <uri>`; brak połączenia / limit czasu → `[STD-HTTP-NET] ...`. Nagłówka `Authorization` i hasła nie ma w żadnym komunikacie |
| `Get-StdBytes -Location <string> [-Session <obj>] [-TimeoutSec 30]` | `https://` → `Invoke-StdHttpRequest`; ścieżka UNC → najpierw `Test-StdTcpPort <host> 445` (3 s), potem `[IO.File]::ReadAllBytes`; ścieżka lokalna → `ReadAllBytes`; brak pliku → `[STD-SRC-NOTFOUND] <location>` |
| `Join-StdLocation -Base <string> -Child <string>` | dla `https://` łączy przez `/` (bez podwójnych ukośników, segmenty kodowane `EscapeDataString`); dla ścieżek łączy przez `Join-Path`, zamieniając `/` na `\` (na Linuksie odwrotnie) |
| `Test-StdTcpPort -HostName <string> -Port <int> [-TimeoutMs 3000]` | `[bool]`, `TcpClient.BeginConnect` + `WaitOne` |

### 3.3 `Modules\Std.Crypto.ps1`

| Funkcja | Opis |
| --- | --- |
| `Get-StdSha256 -Path <string>` / `-Bytes <byte[]>` | małe litery HEX |
| `Test-StdFileHash -Path <string> -Expected <string>` | `[bool]`, bez względu na wielkość liter |
| `New-StdSignedEnvelope -Payload <byte[]> -Certificate <X509Certificate2>` | certyfikat z kluczem prywatnym RSA; zwraca JSON koperty (tekst) |
| `Read-StdSignedEnvelope -Json <string> \| -Bytes <byte[]> [-TrustedCertificates <X509Certificate2[]>]` | zwraca `@{ Kid; PayloadBytes; PayloadText; Signer }`. Błędy: `[STD-SIG-FORMAT]` zły JSON / brak pól / `kind` ≠ `std.signed`, `[STD-SIG-ALG]` `alg` ≠ `RS256`, `[STD-SIG-NOTRUST]` pusta lista zaufanych, `[STD-SIG-UNTRUSTED]` nieznany `kid`, `[STD-SIG-INVALID]` podpis się nie zgadza |
| `New-StdSealedStore -Password <SecureString> [-Iterations 600000]` | pusty magazyn sekretów (hashtable jak w 2.7) |
| `Set-StdSealedItem -Store <hashtable> -Name <string> -Value <SecureString> -Password <SecureString>` | dodaje/podmienia pozycję; najpierw `Test-StdSealedPassword` |
| `Test-StdSealedPassword -Store <hashtable> -Password <SecureString>` | `[bool]` (pole `check`) |
| `Unprotect-StdSealedItem -Store <hashtable> -Name <string> -Password <SecureString>` | zwraca `[string]` i rejestruje go w masce logu. Błędy: `[STD-SEC-PASSWORD]`, `[STD-SEC-MISSING]`, `[STD-SEC-TAMPER]` |
| `Get-StdSecretReference -Text <string>` | nazwy z `{{secret:nazwa}}` (unikalne) |
| `Expand-StdSecretTemplate -Text <string> -Resolver <scriptblock>` | podmienia `{{secret:nazwa}}` wynikiem `& $Resolver nazwa`, każdą wartość dodaje do maski logu |

### 3.4 `Modules\Std.Config.ps1`

| Funkcja | Opis |
| --- | --- |
| `Get-StdTrustedCertificates` | `X509Certificate2[]` z `Std.TrustedSigners.psd1` (puste wpisy pomijane) |
| `Get-StdBootstrap [-Server <string>] [-ScriptRoot <string>]` | kolejność: `-Server`, `<ScriptRoot>\bootstrap.json`, `Get-StdPath Bootstrap`; zwraca hashtable 2.3 albo `$null`; adres nie-`https` → `[STD-CFG-BOOTSTRAP]` |
| `Save-StdBootstrapLocal -Bootstrap <hashtable>` | zapis do `Get-StdPath Bootstrap` |
| `Get-StdSeenRevision -Key <string>` / `Set-StdSeenRevision -Key <string> -Revision <int>` | `revisions.json`; `Set` nigdy nie zmniejsza wartości |
| `Get-StdClientIndex -Bootstrap <hashtable> [-Session <obj>] [-TrustedCertificates ...] [-UsbRoot <string>]` | jak `Get-StdRelease`, dla `clients.sdt`, klucz rewizji `clients`; offline (cache / pendrive) dozwolony zawsze; wynik `@{ Index; Source; Revision; Attempts }` |
| `Get-StdRelease -ClientId <string> -Bootstrap <hashtable> [-Session <obj>] [-TrustedCertificates ...] [-UsbRoot <string>]` | źródła po kolei: serwer (`<server>clients/<id>/release.sdt`), kopie lustrzane, cache (`Cache\<id>\release.sdt`), pendrive (`<UsbRoot>\offline\<id>\release.sdt`). Dla każdego: pobranie → `Read-StdSignedEnvelope` → kontrole niżej. Pierwsze poprawne wygrywa, trafia do cache, `Set-StdSeenRevision`. Wynik `@{ Release (ordered hashtable); Source ('server'\|'mirror'\|'cache'\|'usb'); Revision; Kid; Attempts (lista @{ Source; Location; Error }) }`. Gdy żadne nie pasuje: `[STD-CFG-UNAVAILABLE]` z listą prób |
| `Test-StdRelease -Release <hashtable> [-Unpublished]` | lista błędów `string[]` (pusta = OK), reguły niżej; `-Unpublished` pomija `revision` / `publishedAt` |
| `Merge-StdConfig -Base <hashtable> -Override <hashtable>` | nowy obiekt; hashtable łączone rekurencyjnie, tablice i wartości proste zastępowane, `$null` w `Override` = zostaw `Base` |
| `Resolve-StdEffectiveConfig -Release <hashtable> -SiteId <string> -ProfileName <string>` | `@{ ClientId; ClientName; Revision; Sources; Catalog; Site (bez profiles); ProfileName; Apps; Tasks; Bloatware; Uninstall; SystemSettings; Scripts; Locked }`: `Tasks` = `defaults.tasks` + `profile.tasks`; `Scripts` = `defaults.scripts` + `profile.scripts`; nieznana lokalizacja / profil → `[STD-CFG-SITE]` / `[STD-CFG-PROFILE]` |

Kontrole w `Get-StdRelease` (każda porażka = to źródło odpada, próba idzie do `Attempts`, sprawdzamy następne):

1. `kind` = `std.release`, `schemaVersion` = 4, `client.id` = `-ClientId`;
2. `minClientVersion` ≤ `Get-StdVersion`, inaczej `[STD-CFG-CLIENT-OLD]`;
3. `expiresAt` + 5 minut ≥ teraz (UTC), inaczej `[STD-CFG-EXPIRED]`;
4. `revision` ≥ `Get-StdSeenRevision <id>`, inaczej `[STD-CFG-ROLLBACK]`;
5. cache i pendrive tylko gdy `allowOffline` = true, inaczej `[STD-CFG-OFFLINE]`;
6. `Test-StdRelease` bez błędów, inaczej `[STD-CFG-INVALID]` z pierwszymi błędami.

Reguły `Test-StdRelease` (komunikaty po polsku, każdy z nazwą pola):

- pola wymagane i typy z 2.6; `client.id`, klucze katalogu, `sites[].id` według wzorców; unikalne `sites[].id`;
- `sources.*`: `type` = `https` (wtedy `url` od `https://`) albo `unc` (wtedy `path`); `auth` = `session` / `none`;
- katalog: `name` niepusty; co najmniej jedno z `file`, `url`, `winget`; `url` tylko `https://`; `sha256` = 64 HEX;
  `timeoutMin` 1-240; `successCodes` liczby całkowite; `whenDetected` `skip` / `reinstall`; `sourceOrder` wskazuje istniejące `sources`;
  przy `requireSha256: true` każda pozycja z `file` / `url` ma `sha256`;
- `defaults.tasks` i `profiles.*.tasks`: tylko identyfikatory z 2.8, wartości bool;
- `systemSettings.registry[]`: `hive`, `path`, `name`, `type` z listy;
- `sites[].sourceOrder` wskazuje istniejące `sources`; `computerName.maxLength` ≤ 15; wstawki tylko `{serial}`, `{site}`, `{random}`;
- `profiles.*.apps` wskazują istniejące klucze katalogu; każda lokalizacja ma co najmniej jeden profil;
- **`domainJoin` i `localAdmin` nie mogą mieć żadnego pola, którego nazwa pasuje do `pass|pwd|hasl|secret|credential|user(name)?|login|account`**
  (wyjątek: `localAdmin.username`). Komunikat: „Dane logowania do domeny i hasło lokalnego administratora wpisuje się
  ręcznie przy wdrożeniu - nie mogą być w konfiguracji.”;
- w całym dokumencie poza `secrets` żaden klucz o nazwie `password`, `passwd`, `pwd`, `haslo`, `hasło`, `apitoken`, `token`, `secret`
  (bez względu na wielkość liter) nie może mieć niepustej wartości tekstowej;
- każde `{{secret:nazwa}}` (w `args`, `url`, `scripts[].args`, `wifi[].keyRef` jako nazwa) musi istnieć w `secrets.items`.

### 3.5 `Modules\Std.Legacy.ps1`

`ConvertFrom-StdLegacyConfig -Path <string> | -InputObject <obj> -ClientId <string> -ClientName <string>` →
`@{ Document (ordered hashtable, payload v4 bez revision/publishedAt); Warnings (string[]); SecretsToSet (string[]) }`.

Mapowanie konfiguracji 3.1 (`config.json`):

| 3.1 | v4 |
| --- | --- |
| `InstallSourcePaths.network` | `sources.lan = { type: unc, path, auth: session }` |
| `InstallSourcePaths.web` | `sources.web = { type: https, url, auth: session }`; adres `http://` → zostaje w `url` plus ostrzeżenie (walidacja go odrzuci, trzeba zmienić na `https://`) |
| `DefaultInstallSource` | `sites[0].sourceOrder`: `network` → `[lan, web]`, `web` / `winget` → `[web, lan]` (tylko istniejące) |
| `Programs.<n>` | `catalog.<n>` (klucz oczyszczony do wzorca, kolizje z sufiksem `-2`): `DownloadUrl` → `url`; `WingetId` → `winget` (+ `WingetArgs` → `wingetArgs`); `FileName` → `file` + `SilentArgs` → `args`, ale przy `DefaultInstallSource = winget` i braku `WingetId` → `winget = FileName`, `wingetArgs = WingetArgs` albo `SilentArgs`; `name` = nazwa programu |
| `Programs.*.Enabled = true` | profil `Domyślny` z tymi programami |
| `Profiles.<n>` | `sites[0].profiles.<n>.apps` (nieznane programy pominięte z ostrzeżeniem) |
| `TeamViewer` | `catalog.TeamViewer` (`file` albo `winget` jak wyżej); w `Arguments` wartości `APITOKEN=` i `CUSTOMCONFIGID=` zamienione na `{{secret:teamviewer-apitoken}}` / `{{secret:teamviewer-customconfigid}}` i dopisane do `SecretsToSet` |
| `AntyVirus` | `catalog.AV` (`name` = `Antywirus`); własne `InstallSourcePaths` AV → `sources.av-lan` / `av-web` i `catalog.AV.sourceOrder`; `Credentials` **pominięte** (ostrzeżenie) |
| `DefaultCheckboxes.InstallTeamViewer` / `InstallAV` = true | `TeamViewer` / `AV` dopisane do wszystkich profili |
| `WebAuth` | pominięte (ostrzeżenie: pobieranie idzie z kontem technika) |
| `DomainJoin.DomainName` | `sites[0].domainJoin.domain`; `Username` **pominięty** (ostrzeżenie: login wpisywany ręcznie) |
| `LocalAdmin.Username` | `sites[0].localAdmin.username` |
| `WiFiProfile.FileName` (tekst albo tablica) | `sites[0].wifi = [ { file } ]` + ostrzeżenie |
| `SystemSettings.<flaga>` | `defaults.systemSettings.<flaga>` |
| `SystemSettings.CustomRegistry[]` (`Path` `HKLM:\...`, `Name`, `Value`, `PropertyType`) | `defaults.systemSettings.registry[]` (`hive` z prefiksu, `path` bez prefiksu, `name`, `type`, `value`) |
| `PostInstallScripts[]` | `defaults.scripts = [ { file } ]` |
| `DefaultCheckboxes` | `defaults.tasks`: `WaitForNetwork`, `SuspendHibernation`→`SuspendSleep`, `ImportWiFiProfile`→`ConnectWiFi`, `RemoveBloatware`, `UninstallMicrosoft365`→`UninstallM365`, `UninstallOneDrive`, `InstallApplications`→`InstallApps`, `CreateLocalAdmin`, `ChangeComputerName`→`RenameComputer`, `JoinDomain`, `ChangeSystemSettings`→`SystemSettings`, `RunPostInstallScripts`→`RunScripts`, `ExportHardwareAudit`→`HardwareAudit`, `EnableBitLocker`, `AutoReboot`→`RebootAtEnd`; `DryRun`, `CreateRestorePoint`, `RunWindowsUpdate`, `JoinIntune`, `InstallTeamViewer`, `InstallAV` pominięte (ostrzeżenie przy true) |
| `HardwareAudit.ExportPath`, `AutoUpdate`, `CustomWebDataLocation`, `DarkTheme` | pominięte (ostrzeżenie, poza `DarkTheme`) |
| inne klucze | ostrzeżenie |

Jedna lokalizacja `default` („Domyślna”). Wynik musi przejść `Test-StdRelease -Unpublished` z wyjątkiem błędów o brakujących
sekretach (z `SecretsToSet`) i adresach `http://` (z ostrzeżeń). **Żadna wartość hasła ani tokenu z 3.1 nie trafia do wyniku.**

## 4. Narzędzia (`v4\Tools\`)

### 4.1 `New-StdSigningCertificate.ps1` (Windows)

Parametry: `[-Name 'STD - podpis konfiguracji']`, `[-Years 5]`, `[-Exportable]`, `[-OutPath .]`.
Tworzy certyfikat RSA 3072 / SHA-256 w `Cert:\CurrentUser\My` (`KeyUsage DigitalSignature`, EKU „Document Signing”
`1.3.6.1.4.1.311.10.3.12`, klucz domyślnie nieeksportowalny), zapisuje `.cer` i wypisuje gotowy wpis do `Std.TrustedSigners.psd1`.

### 4.2 `Publish-StdRelease.ps1`

```text
-SourcePath <plik v4 JSON bez podpisu>   (wymagany)
-Certificate <X509Certificate2> | -CertificateThumbprint <odcisk w Cert:\CurrentUser\My>
-Destination <katalog danych serwera (lokalny / UNC, z podkatalogiem std) | https://host/std-admin/>   (wymagany)
[-Credential <pscredential>]  (dla https)
[-ChangeNote <string>] [-ExpiresInDays 90] [-PublishedBy $env:USERNAME] [-WhatIf]
```

1. Wczytanie i `Test-StdRelease -Unpublished`; błędy → przerwanie z listą.
2. Poprzednia rewizja: istniejący `clients/<id>/release.sdt` w miejscu docelowym (sprawdzony tym samym certyfikatem) albo `revision` z pliku źródłowego; nowa = max + 1.
3. Ustawienie `revision`, `publishedAt` (UTC `o`), `publishedBy`, `expiresAt`, `changeNote`; payload = UTF-8 bez BOM `ConvertTo-StdJson`.
4. Podpis → `clients/<id>/release.sdt` i `clients/<id>/history/<rev>.sdt`.
5. Indeks: wczytanie `clients.sdt` (jeśli jest), dopisanie / podmiana klienta, `revision + 1`, podpis, zapis.
6. `https://`: `MKCOL` brakujących katalogów (405 = już jest), potem `PUT`; inaczej zapis plików (katalogi tworzone).
7. Wynik: `@{ ClientId; Revision; ReleasePath; IndexRevision; Sha256 }`.

### 4.3 `Convert-StdLegacyConfig.ps1`

`-Path <config.json> -ClientId <id> -ClientName <nazwa> [-OutPath <plik v4>]` - wywołuje `ConvertFrom-StdLegacyConfig`,
zapisuje dokument (bez podpisu) i wypisuje ostrzeżenia oraz listę sekretów do ustawienia.

## 5. Serwer

### 5.1 Wspólne

- Ten sam układ URL (2.4) i te same role: **technik** (odczyt `/std/`, zapis do `/std/drop/`) i **administrator** (wszystko,
  plus `/std-admin/`). Logowanie: Basic, wyłącznie przez HTTPS (HTTP przekierowany na HTTPS albo wyłączony).
- Typy MIME: `.sdt` → `application/json`; `.msi`, `.exe`, `.msix`, `.zip`, `.ps1`, `.xml`, `.cab` → `application/octet-stream`. Bez listowania katalogów.
- Duże pliki: do 4 GB w `/std-admin/` (instalatory), do 50 MB w `/std/drop/`.
- Nagłówki: `Strict-Transport-Security: max-age=31536000`, `X-Content-Type-Options: nosniff`, `Cache-Control: no-cache` dla `.sdt`.
- **Aktualizacja = ponowne uruchomienie instalki.** Instalka zapisuje ustawienia przy pierwszej instalacji i przy kolejnym
  uruchomieniu bierze je jako domyślne (parametry z wiersza poleceń je nadpisują i też są zapisywane). Nigdy nie usuwa danych
  (`std/`), kont ani plików z hasłami. Przed zmianą konfiguracji serwera WWW robi kopię zapasową i po nieudanym teście ją przywraca.
- Na końcu wypisuje adres, role i gotowy `bootstrap.json` dla pendrive'a.

### 5.2 Windows Server + IIS: `Server\IIS\Install-StdServer.ps1`

Parametry: `[-HostName]`, `[-CertificateThumbprint]` (certyfikat w `Cert:\LocalMachine\My`), `[-Port 443]`, `[-SiteName 'STD']`,
`[-DataPath 'C:\STD-Server\data']`, `[-TechnicianGroup 'STD-Technicy']`, `[-AdminGroup 'STD-Administratorzy']`, `[-SkipFeatureInstall]`.

- Ustawienia: `C:\ProgramData\STD-Server\settings.json` (`HostName`, `CertificateThumbprint`, `Port`, `SiteName`, `DataPath`, grupy, `InstalledVersion`, daty); kopie zapasowe w `C:\ProgramData\STD-Server\backup\<data>\` + `Backup-WebConfiguration`.
- Wymaga administratora. Funkcje: `Web-Server`, `Web-Static-Content`, `Web-Http-Redirect`, `Web-Basic-Auth`, `Web-Url-Auth`,
  `Web-DAV-Publishing`, `Web-Filtering`, `Web-Mgmt-Tools` (`Install-WindowsFeature`; na Windows 10/11 `Enable-WindowsOptionalFeature` z odpowiednikami `IIS-*`).
- Katalogi `<DataPath>\std\{clients,files,tool,drop}`; lokalne grupy (jeśli ich nie ma).
- NTFS (IIS przy Basic działa w imieniu zalogowanego użytkownika): technicy - odczyt `std\` poza `drop`, w `drop` tylko tworzenie plików (bez odczytu i listowania);
  administratorzy - modyfikacja wszędzie; `SYSTEM` i `Administrators` - pełne.
- Witryna `<SiteName>` (`https://<HostName>:<Port>`, SNI, certyfikat po odcisku), aplikacja / katalog wirtualny `std` i `std-admin` → `<DataPath>\std`.
  Uwierzytelnianie: anonimowe wyłączone, Basic włączone, `requireSSL`. Autoryzacja URL: `/std` - obie grupy, czasowniki `GET,HEAD`; `/std/drop` - obie grupy `PUT`;
  `/std-admin` - tylko administratorzy, wszystkie czasowniki. WebDAV: włączony dla `std-admin` (reguła dla administratorów: Read, Write, Source) i dla `std/drop` (technicy: Write).
- `requestFiltering`: `maxAllowedContentLength` 4 GB dla `std-admin`, 50 MB dla `std/drop`; mapy MIME z 5.1; `directoryBrowse` wyłączone.
- `Server\IIS\Add-StdServerUser.ps1 -UserName <n> -Role Technician|Admin [-Create] [-Password <SecureString>]` - dodaje (albo tworzy) lokalne konto do grupy.

### 5.3 Linux + nginx: `Server/Linux/install-std-server.sh`

```text
install-std-server.sh [--domain NAZWA] [--cert PLIK --key PLIK | --self-signed] [--port 443]
                      [--data-dir /srv/std] [--yes]
```

- Ustawienia: `/etc/std-server/settings.conf` (`KEY=wartość`, czytane parserem, nie przez `source`); kopie zapasowe w `/var/lib/std-server/backup/<data>/`.
- Wymaga roota. Debian / Ubuntu (`apt`): `nginx`, `apache2-utils`, `openssl`; RHEL / Rocky / Alma (`dnf`): `nginx`, `httpd-tools`, `openssl`
  (+ `semanage fcontext` / `restorecon` przy włączonym SELinux).
- Dane: `<data-dir>/std/{clients,files,tool,drop}`, właściciel użytkownik nginx (`www-data` / `nginx`), prawa 750 / 640.
- Hasła: `/etc/std-server/technicians.htpasswd`, `admins.htpasswd` (bcrypt, `root:<grupa nginx> 640`) i generowany `readers.htpasswd` (suma obu).
- `/etc/nginx/conf.d/std-server.conf` z szablonu w skrypcie: `server` 443 `ssl` (`TLSv1.2 TLSv1.3`), `server` 80 → 301 na HTTPS,
  `location /std/` (`alias`, `auth_basic` → `readers`, `limit_except GET HEAD { deny all; }`), `location /std/drop/` (`readers`, `dav_methods PUT`,
  `create_full_put_path on`, `limit_except PUT { deny all; }`, 50 MB), `location /std-admin/` (`admins`, `dav_methods PUT DELETE MKCOL`,
  `create_full_put_path on`, 4 GB), `autoindex off`, typy MIME z 5.1, nagłówki z 5.1. Przed `reload`: `nginx -t`; błąd → przywrócenie kopii.
- `/usr/local/sbin/std-server` (kopiowany przez instalkę): `add-tech <login>`, `add-admin <login>`, `remove <login>`, `list`, `status`, `rebuild`;
  hasło pytane albo z `--password-stdin`; po zmianie odbudowa `readers.htpasswd` i `reload`.
- `--self-signed` tworzy certyfikat w `/etc/std-server/tls/` (do testów).

### 5.4 Testy instalek w CI

- `.github/workflows/v4-server-linux.yml` (`ubuntu-latest`): instalacja `--self-signed --domain localhost --yes`, konta testowe przez
  `std-server ... --password-stdin`, `curl -k`: GET bez hasła → 401, technik GET `clients.sdt` → 200, technik PUT do `drop` → 201/204,
  technik GET `drop` → 403, technik PUT do `std-admin` → 401/403, administrator PUT do `std-admin` → 201/204, HTTP → 301;
  potem ponowne uruchomienie instalki (aktualizacja) i sprawdzenie, że konta, ustawienia i pliki zostały.
- `.github/workflows/v4-server-iis.yml` (`windows-latest`): certyfikat testowy w `Cert:\LocalMachine\My`, instalacja, lokalne konta testowe
  przez `Add-StdServerUser.ps1 -Create`, te same sprawdzenia (`curl.exe -k`), ponowne uruchomienie i sprawdzenie zachowania.
- Oba workflow: `on: push` (ścieżki `v4/Server/**` i plik workflow) oraz `workflow_dispatch`.
