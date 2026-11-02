# Serwer STD na Windows Server + IIS

Ta instalka przygotowuje serwer, z którego STD (Smart Tool for Deployment) pobiera podpisane konfiguracje
i instalatory, i na który wysyła raporty. Serwer to zwykłe pliki przez HTTPS, bez bazy danych i bez API.
Szczegóły techniczne: [SPEC.md](../../SPEC.md), sekcje 2.4 i 5.

Pliki w tym folderze:

| Plik | Do czego |
| --- | --- |
| `Install-StdServer.ps1` | instalacja i aktualizacja serwera |
| `Add-StdServerUser.ps1` | dodanie konta technika albo administratora |

## Wymagania

- Windows Server 2016 lub nowszy (działa też na Windows 10 / 11 - do testów), Windows PowerShell 5.1.
- Konto administratora serwera.
- Nazwa DNS serwera, np. `pobieranie.firma.pl`, wskazująca na ten serwer.
- Certyfikat TLS dla tej nazwy w magazynie komputera (`Cert:\LocalMachine\My`), **z kluczem prywatnym**.
  Komputery, na których działa STD, muszą temu certyfikatowi ufać (certyfikat z firmowego CA albo publiczny).
- Wolne porty 443 (HTTPS) i 80 (przekierowanie na HTTPS) dla tej nazwy.

## Instalacja

1. Skopiuj na serwer cały folder `v4` (instalka czyta numer wersji z `v4\STD.psd1`).
2. Zaimportuj certyfikat, jeśli go jeszcze nie ma:

   ```powershell
   Import-PfxCertificate -FilePath C:\Temp\pobieranie.pfx -CertStoreLocation Cert:\LocalMachine\My -Password (Read-Host -AsSecureString 'Hasło pliku PFX')
   ```

   Odcisk certyfikatu (Thumbprint) zobaczysz poleceniem:

   ```powershell
   Get-ChildItem Cert:\LocalMachine\My | Format-Table Thumbprint, Subject, NotAfter
   ```

3. Uruchom instalkę w PowerShellu:

   ```powershell
   cd C:\STD\v4\Server\IIS
   powershell -ExecutionPolicy Bypass -File .\Install-StdServer.ps1 -HostName pobieranie.firma.pl -CertificateThumbprint <odcisk>
   ```

   Jeśli PowerShell nie działa jako administrator, instalka sama uruchomi się ponownie jako administrator
   (pojawi się okno UAC). Okno instalki czeka wtedy na koniec pracy i na Enter, żeby dało się przeczytać wynik.

4. Dodaj konta (rozdział „Konta”).
5. Skopiuj wypisany na końcu `bootstrap.json` na pendrive (rozdział „Plik bootstrap.json dla pendrive'a”).

Instalka nie pyta o żadne hasła i żadnych haseł nie zapisuje.

### Parametry

| Parametr | Domyślnie | Opis |
| --- | --- | --- |
| `-HostName` | pełna nazwa komputera | nazwa w adresie `https://<HostName>/std/` |
| `-CertificateThumbprint` | najnowszy ważny certyfikat dla `HostName` | odcisk certyfikatu w `Cert:\LocalMachine\My` |
| `-Port` | `443` | port HTTPS (port 80 jest zajęty przez przekierowanie) |
| `-SiteName` | `STD` | nazwa witryny i puli aplikacji w IIS |
| `-DataPath` | `C:\STD-Server\data` | katalog danych; pliki STD są w `<DataPath>\std` |
| `-TechnicianGroup` | `STD-Technicy` | lokalna grupa techników (albo istniejąca grupa domenowa `FIRMA\grupa`) |
| `-AdminGroup` | `STD-Administratorzy` | lokalna grupa administratorów STD (albo grupa domenowa) |
| `-SkipFeatureInstall` | | nie instaluje funkcji Windows (IIS jest już zainstalowany) |

### Co instalka robi

1. Instaluje brakujące funkcje IIS: `Web-Server`, `Web-Static-Content`, `Web-Http-Redirect`, `Web-Basic-Auth`,
   `Web-Url-Auth`, `Web-DAV-Publishing`, `Web-Filtering`, `Web-Mgmt-Tools` (oraz konsolę IIS i skrypty zarządzania).
   Na Windows 10 / 11 włącza odpowiedniki `IIS-*`.
2. Robi kopię zapasową (rozdział „Kopie zapasowe i logi”).
3. Tworzy katalogi `<DataPath>\std\clients`, `files`, `files\common`, `tool`, `drop` i grupy (jeśli ich nie ma).
4. Ustawia uprawnienia NTFS (tabela niżej).
5. Konfiguruje IIS: pula `STD`, witryna `STD` (`https://<HostName>:<Port>`, SNI), katalogi wirtualne `/std` i `/std-admin`
   wskazujące na `<DataPath>\std`, witryna `STD-HTTP` na porcie 80 (przekierowanie 301 na HTTPS).
6. Sprawdza, czy serwer odpowiada przez HTTPS właściwym certyfikatem i czy `/std/` wymaga logowania.
   Jeśli nie - przywraca konfigurację IIS z kopii i kończy się błędem.
7. Zapisuje ustawienia w `C:\ProgramData\STD-Server\settings.json` i wypisuje adres, role oraz `bootstrap.json`.

### Kto ma dostęp do czego

| Adres | Technik | Administrator |
| --- | --- | --- |
| `/std/` (konfiguracje, instalatory, wydania STD) | odczyt (GET, HEAD) | odczyt (GET, HEAD) |
| `/std/drop/` (raporty z wdrożeń) | tylko zapis nowych plików (PUT), bez odczytu i listy plików | zapis (PUT) |
| `/std-admin/` (to samo drzewo, publikacja) | brak | wszystko: GET, PUT, DELETE, MKCOL |

- Logowanie: Basic, tylko przez HTTPS. Bez logowania serwer odpowiada 401. Listowanie katalogów jest wyłączone.
- Limity wysyłania: 4 GB w `/std-admin/` (instalatory), 50 MB w `/std/drop/`.
- Nagłówki: `Strict-Transport-Security: max-age=31536000`, `X-Content-Type-Options: nosniff`, `Cache-Control: no-cache`.
- Typy plików: `.sdt` jako `application/json`; `.msi`, `.exe`, `.msix`, `.zip`, `.ps1`, `.xml`, `.cab` i wszystkie inne
  rozszerzenia jako `application/octet-stream` (bez tego IIS nie wydałby np. plików `.msu` czy `.7z`).

Uprawnienia NTFS (IIS przy logowaniu Basic czyta i zapisuje pliki w imieniu zalogowanego użytkownika, więc NTFS
jest drugą, niezależną warstwą ochrony):

| Katalog | Kto | Prawa |
| --- | --- | --- |
| `<DataPath>\std` i podkatalogi | SYSTEM, Administratorzy | pełna kontrola |
| | grupa administratorów STD | modyfikacja |
| | grupa techników | odczyt i wykonanie |
| | IIS_IUSRS (proces IIS) | odczyt |
| `<DataPath>\std\drop` (bez dziedziczenia z `std`) | SYSTEM, Administratorzy | pełna kontrola |
| | grupa administratorów STD | modyfikacja |
| | grupa techników | tylko „Utwórz pliki”, przechodzenie i odczyt atrybutów katalogu; bez listy plików i bez dostępu do cudzych plików |
| | TWÓRCA-WŁAŚCICIEL | zapis (bez odczytu) tylko we własnym, właśnie utworzonym pliku |
| | IIS_IUSRS | odczyt katalogów (nie plików) |
| `C:\inetpub\<SiteName>-wwwroot` (pusty katalog witryny) | SYSTEM, Administratorzy, IIS_IUSRS (odczyt) | |
| `C:\ProgramData\STD-Server` (ustawienia, kopie, logi) | SYSTEM, Administratorzy | |

Instalka nie odblokowuje globalnie sekcji logowania IIS: wszystko jest zapisane w `applicationHost.config`,
więc plik `web.config` podłożony w danych nie może włączyć logowania anonimowego.

## Aktualizacja

**Aktualizacja = ponowne uruchomienie instalki** (najlepiej z nowszego folderu `v4`):

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-StdServer.ps1
```

- Bez parametrów instalka bierze ustawienia z `C:\ProgramData\STD-Server\settings.json` (nazwa, certyfikat, port,
  witryna, katalog danych, grupy). Na początku wypisuje każde ustawienie i skąd pochodzi (`parametr`, `zapisane`, `domyślne`).
- Parametr podany w wierszu poleceń zmienia ustawienie i zostaje zapisany na następne razy.
- Instalka **nigdy nie usuwa** danych (`<DataPath>\std`), kont, grup ani członków grup. Opublikowane konfiguracje,
  instalatory i raporty zostają.
- Przed zmianą robi kopię zapasową. Gdy test po zmianie się nie uda, przywraca poprzednią konfigurację IIS.
- Wszystkie kroki można powtarzać: drugie uruchomienie z tymi samymi ustawieniami niczego nie psuje.

Zmiana ustawienia, przykłady:

```powershell
.\Install-StdServer.ps1 -CertificateThumbprint <nowy odcisk>      # nowy certyfikat
.\Install-StdServer.ps1 -HostName std.firma.pl                     # nowa nazwa (pamiętaj o certyfikacie i bootstrap.json)
.\Install-StdServer.ps1 -DataPath D:\STD-Server\data               # nowy katalog danych
```

Przy zmianie katalogu danych stare pliki zostają na miejscu - przenieś je ręcznie
(np. `robocopy C:\STD-Server\data D:\STD-Server\data /E /COPY:DAT`) i uruchom instalkę jeszcze raz, żeby ustawiła uprawnienia.
Zmiana nazwy witryny (`-SiteName`) nie jest obsługiwana: usuń najpierw starą witrynę w Menedżerze IIS.

## Konta

Technicy i administratorzy logują się w STD kontem Windows z tego serwera (albo z domeny).
Rolę wyznacza członkostwo w grupie.

```powershell
# nowe konto technika (skrypt zapyta o hasło dwa razy)
.\Add-StdServerUser.ps1 -UserName jan.kowalski -Role Technician -Create

# nowe konto administratora STD
.\Add-StdServerUser.ps1 -UserName anna.nowak -Role Admin -Create

# istniejące konto (lokalne albo domenowe) - tylko dodanie do grupy
.\Add-StdServerUser.ps1 -UserName 'FIRMA\piotr.wisniewski' -Role Technician
```

- Skrypt sam uruchamia się jako administrator, jeśli trzeba. Hasła nie przekazuje do nowego okna - zapyta o nie tam.
- Hasło jest pytane bez wyświetlania i nigdzie nie jest zapisywane. Musi spełniać zasady haseł Windows
  (zwykle co najmniej 8 znaków, trzy rodzaje znaków).
- Istniejącego konta skrypt nie zmienia (także hasła).
- Nowe konta lokalne mają hasło bez wygasania: logowanie Basic nie ma ekranu zmiany hasła, a wygasłe hasło
  oznacza nagłą odmowę dostępu w terenie. Nowe konta nie są w grupie „Użytkownicy”, więc nie zalogują się na serwer
  pulpitem.
- Login w STD: konto lokalne - sam login (`jan.kowalski`); konto domenowe - `FIRMA\login`.

Inne czynności (wbudowane polecenia Windows):

```powershell
Set-LocalUser -Name jan.kowalski -Password (Read-Host -AsSecureString 'Nowe hasło')   # zmiana hasła
Remove-LocalGroupMember -Group STD-Technicy -Member jan.kowalski                     # odebranie roli
Disable-LocalUser -Name jan.kowalski                                                 # zablokowanie konta
Get-LocalGroupMember -Group STD-Technicy                                             # lista techników
```

## Certyfikat

- Certyfikat musi być w `Cert:\LocalMachine\My`, mieć klucz prywatny i obejmować nazwę `HostName`
  (w polu SAN; certyfikat wieloznaczny `*.firma.pl` też działa).
- Instalka ostrzega, gdy certyfikat wygasa w ciągu 30 dni, i kończy się błędem, gdy już wygasł.
- Wymiana certyfikatu: zaimportuj nowy, potem `.\Install-StdServer.ps1 -CertificateThumbprint <nowy odcisk>`.
- Certyfikat z `New-SelfSignedCertificate` nadaje się tylko do testów: STD na komputerach odrzuci połączenie
  z certyfikatem, któremu komputer nie ufa.

## Plik bootstrap.json dla pendrive'a

Na końcu instalka wypisuje gotowy plik i zapisuje jego kopię w `C:\ProgramData\STD-Server\bootstrap.json`, np.:

```json
{"server":"https://pobieranie.firma.pl/std/","mirrors":[],"auth":{"mode":"basic"},"proxy":null}
```

Skopiuj go na pendrive obok `Start-STD.cmd`. Plik nie jest tajny (nie ma w nim haseł). W `mirrors` możesz dopisać
ścieżki do kopii tego samego drzewa, np. `"\\\\srv-files\\STD$\\std\\"` (SPEC 2.3).

## Publikacja konfiguracji

Administrator publikuje podpisaną konfigurację narzędziem `v4\Tools\Publish-StdRelease.ps1`:

- przez HTTPS: `-Destination https://pobieranie.firma.pl/std-admin/` i konto administratora STD (`-Credential`);
- na samym serwerze: `-Destination C:\STD-Server\data\std`.

Instalatory klientów kopiuje się do `<DataPath>\std\files\<klient>\` (wspólne do `files\common\`)
albo wysyła przez `/std-admin/files/...`.

## Kopie zapasowe i logi

| Co | Gdzie |
| --- | --- |
| Ustawienia | `C:\ProgramData\STD-Server\settings.json` |
| Kopie przed każdym uruchomieniem | `C:\ProgramData\STD-Server\backup\<data-godzina>\` (`settings.json`, pliki konfiguracji IIS, uprawnienia danych `acl-std.txt`, wiązania certyfikatów `http-sslcert.txt`) |
| Kopie konfiguracji IIS | `Backup-WebConfiguration` o nazwie `STD-<data-godzina>` (lista: `Get-WebConfigurationBackup`) |
| Log instalki | `C:\ProgramData\STD-Server\logs\install-<data-godzina>.log` |
| Logi IIS | `C:\inetpub\logs\LogFiles\W3SVC<id>\` |

Ręczne przywrócenie konfiguracji IIS: `Restore-WebConfiguration -Name STD-<data-godzina>`.
Ręczne przywrócenie uprawnień danych: `icacls <DataPath> /restore <kopia>\acl-std.txt` (np. `icacls C:\STD-Server\data /restore ...`).
Stare kopie można usuwać ręcznie; instalka ich nie kasuje.

## Problemy

| Objaw | Co sprawdzić |
| --- | --- |
| Instalka: `[STD-SRV-CERT]` | Certyfikat nie jest w `Cert:\LocalMachine\My`, nie ma klucza prywatnego albo wygasł. Podaj `-CertificateThumbprint`. |
| Instalka: `[STD-SRV-TEST]` | Serwer nie odpowiedział po zmianie (konfiguracja została przywrócona). Sprawdź log instalki, „Podgląd zdarzeń” → System (źródło WAS, HttpEvent) i czy port 443 nie jest zajęty: `netsh http show sslcert`, `netstat -ano \| findstr :443`. |
| Instalka: `[STD-SRV-SITE]` | Witryna o tej nazwie albo wiązanie `https://<HostName>:<Port>` już istnieje w innej witrynie. Zmień `-SiteName`, `-HostName` albo `-Port`. |
| Instalka: `[STD-SRV-FEATURE]` | Nie udało się zainstalować funkcji Windows. Zainstaluj je w Menedżerze serwera i uruchom z `-SkipFeatureInstall`. Jeśli system prosi o restart - uruchom ponownie serwer i instalkę. |
| Instalka: `[STD-SRV-SETTINGS]` | `settings.json` jest uszkodzony. Przywróć go z `backup\<data>\settings.json`. |
| „Port 80 ... zajmuje witryna” | Inna witryna ma `http://<HostName>`. Przekierowanie HTTP zostało pominięte, HTTPS działa. |
| STD: poprawne hasło, a serwer odpowiada 401 | Konto jest w grupie? (`Get-LocalGroupMember STD-Technicy`). Konto włączone? Konto domenowe wpisane jako `FIRMA\login`? |
| 401 przy zapisie raportu (`/std/drop/`) | Plik o tej nazwie już istnieje i jest cudzy (technicy nie nadpisują cudzych raportów) albo konto nie jest w grupie. |
| 404.3 | Brak typu MIME - nie powinno wystąpić; uruchom ponownie instalkę. |
| 404.13 | Plik większy niż limit (4 GB `/std-admin/`, 50 MB `/std/drop/`). |
| 404.11 | Adres z podwójnie kodowanymi znakami - nie używaj `%25` w nazwach plików. |
| STD: błąd certyfikatu | Komputer nie ufa certyfikatowi albo nazwa w `bootstrap.json` nie zgadza się z certyfikatem. |

W logach IIS (`C:\inetpub\logs\LogFiles`) kolumny `sc-status` i `sc-substatus` mówią dokładnie, co się stało:
401.1 - złe hasło, 401.2 - brak uprawnień według reguł autoryzacji (np. technik w `/std-admin/`), 401.3 - brak
uprawnień NTFS. Logi IIS nie zawierają haseł.

Kody błędów instalki i skryptu kont:

| Kod | Znaczenie |
| --- | --- |
| `STD-SRV-OS` | uruchomiono poza Windows |
| `STD-SRV-ADMIN` | brak uprawnień administratora (odmowa w oknie UAC) |
| `STD-SRV-PARAM` | nieprawidłowy parametr |
| `STD-SRV-SETTINGS` | uszkodzony `settings.json` |
| `STD-SRV-FEATURE` | instalacja funkcji Windows |
| `STD-SRV-IIS` | brak bibliotek IIS |
| `STD-SRV-CERT` | certyfikat |
| `STD-SRV-SITE` | konflikt witryny albo wiązania |
| `STD-SRV-GROUP` | grupa nie istnieje albo nie da się dodać członka |
| `STD-SRV-ACL` | uprawnienia NTFS (`icacls`) |
| `STD-SRV-TLS` | powiązanie certyfikatu w http.sys |
| `STD-SRV-TEST` | test po zmianie konfiguracji (konfiguracja przywrócona) |
| `STD-SRV-USER` | konto: nie istnieje, zły login, nie da się utworzyć |
| `STD-SRV-PASSWORD` | brak hasła dla nowego konta albo hasła się różniły |

Serwer na kontrolerze domeny nie jest obsługiwany (tam nie ma grup lokalnych) - użyj osobnego serwera członkowskiego.

## Test w CI

Workflow `.github/workflows/v4-server-iis.yml` (Windows Server, Windows PowerShell 5.1) instaluje serwer z certyfikatem
testowym dla `localhost`, tworzy konta testowe, sprawdza dostęp `curl.exe -k` (logowanie, odczyt, zapis do `drop`,
brak odczytu `drop`, `/std-admin/`, przekierowanie HTTP), potem uruchamia instalkę ponownie bez parametrów
i sprawdza, że ustawienia, konta, grupy i pliki zostały.
