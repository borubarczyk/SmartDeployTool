# Serwer STD na Linuksie (nginx)

Serwer plików dla STD 4 (Smart Tool for Deployment). Udostępnia przez HTTPS podpisane konfiguracje
klientów, instalatory i wydania narzędzia, a od techników przyjmuje raporty. Nie ma bazy danych ani API:
to zwykłe pliki, nginx i konta z hasłami.

Pliki w tym katalogu:

| Plik | Do czego |
| --- | --- |
| `install-std-server.sh` | instalka; ponowne uruchomienie = aktualizacja |
| `std-server` | narzędzie do kont i stanu serwera (instalka kopiuje je do `/usr/local/sbin/std-server`) |

Umowa techniczna (adresy, role, nagłówki): [SPEC.md](../../SPEC.md), sekcja 5.

## Role

| Rola | Co może |
| --- | --- |
| technik | czyta `https://<serwer>/std/` (konfiguracje, instalatory, wydania STD) i zapisuje raporty do `/std/drop/` (bez odczytu i listowania) |
| administrator | to samo co technik oraz `https://<serwer>/std-admin/`: odczyt, zapis, usuwanie i tworzenie katalogów (publikacja) |

Logowanie: login i hasło (Basic), tylko przez HTTPS. Zapytania na port 80 są przekierowywane na HTTPS.

Na serwerze **nie ma** loginu ani hasła do domeny ani hasła lokalnego administratora komputera.
Te dane technik zawsze wpisuje ręcznie przy wdrożeniu i STD nigdzie ich nie zapisuje.

## Co jest potrzebne

- Debian 12, Ubuntu 22.04 / 24.04 albo RHEL / Rocky / Alma 8 i 9, konto root (`sudo`).
- Nazwa serwera w DNS, np. `pobieranie.firma.pl`, i otwarte porty 80 oraz 443.
- Certyfikat TLS dla tej nazwy. Komputer przygotowywany przez STD zwykle nie jest jeszcze w domenie,
  więc nie ufa firmowemu urzędowi certyfikacji. Najprościej użyć certyfikatu publicznego
  (Let's Encrypt albo kupionego).
- Instalka sama doinstaluje brakujące pakiety: `nginx`, `apache2-utils` (`htpasswd`) i `openssl`
  na Debianie / Ubuntu albo `nginx`, `httpd-tools` i `openssl` na RHEL / Rocky / Alma.

## Instalacja

Skopiuj cały katalog `v4/Server/Linux` na serwer (instalka potrzebuje pliku `std-server` obok siebie) i uruchom:

```bash
cd Linux
sudo bash install-std-server.sh --domain pobieranie.firma.pl \
    --cert /etc/ssl/std/fullchain.pem --key /etc/ssl/std/privkey.pem
```

Bez `--yes` instalka pokazuje podsumowanie i pyta „Kontynuować? [t/N]”. Przy pierwszej instalacji
proponuje też utworzenie pierwszego administratora. Na końcu wypisuje adres serwera, role
i gotowy plik `bootstrap.json` na pendrive.

### Parametry

| Parametr | Domyślnie | Opis |
| --- | --- | --- |
| `--domain NAZWA` | nazwa komputera | nazwa serwera w adresie (litery, cyfry, `-`, `.` albo adres IPv4) |
| `--cert PLIK --key PLIK` | - | certyfikat (PEM, najlepiej pełny łańcuch) i klucz prywatny (PEM, bez hasła) |
| `--self-signed` | - | certyfikat samopodpisany w `/etc/std-server/tls/`, **tylko do testów** |
| `--port PORT` | `443` | port HTTPS; port 80 zawsze tylko przekierowuje na HTTPS |
| `--data-dir KATALOG` | `/srv/std` | katalog danych; pliki serwera są w `KATALOG/std/` |
| `--yes` | - | bez pytań (skrypty, CI) |

Przy pierwszej instalacji trzeba podać certyfikat (`--cert` i `--key`) albo `--self-signed`.
Przy kolejnych uruchomieniach wszystkie parametry są opcjonalne.

### Certyfikat Let's Encrypt

Serwer musi być widoczny z internetu na porcie 80 pod podaną nazwą.

1. Instalacja z certyfikatem tymczasowym. Port 80 od razu obsługuje pliki weryfikacji Let's Encrypt
   (`/.well-known/acme-challenge/` z katalogu `/var/lib/std-server/acme`):

   ```bash
   sudo bash install-std-server.sh --domain pobieranie.firma.pl --self-signed --yes
   ```

2. Pobranie certyfikatu:

   ```bash
   sudo apt install certbot        # albo: sudo dnf install certbot
   sudo certbot certonly --webroot -w /var/lib/std-server/acme -d pobieranie.firma.pl
   ```

3. Przełączenie serwera na ten certyfikat (ścieżki z `live/`, bez rozwijania dowiązań):

   ```bash
   sudo bash install-std-server.sh \
       --cert /etc/letsencrypt/live/pobieranie.firma.pl/fullchain.pem \
       --key /etc/letsencrypt/live/pobieranie.firma.pl/privkey.pem --yes
   ```

4. Odnawianie: certbot odnawia certyfikat sam. Żeby nginx wczytał nowy plik, dodaj skrypt:

   ```bash
   printf '#!/bin/sh\nnginx -t && systemctl reload nginx\n' | sudo tee /etc/letsencrypt/renewal-hooks/deploy/std-server-reload
   sudo chmod 755 /etc/letsencrypt/renewal-hooks/deploy/std-server-reload
   sudo certbot renew --dry-run
   ```

### Własny certyfikat

- Plik certyfikatu w formacie PEM: najpierw certyfikat serwera, potem certyfikaty pośrednie.
- Klucz prywatny w formacie PEM, bez hasła, z prawami `600` (czyta go tylko root).
- Instalka sprawdza, czy klucz pasuje do certyfikatu, i ostrzega, gdy certyfikat wygasł, wygasa w ciągu 14 dni
  albo nie obejmuje nazwy z `--domain`.

Certyfikat z pliku `.pfx` (Windows) zamienisz tak (przy starszych plikach w OpenSSL 3 może być potrzebne `-legacy`):

```bash
openssl pkcs12 -in serwer.pfx -nokeys -out fullchain.pem
openssl pkcs12 -in serwer.pfx -nocerts -nodes -out privkey.pem
chmod 600 privkey.pem
```

Sprawdź kolejność w `fullchain.pem`: pierwszy blok to certyfikat serwera.

### Certyfikat samopodpisany (testy)

`--self-signed` tworzy certyfikat RSA 3072 na 825 dni w `/etc/std-server/tls/`. Komputery z STD go odrzucą,
bo nikt go nie podpisał. Nadaje się tylko do testów (np. `curl -k`). Przy aktualizacji certyfikat zostaje,
chyba że zmieniła się nazwa serwera albo zostało mniej niż 30 dni ważności - wtedy powstaje nowy.

## Co instalka tworzy

| Ścieżka | Zawartość |
| --- | --- |
| `/etc/std-server/settings.conf` | ustawienia (`KLUCZ=wartość`); instalka czyta je parserem, nic z nich nie jest wykonywane |
| `/etc/std-server/technicians.htpasswd` | konta techników (bcrypt), `root:<grupa nginx> 640` |
| `/etc/std-server/admins.htpasswd` | konta administratorów (bcrypt), `root:<grupa nginx> 640` |
| `/etc/std-server/readers.htpasswd` | suma obu plików, tworzona automatycznie - nie edytuj |
| `/etc/std-server/bootstrap.json` | gotowy plik na pendrive |
| `/etc/std-server/tls/` | certyfikat samopodpisany (tylko przy `--self-signed`) |
| `/etc/nginx/conf.d/std-server.conf` | konfiguracja nginx z szablonu w instalce |
| `/usr/local/sbin/std-server` | narzędzie do kont |
| `/srv/std/std/` | dane: `clients/`, `files/`, `tool/`, `drop/`; właściciel użytkownik nginx, katalogi 750, pliki 640 |
| `/srv/std/upload-tmp/` | pliki tymczasowe nginx przy wgrywaniu (na tym samym dysku co dane) |
| `/var/lib/std-server/backup/<data>/` | kopie zapasowe poprzednich wersji plików |
| `/var/lib/std-server/acme/` | pliki weryfikacji Let's Encrypt |

Drzewo danych jest takie samo jak na serwerze Windows (IIS) i w kopii lustrzanej:

```text
/std/                              odczyt: technicy i administratorzy (GET, HEAD)
  clients.sdt                      podpisany indeks klientów
  clients/<id>/release.sdt         podpisana konfiguracja klienta
  clients/<id>/history/<rev>.sdt   archiwum opublikowanych rewizji
  files/<id>/...                   instalatory klienta
  files/common/...                 instalatory wspólne
  tool/STD-<wersja>.zip            wydania narzędzia
  drop/                            tylko zapis (PUT) dla techników, do 50 MB na plik
/std-admin/                        to samo drzewo, tylko administratorzy: GET, PUT, DELETE, MKCOL, do 4 GB na plik
```

Instalka **nie usuwa, nie wyłącza i nie zmienia** plików nginx z dystrybucji (np. domyślnej strony
`/etc/nginx/sites-enabled/default`). Dodaje tylko własny plik `conf.d/std-server.conf`. Serwer STD odpowiada
na nazwę z `--domain`. Zapytania z inną nazwą obsługuje strona oznaczona jako domyślna (`default_server`);
gdy takiej nie ma, nginx bierze pierwszą wczytaną, czyli zwykle serwer STD (pliki z `conf.d/` są wczytywane
przed `sites-enabled/`).

Usunięcie (DELETE) całego drzewa `/std-admin/` i katalogów `clients/`, `files/`, `tool/`, `drop/` jest
zablokowane (403), żeby błąd w skrypcie publikacji nie skasował wszystkiego. Pojedyncze pliki i podkatalogi
administrator usuwa normalnie.

## Aktualizacja

Aktualizacja to ponowne uruchomienie instalki, także w nowszej wersji:

```bash
sudo bash install-std-server.sh --yes
```

- Bez parametrów instalka bierze zapisane ustawienia z `/etc/std-server/settings.conf`.
- Parametr podany w wierszu poleceń zmienia ustawienie i jest zapisywany, np. `--port 8443`.
- Konta, pliki haseł, opublikowane konfiguracje i instalatory zostają. Instalka niczego w nich nie usuwa
  i niczego nie nadpisuje; poprawia tylko właściciela i prawa (750 / 640).
- Przed każdą zmianą instalka kopiuje poprzednie `std-server.conf`, `settings.conf` i `std-server`
  do `/var/lib/std-server/backup/<data>/`.
- Nowa konfiguracja nginx jest sprawdzana (`nginx -t`) i na koniec sprawdzane jest, czy serwer odpowiada.
  Gdy coś się nie uda, instalka sama przywraca poprzednią konfigurację, nie zapisuje nowych ustawień
  i kończy się błędem z opisem przyczyny.
- Zmiana `--data-dir` nie przenosi danych ze starego katalogu (i go nie usuwa). Przenieś pliki ręcznie.

Ręczne zmiany w `/etc/nginx/conf.d/std-server.conf` zostaną nadpisane przy kolejnej aktualizacji
(poprzednia wersja trafi do kopii zapasowej).

## Konta techników i administratorów

```bash
sudo std-server add-tech jan.kowalski      # nowy technik albo zmiana hasła technika
sudo std-server add-admin anna.nowak       # nowy administrator albo zmiana hasła administratora
sudo std-server remove jan.kowalski        # usunięcie konta
sudo std-server list                       # lista kont i ról
sudo std-server status                     # stan serwera, adres, role, bootstrap.json
sudo std-server rebuild                    # odbudowa readers.htpasswd, prawa danych, przeładowanie nginx
```

- Hasło jest pytane dwa razy i nie jest widoczne przy wpisywaniu. Co najmniej 8 znaków, najwyżej 72 bajty.
- W skryptach hasło podaje się na standardowe wejście, nigdy w wierszu poleceń:

  ```bash
  printf '%s\n' "$HASLO" | sudo std-server add-tech jan.kowalski --password-stdin
  ```

- Login: litery, cyfry oraz `.` `_` `@` `-` (do 64 znaków).
- `add-admin` dla istniejącego technika (albo odwrotnie) zmienia rolę konta i ustawia nowe hasło.
- Hasła są zapisywane jako skróty bcrypt. Jawnego hasła nie ma w żadnym pliku ani logu.
- Zmiana konta działa od razu (nginx czyta plik kont przy każdym zapytaniu); `std-server` i tak przeładowuje nginx.

## Wgrywanie plików (publikacja)

- Konfiguracje: `v4\Tools\Publish-StdRelease.ps1` z `-Destination https://pobieranie.firma.pl/std-admin/`
  i kontem administratora (`-Credential`). Narzędzie samo tworzy katalogi (MKCOL) i wgrywa pliki (PUT).
- Instalatory przez HTTPS (curl zapyta o hasło):

  ```bash
  curl -u anna.nowak -T 7z2408-x64.exe https://pobieranie.firma.pl/std-admin/files/lizard/7z2408-x64.exe
  ```

  PUT sam tworzy brakujące katalogi. MKCOL tworzy jeden katalog i wymaga `/` na końcu adresu.
- Instalatory bezpośrednio na dysk (scp, rsync) do `/srv/std/std/files/...`, a potem
  `sudo std-server rebuild`, żeby nginx miał prawa do nowych plików.

## Plik bootstrap.json na pendrive

Instalka i `std-server status` wypisują gotowy plik; jego kopia jest w `/etc/std-server/bootstrap.json`.
Zapisz go na pendrivie obok `Start-STD.cmd`:

```json
{
  "server": "https://pobieranie.firma.pl/std/",
  "mirrors": [],
  "auth": { "mode": "basic" },
  "proxy": null
}
```

`mirrors` to opcjonalne kopie tego samego drzewa co `/std/`, np. `"\\\\srv-files\\STD$\\std\\"`.
Plik nie jest tajny: nie ma w nim haseł, a konfiguracje i tak są sprawdzane po podpisie.

## Sprawdzenie, czy serwer działa

```bash
sudo std-server status
curl -s -o /dev/null -w '%{http_code}\n' https://pobieranie.firma.pl/std/clients.sdt              # 401
curl -s -o /dev/null -w '%{http_code}\n' -u jan.kowalski https://pobieranie.firma.pl/std/clients.sdt   # 200 (po publikacji)
curl -s -o /dev/null -w '%{http_code}\n' -u jan.kowalski https://pobieranie.firma.pl/std/drop/      # 403
curl -s -o /dev/null -w '%{http_code}\n' http://pobieranie.firma.pl/std/                            # 301
```

Przy certyfikacie samopodpisanym dodaj `-k`. Pełny zestaw sprawdzeń jest w workflow
`.github/workflows/v4-server-linux.yml`.

## Rozwiązywanie problemów

Najpierw `sudo std-server status`, potem logi: `/var/log/nginx/std-server.error.log`
i `/var/log/nginx/std-server.access.log`.

| Objaw | Przyczyna i rozwiązanie |
| --- | --- |
| 401 mimo dobrego hasła | Sprawdź `sudo std-server list`. W logu błędów: `user "..." was not found` (brak konta albo literówka w loginie) albo `password mismatch` (ustaw hasło ponownie przez `add-tech` / `add-admin`). |
| 401 dla technika w `/std-admin/` | Tak ma być: `/std-admin/` jest tylko dla administratorów. |
| 403 przy GET `/std/drop/` | Tak ma być: `drop/` przyjmuje tylko zapis (PUT). Raporty odczytuje administrator przez `/std-admin/drop/`. |
| 403 dla pliku wgranego przez scp | Złe prawa albo właściciel. Uruchom `sudo std-server rebuild`. |
| 404 | Zły adres: pliki są pod `/std/...` i `/std-admin/...` (małe litery, z `/` po nazwie katalogu). |
| 409 przy MKCOL | Brak katalogu nadrzędnego albo adres bez `/` na końcu. |
| 413 | Plik za duży: do 50 MB w `/std/drop/`, do 4 GB w `/std-admin/`. |
| 500 i `crypt_r() failed` w logu | System nie obsługuje haseł bcrypt. Użyj obsługiwanej wersji systemu (patrz wyżej). |
| Instalka: „Test nowej konfiguracji nginx (nginx -t) nie przeszedł” | Poprzednia konfiguracja została przywrócona, a powód jest wypisany nad komunikatem. Jeśli instalka pisze, że konfiguracja sprzed zmian też nie przechodzi testu, błąd jest w innym pliku nginx - uruchom `sudo nginx -t`. |
| `Address family not supported` | System bez IPv6, a inna strona ma `listen [::]:80` (zwykle `/etc/nginx/sites-enabled/default`). Usuń te linie albo wyłącz tę stronę (`rm /etc/nginx/sites-enabled/default` usuwa tylko dowiązanie) i uruchom instalkę ponownie. |
| `could not build server_names_hash` | Długa nazwa serwera. Dopisz `server_names_hash_bucket_size 128;` w sekcji `http` pliku `/etc/nginx/nginx.conf`. |
| Instalka: „Serwer nie odpowiada zgodnie z oczekiwaniem” | Zwykle port 80 albo 443 jest zajęty przez inny program - sprawdź `sudo ss -ltnp`. Konfiguracja została przywrócona. |
| Komputery z STD nie łączą się, przeglądarka zgłasza błąd certyfikatu | Certyfikat samopodpisany, wygasły, bez nazwy serwera albo bez certyfikatów pośrednich. Sprawdź: `openssl s_client -connect pobieranie.firma.pl:443 -servername pobieranie.firma.pl`. |
| Brak połączenia z innych komputerów | Zapora. Instalka otwiera porty w `firewalld` i `ufw`, jeśli są włączone. Inne zapory (np. w chmurze) ustaw ręcznie: TCP 80 i 443. |
| SELinux (RHEL / Rocky / Alma) blokuje dostęp | Instalka ustawia etykiety `httpd_sys_rw_content_t` dla danych. Po ręcznym przeniesieniu plików: `sudo restorecon -R /srv/std`; odmowy: `sudo ausearch -m avc -ts recent`. |
| Kontener bez systemd | Instalka uruchamia nginx bezpośrednio (`nginx`) i przeładowuje go przez `nginx -s reload`. Po restarcie kontenera uruchom `nginx` ręcznie. |

Przywrócenie kopii zapasowej konfiguracji nginx:

```bash
sudo ls /var/lib/std-server/backup/
sudo cp /var/lib/std-server/backup/<data>/std-server.conf /etc/nginx/conf.d/std-server.conf
sudo nginx -t && sudo systemctl reload nginx
```

## Odinstalowanie

Instalka nie ma trybu odinstalowania, bo nigdy nie usuwa danych. Żeby wyłączyć serwer STD:

```bash
sudo rm /etc/nginx/conf.d/std-server.conf
sudo nginx -t && sudo systemctl reload nginx
```

Dane i konta zostają w `/srv/std`, `/etc/std-server` i `/var/lib/std-server`. Usuń je ręcznie,
gdy na pewno nie będą potrzebne.
