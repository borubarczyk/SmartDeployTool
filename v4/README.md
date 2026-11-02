# STD 4 (Smart Tool for Deployment)

STD 4 to nowa wersja narzędzia, które przygotowuje komputery z Windows 10 22H2 i Windows 11 do pracy
(Windows PowerShell 5.1, .NET Framework 4.8). Powstaje w tym folderze. STD 3.1 z katalogu głównego
repozytorium działa bez zmian do końca pilotażu.

Najważniejsze zasady:

- Konfiguracja klienta leży na serwerze STD i jest podpisana. STD na komputerze tylko ją pobiera
  i sprawdza podpis - nigdy jej nie zmienia.
- **Login i hasło do domeny oraz hasło lokalnego administratora technik zawsze wpisuje ręcznie.**
  Nie ma ich w konfiguracji, w pliku stanu, w logu ani w żadnym innym pliku. Walidacja odrzuca
  konfigurację, która je zawiera.
- Bez PS2EXE: STD to pliki `.ps1` i `Start-STD.cmd`. STD sprawdza, czy działa jako administrator,
  a jeśli nie, sam uruchamia się ponownie jako administrator (`Test-StdIsAdmin`, `Start-StdElevated`).
- Informacje o systemie i Deinstalator zostają: usuwanie preinstalowanych programów jest częścią
  przygotowania komputera.
- Aktualizacja STD zachowuje ustawienia lokalne (`%ProgramData%\STD`: adres serwera, motyw, ostatni
  wybór, kopie konfiguracji). Aktualizacja serwera (ponowne uruchomienie instalki) zachowuje ustawienia,
  konta i opublikowane pliki.

Plan i decyzje: [PLAN.md](PLAN.md). Umowa techniczna (funkcje, formaty, kody błędów): [SPEC.md](SPEC.md).

## Stan

| Etap | Stan |
| --- | --- |
| 1. Fundament | gotowy |
| 1b. Serwer (IIS i nginx) | gotowy, testy instalacji w CI |
| 2. Silnik i okna | gotowy (testy w CI przechodzą, podgląd okien bez uwag) |
| 3. Kroki i ekran pytań | następny |

Etap 1 zawiera:

- moduł `STD.psd1` / `STD.psm1` z plikami w `Modules\`:
  - `Std.Core.ps1` - wersja, lokalne katalogi danych, JSON, log z maskowaniem sekretów, sprawdzenie
    uprawnień i ponowne uruchomienie jako administrator;
  - `Std.Http.ps1` - HTTPS (TLS 1.2 / 1.3), logowanie Basic, pobieranie z serwera, udziału i dysku;
  - `Std.Crypto.ps1` - SHA-256, podpisana koperta `*.sdt`, zaszyfrowane sekrety, wstawki `{{secret:nazwa}}`;
  - `Std.Config.ps1` - plik startowy, pobranie konfiguracji (serwer, kopie lustrzane, cache, pendrive),
    ochrona przed starszą rewizją, walidacja, konfiguracja wynikowa dla lokalizacji i profilu;
  - `Std.Legacy.ps1` - konwersja `config.json` z STD 3.1;
  - `Std.TrustedSigners.psd1` - zaufane certyfikaty podpisu (w repozytorium pusta lista);
- narzędzia w `Tools\`: `New-StdSigningCertificate.ps1`, `Convert-StdLegacyConfig.ps1`, `Publish-StdRelease.ps1`;
- schemat konfiguracji `schema\std-release.v4.schema.json` i przykład `Tests\Fixtures\release.v4.json`;
- testy Pester 5 w `Tests\` i test dymny `Tests\Invoke-StdSmoke.ps1`.

Etap 2 zawiera:

- `Modules\Std.Engine.ps1` - silnik wdrożenia w osobnym wątku: kroki z `Steps\`, „Sprawdź” (plan bez zmian
  w systemie), „Wdróż”, pauza i przerywanie, plik stanu `state.json` (bez haseł), log wdrożenia
  `%ProgramData%\STD\Logs\<wdrożenie>\deploy.log` i `events.jsonl`, uruchamianie programów, informacje o komputerze;
- `Modules\Std.App.ps1` - start programu: kopia narzędzia do `%ProgramData%\STD\App\<wersja>`, usuwanie starych
  wersji, argumenty ponownego uruchomienia;
- `Steps\Network.ps1` (czekanie na sieć) i `Steps\Power.ps1` (wstrzymanie usypiania) - pierwsze kroki;
  pozostałe kroki z 3.1 i ekran pytań to etap 3;
- `GUI\` - motyw jasny i ciemny (`Std.Theme.ps1`, kolory i style z 3.1), zadania w tle dla okien (`Std.UI.ps1`)
  i ekrany w `GUI\Screens\`;
- `SmartToolforDeployment.ps1` i `Start-STD.cmd` - program i plik startowy;
- `Tools\Render-StdUiPreview.ps1` - podgląd wszystkich okien w PNG (workflow `v4-ui-preview.yml`).

## Uruchomienie z pendrive'a (technik)

Na pendrive kopiujemy cały folder `v4` i kładziemy obok `Start-STD.cmd` plik `bootstrap.json` z adresem serwera
(wypisuje go instalka serwera). Opcjonalnie: `offline\` (kopie konfiguracji do pracy bez serwera) i `wifi\*.xml`
(profile sieci Wi-Fi).

1. Dwuklik `Start-STD.cmd`. STD prosi o uprawnienia administratora (okno UAC), kopiuje się do
   `%ProgramData%\STD\App\<wersja>` i uruchamia się stamtąd - pendrive nie jest potrzebny w trakcie wdrożenia
   (tylko do pobrania kopii offline i profili Wi-Fi przed startem).
2. **Połącz z siecią** - tylko gdy komputer nie ma sieci: lista sieci Wi-Fi, hasło sieci, import profili
   z pendrive'a, „Odśwież”, „Pomiń”.
3. **Logowanie** - konto technika na serwerze STD. Gdy nie ma `bootstrap.json`, technik wpisuje adres serwera
   (STD go zapamięta). „Pracuj offline”, gdy jest kopia listy klientów.
4. **Okno główne** - wybór klient → lokalizacja → profil (ostatnio używane na górze), podsumowanie konfiguracji,
   karta „Komputer”, przyciski „Sprawdź” (plan: co zostanie zrobione, co już jest, gdzie jest problem) i „Wdróż”.
   W trakcie wdrożenia: lista kroków, pasek postępu, dziennik na żywo, Pauza / Wznów / Przerwij. Na końcu
   podsumowanie OK / Pominięto / Błąd, „Ponów nieudane” i „Kopiuj podsumowanie”.

Parametry (także przez `Start-STD.cmd`): `-Server https://.../std/` (zamiast `bootstrap.json`), `-NoCopy` (praca
prosto z pendrive'a, bez kopii). Nieoczekiwany błąd programu trafia do `%TEMP%\STD-crash.log`.

## Pierwsza konfiguracja (administrator)

1. Certyfikat do podpisu (Windows, raz): `Tools\New-StdSigningCertificate.ps1`. Wypisany wpis wklej do
   `Modules\Std.TrustedSigners.psd1`.
2. Konwersja starej konfiguracji:
   `Tools\Convert-StdLegacyConfig.ps1 -Path .\config.json -ClientId lizard -ClientName 'Lizard-Tech'`.
   Popraw ostrzeżenia (np. adresy `http://` na `https://`).
3. Sekrety z listy „do ustawienia” (np. token TeamViewer). Do czasu edytora administratora (etap 5):

   ```powershell
   Import-Module .\v4\STD.psd1
   $doc = ConvertFrom-StdJson -Json (Get-Content -LiteralPath '.\lizard.release.v4.json' -Raw -Encoding UTF8)
   $haslo = Read-Host -Prompt 'Hasło wdrożeniowe klienta' -AsSecureString
   $doc['secrets'] = New-StdSealedStore -Password $haslo
   Set-StdSealedItem -Store $doc['secrets'] -Name 'teamviewer-apitoken' -Value (Read-Host -Prompt 'Token TeamViewer' -AsSecureString) -Password $haslo
   Write-StdJsonFile -Path '.\lizard.release.v4.json' -InputObject $doc
   ```

4. Publikacja: `Tools\Publish-StdRelease.ps1 -SourcePath .\lizard.release.v4.json -CertificateThumbprint <odcisk> -Destination https://<serwer>/std-admin/ -Credential (Get-Credential)`.

## Serwer

Dwie instalki, każda z opisem instalacji, aktualizacji i kont techników:

- Windows Server + IIS: [Server/IIS/Install-StdServer.ps1](Server/IIS/README.md), konta: `Server\IIS\Add-StdServerUser.ps1`;
- Linux + nginx: [Server/Linux/install-std-server.sh](Server/Linux/README.md), konta: `std-server add-tech <login>`.

Aktualizacja serwera to ponowne uruchomienie instalki w nowszej wersji. Instalka bierze zapisane ustawienia,
robi kopię zapasową konfiguracji serwera WWW i nie usuwa danych ani kont. Na końcu wypisuje gotowy
`bootstrap.json` dla pendrive'a. Testy instalek w CI: `.github/workflows/v4-server-iis.yml`
i `.github/workflows/v4-server-linux.yml`.

## Testy

Testy Pester 5 (Windows PowerShell 5.1, Pester 5.6.1) uruchamia workflow `pester.yml` razem z testami 3.1.
Ręcznie:

```powershell
Install-Module -Name Pester -RequiredVersion 5.6.1 -Scope CurrentUser -Force -SkipPublisherCheck
Import-Module Pester -RequiredVersion 5.6.1
Invoke-Pester -Path .\v4\Tests -Output Detailed
```

Testy nie wymagają administratora, sieci ani magazynu certyfikatów (certyfikaty testowe powstają w pamięci,
dane lokalne trafiają do `TestDrive:` przez `$env:STD_DATA_ROOT`). Testy okien (`Std.UI.Tests.ps1`) budują
każde okno w obu motywach - potrzebują Windows i wątku STA (zwykły `powershell.exe`); gdzie indziej te bloki
są pomijane. Testy silnika (`Std.Engine.Tests.ps1`) uruchamiają prawdziwy silnik z testowymi krokami.

Podgląd okien (każde okno w obu motywach jako PNG, z kontrolą uciętego tekstu, kolorów i kontrastu):

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\v4\Tools\Render-StdUiPreview.ps1 -Theme Dark
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\v4\Tools\Render-StdUiPreview.ps1 -Theme Light -FailOnFindings
```

Pliki trafiają do `v4\ui-preview\<motyw>\`. W GitHub Actions robi to workflow `v4-ui-preview.yml` (artefakt
`v4-ui-preview`).

Test dymny przechodzi całą drogę konfiguracji: certyfikat w pamięci, konwersja `config.example.json`,
sekrety, publikacja do katalogu serwera, pobranie z kopii lustrzanej przy niedostępnym serwerze,
konfiguracja wynikowa, podstawienie sekretów w argumentach TeamViewer i maskowanie ich w logu.
Działa w Windows PowerShell 5.1 i w PowerShell 7 (także na Linuksie), bez Pestera:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\v4\Tests\Invoke-StdSmoke.ps1
pwsh -NoProfile -File ./v4/Tests/Invoke-StdSmoke.ps1 -KeepTemp
```

Wypisuje PASS / FAIL dla każdego kroku i kończy się kodem 1, gdy któryś krok się nie udał.

## Pliki

Pliki `.ps1`, `.psm1` i `.psd1` są w UTF-8 z BOM (bez BOM Windows PowerShell 5.1 psuje polskie znaki),
pozostałe w UTF-8 bez BOM. Końce linii: LF, poza `Start-STD.cmd` (CRLF i bez polskich znaków, bo czyta go
`cmd.exe`; pilnuje tego `.gitattributes`). Szczegóły: [SPEC.md](SPEC.md), sekcja 0.
