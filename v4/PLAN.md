# STD 4 (Smart Tool for Deployment) - plan

STD 4 to nowa, mniejsza wersja narzędzia do przygotowania komputera z Windows. Powstaje w folderze `v4/` tego repozytorium.
STD 3.1 w katalogu głównym działa bez zmian do końca pilotażu, potem STD 4 je zastąpi.

Szczegóły techniczne (format konfiguracji, funkcje modułów, serwer) są w [SPEC.md](SPEC.md).

## Decyzje właściciela (1 listopada 2026)

| Temat | Decyzja |
| --- | --- |
| Repozytorium | To samo repo (SmartDeployTool), STD 4 w folderze `v4/`. |
| Konfiguracja | Na serwerze, podpisana. STD na komputerze tylko ją pobiera i sprawdza podpis, nigdy jej nie zmienia. Zmienia ją edytor administratora przyciskiem „Opublikuj”. |
| Serwer | Dwie instalki: **Windows Server + IIS** oraz **Linux + nginx**. Na start same pliki (HTTPS + konta techników), bez API. |
| Aktualizacja serwera | Ponowne uruchomienie instalki w nowszej wersji zachowuje ustawienia, konta techników, opublikowane konfiguracje i pliki instalatorów. |
| Aktualizacja STD na komputerze | Zachowuje ustawienia lokalne: adres serwera, motyw, ostatnio wybranego klienta, lokalizację i profil, kopie konfiguracji. |
| Dołączenie do domeny | Login i hasło konta **zawsze wpisywane ręcznie** przy wdrożeniu, **nigdy nigdzie zapisywane**: ani w konfiguracji, ani w pliku stanu, ani w logu. W konfiguracji jest tylko nazwa domeny i OU. |
| Hasło lokalnego administratora | Tak samo: zawsze wpisywane ręcznie, nigdy zapisywane. W konfiguracji tylko nazwa konta. |
| Uruchamianie | Bez PS2EXE: plik `.ps1` i `Start-STD.cmd`. STD sprawdza, czy działa jako administrator; jeśli nie, sam uruchamia się ponownie jako administrator. |
| Informacje o systemie, Deinstalator | **Zostają.** STD przygotowuje komputer, a usuwanie preinstalowanych programów jest częścią tej pracy. |
| Minimalny system | Windows 10 22H2 i Windows 11 (.NET Framework 4.8, Windows PowerShell 5.1). |

## Co się zmienia dla technika

1. Start z pendrive'a (`Start-STD.cmd`). Bez uprawnień administratora STD sam prosi o podniesienie uprawnień.
2. Jeśli komputer nie ma sieci: ekran „Połącz z siecią” (Wi-Fi albo kabel).
3. Logowanie do serwera STD (konto technika).
4. Wybór: klient → lokalizacja → profil. Ostatnio używane są na górze listy.
5. **Jeden ekran z pytaniami przed startem**:
   - nazwa komputera;
   - login i hasło do domeny;
   - hasło lokalnego administratora;
   - aplikacje i zadania z profilu;
   - karta z informacjami o komputerze.

   Hasła są trzymane tylko w pamięci i kasowane zaraz po użyciu.
6. **„Sprawdź”** pokazuje plan: co zostanie zrobione, co już jest, gdzie jest problem.
7. **„Wdróż”**. Dalej STD o nic nie pyta. Wdrożenie działa w osobnym wątku, więc okno się nie zawiesza. Pauza i Przerwij są dostępne w każdej chwili.
8. Dołączenie do domeny, zmiana nazwy i konto lokalnego administratora są w pierwszej fazie, przed restartem. Po restarcie STD sam wznawia pracę i nie potrzebuje haseł. Jeśli pierwsza faza się przerwie, STD po wznowieniu zapyta o hasła ponownie.
9. Na końcu uczciwe podsumowanie (OK / Pominięto / Błąd), „Ponów nieudane” i raport wysyłany na serwer.

## Zakres STD 4

**Zostaje (często przebudowane):**
- czekanie na sieć i Wi-Fi;
- wstrzymanie usypiania;
- usuwanie bloatware (aplikacje Sklepu);
- **Deinstalator** (okno do ręcznego usuwania programów i krok usuwający preinstalowane programy z listy w konfiguracji);
- odinstalowanie Microsoft 365 i OneDrive;
- instalacja aplikacji (udział / HTTPS / winget, z wykrywaniem, czy program już jest, i sprawdzaniem SHA-256), w tym TeamViewer i antywirus jako zwykłe pozycje katalogu;
- konto lokalnego administratora;
- zmiana nazwy i dołączenie do domeny;
- ustawienia systemu i rejestr;
- skrypty poinstalacyjne;
- BitLocker (klucz odzyskiwania do AD i zaszyfrowany na serwer, nie do pliku tekstowego);
- audyt sprzętu (w raporcie);
- **Informacje o systemie**;
- motyw jasny i ciemny;
- ikona;
- podgląd okien w CI.

**Odpada:**
- PIN 2137 i okno powitalne (zastępuje je logowanie do serwera);
- edytor ustawień w kliencie (przechodzi do edytora administratora);
- przeglądarka logów (zastępuje ją panel dziennika na żywo i log każdego wdrożenia);
- punkt przywracania;
- ikona w zasobniku;
- Dry-Run (zastępuje go „Sprawdź”);
- zaślepka Intune;
- „Windows Update”, który tylko otwiera panel (prawdziwa instalacja aktualizacji ewentualnie w 4.1);
- dźwięk przy każdym błędzie.

## Etapy

| Etap | Zakres | Stan |
| --- | --- | --- |
| 1. Fundament | Moduły: log z maskowaniem, JSON, HTTP, podpis i weryfikacja, sekrety, konfiguracja (pobranie, kopia, ochrona przed starszą wersją), konwerter starego `config.json`; narzędzia „utwórz certyfikat” i „opublikuj”; testy | kod gotowy, testy w CI |
| 1b. Serwer | Instalka IIS (Windows Server) i nginx (Linux), dodawanie kont techników, aktualizacja z zachowaniem ustawień, testy instalacji w CI | kod gotowy, testy w CI |
| 2. Silnik i okna | Wdrożenie w osobnym wątku, pauza i przerywanie, plik stanu, okno wdrożenia, motyw, ekrany „Połącz” i „Logowanie”, wybór klienta, samodzielne podniesienie uprawnień | następny |
| 3. Kroki | Przeniesienie zadań z 3.1, instalator z wykrywaniem i SHA-256, ekran pytań, „Sprawdź”, Deinstalator, Informacje o systemie | |
| 4. Restart i raport | Wznowienie po restarcie, fazy, „Ponów nieudane”, raport, klucz BitLocker na serwer, „Zgłoś problem” | |
| 5. Edytor administratora | Z obecnych Ustawień: klienci i lokalizacje, SHA-256 z pliku, „Opublikuj”, import starego `config.json` | |
| Pilot | Jeden klient, 3.1 jako zapas | |
| 6. API (opcjonalnie) | Logowanie z MFA, wydawanie sekretów, historia wdrożeń | |

## Pliki

```text
v4/
  PLAN.md, SPEC.md, README.md
  STD.psd1, STD.psm1            moduł STD (wersja, ładowanie plików z Modules\)
  Modules\                      Std.Core.ps1, Std.Http.ps1, Std.Crypto.ps1, Std.Config.ps1, Std.Legacy.ps1, Std.TrustedSigners.psd1
  schema\                       std-release.v4.schema.json
  Tools\                        New-StdSigningCertificate.ps1, Publish-StdRelease.ps1, Convert-StdLegacyConfig.ps1
  Server\IIS\                   Install-StdServer.ps1, Add-StdServerUser.ps1, README.md (instalacja, aktualizacja, konta, bootstrap.json)
  Server\Linux\                 install-std-server.sh, std-server, README.md (to samo dla nginx)
  Tests\                        *.Tests.ps1 (Pester 5, Windows PowerShell 5.1), Invoke-StdSmoke.ps1 (test od początku do końca), Fixtures\
```
