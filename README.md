# Smart Tool for Deployment (STD)

**Smart Tool for Deployment** to zaawansowane narzędzie z graficznym interfejsem (GUI) oparte na technologii Windows Presentation Foundation (WPF) i PowerShell. Zostało stworzone, aby w pełni zautomatyzować i ustandaryzować proces wdrażania oraz konfiguracji nowych stacji roboczych z systemem Windows przez inżynierów IT.

## Główne funkcjonalności

- **Zabezpieczenie logowaniem:** Dostęp do uruchomienia narzędzia oraz zmian w oknie `Ustawień` chroniony jest kodem PIN (domyślnie: `2137`).
- **Nowoczesny interfejs GUI (WPF):** Ten sam system wyglądu co w narzędziach ServerReview i AD Manager Diamond: grafitowo-granatowe tło, niebieski akcent, karty, kontrolki o wysokości 30 px, tabele z co drugim wierszem wyróżnionym, zaokrąglone menu i podpowiedzi, ikony z systemowej czcionki Segoe Fluent Icons / Segoe MDL2 Assets zamiast emoji, pasek tytułu okna w kolorach motywu oraz krótkie powiadomienia (toast) w rogu okna zamiast zbędnych okienek. Motyw ciemny (domyślny, kolory 1:1 z narzędzi referencyjnych) i jasny przełącza przycisk z ikoną słońca/księżyca w pasku nagłówka - zmiana działa od razu we wszystkich otwartych oknach i jest zapamiętywana w `config.json`. Wszystkie okna (logowanie, powitanie, komunikaty z kolorowym kafelkiem ikony, okna dialogowe, Logi, Informacje o systemie, Deinstalator, Ustawienia) korzystają z jednego wspólnego zestawu stylów. Ikona narzędzia (pasek tytułu, pasek zadań, zasobnik systemowy) jest rysowana w kodzie - ten sam kafelek z paczką co logo w nagłówku - więc nie wymaga pliku `.ico` obok skryptu.
- **Czytelna Przeglądarka Logów (widok tabelaryczny):** Każdy wpis loguje pełną datę i godzinę, poziom (`INFO`/`ERROR`) oraz kontekst (`System`/`Automat`/`Użytkownik`). Przeglądarka logów prezentuje to w tabeli z sortowalnymi (kliknięcie nagłówka) i filtrowalnymi kolumnami **Rodzaj | Data zdarzenia | Kontekst | Informacja** - z kolorowymi pastylkami i kolorem wiersza wg rodzaju zdarzenia (błąd, sukces, ostrzeżenie, Dry-Run), wyszukiwarką na żywo oraz eksportem czytelnego **raportu HTML** (do wysłania klientowi/dołączenia do zgłoszenia).
- **Prosty układ okna głównego (jak w ServerReview / AD Manager Diamond):** u góry pasek nagłówka z logo, nazwą i wersją narzędzia, loginem operatora, pastylką stanu sieci oraz przyciskami z ikonami *Logi*, *Narzędzia* (menu), *Ustawienia* (koło zębate) i przełącznikiem motywu; pod nim profil wdrożenia, przycisk *Wybierz aplikacje*, licznik zaznaczonych zadań i *Wszystkie* / *Żadne*; pośrodku trzy karty z zadaniami pogrupowanymi w sekcje (Przygotowanie, Czyszczenie systemu, Oprogramowanie, Konta i domena, System i bezpieczeństwo, Zakończenie - pełny opis zadania w podpowiedzi); na dole pasek wdrożenia z kropką stanu (gotowe / w trakcie / pauza / błąd), postępem, stoperem i jedynym głównym przyciskiem *Rozpocznij konfigurację* (Pauza/Przerwij pojawiają się w trakcie wdrożenia). Okno mieści się w całości na ekranie 1024×768.
- **Monitorowanie połączenia (Dynamiczny Status Sieci):** Kropka wizualnie informująca o poprawności adresu IP oraz o dostępności repozytorium WWW (żądanie HTTP HEAD wykonywane w osobnym wątku, więc nawet niedostępny serwer nie spowalnia interfejsu).
- **Zarządzanie Oprogramowaniem:**
  - Automatyczna, cicha instalacja wybranych aplikacji z różnych źródeł (ścieżki sieciowe UNC, zasoby Web/HTTP(S), oraz pakiety **Winget**).
  - **Katalog programów winget** (Ustawienia → Programy → *Dodaj z katalogu winget...*): ok. 40 sprawdzonych programów (przeglądarki, biuro i PDF, chmura, narzędzia, administracja, środowiska uruchomieniowe) do zaznaczenia jednym kliknięciem. Program z katalogu ma tylko `WingetId`, więc instaluje się przez winget przy każdym źródle globalnym - bez instalatora na udziale. Programy, które już są w konfiguracji (ta sama nazwa, alias albo Winget ID), są wyszarzone.
  - Zintegrowane uwierzytelnianie (WebAuth) do pobierania zastrzeżonych instalatorów HTTP.
  - Bieżący wgląd w działanie deinstalacji/instalacji za sprawą dynamicznych komunikatów nad paskiem postępu wdrożenia.
  - **Pełna kontrola:** Możliwość **wstrzymania (Pauza)**, **wznowienia** lub **przerwania** wdrożenia w każdym momencie jego trwania.
- **Wbudowany Deinstalator Zbiorczy:** Dedykowany moduł pozwalający na masowe i ciche usuwanie zainstalowanego na stacji oprogramowania. Zawiera system "Zabijania Procesu" (Kill Process), statusy na żywo jako kolorowe pastylki (Usunięto / Błąd / Brak polecenia) oraz możliwość eksportowania listy programów do CSV.
  - **Narzędzia poboczne:** "Podwójne kliknięcie" wywołujące interaktywny deinstalator wybranej aplikacji, funkcja "Pokaż w eksploratorze", podgląd statusów powiązany na żywo z `PID` i wywoływanym poleceniem ukrytym w systemie (Cmd).
- **Graficzny Edytor Konfiguracji (okno Ustawienia):** Pełne zarządzanie plikiem `config.json` z poziomu GUI – jedno okno z nawigacją po lewej (ikona i nazwa sekcji, jak w AD Manager Diamond): *Źródła*, *Domena i Wi-Fi*, *TeamViewer i AV*, *Programy*, *Profile*, *Zadania* (domyślnie zaznaczone zadania), *Rejestr*, *Skrypty* i *Zaawansowane* (kopia zapasowa, aktualizacje narzędzia). Listy z przyciskami Dodaj/Edytuj/Usuń są bezpośrednio w zakładkach, zmiany trafiają do pliku dopiero po *Zapisz* (wskaźnik „Niezapisane zmiany” i pytanie przy zamykaniu).
- **Skrypty Poinstalacyjne (Post-Install):** Możliwość pobrania i wykonania dodatkowych własnych skryptów `.ps1`, `.bat` tuż po zakończeniu instalacji oprogramowania.
- **Audyt Sprzętowy:** Zbieranie szczegółowych informacji o PC (Model, CPU, RAM, NVMe, Sieć) i automatyczny zrzut do pliku na serwerze po udanej konfiguracji.
- **Szyfrowanie BitLocker:** Zautomatyzowany mechanizm szyfrowania sprzętowego dysku systemowego z automatycznym generowaniem i eksportem bezpiecznego klucza odzyskiwania na serwer.
- **System Tweaks i Bloatware:** Moduły wymuszające optymalizację Windows, usuwające śmieciowe aplikacje (TikTok, Xbox itp.), wyłączające usługi zbierające dane telemetryczne czy Cortanę.
- **Narzędzia (menu przycisku *Narzędzia*):** Błyskawiczny dostęp do *Informacji o systemie* (z kopiowaniem do schowka), deinstalatora, najważniejszych konsol Windows (Właściwości systemu, Zarządzanie komputerem, Edytor rejestru, Urządzenia i drukarki) oraz edycji i przeładowania `config.json`.
- **Tryb testowy (Dry-Run):** Opcjonalny checkbox na ekranie głównym pozwala zasymulować pełne wdrożenie - narzędzie loguje, co zostałoby zrobione (instalacje, deinstalacje, zmiany rejestru, dołączenie do domeny, BitLocker itd.), ale nie wprowadza żadnej rzeczywistej zmiany w systemie. Przydatne do weryfikacji poprawności konfiguracji przed wdrożeniem na realnej stacji.
- **Punkt przywracania i checkpoint wdrożenia:** Opcjonalny checkbox tworzy punkt przywracania systemu Windows tuż przed startem wdrożenia (System Restore). Niezależnie od tego, postęp każdego wdrożenia jest na bieżąco zapisywany do `C:\deploy-checkpoint.json`, co ułatwia diagnozę, na którym kroku wdrożenie zostało przerwane w razie awarii.
- **Aktualizacje narzędzia:** Przycisk *"Sprawdź teraz"* w zakładce *Zaawansowane* okna Ustawień (tam też ścieżka aktualizacji i opcja *Sprawdzaj przy starcie* - wtedy narzędzie samo sprawdza wersję po otwarciu okna głównego) porównuje lokalną wersję z plikiem `std_version.json` publikowanym pod skonfigurowaną ścieżką (`AutoUpdate.VersionCheckPath` w `config.json` - UNC lub URL) i, po potwierdzeniu, pobiera oraz podmienia plik `.ps1`, po czym uruchamia narzędzie ponownie. Przed podmianą sprawdzana jest składnia pobranego skryptu oraz - jeśli `std_version.json` zawiera pole `Sha256` - jego suma kontrolna (`(Get-FileHash .\SmartToolforDeployment.ps1 -Algorithm SHA256).Hash`). Mechanizm podmienia tylko plik `.ps1`, nie skompilowany `.exe`.

## Wymagania

- **Windows 10/11**
- **PowerShell 5.1+** (Skrypt wymaga uruchomienia jako Administrator)
- **.NET Framework / WPF** (Wbudowane domyślnie w system Windows)
- Moduł **ActiveDirectory (RSAT)** – wymagany tylko, jeśli zamierzasz korzystać z funkcji podłączania do domeny lokalnej.

## Szybki start

1. Prawym przyciskiem myszy kliknij plik `SmartToolforDeployment.ps1` i wybierz opcję **Uruchom z programem PowerShell** (Upewnij się, że okno wywoła się z prawami Administratora).
2. W oknie autoryzacji podaj swój identyfikator (np. Imię) oraz wpisz PIN: `2137`.
3. Przy pierwszym uruchomieniu, jeśli brakuje pliku konfiguracyjnego, narzędzie zaproponuje wygenerowanie podstawowego szablonu `config.json`.
4. Skonfiguruj ścieżki sieciowe i pakiety klikając przycisk **Ustawienia** (ikona koła zębatego) w pasku u góry okna.
5. Zaznacz pożądane zadania instalacyjne na głównym ekranie, wybierz aplikacje z listy.
6. Kliknij niebieski przycisk **"Rozpocznij konfigurację"** (w trybie testowym: *"Rozpocznij symulację"*) i śledź pasek stanu na dole okna. Wszystkie operacje będą też na bieżąco zapisywane na dysku (`C:\deploy-log.txt`, format `[yyyy-MM-dd HH:mm:ss] [INFO|ERROR] [Kontekst] treść`) i dostępne w tabelarycznej **Przeglądarce Logów** (przycisk *"Logi"*) - sortowalne kolumny Rodzaj/Data/Kontekst/Informacja, filtry, wyszukiwarka i eksport raportu HTML.

## Przykładowa struktura `config.json`

Plik generuje się i jest zarządzany automatycznie przez aplikację, lecz można też edytować go w Notatniku. Plik `config.json` zawiera dane konkretnego środowiska (ścieżki, loginy, hasła), dlatego nie jest przechowywany w repozytorium (`.gitignore`). Gotowy przykład bez danych wrażliwych znajdziesz w `config.example.json` - skopiuj go jako `config.json` obok skryptu i uzupełnij. Nowy plik przykładowy z własnej konfiguracji wygenerujesz skryptem `Create-ConfigExample.ps1`.

```json
{
    "DefaultInstallSource": "network",
    "InstallSourcePaths": {
        "network": "\\\\SERWER\\Instalki\\",
        "web": "https://pobieranie.mojafirma.pl/apps/"
    },
    "DomainJoin": {
        "DomainName": "firma.local",
        "Username": "ad\\administrator"
    },
    "LocalAdmin": {
        "Username": "LokalnyIT"
    },
    "WebAuth": {
        "Username": "admin",
        "Password": "Password123!"
    },
    "HardwareAudit": {
        "ExportPath": "\\\\SERWER\\Audyty\\"
    },
    "PostInstallScripts": [
        "SkryptDrukarki.ps1"
    ],
    "Profiles": {
        "Standard": ["7-Zip", "Google Chrome"],
        "Programiści": ["7-Zip", "Google Chrome", "Git"]
    },
    "Programs": {
        "7-Zip": {
            "Enabled": true,
            "FileName": "7z.exe",
            "SilentArgs": "/S"
        },
        "Google Chrome": {
            "Enabled": true,
            "FileName": "ChromeSetup.exe",
            "SilentArgs": "/silent /install",
            "WingetId": "Google.Chrome"
        },
        "Git": {
            "Enabled": false,
            "FileName": "",
            "SilentArgs": "",
            "WingetId": "Git.Git",
            "WingetArgs": "--scope machine"
        }
    },
    "SystemSettings": {
        "DisableDeliveryOptimization": true,
        "EnableWin10StartMenu": true,
        "DisableTelemetry": true,
        "DisableCortana": true,
        "DisableFastStartup": true,
        "DisableNewsAndInterests": true,
        "CustomRegistry": [
            {
                "Path": "HKLM:\\SOFTWARE\\MojaFirma",
                "Name": "WdrozenieZakonczone",
                "Value": "1",
                "PropertyType": "DWord"
            }
        ]
    },
    "DefaultCheckboxes": {
        "WaitForNetwork": true,
        "RemoveBloatware": true,
        "InstallApplications": true,
        "RunPostInstallScripts": false,
        "DryRun": false,
        "CreateRestorePoint": false
    },
    "AutoUpdate": {
        "Enabled": false,
        "VersionCheckPath": "\\\\SERWER\\Instalki\\STD\\"
    },
    "DarkTheme": true
}
```

### Skąd instaluje się program

| Wpis programu | Źródło globalne `network` / `web` | Źródło globalne `winget` |
| --- | --- | --- |
| `DownloadUrl` (własny adres URL) | zawsze ten adres | zawsze ten adres |
| `FileName` (+ ewentualnie `WingetId`) | plik z udziału / spod adresu web, argumenty `SilentArgs` | winget z `WingetId`, a bez niego `FileName` jako Winget ID (dawny zapis) |
| tylko `WingetId` (np. z katalogu) | winget | winget |

`WingetArgs` to dodatkowe argumenty winget (np. `--scope machine` - instalacja dla wszystkich użytkowników, gdy pakiet ma też wariant tylko dla bieżącego). Brak programu winget na komputerze zgłasza walidacja przed startem (ostrzeżenie z listą pozycji, których dotyczy). Nowy, domyślny `config.json` zawiera na start kilka programów z katalogu, więc profile *Standard* i *Księgowość* od razu mają co zaznaczyć.

## Licencja

Copyright (c) 2025-2026 Borys Kaleta. **Wszelkie prawa zastrzeżone** - [licencja zastrzeżona](LICENSE).

Używanie, kopiowanie, modyfikowanie i rozpowszechnianie programu (w całości lub w części) wymaga uprzedniej zgody autora. Publiczna dostępność kodu nie oznacza zgody na korzystanie z niego.
