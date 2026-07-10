# Smart Tool for Deployment (STD)

**Smart Tool for Deployment** to zaawansowane narzędzie z graficznym interfejsem (GUI) oparte na technologii Windows Presentation Foundation (WPF) i PowerShell. Zostało stworzone, aby w pełni zautomatyzować i ustandaryzować proces wdrażania oraz konfiguracji nowych stacji roboczych z systemem Windows przez inżynierów IT.

## Główne funkcjonalności

- **Zabezpieczenie logowaniem:** Dostęp do uruchomienia narzędzia oraz zmian w oknie `Ustawień` chroniony jest kodem PIN (domyślnie: `2137`).
- **Nowoczesny interfejs GUI (WPF):** Spójny i responsywny układ z obsługą motywu Jasnego (☀️) i Ciemnego (🌙). Wszystkie komunikaty, komunikaty o błędach i prośby o potwierdzenie są renderowane w natywnym motywie aplikacji.
- **Monitorowanie połączenia (Dynamiczny Status Sieci):** Działająca w tle (nie obciążająca interfejsu) pulsująca kropka wizualnie informująca o poprawności adresu IP oraz natychmiastowej weryfikacji poprawności repozytorium WWW (żądania HTTP HEAD).
- **Zarządzanie Oprogramowaniem:**
  - Automatyczna, cicha instalacja wybranych aplikacji z różnych źródeł (ścieżki sieciowe UNC, zasoby Web/HTTP(S), oraz pakiety **Winget**).
  - Zintegrowane uwierzytelnianie (WebAuth) do pobierania zastrzeżonych instalatorów HTTP.
  - Bieżący wgląd w działanie deinstalacji/instalacji za sprawą dynamicznych komunikatów nad paskiem postępu wdrożenia.
  - **Pełna kontrola:** Możliwość **wstrzymania (Pauza)**, **wznowienia** lub **przerwania** wdrożenia w każdym momencie jego trwania.
- **Wbudowany Deinstalator Zbiorczy:** Dedykowany moduł pozwalający na masowe i ciche usuwanie zainstalowanego na stacji oprogramowania. Zawiera system "Zabijania Procesu" (Kill Process), statusy na żywo (Sukces/Błąd) oraz możliwość eksportowania listy programów do CSV.
  - **Narzędzia poboczne:** "Podwójne kliknięcie" wywołujące interaktywny deinstalator wybranej aplikacji, funkcja "Pokaż w eksploratorze", podgląd statusów powiązany na żywo z `PID` i wywoływanym poleceniem ukrytym w systemie (Cmd).
- **Graficzny Edytor Konfiguracji:** Pełne zarządzanie plikiem `config.json` z poziomu GUI – bez konieczności ręcznego edytowania kodu. Dodawaj i klonuj programy, zarządzaj niestandardowymi wpisami rejestru i ustalaj opcje domyślne.
- **Skrypty Poinstalacyjne (Post-Install):** Możliwość pobrania i wykonania dodatkowych własnych skryptów `.ps1`, `.bat` tuż po zakończeniu instalacji oprogramowania.
- **Audyt Sprzętowy:** Zbieranie szczegółowych informacji o PC (Model, CPU, RAM, NVMe, Sieć) i automatyczny zrzut do pliku na serwerze po udanej konfiguracji.
- **Szyfrowanie BitLocker:** Zautomatyzowany mechanizm szyfrowania sprzętowego dysku systemowego z automatycznym generowaniem i eksportem bezpiecznego klucza odzyskiwania na serwer.
- **System Tweaks i Bloatware:** Moduły wymuszające optymalizację Windows, usuwające śmieciowe aplikacje (TikTok, Xbox itp.), wyłączające usługi zbierające dane telemetryczne czy Cortanę.
- **Szybkie Narzędzia (Quick Tools):** Błyskawiczny dostęp do najważniejszych konsol Windows (Zarządzanie komputerem, Edytor Rejestru) oraz szczegółowych *Informacji o systemie* z opcją kopiowania prosto do schowka.

## Wymagania

- **Windows 10/11**
- **PowerShell 5.1+** (Skrypt wymaga uruchomienia jako Administrator)
- **.NET Framework / WPF** (Wbudowane domyślnie w system Windows)
- Moduł **ActiveDirectory (RSAT)** – wymagany tylko, jeśli zamierzasz korzystać z funkcji podłączania do domeny lokalnej.

## Szybki start

1. Prawym przyciskiem myszy kliknij plik `SmartToolforDeployment.ps1` i wybierz opcję **Uruchom z programem PowerShell** (Upewnij się, że okno wywoła się z prawami Administratora).
2. W oknie autoryzacji podaj swój identyfikator (np. Imię) oraz wpisz PIN: `2137`.
3. Przy pierwszym uruchomieniu, jeśli brakuje pliku konfiguracyjnego, narzędzie zaproponuje wygenerowanie podstawowego szablonu `config.json`.
4. Skonfiguruj ścieżki sieciowe i pakiety klikając **Ustawienia...**.
5. Zaznacz pożądane zadania instalacyjne na głównym ekranie, wybierz aplikacje z listy.
6. Kliknij zielony przycisk **"ROZPOCZNIJ KONFIGURACJĘ"** i śledź logi oraz pasek postępu. Wszystkie operacje będą też na bieżąco zapisywane na dysku (`C:\deploy-log.txt`).

## Przykładowa struktura `config.json`

Plik generuje się i jest zarządzany automatycznie przez aplikację, lecz można też edytować go w Notatniku.

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
        "Standard": ["7-Zip", "Google Chrome", "PowerToys"],
        "Księgowość": ["7-Zip", "Google Chrome", "Szafir_KIR"]
    },
    "Programs": {
        "7-Zip": {
            "Enabled": true,
            "FileName": "7z.exe",
            "SilentArgs": "/S"
        },
        "Google Chrome": {
            "Enabled": true,
            "FileName": "Google.Chrome",
            "SilentArgs": ""
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
        "RunPostInstallScripts": false
    },
    "DarkTheme": true
}
```