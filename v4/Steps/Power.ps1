# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Steps\Power.ps1 - krok „Wstrzymanie usypiania” (SPEC 9).
#
# SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED) w wątku silnika
# (Set-StdKeepAwake z modułu STD): komputer nie zaśnie i nie wygasi ekranu do końca wdrożenia.
# Silnik sam przywraca zwykłe usypianie, gdy wdrożenie się kończy (także po błędzie i przerwaniu).
# Krok jest wcześnie (Order 5), żeby komputer nie zasnął już w czasie czekania na sieć.

$powerStep = @{
    Id     = 'Power.SuspendSleep'
    Title  = 'Wstrzymanie usypiania'
    Phase  = 1
    Order  = 5
    TaskId = 'SuspendSleep'
    Detect = {
        param($Context)
        if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
            return @{ State = 'Blocked'; Detail = 'Wstrzymanie usypiania działa tylko w systemie Windows.' }
        }
        return @{ State = 'Todo'; Detail = 'Komputer nie zaśnie i nie wygasi ekranu do końca wdrożenia.' }
    }
    Run    = {
        param($Context)
        if (Set-StdKeepAwake -Enabled $true) {
            return @{ Result = 'Ok'; Detail = 'Usypianie i wygaszanie ekranu wstrzymane do końca wdrożenia.' }
        }
        return @{ Result = 'Error'; Detail = 'Nie udało się wstrzymać usypiania komputera (SetThreadExecutionState).' }
    }
}
Register-StdStep @powerStep
Remove-Variable -Name powerStep -ErrorAction SilentlyContinue
