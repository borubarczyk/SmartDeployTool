# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Moduł główny STD. Wczytuje pliki z Modules\ w stałej kolejności (SPEC 0.3). Wszystkie pliki są
# dot-source'owane w zakresie tego modułu, więc funkcje i zmienne $script: (odbiorniki logu, maski)
# są wspólne. Brakujący plik jest pomijany z ostrzeżeniem, żeby moduł dało się wczytać także wtedy,
# gdy część plików jeszcze powstaje.

$script:StdModuleFiles = @(
    'Std.Core.ps1'
    'Std.Http.ps1'
    'Std.Crypto.ps1'
    'Std.Config.ps1'
    'Std.Legacy.ps1'
)

$script:StdModulesDir = Join-Path -Path $PSScriptRoot -ChildPath 'Modules'

foreach ($stdModuleFile in $script:StdModuleFiles) {
    $stdModulePath = Join-Path -Path $script:StdModulesDir -ChildPath $stdModuleFile
    if (Test-Path -LiteralPath $stdModulePath -PathType Leaf) {
        . $stdModulePath
    } else {
        Write-Warning "STD: brak pliku modułu $stdModulePath - pomijam go."
    }
}
Remove-Variable -Name stdModuleFile, stdModulePath -ErrorAction SilentlyContinue

# Przy usunięciu modułu (także Import-Module -Force) zamykamy odbiorniki logu, żeby pliki logu
# nie zostały zablokowane przez stare uchwyty.
$ExecutionContext.SessionState.Module.OnRemove = {
    try {
        if ($null -ne $script:StdLogSinks -and (Get-Command -Name 'Unregister-StdLogSink' -ErrorAction SilentlyContinue)) {
            foreach ($sinkName in @($script:StdLogSinks.get_Keys())) {
                try { Unregister-StdLogSink -Name $sinkName } catch { }
            }
        }
    } catch { }
}

Export-ModuleMember -Function '*-Std*'
