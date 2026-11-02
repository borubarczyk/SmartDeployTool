# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Manifest modułu STD. Wersja narzędzia jest wpisana tylko tutaj (ModuleVersion);
# reszta kodu bierze ją z Get-StdVersion.
@{
    RootModule           = 'STD.psm1'
    ModuleVersion        = '4.0.0'
    GUID                 = 'ac626bb6-e29d-4e74-8f2b-71ddaf7691e9'
    Author               = 'Borys Kaleta'
    CompanyName          = 'Borys Kaleta'
    Copyright            = '(c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone.'
    Description          = 'STD (Smart Tool for Deployment) - przygotowanie komputerów z Windows: log, JSON, HTTP, podpis i konfiguracja.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = '*'
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags = @('STD', 'Deployment', 'Windows')
        }
    }
}
