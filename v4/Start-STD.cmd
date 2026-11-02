@echo off
rem STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzezone - warunki w pliku LICENSE.
rem
rem Uruchomienie STD z pendrive'a albo z dowolnego folderu: dwuklik albo Start-STD.cmd [parametry],
rem np. Start-STD.cmd -Server https://pobieranie.firma.pl/std/
rem STD sam poprosi o uprawnienia administratora (okno UAC) i skopiuje sie na dysk komputera.
rem
rem Plik celowo bez polskich znakow: cmd.exe czyta go w stronie kodowej OEM. Konce linii CRLF.
rem Zawsze Windows PowerShell 5.1 w trybie STA (wymagany przez okna). Na 64-bitowym Windows zawsze
rem 64-bitowy PowerShell, takze gdy ten plik uruchomi program 32-bitowy (folder Sysnative).

setlocal
set "STD_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "STD_PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%STD_PS%" set "STD_PS=powershell.exe"

"%STD_PS%" -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0SmartToolforDeployment.ps1" %*
exit /b %ERRORLEVEL%
