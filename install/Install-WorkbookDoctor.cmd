@echo off
setlocal
title Workbook Doctor - install

rem ===========================================================================
rem Installs a built WorkbookDoctor.xlam for the current user.
rem
rem Put this file and WorkbookDoctor.xlam in the same folder and double-click
rem this one. Nothing here needs admin rights, the VBA trust setting, or any
rem change to the PowerShell execution policy - it only copies a file and asks
rem Excel to switch it on.
rem ===========================================================================

set "SOURCE=%~dp0WorkbookDoctor.xlam"
set "TARGET=%APPDATA%\Microsoft\AddIns"

echo.
echo   Workbook Doctor
echo   ---------------
echo.

if not exist "%SOURCE%" (
    echo   WorkbookDoctor.xlam is not in this folder.
    echo.
    echo   Put it next to this file and run this again. If you have not built
    echo   it yet, see addin\INSTALL.md.
    echo.
    pause
    exit /b 1
)

if not exist "%TARGET%" mkdir "%TARGET%" >nul 2>&1

echo   Copying the add-in...
copy /Y "%SOURCE%" "%TARGET%\WorkbookDoctor.xlam" >nul
if errorlevel 1 (
    echo   Could not copy it to %TARGET%
    echo   Close Excel and try again - the file is locked while the add-in is loaded.
    echo.
    pause
    exit /b 1
)

echo   Switching it on in Excel...
rem -Command runs inline code, which the execution policy does not govern, so
rem this works on a machine where .ps1 files are blocked.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "try { $x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $null = $x.Workbooks.Add(); $a = $x.AddIns.Add('%TARGET%\WorkbookDoctor.xlam', $false); $a.Installed = $true; $x.Quit(); exit 0 } catch { exit 1 }"

if errorlevel 1 (
    echo.
    echo   The add-in is in place but Excel would not switch it on by itself.
    echo   Do it once by hand:
    echo.
    echo     Excel ^> File ^> Options ^> Add-ins
    echo     Manage: Excel Add-ins ^> Go... ^> tick "Workbook Doctor"
    echo.
) else (
    echo.
    echo   Installed.
    echo.
    echo   Open Excel and look for "Workbook Doctor" on the Add-ins tab.
    echo   Start with Audit ^> Audit workbook - it only reads and reports.
    echo.
)

pause
endlocal
