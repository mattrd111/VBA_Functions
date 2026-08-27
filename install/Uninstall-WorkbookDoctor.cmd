@echo off
setlocal
title Workbook Doctor - remove

set "TARGET=%APPDATA%\Microsoft\AddIns\WorkbookDoctor.xlam"

echo.
echo   Removing Workbook Doctor...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "try { $x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $null = $x.Workbooks.Add(); foreach ($a in $x.AddIns) { if ($a.Name -eq 'WorkbookDoctor.xlam') { $a.Installed = $false } }; $x.Quit() } catch { }"

if exist "%TARGET%" del /Q "%TARGET%" >nul 2>&1

if exist "%TARGET%" (
    echo   Switched off, but the file could not be deleted.
    echo   Close Excel completely and run this again.
) else (
    echo   Removed.
)
echo.
pause
endlocal
