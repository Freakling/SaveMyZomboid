@echo off
setlocal

:: Get directory of this batch file
set SCRIPT_DIR=%~dp0

:: Run the PowerShell script in bypass mode
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SaveMyZomboid.ps1"

endlocal
exit /b
