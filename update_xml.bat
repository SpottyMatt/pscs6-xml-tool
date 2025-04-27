@echo off
setlocal enabledelayedexpansion

echo Running update script...

if "%~1"=="" (
    powershell -ExecutionPolicy Bypass -File "%~dp0update_xml.ps1"
) else (
    powershell -ExecutionPolicy Bypass -File "%~dp0update_xml.ps1" -XmlFilePath "%~1"
)

echo.
echo Script execution completed.

:: Give PowerShell a moment to finish writing to the log file
timeout /T 2 >nul

:: Display log file if it exists
if exist "%~dp0update_log.txt" (
    echo.
    echo Log file content:
    echo =======================================
    type "%~dp0update_log.txt"
    echo =======================================
)
pause
