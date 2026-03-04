@echo off
setlocal

set "PORT=%~1"
if "%PORT%"=="" set "PORT=8091"

cd /d "%~dp0"

echo Starting standalone OCR Picking Ticket server...
echo Folder: %cd%
echo URL: http://127.0.0.1:%PORT%/
echo.

start "" "http://127.0.0.1:%PORT%/"

where py >nul 2>nul
if %errorlevel%==0 (
  py -m http.server %PORT%
  goto :eof
)

where python >nul 2>nul
if %errorlevel%==0 (
  python -m http.server %PORT%
  goto :eof
)

echo Python launcher not found. Falling back to PowerShell static server...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_standalone.ps1" -Port %PORT%
if %errorlevel% neq 0 (
  echo.
  echo PowerShell fallback failed. Install Python 3 or run from VS Code Live Server.
  pause
)
