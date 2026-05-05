@echo off
setlocal

set "APP_ROOT=%~dp0.."
cd /d "%APP_ROOT%"
set "PYTHONPATH=%APP_ROOT%\src"

start "" "http://127.0.0.1:8765/"

where py >nul 2>nul
if %ERRORLEVEL%==0 (
  py -3 -m pella_order_automation.web_server
) else (
  python -m pella_order_automation.web_server
)

if errorlevel 1 pause
