@echo off
setlocal

set "APP_ROOT=%~dp0.."
cd /d "%APP_ROOT%"
set "PYTHONPATH=%APP_ROOT%\src"

where py >nul 2>nul
if %ERRORLEVEL%==0 (
  py -3 -m pella_order_automation.gui
) else (
  python -m pella_order_automation.gui
)

if errorlevel 1 pause
