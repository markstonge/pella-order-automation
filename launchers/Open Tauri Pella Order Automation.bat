@echo off
setlocal

set "APP_ROOT=%~dp0.."
cd /d "%APP_ROOT%"

where npm >nul 2>nul
if errorlevel 1 (
  echo npm was not found. Please install Node.js, then try again.
  pause
  exit /b 1
)

where cargo >nul 2>nul
if errorlevel 1 (
  echo Cargo was not found. Please install Rust from https://rustup.rs/, then try again.
  pause
  exit /b 1
)

if not exist "%APP_ROOT%\node_modules" (
  echo Installing app dependencies...
  npm install
  if errorlevel 1 pause & exit /b 1
)

set "PYTHONPATH=%APP_ROOT%\src"
npm run desktop:dev

if errorlevel 1 pause
