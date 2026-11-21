@echo off
setlocal
title MSP Ops Center - build the new interface
cd /d "%~dp0"
rem Run this on a PC with Node.js 20 or newer. The server itself never needs Node.js.
where npm >nul 2>&1
if errorlevel 1 (
  echo Node.js is not installed. Install the LTS version from your software centre, then run this again.
  pause
  exit /b 1
)
echo Installing the exact package versions from package-lock.json ...
call npm ci --no-audit --no-fund
if errorlevel 1 goto :fail
echo Checking types and building ...
call npm run build
if errorlevel 1 goto :fail
echo.
echo Done. The new interface is in ..\static\app. Copy the static folder to the server with msp_runbook.py.
pause
exit /b 0
:fail
echo.
echo The build failed. The messages above say why.
pause
exit /b 1
