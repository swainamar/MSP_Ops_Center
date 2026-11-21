@echo off
setlocal
title MSP Ops Center - load the demo
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run Start_MSP_Ops_Center.bat once first so the environment is installed.
  pause
  exit /b 1
)
echo This loads DEMO users, teams, customers and data into the data folder in msp_settings.ini.
echo Every demo account gets the same password. Use this only on a demo copy, never on a live installation.
echo.
choice /C YN /M "Load the demo now"
if errorlevel 2 exit /b 0
".venv\Scripts\python.exe" msp_demo.py
echo.
echo If MSP Ops Center is running, refresh the browser page to see the demo.
pause
