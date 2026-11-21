@echo off
setlocal EnableExtensions
title MSP Ops Center
cd /d "%~dp0"

rem ======================================================================
rem  MSP Ops Center launcher. Double-click to start. Keep the window open.
rem  First run: creates a private Python environment, installs the needed
rem  packages and writes msp_settings.ini (edit it in Notepad).
rem ======================================================================

set "PY="
set "OKVER=import sys; sys.exit(0 if (3, 11) <= sys.version_info[:2] <= (3, 14) else 1)"
rem Ask for a supported Python by version. "py -3" would take the newest one installed, which may not be supported.
for %%V in (3.12 3.13 3.11 3.14) do (
  if not defined PY (
    py -%%V -c "%OKVER%" >nul 2>&1 && set "PY=py -%%V"
  )
)
if not defined PY (
  python -c "%OKVER%" >nul 2>&1 && set "PY=python"
)
if not defined PY (
  where py >nul 2>&1 && goto :oldpython
  where python >nul 2>&1 && goto :oldpython
  goto :nopython
)
%PY% -c "import sys; print('Using Python ' + sys.version.split()[0])"

rem An environment made earlier with an unsupported Python is rebuilt with the one chosen above.
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -c "%OKVER%" >nul 2>&1
  if errorlevel 1 (
    echo The private environment was made with an unsupported Python version. Rebuilding it ...
    rmdir /s /q ".venv"
  )
)
if not exist ".venv\Scripts\python.exe" (
  echo Setting up MSP Ops Center for the first time. This takes a minute...
  %PY% -m venv .venv
  if errorlevel 1 goto :venvfail
)
set "VPY=%~dp0.venv\Scripts\python.exe"

rem Packages are installed only from requirements.txt: exact versions, each file checked against its SHA-256 hash.
rem A tampered package or an unexpected new release is refused (pip --require-hashes).
"%VPY%" -c "import openpyxl, paramiko, cryptography, defusedxml, holidays" >nul 2>&1
if errorlevel 1 (
  if not exist "requirements.txt" goto :noreq
  echo Installing the required packages: exact, verified versions from requirements.txt ...
  "%VPY%" -m pip install --disable-pip-version-check --quiet --require-hashes --only-binary=:all: -r requirements.txt
  if errorlevel 1 goto :pipfail
)
if not exist "msp_roster.py" echo Note: msp_roster.py is not in this folder, so the Shift roster page is switched off.
if not exist "msp_ai.py" echo Note: msp_ai.py is not in this folder, so the Ask assistant is switched off.
if not exist "msp_org.py" echo Note: msp_org.py is not in this folder, so the Teams page is switched off.
if not exist "msp_access.py" echo WARNING: msp_access.py is missing, so team data is NOT isolated. Copy it next to msp_runbook.py.
if not exist "static\vendor\xterm.js" echo WARNING: the static folder is missing. Copy it next to msp_runbook.py, or fonts and terminals will not load.

if not exist "msp_settings.ini" (
  "%VPY%" msp_runbook.py --write-settings-template msp_settings.ini >nul
  echo Created msp_settings.ini. Edit it in Notepad to change the port, storage or sign-in options.
)

echo.
echo Starting MSP Ops Center. Keep this window open while people use the portal.
echo To stop it, close this window or press Ctrl+C.
echo.
"%VPY%" msp_runbook.py --config "%~dp0msp_settings.ini" --open-browser
echo.
echo MSP Ops Center has stopped.
pause
exit /b 0

:nopython
echo Python 3.10 or newer is not installed.
echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH",
echo then double-click this file again.
start "" "https://www.python.org/downloads/"
pause
exit /b 1

:oldpython
echo MSP Ops Center needs Python 3.11, 3.12, 3.13 or 3.14, and none of these was found.
echo Python 3.10 and older no longer receive security fixes. Python 3.12 is recommended.
echo Install Python 3.12 from https://www.python.org/downloads/ and tick "Add Python to PATH".
pause
exit /b 1

:venvfail
echo Could not create the private Python environment in "%~dp0.venv".
echo Check that this folder is not read-only, then try again.
pause
exit /b 1

:noreq
echo requirements.txt is missing. Copy it next to msp_runbook.py; it lists the exact, verified package versions.
pause
exit /b 1

:pipfail
echo Package installation failed.
echo  - If pip says the packages DO NOT MATCH THE HASHES: look at the file name it shows. A "cp3XX" part for a
echo    Python version other than 3.11 to 3.14 means an unsupported Python; install Python 3.12 and run this again.
echo    Otherwise something changed the download: do not continue, and tell your administrator.
echo  - If your network uses a proxy, set it first, for example:
echo     set HTTPS_PROXY=http://proxy.company.local:8080
echo and run this file again from that same window.
pause
exit /b 1
