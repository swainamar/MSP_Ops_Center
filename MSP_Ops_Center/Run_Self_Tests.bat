@echo off
setlocal
title MSP Ops Center - self test
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run Start_MSP_Ops_Center.bat once first so the environment is installed.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m unittest -v test_msp_runbook.py
if exist "test_msp_roster.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_roster.py
if exist "test_msp_ai.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_ai.py
if exist "test_msp_org.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_org.py
if exist "test_msp_demo.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_demo.py
if exist "test_msp_access.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_access.py
if exist "test_msp_ai_org.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_ai_org.py
if exist "test_msp_security.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_security.py
if exist "test_msp_ai_eval.py" ".venv\Scripts\python.exe" -m unittest -v test_msp_ai_eval.py
pause
