@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  echo Please run: py -m venv .venv
  echo Then run: .venv\Scripts\python -m pip install -r requirements.txt
  pause
  exit /b 1
)
.venv\Scripts\python.exe reconcile.py
pause
