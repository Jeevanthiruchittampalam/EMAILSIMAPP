@echo off
REM Double-click to run. Sets up local environment the first time.
setlocal enabledelayedexpansion
cd /d "%~dp0"

set APP_FILE=email_filter_app.py
if not exist "%APP_FILE%" set APP_FILE=email-filter-app.py

if not exist .venv (
  echo Creating local environment...
  py -3 -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip >NUL
pip install -r requirements.txt >NUL

echo Starting Email Filter app...
python "%APP_FILE%"
pause
