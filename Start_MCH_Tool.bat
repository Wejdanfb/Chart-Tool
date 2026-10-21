@echo off
REM Double-click to start the Chart Review Report tool on Windows.
REM First run installs the required packages; later runs just start the tool.
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)

%PY% --version >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.9 or newer from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" during setup, then double-click this file again.
  pause
  exit /b 1
)

%PY% -m pip install -r requirements.txt
if errorlevel 1 (
  echo Could not install the required packages. Check your internet connection and try again.
  pause
  exit /b 1
)

echo.
echo Starting the tool - your browser will open at http://localhost:5050
echo Keep this window open while you use it. Close it (or press Ctrl+C) to stop.
start "" http://localhost:5050
%PY% app.py
pause
