@echo off
REM ============================================================
REM  PDF to PowerPoint - local, private converter
REM  Double-click this file to start the tool. The browser opens
REM  automatically. Close the black window to stop the tool.
REM ============================================================
title PDF to PowerPoint
cd /d "%~dp0"

set "VENV_PY=%~dp0venv\Scripts\python.exe"

REM ---- First run: create the virtual environment if missing ----
if not exist "%VENV_PY%" (
  echo.
  echo First-time setup - this happens only once and takes a minute...
  echo Creating a private Python environment...
  where py >nul 2>nul
  if %errorlevel%==0 (
    py -m venv "%~dp0venv"
  ) else (
    python -m venv "%~dp0venv"
  )
  if not exist "%VENV_PY%" (
    echo.
    echo ERROR: Could not create the Python environment.
    echo Please install Python 3 from https://www.python.org/downloads/
    echo and be sure to tick "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
  )
  echo Installing components ^(no files leave your PC^)...
  "%VENV_PY%" -m pip install --upgrade pip >nul
  "%VENV_PY%" -m pip install -r "%~dp0requirements.txt"
  if %errorlevel% neq 0 (
    echo.
    echo ERROR: Setup failed while installing components.
    pause
    exit /b 1
  )
  echo Setup complete.
  echo.
)

echo Starting PDF to PowerPoint...
echo A browser tab will open at http://127.0.0.1:5000/
echo Keep this window open while you work. Close it to stop.
echo.
"%VENV_PY%" "%~dp0app\server.py"

if %errorlevel% neq 0 (
  echo.
  echo The tool stopped unexpectedly. See the message above.
  pause
)
