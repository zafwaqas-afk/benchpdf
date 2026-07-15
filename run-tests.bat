@echo off
REM ============================================================
REM  Fidelity regression suite — double-click to run.
REM  Asserts the conversion engine's placement invariants on the
REM  committed fixtures (no overlap, no fragmentation, native
REM  tables, zero insets, consistent fonts, golden-layout drift).
REM  No engine change should ship unless this ends GREEN.
REM ============================================================
title Conversion fidelity suite
cd /d "%~dp0"

set "VENV_PY=%~dp0venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo Python environment not found. Double-click "PDF to PowerPoint.bat" once
  echo to set it up, then run this again.
  echo.
  pause
  exit /b 1
)

echo Running the conversion fidelity suite...
echo.
"%VENV_PY%" "%~dp0tests\fidelity_suite.py"
set RC=%errorlevel%
echo.
if %RC%==0 (
  echo ============================================================
  echo   RESULT: GREEN - all fidelity invariants hold.
  echo ============================================================
) else (
  echo ============================================================
  echo   RESULT: RED - a regression was detected. Do not ship.
  echo   Scroll up for the failing checks.
  echo ============================================================
)
echo.
pause
exit /b %RC%
