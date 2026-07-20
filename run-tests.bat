@echo off
REM ============================================================
REM  Fidelity regression suite — double-click to run.
REM  Runs the SAME fixtures and the SAME assertions against EVERY
REM  registered engine (tests/engines.py). An engine that is not in
REM  that registry may not be linked from the site.
REM
REM  Asserts each engine's placement invariants on the
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
"%VENV_PY%" "%~dp0tests\fidelity_suite.py" %*
set RC=%errorlevel%
echo.
if %RC%==0 (
  echo ============================================================
  echo   RESULT: GREEN - every SHIPPING engine holds every invariant.
  echo ============================================================
) else (
  echo ============================================================
  echo   RESULT: RED - a SHIPPING engine regressed. Do not ship.
  echo   Scroll up for the failing checks.
  echo ============================================================
)
echo.
pause
exit /b %RC%
