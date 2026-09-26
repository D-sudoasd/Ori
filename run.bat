@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 (
  echo Failed to enter script folder.
  pause
  exit /b 1
)
where py >nul 2>&1
if errorlevel 1 (
  echo Failed to start. Need Python launcher: py -3
  pause
  exit /b 1
)
py -3 spectra_to_origin.py %*
if errorlevel 1 (
  echo SpectraToOrigin failed. See the error above.
  pause
  exit /b 1
)
