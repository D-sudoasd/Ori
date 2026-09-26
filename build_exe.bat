@echo off
setlocal
cd /d "%~dp0"
py -3 -m pip install -r requirements-build.txt
if errorlevel 1 (
  echo Failed to install PyInstaller.
  exit /b 1
)
py -3 -m PyInstaller --noconfirm --clean SpectraToOrigin.spec
if errorlevel 1 (
  echo PyInstaller failed.
  exit /b 1
)
echo Built dist\SpectraToOrigin.exe
echo Built dist\DataToOriginCLI.exe
