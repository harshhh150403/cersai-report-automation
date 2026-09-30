@echo off
rem One-time setup: creates the .venv folder with Python and the packages the tool needs.
cd /d "%~dp0"
echo Setting up the CERSAI tool...
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 -m venv .venv
) else (
  python -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
  echo.
  echo Python 3 was not found. Install it from https://www.python.org/downloads/
  echo ^(tick "Add python.exe to PATH"^), then run Setup.bat again.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Installing the packages failed. Check the internet connection and run Setup.bat again.
  pause
  exit /b 1
)
echo.
echo Done. From now on use the three numbered .bat files.
pause
