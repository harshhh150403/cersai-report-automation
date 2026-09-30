@echo off
rem Step 2: applies the required-details file and makes the final Excel.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run Setup.bat first ^(one time only^).
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m cersai_tool run2 %*
echo.
pause
