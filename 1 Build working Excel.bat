@echo off
rem Step 1: builds the working Excel and the required-details file.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run Setup.bat first ^(one time only^).
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m cersai_tool run1 %*
echo.
pause
