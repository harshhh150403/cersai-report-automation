@echo off
rem Step 3: checks the final Excel and writes the .DAT file.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run Setup.bat first ^(one time only^).
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m cersai_tool run3 %*
echo.
pause
