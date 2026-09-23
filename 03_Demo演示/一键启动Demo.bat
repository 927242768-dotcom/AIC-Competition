@echo off
setlocal
cd /d "%~dp0"
for %%F in (*.exe) do (
  start "" "%%~fF"
  exit /b 0
)
echo [ERROR] Desktop application EXE was not found.
pause
exit /b 1
