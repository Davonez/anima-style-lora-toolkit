@echo off
rem Preview only: never writes any .txt
setlocal
set "PS1=%~dp0AnimaTagCleaner.ps1"
set "TGT=%~1"
if not exist "%PS1%" (
  echo [ERROR] AnimaTagCleaner.ps1 not found next to this .bat
  pause
  exit /b 1
)
if "%TGT%"=="" goto pick
if "%TGT:~-1%"=="\" set "TGT=%TGT:~0,-1%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Folder "%TGT%" -DryRun
goto end
:pick
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -DryRun
:end
echo.
echo (press any key to close)
pause >nul
