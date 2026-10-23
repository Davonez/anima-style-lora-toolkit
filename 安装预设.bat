@echo off
rem Install the bundled Next-Trainer preset into <Next-Trainer>\config\presets
setlocal
set "PS1=%~dp0InstallPreset.ps1"
if not exist "%PS1%" (
  echo [ERROR] InstallPreset.ps1 not found next to this .bat
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
echo.
echo (press any key to close)
pause >nul
