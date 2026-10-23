@echo off
rem Brightness / saturation / pixel-diff report
rem A = first group (training images or no-LoRA output), B = second group (LoRA output)
setlocal
set "PS1=%~dp0StyleCompare.ps1"
if not exist "%PS1%" (
  echo [ERROR] StyleCompare.ps1 not found next to this .bat
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
echo.
echo (press any key to close)
pause >nul
