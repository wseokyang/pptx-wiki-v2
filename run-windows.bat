@echo off
setlocal
if "%~1"=="" (
  echo Usage: drag a .pptx file onto run-windows.bat
  echo    or: run-windows.bat "C:\path\to\deck.pptx"
  exit /b 2
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-windows.ps1" -PptxPath "%~f1"
exit /b %ERRORLEVEL%
