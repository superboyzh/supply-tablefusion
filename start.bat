@echo off
set PORT=18080

if exist app.exe (
  app.exe
) else (
  echo app.exe not found. Run scripts\build-windows.bat first.
  exit /b 1
)
