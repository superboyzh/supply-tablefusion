@echo off
setlocal
cd /d "%~dp0\.."

where fnm >nul 2>nul
if %errorlevel%==0 (
  call fnm env --use-on-cd > "%TEMP%\fnm_env.cmd"
  call "%TEMP%\fnm_env.cmd"
  fnm use 24
)

cd web
call npm install
if errorlevel 1 exit /b 1
call npm run build
if errorlevel 1 exit /b 1

cd ..
go mod tidy
if errorlevel 1 exit /b 1
set GOOS=windows
set GOARCH=amd64
go build -o app.exe .
