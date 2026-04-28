@echo off
setlocal
cd /d "%~dp0"

where fnm >nul 2>nul
if %errorlevel%==0 (
  call fnm env --use-on-cd > "%TEMP%\fnm_env.cmd"
  call "%TEMP%\fnm_env.cmd"
  fnm use 24
)

echo Building frontend...
cd web
call npm install
if errorlevel 1 exit /b 1
call npm run build
if errorlevel 1 exit /b 1

echo Testing backend...
cd ..
go mod tidy
if errorlevel 1 exit /b 1
go test ./...
if errorlevel 1 exit /b 1

echo Building Windows executable...
if exist output rmdir /s /q output
mkdir "output\表格转换工具"
set GOOS=windows
set GOARCH=amd64
go build -o "output\表格转换工具\app.exe" .
if errorlevel 1 exit /b 1
copy /Y start.bat "output\表格转换工具\start.bat" >nul

echo Done: output\表格转换工具
