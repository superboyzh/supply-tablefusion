#!/usr/bin/env sh
set -e

ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$ROOT"

if command -v fnm >/dev/null 2>&1; then
  eval "$(fnm env)"
  fnm use 24
fi

echo "Building frontend..."
cd web
npm install
npm run build

echo "Testing backend..."
cd "$ROOT"
go mod tidy
go test ./...

echo "Building Windows executable..."
rm -rf output
mkdir -p "output/表格转换工具"
GOOS=windows GOARCH=amd64 go build -o "output/表格转换工具/app.exe" .
cp start.bat "output/表格转换工具/start.bat"

echo "Done: output/表格转换工具"
