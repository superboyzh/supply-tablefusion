#!/usr/bin/env sh
set -e

ROOT=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
cd "$ROOT"

if command -v fnm >/dev/null 2>&1; then
  eval "$(fnm env)"
  fnm use 24
fi

cd web
npm install
npm run build

cd "$ROOT"
go mod tidy
go build -o app .
