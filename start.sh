#!/usr/bin/env sh
set -e

DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$DIR"

PORT=${PORT:-8080}

if [ -x "./app" ]; then
  APP="./app"
elif [ -x "./app.exe" ]; then
  APP="./app.exe"
else
  echo "app binary not found. Run scripts/build.sh first."
  exit 1
fi

PORT=$PORT "$APP"
