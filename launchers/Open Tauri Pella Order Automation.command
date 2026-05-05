#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
APP_ROOT="${SCRIPT_DIR:h}"
cd "$APP_ROOT"

if ! command -v npm >/dev/null 2>&1; then
  echo "npm was not found. Please install Node.js, then try again."
  read -r "?Press Return to close."
  exit 1
fi

if ! command -v cargo >/dev/null 2>&1; then
  echo "Cargo was not found. Please install Rust from https://rustup.rs/, then try again."
  read -r "?Press Return to close."
  exit 1
fi

if [[ ! -d "$APP_ROOT/node_modules" ]]; then
  echo "Installing app dependencies..."
  npm install
fi

export PYTHONPATH="$APP_ROOT/src"
npm run desktop:dev
