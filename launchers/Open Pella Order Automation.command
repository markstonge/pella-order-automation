#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
APP_ROOT="${SCRIPT_DIR:h}"
cd "$APP_ROOT"

if [[ -n "$PELLA_PYTHON" && -x "$PELLA_PYTHON" ]]; then
  PY="$PELLA_PYTHON"
elif [[ -x "/Users/mark/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3" ]]; then
  PY="/Users/mark/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
else
  PY="$(command -v python3)"
fi

export PYTHONPATH="$APP_ROOT/src"
"$PY" -m pella_order_automation.gui
