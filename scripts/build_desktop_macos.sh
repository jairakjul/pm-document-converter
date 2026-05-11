#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PYTHON_EXE="$PROJECT_ROOT/.venv/bin/python"
SPEC_PATH="$PROJECT_ROOT/packaging/PMDocumentConverter_mac.spec"

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "This script must be run on macOS to build a .app bundle." >&2
  exit 1
fi

if [[ ! -x "$PYTHON_EXE" ]]; then
  echo "Python virtual environment not found: $PYTHON_EXE" >&2
  echo "Create it with: python3 -m venv .venv && .venv/bin/python -m pip install -r requirements.txt" >&2
  exit 1
fi

if [[ ! -f "$SPEC_PATH" ]]; then
  echo "PyInstaller spec not found: $SPEC_PATH" >&2
  exit 1
fi

cd "$PROJECT_ROOT"
rm -rf build dist
"$PYTHON_EXE" -m PyInstaller --noconfirm --clean "$SPEC_PATH"

APP_PATH="$PROJECT_ROOT/dist/PMDocumentConverter.app"
if [[ ! -d "$APP_PATH" ]]; then
  echo "Build finished but app bundle was not found: $APP_PATH" >&2
  exit 1
fi

echo "Build completed: $APP_PATH"
