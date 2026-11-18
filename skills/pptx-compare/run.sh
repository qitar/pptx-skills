#!/usr/bin/env bash
set -o errexit
set -o errtrace
set -o pipefail
trap 'echo "Error on line $LINENO. Exit code: $?" >&2' ERR

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$HERE/.venv"

find_python() {
  for cmd in python3.13 python3.12 python3.11 python3.10 python3 python; do
    if command -v "$cmd" >/dev/null 2>&1; then
      if "$cmd" -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" 2>/dev/null; then
        echo "$cmd"
        return 0
      fi
    fi
  done
  return 1
}

if [[ ! -d "$VENV" ]]; then
  PYTHON_CMD=$(find_python) || {
    echo "Error: Python >= 3.10 is required but not found" >&2
    exit 1
  }
  "$PYTHON_CMD" -m venv "$VENV"
fi

if [[ "$OSTYPE" == "msys" || "$OSTYPE" == "cygwin" ]]; then
  PIP="$VENV/Scripts/pip"
  PYTHON="$VENV/Scripts/python"
else
  PIP="$VENV/bin/pip"
  PYTHON="$VENV/bin/python"
fi

"$PIP" install -r "$HERE/requirements.txt" -q

STRATEGY="${1:?Usage: run.sh <lite|full> <before.pptx> <after.pptx> [options]}"
shift

if [[ "$STRATEGY" == "lite" ]]; then
  PYTHONPATH="$HERE" "$PYTHON" "$HERE/slide-compare-lite/scripts/main.py" "$@"
elif [[ "$STRATEGY" == "full" ]]; then
  PYTHONPATH="$HERE" "$PYTHON" "$HERE/slide-compare/scripts/main.py" "$@"
else
  echo "Error: Unknown strategy '$STRATEGY'. Use 'lite' or 'full'." >&2
  exit 1
fi
