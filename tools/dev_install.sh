#!/usr/bin/env bash
# Build and install openrsxl into ./.venv (editable). On Windows a loaded
# extension DLL cannot be overwritten but it can be renamed, so move it away.
set -e
cd "$(dirname "$0")/.."
for f in python/openrsxl/_openrsxl*.pyd; do
  [ -e "$f" ] && mv "$f" "${TMPDIR:-${TEMP:-/tmp}}/openrsxl_old_$(date +%s%N).pyd" || true
done
VIRTUAL_ENV="$(pwd)/.venv" maturin develop --release "$@" 2>&1 | tail -1
