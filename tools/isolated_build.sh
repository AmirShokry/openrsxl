#!/usr/bin/env bash
# Build the extension and assemble an isolated copy of the package in $1
# (use with PYTHONPATH=$1) without touching the editable install.
set -e
cd "$(dirname "$0")/.."
OUT="$1"
PYO3_PYTHON="$(pwd)/.venv/Scripts/python.exe" cargo build --release $CARGO_FLAGS 2>&1 | grep -E "^error" -A 7 || true
rm -rf "$OUT/openrsxl"
mkdir -p "$OUT"
cp -r python/openrsxl "$OUT/openrsxl"
rm -f "$OUT"/openrsxl/_openrsxl*.pyd
cp target/release/_openrsxl.dll "$OUT/openrsxl/_openrsxl.cp314-win_amd64.pyd"
