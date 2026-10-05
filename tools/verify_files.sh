#!/usr/bin/env bash
# Differential verification of whole workbooks (digest based, see digest.py):
# full mode state, read-only state, save after load (lazy) and save after
# full access.  Usage: tools/verify_files.sh <workdir> file.xlsx...
set -u
cd "$(dirname "$0")/.."
WORK="$1"; shift
PY=.venv/Scripts/python
status=0
for f in "$@"; do
  D="$WORK/$(basename "$f" .xlsx)"
  rm -rf "$D"
  for L in openpyxl openrsxl; do
    $PY tools/digest.py $L "$f" "$D" 2>/dev/null
    $PY tools/digest.py $L "$f" "$D" --read-only 2>/dev/null
  done
  $PY tools/digest.py openrsxl "$f" "$D" --save-only 2>/dev/null
  echo "== $f"
  $PY tools/digest.py --compare "$D" | tail -1 || status=1
done
exit $status
