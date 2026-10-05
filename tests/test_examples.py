"""The scripts in examples/ (documentation) must keep working."""

import glob
import os
import subprocess
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXAMPLES = sorted(glob.glob(os.path.join(ROOT, "examples", "*.py")))


@pytest.mark.parametrize("script", EXAMPLES, ids=os.path.basename)
def test_example_runs(script, tmp_path):
    args = [sys.executable, script]
    if os.path.basename(script) == "quickstart.py":
        args.append(str(tmp_path))
    res = subprocess.run(args, capture_output=True, text=True, cwd=ROOT, timeout=300)
    assert res.returncode == 0, res.stdout + res.stderr
    assert res.stdout.strip()
