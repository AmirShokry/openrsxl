"""
Port openpyxl's upstream test-suite so that it exercises openrsxl.

Usage::

    python tools/port_upstream_tests.py <path-to-openpyxl-source-checkout>

The upstream tests live inside the openpyxl package (``openpyxl/*/tests``) and
use package-relative imports.  This script copies every ``tests`` directory
(including fixture data) into ``tests/upstream`` keeping the directory layout,
and rewrites:

* ``openpyxl.<module>`` references -> ``openrsxl.<module>``
* ``import openpyxl`` / ``from openpyxl import`` -> openrsxl
* package-relative imports -> absolute ``openrsxl`` imports
* ``openpyxl.tests.helper`` -> ``upstream_helper``
* test modules importing other test modules -> ``upstream.<path>``

Strings that merely *mention* openpyxl (eg. the default ``creator`` written
into docProps/core.xml) are left untouched, because openrsxl reproduces them
verbatim.
"""

import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEST = os.path.join(ROOT, "tests", "upstream")

MODULE_REF = re.compile(r"\bopenpyxl(?=\.[A-Za-z_])")
IMPORT_STMT = re.compile(r"^(\s*)import openpyxl\b", re.M)
FROM_STMT = re.compile(r"^(\s*)from openpyxl import\b", re.M)
REL_IMPORT = re.compile(r"^(\s*)from (\.+)\s*([\w.]*)\s+import", re.M)
TEST_MODULE_REF = re.compile(r"\bopenrsxl((?:\.\w+)*\.tests\.\w+)")


def rewrite(text, parts):
    text = text.replace("openpyxl.tests.helper", "upstream_helper")

    def rel(m):
        indent, dots, mod = m.groups()
        level = len(dots)
        base = parts[: len(parts) - (level - 1)] if level > 1 else parts
        target = ".".join(["openrsxl"] + base + ([mod] if mod else []))
        return f"{indent}from {target} import "

    text = REL_IMPORT.sub(rel, text)
    text = MODULE_REF.sub("openrsxl", text)
    text = TEST_MODULE_REF.sub(r"upstream\1", text)
    # reprs of test classes: pytest's importlib mode names test modules
    # relative to tests/upstream
    text = text.replace("<upstream.", "<")
    text = IMPORT_STMT.sub(r"\1import openrsxl; import openrsxl as openpyxl", text)
    text = FROM_STMT.sub(r"\1from openrsxl import", text)
    return text


def main(src):
    pkg = os.path.join(src, "openpyxl")
    keep = {"conftest.py", "pytest.ini"}  # hand-written, not generated
    if os.path.exists(DEST):
        for name in os.listdir(DEST):
            if name in keep:
                continue
            path = os.path.join(DEST, name)
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
    for dirpath, dirnames, filenames in os.walk(pkg):
        dirnames[:] = [d for d in dirnames if d != "__pycache__"]
        rel = os.path.relpath(dirpath, pkg)
        parts = [] if rel == "." else rel.replace("\\", "/").split("/")
        if "tests" not in parts:
            continue
        out_dir = os.path.join(DEST, rel)
        os.makedirs(out_dir, exist_ok=True)
        for fn in filenames:
            s = os.path.join(dirpath, fn)
            d = os.path.join(out_dir, fn)
            if fn == "__init__.py":
                continue  # importlib import mode; no packages needed
            if fn.endswith(".py"):
                with open(s, encoding="utf-8") as f:
                    text = f.read()
                if fn == "helper.py" and parts == ["tests"]:
                    d = os.path.join(DEST, "upstream_helper.py")
                with open(d, "w", encoding="utf-8", newline="") as f:
                    f.write(rewrite(text, parts))
            else:
                shutil.copy2(s, d)


if __name__ == "__main__":
    main(sys.argv[1])
