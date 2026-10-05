"""
Benchmark openrsxl.extended streaming features in a fresh process.

    python tools/bench_extended.py <file.xlsx> [flag ...] [--cells] [--lib openpyxl]

flags: names of openrsxl.extended.load_workbook options (formula_and_value,
read_comments, ...) or "all". Without flags: plain read-only mode.
Prints load / iteration time, the process' peak RSS (``peak_rss_mb``) and its
increase over the baseline taken before importing the library (``peak_mb``).
"""

import argparse
import importlib
import os
import sys
import time
import warnings

import psutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bench import peak_rss  # noqa: E402

ALL = [
    "formula_and_value",
    "read_comments",
    "read_hyperlinks",
    "read_merged_cells",
    "read_dimensions",
    "read_sheet_properties",
    "read_data_validations",
    "read_conditional_formatting",
    "read_tables",
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("flags", nargs="*")
    ap.add_argument("--cells", action="store_true", help="iterate cells (default: values)")
    ap.add_argument("--lib", default="openrsxl.extended")
    a = ap.parse_args()
    warnings.simplefilter("ignore")
    flags = ALL if "all" in a.flags else a.flags
    base = psutil.Process().memory_info().rss
    lib = importlib.import_module(a.lib)
    kw = {f: True for f in flags}
    t0 = time.perf_counter()
    wb = lib.load_workbook(a.path, read_only=True, **kw)
    out = {"lib": a.lib, "flags": ",".join(flags) or "-", "load_s": round(time.perf_counter() - t0, 3)}
    t0 = time.perf_counter()
    n = 0
    for ws in wb.worksheets:
        if a.cells:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        n += 1
                    if "formula_and_value" in flags:
                        _ = c.formula, c.cached_value
        else:
            for row in ws.iter_rows(values_only=True):
                n += len(row)
    out["iter_s"] = round(time.perf_counter() - t0, 3)
    out["n"] = n
    peak = peak_rss()
    out["peak_rss_mb"] = round(peak / 2**20, 1)
    out["peak_mb"] = round((peak - base) / 2**20, 1)
    print(out)


if __name__ == "__main__":
    main()
