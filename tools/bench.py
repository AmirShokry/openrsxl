"""
Benchmark one operation of openpyxl or openrsxl in a fresh process.

    python tools/bench.py <lib> <file.xlsx> [--read-only] [--save] [--values] [--cells]

Prints a dict with timings (seconds) and memory (MB):

* ``base_mb``: resident set size after importing psutil, before importing
  the library;
* ``peak_rss_mb``: the process' peak resident set size (Windows: peak
  working set, elsewhere ``ru_maxrss``);
* ``peak_mb``: ``peak_rss_mb - base_mb``, the memory the operation needed.

Saving writes to a temporary file on disk.
"""

import argparse
import importlib
import os
import sys
import tempfile
import time
import warnings

import psutil


def peak_rss():
    """Peak resident set size of this process, in bytes."""
    mi = psutil.Process().memory_info()
    peak = getattr(mi, "peak_wset", None)
    if peak:
        return peak
    import resource

    kb = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return kb if sys.platform == "darwin" else kb * 1024  # macOS reports bytes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("lib")
    ap.add_argument("path")
    ap.add_argument("--read-only", action="store_true")
    ap.add_argument("--save", action="store_true")
    ap.add_argument("--values", action="store_true", help="iterate all values")
    ap.add_argument("--cells", action="store_true", help="iterate all cells")
    a = ap.parse_args()
    warnings.simplefilter("ignore")

    base = psutil.Process().memory_info().rss
    lib = importlib.import_module(a.lib)
    t0 = time.perf_counter()
    wb = lib.load_workbook(a.path, read_only=a.read_only)
    out = {"lib": a.lib, "load_s": round(time.perf_counter() - t0, 3)}
    if a.values or (a.read_only and not a.cells):
        t0 = time.perf_counter()
        n = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                n += len(row)
        out["values_s"] = round(time.perf_counter() - t0, 3)
        out["values"] = n
    if a.cells:
        t0 = time.perf_counter()
        n = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        n += 1
        out["cells_s"] = round(time.perf_counter() - t0, 3)
    if a.save and not a.read_only:
        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        t0 = time.perf_counter()
        wb.save(path)
        out["save_s"] = round(time.perf_counter() - t0, 3)
        os.remove(path)
    peak = peak_rss()
    out["base_mb"] = round(base / 2**20, 1)
    out["peak_rss_mb"] = round(peak / 2**20, 1)
    out["peak_mb"] = round((peak - base) / 2**20, 1)
    print(out)


if __name__ == "__main__":
    main()
