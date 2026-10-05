"""
Benchmark creating a workbook from scratch.

    python tools/bench_write.py <lib> [--write-only] [--rows N]
"""

import argparse
import datetime
import importlib
import os
import sys
import tempfile
import time

import psutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bench import peak_rss  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("lib")
    ap.add_argument("--write-only", action="store_true")
    ap.add_argument("--rows", type=int, default=200_000)
    a = ap.parse_args()
    proc = psutil.Process()
    base = proc.memory_info().rss
    lib = importlib.import_module(a.lib)
    t0 = time.perf_counter()
    wb = lib.Workbook(write_only=a.write_only)
    ws = wb.create_sheet("data") if a.write_only else wb.active
    d = datetime.datetime(2020, 1, 1)
    for r in range(a.rows):
        ws.append([r, r * 1.5, f"name {r % 1000}", "constant", r % 2 == 0, d, None, f"=A{r + 1}*2", -r, 0.1 * r])
    t_fill = time.perf_counter() - t0
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    t0 = time.perf_counter()
    wb.save(path)
    t_save = time.perf_counter() - t0
    os.remove(path)
    peak = peak_rss()
    print(
        {
            "lib": a.lib,
            "write_only": a.write_only,
            "fill_s": round(t_fill, 2),
            "save_s": round(t_save, 2),
            "peak_rss_mb": round(peak / 2**20, 1),
            "peak_mb": round((peak - base) / 2**20, 1),
        }
    )


if __name__ == "__main__":
    main()
