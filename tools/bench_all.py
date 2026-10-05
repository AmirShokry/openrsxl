"""
Run the benchmark matrix (each configuration in a fresh process, one at a
time) and print a markdown table.

    python tools/bench_all.py file1.xlsx [file2.xlsx ...]

Columns: total time of the operation (load + iteration / save), the
speed-up, and memory as "peak RSS (+increase over the process baseline)".
"""

import ast
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIGS = [
    ("load + save", ["--save"]),
    ("load + iterate values", ["--values"]),
    ("load + iterate cells", ["--cells"]),
    ("read-only: iterate values", ["--read-only"]),
    ("read-only: iterate cells", ["--read-only", "--cells"]),
]


def run(lib, path, args):
    out = subprocess.run(
        [sys.executable, os.path.join(HERE, "bench.py"), lib, path] + args, capture_output=True, text=True, check=True
    ).stdout
    line = [x for x in out.splitlines() if x.startswith("{")][-1]
    return ast.literal_eval(line)


def total(r):
    return sum(v for k, v in r.items() if k.endswith("_s"))


def fmt_time(r):
    parts = [f"{r['load_s']:.2f}"]
    for k in ("values_s", "cells_s", "save_s"):
        if k in r:
            parts.append(f"{r[k]:.2f}")
    return f"{total(r):.2f} s" + (f" ({' + '.join(parts)})" if len(parts) > 1 else "")


def fmt_mem(r):
    return f"{r['peak_rss_mb']:.0f} MB (+{r['peak_mb']:.0f})"


def main():
    print(
        "| file | operation | openpyxl time | openrsxl time | speed-up | openpyxl peak RSS (+increase) | openrsxl peak RSS (+increase) |"
    )
    print("|---|---|---|---|---|---|---|")
    for path in sys.argv[1:]:
        for label, args in CONFIGS:
            a = run("openpyxl", path, args)
            b = run("openrsxl", path, args)
            print(
                f"| {os.path.basename(path)} | {label} | {fmt_time(a)} | {fmt_time(b)} | "
                f"{total(a) / total(b):.1f}x | {fmt_mem(a)} | {fmt_mem(b)} |",
                flush=True,
            )


if __name__ == "__main__":
    main()
