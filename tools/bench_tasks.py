"""
Benchmark the same *task* done the openpyxl way and the openrsxl.extended way
(each measurement in a fresh process, one at a time).

    python tools/bench_tasks.py <file.xlsx> [file2.xlsx ...]        # markdown table
    python tools/bench_tasks.py --one <task> <impl> <file.xlsx>     # one measurement

Tasks:

* ``formulas_and_values``: the formula *and* the cached (saved) value of every
  cell. openpyxl can only do this in full mode, loading the workbook twice
  (``data_only=False`` and ``data_only=True``); openrsxl.extended streams it
  in one read-only pass (``formula_and_value=True``).
* ``merged_comments_links``: merged ranges, and the comment and hyperlink of
  every cell. openpyxl: full mode; openrsxl.extended: streaming with
  ``read_merged_cells``, ``read_comments``, ``read_hyperlinks``.

Implementations: ``openpyxl`` (full mode), ``openrsxl`` (the same code, full
mode, drop-in replacement) and ``extended`` (openrsxl.extended streaming).
Memory: peak RSS of the process and its increase over the baseline taken
before importing the library.
"""

import ast
import os
import subprocess
import sys
import time
import warnings

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from bench import peak_rss  # noqa: E402

TASKS = ("formulas_and_values", "merged_comments_links")
IMPLS = ("openpyxl", "openrsxl", "extended")


def full_formulas_and_values(lib, path):
    wf = lib.load_workbook(path)
    wv = lib.load_workbook(path, data_only=True)
    n = f = 0
    for sf, sv in zip(wf.worksheets, wv.worksheets):
        for rf, rv in zip(sf.iter_rows(), sv.iter_rows()):
            for cf, cv in zip(rf, rv):
                formula = cf.value if cf.data_type == "f" else None
                _value = cv.value  # (read, as the other implementations do)
                n += 1
                f += formula is not None
    return n, f


def ext_formulas_and_values(path):
    import openrsxl.extended as ext

    wb = ext.load_workbook(path, read_only=True, formula_and_value=True)
    n = f = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                formula, _value = c.formula, c.cached_value
                n += 1
                f += formula is not None
    return n, f


def full_features(lib, path):
    wb = lib.load_workbook(path)
    n = k = 0
    for ws in wb.worksheets:
        k += len(ws.merged_cells.ranges)
        for row in ws.iter_rows():
            for c in row:
                comment = getattr(c, "comment", None)
                link = getattr(c, "hyperlink", None)
                n += (comment is not None) + (link is not None)
    return n, k


def ext_features(path):
    import openrsxl.extended as ext

    wb = ext.load_workbook(path, read_only=True, read_merged_cells=True, read_comments=True, read_hyperlinks=True)
    n = k = 0
    for ws in wb.worksheets:
        k += len(ws.merged_cells.ranges)
        for row in ws.iter_rows():
            for c in row:
                comment = getattr(c, "comment", None)
                link = getattr(c, "hyperlink", None)
                n += (comment is not None) + (link is not None)
    return n, k


def one(task, impl, path):
    import psutil

    warnings.simplefilter("ignore")
    base = psutil.Process().memory_info().rss
    t0 = time.perf_counter()
    if impl == "extended":
        res = ext_formulas_and_values(path) if task == TASKS[0] else ext_features(path)
    else:
        lib = __import__(impl)
        res = full_formulas_and_values(lib, path) if task == TASKS[0] else full_features(lib, path)
    dt = time.perf_counter() - t0
    peak = peak_rss()
    print(
        {
            "task": task,
            "impl": impl,
            "time_s": round(dt, 3),
            "result": res,
            "peak_rss_mb": round(peak / 2**20, 1),
            "peak_mb": round((peak - base) / 2**20, 1),
        }
    )


def run(task, impl, path):
    out = subprocess.run(
        [sys.executable, os.path.abspath(__file__), "--one", task, impl, path],
        capture_output=True,
        text=True,
        check=True,
    ).stdout
    return ast.literal_eval([x for x in out.splitlines() if x.startswith("{")][-1])


def main(argv):
    if argv[:1] == ["--one"]:
        one(*argv[1:4])
        return
    print("| file | task | openpyxl (full mode) | openrsxl (full mode) | openrsxl.extended (streaming) |")
    print("|---|---|---|---|---|")
    for path in argv:
        for task in TASKS:
            cells = []
            results = set()
            for impl in IMPLS:
                r = run(task, impl, path)
                results.add(r["result"])
                cells.append(f"{r['time_s']:.2f} s, {r['peak_rss_mb']:.0f} MB (+{r['peak_mb']:.0f})")
            note = "" if len(results) == 1 else f" (results differ: {results})"
            print(f"| {os.path.basename(path)} | {task}{note} | " + " | ".join(cells) + " |", flush=True)


if __name__ == "__main__":
    main(sys.argv[1:])
