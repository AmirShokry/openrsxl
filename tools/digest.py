"""
Memory and disk friendly differential check for very large workbooks.

    python tools/digest.py <lib> <file.xlsx> <outdir> [--read-only] [--data-only]

Loads the workbook with one library and writes a digest of the complete
observable state (same normalisation as tests/oracle/compare.py):

* workbook and worksheet level state, verbatim;
* one line per row: row number, number of cells and the SHA-1 of the full
  state of every cell in the row (style properties are referenced through a
  table written once at the end, so they are fully compared too);
* the workbook saved (with fixed timestamps) to <outdir>/<lib>.xlsx.

Run once per library (each in its own process, so peak memory stays that of
one library) and compare with:

    python tools/digest.py --compare <outdir>
"""

import hashlib
import importlib
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tests"))

from oracle.compare import FIXED, cell_state, load, sheet_state, workbook_state  # noqa: E402

STYLE_KEYS = ("number_format", "font", "fill", "border", "alignment", "protection")
NUL = bytes([0])


def compact_state(c, styles):
    """cell state with the style part replaced by a reference into `styles`"""
    st = cell_state(c)
    if "font" in st:
        sty = tuple(repr(st.pop(k)) for k in STYLE_KEYS)
        st["style_ref"] = styles.setdefault(sty, len(styles))
    return repr(st)


def row_digest(cells, styles):
    h = hashlib.sha1()
    for c in cells:
        s = compact_state(c, styles) if hasattr(c, "coordinate") else repr(c)
        h.update(s.encode("utf-8", "surrogatepass"))
        h.update(NUL)
    return h.hexdigest()


def dump(libname, path, outdir, read_only=False, data_only=False):
    lib = importlib.import_module(libname)
    os.makedirs(outdir, exist_ok=True)
    tag = libname + ("-ro" if read_only else "") + ("-do" if data_only else "")
    wb, err, msgs = load(lib, path, read_only=read_only, data_only=data_only)
    styles = {}
    with open(os.path.join(outdir, tag + ".digest"), "w", encoding="utf-8") as out:
        out.write(repr(("load", err, msgs)) + "\n")
        if wb is None:
            return
        out.write(repr(("workbook", workbook_state(wb))) + "\n")
        for ws in wb.worksheets:
            out.write(f"sheet {ws.title!r}\n")
            if read_only:
                for i, row in enumerate(ws.iter_rows()):
                    out.write(f"{i} {len(row)} {row_digest(row, styles)}\n")
                continue
            st = sheet_state(ws)
            st.pop("cells")
            out.write(repr(("sheet", st)) + "\n")
            rows = {}
            for (r, col), c in ws._cells.items():
                rows.setdefault(r, []).append(c)
            for r, cells in rows.items():
                out.write(f"{r} {len(cells)} {row_digest(cells, styles)}\n")
        for sty, idx in sorted(styles.items(), key=lambda x: x[1]):
            out.write(f"style {idx} {sty!r}\n")
    if not read_only:
        wb.properties.created = FIXED
        wb.properties.modified = FIXED
        wb.save(os.path.join(outdir, tag + ".xlsx"))


def compare(outdir):
    ok = True
    for suffix in ("", "-ro", "-do", "-ro-do"):
        a = os.path.join(outdir, "openpyxl" + suffix + ".digest")
        b = os.path.join(outdir, "openrsxl" + suffix + ".digest")
        if not (os.path.exists(a) and os.path.exists(b)):
            continue
        n = 0
        with open(a, encoding="utf-8") as fa, open(b, encoding="utf-8") as fb:
            for la, lb in zip(fa, fb):
                n += 1
                if la != lb.replace("openrsxl", "openpyxl"):
                    ok = False
                    print(f"digest{suffix} line {n} differs:\n  {la[:400]}\n  {lb[:400]}")
                    break
            else:
                if fa.read() or fb.read():
                    ok = False
                    print(f"digest{suffix}: different number of lines")
        print(f"digest{suffix}: {n} lines compared")
    for xa, xb in (
        (os.path.join(outdir, "openpyxl.xlsx"), os.path.join(outdir, "openrsxl.xlsx")),
        (os.path.join(outdir, "openpyxl.xlsx"), os.path.join(outdir, "openrsxl-direct.xlsx")),
        (os.path.join(outdir, "openpyxl-do.xlsx"), os.path.join(outdir, "openrsxl-do.xlsx")),
    ):
        if not (os.path.exists(xa) and os.path.exists(xb)):
            continue
        print("comparing", os.path.basename(xa), "with", os.path.basename(xb))
        za, zb = zipfile.ZipFile(xa), zipfile.ZipFile(xb)
        if za.namelist() != zb.namelist():
            ok = False
            print("saved part names differ")
        for name in za.namelist():
            da, db = za.read(name), zb.read(name)
            if name == "docProps/core.xml":
                da = re.sub(rb"<dcterms:modified[^<]*</dcterms:modified>", b"", da)
                db = re.sub(rb"<dcterms:modified[^<]*</dcterms:modified>", b"", db)
            if da != db:
                ok = False
                i = next((i for i in range(min(len(da), len(db))) if da[i] != db[i]), min(len(da), len(db)))
                print(f"saved part {name} differs at byte {i}: {da[max(0, i-80):i+80]!r} != {db[max(0, i-80):i+80]!r}")
        print(f"saved parts compared: {len(za.namelist())}")
    print("IDENTICAL" if ok else "DIFFERENT")
    return ok


if __name__ == "__main__":
    args = sys.argv[1:]
    if args[0] == "--compare":
        sys.exit(0 if compare(args[1]) else 1)
    ro = "--read-only" in args
    if "--save-only" in args:
        # load and save immediately (no cell is accessed in between)
        args = [a for a in args if a != "--save-only"]
        lib = importlib.import_module(args[0])
        os.makedirs(args[2], exist_ok=True)
        wb = lib.load_workbook(args[1])
        wb.properties.created = FIXED
        wb.properties.modified = FIXED
        wb.save(os.path.join(args[2], args[0] + "-direct.xlsx"))
        sys.exit(0)
    do = "--data-only" in args
    args = [a for a in args if a not in ("--read-only", "--data-only")]
    dump(args[0], args[1], args[2], read_only=ro, data_only=do)
