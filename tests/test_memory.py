"""
Memory of workbooks the program is done with.

A workbook and its worksheets reference each other (as in openpyxl), so a
dropped workbook is freed by Python's cycle collector. openrsxl keeps cells
in Rust and creates few Python objects, so that collector may not run for
a long time. These tests check, with the collector switched off, that

* ``close()`` of a read-only workbook drops its shared strings at once
  (a row iteration already running keeps working);
* the next ``load_workbook`` collects workbooks closed and dropped before;
* ``openrsxl.extended.release(wb)`` frees a workbook in any mode at once,
  even while its worksheets or cells are still referenced;
* repeated load / close cycles do not keep Python objects alive.
"""

import gc
import os
import weakref

import pytest

import openrsxl
from openrsxl import _native
from openrsxl.extended import load_workbook as ext_load
from openrsxl.extended import release


@pytest.fixture
def book(tmp_path):
    """a workbook with a shared-strings table (XlsxWriter writes one;
    openpyxl / openrsxl write inline strings)"""
    xlsxwriter = pytest.importorskip("xlsxwriter")
    p = str(tmp_path / "book.xlsx")
    wb = xlsxwriter.Workbook(p)
    ws = wb.add_worksheet("Sheet")
    for r in range(300):
        ws.write_row(r, 0, [f"text {(r + 1) % 37}", r + 1, (r + 1) * 0.5])
        ws.write_formula(r, 3, f"=B{r + 1}*2", None, (r + 1) * 2)
    wb.add_worksheet("Two").write(0, 0, "two")
    wb.close()
    return p


@pytest.fixture
def no_gc():
    gc.collect()
    _native._closed_pending = False
    was = gc.isenabled()
    gc.disable()
    yield
    if was:
        gc.enable()
    gc.collect()


LOADERS = {
    "full": lambda p: openrsxl.load_workbook(p),
    "data_only": lambda p: openrsxl.load_workbook(p, data_only=True),
    "read_only": lambda p: openrsxl.load_workbook(p, read_only=True),
    "ext_read_only": lambda p: ext_load(p, read_only=True, read_all=True),
    "ext_full": lambda p: ext_load(p, read_all=True),
}


def test_read_only_close_drops_shared_strings(book):
    wb = openrsxl.load_workbook(book, read_only=True)
    ws = wb["Sheet"]
    rows = ws.iter_rows(values_only=True)
    first = next(rows)
    assert first[0] == "text 1" and ws._shared_strings
    wb.close()
    assert ws._shared_strings == []
    # the iteration already running still reads every row correctly
    rest = list(rows)
    assert len(rest) == 299 and rest[-1][0] == "text 4"
    # a new iteration fails as in openpyxl (the archive is closed)
    with pytest.raises(ValueError):
        list(ws.iter_rows())


def test_full_mode_close_keeps_workbook_usable(book, tmp_path):
    """close() keeps openpyxl's meaning in full mode"""
    wb = openrsxl.load_workbook(book)
    wb.close()
    assert wb["Sheet"]["A2"].value == "text 2"
    wb.save(str(tmp_path / "again.xlsx"))


@pytest.mark.parametrize("mode", sorted(LOADERS))
def test_next_load_collects_closed(book, no_gc, mode):
    wb = LOADERS[mode](book)
    for ws in wb.worksheets:
        for _ in ws.iter_rows(values_only=True):
            pass
    ref = weakref.ref(wb)
    wb.close()
    del wb, ws
    assert ref() is not None  # a cycle: waits for the collector
    other = LOADERS[mode](book)
    assert ref() is None  # collected by the next load
    other.close()


def test_collect_only_once(book, no_gc, monkeypatch):
    """A closed workbook still in use costs one collection, not one per load"""
    kept = openrsxl.load_workbook(book)
    kept.close()
    calls = []
    monkeypatch.setattr(_native.gc, "collect", lambda *a: calls.append(1) or 0)
    for _ in range(3):
        openrsxl.load_workbook(book)
    assert len(calls) == 1
    assert kept["Sheet"]["B3"].value == 3


@pytest.mark.parametrize("mode", sorted(LOADERS))
def test_release(book, no_gc, mode):
    wb = LOADERS[mode](book)
    ws = wb.worksheets[0]
    for row in ws.iter_rows():
        cell = row[0]
    refs = [weakref.ref(wb), weakref.ref(wb.worksheets[1])]
    store = getattr(ws, "_cells", None)
    release(wb)
    del wb
    # the workbook and the other worksheets are freed without the collector,
    # although `ws` and `cell` are still referenced
    assert [r() for r in refs] == [None, None]
    if store is not None:
        assert len(store) == 0
    # the worksheet is emptied: no workbook, no cells
    assert vars(ws) == {} and getattr(ws, "parent", None) is None
    del cell


@pytest.mark.parametrize("mode", sorted(LOADERS))
def test_loaded_workbook_has_no_cycles(book, no_gc, mode):
    """Only the workbook <-> worksheets references (openpyxl's object model)
    hold a loaded workbook: release() breaks them, nothing else remains"""
    wb = LOADERS[mode](book)
    ref = weakref.ref(wb)
    release(wb)
    del wb
    assert ref() is None


@pytest.mark.parametrize("mode", sorted(LOADERS))
def test_no_growth(book, mode):
    """Python objects alive after many load / use / close cycles stay flat"""

    def cycle():
        wb = LOADERS[mode](book)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    _ = c.value
        wb.close()

    for _ in range(3):
        cycle()
    gc.collect()
    before = len(gc.get_objects())
    for _ in range(20):
        cycle()
    gc.collect()
    assert len(gc.get_objects()) - before < 200


# -- parity with openpyxl: close() must behave exactly as openpyxl's ----------------
def _outcome(f):
    """(result, None) or (None, (exception type name, message))"""
    try:
        return f(), None
    except Exception as e:  # noqa: BLE001 - compared between the libraries
        return None, (type(e).__name__, str(e))


def _close_scenarios(lib, path, tmp):
    """Everything observable around close(), as plain data"""
    out = {}
    # read-only: iteration started before close continues; new ones fail
    wb = lib.load_workbook(path, read_only=True)
    ws = wb["Sheet"]
    rows = ws.iter_rows(values_only=True)
    first = next(rows)
    pending = ws.iter_rows(min_row=5, max_row=6, values_only=True)  # not started
    wb.close()
    out["ro_first"] = first
    out["ro_rest"] = _outcome(lambda: list(rows))
    out["ro_pending"] = _outcome(lambda: list(pending))
    out["ro_new_iter"] = _outcome(lambda: list(ws.iter_rows()))
    out["ro_values"] = _outcome(lambda: list(ws.values))
    out["ro_cell"] = _outcome(lambda: ws["A1"].value)
    out["ro_title"] = ws.title
    out["ro_sheetnames"] = wb.sheetnames
    out["ro_dims"] = _outcome(lambda: ws.calculate_dimension())
    out["ro_close_twice"] = _outcome(wb.close)
    # full mode: still usable and savable after close
    for kw in ({}, {"data_only": True}, {"keep_vba": True}):
        wb = lib.load_workbook(path, **kw)
        wb.close()
        key = "full" + "".join(f"_{k}" for k in kw)
        out[key + "_values"] = [[c.value for c in r] for r in wb["Sheet"].iter_rows(max_row=5)]
        out[key + "_types"] = [[c.data_type for c in r] for r in wb["Sheet"].iter_rows(max_row=5)]
        wb["Sheet"]["A1"] = "edited"
        target = os.path.join(tmp, f"{lib.__name__}{key}.xlsx")
        out[key + "_save"] = _outcome(lambda: wb.save(target))
        out[key + "_close_twice"] = _outcome(wb.close)
        out[key + "_sheetnames"] = wb.sheetnames
    # write-only
    wb = lib.Workbook(write_only=True)
    ws = wb.create_sheet("W")
    ws.append([1, "two"])
    wb.close()
    out["wo_append_after_close"] = _outcome(lambda: ws.append([3]))
    out["wo_save"] = _outcome(lambda: wb.save(os.path.join(tmp, f"{lib.__name__}_wo.xlsx")))
    # a new workbook
    wb = lib.Workbook()
    wb.close()
    wb.active["A1"] = 5
    out["new_after_close"] = wb.active["A1"].value
    return out


def test_close_parity(book, tmp_path):
    import openpyxl

    ours = _close_scenarios(openrsxl, book, str(tmp_path))
    theirs = _close_scenarios(openpyxl, book, str(tmp_path))
    assert ours.keys() == theirs.keys()
    for k in theirs:
        assert ours[k] == theirs[k], k


def test_load_parity_after_closed(book):
    """A load that runs the collector first returns the same workbook"""
    import openpyxl

    def snapshot(lib):
        wb = lib.load_workbook(book)
        data = (wb.sheetnames, [[(c.value, c.data_type, c.style_id) for c in r] for r in wb["Sheet"].iter_rows()])
        return data

    first = openrsxl.load_workbook(book)
    first.close()  # pending: the next load collects
    assert snapshot(openrsxl)[0] == snapshot(openpyxl)[0]
    assert snapshot(openrsxl)[1] == snapshot(openpyxl)[1]
