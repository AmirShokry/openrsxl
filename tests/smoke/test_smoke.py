"""
Smoke test of an installed openrsxl without the test dependencies (no
openpyxl oracle, numpy, lxml ...): used where only openrsxl and pytest can
be installed (Pyodide / WebAssembly builds).
"""

import datetime
import io
import os
import platform

import pytest

import openrsxl
import openrsxl.extended
from openrsxl import _native
from openrsxl.comments import Comment


def make_book():
    wb = openrsxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["text", 1, 2.5, True, datetime.datetime(2024, 2, 29, 12, 30), None, "=B1+C1"])
    ws["A3"] = "merged"
    ws.merge_cells("A3:C4")
    ws["A5"].hyperlink = "https://example.com"
    ws["B5"].comment = Comment("note", "me")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.mark.skipif(
    platform.python_implementation() != "CPython" or os.environ.get("OPENRSXL_PURE_PYTHON"),
    reason="the Rust engine is used on CPython only (unless disabled)",
)
def test_native_engine_is_used():
    assert _native.ENABLED


def test_roundtrip_values_and_types():
    data = make_book()
    for read_only in (False, True):
        wb = openrsxl.load_workbook(io.BytesIO(data), read_only=read_only)
        row = next(wb["Data"].iter_rows(max_row=1, values_only=True))
        assert row == ("text", 1, 2.5, True, datetime.datetime(2024, 2, 29, 12, 30), None, "=B1+C1")
        assert [type(v) for v in row[:5]] == [str, int, float, bool, datetime.datetime]


def test_full_mode_features():
    wb = openrsxl.load_workbook(io.BytesIO(make_book()))
    ws = wb["Data"]
    assert ws["Z99"].value is None
    assert [str(r) for r in ws.merged_cells] == ["A3:C4"]
    assert ws["A5"].hyperlink.target == "https://example.com"
    assert ws["B5"].comment.text == "note"


def test_extended_streaming():
    wb = openrsxl.extended.load_workbook(
        io.BytesIO(make_book()),
        read_only=True,
        formula_and_value=True,
        read_comments=True,
        read_hyperlinks=True,
        read_merged_cells=True,
    )
    ws = wb["Data"]
    first = next(ws.iter_rows(max_row=1))
    assert first[6].formula == "=B1+C1" and first[6].cached_value is None
    assert [str(r) for r in ws.merged_cells] == ["A3:C4"]
    cells = {c.coordinate: c for row in ws.iter_rows() for c in row if hasattr(c, "coordinate")}
    assert cells["A5"].hyperlink.target == "https://example.com"
    assert cells["B5"].comment.text == "note"
    assert type(cells["B3"]).__name__ == "MergedCell"
    wb.close()
