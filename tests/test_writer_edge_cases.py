"""
Writer edge cases: the same workbook is built through the public API with
openpyxl and with openrsxl, saved, and the packages must be byte-identical
(or both must raise the same exception).  Saved files are then re-loaded by
both libraries and compared.
"""

import datetime
import importlib
import io
from decimal import Decimal

import pytest
from oracle.compare import ADDR_RE, compare_workbooks, norm_text, saved_parts

try:
    import numpy as np
except ImportError:  # pragma: no cover
    np = None

INTS = [0, 1, -1, 7, 2**53, 2**53 + 1, 2**63 - 1, 2**63, 10**30, -(10**30), True, False]
FLOATS = [
    0.0,
    -0.0,
    0.1,
    1 / 3,
    1e-7,
    1e-5,
    1e-4,
    1e15,
    1e16,
    1e17,
    123456789.123456789,
    float("nan"),
    float("inf"),
    float("-inf"),
    5e-324,
    1.7976931348623157e308,
    0.1 + 0.2,
    2.5,
    -2.5e-10,
    9007199254740993.0,
    1234567890123456.7,
]
OTHER_NUMS = [Decimal("1.5"), Decimal("1e400"), Decimal("NaN")]
if np is not None:
    OTHER_NUMS += [np.int64(5), np.float32(1.1), np.float64(2.5), np.bool_(True), np.uint8(200), np.float16(0.5)]
STRINGS = [
    "",
    " ",
    "  a ",
    "\t",
    "a\nb",
    "a\r\nb",
    "\r",
    "<&>\"'",
    "é中\U0001f600",
    " ",
    "\x85",
    "\xa0x\xa0",
    "=SUM(A1)",
    "=",
    "==",
    "#N/A",
    "#REF!",
    "x" * 40000,
    "]]>",
    "a\tb",
    "　x",
    "\x1cx\x1d",
    "�",
    "\U0010ffff",
    "&amp;",
    " \n",
    "\n",
    b"bytes",
    b"caf\xc3\xa9",
]
DATES = [
    datetime.datetime(2020, 1, 2, 3, 4, 5),
    datetime.datetime(2020, 1, 2, 3, 4, 5, 123456),
    datetime.date(1900, 1, 1),
    datetime.date(1900, 2, 28),
    datetime.date(1900, 3, 1),
    datetime.date(1899, 12, 30),
    datetime.time(12, 30),
    datetime.time(0, 0, 0, 1),
    datetime.timedelta(days=1.5),
    datetime.timedelta(seconds=-1),
    datetime.timedelta(0),
    datetime.datetime(9999, 12, 31, 23, 59, 59),
    datetime.datetime(1, 1, 1),
]


def lib(name):
    return importlib.import_module(name)


def build_values(L, values, **wbkw):
    wb = L.Workbook()
    for k, v in wbkw.items():
        setattr(wb, k, v)
    ws = wb.active
    for i, v in enumerate(values, 1):
        ws.cell(row=i, column=1).value = v
        ws.cell(row=i, column=2, value=v)
    return wb


def result(name, builder):
    L = lib(name)
    try:
        wb = builder(L)
        names, parts = saved_parts(wb)
        return ("ok", names, parts)
    except Exception as e:
        return ("error", type(e).__name__, ADDR_RE.sub("", norm_text(str(e))))


def assert_same(builder):
    a = result("openpyxl", builder)
    b = result("openrsxl", builder)
    assert a[0] == b[0], (a[:3] if a[0] == "error" else "ok", b[:3] if b[0] == "error" else "ok")
    if a[0] == "error":
        assert a == b
        return None
    assert a[1] == b[1]
    for name in a[1]:
        x, y = a[2][name], b[2][name]
        if x != y:
            i = next((i for i in range(min(len(x), len(y))) if x[i] != y[i]), min(len(x), len(y)))
            pytest.fail(f"{name} differs at {i}: {x[max(0, i-100):i+100]!r} != {y[max(0, i-100):i+100]!r}")
    # round trip through both readers
    b = io.BytesIO()
    import zipfile

    with zipfile.ZipFile(b, "w") as z:
        for name in a[1]:
            z.writestr(name, a[2][name])
    diffs = compare_workbooks(b.getvalue(), check_save=False)
    assert not diffs, diffs
    return a


VALUE_GROUPS = {"ints": INTS, "floats": FLOATS, "other_nums": OTHER_NUMS, "dates": DATES}


@pytest.mark.parametrize("group", list(VALUE_GROUPS))
def test_value_groups(group):
    assert_same(lambda L: build_values(L, VALUE_GROUPS[group]))


@pytest.mark.parametrize("idx", range(len(STRINGS)))
def test_strings(idx):
    assert_same(lambda L: build_values(L, [STRINGS[idx]]))


@pytest.mark.parametrize("value", INTS + FLOATS + DATES, ids=repr)
def test_single_values(value):
    assert_same(lambda L: build_values(L, [value]))


def test_iso_dates():
    assert_same(lambda L: build_values(L, DATES, iso_dates=True))


def test_epoch_1904():
    def b(L):
        wb = build_values(L, DATES[:6])
        wb.epoch = importlib.import_module(L.__name__ + ".utils.datetime").MAC_EPOCH
        return wb

    assert_same(b)


def test_timezone_error():
    assert_same(lambda L: build_values(L, [datetime.datetime(2020, 1, 1, tzinfo=datetime.timezone.utc)]))


def test_illegal_characters():
    assert_same(lambda L: build_values(L, ["a\x01b"]))


def test_lone_surrogate():
    assert_same(lambda L: build_values(L, ["\ud800"]))


def test_rich_text():
    def b(L):
        from importlib import import_module

        rt = import_module(L.__name__ + ".cell.rich_text")
        tx = import_module(L.__name__ + ".cell.text")
        wb = L.Workbook()
        ws = wb.active
        ws["A1"] = rt.CellRichText("plain ", rt.TextBlock(tx.InlineFont(b=True, sz=14), "bold"), " end ")
        ws["A2"] = rt.CellRichText()
        ws["A3"] = rt.CellRichText(" x ")
        return wb

    assert_same(b)


def test_array_and_datatable_formulas():
    def b(L):
        f = importlib.import_module(L.__name__ + ".worksheet.formula")
        wb = L.Workbook()
        ws = wb.active
        ws["A1"] = f.ArrayFormula("A1:A3", "=SUM(B1:B3*C1:C3)")
        ws["B1"] = f.DataTableFormula(ref="B1:C3", dt2D=True, r1="A1", r2="A2")
        ws["C1"] = "=A1+B1"
        ws["D1"] = '="a<b"&"&"'
        return wb

    assert_same(b)


def test_styles_everything():
    def b(L):
        st = importlib.import_module(L.__name__ + ".styles")
        wb = L.Workbook()
        ws = wb.active
        ws["A1"] = 1
        ws["A1"].font = st.Font(name="Arial", bold=True, italic=True, size=13, color="FF00FF00", underline="double")
        ws["B1"].fill = st.PatternFill("solid", fgColor="FFFF0000")
        ws["C1"].border = st.Border(left=st.Side(style="thin", color="FF000000"), diagonal=st.Side(style="dashed"))
        ws["D1"].alignment = st.Alignment(horizontal="center", wrap_text=True, text_rotation=45)
        ws["E1"].protection = st.Protection(locked=False, hidden=True)
        ws["F1"] = 3.14159
        ws["F1"].number_format = "0.00%"
        ws["G1"] = datetime.datetime(2020, 5, 6)
        ws["G1"].number_format = "yyyy-mm-dd hh:mm"
        ws["H1"] = "styled"
        ws["H1"].style = "Good"
        ws["I1"].style = "Hyperlink"
        ns = st.NamedStyle(name="mine", font=st.Font(bold=True))
        wb.add_named_style(ns)
        ws["J1"] = "named"
        ws["J1"].style = "mine"
        ws["K1"].number_format = "General"
        ws["L1"].quotePrefix = True
        ws["M1"].pivotButton = True
        for col in "NOPQ":
            ws[f"{col}2"].font = st.Font(color="FF123456")
        return wb

    assert_same(b)


def test_hyperlinks_comments_merged():
    def b(L):
        cm = importlib.import_module(L.__name__ + ".comments")
        st = importlib.import_module(L.__name__ + ".styles")
        wb = L.Workbook()
        ws = wb.active
        ws["A1"].hyperlink = "https://example.com/?a=1&b=<2>"
        ws["A2"] = "text"
        ws["A2"].hyperlink = "#Sheet!B2"
        ws["A3"].value = "loc"
        ws["A3"].hyperlink = importlib.import_module(L.__name__ + ".worksheet.hyperlink").Hyperlink(
            ref="A3", location="'Sheet'!C3", tooltip="tip"
        )
        ws["B1"].comment = cm.Comment("hello <world> & co", "me")
        ws["B2"] = 5
        ws["B2"].comment = cm.Comment("c2", "you", height=50, width=80)
        ws["C5"] = "merged"
        ws["C5"].border = st.Border(top=st.Side(style="thick"), bottom=st.Side(style="thin"))
        ws.merge_cells("C5:E7")
        ws.merge_cells("G1:G4")
        return wb

    assert_same(b)


def test_dimensions():
    def b(L):
        st = importlib.import_module(L.__name__ + ".styles")
        wb = L.Workbook()
        ws = wb.active
        ws["A1"] = 1
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[3].hidden = True
        ws.row_dimensions[5].outlineLevel = 2
        ws.row_dimensions[7].font = st.Font(bold=True)
        ws.row_dimensions[2].height = 12.75
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].hidden = True
        ws.column_dimensions.group("D", "F", outline_level=1, hidden=True)
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = "A1:C10"
        ws.print_area = "A1:D20"
        return wb

    assert_same(b)


def test_data_types_manually_set():
    def b(L):
        wb = L.Workbook()
        ws = wb.active
        ws["A1"] = 5
        ws["A1"].data_type = "s"
        ws["A2"] = "x"
        ws["A2"].data_type = "n"
        ws["A3"] = "q"
        ws["A3"].data_type = "e"
        ws["A4"] = 1
        ws["A4"].data_type = "b"
        ws["A5"] = "abc"
        ws["A5"].data_type = 'zz<"'
        ws["A6"] = 2
        ws["A6"].data_type = "f"
        return wb

    assert_same(b)


def test_bad_value_type():
    def b(L):
        wb = L.Workbook()
        ws = wb.active
        ws["A1"]._value = [1, 2]
        return wb

    assert_same(b)


def test_cell_api_errors():
    def b(L):
        wb = L.Workbook()
        ws = wb.active
        ws["A1"] = object()
        return wb

    assert_same(b)


def test_many_cells_sparse_and_unordered():
    def b(L):
        wb = L.Workbook()
        ws = wb.active
        for r in (500, 3, 70, 1, 2):
            for c in (30, 1, 5):
                ws.cell(row=r, column=c, value=r * c)
        ws.cell(row=1, column=16384, value="last col")
        ws.cell(row=1048576, column=1, value="last row")
        ws.cell(row=10, column=10)  # empty unstyled cell is skipped
        return wb

    assert_same(b)


def test_multiple_sheets_and_copy():
    def b(L):
        wb = L.Workbook()
        ws = wb.active
        ws["A1"] = "orig"
        ws["B2"] = 2.5
        cp = wb.copy_worksheet(ws)
        cp["C3"] = "copy"
        wb.create_sheet("Empty")
        wb.create_chartsheet("Chart")
        return wb

    assert_same(b)


def test_insert_delete_move():
    def b(L):
        wb = L.Workbook()
        ws = wb.active
        for r in range(1, 8):
            for c in range(1, 6):
                ws.cell(r, c, f"=A{r}+{c}" if c == 5 else r * 10 + c)
        ws.insert_rows(2, 2)
        ws.delete_cols(3)
        ws.move_range("A1:B3", rows=4, cols=2, translate=True)
        ws.insert_cols(1)
        ws.delete_rows(7, 3)
        return wb

    assert_same(b)


def test_append_and_iterables():
    def b(L):
        wb = L.Workbook()
        ws = wb.active
        ws.append([1, "two", 3.0, None, True, datetime.date(2020, 1, 1)])
        ws.append({"A": 1, "C": 3})
        ws.append({2: "b", 4: "d"})
        ws.append(range(5))
        ws.append(x for x in "abc")
        return wb

    assert_same(b)


WRITE_ONLY_VALUES = INTS + FLOATS + DATES + [s for s in STRINGS if s != "x" * 40000] + [None]


def test_write_only():
    def b(L):
        st = importlib.import_module(L.__name__ + ".styles")
        cm = importlib.import_module(L.__name__ + ".comments")
        cell_mod = importlib.import_module(L.__name__ + ".cell")
        wb = L.Workbook(write_only=True)
        ws = wb.create_sheet("WO")
        ws.column_dimensions["B"].width = 30
        ws.row_dimensions[2].height = 40
        for v in WRITE_ONLY_VALUES:
            ws.append([v, None, v])
        c = cell_mod.WriteOnlyCell(ws, value="styled")
        c.font = st.Font(bold=True)
        c.comment = cm.Comment("note", "me")
        h = cell_mod.WriteOnlyCell(ws, value="link")
        h.hyperlink = "https://example.com"
        ws.append([c, h, "=SUM(A1:A3)", 5])
        ws.append([])
        ws.append(range(3))
        wb.create_sheet("Empty")
        return wb

    assert_same(b)


def test_write_only_no_rows():
    def b(L):
        wb = L.Workbook(write_only=True)
        wb.create_sheet("x")
        return wb

    assert_same(b)


def test_template_and_properties():
    def b(L):
        wb = L.Workbook()
        wb.template = True
        wb.properties.title = "T <&>"
        wb.properties.creator = "someone"
        ws = wb.active
        ws.title = "Data & <More>"
        ws["A1"] = "x"
        wb.create_named_range if hasattr(wb, "create_named_range") else None
        dn = importlib.import_module(L.__name__ + ".workbook.defined_name")
        wb.defined_names["myname"] = dn.DefinedName("myname", attr_text="'Data & <More>'!$A$1")
        return wb

    assert_same(b)
