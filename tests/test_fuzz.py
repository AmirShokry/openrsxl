"""
Property based differential tests (hypothesis) against openpyxl.
"""

import io
import os
import struct

import openpyxl
from fixtures import cells, make_xlsx, sheet_xml
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st
from oracle.compare import compare_workbooks, saved_parts

import openrsxl

EXAMPLES = int(os.environ.get("OPENRSXL_FUZZ_EXAMPLES", "300"))
SETTINGS = settings(max_examples=EXAMPLES, deadline=None, suppress_health_check=list(HealthCheck))


def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------- formulas

FORMULA_ATOMS = st.sampled_from(
    [
        "A1",
        "$A$1",
        "A$1",
        "$A1",
        "B2:C3",
        "A:A",
        "$B:C",
        "1:1",
        "$2:3",
        "Sheet1!A1",
        "'My Sheet'!B2",
        "'it''s'!C3",
        "[1]Sheet1!A1",
        "myName",
        "SUM(",
        "IF(",
        ")",
        ",",
        ";",
        "(",
        "{",
        "}",
        "+",
        "-",
        "*",
        "/",
        "^",
        "&",
        "=",
        "<>",
        ">=",
        "<=",
        "<",
        ">",
        "%",
        " ",
        "\n",
        '"str"',
        '"a""b"',
        "1",
        "1.5",
        "1E+3",
        "2e-1",
        "TRUE",
        "FALSE",
        "#REF!",
        "#N/A",
        "#DIV/0!",
        "XFD1048576",
        "ZZZ1",
        "a1",
        "zz10",
        "A0",
        "A1:A2:A3",
        "Table1[Col]",
        "Table1[[#This Row],[a]]",
        "!",
        ":",
        "'",
        '"',
        "[",
        "]",
        "$",
        "#",
        "INF",
        "nan",
        "1_0",
        "é",
        "١",
        "R1C1",
        "A1048577",
        "AAAA1",
        "_xlfn.X(",
    ]
)

formulas = st.lists(FORMULA_ATOMS, min_size=0, max_size=12).map(lambda xs: "".join(xs))
coords = st.tuples(st.integers(1, 30), st.integers(1, 30))


def coord(rc):
    r, c = rc
    return openpyxl.utils.get_column_letter(c) + str(r)


@SETTINGS
@given(formulas, coords, st.lists(coords, min_size=1, max_size=4))
def test_shared_formula_translation(formula, master, deps):
    rows = {}
    rows.setdefault(master[0], []).append(
        f'<c r="{coord(master)}"><f t="shared" ref="A1:Z30" si="0">{xml_escape(formula)}</f></c>'
    )
    for d in deps:
        if d == master:
            continue
        rows.setdefault(d[0], []).append(f'<c r="{coord(d)}"><f t="shared" si="0"/></c>')
    data = "".join(f'<row r="{r}">' + "".join(cs) + "</row>" for r, cs in rows.items())
    xlsx = make_xlsx([data])
    assert not compare_workbooks(xlsx, check_save=False)
    assert not compare_workbooks(xlsx, check_save=False, read_only=True)


@SETTINGS
@given(formulas)
def test_tokenizer_matches(formula):
    from openpyxl.formula.tokenizer import Tokenizer as A

    from openrsxl.formula.tokenizer import Tokenizer as B

    def run(T):
        try:
            return [(t.value, t.type, t.subtype) for t in T("=" + formula).items]
        except Exception as e:
            return (type(e).__name__, str(e))

    assert run(A) == run(B)


# ---------------------------------------------------------------- numbers

NUMERIC_TEXT = st.text(alphabet="0123456789.eE+-_ \tinfaINFA١x", min_size=1, max_size=12)


@SETTINGS
@given(st.lists(NUMERIC_TEXT, min_size=1, max_size=5), st.sampled_from(["0", "1", "2", "6"]))
def test_number_parsing(texts, style):
    cs = "".join(f'<c r="A{i}" s="{style}"><v>{xml_escape(t)}</v></c>' for i, t in enumerate(texts, 1))
    xlsx = make_xlsx([f'<row r="1">{cs}</row>'])
    assert not compare_workbooks(xlsx, check_save=False)


floats = st.floats(allow_nan=True, allow_infinity=True)
any_bits = st.integers(0, 2**64 - 1).map(lambda i: struct.unpack("<d", struct.pack("<Q", i))[0])


def save_values(L, values):
    wb = L.Workbook()
    ws = wb.active
    for i, v in enumerate(values, 1):
        ws.cell(row=i, column=1, value=v)
    return saved_parts(wb)[1]["xl/worksheets/sheet1.xml"]


@SETTINGS
@given(st.lists(st.one_of(floats, any_bits, st.integers(-(2**70), 2**70)), min_size=1, max_size=20))
def test_number_writing(values):
    assert save_values(openpyxl, values) == save_values(openrsxl, values)


@SETTINGS
@given(st.lists(st.one_of(floats, any_bits), min_size=1, max_size=20))
def test_number_roundtrip(values):
    """Floats written by openpyxl and read back by both libraries."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, v in enumerate(values, 1):
        ws.cell(row=i, column=1, value=v)
    b = io.BytesIO()
    wb.save(b)
    assert not compare_workbooks(b.getvalue(), check_save=False)


@SETTINGS
@given(
    st.lists(st.floats(min_value=-1e7, max_value=3e6, allow_nan=False), min_size=1, max_size=20),
    st.sampled_from(["1", "2", "4", "5"]),
)
def test_date_serials(values, style):
    cs = "".join(f'<c r="A{i}" s="{style}"><v>{v!r}</v></c>' for i, v in enumerate(values, 1))
    xlsx = make_xlsx([f'<row r="1">{cs}</row>'])
    assert not compare_workbooks(xlsx, check_save=False)


# ---------------------------------------------------------------- strings

text_values = st.text(min_size=0, max_size=30)


@SETTINGS
@given(st.lists(text_values, min_size=1, max_size=10))
def test_string_writing(values):
    def build(L):
        try:
            return save_values(L, values)
        except Exception as e:
            return (type(e).__name__, str(e))

    assert build(openpyxl) == build(openrsxl)


@SETTINGS
@given(
    st.lists(
        st.text(alphabet=st.characters(blacklist_categories=("Cs",), blacklist_characters="\x00"), max_size=20),
        min_size=1,
        max_size=10,
    )
)
def test_string_roundtrip(values):
    """Strings written by openpyxl, read back (inline strings)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, v in enumerate(values, 1):
        try:
            ws.cell(row=i, column=1, value=v)
        except Exception:
            pass
    b = io.BytesIO()
    wb.save(b)
    assert not compare_workbooks(b.getvalue(), check_save=False)
    assert not compare_workbooks(b.getvalue(), check_save=False, read_only=True)


# ---------------------------------------------------------------- XML robustness

BASE_DOC = sheet_xml(
    cells(
        '<c r="A1" t="s"><v>1</v></c>',
        '<c r="B1"><v>2.5</v></c>',
        '<c r="C1" t="inlineStr"><is><t xml:space="preserve"> x </t></is></c>',
        '<c r="D1"><f t="shared" ref="D1:D2" si="0">A1+B1</f><v>3</v></c>',
    )
    + cells('<c r="D2"><f t="shared" si="0"/></c>', '<c r="E2" t="str"><v>a&amp;b</v></c>', r=2),
    before='<sheetPr><outlinePr summaryBelow="0"/></sheetPr><cols><col min="1" max="2" width="9"/></cols>',
    after='<mergeCells count="1"><mergeCell ref="F1:G2"/></mergeCells>',
)

MUTATION_CHARS = [
    "<",
    ">",
    "&",
    '"',
    "'",
    "/",
    "=",
    " ",
    "\n",
    "\r",
    "\t",
    "!",
    "?",
    "-",
    "[",
    "]",
    ";",
    "#",
    "x",
    ":",
    "\x00",
    "\x01",
    "é",
    "￾",
    "&amp;",
    "&#0;",
    "&#x41;",
    "<!--",
    "-->",
    "<![CDATA[",
    "]]>",
    "<?pi?>",
    "<a>",
    "</a>",
    "<b/>",
    'q="1"',
    'xmlns:z="u"',
]


@SETTINGS
@given(
    st.lists(
        st.tuples(
            st.integers(0, len(BASE_DOC)),
            st.sampled_from(["ins", "del", "dup"]),
            st.sampled_from(MUTATION_CHARS),
            st.integers(1, 6),
        ),
        min_size=1,
        max_size=3,
    )
)
def test_xml_mutations(mutations):
    doc = BASE_DOC
    for pos, kind, chars, n in mutations:
        pos = min(pos, len(doc))
        if kind == "ins":
            doc = doc[:pos] + chars + doc[pos:]
        elif kind == "del":
            doc = doc[:pos] + doc[pos + n :]
        else:
            doc = doc[:pos] + doc[pos : pos + n] * 2 + doc[pos + n :]
    xlsx = make_xlsx(raw_sheets=[doc.encode("utf-8", "surrogatepass")])
    assert not compare_workbooks(xlsx, check_save=False)
    assert not compare_workbooks(xlsx, check_save=False, read_only=True)


SST_DOC = (
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<si><t>a</t></si><si><r><rPr><b/></rPr><t>b</t></r><r><t xml:space="preserve"> c</t></r></si>'
    "<si><t>x005F_y</t></si></sst>"
)


@SETTINGS
@given(
    st.lists(
        st.tuples(
            st.integers(0, len(SST_DOC)),
            st.sampled_from(["ins", "del"]),
            st.sampled_from(MUTATION_CHARS + ["<si><t>n</t></si>", "<r><t>z</t></r>", "<rPh><t>p</t></rPh>"]),
            st.integers(1, 6),
        ),
        min_size=1,
        max_size=3,
    )
)
def test_shared_string_mutations(mutations):
    doc = SST_DOC
    for pos, kind, chars, n in mutations:
        pos = min(pos, len(doc))
        if kind == "ins":
            doc = doc[:pos] + chars + doc[pos:]
        else:
            doc = doc[:pos] + doc[pos + n :]
    sheet = cells(*[f'<c r="A{i}" t="s"><v>{i}</v></c>' for i in range(3)])
    xlsx = make_xlsx([sheet], strings=doc.encode("utf-8", "surrogatepass"))
    assert not compare_workbooks(xlsx, check_save=False)
