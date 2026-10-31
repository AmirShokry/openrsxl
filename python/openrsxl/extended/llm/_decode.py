"""decompress: openrsxl-llm document -> cells and openpyxl objects (the
reference reader of the format; it does not need the native engine)"""

import datetime
import json
import os
import re

from openrsxl.cell.text import Text  # noqa: F401 - imported for from_tree users
from openrsxl.comments import Comment
from openrsxl.formatting.rule import Rule
from openrsxl.packaging.core import DocumentProperties
from openrsxl.packaging.custom import CustomPropertyList
from openrsxl.styles.alignment import Alignment
from openrsxl.styles.borders import Border
from openrsxl.styles.differential import DifferentialStyle
from openrsxl.styles.fills import Fill
from openrsxl.styles.fonts import Font
from openrsxl.styles.protection import Protection
from openrsxl.utils.cell import column_index_from_string, get_column_letter
from openrsxl.workbook.defined_name import DefinedName
from openrsxl.workbook.properties import CalcProperties
from openrsxl.workbook.protection import WorkbookProtection
from openrsxl.workbook.views import BookView
from openrsxl.worksheet.datavalidation import DataValidation
from openrsxl.worksheet.formula import ArrayFormula, DataTableFormula
from openrsxl.worksheet.hyperlink import Hyperlink
from openrsxl.worksheet.table import Table

from . import _cx
from ._legend import FORMAT
from ._model import COLUMN_FIELDS, ROW_FIELDS, SHEET_TAGS, parse_field, read_object

_decoder = json.JSONDecoder()


class DecodeError(ValueError):
    pass


# -- value tokens ---------------------------------------------------------------
_INT = re.compile(r"[+-]?[0-9]+\Z")
_NUM = re.compile(r"[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([eE][+-]?[0-9]+)?\Z")
# (ASCII digits only: Python's \d also matches other scripts' digits)
_TIME = r"([0-9]{2}):([0-9]{2}):([0-9]{2})(?:\.([0-9]{1,6}))?"
_DATETIME = re.compile(r"([0-9]{4})-([0-9]{2})-([0-9]{2})(?: " + _TIME + r")?\Z")
_TIMEONLY = re.compile(_TIME + r"\Z")
_TD = re.compile(r"(-?)([0-9]+):([0-9]{2}):([0-9]{2})(?:\.([0-9]{6}))?\Z")
ERROR = "e"


def _us(frac):
    return int(frac.ljust(6, "0")) if frac else 0


def parse_token(tok):
    """(value, data_type) of a value token: data_type is 'n' (number), 's'
    (text), 'b' (boolean), 'e' (error) or 'd' (date / time / duration)"""
    if not tok:
        raise DecodeError("empty token")
    c = tok[0]
    if c == '"':
        return json.loads(tok), "s"
    if c == "#":
        return tok, ERROR
    if c == "@":
        kind, _, payload = tok[1:].partition(":")
        if kind == "err":
            return json.loads(payload), ERROR
        if kind == "date":
            return datetime.date.fromisoformat(payload), "d"
        if kind == "td":
            m = _TD.match(payload)
            if not m:
                raise DecodeError(f"bad duration {tok!r}")
            sign, h, mi, s, us = m.groups()
            td = datetime.timedelta(hours=int(h), minutes=int(mi), seconds=int(s), microseconds=_us(us))
            return (-td if sign else td), "d"
        if kind == "float":
            return float(payload), "n"
        if kind == "dt":
            return datetime.datetime.fromisoformat(payload), "d"
        if kind == "time":
            return datetime.time.fromisoformat(payload), "d"
        raise DecodeError(f"unknown typed token {tok!r}")
    if tok == "TRUE":
        return True, "b"
    if tok == "FALSE":
        return False, "b"
    if _INT.match(tok):
        return int(tok), "n"
    if _NUM.match(tok):
        return float(tok), "n"
    m = _DATETIME.match(tok)
    if m:
        y, mo, d, h, mi, s, us = m.groups()
        if h is None:
            return datetime.datetime(int(y), int(mo), int(d)), "d"
        return datetime.datetime(int(y), int(mo), int(d), int(h), int(mi), int(s), _us(us)), "d"
    m = _TIMEONLY.match(tok)
    if m:
        h, mi, s, us = m.groups()
        return datetime.time(int(h), int(mi), int(s), _us(us)), "d"
    return tok, "s"


def _scan_token(s, i):
    """End of the token starting at s[i] (a grid field: up to '|' or the
    end; JSON strings may contain '|')"""
    n = len(s)
    if i < n and s[i] == '"':
        _, j = _decoder.raw_decode(s, i)
        return j
    if s.startswith("@err:", i) and s.startswith('"', i + 5):
        _, j = _decoder.raw_decode(s, i + 5)
        return j
    j = s.find("|", i)
    return n if j < 0 else j


_JUMP = re.compile(r"([A-Z]{1,3})=")
_RUN = re.compile(r"\*([0-9]+)")


_LABEL = re.compile(r"([0-9]+):")


def parse_grid_line(line, start_col, row=None):
    """[(row, col, token)] of a grid line 'ROW:f1|f2|...'; a line without
    the 'ROW:' label is row `row` (the row after the previous line)"""
    m = _LABEL.match(line)
    if m:
        row = int(m.group(1))
        rest = line[m.end() :]
    elif row is None:
        raise DecodeError(f"bad grid line {line[:40]!r}")
    else:
        rest = line
    out = []
    col = start_col
    i = 0
    n = len(rest)
    while True:
        if i < n and rest[i] != "|":
            m = _JUMP.match(rest, i)
            if m:
                col = column_index_from_string(m.group(1))
                i = m.end()
            j = _scan_token(rest, i)
            tok = rest[i:j]
            count = 1
            if tok[:1] != '"' and not tok.startswith('@err:"'):
                m = re.search(r"\*([0-9]+)\Z", tok)
                if m:
                    count = int(m.group(1))
                    tok = tok[: m.start()]
            else:
                m = _RUN.match(rest, j)
                if m:
                    count = int(m.group(1))
                    j = m.end()
            for k in range(count):
                out.append((row, col + k, tok))
            col += count
            i = j
        else:
            col += 1
        if i >= n:
            break
        if rest[i] != "|":
            raise DecodeError(f"bad grid field at {i} in {line[:60]!r}")
        i += 1
        if i == n:
            break
    return out


_REF = re.compile(r"\$?([A-Z]{1,3})\$?([0-9]+)\Z")


def parse_cell(ref):
    m = _REF.match(ref)
    if not m:
        raise DecodeError(f"bad cell reference {ref!r}")
    return int(m.group(2)), column_index_from_string(m.group(1))


def iter_range(rng):
    """(row, col) of every cell of 'A1', 'A1:B3' or 'A6:Q98/4' (columns A:Q
    of rows 6, 10, ..., 98)"""
    rng, _, step = rng.partition("/")
    step = int(step) if step else 1
    a, _, b = rng.partition(":")
    r0, c0 = parse_cell(a)
    r1, c1 = parse_cell(b) if b else (r0, c0)
    for r in range(r0, r1 + 1, step):
        for c in range(c0, c1 + 1):
            yield r, c


def iter_ranges(ranges):
    for rng in ranges.split(","):
        yield from iter_range(rng)


def coordinate(row, col):
    return f"{get_column_letter(col)}{row}"


# -- formula fills (exactly the encoder's semantics: openrsxl's ftemplate) -------
def _ident(ch):
    return ("0" <= ch <= "9") or ("A" <= ch <= "Z") or ("a" <= ch <= "z") or ch in "_.$"


def _template(f, row, col):
    """Literal text and relative / absolute references of `f` at (row, col)"""
    parts = []
    n = len(f)
    lit = 0
    i = 0
    while i < n:
        ch = f[i]
        if ch == '"' or ch == "'":
            q = ch
            i += 1
            while i < n:
                if f[i] == q:
                    if i + 1 < n and f[i + 1] == q:
                        i += 2
                        continue
                    break
                i += 1
            i += 1
            continue
        if (ch == "$" or "A" <= ch <= "Z") and (i == 0 or not _ident(f[i - 1])):
            j = i
            col_abs = ch == "$"
            if col_abs:
                j += 1
            ls = j
            while j < n and j - ls < 4 and "A" <= f[j] <= "Z":
                j += 1
            le = j
            row_abs = j < n and f[j] == "$"
            if row_abs:
                j += 1
            ds = j
            while j < n and j - ds < 8 and "0" <= f[j] <= "9":
                j += 1
            de = j
            if (
                1 <= le - ls <= 3
                and 1 <= de - ds <= 7
                and f[ds] != "0"
                and (j >= n or not (_ident(f[j]) or f[j] in "(!"))
            ):
                if lit < i:
                    parts.append(f[lit:i])
                c = column_index_from_string(f[ls:le])
                r = int(f[ds:de])
                parts.append((col_abs, c if col_abs else c - col, row_abs, r if row_abs else r - row))
                i = j
                lit = i
                continue
        i += 1
    if lit < n:
        parts.append(f[lit:])
    return parts


def fill_formula(formula, row, col, row2, col2):
    """The formula written for (row, col) as it reads in (row2, col2): its
    relative references shifted like an Excel fill (None: off the sheet)"""
    out = []
    for p in _template(formula, row, col):
        if isinstance(p, str):
            out.append(p)
            continue
        col_abs, cv, row_abs, rv = p
        c = cv if col_abs else col2 + cv
        r = rv if row_abs else row2 + rv
        if not 1 <= c <= 18278 or r < 1:
            return None
        out.append(("$" if col_abs else "") + get_column_letter(c) + ("$" if row_abs else "") + str(r))
    return "".join(out)


# -- decoded model ------------------------------------------------------------------
class StyleDef:
    """A cell style of the styles section (openpyxl objects)"""

    __slots__ = (
        "font",
        "fill",
        "border",
        "number_format",
        "alignment",
        "protection",
        "quotePrefix",
        "pivotButton",
        "style",
        "xfId",
    )

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))

    def __repr__(self):
        return f"<StyleDef numFmt={self.number_format!r} style={self.style!r}>"


class DecodedSheet:
    """
    One sheet of a decoded document. Cells are keyed by (row, column):

    * ``values``: the values (data_only view: formula results), Python
      objects as openpyxl returns them; ``data_types``: their openpyxl
      data_type ('n', 's', 'b', 'e', 'd')
    * ``formulas``: formula text (``=...``), ``ArrayFormula`` or
      ``DataTableFormula``
    * ``style_ids``: the style of each non-default cell; ``styles``: style
      id -> StyleDef
    * ``merged`` (range strings), ``hyperlinks`` / ``comments`` (per cell),
      ``column_dimensions`` / ``row_dimensions`` (dicts of fields),
      ``validations``, ``conditional_formatting`` ([(sqref, Rule)]),
      ``tables``, ``sheet`` (attribute -> openpyxl object), print titles /
      area
    * ``filter``: what the encoder left out of the cell sections (and of
      hyperlinks / comments): ``"skip=hidden"`` (hidden rows and columns),
      ``"area=..."`` (cells outside the print area); [] = nothing
    """

    def __init__(self, index, name, dimension=None):
        self.index = index
        self.name = name
        self.dimension = dimension
        self.filter = []
        self.analysis = None
        self.encoding = None
        self.values = {}
        self.data_types = {}
        self.formulas = {}
        self.style_ids = {}
        self.styles = {}
        self.merged = []
        self.hyperlinks = {}
        self.comments = {}
        self.column_dimensions = {}
        self.row_dimensions = {}
        self.validations = []
        self.conditional_formatting = []
        self.tables = []
        self.sheet = {}
        self.print_title_rows = None
        self.print_title_cols = None
        self.print_area = ""
        self.sections = []

    def cell(self, ref):
        """dict(value, formula, style) of a cell ('B5' or (row, col))"""
        key = parse_cell(ref) if isinstance(ref, str) else tuple(ref)
        sid = self.style_ids.get(key, 0)
        return {
            "value": self.values.get(key),
            "data_type": self.data_types.get(key),
            "formula": self.formulas.get(key),
            "style_id": sid,
            "style": self.styles.get(sid),
            "hyperlink": self.hyperlinks.get(key),
            "comment": self.comments.get(key),
        }

    def __repr__(self):
        return f"<DecodedSheet {self.name!r} {len(self.values)} values {len(self.formulas)} formulas>"


class DecodedWorkbook:
    """A decoded document: ``sheets`` (DecodedSheet, also by name) and the
    workbook section (``sheet_entries``, ``defined_names``, ``views``,
    ``calculation``, ``security``, ``properties``, ``custom_doc_props``,
    ``date1904``, ``code_name``, ``active``)"""

    def __init__(self):
        self.format = FORMAT
        self.source = None
        self.included = None
        self.legend = None
        self.sheets = []
        self.sheet_entries = []
        self.defined_names = []
        self.views = []
        self.calculation = None
        self.security = None
        self.properties = None
        self.custom_doc_props = None
        self.date1904 = False
        self.code_name = None
        self.active = 0
        self.styles = {}  # the style table (## styles): id -> StyleDef

    def __getitem__(self, name):
        for s in self.sheets:
            if s.name == name:
                return s
        raise KeyError(name)

    @property
    def sheetnames(self):
        return [s.name for s in self.sheets]

    def to_workbook(self):
        """An openpyxl (openrsxl) Workbook rebuilt from the document: values
        (formulas where there is one), styles and every included section"""
        from ._rebuild import to_workbook

        return to_workbook(self)

    def __repr__(self):
        return f"<DecodedWorkbook {self.source!r} {self.sheetnames}>"


# -- sections ---------------------------------------------------------------------
def _values(sheet, encoding, start, text):
    sheet.encoding = encoding
    start_col = column_index_from_string(start) if start else 1
    row = None  # row of the previous grid line
    above = {}  # its tokens by column (ditto marks)
    for line in text.split("\n"):
        if not line:
            continue
        # (hybrid: the block lines come first, and the first grid line is
        # labelled; every line after it is a grid line)
        if encoding in ("grid", "hybrid") and (row is not None or _LABEL.match(line)):
            m = _LABEL.match(line)
            row = int(m.group(1)) if m else row + 1
            cells = []
            for r, c, tok in parse_grid_line(line, start_col, row):
                if tok == "^":  # the token of the line above, same column
                    if c not in above:
                        raise DecodeError(f"ditto mark without a value above it at {coordinate(r, c)}")
                    tok = above[c]
                cells.append((r, c, tok))
            above = {c: tok for _, c, tok in cells}
        else:
            ranges, sep, tok = line.partition("=")
            if not sep:
                raise DecodeError(f"bad values line {line[:40]!r}")
            cells = ((r, c, tok) for r, c in iter_ranges(ranges))
        for r, c, tok in cells:
            v, dt = parse_token(tok)
            sheet.values[(r, c)] = v
            sheet.data_types[(r, c)] = dt


def _bracket_end(s, i):
    """index of the ']' closing the bracket opened before s[i]"""
    n = len(s)
    while i < n:
        if s[i] == '"':
            _, i = _decoder.raw_decode(s, i)
            continue
        if s[i] == "]":
            return i
        i += 1
    raise DecodeError("missing ']'")


_SPAN = r"[A-Z]{1,3}[0-9]+(?::[A-Z]{1,3}[0-9]+)?(?:/[0-9]+)?"
_HEAD = re.compile(rf"({_SPAN}(?:,{_SPAN})*)")


def _formulas(sheet, text):
    for line in text.split("\n"):
        if not line:
            continue
        m = _HEAD.match(line)
        if not m:
            raise DecodeError(f"bad formula line {line[:40]!r}")
        rng = m.group(1)
        i = m.end()
        kind = None
        attrs = ""
        if i < len(line) and line[i] == "[":
            j = _bracket_end(line, i + 1)
            kind, _, attrs = line[i + 1 : j].partition(" ")
            i = j + 1
        rest = line[i:]
        if rest.startswith("="):
            text_ = rest
        elif rest.startswith("~"):
            text_ = json.loads(rest[1:])
        else:
            text_ = None
        r0, c0 = parse_cell(rng.split(",")[0].split("/")[0].split(":")[0])
        if kind == "array":
            sheet.formulas[(r0, c0)] = ArrayFormula(attrs or None, text_)
        elif kind == "dataTable":
            el = _cx.decode_items("f", attrs)
            kw = dict(el.attrib)
            kw.setdefault("ref", None)
            sheet.formulas[(r0, c0)] = DataTableFormula(**kw)
        elif kind is not None:
            raise DecodeError(f"unknown formula kind {kind!r}")
        else:
            for r, c in iter_ranges(rng):
                if (r, c) == (r0, c0):
                    sheet.formulas[(r, c)] = text_
                else:
                    f = fill_formula(text_, r0, c0, r, c)
                    if f is None:
                        raise DecodeError(f"formula of {rng} cannot be filled to {coordinate(r, c)}")
                    sheet.formulas[(r, c)] = f


_STYLE_DEF = re.compile(r"s([0-9]+): ?")
_PART_DEF = re.compile(r"(font|fill|border) ([fpb])([0-9]+):(?: |$)")


def _styles(sheet, text, number_format_of=None):
    fonts, fills, borders = {}, {}, {}
    raw = {}
    for line in text.split("\n"):
        if not line:
            continue
        m = _PART_DEF.match(line)
        if m:
            kind, _, idx = m.groups()
            el = _cx.decode_items(kind, line[m.end() :])
            if kind == "font":
                fonts[int(idx)] = Font.from_tree(el)
            elif kind == "fill":
                fills[int(idx)] = Fill.from_tree(el)
            else:
                borders[int(idx)] = Border.from_tree(el)
            continue
        m = _STYLE_DEF.match(line)
        if m:
            raw[int(m.group(1))] = _cx.decode_items("xf", line[m.end() :])
            continue
        ranges, sep, sid = line.rpartition("=")
        if not sep or not sid.startswith("s"):
            raise DecodeError(f"bad styles line {line[:40]!r}")
        sid = int(sid[1:])
        for key in iter_ranges(ranges):
            sheet.style_ids[key] = sid

    def resolve(el, base):
        d = dict(base) if base else {}
        a = el.attrib
        if "font" in a:
            d["font"] = fonts[int(a["font"][1:])]
        if "fill" in a:
            d["fill"] = fills[int(a["fill"][1:])]
        if "border" in a:
            d["border"] = borders[int(a["border"][1:])]
        if "numFmt" in a:
            d["number_format"] = a["numFmt"]
        for k in ("quotePrefix", "pivotButton"):
            if k in a:
                d[k] = bool(int(a[k]))
        if "style" in a:
            d["style"] = a["style"]
            d["xfId"] = None
        if "xfId" in a:
            d["xfId"] = int(a["xfId"])
            d["style"] = None
        for child in el:
            if child.tag == "alignment":
                d["alignment"] = Alignment.from_tree(child)
            elif child.tag == "protection":
                d["protection"] = Protection.from_tree(child)
        return d

    base = {"quotePrefix": False, "pivotButton": False}
    if 0 in raw:
        base = resolve(raw[0], base)
    for sid, el in raw.items():
        sheet.styles[sid] = StyleDef(**(base if sid == 0 else resolve(el, base)))


def _split_ranges(line):
    ranges, _, rest = line.partition(" ")
    return ranges, rest


def _hyperlinks(sheet, text):
    for line in text.split("\n"):
        if not line:
            continue
        ranges, items = _split_ranges(line)
        attrs = dict(_cx.decode_items("hyperlink", items).attrib)
        for r, c in iter_ranges(ranges):
            sheet.hyperlinks[(r, c)] = Hyperlink(ref=coordinate(r, c), **attrs)


def _number(raw):
    try:
        return int(raw)
    except ValueError:
        return float(raw)


def _comments(sheet, text):
    for line in text.split("\n"):
        if not line:
            continue
        ranges, items = _split_ranges(line)
        el = _cx.decode_items("comment", items)
        a = el.attrib
        for r, c in iter_ranges(ranges):
            sheet.comments[(r, c)] = Comment(
                el.text,
                a.get("author"),
                height=_number(a["height"]) if "height" in a else 79,
                width=_number(a["width"]) if "width" in a else 144,
            )


def _dimensions(sheet, text):
    col_fields = dict(COLUMN_FIELDS)
    row_fields = dict(ROW_FIELDS)
    for line in text.split("\n"):
        if not line:
            continue
        el = _cx.decode(line)
        a = dict(el.attrib)
        style = a.pop("s", None)
        if el.tag == "col":
            d = dict(col_fields)
            index = a.pop("index", None)
            for k, v in a.items():
                d[k] = parse_field(k, v)
            d["style_id"] = int(style[1:]) if style else None
            key = index or get_column_letter(d["min"])
            sheet.column_dimensions[key] = d
        elif el.tag == "row":
            span = a.pop("r")
            lo, _, hi = span.partition(":")
            d = dict(row_fields)
            for k, v in a.items():
                d[k] = parse_field(k, v)
            d["style_id"] = int(style[1:]) if style else None
            for r in range(int(lo), int(hi or lo) + 1):
                sheet.row_dimensions[r] = dict(d)
        else:
            raise DecodeError(f"bad dimensions line {line[:40]!r}")


def _validations(sheet, text):
    for line in text.split("\n"):
        if line:
            sheet.validations.append(DataValidation.from_tree(_cx.decode(line)))


def _conditional(sheet, text):
    for line in text.split("\n"):
        if not line:
            continue
        ranges, item = _split_ranges(line)
        el = _cx.decode(item)
        dxf = None
        for child in list(el):
            if child.tag == "dxf":
                dxf = DifferentialStyle.from_tree(child)
                el.remove(child)
        rule = Rule.from_tree(el)
        rule.dxf = dxf
        sheet.conditional_formatting.append((ranges.replace(",", " "), rule))


def _tables(sheet, text):
    for line in text.split("\n"):
        if line:
            sheet.tables.append(Table.from_tree(_cx.decode(line)))


def _sheet(sheet, text):
    for line in text.split("\n"):
        if not line:
            continue
        el = _cx.decode(line)
        if el.tag == "printTitles":
            sheet.print_title_rows = el.get("rows")
            sheet.print_title_cols = el.get("cols")
        elif el.tag == "printArea":
            sheet.print_area = el.get("val")
        elif el.tag in SHEET_TAGS:
            attr, cls = SHEET_TAGS[el.tag]
            sheet.sheet[attr] = read_object(cls, el)
        else:
            raise DecodeError(f"unknown sheet setting {el.tag!r}")


_CORE = "{%s}coreProperties" % _cx.PREFIXES["cp"]


def _workbook(book, text):
    for line in text.split("\n"):
        if not line:
            continue
        el = _cx.decode(line)
        tag = el.tag
        if tag == "sheet":
            book.sheet_entries.append(dict(el.attrib))
        elif tag == "workbook":
            book.date1904 = el.get("date1904") == "1"
            book.code_name = el.get("codeName")
        elif tag == "active":
            book.active = int(el.get("val"))
        elif tag == "workbookView":
            book.views.append(BookView.from_tree(el))
        elif tag == "calcPr":
            book.calculation = CalcProperties.from_tree(el)
        elif tag == "workbookProtection":
            book.security = WorkbookProtection.from_tree(el)
        elif tag == "definedName":
            book.defined_names.append(DefinedName.from_tree(el))
        elif tag == _CORE:
            book.properties = DocumentProperties.from_tree(el)
        elif tag == "Properties" or tag.endswith("}Properties"):
            book.custom_doc_props = CustomPropertyList.from_tree(el)
        else:
            raise DecodeError(f"unknown workbook item {tag!r}")


def _section(book, sheet, name, encoding, start, text):
    if name == "values":
        _values(sheet, encoding, start, text)
    elif name == "formulas":
        _formulas(sheet, text)
    elif name == "styles":
        _styles(sheet, text)
    elif name == "merged":
        sheet.merged = [r for r in text.strip().split(",") if r]
    elif name == "hyperlinks":
        _hyperlinks(sheet, text)
    elif name == "comments":
        _comments(sheet, text)
    elif name == "dimensions":
        _dimensions(sheet, text)
    elif name == "validations":
        _validations(sheet, text)
    elif name == "conditional_formatting":
        _conditional(sheet, text)
    elif name == "tables":
        _tables(sheet, text)
    elif name == "sheet":
        _sheet(sheet, text)
    elif name == "analysis":
        sheet.analysis = text
    else:
        raise DecodeError(f"unknown section {name!r}")
    sheet.sections.append(name)


# -- documents --------------------------------------------------------------------
_SHEET_HEAD = re.compile(
    r'## sheet ([0-9]+) ("(?:[^"\\]|\\.)*")(?: dim=(\S+))?((?: (?:skip|area|rows|cols|cells)=\S+)*)\Z'
)


def _parse_text(text):
    book = DecodedWorkbook()
    lines = text.split("\n")
    if not lines or not lines[0].startswith("#" + FORMAT):
        raise DecodeError(f"not an {FORMAT} document")
    m = re.search(r' source=("(?:[^"\\]|\\.)*")', lines[0])
    if m:
        book.source = json.loads(m.group(1))
    m = re.search(r" sections=(\S*)", lines[0])
    if m:
        book.included = [x for x in m.group(1).split(",") if x]
    sheet = None
    cur = None  # (kind, name, encoding, start, lines)

    def close():
        if cur is None:
            return
        kind, name, enc, start, body = cur
        body = "\n".join(body)
        if kind == "legend":
            book.legend = body
        elif kind == "workbook":
            _workbook(book, body)
        elif kind == "styles":
            _style_table(book, body)
        else:
            _section(book, sheet, name, enc, start, body)

    for line in lines[1:]:
        if line.startswith("## "):
            close()
            cur = None
            if line == "## legend":
                cur = ("legend", None, None, None, [])
            elif line == "## workbook":
                cur = ("workbook", None, None, None, [])
            elif line == "## styles":
                cur = ("styles", None, None, None, [])
            else:
                m = _SHEET_HEAD.match(line)
                if not m:
                    raise DecodeError(f"bad header {line!r}")
                sheet = DecodedSheet(int(m.group(1)), json.loads(m.group(2)), m.group(3))
                sheet.filter = m.group(4).split()
                book.sheets.append(sheet)
            continue
        if line.startswith("### "):
            close()
            if sheet is None:
                raise DecodeError(f"section outside a sheet: {line!r}")
            parts = line[4:].split(" ")
            name = parts[0]
            enc = start = None
            for p in parts[1:]:
                if p.startswith("start="):
                    start = p[6:]
                else:
                    enc = p
            cur = ("section", name, enc, start, [])
            continue
        if cur is not None:
            cur[4].append(line)
        elif line:
            raise DecodeError(f"text outside a section: {line[:40]!r}")
    close()
    _share_styles(book)
    return book


def _style_table(book, text):
    """the workbook's ## styles: style parts and cell styles"""
    table = DecodedSheet(0, "")
    _styles(table, text)
    if table.style_ids:
        raise DecodeError("cell ranges in the workbook style table")
    book.styles.update(table.styles)


def _share_styles(book):
    """every sheet sees the workbook's style table"""
    for sheet in book.sheets:
        for sid, sd in book.styles.items():
            sheet.styles.setdefault(sid, sd)


def _parse_dict(d):
    if d.get("format") != FORMAT:
        raise DecodeError(f"not an {FORMAT} document")
    book = DecodedWorkbook()
    book.source = d.get("source")
    book.included = d.get("sections")
    book.legend = d.get("legend")
    if d.get("workbook") is not None:
        _workbook(book, d["workbook"])
    if d.get("styles") is not None:
        _style_table(book, d["styles"])
    for s in d.get("sheets", []):
        sheet = DecodedSheet(s["index"], s["name"], s.get("dimension"))
        sheet.filter = list(s.get("filter") or [])
        book.sheets.append(sheet)
        for name, sec in s["sections"].items():
            _section(book, sheet, name, sec.get("encoding"), sec.get("start"), sec["text"])
        if s.get("analysis") is not None:
            sheet.analysis = s["analysis"]
    _share_styles(book)
    return book


def decompress(source):
    """
    Decode an openrsxl-llm document: a CompressedWorkbook, its text, its
    JSON (str or dict) or the path of a .txt / .json file. Returns a
    DecodedWorkbook.
    """
    from ._encode import CompressedWorkbook

    if isinstance(source, CompressedWorkbook):
        return _parse_dict(source.to_dict())
    if isinstance(source, dict):
        return _parse_dict(source)
    if isinstance(source, os.PathLike) or (
        isinstance(source, str) and not source.startswith(("#", "{")) and os.path.exists(source)
    ):
        with open(source, encoding="utf-8") as f:
            source = f.read()
    if isinstance(source, bytes):
        source = source.decode("utf-8")
    if source.lstrip().startswith("{"):
        return _parse_dict(json.loads(source))
    return _parse_text(source)
