"""compress_workbook: workbook -> openrsxl-llm document (see __init__)"""

import bisect
import json
import os
import warnings
from collections import OrderedDict

from openrsxl import _native
from openrsxl.styles.numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_MAX_SIZE
from openrsxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries
from openrsxl.utils.datetime import CALENDAR_MAC_1904
from openrsxl.workbook.properties import CalcProperties
from openrsxl.workbook.protection import WorkbookProtection
from openrsxl.worksheet._reader import WorkSheetParser
from openrsxl.worksheet.formula import ArrayFormula, DataTableFormula
from openrsxl.xml.constants import ARC_CORE
from openrsxl.xml.functions import tostring

from . import _cx
from ._legend import FORMAT, legend
from ._model import SHEET_OBJECTS, column_dimension_items, read_object, row_dimension_items

#: sections of a sheet, in document order
SHEET_SECTIONS = (
    "analysis",
    "values",
    "formulas",
    "styles",
    "merged",
    "hyperlinks",
    "comments",
    "dimensions",
    "validations",
    "conditional_formatting",
    "tables",
    "sheet",
)

#: every section name accepted by `include` / `exclude`
SECTIONS = SHEET_SECTIONS + ("workbook",)

GROUPS = {
    "all": SECTIONS,
    "cells": ("values", "formulas", "styles"),
    "metadata": (
        "merged",
        "hyperlinks",
        "comments",
        "dimensions",
        "validations",
        "conditional_formatting",
        "tables",
        "sheet",
        "workbook",
    ),
}

OUTPUTS = ("memory", "text", "json")

ENCODINGS = ("auto", "grid", "index", "hybrid")

_MAX_ROW = 0xFFFFFFFF
_MAX_COL = 16384


class CellFilter:
    """The cells a sheet's document leaves out (``skip_hidden`` /
    ``print_area_only``): hidden rows and columns, cells outside the print
    area. ``skip_rows`` / ``skip_cols``: sorted, disjoint (first, last)
    ranges; ``areas``: (row1, col1, row2, col2) rectangles, None = the
    whole sheet."""

    def __init__(self, skip_rows=(), skip_cols=(), areas=None, items=()):
        self.skip_rows = _merge_ranges(skip_rows)
        self.skip_cols = _merge_ranges(skip_cols)
        self.areas = None if areas is None else list(areas)
        self.items = list(items)
        self._row_starts = [r[0] for r in self.skip_rows]
        self._col_starts = [c[0] for c in self.skip_cols]

    @property
    def active(self):
        return bool(self.skip_rows or self.skip_cols or self.areas is not None or self.items)

    @staticmethod
    def _in(ranges, starts, x):
        i = bisect.bisect_right(starts, x) - 1
        return i >= 0 and ranges[i][1] >= x

    def keeps(self, row, col):
        if self.skip_cols and self._in(self.skip_cols, self._col_starts, col):
            return False
        if self.skip_rows and self._in(self.skip_rows, self._row_starts, row):
            return False
        if self.areas is None:
            return True
        return any(r1 <= row <= r2 and c1 <= col <= c2 for r1, c1, r2, c2 in self.areas)

    def options(self):
        """options of the Rust SheetEncoder"""
        return dict(skip_rows=self.skip_rows, skip_cols=self.skip_cols, areas=self.areas)

    def header_items(self):
        """items of the sheet header: ``skip=hidden``, ``area=...``,
        ``rows=...``, ``cols=...``, ``cells=...``"""
        return list(self.items)


def _merge_ranges(ranges):
    out = []
    for a, b in sorted(ranges):
        if out and a <= out[-1][1] + 1:
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return [tuple(r) for r in out]


def _hidden_rows(ws):
    """(first, last) ranges of the hidden rows: rows marked hidden - and,
    with sheetFormatPr zeroHeight (rows hidden by default), every row
    without a visible row dimension"""
    hidden = []
    visible = []
    for idx, rd in ws.row_dimensions.items():
        (hidden if rd.hidden else visible).append((int(idx), int(idx)))
    fmt = getattr(ws, "sheet_format", None)
    if fmt is not None and fmt.zeroHeight:
        hidden = []
        prev = 0
        for a, b in _merge_ranges(visible):
            if a > prev:
                hidden.append((prev, a - 1))
            prev = b + 1
        if prev <= _MAX_ROW:
            hidden.append((prev, _MAX_ROW))
    return _merge_ranges(hidden)


def _hidden_cols(ws):
    out = []
    for key, cd in ws.column_dimensions.items():
        if not cd.hidden:
            continue
        lo = cd.min if cd.min is not None else column_index_from_string(key)
        hi = cd.max if cd.max is not None else lo
        out.append((max(1, lo), min(_MAX_COL, hi)))
    return _merge_ranges(out)


def _print_areas(ws):
    """(rectangles, text) of the sheet's print area and print titles (title
    rows across the columns of the area, title columns across its rows),
    or (None, None) when the sheet has no print area"""
    pa = getattr(ws, "_print_area", None)
    ranges = list(pa.sorted()) if pa is not None and pa.ranges else []
    if not ranges:
        return None, None
    rects = [(r.min_row or 1, r.min_col or 1, r.max_row or _MAX_ROW, r.max_col or _MAX_COL) for r in ranges]
    text = ",".join(str(r.coord) for r in ranges)
    r_lo, c_lo = min(r[0] for r in rects), min(r[1] for r in rects)
    r_hi, c_hi = max(r[2] for r in rects), max(r[3] for r in rects)
    titles = []
    rows = getattr(ws, "_print_rows", None)
    if rows is not None and rows.min_row is not None:
        rects.append((rows.min_row, c_lo, rows.max_row or rows.min_row, c_hi))
        titles.append(str(rows).replace("$", ""))
    cols = getattr(ws, "_print_cols", None)
    if cols is not None and cols.min_col is not None:
        lo = column_index_from_string(cols.min_col)
        hi = column_index_from_string(cols.max_col or cols.min_col)
        rects.append((r_lo, lo, r_hi, hi))
        titles.append(str(cols).replace("$", ""))
    if titles:
        text += "+titles:" + ",".join(titles)
    return rects, text


# -- selection of rows, columns and ranges ------------------------------------------
def _row_number(x):
    t = str(x).strip().replace("$", "")
    if not t.isdigit():
        raise ValueError(f"bad row {x!r}")
    return int(t)


def _col_number(x):
    if isinstance(x, int) and not isinstance(x, bool):
        return x
    t = str(x).strip().replace("$", "").upper()
    if t.isdigit():
        return int(t)
    if not t.isalpha() or len(t) > 3:
        raise ValueError(f"bad column {x!r}")
    return column_index_from_string(t)


def _segments(spec, what, conv, hi):
    """sorted, merged (first, last) segments of a rows / columns spec: 5,
    "5", "5:9", "A:D", "1:10,20:30", (first, last), or a list of these"""
    if spec is None:
        return None
    out = []

    def check(a, b, item):
        if not 1 <= a <= b <= hi:
            raise ValueError(f"bad {what} selection {item!r}: needs 1 <= first <= last <= {hi}")
        out.append((a, b))

    def add(item):
        if isinstance(item, bool):
            raise TypeError(f"bad {what} selection {item!r}")
        if isinstance(item, int):
            check(item, item, item)
        elif isinstance(item, str):
            parts = [p for p in item.replace(" ", ",").split(",") if p]
            if not parts:
                raise ValueError(f"empty {what} selection")
            for part in parts:
                ends = part.split(":")
                if len(ends) > 2 or not all(ends):
                    raise ValueError(f"bad {what} selection {part!r}")
                check(conv(ends[0]), conv(ends[-1]), part)
        elif isinstance(item, tuple) and len(item) == 2:
            check(conv(item[0]), conv(item[1]), item)
        else:
            raise TypeError(f"bad {what} selection {item!r}")

    if isinstance(spec, (list, set, frozenset)):
        if not spec:
            raise ValueError(f"empty {what} selection")
        for item in spec:
            add(item)
    else:
        add(spec)
    return _merge_ranges(out)


def _complement(segments, lo, hi):
    """the (first, last) ranges of lo..hi outside `segments`"""
    out, nxt = [], lo
    for a, b in segments:
        if a > nxt:
            out.append((nxt, a - 1))
        nxt = max(nxt, b + 1)
    if nxt <= hi:
        out.append((nxt, hi))
    return out


def _rect_text(r1, c1, r2, c2):
    if r1 == 1 and r2 >= _MAX_ROW:
        a, b = get_column_letter(c1), get_column_letter(c2)
        return f"{a}:{b}"
    if c1 == 1 and c2 == _MAX_COL:
        return f"{r1}:{r2}"
    a = f"{get_column_letter(c1)}{r1}"
    b = f"{get_column_letter(c2)}{r2}"
    return a if a == b else f"{a}:{b}"


def _range_rects(spec):
    """(row1, col1, row2, col2) rectangles of a ranges spec: "A1:D20",
    "A1:D20,F5:H9", "B:D" (whole columns), "3:7" (whole rows), "C4", or a
    list of these"""
    if spec is None:
        return None
    items = list(spec) if isinstance(spec, (list, tuple, set, frozenset)) else [spec]
    if not items:
        raise ValueError("empty ranges selection")
    out = []
    for item in items:
        if not isinstance(item, str):
            raise TypeError(f"bad range {item!r}: ranges are strings such as 'A1:D20'")
        parts = [p for p in item.replace(" ", ",").split(",") if p]
        if not parts:
            raise ValueError("empty ranges selection")
        for part in parts:
            if "!" in part:
                raise ValueError(f"bad range {part!r}: give ranges per sheet with a dict {{sheet name: ranges}}")
            try:
                c1, r1, c2, r2 = range_boundaries(part.replace("$", "").upper())
            except (ValueError, TypeError) as e:
                raise ValueError(f"bad range {part!r}") from e
            r1, r2 = (r1 or 1), (r2 or _MAX_ROW)
            c1, c2 = (c1 or 1), (c2 or _MAX_COL)
            if not (1 <= r1 <= r2 and 1 <= c1 <= c2 <= _MAX_COL):
                raise ValueError(f"bad range {part!r}")
            out.append((r1, c1, r2, c2))
    return out


def _intersect(a, b):
    """the parts of the rectangles `a` inside the rectangles `b`"""
    out = []
    for r1, c1, r2, c2 in a:
        for s1, d1, s2, d2 in b:
            t = (max(r1, s1), max(c1, d1), min(r2, s2), min(c2, d2))
            if t[0] <= t[2] and t[1] <= t[3]:
                out.append(t)
    return out


class Selection:
    """The rows / columns / ranges options: one spec for every sheet, or a
    dict {sheet name: spec} (sheets not in it: not restricted). Parsed when
    created, so bad input fails before any sheet is read."""

    KEYS = ("rows", "columns", "ranges")

    def __init__(self, rows=None, columns=None, ranges=None):
        self.raw = dict(rows=rows, columns=columns, ranges=ranges)
        for key, spec in self.raw.items():
            for one in spec.values() if isinstance(spec, dict) else [spec]:
                self._parse(key, one)

    @property
    def given(self):
        return any(v is not None for v in self.raw.values())

    def sheet_names(self):
        """sheet names used as keys of dict options"""
        names = set()
        for spec in self.raw.values():
            if isinstance(spec, dict):
                names.update(spec)
        return names

    @staticmethod
    def _parse(key, spec):
        if key == "rows":
            return _segments(spec, "row", _row_number, _MAX_ROW)
        if key == "columns":
            return _segments(spec, "column", _col_number, _MAX_COL)
        return _range_rects(spec)

    def for_sheet(self, name):
        """(row segments, column segments, rectangles) of a sheet, None
        where not restricted"""
        out = []
        for key in self.KEYS:
            spec = self.raw[key]
            if isinstance(spec, dict):
                spec = spec.get(name)
            out.append(self._parse(key, spec))
        return tuple(out)


def _cell_filter(ws, skip_hidden, print_area_only, selection=None):
    """CellFilter of a sheet: skip_hidden, print_area_only and the rows /
    columns / ranges selection (a cell is kept when every one keeps it)"""
    items = []
    skip_rows = _hidden_rows(ws) if skip_hidden else []
    skip_cols = _hidden_cols(ws) if skip_hidden else []
    if skip_rows or skip_cols:
        items.append("skip=hidden")
    areas = None
    if print_area_only:
        areas, text = _print_areas(ws)
        if areas is not None:
            items.append(f"area={text}")
    if selection is not None and selection.given:
        rows, cols, rects = selection.for_sheet(ws.title)
        if rows is not None:
            skip_rows = list(skip_rows) + _complement(rows, 0, _MAX_ROW)
            items.append("rows=" + ",".join(f"{a}:{b}" for a, b in rows))
        if cols is not None:
            skip_cols = list(skip_cols) + _complement(cols, 1, _MAX_COL)
            items.append("cols=" + ",".join(f"{get_column_letter(a)}:{get_column_letter(b)}" for a, b in cols))
        if rects is not None:
            areas = rects if areas is None else _intersect(areas, rects)
            items.append("cells=" + ",".join(_rect_text(*r) for r in rects))
    return CellFilter(skip_rows, skip_cols, areas, items)


def _resolve_sections(include, exclude):
    def expand(spec):
        if spec is None:
            return set()
        if isinstance(spec, str):
            spec = [spec]
        out = set()
        for s in spec:
            if s in GROUPS:
                out.update(GROUPS[s])
            elif s in SECTIONS:
                out.add(s)
            else:
                raise ValueError(f"unknown section {s!r}; sections: {', '.join(SECTIONS)}; groups: {', '.join(GROUPS)}")
        return out

    chosen = expand(include) - expand(exclude)
    return tuple(s for s in SECTIONS if s in chosen)


def _resolve_outputs(output):
    if isinstance(output, str):
        output = [output]
    out = set()
    for o in output:
        if o == "all":
            out.update(OUTPUTS)
        elif o in OUTPUTS:
            out.add(o)
        else:
            raise ValueError(f"unknown output {o!r}; one or more of {', '.join(OUTPUTS)} or 'all'")
    if not out:
        raise ValueError("no output")
    return out


# -- results -------------------------------------------------------------------
class Section:
    """One section of a sheet (or the workbook): its header line and text"""

    __slots__ = ("name", "encoding", "start", "text", "tokens", "size")

    def __init__(self, name, text, encoding=None, start=None, tokens=None, size=None):
        self.name = name
        self.text = text
        self.encoding = encoding
        self.start = start
        self.tokens = tokens
        self.size = size

    @property
    def header(self):
        h = f"### {self.name}"
        if self.encoding:
            h += f" {self.encoding}"
        if self.start:
            h += f" start={self.start}"
        return h

    def to_dict(self):
        d = {}
        if self.encoding:
            d["encoding"] = self.encoding
        if self.start:
            d["start"] = self.start
        d["text"] = self.text
        return d

    def __repr__(self):
        return f"<Section {self.header!r} ~{self.tokens} tokens>"


class CompressedSheet:
    """A compressed worksheet: ``sections`` maps section names to Section"""

    def __init__(self, index, name, dimension, analysis):
        self.index = index
        self.name = name
        self.dimension = dimension
        self.analysis = analysis
        self.sections = OrderedDict()
        self.used_styles = set()  # style ids of its cells / dimensions
        #: what the sheet's cell sections leave out: ["skip=hidden", "area=..."]
        self.filter = []

    @property
    def header(self):
        h = f"## sheet {self.index} {json.dumps(self.name, ensure_ascii=False)}"
        if self.dimension:
            h += f" dim={self.dimension}"
        for item in self.filter:
            h += " " + item
        return h

    def text(self, sections=None):
        """The sheet as text (only the given sections if not None)"""
        parts = [self.header]
        for name, sec in self.sections.items():
            if sections is not None and name not in sections:
                continue
            if sec.text is None:
                raise ValueError(f"section {name!r} of sheet {self.name!r} was not kept in memory")
            parts.append(sec.header)
            if sec.text:
                parts.append(sec.text.rstrip("\n"))
        return "\n".join(parts) + "\n"

    def to_dict(self, sections=None):
        return {
            "index": self.index,
            "name": self.name,
            "dimension": self.dimension,
            **({"filter": list(self.filter)} if self.filter else {}),
            "analysis": self.analysis,
            "sections": {n: s.to_dict() for n, s in self.sections.items() if sections is None or n in sections},
        }

    def __getitem__(self, name):
        return self.sections[name]

    def __repr__(self):
        return f"<CompressedSheet {self.name!r} {list(self.sections)}>"


class CompressedWorkbook:
    """
    Result of ``compress_workbook``.

    ``text()`` / ``to_json()`` assemble the document (or a part of it:
    ``sections=``, ``sheets=``); ``sheets`` lists the CompressedSheet objects
    (also by name: ``result["Sheet1"]``); ``files`` the files written;
    ``tokens`` the estimated LLM tokens of the whole document.
    """

    def __init__(self, source, sections, legend_text, workbook, styles=None):
        self.format = FORMAT
        self.source = source
        self.included = sections
        self.legend = legend_text
        self.workbook = workbook  # Section or None
        self.styles = styles  # StyleTable or None
        self.sheets = []
        self.files = {}
        self.in_memory = True

    def __getitem__(self, name):
        for s in self.sheets:
            if s.name == name:
                return s
        raise KeyError(name)

    def _sheets(self, sheets):
        if sheets is None:
            return self.sheets
        if isinstance(sheets, str):
            sheets = [sheets]
        return [self[s] for s in sheets]

    def header(self):
        """first line: format, source and the included sections (an included
        section absent from a sheet is empty)"""
        h = f"#{FORMAT}"
        if self.source:
            h += f" source={json.dumps(self.source, ensure_ascii=False)}"
        return h + " sections=" + ",".join(self.included)

    def text(self, sections=None, sheets=None, legend=None):
        """
        The document as text. ``sections``: names of the sections to keep;
        ``sheets``: names of the sheets to keep; ``legend``: include the
        legend (default: if the document was made with one).
        """
        if not self.in_memory:
            raise ValueError("the document was not kept in memory (output without 'memory')")
        if isinstance(sections, str):
            sections = [sections]
        parts = [self.header()]
        if legend is None:
            legend = bool(self.legend)
        if legend and self.legend:
            parts.append("## legend")
            parts.append(self.legend)
        if self.workbook is not None and (sections is None or "workbook" in sections):
            parts.append("## workbook")
            if self.workbook.text:
                parts.append(self.workbook.text.rstrip("\n"))
        table = self._style_text(sections, sheets)
        if table is not None:
            parts.append("## styles")
            parts.append(table.rstrip("\n"))
        text = "\n".join(parts) + "\n"
        return text + "".join(s.text(sections) for s in self._sheets(sheets))

    def _style_text(self, sections, sheets):
        """the style table of these sheets (only the styles they use)"""
        if self.styles is None or (sections is not None and "styles" not in sections):
            return None
        if sheets is None:
            return self.styles.text()
        return self.styles.text(set().union(*(s.used_styles for s in self._sheets(sheets))))

    def to_dict(self, sections=None, sheets=None):
        if not self.in_memory:
            raise ValueError("the document was not kept in memory (output without 'memory')")
        d = {"format": FORMAT, "source": self.source, "sections": list(self.included)}
        if self.legend:
            d["legend"] = self.legend
        if self.workbook is not None and (sections is None or "workbook" in sections):
            d["workbook"] = self.workbook.text
        table = self._style_text(sections, sheets)
        if table is not None:
            d["styles"] = table
        d["sheets"] = [s.to_dict(sections) for s in self._sheets(sheets)]
        return d

    def to_json(self, sections=None, sheets=None, indent=None):
        return json.dumps(self.to_dict(sections, sheets), ensure_ascii=False, indent=indent)

    @property
    def tokens(self):
        """Estimated LLM tokens of the text document"""
        if self.in_memory:
            return estimate_tokens(self.text())
        total = estimate_tokens(self.header() + (self.legend or ""))
        if self.workbook is not None:
            total += self.workbook.tokens or 0
        if self.styles is not None:
            total += estimate_tokens(self.styles.text())
        for s in self.sheets:
            total += sum(sec.tokens or 0 for sec in s.sections.values())
        return total

    def __repr__(self):
        return f"<CompressedWorkbook {self.source!r} {len(self.sheets)} sheets ~{self.tokens} tokens>"


def estimate_tokens(text):
    """Estimated number of LLM tokens of `text` (BPE tokenizers such as
    o200k / cl100k; exact counts: ``count_tokens``)"""
    return _native._rs.llm_estimate_tokens(text)


def count_tokens(text, encoding="o200k_base"):
    """Exact token count with tiktoken (``pip install tiktoken``)"""
    import tiktoken

    return len(tiktoken.get_encoding(encoding).encode(text, disallowed_special=()))


# -- helpers -------------------------------------------------------------------
def _formula_info(obj):
    """Rust helper: description of a formula object"""
    if isinstance(obj, ArrayFormula):
        return ("array", None if obj.ref is None else str(obj.ref), str(obj.text if obj.text is not None else "="))
    if isinstance(obj, DataTableFormula):
        return ("table", _data_table_items(obj))
    return ("text", str(obj))


_DT_DEFAULTS = {
    "ref": None,
    "ca": False,
    "dt2D": False,
    "dtr": False,
    "r1": None,
    "r2": None,
    "del1": False,
    "del2": False,
}


def _data_table_items(obj):
    parts = []
    for k, default in _DT_DEFAULTS.items():
        v = getattr(obj, k, default)
        if v is default or v == default and type(v) is type(default):
            continue
        if v is True:
            v = "1"
        parts.append(f"{k}={_cx.scalar(v)}")
    return " ".join(parts)


def _rects(cells):
    """Exact rectangle cover of a set of (row, col), row-major order"""
    rows = {}
    for r, c in cells:
        rows.setdefault(r, []).append(c)
    rects = []
    open_ = {}  # (c0, c1) -> rect index, for the previous row
    prev = None
    for r in sorted(rows):
        cols = sorted(set(rows[r]))
        runs = []
        for c in cols:
            if runs and runs[-1][1] + 1 == c:
                runs[-1][1] = c
            else:
                runs.append([c, c])
        new_open = {}
        for c0, c1 in runs:
            idx = open_.get((c0, c1)) if prev is not None and r == prev + 1 else None
            if idx is None:
                rects.append([r, c0, r, c1])
                idx = len(rects) - 1
            else:
                rects[idx][2] = r
            new_open[(c0, c1)] = idx
        open_ = new_open
        prev = r
    return rects


def _spans(rects):
    """Single-row rectangles of the same columns repeating every d > 1 rows
    (at least 3 times) as strided spans [r0, c0, r1, c1, step]; the same
    rule as the Rust encoder. Sorted by top-left cell."""
    out = []
    rows = {}
    for r in rects:
        if r[0] == r[2]:
            rows.setdefault((r[1], r[3]), []).append(r[0])
        else:
            out.append([r[0], r[1], r[2], r[3], 1])
    for (c0, c1), rs in rows.items():
        i, n = 0, len(rs)
        while i < n:
            if i + 2 < n:
                d = rs[i + 1] - rs[i]
                if d > 1 and rs[i + 2] - rs[i + 1] == d:
                    j = i + 2
                    while j + 1 < n and rs[j + 1] - rs[j] == d:
                        j += 1
                    out.append([rs[i], c0, rs[j], c1, d])
                    i = j + 1
                    continue
            out.append([rs[i], c0, rs[i], c1, 1])
            i += 1
    out.sort(key=lambda s: (s[0], s[1]))
    return out


def _span_text(s):
    a = f"{get_column_letter(s[1])}{s[0]}"
    if s[2] != s[0] or s[3] != s[1]:
        a += f":{get_column_letter(s[3])}{s[2]}"
    if s[4] > 1:
        a += f"/{s[4]}"
    return a


def _ranges_text(cells):
    return ",".join(_span_text(s) for s in _spans(_rects(cells)))


def _xml(el):
    return tostring(el) if el is not None else None


def _is_default(obj, default):
    try:
        return _xml(obj.to_tree()) == _xml(default.to_tree())
    except Exception:
        return False


# -- sections -------------------------------------------------------------------
def _number_format(wb, fmt_id):
    if fmt_id < BUILTIN_FORMATS_MAX_SIZE:
        return BUILTIN_FORMATS.get(fmt_id, "General")
    return wb._number_formats[fmt_id - BUILTIN_FORMATS_MAX_SIZE]


class StyleTable:
    """
    The workbook's style table (``## styles``): the font / fill / border
    parts and the cell styles the encoded sheets use, once for the whole
    workbook (the sheets' ``### styles`` sections only map cells to style
    ids). ``text(ids)``: the lines that styles `ids` need (s0 always).
    """

    _PARTS = ("font", "fill", "border")

    def __init__(self, parts, cells):
        self.parts = parts  # {(kind, id): line}, in document order
        self.cells = cells  # {style id: (line, (fontId, fillId, borderId))}

    def text(self, ids=None):
        ids = sorted(self.cells) if ids is None else sorted((set(ids) | {0}) & set(self.cells))
        need = {(kind, self.cells[s][1][k]) for s in ids for k, kind in enumerate(self._PARTS)}
        lines = [line for key, line in self.parts.items() if key in need]
        lines += [self.cells[s][0] for s in ids]
        return "".join(line + "\n" for line in lines)

    def __repr__(self):
        return f"<StyleTable {len(self.cells)} styles>"


def _style_table(wb, used):
    """StyleTable of the cell styles `used` (and s0)"""
    styles = wb._cell_styles
    used = sorted(set(used) | {0})
    used = [i for i in used if i < len(styles)]
    s0 = styles[0]
    parts = OrderedDict()
    for attr, pool, prefix, kind in (
        ("fontId", wb._fonts, "f", "font"),
        ("fillId", wb._fills, "p", "fill"),
        ("borderId", wb._borders, "b", "border"),
    ):
        for i in sorted({getattr(styles[s], attr) for s in used}):
            items = _cx.encode_items(pool[i].to_tree())
            parts[(kind, i)] = f"{kind} {prefix}{i}: {items}" if items else f"{kind} {prefix}{i}:"
    names = wb._named_styles.names
    cells = OrderedDict()
    for s in used:
        sa = styles[s]
        items = []
        first = s == 0
        for attr, prefix, key in (("fontId", "f", "font"), ("fillId", "p", "fill"), ("borderId", "b", "border")):
            if first or getattr(sa, attr) != getattr(s0, attr):
                items.append(f"{key}={prefix}{getattr(sa, attr)}")
        if first or sa.numFmtId != s0.numFmtId:
            items.append(f"numFmt={_cx.scalar(_number_format(wb, sa.numFmtId))}")
        if first or sa.alignmentId != s0.alignmentId:
            items.append(_cx.encode(wb._alignments[sa.alignmentId].to_tree()))
        if first or sa.protectionId != s0.protectionId:
            items.append(_cx.encode(wb._protections[sa.protectionId].to_tree()))
        for attr in ("quotePrefix", "pivotButton"):
            v = getattr(sa, attr)
            if (first and v) or (not first and v != getattr(s0, attr)):
                items.append(f"{attr}={int(v)}")
        if first or sa.xfId != s0.xfId:
            if sa.xfId < len(names):
                items.append(f"style={_cx.scalar(names[sa.xfId])}")
            else:
                items.append(f"xfId={sa.xfId}")
        cells[s] = (f"s{s}: " + " ".join(items), (sa.fontId, sa.fillId, sa.borderId))
    return StyleTable(parts, cells)


def _merged_section(ws):
    return ",".join(str(r.coord) for r in ws.merged_cells.ranges) + "\n" if ws.merged_cells.ranges else ""


def _per_cell(ws, key, cell_filter=None):
    """{(row, col): object} of an overlay feature (hyperlink, comment), as
    full mode binds them (only cells the filter keeps)"""
    out = {}
    for row, cols in ws._overlay.items():
        for col, spec in cols.items():
            if cell_filter is not None and not cell_filter.keeps(row, col):
                continue
            if key == "links":
                links = spec.get("links")
                if links:
                    out[(row, col)] = links[-1][1]
            else:
                c = spec.get("comment")
                if c is not None:
                    out[(row, col)] = c
    return out


_LINK_ATTRS = ("target", "location", "tooltip", "display")


def _link_items(link):
    return " ".join(f"{a}={_cx.scalar(getattr(link, a))}" for a in _LINK_ATTRS if getattr(link, a, None) is not None)


def _grouped(per_cell, items_of):
    """lines 'RANGES items' grouping cells with identical items (in order
    of their first cell)"""
    groups = OrderedDict()
    for coord in sorted(per_cell):
        groups.setdefault(items_of(per_cell[coord]), []).append(coord)
    lines = []
    for items, cells in groups.items():
        lines.append(f"{_ranges_text(cells)} {items}" if items else _ranges_text(cells))
    # row-major order of the first cell of each line
    return "".join(line + "\n" for line in lines)


def _hyperlinks_section(ws, cell_filter=None):
    return _grouped(_per_cell(ws, "links", cell_filter), _link_items)


def _comment_items(c):
    items = []
    if c.author is not None:
        items.append(f"author={_cx.scalar(c.author)}")
    if c.width != 144:
        items.append(f"width={_cx.scalar(c.width)}")
    if c.height != 79:
        items.append(f"height={_cx.scalar(c.height)}")
    items.append(json.dumps(c.content if c.content is not None else "", ensure_ascii=False))
    return " ".join(items)


def _comments_section(ws, cell_filter=None):
    return _grouped(_per_cell(ws, "comments", cell_filter), _comment_items)


def _style_id(wb, style_array):
    """index of the first cell style equal to `style_array` (IndexedList's
    own index() can be stale)"""
    if style_array is None or not any(style_array):
        return None
    ids = wb.__dict__.get("_llm_style_ids")
    if ids is None:
        ids = {}
        for i, sa in enumerate(wb._cell_styles):
            ids.setdefault(tuple(sa), i)
        wb._llm_style_ids = ids
    return ids.get(tuple(style_array))


def _dimensions_section(ws, used_styles):
    wb = ws.parent
    lines = []
    for key, cd in ws.column_dimensions.items():
        sid = _style_id(wb, cd._style)
        if sid is not None:
            used_styles.add(sid)
        items = column_dimension_items(key, cd, sid)
        lines.append(f"col({items})")
    # rows: consecutive rows with identical settings in one line
    prev = None
    for idx in sorted(ws.row_dimensions):
        rd = ws.row_dimensions[idx]
        sid = _style_id(wb, rd._style)
        if sid is not None:
            used_styles.add(sid)
        items = row_dimension_items(rd, sid)
        if prev is not None and prev[2] == items and prev[1] + 1 == idx:
            prev[1] = idx
            continue
        if prev is not None:
            lines.append(_row_line(prev))
        prev = [idx, idx, items]
    if prev is not None:
        lines.append(_row_line(prev))
    return "".join(line + "\n" for line in lines)


def _row_line(p):
    r = str(p[0]) if p[0] == p[1] else f"{p[0]}:{p[1]}"
    return f"row(r={r} {p[2]})" if p[2] else f"row(r={r})"


def _validations_section(ws):
    dvs = ws.data_validations.dataValidation if ws.data_validations else []
    return "".join(_cx.encode_min(dv) + "\n" for dv in dvs)


def _conditional_section(ws):
    lines = []
    for cf in ws.conditional_formatting:
        ranges = str(cf.sqref).replace(" ", ",")
        for rule in cf.rules:
            el = rule.to_tree()
            el.attrib.pop("dxfId", None)  # the format itself follows: dxf(...)
            el = _cx.minimal(rule, el)
            if rule.dxf is not None:
                el.append(_cx.minimal(rule.dxf))
            lines.append(f"{ranges} {_cx.encode(el)}")
    return "".join(line + "\n" for line in lines)


def _tables_section(ws):
    return "".join(_cx.encode_min(t) + "\n" for t in ws.tables.values())


def _raw_elements(ws):
    """{local tag: element} of the worksheet's parts outside <sheetData>, as
    in the file"""
    out = {}
    try:
        if not _native.ENABLED:
            raise _native.NativeFallback("disabled")
        with ws._get_source() as src:
            res = _native._rs.prescan_sheet(src, False, None, _native.reader_helpers())
        for snippet, ns in list(res["head"]) + list(res["tail"]):
            el = _native.wrapped_element(snippet, ns)
            out.setdefault(el.tag.rsplit("}", 1)[-1], el)
    except _native.NativeFallback:
        from openrsxl.xml.functions import iterparse

        depth = 0
        with ws._get_source() as src:
            for ev, el in iterparse(src, events=("start", "end")):
                if ev == "start":
                    depth += 1
                    continue
                depth -= 1
                if depth == 1:
                    tag = el.tag.rsplit("}", 1)[-1]
                    if tag != "sheetData":
                        out.setdefault(tag, el)
                elif depth == 2 and el.tag.endswith("}row"):
                    el.clear()
    return out


def _sheet_section(ws):
    lines = []
    raw = None
    for attr, cls, factory in SHEET_OBJECTS:
        obj = getattr(ws, attr, None)
        if obj is None:
            continue
        el = obj.to_tree()
        if el is None:
            continue
        default = factory()
        if default is not None and _xml(default.to_tree()) == _xml(el):
            continue
        small = _cx.minimal(obj, el, cls, read_object)
        if small is el and not _cx.verifies(el, el, cls, read_object):
            # openpyxl does not read its own serialisation back exactly (eg.
            # header / footer escapes): the element of the file, from which
            # openpyxl built the object, reads back by definition
            if raw is None:
                raw = _raw_elements(ws)
            tag = el.tag.rsplit("}", 1)[-1]
            src = raw.get(tag)
            if src is not None and _cx.verifies(src, el, cls, read_object):
                small = src
            else:
                warnings.warn(f"openrsxl.extended.llm: {tag} of sheet {ws.title!r} may not decode exactly")
        lines.append(_cx.encode(small))
    items = []
    if ws.print_title_rows:
        items.append(f"rows={_cx.scalar(ws.print_title_rows)}")
    if ws.print_title_cols:
        items.append(f"cols={_cx.scalar(ws.print_title_cols)}")
    if items:
        lines.append(f"printTitles({' '.join(items)})")
    if ws.print_area:
        lines.append(f"printArea:{_cx.scalar(ws.print_area)}")
    return "".join(line + "\n" for line in lines)


def _has_part(wb, name):
    archive = getattr(wb, "_archive", None)
    if archive is None:
        return True
    try:
        archive.getinfo(name)
    except KeyError:
        return False
    return True


def _missing_times(wb):
    """'created' / 'modified' absent from the file's core properties"""
    try:
        from openrsxl.xml.functions import fromstring

        root = fromstring(wb._archive.read(ARC_CORE))
    except Exception:
        return set()
    present = {child.tag.rsplit("}", 1)[-1] for child in root}
    return {"created", "modified"} - present


def _workbook_section(wb):
    lines = []
    for ws in wb._sheets:
        items = [f"name={_cx.scalar(ws.title)}"]
        if ws.sheet_state != "visible":
            items.append(f"state={ws.sheet_state}")
        if not hasattr(ws, "iter_rows"):
            items.append("type=chartsheet")
        lines.append(f"sheet({' '.join(items)})")
    props = []
    if wb.epoch == CALENDAR_MAC_1904:
        props.append("date1904=1")
    if wb.code_name:
        props.append(f"codeName={_cx.scalar(wb.code_name)}")
    if props:
        lines.append(f"workbook({' '.join(props)})")
    if wb._active_sheet_index:
        lines.append(f"active:{wb._active_sheet_index}")
    for view in wb.views or []:
        lines.append(_cx.encode_min(view))
    if wb.calculation is not None and not _is_default(wb.calculation, CalcProperties()):
        lines.append(_cx.encode_min(wb.calculation))
    if wb.security is not None and not _is_default(wb.security, WorkbookProtection()):
        # WorkbookProtection serialises as "workbookPr": named after its role
        lines.append(f"workbookProtection({_cx.encode_items(_cx.minimal(wb.security))})")
    for dn in wb.defined_names.values():
        lines.append(_cx.encode_min(dn))
    for ws in wb._sheets:
        for dn in getattr(ws, "defined_names", {}).values():
            lines.append(_cx.encode_min(dn))
    if wb.properties is not None and _has_part(wb, ARC_CORE):
        # (without docProps/core.xml openpyxl makes up properties, and the
        # times the file lacks are "now": not data)
        el = _cx.minimal(wb.properties)
        missing = _missing_times(wb)
        for child in list(el):
            if child.tag.rsplit("}", 1)[-1] in missing:
                el.remove(child)
        lines.append(_cx.encode(el))
    if wb.custom_doc_props is not None and len(wb.custom_doc_props):
        lines.append(_cx.encode_min(wb.custom_doc_props))
    return "".join(line + "\n" for line in lines)


def _analysis_text(a):
    def fmt(v):
        if isinstance(v, float):
            return repr(v)
        return str(v)

    kinds = ",".join(f"{k}:{v}" for k, v in a["kinds"].items())
    est = ",".join(f"{k}:{v}" for k, v in a["estimated_tokens"].items() if v is not None)
    items = [
        f"cells={a['cells']}",
        f"rows={a['rows']}",
        f"values={a['values_range']}" if a["values_range"] else None,
        f"density={fmt(a['density'])}",
        f"distinct_ratio={fmt(a['distinct_ratio'])}" if a["distinct_ratio"] is not None else None,
        f"kinds={kinds}" if kinds else None,
        f"formulas={a['formulas']}" if a.get("formulas") else None,
        f"formula_groups={a['formula_groups']}" if a.get("formulas") else None,
        f"styled_cells={a['styled_cells']}" if a.get("styled_cells") else None,
        f"skipped_cells={a['skipped_cells']}" if a.get("skipped_cells") else None,
        f"est_tokens={est}" if est else None,
        f"encoding={a['encoding']}",
    ]
    line = " ".join(i for i in items if i)
    return line + "\nreason: " + a["reason"] + "\n"


# -- cell passes ----------------------------------------------------------------
def _parser_for(ws, data_only, source=None):
    wb = ws.parent
    return WorkSheetParser(
        source,
        ws._shared_strings,
        data_only=data_only,
        epoch=wb.epoch,
        date_formats=wb._date_formats,
        timedelta_formats=wb._timedelta_formats,
    )


def _feed_native(enc, ws, n_styles):
    parser = _parser_for(ws, False)
    parser.formula_and_value = True
    with ws._get_source() as src:
        enc.feed_native(src, parser, _native.reader_helpers(), _formula_info, n_styles)


def _feed_python(enc, ws, n_styles):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        with ws._get_source() as s1, ws._get_source() as s2:
            frows = _parser_for(ws, False, s1).parse()
            vrows = _parser_for(ws, True, s2).parse()
            for (_, fcells), (_, vcells) in zip(frows, vrows):
                batch = []
                for fc, vc in zip(fcells, vcells):
                    formula = fc["value"] if fc["data_type"] == "f" else None
                    batch.append((vc["row"], vc["column"], vc["value"], vc["data_type"], formula, vc["style_id"]))
                if batch:
                    enc.feed(batch, _formula_info, n_styles)


def _run_pass(enc, ws, n, mode=None):
    """Run a pass; returns (result, mode). Pass 1 tries the native reader,
    then the Python one, and buffers unordered sheets; pass 2 repeats the
    mode that worked."""
    n_styles = len(ws.parent._cell_styles)
    modes = [mode] if mode else (["native"] if _native.ENABLED else []) + ["python", "buffered"]
    last = None
    for m in modes:
        enc.begin_pass(n, buffered=(m == "buffered"))
        try:
            if m == "native":
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    _feed_native(enc, ws, n_styles)
            else:
                _feed_python(enc, ws, n_styles)
            return enc.end_pass(), m
        except _native.NativeFallback as e:
            last = e
            continue
        except ValueError as e:
            if "cells out of order" in str(e) and m != "buffered":
                last = e
                continue
            raise
    raise last


# -- main -------------------------------------------------------------------------
def _flags(sections):
    s = set(sections)
    return dict(
        read_merged_cells=bool(s & {"merged", "hyperlinks", "comments"}),
        read_hyperlinks="hyperlinks" in s,
        read_comments="comments" in s,
        read_dimensions="dimensions" in s,
        read_sheet_properties="sheet" in s,
        read_data_validations="validations" in s,
        read_conditional_formatting="conditional_formatting" in s,
        read_tables="tables" in s,
    )


class _Out:
    """Text / JSON files written while compressing"""

    def __init__(self, text_path, json_path):
        self.t = open(text_path, "w", encoding="utf-8", newline="\n") if text_path else None
        self.j = open(json_path, "w", encoding="utf-8", newline="\n") if json_path else None
        self._first_sheet = True
        self._first_section = True

    def text(self, s):
        if self.t is not None and s:
            self.t.write(s)

    def json_raw(self, s):
        if self.j is not None:
            self.j.write(s)

    def json_str(self, s):
        # chunk of a JSON string (without quotes)
        if self.j is not None and s:
            self.j.write(json.dumps(s, ensure_ascii=False)[1:-1])

    def close(self):
        for f in (self.t, self.j):
            if f is not None:
                f.close()


def _default_base(filename):
    if isinstance(filename, (str, os.PathLike)):
        base = os.fspath(filename)
        root, ext = os.path.splitext(base)
        return root
    return None


_OPTIONS = dict(
    include="all",
    exclude=None,
    encoding="auto",
    legend=True,
    jump_gap=4,
    runs=True,
    row_anchor=10,
    ditto=True,
    index_limit=64 << 20,
    keep_links=True,
    skip_hidden=False,
    skip_hidden_sheets=False,
    print_area_only=False,
    rows=None,
    columns=None,
    ranges=None,
)


def _file_paths(filename, output, path, suffix=""):
    """(keep in memory, text path, json path) of an ``output`` / ``path``"""
    outputs = _resolve_outputs(output)
    text_path = json_path = None
    if outputs & {"text", "json"}:
        base = path if path is not None else _default_base(filename)
        if base is None:
            raise ValueError("path= is required to write files when the source is a file object")
        base = os.fspath(base)
        if base.endswith((".txt", ".json")):
            base = os.path.splitext(base)[0]
        if base.endswith(".llm"):
            base = base[:-4]
        base += suffix
        parent = os.path.dirname(base)
        if parent:
            os.makedirs(parent, exist_ok=True)
        if "text" in outputs:
            text_path = base + ".llm.txt"
        if "json" in outputs:
            json_path = base + ".llm.json"
    return "memory" in outputs, text_path, json_path


class _Job:
    """A workbook opened for compression with the given options"""

    def __init__(self, filename, options, sheets=None):
        unknown = set(options) - set(_OPTIONS)
        if unknown:
            raise TypeError(f"unexpected keyword argument {sorted(unknown)[0]!r}")
        o = dict(_OPTIONS, **options)
        if _native._rs is None:  # pragma: no cover
            raise ImportError("openrsxl.extended.llm needs the native engine")
        if o["encoding"] not in ENCODINGS:
            raise ValueError(f"encoding must be one of {', '.join(ENCODINGS)}")
        self.filename = filename
        self.sections = _resolve_sections(o["include"], o["exclude"])
        self.encoding = o["encoding"]
        self.legend = o["legend"]
        self.tuning = {k: o[k] for k in ("jump_gap", "runs", "row_anchor", "ditto", "index_limit")}
        self.scope = {k: bool(o[k]) for k in ("skip_hidden", "skip_hidden_sheets", "print_area_only")}
        self.scope["selection"] = Selection(o["rows"], o["columns"], o["ranges"])
        flags = _flags(self.sections)
        if o["skip_hidden"]:
            flags["read_dimensions"] = True  # hidden rows / columns
            flags["read_sheet_properties"] = True  # sheetFormatPr zeroHeight
        if o["print_area_only"]:
            flags["read_sheet_properties"] = True  # print area and titles
        from openrsxl.extended._load import _load_for

        # the metadata of the selected sheets only: the others are not read
        self.wb = _load_for(filename, sheets, keep_links=o["keep_links"], **flags)
        titles = {ws.title for ws in self.wb._sheets if hasattr(ws, "iter_rows")}
        unknown = sorted(self.scope["selection"].sheet_names() - titles)
        if unknown:
            self.wb.close()
            raise KeyError(f"rows / columns / ranges: no worksheet named {unknown[0]!r}")

    def worksheet_names(self, sheets=None):
        """names of the worksheets `sheets` selects, in workbook order"""
        return [ws.title for ws in _select(self.wb, sheets, self.scope)]

    def run(self, sheets, output, path, suffix=""):
        keep, text_path, json_path = _file_paths(self.filename, output, path, suffix)
        return _compress(
            self.wb,
            self.filename,
            self.sections,
            sheets,
            self.encoding,
            keep,
            text_path,
            json_path,
            self.legend,
            self.tuning,
            self.scope,
        )

    def close(self):
        self.wb.close()


def _select(wb, sheet_names, scope):
    """the worksheets to compress: `sheet_names` (a name or names; None:
    all), in workbook order, without hidden ones if skip_hidden_sheets"""
    worksheets = [ws for ws in wb._sheets if hasattr(ws, "iter_rows")]
    if sheet_names is not None:
        if isinstance(sheet_names, str):
            sheet_names = [sheet_names]
        sheet_names = list(sheet_names)
        names = {ws.title for ws in worksheets}
        missing = [n for n in sheet_names if n not in names]
        if missing:
            if missing[0] in {ws.title for ws in wb._sheets}:
                raise KeyError(f"{missing[0]!r} is a chartsheet, not a worksheet")
            raise KeyError(f"no worksheet named {missing[0]!r}")
        wanted = set(sheet_names)
        worksheets = [ws for ws in worksheets if ws.title in wanted]
    if scope.get("skip_hidden_sheets"):
        worksheets = [ws for ws in worksheets if ws.sheet_state == "visible"]
    return worksheets


def compress_workbook(
    filename,
    *,
    include="all",
    exclude=None,
    sheets=None,
    encoding="auto",
    output="memory",
    path=None,
    legend=True,
    jump_gap=4,
    runs=True,
    row_anchor=10,
    ditto=True,
    index_limit=64 << 20,
    keep_links=True,
    skip_hidden=False,
    skip_hidden_sheets=False,
    print_area_only=False,
    rows=None,
    columns=None,
    ranges=None,
):
    """Compress a workbook - or the worksheets named in ``sheets`` (a name
    or a list of names) - into an LLM readable, lossless text document
    (full documentation: ``help(openrsxl.extended.llm)``)"""
    options = {k: v for k, v in locals().items() if k in _OPTIONS}
    job = _Job(filename, options, sheets)
    try:
        return job.run(sheets, output, path)
    finally:
        job.close()


def _safe_name(name):
    """a sheet name usable in a file name"""
    out = "".join(c if c.isalnum() or c in "-_ ." else "_" for c in name).strip(" .")
    return out or "sheet"


def compress_sheets(
    filename,
    *,
    include="all",
    exclude=None,
    sheets=None,
    encoding="auto",
    output="memory",
    path=None,
    legend=True,
    jump_gap=4,
    runs=True,
    row_anchor=10,
    ditto=True,
    index_limit=64 << 20,
    keep_links=True,
    skip_hidden=False,
    skip_hidden_sheets=False,
    print_area_only=False,
    rows=None,
    columns=None,
    ranges=None,
):
    """
    Compress one worksheet at a time: yields ``(sheet name,
    CompressedWorkbook)`` for each worksheet ``sheets`` selects (default:
    all), in workbook order. The workbook is opened once; each document is
    complete on its own (legend, ``## workbook``, the styles its sheet uses)
    and equal to ``compress_workbook(filename, sheets=[name], ...)``. With
    ``output`` "text" / "json", each sheet gets its own files:
    ``<path>.<sheet name>.llm.txt`` (characters not allowed in file names
    become ``_``). A document is finished before the next sheet is read:
    stop iterating to skip the remaining sheets.
    """
    options = {k: v for k, v in locals().items() if k in _OPTIONS}
    job = _Job(filename, options, sheets)
    try:
        used = set()
        for name in job.worksheet_names(sheets):
            suffix = "." + _safe_name(name)
            n = 2
            while suffix.lower() in used:
                suffix = f".{_safe_name(name)}~{n}"
                n += 1
            used.add(suffix.lower())
            yield name, job.run([name], output, path, suffix)
    finally:
        job.close()


def _compress(
    wb, filename, sections, sheet_names, encoding, keep, text_path, json_path, want_legend, tuning, scope=None
):
    scope = scope or {}
    secset = set(sections)
    cell_secs = secset & {"values", "formulas", "styles"}
    worksheets = _select(wb, sheet_names, scope)
    # hidden sheets left out: the legend says so
    skipped_sheets = bool(scope.get("skip_hidden_sheets")) and len(worksheets) < len(_select(wb, sheet_names, {}))
    position = {id(ws): i for i, ws in enumerate(wb._sheets, 1)}
    # the analysis measures everything, even what is not encoded
    measure = "analysis" in secset
    opts = dict(
        values="values" in secset or measure,
        formulas="formulas" in secset or measure,
        styles="styles" in secset or measure,
        encoding=encoding,
        **tuning,
    )

    # pass 1 of every sheet (analysis, formulas, style map, index values) and
    # the metadata sections - everything but the grids - so that the legend
    # describes exactly what the document contains
    jobs = []
    for ws in worksheets:
        enc = mode = analysis = None
        p1 = {}
        cf = _cell_filter(ws, scope.get("skip_hidden"), scope.get("print_area_only"), scope.get("selection"))
        if cell_secs or measure:
            enc = _native._rs.SheetEncoder(dict(opts, **cf.options()) if cf.active else opts)
            analysis, mode = _run_pass(enc, ws, 1)
            p1 = enc.sections()
            enc.release()
        used_styles = set(analysis["used_styles"]) if analysis is not None else set()
        texts = OrderedDict()
        if "analysis" in secset:
            texts["analysis"] = _analysis_text(analysis)
        if "values" in secset:
            texts["values"] = p1["values"]  # None: the grid of pass 2
        if "formulas" in secset:
            texts["formulas"] = p1["formulas"]
        meta = OrderedDict()
        if "merged" in secset:
            meta["merged"] = _merged_section(ws)
        if "hyperlinks" in secset:
            meta["hyperlinks"] = _hyperlinks_section(ws, cf if cf.active else None)
        if "comments" in secset:
            meta["comments"] = _comments_section(ws, cf if cf.active else None)
        if "dimensions" in secset:
            meta["dimensions"] = _dimensions_section(ws, used_styles)
        if "validations" in secset:
            meta["validations"] = _validations_section(ws)
        if "conditional_formatting" in secset:
            meta["conditional_formatting"] = _conditional_section(ws)
        if "tables" in secset:
            meta["tables"] = _tables_section(ws)
        if "sheet" in secset:
            meta["sheet"] = _sheet_section(ws)
        if "styles" in secset:
            # (the style table itself is the workbook's ## styles)
            texts["styles"] = p1["style_map"]
        texts.update(meta)
        jobs.append((ws, enc, mode, analysis, p1, texts, used_styles, cf))

    # the legend: the parts the document uses
    parts = set()
    if skipped_sheets:
        parts.add("filter")
    for _, _, _, analysis, _, texts, _, cf in jobs:
        if cf.active:
            parts.add("filter")
        for name, t in texts.items():
            if name == "values":
                if analysis["cells"]:
                    parts.add(analysis["encoding"])
            elif t:
                parts.add(name)
    wbsec = None
    if "workbook" in secset:
        t = _workbook_section(wb)
        wbsec = Section("workbook", t, tokens=estimate_tokens(t), size=len(t))
        parts.add("workbook")
    table = None
    if "styles" in secset and jobs:
        table = _style_table(wb, set().union(*(job[6] for job in jobs)))
        parts.add("styles")
    legend_text = legend(parts) if want_legend else None
    source = os.path.basename(os.fspath(filename)) if isinstance(filename, (str, os.PathLike)) else None
    result = CompressedWorkbook(source, sections, legend_text, wbsec, table)
    result.in_memory = keep

    out = _Out(text_path, json_path)
    try:
        head = result.header() + "\n"
        if legend_text:
            head += "## legend\n" + legend_text + "\n"
        if wbsec is not None:
            head += "## workbook\n" + wbsec.text
        if table is not None:
            head += "## styles\n" + table.text()
        out.text(head)
        out.json_raw(
            '{"format": %s, "source": %s, "sections": %s'
            % (json.dumps(FORMAT), json.dumps(source, ensure_ascii=False), json.dumps(list(sections)))
        )
        if legend_text:
            out.json_raw(', "legend": ' + json.dumps(legend_text, ensure_ascii=False))
        if wbsec is not None:
            out.json_raw(', "workbook": ' + json.dumps(wbsec.text, ensure_ascii=False))
        if table is not None:
            out.json_raw(', "styles": ' + json.dumps(table.text(), ensure_ascii=False))
        out.json_raw(', "sheets": [')

        for k, (ws, enc, mode, analysis, p1, texts, used, cf) in enumerate(jobs):
            n = position[id(ws)]  # the sheet's position in the workbook
            dimension = analysis["dimension"] if analysis else None
            sheet = CompressedSheet(n, ws.title, dimension, analysis)
            sheet.used_styles = used
            sheet.filter = cf.header_items()
            out.text(sheet.header + "\n")
            if k > 0:
                out.json_raw(", ")
            filt = ('"filter": %s, ' % json.dumps(sheet.filter)) if sheet.filter else ""
            out.json_raw(
                '{"index": %d, "name": %s, "dimension": %s, %s"analysis": %s, "sections": {'
                % (n, json.dumps(ws.title, ensure_ascii=False), json.dumps(dimension), filt, json.dumps(analysis))
            )
            first = [True]

            def emit(sec, stream=None, sheet=sheet, first=first):
                out.text(sec.header + "\n")
                if not first[0]:
                    out.json_raw(", ")
                first[0] = False
                d = sec.to_dict()
                d.pop("text")
                head = json.dumps(sec.name) + ": " + json.dumps(d)[:-1]
                out.json_raw(head + (', "text": "' if d else '"text": "'))
                if stream is not None:
                    stream()
                else:
                    out.text(sec.text)
                    out.json_str(sec.text)
                out.json_raw('"}')
                sheet.sections[sec.name] = sec
                if not keep:
                    sec.text = None

            for name, t in texts.items():
                if name != "values":
                    if t:
                        emit(Section(name, t, tokens=estimate_tokens(t), size=len(t)))
                    continue
                if not analysis["cells"]:
                    continue
                enc_name = analysis["encoding"]
                start = analysis["grid_start"] if enc_name in ("grid", "hybrid") else None
                if t is not None:
                    emit(Section("values", t, enc_name, start, estimate_tokens(t), len(t)))
                    continue
                blocks = p1["blocks"]
                sec = Section("values", None, enc_name, start)

                def stream(enc=enc, mode=mode, ws=ws, blocks=blocks, sec=sec):
                    out.text(blocks)
                    out.json_str(blocks)
                    if keep:
                        enc.set_sink(None)
                    else:

                        def sink(chunk):
                            out.text(chunk)
                            out.json_str(chunk)

                        enc.set_sink(sink)
                    _run_pass(enc, ws, 2, mode)
                    tokens, size = enc.grid_size
                    grid = enc.take_grid()
                    if keep:
                        sec.text = blocks + grid
                        out.text(grid)
                        out.json_str(grid)
                    sec.tokens = tokens + estimate_tokens(blocks)
                    sec.size = size + len(blocks)

                emit(sec, stream)
            out.json_raw("}}")
            result.sheets.append(sheet)
            if enc is not None:
                enc.release()
        out.json_raw("]}")
    finally:
        out.close()
    if text_path:
        result.files["text"] = text_path
    if json_path:
        result.files["json"] = json_path
    return result
