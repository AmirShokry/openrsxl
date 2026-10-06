"""
openrsxl.extended: streaming (read-only) worksheets with the features that
openpyxl only offers in its full (non read-only) mode.

Every feature is opt-in (keyword arguments of
``openrsxl.extended.load_workbook``) and reproduces what openpyxl's full mode
returns for the same file - same attribute names, same types, same values:

=============================  ==============================================
flag                           adds
=============================  ==============================================
``formula_and_value``          ``cell.formula`` and ``cell.cached_value``
``read_comments``              ``cell.comment``
``read_hyperlinks``            ``cell.hyperlink`` (and the value openpyxl
                               gives to empty hyperlink cells)
``read_merged_cells``          ``ws.merged_cells`` and ``MergedCell`` objects
``read_dimensions``            ``ws.row_dimensions``, ``ws.column_dimensions``
``read_sheet_properties``      ``ws.sheet_properties``, ``views``,
                               ``freeze_panes``, print / page settings,
                               ``protection``, ``auto_filter``, ...
``read_data_validations``      ``ws.data_validations``
``read_conditional_formatting`` ``ws.conditional_formatting``
``read_tables``                ``ws.tables``
``create_empty_cells``         a cell at every empty position, as full mode
                               creates on access (``coordinate``, default
                               style), instead of the shared ``EmptyCell``
=============================  ==============================================

Memory stays independent of the number of cells: rows are streamed exactly
as in read-only mode; only the requested features themselves (merged ranges,
hyperlinks, comments, dimensions...) are kept in memory.
"""

import warnings
from collections import defaultdict
from copy import copy

from openrsxl import _native
from openrsxl.cell.cell import Cell
from openrsxl.cell.cell import MergedCell as _MergedCell
from openrsxl.cell.read_only import EmptyCell as _EmptyCell
from openrsxl.cell.read_only import ReadOnlyCell as _ReadOnlyCell
from openrsxl.comments.comment_sheet import CommentSheet
from openrsxl.formatting.formatting import ConditionalFormattingList
from openrsxl.styles.cell_style import StyleArray
from openrsxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries
from openrsxl.worksheet._read_only import ReadOnlyWorksheet as _ReadOnlyWorksheet
from openrsxl.worksheet._reader import WorkSheetParser
from openrsxl.worksheet.cell_range import CellRange, MultiCellRange
from openrsxl.worksheet.datavalidation import DataValidationList
from openrsxl.worksheet.dimensions import (
    ColumnDimension,
    DimensionHolder,
    RowDimension,
    SheetFormatProperties,
)
from openrsxl.worksheet.filters import AutoFilter
from openrsxl.worksheet.merge import MergedCellRange
from openrsxl.worksheet.page import PageMargins, PrintOptions, PrintPageSetup
from openrsxl.worksheet.pagebreak import ColBreak, RowBreak
from openrsxl.worksheet.print_settings import PrintArea
from openrsxl.worksheet.properties import WorksheetProperties
from openrsxl.worksheet.protection import SheetProtection
from openrsxl.worksheet.scenario import ScenarioList
from openrsxl.worksheet.table import Table, TableList
from openrsxl.worksheet.views import SheetViewList
from openrsxl.worksheet.worksheet import Worksheet
from openrsxl.xml.constants import COMMENTS_NS
from openrsxl.xml.functions import fromstring

FLAGS = (
    "formula_and_value",
    "read_comments",
    "read_hyperlinks",
    "read_merged_cells",
    "read_dimensions",
    "read_sheet_properties",
    "read_data_validations",
    "read_conditional_formatting",
    "read_tables",
    "create_empty_cells",
)

COMMENT_WARNING = """Cell '{0}':{1} is part of a merged range but has a comment which will be removed because merged cells cannot contain any data."""

_MISSING = object()


class ReadOnlyCell(_ReadOnlyCell):
    """
    Read-only cell of openrsxl.extended streaming worksheets.

    Attributes added by the opt-in features (AttributeError otherwise):

    * ``formula``: the formula as openpyxl's full mode returns it
      (``data_only=False``): a ``str`` starting with ``=``, an
      ``ArrayFormula`` or a ``DataTableFormula``; None for other cells.
    * ``cached_value``: the value openpyxl returns with ``data_only=True``
      (the result saved by Excel for formula cells).
    * ``hyperlink``: a ``Hyperlink`` or None.
    * ``comment``: a ``Comment`` or None.
    """

    # plain slots (attribute access at C speed); a slot of a disabled
    # feature is never set and raises AttributeError like any attribute
    # that read-only cells do not have
    __slots__ = ("formula", "cached_value", "hyperlink", "comment")

    def __eq__(self, other):
        for a in _ReadOnlyCell.__slots__:
            if getattr(self, a) != getattr(other, a):
                return
        for a in ReadOnlyCell.__slots__:
            if getattr(self, a, _MISSING) != getattr(other, a, _MISSING):
                return
        return True

    def __ne__(self, other):
        return not self.__eq__(other)

    @property
    def style_array(self):
        # cells created by a feature (a link or comment on an empty cell, the
        # top-left cell of a merged range) carry their own StyleArray, as
        # full mode's Cell._style: the workbook's style registry stays
        # exactly openpyxl's
        sid = self._style_id
        if sid.__class__ is StyleArray:
            return sid
        return self.parent.parent._cell_styles[sid]

    @property
    def has_style(self):
        sid = self._style_id
        if sid.__class__ is StyleArray:
            return any(sid)
        return sid != 0


class EmptyCell(_EmptyCell):
    """Missing cells of openrsxl.extended streaming worksheets."""

    __slots__ = ()
    formula = None
    cached_value = None
    hyperlink = None
    comment = None


EMPTY_CELL = EmptyCell()


class MergedCell(_MergedCell):
    """MergedCell of openrsxl.extended (adds formula / cached_value)."""

    __slots__ = ()
    formula = None
    cached_value = None


class _Options:
    __slots__ = FLAGS

    def __init__(self, **kw):
        for k in FLAGS:
            setattr(self, k, bool(kw.get(k, False)))

    @property
    def cell_level(self):
        return (
            self.formula_and_value
            or self.read_comments
            or self.read_hyperlinks
            or self.read_merged_cells
            or self.create_empty_cells
        )

    @property
    def needs_tail(self):
        return (
            self.read_hyperlinks
            or self.read_merged_cells
            or self.read_sheet_properties
            or self.read_data_validations
            or self.read_conditional_formatting
            or self.read_tables
            or self.read_dimensions
        )


def _scan(ws, want_rows=False, corners=None):
    """
    Pre-scan a worksheet: returns a WorkSheetParser filled by openpyxl's own
    handlers with everything outside the cell data (+ row dimensions and
    the style ids of the requested cells).
    """
    parser = WorkSheetParser(
        None,
        [],
        data_only=ws.parent.data_only,
        epoch=ws.parent.epoch,
        date_formats=ws.parent._date_formats,
        timedelta_formats=ws.parent._timedelta_formats,
    )
    parser._corner_styles = {}
    result = None
    if _native.ENABLED:
        try:
            with ws._get_source() as src:
                result = _native._rs.prescan_sheet(src, want_rows, corners, _native.reader_helpers())
        except _native.NativeFallback:
            result = None
    if result is not None:
        try:
            # handlers may warn (eg. unsupported extensions): openpyxl's
            # read-only mode reports those while iterating, not here
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                for snippet, ns in result["head"]:
                    parser._dispatch_snippet(snippet, ns)
                for snippet, ns in result["tail"]:
                    parser._dispatch_snippet(snippet, ns)
            for idx, attrs in result["rows"]:
                parser.row_dimensions[str(idx)] = dict(attrs)
            for coord, s in result["corners"].items():
                parser._corner_styles[coord] = s
            return parser
        except _native.NativeFallback:
            pass
    # exact (slower) path: the original parser over the whole document
    parser = _python_scan(ws, corners)
    return parser


class _CornerParser(WorkSheetParser):
    """WorkSheetParser recording the 's' attribute of some cells."""

    def parse_row(self, row):
        # WorkSheetParser.parse_row without interpreting the cells (their
        # errors belong to the iteration, as in read-only mode)
        attrs = dict(row.attrib)
        if "r" in attrs:
            try:
                self.row_counter = int(attrs["r"])
            except ValueError:
                val = float(attrs["r"])
                if val.is_integer():
                    self.row_counter = int(val)
                else:
                    raise ValueError(f"{attrs['r']} is not a valid row number")
        else:
            self.row_counter += 1
        self.col_counter = 0
        keys = {k for k in attrs if not k.startswith("{")}
        if keys - {"r", "spans"}:
            self.row_dimensions[str(self.row_counter)] = attrs
        if self._corners:
            for el in row:
                coordinate = el.get("r")
                if coordinate:
                    r, c = coordinate_to_tuple(coordinate)
                    self.col_counter = c
                else:
                    self.col_counter += 1
                    r, c = self.row_counter, self.col_counter
                if (r, c) in self._corners:
                    self._corner_styles[(r, c)] = el.get("s")
        return self.row_counter, []


def _python_scan(ws, corners):
    with ws._get_source() as src:
        parser = _CornerParser(
            src,
            ws._shared_strings,
            data_only=ws.parent.data_only,
            epoch=ws.parent.epoch,
            date_formats=ws.parent._date_formats,
            timedelta_formats=ws.parent._timedelta_formats,
        )
        parser._corners = set(corners or ())
        parser._corner_styles = {}
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            for _ in parser.parse():
                pass
    return parser


def _style_id_of(s):
    """`element.get('s', 0)`; `int()` if not empty (WorkSheetParser.parse_cell)"""
    if s is None:
        return 0
    if s:
        return int(s)
    return s


class _MergeInfo:
    """Merged ranges with the styles openpyxl's full mode gives their cells."""

    # ranges up to this size are formatted directly, larger ones through a
    # representative range of at most 3x3 cells (same edge classes)
    SMALL = 64

    def __init__(self, ws, ranges):
        self.ws = ws
        self.ranges = ranges  # list of CellRange, file order
        self.by_row = sorted(range(len(ranges)), key=lambda i: ranges[i].min_row)
        self._styles = None  # computed on first cell iteration
        self.shadow_starts = {}  # range index -> start_cell of the simulation

    # -- geometry ---------------------------------------------------------
    def bounds(self):
        if not self.ranges:
            return None
        return (
            min(r.min_row for r in self.ranges),
            max(r.max_row for r in self.ranges),
            min(r.min_col for r in self.ranges),
            max(r.max_col for r in self.ranges),
        )

    # rows per bucket of the lookup index
    BUCKET = 64

    def _candidates(self, row):
        index = self.__dict__.get("_index")
        if index is None:
            # ranges spanning many buckets (eg. whole columns) are kept in a
            # separate list so the index stays proportional to the ranges
            index, tall = {}, []
            for i, r in enumerate(self.ranges):
                first, last = r.min_row // self.BUCKET, r.max_row // self.BUCKET
                if last - first > 16:
                    tall.append(i)
                    continue
                for b in range(first, last + 1):
                    index.setdefault(b, []).append(i)
            self._index, self._tall = index, tall
        found = index.get(row // self.BUCKET, ())
        if self._tall:
            return sorted(set(found).union(self._tall))
        return found

    def merged_coord(self, row, col):
        """The first range (file order) containing (row, col), if (row, col)
        is not its top-left cell (cells openpyxl turns into MergedCell)"""
        for i in self._candidates(row):
            r = self.ranges[i]
            if r.min_row <= row <= r.max_row and r.min_col <= col <= r.max_col:
                if (row, col) != (r.min_row, r.min_col):
                    return r
        return None

    # -- styles -------------------------------------------------------------
    def _groups(self):
        """Indices of ranges grouped by overlap (overlapping ranges interact)"""
        n = len(self.ranges)
        parent = list(range(n))

        def find(a):
            while parent[a] != a:
                parent[a] = parent[parent[a]]
                a = parent[a]
            return a

        order = sorted(range(n), key=lambda i: (self.ranges[i].min_row, self.ranges[i].min_col))
        active = []
        for i in order:
            ri = self.ranges[i]
            active = [j for j in active if self.ranges[j].max_row >= ri.min_row]
            for j in active:
                rj = self.ranges[j]
                if not (rj.max_col < ri.min_col or ri.max_col < rj.min_col):
                    parent[find(i)] = find(j)
            active.append(i)
        groups = defaultdict(list)
        for i in range(n):
            groups[find(i)].append(i)
        return [sorted(g) for g in groups.values()]

    def compute_styles(self):
        """
        Run openpyxl's own merge logic (MergedCellRange, _clean_merge_range)
        on a detached worksheet holding only the corner cells.
        """
        if self._styles is not None:
            return self._styles
        ws = self.ws
        wb = ws.parent
        corners = set()
        for r in self.ranges:
            corners.add((r.min_row, r.min_col))
            corners.add((r.max_row, r.max_col))
        parser = _scan(ws, want_rows=False, corners=corners)
        present = {coord: _style_id_of(s) for coord, s in parser._corner_styles.items()}
        styles = {}  # range index -> callable(row, col) -> spec

        for group in self._groups():
            shadow = Worksheet(wb)
            if (
                len(group) == 1
                and self.ranges[group[0]].size["rows"] * self.ranges[group[0]].size["columns"] > self.SMALL
            ):
                rng = self.ranges[group[0]]
                rep = self._representative(rng)
                self._place(shadow, present, rng, rep)
                mcr = MergedCellRange(shadow, rep.coord)
                shadow._clean_merge_range(mcr)
                styles[group[0]] = (rep, self._results(shadow, rep))
                continue
            for i in group:
                rng = self.ranges[i]
                for coord in ((rng.min_row, rng.min_col), (rng.max_row, rng.max_col)):
                    if coord in present and coord not in shadow._cells:
                        shadow._cells[coord] = Cell(
                            shadow, row=coord[0], column=coord[1], style_array=wb._cell_styles[present[coord]]
                        )
            for i in group:
                mcr = MergedCellRange(shadow, self.ranges[i].coord)
                shadow._clean_merge_range(mcr)
                # the cell object full mode keeps as this range's start_cell
                # (later overlapping ranges may replace it in ws._cells)
                self.shadow_starts[i] = mcr.start_cell
            for i in group:
                styles[i] = (None, self._results(shadow, self.ranges[i]))
        self._styles = styles
        return styles

    @staticmethod
    def _representative(rng):
        h = min(rng.size["rows"], 3)
        w = min(rng.size["columns"], 3)
        return CellRange(
            min_col=rng.min_col, min_row=rng.min_row, max_col=rng.min_col + w - 1, max_row=rng.min_row + h - 1
        )

    @staticmethod
    def _place(shadow, present, rng, rep):
        wb = shadow.parent
        start = (rng.min_row, rng.min_col)
        end = (rng.max_row, rng.max_col)
        if start in present:
            shadow._cells[start] = Cell(
                shadow, row=start[0], column=start[1], style_array=wb._cell_styles[present[start]]
            )
        if end in present:
            rep_end = (rep.max_row, rep.max_col)
            shadow._cells[rep_end] = Cell(
                shadow, row=rep_end[0], column=rep_end[1], style_array=wb._cell_styles[present[end]]
            )

    @staticmethod
    def _results(shadow, rng):
        out = {}
        for coord in rng.cells:
            out[coord] = shadow._cells[coord]
        return out

    def spec(self, idx, row, col):
        """The shadow cell (Cell or MergedCell) for (row, col) of range idx"""
        rep, results = self.compute_styles()[idx]
        if rep is None:
            return results[(row, col)]
        rng = self.ranges[idx]

        def map_(v, lo, hi, rep_lo, rep_hi):
            if v == lo:
                return rep_lo
            if v == hi:
                return rep_hi
            return rep_lo + 1

        r = map_(row, rng.min_row, rng.max_row, rep.min_row, rep.max_row)
        c = map_(col, rng.min_col, rng.max_col, rep.min_col, rep.max_col)
        return results[(r, c)]


class StreamedMergedCellRange(MergedCellRange):
    """
    MergedCellRange of a streamed worksheet.

    ``start_cell`` is the top-left cell exactly as ``iter_rows`` yields it
    (value, borders of the range, hyperlink, comment...), like full mode's
    ``start_cell``. It is read on first access, for all ranges of the sheet
    in one streaming pass, and kept (one cell per range).
    """

    @classmethod
    def _new(cls, ws, coord):
        mcr = cls.__new__(cls)
        CellRange.__init__(mcr, range_string=coord)
        mcr.ws = ws
        return mcr

    @property
    def start_cell(self):
        return self.ws._merged_start_cells().get(self.coord)

    def __copy__(self):
        return self._new(self.ws, self.coord)


StreamedMergedCellRange.__name__ = StreamedMergedCellRange.__qualname__ = "MergedCellRange"


class ReadOnlyWorksheet(_ReadOnlyWorksheet):
    """
    Streaming worksheet of openrsxl.extended.

    Identical to openpyxl's ``ReadOnlyWorksheet`` plus the attributes of the
    enabled features, with the names and types of openpyxl's ``Worksheet``.
    """

    # Worksheet API made available by the features
    sheet_view = Worksheet.sheet_view
    selected_cell = Worksheet.selected_cell
    active_cell = Worksheet.active_cell
    show_gridlines = Worksheet.show_gridlines
    freeze_panes = Worksheet.freeze_panes
    print_title_rows = Worksheet.print_title_rows
    print_title_cols = Worksheet.print_title_cols
    print_titles = Worksheet.print_titles
    print_area = Worksheet.print_area
    tables = Worksheet.tables
    _add_row = Worksheet._add_row
    _add_column = Worksheet._add_column

    def __init__(self, parent_workbook, title, worksheet_path, shared_strings, options=None):
        self._ext = options or _Options()
        self._overlay = {}
        self._merge = None
        super().__init__(parent_workbook, title, worksheet_path, shared_strings)
        if self._ext.read_sheet_properties:
            # defaults of Worksheet._setup, overwritten by the file's values
            self._print_rows = None
            self._print_cols = None
            self._print_area = PrintArea()

    # -- dimensions --------------------------------------------------------
    @property
    def max_row(self):
        mr = self._max_row
        b = self._ext_bounds
        if b is not None and mr is not None:
            return max(mr, b[1])
        return mr

    @property
    def max_column(self):
        mc = self._max_column
        b = self._ext_bounds
        if b is not None and mc is not None:
            return max(mc, b[3])
        return mc

    _ext_bounds = None

    # -- binding (at load time, like openpyxl's full mode) ------------------
    def _ext_bind(self, archive, rels):
        opt = self._ext
        self._rels = rels
        parser = None
        if opt.needs_tail:
            parser = _scan(self, want_rows=opt.read_dimensions)
        if opt.read_merged_cells:
            ranges = []
            if parser.merged_cells:
                for cr in parser.merged_cells.mergeCell:
                    ranges.append(cr.coord)
            self._bind_merged(ranges)
        if opt.read_hyperlinks:
            self._bind_hyperlinks(parser.hyperlinks.hyperlink)
        if opt.read_comments:
            self._bind_comments(archive, rels)
        if opt.read_conditional_formatting:
            self.conditional_formatting = ConditionalFormattingList()
            for cf in parser.formatting:
                for rule in cf.rules:
                    if rule.dxfId is not None:
                        rule.dxf = self.parent._differential_styles[rule.dxfId]
                    self.conditional_formatting[cf] = rule
        if opt.read_data_validations:
            self.data_validations = getattr(parser, "data_validations", None) or DataValidationList()
        if opt.read_dimensions:
            self._bind_dimensions(parser)
        if opt.read_sheet_properties:
            self._bind_properties(parser)
        if opt.read_tables:
            self._tables = TableList()
            for t in parser.tables.tablePart:
                rel = self._rels.get(t.id)
                table = Table.from_tree(fromstring(archive.read(rel.Target)))
                self._tables.add(table)
        self._ext_compute_bounds()

    def _bind_merged(self, refs):
        mcrs = []
        ranges = []
        for ref in refs:
            mcrs.append(StreamedMergedCellRange._new(self, ref))
            ranges.append(CellRange(ref))
        self.merged_cells = MultiCellRange(mcrs)
        self._merge = _MergeInfo(self, ranges) if ranges else None

    def _merged_start_cells(self):
        """
        start_cell of every merged range, by range coordinate (one streaming
        pass, cached): the top-left cell as iter_rows yields it, or - for a
        range whose top-left cell another (overlapping) range covers - the
        cell full mode keeps there: the file's cell with the styles of the
        merge, no hyperlink / comment (bound later to the MergedCell).
        """
        cells = self.__dict__.get("_start_cells")
        if cells is not None:
            return cells
        cells = {}
        merge = self._merge
        if merge is None:
            self._start_cells = cells
            return cells
        merge.compute_styles()
        orphans = {}
        for i, r in enumerate(merge.ranges):
            if merge.merged_coord(r.min_row, r.min_col) is not None:
                orphans.setdefault(r.coord, i)
        starts = {(r.min_row, r.min_col) for r in merge.ranges}
        min_row = min(r for r, _ in starts)
        max_row = max(r for r, _ in starts)
        min_col = min(c for _, c in starts)
        max_col = max(c for _, c in starts)
        orphan_coords = {(merge.ranges[i].min_row, merge.ranges[i].min_col) for i in orphans.values()}
        base_cells = {}
        base = self._ext_base_rows(min_col, min_row, max_col, max_row, False)
        if orphan_coords:
            base = _record(base, min_row, min_col, orphan_coords, base_cells)
        live = {}
        # by position, as iter_rows yields them (a cell's own coordinate can
        # differ from its position, eg. r="A0" in <row r="1">)
        for r_idx, row in enumerate(self._apply_overlays(base, min_col, min_row, max_col, max_row, False), min_row):
            for c_idx, c in enumerate(row, min_col):
                if (r_idx, c_idx) in starts:
                    live[(r_idx, c_idx)] = c
        cell_cls = self._ext_cell_cls()
        extra = self._ext_extra_slots()
        for i, r in enumerate(merge.ranges):
            if r.coord in cells:
                continue
            key = (r.min_row, r.min_col)
            if r.coord not in orphans:
                cells[r.coord] = live.get(key)
                continue
            shadow = merge.shadow_starts[i]
            if isinstance(shadow, _MergedCell):
                c = MergedCell(self, key[0], key[1])
                c._style = copy(shadow._style)
            else:
                c = base_cells.get(key)
                if c is None:
                    c = self._new_cell(cell_cls, extra, key[0], key[1])
                c._style_id = StyleArray(shadow._style)
            cells[r.coord] = c
        self._start_cells = cells
        return cells

    def _overlay_at(self, row, col):
        rows = self._overlay.setdefault(row, {})
        spec = rows.get(col)
        if spec is None:
            spec = rows[col] = {}
        return spec

    def _cell_exists_merged(self, row, col):
        return self._merge is not None and self._merge.merged_coord(row, col) is not None

    def _bind_hyperlinks(self, links):
        for link in links:
            if link.id:
                rel = self._rels.get(link.id)
                link.target = rel.Target
            if ":" in link.ref:
                min_col, min_row, max_col, max_row = range_boundaries(link.ref)
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        if self._cell_exists_merged(row, col):
                            continue  # MergedCell.hyperlink cannot be set
                        self._overlay_at(row, col).setdefault("links", []).append(("copy", link))
            else:
                row, col = coordinate_to_tuple(link.ref)
                if self._merge is not None:
                    if self._merge.merged_coord(row, col) is not None:
                        # WorksheetReader.normalize_merged_cell_link: first
                        # range of ws.merged_cells (a set) containing the cell
                        for rng in self.merged_cells:
                            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                                row, col = rng.min_row, rng.min_col
                                break
                        if self._merge.merged_coord(row, col) is not None:
                            # overlapping merged ranges: openpyxl fails too
                            raise AttributeError("'MergedCell' object attribute 'hyperlink' is read-only")
                self._overlay_at(row, col).setdefault("links", []).append(("same", link))

    def _bind_comments(self, archive, rels):
        for r in rels.find(COMMENTS_NS):
            src = archive.read(r.target)
            comment_sheet = CommentSheet.from_tree(fromstring(src))
            for ref, comment in comment_sheet.comments:
                row, col = coordinate_to_tuple(ref)
                if self._cell_exists_merged(row, col):
                    warnings.warn(COMMENT_WARNING.format(self.title, f"{get_column_letter(col)}{row}"))
                    continue
                self._overlay_at(row, col)["comment"] = comment

    def _bind_dimensions(self, parser):
        self.row_dimensions = DimensionHolder(worksheet=self, default_factory=self._add_row)
        self.column_dimensions = DimensionHolder(worksheet=self, default_factory=self._add_column)
        for col, cd in parser.column_dimensions.items():
            if "style" in cd:
                key = int(cd["style"])
                cd["style"] = self.parent._cell_styles[key]
            self.column_dimensions[col] = ColumnDimension(self, **cd)
        for row, rd in parser.row_dimensions.items():
            if "s" in rd:
                key = int(rd["s"])
                rd["s"] = self.parent._cell_styles[key]
            self.row_dimensions[int(row)] = RowDimension(self, **rd)

    def _bind_properties(self, parser):
        defaults = dict(
            print_options=PrintOptions(),
            page_margins=PageMargins(),
            page_setup=PrintPageSetup(worksheet=self),
            HeaderFooter=None,
            auto_filter=AutoFilter(),
            sheet_properties=WorksheetProperties(),
            views=SheetViewList(),
            sheet_format=SheetFormatProperties(),
            row_breaks=RowBreak(),
            col_breaks=ColBreak(),
            scenarios=ScenarioList(),
            protection=SheetProtection(),
        )
        from openrsxl.worksheet.header_footer import HeaderFooter

        defaults["HeaderFooter"] = HeaderFooter()
        for k, v in defaults.items():
            setattr(self, k, v)
        for k in (
            "print_options",
            "page_margins",
            "page_setup",
            "HeaderFooter",
            "auto_filter",
            "sheet_properties",
            "views",
            "sheet_format",
            "row_breaks",
            "col_breaks",
            "scenarios",
            "protection",
        ):
            v = getattr(parser, k, None)
            if v is not None:
                setattr(self, k, v)
        # full mode keeps the legacy drawing only for workbooks with VBA
        self.legacy_drawing = None

    def _ext_compute_bounds(self):
        """Cells created by the features extend the sheet like in full mode"""
        rows, cols = [], []
        if self._overlay:
            rows.extend(self._overlay)
            cols.extend(c for r in self._overlay.values() for c in r)
        if self._merge is not None:
            b = self._merge.bounds()
            rows += [b[0], b[1]]
            cols += [b[2], b[3]]
        if rows:
            self._ext_bounds = (min(rows), max(rows), min(cols), max(cols))

    # -- rows ----------------------------------------------------------------
    def _ext_cell_cls(self):
        return ReadOnlyCell if self._ext.cell_level else _ReadOnlyCell

    def _ext_extra_slots(self):
        extra = []
        if self._ext.read_hyperlinks:
            extra.append("hyperlink")
        if self._ext.read_comments:
            extra.append("comment")
        return extra

    def _cells_by_row(self, min_col, min_row, max_col, max_row, values_only=False):
        max_col = max_col or self.max_column
        max_row = max_row or self.max_row
        rows = self._ext_base_rows(min_col, min_row, max_col, max_row, values_only)
        if self._overlay or self._merge is not None:
            rows = self._apply_overlays(rows, min_col, min_row, max_col, max_row, values_only)
        if self._ext.create_empty_cells and not values_only:
            rows = self._create_empty(rows, min_row, min_col)
        return rows

    def _create_empty(self, rows, min_row, min_col):
        """
        create_empty_cells: the cell full mode creates when an empty position
        is accessed (``coordinate``, value None, default style), instead of
        read-only mode's shared EmptyCell; rows keep read-only mode's shape
        """
        # _new_cell, inlined (empty sheets can have millions of positions);
        # one default StyleArray for all of them, as read-only cells share
        # theirs through wb._cell_styles
        cell_cls = self._ext_cell_cls()
        new = cell_cls.__new__
        style = StyleArray()
        names = (["formula", "cached_value"] if self._ext.formula_and_value else []) + self._ext_extra_slots()
        empty = EMPTY_CELL
        ws = self

        def make(r, col):
            c = new(cell_cls)
            c.parent = ws
            c.row = r
            c.column = col
            c._value = None
            c.data_type = "n"
            c._style_id = style
            for name in names:
                setattr(c, name, None)
            return c

        for r, row in enumerate(rows, min_row):
            yield tuple([make(r, col) if c is empty else c for col, c in enumerate(row, min_col)])

    def _ext_base_rows(self, min_col, min_row, max_col, max_row, values_only):
        """openpyxl's read-only row stream (with extended cells)"""
        filler = None if values_only else (EMPTY_CELL if self._ext.cell_level else _EMPTY())
        empty_row = []
        if max_col is not None:
            empty_row = (filler,) * (max_col + 1 - min_col)

        counter = min_row
        idx = 1
        cell_cls = self._ext_cell_cls()
        extra = self._ext_extra_slots()
        with self._get_source() as src:
            parser = WorkSheetParser(
                src,
                self._shared_strings,
                data_only=self.parent.data_only,
                epoch=self.parent.epoch,
                date_formats=self.parent._date_formats,
                timedelta_formats=self.parent._timedelta_formats,
            )
            parser.formula_and_value = self._ext.formula_and_value
            rows = None
            if _native.ENABLED:
                reader = _native._rs.RowReader(src, parser, _native.reader_helpers())
                rows = self._native_rows(reader, parser, min_col, max_col, values_only, cell_cls, filler, extra)
            else:
                rows = self._python_rows(src, 0, min_col, max_col, values_only, filler)
            for idx, row in rows:
                if max_row is not None and idx > max_row:
                    break
                for _ in range(counter, idx):
                    counter += 1
                    yield empty_row
                if counter <= idx:
                    counter += 1
                    yield row

        if max_row is not None and max_row < idx:
            for _ in range(counter, max_row + 1):
                yield empty_row

    def _native_rows(self, reader, parser, min_col, max_col, values_only, cell_cls, filler, extra):
        actions = reader.actions
        consumed = 0
        while True:
            try:
                item = reader.next_row(self, min_col, max_col, values_only, cell_cls, filler, extra)
            except _native.NativeFallback:
                item = False
            if actions:
                pending = list(actions)
                del actions[:]
                try:
                    _native.replay(parser, pending)
                except _native.NativeFallback:
                    item = False
            if item is False:
                with self._get_source() as src:
                    yield from self._python_rows(src, consumed, min_col, max_col, values_only, filler)
                return
            if item is None:
                return
            consumed += 1
            yield item

    def _python_rows(self, src, skip, min_col, max_col, values_only, filler):
        """Exact Python implementation (also of formula_and_value)."""
        kw = dict(
            epoch=self.parent.epoch,
            date_formats=self.parent._date_formats,
            timedelta_formats=self.parent._timedelta_formats,
        )
        data_only = self.parent.data_only
        rows = WorkSheetParser(src, self._shared_strings, data_only=data_only, **kw).parse()
        if self._ext.formula_and_value:
            # the other reading of the same rows, warnings are reported once
            other = WorkSheetParser(self._get_source(), self._shared_strings, data_only=not data_only, **kw).parse()
            if data_only:
                pairs = ((r, o, r) for r, o in zip(rows, _quiet(other)))
            else:
                pairs = ((r, r, o) for r, o in zip(rows, _quiet(other)))
        else:
            pairs = ((r, None, None) for r in rows)
        cell_cls = self._ext_cell_cls()
        extra = self._ext_extra_slots()
        for n, (prim, form, val) in enumerate(pairs):
            if n < skip:
                continue
            idx, cells = prim
            yield idx, self._ext_get_row(cells, form, val, min_col, max_col, values_only, filler, cell_cls, extra)

    def _ext_get_row(self, row, form, val, min_col, max_col, values_only, filler, cell_cls, extra):
        if not row and not max_col:
            return ()
        max_col = max_col or row[-1]["column"]
        width = max_col + 1 - min_col
        new_row = [filler] * width
        fcells = form[1] if form is not None else None
        vcells = val[1] if val is not None else None
        for k, cell in enumerate(row):
            counter = cell["column"]
            if min_col <= counter <= max_col:
                i = counter - min_col
                if values_only:
                    new_row[i] = cell["value"]
                    continue
                c = cell_cls.__new__(cell_cls)
                _ReadOnlyCell.__init__(c, self, **cell)
                if fcells is not None:
                    f = fcells[k]
                    c.formula = f["value"] if f["data_type"] == "f" else None
                    c.cached_value = vcells[k]["value"]
                for name in extra:
                    setattr(c, name, None)
                new_row[i] = c
        return tuple(new_row)

    def _get_cell(self, row, column):
        """Cells are returned by a generator which can be empty"""
        for cells in self._cells_by_row(column, row, column, row):
            if cells:
                return cells[0]
        if self._ext.create_empty_cells:
            # beyond the rows of the file: full mode creates the cell too
            return self._new_cell(self._ext_cell_cls(), self._ext_extra_slots(), row, column)
        return EMPTY_CELL if self._ext.cell_level else _EMPTY()

    # -- overlays -------------------------------------------------------------
    def _apply_overlays(self, base, min_col, min_row, max_col, max_row, values_only):
        overlay = self._overlay
        merge = self._merge
        cell_cls = self._ext_cell_cls()
        extra = self._ext_extra_slots()
        if merge is not None and not values_only:
            merge.compute_styles()  # (cached for the whole sheet)
        # merged ranges sorted by start row, swept as rows advance
        pending = list(merge.by_row) if merge is not None else []
        pi = 0
        active = []
        row_idx = min_row - 1
        if self._ext_bounds is not None:
            # rows created by the features after the last row of the file
            last = self._ext_bounds[1]
            if max_row is not None:
                last = min(last, max_row)
            if max_col is not None:
                filler = None if values_only else (EMPTY_CELL if self._ext.cell_level else _EMPTY())
                pad = (filler,) * max(0, max_col + 1 - min_col)
            else:
                pad = ()
            base = _chain_rows(base, last, min_row, pad)
        for row in base:
            row_idx += 1
            if merge is not None:
                while pi < len(pending) and merge.ranges[pending[pi]].min_row <= row_idx:
                    active.append(pending[pi])
                    pi += 1
                if active:
                    active = [i for i in active if merge.ranges[i].max_row >= row_idx]
            specs = overlay.get(row_idx)
            if not specs and not active:
                yield row
                continue
            width = len(row)
            needed = width
            if specs:
                needed = max(needed, max(specs) + 1 - min_col)
            for i in active:
                needed = max(needed, merge.ranges[i].max_col + 1 - min_col)
            if max_col is not None:
                needed = min(needed, max_col + 1 - min_col)
            cells = list(row)
            if needed > width:
                filler = None if values_only else (EMPTY_CELL if self._ext.cell_level else _EMPTY())
                cells.extend([filler] * (needed - width))
            # 1. merged cells
            for i in active:
                rng = merge.ranges[i]
                for col in range(max(rng.min_col, min_col), min(rng.max_col, min_col + len(cells) - 1) + 1):
                    k = col - min_col
                    # the top-left cell stays a cell unless another
                    # (overlapping) range covers it
                    if (row_idx, col) == (rng.min_row, rng.min_col) and merge.merged_coord(row_idx, col) is None:
                        if values_only:
                            continue
                        shadow = merge.spec(i, row_idx, col)
                        cells[k] = self._styled_start(cells[k], row_idx, col, shadow, cell_cls, extra)
                    elif values_only:
                        cells[k] = None
                    else:
                        shadow = merge.spec(i, row_idx, col)
                        m = MergedCell(self, row_idx, col)
                        if shadow._style is not None:
                            m._style = copy(shadow._style)
                        cells[k] = m
            # 2. hyperlinks, 3. comments
            if specs:
                for col, spec in specs.items():
                    k = col - min_col
                    if not 0 <= k < len(cells):
                        continue
                    cells[k] = self._apply_spec(cells[k], row_idx, col, spec, values_only, cell_cls, extra)
            yield tuple(cells)

    def _new_cell(self, cell_cls, extra, row, col, value=None, data_type="n", style_id=None):
        # full mode creates these cells with the default style (all zero),
        # not the workbook's first cell style
        if style_id is None:
            style_id = StyleArray()
        c = cell_cls.__new__(cell_cls)
        _ReadOnlyCell.__init__(c, self, row, col, value, data_type, style_id)
        if self._ext.formula_and_value:
            c.formula = None
            c.cached_value = value
        for name in extra:
            setattr(c, name, None)
        return c

    def _styled_start(self, cell, row, col, shadow, cell_cls, extra):
        """top-left cell of a merged range with the borders of the range"""
        # its own StyleArray (see ReadOnlyCell.style_array)
        style_id = StyleArray(shadow._style) if shadow._style is not None else StyleArray()
        if not isinstance(cell, _ReadOnlyCell):
            # openpyxl creates the cell
            return self._new_cell(cell_cls, extra, row, col, style_id=style_id)
        cell._style_id = style_id
        return cell

    def _apply_spec(self, cell, row, col, spec, values_only, cell_cls, extra):
        links = spec.get("links")
        comment = spec.get("comment")
        if values_only:
            value = cell
            if links:
                for _how, link in links:
                    if value is None:
                        value = _filled(self, row, col, link)[0]
            return value
        if not isinstance(cell, (_ReadOnlyCell, _MergedCell)):
            cell = self._new_cell(cell_cls, extra, row, col)
        if isinstance(cell, _MergedCell):
            return cell
        if links:
            dual = self._ext.formula_and_value
            if dual:
                # openpyxl fills empty cells with the link target in each
                # reading (data_only=False / True) independently
                fv = cell.formula if cell.formula is not None else cell.cached_value
                dv = cell.cached_value
            for _how, link in links:
                link = copy(link)
                link.ref = cell.coordinate
                cell.hyperlink = link
                if cell._value is None:
                    value, data_type = _filled(self, row, col, link)
                    cell._value = value
                    cell.data_type = data_type
                if dual:
                    if fv is None:
                        fv, fdt = _filled(self, row, col, link)
                        if fdt == "f":
                            cell.formula = fv
                    if dv is None:
                        dv = _filled(self, row, col, link)[0]
                        cell.cached_value = dv
        if comment is not None:
            comment = copy(comment)
            comment.bind(cell)
            cell.comment = comment
        return cell


def _record(base, min_row, min_col, coords, out):
    """Pass rows through, keeping copies of the cells at `coords` as read
    from the file (before the features modify them)"""
    idx = min_row - 1
    for row in base:
        idx += 1
        for k, c in enumerate(row):
            if (idx, min_col + k) in coords and isinstance(c, _ReadOnlyCell):
                out[(idx, min_col + k)] = copy(c)
        yield row


def _quiet(gen):
    """Iterate `gen` with its warnings suppressed"""
    while True:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            try:
                item = next(gen)
            except StopIteration:
                return
        yield item


def _chain_rows(base, last, min_row, pad):
    idx = min_row - 1
    for row in base:
        idx += 1
        yield row
    for _ in range(idx + 1, last + 1):
        yield pad


def _EMPTY():
    from openrsxl.cell.read_only import EMPTY_CELL as base

    return base


def _filled(ws, row, col, link):
    """Value and data type openpyxl gives an empty cell receiving `link`"""
    c = Cell(ws, row=row, column=col)
    c.value = link.target or link.location
    return c._value, c.data_type
