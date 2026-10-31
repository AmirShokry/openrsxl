"""
openrsxl.extended.llm
=====================

Lossless, LLM readable compression of Excel workbooks.

``compress_workbook`` streams every worksheet through the Rust engine (no
cell objects, memory independent of the size of the sheet), measures it,
chooses the cheapest readable encoding for its values and writes a plain
text document - separable sections for values, formulas, styles and every
kind of metadata - that a language model reads directly and that
``decompress`` turns back into exactly what openpyxl reads::

    from openrsxl.extended.llm import compress_workbook, decompress

    doc = compress_workbook("book.xlsx")                 # everything, in memory
    prompt = doc.text()                                  # the whole document
    values = doc.text(sections=["values"])               # only the values
    doc["Sales"]["formulas"].text                        # one section
    doc = compress_workbook("book.xlsx", include="values", output="text")  # book.llm.txt
    doc = compress_workbook("book.xlsx", include=["formulas", "styles"],
                            output="all", path="out/book")  # memory + .txt + .json
    book = decompress(doc)            # or decompress("out/book.llm.json")
    book["Sales"].values[(2, 3)], book["Sales"].formulas[(2, 4)]
    wb = book.to_workbook()           # an openpyxl Workbook again

Lossless means: for every included section, decoding gives back the same
values (type and value, as openpyxl reads them with ``data_only=True``), the
same formulas (``data_only=False``: text, ``ArrayFormula``,
``DataTableFormula``), the same styles (font, fill, border, number format,
alignment, protection, named style of every cell), and openpyxl objects
equal to the ones openpyxl's full mode loads for merged cells,
hyperlinks, comments, dimensions, data validation, conditional formatting,
tables, sheet settings and workbook settings. Nothing is rounded, sampled
or truncated. Cells are positioned by their own coordinate (as in full
mode) and hold what the file contains (values in merged cells that are not
top-left cells are kept; hyperlinks do not change values). Not encoded:
charts, images and drawings, pivot tables, VBA, external link caches,
rich text runs inside cells (the text is kept), theme XML, and package
internals such as relationship ids.

Sections
--------

``include`` / ``exclude`` take section names or groups (``"all"``, the
default; ``"cells"`` = values, formulas, styles; ``"metadata"`` = the rest):

=========================  ==================================================
section                    content
=========================  ==================================================
``values``                 cell values (formula cells: their cached result)
``formulas``               formulas, fills grouped into ranges
``styles``                 style table + map of the cells' styles
``merged``                 merged ranges
``hyperlinks``             per-cell hyperlinks (grouped into ranges)
``comments``               per-cell comments (notes)
``dimensions``             column widths, row heights, hidden, outline
``validations``            data validation rules
``conditional_formatting`` conditional formatting rules and their formats
``tables``                 Excel tables
``sheet``                  sheet settings: views / frozen panes, page setup,
                           print titles / area, protection, auto filter...
``workbook``               sheet list and states, defined names, views,
                           calculation settings, document properties
``analysis``               the measurements and the chosen encoding
=========================  ==================================================

Outputs
-------

``output``: ``"memory"`` (default: the returned CompressedWorkbook holds the
text), ``"text"`` (``<path>.llm.txt``), ``"json"`` (``<path>.llm.json``),
``"all"`` or a list of them.

**Give a model the text form** (``doc.text()`` or the ``.llm.txt`` file): it
is the default and the cheapest. The JSON form (opt-in: ``output="json"``,
``doc.to_json()``) holds the same section texts inside a JSON envelope:
1-2 % more tokens on large workbooks (escaped line breaks and quotes,
keys, the analysis repeated as an object), about 50 % more on small ones
(tiktoken o200k_base; docs/LLM_ENCODING.md has the measurements). Use it
when a program consumes the document. ``path`` defaults to the workbook's path
without its extension. Without ``"memory"`` the values are streamed to the
files as they are encoded and nothing is kept: memory stays flat whatever
the size of the sheet. The JSON form holds the same section texts:
``{"format", "source", "sections", "legend", "workbook", "sheets": [{"index",
"name", "dimension", "analysis", "sections": {"values": {"encoding",
"start", "text"}, "formulas": {"text"}, ...}}]}``.

Filters (opt-in, not lossless)
------------------------------

To spend no tokens on what a reader of the sheet never sees:

``skip_hidden=True``
    leaves out the cells of hidden rows and hidden columns (rows marked
    hidden, collapsed outline groups; with ``sheetFormatPr zeroHeight`` every
    row without a visible row dimension): their values, formulas, styles,
    hyperlinks and comments.
``skip_hidden_sheets=True``
    leaves out hidden and very hidden worksheets (also when named in
    ``sheets=``). The ``workbook`` section still lists them with their state.
``print_area_only=True``
    on sheets with a print area, keeps only the cells inside it, plus the
    print titles (title rows across the area's columns, title columns across
    its rows). Sheets without a print area are kept whole.
``rows=`` / ``columns=`` / ``ranges=``
    keep only the cells of these rows, columns and ranges - several
    segments each: ``rows="1:100,200:300"`` (or ``5``, ``(1, 100)``, a list),
    ``columns="A:D,F"`` (letters or numbers), ``ranges="A1:D20,F5:H9"`` or
    ``ranges=["A1:D20", "F5:H9"]`` (start:end rectangles, encoded as blocks
    of rows, not cell by cell; ``"K2"`` one cell, ``"B:D"`` whole columns,
    ``"3:7"`` whole rows).
    A cell is kept when it is in the rows, the columns and the ranges given.
    One value applies to every sheet; a dict ``{sheet name: value}`` gives
    each sheet its own (sheets not in it are not restricted).

The filters apply before the analysis, so the encoding is chosen for the
cells that remain (``analysis["skipped_cells"]`` counts the others). A
filtered sheet says so in its header - ``## sheet 1 "S" dim=A1:F40
skip=hidden area=A1:F40+titles:1:2`` - (JSON: ``"filter"``; decoded:
``DecodedSheet.filter``) and the legend explains it: the left-out cells are
not empty, only not shown. Range metadata (merged cells, dimensions,
validations, conditional formats, tables, sheet settings) is kept whole.
Without filters, or when a flag finds nothing to leave out, the document is
exactly the lossless one.

How the values encoding is chosen
---------------------------------

Pass 1 streams the sheet once and measures it: number of cells, bounding
box and density (cells / area), types, distinct values and how they cluster
into rectangles of one repeated value, and the estimated token cost of each
encoding:

``grid``
    rows as ``ROW:v1|v2|...`` (positional like CSV, empty cells cost almost
    nothing, long gaps become a column jump ``K=v``, repeats ``v*N``). The
    cheapest form of dense tables and of most real sheets.
``index``
    the inverted index of SpreadsheetLLM (Microsoft, 2024): one line
    ``RANGES=v`` per distinct value. Wins on sparse sheets and on sheets
    dominated by a few repeated values (forms, matrices of flags / zeros).
``hybrid``
    the large rectangles of one repeated value (constant columns, repeated
    categories in sorted data) as ``RANGES=v`` lines, everything else as a
    grid. Wins on dense tables with long runs of identical values.

``index`` is chosen only when it is clearly cheaper than the grid (below
``0.85 x`` grid), ``hybrid`` below ``0.92 x`` - rows read as rows are easier
for a model, so the grid wins ties. ``encoding=`` forces one. The analysis
(and the reason of the choice) is in the ``analysis`` section and in
``sheet.analysis``.

How to read a document (for AI agents)
--------------------------------------

Every document starts with ``#openrsxl-llm/1`` and (unless
``legend=False``) a legend explaining exactly the encodings it uses. The
complete grammar - ``LEGEND`` below, and docs/LLM_ENCODING.md - in short:

* ``## sheet N "Name" dim=A1:F120`` opens a sheet (N: its position in the
  workbook, as in ``## workbook``), ``### <section>`` a
  section (``### values grid start=A``: the encoding and the grid's first
  column).
* Value tokens: ``12`` integer, ``-3.5`` / ``1e-05`` float, ``TRUE`` /
  ``FALSE``, ``#N/A`` (``#``...) error, ``2024-01-05`` date (midnight),
  ``2024-01-05 13:45:00`` datetime, ``13:45:00`` time, ``"..."`` text in
  JSON syntax (only when bare text would be ambiguous), ``@date:``,
  ``@td:`` (duration), ``@err:"..."``, ``@float:inf``; anything else is
  text exactly as written.
* grid: ``ROW:f1|f2|...`` - f1 is column ``start``, then one column per
  field; empty field = empty cell; ``COL=v`` jumps to COL; ``v*N`` = N
  cells with v; ``^`` = the value of the cell above (the previous line,
  same column; never on a labelled line, so a value is never more than 9
  lines up); a line without ``ROW:`` is the row after the previous
  line (every row that is a multiple of 10 is labelled, so a row number is
  never more than 9 lines away); missing rows are empty.
* index: ``RANGES=v`` - v in every cell of RANGES.
* hybrid: ``RANGES=v`` block lines, then grid lines from the first
  ``ROW:`` line on (block cells are empty in the grid).
* formulas: ``D2:D100=B2*C2`` - D2 has ``=B2*C2``, D3 ``=B3*C3`` ... (the
  relative references shift like an Excel fill, ``$`` parts stay);
  ``E5~"..."`` a formula as JSON string; ``F1[array F1:F3]=...`` array
  formula; ``G1[dataTable ...]`` data table.
* styles: the workbook's ``## styles`` table, written once before the
  sheets - ``font fN:`` / ``fill pN:`` / ``border bN:`` parts, ``sN:`` cell
  styles (keys not given are those of ``s0``); each sheet's ``### styles``
  has ``RANGES=sN`` lines mapping its cells to styles (all other cells:
  ``s0``). ``CompressedWorkbook.styles`` is the table.
* metadata sections are one item per line in compact XML: ``name(...)``
  element, ``k=v`` attribute, ``k:v`` child with a ``val`` attribute, bare
  ``k`` empty child, ``"..."`` text - the OOXML vocabulary Excel uses.

Functions
---------

``compress_workbook(filename, *, include="all", exclude=None, sheets=None,
encoding="auto", output="memory", path=None, legend=True, jump_gap=4,
runs=True, row_anchor=10, ditto=True, index_limit=64 MiB, keep_links=True,
skip_hidden=False, skip_hidden_sheets=False, print_area_only=False,
rows=None, columns=None, ranges=None)``
    -> ``CompressedWorkbook``. ``sheets``: the worksheet(s) to compress, a
    name or a list of names (default: all; written in workbook order).
    The other sheets are not read. ``jump_gap``: empty cells after which the grid
    writes a column jump. ``runs``: ``v*N`` repeats. ``row_anchor``: grid lines
    of consecutive rows drop their ``ROW:`` label, except rows that are
    multiples of ``row_anchor`` (1: every line labelled). ``ditto``: ``^``
    for a value equal to the one above it. ``index_limit``: memory
    the analysis may use for the inverted index of a sheet (above it, the
    sheet has too many distinct values for an index to win: grid).
    ``skip_hidden`` / ``skip_hidden_sheets`` / ``print_area_only`` /
    ``rows`` / ``columns`` / ``ranges``: see Filters.
``compress_sheets(filename, *, sheets=None, ...)``
    the same options; yields ``(sheet name, CompressedWorkbook)`` one
    worksheet at a time, in workbook order: the workbook is opened once and
    each document is complete on its own (equal to ``compress_workbook(...,
    sheets=[name])``). Files: ``<path>.<sheet name>.llm.txt`` / ``.json``.
``analyze_workbook(filename, sheets=None, encoding="auto", **options)``
    -> {sheet name: analysis dict} (pass 1 only: nothing is encoded).
``decompress(source)``
    -> ``DecodedWorkbook`` from a CompressedWorkbook, its text or JSON, or a
    .txt / .json path.
``estimate_tokens(text)`` / ``count_tokens(text, encoding="o200k_base")``
    estimated (built in) / exact (tiktoken) token counts.
"""

from ._decode import (
    DecodedSheet,
    DecodedWorkbook,
    DecodeError,
    StyleDef,
    decompress,
    fill_formula,
    parse_token,
)
from ._encode import (
    GROUPS,
    SECTIONS,
    CompressedSheet,
    CompressedWorkbook,
    Section,
    compress_sheets,
    compress_workbook,
    count_tokens,
    estimate_tokens,
)
from ._legend import FORMAT, LEGEND, legend

__all__ = [
    "FORMAT",
    "GROUPS",
    "LEGEND",
    "SECTIONS",
    "CompressedSheet",
    "CompressedWorkbook",
    "DecodeError",
    "DecodedSheet",
    "DecodedWorkbook",
    "Section",
    "StyleDef",
    "analyze_workbook",
    "compress_sheets",
    "compress_workbook",
    "count_tokens",
    "decompress",
    "estimate_tokens",
    "fill_formula",
    "legend",
    "parse_token",
]


def analyze_workbook(filename, sheets=None, encoding="auto", **options):
    """{sheet name: analysis} - the measurements of pass 1 and the values
    encoding compress_workbook would choose, without encoding anything
    (``options``: compress_workbook's, eg. skip_hidden=True)"""
    doc = compress_workbook(filename, include=["analysis"], sheets=sheets, encoding=encoding, legend=False, **options)
    return {s.name: s.analysis for s in doc.sheets}
