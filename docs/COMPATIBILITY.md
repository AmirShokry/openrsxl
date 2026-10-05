# Compatibility with openpyxl

openrsxl targets **openpyxl 3.1.5**. The goal is a 100% match of everything
observable through the public API: names, signatures, return values and their
Python types, exceptions (type and message), warnings (category and message)
and the bytes written by `Workbook.save`.

This document lists every known difference. None of them affects the values,
types or files produced.

## Differences that cannot be avoided

1. **Module paths.** Classes live in `openrsxl.*` instead of `openpyxl.*`, so
   `Font.__module__ == "openrsxl.styles.fonts"` and default reprs read
   `<openrsxl.styles.fonts.Font object> ...`. Objects of the two libraries are
   distinct types: an `openpyxl.styles.Font` cannot be assigned to an openrsxl
   cell (the typed descriptors reject it, just as they would reject any other
   foreign type).
2. **Warning locations.** Warnings carry the same category and message, but
   their `filename`/`lineno` point into openrsxl's modules. Warnings raised
   while reading worksheets (unsupported extensions, out-of-range dates, ...)
   are emitted by `openrsxl.worksheet._reader`, as in openpyxl, but from a
   different line. Filters based on message, category or module name
   (`openrsxl...` instead of `openpyxl...`) work as usual.
3. **Tracebacks.** Frames of the Rust engine do not appear in tracebacks and
   the chain of frames above an exception is different. The exception object
   itself (type, message, args) is identical.
4. **`openrsxl.extended`.** An extra namespace (see README); it does not exist
   in openpyxl.
5. **Python implementations.** CPython 3.10 or later. The Rust engine
   creates Python objects by writing their slots directly (CPython object
   layout): on PyPy and GraalPy (no wheels; built from the source
   distribution) it is disabled and openpyxl's Python code runs - the same
   results without the speed-up. On free-threaded CPython
   (3.14t) the engine runs without the GIL: different workbooks can be used
   by different threads in parallel, while one worksheet used *concurrently*
   by several threads raises `RuntimeError: Already borrowed` (Rust borrow
   checking) where openpyxl's dict-based objects would race silently.
6. **`Worksheet._cells` of loaded worksheets** (a private attribute) is a
   `openrsxl.worksheet._cellstore.CellStore` instead of a `dict`. It keeps
   cells in compact form (calamine style) and creates `Cell` objects on first
   access. It implements the complete dict interface with dict semantics
   (insertion order, `KeyError`, views, `popitem`, iteration errors when
   mutated, keys equal to integer pairs such as `(1.0, 2)` ...), so code
   using `ws._cells` like a dict keeps working, but `type(ws._cells) is dict`
   / `isinstance(ws._cells, dict)` are false (`collections.abc.MutableMapping`
   is true). Worksheets created with `Workbook()` / `create_sheet()` still use
   a plain dict.

## Deliberate choices (for byte-identical behaviour)

* `__version__` is `"3.1.5"`, the openpyxl version whose contract is
  implemented (code checking the openpyxl version keeps working). The
  version of the openrsxl package itself is
  `importlib.metadata.version("openrsxl")`.
* Strings that openpyxl writes or reports are kept verbatim, including those
  that mention openpyxl: the default document `creator` (`"openpyxl"`), error
  messages such as *"openpyxl does not support the old .xls file format"*,
  warnings such as *"Workbook contains no stylesheet, using openpyxl's
  defaults"*, and the `openpyxl.` prefix of temporary files.
* The environment variables `OPENPYXL_LXML` and `OPENPYXL_DEFUSEDXML` are
  honoured exactly as openpyxl does, so both libraries pick the same XML
  backend in a given environment. openrsxl reproduces the output of *both*
  backends (lxml and ElementTree differ in many small ways, e.g. `<v />` vs
  `<v></v>`, character references for non-ASCII text).
* openpyxl's quirks are reproduced on purpose (bug compatibility), e.g.
  `_x005F_` sequences being stripped from shared strings, the sheet prefix
  of a defined name being dropped when a shared formula is translated,
  temporary files that cannot be removed after a failed save on Windows,
  string truncation at 32,767 characters, `int("1_000")`-style number parsing.

## Edge cases handled by the Python fallback

The Rust engine hands over to the original Python implementation for inputs
it does not handle natively (DOCTYPE declarations, non-UTF-8 encodings,
malformed XML, rich text, unusual value types, ...). Results are identical,
with these theoretical exceptions:

* **Read-only mode, malformed XML found after rows were already produced:**
  the Python parser resumes the stream after the rows already returned; any
  warnings emitted for top-level worksheet elements *before* that point
  (e.g. "extension is not supported") are emitted a second time.
* **Full mode, a `<row>` element outside `<sheetData>`:** handled by the
  Python implementation; warnings for top-level elements preceding it may be
  emitted twice.
* **Malformed XML exactly at a 16 KB block boundary.** openpyxl's parser
  (`ElementTree.iterparse`) feeds expat 16 KB at a time and a syntax error
  hides every row completed in the same block. openrsxl reproduces this
  block logic; only when the offending token itself straddles a block
  boundary could the *last* row before the error be reported differently in
  read-only mode.

Neither situation occurs in files produced by Excel, LibreOffice or openpyxl.

## Not differences

* Peak memory and speed: openrsxl is faster and uses the same or (much)
  less memory in every measured scenario (see BENCHMARKS.md). `Cell`
  objects of a loaded sheet are created on first access instead of during
  `load_workbook`; even that first pass over all cells is faster than
  openpyxl's iteration over its pre-built cells.
* Threads: workbooks, worksheets, cells and read-only generators can be used
  from any thread, as with openpyxl. Like openpyxl, openrsxl holds the GIL
  while working.
* `copy.copy` / `copy.deepcopy` / `pickle` of cells, worksheets and
  workbooks behave like openpyxl's (a deep-copied or unpickled loaded
  worksheet holds a plain dict of cells).
* Bugs of openpyxl 3.1.5 are reproduced, not fixed (eg. temporary files left
  behind after a failed save on Windows).

## openrsxl.extended streaming features

`openrsxl.extended.load_workbook(..., read_only=True, <flags>)` reproduces,
feature by feature, what openpyxl's full mode returns for the same file; the
test oracle compares every cell (type, value, data type, styles, formula,
cached value, hyperlink, comment) and every worksheet attribute with openpyxl
(`tests/test_extended.py`, `tests/oracle/extended.py`). Known limitations and
choices:

* **Names.** openpyxl has no attribute for "the formula" or "the cached
  value" of a cell in any mode (the formula is the `value` with
  `data_only=False`, the cached value is the `value` with `data_only=True`).
  `cell.formula` and `cell.cached_value` are therefore new names; their
  values are exactly those two openpyxl values. Every other added attribute
  reuses openpyxl's full-mode name (`hyperlink`, `comment`, `merged_cells`,
  `row_dimensions`, `freeze_panes`, ...).
* **Disabled features** do not exist: `cell.formula` without
  `formula_and_value=True` raises `AttributeError`, like any attribute
  missing in openpyxl's read-only mode (the added attributes are plain slots,
  read at C speed). With no flag set, `openrsxl.extended.load_workbook`
  returns the plain openpyxl read-only objects.
* **Full mode.** The `read_*` flags have no effect without `read_only=True`
  (full mode reads everything anyway). `formula_and_value=True` requires
  `read_only=True` (`ValueError` otherwise).
* **values_only.** `iter_rows(values_only=True)` yields `value` (following
  `data_only`); the other features are only visible on cell objects, except
  that merged cells are `None` and hyperlink targets fill empty cells, as in
  full mode.
* **Dimensions.** As in openpyxl's read-only mode, rows and columns come from
  the sheet's `<dimension>` element; file cells outside it are not produced.
  Cells *created* by a feature (merged cells, linked or commented empty cells,
  as full mode creates them) extend `max_row` / `max_column` when needed.
* **MergedCellRange.start_cell** is the top-left cell as full mode has it
  (value, borders of the range, hyperlink, comment, formula / cached value),
  including full mode's quirk for overlapping ranges (the range keeps the
  cell that a later range replaced by a `MergedCell`). It is read lazily: the
  first access streams the rows of the sheet that hold top-left cells once,
  then the cells of all ranges are kept (one per range). The ranges are
  instances of a subclass named `MergedCellRange`
  (`isinstance(r, openrsxl.worksheet.merge.MergedCellRange)` is true,
  `type(r) is MergedCellRange` is not). As with every streamed cell, the
  object is not the one a later `iter_rows` yields; changing it changes
  nothing else.
* **Shared state.** Cells, hyperlinks and comments are new objects on every
  iteration (as all cells of read-only mode); a hyperlink or comment object
  is bound to the cell yielded in that iteration.
* **Overlapping merged ranges** (invalid for Excel) are handled like full
  mode, including the `AttributeError` openpyxl raises at load time for a
  single-cell hyperlink inside them.
* **Not streamed yet:** images, charts, drawings and pivot tables
  (`ws._images`, `ws._charts`, `ws._pivots`), `legacy_drawing` (kept by
  openpyxl only with `keep_vba`; always `None` here) and rich text
  (`rich_text=True` has the same effect as in openpyxl's read-only mode).
* **Warnings.** Warnings of worksheet content (unsupported extensions,
  invalid dates, ...) are emitted while iterating, as in openpyxl's read-only
  mode, not while loading. The "merged range but has a comment" warning of
  `read_comments` + `read_merged_cells` is emitted while loading, as in full
  mode.
* **Malformed cell data.** Full mode parses all cells at load time, so a
  worksheet whose `<sheetData>` is malformed fails in `load_workbook`. The
  streaming features read cell data while iterating, like openpyxl's
  read-only mode: the same exception is raised while iterating (or while
  loading, when the Python fallback has to read the whole sheet, eg. for
  malformed XML). Data that only full mode rejects - eg. cells with `s=""`
  or a style index beyond the stylesheet, which read-only mode accepts until
  a style is used - behaves as in read-only mode; a feature that needs the
  faulty data (the style of a merged range's corner cell) raises full mode's
  error while iterating.
* **Cell positions** follow read-only mode: a cell whose coordinate
  disagrees with its `<row>` element (eg. `r="A0"` in `<row r="1">`) is
  yielded where the row puts it, as openpyxl's read-only mode does (full
  mode files it under its own coordinate). `range.start_cell` is always the
  cell `iter_rows` yields at the top-left position.
* **Cells created or restyled by a feature** (a hyperlink or comment on an
  empty cell gets full mode's default style; the top-left cell of a merged
  range gets the borders of the range) carry their own style: their private
  `_style_id` attribute holds a `StyleArray` instead of an index into
  `wb._cell_styles`, which streaming never changes (so the style ids
  openpyxl hands out later stay the same).
* **Huge dimensions.** As in openpyxl's read-only mode, iteration covers the
  sheet's declared `<dimension>`: a sheet claiming `A1:XFD1048576` yields 17
  billion (empty) cells. Bound the iteration (`iter_rows(max_row=...,
  max_col=...)`) or call `ws.reset_dimensions()`.
* **Cost.** Each feature is only paid when enabled. Sheet-level features
  scan the parts of the sheet outside `<sheetData>` once at load time (the
  cell data is skipped with byte searches). `formula_and_value` interprets
  each formula cell twice (formula and cached value) in the same pass;
  `read_merged_cells` computes the styles of merged cells once per worksheet
  with openpyxl's own merge code on a 3x3 representative of each range.
