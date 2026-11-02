# openrsxl

[![PyPI](https://img.shields.io/pypi/v/openrsxl.svg)](https://pypi.org/project/openrsxl/)
[![Python versions](https://img.shields.io/pypi/pyversions/openrsxl.svg)](https://pypi.org/project/openrsxl/)
[![CI](https://github.com/AmirShokry/openrsxl/actions/workflows/CI.yml/badge.svg)](https://github.com/AmirShokry/openrsxl/actions/workflows/CI.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/AmirShokry/openrsxl/blob/main/LICENSE)

**openrsxl** is a drop-in replacement for [openpyxl](https://openpyxl.readthedocs.io)
3.1.5 whose performance-critical parts are written in Rust.

* **Same API** - every module, class, function, signature, default, constant,
  exception and warning of openpyxl exists under the same name in `openrsxl`.
* **Same results** - the same values with the same Python types (an empty
  cell is `None`), the same errors for broken files, and byte-for-byte the
  same files when saving.
* **Faster and lighter** - load + save is 7-12x faster with 6-17x less
  memory, read-only streaming 4-13x faster ([benchmarks](#performance)).
* **`openrsxl.extended`** - openpyxl's API plus things openpyxl cannot do:
  streaming (read-only mode) access to formulas *and* their cached values,
  comments, hyperlinks, merged cells, dimensions, sheet properties, data
  validation, conditional formatting, tables and full mode's cells at empty
  positions - each behind its own flag; and `openrsxl.extended.llm`,
  lossless compression of workbooks into compact text for language models.

## Installation

```bash
pip install openrsxl
```

Pre-built wheels for CPython 3.10 - 3.14 (including free-threaded 3.14t) on
Linux (glibc: x86_64, aarch64, i686, armv7, ppc64le, s390x; musl: x86_64,
aarch64, armv7), macOS (x86_64, arm64) and Windows (x64, x86, arm64); no Rust
toolchain needed. The only runtime dependency is `et_xmlfile`, like
openpyxl. `lxml`, `defusedxml`, `pillow`, `numpy` and `pandas` are optional
exactly as for openpyxl.

## Quick start

Replace `openpyxl` by `openrsxl` in your imports - nothing else changes:

```python
import openrsxl                                  # instead of: import openpyxl
from openrsxl.styles import Font, PatternFill    # instead of: from openpyxl.styles import ...

# read
wb = openrsxl.load_workbook("report.xlsx")
ws = wb["Sales"]
print(ws["A1"].value)           # same value and Python type as openpyxl
print(ws["Z999"].value)         # None for an empty cell, as in openpyxl
for row in ws.iter_rows(min_row=2, values_only=True):
    print(row)

# write
ws["B2"] = 42
ws["B2"].font = Font(bold=True)
ws["C2"] = "=B2*2"
ws.merge_cells("D2:E3")
wb.save("report-out.xlsx")      # byte-identical to what openpyxl writes

# stream a huge file with constant memory (openpyxl's read-only mode)
wb = openrsxl.load_workbook("huge.xlsx", read_only=True)
total = sum(row[3] or 0 for row in wb.active.iter_rows(min_row=2, values_only=True))
wb.close()

# write a huge file (openpyxl's write-only mode)
wb = openrsxl.Workbook(write_only=True)
ws = wb.create_sheet("log")
for i in range(1_000_000):
    ws.append([i, f"event {i}", i * 0.5])
wb.save("log.xlsx")
```

Existing code using openpyxl can also keep its imports and alias the module
once (`import openrsxl as openpyxl`), or switch a whole application -
including libraries that import openpyxl themselves, such as pandas'
`read_excel(engine="openpyxl")` - with one call at start-up, before
anything imports openpyxl:

```python
import openrsxl.extended
openrsxl.extended.install_as_openpyxl()   # `import openpyxl` now imports openrsxl

from openpyxl.worksheet.formula import ArrayFormula   # openrsxl's class
```

Then `isinstance` checks written for openpyxl's classes hold for openrsxl's
objects (without it, the classes of the two libraries are distinct, see
below). Supported formats are openpyxl's: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`.

See [`examples/quickstart.py`](https://github.com/AmirShokry/openrsxl/blob/main/examples/quickstart.py) for a complete,
runnable example.

## `openrsxl.extended`: streaming with full-mode features

openpyxl offers two ways to read a file:

* **full mode** (`load_workbook(path)`) gives everything (formulas, comments,
  hyperlinks, merged cells, ...) but keeps every cell in memory - a 100 MB
  file can take minutes and gigabytes;
* **read-only mode** (`read_only=True`) streams rows with constant memory but
  drops everything except values and styles.

`openrsxl.extended` is the same API as `openrsxl` (`openrsxl.extended.styles`,
`openrsxl.extended.load_workbook`, ... are the same objects) whose
`load_workbook` accepts extra keyword-only flags that add full-mode features to
read-only streaming. Rows are still streamed, so memory does not grow with the
number of cells; each feature is only paid for when enabled, and without any
flag the result is exactly openpyxl's read-only mode.

```python
from openrsxl.extended import load_workbook

wb = load_workbook("book.xlsx", read_only=True,
                   formula_and_value=True,   # formula and saved result in one pass
                   read_comments=True, read_hyperlinks=True, read_merged_cells=True)
ws = wb["Sheet1"]
for row in ws.iter_rows():
    for cell in row:
        cell.value            # follows data_only, exactly as in openpyxl
        cell.formula          # "=SUM(A1:A9)" (or None)
        cell.cached_value     # the result saved in the file, e.g. 45
        cell.hyperlink        # Hyperlink or None
        cell.comment          # Comment or None
print(ws.merged_cells.ranges)
wb.close()
```

### The flags

All flags are keyword-only and need `read_only=True` (full mode already
reads everything: the flags are ignored there, except an explicit
`formula_and_value=True`, which raises `ValueError`).

**`read_all=True` turns every feature on.** A feature flag you pass
explicitly always wins over `read_all`, so you can start from everything and
switch features off, or leave `read_all` off and switch features on:

```python
from openrsxl.extended import load_workbook

load_workbook(path, read_only=True, read_all=True)                       # every feature
load_workbook(path, read_only=True, read_all=True, read_comments=False,
              create_empty_cells=False)                                   # all but these two
load_workbook(path, read_only=True, read_comments=True)                  # only comments
load_workbook(path, read_only=True)                                       # none: openpyxl's read-only mode
```

The feature flags default to `None` ("not given": follow `read_all`, which
defaults to `False`); `True` / `False` enable / disable a feature whatever
`read_all` says.

| flag | what it adds (names and types of openpyxl's full mode) |
|---|---|
| `read_all` | every feature below that is not given explicitly. |
| `formula_and_value` | `cell.formula`: the formula as openpyxl reads it with `data_only=False` - a `str` starting with `=`, an `ArrayFormula` or a `DataTableFormula` - or `None` for cells without formula; `cell.cached_value`: the value openpyxl reads with `data_only=True` (the result Excel saved). Both in one pass, where openpyxl needs two full loads. |
| `read_comments` | `cell.comment` (`Comment` or `None`). |
| `read_hyperlinks` | `cell.hyperlink` (`Hyperlink` or `None`); as in full mode, an empty cell with a link gets the link target as its value. |
| `read_merged_cells` | `ws.merged_cells` (the ranges); cells covered by a range are `MergedCell` objects (value `None`); the top-left cell gets the borders of the range; `range.start_cell` is that cell. |
| `read_dimensions` | `ws.row_dimensions` and `ws.column_dimensions` (heights, widths, hidden, outline levels, styles). |
| `read_sheet_properties` | `ws.sheet_properties`, `ws.views`, `ws.freeze_panes`, `ws.sheet_format`, `ws.print_options`, `ws.page_margins`, `ws.page_setup`, `ws.HeaderFooter`, `ws.auto_filter`, `ws.protection`, `ws.row_breaks`, `ws.col_breaks`, `ws.scenarios`, `ws.print_title_rows` / `print_title_cols` / `print_area`. |
| `read_data_validations` | `ws.data_validations`. |
| `read_conditional_formatting` | `ws.conditional_formatting`. |
| `read_tables` | `ws.tables`. |
| `create_empty_cells` | a cell at every empty position, like the one full mode creates when the position is accessed: `coordinate`, `row`, `column`, value `None`, default style (and `None` for the enabled cell features), instead of read-only mode's shared, coordinate-less `EmptyCell`. Created on the fly (nothing is kept); rows keep read-only mode's shape. |

**Compatibility rule.** Everything openpyxl's read-only mode already has
(`iter_rows`, `rows`, `values`, `ws.cell()`, `ws["A1:C3"]`, `max_row`,
`calculate_dimension()`, `reset_dimensions()`, cell attributes, errors)
keeps its interface, return types and errors with any flag, so code written
for openpyxl's read-only mode runs unchanged; a feature only changes what
openpyxl's full mode would return differently for the same file (a cell
covered by a merged range is a `MergedCell`, an empty position is a cell
with `create_empty_cells`, cells created by a feature can extend
`max_row` / `max_column`).

**Why `create_empty_cells` is not the default.** It gives full mode's cells,
but the default must stay exactly openpyxl's read-only mode (the shared
`EmptyCell`), so that `openrsxl.extended.load_workbook` is a drop-in
replacement without flags. `read_all=True` includes it; add
`create_empty_cells=False` to keep read-only mode's `EmptyCell` while
enabling everything else (sparse sheets: no object per empty position).

### Access options

These change how a sheet is accessed rather than what a cell holds, so
`read_all` does not include them (default `False`, read-only workbooks):

| option | effect |
|---|---|
| `exact_dimensions` | `min_row`, `max_row`, `min_column`, `max_column` - and so the rows and columns of `iter_rows()` - from the cells present, as full mode computes them (every cell of the file, also one with a style only, plus the cells the enabled features create), instead of the sheet's `<dimension>` element. Computed by the load-time pre-scan (byte searches, no cell objects). |
| `random_access` | `ws.cell()` and `ws[...]` (a cell, a range, a column `"B"`, a row `5`) without reading the sheet again on every call. The first random access builds a store of the sheet's cells in one streaming pass; later calls, and `iter_rows`, are served from it with the same cell objects. Ranges, rows, columns and `iter_rows` hold every position of their bounds, as full mode (read-only mode's `iter_rows` stops at the last row of the file); `ws.iter_cols()` / `ws.columns` exist with it. `ws._cells` is a read-only mapping `{(row, column): cell}` of the cells full mode keeps (`.get`, `in`, `items()`, `len`, `bool`). `openrsxl.extended.release_cells(ws)` (or a workbook) frees the store. |
| `reopen` | a workbook closed with `close()` reopens its file (the path, or the file object it was loaded from) when it is used again, as a full-mode workbook stays usable after `close()`. Its shared strings are kept. A pass started before `close()` goes on in any case. |

`skip_invalid_defined_names=True` (any mode) skips a print titles defined
name that openpyxl cannot parse (eg. `#N/A`, which makes openpyxl's
`load_workbook` fail) with a warning.

```python
from openrsxl.extended import load_workbook, release_cells

wb = load_workbook("book.xlsx", read_only=True, read_all=True,
                   exact_dimensions=True, random_access=True, reopen=True)
ws = wb["Data"]
print(ws.max_row, ws.max_column)               # the cells' bounds, as full mode
print(ws.cell(5, 2).value, ws["B5"].coordinate)  # one pass on first use, then a lookup
total = sum(c.value or 0 for c in ws["C"])     # a column, every position
present = ws._cells.get((5, 2))                # None where full mode has no cell
release_cells(ws)                              # free the store (rebuilt on demand)
```

Streaming workbooks can be shared by threads (a GIL build of Python): each
pass reads its own stream of the file, and what is built lazily (the
random access store, merged-cell styles, `start_cell`) is built once under a
per-sheet lock. Warnings of a sheet's data (eg. "Data Validation extension
is not supported") are reported by its first complete pass only, as full
mode reports them once per load.

Unlike full mode, reading a position never creates a cell: full mode's
`iter_rows(max_row=50)` or `ws.cell(50, 1)` add cells and raise `max_row`;
streaming leaves the sheet as the file has it.

Every value is what openpyxl's full mode returns for the same file - same
attribute names, same types, same values. The only new names are
`cell.formula` and `cell.cached_value`, because openpyxl has no attribute for
them in any mode (it exposes the formula or the cached value as `cell.value`,
depending on `data_only`).

### Examples

**Formulas and their results in one pass** - `cell.value` keeps following
`data_only`:

```python
from openrsxl.extended import load_workbook

wb = load_workbook("book.xlsx", read_only=True, formula_and_value=True)
for row in wb.active.iter_rows(min_row=2, max_col=4):
    for cell in row:
        if cell.formula is not None:
            print(cell.coordinate, cell.formula, "=", cell.cached_value)
# with data_only=True, cell.value is the cached value (as in openpyxl) and
# cell.formula still gives the formula
```

**Comments and hyperlinks:**

```python
from openrsxl.extended import load_workbook

wb = load_workbook("book.xlsx", read_only=True, read_comments=True, read_hyperlinks=True)
for row in wb.active.iter_rows():
    for cell in row:
        if cell.comment:                         # None if absent (empty and merged cells too)
            print(cell.coordinate, cell.comment.author, cell.comment.text)
        if cell.hyperlink:
            print(cell.coordinate, cell.hyperlink.target or cell.hyperlink.location)
```

**Every position as a cell, as in full mode** - read-only mode yields one
shared `EmptyCell` without coordinate for every empty position:

```python
from openrsxl.extended import load_workbook

wb = load_workbook("book.xlsx", read_only=True, create_empty_cells=True)
for row in wb.active.iter_rows():
    for cell in row:
        print(cell.coordinate, cell.value)       # empty positions too: None, default style
```

**Merged cells:**

```python
from openrsxl.extended import load_workbook

wb = load_workbook("book.xlsx", read_only=True, read_merged_cells=True)
ws = wb.active
for rng in ws.merged_cells.ranges:
    print(rng.coord, rng.start_cell.value)       # the top-left cell, as full mode has it
for row in ws.iter_rows(values_only=True):
    print(row)                                   # covered cells are None
```

**Sheet layout and rules:**

```python
from openrsxl.extended import load_workbook

wb = load_workbook("book.xlsx", read_only=True, read_dimensions=True, read_sheet_properties=True,
                   read_data_validations=True, read_conditional_formatting=True, read_tables=True)
ws = wb.active
print(ws.column_dimensions["A"].width, ws.row_dimensions[1].height, ws.freeze_panes)
for dv in ws.data_validations.dataValidation:
    print(dv.type, dv.formula1, dv.sqref)
for cf in ws.conditional_formatting:
    print(cf.sqref, [rule.type for rule in cf.rules])
for table in ws.tables.values():
    print(table.displayName, table.ref)
```

A feature that is not enabled does not exist, exactly like the attributes
openpyxl's read-only cells do not have: `cell.formula` without
`formula_and_value=True` raises `AttributeError`. With `values_only=True`,
rows hold `cell.value` only (merged cells are `None`, link targets fill empty
linked cells, as in full mode).

[`examples/extended_streaming.py`](https://github.com/AmirShokry/openrsxl/blob/main/examples/extended_streaming.py) is a
runnable tour of every flag.

### Freeing memory

`wb.close()` behaves as in openpyxl: a read-only workbook closes its file
(openrsxl also drops its shared strings at once), a full-mode workbook stays
usable - so its cells stay in memory while it is referenced. A workbook and
its worksheets reference each other, so a dropped workbook waits for
Python's cycle collector; openrsxl runs it before the next `load_workbook`
when a closed workbook is still in memory. To return the memory of a
workbook at once, in any mode:

```python
from openrsxl.extended import release

release(wb)        # closes it, empties its cells, breaks the references; wb is unusable afterwards
```

## `openrsxl.extended.llm`: lossless compression for language models

`compress_workbook` turns a workbook into a compact text document that a
language model reads directly and that decodes back to exactly what
openpyxl reads - values with their types, formulas, every cell's style,
merged cells, hyperlinks, comments, dimensions, data validation,
conditional formatting, tables, sheet and workbook settings - each in its
own section, so you send a model only what it needs:

```python
from openrsxl.extended.llm import compress_sheets, compress_workbook, decompress

doc = compress_workbook("book.xlsx")                       # all sections, in memory
prompt = doc.text()                                        # the document (with a legend)
compress_workbook("book.xlsx", include="values").text()    # values only
compress_workbook("book.xlsx", include=["formulas", "styles"], output="all", path="out/book")
                                                           # memory + out/book.llm.txt + .llm.json
book = decompress("out/book.llm.json")                     # back to cells and openpyxl objects
book["Sales"].cell("D2")                                   # value, formula, style, link, comment

compress_workbook("book.xlsx", sheets="Sales")             # one worksheet (the others are not read)
for name, doc in compress_sheets("book.xlsx"):             # one document per worksheet, workbook opened once
    ...
```

```
## styles
...
s1: font=f1 fill=p2
## sheet 1 "Sales" dim=A1:E27
### values hybrid start=A
A2:A13=North
A14:A25=South
1:Region|Month|Units|Price|Revenue
|2024-01-01|6|2.5
|2024-02-01|9|2.5
...
### formulas
E2:E25=C2*D2
E27=SUM(E2:E25)
### styles
A1:E1=s1
```

* **Streamed by the Rust engine** - no cell objects, flat memory; written to
  files as it goes (`large_sample.xlsx`, 98 MB, 10 M cells: 28 s, 212 MB peak).
* **Measured per sheet** - density, distinct values, repetition, estimated
  tokens - to pick the cheapest readable encoding: a CSV-like `grid`, the
  `index` of SpreadsheetLLM (value -> ranges) for sparse or repetitive
  sheets, or `hybrid` (repeated blocks + grid). Formula fills become one
  line (`E2:E25=C2*D2`), periodic rows a strided range (`B6:Q98/4`).
* **Small** - against the same information loaded uncompressed (cells as
  JSON with their types, styles.xml, the metadata XML), the whole document
  saves 97 % of the tokens on `sample.xlsx`, 93 % on `large_sample.xlsx`
  and 84 % on generated financial workbooks; values alone 80-92 %, formulas
  99.4-99.97 %, styles 94-99.9 % (`tools/llm_bench.py`). Ledgers stay
  row-local: a value repeated down a column is `^`, and rows are labelled
  only where the number cannot be inferred (and every 10th row).
* **Lossless by default, focused on request** - `skip_hidden=True` leaves
  out hidden rows and columns, `skip_hidden_sheets=True` hidden sheets,
  `print_area_only=True` everything outside the print area (and print
  titles), `rows="1:100,200:300"` / `columns="A:D,F"` /
  `ranges=["A1:D20", "F5:H9"]` keep only those segments (per sheet with a dict; `merged` then
  lists only the merged ranges they touch); a filtered sheet is marked as such in the document.
  `style_keys=["font.b", "fill", "numFmt", ...]` writes only those style keys and
  `resolve_colors=True` turns theme / indexed colours into RGB.
* **Built for citing cells** - `cite_cells=True` names every value's cell: the
  `index` encoding unless it costs more than 1.3 x the grid, otherwise a grid
  with `COL=v` on every field and a row label on every line (still lossless).
  Models copying addresses from a positional grid were seen to shift values by a
  column; this leaves nothing to count.
* **Cheapest form by default** - the text document (`doc.text()`,
  `output="text"`). JSON is opt-in (`output="json"`, `doc.to_json()`) for
  programs: the same content costs 1-2 % more tokens on large workbooks and
  about 50 % more on small ones.
* **Self-describing** - every document starts with a legend explaining the
  encodings it uses (`legend=False` + `legend(doc)` to send it once for several
  documents); dates are written as openpyxl reads them (`2025-04-06 00:00:00`);
  the format is specified in
  [`docs/LLM_ENCODING.md`](https://github.com/AmirShokry/openrsxl/blob/main/docs/LLM_ENCODING.md).

[`examples/llm_compression.py`](https://github.com/AmirShokry/openrsxl/blob/main/examples/llm_compression.py)
is a runnable tour.

## Performance

Windows 10, Python 3.14, openpyxl 3.1.5 with lxml; each measurement in a
fresh process, one at a time ([`docs/BENCHMARKS.md`](https://github.com/AmirShokry/openrsxl/blob/main/docs/BENCHMARKS.md) has
every file and mode, and how to reproduce). Memory: the process' peak
resident set size, with its increase over the baseline (before importing the
library) in parentheses.

| `large_sample.xlsx` - 98 MB, 25 sheets, 10 M cells | openpyxl | openrsxl |
|---|---|---|
| load + save | 228.0 s - 4 617 MB (+4 599) | **18.2 s - 291 MB (+273)** |
| load + iterate all values | 146.8 s - 4 594 MB (+4 576) | **7.8 s - 359 MB (+340)** |
| read-only: iterate all values | 101.8 s - 149 MB (+131) | **7.8 s - 119 MB (+100)** |
| formula *and* cached value of every cell | full mode, loaded twice: > 270 s, > 4.5 GB | **`openrsxl.extended`, one streaming pass: 11.3 s - 119 MB (+100)** |

| 100 000-row files written by openpyxl, and a 0.6 M cell workbook | speed-up | memory |
|---|---|---|
| load + save | 7.0x - 12.3x | 5.6x - 12x less |
| load + iterate all values | 7.8x - 13.8x | 3x - 11x less |
| read-only: iterate all values / cells | 3.7x - 9.6x | less (43-49 MB vs 51-61 MB) |
| write-only: write 200 000 rows | 2.6x (fill), same save | the same (43 MB) |

The streaming features keep read-only mode's flat memory: all nine `read_*` /
`formula_and_value` flags add ~2 MB to the 119 MB of `large_sample.xlsx`
(whose peak is its shared strings table) and 10x the cells give the same
peak. `create_empty_cells` adds no memory and ~0.2 µs per empty position.

## Compatibility and limitations

openrsxl targets **openpyxl 3.1.5** and reproduces its observable behaviour
(values, types, errors, warnings, saved bytes - including openpyxl's bugs).
What differs, in full in [`docs/COMPATIBILITY.md`](https://github.com/AmirShokry/openrsxl/blob/main/docs/COMPATIBILITY.md):

* **Module names.** Classes live in `openrsxl.*` (`Font.__module__ ==
  "openrsxl.styles.fonts"`), so openpyxl and openrsxl objects cannot be
  mixed, and `isinstance` checks against openpyxl's classes are false - eg.
  for the `ArrayFormula` / `DataTableFormula` values of formula cells -
  unless `openrsxl.extended.install_as_openpyxl()` makes openpyxl's names
  point to openrsxl (above; the classes then also report openpyxl's module
  names, eg. in default reprs). Code that must accept both libraries' objects
  can test the attributes they share (`value.t == "array"`, `value.ref`,
  `value.text`). Warnings carry the same category and message but other
  `filename`/`lineno`; tracebacks differ.
* **`openrsxl.__version__` is `"3.1.5"`**, the openpyxl version whose API is
  implemented (code checking the openpyxl version keeps working); the
  package's own version is `importlib.metadata.version("openrsxl")`.
* **`ws._cells` of loaded worksheets** (a private attribute) is a
  dict-compatible `CellStore` (compact cells, `Cell` objects created on first
  access): `isinstance(ws._cells, dict)` is false, everything else behaves
  like the dict.
* **CPython 3.10+.** The Rust engine writes Python objects directly
  (CPython object layout). On other implementations (PyPy, GraalPy - built
  from source) it is disabled and openpyxl's Python code runs: same results,
  no speed-up. `OPENRSXL_PURE_PYTHON=1` disables the engine on CPython too
  (debugging aid).
* **Free-threaded Python (3.14t):** the engine runs without the GIL; one
  worksheet used *concurrently* by several threads raises `RuntimeError:
  Already borrowed` instead of racing (openpyxl objects are not thread-safe
  either). Different workbooks in different threads work in parallel.
* **Python fallbacks.** Input the engine cannot reproduce with certainty
  (DOCTYPEs, non UTF-8 encodings, malformed XML, ...) is handled by
  openpyxl's own code; in read-only mode, warnings for elements before a
  malformed spot can be emitted twice.

`openrsxl.extended` (streaming features):

* `cell.formula` and `cell.cached_value` are new names (see above); a
  disabled feature raises `AttributeError`.
* Rows and columns are those of read-only mode (unless `exact_dimensions`:
  full mode's bounds): cells outside the sheet's
  `<dimension>` are not produced (cells *created* by a feature - a merged or
  linked empty cell - extend `max_row` / `max_column`), cell positions follow
  the `<row>` elements, and a sheet that claims a huge dimension yields that
  many (empty) cells, as in openpyxl (use `iter_rows(max_row=..., max_col=...)`
  or `ws.reset_dimensions()`).
* Empty positions are read-only mode's coordinate-less `EmptyCell` unless
  `create_empty_cells=True`, which creates full mode's cell for each of
  them (about 0.2 µs per empty position, no memory).
* Cell data that makes full mode fail in `load_workbook` (eg. an invalid
  style index) behaves as in read-only mode: the error is raised when the
  faulty data is used - when the style is read, or while iterating if a
  feature needs it (the style of a merged range's corner cell).
* Merged ranges are instances of a `MergedCellRange` subclass (`isinstance`
  holds); `start_cell` is computed on first access; cells and their links /
  comments are new objects on every iteration.
* Not streamed (yet): images, charts, drawings, pivot tables,
  `ws.legacy_drawing`; `rich_text=True` behaves as in read-only mode.

## Security

openrsxl reads untrusted files the way openpyxl does, and is tested for it:

* The Rust XML tokenizer knows only the five predefined entities and
  character references; it never resolves external entities or touches the
  network or the file system. Documents with a `DOCTYPE` are parsed by
  openpyxl's own code path (expat with its entity-expansion limits, lxml with
  `resolve_entities=False`, defusedxml when installed and enabled) - exactly
  as openpyxl.
* Hostile input is tested in every load mode against openpyxl
  ([`tests/test_security.py`](https://github.com/AmirShokry/openrsxl/blob/main/tests/test_security.py)): XXE in every part,
  billion laughs, 100 000-level element nesting, 100 000 comments around the
  root, numbers and coordinates beyond every integer limit, path traversal in
  relationships, duplicate zip members - and the fuzzing corpus of Apache POI
  (its crash test cases and XML bombs). Malformed input never panics or
  crashes the interpreter: it raises openpyxl's exception.
* Like openpyxl, decompression is not size-limited: a "zip bomb" is read as
  far as the data asks for (read-only and streaming modes keep memory flat).

## How it works

openrsxl = openpyxl's object model (ported, MIT licensed - see
[`NOTICE`](https://github.com/AmirShokry/openrsxl/blob/main/NOTICE)) + a Rust engine for everything that touches cell data:

| openpyxl (pure Python) | openrsxl |
|---|---|
| `xml.etree` iterparse of worksheets | streaming Rust XML tokenizer (`src/xml.rs`) |
| `WorkSheetParser.parse_cell` / `bind_cells` | `src/sheet.rs` |
| `Worksheet._cells` dict of `Cell` objects | `CellStore` (`src/store.rs`): compact cells, `Cell` objects created on access |
| shared strings table | `src/strings.rs`, `src/text.rs` |
| shared-formula `Translator` / `Tokenizer` | `src/formula.rs` (compiled templates) |
| `from_excel` date conversion | `src/dates.rs` |
| `<sheetData>` serialisation (lxml and ElementTree output) | `src/writer.rs` |
| read-only row streaming | `RowReader` in `src/lib.rs` |
| (extended) sheet parts around `<sheetData>` | `src/prescan.rs` |

Like [calamine](https://github.com/tafia/calamine) /
[python-calamine](https://github.com/dimastbk/python-calamine), cells of a
loaded worksheet are kept in compact Rust form (10 bytes per cell, shared
strings by index, formulae de-duplicated into position-independent
templates) and Python objects are only created when needed; `save()` and
`iter_rows(values_only=True)` work directly on the compact form. Whenever
the engine meets input it cannot reproduce with certainty it hands over to
openpyxl's original Python code, so the result is always openpyxl's.

## How it is tested

openpyxl itself is the test oracle (a test dependency only): every test
compares openrsxl with openpyxl.

* openpyxl's complete test-suite, ported (2 600 tests), with lxml and with
  ElementTree;
* 2 600+ differential tests: every public name and signature, ~375
  one-feature reader fixtures in 5 load modes, every value type written,
  hostile input, `openrsxl.extended` against full mode, hypothesis-driven
  fuzzing of XML, numbers, formulae and random sequences of API calls on
  loaded workbooks;
* a corpus of ~1 960 real-world files from the test-suites of XlsxWriter
  (Excel-made files), Apache POI, ClosedXML, calamine, pandas, exceljs,
  python-calamine and openpyxl, compared in 10 modes each (`tools/corpus.py`):
  every cell, style and worksheet attribute, warnings, errors and the saved
  package byte for byte; `openrsxl.extended` against openpyxl's full mode.

```bash
pip install -e .[test]                 # or: maturin develop --release
pytest                                 # differential, security, fuzz tests
cd tests/upstream && pytest            # openpyxl's own test-suite
python tools/corpus.py <directory>     # whole-file differential run
```

[`docs/TESTING.md`](https://github.com/AmirShokry/openrsxl/blob/main/docs/TESTING.md) is the verification report of this
release (corpus, results, security review); see
[`docs/DEVELOPMENT.md`](https://github.com/AmirShokry/openrsxl/blob/main/docs/DEVELOPMENT.md) for the layout and tools.

## License

MIT. openrsxl contains code derived from openpyxl 3.1.5 (MIT licensed,
copyright (c) 2010 openpyxl); see [`NOTICE`](https://github.com/AmirShokry/openrsxl/blob/main/NOTICE).
