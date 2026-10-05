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
  validation, conditional formatting and tables - each behind its own flag.

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
once (`import openrsxl as openpyxl`). Supported formats are openpyxl's:
`.xlsx`, `.xlsm`, `.xltx`, `.xltm`.

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

All flags are keyword-only, default to `False` and need `read_only=True`
(full mode already reads everything; `formula_and_value` raises `ValueError`
without `read_only=True`, the other flags are ignored).

| flag | what it adds (names and types of openpyxl's full mode) |
|---|---|
| `formula_and_value` | `cell.formula`: the formula as openpyxl reads it with `data_only=False` - a `str` starting with `=`, an `ArrayFormula` or a `DataTableFormula` - or `None` for cells without formula; `cell.cached_value`: the value openpyxl reads with `data_only=True` (the result Excel saved). Both in one pass, where openpyxl needs two full loads. |
| `read_comments` | `cell.comment` (`Comment` or `None`). |
| `read_hyperlinks` | `cell.hyperlink` (`Hyperlink` or `None`); as in full mode, an empty cell with a link gets the link target as its value. |
| `read_merged_cells` | `ws.merged_cells` (the ranges); cells covered by a range are `MergedCell` objects (value `None`); the top-left cell gets the borders of the range; `range.start_cell` is that cell. |
| `read_dimensions` | `ws.row_dimensions` and `ws.column_dimensions` (heights, widths, hidden, outline levels, styles). |
| `read_sheet_properties` | `ws.sheet_properties`, `ws.views`, `ws.freeze_panes`, `ws.sheet_format`, `ws.print_options`, `ws.page_margins`, `ws.page_setup`, `ws.HeaderFooter`, `ws.auto_filter`, `ws.protection`, `ws.row_breaks`, `ws.col_breaks`, `ws.scenarios`, `ws.print_title_rows` / `print_title_cols` / `print_area`. |
| `read_data_validations` | `ws.data_validations`. |
| `read_conditional_formatting` | `ws.conditional_formatting`. |
| `read_tables` | `ws.tables`. |

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
        if getattr(cell, "comment", None):       # EmptyCell / MergedCell have no comment
            print(cell.coordinate, cell.comment.author, cell.comment.text)
        if getattr(cell, "hyperlink", None):
            print(cell.coordinate, cell.hyperlink.target or cell.hyperlink.location)
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

The streaming features keep read-only mode's flat memory: all nine flags add
~2 MB to the 119 MB of `large_sample.xlsx` (whose peak is its shared strings
table) and 10x the cells give the same peak.

## Compatibility and limitations

openrsxl targets **openpyxl 3.1.5** and reproduces its observable behaviour
(values, types, errors, warnings, saved bytes - including openpyxl's bugs).
What differs, in full in [`docs/COMPATIBILITY.md`](https://github.com/AmirShokry/openrsxl/blob/main/docs/COMPATIBILITY.md):

* **Module names.** Classes live in `openrsxl.*` (`Font.__module__ ==
  "openrsxl.styles.fonts"`), so openpyxl and openrsxl objects cannot be
  mixed. Warnings carry the same category and message but other
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
* Rows and columns are those of read-only mode: cells outside the sheet's
  `<dimension>` are not produced (cells *created* by a feature - a merged or
  linked empty cell - extend `max_row` / `max_column`), cell positions follow
  the `<row>` elements, and a sheet that claims a huge dimension yields that
  many (empty) cells, as in openpyxl (use `iter_rows(max_row=..., max_col=...)`
  or `ws.reset_dimensions()`).
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
