# Verification report

How each openrsxl release was verified before it was published. openpyxl
3.1.5 is the oracle throughout: every check compares openrsxl with openpyxl
itself.

## openrsxl 0.2.0

New: `create_empty_cells`, `install_as_openpyxl()` and a fix of the import
aliases of `openrsxl.extended` (CHANGELOG.md).

* Test suites: openrsxl's own suite, 2 644 tests, all pass (lxml); its
  reader, writer, extended and security tests also with ElementTree
  (2 229); openpyxl's ported suite: 2 588 pass, as for 0.1.0.
* `create_empty_cells` is compared with the cell openpyxl's full mode
  creates when the position is accessed - coordinate, value, data type,
  every style attribute, the enabled cell features - in every test of the
  extended oracle: all flags, each flag alone, each pair of the cell-level
  flags, 21 raw-XML edge cases, 60 random workbooks, random sub-rectangles.
  Deliberately broken variants (another default style, cells one row off,
  an `EmptyCell` left in place) are all detected by the oracle.
* Real-world files: all 1 972 files that the 0.1.0 runs matched (the
  corpus below, the XlsxWriter files with unusual options and six more test
  files of python-calamine), re-checked in extended mode with every flag
  (also with `data_only=True`) and with a random subset of the flags per
  file: 1 969 identical - among them the three heaviest, compared one at a
  time with larger limits; the other three declare a dimension of 17
  billion cells, so their streaming checks were skipped, as for 0.1.0.
* `install_as_openpyxl()` runs in fresh interpreters with every warning
  turned into an error: class identity, `isinstance` of formula values in
  full, read-only and streaming modes, saving through the alias (relative
  imports of aliased modules), refusal after the real openpyxl was
  imported, pandas' `read_excel(engine="openpyxl")` and `to_excel`.

## openrsxl 0.1.0

### Test suites

| suite | tests | result |
|---|---|---|
| openrsxl's own suite (`pytest`) | 2 638 | all pass (lxml and ElementTree back-ends) |
| openpyxl's suite, ported (`tests/upstream`) | 2 605 | 2 588 pass, the rest skipped / expected failures - the same as openpyxl's own run with the same packages (lxml back-end); also with ElementTree |
| Python fallbacks (`OPENRSXL_PURE_PYTHON=1`) | `test_extended`, smoke | all pass |
| hypothesis properties (API sequences, CellStore, XML / numbers / formulae) | 300 examples each in every run, 3 000 each before the release | all pass |

Interpreters: CPython 3.14 (all suites), 3.10 (locally built wheel), 3.14t
free-threaded (2 135 differential tests against the installed wheel); in CI
CPython 3.10 - 3.14 and 3.14t on Linux, and the built wheels on Linux,
macOS and Windows (x86-64 and ARM64). On PyPy the Rust engine is disabled
(it crashed PyPy's cpyext layer); openpyxl's Python code runs there.

### Real-world corpus

`tools/corpus.py` compares every file in 10 modes - full mode, `data_only`,
read-only (both), `rich_text`, `keep_vba`, `keep_links=False`, the saved
package byte for byte (right after loading and after reading every cell),
and `openrsxl.extended` with every flag (and a random subset) against
openpyxl's full mode.

| source | files | what they are |
|---|---|---|
| [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) test-suite | 994 | files saved by Excel: charts, images, comments, tables, autofilters, data validation, conditional formats, defined names, rich strings, ... |
| [Apache POI](https://github.com/apache/poi) `test-data/spreadsheet` | 367 | real-world files from bug reports, LibreOffice / Google / WPS exports, strict OOXML, encrypted files, fuzzer crash cases, XML bombs |
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) resources | 385 | .NET-written files, pivot tables, `TryToLoad` problem files, LibreOffice files |
| [calamine](https://github.com/tafia/calamine) tests | 70 | edge cases of a Rust reader (empty `s` attributes, huge dimensions, ...) |
| [pandas](https://github.com/pandas-dev/pandas) excel test data | 61 | files of pandas' openpyxl engine tests |
| [exceljs](https://github.com/exceljs/exceljs) | 37 | JavaScript-written files, a 150 000-row workbook |
| [python-calamine](https://github.com/dimastbk/python-calamine), openpyxl | 46 | |
| **total** | **1 960** | |

Results (lxml back-end): **all 1 960 files identical in all 10 modes**. Two
of them declare a dimension of 17 billion cells, so their read-only /
streaming checks were not run (read-only mode would yield 17 billion cells,
in both libraries). The five largest files - 2.1 million cells in 150 000
rows, thousands of merged ranges, a merged range covering a whole column
(for which openpyxl's full mode creates a million `MergedCell` objects), a
1 MB string shared by 12 000 cells - were compared one file at a time with
larger limits: the oracle (openpyxl's full mode, loaded several times per
check) needs up to 92 minutes per file for them.

ElementTree back-end (`OPENPYXL_LXML=False`), every 4th file (487 files,
full mode, `data_only`, read-only, extended): 486 identical; the oracle
(openpyxl loaded twice in full mode) exceeded the 3 GB limit of the run on
the remaining file, which is identical with the lxml back-end.

Files written by XlsxWriter with unusual options (constant-memory mode with
inline strings, rich strings, dynamic arrays, 1904 dates, cells at XFD /
row 1 048 576, chartsheets, protection, outline levels): 6 of 6 identical
in all 10 modes (the file with cells at XFD1048576 without the streaming
checks).

### Security review

* Hostile input (`tests/test_security.py`, 72 tests, every load mode and the
  extended features): XXE in every package part (no file content leaks),
  billion laughs, 100 000-level nesting at every level, 100 000 comments /
  processing instructions around the root, digit strings beyond Python's
  integer limits, coordinates and row numbers beyond every integer type,
  invalid style indices, malformed character references, relationship
  targets outside the package, duplicate zip members.
* Apache POI's fuzzer crash files and XML bombs (in the corpus): openpyxl's
  result or exception, no crash.
* Dependencies (OSV database, 2026-10-05): 15 locked Rust crates and the
  Python packages of the test environment - no known vulnerability. All
  dependencies, GitHub actions (pinned by commit, each pin verified against
  its release tag) and lint hooks are at their latest versions.
* Rust code: `cargo clippy -D warnings`; every `unsafe` block reviewed; no
  panic can reach Python (malformed input makes the engine fall back to
  openpyxl's code); checked arithmetic on row / column counters; no
  recursion on input structure; 32-bit (i686 / armv7) and big-endian
  (s390x) builds reviewed and compiled.

### Defects found and fixed by this verification

* Extended streaming: a malformed character reference (`&#x;`, `&#0;`, ...)
  in a `<row>` / `<c>` attribute made the pre-scan panic (`PanicException`)
  instead of raising openpyxl's `ParseError`.
* Extended streaming: row numbers beyond 64 bits raised `OverflowError` while
  loading (now handled by the Python scan, like openpyxl).
* Saving a loaded workbook without touching its cells wrote other style
  indices than openpyxl when openpyxl's own style registry quirk applies
  (custom number formats renumbered after hashing).
* Saving a workbook with row numbers beyond 64 bits raised `TypeError`.
* Row counters could wrap around after row 9 223 372 036 854 775 807 instead
  of falling back to Python's integers.
* Extended streaming: styles of merged cells were added to the workbook's
  style registry, changing the style ids openpyxl hands out later; cells
  created by a feature used the workbook's first cell style instead of the
  default style; `range.start_cell` was `None` for a cell whose coordinate
  disagrees with its row.
* The XML tokenizer recursed once per whitespace run around the root element
  (now a loop: no stack exhaustion on any target).
* The extension did not compile for Python 3.10 (pyo3-ffi declares
  `PyMemberDescrObject.d_member` with the wrong type before 3.11) nor for
  32-bit targets (a shift overflow), and crashed PyPy.
* Packaging: the source distribution would have included a debug-symbols
  file and openpyxl's source tree.
