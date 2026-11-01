# Changelog

## 0.3.1

* Performance (extended streaming): sheets with many merged cells,
  hyperlinks or both. Each merged range ran openpyxl's merge logic on a new
  worksheet; ranges that overlap no other one now share one result per
  shape and corner styles (the same cells and styles: the logic depends on
  nothing else), on one worksheet reused throughout. The ranges containing a
  row are looked up by row, and a hyperlink's relationship by Id instead of
  a scan of all relationships per link (also in full mode). 20k merged
  ranges: 7.2 s / +311 MB -> 0.9 s / +53 MB (openpyxl full mode: 2.5 s /
  +84 MB); 40k ranges, 20 per row: 20.1 s / +577 MB -> 1.4 s / +74 MB
  (openpyxl: 3.6 s / +78 MB); 20k hyperlinks: 8.8 s -> 0.8 s (openpyxl:
  9.4 s). Results unchanged; without the streaming flags nothing changes.

## 0.3.0

* `openrsxl.extended.llm`: lossless, LLM readable compression of workbooks.
  `compress_workbook(path, include=..., output=...)` streams every
  worksheet through the Rust engine and writes a text document with
  separable sections - values, formulas, styles, merged cells, hyperlinks,
  comments, dimensions, data validation, conditional formatting, tables,
  sheet and workbook settings, analysis - to memory, a `.llm.txt` and / or a
  `.llm.json` file. Each sheet is measured first (density, distinct values,
  repetition, estimated tokens) to choose its values encoding: `grid`
  (CSV-like rows), `index` (value -> ranges) or `hybrid`; formula fills are
  one line, periodic rows strided ranges. `decompress` reads a document back
  into cells and openpyxl objects (`to_workbook()` rebuilds a Workbook),
  `analyze_workbook` returns the measurements. Format and decoding guide:
  `docs/LLM_ENCODING.md`; the legend at the top of every document explains
  the encodings it uses. Grid rows are labelled only where the row number
  cannot be inferred (plus every 10th row, `row_anchor=`), values repeated
  down a column are `^` (`ditto=`), the style table is written once per
  workbook (`## styles`). `tools/llm_bench.py` measures tokens against the
  same information loaded uncompressed; `tools/make_financial_files.py`
  generates financial workbooks for tests and benchmarks. Opt-in filters
  spend no tokens on what is not shown: `skip_hidden=True` (cells of hidden
  rows and columns), `skip_hidden_sheets=True`, `print_area_only=True` (the
  print area and print titles); a filtered sheet says so in its header
  (`skip=hidden`, `area=...`). One sheet at a time: `sheets="Name"`
  / `sheets=[...]` compress only those worksheets (the others are not read,
  their metadata included), `compress_sheets(path)` yields one complete
  document per worksheet with the workbook opened once; `## sheet N` is the
  sheet's position in the workbook. The text form is the default and the
  cheapest output (JSON: 1-2 % more tokens on large workbooks). The
  `openrsxl` namespace is unchanged.
* `compress_workbook(..., rows=, columns=, ranges=)` (also
  `compress_sheets`, `analyze_workbook`): compress only these rows /
  columns / ranges, several segments each (`rows="1:100,200:300"`,
  `columns="A:D,F"`, `ranges="A1:D20,F5:H9"`), the same for every sheet or
  per sheet (`{sheet name: value}`); combined with the other filters; the
  sheet header says what was selected (`rows=`, `cols=`, `cells=`).
* Memory: `close()` of a read-only workbook drops its shared strings at
  once; the default sheet `load_workbook` discards no longer keeps the
  workbook in a reference cycle; a workbook closed and dropped is collected
  before the next `load_workbook` (openrsxl creates few Python objects, so
  Python's cycle collector could wait long). `openrsxl.extended.release(wb)`
  frees a workbook at once in any mode. `close()` keeps openpyxl's
  behaviour (a full-mode workbook stays usable); tests/test_memory.py
  compares every behaviour around `close()` with openpyxl. Rust: references
  could leak when building a read-only cell failed.
* Dependencies: pyo3 0.30.1; GitHub Actions setup-python 7.1.0, setup-uv
  10.3.3, upload-artifact 7.0.2, download-artifact 8.0.2, rust-cache 2.9.3,
  gh-action-pypi-publish 1.14.3; pre-commit ruff 0.17.3, black 26.10.1,
  pyupgrade 3.22.0, mypy 2.4.1.

## 0.2.2

* `openrsxl.extended.load_workbook(..., read_all=True)`: every streaming
  feature with one flag. A feature flag given explicitly wins over it
  (`read_all=True, read_comments=False`: everything but comments); the
  feature flags now default to `None` ("not given", following `read_all`,
  which defaults to `False`), so a call without flags is still exactly
  openpyxl's read-only mode.
* Fixed (extended): the empty cells and merged cells of a streamed sheet had
  the attributes of every cell feature (`cell.comment` was `None` without
  `read_comments`); they now have those of the enabled features only, like
  the other cells (`AttributeError` otherwise; `MergedCell` without
  `formula_and_value` is openpyxl's own).
* Fixed (extended): `create_empty_cells` created cells at positions before
  A1 (`iter_rows(min_row=-1)` gave cells `A-1`, `A0`); full mode has no cell
  there, so these stay read-only mode's `EmptyCell`.
* Fixed (extended): `pickle` of a streamed `MergedCellRange` failed on its
  class name; it now fails like openpyxl's read-only objects (the worksheet
  holds the open file).

## 0.2.1

* Fixed: after a failed `save` (eg. a timezone-aware datetime, TypeError)
  the temporary worksheet file stayed open until garbage collection
  (`ResourceWarning: unclosed file`, file not removable on Windows); it is
  now closed when the writer is discarded, as in openpyxl.
* Fixed: a write-only workbook discarded without `save` (eg. after an
  `append` that raised) kept its temporary worksheet file open until the
  process exited: the native row writer was part of a worksheet <-> writer
  reference cycle the garbage collector could not see (atexit then failed
  to remove the file on Windows). It now supports garbage collection.
* Fixed: `ws._cells` of loaded worksheets - its views were not reversible
  and had no `.mapping` (dict views have both) and a different repr; an
  unhashable key raised `unhashable type` instead of the message of a dict
  (Python 3.14: "cannot use 'list' as a dict key ..."); `pop()` of an
  empty store hashed the key (dict raises KeyError / returns the default).
* Memory: iterating the keys of `ws._cells` (`list(ws._cells)`, `items()`,
  `reversed`, `popitem`) built a new row and column int with every key
  tuple, and a new tuple on every pass. Key tuples now share their ints, as
  the keys of openpyxl's dict do, and a cell object keeps the `(row, col)`
  tuple it was looked up with. Peak memory of dict operations on a sheet
  with 1M cells: 515 MB -> 419 MB (openpyxl: 459 MB); time unchanged.
* `tests/test_api_parity.py`: API parity checks beyond cell values (see
  docs/DEVELOPMENT.md).

## 0.2.0

* `openrsxl.extended.load_workbook(..., read_only=True,
  create_empty_cells=True)`: every empty position of a row is a cell like
  the one openpyxl's full mode creates when the position is accessed -
  `coordinate`, `row`, `column`, value None, default style - instead of
  read-only mode's shared, coordinate-less `EmptyCell`. Created on the fly
  (no memory), about 0.2 µs per empty position. Without the flag nothing
  changes.
* `openrsxl.extended.install_as_openpyxl()`: called once at start-up, it
  makes `import openpyxl` (every `openpyxl.*` module, also inside other
  libraries such as pandas) import openrsxl, so `isinstance` checks written
  for openpyxl's classes - eg. `ArrayFormula` / `DataTableFormula` formula
  values - hold for openrsxl's objects. Without it the classes of the two
  libraries stay distinct (documented, with the attributes both share).
* Fixed: importing a module through its `openrsxl.extended.<module>` alias
  replaced the `__spec__` of the `openrsxl.<module>` module; its
  function-level relative imports (eg. when saving a workbook) then emitted
  `DeprecationWarning: __package__ != __spec__.parent`.

## 0.1.0

First release.

* `openrsxl`: the complete openpyxl 3.1.5 API (every module, class, function,
  signature, default, exception and warning) with a Rust engine for reading
  and writing cell data: same values, same Python types, byte-identical saved
  files, much faster loading / saving / iteration with a fraction of the
  memory.
* `openrsxl.extended`: openpyxl's API plus streaming (read-only) access to
  what openpyxl only offers in full mode, each behind its own flag:
  `formula_and_value`, `read_comments`, `read_hyperlinks`,
  `read_merged_cells`, `read_dimensions`, `read_sheet_properties`,
  `read_data_validations`, `read_conditional_formatting`, `read_tables`.
