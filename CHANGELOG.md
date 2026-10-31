# Changelog

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
