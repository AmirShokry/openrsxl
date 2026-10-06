# Changelog

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
