# Changelog

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
