# Benchmarks

openrsxl vs openpyxl 3.1.5 on the same machine and Python (Windows 10, 16
cores, Python 3.14, lxml installed, so both libraries use their lxml
back-end). Every number comes from a fresh process running one operation,
one process at a time. Memory is reported as **peak RSS (+increase)**: the
peak resident set size of the process (Windows: peak working set) and, in
parentheses, its increase over the baseline measured before the library was
imported (~19 MB of Python and psutil). Saving writes to a file on disk.

Reproduce with:

```bash
python tools/make_bench_files.py benchfiles --rows 100000
python tools/bench_all.py benchfiles/*.xlsx your-file.xlsx      # openpyxl vs openrsxl
python tools/bench_tasks.py your-file.xlsx                      # full mode vs openrsxl.extended
python tools/bench_extended.py your-file.xlsx all --cells       # streaming features
python tools/bench_write.py openpyxl; python tools/bench_write.py openrsxl [--write-only]
```

Files:

* `bench_*.xlsx` - 100 000 rows each, written by openpyxl
  (`tools/make_bench_files.py`): `numbers`, `strings`, `dates`, `sparse` (3
  values in 200 columns per row), `styled`, `formulas` (including formulae
  with unique constants).
* `sample.xlsx` - a private real-world workbook (not distributed): 6 MB, 9
  sheets, 0.6 M cells, many formulae, shared strings, a table, comments,
  dates.
* `large_sample.xlsx` - a private real-world workbook (not distributed):
  98 MB, 25 sheets, 10 M cells, 670 MB of sheet XML, 5.6 M formulae, shared
  strings, charts, tables, comments, conditional formatting, extensions.

## Reading and saving existing files

Time is the total of the operation, with load + iteration / save in
parentheses.

| file | operation | openpyxl time | openrsxl time | speed-up | openpyxl peak RSS (+increase) | openrsxl peak RSS (+increase) |
|---|---|---|---|---|---|---|
| sample.xlsx | load + save | 11.66 s (5.88 + 5.78) | 1.49 s (0.66 + 0.83) | 7.8x | 368 MB (+350) | 59 MB (+40) |
| sample.xlsx | load + iterate values | 6.20 s (5.77 + 0.43) | 0.79 s (0.67 + 0.13) | 7.8x | 336 MB (+317) | 56 MB (+37) |
| sample.xlsx | load + iterate cells | 6.34 s (5.87 + 0.47) | 1.09 s (0.69 + 0.41) | 5.8x | 335 MB (+316) | 261 MB (+242) |
| sample.xlsx | read-only: iterate values | 4.17 s (0.41 + 3.75) | 0.51 s (0.04 + 0.47) | 8.2x | 51 MB (+32) | 48 MB (+30) |
| sample.xlsx | read-only: iterate cells | 4.59 s (0.43 + 4.16) | 0.65 s (0.04 + 0.61) | 7.1x | 51 MB (+33) | 49 MB (+30) |
| bench_numbers.xlsx | load + save | 14.72 s (7.63 + 7.09) | 1.64 s (0.45 + 1.19) | 9.0x | 529 MB (+510) | 61 MB (+42) |
| bench_numbers.xlsx | load + iterate values | 7.82 s (7.08 + 0.74) | 0.60 s (0.52 + 0.08) | 13.1x | 438 MB (+419) | 58 MB (+39) |
| bench_numbers.xlsx | load + iterate cells | 7.78 s (7.01 + 0.78) | 0.89 s (0.49 + 0.40) | 8.7x | 439 MB (+420) | 324 MB (+305) |
| bench_numbers.xlsx | read-only: iterate values | 6.50 s (2.11 + 4.39) | 0.92 s (0.40 + 0.52) | 7.1x | 61 MB (+42) | 43 MB (+24) |
| bench_numbers.xlsx | read-only: iterate cells | 6.69 s (1.98 + 4.72) | 1.10 s (0.40 + 0.71) | 6.1x | 60 MB (+42) | 43 MB (+25) |
| bench_strings.xlsx | load + save | 23.28 s (14.47 + 8.81) | 1.89 s (1.12 + 0.77) | 12.3x | 560 MB (+541) | 77 MB (+58) |
| bench_strings.xlsx | load + iterate values | 14.90 s (14.10 + 0.80) | 1.21 s (1.10 + 0.10) | 12.4x | 470 MB (+451) | 77 MB (+58) |
| bench_strings.xlsx | load + iterate cells | 15.31 s (14.53 + 0.78) | 1.51 s (1.11 + 0.40) | 10.1x | 470 MB (+451) | 354 MB (+335) |
| bench_strings.xlsx | read-only: iterate values | 14.17 s (2.78 + 11.39) | 1.48 s (0.45 + 1.03) | 9.6x | 60 MB (+42) | 43 MB (+24) |
| bench_strings.xlsx | read-only: iterate cells | 14.65 s (2.75 + 11.89) | 2.07 s (0.50 + 1.57) | 7.1x | 61 MB (+42) | 44 MB (+25) |
| bench_dates.xlsx | load + save | 13.30 s (6.45 + 6.85) | 1.89 s (0.76 + 1.13) | 7.0x | 412 MB (+393) | 64 MB (+45) |
| bench_dates.xlsx | load + iterate values | 7.09 s (6.56 + 0.53) | 0.62 s (0.52 + 0.10) | 11.5x | 366 MB (+347) | 60 MB (+42) |
| bench_dates.xlsx | load + iterate cells | 7.08 s (6.55 + 0.53) | 0.87 s (0.54 + 0.33) | 8.2x | 366 MB (+347) | 254 MB (+235) |
| bench_dates.xlsx | read-only: iterate values | 6.23 s (1.52 + 4.71) | 0.89 s (0.29 + 0.60) | 7.0x | 53 MB (+34) | 43 MB (+25) |
| bench_dates.xlsx | read-only: iterate cells | 6.52 s (1.54 + 4.98) | 1.13 s (0.31 + 0.82) | 5.8x | 52 MB (+33) | 43 MB (+25) |
| bench_sparse.xlsx | load + save | 4.94 s (2.59 + 2.35) | 0.68 s (0.17 + 0.51) | 7.2x | 203 MB (+185) | 52 MB (+33) |
| bench_sparse.xlsx | load + iterate values | 40.34 s (2.32 + 38.01) | 3.81 s (0.17 + 3.64) | 10.6x | 4129 MB (+4110) | 1362 MB (+1343) |
| bench_sparse.xlsx | load + iterate cells | 40.13 s (2.29 + 37.85) | 7.38 s (0.17 + 7.20) | 5.4x | 4129 MB (+4110) | 3400 MB (+3382) |
| bench_sparse.xlsx | read-only: iterate values | 2.51 s (0.74 + 1.77) | 0.43 s (0.15 + 0.28) | 5.8x | 53 MB (+34) | 44 MB (+25) |
| bench_sparse.xlsx | read-only: iterate cells | 2.86 s (0.74 + 2.12) | 0.77 s (0.14 + 0.63) | 3.7x | 53 MB (+34) | 43 MB (+24) |
| bench_styled.xlsx | load + save | 16.82 s (9.39 + 7.43) | 1.71 s (0.65 + 1.07) | 9.8x | 452 MB (+433) | 61 MB (+43) |
| bench_styled.xlsx | load + iterate values | 9.89 s (9.31 + 0.58) | 0.72 s (0.65 + 0.07) | 13.8x | 379 MB (+361) | 61 MB (+42) |
| bench_styled.xlsx | load + iterate cells | 9.97 s (9.39 + 0.59) | 1.28 s (0.93 + 0.35) | 7.8x | 380 MB (+361) | 281 MB (+262) |
| bench_styled.xlsx | read-only: iterate values | 8.77 s (1.97 + 6.81) | 1.09 s (0.39 + 0.70) | 8.1x | 52 MB (+34) | 43 MB (+24) |
| bench_styled.xlsx | read-only: iterate cells | 9.71 s (2.17 + 7.54) | 1.24 s (0.39 + 0.84) | 7.8x | 52 MB (+34) | 44 MB (+25) |
| bench_formulas.xlsx | load + save | 13.31 s (6.95 + 6.36) | 1.44 s (0.53 + 0.91) | 9.2x | 467 MB (+448) | 62 MB (+43) |
| bench_formulas.xlsx | load + iterate values | 7.51 s (6.82 + 0.68) | 0.63 s (0.52 + 0.10) | 11.9x | 395 MB (+376) | 62 MB (+44) |
| bench_formulas.xlsx | load + iterate cells | 7.32 s (6.69 + 0.63) | 0.85 s (0.50 + 0.35) | 8.6x | 396 MB (+377) | 296 MB (+277) |
| bench_formulas.xlsx | read-only: iterate values | 6.49 s (2.12 + 4.38) | 0.81 s (0.36 + 0.45) | 8.0x | 61 MB (+42) | 43 MB (+25) |
| bench_formulas.xlsx | read-only: iterate cells | 6.77 s (2.09 + 4.69) | 0.92 s (0.33 + 0.59) | 7.4x | 61 MB (+42) | 44 MB (+25) |
| large_sample.xlsx | load + save | 228.03 s (134.77 + 93.26) | 18.20 s (6.84 + 11.35) | 12.5x | 4617 MB (+4599) | 291 MB (+273) |
| large_sample.xlsx | load + iterate values | 146.81 s (137.60 + 9.21) | 7.82 s (6.52 + 1.31) | 18.8x | 4594 MB (+4576) | 359 MB (+340) |
| large_sample.xlsx | load + iterate cells | 143.35 s (134.09 + 9.26) | 14.00 s (6.65 + 7.35) | 10.2x | 4497 MB (+4479) | 3498 MB (+3480) |
| large_sample.xlsx | read-only: iterate values | 101.81 s (4.91 + 96.90) | 7.76 s (0.39 + 7.37) | 13.1x | 149 MB (+131) | 119 MB (+100) |
| large_sample.xlsx | read-only: iterate cells | 111.04 s (4.95 + 106.10) | 9.76 s (0.36 + 9.39) | 11.4x | 149 MB (+130) | 119 MB (+100) |

## The same task done with openpyxl and with openrsxl.extended

`tools/bench_tasks.py`: openpyxl can only give the formula *and* the cached
value of every cell, or merged ranges and the comment / hyperlink of every
cell, in full mode (loading the workbook twice for formulas + values);
openrsxl runs the same code (drop-in), openrsxl.extended streams it in one
read-only pass.

| file | task | openpyxl (full mode) | openrsxl (full mode) | openrsxl.extended (streaming) |
|---|---|---|---|---|
| sample.xlsx | formulas_and_values | 13.01 s, 582 MB (+566) | 3.12 s, 439 MB (+422) | 1.13 s, 48 MB (+32) |
| sample.xlsx | merged_comments_links | 7.16 s, 335 MB (+318) | 1.58 s, 261 MB (+244) | 1.17 s, 49 MB (+32) |
| bench_formulas.xlsx | formulas_and_values | 14.88 s, 699 MB (+682) | 3.18 s, 506 MB (+489) | 1.48 s, 44 MB (+27) |
| bench_formulas.xlsx | merged_comments_links | 7.81 s, 396 MB (+379) | 1.48 s, 296 MB (+279) | 1.67 s, 44 MB (+27) |

On `large_sample.xlsx` openpyxl needs two full loads (2 x ~135 s, 4.5 GB
each) for formulas and cached values; openrsxl.extended streams them in
11.3 s with a 119 MB peak (below).

## openrsxl.extended streaming features

`tools/bench_extended.py <file> [flags] --cells`: read-only load, then every
cell is visited (reading `value`, and `formula` / `cached_value` when
enabled).

| file | mode | load | iterate cells | peak RSS (+increase) |
|---|---|---|---|---|
| sample.xlsx (616k cells) | openpyxl read-only | 0.39 s | 4.04 s | 51 MB (+32) |
|  | openrsxl.extended, no flag (= read-only mode) | 0.03 s | 0.56 s | 49 MB (+30) |
|  | openrsxl.extended, `formula_and_value` | 0.04 s | 0.64 s | 49 MB (+30) |
|  | openrsxl.extended, `read_merged_cells` | 0.11 s | 0.56 s | 49 MB (+30) |
|  | openrsxl.extended, all nine flags | 0.13 s | 0.61 s | 49 MB (+30) |
| bench_formulas.xlsx (800k cells) | openpyxl read-only | 2.03 s | 4.43 s | 61 MB (+42) |
|  | openrsxl.extended, no flag (= read-only mode) | 0.33 s | 0.60 s | 44 MB (+25) |
|  | openrsxl.extended, `formula_and_value` | 0.32 s | 0.65 s | 44 MB (+25) |
|  | openrsxl.extended, `read_merged_cells` | 0.45 s | 0.69 s | 44 MB (+25) |
|  | openrsxl.extended, all nine flags | 0.43 s | 0.67 s | 44 MB (+25) |
| large_sample.xlsx (9.9M cells) | openpyxl read-only | 4.96 s | 106.79 s | 149 MB (+130) |
|  | openrsxl.extended, no flag (= read-only mode) | 0.36 s | 9.15 s | 119 MB (+100) |
|  | openrsxl.extended, `formula_and_value` | 0.45 s | 10.89 s | 119 MB (+100) |
|  | openrsxl.extended, `read_merged_cells` | 1.51 s | 10.72 s | 120 MB (+101) |
|  | openrsxl.extended, all nine flags | 1.69 s | 11.88 s | 120 MB (+102) |

Memory does not grow with the number of cells: the peak of a 9.9 M cell
workbook is its shared strings table, as in plain read-only mode, and every
feature together adds ~2 MB to it.

## Creating workbooks

`tools/bench_write.py`: 200 000 rows x 10 columns (ints, floats, strings,
booleans, datetimes, a formula, None).

| mode | library | fill | save | peak RSS (+increase) |
|---|---|---|---|---|
| normal | openpyxl | 5.26 s | 13.52 s | 725 MB (+706) |
| write-only | openpyxl | 16.85 s | 0.92 s | 43 MB (+24) |
| normal | openrsxl | 4.91 s | 3.05 s | 589 MB (+570) |
| write-only | openrsxl | 6.51 s | 0.91 s | 43 MB (+24) |

Filling a normal worksheet runs openpyxl's own Python code (`Cell` objects
are created by `ws.append`), hence similar times.

## Where the time goes

* Loading uses a streaming XML tokenizer with a fast path for plain `<c>`
  elements; cells are stored compactly (10 bytes each) and Python objects
  are created on demand.
* The first iteration over *cell objects* of a loaded sheet creates them
  (openpyxl creates them in `load_workbook`), hence the higher memory of
  "load + iterate cells" compared with "load + iterate values".
* About half of the time of `save` on large files is zlib compression
  (`zipfile`), kept identical to openpyxl so that saved files are byte for
  byte the same.
* `formula_and_value` interprets each formula cell twice in the same pass;
  sheet-level features read the parts of each sheet outside `<sheetData>` at
  load time (the cell data is skipped with byte searches); `read_merged_cells`
  needs the styles of the corner cells of each range, which come before
  `<mergeCells>` in the file, so sheets with merged cells are skipped
  through once more on first iteration.
