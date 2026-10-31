# openrsxl-llm/1: lossless, LLM readable workbooks

`openrsxl.extended.llm` turns a workbook into a plain text document that a
language model reads directly - values, formulas, styles and every kind of
metadata in separate sections - and that decodes back to exactly what
openpyxl reads. This page is the specification of the format (written so
that an AI agent can read a document without any code), the reasons behind
it, and the measurements.

* [Using it](#using-it)
* [Reading a document](#reading-a-document) - the format, section by section
* [Decoding algorithm](#decoding-algorithm)
* [How the encoding is chosen](#how-the-encoding-is-chosen)
* [What "lossless" covers](#what-lossless-covers)
* [Measurements](#measurements)
* [Design notes and research](#design-notes-and-research)

## Using it

```python
from openrsxl.extended.llm import analyze_workbook, compress_workbook, decompress

doc = compress_workbook("book.xlsx")                     # everything, in memory
prompt = doc.text()                                      # the document
doc.text(sections=["values"], sheets=["Sales"])          # a part of it
doc["Sales"]["formulas"].text                            # one section
doc.to_json()                                            # the JSON form (opt-in, more tokens)

compress_workbook("book.xlsx", include="values")         # values only
compress_workbook("book.xlsx", include=["formulas", "styles"])
compress_workbook("book.xlsx", include="all", exclude="styles")
compress_workbook("book.xlsx", output="text")            # book.llm.txt, streamed
compress_workbook("book.xlsx", output="all", path="out/book")  # memory + out/book.llm.txt + .json

analyze_workbook("book.xlsx")      # {sheet: measurements and chosen encoding}
book = decompress(doc)             # or decompress("out/book.llm.json") / the text
book["Sales"].values[(2, 3)]       # (row, column) -> value as openpyxl reads it
book["Sales"].cell("D2")           # value, formula, style, hyperlink, comment
wb = book.to_workbook()            # an openpyxl Workbook again
```

| argument | |
|---|---|
| `include` / `exclude` | section names (below) or groups: `"all"` (default), `"cells"` = values + formulas + styles, `"metadata"` = the rest |
| `sheets` | the worksheet(s) to compress: a name or a list of names (default: all). The other sheets are not read. See [One sheet at a time](#one-sheet-at-a-time) |
| `encoding` | values encoding: `"auto"` (default: chosen per sheet), `"grid"`, `"index"`, `"hybrid"` |
| `output` | `"memory"` (default: the text form, `doc.text()`), `"text"` (`<path>.llm.txt`), `"json"` (`<path>.llm.json`, opt-in), `"all"` or a list - see [Which output form](#which-output-form-to-give-a-model) |
| `path` | base path of the files (default: the workbook's path without extension) |
| `legend` | put the legend (how to read the document) at the top (default True) |
| `jump_gap`, `runs`, `row_anchor`, `ditto`, `index_limit` | tuning: empty cells before a grid column jump (4), `v*N` repeats (on), grid rows always labelled (multiples of 10; 1 = every row), `^` ditto marks (on), memory of the analysis' inverted index (64 MiB) |
| `skip_hidden` | leave out the cells of hidden rows and columns (default False; see [Filters](#filters-opt-in)) |
| `skip_hidden_sheets` | leave out hidden and very hidden worksheets (default False) |
| `print_area_only` | on sheets with a print area, keep only its cells and the print titles (default False) |
| `rows`, `columns`, `ranges` | keep only these rows / columns / ranges, several segments each (default: all); see [Selecting rows, columns and ranges](#selecting-rows-columns-and-ranges) |

Sections: `values`, `formulas`, `styles`, `merged`, `hyperlinks`,
`comments`, `dimensions`, `validations`, `conditional_formatting`,
`tables`, `sheet`, `workbook`, `analysis`.

Every worksheet is **streamed** by the Rust engine: no cell objects are
created and memory does not grow with the size of a sheet. Without
`"memory"` in `output` the values are written to the files as they are
encoded (a 98 MB workbook with 10 million cells compresses in under 30 s
with a 212 MB peak - the formulas and style maps of all sheets
are kept until they are written).

## One sheet at a time

```python
from openrsxl.extended.llm import compress_sheets, compress_workbook

doc = compress_workbook("book.xlsx")                       # every worksheet
doc = compress_workbook("book.xlsx", sheets="Sales")       # one worksheet
doc = compress_workbook("book.xlsx", sheets=["Sales", "Costs"])  # some of them

for name, doc in compress_sheets("book.xlsx"):             # one document per sheet
    send_to_model(doc.text())
for name, doc in compress_sheets("book.xlsx", sheets=["Sales", "Costs"], output="text", path="out/book"):
    ...                                                    # out/book.Sales.llm.txt, out/book.Costs.llm.txt
```

* `sheets=` selects worksheets by name; only those are read (a 1-sheet
  document of a 50-sheet workbook costs the time of that sheet). The
  document still has the whole `## workbook` section (every sheet's name and
  state), and `## styles` only has the styles of the selected sheets.
* `compress_sheets` opens the workbook once and yields `(name,
  CompressedWorkbook)` per worksheet, in workbook order; each document is
  complete on its own and equal to `compress_workbook(..., sheets=[name])`.
  A document is finished before the next sheet is read, so memory holds one
  sheet's document at a time; stop iterating to skip the rest. With file
  outputs each sheet gets `<path>.<sheet name>.llm.txt` / `.json`
  (characters not allowed in file names become `_`).
* Every option works with both, filters included
  (`skip_hidden_sheets=True` drops a hidden sheet even when named).
* Sheet numbers (`## sheet N`) are the sheets' positions in the workbook, so
  documents made one sheet at a time agree with each other and with the
  document of all sheets.
* A name that is not a worksheet raises `KeyError` (chartsheets have no
  cells: they are listed in `## workbook` only).

## Which output form to give a model

**The text form is the cheapest, and it is the default.** `compress_workbook`
returns the document in memory and `doc.text()` - or the `.llm.txt` file of
`output="text"` - is exactly that text. The JSON form is only produced on
request (`output="json"` / `"all"`, `doc.to_json()`): it carries the same
section texts, so it never says more, and it always costs more tokens.

Measured with tiktoken `o200k_base` (`count_tokens`):

| workbook | text | JSON | JSON, indented |
|---|---:|---:|---:|
| `sample.xlsx` (6 MB, everything) | 1 087 256 | +1.5 % | +1.6 % |
| `sample.xlsx`, values only | 1 064 213 | +1.5 % | +1.6 % |
| 8 financial workbooks (`tools/make_financial_files.py`), everything | 149 k - 247 k each | +1.0 to +1.7 % | +1.3 to +2.2 % |
| the same, values only | 140 k - 230 k each | +0.9 to +1.6 % | +1.2 to +1.9 % |
| a small two-sheet workbook | 1 563 | **+47.5 %** | +62 % |
| all 20 documents | 4 958 516 | +1.4 % | +1.7 % |

Where the JSON overhead comes from: every line break of a section is an
escaped `\n`, every quote of a value an escaped `\"`, the keys and
punctuation of the envelope, and the analysis of each sheet repeated as a
JSON object. On large sheets this is 1-2 %; on small documents, where the
envelope dominates, it is half again as many tokens.

Use the JSON form only when a program, not a model, consumes the document
(it parses without a custom reader: `json.load` then
`decompress(...)`), or to store it next to other JSON. Both forms decode to
the same workbook. The defaults of the other choices are the cheapest too:
`encoding="auto"` picks the cheapest readable values encoding per sheet,
`legend=True` costs at most about 920 tokens and can be turned off
(`legend=False`) when the reader already knows the format, and `include=`
/ the filters cut what is not needed.

## Reading a document

A complete small document (the demo of `examples/llm_compression.py`,
legend left out):

```
#openrsxl-llm/1 source="demo.xlsx" sections=analysis,values,formulas,styles,comments,sheet,workbook
## workbook
sheet(name=Sales)
sheet(name=Notes)
workbookView
## styles
font f0: name:Calibri family:2 color(theme=1) sz:11 scheme:minor
font f1: b:1
fill p0: patternFill
fill p2: patternFill(patternType=solid fgColor(rgb=00DDEBF7))
border b0: left right top bottom diagonal
s0: font=f0 fill=p0 border=b0 numFmt=General alignment protection(locked=1 hidden=0) style=Normal
s1: font=f1 fill=p2
s2: numFmt="mmm yyyy"
## sheet 1 "Sales" dim=A1:E27
### analysis
cells=102 rows=26 values=A1:E27 density=0.7556 distinct_ratio=0.4412 kinds=number:48,text:30,date:24 formulas=25 formula_groups=2 styled_cells=29 est_tokens=grid:407,index:333,hybrid:289 encoding=hybrid
reason: dense sheet (density 0.76) with blocks of repeated values: hybrid ~71% of grid
### values hybrid start=A
A2:A13=North
D2:D25=2.5
A14:A25=South
1:Region|Month|Units|Price|Revenue
|2024-01-01|6
|2024-02-01|9
...
10:|2024-09-01|30
|2024-10-01|33
...
20:|2024-07-01|57
...
|2024-12-01|75
27:Total
### formulas
E2:E25=C2*D2
E27=SUM(E2:E25)
### styles
A1:E1=s1
B2:B25=s2
### comments
A1 author=analyst "North / South split"
### sheet
sheetViews(sheetView(pane(ySplit=1 topLeftCell=A2 activePane=bottomLeft state=frozen) selection(pane=bottomLeft activeCell=A1 sqref=A1)))
## sheet 2 "Notes" dim=B2:H30
### values grid start=B
2:Prices exclude VAT
9:Reviewed
30:H=Reviewed
```

Reading it: A2 is `North` (block `A2:A13`), B3 is the date 2024-02-01 (the
second line after `1:`, so row 3), C3 is
9, D3 is 2.5 (block), E3 holds `=C3*D3` (the fill `E2:E25=C2*D2` shifted by
one row) whose value is not stored in this file, A1 is bold on a light blue
fill and has a comment, B2:B25 are formatted `mmm yyyy`, the first row is
frozen; on Notes, H30 is `Reviewed`.

### Structure

* Line 1: `#openrsxl-llm/1`, the source file name and the included
  sections. A section that is included but absent from a sheet is empty
  (no formulas, no styles beyond `s0`...).
* `## legend`: how to read the document (only the parts it uses; every line
  starts with `- ` or two spaces).
* `## workbook`: the workbook section.
* `## sheet N "Name" dim=A1:F120`: a worksheet (N: its position in the
  workbook from 1, as listed in `## workbook` - the same in a document of one
  sheet and in one of all sheets; the name as a JSON
  string, the bounding box of all its cells). A filtered sheet adds
  `skip=hidden` and / or `area=RANGES[+titles:ROWS,COLS]` (see below).
* `### <section> [encoding] [start=COL]`: a section of the current sheet.

No line of a section starts with `#`.

### Filters (opt-in)

By default nothing is left out. Three flags trade completeness for tokens
when the reader only needs what a person sees or prints:

| flag | leaves out | header item |
|---|---|---|
| `skip_hidden=True` | cells of hidden rows and hidden columns (rows marked hidden, collapsed outline groups; with `sheetFormatPr zeroHeight`, every row without a visible row dimension) | `skip=hidden` |
| `skip_hidden_sheets=True` | hidden and very hidden worksheets, also when named in `sheets=` (the `## workbook` section still lists them with `state=`) | - |
| `print_area_only=True` | on sheets with a print area, every cell outside it, except the print titles (title rows across the area's columns, title columns across its rows); sheets without a print area are kept whole | `area=A1:F40,H1:J9+titles:1:2,A:A` |

* A left-out cell is dropped from `values`, `formulas`, `styles`,
  `hyperlinks` and `comments`. Range metadata (merged cells, dimensions,
  validations, conditional formats, tables, sheet settings) is kept whole,
  so a reader still sees which rows are hidden and what the print area is.
* The filter applies before the analysis: the encoding is chosen for the
  cells that remain, `dim=` is their bounding box, and
  `analysis["skipped_cells"]` counts the others.
* A decoder must not treat left-out cells as empty: `DecodedSheet.filter`
  (JSON: `"filter"`) lists the header items, and the legend explains them.
* When a flag finds nothing to leave out, the document is byte for byte the
  lossless one.

### Selecting rows, columns and ranges

```python
compress_workbook("book.xlsx", rows="1:100,200:300")           # two row segments
compress_workbook("book.xlsx", columns="A:D,F")                # columns A..D and F
compress_workbook("book.xlsx", rows="2:500", columns="B:H")    # rows x columns
compress_workbook("book.xlsx", ranges="A1:D20,F5:H9")          # start:end rectangles, all else ignored
compress_workbook("book.xlsx", ranges=["A1:D20", "F5:H9", "K2"])   # the same as a list (+ one cell)
compress_workbook("book.xlsx", ranges={"Sales": "A1:F40", "Costs": "B2:D9"})   # per sheet
```

| option | accepts | header item |
|---|---|---|
| `rows` | `5`, `"5"`, `"5:9"`, `"1:10,20:30"`, `(first, last)`, or a list of these | `rows=1:10,20:30` |
| `columns` | `"B"`, `"A:D,F"`, `2`, `("B", "D")` / `(2, 4)`, or a list of these | `cols=A:D,F:F` |
| `ranges` | `"A1:D20"`, `"A1:D20,F5:H9"`, `"C4"`, `"B:D"` (whole columns), `"3:7"` (whole rows), or a list of these | `cells=A1:D20,F5:H9` |

* Each entry of `ranges` is a rectangle given by its start and end
  (`"A1:D20"`); a list or a comma separated string gives several. The
  selected rectangles are encoded as blocks of rows like any sheet (a
  rectangle of 20 rows is 20 grid lines), never cell by cell; a single cell
  (`"K2"`) is simply a 1 x 1 rectangle. Overlapping rectangles are fine.
  Example - `ranges=["A1:C4", "F6:H9"]` on a sheet of numbers:

  ```
  ## sheet 1 "Sheet" dim=A1:H9 cells=A1:C4,F6:H9
  ### values grid start=A
  1:11|12|13
  21|22|23
  31|32|33
  41|42|43
  6:F=66|67|68
  F=76|77|78
  F=86|87|88
  F=96|97|98
  ```

* A cell is kept when it is in the rows, the columns and the ranges given
  (options not given do not restrict), and when the other filters keep it:
  with `print_area_only=True` the ranges are intersected with the print area.
* One value applies to every selected sheet; a dict `{sheet name: value}`
  gives each sheet its own, and sheets not in it are not restricted. A
  dict key that is not a worksheet raises `KeyError`.
* `$` signs are ignored; `"Sheet!A1"` is refused - use the dict form.
  Invalid input (row 0, `last < first`, a column past `XFD`, an empty
  selection) raises `ValueError` / `TypeError` before anything is read.
* Like the other filters, the selection applies before the analysis: the
  encoding is chosen for the selected cells, `dim=` is their bounding box,
  `analysis["skipped_cells"]` counts the others, and range metadata (merged
  cells, dimensions, validations...) is kept whole.

### Cell references and ranges

A1 references (columns `A`..`ZZZ`, rows from 1). `RANGES` is a
comma-separated list of:

| range | cells |
|---|---|
| `B2` | one cell |
| `B2:D9` | a rectangle |
| `B6:Q98/4` | columns B..Q of rows 6, 10, 14, ... 98 (every 4th row; `r1` is the last row) |

### Value tokens

| token | value |
|---|---|
| `12`, `-3`, `123456789012345678901234` | integer |
| `-3.5`, `1e-05`, `1.0`, `1e+16` | float (has `.` or an exponent) |
| `TRUE`, `FALSE` | boolean |
| `#N/A`, `#DIV/0!`, `#VALUE!` ... (starts with `#`) | error |
| `2024-01-05` | date (a datetime at midnight, as openpyxl reads a date cell) |
| `2024-01-05 13:45:00`, `2024-01-05 13:45:00.000500` | datetime |
| `13:45:00`, `13:45:00.000500` | time |
| `"..."` | text, JSON string syntax |
| `@date:2024-01-05` | a date without time (ISO 8601 cells) |
| `@td:36:15:00`, `@td:-0:00:01.500000` | duration (`[-]hours:MM:SS[.ffffff]`) |
| `@err:"..."` | an error whose text cannot be written bare |
| `@float:inf`, `@float:-inf`, `@float:nan` | non-finite floats |
| anything else | text, exactly as written |

Text is quoted only when the bare form would read as something else: it is
empty, has spaces at an end, contains `|` or a control character (line
breaks...), starts with `"` `#` `@` `=`, is `TRUE` / `FALSE`, looks like a
number (`"007"`, `"1e5"`), a date or time, ends with `*` + digits or starts
with `COL=`. Numbers are Python's shortest round-trip form; `1` (integer)
and `1.0` (float) are different values.

### `### values grid start=X`

One line per row that has values: `ROW:field|field|...`. The first field is
column X, each following field the next column.

The `ROW:` label is left out when the line is the row right after the
previous line: `5:a|b` then `c|d` = row 5 then row 6. The first line, every
line after a gap of empty rows, and every row that is a multiple of 10
keep the label. A row number is never more than 9 lines up, and the label
costs nothing on the consecutive rows of a table. (A line that would start
with digits and `:` without its label, such as a time `13:45:00`, keeps
it.) Fields:

| field | meaning |
|---|---|
| (empty) | an empty cell |
| `v` | token v |
| `v*N` | token v in N consecutive cells |
| `COL=v` / `COL=v*N` | jump to column COL (the cells in between are empty) |
| `^` / `^*N` | ditto: the same value as the cell above - the previous line's field in the same column (N cells: each takes the value above it) |

Rows that are not listed are empty; nothing follows the last value of a
row. `3:a||b*3|K=x` = A3 `a`, B3 empty, C3:E3 `b`, K3 `x` (with `start=A`).

Ditto marks are what make ledgers and journals cheap without moving their
data out of the rows: the date, journal number, account or currency that
repeats down a column is written once and `^` below it. They are used
only on unlabelled lines (the row right after the line above), never on a
labelled one: every anchor row spells its values out, so any value is at
most 9 lines up. `^` is used only when the value costs more than one
token. A text `^` is written `"^"`.

```
10:2024-01-11|JE4931|1148|Receivables - D41|CC118|CHF|Invoice 50158||21317.84
^|^|4027|Revenue - H26|^|^|Payment 11202|880.5|
^|JE4932|^|^|CC330|^|Accrual 77812|1520.0|
```

= row 11 has the date and journal of row 10, row 12 the date of row 11,
account `4027`, name `Revenue - H26` and currency `CHF`. Three or more
ditto marks side by side are `^*N`; a whole repeated row is `^*N` alone.

### `### values index`

The inverted index: one line per distinct value, `RANGES=v`, in order of
first appearance. `B2,B9:B12=Alice` = B2, B9, B10, B11, B12 hold `Alice`.
Cells not listed are empty. (Everything after the first `=` is the token.)

### `### values hybrid start=X`

First block lines `RANGES=v` (large areas of one repeated value), then grid
lines as in `grid`: the first grid line is labelled `ROW:`, and every line
from it on is a grid line. A cell covered by a block
is an empty field in the grid.

### `### formulas`

| line | meaning |
|---|---|
| `RANGES=FORMULA` | the first cell of RANGES holds `=FORMULA`; every other cell holds the same formula with its relative references shifted by the cell's offset from that first cell |
| `CELL~"..."` | the formula (whole text, JSON string) - when it has line breaks or does not start with `=` |
| `CELL[array REF]=FORMULA` | an array (CSE) formula over REF, stored in CELL |
| `CELL[dataTable ATTRS]` | a what-if data table (`ref=`, `dt2D=`, `r1=` ... as in OOXML) |

The shift is Excel's copy / fill: in `D2:D100=B2*C2`, D3 is `=B3*C3`; `$B$2`
never moves, `$B2` keeps its column, `B$2` its row; text in quotes and sheet
names do not change. Lines are in order of their first cell. Formulas
whose fill Excel would treat differently (whole-column ranges such as `A:A`
across columns, whole rows `1:1` across rows, structured / external
references `Table1[Col]` across columns) are never grouped in that
direction - every line reads correctly with Excel's rules. The values of
formula cells (the results last calculated by Excel) are in the values
section.

### `## styles` and `### styles`

Styles belong to the workbook, so the table is written once, as the
workbook-level `## styles` (before the sheets), with the styles the encoded
sheets use; each sheet's `### styles` only maps its cells:

```
## styles
font fN: <compact XML of <font>>        a font (the workbook's font N)
fill pN: <compact XML of <fill>>        a fill
border bN: <compact XML of <border>>    a border
sN: font=fN fill=pN border=bN numFmt=CODE alignment(...) protection(...) quotePrefix=1 pivotButton=1 style=NAME
### styles                              (in a sheet)
RANGES=sN                               the cells with style N
```

`s0` lists every key; another `sN` lists only the keys that differ from
`s0`. Cells not mapped have `s0`; a sheet whose cells all have `s0` has no
`### styles`. A part of a document (`doc.text(sheets=["Sales"])`) carries
the style table lines its sheets need. `numFmt` is the number format code
(`General`, `0.00%`, `"mmm yyyy"`), `style` the named style.

### Compact XML (styles and metadata)

The metadata sections use the OOXML vocabulary Excel itself uses, in one line
per item:

| notation | XML |
|---|---|
| `name(...)` | element |
| `k=v` | attribute `k="v"` |
| `k:v` | child element `<k val="v"/>` |
| `k` | empty child element `<k/>` |
| `"..."` | text content (JSON string) |
| `{r}id`, `{dc}creator` | namespaced names (`r` relationships, `dc` / `dcterms` / `cp` document properties, ...) |

Values with spaces, `(`, `)` or `"` are JSON strings. Attributes that equal
openpyxl's defaults are left out (each item is checked: reading it back
gives exactly the original object).

### Metadata sections

| section | lines |
|---|---|
| `merged` | merged ranges, comma-separated; the top-left cell holds the value and style of the range |
| `hyperlinks` | `RANGES target=URL location=... tooltip=... display=...` |
| `comments` | `RANGES author=... width=.. height=.. "text"` (width / height only when not 144 / 79) |
| `dimensions` | `col(min=1 max=3 width=12.5 hidden=1 s=sN ...)`; `row(r=5 ht=30.0)`, `row(r=8:20 ...)` for a span of rows with the same settings |
| `validations` | `dataValidation(...)` per rule (`sqref` = cells) |
| `conditional_formatting` | `RANGES cfRule(... dxf(...))` per rule, in order; `dxf` = the format applied |
| `tables` | `table(...)` per table |
| `sheet` | `sheetPr`, `sheetViews` (frozen panes, selection, zoom), `sheetFormatPr`, `printOptions`, `pageMargins`, `pageSetup`, `headerFooter`, `autoFilter`, `sheetProtection`, `rowBreaks`, `colBreaks`, `scenarios`, `printTitles(rows= cols=)`, `printArea:` - only the settings that differ from Excel's defaults |
| `workbook` | `sheet(name= state=hidden type=chartsheet)` in order, `workbook(date1904=1 codeName=)`, `active:N`, `workbookView`, `calcPr`, `workbookProtection`, `definedName(name= localSheetId= "formula")`, `{cp}coreProperties(...)`, custom properties |
| `analysis` | measurements and the reason of the values encoding (not cell data) |

## Decoding algorithm

The reference decoder is `openrsxl/extended/llm/_decode.py` (pure Python,
about 600 lines). In short:

```
for each line:
  "## sheet N \"name\" ..."  -> new sheet
  "### name [enc] [start=X]" -> new section
values grid / hybrid:
  the grid starts at the first line 'DIGITS:' (hybrid: block lines before);
  each grid line: row = its 'ROW:' label, or the previous line's row + 1;
  fields = split at '|' (a field starting
    with '"' or '@err:"' is a JSON string first, then maybe '*N');
    col = X; for each field: empty -> col += 1; else optional 'COL=' jump,
    token, optional '*N'; cells col..col+N-1 = token; col += N
  other lines (index, hybrid blocks): RANGES '=' token
formulas: RANGES ('=' body | '~' json | '[array REF]=' body | '[dataTable ...]');
  anchor = first cell; each cell = shift(anchor formula, cell - anchor)
## styles: part lines, sN lines (missing key = s0's); ### styles: RANGES=sN
```

`decompress()` returns cells keyed by `(row, column)` and openpyxl objects
(`Font`, `PatternFill`, `DataValidation`, `Rule`, `Table`, `Hyperlink`,
`Comment`, ...); `to_workbook()` rebuilds a Workbook from them.

## How the encoding is chosen

Each sheet is streamed once (pass 1) and measured: number of cells,
bounding box and **density** (cells / area), value types, **distinct
values** and how each value's cells cluster into rectangles and strided
rows (exact cover, built while streaming), formula templates, styles. From
these the token cost of each values encoding is estimated (a BPE-like
estimate within a few percent of tiktoken's o200k counts):

* **grid** - cost of the row lines: fields, row labels (only where
  needed), empty runs (`|||` is one token), column jumps, `v*N` repeats,
  `^` ditto marks (a value equal to the one above costs one token).
* **index** - cost of one line per distinct value with its ranges.
* **hybrid** - grid cost minus what the blocks save (rectangles of 3+ rows
  of one value, strided rows repeated 3+ times, kept when cheaper as a
  range line than in the grid - where the rows below a block's first are
  ditto marks, so short vertical repeats stay in the rows and only long
  constant runs such as a currency column become blocks).

The cheapest wins, with a bias for the grid (rows read as rows are the
easiest for a model): index below 0.85 x grid, hybrid below 0.92 x grid.
In practice:

| sheet | typically |
|---|---|
| ledgers, journals, registers, payroll, FX tables, P&L / trial balance layouts | grid (repeats down a column as `^`) |
| tables grouped by long constant categories (budget by department), constant columns | hybrid |
| sparse sheets, forms, matrices of flags or zeros, periodic returns (VAT, monthly blocks) | index |

The rules are general - they measure the sheet, not its kind - and are
checked against a population of financial workbooks rather than one file:
`tools/make_financial_files.py` draws ledgers, trial balances, monthly
P&Ls, AR ageing, budget vs actual, cash flow statements, FX rates, payroll
and fixed asset registers from seeds. On 8 such workbooks (72 sheets) the
automatic choice costs 2.3 % more tokens than the cheapest encoding chosen
in hindsight, all of it the readability bias (ledgers kept as row-local
grids where `hybrid` would save 7 %; `encoding="hybrid"` takes it).

The analysis keeps the inverted index in memory only up to `index_limit`
(64 MiB); a sheet with more distinct data than that cannot win as an index
and is encoded as a grid. Pass 2 (grid and hybrid) streams the sheet again
and writes the rows as they come.

## What "lossless" covers

Decoding gives back, for every included section:

* values: value and type exactly as openpyxl reads them with
  `data_only=True` (formula cells: the cached result), at each cell's own
  coordinate - also out-of-order rows (as full mode, which read-only mode
  would skip), values inside merged ranges and empty-cell hyperlinks are
  the file's (openpyxl's full mode would replace them);
* formulas: as openpyxl reads them with `data_only=False` (text,
  `ArrayFormula`, `DataTableFormula`), shared formulas expanded;
* styles: font, fill, border, number format, alignment, protection,
  quotePrefix, pivotButton and named style of every cell in the file
  (styled empty cells included);
* metadata: objects equal (same XML) to those of openpyxl's full mode.

Not encoded: charts, images and drawings, pivot tables, VBA, external link
caches, the theme, rich text runs inside cells (their text is kept),
relationship ids and other package internals. The test-suite
(`tests/test_llm.py`, oracle `tests/oracle/llm.py`) compares every
decoded value, formula, cell style and object with the real openpyxl, for
every encoding and output form, on synthetic workbooks, raw XML edge cases,
openpyxl's own test files and large real workbooks.

## Measurements

Exact token counts (tiktoken `o200k_base`), `python tools/llm_bench.py
book.xlsx --types`. For each selection the **baseline** is the same
information loaded without compression, losslessly: every cell as JSON
`{"A1": {"v": value, "t": type}}` (values), `{"D2": "=B2*C2"}`
(formulas), `styles.xml` plus `{"A1": style id}` (styles), the XML of the
metadata parts (metadata), and all of it at once (everything). **Ours** is
the complete document a model receives for that selection - legend,
headers and analysis included. **Saved** = 1 - ours / baseline.

| loaded | `sample.xlsx` baseline | ours | saved | `large_sample.xlsx` baseline | ours | saved | 8 financial workbooks baseline | ours | saved |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| everything | 39 755 578 | 1 087 256 | 97.3 % | 422 325 682 | 30 075 118 | 92.9 % | 8 895 464 | 1 449 990 | 83.7 % |
| values | 13 368 902 | 1 064 213 | 92.0 % | 189 680 365 | 29 456 507 | 84.5 % | 6 562 691 | 1 354 686 | 79.4 % |
| formulas | 24 509 798 | 17 983 | 99.9 % | 189 252 923 | 47 430 | 99.97 % | 1 309 959 | 7 386 | 99.4 % |
| styles | 3 550 441 | 2 171 | 99.9 % | 70 185 221 | 555 671 | 99.2 % | 1 298 496 | 81 022 | 93.8 % |
| metadata | 7 059 | 2 832 | 59.9 % | 24 304 | 14 574 | 40.0 % | 18 536 | 6 400 | 65.5 % |

`sample.xlsx`: 6 MB, 9 sheets, 616 000 cells (424 000 formulas).
`large_sample.xlsx`: 98 MB, 25 sheets, 9.7 million values, 5.7 million
formulas. Financial workbooks: `tools/make_financial_files.py`, seeds 1-8.
Metadata saves least: it is small, and the legend lines that explain it
are a large part of a small total.

Per value type (a copy of the workbook holding only the cells of one type,
in their places; same baseline):

| value type | `sample.xlsx` cells | baseline | ours | saved | 2 000 x 10 dense (random) saved | 10 % sparse saved |
|---|---:|---:|---:|---:|---:|---:|
| integer | 6 231 | 106 031 | 19 947 | 81.2 % | 82.7 % | 73.0 % |
| decimal | 400 454 | 8 908 871 | 462 684 | 94.8 % | 74.1 % | 64.3 % |
| text | 192 573 | 3 702 910 | 570 092 | 84.6 % | 82.9 % | 84.7 % |
| date / datetime | 17 008 | 475 262 | 20 782 | 95.6 % | 89.8 % / 52.6 % | 81.0 % / 46.3 % |
| boolean | - | - | - | - | 89.4 % | 75.7 % |
| error | - | - | - | - | 88.2 % | 85.0 % |

The random columns are the hard case (no value repeats, nothing to share):
what is saved there is the JSON structure; distinct date-times save least
because a full timestamp is most of its own cost. Real sheets repeat, so
they save more (`sample.xlsx`'s decimals: 94.8 %).

Lossy formats for reference (values of `sample.xlsx`): CSV 4 107 699
tokens, a Markdown table 4 637 552, SpreadsheetLLM's vanilla `A1,v|`
6 540 899 - none of them keeps types, formulas or styles.

Memory: `large_sample.xlsx` compresses in 28 s with a peak of 212 MB
written to files (236 MB kept in memory; 93 MB of it is the interpreter
and the workbook's shared strings). Formulas, styles or metadata alone
stay at 93-120 MB.

The legend costs at most about 920 tokens (only the parts a document uses);
`legend=False` leaves it out when the reader already knows the format.

## Design notes and research

* **SpreadsheetLLM** (Tian et al., Microsoft, EMNLP 2024,
  [arXiv 2407.09025](https://arxiv.org/abs/2407.09025)) encodes sheets for
  LLMs with three modules: structural anchors (drops rows and columns far
  from table boundaries - lossy), the **inverted index** (value -> address
  ranges - lossless) and data-format aggregation (replaces numbers by their
  format type - lossy). Only the inverted index is lossless; it is the
  `index` encoding here (with strided rows added), used where it wins:
  sparse and repetitive sheets.
* Table formats for LLMs: CSV-like rows are the most token-efficient,
  JSON / HTML several times more expensive (TQA-Bench, arXiv 2411.19504);
  per-cell address listings - the SpreadsheetLLM baseline - are the most
  expensive. The `grid` is a CSV-like row form with column jumps instead of
  long runs of empty fields and `v*N` repeats. Row numbers: models cannot
  count lines reliably, so rows are labelled - but a label on every line
  costs about 2 tokens per row, which on narrow or sparse tables is as much
  as the data (it made sparse number columns 25-37 % more expensive than
  CSV). Labels are therefore kept where the row cannot be inferred (after a
  gap of empty rows) and at anchors every 10th row, so a row number is
  never more than 9 lines up: 7-15 % fewer tokens on such tables, 2-7 % on
  dense ones.
* Tokenizer facts used (o200k / cl100k): runs of `|` merge (`|||` is one
  token), so empty cells are almost free; JSON quotes merge with the
  delimiters (`"|`), so quoting costs about one token; ISO dates cost 6
  tokens (the clearest form, kept); short numbers 1-3 tokens.
* Formulas: a fill (the same formula copied over a range) is one line,
  checked byte for byte against every cell (the formula templates of the
  Rust engine); fills that repeat every N rows become strided ranges.
* Lossless by construction: every grouping is verified on the actual data
  while encoding (formula templates, rectangle covers, metadata items read
  back through openpyxl), never assumed.
