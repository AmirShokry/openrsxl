"""
The legend: how to read an openrsxl-llm document, written for the model
that receives it (every line starts with "- " or two spaces). ``legend(parts)`` assembles only the parts a document
uses; ``LEGEND`` is the complete text (also in docs/LLM_ENCODING.md).
"""

FORMAT = "openrsxl-llm/1"

_PARTS = {
    "intro": """{format}: lossless text form of an Excel workbook (nothing rounded or omitted). '## ' / '### ' lines open sections.
Cells are A1 refs; RANGES = comma separated cells, rectangles (B2,D2:F9) or strided rows (A6:Q98/4 = A:Q of rows 6,10,...,98).""",
    "filter": """Filtered sheets: a sheet header with 'skip=hidden' leaves out the cells (values, formulas, styles, hyperlinks, comments) of hidden
rows and columns; 'area=RANGES' (+titles:ROWS,COLS = print titles) keeps only the cells of the print area; 'rows=1:9,20:30' / 'cols=A:D'
/ 'cells=A1:D20,F5:H9' keep only the cells of those rows / columns / ranges (all items of a header apply). Other cells of such a sheet are
not empty, only not shown. Sheets missing from the document but listed in ## workbook were left out.""",
    "tokens": """Values: 12 int; -3.5, 1e-05 float; TRUE/FALSE; #N/A (starts with #) error; 2024-01-05 date; 2024-01-05 13:45:00[.ffffff] datetime;
13:45:00 time; "..." JSON string = text that would otherwise read as another type; @date:Y-M-D, @td:[-]H:MM:SS duration, @err:"...",
@float:inf; any other token is text as written.""",
    "grid": """### values grid start=X: lines 'ROW:f1|f2|...', f1 in column X, then one column per field; empty field = empty cell; 'COL=v' jumps to
column COL; 'v*N' = v in N cells; '^' = the same value as the cell above (previous line, same column). A line without 'ROW:' is the row after the previous line (labels are kept after gaps and every
10th row by default); unlisted rows are empty.""",
    "index": """### values index: 'RANGES=v' = v in every cell of RANGES; unlisted cells are empty.""",
    "hybrid": """### values hybrid start=X: 'RANGES=v' block lines, then grid lines (from the first 'ROW:' line on); block cells are empty fields in the grid.""",
    "formulas": """### formulas: 'RANGES=FORMULA': the first cell has =FORMULA, each other cell the same formula copied to it (relative refs shifted by
its offset from the first cell, as Excel copies; $ parts fixed); 'CELL~"..."' JSON formula; 'CELL[array REF]=...' array formula;
'CELL[dataTable ...]' data table. Formula results are in the values.""",
    "styles": """## styles (the workbook's, once): 'font fN:' / 'fill pN:' / 'border bN:' parts (compact XML); 'sN:' style N = font fill border
numFmt alignment protection quotePrefix pivotButton style(named), missing keys as in s0. ### styles of a sheet: 'RANGES=sN' cells
with style N; all other cells s0.""",
    "cx": """\
Compact XML: name(...) element; k=v attribute; k:v child <k val="v"/>; bare k empty child <k/>; "..." text (JSON); values with spaces, ( ) or " are JSON strings.""",
    "merged": """\
### merged: merged ranges; the top-left cell holds the value and style of the whole range.""",
    "hyperlinks": """\
### hyperlinks: 'RANGES target=URL location=... tooltip=... display=...' - each cell of RANGES has that hyperlink.""",
    "comments": """\
### comments: 'RANGES author=... width=.. height=.. "text"' - the comment (note) of each cell of RANGES.""",
    "dimensions": """\
### dimensions: 'col(min=1 max=3 width=..)' column widths / hidden / outline / style s=sN; 'row(r=5:9 ht=..)' row heights etc. (r: one row or a span).""",
    "validations": """\
### validations: one dataValidation(...) per line: sqref = cells, formula1 / formula2 = the rule's values (OOXML attributes).""",
    "conditional_formatting": """\
### conditional_formatting: 'RANGES cfRule(...)' one rule per line in priority order; dxf(...) is the format applied when the rule matches.""",
    "tables": """\
### tables: one table(...) per line (OOXML: ref = range, tableColumns, autoFilter, tableStyleInfo).""",
    "sheet": """\
### sheet: sheet level settings in compact XML (sheetPr, sheetViews with frozen panes, sheetFormatPr, page setup, print titles/area, protection ...);
settings not listed have Excel's defaults.""",
    "workbook": """\
## workbook: sheet(name=.. state=..) in workbook order, definedName(name=.. "formula"), workbookView, calcPr, coreProperties ... in compact XML.""",
    "analysis": """\
### analysis: how the sheet was measured and why its values encoding was chosen (informational, not cell data).""",
}

ORDER = (
    "intro",
    "filter",
    "tokens",
    "grid",
    "index",
    "hybrid",
    "formulas",
    "styles",
    "cx",
    "merged",
    "hyperlinks",
    "comments",
    "dimensions",
    "validations",
    "conditional_formatting",
    "tables",
    "sheet",
    "workbook",
    "analysis",
)

# parts that need the compact XML explanation
_CX_USERS = {"styles", "dimensions", "validations", "conditional_formatting", "tables", "sheet", "workbook"}


def legend(parts=None):
    """The legend for the given parts (all if None)"""
    if parts is None:
        wanted = set(ORDER)
    else:
        wanted = {"intro"} | set(parts)
        if wanted & {"grid", "index", "hybrid"}:
            wanted.add("tokens")
        if wanted & _CX_USERS:
            wanted.add("cx")
    # no line starts with '#': documents are split at '## ' / '### ' lines
    out = []
    for p in ORDER:
        if p in wanted:
            lines = _PARTS[p].format(format=FORMAT).split("\n")
            out.append("- " + lines[0])
            out.extend("  " + line for line in lines[1:])
    return "\n".join(out)


LEGEND = legend()
