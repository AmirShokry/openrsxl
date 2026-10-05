"""
openrsxl.extended: openpyxl's read-only (streaming) mode, plus the things
openpyxl only offers in its full mode - each one behind its own flag.

    python examples/extended_streaming.py [book.xlsx]

Without an argument a demo workbook is created first. Every flag is
demonstrated in its own section below; all of them can be combined.
"""

import os
import re
import sys
import tempfile
import zipfile

import openrsxl
from openrsxl.extended import load_workbook


def make_demo(path):
    """A small workbook using every feature (written with openrsxl)."""
    from openrsxl.comments import Comment
    from openrsxl.formatting.rule import CellIsRule
    from openrsxl.styles import Font
    from openrsxl.worksheet.datavalidation import DataValidation
    from openrsxl.worksheet.table import Table

    wb = openrsxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["region", "units", "price", "revenue"])
    for r, (region, units, price) in enumerate([("North", 10, 2.5), ("South", 4, 10.0), ("East", 7, 3.0)], start=2):
        ws.append([region, units, price, f"=B{r}*C{r}"])
    ws["A6"] = "Total"
    ws["D6"] = "=SUM(D2:D4)"
    ws.merge_cells("A6:C6")
    ws["A2"].comment = Comment("Best region this year", "analyst")
    ws["F1"].hyperlink = "https://example.com/report"  # an empty cell with a link
    ws["A3"].hyperlink = "#Sales!D6"
    ws.row_dimensions[1].height = 24
    ws.column_dimensions["A"].width = 18
    ws.freeze_panes = "B2"
    ws.print_title_rows = "1:1"
    dv = DataValidation(type="list", formula1='"North,South,East,West"')
    dv.add("A2:A5")
    ws.add_data_validation(dv)
    ws.conditional_formatting.add("B2:B5", CellIsRule(operator="greaterThan", formula=["5"], font=Font(bold=True)))
    ws.add_table(Table(displayName="SalesTable", ref="A1:D4"))
    wb.save(path)
    # openpyxl / openrsxl never calculate formulas, so the file has no saved
    # results yet; files saved by Excel or LibreOffice have them. Simulate
    # that here by adding the results Excel would store:
    _add_cached_values(path, "xl/worksheets/sheet1.xml", {"D2": 25, "D3": 40, "D4": 21, "D6": 86})


def _add_cached_values(path, part, values):
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    xml = parts[part].decode("utf-8")
    for coord, value in values.items():
        xml = re.sub(rf'(<c r="{coord}"[^>]*>\s*<f>[^<]*</f>)\s*(<v\s*/>|<v></v>)?', rf"\1<v>{value}</v>", xml)
    parts[part] = xml.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def section(title):
    print(f"\n--- {title} ---")


def main(path):
    section("no flag: exactly openpyxl's read-only mode")
    wb = load_workbook(path, read_only=True)
    ws = wb.worksheets[0]
    print([c.value for c in next(ws.iter_rows())])
    wb.close()

    section("formula_and_value: the formula and the saved result, in one pass")
    wb = load_workbook(path, read_only=True, formula_and_value=True)
    for r, (cell,) in enumerate(wb.worksheets[0].iter_rows(min_row=2, max_row=6, min_col=4, max_col=4), start=2):
        # cell.value follows data_only (here: the formula, as openpyxl);
        # cell.formula: the formula or None; cell.cached_value: the result.
        # (row 5 is empty: an EmptyCell, whose attributes are all None)
        print(f"D{r}", repr(cell.value), repr(cell.formula), repr(cell.cached_value))
    wb.close()
    # with data_only=True, cell.value is the saved result, as in openpyxl
    wb = load_workbook(path, read_only=True, data_only=True, formula_and_value=True)
    cell = next(wb.worksheets[0].iter_rows(min_row=6, min_col=4, max_col=4))[0]
    print("data_only=True:", repr(cell.value), repr(cell.formula), repr(cell.cached_value))
    wb.close()

    section("read_comments: cell.comment")
    wb = load_workbook(path, read_only=True, read_comments=True)
    for row in wb.worksheets[0].iter_rows():
        for cell in row:
            if getattr(cell, "comment", None) is not None:
                print(cell.coordinate, repr(cell.comment.text), "by", cell.comment.author)
    wb.close()

    section("read_hyperlinks: cell.hyperlink")
    wb = load_workbook(path, read_only=True, read_hyperlinks=True)
    for row in wb.worksheets[0].iter_rows():
        for cell in row:
            if getattr(cell, "hyperlink", None) is not None:
                link = cell.hyperlink
                # as in full mode, an empty linked cell gets the target as value
                print(cell.coordinate, repr(cell.value), "->", link.target or link.location)
    wb.close()

    section("read_merged_cells: ws.merged_cells, MergedCell objects, range.start_cell")
    wb = load_workbook(path, read_only=True, read_merged_cells=True)
    ws = wb.worksheets[0]
    print("ranges:", [str(r) for r in ws.merged_cells])
    for r in ws.merged_cells:
        # the top-left cell as full mode has it (value, borders of the range)
        print(r.coord, "start cell:", repr(r.start_cell.value))
    print("A6:C6 ->", [type(c).__name__ for c in next(ws.iter_rows(min_row=6, max_row=6, max_col=3))])
    wb.close()

    section("read_dimensions: ws.row_dimensions / ws.column_dimensions")
    wb = load_workbook(path, read_only=True, read_dimensions=True)
    ws = wb.worksheets[0]
    print("row 1 height:", ws.row_dimensions[1].height, "| column A width:", ws.column_dimensions["A"].width)
    wb.close()

    section("read_sheet_properties: views, freeze panes, print settings, ...")
    wb = load_workbook(path, read_only=True, read_sheet_properties=True)
    ws = wb.worksheets[0]
    print(
        "freeze_panes:",
        ws.freeze_panes,
        "| print_title_rows:",
        ws.print_title_rows,
        "| orientation:",
        ws.page_setup.orientation,
    )
    wb.close()

    section("read_data_validations / read_conditional_formatting / read_tables")
    wb = load_workbook(
        path, read_only=True, read_data_validations=True, read_conditional_formatting=True, read_tables=True
    )
    ws = wb.worksheets[0]
    for dv in ws.data_validations.dataValidation:
        print("validation:", dv.type, dv.formula1, "on", dv.sqref)
    for cf in ws.conditional_formatting:
        print("conditional formatting:", cf.sqref, [rule.type for rule in cf.rules])
    for table in ws.tables.values():  # (as in openpyxl, .items() gives (name, ref) pairs)
        print("table:", table.displayName, table.ref)
    wb.close()

    section("all flags at once, values_only")
    flags = dict(
        formula_and_value=True,
        read_comments=True,
        read_hyperlinks=True,
        read_merged_cells=True,
        read_dimensions=True,
        read_sheet_properties=True,
        read_data_validations=True,
        read_conditional_formatting=True,
        read_tables=True,
    )
    wb = load_workbook(path, read_only=True, **flags)
    for row in wb.worksheets[0].iter_rows(values_only=True):
        print(row)  # merged cells are None, link targets fill empty cells (full mode)
    wb.close()

    section("disabled features do not exist (as in openpyxl's read-only mode)")
    wb = load_workbook(path, read_only=True, read_comments=True)
    cell = next(wb.worksheets[0].iter_rows())[0]
    try:
        print(cell.formula)
    except AttributeError as err:
        print("AttributeError:", err)
    wb.close()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            demo = os.path.join(tmp, "demo.xlsx")
            make_demo(demo)
            main(demo)
