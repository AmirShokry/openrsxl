"""
openrsxl.extended.llm: lossless, LLM readable compression of a workbook.

    python examples/llm_compression.py [book.xlsx]

Without an argument a demo workbook is created first. The script prints the
analysis of every sheet, the compressed document, its parts (values only,
formulas only...), writes the .llm.txt / .llm.json files and decodes the
document again to show that nothing was lost.
"""

import datetime
import os
import sys
import tempfile

import openrsxl
from openrsxl.extended.llm import analyze_workbook, compress_workbook, decompress


def make_demo(path):
    from openrsxl.comments import Comment
    from openrsxl.styles import Font, PatternFill

    wb = openrsxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Region", "Month", "Units", "Price", "Revenue"])
    for r in range(2, 26):
        region = "North" if r < 14 else "South"
        ws.append([region, datetime.datetime(2024, (r - 2) % 12 + 1, 1), r * 3, 2.5, f"=C{r}*D{r}"])
    ws["A27"] = "Total"
    ws["E27"] = "=SUM(E2:E25)"
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for r in range(2, 26):
        ws.cell(r, 2).number_format = "mmm yyyy"
    ws["A1"].comment = Comment("North / South split", "analyst")
    ws.freeze_panes = "A2"
    for r in range(6, 10):
        ws.row_dimensions[r].hidden = True
    ws.print_area = "A1:E13"
    notes = wb.create_sheet("Notes")
    notes["B2"] = "Prices exclude VAT"
    notes["B9"] = "Reviewed"
    notes["H30"] = "Reviewed"
    wb.save(path)


def main():
    tmp = tempfile.mkdtemp()
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = os.path.join(tmp, "demo.xlsx")
        make_demo(path)

    # 1. the analysis: how each sheet is measured and encoded
    for name, a in analyze_workbook(path).items():
        print(f"{name}: {a['cells']} cells, density {a['density']}, -> {a['encoding']} ({a['reason']})")

    # 2. the whole document, in memory
    doc = compress_workbook(path)
    print(f"\n--- document (~{doc.tokens} tokens) ---")
    print(doc.text())

    # 3. parts of it: one section, one sheet, no legend
    print("--- values only ---")
    print(compress_workbook(path, include="values", legend=False).text())
    print("--- formulas of 'Sales' ---")
    print(doc["Sales"]["formulas"].text)

    # 3b. only what a reader sees (opt-in, leaves cells out): hidden rows and
    # columns, hidden sheets, cells outside the print area
    seen = compress_workbook(
        path, include="values", legend=False, skip_hidden=True, skip_hidden_sheets=True, print_area_only=True
    )
    print("--- values a reader sees (Sales rows 6-9 hidden, print area A1:E13) ---")
    print(seen.text())

    # 4. files (streamed: the values are written as they are encoded)
    files = compress_workbook(path, output=["text", "json"], path=os.path.join(tmp, "demo")).files
    print("--- files ---")
    for kind, p in files.items():
        print(kind, os.path.basename(p), os.path.getsize(p), "bytes")

    # 5. back again: the same values, formulas and styles openpyxl reads
    book = decompress(files["json"])
    sales = book["Sales"]
    print("\n--- decoded ---")
    print("A2:", sales.cell("A2")["value"], "| E3:", sales.cell("E3")["formula"], "| B2:", sales.cell("B2")["value"])
    print("B2 number format:", sales.cell("B2")["style"].number_format)
    wb = book.to_workbook()  # an openpyxl Workbook again
    print("rebuilt sheets:", wb.sheetnames)


if __name__ == "__main__":
    main()
