"""
openrsxl is a drop-in replacement for openpyxl: the same API, the same
results, the same files - only the imports change.

    python examples/quickstart.py [output-directory]
"""

import datetime
import os
import sys
import tempfile

import openrsxl  # instead of: import openpyxl
from openrsxl.comments import Comment
from openrsxl.styles import Alignment, Border, Font, PatternFill, Side


def write_workbook(path):
    wb = openrsxl.Workbook()
    ws = wb.active
    ws.title = "Orders"

    # values of every type: the same Python types as openpyxl
    ws.append(["order", "customer", "quantity", "price", "total", "shipped", "date"])
    rows = [
        (1001, "Ada", 3, 9.99, True, datetime.datetime(2024, 1, 5, 14, 30)),
        (1002, "Grace", 1, 249.0, False, datetime.datetime(2024, 1, 6, 9, 0)),
        (1003, "Linus", 12, 0.5, True, datetime.datetime(2024, 1, 7, 18, 45)),
    ]
    for r, (order, customer, qty, price, shipped, date) in enumerate(rows, start=2):
        ws.append([order, customer, qty, price, f"=C{r}*D{r}", shipped, date])
    ws["G6"] = "=SUM(E2:E4)"

    # styles, number formats, merged cells, comments, hyperlinks
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=5):
        for cell in row:
            cell.number_format = "#,##0.00"
    ws["A6"] = "Grand total"
    ws.merge_cells("A6:F6")
    ws["A6"].border = Border(top=Side(style="thin"))
    ws["B2"].comment = Comment("Our first customer", "sales")
    ws["A8"] = "Documentation"
    ws["A8"].hyperlink = "https://openpyxl.readthedocs.io"
    ws.column_dimensions["G"].width = 20
    ws.freeze_panes = "A2"

    wb.save(path)  # byte for byte the file openpyxl would write


def read_workbook(path):
    wb = openrsxl.load_workbook(path)
    ws = wb["Orders"]
    print("dimensions:", ws.dimensions)
    print("A2:", repr(ws["A2"].value), "B2:", repr(ws["B2"].value), "E2:", repr(ws["E2"].value))
    print("G2:", repr(ws["G2"].value))  # datetime.datetime
    print("Z99:", repr(ws["Z99"].value))  # None, like openpyxl
    print("merged:", ws.merged_cells.ranges)
    print("comment:", ws["B2"].comment.text, "| link:", ws["A8"].hyperlink.target)
    for row in ws.iter_rows(min_row=1, max_row=4, max_col=4, values_only=True):
        print("  ", row)


def stream_workbook(path):
    # read-only mode streams rows with constant memory, as in openpyxl
    wb = openrsxl.load_workbook(path, read_only=True)
    total = 0
    for row in wb["Orders"].iter_rows(min_row=2, max_row=4, values_only=True):
        total += row[2] * row[3]
    print("streamed total:", round(total, 2))
    wb.close()


def write_only(path, rows=10_000):
    # write-only mode writes rows without keeping them in memory
    wb = openrsxl.Workbook(write_only=True)
    ws = wb.create_sheet("Log")
    for i in range(rows):
        ws.append([i, f"event {i}", i * 0.5])
    wb.save(path)
    print(f"wrote {rows} rows to", os.path.basename(path))


def main(out_dir):
    path = os.path.join(out_dir, "orders.xlsx")
    write_workbook(path)
    read_workbook(path)
    stream_workbook(path)
    write_only(os.path.join(out_dir, "log.xlsx"))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            main(tmp)
