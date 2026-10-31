"""DecodedWorkbook.to_workbook: an openpyxl Workbook from a decoded document"""

from copy import copy

from openrsxl.cell.cell import Cell, MergedCell
from openrsxl.utils.datetime import CALENDAR_MAC_1904
from openrsxl.utils.indexed_list import IndexedList
from openrsxl.worksheet.dimensions import ColumnDimension, RowDimension


def _set_defaults(wb, s0):
    """Make the workbook's default style parts those of the document's s0
    (cells without a style read as s0)"""
    if s0 is None:
        return
    for attr, value in (
        ("_fonts", s0.font),
        ("_fills", s0.fill),
        ("_borders", s0.border),
        ("_alignments", s0.alignment),
        ("_protections", s0.protection),
    ):
        if value is None:
            continue
        items = list(getattr(wb, attr))
        items[0] = copy(value)
        setattr(wb, attr, IndexedList(items))


def _apply_style(cell, sd):
    if sd.font is not None:
        cell.font = copy(sd.font)
    if sd.fill is not None:
        cell.fill = copy(sd.fill)
    if sd.border is not None:
        cell.border = copy(sd.border)
    if sd.alignment is not None:
        cell.alignment = copy(sd.alignment)
    if sd.protection is not None:
        cell.protection = copy(sd.protection)
    if sd.number_format is not None:
        cell.number_format = sd.number_format
    if sd.style is not None and sd.style in cell.parent.parent.named_styles:
        cell.style = sd.style
        # the named style resets the parts: apply them again
        for a in ("font", "fill", "border", "alignment", "protection"):
            v = getattr(sd, a)
            if v is not None:
                setattr(cell, a, copy(v))
        if sd.number_format is not None:
            cell.number_format = sd.number_format
    cell._style.quotePrefix = int(bool(sd.quotePrefix))
    cell._style.pivotButton = int(bool(sd.pivotButton))


def _fill_sheet(wb, ws, s):
    arrays = {}  # style id -> StyleArray of this workbook

    def style_array(sid):
        a = arrays.get(sid)
        if a is None:
            sd = s.styles.get(sid)
            probe = Cell(ws, row=1, column=1)  # detached
            if sd is not None:
                _apply_style(probe, sd)
            a = arrays[sid] = copy(probe._style)
        return a

    if s.styles:
        _set_defaults(wb, s.styles.get(0))
    keys = set(s.values) | set(s.formulas) | set(s.style_ids)
    for key in sorted(keys):
        r, c = key
        cell = ws.cell(row=r, column=c)
        if key in s.formulas:
            cell._value = s.formulas[key]
            cell.data_type = "f"
        elif key in s.values:
            cell._value = s.values[key]
            cell.data_type = s.data_types.get(key, "n")
        sid = s.style_ids.get(key)
        if sid:
            cell._style = copy(style_array(sid))
    for rng in s.merged:
        ws.merge_cells(rng)
    for (r, c), link in s.hyperlinks.items():
        cell = ws.cell(row=r, column=c)
        if not isinstance(cell, MergedCell):
            cell.hyperlink = copy(link)
    for (r, c), comment in s.comments.items():
        cell = ws.cell(row=r, column=c)
        if not isinstance(cell, MergedCell):
            cell.comment = copy(comment)
    for key, d in s.column_dimensions.items():
        d = dict(d)
        sid = d.pop("style_id", None)
        cd = ColumnDimension(ws, index=key, **d)
        if sid:
            cd._style = copy(style_array(sid))
        ws.column_dimensions[key] = cd
    for idx, d in s.row_dimensions.items():
        d = dict(d)
        sid = d.pop("style_id", None)
        rd = RowDimension(ws, index=idx, **d)
        if sid:
            rd._style = copy(style_array(sid))
        ws.row_dimensions[idx] = rd
    for dv in s.validations:
        ws.data_validations.dataValidation.append(copy(dv))
    for sqref, rule in s.conditional_formatting:
        ws.conditional_formatting.add(sqref, rule)
    for table in s.tables:
        ws.add_table(table)
    for attr, obj in s.sheet.items():
        setattr(ws, attr, obj)
    if s.print_title_rows:
        ws.print_title_rows = s.print_title_rows
    if s.print_title_cols:
        ws.print_title_cols = s.print_title_cols
    if s.print_area:
        ws.print_area = s.print_area


def to_workbook(book):
    from openrsxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    if book.date1904:
        wb.epoch = CALENDAR_MAC_1904
    if book.code_name:
        wb.code_name = book.code_name
    by_name = {s.name: s for s in book.sheets}
    entries = book.sheet_entries or [{"name": s.name} for s in book.sheets]
    for e in entries:
        if e.get("type") == "chartsheet":
            cs = wb.create_chartsheet(e["name"])
            cs.sheet_state = e.get("state", "visible")
            continue
        ws = wb.create_sheet(e["name"])
        ws.sheet_state = e.get("state", "visible")
        s = by_name.get(e["name"])
        if s is not None:
            _fill_sheet(wb, ws, s)
    for dn in book.defined_names:
        if dn.localSheetId is not None:
            idx = int(dn.localSheetId)
            if idx < len(wb._sheets):
                wb._sheets[idx].defined_names[dn.name] = copy(dn)
        else:
            wb.defined_names[dn.name] = copy(dn)
    if book.views:
        wb.views = [copy(v) for v in book.views]
    if book.calculation is not None:
        wb.calculation = copy(book.calculation)
    if book.security is not None:
        wb.security = copy(book.security)
    if book.properties is not None:
        wb.properties = copy(book.properties)
    if book.custom_doc_props is not None:
        wb.custom_doc_props = copy(book.custom_doc_props)
    if wb._sheets:
        wb.active = min(book.active, len(wb._sheets) - 1)
    return wb
