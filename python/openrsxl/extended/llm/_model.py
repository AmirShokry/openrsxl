"""openpyxl objects of the metadata sections, shared by encoder and decoder"""

from openrsxl.utils.cell import get_column_letter
from openrsxl.worksheet.dimensions import ColumnDimension, RowDimension, SheetFormatProperties
from openrsxl.worksheet.filters import AutoFilter
from openrsxl.worksheet.header_footer import HeaderFooter
from openrsxl.worksheet.page import PageMargins, PrintOptions, PrintPageSetup
from openrsxl.worksheet.pagebreak import ColBreak, RowBreak
from openrsxl.worksheet.properties import WorksheetProperties
from openrsxl.worksheet.protection import SheetProtection
from openrsxl.worksheet.scenario import ScenarioList
from openrsxl.worksheet.views import SheetViewList

#: (worksheet attribute, class, default factory) of the "sheet" section
SHEET_OBJECTS = (
    ("sheet_properties", WorksheetProperties, WorksheetProperties),
    ("views", SheetViewList, SheetViewList),
    ("sheet_format", SheetFormatProperties, SheetFormatProperties),
    ("print_options", PrintOptions, PrintOptions),
    ("page_margins", PageMargins, PageMargins),
    ("page_setup", PrintPageSetup, PrintPageSetup),
    ("HeaderFooter", HeaderFooter, HeaderFooter),
    ("auto_filter", AutoFilter, AutoFilter),
    ("protection", SheetProtection, SheetProtection),
    ("row_breaks", RowBreak, RowBreak),
    ("col_breaks", ColBreak, ColBreak),
    ("scenarios", ScenarioList, ScenarioList),
)

#: XML tag -> (worksheet attribute, class)
SHEET_TAGS = {
    "sheetPr": ("sheet_properties", WorksheetProperties),
    "sheetViews": ("views", SheetViewList),
    "sheetFormatPr": ("sheet_format", SheetFormatProperties),
    "printOptions": ("print_options", PrintOptions),
    "pageMargins": ("page_margins", PageMargins),
    "pageSetup": ("page_setup", PrintPageSetup),
    "headerFooter": ("HeaderFooter", HeaderFooter),
    "autoFilter": ("auto_filter", AutoFilter),
    "sheetProtection": ("protection", SheetProtection),
    "rowBreaks": ("row_breaks", RowBreak),
    "colBreaks": ("col_breaks", ColBreak),
    "scenarios": ("scenarios", ScenarioList),
}


def read_object(cls, el):
    """`cls` from an element as openpyxl's worksheet parser reads it (the
    password of a sheet protection is stored hashed)"""
    obj = cls.from_tree(el)
    if cls is SheetProtection:
        password = el.get("password")
        if password is not None:
            obj.set_password(password, True)
    return obj


def _num(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, float):
        return repr(v)
    return str(v)


_COLUMN = ColumnDimension(None)
_ROW = RowDimension(None)

#: (field, default) of column dimensions (model attributes of ColumnDimension)
COLUMN_FIELDS = tuple(
    (f, getattr(_COLUMN, f)) for f in ("min", "max", "width", "bestFit", "hidden", "outlineLevel", "collapsed")
)

#: (field, default) of row dimensions
ROW_FIELDS = tuple((f, getattr(_ROW, f)) for f in ("ht", "hidden", "outlineLevel", "collapsed", "thickBot", "thickTop"))


def column_dimension_items(key, cd, style_id):
    items = []
    if cd.min is None or get_column_letter(cd.min) != key:
        items.append(f"index={key}")
    for field, default in COLUMN_FIELDS:
        v = getattr(cd, field)
        if not (v == default and type(v) is type(default)):
            items.append(f"{field}={'none' if v is None else _num(v)}")
    if style_id is not None:
        items.append(f"s=s{style_id}")
    return " ".join(items)


def row_dimension_items(rd, style_id):
    items = []
    for field, default in ROW_FIELDS:
        v = getattr(rd, field)
        if not (v == default and type(v) is type(default)):
            items.append(f"{field}={'none' if v is None else _num(v)}")
    if style_id is not None:
        items.append(f"s=s{style_id}")
    return " ".join(items)


def parse_field(field, raw):
    """Value of a dimension field written by column/row_dimension_items"""
    if raw == "none":
        return None
    if field in ("min", "max", "outlineLevel"):
        return int(raw)
    if field in ("width", "ht"):
        return float(raw)
    return raw == "1"
