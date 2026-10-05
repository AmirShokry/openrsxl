"""openrsxl.extended.load_workbook: openpyxl's load_workbook + opt-in features"""

from openrsxl.packaging.relationship import RelationshipList, get_dependents, get_rels_path
from openrsxl.reader.excel import KEEP_VBA, ExcelReader

from ._streaming import FLAGS, ReadOnlyWorksheet, _Options


class _ExtendedReader(ExcelReader):

    def __init__(self, *args, options=None, **kw):
        super().__init__(*args, **kw)
        self._ext = options

    def read_worksheets(self):
        if not self.read_only:
            return super().read_worksheets()
        for sheet, rel in self.parser.find_sheets():
            if rel.target not in self.valid_files:
                continue
            if "chartsheet" in rel.Type:
                self.read_chartsheet(sheet, rel)
                continue
            rels_path = get_rels_path(rel.target)
            rels = RelationshipList()
            if rels_path in self.valid_files:
                rels = get_dependents(self.archive, rels_path)
            ws = ReadOnlyWorksheet(self.wb, sheet.name, rel.target, self.shared_strings, self._ext)
            ws.sheet_state = sheet.state
            self.wb._sheets.append(ws)
            ws._ext_bind(self.archive, rels)


def load_workbook(
    filename,
    read_only=False,
    keep_vba=KEEP_VBA,
    data_only=False,
    keep_links=True,
    rich_text=False,
    *,
    formula_and_value=False,
    read_comments=False,
    read_hyperlinks=False,
    read_merged_cells=False,
    read_dimensions=False,
    read_sheet_properties=False,
    read_data_validations=False,
    read_conditional_formatting=False,
    read_tables=False,
):
    """Open the given filename and return the workbook

    Same as ``openpyxl.load_workbook``, plus opt-in features for read-only
    (streaming) workbooks. Each one adds what openpyxl's full mode offers,
    with the same names and types:

    :param formula_and_value: ``cell.formula`` (the formula, as read with
        ``data_only=False``, None for other cells) and ``cell.cached_value``
        (the value as read with ``data_only=True``)
    :param read_comments: ``cell.comment``
    :param read_hyperlinks: ``cell.hyperlink``
    :param read_merged_cells: ``ws.merged_cells``; cells covered by a merged
        range are ``MergedCell`` objects
    :param read_dimensions: ``ws.row_dimensions`` / ``ws.column_dimensions``
    :param read_sheet_properties: ``ws.sheet_properties``, ``views``,
        ``freeze_panes``, ``sheet_format``, ``print_options``,
        ``page_margins``, ``page_setup``, ``HeaderFooter``, ``auto_filter``,
        ``protection``, ``row_breaks``, ``col_breaks``, ``scenarios``,
        print titles / area
    :param read_data_validations: ``ws.data_validations``
    :param read_conditional_formatting: ``ws.conditional_formatting``
    :param read_tables: ``ws.tables``

    The ``read_*`` flags have no effect without ``read_only=True`` (full mode
    always reads everything); ``formula_and_value`` requires
    ``read_only=True``. Without any flag the result is exactly openpyxl's.

    Example::

        wb = load_workbook("book.xlsx", read_only=True, formula_and_value=True,
                           read_comments=True, read_merged_cells=True)
        for row in wb.active.iter_rows():
            for cell in row:
                cell.value, cell.formula, cell.cached_value, cell.comment

    Rows are streamed as in read-only mode (memory does not grow with the
    number of cells) and every value is what openpyxl's full mode returns
    for the same file. ``cell.value`` follows ``data_only`` as in openpyxl;
    the attribute of a disabled feature does not exist (AttributeError);
    ``iter_rows(values_only=True)`` yields values only (merged cells are
    None, link targets fill empty linked cells, as in full mode).

    Limitations (docs/COMPATIBILITY.md): rows / columns and cell positions
    are those of read-only mode (the sheet's ``<dimension>``); errors in
    cell data that make full mode fail at load time are raised while
    iterating; images, charts, drawings and pivot tables are not streamed.
    """
    options = _Options(
        formula_and_value=formula_and_value,
        read_comments=read_comments,
        read_hyperlinks=read_hyperlinks,
        read_merged_cells=read_merged_cells,
        read_dimensions=read_dimensions,
        read_sheet_properties=read_sheet_properties,
        read_data_validations=read_data_validations,
        read_conditional_formatting=read_conditional_formatting,
        read_tables=read_tables,
    )
    if formula_and_value and not read_only:
        raise ValueError("formula_and_value=True requires read_only=True")
    if not read_only or not any(getattr(options, k) for k in FLAGS):
        # exactly openpyxl
        reader = ExcelReader(filename, read_only, keep_vba, data_only, keep_links, rich_text)
    else:
        reader = _ExtendedReader(filename, read_only, keep_vba, data_only, keep_links, rich_text, options=options)
    reader.read()
    return reader.wb


open = load_workbook
