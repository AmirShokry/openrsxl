"""openrsxl.extended.load_workbook: openpyxl's load_workbook + opt-in features"""

import types
from warnings import warn

from openrsxl.packaging.relationship import RelationshipList, get_dependents, get_rels_path
from openrsxl.reader.excel import KEEP_VBA, ExcelReader
from openrsxl.worksheet.print_settings import PrintArea, PrintTitles

from ._streaming import FLAGS, ReadOnlyWorksheet, _Options


class _ExtendedReader(ExcelReader):

    def __init__(self, *args, options=None, only=None, skip_invalid_defined_names=False, **kw):
        super().__init__(*args, **kw)
        self._ext = options or _Options()
        # names of the sheets that get the features (None: all); the others
        # are plain read-only worksheets, their XML is not scanned at load
        self._only = only
        self._skip_names = skip_invalid_defined_names

    def read_workbook(self):
        super().read_workbook()
        if self._skip_names:
            self.parser.assign_names = types.MethodType(_assign_names, self.parser)

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
            options = self._ext if self._only is None or sheet.name in self._only else _Options()
            ws = ReadOnlyWorksheet(self.wb, sheet.name, rel.target, self.shared_strings, options)
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
    read_all=False,
    formula_and_value=None,
    read_comments=None,
    read_hyperlinks=None,
    read_merged_cells=None,
    read_dimensions=None,
    read_sheet_properties=None,
    read_data_validations=None,
    read_conditional_formatting=None,
    read_tables=None,
    create_empty_cells=None,
    exact_dimensions=False,
    random_access=False,
    reopen=False,
    skip_invalid_defined_names=False,
):
    """Open the given filename and return the workbook

    Same as ``openpyxl.load_workbook``, plus opt-in features for read-only
    (streaming) workbooks. Each one adds what openpyxl's full mode offers,
    with the same names and types:

    :param read_all: every feature below; a feature given explicitly wins,
        so ``read_all=True, read_comments=False`` is every feature except
        comments
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
    :param create_empty_cells: every empty position of a row is a cell like
        the one full mode creates when it is accessed - ``coordinate``,
        ``row``, ``column``, value None, default style, None for the enabled
        features - instead of read-only mode's shared, coordinate-less
        ``EmptyCell``; created on the fly, nothing is kept

    Options outside ``read_all`` (default False; read-only workbooks):

    :param exact_dimensions: ``min_row``, ``max_row``, ``min_column``,
        ``max_column`` - and so the rows and columns of ``iter_rows()`` -
        from the cells present (every cell of the file, also one with a
        style only, and the cells the enabled features create), as full
        mode, instead of the sheet's ``<dimension>`` element; computed by
        the pre-scan at load time
    :param random_access: ``ws.cell()`` and ``ws[...]`` (cells, ranges,
        columns, rows) without reading the sheet again on every call: the
        first one builds a store of the sheet's cells (one streaming pass),
        later calls - and ``iter_rows`` - are served from it; the same cell
        objects as ``iter_rows``. Ranges, rows, columns and ``iter_rows``
        hold every position of their bounds, as full mode (also after the
        last row of the file, where read-only mode stops). ``ws._cells`` is a read-only mapping
        {(row, column): cell} of the cells full mode keeps.
        ``openrsxl.extended.release_cells(ws)`` frees the store
    :param reopen: a workbook closed with ``close()`` reopens its file (the
        path, or the file object it was loaded from) when it is used again,
        as full mode stays usable after ``close()``; the shared strings are
        kept

    :param skip_invalid_defined_names: (any mode) a print titles defined
        name that openpyxl cannot parse (eg. ``#N/A``) is skipped with a
        warning instead of failing the load

    A feature flag left unset (None) follows ``read_all``; True / False
    enable / disable it whatever ``read_all`` is. The flags have no effect
    without ``read_only=True`` (full mode always reads everything and
    creates cells on access), except an explicit ``formula_and_value=True``,
    which requires ``read_only=True`` (ValueError). Without any flag the
    result is exactly openpyxl's.

    Example::

        wb = load_workbook("book.xlsx", read_only=True, formula_and_value=True,
                           read_comments=True, read_merged_cells=True)
        for row in wb.active.iter_rows():
            for cell in row:
                cell.value, cell.formula, cell.cached_value, cell.comment

        # everything full mode has, except cells at empty positions
        wb = load_workbook("book.xlsx", read_only=True, read_all=True,
                           create_empty_cells=False)

    Rows are streamed as in read-only mode (memory does not grow with the
    number of cells) and every value is what openpyxl's full mode returns
    for the same file. ``cell.value`` follows ``data_only`` as in openpyxl;
    the attribute of a disabled feature does not exist (AttributeError);
    ``iter_rows(values_only=True)`` yields values only (merged cells are
    None, link targets fill empty linked cells, as in full mode).

    Warnings of a sheet's data (eg. unsupported extensions) are reported by
    its first complete pass only, as full mode reports them once per load.

    Limitations (docs/COMPATIBILITY.md): rows / columns and cell positions
    are those of read-only mode (the sheet's ``<dimension>``, unless
    ``exact_dimensions``); iterating or accessing a position does not
    create a cell or extend the sheet as full mode does; errors in
    cell data that make full mode fail at load time are raised while
    iterating; images, charts, drawings and pivot tables are not streamed.
    """
    given = dict(
        formula_and_value=formula_and_value,
        read_comments=read_comments,
        read_hyperlinks=read_hyperlinks,
        read_merged_cells=read_merged_cells,
        read_dimensions=read_dimensions,
        read_sheet_properties=read_sheet_properties,
        read_data_validations=read_data_validations,
        read_conditional_formatting=read_conditional_formatting,
        read_tables=read_tables,
        create_empty_cells=create_empty_cells,
    )
    if formula_and_value and not read_only:
        raise ValueError("formula_and_value=True requires read_only=True")
    options = _Options(
        **{k: bool(read_all) if v is None else v for k, v in given.items()},
        exact_dimensions=exact_dimensions,
        random_access=random_access,
        reopen=reopen,
    )
    streaming = read_only and options.any()
    if not streaming and not skip_invalid_defined_names:
        # exactly openpyxl
        reader = ExcelReader(filename, read_only, keep_vba, data_only, keep_links, rich_text)
    else:
        reader = _ExtendedReader(
            filename,
            read_only,
            keep_vba,
            data_only,
            keep_links,
            rich_text,
            options=options if streaming else None,
            skip_invalid_defined_names=skip_invalid_defined_names,
        )
    reader.read()
    if streaming and options.reopen:
        reader.wb._ext_source = filename
    return reader.wb


def _assign_names(self):
    """WorkbookParser.assign_names; a print titles definition openpyxl
    cannot parse (ValueError, eg. #N/A) is skipped with a warning"""
    for idx, names in self.defined_names.by_sheet().items():
        if idx == "global":
            self.wb.defined_names = names
            continue

        try:
            sheet = self.wb._sheets[idx]
        except IndexError:
            warn(f"Defined names for sheet index {idx} cannot be located")
            continue

        for name, defn in names.items():
            reserved = defn.is_reserved
            if reserved is None:
                sheet.defined_names[name] = defn

            elif reserved == "Print_Titles":
                try:
                    titles = PrintTitles.from_string(defn.value)
                except ValueError:
                    warn(f"Print titles cannot be set to Defined name: {defn.value}.")
                    continue
                sheet._print_rows = titles.rows
                sheet._print_cols = titles.cols
            elif reserved == "Print_Area":
                try:
                    sheet._print_area = PrintArea.from_string(defn.value)
                except TypeError:
                    warn(f"Print area cannot be set to Defined name: {defn.value}.")
                    continue


open = load_workbook


def _load_for(filename, sheets, keep_links=True, **flags):
    """Read-only workbook with the features `flags` bound only on the sheets
    named in `sheets` (None: all) - for openrsxl.extended.llm, which reads
    the metadata of the sheets it compresses only"""
    options = _Options(**flags)
    only = None if sheets is None else set([sheets] if isinstance(sheets, str) else sheets)
    if not any(getattr(options, k) for k in FLAGS):
        # exactly openpyxl, as load_workbook
        reader = ExcelReader(filename, True, KEEP_VBA, False, keep_links, False)
    else:
        reader = _ExtendedReader(filename, True, KEEP_VBA, False, keep_links, False, options=options, only=only)
    reader.read()
    return reader.wb
