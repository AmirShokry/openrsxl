"""
Glue between the Rust engine (``openrsxl._openrsxl``) and the Python object
model.

The Rust engine handles the hot paths (shared strings, worksheet cells,
``<sheetData>`` serialisation).  Whenever it meets input it cannot reproduce
with certainty it raises ``NativeFallback`` and the original pure Python
implementation is used instead, so observable behaviour is always identical
to openpyxl.

Set the environment variable ``OPENRSXL_PURE_PYTHON=1`` to disable the
native engine entirely (useful for debugging / differential testing).

The engine creates Python objects by writing their slots directly, which
needs CPython's object layout: on other implementations (PyPy's cpyext
emulation, GraalPy) it is disabled and openpyxl's Python code is used.
"""

import os
import platform
from io import BytesIO

from openrsxl.xml.functions import fromstring

try:
    from openrsxl import _openrsxl as _rs
except ImportError:  # pragma: no cover - source checkout without build
    _rs = None

ENABLED = (
    _rs is not None
    and platform.python_implementation() == "CPython"
    and os.environ.get("OPENRSXL_PURE_PYTHON", "") not in ("1", "True", "true")
)

if _rs is not None:
    NativeFallback = _rs.NativeFallback
else:  # pragma: no cover

    class NativeFallback(Exception):  # type: ignore[no-redef]
        pass


class open_source:
    """
    Context manager giving a seekable binary stream for `source`, which may
    be a file-like object or a path (as accepted by ElementTree.iterparse).
    """

    def __init__(self, source):
        self.source = source
        self.owned = None

    def __enter__(self):
        src = self.source
        if hasattr(src, "read"):
            return src
        self.owned = open(src, "rb")
        return self.owned

    def __exit__(self, *exc):
        if self.owned is not None:
            self.owned.close()


def usable(source):
    """Can the native engine read from `source`?"""
    if hasattr(source, "read"):
        return hasattr(source, "seek")
    return isinstance(source, (str, bytes)) or hasattr(source, "__fspath__")


def _escape_attr(value):
    return value.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")


def wrap(snippet, namespaces):
    """
    Parse an element snippet captured by the native reader, restoring the
    namespace declarations that were in scope in the original document.
    """
    decls = []
    seen = set()
    for prefix, uri in namespaces:
        seen.add(prefix)
    for prefix, uri in namespaces:
        if prefix:
            decls.append(f' xmlns:{prefix}="{_escape_attr(uri)}"')
        else:
            decls.append(f' xmlns="{_escape_attr(uri)}"')
    head = ("<__openrsxl_wrap%s>" % "".join(decls)).encode("utf-8")
    return head + snippet + b"</__openrsxl_wrap>"


def wrapped_element(snippet, namespaces):
    root = fromstring(wrap(snippet, namespaces))
    return root[0]


# --------------------------------------------------------------------------
# helpers called from Rust


def _row_index(value):
    # mirrors WorkSheetParser.parse_row
    try:
        return int(value)
    except ValueError:
        val = float(value)
        if val.is_integer():
            return int(val)
        else:
            raise ValueError(f"{value} is not a valid row number")


def _date_warning(coordinate, value):
    return f"""Cell {coordinate} is marked as a date but the serial value {value} is outside the limits for dates. The cell will be treated as an error."""


def _text_content(snippet, namespaces):
    from openrsxl.cell.text import Text

    return Text.from_tree(wrapped_element(snippet, namespaces)).content


def _si_content(snippet, namespaces):
    return _text_content(snippet, namespaces).replace("x005F_", "")


def _rich_inline(snippet, namespaces):
    from openrsxl.worksheet._reader import parse_richtext_string

    return parse_richtext_string(wrapped_element(snippet, namespaces))


def _validate_rpr(snippet, namespaces):
    from openrsxl.cell.text import InlineFont

    try:
        InlineFont.from_tree(wrapped_element(snippet, namespaces))
    except Exception:
        return False
    return True


_HELPERS = None


def reader_helpers():
    global _HELPERS
    if _HELPERS is None:
        from openrsxl.cell.text import RichText, Text
        from openrsxl.formula.tokenizer import TokenizerError
        from openrsxl.formula.translate import Translator, TranslatorError
        from openrsxl.utils.cell import coordinate_to_tuple
        from openrsxl.utils.datetime import from_excel, from_ISO8601
        from openrsxl.worksheet.formula import ArrayFormula, DataTableFormula

        _HELPERS = dict(
            int=int,
            float=float,
            coordinate_to_tuple=coordinate_to_tuple,
            row_index=_row_index,
            from_excel=from_excel,
            from_ISO8601=from_ISO8601,
            date_warning=_date_warning,
            ArrayFormula=ArrayFormula,
            DataTableFormula=DataTableFormula,
            Translator=Translator,
            TokenizerError=TokenizerError,
            TranslatorError=TranslatorError,
            text_content=_text_content,
            si_content=_si_content,
            rich_inline=_rich_inline,
            validate_rpr=_validate_rpr,
            text_dangerous=sorted(set(dir(Text)) - {"t", "r", "rPh", "phoneticPr"}),
            rich_dangerous=sorted(set(dir(RichText)) - {"t", "rPr"}),
        )
    return _HELPERS


def replay(parser, actions):
    """
    Replay recorded actions (top-level elements and warnings) on a
    WorkSheetParser exactly as its ``parse`` loop would have handled them.
    """
    for action in actions:
        kind = action[0]
        if kind == "warn":
            parser._warn(action[1])
        else:
            parser._dispatch_snippet(action[1], action[2])


# --------------------------------------------------------------------------
# writer helpers


def _add_comment(ws, cell):
    from openrsxl.comments.comment_sheet import CommentRecord

    comment = CommentRecord.from_cell(cell)
    ws._comments.append(comment)


def _render_cell(ws, cell, styled):
    """Render one cell with the original openpyxl code (fallback)."""
    from io import StringIO

    from openrsxl import LXML
    from openrsxl.cell._writer import write_cell
    from openrsxl.xml.functions import xmlfile

    if LXML:
        out = BytesIO()
        with xmlfile(out) as xf:
            with xf.element("w"):
                write_cell(xf, ws, cell, styled)
        return out.getvalue()[3:-4]
    out = StringIO()
    with xmlfile(out, encoding="unicode") as xf:
        with xf.element("w"):
            write_cell(xf, ws, cell, styled)
    return out.getvalue()[3:-4]


def _date_value(ws, cell):
    # mirrors openpyxl.cell._writer._set_attributes for data_type "d"
    from datetime import timedelta

    from openrsxl.utils.datetime import to_excel, to_ISO8601

    value = cell._value
    if hasattr(value, "tzinfo") and value.tzinfo is not None:
        raise TypeError(
            "Excel does not support timezones in datetimes. "
            "The tzinfo in the datetime/time object must be set to None."
        )

    if cell.parent.parent.iso_dates and not isinstance(value, timedelta):
        return to_ISO8601(value), True
    return to_excel(value, cell.parent.parent.epoch), False


def _date_value_raw(ws, value):
    # _set_attributes for data_type "d", given the value
    from datetime import timedelta

    from openrsxl.utils.datetime import to_excel, to_ISO8601

    if hasattr(value, "tzinfo") and value.tzinfo is not None:
        raise TypeError(
            "Excel does not support timezones in datetimes. "
            "The tzinfo in the datetime/time object must be set to None."
        )
    wb = ws.parent
    if wb.iso_dates and not isinstance(value, timedelta):
        return to_ISO8601(value), True
    return to_excel(value, wb.epoch), False


_WHELPERS = None


def writer_helpers():
    global _WHELPERS
    if _WHELPERS is None:
        from openrsxl.compat import safe_string
        from openrsxl.styles.cell_style import StyleArray

        _WHELPERS = dict(
            StyleArray=StyleArray,
            add_comment=_add_comment,
            write_cell=_render_cell,
            date_value=_date_value,
            date_value_raw=_date_value_raw,
            safe_string=safe_string,
        )
    return _WHELPERS
