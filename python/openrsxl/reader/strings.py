# Copyright (c) 2010-2024 openpyxl

from openrsxl.cell.text import Text

from openrsxl.xml.functions import iterparse
from openrsxl.xml.constants import SHEET_MAIN_NS
from openrsxl.cell.rich_text import CellRichText


def read_string_table(xml_source):
    """Read in all shared strings in the table"""
    from openrsxl import _native
    if _native.ENABLED and hasattr(xml_source, "read") and hasattr(xml_source, "seek"):
        try:
            return _native._rs.read_string_table(xml_source, _native.reader_helpers())
        except _native.NativeFallback:
            xml_source.seek(0)
    return _read_string_table(xml_source)


def _read_string_table(xml_source):
    """Read in all shared strings in the table (pure Python)"""

    strings = []
    STRING_TAG = '{%s}si' % SHEET_MAIN_NS

    for _, node in iterparse(xml_source):
        if node.tag == STRING_TAG:
            text = Text.from_tree(node).content
            text = text.replace('x005F_', '')
            node.clear()

            strings.append(text)

    return strings


def read_rich_text(xml_source):
    """Read in all shared strings in the table"""

    strings = []
    STRING_TAG = '{%s}si' % SHEET_MAIN_NS

    for _, node in iterparse(xml_source):
        if node.tag == STRING_TAG:
            text = CellRichText.from_tree(node)
            if len(text) == 0:
                text = ''
            elif len(text) == 1 and isinstance(text[0], str):
                text = text[0]
            node.clear()

            strings.append(text)

    return strings
