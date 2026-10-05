# Copyright (c) 2010-2024 openpyxl

from upstream_helper import compare_xml

from openrsxl.xml.functions import tostring


def test_related():
    from openrsxl.worksheet.related import  Related
    rel = Related(id="rId1")
    expected = """
    <drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
    """
    xml = tostring(rel.to_tree("drawing"))
    diff = compare_xml(xml, expected)
    assert diff is None, diff
