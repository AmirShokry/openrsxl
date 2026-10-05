# Copyright (c) 2010-2024 openpyxl
import pytest

from openrsxl.xml.functions import fromstring, tostring
from upstream_helper import compare_xml

@pytest.fixture
def UpDownBars():
    from openrsxl.chart.updown_bars import  UpDownBars
    return UpDownBars


class TestUpDownBars:

    def test_ctor(self, UpDownBars):
        bars = UpDownBars(gapWidth=150)
        xml = tostring(bars.to_tree())
        expected = """
        <upbars>
          <gapWidth val="150"/>
        </upbars>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, UpDownBars):
        src = """
        <upDownBars>
          <gapWidth val="156"/>
        </upDownBars>
        """
        node = fromstring(src)
        bars = UpDownBars.from_tree(node)
        assert bars == UpDownBars(gapWidth=156)
