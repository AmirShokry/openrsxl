# Copyright (c) 2010-2024 openpyxl

from openrsxl.xml.constants import CHART_NS

from openrsxl.descriptors.serialisable import Serialisable
from openrsxl.descriptors.excel import Relation


class ChartRelation(Serialisable):

    tagname = "chart"
    namespace = CHART_NS

    id = Relation()

    def __init__(self, id):
        self.id = id
