# Copyright (c) 2010-2024 openpyxl

from openrsxl.descriptors.serialisable import Serialisable
from openrsxl.descriptors import (
    Sequence
)
from openrsxl.descriptors.excel import (
    Relation,
)

class ExternalReference(Serialisable):

    tagname = "externalReference"

    id = Relation()

    def __init__(self, id):
        self.id = id
