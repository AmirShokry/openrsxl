# Copyright (c) 2010-2024 openpyxl


from openrsxl.descriptors.serialisable import Serialisable
from openrsxl.descriptors import (
    Sequence,
    Alias
)


class AuthorList(Serialisable):

    tagname = "authors"

    author = Sequence(expected_type=str)
    authors = Alias("author")

    def __init__(self,
                 author=(),
                ):
        self.author = author
