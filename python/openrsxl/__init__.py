# Copyright (c) 2010-2024 openpyxl

DEBUG = False

from openrsxl.compat.numbers import NUMPY
from openrsxl.xml import DEFUSEDXML, LXML
from openrsxl.workbook import Workbook
from openrsxl.reader.excel import load_workbook as open
from openrsxl.reader.excel import load_workbook
import openrsxl._constants as constants

# Expose constants especially the version number

__author__ = constants.__author__
__author_email__ = constants.__author_email__
__license__ = constants.__license__
__maintainer_email__ = constants.__maintainer_email__
__url__ = constants.__url__
__version__ = constants.__version__
