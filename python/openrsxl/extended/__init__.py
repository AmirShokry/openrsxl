"""
openrsxl.extended
=================

openpyxl's API plus features openpyxl does not have.

Every public name of ``openrsxl`` is available here and every submodule can
be imported through this package; submodules are *aliases* of the
``openrsxl`` modules (the very same objects, e.g.
``openrsxl.extended.styles.Font is openrsxl.styles.Font``)::

    from openrsxl.extended import load_workbook, Workbook
    from openrsxl.extended.styles import Font

Additions:

``load_workbook(..., read_only=True, <flags>)``
    streaming (read-only) workbooks with what openpyxl only offers in full
    mode - formulas *and* cached values in one pass, comments, hyperlinks,
    merged cells, row / column dimensions, sheet properties, data
    validation, conditional formatting and tables - each behind its own
    keyword-only flag (``formula_and_value``, ``read_comments``,
    ``read_hyperlinks``, ``read_merged_cells``, ``read_dimensions``,
    ``read_sheet_properties``, ``read_data_validations``,
    ``read_conditional_formatting``, ``read_tables``). Without a flag it is
    exactly openpyxl's ``load_workbook``. See ``help(load_workbook)``, the
    README and docs/COMPATIBILITY.md (limitations).

New extended features are added to this package without affecting the
openpyxl-compatible ``openrsxl`` namespace; real modules of
``openrsxl/extended/`` take precedence over the aliases.
"""

import importlib
import importlib.abc
import importlib.util
import os
import sys

import openrsxl as _base
from openrsxl import *  # noqa: F401,F403 - re-export the public API
from openrsxl import (  # noqa: F401 - names not covered by `import *`
    DEBUG,
    DEFUSEDXML,
    LXML,
    NUMPY,
    Workbook,
    __author__,
    __author_email__,
    __license__,
    __maintainer_email__,
    __url__,
    __version__,
    constants,
    load_workbook,
    open,
)

_PREFIX = __name__ + "."
_HERE = os.path.dirname(__file__)


class _AliasLoader(importlib.abc.Loader):

    def __init__(self, target):
        self.target = target

    def create_module(self, spec):
        return importlib.import_module(self.target)

    def exec_module(self, module):
        # the aliased module is already initialised
        pass


class _AliasFinder(importlib.abc.MetaPathFinder):
    """Resolve ``openrsxl.extended.<x>`` to ``openrsxl.<x>``."""

    def find_spec(self, fullname, path=None, target=None):
        if not fullname.startswith(_PREFIX):
            return None
        rest = fullname[len(_PREFIX) :]
        # real extended modules win
        candidate = os.path.join(_HERE, *rest.split("."))
        if os.path.exists(candidate + ".py") or os.path.isdir(candidate):
            return None
        target_name = "openrsxl." + rest
        try:
            target_spec = importlib.util.find_spec(target_name)
        except (ImportError, ValueError):
            return None
        if target_spec is None:
            return None
        return importlib.util.spec_from_loader(
            fullname, _AliasLoader(target_name), is_package=target_spec.submodule_search_locations is not None
        )


if not any(isinstance(f, _AliasFinder) for f in sys.meta_path):
    sys.meta_path.insert(0, _AliasFinder())


def __getattr__(name):
    # `openrsxl.extended.styles` without an explicit import
    try:
        return importlib.import_module(_PREFIX + name)
    except ImportError:
        pass
    try:
        return getattr(_base, name)
    except AttributeError:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}") from None


def __dir__():
    return sorted(set(globals()) | set(dir(_base)))


# -- extended features -------------------------------------------------------
from ._load import load_workbook, open  # noqa: E402,F401,F811
