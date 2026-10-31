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
    ``read_conditional_formatting``, ``read_tables``), and full mode's cells
    at empty positions (``create_empty_cells``); ``read_all=True`` enables
    every feature not given explicitly (``read_all=True,
    read_comments=False``: all but comments). Without a flag it is exactly
    openpyxl's ``load_workbook``. See ``help(load_workbook)``, the
    README and docs/COMPATIBILITY.md (limitations).

``openrsxl.extended.llm``
    lossless, LLM readable compression of workbooks: ``compress_workbook``
    (separable sections - values, formulas, styles, metadata - to memory,
    text or JSON files, each sheet's encoding chosen by measuring it),
    ``decompress``, ``analyze_workbook``. See ``help(openrsxl.extended.llm)``
    and docs/LLM_ENCODING.md.

``release(wb)``
    free the memory of a workbook at once (``close()`` keeps openpyxl's
    meaning: a full-mode workbook stays usable, so its cells stay in memory).

``install_as_openpyxl()``
    make ``import openpyxl`` (in this process, also inside other libraries)
    import openrsxl, so that code written for openpyxl - including its
    ``isinstance`` checks - uses openrsxl unchanged.

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
        self._spec = None

    def create_module(self, spec):
        module = importlib.import_module(self.target)
        self._spec = module.__spec__
        return module

    def exec_module(self, module):
        # the aliased module is already initialised; the import system has
        # just set its __spec__ to the alias' spec: restore its own (relative
        # imports and importlib.reload rely on it)
        module.__spec__ = self._spec


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


class _OpenpyxlFinder(importlib.abc.MetaPathFinder):
    """Resolve ``openpyxl`` / ``openpyxl.<x>`` to ``openrsxl`` / ``openrsxl.<x>``."""

    def find_spec(self, fullname, path=None, target=None):
        if fullname != "openpyxl" and not fullname.startswith("openpyxl."):
            return None
        target_name = "openrsxl" + fullname[len("openpyxl") :]
        try:
            target_spec = importlib.util.find_spec(target_name)
        except (ImportError, ValueError):
            return None
        if target_spec is None:
            return None
        return importlib.util.spec_from_loader(
            fullname, _AliasLoader(target_name), is_package=target_spec.submodule_search_locations is not None
        )


def install_as_openpyxl():
    """
    Make ``import openpyxl`` import openrsxl, in this process.

    Every ``openpyxl`` module - ``import openpyxl``, ``from openpyxl.styles
    import Font``, also inside other libraries such as pandas - is then the
    corresponding ``openrsxl`` module: ``openpyxl.worksheet.formula.ArrayFormula
    is openrsxl.worksheet.formula.ArrayFormula``, so ``isinstance`` checks
    written for openpyxl hold for openrsxl's objects. Call it once at start-up,
    before anything imports openpyxl (RuntimeError otherwise: objects of the
    real openpyxl would exist next to openrsxl's); calling it again is a
    no-op. Like PyMySQL's ``install_as_MySQLdb()``.
    """
    for name, module in list(sys.modules.items()):
        if name == "openpyxl" or name.startswith("openpyxl."):
            if module is not sys.modules.get("openrsxl" + name[len("openpyxl") :]):
                raise RuntimeError(
                    "openpyxl is already imported: call install_as_openpyxl() before anything imports openpyxl"
                )
    if not any(isinstance(f, _OpenpyxlFinder) for f in sys.meta_path):
        sys.meta_path.insert(0, _OpenpyxlFinder())


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


def release(wb):
    """
    Free the memory of a workbook now, in every mode, and leave it unusable.

    ``wb.close()`` keeps openpyxl's meaning: a full-mode workbook stays
    usable after it, so its cells stay in memory; a read-only one drops its
    shared strings. A workbook and its worksheets reference each other, so
    once dropped the rest waits for Python's cycle collector.
    ``release(wb)`` closes the workbook, empties the cell stores and breaks
    those references: the memory is returned at once (whatever other
    references to the workbook, its worksheets or cells remain, they no
    longer hold it). Any use of the workbook or its worksheets afterwards
    fails. Returns None.
    """
    wb.close()
    sheets = list(getattr(wb, "_sheets", ()))
    for ws in sheets:
        cells = getattr(ws, "_cells", None)
        if cells is not None and hasattr(cells, "clear"):
            cells.clear()
        d = getattr(ws, "__dict__", None)
        if d is not None:
            d.clear()
    archive = getattr(wb, "vba_archive", None)
    if archive is not None:
        try:
            archive.close()
        except Exception:  # pragma: no cover - already closed
            pass
    wb.__dict__.clear()
