# Root conftest for the ported upstream openpyxl test-suite (see
# tools/port_upstream_tests.py).  Mirrors the upstream root conftest.py.

import os
import sys
import platform

import pytest

sys.path.insert(0, os.path.dirname(__file__))  # for upstream_helper
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))  # upstream.* namespace


def pytest_runtest_setup(item):
    from openrsxl import DEFUSEDXML, LXML
    if isinstance(item, pytest.Function):
        try:
            from PIL import Image
        except ImportError:
            Image = False
        if item.get_closest_marker("pil_required") and Image is False:
            pytest.skip("PIL must be installed")
        elif item.get_closest_marker("pil_not_installed") and Image:
            pytest.skip("PIL is installed")
        elif item.get_closest_marker("not_py33"):
            pytest.skip("Ordering is not a given in Python 3")
        elif item.get_closest_marker("defusedxml_required"):
            if LXML or not DEFUSEDXML:
                pytest.skip("defusedxml is required to guard against these vulnerabilities")
        elif item.get_closest_marker("lxml_required"):
            if not LXML:
                pytest.skip("LXML is required for some features such as schema validation")
        elif item.get_closest_marker("lxml_buffering"):
            from lxml.etree import LIBXML_VERSION
            if LIBXML_VERSION < (3, 4, 0, 0):
                pytest.skip("LXML >= 3.4 is required")
        elif item.get_closest_marker("no_lxml"):
            if LXML:
                pytest.skip("LXML has a different interface")
        elif item.get_closest_marker("numpy_required"):
            from openrsxl import NUMPY
            if not NUMPY:
                pytest.skip("Numpy must be installed")
        elif item.get_closest_marker("pandas_required"):
            try:
                import pandas  # noqa
            except ImportError:
                pytest.skip("Pandas must be installed")
        elif item.get_closest_marker("no_pypy"):
            if platform.python_implementation() == "PyPy":
                pytest.skip("Skipping pypy")


def pytest_collection_modifyitems(items):
    # openpyxl 3.1.5's own test expects lxml.etree.fromstring() to reject a
    # file object with ValueError; lxml >= 6 raises TypeError, so the test
    # fails in the same way for openpyxl itself (not an openrsxl difference)
    try:
        from lxml.etree import LXML_VERSION
    except ImportError:
        return
    if LXML_VERSION < (6,):
        return
    for item in items:
        path = str(item.fspath).replace(os.sep, "/")
        if path.endswith("xml/tests/test_functions.py") and getattr(item, "originalname", "") == "test_iterparse":
            item.add_marker(pytest.mark.xfail(raises=TypeError, strict=True,
                                              reason="openpyxl's test predates lxml 6 (TypeError, not ValueError)"))
