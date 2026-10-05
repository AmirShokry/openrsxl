"""
openrsxl: compact cell storage for loaded worksheets.

`CellStore` replaces the ``Worksheet._cells`` dictionary of worksheets read
from a file.  Following calamine / python-calamine, cell values are kept in
compact Rust form and Python objects are only created on demand: a `Cell`
object is created when a cell is accessed (and then cached, so identity
semantics are those of a dict).  The class implements the complete dict
interface with dict semantics (insertion order, KeyError, views, ...).

Known difference (docs/COMPATIBILITY.md): it is not a ``dict`` subclass, so
``isinstance(ws._cells, dict)`` is false (``collections.abc.MutableMapping``
is true). On free-threaded Python, one store used concurrently by several
threads raises ``RuntimeError: Already borrowed`` where a dict would race.
"""

from collections.abc import Mapping, MutableMapping

from openrsxl._openrsxl import CellStore as _CellStore


class CellStore(_CellStore, MutableMapping):

    __slots__ = ()
    __hash__ = None  # type: ignore[assignment]  # unhashable, as dict

    def __repr__(self):
        return repr(dict(self.items()))

    def __eq__(self, other):
        if not isinstance(other, Mapping):
            return NotImplemented
        return dict(self.items()) == dict(other.items())

    def copy(self):
        return dict(self.items())

    def __copy__(self):
        return dict(self.items())

    def __reduce__(self):
        return (dict, (list(self.items()),))

    def __or__(self, other):
        if not isinstance(other, Mapping):
            return NotImplemented
        new = dict(self.items())
        new.update(other)
        return new

    def __ror__(self, other):
        if not isinstance(other, Mapping):
            return NotImplemented
        new = dict(other)
        new.update(self.items())
        return new

    def __ior__(self, other):
        self.update(other)
        return self

    @classmethod
    def fromkeys(cls, iterable, value=None):
        return dict.fromkeys(iterable, value)
