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

from collections.abc import ItemsView, KeysView, Mapping, MutableMapping, ValuesView
from types import MappingProxyType

from openrsxl._openrsxl import CellStore as _CellStore

_MISSING = object()


class _ViewMixin:
    """What dict views add to collections.abc's: reversed(), .mapping, repr."""

    __slots__ = ()

    @property
    def mapping(self):
        return MappingProxyType(self._mapping)

    def __repr__(self):
        return f"{self._dict_name}({list(self)!r})"


class _Keys(_ViewMixin, KeysView):
    __slots__ = ()
    _dict_name = "dict_keys"

    def __reversed__(self):
        return reversed(self._mapping)


class _Values(_ViewMixin, ValuesView):
    __slots__ = ()
    _dict_name = "dict_values"

    def __reversed__(self):
        m = self._mapping
        return (m[k] for k in reversed(m))


class _Items(_ViewMixin, ItemsView):
    __slots__ = ()
    _dict_name = "dict_items"

    def __reversed__(self):
        m = self._mapping
        return ((k, m[k]) for k in reversed(m))


class CellStore(_CellStore, MutableMapping):

    __slots__ = ()
    __hash__ = None  # type: ignore[assignment]  # unhashable, as dict

    def keys(self):
        return _Keys(self)

    def values(self):
        return _Values(self)

    def items(self):
        return _Items(self)

    def pop(self, key, default=_MISSING):
        if not len(self):
            # dict.pop of an empty dict does not hash the key
            if default is _MISSING:
                raise KeyError(key)
            return default
        if default is _MISSING:
            return MutableMapping.pop(self, key)
        return MutableMapping.pop(self, key, default)

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
