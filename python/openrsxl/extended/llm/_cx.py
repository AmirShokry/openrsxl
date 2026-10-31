"""
Compact XML ("CX"): a reversible one-line notation for the XML elements
openpyxl writes for styles, sheet properties, validations, ...

========================  =================================================
CX                        XML
========================  =================================================
``name(...)``             element ``<name ...>...</name>``
``k=v``                   attribute ``k="v"``
``k:v``                   child element ``<k val="v"/>``
``k``                     empty child element ``<k/>``
``"..."``                 text content (JSON string)
``{r}id``                 namespaced name (``r``: relationships, ``x14``,
                          ``dc``, ...; or ``{uri}name``)
========================  =================================================

Items are separated by one space; attributes come first, then the text,
then the children in document order. A value is written as is unless it is
empty or contains white space, ``(``, ``)`` or ``"`` - then it is a JSON
string. Every element openpyxl serialises (``obj.to_tree()``) round-trips
exactly: ``decode(encode(el))`` gives the same tag, attributes, text and
children.
"""

import json
import re

from openrsxl.xml.constants import (
    COREPROPS_NS,
    CUSTPROPS_NS,
    DCORE_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX,
    DRAWING_NS,
    REL_NS,
    SHEET_DRAWING_NS,
    SHEET_MAIN_NS,
    VTYPES_NS,
    XSI_NS,
)
from openrsxl.xml.functions import Element

# prefix -> uri
PREFIXES = {
    "r": REL_NS,
    "cp": COREPROPS_NS,
    "dc": DCORE_NS,
    DCTERMS_PREFIX: DCTERMS_NS,
    "xsi": XSI_NS,
    "vt": VTYPES_NS,
    "op": CUSTPROPS_NS,
    "a": DRAWING_NS,
    "xdr": SHEET_DRAWING_NS,
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
URIS = {v: k for k, v in PREFIXES.items()}

_PLAIN = re.compile(r'[^\s()"]+\Z')
_NAME_END = " \t\r\n()=:"


def scalar(value):
    """A value: as is, or a JSON string"""
    value = str(value)
    if _PLAIN.match(value):
        return value
    return json.dumps(value, ensure_ascii=False)


def name(qname):
    """`{uri}local` -> `local` (main / no namespace), `{prefix}local`"""
    if qname[0] != "{":
        return qname
    uri, local = qname[1:].split("}", 1)
    if uri == SHEET_MAIN_NS:
        return local
    prefix = URIS.get(uri)
    if prefix is not None:
        return f"{{{prefix}}}{local}"
    return qname


def encode_items(el, skip_attrs=()):
    """The content of `el` (attributes, text, children) as CX items"""
    parts = []
    for k, v in el.attrib.items():
        if k in skip_attrs:
            continue
        parts.append(f"{name(k)}={scalar(v)}")
    if el.text is not None:
        parts.append(json.dumps(el.text, ensure_ascii=False))
    for child in el:
        parts.append(encode(child))
    return " ".join(parts)


def encode(el):
    """One element as a CX item"""
    tag = name(el.tag)
    has_text = el.text is not None
    if not el.attrib and len(el) == 0 and not has_text:
        return tag
    if len(el) == 0 and not has_text and list(el.attrib) == ["val"]:
        return f"{tag}:{scalar(el.attrib['val'])}"
    return f"{tag}({encode_items(el)})"


# -- decoding ------------------------------------------------------------------
class CXError(ValueError):
    pass


def _qname(n):
    if n[0] != "{":
        return n
    pre, local = n[1:].split("}", 1)
    return f"{{{PREFIXES.get(pre, pre)}}}{local}"


_decoder = json.JSONDecoder()


class _Parser:

    def __init__(self, text):
        self.s = text
        self.i = 0

    def ws(self):
        s, i = self.s, self.i
        while i < len(s) and s[i] in " \t\r\n":
            i += 1
        self.i = i

    def string(self):
        try:
            value, end = _decoder.raw_decode(self.s, self.i)
        except ValueError as e:
            raise CXError(f"bad string at {self.i}: {self.s[self.i:self.i + 30]!r}") from e
        if not isinstance(value, str):
            raise CXError(f"expected a string at {self.i}")
        self.i = end
        return value

    def scalar(self):
        if self.i < len(self.s) and self.s[self.i] == '"':
            return self.string()
        s, i = self.s, self.i
        j = i
        while j < len(s) and s[j] not in ' \t\r\n()"':
            j += 1
        if j == i:
            raise CXError(f"expected a value at {i}")
        self.i = j
        return s[i:j]

    def key(self):
        s, i = self.s, self.i
        j = i
        if j < len(s) and s[j] == "{":
            j = s.index("}", j) + 1
        while j < len(s) and s[j] not in _NAME_END:
            j += 1
        if j == i:
            raise CXError(f"expected a name at {i}: {s[i:i + 30]!r}")
        self.i = j
        return s[i:j]

    def items(self, el, until=None):
        """Parse items into `el` until `until` (')') or the end"""
        s = self.s
        while True:
            self.ws()
            if self.i >= len(s):
                if until is not None:
                    raise CXError("missing ')'")
                return el
            c = s[self.i]
            if c == until:
                self.i += 1
                return el
            if c == '"':
                el.text = self.string()
                continue
            k = self.key()
            nxt = s[self.i] if self.i < len(s) else ""
            if nxt == "=":
                self.i += 1
                el.set(_qname(k), self.scalar())
            elif nxt == ":":
                self.i += 1
                child = Element(_qname(k))
                child.set("val", self.scalar())
                el.append(child)
            elif nxt == "(":
                self.i += 1
                child = Element(_qname(k))
                self.items(child, ")")
                el.append(child)
            else:
                el.append(Element(_qname(k)))


def decode_items(tag, text):
    """Build element `tag` from CX items"""
    el = Element(_qname(tag))
    _Parser(text).items(el)
    return el


def decode(text):
    """One CX item (``name(...)``, ``name:v``, ``name``) -> element"""
    root = Element("_")
    _Parser(text).items(root)
    if len(root) != 1 or root.attrib or root.text is not None:
        raise CXError(f"expected one element: {text[:40]!r}")
    return root[0]


def split_items(text):
    """Split CX items at the top level (keeps their text)"""
    p = _Parser(text)
    out = []
    s = text
    while True:
        p.ws()
        if p.i >= len(s):
            return out
        start = p.i
        if s[p.i] == '"':
            p.string()
        else:
            p.key()
            nxt = s[p.i] if p.i < len(s) else ""
            if nxt in "=:":
                p.i += 1
                p.scalar()
            elif nxt == "(":
                p.i += 1
                p.items(Element("_"), ")")
        out.append(s[start : p.i])


# -- minimal elements ------------------------------------------------------------
_DEFAULTS: dict = {}
_XSI_TYPE = "{%s}type" % XSI_NS


def _default_attrs(cls):
    """Attributes of `cls()` serialised (None if cls() cannot be built)"""
    if cls not in _DEFAULTS:
        try:
            el = cls().to_tree()
            _DEFAULTS[cls] = dict(el.attrib) if el is not None else None
        except Exception:
            _DEFAULTS[cls] = None
    return _DEFAULTS[cls]


def _child_class(cls, tag):
    local = tag.rsplit("}", 1)[-1]
    desc = getattr(cls, local, None)
    t = getattr(desc, "expected_type", None)
    if isinstance(t, type) and hasattr(t, "from_tree"):
        return t
    return None


def _stripped(el, cls):
    """Copy of `el` without the attributes that `cls.from_tree` restores by
    itself (equal to the defaults, xmlns, xsi:type), recursively"""
    new = Element(el.tag)
    defaults = _default_attrs(cls) if cls is not None else None
    for k, v in el.attrib.items():
        if k in ("xmlns", _XSI_TYPE):
            continue
        if defaults is not None and defaults.get(k) == v:
            continue
        new.set(k, v)
    new.text = el.text
    for child in el:
        new.append(_stripped(child, _child_class(cls, child.tag) if cls is not None else None))
    return new


def minimal(obj, el=None, cls=None, reader=None):
    """The shortest element that `reader` (default ``cls.from_tree``) turns
    back into an object serialising exactly as `el` (default:
    ``obj.to_tree()``); verified, else `el` itself (see ``verifies``)"""
    if el is None:
        el = obj.to_tree()
    cls = cls or type(obj)
    for candidate in (_stripped(el, cls), el):
        for c in (candidate, _with_unset(candidate, obj, cls)):
            if c is not None and verifies(c, el, cls, reader):
                return c
    return el


def verifies(candidate, el, cls, reader=None):
    """Does reading `candidate` give an object serialising as `el`?"""
    from openrsxl.xml.functions import tostring

    try:
        obj = reader(cls, candidate) if reader is not None else cls.from_tree(candidate)
        back = obj.to_tree()
    except Exception:
        return False
    return back is not None and tostring(back) == tostring(el)


def _with_unset(el, obj, cls):
    """`el` plus an empty child for every element that is unset (None) in
    `obj` but not in ``cls()`` - from_tree reads an empty element as None
    (e.g. a workbook without creator: DocumentProperties defaults to one)"""
    try:
        default = cls()
    except Exception:
        return None
    names = [
        n
        for n in (getattr(cls, "__elements__", None) or ())
        if getattr(obj, n, None) is None and getattr(default, n, None) is not None
    ]
    if not names:
        return None
    new = Element(el.tag, dict(el.attrib))
    new.text = el.text
    for child in el:
        new.append(child)
    for n in names:
        new.append(Element(n))
    return new


def encode_min(obj, el=None, cls=None):
    """CX of the minimal element of `obj`"""
    return encode(minimal(obj, el, cls))
