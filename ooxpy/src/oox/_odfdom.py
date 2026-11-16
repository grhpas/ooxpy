"""odfpy-compatible DOM surface for ``OoxElement`` (mixed in; no imports of element.py).

odfpy's ``Element`` is a minidom-like node.  This mixin gives ``OoxElement`` the
same method names so odfpy-style code ports mechanically.  Two semantics are
kept from odfpy on purpose:

* ``addElement`` checks the grammar (and, in oox, inserts in schema order);
  ``appendChild`` / ``insertBefore`` are raw DOM operations that do neither.
* ``childNodes`` includes a text node when the element has text, so the usual
  ``if node.nodeType == node.TEXT_NODE: node.data`` loops keep working.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import IO, TYPE_CHECKING, Any


class TextNode:
    """Read/write view of an element's leading text (odfpy ``element.Text``)."""

    ELEMENT_NODE, TEXT_NODE = 1, 3
    nodeType = TEXT_NODE

    def __init__(self, parent) -> None:
        self.parentNode = parent

    @property
    def data(self) -> str:
        return self.parentNode.text or ""

    @data.setter
    def data(self, value: str) -> None:
        self.parentNode.text = value

    def __str__(self) -> str:
        return self.data

    def __eq__(self, other: object) -> bool:
        return isinstance(other, TextNode) and other.parentNode is self.parentNode

    __hash__ = None  # type: ignore[assignment]


class OdfDomMixin:
    if TYPE_CHECKING:  # provided by the host class (OoxElement / lxml ElementBase)
        def __getattr__(self, name: str) -> Any: ...
        def __iter__(self) -> Any: ...
        def __len__(self) -> int: ...

    ELEMENT_NODE, TEXT_NODE = 1, 3
    nodeType = ELEMENT_NODE

    # -- naming -----------------------------------------------------------------
    @property
    def tagName(self) -> str | None:
        return self.tag_name

    # -- navigation -------------------------------------------------------------
    @property
    def childNodes(self) -> list:
        nodes: list = [TextNode(self)] if self.text else []
        nodes.extend(self)
        return nodes

    @property
    def parentNode(self):
        return self.getparent()

    @property
    def firstChild(self):
        nodes = self.childNodes
        return nodes[0] if nodes else None

    @property
    def lastChild(self):
        nodes = self.childNodes
        return nodes[-1] if nodes else None

    @property
    def nextSibling(self):
        return self.getnext()

    @property
    def previousSibling(self):
        return self.getprevious()

    def hasChildNodes(self) -> bool:
        return bool(self.text) or len(self) > 0

    # -- raw DOM mutation (no grammar, no ordering -- as in odfpy) -------------
    def appendChild(self, new):
        self.append(new)
        return new

    def insertBefore(self, new, ref):
        if ref is None:
            self.append(new)
        else:
            ref.addprevious(new)
        return new

    def removeChild(self, old):
        self.remove(old)
        return old

    def replaceChild(self, new, old):
        self.replace(old, new)
        return old

    # -- attributes -------------------------------------------------------------
    def removeAttribute(self, attr: str, check_grammar: bool = True) -> None:
        self.set_attr(attr, None)

    def setAttrNS(self, namespace: str, localpart: str, value: object) -> None:
        self.set_attr(f"{{{namespace}}}{localpart}" if namespace else localpart, value)

    def getAttrNS(self, namespace: str, localpart: str) -> str | None:
        return self.get(f"{{{namespace}}}{localpart}" if namespace else localpart)

    def removeAttrNS(self, namespace: str, localpart: str) -> None:
        self.attrib.pop(f"{{{namespace}}}{localpart}" if namespace else localpart, None)

    def allowed_attributes(self) -> list[tuple[str | None, str]] | None:
        from . import grammar

        info = self.info
        if info is None:
            return None
        out = []
        for spec in info.attrs.values():
            prefix, _, local = spec.qname.rpartition(":")
            out.append((grammar.URI[prefix] if prefix else None, local))
        return out

    # -- queries ----------------------------------------------------------------
    def getElementsByType(self, element: Callable) -> list:
        """``element`` is a factory function, e.g. ``getElementsByType(w.P)``."""
        tag = element(check_grammar=False).tag
        return list(self.iter(tag))

    def isInstanceOf(self, element: Callable) -> bool:
        return self.tag == element(check_grammar=False).tag

    # -- output -----------------------------------------------------------------
    def toXml(self, level: int = 0, f: IO[str] | None = None) -> str:
        xml = self.to_xml()
        if f is not None:
            f.write(xml)
        return xml
