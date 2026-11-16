"""Schema-aware lxml element -- the odfpy ``Element`` equivalent.

Design choices (see docs/ARCHITECTURE.md):

* **lxml-backed**, not a hand-rolled node tree like odfpy's.  Parsing, XPath,
  C-speed serialization and exact round-tripping of unknown markup come free.
  One element class (``OoxElement``) is registered for *every* tag; schema
  knowledge comes from the generated tables, so there are no 4,000 classes.
* **Schema-ordered insertion.**  Word rejects ``<w:pPr>`` children out of
  sequence order, so ``add()`` inserts each child at its schema position
  rather than appending.
* **Validation at the edges** (construction / ``add``), switchable with
  ``lenient()`` for repairing messy real-world files.  Loading never validates.
"""

from __future__ import annotations

import contextlib
import contextvars
import difflib
from collections.abc import Iterator
from typing import Any

from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

from . import attrconverters, grammar
from ._odfdom import OdfDomMixin
from .errors import DuplicateChild, InvalidAttribute, InvalidChild, MissingAttribute, PackageError

_STRICT: contextvars.ContextVar[bool] = contextvars.ContextVar("oox_strict", default=True)
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


@contextlib.contextmanager
def lenient() -> Iterator[None]:
    """Turn construction-time checks off inside the block: unknown attributes
    and children, children out of place and invalid values are all accepted
    (``validate()`` still reports them).  For repairing messy input, or for
    markup Office reads although the schema doesn't allow it."""
    token = _STRICT.set(False)
    try:
        yield
    finally:
        _STRICT.reset(token)


def strict() -> bool:
    return _STRICT.get()


_TRUSTED: contextvars.ContextVar[bool] = contextvars.ContextVar("oox_trusted", default=False)


@contextlib.contextmanager
def trusted() -> Iterator[None]:
    """Inside oox's own builders: skip the attach-time re-check of ambiguous
    elements (``x:row``, ``x:c`` ...), whose structure the builder guarantees.
    Everything else (unknown children and attributes, values) is still checked."""
    token = _TRUSTED.set(True)
    try:
        yield
    finally:
        _TRUSTED.reset(token)


# Office 2010 renames of Office 2007 markup (bidi-neutral names): strip_for
# writes the old name instead of dropping the information.
_OLD_NAMES = {"w:start": "w:left", "w:end": "w:right",
              "w:startChars": "w:leftChars", "w:endChars": "w:rightChars"}


class OoxElement(OdfDomMixin, etree.ElementBase):
    """An OOXML element.  Create via factories: ``from oox.ns import w; w.P()``.

    Also speaks odfpy's DOM dialect (``childNodes``, ``appendChild``,
    ``getElementsByType`` ...) -- see ``oox._odfdom``.
    """

    # -- identity -------------------------------------------------------------
    @property
    def tag_name(self) -> str | None:
        """Prefixed name with canonical prefix, e.g. ``"w:p"`` (odfpy: ``tagName``)."""
        return grammar.qname_of(self.tag)

    @property
    def qname(self) -> tuple[str | None, str]:
        """``(namespace URI, local name)`` -- same shape as odfpy's ``Element.qname``."""
        q = etree.QName(self.tag)
        return (q.namespace, q.localname)

    @property
    def schema_key(self) -> str | None:
        """``TYPE/qname`` of the content model in effect, resolved from context."""
        q = self.tag_name
        if q is None:
            return None
        keys = grammar.candidates(q)
        if len(keys) == 1:
            return keys[0]
        if not keys:
            return None
        parent = self.getparent()
        if parent is None or not isinstance(parent, OoxElement):
            return None
        pkey = parent.schema_key
        return grammar.child_key(pkey, q) if pkey else None

    @property
    def info(self) -> grammar.TypeInfo | None:
        """The schema type of this element (``grammar.TypeInfo``: attributes, children,
        order), or None for unknown elements."""
        key = self.schema_key
        if key is not None:
            return grammar.type_info(key)
        return grammar.detached_info(self.tag_name) if self.tag_name else None

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {self.tag_name or self.tag} [{self.schema_key or '?'}] at {id(self):#x}>"

    # -- attributes -----------------------------------------------------------
    def set_attr(self, name: str, value: object) -> OoxElement:
        """Set by kwarg name (``val``), qname (``w:val``) or Clark name.

        ``None`` removes the attribute.  Values are converted/validated by kind.
        """
        info = self.info
        spec = info.attr(name) if info else None
        if spec is None:
            if strict() and info is not None:
                allowed = sorted(info.kwargs)
                hint = difflib.get_close_matches(name, allowed, n=3)
                raise InvalidAttribute(
                    f"{self.tag_name} has no attribute {name!r}"
                    + (f"; did you mean {', '.join(hint)}?" if hint else
                       f"; allowed: {', '.join(allowed) or '(none)'}"))
            # unknown: accept qname / Clark / plain names verbatim
            clark = name if name.startswith("{") else (
                grammar.clark(name) if ":" in name else name)
            if value is None:
                self.attrib.pop(clark, None)
            else:
                self.set(clark, str(value))
            return self
        if value is None:
            self.attrib.pop(spec.clark, None)
        else:
            self.set(spec.clark, attrconverters.to_xml_lenient(spec, value))
        return self

    def get_attr(self, name: str, default: str | None = None) -> str | None:
        """The value of attribute ``name`` (``"val"``, ``"w:val"`` or a Clark name), or
        ``default``."""
        info = self.info
        spec = info.attr(name) if info else None
        clark = spec.clark if spec else (grammar.clark(name) if ":" in name else name)
        return self.get(clark, default)

    # odfpy-compatible spellings
    def setAttribute(self, name: str, value: object) -> None:
        """odfpy's ``setAttribute``: same as ``set_attr``."""
        self.set_attr(name, value)

    def getAttribute(self, name: str) -> str | None:
        """odfpy's ``getAttribute``: the attribute's value, or None."""
        return self.get_attr(name)

    # -- children -------------------------------------------------------------
    def add(self, *items: OoxElement | str, replace: bool = False) -> OoxElement:
        """Add children in schema order; strings are added as text.

        ``replace=True`` swaps out an existing child with the same tag (e.g. a
        second ``w:pPr``) instead of raising ``DuplicateChild``.  It also
        applies where the schema allows repeats (``w:trPr`` is an unbounded
        choice): the first child with that tag is replaced.
        Returns ``self`` so calls can be chained.
        """
        for item in items:
            if isinstance(item, str):
                self.add_text(item)
                continue
            if not isinstance(item, etree._Element):
                raise TypeError(f"cannot add {type(item).__name__} to {self.tag_name}")
            self._add_child(item, replace)
        return self

    def _add_child(self, item: etree._Element, replace: bool) -> None:
        info = self.info
        q = grammar.qname_of(item.tag)
        spec = info.children.get(q) if (info and q) else None
        if spec is None:
            if info is not None and strict() and not info.any_child:
                allowed = ", ".join(sorted(info.children)) or "(no element children)"
                raise InvalidChild(f"<{q or item.tag}> is not allowed in <{self.tag_name}>; "
                                   f"allowed: {allowed}")
            self.append(item)
            return
        if spec.max_one or replace:
            existing = next((c for c in self.iterchildren(item.tag)), None)
            if existing is not None and existing is not item:
                if replace:
                    self.replace(existing, item)
                    self._check_attached(item, q)
                    return
                if strict():
                    raise DuplicateChild(
                        f"<{self.tag_name}> allows one <{q}>; pass replace=True to swap it")
        assert info is not None  # spec came from info
        where = self._position_for(info, spec.rank) if info.ordered and spec.rank else None
        if where is None:
            self.append(item)
        elif where is self:
            self.insert(0, item)
        else:
            where.addnext(item)
        self._check_attached(item, q)

    def _position_for(self, info: grammar.TypeInfo, rank: tuple[int, ...]):
        """Where a child of ``rank`` goes: ``None`` = append, ``self`` = first,
        else the sibling to insert after (the last one with rank <= ``rank``;
        unknown children such as ``mc:AlternateContent`` stay after their
        known predecessor).  Scans from the end: appending in order, or just
        before a trailing ``w:sectPr``, costs one or two steps (lxml's
        ``len()`` and indexing walk every child)."""
        first = True
        for c in self.iterchildren(reversed=True):
            cs = info.children.get(grammar.qname_of(c.tag) or "")
            if cs is not None and cs.rank <= rank:
                return None if first else c
            first = False
        return None if first else self

    def _check_attached(self, item: etree._Element, q: str | None) -> None:
        """Ambiguous elements were checked against a union; re-check precisely."""
        if not strict() or q is None or _TRUSTED.get() or len(grammar.candidates(q)) < 2:
            return
        if isinstance(item, OoxElement):
            item._reorder()  # children were added before the model was known
            problems = item.validate(recursive=True)
            if problems:
                item.getparent().remove(item)
                raise InvalidChild(f"<{q}> in <{self.tag_name}>: {problems[0]}")

    def _reorder(self) -> None:
        """Stable-sort children into schema order (recursively).

        Unknown children (e.g. ``mc:AlternateContent``) travel with the
        preceding known sibling.
        """
        info = self.info
        if info is not None and info.ordered and len(self) > 1:
            keyed, rank = [], ()
            for i, c in enumerate(self):
                cs = info.children.get(grammar.qname_of(c.tag) or "")
                if cs is not None and cs.rank:
                    rank = cs.rank
                keyed.append((rank, i, c))
            if any(keyed[i][0] > keyed[i + 1][0] for i in range(len(keyed) - 1)):
                for _, _, c in sorted(keyed, key=lambda t: (t[0], t[1])):
                    self.append(c)  # append moves the node
        for c in self:
            if isinstance(c, OoxElement):
                c._reorder()

    def add_text(self, text: str) -> OoxElement:
        """Append ``text`` to this element (after its last child, if any); refused for
        elements whose schema type holds no text.  Returns the element."""
        info = self.info
        if strict() and info is not None and not info.leaf_text:
            raise InvalidChild(f"<{self.tag_name}> cannot contain text")
        if len(self):
            last = self[-1]
            last.tail = (last.tail or "") + text
        else:
            self.text = (self.text or "") + text
        full = self.text or ""
        if (info and "xml:space" in info.attrs
                and (full != full.strip() or "  " in full or "\n" in full or "\t" in full)):
            self.set(XML_SPACE, "preserve")  # Word drops edge whitespace otherwise
        return self

    # odfpy-compatible spellings
    def addElement(self, element: OoxElement, check_grammar: bool = True) -> None:
        """odfpy's ``addElement``: ``add`` in schema order; ``check_grammar=False`` skips
        validation, as in odfpy."""
        if check_grammar:
            self.add(element)
        else:
            with lenient():
                self.add(element)

    def addText(self, text: str) -> None:
        """odfpy's ``addText``: append text (see ``add_text``)."""
        self.add_text(text)

    def first(self, qname: str) -> OoxElement | None:
        """First direct child with the given qname (``"w:pPr"``)."""
        return next(self.iterchildren(grammar.clark(qname)), None)

    def need(self, qname: str) -> OoxElement:
        """Like ``first``, but a missing child raises ``PackageError`` naming it.

        Use where the structure is required: files from other producers can
        be well-formed yet lack it, and a clear error beats ``'NoneType' has
        no attribute ...``.
        """
        el = self.first(qname)
        if el is None:
            raise PackageError(f"<{self.tag_name}> has no <{qname}> (unexpected document structure)")
        return el

    # -- validation -----------------------------------------------------------
    def validate(self, recursive: bool = True, target: str | None = None,
                 level: str = "basic") -> list[str]:
        """Return a list of problems (empty == valid as far as we model it).

        Checks: unknown/missing attributes, attribute values (enums, numbers,
        lengths, patterns), allowed children, singleton children, sequence
        order, stray text, and mandatory children (those required regardless
        of choices).  ``level="full"`` also matches each element's children
        against its whole content model (``oox.contentmodel``): exclusive
        choices, required alternatives, occurrence counts, unfinished groups.

        ``target="Office2007"`` (or Office2010, ...) also reports markup that
        version would not understand: elements and attributes added later to
        namespaces it knows.  Newer *namespaces* (w14, ...) are fine -- they
        are declared ``mc:Ignorable`` -- as is anything inside ``extLst`` or
        ``mc:AlternateContent``.
        """
        if level not in ("basic", "full"):
            raise ValueError("level must be 'basic' or 'full'")
        problems: list[str] = []
        self._validate_into(problems, recursive)
        if level == "full":
            content: list[str] = []
            self._content_into(content, recursive)
            basic = set(problems)
            for m in content:   # "incomplete, missing <x>" repeats a basic finding
                tag, _, rest = m.partition(": incomplete, missing ")
                if rest and f"{tag}: missing required child {rest}" in basic:
                    continue
                problems.append(m)
        if target is not None:
            if target not in grammar.OFFICE_VERSIONS:
                raise ValueError(f"target must be one of {grammar.OFFICE_VERSIONS}")
            self._target_into(problems, grammar.version_rank(target), target)
        return problems

    def _target_into(self, out: list[str], rank: int, target: str) -> None:
        out.extend(msg for _, _, msg in self.target_findings(target))

    def _content_into(self, out: list[str], recursive: bool) -> None:
        from . import contentmodel

        for el in (self.iter() if recursive else [self]):
            if not isinstance(el, OoxElement):
                continue
            key = el.schema_key
            if key is None:
                continue
            kids = [q for c in el if isinstance(c, etree._Element)
                    and (q := grammar.qname_of(c.tag)) is not None]
            msg = contentmodel.check(key, el.tag_name or "?", kids)
            if msg:
                out.append(msg)

    def target_findings(self, target: str) -> list[tuple[OoxElement, str | None, str]]:
        """Markup ``target`` (e.g. "Office2007") won't understand, as
        ``(element, attribute Clark name or None, message)``: the attribute,
        or (``None``) the whole element, was added to its namespace later."""
        if target not in grammar.OFFICE_VERSIONS:
            raise ValueError(f"target must be one of {grammar.OFFICE_VERSIONS}")
        rank = grammar.version_rank(target)
        found: list[tuple[OoxElement, str | None, str]] = []

        def ns_rank(qname: str | None) -> int:
            prefix = (qname or "").rpartition(":")[0]
            return grammar.version_rank(grammar.VERSION.get(prefix))

        def walk(el: OoxElement) -> None:
            tag = el.tag_name
            if tag is None or tag.startswith("mc:") or tag.rpartition(":")[2] == "extLst":
                return  # unknown, AlternateContent, extension lists: own compat rules
            if ns_rank(tag) > rank:
                return  # whole subtree in a newer, ignorable namespace
            info = el.info
            if info is not None:
                for clark in el.attrib:
                    q = grammar.qname_of(clark)
                    spec = info.attrs.get(q) if q else None
                    if spec is None or ns_rank(q) > rank:
                        continue
                    if grammar.version_rank(spec.version) > rank:
                        found.append((el, clark, f"<{tag}>: attribute {q} needs {spec.version} "
                                                 f"(not understood by {target})"))
                for c in el:
                    if not isinstance(c, OoxElement):
                        continue
                    cs = info.children.get(c.tag_name or "")
                    if cs is not None and ns_rank(c.tag_name) <= rank and \
                            grammar.version_rank(cs.version) > rank:
                        found.append((c, None, f"<{tag}>: child <{c.tag_name}> needs "
                                               f"{cs.version} (not understood by {target})"))
            for c in el:
                if isinstance(c, OoxElement):
                    walk(c)

        walk(self)
        return found

    def strip_for(self, target: str) -> list[str]:
        """Make the tree readable by ``target`` (see ``target_findings``):
        markup that was only renamed later goes back to its old name (Word
        2010's ``w:start``/``w:end`` -> ``w:left``/``w:right``), anything else
        is removed.  Returns a message per change."""
        changes = []
        for el, clark, msg in self.target_findings(target):
            q = grammar.qname_of(clark) if clark else el.tag_name
            old = _OLD_NAMES.get(q or "")
            if clark is not None:
                value = el.attrib.pop(clark, None)
                if old and value is not None and grammar.clark(old) not in el.attrib:
                    el.set(grammar.clark(old), value)
                    msg += f"; renamed to {old}"
            elif el.getparent() is not None:
                parent = el.getparent()
                if old and parent.find(grammar.clark(old)) is None:
                    el.tag = grammar.clark(old)
                    msg += f"; renamed to <{old}>"
                else:
                    parent.remove(el)
            changes.append(msg)
        return changes

    def _validate_into(self, out: list[str], recursive: bool) -> None:
        info = self.info
        where = self.tag_name or self.tag
        if info is not None:
            present = set()
            for clark, value in self.attrib.items():
                q = grammar.qname_of(clark)
                spec = info.attrs.get(q) if q else None
                if spec is None:
                    if not clark.startswith("{http://www.w3.org/2000/xmlns/") and \
                            not (q or "").startswith(("mc:", "xml:")):
                        out.append(f"<{where}>: unexpected attribute {q or clark}")
                    continue
                present.add(spec.qname)
                try:
                    attrconverters.to_xml(spec, value)
                except Exception as e:
                    out.append(f"<{where}>: {e}")
            for spec in info.attrs.values():
                if spec.required and spec.qname not in present:
                    out.append(f"<{where}>: missing required attribute {spec.qname}")
            if (self.text or "").strip() and not info.leaf_text:
                out.append(f"<{where}>: unexpected text")
            last_rank: tuple[int, ...] = ()
            seen_single: set[str] = set()
            for c in self:
                cq = grammar.qname_of(c.tag)
                if cq is None or cq.startswith("mc:"):
                    continue
                cs = info.children.get(cq)
                if cs is None:
                    if not info.any_child:
                        out.append(f"<{where}>: child <{cq}> not allowed")
                    continue
                if cs.max_one:
                    if cq in seen_single:
                        out.append(f"<{where}>: more than one <{cq}>")
                    seen_single.add(cq)
                if info.ordered and cs.rank and cs.rank < last_rank:
                    out.append(f"<{where}>: <{cq}> is out of schema order")
                last_rank = max(last_rank, cs.rank)
            present_kids = {grammar.qname_of(c.tag) for c in self}
            for cq, cs in info.children.items():
                if cs.required and cq not in present_kids:
                    out.append(f"<{where}>: missing required child <{cq}>")
            from .semantics import RULES

            rule = RULES.get(info.key)
            if rule is not None:
                out.extend(rule(self))
        if recursive:
            for c in self:
                if isinstance(c, OoxElement):
                    c._validate_into(out, True)

    # -- serialization --------------------------------------------------------
    def to_xml(self, pretty: bool = False) -> str:
        """This element and its subtree as an XML string."""
        return etree.tostring(self, encoding="unicode", pretty_print=pretty)


# One parser for everything: our element class, and safe defaults for
# untrusted input (no entity expansion, no network).
_lookup = etree.ElementDefaultClassLookup(element=OoxElement)
PARSER = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
PARSER.set_element_class_lookup(_lookup)


def parse_bytes(data: bytes) -> OoxElement:
    return etree.fromstring(data, PARSER)


def make(qname: str, key: str | None, children: tuple, text: str | None,
         attrs: dict[str, object]) -> OoxElement:
    """Factory backend used by every generated ``oox.ns.*`` function.

    Accepts odfpy's ``Element(...)`` keywords as well: ``attributes={...}``,
    ``qattributes={(uri, local): value}``, ``parent=element`` and
    ``check_grammar=False``.
    """
    attributes: Any = attrs.pop("attributes", None) or {}
    qattributes: Any = attrs.pop("qattributes", None) or {}
    parent: Any = attrs.pop("parent", None)
    check = attrs.pop("check_grammar", True)
    merged = {**attributes, **{f"{{{u}}}{loc}" if u else loc: v
                               for (u, loc), v in qattributes.items()}, **attrs}
    if check:
        el = _make(qname, key, children, text, merged)
    else:
        with lenient():
            el = _make(qname, key, children, text, merged)
    if parent is not None:
        parent.addElement(el, check_grammar=check)
    return el


def _make(qname: str, key: str | None, children: tuple, text: str | None,
          attrs: dict[str, object]) -> OoxElement:
    prefix, _, local = qname.rpartition(":")
    uri = grammar.URI[prefix]
    el: OoxElement = PARSER.makeelement(f"{{{uri}}}{local}", nsmap={prefix: uri})
    info = grammar.type_info(key) if key else grammar.detached_info(qname)
    for name, value in attrs.items():
        spec = info.attr(name) if info else None
        if spec is None:
            el.set_attr(name, value)  # raises with suggestions in strict mode
        elif value is not None:
            el.set(spec.clark, attrconverters.to_xml_lenient(spec, value))
    if strict() and info is not None:
        missing = [s.qname for s in info.attrs.values()
                   if s.required and el.get(s.clark) is None]
        if missing:
            raise MissingAttribute(f"<{qname}> requires {', '.join(missing)}")
    if children:
        el.add(*children)
    if text is not None:
        el.add_text(text)
    return el
