"""Runtime view of the generated grammar tables.

Grammar modules are imported lazily, one per type prefix, so ``import oox``
stays cheap and a WordprocessingML-only program never loads SpreadsheetML.

The central idea is *context-dependent type resolution*: an element's content
model (its "key", e.g. ``w:CT_ParaRPr/w:rPr``) is decided by its parent's
content model, exactly as in XSD.  Most qnames have a single key; the 191 that
don't are resolved on demand through the parent chain.  Nothing is stored on
the element itself (lxml proxies are not stable, and the XML stays clean).
"""

from __future__ import annotations

import importlib
import keyword
from dataclasses import dataclass, field
from functools import cache

from ._generated.index import CANDIDATES
from ._generated.namespaces import NAMESPACES

FLAG_LEAF_TEXT, FLAG_ANY, FLAG_ORDERED, FLAG_ABSTRACT, FLAG_PART_ROOT = 1, 2, 4, 8, 16

URI: dict[str, str] = {p: u for p, u, _ in NAMESPACES}
PREFIX: dict[str, str] = {u: p for p, u, _ in NAMESPACES}
VERSION: dict[str, str | None] = {p: v for p, _, v in NAMESPACES}

# Factory keywords with special meaning (odfpy's Element(...) signature).  Two
# of them are also real OOXML attribute local names somewhere ("text",
# "parent"); those attributes stay reachable as ``prefix_local`` / "p:local".
OFFICE_VERSIONS = ("Office2007", "Office2010", "Office2013", "Office2016", "Office2019",
                   "Office2021", "Microsoft365")


def version_rank(v: str | None) -> int:
    """Office2007 (or None) -> 0, Office2010 -> 1, ...; unknown -> 0."""
    return OFFICE_VERSIONS.index(v) if v in OFFICE_VERSIONS else 0


RESERVED_KWARGS = frozenset({"text", "attributes", "qattributes", "parent", "check_grammar"})


def _loose(name: str) -> str:
    """odfpy's attribute spelling: lowercase, no dashes (``style-name`` -> ``stylename``)."""
    return name.lower().replace("-", "").replace("_", "")


@dataclass(frozen=True, slots=True)
class AttrSpec:
    qname: str
    kind: str
    required: bool
    enum: str | None
    version: str | None = None  # first Office version that knows it (None = 2007)
    lexical: tuple | None = None  # see oox.lexical

    @property
    def clark(self) -> str:
        return clark(self.qname)


@dataclass(frozen=True, slots=True)
class ChildSpec:
    key: str
    rank: tuple[int, ...]
    max_one: bool
    required: bool = False
    version: str | None = None  # Office version that added it to this parent


@dataclass(frozen=True, slots=True)
class TypeInfo:
    key: str
    qname: str | None
    flags: int
    attrs: dict[str, AttrSpec]  # by qname
    children: dict[str, ChildSpec]  # by child qname
    kwargs: dict[str, AttrSpec] = field(repr=False)  # python kwarg -> spec
    version: str | None = None

    leaf_text = property(lambda self: bool(self.flags & FLAG_LEAF_TEXT))
    any_child = property(lambda self: bool(self.flags & FLAG_ANY))
    ordered = property(lambda self: bool(self.flags & FLAG_ORDERED))

    def attr(self, name: str) -> AttrSpec | None:
        """Look up by kwarg (``val``/``w_val``), qname (``w:val``) or Clark name."""
        if name in self.kwargs:
            return self.kwargs[name]
        if name in self.attrs:
            return self.attrs[name]
        if name.startswith("{"):
            q = qname_of(name)
            return self.attrs.get(q) if q else None
        # odfpy-style fallback: case-insensitive, dash-less, if unambiguous
        hits = [s for s in self.attrs.values()
                if _loose(s.qname.rpartition(":")[2]) == _loose(name)]
        return hits[0] if len(hits) == 1 else None


def clark(qname: str) -> str:
    prefix, _, local = qname.rpartition(":")
    return f"{{{URI[prefix]}}}{local}" if prefix else local


_QNAMES: dict[str, str | None] = {}


def qname_of(tag: str) -> str | None:
    """``{uri}local`` -> ``prefix:local`` using canonical prefixes (memoized:
    a document has few distinct tags, looked up very often)."""
    if not isinstance(tag, str):  # comments / PIs have callable tags
        return None
    hit = _QNAMES.get(tag, _MISSING)
    if hit is not _MISSING:
        return hit  # type: ignore[return-value]
    if tag[:1] != "{":
        out: str | None = tag
    else:
        uri, _, local = tag[1:].partition("}")
        prefix = PREFIX.get(uri)
        out = f"{prefix}:{local}" if prefix is not None else None
    if len(_QNAMES) < 100000:
        _QNAMES[tag] = out
    return out


_MISSING = object()


def _kwarg_map(attrs: dict[str, AttrSpec], element_prefix: str = "") -> dict[str, AttrSpec]:
    locals_: dict[str, list[AttrSpec]] = {}
    for spec in attrs.values():
        locals_.setdefault(spec.qname.rpartition(":")[2], []).append(spec)
    out: dict[str, AttrSpec] = {}
    for local, specs in locals_.items():
        ident = local.replace("-", "_").replace(".", "_")
        for spec in specs:  # always: prefix_local
            prefix = spec.qname.rpartition(":")[0]
            if prefix:
                out[f"{prefix}_{ident}"] = spec
        # plain name: the only attribute with this local name, or the one
        # unqualified attribute among several (``id`` next to ``r:id``)
        unqualified = [s for s in specs if ":" not in s.qname]
        plain = specs[0] if len(specs) == 1 else (unqualified[0] if len(unqualified) == 1 else None)
        if plain is not None and ident in RESERVED_KWARGS and element_prefix and \
                ":" not in plain.qname:
            # an unqualified attribute named like an odfpy keyword (cfRule/@text,
            # x:sheetView/@parent ...): reachable as <element prefix>_<name>
            out.setdefault(f"{element_prefix}_{ident}", plain)
        if plain is None or not ident.isidentifier() or ident in RESERVED_KWARGS:
            continue
        out[ident + "_" if keyword.iskeyword(ident) else ident] = plain  # def -> def_
    return out


def _build(key: str, entry: tuple) -> TypeInfo:
    qname, flags, attrs, kids, version = entry
    amap = {a[0]: AttrSpec(*a) for a in attrs}
    cmap: dict[str, ChildSpec] = {}
    for ckey, rank, max_one, required, cversion in kids:
        cq = ckey.partition("/")[2]
        cmap.setdefault(cq, ChildSpec(ckey, tuple(rank), max_one, required, cversion))
    return TypeInfo(key, qname, flags, amap, cmap, _kwarg_map(amap, qname.partition(":")[0]),
                    version)


@cache
def _table(mod: str) -> dict:
    return importlib.import_module(f"oox._generated.g_{mod}").TYPES


@cache
def type_info(key: str) -> TypeInfo | None:
    mod = key.partition("/")[0].partition(":")[0]
    try:
        entry = _table(mod).get(key)
    except ModuleNotFoundError:
        return None
    return _build(key, entry) if entry else None


def candidates(qname: str) -> tuple[str, ...]:
    return CANDIDATES.get(qname, ())


@cache
def child_key(parent_key: str, qname: str) -> str | None:
    info = type_info(parent_key)
    if info is None:
        return None
    spec = info.children.get(qname)
    return spec.key if spec else None


@cache
def detached_info(qname: str) -> TypeInfo | None:
    """Content model for an element that has no parent yet.

    Unique qname -> its model.  Ambiguous qname -> the *union* of all models
    (permissive), with an attribute required only if every model requires it.
    The precise model is checked again when the element is attached.
    """
    keys = candidates(qname)
    if not keys:
        return None
    if len(keys) == 1:
        return type_info(keys[0])
    infos = [i for i in (type_info(k) for k in keys) if i]
    attrs: dict[str, AttrSpec] = {}
    for i in infos:
        for q, a in i.attrs.items():
            prev = attrs.setdefault(q, a)
            if (prev.kind, prev.enum, prev.lexical) == (a.kind, a.enum, a.lexical):
                continue
            if (prev.kind, prev.enum) == (a.kind, a.enum):  # only value shapes differ
                attrs[q] = AttrSpec(q, a.kind, prev.required, a.enum, prev.version,
                                    ("union", tuple(x for x in (prev.lexical, a.lexical) if x))
                                    if prev.lexical and a.lexical else None)
                continue
            if prev.kind == a.kind == "enum":
                # w:jc/@w:val is ST_Jc or ST_JcTable: accept either set now
                keys = sorted(set((prev.enum or "").split("|")) | {a.enum or ""})
                attrs[q] = AttrSpec(q, "enum", prev.required, "|".join(k for k in keys if k))
            else:
                # x:c/@t is bool in calcChain, enum in a worksheet: accept the
                # lexical value now; the precise model re-checks on attach.
                attrs[q] = AttrSpec(q, "string", prev.required, None)
    for q, a in list(attrs.items()):
        if a.required and not all(q in i.attrs and i.attrs[q].required for i in infos):
            attrs[q] = AttrSpec(a.qname, a.kind, False, a.enum)
    children: dict[str, ChildSpec] = {}
    for i in infos:
        for q, c in i.children.items():
            children.setdefault(q, ChildSpec(c.key, (), False))
    flags = 0
    for i in infos:
        flags |= i.flags & (FLAG_LEAF_TEXT | FLAG_ANY)
    return TypeInfo(f"<union {qname}>", qname, flags, attrs, children,
                    _kwarg_map(attrs, qname.partition(":")[0]))


# Values ISO/IEC 29500 Transitional allows but the SDK's enumerations omit,
# accepted where Office has been seen to accept them.  Evidence: LibreOffice
# writes them (tools/interop.py) and Word 2007 opens them without a prompt,
# lists and tables intact (observe case word-94, 2026-11-15).
EXTRA_ENUM_VALUES: dict[str, tuple[str, ...]] = {
    "w:ST_LevelJustification": ("start", "end"),      # ST_Jc in ISO Transitional
    "w:ST_TblJc": ("start", "end"),                   # ST_JcTable
    "w:ST_OnOffOnly": ("true", "false", "1", "0"),    # ST_OnOff (cantSplit, tblHeader)
}


def enum_values(enum_key: str) -> tuple[str, ...]:
    """Allowed values; ``"a|b"`` (from a union of models) merges both sets."""
    from ._generated.enums import ENUMS

    if "|" not in enum_key:
        return ENUMS.get(enum_key, ()) + EXTRA_ENUM_VALUES.get(enum_key, ())
    out: dict[str, None] = {}
    for k in enum_key.split("|"):
        out.update(dict.fromkeys(ENUMS.get(k, ()) + EXTRA_ENUM_VALUES.get(k, ())))
    return tuple(out)
