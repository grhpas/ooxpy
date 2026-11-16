"""Source-agnostic intermediate representation (IR) of the OOXML vocabulary.

Every frontend (Open XML SDK metadata today, ECMA-376 XSD later) produces a
``Schema``; the emitter only ever sees the IR.  This keeps us decoupled from
the SDK's JSON layout, which is an internal artifact of the .NET project and
not a public contract.

Terminology
-----------
qname
    ``prefix:local`` using the *canonical* prefix from the namespace table,
    e.g. ``w:rPr``.
key
    ``TYPE/qname`` -- an element *bound to a content model*, e.g.
    ``w:CT_ParaRPr/w:rPr``.  191 qnames have more than one key (``w:rPr`` has
    seven), which is why the runtime resolves an element's key from its parent
    rather than from its qname alone.
rank
    A tuple used to keep children in schema order.  Siblings are ordered by
    rank; equal ranks keep insertion order (choices, repeating groups).
"""

from __future__ import annotations

from dataclasses import dataclass, field

UNBOUNDED = None


@dataclass(frozen=True, slots=True)
class Namespace:
    prefix: str
    uri: str
    version: str | None = None  # None == Office 2007 / ECMA-376 baseline


@dataclass(frozen=True, slots=True)
class EnumDef:
    key: str  # unique id, e.g. "w:ST_Jc"
    values: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class AttrDef:
    qname: str  # "w:val", or "val" for unqualified attributes
    kind: str  # converter kind: "string", "bool", "int", "enum", ...
    required: bool = False
    enum: str | None = None  # EnumDef.key
    version: str | None = None
    doc: str = ""
    # Lexical constraints (AND of items; an item is a constraint or a
    # ("union", (constraint, ...)) where any member may match).  Constraints:
    #   ("num", is_float, min_incl, max_incl, min_excl, max_excl)
    #   ("str", pattern, min_len, max_len, exact_len)
    #   ("enum", enum_key)
    lexical: tuple | None = None


OFFICE_VERSIONS = ("Office2007", "Office2010", "Office2013", "Office2016", "Office2019",
                   "Office2021", "Microsoft365")


@dataclass(frozen=True, slots=True)
class ChildRef:
    key: str  # TYPE/qname of the child
    qname: str
    rank: tuple[int, ...]
    max_one: bool  # at most one occurrence in this parent
    version: str | None = None
    required: bool = False  # must be present (min>=1, no optional/choice ancestor)


@dataclass(slots=True)
class TypeDef:
    key: str
    qname: str | None  # None for abstract base types
    class_name: str  # friendly SDK name, used for aliases (e.g. "Paragraph")
    summary: str = ""
    attrs: list[AttrDef] = field(default_factory=list)
    children: list[ChildRef] = field(default_factory=list)
    leaf_text: bool = False  # simple content: element holds text
    any_child: bool = False  # xsd:any present -- accept foreign children
    ordered: bool = True  # False if the content model (particle) is unknown
    abstract: bool = False
    part_root: bool = False
    version: str | None = None
    # the content model as a tree, for full validation (oox.contentmodel):
    # ("E", qname, min, max) | ("A", min, max) | ("S"|"C"|"L", min, max, (items...));
    # max None = unbounded.  None if unknown or overridden as unordered.
    particle: tuple | None = None


@dataclass(slots=True)
class Schema:
    source: str  # human-readable provenance, embedded in generated headers
    namespaces: list[Namespace]
    types: dict[str, TypeDef]
    enums: dict[str, EnumDef]
    parts: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
