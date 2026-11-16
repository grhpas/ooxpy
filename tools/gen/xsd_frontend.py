"""Frontend: ECMA-376 / ISO/IEC 29500 XSD files -> IR (Phase 3).

Used as an oracle today: ``tools.gen.irdiff`` compares its IR with the SDK
frontend's and reports every deviation.  It is not wired into the emitter, so
the generated code stays the SDK's.

    from tools.gen import xsd_frontend
    schema = xsd_frontend.load("path/to/schemas", prefixes={uri: prefix, ...})

What it handles (everything the Transitional schemas and Microsoft's Word
extension schemas use): named and anonymous complex types; sequence, choice,
all, group references (inlined as a sequence with the reference's
occurrence), element references across namespaces, ``xsd:any``;
complexContent extension (base particle, then the extension's, attributes
merged) and restriction (own content); simpleContent (text plus
attributes); attribute groups; simple types with enumerations, restriction
chains, unions and lists; elementFormDefault/attributeFormDefault.

Keys follow the SDK's convention, ``TYPE/qname`` with canonical prefixes
(``w:CT_RPr/w:rPr``); elements of a simple type get ``xsd:string/c:f``-style
keys like the SDK's.  Anonymous complex types are named
``<prefix>:anon_<element>``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

from .ir import AttrDef, ChildRef, EnumDef, Namespace, Schema, TypeDef

XS = "http://www.w3.org/2001/XMLSchema"
XML_NS = "http://www.w3.org/XML/1998/namespace"
_Q = f"{{{XS}}}"

_BUILTIN_KIND = {
    "string": "string", "token": "string", "normalizedString": "string", "NCName": "string",
    "Name": "string", "QName": "string", "ID": "string", "IDREF": "string", "anyURI": "string",
    "language": "string", "NMTOKEN": "string", "NMTOKENS": "list",
    "boolean": "bool",
    "int": "int", "integer": "int", "long": "int", "short": "int", "byte": "int",
    "unsignedInt": "int", "unsignedLong": "int", "unsignedShort": "int", "unsignedByte": "int",
    "nonNegativeInteger": "int", "positiveInteger": "int", "negativeInteger": "int",
    "nonPositiveInteger": "int",
    "double": "float", "float": "float", "decimal": "float",
    "dateTime": "datetime", "date": "datetime", "time": "datetime",
    "hexBinary": "hex", "base64Binary": "base64",
}


class _Doc:
    """One parsed schema file."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.root = etree.parse(str(path)).getroot()
        self.tns = self.root.get("targetNamespace") or ""
        self.elem_qualified = self.root.get("elementFormDefault") == "qualified"
        self.attr_qualified = self.root.get("attributeFormDefault") == "qualified"

    def resolve(self, node, qname: str) -> tuple[str, str]:
        """``prefix:local`` in a schema attribute -> (namespace URI, local)."""
        prefix, _, local = qname.rpartition(":")
        uri = node.nsmap.get(prefix or None)
        if uri is None:
            uri = XML_NS if prefix == "xml" else (self.tns if not prefix else "")
        return uri, local


def _occurs(node) -> tuple[int, int | None]:
    mn = int(node.get("minOccurs", "1"))
    mx_text = node.get("maxOccurs", "1")
    return mn, None if mx_text == "unbounded" else int(mx_text)


def load(schema_dir: str | Path, prefixes: dict[str, str] | None = None,
         source: str | None = None) -> Schema:
    """Every ``*.xsd`` under ``schema_dir``.  ``prefixes`` maps namespace URIs to
    canonical prefixes (default: oox's namespace table); namespaces without
    one get a ``ns<n>`` prefix and a warning."""
    if prefixes is None:
        from oox.grammar import URI

        prefixes = {}
        for p, u in URI.items():
            prefixes.setdefault(u, p)
    prefixes = dict(prefixes)
    warnings: list[str] = []
    docs = [_Doc(p) for p in sorted(Path(schema_dir).rglob("*.xsd"))]

    # global definitions, by (uri, local)
    ctypes: dict[tuple[str, str], tuple[_Doc, Any]] = {}
    stypes: dict[tuple[str, str], tuple[_Doc, Any]] = {}
    groups: dict[tuple[str, str], tuple[_Doc, Any]] = {}
    agroups: dict[tuple[str, str], tuple[_Doc, Any]] = {}
    gelements: dict[tuple[str, str], tuple[_Doc, Any]] = {}
    gattrs: dict[tuple[str, str], tuple[_Doc, Any]] = {}
    tables = {"complexType": ctypes, "simpleType": stypes, "group": groups,
              "attributeGroup": agroups, "element": gelements, "attribute": gattrs}
    for d in docs:
        for node in d.root:
            if not isinstance(node.tag, str) or not node.tag.startswith(_Q):
                continue
            table = tables.get(node.tag[len(_Q):])
            if table is not None and node.get("name"):
                key = (d.tns, node.get("name"))
                if key in table:
                    warnings.append(f"duplicate definition {key} in {d.path.name}")
                table[key] = (d, node)

    declared: dict[str, str] = {}          # prefixes the schema files themselves use
    for d in docs:
        for p, u in d.root.nsmap.items():
            if p and u not in declared:
                declared[u] = p

    def prefix_of(uri: str) -> str:
        if uri == XML_NS:
            return "xml"
        if uri not in prefixes:
            taken = set(prefixes.values())
            p = declared.get(uri)
            prefixes[uri] = p if p and p not in taken else f"ns{len(prefixes)}"
            warnings.append(f"no canonical prefix for {uri}; using {prefixes[uri]!r}")
        return prefixes[uri]

    def qn(uri: str, local: str) -> str:
        return f"{prefix_of(uri)}:{local}" if uri else local

    # --- simple types --------------------------------------------------------------
    enums: dict[str, EnumDef] = {}
    st_cache: dict[tuple[str, str], tuple[str, str | None]] = {}

    def simple(d: _Doc, node, name_hint: str) -> tuple[str, str | None]:
        """(kind, enum key) of a simpleType node."""
        for child in node:
            if not isinstance(child.tag, str):
                continue
            tag = child.tag[len(_Q):] if child.tag.startswith(_Q) else ""
            if tag == "restriction":
                values = [e.get("value") for e in child if e.tag == _Q + "enumeration"]
                if values:
                    enums[name_hint] = EnumDef(name_hint, tuple(values))
                    return "enum", name_hint
                base = child.get("base")
                if base:
                    return st_kind(*d.resolve(child, base))
                inner = child.find(_Q + "simpleType")
                return simple(d, inner, name_hint) if inner is not None else ("string", None)
            if tag == "union":
                members = [st_kind(*d.resolve(child, m))
                           for m in (child.get("memberTypes") or "").split()]
                members += [simple(d, s, name_hint + "#u") for s in child.findall(_Q + "simpleType")]
                vals: list[str] = []
                for kind, key in members:
                    if kind == "bool":            # s:ST_OnOff = xsd:boolean | on/off
                        vals += [v for v in ("true", "false", "1", "0") if v not in vals]
                        continue
                    if kind != "enum" or key is None:
                        return "string", None   # a union with a non-enumerated member
                    vals += [v for v in enums[key].values if v not in vals]
                enums[name_hint] = EnumDef(name_hint, tuple(vals))
                return "enum", name_hint
            if tag == "list":
                return "list", None
        return "string", None

    def st_kind(uri: str, local: str) -> tuple[str, str | None]:
        if uri == XS:
            return _BUILTIN_KIND.get(local, "string"), None
        key = (uri, local)
        if key not in st_cache:
            if key not in stypes:
                st_cache[key] = ("string", None)
                warnings.append(f"unknown simple type {qn(uri, local)}")
            else:
                st_cache[key] = ("string", None)   # guards recursion
                d, node = stypes[key]
                st_cache[key] = simple(d, node, qn(uri, local))
        return st_cache[key]

    # --- attributes ----------------------------------------------------------------
    def attr_def(d: _Doc, node) -> AttrDef | None:
        if node.get("use") == "prohibited":
            return None
        ref = node.get("ref")
        if ref:
            uri, local = d.resolve(node, ref)
            gd, gnode = gattrs.get((uri, local), (d, None))
            kind, enum = attr_type(gd, gnode) if gnode is not None else ("string", None)
            return AttrDef(qn(uri, local), kind, node.get("use") == "required", enum)
        name = node.get("name")
        form = node.get("form")
        qualified = form == "qualified" or (form is None and d.attr_qualified)
        kind, enum = attr_type(d, node, f"{prefix_of(d.tns)}:anon_{name}")
        return AttrDef(qn(d.tns, name) if qualified else name, kind,
                       node.get("use") == "required", enum)

    def attr_type(d: _Doc, node, hint: str = "") -> tuple[str, str | None]:
        t = node.get("type")
        if t:
            return st_kind(*d.resolve(node, t))
        inner = node.find(_Q + "simpleType")
        return simple(d, inner, hint) if inner is not None else ("string", None)

    def attrs_of(d: _Doc, node, out: dict[str, AttrDef]) -> bool:
        """Collect attributes (and attribute groups) of a type node; returns
        True if it has anyAttribute."""
        any_attr = False
        for child in node:
            if child.tag == _Q + "attribute":
                a = attr_def(d, child)
                if a is not None:
                    out[a.qname] = a
            elif child.tag == _Q + "attributeGroup":
                uri, local = d.resolve(child, child.get("ref"))
                if (uri, local) in agroups:
                    gd, gnode = agroups[(uri, local)]
                    any_attr |= attrs_of(gd, gnode, out)
                else:
                    warnings.append(f"unknown attribute group {qn(uri, local)}")
            elif child.tag == _Q + "anyAttribute":
                any_attr = True
        return any_attr

    # --- particles -----------------------------------------------------------------
    types: dict[str, TypeDef] = {}
    pending: list[tuple[str, str, _Doc, Any]] = []   # (key, qname, doc, type node)

    def element_key(d: _Doc, node, owner_uri: str, local_scope: bool) -> tuple[str, str]:
        """(key, qname) of an element declaration; queues its type."""
        ref = node.get("ref")
        if ref:
            uri, local = d.resolve(node, ref)
            if (uri, local) not in gelements:
                warnings.append(f"unknown element {qn(uri, local)}")
                return f"?/{qn(uri, local)}", qn(uri, local)
            gd, gnode = gelements[(uri, local)]
            return element_key(gd, gnode, uri, False)
        name = node.get("name")
        form = node.get("form")
        qualified = not local_scope or form == "qualified" or \
            (form is None and d.elem_qualified)
        uri = d.tns if qualified else ""
        qname = qn(uri, name)
        t = node.get("type")
        if t:
            turi, tlocal = d.resolve(node, t)
            if turi == XS or (turi, tlocal) in stypes:
                key = f"{'xsd:' + tlocal if turi == XS else qn(turi, tlocal)}/{qname}"
                if turi != XS:
                    key = f"xsd:string/{qname}" if st_kind(turi, tlocal)[0] != "enum" \
                        else f"{qn(turi, tlocal)}/{qname}"
                if key not in types:
                    types[key] = TypeDef(key, qname, "", leaf_text=True)
                return key, qname
            if (turi, tlocal) not in ctypes:
                warnings.append(f"unknown type {qn(turi, tlocal)} of {qname}")
                return f"?/{qname}", qname
            key = f"{qn(turi, tlocal)}/{qname}"
            if key not in types:
                types[key] = TypeDef(key, qname, "")      # placeholder, filled below
                td, tnode = ctypes[(turi, tlocal)]
                pending.append((key, qname, td, tnode))
            return key, qname
        inner = node.find(_Q + "complexType")
        if inner is None:                                  # no type: anyType text
            key = f"xsd:anyType/{qname}"
            types.setdefault(key, TypeDef(key, qname, "", leaf_text=True))
            return key, qname
        key = f"{prefix_of(d.tns)}:anon_{name}/{qname}"
        if key not in types:
            types[key] = TypeDef(key, qname, "")
            pending.append((key, qname, d, inner))
        return key, qname

    def particle(d: _Doc, node, keys: dict[str, str]) -> tuple | None:
        tag = node.tag[len(_Q):]
        mn, mx = _occurs(node)
        if tag == "element":
            key, qname = element_key(d, node, d.tns, True)
            keys[qname] = key
            return ("E", qname, mn, mx)
        if tag == "any":
            return ("A", mn, mx)
        if tag == "group":
            uri, local = d.resolve(node, node.get("ref"))
            if (uri, local) not in groups:
                warnings.append(f"unknown group {qn(uri, local)}")
                return None
            gd, gnode = groups[(uri, local)]
            inner = [particle(gd, c, keys) for c in gnode
                     if isinstance(c.tag, str) and c.tag[len(_Q):] in ("sequence", "choice", "all")]
            return ("S", mn, mx, tuple(i for i in inner if i is not None))
        if tag in ("sequence", "choice", "all"):
            items = [particle(d, c, keys) for c in node
                     if isinstance(c.tag, str) and c.tag.startswith(_Q)
                     and c.tag[len(_Q):] in ("element", "any", "group", "sequence", "choice")]
            code = {"sequence": "S", "choice": "C", "all": "L"}[tag]
            return (code, mn, mx, tuple(i for i in items if i is not None))
        return None

    def content(d: _Doc, tnode, keys: dict[str, str], attrs: dict[str, AttrDef]
                ) -> tuple[tuple | None, bool]:
        """(particle, leaf_text) of a complexType node, following extensions."""
        mixed = tnode.get("mixed") == "true"
        for child in tnode:
            if not isinstance(child.tag, str):
                continue
            tag = child.tag[len(_Q):] if child.tag.startswith(_Q) else ""
            if tag in ("sequence", "choice", "all", "group"):
                p = particle(d, child, keys)
                attrs_of(d, tnode, attrs)
                return p, mixed
            if tag == "simpleContent":
                ext = next((c for c in child if isinstance(c.tag, str)), None)
                if ext is not None:
                    base = ext.get("base")
                    if base:
                        buri, blocal = d.resolve(ext, base)
                        if (buri, blocal) in ctypes:          # simpleContent from a complex type
                            bd, bnode = ctypes[(buri, blocal)]
                            content(bd, bnode, keys, attrs)
                    attrs_of(d, ext, attrs)
                return None, True
            if tag == "complexContent":
                ext = next((c for c in child if isinstance(c.tag, str)), None)
                if ext is None:
                    return None, mixed
                own = next((particle(d, c, keys) for c in ext if isinstance(c.tag, str)
                            and c.tag[len(_Q):] in ("sequence", "choice", "all", "group")), None)
                if ext.tag == _Q + "extension":
                    buri, blocal = d.resolve(ext, ext.get("base"))
                    base_p, base_text = None, False
                    if (buri, blocal) in ctypes:
                        bd, bnode = ctypes[(buri, blocal)]
                        base_p, base_text = content(bd, bnode, keys, attrs)
                    attrs_of(d, ext, attrs)
                    parts = tuple(p for p in (base_p, own) if p is not None)
                    p = None if not parts else parts[0] if len(parts) == 1 else ("S", 1, 1, parts)
                    return p, mixed or base_text
                attrs_of(d, ext, attrs)                    # restriction: own content only
                return own, mixed
        attrs_of(d, tnode, attrs)
        return None, mixed

    # roots: every global element
    for (uri, _local), (d, node) in sorted(gelements.items(), key=lambda kv: kv[0]):
        element_key(d, node, uri, False)
    while pending:
        key, qname, d, tnode = pending.pop()
        keys: dict[str, str] = {}
        attrs: dict[str, AttrDef] = {}
        p, leaf = content(d, tnode, keys, attrs)
        walker = _Walker()
        if p is not None:
            walker.walk(p, keys)
        tname = key.partition("/")[0]
        types[key] = TypeDef(
            key=key, qname=qname, class_name=tname.partition(":")[2], attrs=list(attrs.values()),
            children=list(walker.children.values()), leaf_text=leaf, any_child=walker.any_child,
            abstract=tnode.get("abstract") == "true", particle=p)

    namespaces = [Namespace(prefixes[u], u) for u in sorted({d.tns for d in docs}) if u]
    return Schema(source=source or f"XSD files in {Path(schema_dir).name}",
                  namespaces=namespaces, types=types, enums=enums, warnings=warnings)


class _Walker:
    """ChildRefs from a particle tree (the same ranking rules as the SDK frontend)."""

    def __init__(self) -> None:
        self.children: dict[str, ChildRef] = {}
        self.any_child = False

    def walk(self, p: tuple, keys: dict[str, str], rank: tuple[int, ...] = (),
             frozen: bool = False, required: bool = True) -> None:
        if p[0] == "E":
            _, qname, mn, mx = p
            key = keys[qname]
            max_one = (not frozen) and mx == 1
            req = required and mn >= 1
            prev = self.children.get(key)
            if prev is not None:
                rank, max_one, req = min(rank, prev.rank), False, req and prev.required
            self.children[key] = ChildRef(key, qname, rank, max_one, None, req)
            return
        if p[0] == "A":
            self.any_child = True
            return
        code, mn, mx, items = p
        frozen = frozen or mx is None or mx > 1
        is_choice = code == "C" and len(items) > 1
        child_req = required and mn >= 1 and not is_choice
        if code == "S" and len(items) > 1:
            for i, item in enumerate(items):
                self.walk(item, keys, rank if frozen else (*rank, i), frozen, child_req)
        else:
            for item in items:
                self.walk(item, keys, rank, frozen, child_req)
