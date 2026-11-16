"""Frontend: Open XML SDK metadata (``dotnet/Open-XML-SDK/data``) -> IR.

Facts about the input that this module relies on (verified against the SDK
data at the pinned commit; see tests/test_generator.py):

* ``schemas/*.json``: ``{"TargetNamespace", "Types": [...], "Enums": [...]}``.
* A type's ``Name`` is ``TYPE/qname``; abstract bases are ``TYPE/`` (empty
  element part) and carry the attributes/particle shared by derived elements
  such as ``w:CT_OnOff/w:b`` (``IsDerived: true``).
* Particle ``Occurs``: absent -> exactly once; present -> list of
  ``{"Min", "Max"}`` where a missing Min is 0 and a missing/0 Max is
  unbounded.  Extra entries carry ``Version``/``IncludeVersion`` (Office
  version-specific cardinality); we take the first unversioned entry.
* Attribute ``Type`` is a .NET type (``EnumValue<Ns.FooValues>`` etc.); the
  .NET heuristics (ClassName, BaseClass, CompositeType) are *not* treated as
  schema facts, only as naming hints.
"""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

from .ir import OFFICE_VERSIONS, AttrDef, ChildRef, EnumDef, Namespace, Schema, TypeDef

_KIND_BY_DOTNET = {
    "StringValue": "string",
    "BooleanValue": "bool",
    "OnOffValue": "bool",
    "TrueFalseValue": "tf",
    "TrueFalseBlankValue": "tf",
    "EnumValue": "enum",
    "ByteValue": "int",
    "SByteValue": "int",
    "Int16Value": "int",
    "UInt16Value": "int",
    "Int32Value": "int",
    "UInt32Value": "int",
    "Int64Value": "int",
    "UInt64Value": "int",
    "IntegerValue": "int",
    "DoubleValue": "float",
    "SingleValue": "float",
    "DecimalValue": "float",
    "DateTimeValue": "datetime",
    "ListValue": "list",
    "HexBinaryValue": "hex",
    "Base64BinaryValue": "base64",
}


_FLOAT_TYPES = {"xsd:double", "xsd:float", "xsd:decimal"}


def _version_rank(v: str | None) -> int:
    return OFFICE_VERSIONS.index(v) if v in OFFICE_VERSIONS else -1


def _compile_lexical(validators: list[dict], resolve_enum, file_stem: str,
                     attr_kind: str) -> tuple | None:
    """SDK String/Number/Enum validators -> lexical constraint tuple.

    Validators may be tied to Office versions: ``w:before`` is integer-only
    for Office 2007 but a union of integer | "1.5in"-style measure from 2010.
    We validate against the newest rules (the most permissive, and what
    current Office writes): unversioned validators always apply; among
    versioned ones only the newest version's set is kept.
    """
    vals = [v for v in validators if v["Name"] in ("StringValidator", "NumberValidator",
                                                   "EnumValidator")]
    versioned = [v for v in vals if "Version" in v]
    if versioned:
        newest = max(_version_rank(v["Version"]) for v in versioned)
        vals = [v for v in vals if "Version" not in v or _version_rank(v["Version"]) == newest]

    def one(v: dict):
        args = {a.get("Name"): a.get("Value") for a in v.get("Arguments", [])}
        vtype = v.get("Type") or ""
        if v["Name"] == "NumberValidator" and vtype == "xsd:boolean":
            return ("str", "true|false|1|0", None, None, None)
        if v["Name"] == "NumberValidator":
            is_float = vtype in _FLOAT_TYPES or (not vtype and attr_kind == "float")
            num = lambda k: (float(args[k]) if is_float else int(args[k])) if k in args else None
            lo_excl = num("MinExclusive")
            lo_incl = num("MinInclusive")
            if args.get("IsPositive") == "True" and lo_incl is None and lo_excl is None:
                lo_excl = 0
            if (args.get("IsNonNegative") == "True" or "unsigned" in vtype.lower()) \
                    and lo_incl is None and lo_excl is None:
                lo_incl = 0
            return ("num", is_float, lo_incl, num("MaxInclusive"), lo_excl, num("MaxExclusive"))
        if v["Name"] == "EnumValidator":
            key = resolve_enum(file_stem, v.get("Type", ""), by_type=True)
            return ("enum", key) if key else None
        pat = args.get("Pattern")
        ints = lambda k: int(args[k]) if k in args else None
        if attr_kind == "hex" or "Hex" in vtype:
            # hexBinary: lengths count BYTES (w:color/@w:val Length 3 = "1F3864")
            dbl = lambda n: None if n is None else 2 * n
            return ("str", pat or "[0-9A-Fa-f]*", dbl(ints("MinLength")), dbl(ints("MaxLength")),
                    dbl(ints("Length")))
        c = ("str", pat, ints("MinLength"), ints("MaxLength"), ints("Length"))
        return c if any(x is not None for x in c[1:]) else None

    items: list = []
    unions: dict[int, list] = {}
    for v in vals:
        c = one(v)
        if "UnionId" in v:
            unions.setdefault(v["UnionId"], []).append(c)  # None member = "anything"
        elif c is not None:
            items.append(c)
    for _, members in sorted(unions.items()):
        if any(m is None for m in members):
            continue  # an unconstrained alternative accepts everything
        items.append(("union", tuple(members)) if len(members) > 1 else members[0])
    return tuple(items) or None


def _occurs(p: dict) -> tuple[int, int | None]:
    occ = p.get("Occurs")
    if occ is None:
        return 1, 1
    entry = next((o for o in occ if "Version" not in o and "IncludeVersion" not in o), occ[0])
    mx = entry.get("Max", 0)
    return entry.get("Min", 0), (None if mx == 0 else mx)


def _repeatable(p: dict) -> bool:
    _, mx = _occurs(p)
    return mx is None or mx > 1


def _particle_tree(p: dict) -> tuple:
    """The SDK particle as a compact tree (see ``ir.TypeDef.particle``)."""
    kind = p.get("Kind")
    mn, mx = _occurs(p)
    if kind is None:
        return ("E", p["Name"].partition("/")[2], mn, mx)
    if kind == "Any":
        return ("A", mn, mx)
    items = tuple(_particle_tree(i) for i in p.get("Items", []))
    code = {"Choice": "C", "All": "L"}.get(kind, "S")  # Sequence and Group: in order
    return (code, mn, mx, items)


class _ParticleWalker:
    """Flattens a particle tree into ChildRefs with ordering ranks.

    Ordering is only enforced where the schema makes it deterministic: inside
    a repeatable compositor (e.g. ``(a, b)*`` or an unbounded choice) every
    descendant shares the compositor's rank, so insertion order is preserved
    instead of wrongly regrouping ``a b a`` into ``a a b``.
    """

    def __init__(self) -> None:
        self.children: dict[str, ChildRef] = {}
        self.any_child = False

    def walk(self, p: dict, rank: tuple[int, ...] = (), frozen: bool = False,
             required: bool = True) -> None:
        kind = p.get("Kind")
        mn, _ = _occurs(p)
        if kind is None:  # element
            key = p["Name"]
            qname = key.partition("/")[2]
            _, mx = _occurs(p)
            max_one = (not frozen) and mx == 1
            req = required and mn >= 1
            prev = self.children.get(key)
            if prev is not None:  # appears twice in the model
                rank, max_one, req = min(rank, prev.rank), False, req and prev.required
            self.children[key] = ChildRef(key, qname, rank, max_one, p.get("InitialVersion"), req)
            return
        if kind == "Any":
            self.any_child = True
            return
        frozen = frozen or _repeatable(p)
        items = p.get("Items", [])
        is_choice = kind == "Choice" and len(items) > 1
        child_req = required and mn >= 1 and not is_choice
        if kind == "Sequence" or (kind == "Group" and len(items) > 1):
            for i, item in enumerate(items):
                self.walk(item, rank if frozen else (*rank, i), frozen, child_req)
        else:  # Choice, All, single-item Group
            for item in items:
                self.walk(item, rank, frozen, child_req)


def _git_rev(path: Path) -> str:
    try:
        out = subprocess.run(
            ["git", "-C", str(path), "log", "-1", "--format=%H %cs"],
            capture_output=True, text=True, check=True, timeout=10,
        )
        return out.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return "unknown revision"


def load(data_dir: str | Path) -> Schema:
    data_dir = Path(data_dir)
    warnings: list[str] = []

    namespaces = [
        Namespace(n["Prefix"], n["Uri"], n.get("Version"))
        for n in json.loads((data_dir / "namespaces.json").read_text("utf-8-sig"))
        if n["Prefix"]
    ]

    raw: dict[str, dict] = {}
    raw_enums: list[tuple[str, dict]] = []  # (file stem, enum)
    attr_enum_refs: list[tuple[str, str]] = []  # (file stem, dotnet type)
    for f in sorted((data_dir / "schemas").glob("*.json")):
        doc = json.loads(f.read_text("utf-8-sig"))
        for e in doc.get("Enums", []):
            raw_enums.append((f.stem, e))
        for t in doc.get("Types", []):
            t["_file"] = f.stem
            raw[t["Name"]] = t
            for a in t.get("Attributes", []):
                if a["Type"].startswith("EnumValue<"):
                    attr_enum_refs.append((f.stem, a["Type"][10:-1]))

    # --- enums: map "<dotnet ns>.<Name>" -> EnumDef.key ---------------------
    enums: dict[str, EnumDef] = {}
    by_file_name: dict[tuple[str, str], str] = {}
    by_name: dict[str, list[str]] = {}
    for stem, e in raw_enums:
        key = e.get("Type") or f"{stem}:{e['Name']}"
        if key in enums:
            key = f"{key}#{stem}"
        enums[key] = EnumDef(key, tuple(fa["Value"] for fa in e.get("Facets", [])))
        by_file_name[(stem, e["Name"])] = key
        by_name.setdefault(e["Name"], []).append(key)
    dotnet_ns_file: dict[str, str] = {}
    for stem, dotnet in attr_enum_refs:  # learn .NET namespace -> schema file
        ns, _, name = dotnet.rpartition(".")
        if (stem, name) in by_file_name:
            dotnet_ns_file.setdefault(ns, stem)

    def resolve_enum(stem: str, dotnet: str, by_type: bool = False) -> str | None:
        if by_type:  # EnumValidator gives the schema type, e.g. "w:ST_Jc"
            return dotnet if dotnet in enums else None
        ns, _, name = dotnet.rpartition(".")
        for s in (dotnet_ns_file.get(ns), stem):
            if s and (s, name) in by_file_name:
                return by_file_name[(s, name)]
        cands = by_name.get(name, [])
        if len(cands) == 1:
            return cands[0]
        warnings.append(f"unresolved enum {dotnet} (from {stem})")
        return None

    # --- types ---------------------------------------------------------------
    by_class = {t["ClassName"]: t for t in raw.values()}

    def base_of(t: dict) -> dict | None:
        tpart = t["Name"].partition("/")[0]
        b = raw.get(tpart + "/")
        if b is None and t.get("IsDerived"):
            b = by_class.get(t.get("BaseClass", ""))
        return b if b is not t else None

    def merged(t: dict, field: str):
        if field in t:
            return t[field]
        b = base_of(t)
        return merged(b, field) if b is not None else None

    types: dict[str, TypeDef] = {}
    for key, t in raw.items():
        tpart, _, qname = key.partition("/")
        attrs = []
        for a in merged(t, "Attributes") or []:
            dotnet = a["Type"]
            kind = _KIND_BY_DOTNET.get(dotnet.partition("<")[0], "string")
            required = any(
                v["Name"] == "RequiredValidator" and "Version" not in v
                for v in a.get("Validators", [])
            )
            enum = resolve_enum(t["_file"], dotnet[10:-1]) if kind == "enum" else None
            if kind == "enum" and enum is None:
                kind = "string"
            version = a.get("Version") or next(
                (arg.get("Value") for v in a.get("Validators", [])
                 if v["Name"] == "OfficeVersionValidator" for arg in v.get("Arguments", [])), None)
            lexical = None if kind == "enum" else _compile_lexical(
                a.get("Validators", []), resolve_enum, t["_file"], kind)
            attrs.append(AttrDef(a["QName"].lstrip(":"), kind, required, enum,
                                 version, a.get("PropertyComments", ""), lexical))

        walker = _ParticleWalker()
        particle = merged(t, "Particle")
        ordered = True
        if particle is not None:
            walker.walk(particle)
        else:
            kids = merged(t, "Children") or []
            if kids:
                ordered = False
                for c in kids:
                    walker.children.setdefault(
                        c["Name"], ChildRef(c["Name"], c["Name"].partition("/")[2], (), False))
        for extra in merged(t, "AdditionalElements") or []:
            if extra not in walker.children:
                ordered = False
                walker.children[extra] = ChildRef(extra, extra.partition("/")[2], (), False)

        leaf_text = bool(merged(t, "IsLeafText")) or tpart.startswith("xsd:")
        types[key] = TypeDef(
            key=key,
            qname=qname or None,
            class_name=t["ClassName"],
            summary=(t.get("Summary") or "").strip(),
            attrs=attrs,
            children=list(walker.children.values()),
            leaf_text=leaf_text,
            any_child=walker.any_child,
            ordered=ordered,
            abstract=bool(t.get("IsAbstract")) or not qname,
            part_root="Part" in t,
            version=t.get("Version"),
            particle=_particle_tree(particle) if particle is not None and ordered else None,
        )

    missing = {c.key for td in types.values() for c in td.children if c.key not in types}
    for m in sorted(missing):
        warnings.append(f"child reference without definition: {m}")

    parts = [json.loads(f.read_text("utf-8-sig")) for f in sorted((data_dir / "parts").glob("*.json"))]

    return Schema(
        source=f"Open XML SDK data @ {_git_rev(data_dir)}",
        namespaces=namespaces,
        types=types,
        enums=enums,
        parts=parts,
        warnings=warnings,
    )
