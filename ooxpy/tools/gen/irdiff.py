"""Compare two IRs (the SDK frontend's and the XSD frontend's) and report
every deviation, for triage (Phase 3).

    python -m tools.gen.irdiff --sdk-data /path/to/Open-XML-SDK/data \\
        --xsd /path/to/schemas [--out docs/xsd-diff] [--json]

Pairing.  A type in one IR is paired with a type in the other when the keys
are equal (``w:CT_RPr/w:rPr``), and otherwise through context: children with
the same qname of two paired parents are paired too, so a renamed type
(the SDK's ``w:CT_OnOffOnly/w:cantSplit`` for the XSD's
``w:CT_OnOff/w:cantSplit``) is still compared, and reported as renamed.

Comparison is limited to the XSD's namespaces: the SDK's Microsoft extension
elements and attributes (w14, x14, ...) inside a Transitional type are
listed as extensions, not as differences.  Content models are compared after
normalising both trees the same way (groups inlined, single-item and
nested 1..1 compositors flattened).

Categories (one finding per type pair and category):
    only_sdk, only_xsd       types present on one side
    renamed                  paired through context, different type names
    children                 child qnames on one side only
    model                    same children, different content model
                             (order, choice vs sequence, occurrences)
    attrs                    attributes on one side only
    required                 an attribute required on one side only
    enum                     enumeration values differ
    text                     simple content on one side only
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

from .ir import Schema, TypeDef


# -- normalisation ------------------------------------------------------------------
def _ns(qname: str) -> str:
    return qname.partition(":")[0] if ":" in qname else ""


def normalize(p: tuple | None, keep_ns: set[str] | None = None) -> tuple | None:
    """A canonical content model: elements outside ``keep_ns`` dropped, groups
    and 1..1 nesting flattened, single-item compositors unwrapped."""
    if p is None:
        return None
    if p[0] == "E":
        if keep_ns is not None and _ns(p[1]) not in keep_ns:
            return None
        return p
    if p[0] == "A":
        return p
    code, mn, mx, items = p
    out: list[tuple] = []
    for item in items:
        n = normalize(item, keep_ns)
        if n is None:
            continue
        # splice a 1..1 sequence into a sequence, a 1..1 choice into a choice
        if n[0] == code and n[0] in ("S", "C") and n[1:3] == (1, 1):
            out.extend(n[3])
        else:
            out.append(n)
    if not out:
        return None
    if len(out) == 1:
        only = out[0]
        if (mn, mx) == (1, 1):
            return only
        o_mn, o_mx = (only[1], only[2]) if only[0] != "E" else (only[2], only[3])
        if (mn, mx) == (0, 1) and (o_mn == 0 or (o_mn == 1 and o_mx is None)):
            # (x?)? = x?, (x*)? = x*, (x+)? = x*: same language
            if only[0] == "E":
                return ("E", only[1], 0, o_mx)
            return (only[0], 0, o_mx, *only[3:]) if only[0] != "A" else ("A", 0, o_mx)
        if only[0] == "E" and only[2:4] == (1, 1):
            return ("E", only[1], mn, mx)
        if only[0] in ("S", "C", "L") and only[1:3] == (1, 1):
            return (only[0], mn, mx, only[3])
        code = "S"          # one item: sequence and choice mean the same
    return (code, mn, mx, tuple(out))


def render(p: tuple | None) -> str:
    if p is None:
        return "(empty)"

    def occ(mn, mx):
        if (mn, mx) == (1, 1):
            return ""
        if (mn, mx) == (0, 1):
            return "?"
        if (mn, mx) == (0, None):
            return "*"
        if (mn, mx) == (1, None):
            return "+"
        return f"{{{mn},{'' if mx is None else mx}}}"

    if p[0] == "E":
        return p[1] + occ(p[2], p[3])
    if p[0] == "A":
        return "any" + occ(p[1], p[2])
    code, mn, mx, items = p
    sep = {"S": ", ", "C": " | ", "L": " & "}[code]
    return "(" + sep.join(render(i) for i in items) + ")" + occ(mn, mx)


def _qnames(p: tuple | None) -> set[str]:
    if p is None:
        return set()
    if p[0] == "E":
        return {p[1]}
    if p[0] == "A":
        return set()
    out: set[str] = set()
    for i in p[3]:
        out |= _qnames(i)
    return out


# -- pairing ------------------------------------------------------------------------
def pair(sdk: Schema, xsd: Schema) -> set[tuple[str, str]]:
    """(xsd key, sdk key) pairs.  Many-to-many: the SDK specialises some XSD
    types per parent (``a:CT_OfficeArtExtensionList`` becomes
    ``a:CT_BlipExtensionList`` under ``a:blip``), so one XSD type can pair
    with several SDK types, each compared."""
    pairs = {(k, k) for k in xsd.types if k in sdk.types}
    # a qname with exactly one type on each side (w14 elements in extension lists
    # have no shared parent to pair them through)
    by_q_x: dict[str, list[str]] = defaultdict(list)
    by_q_s: dict[str, list[str]] = defaultdict(list)
    for k, t in xsd.types.items():
        if t.qname:
            by_q_x[t.qname].append(k)
    for k, t in sdk.types.items():
        if t.qname:
            by_q_s[t.qname].append(k)
    paired_x = {x for x, _ in pairs}
    for q, xs in by_q_x.items():
        ss = by_q_s.get(q, [])
        if len(xs) == 1 and len(ss) == 1 and xs[0] not in paired_x:
            pairs.add((xs[0], ss[0]))
    queue = list(pairs)
    while queue:
        xk, sk = queue.pop()
        xt, st = xsd.types[xk], sdk.types[sk]
        s_by_q = defaultdict(list)
        for c in st.children:
            s_by_q[c.qname].append(c.key)
        for c in xt.children:
            if c.key not in xsd.types:
                continue
            cands = [k for k in s_by_q.get(c.qname, []) if k in sdk.types]
            if len(cands) == 1 and (c.key, cands[0]) not in pairs:
                pairs.add((c.key, cands[0]))
                queue.append((c.key, cands[0]))
    return pairs


def _relaxed(p: tuple | None) -> tuple | None:
    """The model with every finite maxOccurs above 1 made unbounded."""
    if p is None or p[0] == "A":
        return p
    if p[0] == "E":
        return ("E", p[1], p[2], None if p[3] is None or p[3] > 1 else p[3])
    code, mn, mx, items = p
    return (code, mn, None if mx is None or mx > 1 else mx, tuple(_relaxed(i) for i in items))


# -- comparison ---------------------------------------------------------------------
def _attr_map(t: TypeDef, keep_ns: set[str] | None) -> dict:
    return {a.qname: a for a in t.attrs
            if keep_ns is None or _ns(a.qname) in keep_ns or not _ns(a.qname)}


def compare(sdk: Schema, xsd: Schema) -> dict:
    """``{"findings": [...], "stats": {...}}``; each finding is a dict with
    category, namespace, keys and details."""
    keep_all = {n.prefix for n in xsd.namespaces} | {"xml", ""}
    # Microsoft's extension namespaces (schemas.microsoft.com) extend the
    # standard types from the outside: inside a standard type, the SDK's w14,
    # x14 ... children and attributes are extensions, not differences
    ms = {n.prefix for n in xsd.namespaces if "schemas.microsoft.com" in n.uri} | \
        {n.prefix for n in sdk.namespaces if "schemas.microsoft.com" in n.uri}
    keep_std = keep_all - ms
    pairs = pair(sdk, xsd)
    paired_sdk = {sk for _, sk in pairs}
    paired_xsd = {xk for xk, _ in pairs}
    covered_by_any = 0
    findings: list[dict] = []
    extensions: Counter = Counter()

    def add(cat: str, xk: str | None, sk: str | None, **detail) -> None:
        k = xk or sk or ""
        q = k.partition("/")[2] or k
        findings.append({"category": cat, "ns": _ns(q), "xsd": xk, "sdk": sk, **detail})

    for xk, sk in sorted(pairs):
        xt, st = xsd.types[xk], sdk.types[sk]
        keep = keep_all if _ns(xt.qname or "") in ms else keep_std
        if xk != sk:
            add("renamed", xk, sk)
        s_model = normalize(st.particle, keep)
        x_model = normalize(xt.particle, keep)
        s_all = {c.qname for c in st.children}
        s_kids = {q for q in s_all if _ns(q) in keep}
        for q in s_all - s_kids:
            extensions[_ns(q)] += 1
        x_kids = {c.qname for c in xt.children}
        if xt.any_child and s_kids - x_kids:      # the SDK names what xsd:any allows
            covered_by_any += len(s_kids - x_kids)
            s_kids &= x_kids
        if s_kids != x_kids:
            add("children", xk, sk, only_sdk=sorted(s_kids - x_kids),
                only_xsd=sorted(x_kids - s_kids))
        elif xt.any_child or st.any_child:
            pass                                  # wildcards: models aren't comparable
        elif s_model != x_model and not (st.particle is None and not st.ordered):
            kind = "limits" if _relaxed(s_model) == _relaxed(x_model) else "structure"
            add("model", xk, sk, sdk_model=render(s_model), xsd_model=render(x_model),
                kind=kind)
        elif st.particle is None and not st.ordered and x_model is not None:
            add("model", xk, sk, sdk_model="(unordered: overrides.py)",
                xsd_model=render(x_model), known=True)
        s_attrs, x_attrs = _attr_map(st, keep), _attr_map(xt, None)
        ext_attrs = {a.qname for a in st.attrs} - set(s_attrs)
        for a in ext_attrs:
            extensions[_ns(a)] += 1
        if set(s_attrs) != set(x_attrs):
            add("attrs", xk, sk, only_sdk=sorted(set(s_attrs) - set(x_attrs)),
                only_xsd=sorted(set(x_attrs) - set(s_attrs)))
        for q in sorted(set(s_attrs) & set(x_attrs)):
            sa, xa = s_attrs[q], x_attrs[q]
            if sa.required != xa.required:
                add("required", xk, sk, attr=q, sdk_required=sa.required,
                    xsd_required=xa.required)
            if sa.enum and xa.enum:
                sv = set(sdk.enums[sa.enum].values) if sa.enum in sdk.enums else set()
                xv = set(xsd.enums[xa.enum].values) if xa.enum in xsd.enums else set()
                if sv != xv:
                    add("enum", xk, sk, attr=q, sdk_enum=sa.enum, xsd_enum=xa.enum,
                        only_sdk=sorted(sv - xv), only_xsd=sorted(xv - sv))
        if st.leaf_text != xt.leaf_text and not st.children and not xt.children:
            add("text", xk, sk, sdk_text=st.leaf_text, xsd_text=xt.leaf_text)

    for xk in sorted(set(xsd.types) - paired_xsd):
        add("only_xsd", xk, None)
    for sk in sorted(k for k in sdk.types if k not in paired_sdk
                     and _ns(k.partition("/")[2]) in keep_all and sdk.types[k].qname):
        add("only_sdk", None, sk)
    stats = {
        "xsd_types": len(xsd.types), "sdk_types_in_xsd_namespaces": sum(
            1 for k, t in sdk.types.items() if t.qname and _ns(t.qname) in keep_all),
        "paired": len(pairs), "renamed": sum(1 for x, s in pairs if x != s),
        "sdk_children_covered_by_xsd_any": covered_by_any,
        "by_category": dict(Counter(f["category"] for f in findings)),
        "sdk_extension_items_in_paired_types": dict(extensions.most_common()),
        "xsd_warnings": xsd.warnings,
    }
    return {"findings": findings, "stats": stats}


# -- report -------------------------------------------------------------------------
_TITLES = {
    "children": "Children on one side only",
    "model": "Same children, different content model",
    "attrs": "Attributes on one side only",
    "required": "Required on one side only",
    "enum": "Enumeration values differ",
    "text": "Simple content on one side only",
    "renamed": "Paired through context, different type names",
    "only_sdk": "Types only in the SDK data (in the XSD's namespaces)",
    "only_xsd": "Types only in the XSD",
}


def stricter(f: dict) -> set[str]:
    """Which side rejects what the other allows: "sdk" (oox may flag a file
    ISO/IEC 29500 allows), "xsd" (oox may accept what ISO forbids)."""
    out: set[str] = set()
    cat = f["category"]
    if cat in ("children", "attrs", "enum"):
        if f.get("only_xsd"):
            out.add("sdk")
        if f.get("only_sdk"):
            out.add("xsd")
    elif cat == "required":
        out.add("sdk" if f["sdk_required"] else "xsd")
    return out


def markdown(result: dict, sdk_source: str, xsd_source: str, limit: int = 60) -> str:
    st = result["stats"]
    by_cat: dict[str, list] = defaultdict(list)
    for f in result["findings"]:
        by_cat[f["category"]].append(f)
    lines = ["# SDK data vs XSD: IR diff", "",
             f"- SDK: {sdk_source}", f"- XSD: {xsd_source}",
             f"- XSD types: {st['xsd_types']}; SDK types in the XSD's namespaces: "
             f"{st['sdk_types_in_xsd_namespaces']}; paired: {st['paired']} "
             f"({st['renamed']} through context)", "",
             "| Category | Findings | Namespaces |", "|---|---|---|"]
    for cat in _TITLES:
        items = by_cat.get(cat, [])
        if cat == "model":
            items = sorted(items, key=lambda f: (f.get("kind") == "limits", f["xsd"] or ""))
            by_cat[cat] = items
        ns = Counter(f["ns"] for f in items)
        lines.append(f"| {_TITLES[cat]} (`{cat}`) | {len(items)} | "
                     + ", ".join(f"{n or '-'} {c}" for n, c in ns.most_common(8)) + " |")
    sdk_strict = Counter(f["category"] for f in result["findings"] if "sdk" in stricter(f))
    xsd_strict = Counter(f["category"] for f in result["findings"] if "xsd" in stricter(f))
    lines += ["", "Direction (children, attributes, enumerations, required):",
              "- **SDK stricter** (oox's validator may flag what ISO/IEC 29500 allows): "
              + ", ".join(f"{c} {n}" for c, n in sdk_strict.most_common()),
              "- **XSD stricter** (oox may accept what ISO/IEC 29500 forbids): "
              + ", ".join(f"{c} {n}" for c, n in xsd_strict.most_common())]
    limits = sum(1 for f in by_cat.get("model", []) if f.get("kind") == "limits")
    if limits:
        lines += ["", f"Of the model differences, {limits} differ only in finite maximums "
                  "(the SDK's implementation limits, e.g. `a:tr{1,1000}`); they are listed last."]
    if st["sdk_children_covered_by_xsd_any"]:
        lines += ["", f"{st['sdk_children_covered_by_xsd_any']} SDK children fall under an "
                  "`xsd:any` in the XSD and aren't counted as differences."]
    ext = st["sdk_extension_items_in_paired_types"]
    if ext:
        lines += ["", "SDK extension children/attributes inside paired types (not differences): "
                  + ", ".join(f"{n} {c}" for n, c in list(ext.items())[:15])]
    for cat, title in _TITLES.items():
        items = by_cat.get(cat, [])
        if not items:
            continue
        lines += ["", f"## {title}", ""]
        groups = _grouped(items) if cat != "renamed" else [(f, []) for f in items]
        for f, more in groups[:limit]:
            line = _line(f)
            if more:
                names = ", ".join(f"`{m}`" for m in more[:6]) + (" …" if len(more) > 6 else "")
                what = "elements" if cat == "enum" else "SDK types"
                line += f" (and {len(more)} more {what}: {names})"
            lines.append("- " + line)
        if len(groups) > limit:
            lines.append(f"- … {len(groups) - limit} more (see the JSON)")
    if st["xsd_warnings"]:
        lines += ["", "## XSD frontend warnings", ""] + [f"- {w}" for w in st["xsd_warnings"]]
    return "\n".join(lines) + "\n"


def _grouped(items: list[dict]) -> list[tuple[dict, list[str]]]:
    """Findings that differ only in the SDK type (one XSD type paired with
    several SDK specialisations) shown once."""
    out: dict[str, tuple[dict, list[str]]] = {}
    for f in items:
        enum = f["category"] == "enum"      # the same value sets on many elements: once
        drop = ("sdk", "sdk_enum", "xsd", "ns") if enum else ("sdk", "sdk_enum")
        detail = {k: v for k, v in f.items() if k not in drop}
        key = json.dumps(detail, sort_keys=True, default=str)
        label = (f["xsd"] or "").partition("/")[2] if enum else \
            (f["sdk"] or "").partition("/")[0]
        if key in out:
            out[key][1].append(label)
        else:
            out[key] = (f, [])
    return list(out.values())


def _line(f: dict) -> str:
    cat = f["category"]
    if f["xsd"] and f["sdk"] and f["xsd"] != f["sdk"]:
        key = f"`{f['xsd']}` (SDK: `{f['sdk'].partition('/')[0]}`)"
    else:
        key = f"`{f['xsd'] or f['sdk']}`"
    if cat == "renamed":
        return f"`{f['xsd']}` (XSD) = `{f['sdk']}` (SDK)"
    if cat in ("children", "attrs"):
        parts = []
        if f["only_sdk"]:
            parts.append("SDK only: " + ", ".join(f"`{q}`" for q in f["only_sdk"]))
        if f["only_xsd"]:
            parts.append("XSD only: " + ", ".join(f"`{q}`" for q in f["only_xsd"]))
        return f"{key}: " + "; ".join(parts)
    if cat == "model":
        tag = " (known: overrides.py)" if f.get("known") else \
            " (limits only)" if f.get("kind") == "limits" else ""
        return f"{key}{tag}: SDK `{f['sdk_model']}` vs XSD `{f['xsd_model']}`"
    if cat == "required":
        side = "SDK" if f["sdk_required"] else "XSD"
        return f"{key} `{f['attr']}`: required only in the {side}"
    if cat == "enum":
        parts = []
        if f["only_sdk"]:
            parts.append("SDK only: " + ", ".join(f["only_sdk"]))
        if f["only_xsd"]:
            parts.append("XSD only: " + ", ".join(f["only_xsd"]))
        return f"{key} `{f['attr']}` (`{f['sdk_enum']}` / `{f['xsd_enum']}`): " + "; ".join(parts)
    if cat == "text":
        return f"{key}: text in {'SDK' if f['sdk_text'] else 'XSD'} only"
    return key


def main(argv: list[str] | None = None) -> int:
    from . import overrides, sdk_frontend, xsd_frontend

    ap = argparse.ArgumentParser(prog="python -m tools.gen.irdiff", description=__doc__)
    ap.add_argument("--sdk-data", required=True, type=Path)
    ap.add_argument("--xsd", required=True, type=Path, help="directory of *.xsd files")
    ap.add_argument("--out", type=Path, help="write OUT.md and OUT.json")
    ap.add_argument("--limit", type=int, default=60, help="findings per category in the .md")
    ap.add_argument("--label", help="how to name the XSD set in the report")
    args = ap.parse_args(argv)
    sdk = sdk_frontend.load(args.sdk_data)
    overrides.apply(sdk)
    prefixes: dict[str, str] = {}
    for n in sdk.namespaces:
        prefixes.setdefault(n.uri, n.prefix)
    xsd = xsd_frontend.load(args.xsd, prefixes=prefixes, source=args.label)
    result = compare(sdk, xsd)
    md = markdown(result, sdk.source, xsd.source, args.limit)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.with_suffix(".md").write_text(md, encoding="utf-8")
        args.out.with_suffix(".json").write_text(
            json.dumps(result, indent=1, ensure_ascii=False), encoding="utf-8")
        print(f"wrote {args.out.with_suffix('.md')} and .json")
    else:
        sys.stdout.write(md)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
