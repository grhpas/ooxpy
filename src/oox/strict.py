"""Strict <-> Transitional conformance (ISO/IEC 29500 Strict, ECMA-376 Part 1).

Strict documents use other namespace URIs and relationship types:

    http://schemas.openxmlformats.org/wordprocessingml/2006/main   (Transitional)
    http://purl.oclc.org/ooxml/wordprocessingml/main               (Strict)

oox keeps every tree in Transitional form in memory, so the generated grammar,
validation and all helpers work unchanged.  A package opened as Strict is
written back as Strict (``Package.conformance``); set it to convert.

The translation is textual on namespace declarations and the few attributes
whose *values* are namespace URIs (``a:graphicData/@uri``), which keeps
element prefixes intact.  Two value forms differ as well and are translated
on the parsed tree (``values_to_transitional`` / ``values_to_strict``):

- DrawingML percentages: Strict writes ``"50%"``, Transitional an integer in
  thousandths of a percent (``50000``).  Any integer-typed attribute holding
  a percentage is converted on reading (the part remembers which ones); on
  writing Strict, those plus the known percentage attributes go back.
- ST_OnOff: Strict is ``xsd:boolean``, so ``on``/``off`` become ``true``/``false``.
"""

from __future__ import annotations

import re

T_ROOT = "http://schemas.openxmlformats.org/"
S_ROOT = "http://purl.oclc.org/ooxml/"
_HEADS = {"wordprocessingml", "spreadsheetml", "presentationml", "drawingml",
          "officeDocument", "schemaLibrary"}
_RENAMED = {"extended-properties": "extendedProperties", "custom-properties": "customProperties"}
_UNRENAMED = {v: k for k, v in _RENAMED.items()}
_T_RELS = T_ROOT + "officeDocument/2006/relationships/"
_S_RELS = S_ROOT + "officeDocument/relationships/"


def ns_to_strict(uri: str) -> str:
    if uri.startswith(T_ROOT):
        bits = uri[len(T_ROOT):].split("/")
        if len(bits) == 3 and bits[1] == "2006" and bits[0] in _HEADS:
            return f"{S_ROOT}{bits[0]}/{_RENAMED.get(bits[2], bits[2])}"
    return uri


def ns_to_transitional(uri: str) -> str:
    if uri.startswith(S_ROOT):
        bits = uri[len(S_ROOT):].split("/")
        if len(bits) == 2 and bits[0] in _HEADS:
            return f"{T_ROOT}{bits[0]}/2006/{_UNRENAMED.get(bits[1], bits[1])}"
    return uri


def reltype_to_strict(t: str) -> str:
    if t.startswith(_T_RELS):
        name = t[len(_T_RELS):]
        return _S_RELS + _RENAMED.get(name, name)
    return t


def reltype_to_transitional(t: str) -> str:
    if t.startswith(_S_RELS):
        name = t[len(_S_RELS):]
        return _T_RELS + _UNRENAMED.get(name, name)
    return t


# xmlns / xmlns:p declarations and a:graphicData/@uri (whose value names a namespace)
_URI_ATTRS = re.compile(rb'(\sxmlns(?::[A-Za-z_][\w.\-]*)?|\suri)\s*=\s*(["\'])([^"\']*)\2')


def _rewrite(data: bytes, fn) -> bytes:
    def sub(m: re.Match) -> bytes:
        old = m.group(3).decode("utf-8", "replace")
        new = fn(old)
        return m.group(0) if new == old else m.group(1) + b"=" + m.group(2) + new.encode() + m.group(2)
    return _URI_ATTRS.sub(sub, data)


def xml_to_transitional(data: bytes) -> bytes:
    return _rewrite(data, ns_to_transitional)


def xml_to_strict(data: bytes) -> bytes:
    return _rewrite(data, ns_to_strict)


def is_strict_reltype(t: str) -> bool:
    return t.startswith(S_ROOT)


# -- value forms -------------------------------------------------------------------
_PERCENT = re.compile(r"^(-?\d+(?:\.\d+)?)%$")
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_COLOR_TRANSFORMS = ("tint", "shade", "alpha", "alphaOff", "alphaMod", "hueMod", "sat",
                     "satOff", "satMod", "lum", "lumOff", "lumMod", "red", "redOff", "redMod",
                     "green", "greenOff", "greenMod", "blue", "blueOff", "blueMod")
# DrawingML attributes typed as percentages (ST_Percentage and its restrictions)
KNOWN_PERCENT: frozenset[tuple[str, str]] = frozenset(
    [(_A + t, "val") for t in _COLOR_TRANSFORMS]
    + [(_A + t, side) for t in ("srcRect", "fillRect", "fillToRect", "tileRect")
       for side in ("l", "t", "r", "b")]
    + [(_A + "gs", "pos"), (_A + "spcPct", "val"), (_A + "buSzPct", "val"),
       (_A + "alphaModFix", "amt"), (_A + "tile", "sx"), (_A + "tile", "sy"),
       (_A + "lum", "bright"), (_A + "lum", "contrast"), (_A + "biLevel", "thresh"),
       (_A + "alphaBiLevel", "thresh")])


def _int_attr(el, attr: str) -> bool:
    info = getattr(el, "info", None)
    if info is None:
        return False
    from .grammar import qname_of

    q = qname_of(attr) if attr.startswith("{") else attr
    spec = info.attrs.get(q or "")
    return spec is not None and spec.kind == "int"


def values_to_transitional(root) -> set[tuple[str, str]]:
    """Strict percentages (``"50%"``) on integer attributes -> thousandths.
    Returns the ``(element tag, attribute)`` pairs converted."""
    seen: set[tuple[str, str]] = set()
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        for attr, value in el.attrib.items():
            m = _PERCENT.match(value)
            if m and _int_attr(el, attr):
                el.set(attr, str(round(float(m.group(1)) * 1000)))
                seen.add((el.tag, attr))
    return seen


def values_to_strict(root, percent: set[tuple[str, str]] | frozenset = frozenset()) -> None:
    """In place, on a copy about to be written as Strict: thousandths ->
    ``"50%"`` for ``percent`` and ``KNOWN_PERCENT`` attributes; ST_OnOff
    ``on``/``off`` -> ``true``/``false``."""
    wanted = KNOWN_PERCENT | set(percent)
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        info = None
        for attr, value in el.attrib.items():
            if (el.tag, attr) in wanted and re.fullmatch(r"-?\d+", value):
                el.set(attr, f"{int(value) / 1000:g}%")
            elif value in ("on", "off"):
                if info is None:
                    info = getattr(el, "info", None)
                from .grammar import qname_of

                q = qname_of(attr) if attr.startswith("{") else attr
                spec = info.attrs.get(q or "") if info is not None else None
                if spec is not None and (spec.kind == "bool" or spec.enum == "w:ST_OnOffOnly"):
                    el.set(attr, "true" if value == "on" else "false")
