"""Corrections applied to the IR after the frontend, where the SDK metadata
deviates from ECMA-376.  Every entry needs a reason and evidence; the Phase 3
XSD frontend will find the rest systematically.
"""

from __future__ import annotations

from dataclasses import replace

from .ir import Schema

# ECMA-376 Part 1 sml.xsd declares these as <xsd:choice maxOccurs="unbounded">,
# i.e. children in any order.  The SDK data models them as fixed sequences,
# which flags valid files: openpyxl writes <font><name/><family/><color/><sz/>.
UNORDERED_CHOICE = {
    "x:CT_Font/x:font",
    "x:CT_RPrElt/x:rPr",
}


# Applied at run time for now (oox.lexical.OPEN_ENUMS), so the committed
# generated code stays identical to the pinned SDK's output: enumerations the
# SDK closes although ECMA-376 types the attribute ST_String.
#   w:ST_CharacterSet -- ISO/IEC 29500-1 §17.8.3.2 (IANA character set names)


def apply(schema: Schema) -> None:
    for key in sorted(UNORDERED_CHOICE):
        t = schema.types.get(key)
        if t is None:
            schema.warnings.append(f"override target missing: {key}")
            continue
        t.ordered = False
        t.particle = None
        t.children = [replace(c, rank=(), max_one=False, required=False) for c in t.children]
