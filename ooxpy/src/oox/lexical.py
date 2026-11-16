"""Lexical (value-shape) checks compiled from the schema data.

An attribute's ``lexical`` is a tuple of items that must all hold; an item is
a constraint or ``("union", (constraint, ...))`` where one member suffices:

    ("num", is_float, min_incl, max_incl, min_excl, max_excl)
    ("str", pattern, min_len, max_len, exact_len)     # pattern is an XSD regex
    ("enum", enum_key)

XSD patterns are implicitly anchored, hence ``fullmatch``.  The constraints
follow the newest Office version's rules (see tools/gen/sdk_frontend.py).
"""

from __future__ import annotations

import re
from functools import cache

_INT = re.compile(r"[+-]?\d+")
_FLOAT = re.compile(r"[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?|[+-]?INF|NaN")


# SDK enumerations that ECMA-376 leaves open (ST_String plus a list of
# examples).  Checked at run time until the Phase 3 XSD frontend corrects the
# IR (tools/gen/overrides.py lists them too).
OPEN_ENUMS = {
    # ISO/IEC 29500-1 §17.8.3.2: w:charset/@w:characterSet is ST_String, "the
    # names and aliases listed in the IANA registration table".  The SDK's list
    # misses e.g. windows-1252, which LibreOffice writes for most fonts.
    "w:ST_CharacterSet",
}


@cache
def _regex(pattern: str) -> re.Pattern:
    return re.compile(pattern)


def _check_one(c: tuple, text: str) -> bool:
    kind = c[0]
    if kind == "num":
        _, is_float, lo, hi, lo_x, hi_x = c
        if not (_FLOAT if is_float else _INT).fullmatch(text):
            return False
        v = float(text) if is_float else int(text)
        return ((lo is None or v >= lo) and (hi is None or v <= hi)
                and (lo_x is None or v > lo_x) and (hi_x is None or v < hi_x))
    if kind == "str":
        _, pattern, lo, hi, exact = c
        n = len(text)
        if (lo is not None and n < lo) or (hi is not None and n > hi) or \
                (exact is not None and n != exact):
            return False
        return pattern is None or bool(_regex(pattern).fullmatch(text))
    if kind == "enum":
        from .grammar import enum_values

        if c[1] in OPEN_ENUMS:
            return True
        values = enum_values(c[1])
        return not values or text in values
    if kind == "union":
        return any(_check_one(m, text) for m in c[1])
    return True


def check(lexical: tuple | None, text: str) -> str | None:
    """``None`` if ``text`` is acceptable, else a description of what is expected."""
    if not lexical:
        return None
    for item in lexical:
        if not _check_one(item, text):
            return describe(item)
    return None


def describe(c: tuple) -> str:
    kind = c[0]
    if kind == "union":
        return " or ".join(describe(m) for m in c[1])
    if kind == "num":
        _, is_float, lo, hi, lo_x, hi_x = c
        parts = ["a number" if is_float else "an integer"]
        if lo is not None:
            parts.append(f">= {lo}")
        if lo_x is not None:
            parts.append(f"> {lo_x}")
        if hi is not None:
            parts.append(f"<= {hi}")
        if hi_x is not None:
            parts.append(f"< {hi_x}")
        return " ".join(parts)
    if kind == "str":
        _, pattern, lo, hi, exact = c
        if pattern == "[0-9A-Fa-f]*":
            what = "hex digits"
            if exact is not None:
                return f"{exact} {what}"
        elif pattern and "(mm|cm|in|pt|pc|pi)" in pattern:
            what = ("a measure like 1.5in or -2cm" if pattern.startswith("-") else
                    "a measure like 1.5in") + " (units: mm cm in pt pc pi)"
        elif pattern and "{12}" in pattern and "{8}" in pattern:
            what = "a GUID like {00000000-0000-0000-0000-000000000000}"
        elif pattern:
            what = f"text matching /{pattern}/"
        else:
            what = "text"
        bits = []
        if exact is not None:
            bits.append(f"exactly {exact} characters")
        if lo is not None:
            bits.append(f"at least {lo} characters")
        if hi is not None:
            bits.append(f"at most {hi} characters")
        return what + (f" ({', '.join(bits)})" if bits else "")
    if kind == "enum":
        from .grammar import enum_values

        vals = enum_values(c[1])
        return "one of: " + ", ".join(vals[:12]) + (", ..." if len(vals) > 12 else "")
    return "a valid value"
