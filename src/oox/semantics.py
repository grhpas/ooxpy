"""Rules the schema cannot express, from the prose of ECMA-376.

Each rule takes an element and returns problem strings; ``validate()`` runs
the rule registered for the element's content model.  Add rules here when an
Office check (tools/officecheck observe cases) shows Office enforces them.
"""

from __future__ import annotations

from collections.abc import Callable

from .grammar import qname_of


def _ref_col(ref: str) -> int | None:
    col = 0
    for ch in ref:
        if ch.isalpha():
            col = col * 26 + ord(ch.upper()) - 64
        else:
            break
    return col or None


def _sheet_rows(el) -> list[str]:
    """ECMA-376 Part 1 §18.3.1.73: rows appear in ascending order of r."""
    out, prev = [], 0
    for row in el:
        if qname_of(row.tag) != "x:row":
            continue
        r = row.get("r")
        if r is None or not r.isdigit():
            prev += 1
            continue
        if int(r) <= prev:
            out.append(f"<x:sheetData>: row {r} after row {prev} "
                       "(rows must be in ascending order; Excel repairs such files)")
        prev = max(prev, int(r))
    return out


def _row_cells(el) -> list[str]:
    """§18.3.1.4: cells within a row appear in ascending column order."""
    out, prev = [], 0
    for c in el:
        if qname_of(c.tag) != "x:c":
            continue
        ref = c.get("r")
        col = _ref_col(ref) if ref else None
        if col is None:
            prev += 1
            continue
        if col <= prev:
            out.append(f"<x:row r=\"{el.get('r')}\">: cell {ref} out of column order")
        prev = max(prev, col)
    return out


RULES: dict[str, Callable] = {
    "x:CT_SheetData/x:sheetData": _sheet_rows,
    "x:CT_Row/x:row": _row_cells,
}
