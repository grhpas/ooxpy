"""Worksheet layout: column widths, row heights, merged cells, frozen panes,
autofilter, defined names and data validation -- each with a reader.

    book.set_column_width("Sheet1", "A", 24)          # characters, as Excel shows them
    book.set_column_width("Sheet1", "B:D", 12)
    book.set_row_height("Sheet1", 1, 30)              # points
    book.merge_cells("Sheet1", "A1:D1")
    book.freeze_panes("Sheet1", "A2")                 # keep row 1 visible
    book.auto_filter("Sheet1", "A1:D20")
    book.define_name("Sales", "Sheet1", "B2:D20")
    book.add_validation("Sheet1", "E2:E20", values=["Yes", "No"])   # drop-down list

Column widths are given in characters of the default font (Calibri 11:
7 pixels per digit), and stored the way Excel stores them, so Excel's
"Column Width" dialog shows the number you gave.
"""

from __future__ import annotations

import math
import re
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import x

if TYPE_CHECKING:
    from .spreadsheet import SpreadsheetDocument

X = URI["x"]
MAX_DIGIT_WIDTH = 7  # pixels, Calibri 11 (Excel's default font)
_RANGE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$")


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def parse_range(ref: str) -> tuple[int, int, int, int]:
    """"A1:C3" -> (col1, row1, col2, row2), 1-based, normalised."""
    m = _RANGE.match(ref.strip())
    if not m:
        raise ValueError(f"not an A1 range: {ref!r}")
    c1, r1 = _col_index(m.group(1)), int(m.group(2))
    c2, r2 = (_col_index(m.group(3)), int(m.group(4))) if m.group(3) else (c1, r1)
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def range_ref(c1: int, r1: int, c2: int, r2: int) -> str:
    from .spreadsheet import cell_ref

    a, b = cell_ref(c1, r1), cell_ref(c2, r2)
    return a if a == b else f"{a}:{b}"


def absolute(sheet: str, ref: str) -> str:
    """"Sheet 1", "A1:B5" -> "'Sheet 1'!$A$1:$B$5"."""
    from .spreadsheet import column_letter

    c1, r1, c2, r2 = parse_range(ref)
    a = f"${column_letter(c1)}${r1}"
    b = f"${column_letter(c2)}${r2}"
    name = sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet) else \
        "'" + sheet.replace("'", "''") + "'"
    return f"{name}!{a}" if a == b else f"{name}!{a}:{b}"


# -- columns and rows --------------------------------------------------------------------
def width_to_stored(chars: float) -> float:
    """Characters (as Excel's dialog shows) -> the stored ``width``."""
    return math.trunc((chars * MAX_DIGIT_WIDTH + 5) / MAX_DIGIT_WIDTH * 256) / 256


def stored_to_width(width: float) -> float:
    pixels = math.trunc(((256 * width + math.trunc(128 / MAX_DIGIT_WIDTH)) / 256)
                        * MAX_DIGIT_WIDTH)
    return round((pixels - 5) / MAX_DIGIT_WIDTH, 2)


def _cols(ws: OoxElement) -> dict[int, dict[str, str]]:
    cols = ws.first("x:cols")
    out: dict[int, dict[str, str]] = {}
    if cols is None:
        return out
    for col in cols.iterchildren(f"{{{X}}}col"):
        attrs = {k: v for k, v in col.attrib.items() if k not in ("min", "max")}
        for i in range(int(col.get("min", "1")), int(col.get("max", "1")) + 1):
            out[i] = dict(attrs)
    return out


def _write_cols(ws: OoxElement, spec: dict[int, dict[str, str]]) -> None:
    old = ws.first("x:cols")
    if old is not None:
        ws.remove(old)
    if not spec:
        return
    cols = x.Cols()
    run_start: int | None = None
    keys = sorted(spec)
    for i, c in enumerate(keys):
        if run_start is None:
            run_start = c
        nxt = keys[i + 1] if i + 1 < len(keys) else None
        if nxt is None or nxt != c + 1 or spec[nxt] != spec[c]:  # a run of equal columns ends
            el = x.Col(min=run_start, max=c)
            for k, v in spec[c].items():
                el.set(k, v)
            cols.append(el)
            run_start = None
    ws.add(cols)


def set_column_width(ws: OoxElement, cols: str, chars: float | None) -> None:
    m = re.fullmatch(r"\s*([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?\s*", cols)
    if not m:
        raise ValueError(f"columns are letters like 'B' or 'B:D', not {cols!r}")
    first, last = _col_index(m.group(1)), _col_index(m.group(2) or m.group(1))
    spec = _cols(ws)
    for i in range(min(first, last), max(first, last) + 1):
        attrs = spec.get(i, {})
        if chars is None:
            attrs.pop("width", None)
            attrs.pop("customWidth", None)
        else:
            attrs["width"] = repr(width_to_stored(chars))
            attrs["customWidth"] = "1"
        if attrs:
            spec[i] = attrs
        else:
            spec.pop(i, None)
    _write_cols(ws, spec)


def column_width(ws: OoxElement, col: str) -> float | None:
    attrs = _cols(ws).get(_col_index(col))
    if not attrs or "width" not in attrs:
        return None
    return stored_to_width(float(attrs["width"]))


def _row(book: SpreadsheetDocument, ws: OoxElement, index: int, create: bool) -> OoxElement | None:
    from .spreadsheet import _indexed

    data = ws.first("x:sheetData")
    assert data is not None
    for r_index, r in _indexed(data, "r"):
        if r_index == index:
            return r
        if r_index > index:
            if not create:
                return None
            new = x.Row(r=index)
            r.addprevious(new)
            return new
    if not create:
        return None
    new = x.Row(r=index)
    data.add(new)
    return new


def set_row_height(book: SpreadsheetDocument, ws: OoxElement, row: int, points: float | None) -> None:
    r = _row(book, ws, row, create=points is not None)
    if r is None:
        return
    if points is None:
        r.attrib.pop("ht", None)
        r.attrib.pop("customHeight", None)
    else:
        r.set("ht", repr(float(points)))
        r.set("customHeight", "1")


def row_height(book: SpreadsheetDocument, ws: OoxElement, row: int) -> float | None:
    r = _row(book, ws, row, create=False)
    return float(r.get("ht")) if r is not None and r.get("ht") else None


# -- merged cells ------------------------------------------------------------------------------
def merged_cells(ws: OoxElement) -> list[str]:
    mc = ws.first("x:mergeCells")
    return [] if mc is None else [m.get("ref") for m in mc.iterchildren(f"{{{X}}}mergeCell")]


def merge_cells(ws: OoxElement, ref: str) -> None:
    c1, r1, c2, r2 = parse_range(ref)
    if (c1, r1) == (c2, r2):
        raise ValueError("a merge needs at least two cells")
    for other in merged_cells(ws):
        d1, s1, d2, s2 = parse_range(other)
        if not (c2 < d1 or d2 < c1 or r2 < s1 or s2 < r1):
            raise ValueError(f"{ref} overlaps the merged range {other}")
    mc = ws.first("x:mergeCells")
    if mc is None:
        mc = x.MergeCells()
        ws.add(mc)
    mc.add(x.MergeCell(ref=range_ref(c1, r1, c2, r2)))
    mc.set("count", str(len(mc)))


def unmerge_cells(ws: OoxElement, ref: str) -> None:
    mc = ws.first("x:mergeCells")
    target = range_ref(*parse_range(ref))
    if mc is not None:
        for m in list(mc):
            if m.get("ref") == target:
                mc.remove(m)
        if len(mc) == 0:
            ws.remove(mc)
        else:
            mc.set("count", str(len(mc)))


# -- frozen panes ------------------------------------------------------------------------------
def freeze_panes(ws: OoxElement, ref: str | None) -> None:
    """Freeze rows above and columns left of ``ref`` ("B2"); None unfreezes."""
    from .spreadsheet import split_ref

    views = ws.first("x:sheetViews")
    if views is None:
        views = x.SheetViews(x.SheetView(workbookViewId=0))
        ws.add(views)
    view = views.first("x:sheetView")
    assert view is not None
    for old in list(view.iterchildren(f"{{{X}}}pane", f"{{{X}}}selection")):
        view.remove(old)
    if ref is None:
        return
    col, row = split_ref(ref)
    xs, ys = col - 1, row - 1
    if not xs and not ys:
        return
    active = "bottomRight" if xs and ys else ("topRight" if xs else "bottomLeft")
    attrs: dict[str, Any] = {"topLeftCell": ref.upper(), "activePane": active, "state": "frozen"}
    if xs:
        attrs["xSplit"] = xs
    if ys:
        attrs["ySplit"] = ys
    view.add(x.Pane(**attrs), x.Selection(pane=active, activeCell=ref.upper(),
                                          sqref=ref.upper()))


def frozen_panes(ws: OoxElement) -> str | None:
    views = ws.first("x:sheetViews")
    view = views.first("x:sheetView") if views is not None else None
    pane = view.first("x:pane") if view is not None else None
    if pane is None or pane.get("state") not in ("frozen", "frozenSplit"):
        return None
    if pane.get("topLeftCell"):
        return pane.get("topLeftCell")
    from .spreadsheet import cell_ref

    return cell_ref(int(float(pane.get("xSplit", "0"))) + 1, int(float(pane.get("ySplit", "0"))) + 1)


# -- defined names and autofilter -----------------------------------------------------------------
def define_name(book: SpreadsheetDocument, name: str, refers_to: str, *,
                local_sheet: int | None = None, hidden: bool = False) -> OoxElement:
    if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", name) or \
            re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        raise ValueError(f"invalid name {name!r} (letters, digits, _ and .; not a cell reference)")
    names = book.workbook.first("x:definedNames")
    if names is None:
        names = x.DefinedNames()
        book.workbook.add(names)
    for existing in list(names):
        same_scope = existing.get("localSheetId") == (None if local_sheet is None
                                                      else str(local_sheet))
        if existing.get("name", "").lower() == name.lower() and same_scope:
            names.remove(existing)
    attrs: dict[str, Any] = {"name": name}
    if local_sheet is not None:
        attrs["localSheetId"] = local_sheet
    if hidden:
        attrs["hidden"] = True
    el = x.DefinedName(text=refers_to.lstrip("="), **attrs)
    names.add(el)
    return el


def defined_names(book: SpreadsheetDocument) -> dict[str, str]:
    """``{name: refers_to}``; sheet-scoped names as ``"Sheet!name"``; Excel's
    own hidden names (``_xlnm.*``) included."""
    names = book.workbook.first("x:definedNames")
    out: dict[str, str] = {}
    if names is None:
        return out
    sheets = book.tab_names
    for el in names:
        key = el.get("name") or ""
        if el.get("localSheetId") is not None:
            idx = int(el.get("localSheetId") or 0)
            key = f"{sheets[idx] if idx < len(sheets) else idx}!{key}"
        out[key] = el.text or ""
    return out


def auto_filter(book: SpreadsheetDocument, sheet: str, ref: str | None) -> None:
    ws = book.sheet(sheet)
    old = ws.first("x:autoFilter")
    if old is not None:
        ws.remove(old)
    idx = book.tab_names.index(sheet)
    names = book.workbook.first("x:definedNames")
    if names is not None:
        for el in list(names):
            if el.get("name") == "_xlnm._FilterDatabase" and el.get("localSheetId") == str(idx):
                names.remove(el)
    if ref is None:
        return
    norm = range_ref(*parse_range(ref))
    ws.add(x.AutoFilter(ref=norm))
    define_name(book, "_xlnm._FilterDatabase", absolute(sheet, norm), local_sheet=idx,
                hidden=True)


def auto_filter_range(ws: OoxElement) -> str | None:
    af = ws.first("x:autoFilter")
    return af.get("ref") if af is not None else None


# -- data validation -----------------------------------------------------------------------------
def add_validation(ws: OoxElement, ref: str, *, values: list[str] | None = None,
                   source: str | None = None, whole: tuple[int, int] | None = None,
                   decimal: tuple[float, float] | None = None, allow_blank: bool = True,
                   prompt: str | None = None, error: str | None = None) -> OoxElement:
    """A drop-down list (``values`` or a ``source`` range like "$F$1:$F$5"),
    or a whole/decimal number between two bounds."""
    chosen = [k for k, v in (("values", values), ("source", source), ("whole", whole),
                             ("decimal", decimal)) if v is not None]
    if len(chosen) != 1:
        raise ValueError("give exactly one of values=, source=, whole=, decimal=")
    sqref = " ".join(range_ref(*parse_range(r)) for r in ref.split())
    attrs: dict[str, Any] = {"sqref": sqref, "allowBlank": allow_blank}
    kids: list[OoxElement] = []
    if values is not None:
        joined = ",".join(values)
        if len(joined) > 255 or any("," in v for v in values):
            raise ValueError("list items: no commas, 255 characters in all (use source=)")
        attrs["type"] = "list"
        kids.append(x.Formula1(text=f'"{joined}"'))
    elif source is not None:
        attrs["type"] = "list"
        kids.append(x.Formula1(text=source.lstrip("=")))
    else:
        bounds = whole if whole is not None else decimal
        assert bounds is not None
        attrs["type"] = "whole" if whole is not None else "decimal"
        attrs["operator"] = "between"
        kids += [x.Formula1(text=str(bounds[0])), x.Formula2(text=str(bounds[1]))]
    attrs["showInputMessage"] = True
    attrs["showErrorMessage"] = True
    if prompt:
        attrs["prompt"] = prompt
    if error:
        attrs["error"] = error
    dv = x.DataValidation(*kids, **attrs)
    dvs = ws.first("x:dataValidations")
    if dvs is None:
        dvs = x.DataValidations()
        ws.add(dvs)
    dvs.add(dv)
    dvs.set("count", str(len(dvs)))
    return dv


def validations(ws: OoxElement) -> list[dict[str, Any]]:
    dvs = ws.first("x:dataValidations")
    out = []
    for dv in (dvs.iterchildren(f"{{{X}}}dataValidation") if dvs is not None else []):
        f1, f2 = dv.first("x:formula1"), dv.first("x:formula2")
        item: dict[str, Any] = {"range": dv.get("sqref"), "type": dv.get("type", "none")}
        t1 = f1.text if f1 is not None else None
        if item["type"] == "list" and t1:
            if t1.startswith('"'):
                item["values"] = t1.strip('"').split(",")
            else:
                item["source"] = t1
        else:
            item["min"] = t1
            item["max"] = f2.text if f2 is not None else None
        out.append(item)
    return out


def autofit_columns(book: SpreadsheetDocument, sheet: str, cols: str | None = None, *,
                    min_width: float = 8.43, max_width: float = 80.0, padding: float = 1.5) -> dict[str, float]:
    """Set column widths from the longest displayed text in each column (an
    estimate: characters of the default font, bold or larger text not
    weighted).  Excel shows ``####`` for numbers and dates that don't fit,
    so this is worth calling after filling a sheet.  Returns the widths set."""
    from .spreadsheet import column_letter, split_ref

    longest: dict[int, int] = {}
    merged = [parse_range(m) for m in merged_cells(book.sheet(sheet))]
    for ref in book.cells(sheet, dates=False):
        col, row = split_ref(ref)
        if any(c1 <= col <= c2 and r1 <= row <= r2 and c1 != c2 for c1, r1, c2, r2 in merged):
            continue  # merged across columns: don't widen one column for it
        text = book.display_text(sheet, ref)
        width = max((len(line) for line in text.splitlines()), default=0)
        longest[col] = max(longest.get(col, 0), width)
    if cols is not None:
        first, _, last = cols.partition(":")
        lo, hi = _col_index(first), _col_index(last or first)
        longest = {c: w for c, w in longest.items() if lo <= c <= hi}
    out = {}
    for col, chars in sorted(longest.items()):
        # whole characters: Excel snaps widths to pixels, whole numbers stay exact
        width = float(min(max(math.ceil(chars + padding), min_width), max_width))
        letter = column_letter(col)
        set_column_width(book.sheet(sheet), letter, width)
        out[letter] = width
    return out
