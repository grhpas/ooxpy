"""Excel print setup and sheet options.

    book.page_setup("Sales", orientation="landscape", paper="A4", fit_width=1,
                    print_area="A1:F40", repeat_rows="1:1",
                    header="Quarterly report", footer={"right": "Page {page} of {pages}"})
    book.sheet_options("Sales", tab_color="2F5496")
    book.sheet_options("Lookup", hidden=True)
    book.protect_sheet("Sales", password="secret")

Header and footer text takes ``{page}``, ``{pages}``, ``{date}``, ``{time}``,
``{sheet}`` and ``{file}``; a plain string goes in the centre, a dict places
"left", "center" and "right".  Sheet protection uses Excel's legacy password
hash (what Excel 2007 writes): it stops casual edits, not a determined reader.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import x

if TYPE_CHECKING:
    from .spreadsheet import SpreadsheetDocument

X = URI["x"]
PAPER = {"letter": 1, "legal": 5, "a3": 8, "a4": 9, "a5": 11, "b5": 13, "tabloid": 3}
_TOKENS = {"page": "&P", "pages": "&N", "date": "&D", "time": "&T", "sheet": "&A", "file": "&F"}
_TOKENS_BACK = {v: "{" + k + "}" for k, v in _TOKENS.items()}
MARGINS = {"left": 0.7, "right": 0.7, "top": 0.75, "bottom": 0.75, "header": 0.3, "footer": 0.3}


def _hf_code(spec: str | dict[str, str]) -> str:
    """Friendly header/footer text -> Excel's &L&C&R code."""
    parts = spec if isinstance(spec, dict) else {"center": spec}
    unknown = set(parts) - {"left", "center", "right"}
    if unknown:
        raise ValueError(f"header/footer sections are left, center, right: {sorted(unknown)}")
    out = ""
    for key, code in (("left", "&L"), ("center", "&C"), ("right", "&R")):
        text = parts.get(key)
        if not text:
            continue
        text = text.replace("&", "&&")
        for name, token in _TOKENS.items():
            text = text.replace("{" + name + "}", token)
        out += code + text
    return out


def _hf_text(code: str | None) -> dict[str, str]:
    out: dict[str, str] = {}
    if not code:
        return out
    for key, text in re.findall(r"&([LCR])((?:[^&]|&&|&[^LCR])*)", code):
        text = text.replace("&&", "\0")  # a literal "&" first, so "R&&D" isn't a date code
        for token, name in _TOKENS_BACK.items():
            text = text.replace(token, name)
        out[{"L": "left", "C": "center", "R": "right"}[key]] = text.replace("\0", "&")
    return out


def _sheet_pr(ws: OoxElement) -> OoxElement:
    pr = ws.first("x:sheetPr")
    if pr is None:
        pr = x.SheetPr()
        ws.add(pr)
    return pr


def _rows_cols(spec: str | tuple[int, int] | None, kind: str) -> str | None:
    if spec is None:
        return None
    if isinstance(spec, tuple):
        a, b = spec
    else:
        a, _, b = spec.partition(":")
        b = b or a
    if kind == "rows":
        return f"${int(a)}:${int(b)}"
    return f"${str(a).upper()}:${str(b).upper()}"


def page_setup(book: SpreadsheetDocument, sheet: str, *, orientation: str | None = None,
               paper: str | None = None, fit_width: int | None = None,
               fit_height: int | None = None, scale: int | None = None,
               margins: dict[str, float] | None = None, print_area: str | None = None,
               repeat_rows: str | tuple[int, int] | None = None,
               repeat_cols: str | None = None, header: str | dict | None = None,
               footer: str | dict | None = None, gridlines: bool | None = None,
               center_horizontally: bool | None = None,
               center_vertically: bool | None = None) -> None:
    from .xlsheet import absolute, define_name

    ws = book.sheet(sheet)
    idx = book.tab_names.index(sheet)
    if gridlines is not None or center_horizontally is not None or center_vertically is not None:
        po = ws.first("x:printOptions")
        if po is None:
            po = x.PrintOptions()
            ws.add(po)
        for attr, val in (("gridLines", gridlines), ("horizontalCentered", center_horizontally),
                          ("verticalCentered", center_vertically)):
            if val is not None:
                po.set(attr, "1" if val else "0")
    if margins is not None or ws.first("x:pageMargins") is None:
        m: dict[str, Any] = dict(MARGINS)
        old = ws.first("x:pageMargins")
        if old is not None:
            m.update({k: float(v) for k, v in old.attrib.items() if k in MARGINS})
        unknown = set(margins or {}) - set(MARGINS)
        if unknown:
            raise ValueError(f"margins are {list(MARGINS)} (inches): {sorted(unknown)}")
        m.update(margins or {})
        ws.add(x.PageMargins(**m), replace=True)
    ps = ws.first("x:pageSetup")
    if ps is None:
        ps = x.PageSetup()
        ws.add(ps)
    if orientation is not None:
        if orientation not in ("portrait", "landscape"):
            raise ValueError("orientation is 'portrait' or 'landscape'")
        ps.set("orientation", orientation)
    if paper is not None:
        if paper.lower() not in PAPER:
            raise ValueError(f"paper is one of {list(PAPER)}")
        ps.set("paperSize", str(PAPER[paper.lower()]))
    if scale is not None:
        ps.set("scale", str(int(scale)))
    if fit_width is not None or fit_height is not None:
        sp = _sheet_pr(ws)
        sp.add(x.PageSetUpPr(fitToPage=True), replace=True)
        ps.set("fitToWidth", str(int(fit_width or 0)))
        ps.set("fitToHeight", str(int(fit_height or 0)))
    if header is not None or footer is not None:
        hf = ws.first("x:headerFooter")
        if hf is None:
            hf = x.HeaderFooter()
            ws.add(hf)
        if header is not None:
            hf.add(x.OddHeader(text=_hf_code(header)), replace=True)
        if footer is not None:
            hf.add(x.OddFooter(text=_hf_code(footer)), replace=True)
    if print_area is not None:
        define_name(book, "_xlnm.Print_Area", absolute(sheet, print_area), local_sheet=idx)
    if repeat_rows is not None or repeat_cols is not None:
        from .xlsheet import absolute as _abs

        quoted = _abs(sheet, "A1").split("!")[0]
        targets = [f"{quoted}!{v}" for v in (_rows_cols(repeat_cols, "cols"),
                                             _rows_cols(repeat_rows, "rows")) if v]
        define_name(book, "_xlnm.Print_Titles", ",".join(targets), local_sheet=idx)


def page_setup_info(book: SpreadsheetDocument, sheet: str) -> dict[str, Any]:
    ws = book.sheet(sheet)
    out: dict[str, Any] = {}
    ps = ws.first("x:pageSetup")
    if ps is not None:
        out["orientation"] = ps.get("orientation", "default")
        size = int(ps.get("paperSize", "1"))
        out["paper"] = next((k for k, v in PAPER.items() if v == size), size)
        if ps.get("scale"):
            out["scale"] = int(ps.get("scale"))
        pr = ws.first("x:sheetPr")
        fit = pr.first("x:pageSetUpPr") if pr is not None else None
        if fit is not None and fit.get("fitToPage") in ("1", "true"):
            out["fit_width"] = int(ps.get("fitToWidth", "1"))
            out["fit_height"] = int(ps.get("fitToHeight", "1"))
    m = ws.first("x:pageMargins")
    if m is not None:
        out["margins"] = {k: float(v) for k, v in m.attrib.items() if k in MARGINS}
    hf = ws.first("x:headerFooter")
    if hf is not None:
        head, foot = hf.first("x:oddHeader"), hf.first("x:oddFooter")
        out["header"] = _hf_text(head.text if head is not None else None)
        out["footer"] = _hf_text(foot.text if foot is not None else None)
    po = ws.first("x:printOptions")
    if po is not None:
        out["gridlines"] = po.get("gridLines") in ("1", "true")
    names = book.defined_names()
    for key, label in (("_xlnm.Print_Area", "print_area"), ("_xlnm.Print_Titles", "print_titles")):
        val = names.get(f"{sheet}!{key}")
        if val:
            out[label] = val
    return out


# -- sheet options ------------------------------------------------------------------------
def excel_password_hash(password: str) -> str:
    """Excel's legacy 16-bit sheet/workbook password hash (ECMA-376), as hex."""
    h = 0
    for i, ch in enumerate(password, start=1):
        value = ord(ch) << i
        h ^= (value & 0x7FFF) | (value >> 15)
    h ^= len(password)
    h ^= 0xCE4B
    return f"{h:04X}"


def sheet_options(book: SpreadsheetDocument, sheet: str, *, tab_color: str | None = None,
                  hidden: bool | str | None = None, zoom: int | None = None,
                  gridlines: bool | None = None) -> None:
    """``tab_color`` (hex, "" removes), ``hidden`` (True, False or "very" --
    hidden from Excel's Unhide list too), ``zoom`` (percent), ``gridlines``
    (shown on screen)."""
    from .xlstyles import _rgb

    ws = book.sheet(sheet)
    if tab_color is not None:
        pr = _sheet_pr(ws)
        old = pr.first("x:tabColor")
        if old is not None:
            pr.remove(old)
        if tab_color:
            pr.add(x.TabColor(rgb=_rgb(tab_color)))
    if hidden is not None:
        entry = next(e for e in book._sheet_entries() if e.get("name") == sheet)
        state = {True: "hidden", False: None, "very": "veryHidden"}[hidden]
        if state:
            visible = [e for e in book._sheet_entries()
                       if e.get("state") in (None, "visible") and e is not entry]
            if not visible:
                raise ValueError("a workbook needs at least one visible sheet")
            entry.set("state", state)
            # the active and selected tab must be a visible sheet
            idx = book.tab_names.index(sheet)
            views = book.workbook.first("x:bookViews")
            wbv = views.first("x:workbookView") if views is not None else None
            active = int(wbv.get("activeTab", "0")) if wbv is not None else 0
            if active == idx:
                first_visible = book.tab_names.index(visible[0].get("name"))
                if views is None:
                    views = x.BookViews(x.WorkbookView())
                    book.workbook.add(views)
                wv = views.first("x:workbookView")
                assert wv is not None
                wv.set("activeTab", str(first_visible))
                if first_visible:
                    wv.set("firstSheet", str(min(first_visible, int(wv.get("firstSheet", "0")))))
            for view in ws.iter(f"{{{X}}}sheetView"):
                view.attrib.pop("tabSelected", None)
        elif "state" in entry.attrib:
            del entry.attrib["state"]
    if zoom is not None or gridlines is not None:
        views = ws.first("x:sheetViews")
        if views is None:
            views = x.SheetViews(x.SheetView(workbookViewId=0))
            ws.add(views)
        view = views.first("x:sheetView")
        assert view is not None
        if zoom is not None:
            view.set("zoomScale", str(int(zoom)))
        if gridlines is not None:
            view.set("showGridLines", "1" if gridlines else "0")


def sheet_options_info(book: SpreadsheetDocument, sheet: str) -> dict[str, Any]:
    ws = book.sheet(sheet)
    entry = next(e for e in book._sheet_entries() if e.get("name") == sheet)
    pr = ws.first("x:sheetPr")
    tab = pr.first("x:tabColor") if pr is not None else None
    state = entry.get("state", "visible")
    prot = ws.first("x:sheetProtection")
    views = ws.first("x:sheetViews")
    view = views.first("x:sheetView") if views is not None else None
    return {
        "tab_color": (tab.get("rgb") or "")[-6:] or None if tab is not None else None,
        "hidden": {"visible": False, "hidden": True, "veryHidden": "very"}.get(state, False),
        "protected": prot is not None and prot.get("sheet") in ("1", "true"),
        "zoom": int(view.get("zoomScale", "100")) if view is not None else 100,
        "gridlines": view.get("showGridLines", "1") not in ("0", "false") if view is not None else True,
    }


def protect_sheet(book: SpreadsheetDocument, sheet: str, password: str | None = None, *,
                  allow: tuple[str, ...] = ()) -> None:
    """Protect the sheet's cells (Review > Protect Sheet).  ``allow``: actions
    still permitted, e.g. ("sort", "autoFilter", "formatCells")."""
    ws = book.sheet(sheet)
    attrs: dict[str, Any] = {"sheet": True, "objects": True, "scenarios": True}
    if password:
        attrs["password"] = excel_password_hash(password)
    for action in allow:
        attrs[action] = False  # these attributes say "locked"; False allows
    ws.add(x.SheetProtection(**attrs), replace=True)


def unprotect_sheet(book: SpreadsheetDocument, sheet: str) -> None:
    ws = book.sheet(sheet)
    old = ws.first("x:sheetProtection")
    if old is not None:
        ws.remove(old)
