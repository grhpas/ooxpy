"""Charts on worksheets (Insert > Chart), with series taken from a cell range.

    book.add_chart("Sales", "column", "A1:C5", anchor="E2", title="Sales",
                   y_title="EUR k", legend="bottom")
    book.add_chart("Sales", "pie", "A1:B5", anchor="E20", labels="percent")
    book.charts("Sales")   # [{"anchor": "E2", "type": "column", "title": "Sales", ...}]

The range's first row holds the series names and its first column the
categories (for scatter: the x values); ``series_in="rows"`` reads series
along rows instead.  The chart refers to the cells (``Sales!$B$2:$B$5``) and
caches their current values, so it shows without recalculation.  Size in EMU
(``oox.units``); the chart covers the cells under it, computed from the
sheet's column widths and row heights.  Other options are those of
``oox.charts.build_chart``.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from . import drawing
from .charts import CHART_URI, RT_CHART, Series, build_chart, chart_info
from .element import OoxElement
from .ns import a, c, x, xdr
from .units import Cm

if TYPE_CHECKING:
    from .spreadsheet import SpreadsheetDocument

EMU_PER_PX = 9525
DEFAULT_COL_PX = 64      # 8.43 characters of Calibri 11
DEFAULT_ROW_PT = 15.0


def quote_sheet(name: str) -> str:
    """Sheet name as it appears in a reference: quoted unless it is a plain word."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name) and not re.fullmatch(
            r"[A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def _ref(sheet: str, c1: int, r1: int, c2: int, r2: int) -> str:
    """``Sheet!$A$1:$B$2`` (or a single cell)."""
    from .spreadsheet import cell_ref

    start, end = _abs(cell_ref(c1, r1)), _abs(cell_ref(c2, r2))
    return f"{quote_sheet(sheet)}!{start}" + (f":{end}" if (c1, r1) != (c2, r2) else "")


def _abs(ref: str) -> str:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    assert m
    return f"${m.group(1)}${m.group(2)}"


def _col_px(book: SpreadsheetDocument, sheet: str, col: int) -> int:
    from .spreadsheet import column_letter

    chars = book.column_width(sheet, column_letter(col))
    return DEFAULT_COL_PX if chars is None else int(chars * 7 + 5)


def _row_emu(book: SpreadsheetDocument, sheet: str, row: int) -> int:
    pt = book.row_height(sheet, row)
    return int((DEFAULT_ROW_PT if pt is None else pt) * 12700)


def _span(sizes, start: int, length: int) -> tuple[int, int]:
    """(index, offset) where ``length`` EMU from the start of ``start`` ends."""
    i = start
    while True:
        size = sizes(i)
        if length < size:
            return i, length
        length -= size
        i += 1


def series_from_range(book: SpreadsheetDocument, sheet: str, data: str,
                      series_in: str = "columns") -> tuple[list, str, list[Series]]:
    """(categories, categories_ref, series) for a header-row/label-column range."""
    from .xlsheet import parse_range

    c1, r1, c2, r2 = parse_range(data)
    if c2 - c1 < 1 or r2 - r1 < 1:
        raise ValueError(f"chart data {data!r}: needs a header row, a label column and values")

    def val(col, row):
        from .spreadsheet import cell_ref

        return book.value(sheet, cell_ref(col, row))

    def num(v):
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

    if series_in == "columns":
        cats = [val(c1, r) for r in range(r1 + 1, r2 + 1)]
        cref = _ref(sheet, c1, r1 + 1, c1, r2)
        series = [Series(None if val(col, r1) is None else str(val(col, r1)),
                         [num(val(col, r)) for r in range(r1 + 1, r2 + 1)],
                         name_ref=_ref(sheet, col, r1, col, r1),
                         values_ref=_ref(sheet, col, r1 + 1, col, r2))
                  for col in range(c1 + 1, c2 + 1)]
    elif series_in == "rows":
        cats = [val(col, r1) for col in range(c1 + 1, c2 + 1)]
        cref = _ref(sheet, c1 + 1, r1, c2, r1)
        series = [Series(None if val(c1, r) is None else str(val(c1, r)),
                         [num(val(col, r)) for col in range(c1 + 1, c2 + 1)],
                         name_ref=_ref(sheet, c1, r, c1, r),
                         values_ref=_ref(sheet, c1 + 1, r, c2, r))
                  for r in range(r1 + 1, r2 + 1)]
    else:
        raise ValueError("series_in must be 'columns' or 'rows'")
    for s in series:
        if s.name_ref and s.name is None:
            s.name_ref = None
    return cats, cref, series


def add_chart(book: SpreadsheetDocument, sheet: str, kind: str, data: str, *,
              anchor: str = "E2", width: int = Cm(15), height: int = Cm(7.5),
              series_in: str = "columns", name: str | None = None, **options) -> OoxElement:
    """Add a chart of ``data`` with its top-left corner at ``anchor``; returns
    its ``c:chartSpace``."""
    from .spreadsheet import split_ref

    ws = book._ws(sheet)
    ws_part = book.part_of(ws)
    cats, cref, series = series_from_range(book, sheet, data, series_in)
    space = build_chart(kind, cats, series, categories_ref=cref, **options)
    found = book.related_parts(ws_part, "DrawingsPart")
    if found:
        dpart = found[0]
    else:
        dpart, drid = book.add_part("DrawingsPart", ws_part, element=xdr.WsDr())
        ws.add(x.Drawing(r_id=drid))
    _, rid = book.add_part("ChartPart", dpart, element=space, folder="/xl/charts")
    col, row = split_ref(anchor)
    to_col, col_off = _span(lambda i: _col_px(book, sheet, i) * EMU_PER_PX, col, width)
    to_row, row_off = _span(lambda i: _row_emu(book, sheet, i), row, height)
    n = drawing.max_shape_id(dpart.element) + 1
    frame = xdr.GraphicFrame(
        xdr.NvGraphicFramePr(xdr.CNvPr(id=n, name=name or f"Chart {n}"), xdr.CNvGraphicFramePr()),
        xdr.Xfrm(a.Off(x=0, y=0), a.Ext(cx=0, cy=0)),
        a.Graphic(a.GraphicData(c.Chart(r_id=rid), uri=CHART_URI)),
        macro="")
    dpart.element.add(xdr.TwoCellAnchor(
        xdr.From(xdr.Col(text=str(col - 1)), xdr.ColOff(text="0"),
                 xdr.Row(text=str(row - 1)), xdr.RowOff(text="0")),
        xdr.To(xdr.Col(text=str(to_col - 1)), xdr.ColOff(text=str(col_off)),
               xdr.Row(text=str(to_row - 1)), xdr.RowOff(text=str(row_off))),
        frame, xdr.ClientData(), editAs="oneCell"))
    return space


XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
C_CHART = "{http://schemas.openxmlformats.org/drawingml/2006/chart}chart"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def add_chart_sheet(book: SpreadsheetDocument, name: str, kind: str, sheet: str, data: str, *,
                    series_in: str = "columns", **options) -> OoxElement:
    """A chart sheet (its own tab) showing a chart of ``data`` on ``sheet``."""
    book._check_tab_name(name)
    cats, cref, series = series_from_range(book, sheet, data, series_in)
    space = build_chart(kind, cats, series, categories_ref=cref, **options)
    root = x.Chartsheet(x.SheetPr(), x.SheetViews(x.SheetView(zoomToFit=True, workbookViewId=0)),
                        x.PageMargins(left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3,
                                      footer=0.3))
    cs_part, rid = book.add_part("ChartsheetPart", element=root)
    dpart, drid = book.add_part("DrawingsPart", cs_part, element=xdr.WsDr())
    root.add(x.Drawing(r_id=drid))
    _, crid = book.add_part("ChartPart", dpart, element=space, folder="/xl/charts")
    frame = xdr.GraphicFrame(
        xdr.NvGraphicFramePr(xdr.CNvPr(id=2, name="Chart 1"),
                             xdr.CNvGraphicFramePr(a.GraphicFrameLocks(noGrp=True))),
        xdr.Xfrm(a.Off(x=0, y=0), a.Ext(cx=0, cy=0)),
        a.Graphic(a.GraphicData(c.Chart(r_id=crid), uri=CHART_URI)),
        macro="")
    # Excel's size for a chart sheet on A4/Letter; it zooms to fit the window
    dpart.element.add(xdr.AbsoluteAnchor(xdr.Pos(x=0, y=0), xdr.Ext(cx=9294091, cy=6070023),
                                         frame, xdr.ClientData()))
    sheet_id = max((int(e.get("sheetId")) for e in book._sheet_entries()), default=0) + 1
    book.workbook.need("x:sheets").add(x.Sheet(name=name, sheetId=sheet_id, r_id=rid))
    return space


def charts(book: SpreadsheetDocument, sheet: str) -> list[dict]:
    """Charts on a sheet or chart sheet in drawing order: ``chart_info`` plus
    "anchor" (a worksheet's top-left cell; None on a chart sheet) and "name"."""
    from .spreadsheet import cell_ref

    ws = book.chart_sheet(sheet) if sheet in book.chart_sheet_names else book._ws(sheet)
    out = []
    for dpart in book.related_parts(book.part_of(ws), "DrawingsPart"):
        for anchor in dpart.element:
            if not isinstance(anchor.tag, str):
                continue
            ref = anchor.find(".//" + C_CHART)
            rel = dpart.rels.items.get(ref.get(R_ID, "")) if ref is not None else None
            if rel is None or rel.reltype != RT_CHART:
                continue
            target = dpart.resolve(rel.target)
            if target not in book.package.parts:
                continue
            info = chart_info(book.package.parts[target].element)
            _fill_from_cells(book, info)
            frm = anchor.find(XDR + "from")
            if frm is not None:
                info["anchor"] = cell_ref(int(frm.findtext(XDR + "col") or 0) + 1,
                                          int(frm.findtext(XDR + "row") or 0) + 1)
            if frm is None:
                info["anchor"] = None
            pr = anchor.find(".//" + XDR + "cNvPr")
            info["name"] = pr.get("name") if pr is not None else None
            out.append(info)
    return out


def _cells(book: SpreadsheetDocument, ref: str | None) -> list | None:
    """Values of ``Sheet!$A$1:$A$4`` (one row or column) in this workbook, or None."""
    from .spreadsheet import cell_ref
    from .xlsheet import parse_range

    if not ref or "!" not in ref or ref.startswith("["):
        return None
    sheet, _, rng = ref.rpartition("!")
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    if sheet not in book.sheet_names:          # another workbook, or a chart sheet
        return None
    try:
        c1, r1, c2, r2 = parse_range(rng.replace("$", ""))
    except ValueError:
        return None
    if (c2 - c1 + 1) * (r2 - r1 + 1) > 100_000:
        return None
    return [book.value(sheet, cell_ref(col, row))
            for row in range(r1, r2 + 1) for col in range(c1, c2 + 1)]


def _fill_from_cells(book: SpreadsheetDocument, info: dict) -> None:
    """Values, categories and names a chart has no cache for (openpyxl writes
    references only) are read from the cells they refer to."""
    for ser in info["series"]:
        if not any(v is not None for v in ser["values"]):
            got = _cells(book, ser["values_ref"])
            if got is not None:
                ser["values"] = [v if isinstance(v, (int, float)) and not isinstance(v, bool)
                                 else None for v in got]
        if not ser["categories"]:
            got = _cells(book, ser["categories_ref"])
            if got is not None:
                ser["categories"] = got
        if ser["name"] is None:
            got = _cells(book, ser["name_ref"])
            if got:
                ser["name"] = None if got[0] is None else str(got[0])
    if info["series"]:
        info["categories"] = info["series"][0]["categories"]
