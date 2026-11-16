"""SpreadsheetML (.xlsx): the ``OpenDocumentSpreadsheet()`` equivalent.

    from oox.spreadsheet import SpreadsheetDocument, Formula

    book = SpreadsheetDocument.new()            # one sheet, "Sheet1"
    book.append_row("Sheet1", ["Item", "Qty", "Price", "Total"])
    book.append_row("Sheet1", ["Pens", 3, 1.5, Formula("B2*C2")])
    book.add_picture("Sheet1", "logo.png", anchor="F2")
    book.save("report.xlsx")

    book = SpreadsheetDocument.open("from_excel.xlsx")   # reading works on any producer
    book.value("Sheet1", "B2")          # 3, "text", True, 1.5, CellError("#DIV/0!") ...
    book.values("Sheet1")               # [[...], [...]] dense, from A1
    book.set_value("Sheet1", "D7", 42)  # keeps the cell's existing style

Strings are written as inline strings (``t="inlineStr"``), which every
consumer reads and which avoid managing a shared-string table.  A
shared-strings and number-format/style manager is roadmap phase 2.
"""

from __future__ import annotations

import datetime as _dt
import re
from collections.abc import Iterable
from dataclasses import dataclass

from . import drawing
from .document import RT_OFFICE_DOCUMENT, OoxDocument
from .element import OoxElement
from .errors import PackageError
from .grammar import URI
from .ns import x, xdr
from .opc import Package

X = URI["x"]

CT_SML_MAIN = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    ".xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}

_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_BAD_SHEET_CHARS = set('[]:*?/\\')
CT_CHARTSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"


@dataclass(frozen=True)
class Formula:
    """A cell formula, without the leading ``=``: ``Formula("SUM(A1:A3)")``.

    ``cached`` is the result to store with it (a number, str or bool), so
    readers that don't calculate see a value; Excel recalculates anyway.
    """

    expression: str
    cached: object = None

    def __repr__(self) -> str:
        if self.cached is None:
            return f"Formula(expression={self.expression!r})"
        return f"Formula(expression={self.expression!r}, cached={self.cached!r})"


class CellError(str):
    """An error value read from a cell, e.g. ``CellError("#DIV/0!")``."""

    def __repr__(self) -> str:
        return f"CellError({str(self)!r})"


_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")
_INT = re.compile(r"^-?\d+$")


def rich_text(container: OoxElement) -> str:
    """Plain text of an ``x:si`` / ``x:is``: direct ``x:t`` plus run texts,
    skipping phonetic runs (``x:rPh``); decodes Excel's ``_xHHHH_`` escapes."""
    parts: list[str] = []
    for child in container:
        local = child.tag.rpartition("}")[2] if isinstance(child.tag, str) else ""
        if local == "t":
            parts.append(child.text or "")
        elif local == "r":
            t = child.first("x:t")
            if t is not None:
                parts.append(t.text or "")
    return _ESCAPE.sub(lambda m: chr(int(m.group(1), 16)), "".join(parts))


def _number(text: str) -> int | float:
    return int(text) if _INT.match(text) else float(text)


# -- A1 references -------------------------------------------------------------
def column_letter(col: int) -> str:
    """1 -> "A", 27 -> "AA"."""
    if not 1 <= col <= 16384:
        raise ValueError(f"column out of range: {col}")
    out = ""
    while col:
        col, rem = divmod(col - 1, 26)
        out = chr(65 + rem) + out
    return out


def cell_ref(col: int, row: int) -> str:
    return f"{column_letter(col)}{row}"


def split_ref(ref: str) -> tuple[int, int]:
    """"B3" -> (2, 3)  (1-based column, row)."""
    m = _REF.match(ref)
    if not m:
        raise ValueError(f"not an A1 cell reference: {ref!r}")
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + ord(ch) - 64
    return col, int(m.group(2))


def excel_serial(value: _dt.date | _dt.datetime) -> float:
    """Date/datetime -> Excel serial number (1900 date system)."""
    if isinstance(value, _dt.datetime):
        delta = value - _dt.datetime(1899, 12, 30, tzinfo=value.tzinfo)
    else:
        delta = _dt.datetime.combine(value, _dt.time()) - _dt.datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6


def from_excel_serial(serial: float) -> _dt.datetime:
    """Excel serial number (1900 date system) -> datetime."""
    return _dt.datetime(1899, 12, 30) + _dt.timedelta(days=serial)


def cell(ref: str, value: object) -> OoxElement:
    """Build ``<x:c>`` for a Python value.

    ``None`` -> empty cell, ``bool`` -> boolean, ``int``/``float`` -> number,
    ``str`` -> inline string, ``Formula`` -> formula.  Dates need a date
    number format, which lives in the workbook's stylesheet: use the
    document's ``set_value`` / ``append_row``, which add one.
    """
    split_ref(ref)  # validate
    if value is None:
        return x.C(r=ref)
    if isinstance(value, bool):
        return x.C(x.V(text="1" if value else "0"), r=ref, t="b")
    if isinstance(value, (int, float)):
        return x.C(x.V(text=repr(value) if isinstance(value, float) else str(value)), r=ref)
    if isinstance(value, Formula):
        c = x.C(x.F(text=value.expression), r=ref)
        v = value.cached
        if isinstance(v, bool):
            c.set("t", "b")
            c.add(x.V(text="1" if v else "0"))
        elif isinstance(v, (int, float)):
            c.add(x.V(text=repr(v) if isinstance(v, float) else str(v)))
        elif isinstance(v, str):
            c.set("t", "str")
            c.add(x.V(text=v))
        return c
    if isinstance(value, str):
        return x.C(x.Is(x.T(value)), r=ref, t="inlineStr")
    if isinstance(value, (_dt.date, _dt.datetime)):
        raise TypeError("dates need a date style: use book.set_value()/append_row(), "
                        "or pass excel_serial(d)")
    raise TypeError(f"unsupported cell value type: {type(value).__name__}")


def row(index: int, values: Iterable[object], start_col: int = 1) -> OoxElement:
    """Build ``<x:row r=index>`` with one cell per value (``None`` cells are skipped)."""
    r = x.Row(r=index)
    for i, v in enumerate(values):
        if v is not None:
            r.add(cell(cell_ref(start_col + i, index), v))
    return r


# -- document --------------------------------------------------------------------
class SpreadsheetDocument(OoxDocument):
    MAIN_CONTENT_TYPES = CT_SML_MAIN

    @classmethod
    def new(cls, sheet: str | None = "Sheet1", kind: str = ".xlsx") -> SpreadsheetDocument:
        """A new workbook with one worksheet named ``sheet`` (Excel refuses a workbook
        without one).  ``kind`` ".xlsm", ".xltx" ... sets the content type."""
        pkg = Package()
        part = pkg.add_part("/xl/workbook.xml", CT_SML_MAIN[kind],
                            element=x.Workbook(x.Sheets()))
        pkg.relate(None, part, RT_OFFICE_DOCUMENT)
        doc = cls(pkg)
        if sheet:
            doc.add_sheet(sheet)
        return doc

    @property
    def workbook(self) -> OoxElement:
        """The root ``x:workbook`` element of the main part."""
        return self.main_part.element

    # odfpy-flavoured alias: doc.spreadsheet in odfpy
    spreadsheet = workbook

    def _sheet_entries(self) -> list[OoxElement]:
        sheets = self.workbook.first("x:sheets")
        return list(sheets) if sheets is not None else []

    def _tab_part(self, entry: OoxElement):
        rid = entry.get_attr("r:id")
        rel = self.main_part.rels.items.get(rid or "")
        part = self.package.parts.get(self.main_part.resolve(rel.target)) if rel else None
        if part is None:
            raise PackageError(f"sheet {entry.get('name')!r}: broken relationship {rid!r}")
        return part

    def _is_chart_sheet(self, entry: OoxElement) -> bool:
        try:
            return self._tab_part(entry).content_type == CT_CHARTSHEET
        except PackageError:
            return False

    @property
    def tab_names(self) -> list[str]:
        """Every tab in order: worksheets and chart sheets (``localSheetId``
        and the active tab count these)."""
        return [s.get("name") for s in self._sheet_entries()]

    @property
    def sheet_names(self) -> list[str]:
        """The worksheets, in tab order (chart sheets: ``chart_sheet_names``)."""
        return [s.get("name") for s in self._sheet_entries() if not self._is_chart_sheet(s)]

    @property
    def chart_sheet_names(self) -> list[str]:
        """Names of the chart sheets, in tab order (``tab_names`` has every tab)."""
        return [s.get("name") for s in self._sheet_entries() if self._is_chart_sheet(s)]

    def sheet(self, name: str) -> OoxElement:
        """Root ``<x:worksheet>`` of the named sheet."""
        for entry in self._sheet_entries():
            if entry.get("name") == name:
                part = self._tab_part(entry)
                if part.content_type == CT_CHARTSHEET:
                    raise KeyError(f"{name!r} is a chart sheet; use chart_sheet() or charts()")
                return part.element
        raise KeyError(name)

    def chart_sheet(self, name: str) -> OoxElement:
        """Root ``<x:chartsheet>`` of the named chart sheet."""
        for entry in self._sheet_entries():
            if entry.get("name") == name and self._is_chart_sheet(entry):
                return self._tab_part(entry).element
        raise KeyError(f"no chart sheet {name!r}")

    def add_chart_sheet(self, name: str, kind: str, sheet: str, data: str, *,
                        series_in: str = "columns", **options) -> OoxElement:
        """A tab holding one chart of ``data`` on ``sheet`` (a worksheet), as
        Excel's Move Chart > New sheet; options as ``add_chart``.  Returns the
        chart's ``c:chartSpace``."""
        from .xlcharts import add_chart_sheet

        return add_chart_sheet(self, name, kind, sheet, data, series_in=series_in, **options)

    @property
    def worksheets(self) -> list[OoxElement]:
        """The ``x:worksheet`` elements, in tab order (chart sheets excluded)."""
        return [self.sheet(n) for n in self.sheet_names]

    def _check_tab_name(self, name: str) -> None:
        if not name or len(name) > 31 or _BAD_SHEET_CHARS & set(name) or name[0] == "'":
            raise ValueError(f"invalid sheet name {name!r} (1-31 chars, none of []:*?/\\)")
        if name.lower() in (n.lower() for n in self.tab_names):
            raise ValueError(f"duplicate sheet name {name!r}")

    def add_sheet(self, name: str) -> OoxElement:
        """Append a worksheet; returns its ``<x:worksheet>`` root."""
        self._check_tab_name(name)
        part, rid = self.add_part("WorksheetPart", element=x.Worksheet(x.SheetData()))
        sheet_id = max((int(s.get("sheetId")) for s in self._sheet_entries()), default=0) + 1
        self.workbook.need("x:sheets").add(x.Sheet(name=name, sheetId=sheet_id, r_id=rid))
        return part.element

    def _ws(self, sheet: str | OoxElement) -> OoxElement:
        return self.sheet(sheet) if isinstance(sheet, str) else sheet

    def sheet_data(self, sheet: str | OoxElement) -> OoxElement:
        """The sheet's ``x:sheetData`` element (the rows)."""
        return self._ws(sheet).first("x:sheetData")

    def append_row(self, sheet: str | OoxElement, values: Iterable[object],
                   start_col: int = 1, *, style: int | dict | None = None) -> OoxElement:
        """Add a row after the last one; returns the ``<x:row>``.  Dates get a
        date format; ``style`` (an index from ``cell_style`` or a dict of its
        options) applies to every cell."""
        from .element import trusted

        data = self.sheet_data(sheet)
        last = _last_index(data, "r")
        index = self._style_index(style)   # validated once, outside the trusted block
        with trusted():
            r = x.Row(r=last + 1)
            for i, v in enumerate(values):
                if v is not None or index is not None:
                    r.add(self._cell(cell_ref(start_col + i, last + 1), v, index))
            data.add(r)
        return r

    # -- styles -----------------------------------------------------------------
    def cell_style(self, **fmt) -> int:
        """A cell style index for ``set_value(..., style=)``: number_format,
        bold, italic, underline, strike, size, font, color, fill (hex), border
        ("thin", "medium" ... or {"bottom": "double"}), border_color, align,
        valign, wrap, indent, rotation.  Equal styles share one index."""
        from .xlstyles import StyleSheet

        return StyleSheet(self).xf(**fmt)

    def _style_index(self, style: int | dict | None) -> int | None:
        if style is None or isinstance(style, int):
            return style
        return self.cell_style(**style)

    def _cell(self, ref: str, value: object, style: int | dict | None) -> OoxElement:
        """``cell()`` plus what needs the workbook: dates and styles."""
        from .xlstyles import DATE_FORMAT, DATETIME_FORMAT, TIME_FORMAT

        xf = self._style_index(style)
        if isinstance(value, (_dt.date, _dt.datetime, _dt.time)):
            if isinstance(value, _dt.datetime):
                fmt, value = DATETIME_FORMAT, excel_serial(value)
            elif isinstance(value, _dt.date):
                fmt, value = DATE_FORMAT, excel_serial(value)
            else:
                fmt = TIME_FORMAT
                value = (value.hour * 3600 + value.minute * 60 + value.second) / 86400
            if xf is None:
                xf = self.cell_style(number_format=fmt)
        c = cell(ref, value)
        if xf:
            c.set("s", str(xf))
        return c

    def style_range(self, sheet: str | OoxElement, ref: str, style: int | dict | None = None,
                    **fmt) -> None:
        """Apply a style to every cell in a range (empty cells are created)."""
        from .xlsheet import parse_range

        xf = self._style_index(style if style is not None else fmt)
        c1, r1, c2, r2 = parse_range(ref)
        for rw in range(r1, r2 + 1):
            for col in range(c1, c2 + 1):
                target = self.cell_element(sheet, cell_ref(col, rw))
                if target is None:
                    target = self.set_value(sheet, cell_ref(col, rw), None)
                if xf:
                    target.set("s", str(xf))
                elif "s" in target.attrib:
                    del target.attrib["s"]

    def cell_format(self, sheet: str | OoxElement, ref: str) -> dict:
        """The cell's formatting as ``cell_style`` options (number_format,
        bold, font, fill, border, align ...)."""
        from .xlstyles import StyleSheet

        c = self.cell_element(sheet, ref)
        return StyleSheet(self, create=False).describe(int(c.get("s", "0")) if c is not None else 0)

    def display_text(self, sheet: str | OoxElement, ref: str) -> str:
        """The text Excel shows in the cell (number format applied; formula
        cells show their cached result, if any)."""
        from .xlstyles import format_value

        c = self.cell_element(sheet, ref)
        if c is None:
            return ""
        raw = self._decode(c, self._shared_strings(), False)
        if isinstance(raw, CellError):
            return str(raw)
        return format_value(raw, self.cell_format(sheet, ref)["number_format"])

    # -- reading ----------------------------------------------------------------
    # -- read caches ----------------------------------------------------------------
    # Single-cell reads would otherwise scan the sheet's rows, decode the whole
    # shared-strings table and parse the styles each time.  Each cache keeps
    # the element it was built from and a fingerprint that is cheap to take
    # (counting 50,000 children costs half a millisecond in lxml, so the big
    # ones use their last child); the row index is also checked around each
    # lookup.  Edits -- through oox or directly on the tree -- rebuild a cache
    # on the next read.
    def _cache(self, name: str, element: OoxElement, fingerprint: tuple, build):
        caches = self.__dict__.setdefault("_read_caches", {})
        hit = caches.get(name)
        if hit is not None and hit[0] is element and hit[1] == fingerprint:
            return hit[2]
        value = build()
        caches[name] = (element, fingerprint, value)
        return value

    def _drop_cache(self, name: str) -> None:
        self.__dict__.get("_read_caches", {}).pop(name, None)

    def _shared_strings(self) -> list[str]:
        found = self.related_parts(self.main_part, "SharedStringTablePart")
        if not found:
            return []
        sst = found[0].element
        last = _last_child(sst, "si")
        return self._cache("sst", sst, (last, last.get("count") if last is not None else None),
                           lambda: [rich_text(si) for si in sst])

    def _row_index(self, data: OoxElement, fresh: bool = False):
        """``(sorted row numbers, {number: x:row})``, or None if a row has no ``r``."""
        def build():
            index: dict[int, OoxElement] = {}
            for row in data.iterchildren(f"{{{X}}}row"):
                r = row.get("r")
                if r is None:
                    return None
                index[int(r)] = row
            return sorted(index), index

        name = f"rows:{id(data)}"
        if fresh:
            self._drop_cache(name)
        return self._cache(name, data, (), build)

    def _find_row(self, data: OoxElement, rw: int) -> tuple[bool, OoxElement | None]:
        """``(known, row)``: ``known`` is False when the sheet can't be indexed."""
        import bisect

        tag = f"{{{X}}}row"

        def next_row(el):
            for sib in el.itersiblings():
                if sib.tag == tag:
                    return sib
            return None

        for attempt in (False, True):
            got = self._row_index(data, fresh=attempt)
            if got is None:
                return False, None
            keys, index = got
            row = index.get(rw)
            if row is not None:
                if row.getparent() is data and row.get("r") == str(rw):
                    return True, row
                continue
            # a miss is genuine if the cached neighbours are still neighbours
            at = bisect.bisect_left(keys, rw)
            before = index[keys[at - 1]] if at > 0 else None
            after = index[keys[at]] if at < len(keys) else None
            ok = True
            for key, el in ((keys[at - 1] if at > 0 else None, before),
                            (keys[at] if at < len(keys) else None, after)):
                if el is not None and (el.getparent() is not data or el.get("r") != str(key)):
                    ok = False
            if ok:
                if before is not None:
                    ok = next_row(before) is after
                else:
                    first = next(data.iterchildren(tag), None)
                    ok = first is after
            if ok:
                return True, None
        return False, None

    def cell_element(self, sheet: str | OoxElement, ref: str) -> OoxElement | None:
        """The ``x:c`` element at ``ref``, or None if the cell isn't stored."""
        col, rw = split_ref(ref)
        data = self.sheet_data(sheet)
        known, row = self._find_row(data, rw)
        if known:
            if row is None:
                return None
            for c_index, c in _indexed(row, "c"):
                if c_index == col:
                    return c
                if c_index > col:
                    break
            return None
        for r_index, r in _indexed(data, "r"):
            if r_index == rw:
                for c_index, c in _indexed(r, "c"):
                    if c_index == col:
                        return c
            if r_index > rw:
                break
        return None

    def value(self, sheet: str | OoxElement, ref: str, *, formulas: bool = False,
              dates: bool = True) -> object:
        """Python value of one cell (``None`` if empty or absent).

        Formula cells give their cached result; files written by oox (or
        openpyxl) have no cached result until Excel saves them, so you get
        ``None`` -- pass ``formulas=True`` to get ``Formula`` objects instead.
        Numbers in a date or time format come back as ``datetime`` (pass
        ``dates=False`` for the serial number).
        """
        c = self.cell_element(sheet, ref)
        if c is None:
            return None
        raw = self._decode(c, self._shared_strings(), formulas)
        return self._dated(c, raw, self._date_xfs()) if dates else raw

    def _date_xfs(self) -> set[int]:
        """cellXfs indexes whose number format is a date/time format."""
        from .xlstyles import StyleSheet, is_date_format

        found = self.related_parts(self.main_part, "WorkbookStylesPart")
        root = found[0].element if found else None

        def build() -> set[int]:
            st = StyleSheet(self, create=False)
            out: set[int] = set()
            if st.root is not None:
                xfs = st.root.first("x:cellXfs")
                for i, xf in enumerate(xfs if xfs is not None else []):
                    nid = int(xf.get("numFmtId", "0"))
                    if is_date_format(st.number_format_code(nid), nid):
                        out.add(i)
            return out

        if root is None:
            return build()
        xfs, fmts = root.first("x:cellXfs"), root.first("x:numFmts")
        fp = (xfs, len(xfs) if xfs is not None else -1, fmts, len(fmts) if fmts is not None else -1)
        return self._cache("date_xfs", root, fp, build)

    @staticmethod
    def _dated(c: OoxElement, value: object, date_xfs: set[int]) -> object:
        if isinstance(value, (int, float)) and not isinstance(value, bool) and \
                int(c.get("s", "0")) in date_xfs:
            when = from_excel_serial(float(value))
            # serials are floats: round away the noise, to the millisecond
            ms = round(when.microsecond / 1000)
            when = when.replace(microsecond=0) + _dt.timedelta(milliseconds=ms)
            return when.time() if 0 <= float(value) < 1 else when  # Excel's day 0 is a time
        return value

    def values(self, sheet: str | OoxElement, *, formulas: bool = False,
               dates: bool = True) -> list[list[object]]:
        """All values as a dense list of rows, starting at A1."""
        sparse = self.cells(sheet, formulas=formulas, dates=dates)
        if not sparse:
            return []
        coords = [split_ref(r) for r in sparse]
        width, height = max(c for c, _ in coords), max(r for _, r in coords)
        grid: list[list[object]] = [[None] * width for _ in range(height)]
        for ref, v in sparse.items():
            col, rw = split_ref(ref)
            grid[rw - 1][col - 1] = v
        return grid

    def cells(self, sheet: str | OoxElement, *, formulas: bool = False,
              dates: bool = True) -> dict[str, object]:
        """Non-empty cells as ``{"A1": value, ...}`` in sheet order."""
        sst = self._shared_strings()
        date_xfs = self._date_xfs() if dates else set()
        out: dict[str, object] = {}
        for r_index, r in _indexed(self.sheet_data(sheet), "r"):
            for c_index, c in _indexed(r, "c"):
                v = self._dated(c, self._decode(c, sst, formulas), date_xfs)
                if v is not None:
                    out[cell_ref(c_index, r_index)] = v
        return out

    @staticmethod
    def _decode(c: OoxElement, sst: list[str], formulas: bool) -> object:
        t = c.get("t", "n")
        f, v = c.first("x:f"), c.first("x:v")
        if formulas and f is not None and f.text:
            return Formula(f.text)
        if t == "inlineStr":
            is_ = c.first("x:is")
            return rich_text(is_) if is_ is not None else ""
        if v is None or v.text is None:
            return None
        text = v.text
        if t == "s":
            return sst[int(text)]
        if t == "b":
            return text == "1"
        if t == "e":
            return CellError(text)
        if t == "str":
            return _ESCAPE.sub(lambda m: chr(int(m.group(1), 16)), text)
        if t == "d":
            return _dt.datetime.fromisoformat(text)
        return _number(text)

    # -- editing ----------------------------------------------------------------
    def set_value(self, sheet: str | OoxElement, ref: str, value: object, *,
                  style: int | dict | None = None) -> OoxElement:
        """Write one cell, creating its row if needed and keeping row/column
        order.  Dates and times get a date format.  ``style``: an index from
        ``cell_style`` or a dict of its options; without it an existing
        cell's style is kept."""
        col, rw = split_ref(ref)
        data = self.sheet_data(sheet)
        new = self._cell(cell_ref(col, rw), value, style)
        target_row = None
        last = _last_index(data, "r")
        if rw > last:                     # fast path: writing below the data
            target_row = x.Row(r=rw)
            data.add(target_row)
        elif rw == last:
            target_row = _last_child(data, "row")
        for r_index, r in ([] if target_row is not None else _indexed(data, "r")):
            if r_index == rw:
                target_row = r
                break
            if r_index > rw:
                target_row = x.Row(r=rw)
                r.addprevious(target_row)
                break
        if target_row is None:
            target_row = x.Row(r=rw)
            data.add(target_row)
        if col > _last_index(target_row, "c"):   # fast path: right of the last cell
            target_row.add(new)
            return new
        for c_index, c in _indexed(target_row, "c"):
            if c_index == col:
                if c.get("s") is not None and new.get("s") is None:
                    new.set("s", c.get("s"))
                target_row.replace(c, new)
                return new
            if c_index > col:
                c.addprevious(new)
                return new
        target_row.add(new)
        return new

    # -- shared strings -----------------------------------------------------------
    def share_strings(self) -> int:
        """Move inline strings into the shared-strings table, as Excel stores
        them (smaller files; some readers expect it).  Equal strings share one
        entry.  Returns the number of cells converted."""
        import copy

        from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

        found = self.related_parts(self.main_part, "SharedStringTablePart")
        if found:
            sst = found[0].element
        else:
            sst = x.Sst()
            self.add_part("SharedStringTablePart", element=sst)
        index: dict[bytes, int] = {}
        for i, si in enumerate(sst):
            index.setdefault(b"".join(etree.tostring(c, method="c14n") for c in si), i)
        converted = 0
        for name in self.sheet_names:
            for c in self.sheet(name).iter(f"{{{X}}}c"):
                if c.get("t") != "inlineStr":
                    continue
                is_ = c.first("x:is")
                if is_ is None:
                    continue
                si = x.Si()
                for child in is_:
                    si.append(copy.deepcopy(child))
                key = b"".join(etree.tostring(ch, method="c14n") for ch in si)
                if key not in index:
                    sst.append(si)
                    index[key] = len(sst) - 1
                c.remove(is_)
                c.set("t", "s")
                c.add(x.V(text=str(index[key])))
                converted += 1
        refs = sum(1 for name in self.sheet_names
                   for c in self.sheet(name).iter(f"{{{X}}}c") if c.get("t") == "s")
        sst.set("count", str(refs))
        sst.set("uniqueCount", str(len(sst)))
        return converted

    # -- print setup and sheet options ------------------------------------------
    def page_setup(self, sheet: str, **options) -> None:
        """Orientation, paper, fit to pages, scale, margins, print area,
        repeated rows/columns, header and footer, gridlines, centring; see
        ``oox.xlprint.page_setup``."""
        from .xlprint import page_setup

        page_setup(self, sheet, **options)

    def page_setup_info(self, sheet: str) -> dict:
        """``{"orientation", "paper", "margins", ...}``: the sheet's print setup, with
        ``fit``, ``scale``, ``print_area``, ``repeat_rows``, ``header``/``footer`` where
        set."""
        from .xlprint import page_setup_info

        return page_setup_info(self, sheet)

    def sheet_options(self, sheet: str, **options) -> None:
        """tab_color, hidden (True/False/"very"), zoom, gridlines."""
        from .xlprint import sheet_options

        sheet_options(self, sheet, **options)

    def sheet_options_info(self, sheet: str) -> dict:
        """``{"tab_color", "hidden", "protected", "zoom", "gridlines"}`` for the sheet."""
        from .xlprint import sheet_options_info

        return sheet_options_info(self, sheet)

    def protect_sheet(self, sheet: str, password: str | None = None, *,
                      allow: tuple[str, ...] = ()) -> None:
        """Protect the sheet's cells (Review > Protect Sheet), optionally with Excel's legacy
        password hash.  ``allow``: actions still permitted, e.g. ("sort", "autoFilter",
        "formatCells")."""
        from .xlprint import protect_sheet

        protect_sheet(self, sheet, password, allow=allow)

    def unprotect_sheet(self, sheet: str) -> None:
        """Remove the sheet's protection."""
        from .xlprint import unprotect_sheet

        unprotect_sheet(self, sheet)

    def protect_workbook(self, *, password: str | None = None, structure: bool = True,
                         windows: bool = False) -> None:
        """Lock the structure (sheets can't be added, deleted, moved or renamed)
        and/or windows; see ``oox.protection.protect_workbook``."""
        from .protection import protect_workbook

        protect_workbook(self, password=password, structure=structure, windows=windows)

    def workbook_protection(self) -> dict | None:
        """``{"structure", "windows", "password"}`` (booleans), or None."""
        from .protection import workbook_protection

        return workbook_protection(self)

    def check_workbook_password(self, password: str) -> bool:
        """Whether ``password`` unlocks the workbook structure (True without a password);
        checks Excel 2007's 16-bit hash and Excel 2010's salted one."""
        from .protection import check_workbook_password

        return check_workbook_password(self, password)

    def unprotect_workbook(self) -> None:
        """Remove workbook structure and window protection."""
        from .protection import unprotect_workbook

        unprotect_workbook(self)

    # -- comments and hyperlinks --------------------------------------------------
    def add_comment(self, sheet: str, ref: str, text: str, *, author: str = "oox",
                    visible: bool = False) -> OoxElement:
        """A cell comment (note); replaces an existing one on that cell."""
        from .xlnotes import add_comment

        return add_comment(self, sheet, ref, text, author=author, visible=visible)

    def comments(self, sheet: str) -> dict[str, tuple[str, str]]:
        """``{ref: (author, text)}`` for the sheet's comments; the "Author:" line Excel puts
        first is removed."""
        from .xlnotes import comments

        return comments(self, sheet)

    def add_hyperlink(self, sheet: str, ref: str, target: str, *, text: str | None = None,
                      tooltip: str | None = None, style: bool = True) -> OoxElement:
        """A link in a cell: a URL, or "#Sheet!A1" / "#Name" inside the workbook."""
        from .xlnotes import add_hyperlink

        return add_hyperlink(self, sheet, ref, target, text=text, tooltip=tooltip, style=style)

    def hyperlinks(self, sheet: str) -> dict[str, str]:
        """``{ref: target}``: a URL, or ``#Sheet!A1`` for a link inside the workbook."""
        from .xlnotes import hyperlinks

        return hyperlinks(self, sheet)

    # -- tables -----------------------------------------------------------------
    def add_table(self, sheet: str, ref: str, *, name: str | None = None,
                  style: str | None = "TableStyleMedium2", totals: dict | None = None,
                  stripes: bool = True, first_column: bool = False,
                  last_column: bool = False) -> OoxElement:
        """Make ``ref`` an Excel table (header row = first row); ``totals``
        adds a totals row: {"Amount": "sum", "Region": "Total"}.  See ``oox.xltables``."""
        from .xltables import add_table

        return add_table(self, sheet, ref, name=name, style=style, totals=totals,
                         stripes=stripes, first_column=first_column, last_column=last_column)

    def tables(self, sheet: str) -> list[dict]:
        """``{"name", "ref", "columns", "style", "totals", "header"}`` for each table on the
        sheet."""
        from .xltables import tables

        return tables(self, sheet)

    # -- conditional formatting --------------------------------------------------
    def conditional_format(self, sheet: str, ref: str, rule: str, **options) -> OoxElement:
        """Add a conditional format; see ``oox.xlcond`` for the rules and options."""
        from .xlcond import conditional_format

        return conditional_format(self, sheet, ref, rule, **options)

    def conditional_formats(self, sheet: str) -> list[dict]:
        """``{"range", "type", "priority", ...}`` per rule, in priority order, with
        ``operator``, ``formulas``, ``style``, ``colors``, ``icons``, ``rank`` or ``text``
        where they apply."""
        from .xlcond import conditional_formats

        return conditional_formats(self, sheet)

    # -- layout -------------------------------------------------------------------
    def set_column_width(self, sheet: str, cols: str, chars: float | None) -> None:
        """Width in characters, as Excel's Column Width dialog shows it;
        ``cols`` is "B" or "B:D"; None resets."""
        from .xlsheet import set_column_width

        set_column_width(self._ws(sheet), cols, chars)

    def column_width(self, sheet: str, col: str) -> float | None:
        """Width of column ``col`` ("B") in characters as Excel's dialog shows it, or None
        for the default."""
        from .xlsheet import column_width

        return column_width(self._ws(sheet), col)

    def set_row_height(self, sheet: str, row_index: int, points: float | None) -> None:
        """Set row ``row_index`` (1-based) to ``points``; None returns it to the default."""
        from .xlsheet import set_row_height

        set_row_height(self, self._ws(sheet), row_index, points)

    def autofit_columns(self, sheet: str, cols: str | None = None, **options) -> dict[str, float]:
        """Widen columns to their longest displayed text (an estimate), so
        numbers and dates don't show as ####; see ``oox.xlsheet.autofit_columns``."""
        from .xlsheet import autofit_columns

        return autofit_columns(self, sheet, cols, **options)

    def row_height(self, sheet: str, row_index: int) -> float | None:
        """Height of row ``row_index`` (1-based) in points, or None for the default."""
        from .xlsheet import row_height

        return row_height(self, self._ws(sheet), row_index)

    def merge_cells(self, sheet: str, ref: str) -> None:
        """Merge the range ``ref`` ("A1:C1"); the top-left cell's value is shown."""
        from .xlsheet import merge_cells

        merge_cells(self._ws(sheet), ref)

    def unmerge_cells(self, sheet: str, ref: str) -> None:
        """Undo the merge of ``ref``."""
        from .xlsheet import unmerge_cells

        unmerge_cells(self._ws(sheet), ref)

    def merged_cells(self, sheet: str) -> list[str]:
        """The sheet's merged ranges, in document order."""
        from .xlsheet import merged_cells

        return merged_cells(self._ws(sheet))

    def freeze_panes(self, sheet: str, ref: str | None) -> None:
        """Freeze the rows above and columns left of ``ref`` ("B2"); None unfreezes."""
        from .xlsheet import freeze_panes

        freeze_panes(self._ws(sheet), ref)

    def frozen_panes(self, sheet: str) -> str | None:
        """The top-left unfrozen cell ("B2") if panes are frozen, else None."""
        from .xlsheet import frozen_panes

        return frozen_panes(self._ws(sheet))

    def auto_filter(self, sheet: str, ref: str | None) -> None:
        """Filter buttons on the header row of ``ref``; None removes them."""
        from .xlsheet import auto_filter

        auto_filter(self, sheet, ref)

    def auto_filter_range(self, sheet: str) -> str | None:
        """The sheet's autofilter range ("A1:D20"), or None."""
        from .xlsheet import auto_filter_range

        return auto_filter_range(self._ws(sheet))

    def define_name(self, name: str, sheet: str, ref: str, *, local: bool = False) -> OoxElement:
        """A named range ("Sales" -> Sheet1!$B$2:$D$20); ``local`` scopes it to the sheet."""
        from .xlsheet import absolute, define_name

        return define_name(self, name, absolute(sheet, ref),
                           local_sheet=self.tab_names.index(sheet) if local else None)

    def defined_names(self) -> dict[str, str]:
        """``{name: refers_to}``; sheet-scoped names as ``"Sheet!name"``; Excel's hidden
        names (``_xlnm.*``) included."""
        from .xlsheet import defined_names

        return defined_names(self)

    def add_validation(self, sheet: str, ref: str, **options) -> OoxElement:
        """Data validation: ``values=[...]`` (drop-down), ``source="$F$1:$F$5"``,
        ``whole=(1, 10)`` or ``decimal=(0, 1)``; ``prompt=``, ``error=``."""
        from .xlsheet import add_validation

        return add_validation(self._ws(sheet), ref, **options)

    def validations(self, sheet: str) -> list[dict]:
        """``{"range", "type", ...}`` per data-validation rule (lists add ``values``, bounds
        add ``operator`` and ``formulas``)."""
        from .xlsheet import validations

        return validations(self._ws(sheet))

    def add_chart(self, sheet: str, kind: str, data: str, **options) -> OoxElement:
        """Chart of a cell range (header row = series names, first column =
        categories); see ``oox.xlcharts.add_chart``."""
        from .xlcharts import add_chart

        return add_chart(self, sheet, kind, data, **options)

    def charts(self, sheet: str) -> list[dict]:
        """Charts on a worksheet or chart sheet, in drawing order: ``chart_info`` plus
        "anchor" (the top-left cell; None on a chart sheet) and "name"."""
        from .xlcharts import charts

        return charts(self, sheet)

    def add_picture(self, sheet: str | OoxElement, image, anchor: str = "A1", *,
                    width: int | None = None, height: int | None = None,
                    descr: str = "", fallback=None) -> OoxElement:
        """Float a picture over the sheet with its top-left corner at ``anchor``.

        Size in EMU (``oox.units``); one dimension keeps the aspect ratio.  An
        SVG needs ``fallback=`` (a PNG) for Office before 2016.
        """
        ws = self._ws(sheet)
        ws_part = self.part_of(ws)
        found = self.related_parts(ws_part, "DrawingsPart")
        if found:
            dpart = found[0]
        else:
            dpart, drid = self.add_part("DrawingsPart", ws_part, element=xdr.WsDr())
            ws.add(x.Drawing(r_id=drid))
        rid, info, svg_rid = self.add_picture_image(image, fallback, dpart)
        cx, cy = info.scaled(width, height)
        col, rw = split_ref(anchor)
        n = drawing.max_shape_id(dpart.element) + 1
        pic = drawing.picture(xdr, n, f"Picture {n}", rid, cx, cy, descr=descr, svg_rid=svg_rid)
        dpart.element.add(xdr.OneCellAnchor(
            xdr.From(xdr.Col(text=str(col - 1)), xdr.ColOff(text="0"),
                     xdr.Row(text=str(rw - 1)), xdr.RowOff(text="0")),
            xdr.Ext(cx=cx, cy=cy),
            pic,
            xdr.ClientData(),
        ))
        return pic


def _last_child(parent: OoxElement, local: str):
    for el in parent.iterchildren(reversed=True):
        if isinstance(el.tag, str) and el.tag.endswith("}" + local):
            return el
    return None


def _last_index(parent: OoxElement, kind: str) -> int:
    """Index of the last row (``kind="r"``) or cell (``"c"``), 0 if none.

    Rows and cells are kept in ascending order (Excel requires it), so the
    last element tells -- unless it lacks ``r``, then every index is counted.
    """
    last = _last_child(parent, "row" if kind == "r" else "c")
    if last is None:
        return 0
    ref = last.get("r")
    if ref is not None:
        return int(ref) if kind == "r" else split_ref(ref)[0]
    return max((i for i, _ in _indexed(parent, kind)), default=0)


def _indexed(parent: OoxElement, kind: str):
    """Yield ``(index, element)`` for rows (``kind="r"``) or cells (``"c"``).

    Rows carry ``r="7"``, cells ``r="B7"``; both are optional in the schema,
    in which case the index follows the previous one.
    """
    local = "row" if kind == "r" else "c"
    prev = 0
    for el in parent:
        if not (isinstance(el.tag, str) and el.tag.endswith("}" + local)):
            continue
        ref = el.get("r")
        if ref is None:
            prev += 1
        elif kind == "r":
            prev = int(ref)
        else:
            prev = split_ref(ref)[0]
        yield prev, el
