"""Word tables: building, merged cells, widths, header rows, reading.

    t = doc.add_table([["Region", "Q1", "Q2"], ["North", 120, 135]],
                      header_rows=1, widths=[Cm(5), Cm(3), Cm(3)])
    t.merge(2, 0, 2, 1)                 # cells (2,0)-(2,1) become one
    t.format_cell(0, 0, fill="D9E2F3", align="center")
    t.values()                          # [["Region", "Q1", "Q2"], ...]

Coordinates are 0-based *grid* positions, so they don't shift when cells are
merged: after ``merge(2, 0, 2, 1)``, both ``cell(2, 0)`` and ``cell(2, 1)``
return the merged cell.

How OOXML stores merges: a horizontal merge is one ``w:tc`` with
``w:gridSpan``; a vertical merge keeps a ``w:tc`` in every row, the first with
``w:vMerge val="restart"`` and the ones below with ``w:vMerge`` (continue).
Widths and heights are in points, or a ``Length`` from ``oox.units``.

Word doesn't count a row whose cells all continue vertical merges (Word 2007
reports 3 rows for such a 4-row table, and the row takes no height); LibreOffice
shows it as a thin row.  ``Table.word_row_count`` gives Word's count; avoid
merges that cover a whole row if the layout matters.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import w
from .units import PAGE_LIMIT, twips, twips_to_points

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}
_VALIGN = ("top", "center", "bottom")


def _twips(value: float, name: str = "width") -> int:
    return twips(value, name, PAGE_LIMIT)


def _w(el: OoxElement, attr: str, default: str | None = None) -> str | None:
    return el.get(f"{{{W}}}{attr}", default)


def _span(tc: OoxElement) -> int:
    tcpr = tc.first("w:tcPr")
    gs = tcpr.first("w:gridSpan") if tcpr is not None else None
    return int(_w(gs, "val", "1") or 1) if gs is not None else 1


def _vmerge(tc: OoxElement) -> str | None:
    """None, "restart" or "continue"."""
    tcpr = tc.first("w:tcPr")
    vm = tcpr.first("w:vMerge") if tcpr is not None else None
    if vm is None:
        return None
    return "restart" if _w(vm, "val") == "restart" else "continue"


def _tcpr(tc: OoxElement) -> OoxElement:
    tcpr = tc.first("w:tcPr")
    if tcpr is None:
        tcpr = w.TcPr()
        tc.insert(0, tcpr)
    return tcpr


def _set_row_prop(tr: OoxElement, prop: OoxElement | None, tag: str) -> None:
    """Set (or with prop=None, remove) one row property.  w:trPr is an
    unbounded choice in the schema, so replace= can't be used: it would add a
    second copy."""
    trpr = tr.first("w:trPr")
    if trpr is not None:
        for old in list(trpr.iterchildren(f"{{{W}}}{tag}")):
            trpr.remove(old)
    if prop is None:
        return
    if trpr is None:
        trpr = w.TrPr()
        tr.insert(0, trpr)
    trpr.add(prop)


def _fill_cell(tc: OoxElement, value: Any, bold: bool = False) -> None:
    from . import teletype

    if isinstance(value, OoxElement):
        tc.add(value)
        return
    if isinstance(value, (list, tuple)) and value and all(isinstance(v, OoxElement) for v in value):
        tc.add(*value)
        return
    p = w.P()
    if value is not None and value != "":
        r = w.R(w.RPr(w.B())) if bold else w.R()
        teletype.addTextToElement(r, str(value))
        p.add(r)
    tc.add(p)


class Table:
    def __init__(self, doc: WordprocessingDocument | None, tbl: OoxElement) -> None:
        self.doc, self.element = doc, tbl

    def __repr__(self) -> str:
        return f"<Table {self.n_rows}x{self.n_cols}>"

    # -- structure --------------------------------------------------------------------
    @property
    def rows(self) -> list[OoxElement]:
        """The ``w:tr`` elements."""
        return list(self.element.iterchildren(f"{{{W}}}tr"))

    @property
    def n_rows(self) -> int:
        """Number of rows."""
        return len(self.rows)

    @property
    def word_row_count(self) -> int:
        """Rows as Word counts them: rows made only of vertical-merge
        continuation cells don't count."""
        return sum(1 for tr in self.rows
                   if any(_vmerge(tc) != "continue" for tc in tr.iterchildren(f"{{{W}}}tc")))

    @property
    def n_cols(self) -> int:
        """Number of grid columns."""
        grid = self.element.first("w:tblGrid")
        n = len(list(grid.iterchildren(f"{{{W}}}gridCol"))) if grid is not None else 0
        return max([n, *(len(r) for r in self._grid())])

    def _grid(self) -> list[list[OoxElement | None]]:
        """Grid position -> the w:tc shown there (the origin of a merge)."""
        out: list[list[OoxElement | None]] = []
        for r, tr in enumerate(self.rows):
            row: list[OoxElement | None] = []
            trpr = tr.first("w:trPr")
            gb = trpr.first("w:gridBefore") if trpr is not None else None
            if gb is not None:
                row.extend([None] * int(_w(gb, "val", "0") or 0))
            for tc in tr.iterchildren(f"{{{W}}}tc"):
                col, span = len(row), _span(tc)
                origin: OoxElement | None = tc
                if _vmerge(tc) == "continue" and r > 0 and col < len(out[r - 1]):
                    above = out[r - 1][col]
                    if above is not None:
                        origin = above
                row.extend([origin] * span)
            out.append(row)
        return out

    def _starts(self, r: int) -> dict[int, OoxElement]:
        """Grid column -> w:tc that starts there, in row r (incl. continuation cells)."""
        tr = self.rows[r]
        col = 0
        trpr = tr.first("w:trPr")
        gb = trpr.first("w:gridBefore") if trpr is not None else None
        if gb is not None:
            col = int(_w(gb, "val", "0") or 0)
        out = {}
        for tc in tr.iterchildren(f"{{{W}}}tc"):
            out[col] = tc
            col += _span(tc)
        return out

    def cell(self, row: int, col: int) -> OoxElement:
        """The ``w:tc`` shown at grid position (row, col), 0-based."""
        try:
            tc = self._grid()[row][col]
        except IndexError:
            raise IndexError(f"no cell at ({row}, {col}) in a {self.n_rows}x{self.n_cols} table") from None
        if tc is None:
            raise IndexError(f"no cell at ({row}, {col}): the row starts later (gridBefore)")
        return tc

    def text(self, row: int, col: int) -> str:
        """Text of the cell shown at grid position (``row``, ``col``); paragraphs are joined
        with ``\\n``."""
        from . import teletype

        return "\n".join(teletype.extractText(p) for p in self.cell(row, col).iterchildren(f"{{{W}}}p"))

    def values(self) -> list[list[str | None]]:
        """Text per grid position; positions covered by a merge (other than its
        top-left) are ``None``."""
        from . import teletype

        out: list[list[str | None]] = []
        seen: set[int] = set()
        for row in self._grid():
            vals: list[str | None] = []
            for tc in row:
                if tc is None or id(tc) in seen:
                    vals.append(None)
                    continue
                seen.add(id(tc))
                vals.append("\n".join(teletype.extractText(p)
                                      for p in tc.iterchildren(f"{{{W}}}p")))
            out.append(vals)
        return out

    # -- editing ----------------------------------------------------------------------
    def set_text(self, row: int, col: int, text: Any, *, bold: bool = False) -> OoxElement:
        """Replace the cell's content with one paragraph of ``text``."""
        tc = self.cell(row, col)
        for child in list(tc):
            if child.tag != f"{{{W}}}tcPr":
                tc.remove(child)
        _fill_cell(tc, text, bold)
        return tc

    def format_cell(self, row: int, col: int, *, fill: str | None = None,
                    valign: str | None = None, align: str | None = None,
                    bold: bool | None = None) -> OoxElement:
        """Shading (hex colour), vertical alignment (top/center/bottom),
        paragraph alignment (left/center/right/justify), bold."""
        tc = self.cell(row, col)
        tcpr = _tcpr(tc)
        if fill is not None:
            tcpr.add(w.Shd(val="clear", color="auto", fill=fill), replace=True)
        if valign is not None:
            if valign not in _VALIGN:
                raise ValueError(f"valign is one of {_VALIGN}")
            tcpr.add(w.VAlign(val=valign), replace=True)
        for p in tc.iterchildren(f"{{{W}}}p"):
            if align is not None:
                ppr = p.first("w:pPr")
                if ppr is None:
                    ppr = w.PPr()
                    p.insert(0, ppr)
                ppr.add(w.Jc(val=_ALIGN[align]), replace=True)
            if bold is not None:
                for r in p.iter(f"{{{W}}}r"):
                    rpr = r.first("w:rPr")
                    if bold:
                        if rpr is None:
                            rpr = w.RPr()
                            r.insert(0, rpr)
                        rpr.add(w.B(), replace=True)
                    elif rpr is not None and rpr.first("w:b") is not None:
                        rpr.remove(rpr.first("w:b"))
        return tc

    def merge(self, row1: int, col1: int, row2: int, col2: int) -> OoxElement:
        """Merge the rectangle (row1, col1)-(row2, col2), inclusive; returns the
        merged cell.  Text of the merged cells is kept, as in Word."""
        if row2 < row1 or col2 < col1:
            raise ValueError("give the top-left corner first")
        grid = self._grid()
        if row2 >= len(grid) or any(col2 >= len(grid[r]) for r in range(row1, row2 + 1)):
            raise IndexError("merge range is outside the table")
        # every position must be an unmerged cell inside the rectangle
        inside = {id(grid[r][c]) for r in range(row1, row2 + 1) for c in range(col1, col2 + 1)}
        for r, row in enumerate(grid):
            for c, tc in enumerate(row):
                in_rect = row1 <= r <= row2 and col1 <= c <= col2
                if tc is not None and id(tc) in inside and not in_rect:
                    raise ValueError(f"cell ({r}, {c}) is already merged across the range edge")
        widths = self._grid_twips()
        span = col2 - col1 + 1
        origin: OoxElement | None = None
        for r in range(row1, row2 + 1):
            starts = self._starts(r)
            if col1 not in starts:
                raise ValueError(f"row {r} has no cell starting at column {col1}")
            first = starts[col1]
            for c in range(col1 + 1, col2 + 1):
                tc = starts.get(c)
                if tc is not None:
                    _move_content(tc, origin if origin is not None else first)
                    tc.getparent().remove(tc)
            tcpr = _tcpr(first)
            if span > 1:
                tcpr.add(w.GridSpan(val=span), replace=True)
                if widths:
                    tcpr.add(w.TcW(w=sum(widths[col1:col2 + 1]), type="dxa"), replace=True)
            if row2 > row1:
                tcpr.add(w.VMerge(val="restart") if r == row1 else w.VMerge(), replace=True)
            if origin is None:
                origin = first
            else:
                _move_content(first, origin)
                first.add(w.P())  # every cell keeps a paragraph
        assert origin is not None
        return origin

    @property
    def widths(self) -> list[float]:
        """Column widths from the table grid, in points."""
        return [twips_to_points(x) for x in self._grid_twips()]

    def _grid_twips(self) -> list[int]:
        grid = self.element.first("w:tblGrid")
        if grid is None:
            return []
        return [int(_w(g, "w", "0") or 0) for g in grid.iterchildren(f"{{{W}}}gridCol")]

    def set_widths(self, widths: list[float]) -> None:
        """Fixed column widths in points (or ``Length``); updates the grid, every
        cell and the table width, and switches to a fixed layout."""
        widths = [_twips(x) for x in widths]
        if len(widths) != self.n_cols:
            raise ValueError(f"{len(widths)} widths for {self.n_cols} columns")
        grid = self.element.first("w:tblGrid")
        if grid is None:
            grid = w.TblGrid()
            self.element.add(grid)
        for g in list(grid):
            grid.remove(g)
        grid.add(*(w.GridCol(w=x) for x in widths))
        for r in range(self.n_rows):
            for c, tc in self._starts(r).items():
                _tcpr(tc).add(w.TcW(w=sum(widths[c:c + _span(tc)]), type="dxa"), replace=True)
        tblpr = self.element.first("w:tblPr")
        assert tblpr is not None
        tblpr.add(w.TblW(w=sum(widths), type="dxa"), replace=True)
        tblpr.add(w.TblLayout(type="fixed"), replace=True)

    def set_header_rows(self, n: int) -> None:
        """Repeat the first ``n`` rows at the top of each page."""
        for i, tr in enumerate(self.rows):
            _set_row_prop(tr, w.TblHeader() if i < n else None, "tblHeader")

    @property
    def header_rows(self) -> int:
        """How many leading rows repeat as a header on each page."""
        n = 0
        for tr in self.rows:
            trpr = tr.first("w:trPr")
            if trpr is None or trpr.first("w:tblHeader") is None:
                break
            n += 1
        return n

    def set_row_height(self, row: int, height: float, rule: str = "atLeast") -> None:
        """Height of ``row`` in points (or ``Length``); ``rule`` "atLeast" or
        "exact"."""
        _set_row_prop(self.rows[row], w.TrHeight(val=_twips(height, "height"), hRule=rule),
                      "trHeight")

    def add_row(self, values: list[Any] | None = None) -> OoxElement:
        """Append a row (one cell per grid column) and return it."""
        widths = self._grid_twips()
        tr = w.Tr()
        for c in range(self.n_cols):
            tc = w.Tc(w.TcPr(w.TcW(w=widths[c] if c < len(widths) else 0, type="dxa")))
            _fill_cell(tc, values[c] if values and c < len(values) else None)
            tr.add(tc)
        self.element.add(tr)
        return tr


def _move_content(src: OoxElement, dst: OoxElement) -> None:
    """Move src's non-empty paragraphs (and tables) into dst, as Word does on merge."""
    from . import teletype

    for child in list(src):
        if child.tag == f"{{{W}}}tcPr":
            continue
        if child.tag == f"{{{W}}}p" and not teletype.extractText(child) and \
                child.find(f".//{{{W}}}drawing") is None:
            src.remove(child)
            continue
        dst.append(child)
    # a cell whose only paragraph was empty: drop the empty one if dst gained text
    ps = list(dst.iterchildren(f"{{{W}}}p"))
    if len(ps) > 1 and not teletype.extractText(ps[0]) and ps[0].find(f".//{{{W}}}drawing") is None:
        dst.remove(ps[0])


def add_table(doc: WordprocessingDocument, rows: list[list[Any]], *,
              widths: list[float] | None = None, style: str | None = "Table Grid",
              header_rows: int = 0, bold_header: bool = True, align: str | None = None,
              before: OoxElement | None = None) -> Table:
    """A table from rows of values (str, numbers, None, or paragraphs).

    ``widths`` (points or ``Length`` per column) gives a fixed layout; by default the
    columns share the section's text width.  ``header_rows`` repeat on each
    page and are bold unless ``bold_header=False``.
    """
    n_cols = max((len(r) for r in rows), default=0)
    if n_cols == 0:
        raise ValueError("a table needs at least one cell")
    grid: list[int]
    if widths is None:
        text_width = twips(doc.sections[-1].text_width)
        grid = [text_width // n_cols] * n_cols
        fixed = False
    else:
        grid = [_twips(x) for x in widths]
        if len(grid) != n_cols:
            raise ValueError(f"{len(grid)} widths for {n_cols} columns")
        fixed = True
    tblpr = w.TblPr(w.TblW(w=sum(grid), type="dxa"))
    if style:
        tblpr.add(w.TblStyle(val=doc.style_id(style)))
    if align:
        tblpr.add(w.Jc(val=_ALIGN[align]))
    if fixed:
        tblpr.add(w.TblLayout(type="fixed"))
    tbl = w.Tbl(tblpr, w.TblGrid(*(w.GridCol(w=x) for x in grid)))
    for i, values in enumerate(rows):
        tr = w.Tr(w.TrPr(w.TblHeader())) if i < header_rows else w.Tr()
        for c in range(n_cols):
            tc = w.Tc(w.TcPr(w.TcW(w=grid[c], type="dxa")))
            _fill_cell(tc, values[c] if c < len(values) else None,
                       bold=bold_header and i < header_rows)
            tr.add(tc)
        tbl.add(tr)
    if before is not None:
        prev = before.getprevious()
        if prev is not None and prev.tag == f"{{{W}}}tbl":
            before.addprevious(w.P())
        before.addprevious(tbl)
    else:
        blocks = [el for el in doc.body if el.tag != f"{{{W}}}sectPr"]
        if blocks and blocks[-1].tag == f"{{{W}}}tbl":
            doc.body.add(w.P())  # two tables in a row would join into one
        doc.body.add(tbl)
    return Table(doc, tbl)


def tables(doc: WordprocessingDocument) -> list[Table]:
    """The body's top-level tables, including those inside block-level content
    controls (nested ones: ``Table(None, el)`` on a ``w:tbl`` found in a cell)."""
    from .walk import iter_blocks

    return [Table(doc, t) for t in iter_blocks(doc.body, f"{{{W}}}tbl")]


def ensure_paragraph_after_table(doc: WordprocessingDocument) -> None:
    """Word ends a document (and every cell) with a paragraph; called on save."""
    blocks = [el for el in doc.body if el.tag != f"{{{W}}}sectPr"]
    if blocks and blocks[-1].tag == f"{{{W}}}tbl":
        doc.body.add(w.P())
    for tc in doc.body.iter(f"{{{W}}}tc"):
        kids = [k for k in tc if k.tag != f"{{{W}}}tcPr"]
        if not kids or kids[-1].tag == f"{{{W}}}tbl":
            tc.add(w.P())

