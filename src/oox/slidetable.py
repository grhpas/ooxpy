"""Editing tables already on a slide (made by oox, PowerPoint or anyone else).

    t = deck.table(slide)                 # first table, or by index / shape name
    t.values()                            # [["Region", "Q1"], ["North", "120"], ...]
    t.set_text(1, 1, 125)                 # keeps the cell's font, size and alignment
    t.add_row(["West", 80])               # formatted like the last row
    t.delete_row(2)
    t.add_column(["Q2", 130, 95], index=2)
    t.merge(0, 0, 0, 1)

Rows and columns are 0-based, as in ``oox.tables.Table`` (Word).  Merged cells
follow DrawingML: the top-left cell carries ``gridSpan``/``rowSpan``, the
cells it covers ``hMerge``/``vMerge``; inserting a row or column inside a
merge is refused (unmerge first), deleting one shrinks the merge.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import a
from .units import DRAWING_LIMIT, emu, points

if TYPE_CHECKING:
    from .presentation import PresentationDocument

A = URI["a"]
_MERGE_ATTRS = ("gridSpan", "rowSpan", "hMerge", "vMerge")


def _on(tc: OoxElement, attr: str) -> bool:
    return tc.get(attr) in ("1", "true")


def _int(tc: OoxElement, attr: str) -> int:
    return int(tc.get(attr, "1"))


class SlideTable:
    """A table graphic frame on a slide (``p:graphicFrame`` holding ``a:tbl``)."""

    def __init__(self, deck: PresentationDocument | None, frame: OoxElement) -> None:
        tbl = frame.find(f".//{{{A}}}tbl")
        if tbl is None:
            raise ValueError("not a table graphic frame")
        self.deck, self.frame = deck, frame
        self.element: OoxElement = tbl  # type: ignore[assignment]

    def __repr__(self) -> str:
        return f"<SlideTable {self.n_rows}x{self.n_cols}>"

    # -- structure --------------------------------------------------------------------
    @property
    def rows(self) -> list[OoxElement]:
        """The ``a:tr`` elements."""
        return list(self.element.iterchildren(f"{{{A}}}tr"))

    @property
    def n_rows(self) -> int:
        """Number of rows."""
        return len(self.rows)

    def _grid_cols(self) -> list[OoxElement]:
        grid = self.element.first("a:tblGrid")
        return list(grid.iterchildren(f"{{{A}}}gridCol")) if grid is not None else []

    @property
    def n_cols(self) -> int:
        """Number of grid columns."""
        return len(self._grid_cols())

    @property
    def widths(self) -> list[float]:
        """Column widths in points."""
        return [points(int(g.get("w", "0"))) for g in self._grid_cols()]

    def set_widths(self, widths: list[float]) -> None:
        """Column widths in points (or ``Length``); the frame's width follows."""
        cols = self._grid_cols()
        if len(widths) != len(cols):
            raise ValueError(f"{len(widths)} widths for {len(cols)} columns")
        for g, wd in zip(cols, widths, strict=True):
            g.set("w", str(emu(wd, "widths", DRAWING_LIMIT)))
        self._fit_frame()

    def _tc(self, row: int, col: int) -> OoxElement:
        rows = self.rows
        if not (0 <= row < len(rows) and 0 <= col < self.n_cols):
            raise IndexError(f"no cell at ({row}, {col}) in a {self.n_rows}x{self.n_cols} table")
        cells = list(rows[row].iterchildren(f"{{{A}}}tc"))
        if col >= len(cells):
            raise IndexError(f"row {row} has only {len(cells)} cells")
        return cells[col]

    def _origin(self, row: int, col: int) -> tuple[int, int]:
        """The top-left cell of the merge covering (row, col)."""
        while col > 0 and _on(self._tc(row, col), "hMerge"):
            col -= 1
        while row > 0 and _on(self._tc(row, col), "vMerge"):
            row -= 1
        return row, col

    def cell(self, row: int, col: int) -> OoxElement:
        """The ``a:tc`` shown at (row, col): a merge's top-left cell."""
        return self._tc(*self._origin(row, col))

    def text(self, row: int, col: int) -> str:
        """Text of the cell at (``row``, ``col``); paragraphs are joined with ``\\n``."""
        return _text(self.cell(row, col))

    def values(self) -> list[list[str | None]]:
        """Text per grid position; cells covered by a merge are ``None``."""
        out = []
        for tr in self.rows:
            out.append([None if _on(tc, "hMerge") or _on(tc, "vMerge") else _text(tc)
                        for tc in tr.iterchildren(f"{{{A}}}tc")])
        return out

    def merges(self) -> list[tuple[int, int, int, int]]:
        """``(row1, col1, row2, col2)`` of each merged block."""
        out = []
        for r, tr in enumerate(self.rows):
            for c, tc in enumerate(tr.iterchildren(f"{{{A}}}tc")):
                if _on(tc, "hMerge") or _on(tc, "vMerge"):
                    continue
                gs, rs = _int(tc, "gridSpan"), _int(tc, "rowSpan")
                if gs > 1 or rs > 1:
                    out.append((r, c, r + rs - 1, c + gs - 1))
        return out

    # -- text -------------------------------------------------------------------------
    def set_text(self, row: int, col: int, text: Any, *, size: float | None = None,
                 bold: bool | None = None, color: str | None = None,
                 align: str | None = None) -> OoxElement:
        """Replace the text of the cell shown at (row, col); ``\\n`` starts a
        paragraph.  The first run's font, size and colour and the first
        paragraph's alignment are kept unless given."""
        tc = self.cell(row, col)
        _replace_text(tc, text, size=size, bold=bold, color=color, align=align)
        return tc

    # -- rows -------------------------------------------------------------------------
    def add_row(self, values: list[Any] | None = None, *, index: int | None = None,
                like: int | None = None) -> OoxElement:
        """Insert a row before ``index`` (default: at the end), formatted like
        row ``like`` (default: the row above, or the first row); returns it."""
        rows = self.rows
        index = len(rows) if index is None else index
        if not 0 <= index <= len(rows):
            raise IndexError(f"row index {index} for a table of {len(rows)} rows")
        if 0 < index < len(rows):
            for c, tc in enumerate(rows[index].iterchildren(f"{{{A}}}tc")):
                if _on(tc, "vMerge"):
                    raise ValueError(f"row {index} is inside a merge at column {c}; unmerge first")
        like = (index - 1 if index > 0 else 0) if like is None else like
        tr = copy.deepcopy(rows[like])
        for c, tc in enumerate(tr.iterchildren(f"{{{A}}}tc")):
            for attr in _MERGE_ATTRS:
                if attr in tc.attrib:
                    del tc.attrib[attr]
            value = values[c] if values and c < len(values) else None
            _replace_text(tc, value)
        if index < len(rows):
            rows[index].addprevious(tr)
        else:
            rows[-1].addnext(tr)
        self._fit_frame()
        return tr  # type: ignore[return-value]

    def delete_row(self, row: int) -> None:
        """Remove a row; merges through it shrink (a merge starting in it moves
        down with its text)."""
        rows = self.rows
        if not 0 <= row < len(rows):
            raise IndexError(f"no row {row} in a table of {len(rows)} rows")
        if len(rows) == 1:
            raise ValueError("a table keeps at least one row")
        cells = list(rows[row].iterchildren(f"{{{A}}}tc"))
        for c, tc in enumerate(cells):
            if _on(tc, "hMerge"):
                continue
            rs = _int(tc, "rowSpan")
            if _on(tc, "vMerge"):
                r0, c0 = self._origin(row, c)
                origin = self._tc(r0, c0)
                if _int(origin, "rowSpan") > 1:
                    _set_span(origin, "rowSpan", _int(origin, "rowSpan") - 1)
            elif rs > 1:                         # the merge continues in the next row
                below = self._tc(row + 1, c)
                gs = _int(tc, "gridSpan")
                for attr in ("vMerge", "hMerge"):
                    below.attrib.pop(attr, None)
                _set_span(below, "rowSpan", rs - 1)
                _set_span(below, "gridSpan", gs)
                for k in range(1, gs):            # the cells right of it stay covered
                    right = self._tc(row + 1, c + k)
                    right.attrib.pop("vMerge", None)
                    right.set("hMerge", "1")
                body, keep = tc.first("a:txBody"), below.first("a:txBody")
                if body is not None and keep is not None:
                    keep.addprevious(copy.deepcopy(body))
                    below.remove(keep)
        rows[row].getparent().remove(rows[row])
        self._fit_frame()

    # -- columns ----------------------------------------------------------------------
    def add_column(self, values: list[Any] | None = None, *, index: int | None = None,
                   width: float | None = None, like: int | None = None) -> None:
        """Insert a column before ``index`` (default: at the end), formatted
        like column ``like`` (default: the one left of it); ``width`` in points
        or ``Length`` (default: that column's).  The frame grows by the width."""
        cols = self._grid_cols()
        index = len(cols) if index is None else index
        if not 0 <= index <= len(cols):
            raise IndexError(f"column index {index} for a table of {len(cols)} columns")
        rows = self.rows
        if 0 < index < len(cols):
            for r, tr in enumerate(rows):
                if _on(list(tr.iterchildren(f"{{{A}}}tc"))[index], "hMerge"):
                    raise ValueError(f"column {index} is inside a merge in row {r}; unmerge first")
        like = (index - 1 if index > 0 else 0) if like is None else like
        new_col = a.GridCol(w=emu(width, "width", DRAWING_LIMIT) if width is not None
                            else int(cols[like].get("w", "0")))
        if index < len(cols):
            cols[index].addprevious(new_col)
        else:
            cols[-1].addnext(new_col)
        for r, tr in enumerate(rows):
            cells = list(tr.iterchildren(f"{{{A}}}tc"))
            tc = copy.deepcopy(cells[like])
            for attr in _MERGE_ATTRS:
                tc.attrib.pop(attr, None)
            _replace_text(tc, values[r] if values and r < len(values) else None)
            if index < len(cells):
                cells[index].addprevious(tc)
            else:
                cells[-1].addnext(tc)
        self._fit_frame()

    def delete_column(self, col: int) -> None:
        """Remove a column; merges through it shrink."""
        cols = self._grid_cols()
        if not 0 <= col < len(cols):
            raise IndexError(f"no column {col} in a table of {len(cols)} columns")
        if len(cols) == 1:
            raise ValueError("a table keeps at least one column")
        for tr in self.rows:
            cells = list(tr.iterchildren(f"{{{A}}}tc"))
            tc = cells[col]
            gs = _int(tc, "gridSpan")
            if _on(tc, "hMerge"):
                c0 = col
                while c0 > 0 and _on(cells[c0], "hMerge"):
                    c0 -= 1
                if _int(cells[c0], "gridSpan") > 1:     # the block's cell in this row
                    _set_span(cells[c0], "gridSpan", _int(cells[c0], "gridSpan") - 1)
            elif gs > 1:                          # the merge continues to the right
                right = cells[col + 1]
                right.attrib.pop("hMerge", None)
                _set_span(right, "gridSpan", gs - 1)
                if "rowSpan" in tc.attrib:
                    right.set("rowSpan", tc.get("rowSpan"))
                if _on(tc, "vMerge"):
                    right.set("vMerge", "1")
                body, keep = tc.first("a:txBody"), right.first("a:txBody")
                if body is not None and keep is not None:
                    keep.addprevious(copy.deepcopy(body))
                    right.remove(keep)
            tr.remove(tc)
        cols[col].getparent().remove(cols[col])
        self._fit_frame()

    # -- merging ----------------------------------------------------------------------
    def merge(self, row1: int, col1: int, row2: int, col2: int) -> OoxElement:
        """Merge the block (inclusive, 0-based); the texts of the merged cells
        join the top-left one, as PowerPoint does.  Returns that cell."""
        if row2 < row1 or col2 < col1:
            row1, row2, col1, col2 = min(row1, row2), max(row1, row2), min(col1, col2), \
                max(col1, col2)
        self._tc(row2, col2)                       # bounds
        for (r1, c1, r2, c2) in self.merges():
            if not (r2 < row1 or r1 > row2 or c2 < col1 or c1 > col2):
                raise ValueError(f"overlaps the merged block ({r1}, {c1})-({r2}, {c2})")
        if (row1, col1) == (row2, col2):
            return self._tc(row1, col1)
        origin = self._tc(row1, col1)
        texts = [_text(origin)] if _text(origin) else []
        for r in range(row1, row2 + 1):
            for c in range(col1, col2 + 1):
                if (r, c) == (row1, col1):
                    continue
                tc = self._tc(r, c)
                if _text(tc):
                    texts.append(_text(tc))
                    self.set_text(r, c, "")
                if c > col1:
                    tc.set("hMerge", "1")
                if r > row1:
                    tc.set("vMerge", "1")
                    if c == col1:                  # as PowerPoint writes it
                        _set_span(tc, "gridSpan", col2 - col1 + 1)
        _set_span(origin, "gridSpan", col2 - col1 + 1)
        _set_span(origin, "rowSpan", row2 - row1 + 1)
        if len(texts) > 1:
            self.set_text(row1, col1, "\n".join(texts))
        return origin

    def unmerge(self, row: int, col: int) -> None:
        """Split the merged block shown at (row, col) back into cells."""
        r0, c0 = self._origin(row, col)
        origin = self._tc(r0, c0)
        rs, gs = _int(origin, "rowSpan"), _int(origin, "gridSpan")
        for r in range(r0, r0 + rs):
            for c in range(c0, c0 + gs):
                for attr in _MERGE_ATTRS:
                    self._tc(r, c).attrib.pop(attr, None)

    # -- style ------------------------------------------------------------------------
    def set_style(self, style: str | None = None, *, header: bool | None = None,
                  banded: bool | None = None, first_column: bool | None = None) -> None:
        """A built-in style by name (``oox.slides.TABLE_STYLES``) or GUID, and
        its header row / banded rows / first column switches."""
        from .slides import TABLE_STYLES

        pr = self.element.first("a:tblPr")
        if pr is None:
            pr = a.TblPr()
            self.element.insert(0, pr)
        for attr, on in (("firstRow", header), ("bandRow", banded), ("firstCol", first_column)):
            if on is not None:
                if on:
                    pr.set(attr, "1")
                else:
                    pr.attrib.pop(attr, None)
        if style is not None:
            sid = pr.first("a:tableStyleId")
            guid = TABLE_STYLES.get(style, style)
            if sid is None:
                pr.add(a.TableStyleId(text=guid))
            else:
                sid.text = guid

    # -- frame ------------------------------------------------------------------------
    def _fit_frame(self) -> None:
        """The graphic frame's size follows the grid and the row heights."""
        ext = self.frame.find(f"{{{URI['p']}}}xfrm/{{{A}}}ext")
        if ext is None:
            return
        ext.set("cx", str(sum(int(g.get("w", "0")) for g in self._grid_cols())))
        ext.set("cy", str(sum(int(tr.get("h", "0")) for tr in self.rows)))


def _replace_text(tc: OoxElement, text: Any, *, size: float | None = None,
                  bold: bool | None = None, color: str | None = None,
                  align: str | None = None) -> None:
    """Replace a cell's paragraphs, keeping the first run's and paragraph's properties."""
    from .slides import _ALIGN

    body = tc.first("a:txBody")
    if body is None:
        body = a.TxBody(a.BodyPr(), a.LstStyle())
        tc.insert(0, body)
    first_p = body.first("a:p")
    ppr = copy.deepcopy(first_p.first("a:pPr")) if first_p is not None and \
        first_p.first("a:pPr") is not None else None
    rpr = None
    for r in body.iter(f"{{{A}}}r"):
        if r.find(f"{{{A}}}rPr") is not None:
            rpr = copy.deepcopy(r.find(f"{{{A}}}rPr"))
        break
    end = first_p.first("a:endParaRPr") if first_p is not None else None
    if rpr is None and end is not None:
        rpr = a.RPr(*copy.deepcopy(list(end)), **dict(end.attrib))
    if rpr is None:
        rpr = a.RPr(lang="en-US", dirty=False)
    if size is not None:
        rpr.set("sz", str(int(size * 100)))
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    if color is not None:
        for old in list(rpr):
            if old.tag.split("}")[1] in ("noFill", "solidFill", "gradFill", "blipFill",
                                         "pattFill", "grpFill"):
                rpr.remove(old)
        rpr.add(a.SolidFill(a.SrgbClr(val=color.lstrip("#").upper())))
    if align is not None:
        if align not in _ALIGN:
            raise ValueError(f"align={align!r}: use one of {sorted(_ALIGN)}")
        ppr = ppr if ppr is not None else a.PPr()
        ppr.set("algn", _ALIGN[align])
    for old in list(body.iterchildren(f"{{{A}}}p")):
        body.remove(old)
    lines = "" if text is None else str(text)
    for line in lines.split("\n"):
        para = a.P()
        if ppr is not None:
            para.add(copy.deepcopy(ppr))
        if line:
            para.add(a.R(copy.deepcopy(rpr), a.T(line)))
        else:
            end = a.EndParaRPr(**{k: v for k, v in rpr.attrib.items()
                                  if k in ("lang", "sz", "b", "i")})
            para.add(end)
        body.add(para)


def _set_span(tc: OoxElement, attr: str, n: int) -> None:
    if n > 1:
        tc.set(attr, str(n))
    else:
        tc.attrib.pop(attr, None)


def _text(tc: OoxElement) -> str:
    lines = []
    for para in tc.iter(f"{{{A}}}p"):
        lines.append("".join(t.text or "" for t in para.iter(f"{{{A}}}t")))
    return "\n".join(lines)
