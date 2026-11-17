"""Excel cell comments ("notes") and hyperlinks.

    book.add_comment("Sheet1", "B2", "Checked by finance", author="Ada")
    book.comments("Sheet1")             # {"B2": ("Ada", "Checked by finance")}
    book.add_hyperlink("Sheet1", "A10", "https://example.com", text="Our site")
    book.add_hyperlink("Sheet1", "A11", "#Summary!A1", text="Go to summary")
    book.hyperlinks("Sheet1")           # {"A10": "https://example.com", "A11": "#Summary!A1"}

A comment is stored twice, as Excel does: the text in a comments part, and
the yellow box in a legacy VML drawing that Excel needs to show it.  Hidden
comments (the default) show as a red triangle and pop up on hover;
``visible=True`` keeps the box open.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import x

if TYPE_CHECKING:
    from .opc import Part
    from .spreadsheet import SpreadsheetDocument

X = URI["x"]
R = URI["r"]
RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
VML_HEAD = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:x="urn:schemas-microsoft-com:office:excel">'
            '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="{block}"/></o:shapelayout>'
            '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
            'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
            '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>')
VML_SHAPE = ('<v:shape id="_x0000_s{sid}" type="#_x0000_t202" style="position:absolute;'
             'margin-left:59.25pt;margin-top:1.5pt;width:{w}pt;height:{h}pt;z-index:{z};'
             'visibility:{vis}" fillcolor="#ffffe1" o:insetmode="auto">'
             '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
             '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto">'
             '<div style="text-align:left"></div></v:textbox>'
             '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
             '<x:Anchor>{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>'
             '<x:Row>{row}</x:Row><x:Column>{col}</x:Column>{visible}</x:ClientData></v:shape>')


def _comments_part(book: SpreadsheetDocument, ws: OoxElement, create: bool) -> Part | None:
    part = book.part_of(ws)
    found = book.related_parts(part, "WorksheetCommentsPart")
    if found:
        return found[0]
    if not create:
        return None
    cpart, _ = book.add_part("WorksheetCommentsPart", part,
                             element=x.Comments(x.Authors(), x.CommentList()))
    return cpart


def _vml_blocks(book: SpreadsheetDocument) -> set[int]:
    """Shape-id blocks (``o:idmap data``) used by the workbook's VML drawings."""
    used: set[int] = set()
    for part in book.package.parts.values():
        if part.content_type.endswith("vmlDrawing") and part._blob:
            for m in re.finditer(rb'idmap[^>]*data="([\d,]+)"', part._blob):
                used |= {int(n) for n in m.group(1).split(b",") if n}
    return used


def _set_vml_shape(book: SpreadsheetDocument, ws: OoxElement, row: int, col: int,
                   lines: int, visible: bool) -> None:
    """Add (or replace) the note box for one cell in the sheet's VML drawing.
    Edited as text: VML written by Excel is often not well-formed XML."""
    part = book.part_of(ws)
    vml = book.related_parts(part, "VmlDrawingPart")
    if vml:
        text = (vml[0]._blob or b"").decode("utf-8", "replace")
    else:
        used = _vml_blocks(book)
        block = next(i for i in range(1, 10000) if i not in used)
        text = VML_HEAD.format(block=block) + "</xml>"
    # drop an existing box for this cell
    def keep(m: re.Match[str]) -> str:
        body = m.group(0)
        same = re.search(rf"<x:Row>\s*{row - 1}\s*</x:Row>", body) and \
            re.search(rf"<x:Column>\s*{col - 1}\s*</x:Column>", body)
        return "" if same else body
    text = re.sub(r"<v:shape\b.*?</v:shape>", keep, text, flags=re.S)
    block_m = re.search(r'idmap[^>]*data="(\d+)', text)
    block = int(block_m.group(1)) if block_m else 1
    ids = [int(n) for n in re.findall(r'_x0000_s(\d+)', text)]
    sid = max([1024 * block, *ids]) + 1
    z = len(re.findall(r"<v:shape\b", text)) + 1
    top = max(row - 2, 0)
    shape = VML_SHAPE.format(
        sid=sid, z=z, w=108, h=round(15 * lines, 2), vis="visible" if visible else "hidden",
        anchor=f"{col}, 15, {top}, 10, {col + 2}, 15, {top + lines}, 4", row=row - 1,
        col=col - 1, visible="<x:Visible/>" if visible else "")
    cut = text.rfind("</xml>")
    text = text[:cut] + shape + text[cut:]
    blob = text.encode("utf-8")
    if vml:
        vml[0]._blob = blob
        vml[0]._element = None
        return
    _, rid = book.add_part("VmlDrawingPart", part, blob=blob, ext="vml")
    ws.add(x.LegacyDrawing(r_id=rid), replace=True)


def _text(comment: OoxElement) -> str:
    return "".join(t.text or "" for t in comment.iter(f"{{{X}}}t"))


def add_comment(book: SpreadsheetDocument, sheet: str, ref: str, text: str, *,
                author: str = "oox", visible: bool = False) -> OoxElement:
    from .spreadsheet import cell_ref, split_ref

    ws = book.sheet(sheet)
    ref = cell_ref(*split_ref(ref))
    cpart = _comments_part(book, ws, create=True)
    assert cpart is not None
    root = cpart.element
    authors = root.first("x:authors")
    assert authors is not None
    names = [a.text for a in authors]
    if author not in names:
        authors.add(x.Author(text=author))
        names.append(author)
    clist = root.first("x:commentList")
    assert clist is not None
    for old in list(clist):
        if old.get("ref") == ref:
            clist.remove(old)
    # Excel shows the author in bold on the first line; keep that look
    body = x.Text(x.R(x.RPr(x.B(), x.Sz(val=9), x.Color(indexed=81), x.RFont(val="Tahoma"),
                            x.Family(val=2)), x.T(f"{author}:")),
                  x.R(x.RPr(x.Sz(val=9), x.Color(indexed=81), x.RFont(val="Tahoma"),
                            x.Family(val=2)), x.T("\n" + text, space="preserve")))
    comment = x.Comment(body, ref=ref, authorId=names.index(author))
    clist.add(comment)
    # comments in sheet order (by row, then column), as Excel writes them
    def position(c: OoxElement) -> tuple[int, int]:
        col, row = split_ref(c.get("ref") or "A1")
        return row, col

    items = sorted(list(clist), key=position)
    for c in items:
        clist.append(c)
    col, row = split_ref(ref)
    _set_vml_shape(book, ws, row, col, max(3, text.count("\n") + 2), visible)
    return comment


def comments(book: SpreadsheetDocument, sheet: str) -> dict:
    """``{ref: Comment}``; the "Author:" line Excel puts first is removed."""
    from .records import Comment

    cpart = _comments_part(book, book.sheet(sheet), create=False)
    if cpart is None:
        return {}
    root = cpart.element
    authors_el = root.first("x:authors")
    authors = [a.text or "" for a in authors_el] if authors_el is not None else []
    out = {}
    for c in root.iter(f"{{{X}}}comment"):
        aid = int(c.get("authorId", "0"))
        author = authors[aid] if aid < len(authors) else ""
        text = _text(c)
        if author and text.startswith(author + ":"):
            text = text[len(author) + 1:].lstrip("\n")
        out[c.get("ref")] = Comment(text, author=author or None, cell=c.get("ref"))
    return out


def add_hyperlink(book: SpreadsheetDocument, sheet: str, ref: str, target: str, *,
                  text: str | None = None, tooltip: str | None = None,
                  style: bool = True) -> OoxElement:
    """``target``: a URL (external) or ``#Sheet!A1`` / ``#Name`` (inside the
    workbook).  ``text`` is written into the cell; ``style`` makes it blue and
    underlined, as Excel's Hyperlink style."""
    from .spreadsheet import cell_ref, split_ref

    ws = book.sheet(sheet)
    ref = cell_ref(*split_ref(ref))
    attrs: dict[str, Any] = {"ref": ref}
    if target.startswith("#"):
        attrs["location"] = target[1:]
    else:
        attrs["r_id"] = book.part_of(ws).rels.add(RT_HYPERLINK, target, external=True)
    if text is not None:
        attrs["display"] = text
    if tooltip:
        attrs["tooltip"] = tooltip
    links = ws.first("x:hyperlinks")
    if links is None:
        links = x.Hyperlinks()
        ws.add(links)
    for old in list(links):
        if old.get("ref") == ref:
            links.remove(old)
    link = x.Hyperlink(**attrs)
    links.add(link)
    if text is not None or book.cell_element(ws, ref) is None:
        book.set_value(ws, ref, text if text is not None else target,
                       style={"color": "0563C1", "underline": True} if style else None)
    return link


def hyperlinks(book: SpreadsheetDocument, sheet: str) -> dict:
    """``{ref: Hyperlink}``: ``target`` a URL, or ``#Sheet!A1`` for a link
    inside the workbook (kind "location"), as ``add_hyperlink`` takes it."""
    from .records import Hyperlink

    ws = book.sheet(sheet)
    part = book.part_of(ws)
    links = ws.first("x:hyperlinks")
    out: dict = {}
    for link in (links.iterchildren(f"{{{X}}}hyperlink") if links is not None else []):
        rid, ref = link.get(f"{{{R}}}id"), link.get("ref")
        text = link.get("display")
        if text is None:
            value = book.value(sheet, ref) if ref and ":" not in ref else None
            text = None if value is None else str(value)
        tip = link.get("tooltip")
        if rid:
            rel = part.rels.get(rid)
            out[ref] = Hyperlink(text, target=rel.target if rel is not None else None,
                                 tooltip=tip, cell=ref)
        else:
            out[ref] = Hyperlink(text, target="#" + (link.get("location") or ""), tooltip=tip,
                                 kind="location", cell=ref)
    return out

