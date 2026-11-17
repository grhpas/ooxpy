"""Slide content beyond text boxes: shapes with fills, tables, speaker notes,
picture placeholders, and a reader for what a slide holds.

    deck.add_shape(slide, "roundRect", x=72, y=144, width=216, height=72,
                   text="Step 1", fill="2F5496")
    deck.add_table(slide, [["Region", "Q1"], ["North", 120]], y=Inches(3.5),
                   width=Inches(6), style="Medium Style 2 - Accent 1")
    deck.set_notes(slide, "Mention the North region first.")
    deck.set_placeholder_picture(slide, "pic", "photo.jpg")   # crops to fill, like PowerPoint
    deck.shapes(slide)    # [Shape(kind="table", name=..., text=...), ...]

Tables use PowerPoint's built-in table styles, referenced by GUID
(``TABLE_STYLES`` names a few; any GUID works).  Notes need a notes master,
which is created on first use (decks from PowerPoint usually have one).
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from . import drawing
from .element import OoxElement
from .grammar import URI
from .ns import a, p
from .units import DRAWING_LIMIT, Inches, emu

if TYPE_CHECKING:
    from .presentation import PresentationDocument

TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"
P, A = URI["p"], URI["a"]
# PowerPoint's built-in table styles (Design > Table Styles), by GUID
TABLE_STYLES = {
    "Medium Style 2 - Accent 1": "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    "Medium Style 2 - Accent 2": "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}",
    "Medium Style 2 - Accent 3": "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
    "Medium Style 2 - Accent 4": "{00A15C55-8517-42AA-B614-E9B94910E393}",
    "Medium Style 2 - Accent 5": "{7DF18680-E054-41AD-8BC1-D1AEF772440D}",
    "Medium Style 2 - Accent 6": "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
    "Medium Style 2": "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
    "Light Style 1 - Accent 1": "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
    "No Style, Table Grid": "{5940675A-B579-460E-94D1-54222C63F5DA}",
    "No Style, No Grid": "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
}
# preset geometry -> PowerPoint's MsoAutoShapeType (for the Office check)
SHAPES = {"rect": 1, "roundRect": 5, "ellipse": 9, "triangle": 7, "diamond": 4,
          "hexagon": 10, "octagon": 6, "rightArrow": 33, "leftArrow": 34, "upArrow": 35,
          "downArrow": 36, "homePlate": 51, "chevron": 52}
_ALIGN = {"left": "l", "center": "ctr", "right": "r", "justify": "just"}


def _paragraphs(text: str, *, size: float | None = None, bold: bool = False,
                color: str | None = None, align: str | None = None,
                lang: str = "en-US") -> list[OoxElement]:
    out = []
    for line in str(text).split("\n"):
        rpr = a.RPr(lang=lang, sz=int(size * 100) if size else None,
                    b=True if bold else None, dirty=False)
        if color:
            rpr.add(a.SolidFill(a.SrgbClr(val=color)))
        para = a.P(a.R(rpr, a.T(line))) if line else a.P(a.EndParaRPr(lang=lang))
        if align:
            para.add(a.PPr(algn=_ALIGN[align]))
        out.append(para)
    return out


def _text_of(container) -> str:
    """Paragraph texts (``a:p``) under a txBody, joined with newlines."""
    lines = []
    for para in container.iter(f"{{{A}}}p"):
        lines.append("".join(t.text or "" for t in para.iter(f"{{{A}}}t")))
    return "\n".join(lines)


# -- shapes -------------------------------------------------------------------------
def add_shape(deck: PresentationDocument, slide: OoxElement, shape: str,
              box: tuple[int, int, int, int], *, text: str | None = None, fill: str | None = "4472C4",
              line: str | None = None, line_width: float = 1.0, size: float | None = None,
              bold: bool = False, color: str | None = "FFFFFF", align: str | None = "center",
              name: str | None = None) -> OoxElement:
    """An auto shape (``shape`` = preset geometry: rect, roundRect, ellipse,
    rightArrow, chevron, ...).  ``fill``/``line`` as hex RGB or ``None``;
    ``line_width`` in points; text centred and white by default.  ``box`` is
    (x, y, cx, cy) in EMU (``deck.add_shape`` takes points)."""
    if shape not in SHAPES:
        raise ValueError(f"shape {shape!r}: use one of {sorted(SHAPES)}")
    n = drawing.max_shape_id(slide) + 1
    sp_pr = drawing.xfrm(*box)
    geom = a.PrstGeom(a.AvLst(), prst=shape)
    fill_el = a.SolidFill(a.SrgbClr(val=fill)) if fill else a.NoFill()
    ln = a.Ln(a.SolidFill(a.SrgbClr(val=line)), w=int(line_width * 12700)) if line \
        else a.Ln(a.NoFill())
    body = p.TxBody(a.BodyPr(rtlCol=False, anchor="ctr"), a.LstStyle(),
                    *(_paragraphs(text, size=size, bold=bold, color=color, align=align)
                      if text else [a.P(a.EndParaRPr(lang="en-US"))]))
    sp = p.Sp(
        p.NvSpPr(p.CNvPr(id=n, name=name or f"{shape} {n}"), p.CNvSpPr(), p.NvPr()),
        p.SpPr(sp_pr, geom, fill_el, ln),
        body,
    )
    deck._tree(slide).add(sp)
    return sp


# -- tables -------------------------------------------------------------------------
def add_table(deck: PresentationDocument, slide: OoxElement, rows,
              box: tuple[int, int, int, int | None], *, header: bool = True,
              banded: bool = True, first_column: bool = False,
              style: str | None = "Medium Style 2 - Accent 1",
              widths: list[float] | None = None, row_height: float = Inches(0.4),
              size: float | None = None, merges=(), name: str | None = None) -> OoxElement:
    """A table from a list of rows (values become text; ``None`` is empty).

    ``style`` is a name from ``TABLE_STYLES`` or a GUID; ``header``,
    ``banded``, ``first_column`` switch the style's options.  ``merges`` are
    ``"A1:B1"``-style ranges (columns A.., rows 1..).  ``box`` is (x, y, cx,
    cy) in EMU, cy None for rows x ``row_height``; ``widths`` (per column) and
    ``row_height`` in points or ``Length``.  Returns the ``p:graphicFrame``.
    """
    from .xlsheet import parse_range

    rows = [list(r) for r in rows]
    if not rows or not rows[0]:
        raise ValueError("a table needs at least one row and one column")
    ncols = max(len(r) for r in rows)
    for r in rows:
        r.extend([None] * (ncols - len(r)))
    x, y, cx, cy = box
    cols = ([emu(v, "widths", DRAWING_LIMIT) for v in widths] if widths
            else [cx // ncols] * ncols)
    if len(cols) != ncols:
        raise ValueError(f"{len(cols)} column widths for {ncols} columns")
    height = emu(row_height, "row_height", DRAWING_LIMIT) if cy is None else cy // len(rows)
    span: dict[tuple[int, int], tuple[int, int]] = {}
    covered: dict[tuple[int, int], tuple[bool, bool]] = {}
    for ref in merges:
        c1, r1, c2, r2 = parse_range(ref)
        if c2 > ncols or r2 > len(rows):
            raise ValueError(f"merge {ref} is outside the {len(rows)}x{ncols} table")
        span[(r1 - 1, c1 - 1)] = (r2 - r1 + 1, c2 - c1 + 1)
        for rr in range(r1 - 1, r2):
            for cc in range(c1 - 1, c2):
                if (rr, cc) != (r1 - 1, c1 - 1):
                    covered[(rr, cc)] = (rr > r1 - 1, cc > c1 - 1)
    guid = TABLE_STYLES.get(style, style) if style else None
    tbl_pr = a.TblPr(firstRow=True if header else None, bandRow=True if banded else None,
                     firstCol=True if first_column else None)
    if guid:
        tbl_pr.add(a.TableStyleId(text=guid))
    tbl = a.Tbl(tbl_pr, a.TblGrid(*[a.GridCol(w=w) for w in cols]))
    for ri, row in enumerate(rows):
        tr = a.Tr(h=height)
        for ci, value in enumerate(row):
            text = "" if value is None else str(value)
            tc = a.Tc(a.TxBody(a.BodyPr(), a.LstStyle(), *_paragraphs(text, size=size)),
                      a.TcPr())
            if (ri, ci) in span:
                rs, cs = span[(ri, ci)]
                if cs > 1:
                    tc.set_attr("gridSpan", cs)
                if rs > 1:
                    tc.set_attr("rowSpan", rs)
            if (ri, ci) in covered:
                v, h = covered[(ri, ci)]
                if h:
                    tc.set_attr("hMerge", True)
                if v:
                    tc.set_attr("vMerge", True)
            tr.append(tc)
        tbl.append(tr)
    n = drawing.max_shape_id(slide) + 1
    frame = p.GraphicFrame(
        p.NvGraphicFramePr(p.CNvPr(id=n, name=name or f"Table {n}"),
                           p.CNvGraphicFramePr(a.GraphicFrameLocks(noGrp=True)), p.NvPr()),
        p.Xfrm(a.Off(x=x, y=y), a.Ext(cx=sum(cols), cy=height * len(rows))),
        a.Graphic(a.GraphicData(tbl, uri=TABLE_URI)),
    )
    deck._tree(slide).add(frame)
    return frame


def table_info(frame: OoxElement) -> dict:
    """Rows of cell texts (merged-over cells read as ""), the style GUID and
    name, and the header/banding switches of a table graphic frame."""
    tbl = frame.find(f".//{{{A}}}tbl")
    if tbl is None:
        raise ValueError("not a table")
    pr = tbl.find(f"{{{A}}}tblPr")
    style_el = pr.find(f"{{{A}}}tableStyleId") if pr is not None else None
    guid = style_el.text if style_el is not None else None
    rows, merges = [], []
    for ri, tr in enumerate(tbl.findall(f"{{{A}}}tr")):
        cells = []
        for ci, tc in enumerate(tr.findall(f"{{{A}}}tc")):
            hidden = tc.get("hMerge") in ("1", "true") or tc.get("vMerge") in ("1", "true")
            cells.append("" if hidden else _text_of(tc))
            gs, rs = int(tc.get("gridSpan", "1")), int(tc.get("rowSpan", "1"))
            if (gs > 1 or rs > 1) and not hidden:   # PowerPoint repeats gridSpan below
                from .spreadsheet import cell_ref

                merges.append(f"{cell_ref(ci + 1, ri + 1)}:{cell_ref(ci + gs, ri + rs)}")
        rows.append(cells)

    def on(attr):
        return pr is not None and pr.get(attr) in ("1", "true")

    names = {v: k for k, v in TABLE_STYLES.items()}
    return {"rows": rows, "merges": merges, "style": guid, "style_name": names.get(guid or ""),
            "header": on("firstRow"), "banded": on("bandRow"), "first_column": on("firstCol")}


# -- speaker notes ------------------------------------------------------------------
def _notes_placeholder(shape_id: int, name: str, ph: OoxElement, geometry=None,
                       text: str | None = None, locks: dict | None = None) -> OoxElement:
    sp_pr = p.SpPr(drawing.xfrm(*geometry), a.PrstGeom(a.AvLst(), prst="rect")) if geometry \
        else p.SpPr()
    sp = p.Sp(p.NvSpPr(p.CNvPr(id=shape_id, name=name),
                       p.CNvSpPr(a.SpLocks(noGrp=True, **(locks or {}))), p.NvPr(ph)),
              sp_pr)
    if text is not None or ph.get("type") != "sldImg":
        sp.append(p.TxBody(a.BodyPr(), a.LstStyle(),
                           *(_paragraphs(text) if text else [a.P(a.EndParaRPr(lang="en-US"))])))
    return sp


def _notes_master() -> OoxElement:
    tree = p.SpTree(
        p.NvGrpSpPr(p.CNvPr(id=1, name=""), p.CNvGrpSpPr(), p.NvPr()),
        p.GrpSpPr(a.Xfrm(a.Off(x=0, y=0), a.Ext(cx=0, cy=0),
                         a.ChOff(x=0, y=0), a.ChExt(cx=0, cy=0))),
        _notes_placeholder(2, "Slide Image Placeholder 1", p.Ph(type="sldImg", idx=2),
                           (685800, 1143000, 5486400, 3086100), None,
                           {"noRot": True, "noChangeAspect": True}),
        _notes_placeholder(3, "Notes Placeholder 2", p.Ph(type="body", sz="quarter", idx=3),
                           (685800, 4400550, 5486400, 3600450), ""),
    )
    level = a.Lvl1pPr(a.DefRPr(a.SolidFill(a.SchemeClr(val="tx1")), a.Latin(typeface="+mn-lt"),
                               sz=1200, kern=1200), marL=0, algn="l")
    return p.NotesMaster(
        p.CSld(p.Bg(p.BgRef(a.SchemeClr(val="bg1"), idx=1001)), tree),
        p.ClrMap(bg1="lt1", tx1="dk1", bg2="lt2", tx2="dk2", accent1="accent1",
                 accent2="accent2", accent3="accent3", accent4="accent4", accent5="accent5",
                 accent6="accent6", hlink="hlink", folHlink="folHlink"),
        p.NotesStyle(level),
    )


def _notes_master_part(deck: PresentationDocument):
    found = deck.related_parts(deck.main_part, "NotesMasterPart")
    if found:
        return found[0]
    part, rid = deck.add_part("NotesMasterPart", element=_notes_master(),
                              folder="/ppt/notesMasters")
    lst = deck.presentation.first("p:notesMasterIdLst")
    if lst is None:
        lst = p.NotesMasterIdLst()
        deck.presentation.add(lst)       # schema order: after sldMasterIdLst
    lst.add(p.NotesMasterId(r_id=rid))
    deck.add_part("ThemePart", part, element=drawing.default_theme(), folder="/ppt/theme")
    return part


def _notes_part(deck: PresentationDocument, slide: OoxElement, create: bool):
    sp = deck.part_of(slide)
    found = deck.related_parts(sp, "NotesSlidePart")
    if found or not create:
        return found[0] if found else None
    master = _notes_master_part(deck)
    tree = p.SpTree(
        p.NvGrpSpPr(p.CNvPr(id=1, name=""), p.CNvGrpSpPr(), p.NvPr()),
        p.GrpSpPr(a.Xfrm(a.Off(x=0, y=0), a.Ext(cx=0, cy=0),
                         a.ChOff(x=0, y=0), a.ChExt(cx=0, cy=0))),
        _notes_placeholder(2, "Slide Image Placeholder 1", p.Ph(type="sldImg"), None, None,
                           {"noRot": True, "noChangeAspect": True}),
        _notes_placeholder(3, "Notes Placeholder 2", p.Ph(type="body", idx=1), None, ""),
    )
    notes = p.Notes(p.CSld(tree), p.ClrMapOvr(a.MasterClrMapping()))
    part, _ = deck.add_part("NotesSlidePart", sp, element=notes, folder="/ppt/notesSlides")
    deck.relate(part, master, "NotesMasterPart")
    deck.relate(part, sp, "SlidePart")
    return part


def _notes_body(notes: OoxElement) -> OoxElement | None:
    for sp in notes.iter(f"{{{P}}}sp"):
        ph = sp.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph")
        if ph is not None and ph.get("type") == "body":
            return sp
    return None


def set_notes(deck: PresentationDocument, slide: OoxElement, text: str) -> OoxElement:
    """Set a slide's speaker notes (``\\n`` starts a new paragraph); returns
    the notes body shape."""
    part = _notes_part(deck, slide, create=True)
    assert part is not None
    notes = part.element
    body = _notes_body(notes)
    if body is None:  # a notes slide without a body placeholder: add one
        body = _notes_placeholder(drawing.max_shape_id(notes) + 1, "Notes Placeholder",
                                  p.Ph(type="body", idx=1), None, "")
        notes.need("p:cSld").need("p:spTree").add(body)
    tx = body.first("p:txBody")
    if tx is None:
        tx = p.TxBody(a.BodyPr(), a.LstStyle())
        body.add(tx)
    for old in list(tx.iterchildren(f"{{{A}}}p")):
        tx.remove(old)
    for para in _paragraphs(text):
        tx.append(para)
    return body


def notes(deck: PresentationDocument, slide: OoxElement) -> str | None:
    """A slide's speaker notes, or ``None`` if it has none."""
    part = _notes_part(deck, slide, create=False)
    if part is None:
        return None
    body = _notes_body(part.element)
    return _text_of(body) if body is not None else None


# -- picture placeholders -----------------------------------------------------------
def _xfrm_of(shape) -> tuple[int, int, int, int] | None:
    off = shape.find(f".//{{{A}}}xfrm/{{{A}}}off")
    ext = shape.find(f".//{{{A}}}xfrm/{{{A}}}ext")
    if off is None or ext is None:
        return None
    return int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))


def _inherited_box(deck: PresentationDocument, slide: OoxElement, ph_type: str, idx: int):
    """Position of a placeholder: its own, else the layout's, else the master's
    (matched by idx, then by type)."""
    from .presentation import placeholder_info

    sp = deck.part_of(slide)
    layouts = deck.related_parts(sp, "SlideLayoutPart")
    chain = []
    if layouts:
        chain.append(layouts[0].element)
        masters = deck.related_parts(layouts[0], "SlideMasterPart")
        if masters:
            chain.append(masters[0].element)
    for root in chain:
        tree = root.need("p:cSld").need("p:spTree")
        shapes = [(placeholder_info(s), s) for s in tree if isinstance(s, OoxElement)]
        for match in (lambda i: i[1] == idx, lambda i: i[0] == ph_type,
                      lambda i: ph_type in ("pic", "obj") and i[0] in ("body", "obj")):
            for info, s in shapes:
                if info is not None and match(info) and _xfrm_of(s):
                    return _xfrm_of(s)
    return None


def set_placeholder_picture(deck: PresentationDocument, slide: OoxElement, which, image, *,
                            crop: bool = True, descr: str = "") -> OoxElement:
    """Put a picture into a picture (or content) placeholder.  With ``crop``
    the picture is cropped to fill the placeholder, as PowerPoint does;
    otherwise it is stretched.  Returns the new ``p:pic``."""
    from .presentation import placeholder_info

    shape = deck.placeholder(slide, which)
    ph_type, idx = placeholder_info(shape) or ("obj", 0)
    rid, info = deck.add_image(image, deck.part_of(slide))
    box = _xfrm_of(shape) or _inherited_box(deck, slide, ph_type, idx)
    src = None
    if crop and box and box[3] and info.width_px and info.height_px:
        w, h = info.native_size()          # EMU: respects the image's own dpi per axis
        box_ratio, img_ratio = box[2] / box[3], w / h
        if img_ratio > box_ratio:      # wider: crop left and right
            cut = round((1 - box_ratio / img_ratio) / 2 * 100000)
            src = a.SrcRect(l=cut, r=cut)
        elif img_ratio < box_ratio:    # taller: crop top and bottom
            cut = round((1 - img_ratio / box_ratio) / 2 * 100000)
            src = a.SrcRect(t=cut, b=cut)
    nv = shape.first("p:nvSpPr")
    cnv = nv.first("p:cNvPr") if nv is not None else None
    ph = copy.deepcopy(next(shape.iter(f"{{{P}}}ph")))
    blip_fill = p.BlipFill(drawing.blip(rid), *([src] if src is not None else []),
                           a.Stretch(a.FillRect()))
    old_pr = shape.first("p:spPr")
    sp_pr: OoxElement = copy.deepcopy(old_pr) if old_pr is not None else p.SpPr()
    if sp_pr.find(f"{{{A}}}prstGeom") is None and sp_pr.find(f"{{{A}}}custGeom") is None:
        # PowerPoint 2007 draws nothing for a picture without a geometry when
        # neither layout nor master gives one (Office's master does; oox's and
        # many templates don't): give it its own rectangle
        sp_pr.add(a.PrstGeom(a.AvLst(), prst="rect"))
    pic = p.Pic(
        p.NvPicPr(p.CNvPr(id=int(cnv.get("id")) if cnv is not None else
                          drawing.max_shape_id(slide) + 1,
                          name=cnv.get("name") if cnv is not None else "Picture",
                          descr=descr or None),
                  p.CNvPicPr(a.PicLocks(noGrp=True, noChangeAspect=True)), p.NvPr(ph)),
        blip_fill,
        sp_pr,
    )
    parent = shape.getparent()
    parent.replace(shape, pic)
    return pic


# -- reading ------------------------------------------------------------------------
def _position(el: OoxElement, local: str) -> tuple:
    from .units import points

    xfrm = (el.find(f"{{{P}}}xfrm") if local == "graphicFrame"
            else el.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm") if local == "grpSp"
            else el.find(f"{{{P}}}spPr/{{{A}}}xfrm"))
    off = xfrm.find(f"{{{A}}}off") if xfrm is not None else None
    ext = xfrm.find(f"{{{A}}}ext") if xfrm is not None else None
    if off is None or ext is None:
        return None, None, None, None
    return (points(int(off.get("x", "0"))), points(int(off.get("y", "0"))),
            points(int(ext.get("cx", "0"))), points(int(ext.get("cy", "0"))))


def shapes(deck: PresentationDocument, slide: OoxElement) -> list:
    """A ``Shape`` for what a slide holds, in z-order: kind (placeholder,
    text_box, shape, picture, table, chart, group, other), name, text,
    preset geometry, ``Placeholder``, position and size in points."""
    from .presentation import placeholder_info
    from .records import Placeholder, Shape

    out = []
    for el in deck._tree(slide):
        if not isinstance(el.tag, str):
            continue
        local = el.tag.split("}")[1]
        if local in ("nvGrpSpPr", "grpSpPr", "extLst"):
            continue
        cnv = el.find(f".//{{{P}}}cNvPr")
        name = cnv.get("name") if cnv is not None else None
        info = placeholder_info(el)
        text, geometry = None, None
        if local == "sp":
            geom = el.find(f"{{{P}}}spPr/{{{A}}}prstGeom")
            geometry = geom.get("prst") if geom is not None else None
            body = el.find(f"{{{P}}}txBody")
            text = _text_of(body) if body is not None else None
            box = el.find(f"{{{P}}}nvSpPr/{{{P}}}cNvSpPr")
            if info is not None:
                kind = "placeholder"
            elif box is not None and box.get("txBox") in ("1", "true"):
                kind = "text_box"
            else:
                kind = "shape"
        elif local == "pic":
            kind = "picture"
        elif local == "graphicFrame":
            data = el.find(f".//{{{A}}}graphicData")
            uri = data.get("uri") if data is not None else ""
            kind = {"http://schemas.openxmlformats.org/drawingml/2006/table": "table",
                    "http://schemas.openxmlformats.org/drawingml/2006/chart": "chart"
                    }.get(uri, "other")
            if kind == "table":
                text = "\n".join("\t".join(r) for r in table_info(el)["rows"])
        elif local == "grpSp":
            kind = "group"
        else:
            kind = "other"
        x, y, width, height = _position(el, local)
        out.append(Shape(kind, name=name, text=text,
                         placeholder=Placeholder(info[0], info[1], name, el) if info else None,
                         geometry=geometry, x=x, y=y, width=width, height=height, element=el))
    return out


def tables(deck: PresentationDocument, slide: OoxElement) -> list:
    """A ``SlideTableInfo`` for each table on a slide (cell texts, merges,
    style and switches, name)."""
    from .records import SlideTableInfo

    out = []
    for frame in deck._tree(slide).iterchildren(f"{{{P}}}graphicFrame"):
        if frame.find(f".//{{{A}}}tbl") is not None:
            info = table_info(frame)
            cnv = frame.find(f".//{{{P}}}cNvPr")
            out.append(SlideTableInfo(name=cnv.get("name") if cnv is not None else None, **info))
    return out
