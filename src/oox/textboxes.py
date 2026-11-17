"""Word text boxes, written the way Word 2010 and later write them.

    run = doc.add_text_box("Note\\nSecond paragraph", width=Cm(6), height=Cm(3),
                           x=Cm(10), y=0, wrap="square", fill="FFF2CC")
    doc.text_boxes()          # [TextBox(text="Note\\nSecond paragraph", ...)]
    doc.set_text_box_text(0, "New text")

Each text box is an ``mc:AlternateContent`` in a run: the choice is a
DrawingML shape (``wps:wsp``, Word 2010+) and the fallback the same box as
VML (``v:shape`` with ``v:textbox``), which is what Word 2007 reads.  Both
hold the text in a ``w:txbxContent`` (paragraphs, tables), so the readers
count each box once and the editors change both copies.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, Any

from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

from . import drawing
from .element import OoxElement
from .grammar import URI
from .units import DRAWING_LIMIT, Cm, emu

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W, WP, V, VO, MC = URI["w"], URI["wp"], URI["v"], URI["o"], URI["mc"]
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_BOX_TYPE = "_x0000_t202"   # VML's text box shape type
WRAPS = ("square", "topAndBottom", "front", "behind")
H_RELATIVE = {"column": "text", "page": "page", "margin": "margin", "character": "char"}
V_RELATIVE = {"paragraph": "text", "page": "page", "margin": "margin", "line": "line"}
H_ALIGN = ("left", "center", "right", "inside", "outside")
V_ALIGN = ("top", "center", "bottom", "inside", "outside")
_Z_BASE = 251659264


def _pt(emu: int) -> str:
    v = int(emu) / 12700
    return f"{v:g}pt"


def _content(text_or_paragraphs: Any) -> list[OoxElement]:
    from . import teletype
    from .ns import w

    if isinstance(text_or_paragraphs, str):
        out = []
        for line in text_or_paragraphs.split("\n"):
            p = w.P()
            if line:
                teletype.addTextToElement(p, line)
            out.append(p)
        return out
    items = [getattr(i, "element", i) for i in text_or_paragraphs]  # Table wrappers too
    if not items:
        return [w.P()]
    return items


def add_text_box(doc: WordprocessingDocument, content: Any = "", *, width: float = Cm(6),
                 height: float = Cm(3), x: float = 0, y: float = 0, h_relative: str = "column",
                 v_relative: str = "paragraph", h_align: str | None = None,
                 v_align: str | None = None, wrap: str = "square",
                 fill: str | None = "FFFFFF", line: str | None = "000000",
                 line_width: float = 0.75, paragraph: OoxElement | None = None,
                 name: str | None = None) -> OoxElement:
    """A floating text box; returns the ``w:r`` holding it.

    ``content`` is text (``\\n`` starts a paragraph) or a list of ``w:p``/
    ``w:tbl`` elements (or ``Table`` objects); they are copied in.  Sizes and
    offsets in points (or ``Length``); ``x``/``y`` are offsets from ``h_relative`` (column, page, margin, character) and
    ``v_relative`` (paragraph, page, margin, line); ``h_align`` (left, center,
    right, inside, outside) and ``v_align`` (top, center, bottom, inside,
    outside) replace them.  ``fill``/``line`` are hex colours or None,
    ``line_width`` in points.  Without ``paragraph`` the box is anchored in a
    new paragraph at the end of the body."""
    from .ns import a, v, w, w10, wp, wps

    if wrap not in WRAPS:
        raise ValueError(f"wrap must be one of {WRAPS}")
    if h_relative not in H_RELATIVE or v_relative not in V_RELATIVE:
        raise ValueError(f"h_relative: {sorted(H_RELATIVE)}; v_relative: {sorted(V_RELATIVE)}")
    if h_align not in (None, *H_ALIGN) or v_align not in (None, *V_ALIGN):
        raise ValueError(f"h_align: {H_ALIGN}; v_align: {V_ALIGN}")
    width, height = emu(width, "width", DRAWING_LIMIT), emu(height, "height", DRAWING_LIMIT)
    x, y = emu(x, "x", DRAWING_LIMIT), emu(y, "y", DRAWING_LIMIT)
    paragraphs = _content(content)
    n = drawing.max_shape_id(doc.document) + 1
    name = name or f"Text Box {n}"
    z = _Z_BASE + sum(1 for _ in doc.document.iter(f"{{{WP}}}anchor"))

    # -- DrawingML (Word 2010+) ----------------------------------------------------
    sp_pr = [a.Xfrm(a.Off(x=0, y=0), a.Ext(cx=width, cy=height)),
             a.PrstGeom(a.AvLst(), prst="rect"),
             a.SolidFill(a.SrgbClr(val=fill)) if fill else a.NoFill(),
             a.Ln(a.SolidFill(a.SrgbClr(val=line)), w=int(line_width * 12700)) if line
             else a.Ln(a.NoFill())]
    shape = wps.Wsp(
        wps.CNvSpPr(txBox=True),
        wps.SpPr(*sp_pr),
        wps.Txbx(w.TxbxContent(*[copy.deepcopy(p) for p in paragraphs])),
        wps.BodyPr(a.NoAutofit(), rot=0, vert="horz", wrap="square", lIns=91440, tIns=45720,
                   rIns=91440, bIns=45720, anchor="t", anchorCtr=False))
    wrap_el = {"square": lambda: wp.WrapSquare(wrapText="bothSides"),
               "topAndBottom": wp.WrapTopAndBottom}.get(wrap, wp.WrapNone)()
    anchor = wp.Anchor(
        wp.SimplePos(x=0, y=0),
        wp.PositionH(wp.Align(text=h_align) if h_align else wp.PosOffset(text=str(x)),
                     relativeFrom=h_relative),
        wp.PositionV(wp.Align(text=v_align) if v_align else wp.PosOffset(text=str(y)),
                     relativeFrom=v_relative),
        wp.Extent(cx=width, cy=height),
        wp.EffectExtent(l=0, t=0, r=0, b=0),
        wrap_el,
        wp.DocPr(id=n, name=name),
        wp.CNvGraphicFramePr(),
        a.Graphic(a.GraphicData(shape, uri=WPS_URI)),
        distT=45720, distB=45720, distL=114300, distR=114300, simplePos=False,
        relativeHeight=z, behindDoc=wrap == "behind", locked=False, layoutInCell=True,
        allowOverlap=True)
    choice_drawing = w.Drawing(anchor)

    # -- VML (Word 2007) -------------------------------------------------------------
    style = ";".join([
        "position:absolute", f"margin-left:{_pt(0 if h_align else x)}",
        f"margin-top:{_pt(0 if v_align else y)}",
        f"width:{_pt(width)}", f"height:{_pt(height)}",
        f"z-index:{-z if wrap == 'behind' else z}",
        *([f"mso-position-horizontal:{h_align}"] if h_align else []),
        f"mso-position-horizontal-relative:{H_RELATIVE[h_relative]}",
        *([f"mso-position-vertical:{v_align}"] if v_align else []),
        f"mso-position-vertical-relative:{V_RELATIVE[v_relative]}"])
    vml_attrs: dict = {"id": name.replace(" ", "_"), "type": f"#{_BOX_TYPE}", "style": style}
    if fill:
        vml_attrs["fillcolor"] = f"#{fill}"
    else:
        vml_attrs["filled"] = False
    if line:
        vml_attrs["strokecolor"] = f"#{line}"
        vml_attrs["strokeweight"] = f"{line_width:g}pt"
    else:
        vml_attrs["stroked"] = False
    kids: list[OoxElement] = [v.Textbox(w.TxbxContent(*[copy.deepcopy(p) for p in paragraphs]))]
    if wrap in ("square", "topAndBottom"):
        kids.append(w10.Wrap(type=wrap))
    vshape = v.Shape(*kids, **vml_attrs)
    vshape.set(f"{{{VO}}}spid", f"_x0000_s{1024 + n}")
    pict_kids: list[OoxElement] = []
    if not any(st.get("id") == _BOX_TYPE for st in doc.document.iter(f"{{{V}}}shapetype")):
        vpath = v.Path(gradientshapeok=True)
        vpath.set(f"{{{VO}}}connecttype", "rect")
        st = v.Shapetype(v.Stroke(joinstyle="miter"), vpath,
                         id=_BOX_TYPE, coordsize="21600,21600", path="m,l,21600r21600,l21600,xe")
        st.set(f"{{{VO}}}spt", "202")
        pict_kids.append(st)
    fallback_pict = w.Pict(*pict_kids, vshape)

    # -- the alternate content in a run ---------------------------------------------------
    ac = etree.Element(f"{{{MC}}}AlternateContent", nsmap={"mc": MC})
    choice = etree.SubElement(ac, f"{{{MC}}}Choice", Requires="wps", nsmap={"wps": WPS_URI})
    choice.append(choice_drawing)
    fb = etree.SubElement(ac, f"{{{MC}}}Fallback")
    fb.append(fallback_pict)
    run = w.R()
    run.append(ac)
    if paragraph is None:
        paragraph = w.P()
        doc.body.add(paragraph)
    paragraph.add(run)
    return run


def _boxes(doc: WordprocessingDocument) -> list[list[OoxElement]]:
    """Each text box as the list of its ``w:txbxContent`` copies."""
    out: list[list[OoxElement]] = []
    for ac in doc.document.iter(f"{{{MC}}}AlternateContent"):
        copies = [t for t in ac.iter(f"{{{W}}}txbxContent")]
        if copies and ac.find(f".//{{{MC}}}AlternateContent") is None:
            out.append(copies)
    for tb in doc.document.iter(f"{{{W}}}txbxContent"):
        if not any(anc.tag == f"{{{MC}}}AlternateContent" for anc in tb.iterancestors()):
            out.append([tb])
    return out


def text_boxes(doc: WordprocessingDocument) -> list:
    """A ``TextBox`` (text, paragraphs, vml_only) for each text box in the
    body, once each."""
    from . import teletype
    from .records import TextBox

    out = []
    for copies in _boxes(doc):
        first = copies[0]
        paras = ["".join(teletype.extractText(p)) for p in first.iterchildren(f"{{{W}}}p")]
        out.append(TextBox("\n".join(paras), paragraphs=len(paras),
                           vml_only=all(c.getparent().tag == f"{{{V}}}textbox" for c in copies)))
    return out


def set_text_box_text(doc: WordprocessingDocument, index: int, content: Any) -> None:
    """Replace the content of text box ``index`` (in ``text_boxes()`` order) in
    every copy (DrawingML and VML)."""
    boxes = _boxes(doc)
    if not 0 <= index < len(boxes):
        raise IndexError(f"text box {index}: the document has {len(boxes)}")
    paragraphs = _content(content)
    for tb in boxes[index]:
        for child in list(tb):
            tb.remove(child)
        for p in paragraphs:
            tb.append(copy.deepcopy(p))
