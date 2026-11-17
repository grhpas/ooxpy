"""Floating (anchored) pictures in Word, and a picture reader.

    doc.add_floating_picture("logo.png", width=Cm(4), h_align="right",
                             h_relative="margin", y=Cm(1), wrap="square")
    doc.add_floating_picture("draft.png", width=Cm(12), h_align="center",
                             v_align="center", h_relative="page",
                             v_relative="page", wrap="behind")

A floating picture sits in a run of an *anchor paragraph* (it moves with that
paragraph) and is positioned relative to the page, margin, column or
paragraph, by offset (points or a ``Length``) or alignment.  Wrapping:
``square``, ``tight``, ``through``, ``topAndBottom``, ``front`` (no wrap, in
front of text) and ``behind``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import drawing
from .element import OoxElement
from .grammar import URI
from .units import DRAWING_LIMIT, PAGE_LIMIT, emu

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
WP = URI["wp"]
H_RELATIVE = ("character", "column", "insideMargin", "leftMargin", "margin", "outsideMargin",
              "page", "rightMargin")
V_RELATIVE = ("bottomMargin", "insideMargin", "line", "margin", "outsideMargin", "page",
              "paragraph", "topMargin")
WRAPS = ("square", "tight", "through", "topAndBottom", "front", "behind")
_Z_BASE = 251658240  # Word's base for relativeHeight


def _position(factory: Any, relative: str, allowed: tuple[str, ...], offset: int,
              align: str | None) -> OoxElement:
    from .ns import wp

    if relative not in allowed:
        raise ValueError(f"relative position must be one of {allowed}")
    inner = wp.Align(align) if align else wp.PosOffset(str(int(offset)))
    return factory(inner, relativeFrom=relative)


def add_floating_picture(doc: WordprocessingDocument, image: Any, *,
                         width: float | None = None, height: float | None = None,
                         x: float = 0, y: float = 0, h_relative: str = "column",
                         v_relative: str = "paragraph", h_align: str | None = None,
                         v_align: str | None = None, wrap: str = "square",
                         side: str = "bothSides", paragraph: OoxElement | None = None,
                         descr: str = "", margin: float = 9,
                         fallback: Any = None) -> OoxElement:
    """Insert a floating picture; returns the ``w:r`` holding it.

    Sizes and offsets in points (or ``Length``).  ``x``/``y`` are offsets
    from the ``h_relative``/``v_relative`` reference; ``h_align`` (left,
    center, right, inside, outside) and ``v_align`` (top, center, bottom,
    inside, outside) replace them.  ``margin`` is the gap between picture and
    text on the left and right.
    """
    from .document import picture_size
    from .ns import a, pic, w, wp

    if wrap not in WRAPS:
        raise ValueError(f"wrap must be one of {WRAPS}")
    x, y = emu(x, "x", DRAWING_LIMIT), emu(y, "y", DRAWING_LIMIT)
    margin = emu(margin, "margin", PAGE_LIMIT)
    rid, info, svg_rid = doc.add_picture_image(image, fallback)
    cx, cy = picture_size(info, width, height)
    n = drawing.max_shape_id(doc.document) + 1
    name = f"Picture {n}"
    z = _Z_BASE + sum(1 for _ in doc.document.iter(f"{{{WP}}}anchor"))
    children: list[OoxElement] = [
        wp.SimplePos(x=0, y=0),
        _position(wp.PositionH, h_relative, H_RELATIVE, x, h_align),
        _position(wp.PositionV, v_relative, V_RELATIVE, y, v_align),
        wp.Extent(cx=cx, cy=cy),
        wp.EffectExtent(l=0, t=0, r=0, b=0),
    ]
    if wrap in ("front", "behind"):
        children.append(wp.WrapNone())
    elif wrap == "topAndBottom":
        children.append(wp.WrapTopAndBottom())
    elif wrap == "square":
        children.append(wp.WrapSquare(wrapText=side))
    else:  # tight / through need a wrap polygon: the picture's rectangle
        poly = wp.WrapPolygon(wp.Start(x=0, y=0), wp.LineTo(x=0, y=21600),
                              wp.LineTo(x=21600, y=21600), wp.LineTo(x=21600, y=0),
                              wp.LineTo(x=0, y=0), edited=False)
        children.append((wp.WrapTight if wrap == "tight" else wp.WrapThrough)(poly, wrapText=side))
    children += [
        wp.DocPr(id=n, name=name, descr=descr or None),
        wp.CNvGraphicFramePr(a.GraphicFrameLocks(noChangeAspect=True)),
        a.Graphic(a.GraphicData(drawing.picture(pic, 0, name, rid, cx, cy, descr=descr,
                                                svg_rid=svg_rid),
                                uri=drawing.PIC_URI)),
    ]
    anchor = wp.Anchor(*children, distT=0, distB=0, distL=margin, distR=margin,
                       simplePos=False, relativeHeight=z, behindDoc=wrap == "behind",
                       locked=False, layoutInCell=True, allowOverlap=True)
    run = w.R(w.Drawing(anchor))
    if paragraph is None:
        doc.body.add(w.P(run))
    else:
        paragraph.add(run)
    return run


PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_A, _R, _V, _W10 = URI["a"], URI["r"], URI["v"], URI["w10"]
_SVG_BLIP = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"
_PT = 12700  # EMU per point


def _blip_ids(blip: OoxElement | None) -> dict[str, str | None]:
    if blip is None:
        return {"embed": None, "link": None, "svg": None}
    svg = next(blip.iter(_SVG_BLIP), None)
    return {"embed": blip.get(f"{{{_R}}}embed"), "link": blip.get(f"{{{_R}}}link"),
            "svg": svg.get(f"{{{_R}}}embed") if svg is not None else None}


def _drawing_picture(container: OoxElement) -> dict[str, Any] | None:
    data = container.find(f"{{{_A}}}graphic/{{{_A}}}graphicData")
    if data is None or data.get("uri") != PICTURE_URI:
        return None  # a chart, a shape, a text box, a group ...
    floating = container.tag == f"{{{WP}}}anchor"
    ext = container.first("wp:extent")
    pr = container.first("wp:docPr")
    wrap = "inline"
    if floating:
        for child in container:
            local = child.tag.rpartition("}")[2]
            if local.startswith("wrap"):
                wrap = local[4].lower() + local[5:]
                break
        if wrap == "none":
            wrap = "behind" if container.get("behindDoc") in ("1", "true") else "front"
    return {"kind": "drawing", "floating": floating,
            "name": pr.get("name") if pr is not None else None,
            "descr": pr.get("descr") if pr is not None else None,
            "width": int(ext.get("cx", "0")) if ext is not None else 0,
            "height": int(ext.get("cy", "0")) if ext is not None else 0,
            "wrap": wrap, **_blip_ids(next(data.iter(f"{{{_A}}}blip"), None)),
            "element": container}


def _style(shape: OoxElement) -> dict[str, str]:
    out = {}
    for item in (shape.get("style") or "").split(";"):
        key, _, value = item.partition(":")
        if key.strip():
            out[key.strip().lower()] = value.strip()
    return out


def _length(value: str | None) -> int:
    """A CSS length from a VML style ("72pt", "1in", "2.5cm", "96px") in EMU."""
    if not value:
        return 0
    units = {"pt": _PT, "in": 914400, "cm": 360000, "mm": 36000, "px": 9525, "pc": 12 * _PT}
    for unit, per in units.items():
        if value.endswith(unit):
            try:
                return round(float(value[:-len(unit)]) * per)
            except ValueError:
                return 0
    try:
        return round(float(value) * 9525)  # unitless: pixels
    except ValueError:
        return 0


def _vml_picture(imagedata: OoxElement) -> dict[str, Any]:
    shape = imagedata.getparent()
    style = _style(shape) if shape is not None else {}
    floating = style.get("position") == "absolute"
    wrap = "inline"
    if floating:
        w10 = shape.find(f"{{{_W10}}}wrap") if shape is not None else None
        kind = w10.get("type") if w10 is not None else None
        wrap = {"square": "square", "tight": "tight", "through": "through",
                "topAndBottom": "topAndBottom"}.get(kind or "", "")
        if not wrap:
            z = style.get("z-index", "0")
            wrap = "behind" if z.startswith("-") else "front"
    return {"kind": "vml", "floating": floating,
            "name": shape.get("id") if shape is not None else None,
            "descr": (shape.get("alt") if shape is not None else None)
            or imagedata.get(f"{{{URI['o']}}}title"),
            "width": _length(style.get("width")), "height": _length(style.get("height")),
            "wrap": wrap, "embed": imagedata.get(f"{{{_R}}}id"), "link": None, "svg": None,
            "element": shape if shape is not None else imagedata}


def pictures(doc: WordprocessingDocument) -> list:
    """Pictures in the body, in document order, as Word shows them: DrawingML
    pictures (inline or floating) and VML ones (``v:imagedata``, in older files
    and object previews).  Charts, shapes and text boxes aren't pictures, and
    the copy Word keeps of a text box for older readers is skipped.

    Each is a ``Picture`` (kind "drawing" or "vml", floating, name, descr,
    width, height, wrap, embed, link, svg, element): sizes in points,
    ``wrap`` is ``"inline"`` for inline pictures; ``embed`` is the
    relationship id of the stored image, ``link`` that of a linked (external)
    one, ``svg`` that of an SVG version.  ``doc.image_of(picture)`` returns the
    image itself."""
    from .records import Picture
    from .units import points
    from .walk import iter_story

    out = []
    for el in iter_story(doc.body, f"{{{WP}}}inline", f"{{{WP}}}anchor", f"{{{_V}}}imagedata"):
        found = _vml_picture(el) if el.tag == f"{{{_V}}}imagedata" else _drawing_picture(el)
        if found is not None:
            found["width"], found["height"] = points(found["width"]), points(found["height"])
            out.append(Picture(**found))
    return out
