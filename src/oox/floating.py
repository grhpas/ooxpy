"""Floating (anchored) pictures in Word, and a picture reader.

    doc.add_floating_picture("logo.png", width=Cm(4), h_align="right",
                             h_relative="margin", y=Cm(1), wrap="square")
    doc.add_floating_picture("draft.png", width=Cm(12), h_align="center",
                             v_align="center", h_relative="page",
                             v_relative="page", wrap="behind")

A floating picture sits in a run of an *anchor paragraph* (it moves with that
paragraph) and is positioned relative to the page, margin, column or
paragraph, by offset (EMU, ``oox.units``) or alignment.  Wrapping:
``square``, ``tight``, ``through``, ``topAndBottom``, ``front`` (no wrap, in
front of text) and ``behind``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import drawing
from .element import OoxElement
from .grammar import URI

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
                         width: int | None = None, height: int | None = None,
                         x: int = 0, y: int = 0, h_relative: str = "column",
                         v_relative: str = "paragraph", h_align: str | None = None,
                         v_align: str | None = None, wrap: str = "square",
                         side: str = "bothSides", paragraph: OoxElement | None = None,
                         descr: str = "", margin: int = 114300,
                         fallback: Any = None) -> OoxElement:
    """Insert a floating picture; returns the ``w:r`` holding it.

    ``x``/``y`` are offsets in EMU from the ``h_relative``/``v_relative``
    reference; ``h_align`` (left, center, right, inside, outside) and
    ``v_align`` (top, center, bottom, inside, outside) replace them.
    ``margin`` is the gap between picture and text on the left and right.
    """
    from .ns import a, pic, w, wp

    if wrap not in WRAPS:
        raise ValueError(f"wrap must be one of {WRAPS}")
    rid, info, svg_rid = doc.add_picture_image(image, fallback)
    cx, cy = info.scaled(width, height)
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


def pictures(doc: WordprocessingDocument) -> list[dict[str, Any]]:
    """Pictures in the body, in document order: ``{"floating", "name",
    "descr", "width", "height", "wrap"}`` (EMU; wrap is ``"inline"`` for
    inline pictures)."""
    out = []
    for container in doc.body.iter(f"{{{WP}}}inline", f"{{{WP}}}anchor"):
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
        out.append({
            "floating": floating,
            "name": pr.get("name") if pr is not None else None,
            "descr": pr.get("descr") if pr is not None else None,
            "width": int(ext.get("cx", "0")) if ext is not None else 0,
            "height": int(ext.get("cy", "0")) if ext is not None else 0,
            "wrap": wrap,
        })
    return out
