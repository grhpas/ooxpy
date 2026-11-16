"""Shared DrawingML builders (pictures, shapes, the default theme).

The same picture structure appears under three namespaces:
``pic:pic`` (Word), ``xdr:pic`` (Excel) and ``p:pic`` (PowerPoint).  The
builders take the factory module so one implementation serves all three.
"""

from __future__ import annotations

from types import ModuleType

from .element import OoxElement
from .ns import a

PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def xfrm(x: int, y: int, cx: int, cy: int) -> OoxElement:
    return a.Xfrm(a.Off(x=x, y=y), a.Ext(cx=cx, cy=cy))


def sp_pr(mod: ModuleType, x: int, y: int, cx: int, cy: int, *extra: OoxElement) -> OoxElement:
    return mod.SpPr(xfrm(x, y, cx, cy), a.PrstGeom(a.AvLst(), prst="rect"), *extra)


SVG_BLIP_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"  # Office 2016+ reads this extension


def blip(rid: str, svg_rid: str | None = None) -> OoxElement:
    """``a:blip`` for ``rid``; with ``svg_rid`` the SVG goes in the blip's
    extension list (Office 2016+) and ``rid`` is the raster fallback."""
    b = a.Blip(r_embed=rid)
    if svg_rid:
        from .ns import asvg

        b.add(a.ExtLst(a.Ext(asvg.SvgBlip(r_embed=svg_rid), uri=SVG_BLIP_URI)))
    return b


def picture(mod: ModuleType, shape_id: int, name: str, rid: str, cx: int, cy: int,
            x: int = 0, y: int = 0, descr: str = "", nv_pr: bool = False,
            svg_rid: str | None = None) -> OoxElement:
    """``<mod:pic>`` showing the image related as ``rid``.

    ``nv_pr=True`` adds ``p:nvPr``, which PresentationML requires;
    ``svg_rid`` adds an SVG version shown by Office 2016+ (``rid`` is then the
    fallback).
    """
    nv = mod.NvPicPr(
        mod.CNvPr(id=shape_id, name=name, descr=descr or None),
        mod.CNvPicPr(a.PicLocks(noChangeAspect=True)),
    )
    if nv_pr:
        nv.add(mod.NvPr())
    return mod.Pic(
        nv,
        mod.BlipFill(blip(rid, svg_rid), a.Stretch(a.FillRect())),
        sp_pr(mod, x, y, cx, cy),
    )


def max_shape_id(root: OoxElement) -> int:
    """Highest ``id`` on any non-visual drawing properties in a part."""
    best = 0
    for el in root.iter():
        if isinstance(el.tag, str) and el.tag.endswith(("}cNvPr", "}docPr")):
            try:
                best = max(best, int(el.get("id", "0")))
            except ValueError:
                pass
    return best


def default_theme(name: str = "Office Theme") -> OoxElement:
    """A complete, minimal ``a:theme`` (colour, font and format schemes).

    Every presentation needs one; Word and Excel use one if present.
    Colours are Office's 2013+ defaults.
    """
    def srgb(v: str) -> OoxElement:
        return a.SrgbClr(val=v)

    def ph_fill() -> OoxElement:
        return a.SolidFill(a.SchemeClr(val="phClr"))

    clr = a.ClrScheme(
        a.Dk1(a.SysClr(val="windowText", lastClr="000000")),
        a.Lt1(a.SysClr(val="window", lastClr="FFFFFF")),
        a.Dk2(srgb("44546A")), a.Lt2(srgb("E7E6E6")),
        a.Accent1(srgb("4472C4")), a.Accent2(srgb("ED7D31")), a.Accent3(srgb("A5A5A5")),
        a.Accent4(srgb("FFC000")), a.Accent5(srgb("5B9BD5")), a.Accent6(srgb("70AD47")),
        a.Hlink(srgb("0563C1")), a.FolHlink(srgb("954F72")),
        name="Office",
    )
    fonts = a.FontScheme(
        a.MajorFont(a.Latin(typeface="Calibri Light"), a.Ea(typeface=""), a.Cs(typeface="")),
        a.MinorFont(a.Latin(typeface="Calibri"), a.Ea(typeface=""), a.Cs(typeface="")),
        name="Office",
    )
    fmt = a.FmtScheme(
        a.FillStyleLst(ph_fill(), ph_fill(), ph_fill()),
        a.LnStyleLst(*(a.Ln(ph_fill(), w=w) for w in (6350, 12700, 19050))),
        a.EffectStyleLst(*(a.EffectStyle(a.EffectLst()) for _ in range(3))),
        a.BgFillStyleLst(ph_fill(), ph_fill(), ph_fill()),
        name="Office",
    )
    return a.Theme(a.ThemeElements(clr, fonts, fmt), a.ObjectDefaults(),
                   a.ExtraClrSchemeLst(), name=name)
