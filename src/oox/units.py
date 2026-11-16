"""Length units.  DrawingML measures in EMU (English Metric Units):
914400 per inch, 360000 per cm, 12700 per point.  WordprocessingML page
geometry uses twips (1/20 pt).  All helpers return ``int``.

    from oox.units import Inches, Cm, Pt
    doc.add_picture("logo.png", width=Cm(4))
"""

from __future__ import annotations

EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_MM = 36000
EMU_PER_PT = 12700
EMU_PER_TWIP = 635


def Emu(value: float) -> int:
    return round(value)


def Inches(value: float) -> int:
    return round(value * EMU_PER_INCH)


def Cm(value: float) -> int:
    return round(value * EMU_PER_CM)


def Mm(value: float) -> int:
    return round(value * EMU_PER_MM)


def Pt(value: float) -> int:
    return round(value * EMU_PER_PT)


def Px(value: float, dpi: float = 96) -> int:
    return round(value * EMU_PER_INCH / dpi)


def to_twips(emu: int) -> int:
    """EMU -> twips (for w:pgSz, w:pgMar, w:tblW ...)."""
    return round(emu / EMU_PER_TWIP)
