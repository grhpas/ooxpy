"""Length units.

**oox's rule: a plain number is points, as in Office's dialogs; anything else
is a ``Length``.**  Every size, position, margin, width and indent oox takes
follows it, and every reader reports lengths in points::

    from oox.units import Cm, Inches, Pt, Emu
    doc.add_picture("logo.png", width=Cm(4))     # or width=113.4 (points)
    section.set_margins(left=Inches(1))          # or left=72
    deck.add_shape(slide, "rect", x=Emu(914400)) # EMU from another library

The helpers return a ``Length``: an ``int`` count of EMU (DrawingML's unit:
914400 per inch, 12700 per point) that remembers it is a length.  Excel column
widths are the exception: they are in characters, as in Excel's dialog.

A plain number too large to be meant as points (a raw EMU or twips value from
older code) raises ``ValueError`` naming ``Emu()``, instead of drawing a
picture a few hundred metres wide.
"""

from __future__ import annotations

EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_MM = 36000
EMU_PER_PT = 12700
EMU_PER_TWIP = 635


class Length(int):
    """A length in EMU.  Behaves as an ``int`` everywhere; arithmetic gives
    plain ints."""

    __slots__ = ()

    @property
    def pt(self) -> float:
        """In points, as the readers report lengths."""
        return round(self / EMU_PER_PT, 6)

    @property
    def cm(self) -> float:
        return self / EMU_PER_CM

    @property
    def inches(self) -> float:
        return self / EMU_PER_INCH

    @property
    def twips(self) -> int:
        return round(self / EMU_PER_TWIP)


def Emu(value: float) -> Length:
    return Length(round(value))


def Inches(value: float) -> Length:
    return Length(round(value * EMU_PER_INCH))


def Cm(value: float) -> Length:
    return Length(round(value * EMU_PER_CM))


def Mm(value: float) -> Length:
    return Length(round(value * EMU_PER_MM))


def Pt(value: float) -> Length:
    return Length(round(value * EMU_PER_PT))


def Px(value: float, dpi: float = 96) -> Length:
    return Length(round(value * EMU_PER_INCH / dpi))


def to_twips(emu: int) -> int:
    """EMU -> twips (for w:pgSz, w:pgMar, w:tblW ...)."""
    return round(emu / EMU_PER_TWIP)


def _check(value: object, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError(f"{name} is a number of points or a Length (Cm(2), Inches(1) ...), "
                        f"not {type(value).__name__}")
    return float(value)


def _limit(points: float, value: object, name: str, limit: float | None) -> None:
    if limit is not None and abs(points) > limit:
        raise ValueError(
            f"{name}={value!r} is read as points ({points / 72:,.0f} inches), more than "
            f"the {limit:g} pt allowed here; for EMU write Emu({value!r}), or use "
            "Cm(), Inches() or Pt()")


def emu(value: float, name: str = "length", limit: float | None = None) -> int:
    """Points (or a ``Length``) -> EMU.  ``limit`` (points) catches raw EMU or
    twips passed by mistake."""
    if isinstance(value, Length):
        return int(value)
    points = _check(value, name)
    _limit(points, value, name, limit)
    return round(points * EMU_PER_PT)


def twips(value: float, name: str = "length", limit: float | None = None) -> int:
    """Points (or a ``Length``) -> twips."""
    if isinstance(value, Length):
        return round(value / EMU_PER_TWIP)
    points = _check(value, name)
    _limit(points, value, name, limit)
    return round(points * 20)


def points(emu_value: float) -> float:
    """EMU -> points, as the readers report them (exact enough to write back)."""
    return round(emu_value / EMU_PER_PT, 6)


def twips_to_points(value: float) -> float:
    """Twips -> points."""
    return round(value / 20, 6)


# Plausible maxima in points, to catch raw EMU or twips in old code.
PAGE_LIMIT = 1584.0       # Word's largest page: 22 inches
DRAWING_LIMIT = 4032.0    # PowerPoint's largest slide, 56 inches; also pictures and shapes
