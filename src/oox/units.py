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


def _plain(other: object) -> bool:
    return isinstance(other, (int, float)) and not isinstance(other, (bool, Length))


def _ambiguous(op: str, other: object) -> TypeError:
    return TypeError(
        f"Length {op} {other!r}: a plain number would be points to oox but EMU to int "
        f"arithmetic; write Pt({other!r}), Cm(...) or Emu({other!r})")


class Length(int):
    """A length in EMU that remembers it is a length.

    Arithmetic keeps it a ``Length``: ``Cm(2) + Cm(1)``, ``-Cm(1)``,
    ``abs()``, ``Cm(2) * 1.5`` and ``Cm(3) / 2`` are all lengths, and
    ``sum(lengths)`` works.  ``Length / Length`` is a plain ratio.  Adding a
    plain number other than 0 raises ``TypeError``, because oox reads a plain
    number as points: write ``Cm(2) + Pt(6)``.  Anywhere else it behaves as an
    ``int`` count of EMU."""

    __slots__ = ()

    def __repr__(self) -> str:
        return f"Emu({int(self)})"

    def __str__(self) -> str:  # the number, as XML attributes need it
        return int.__repr__(self)

    def __format__(self, spec: str) -> str:
        return int(self).__format__(spec)

    def __add__(self, other: object) -> Length:  # type: ignore[override]
        if isinstance(other, Length) or (other == 0 and _plain(other)):
            return Length(int(self) + int(other))  # type: ignore[call-overload]
        if _plain(other):
            raise _ambiguous("+", other)
        return NotImplemented

    __radd__ = __add__  # type: ignore[assignment]

    def __sub__(self, other: object) -> Length:  # type: ignore[override]
        if isinstance(other, Length) or (other == 0 and _plain(other)):
            return Length(int(self) - int(other))  # type: ignore[call-overload]
        if _plain(other):
            raise _ambiguous("-", other)
        return NotImplemented

    def __rsub__(self, other: object) -> Length:  # type: ignore[override]
        if other == 0 and _plain(other):
            return Length(-int(self))
        if _plain(other):
            raise _ambiguous("-", other)
        return NotImplemented

    def __mul__(self, other: object) -> Length:  # type: ignore[override]
        if _plain(other):
            return Length(round(int(self) * other))  # type: ignore[operator]
        if isinstance(other, Length):
            raise TypeError("Length * Length is an area, not a length")
        return NotImplemented

    __rmul__ = __mul__  # type: ignore[assignment]

    def __truediv__(self, other: object) -> Length | float:  # type: ignore[override]
        if isinstance(other, Length):
            return int(self) / int(other)
        if _plain(other):
            return Length(round(int(self) / other))  # type: ignore[operator]
        return NotImplemented

    def __floordiv__(self, other: object) -> Length | int:  # type: ignore[override]
        if isinstance(other, Length):
            return int(self) // int(other)
        if _plain(other):
            return Length(int(self) // other)  # type: ignore[operator]
        return NotImplemented

    def __mod__(self, other: object) -> Length:  # type: ignore[override]
        if isinstance(other, Length):
            return Length(int(self) % int(other))
        return NotImplemented

    def __neg__(self) -> Length:
        return Length(-int(self))

    def __pos__(self) -> Length:
        return self

    def __abs__(self) -> Length:
        return Length(abs(int(self)))

    @property
    def pt(self) -> float:
        """In points, as the readers report lengths."""
        return round(int(self) / EMU_PER_PT, 6)

    @property
    def cm(self) -> float:
        return int(self) / EMU_PER_CM

    @property
    def inches(self) -> float:
        return int(self) / EMU_PER_INCH

    @property
    def twips(self) -> int:
        return round(int(self) / EMU_PER_TWIP)


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
    return round(int(emu) / EMU_PER_TWIP)


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
        return round(int(value) / EMU_PER_TWIP)
    points = _check(value, name)
    _limit(points, value, name, limit)
    return round(points * 20)


def points(emu_value: float) -> float:
    """EMU -> points, as the readers report them (exact enough to write back)."""
    return round(float(emu_value) / EMU_PER_PT, 6)


def twips_to_points(value: float) -> float:
    """Twips -> points."""
    return round(float(value) / 20, 6)


# Plausible maxima in points, to catch raw EMU or twips in old code.
PAGE_LIMIT = 1584.0       # Word's largest page: 22 inches
DRAWING_LIMIT = 4032.0    # PowerPoint's largest slide, 56 inches; also pictures and shapes
