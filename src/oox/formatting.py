"""Character and paragraph formatting, custom styles, and the value types
shared with the formatting resolver (``oox.resolve``).

    p = doc.add_paragraph()
    doc.add_run(p, "Warning: ", bold=True, color="C00000")
    doc.add_run(p, "check the figures.", italic=True, size=12, font="Georgia")
    doc.format_paragraph(p, align="center", space_after=12, first_line=Cm(1),
                         line_spacing=LineSpacing.exactly(14))

    doc.add_style("Note", base="Normal", italic=True, color="595959", space_before=6)
    doc.add_paragraph("A note.", style="Note")

    fmt = doc.formatting()           # reads what Word shows; see oox.resolve
    fmt.paragraph(p).line_spacing    # LineSpacing(rule='exact', value=14.0)
    fmt.run(r).bold

Units are points throughout, as in Word's UI: font size, paragraph spacing,
indents, letter spacing, tab positions, border widths.  A ``Length`` from
``oox.units`` (``Cm(1)``, ``Inches(0.5)``) is accepted wherever points are.
Line spacing is a multiple (``1.15``) or a ``LineSpacing``.  The resolver
reports values in the same forms the writers take.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, NamedTuple

from .element import OoxElement
from .grammar import URI
from .ns import w
from .units import PAGE_LIMIT, twips

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}
_ALIGN_BACK = {"left": "left", "start": "left", "center": "center", "right": "right",
               "end": "right", "both": "justify", "distribute": "justify"}
RUN_OPTIONS = ("bold", "italic", "underline", "strike", "double_strike", "size", "font",
               "color", "highlight", "superscript", "subscript", "caps", "small_caps",
               "shading", "letter_spacing", "hidden", "language", "border")
PARAGRAPH_OPTIONS = ("align", "space_before", "space_after", "line_spacing", "indent_left",
                     "indent_right", "first_line", "keep_with_next", "keep_together",
                     "page_break_before", "widow_control", "contextual_spacing",
                     "outline_level", "shading", "borders", "tabs")
BORDER_SIDES = ("top", "left", "bottom", "right", "between", "bar")
_LINE_RULES = {"auto": "auto", "exact": "exact", "at_least": "atLeast"}


# -- value types -------------------------------------------------------------------------------
class LineSpacing(NamedTuple):
    """Line spacing as Word stores it.  ``rule`` is "auto" (``value`` is a
    multiple of single spacing), "exact" or "at_least" (``value`` in points).

        LineSpacing.multiple(1.15)   LineSpacing.exactly(12)   LineSpacing.at_least(14)
    """

    rule: str
    value: float

    @classmethod
    def multiple(cls, n: float) -> LineSpacing:
        """``n`` times single spacing (1.15, 1.5, 2)."""
        return cls("auto", float(n))

    @classmethod
    def exactly(cls, points: float) -> LineSpacing:
        """Exactly ``points`` per line."""
        return cls("exact", float(points))

    @classmethod
    def at_least(cls, points: float) -> LineSpacing:
        """At least ``points`` per line (more for larger text)."""
        return cls("at_least", float(points))


class Border(NamedTuple):
    """One border line: ``style`` ("single", "double", "dotted", ...), ``width``
    and ``space`` (gap to the text) in points, ``color`` as hex or None (auto)."""

    style: str = "single"
    width: float = 0.5
    color: str | None = None
    space: float = 0.0


class Tab(NamedTuple):
    """A tab stop: ``position`` in points from the left indent's origin,
    ``align`` ("left", "center", "right", "decimal", "bar"), ``leader``
    ("none", "dot", "hyphen", "underscore", "heavy", "middleDot")."""

    position: float
    align: str = "left"
    leader: str = "none"


@dataclass(frozen=True)
class Borders:
    """Paragraph borders; each side is a ``Border`` or None."""

    top: Border | None = None
    left: Border | None = None
    bottom: Border | None = None
    right: Border | None = None
    between: Border | None = None
    bar: Border | None = None

    def __bool__(self) -> bool:
        return any(getattr(self, side) is not None for side in BORDER_SIDES)


# -- helpers -----------------------------------------------------------------------------------
def to_twips(value: float) -> int:
    """Points (or a ``Length``) to twips."""
    return twips(value, limit=PAGE_LIMIT)


def _val(el: OoxElement | None, attr: str = "val") -> str | None:
    return el.get(f"{{{W}}}{attr}") if el is not None else None


def _on(el: OoxElement | None) -> bool | None:
    """ST_OnOff: present without val means on."""
    if el is None:
        return None
    return _val(el) not in ("0", "false", "off")


def _hex(color: str) -> str:
    return color.lstrip("#").upper()


def _border_element(factory, border: Border) -> OoxElement:
    attrs: dict[str, Any] = {"val": border.style, "sz": round(border.width * 8),
                             "space": round(border.space),
                             "color": _hex(border.color) if border.color else "auto"}
    return factory(**attrs)


# -- building ----------------------------------------------------------------------------------
def apply_run_format(rpr: OoxElement, **fmt: Any) -> OoxElement:
    """Set run properties on a ``w:rPr`` (None leaves a property alone)."""
    unknown = set(fmt) - set(RUN_OPTIONS)
    if unknown:
        raise TypeError(f"unknown run format {sorted(unknown)}; known: {', '.join(RUN_OPTIONS)}")
    toggles = {"bold": (w.B, w.BCs), "italic": (w.I, w.ICs), "strike": (w.Strike, None),
               "double_strike": (w.Dstrike, None), "caps": (w.Caps, None),
               "small_caps": (w.SmallCaps, None), "hidden": (w.Vanish, None)}
    for name, (factory, cs) in toggles.items():
        value = fmt.get(name)
        if value is None:
            continue
        rpr.add(factory() if value else factory(val=False), replace=True)
        if cs is not None:
            rpr.add(cs() if value else cs(val=False), replace=True)
    if fmt.get("underline") is not None:
        u = fmt["underline"]
        rpr.add(w.U(val="single" if u is True else ("none" if u is False else u)), replace=True)
    if fmt.get("size") is not None:
        half = round(float(fmt["size"]) * 2)
        rpr.add(w.Sz(val=half), replace=True)
        rpr.add(w.SzCs(val=half), replace=True)
    if fmt.get("font") is not None:
        f = fmt["font"]
        rpr.add(w.RFonts(ascii=f, hAnsi=f, cs=f, eastAsia=f), replace=True)
    if fmt.get("color") is not None:
        rpr.add(w.Color(val=_hex(fmt["color"])), replace=True)
    if fmt.get("highlight") is not None:
        rpr.add(w.Highlight(val=fmt["highlight"]), replace=True)
    if fmt.get("superscript") is not None or fmt.get("subscript") is not None:
        pos = ("superscript" if fmt.get("superscript")
               else "subscript" if fmt.get("subscript") else "baseline")
        rpr.add(w.VertAlign(val=pos), replace=True)
    if fmt.get("shading") is not None:
        rpr.add(w.Shd(val="clear", color="auto", fill=_hex(fmt["shading"])), replace=True)
    if fmt.get("letter_spacing") is not None:
        rpr.add(w.Spacing(val=to_twips(fmt["letter_spacing"])), replace=True)
    if fmt.get("language") is not None:
        rpr.add(w.Lang(val=fmt["language"]), replace=True)
    if fmt.get("border") is not None:
        rpr.add(_border_element(w.Bdr, fmt["border"]), replace=True)
    return rpr


def run(text: str, *, style: str | None = None, **fmt: Any) -> OoxElement:
    """A ``w:r`` with ``text`` (tabs and line breaks converted) and formatting:
    bold, italic, underline (True or a Word style such as "double"), strike,
    double_strike, size (pt), font, color (hex), highlight ("yellow" ...),
    superscript, subscript, caps, small_caps, shading (hex fill),
    letter_spacing (pt), hidden, language ("en-US"), border (``Border``);
    ``style`` is a character style id."""
    from . import teletype

    rpr = w.RPr()
    if style:
        rpr.add(w.RStyle(val=style))
    apply_run_format(rpr, **fmt)
    r = w.R(rpr) if len(rpr) else w.R()
    teletype.addTextToElement(r, text)
    return r


def _line_attrs(ls: Any) -> tuple[int, str]:
    if isinstance(ls, LineSpacing):
        if ls.rule not in _LINE_RULES:
            raise ValueError(f"LineSpacing rule is one of {', '.join(_LINE_RULES)}")
        if ls.rule == "auto":
            return round(ls.value * 240), "auto"
        return to_twips(ls.value), _LINE_RULES[ls.rule]
    if isinstance(ls, (int, float)) and not isinstance(ls, bool):
        return round(float(ls) * 240), "auto"
    raise TypeError("line_spacing is a multiple (1.15) or a LineSpacing "
                    "(LineSpacing.exactly(12), LineSpacing.at_least(12))")


def apply_paragraph_format(ppr: OoxElement, **fmt: Any) -> OoxElement:
    unknown = set(fmt) - set(PARAGRAPH_OPTIONS)
    if unknown:
        raise TypeError(f"unknown paragraph format {sorted(unknown)}; "
                        f"known: {', '.join(PARAGRAPH_OPTIONS)}")
    given = {k for k, v in fmt.items() if v is not None}
    if "align" in given:
        ppr.add(w.Jc(val=_ALIGN[fmt["align"]]), replace=True)
    for name, factory in (("keep_with_next", w.KeepNext), ("keep_together", w.KeepLines),
                          ("page_break_before", w.PageBreakBefore),
                          ("widow_control", w.WidowControl),
                          ("contextual_spacing", w.ContextualSpacing)):
        if name in given:
            ppr.add(factory() if fmt[name] else factory(val=False), replace=True)
    if "outline_level" in given:
        level = int(fmt["outline_level"])
        if not 0 <= level <= 9:
            raise ValueError("outline_level is 0-8 (9 = body text)")
        ppr.add(w.OutlineLvl(val=level), replace=True)
    if "shading" in given:
        ppr.add(w.Shd(val="clear", color="auto", fill=_hex(fmt["shading"])), replace=True)
    if "borders" in given:
        borders = fmt["borders"]
        sides = ({s: getattr(borders, s) for s in BORDER_SIDES} if isinstance(borders, Borders)
                 else dict(borders))
        bad = set(sides) - set(BORDER_SIDES)
        if bad:
            raise ValueError(f"border sides are {', '.join(BORDER_SIDES)}")
        pbdr = w.PBdr()
        for side in BORDER_SIDES:
            if sides.get(side) is not None:
                pbdr.add(_border_element(getattr(w, side.capitalize()), sides[side]))
        ppr.add(pbdr, replace=True)
    if "tabs" in given:
        tabs = w.Tabs()
        for tab in fmt["tabs"]:
            t = tab if isinstance(tab, Tab) else Tab(*tab)
            attrs: dict[str, Any] = {"val": t.align, "pos": to_twips(t.position)}
            if t.leader != "none":
                attrs["leader"] = t.leader
            tabs.add(w.Tab(**attrs))
        if len(tabs):
            ppr.add(tabs, replace=True)
        elif ppr.first("w:tabs") is not None:
            ppr.remove(ppr.first("w:tabs"))
    if {"space_before", "space_after", "line_spacing"} & given:
        spacing = ppr.first("w:spacing")
        sp = {k.rpartition("}")[2]: v for k, v in (spacing.attrib.items()
                                                    if spacing is not None else [])}
        if "space_before" in given:
            sp["before"] = to_twips(fmt["space_before"])
            sp.pop("beforeAutospacing", None)
        if "space_after" in given:
            sp["after"] = to_twips(fmt["space_after"])
            sp.pop("afterAutospacing", None)
        if "line_spacing" in given:
            sp["line"], sp["lineRule"] = _line_attrs(fmt["line_spacing"])
        ppr.add(w.Spacing(**sp), replace=True)
    if {"indent_left", "indent_right", "first_line"} & given:
        ind = ppr.first("w:ind")
        attrs = {k.rpartition("}")[2]: v for k, v in (ind.attrib.items() if ind is not None else [])}
        if "indent_left" in given:
            attrs.pop("start", None)
            attrs["left"] = to_twips(fmt["indent_left"])
        if "indent_right" in given:
            attrs.pop("end", None)
            attrs["right"] = to_twips(fmt["indent_right"])
        if "first_line" in given:
            v = to_twips(fmt["first_line"])
            attrs.pop("firstLine", None)
            attrs.pop("hanging", None)
            if v >= 0:
                attrs["firstLine"] = v
            else:
                attrs["hanging"] = -v
        ppr.add(w.Ind(**attrs), replace=True)
    return ppr


def format_paragraph(paragraph: OoxElement, **fmt: Any) -> OoxElement:
    """Paragraph formatting: align (left/center/right/justify), space_before /
    space_after (pt), line_spacing (a multiple or a ``LineSpacing``),
    indent_left / indent_right / first_line (pt or a ``Length``; negative
    first_line = hanging), keep_with_next, keep_together, page_break_before,
    widow_control, contextual_spacing, outline_level (0-8), shading (hex
    fill), borders (``Borders`` or {"bottom": Border(), ...}), tabs (``Tab``
    stops; replaces the paragraph's own)."""
    ppr = paragraph.first("w:pPr")
    if ppr is None:
        ppr = w.PPr()
        paragraph.insert(0, ppr)
    apply_paragraph_format(ppr, **fmt)
    return paragraph


def style_id_for(name: str) -> str:
    """Word's way of making an id from a name: letters and digits only."""
    sid = re.sub(r"[^0-9A-Za-z]", "", name)
    return sid or "Style"


def add_style(doc: WordprocessingDocument, name: str, *, kind: str = "paragraph",
              base: str | None = None, next_style: str | None = None,
              quick: bool = True, **fmt: Any) -> str:
    """Add a custom style and return its id.  ``fmt`` takes the run and
    paragraph options (paragraph styles may use both; ``shading`` on a
    paragraph style shades the paragraph)."""
    if kind not in ("paragraph", "character"):
        raise ValueError("kind is 'paragraph' or 'character'")
    try:
        doc.style_id(name)
    except KeyError:
        pass
    else:
        raise ValueError(f"a style named {name!r} already exists")
    sid = style_id_for(name)
    styles = doc.ensure_styles()
    existing = {st.get(f"{{{W}}}styleId") for st in styles.iterchildren(f"{{{W}}}style")}
    n, base_sid = 1, sid
    while sid in existing:
        n += 1
        sid = f"{base_sid}{n}"
    # "shading" is both a run and a paragraph option: a paragraph style shades the paragraph
    para_fmt = ({k: v for k, v in fmt.items() if k in PARAGRAPH_OPTIONS}
                if kind == "paragraph" else {})
    run_fmt = {k: v for k, v in fmt.items() if k in RUN_OPTIONS and k not in para_fmt}
    if kind == "character" and set(fmt) & (set(PARAGRAPH_OPTIONS) - set(RUN_OPTIONS)):
        raise TypeError("character styles take run formatting only")
    unknown = set(fmt) - set(run_fmt) - set(para_fmt)
    if unknown:
        raise TypeError(f"unknown format {sorted(unknown)}")
    st = w.Style(w.Name(val=name), type=kind, styleId=sid, customStyle=True)
    base = base if base is not None else ("Normal" if kind == "paragraph" else None)
    if base:
        st.add(w.BasedOn(val=doc.style_id(base)))
    if next_style:
        st.add(w.Next(val=doc.style_id(next_style)))
    if quick:
        st.add(w.QFormat())
    if para_fmt:
        st.add(apply_paragraph_format(w.PPr(), **para_fmt))
    if run_fmt:
        st.add(apply_run_format(w.RPr(), **run_fmt))
    styles.add(st)
    return sid
