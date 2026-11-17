"""What Word shows: the effective formatting of paragraphs, runs and table cells.

    fmt = doc.formatting()
    pf = fmt.paragraph(p)     # ParagraphFormat(style='Heading 1', align='left', ...)
    rf = fmt.run(r)           # RunFormat(bold=True, size=16.0, font='Calibri Light', ...)
    fmt.mark(p)               # the paragraph mark's RunFormat
    fmt.cell(tc)              # TableCellFormat(fill='DDDDDD', borders=..., margins=...)
    rf.font_for("α")          # the font Word uses for that character
    fmt.explain(p)            # {"align": "paragraph style 'Heading 1'", ...}

Values are in points and in the forms the writers take (``LineSpacing``,
``Border``, ``Tab``), so a value read here can be written back as it is.

Word builds formatting in layers, each overriding the one before (Word's
"style hierarchy", ECMA-376 Part 1 §17.7.2):

1. document defaults (``w:docDefaults``);
2. the table style, for paragraphs in a table: its whole-table formatting,
   then the conditional formats that apply to the cell (banded columns,
   banded rows, first/last row, first/last column, corner cells, in that
   order) as switched on by the table's ``w:tblLook``;
3. the paragraph style and the styles it is based on;
4. numbering: the list level's paragraph properties (indents, tabs).  When
   the numbering comes from the paragraph style it is applied *before* the
   style, so the style's own indents win (§17.9.x, and what Word does);
5. the character style, for runs;
6. direct formatting.

``explain()`` says which of these layers each value came from.

Toggle properties (bold, italic, caps, small caps, strike, double strike,
hidden, outline, shadow, emboss, imprint, and the complex-script bold and
italic): within a style's ``basedOn`` chain the nearest value wins; a table or
character style that turns one on flips the value below it (bold in a
character style applied to bold text gives regular text, as in Word); a
paragraph style sets the value outright (MS-OI29500 2.1.258: Word resets
rather than toggles there); direct formatting sets it outright.

Word's compatibility rule for tables (MS-DOCX ``overrideTableStyleFontSize
AndJustification``, off unless the setting says otherwise): when the
paragraph style's font size comes from the default paragraph style and is 11
or 12 pt, or its alignment comes from there and is left, the table style's
value is kept.

Where Word and the standard differ, the values are Word's: a paragraph
without ``w:widowControl`` anywhere in its hierarchy has widow control on.

Table cells (``cell()``): borders come from the cell's own ``w:tcBorders``,
then the table style's cell borders (whole table, then each conditional
format, mapped to the cell's edges: a region's outer edge or its inside
lines), then the row's ``w:tblPrEx`` and the table's ``w:tblBorders``.
Between two cells, with no cell spacing, the border Word draws is chosen as
MS-OI29500 2.1.169 says: ``nil`` wins, then ``none`` loses, then the heavier
border (width x style number; dotted and dashed weigh 1), then the style
earlier in Word's list, then the darker colour, then the cell above or to
the left.  Margins, shading and vertical alignment go through the same
layers.

Not modelled: character-unit indents and ``beforeLines``/``afterLines``
spacing, East Asian layout, frames, conditional formatting recorded only in
``w:cnfStyle``, right-to-left tables (``w:bidiVisual``: left and right are
not swapped), diagonal cell borders, the merged look of vertically merged
cells (each ``w:tc`` is reported as stored), and fonts that the theme gives
per script (``a:font script=...``).

A resolver indexes the styles when it is made and does not see later edits
to them; make another one after changing styles.  Reading never changes the
document.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .formatting import _ALIGN_BACK, BORDER_SIDES, HIGHLIGHT_COLORS, Border, Borders, LineSpacing, Tab
from .grammar import URI
from .records import Record

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
A = URI["a"]
_W = f"{{{W}}}"
_RUN_TOGGLES = {"bold": "b", "italic": "i", "caps": "caps", "small_caps": "smallCaps",
                "strike": "strike", "double_strike": "dstrike", "hidden": "vanish",
                "outline": "outline", "shadow": "shadow", "emboss": "emboss",
                "imprint": "imprint", "bold_complex": "bCs", "italic_complex": "iCs"}
TOGGLES = frozenset(_RUN_TOGGLES)
_RUN_FLAGS = {"right_to_left": "rtl", "complex_script": "cs"}  # not toggles
_PARA_FLAGS = {"keep_with_next": "keepNext", "keep_together": "keepLines",
               "page_break_before": "pageBreakBefore", "widow_control": "widowControl",
               "contextual_spacing": "contextualSpacing", "right_to_left": "bidi",
               "suppress_hyphens": "suppressAutoHyphens"}
_FONT_SLOTS = (("font", "ascii", "asciiTheme"), ("font_high_ansi", "hAnsi", "hAnsiTheme"),
               ("font_east_asia", "eastAsia", "eastAsiaTheme"),
               ("font_complex", "cs", "cstheme"))
_RULES_BACK = {"auto": "auto", "exact": "exact", "atLeast": "at_least"}
_WRAPPERS = frozenset({f"{_W}sdt", f"{_W}sdtContent", f"{_W}customXml"})
# The order Word applies conditional table formatting in (§17.7.6); later wins.
_CONDITIONS = ("wholeTable", "band1Vert", "band2Vert", "band1Horz", "band2Horz",
               "firstRow", "lastRow", "firstCol", "lastCol",
               "nwCell", "neCell", "swCell", "seCell")
_CONDITION_NAMES = {"wholeTable": "whole table", "band1Vert": "odd column band",
                    "band2Vert": "even column band", "band1Horz": "odd row band",
                    "band2Horz": "even row band", "firstRow": "first row",
                    "lastRow": "last row", "firstCol": "first column",
                    "lastCol": "last column", "nwCell": "top left cell",
                    "neCell": "top right cell", "swCell": "bottom left cell",
                    "seCell": "bottom right cell"}
_AUTO_SPACING = 14.0  # what Word uses for "Auto" spacing before/after (HTML paragraphs)
_DEFAULT_LOOK = {"firstRow": True, "lastRow": False, "firstColumn": True,
                 "lastColumn": False, "noHBand": False, "noVBand": True}  # Word's 0x04A0
_SINGLE = LineSpacing("auto", 1.0)
_LOOK_BITS = {"firstRow": 0x20, "lastRow": 0x40, "firstColumn": 0x80, "lastColumn": 0x100,
              "noHBand": 0x200, "noVBand": 0x400}
# Word's cell margins when nothing sets them (its "Table Normal" style: 0.075")
_DEFAULT_CELL_MARGINS = {"top": 0.0, "left": 5.4, "bottom": 0.0, "right": 5.4}
DEFAULT = "Word's default"
DIRECT = "direct formatting"


# -- the records ---------------------------------------------------------------------------------
@dataclass(frozen=True)
class RunFormat(Record):
    """Effective run (character) formatting.

    Sizes and ``letter_spacing``, ``position`` (raised > 0, lowered < 0) and
    ``kerning`` (the smallest size kerned; None: off) in points; ``scale``
    in percent; colors as hex ("C00000") or None for automatic;
    ``highlight`` is Word's colour name ("yellow"; ``highlight_color`` gives
    its hex); ``underline`` is None or Word's underline style ("single",
    "double" ...); ``emphasis`` is None or "dot", "comma", "circle",
    "underDot".  ``style`` is the character style's name.

    Fonts: Word picks one of four per character.  ``font`` is for ASCII,
    ``font_high_ansi`` for the rest of Latin, Greek, Cyrillic and symbols,
    ``font_east_asia`` and ``font_complex`` (right-to-left and South-East
    Asian scripts) for theirs; ``font_hint`` ("eastAsia", "cs" or None)
    settles characters that two of them share.  ``font_for(char)`` does the
    choosing.  ``bold_complex``, ``italic_complex`` and ``size_complex``
    apply to complex-script text, ``right_to_left`` and ``complex_script``
    make a whole run use them."""

    bold: bool = False
    italic: bool = False
    underline: str | None = None
    strike: bool = False
    double_strike: bool = False
    size: float = 10.0
    font: str | None = "Times New Roman"
    font_east_asia: str | None = None
    font_complex: str | None = None
    color: str | None = None
    highlight: str | None = None
    superscript: bool = False
    subscript: bool = False
    caps: bool = False
    small_caps: bool = False
    shading: str | None = None
    letter_spacing: float = 0.0
    hidden: bool = False
    language: str | None = None
    border: Border | None = None
    style: str | None = None
    font_high_ansi: str | None = "Times New Roman"
    font_hint: str | None = None
    position: float = 0.0
    scale: int = 100
    kerning: float | None = None
    emphasis: str | None = None
    outline: bool = False
    shadow: bool = False
    emboss: bool = False
    imprint: bool = False
    right_to_left: bool = False
    complex_script: bool = False
    bold_complex: bool = False
    italic_complex: bool = False
    size_complex: float = 10.0

    @property
    def highlight_color(self) -> str | None:
        """The highlight as hex ("FFFF00" for "yellow"), or None."""
        return HIGHLIGHT_COLORS.get(self.highlight) if self.highlight else None

    def font_for(self, char: str) -> str | None:
        """The font Word uses for ``char`` in this run (§17.3.2.26, simplified:
        whole Unicode blocks, not single characters)."""
        return getattr(self, script_slot(char, self.font_hint, self.right_to_left
                                         or self.complex_script))


@dataclass(frozen=True)
class ParagraphFormat(Record):
    """Effective paragraph formatting, in points.  ``first_line`` is negative
    for a hanging indent; ``outline_level`` is 0-8 (Word's dialogs say 1-9)
    or None for body text; ``style`` is the paragraph style's name;
    ``numbering`` is ``(num_id, level)`` for a list paragraph (level 0-8),
    with the ints ``ListItem`` uses.  ``widow_control`` is on unless
    something turns it off, as in Word."""

    style: str | None = None
    align: str = "left"
    space_before: float = 0.0
    space_after: float = 0.0
    line_spacing: LineSpacing = _SINGLE
    indent_left: float = 0.0
    indent_right: float = 0.0
    first_line: float = 0.0
    keep_with_next: bool = False
    keep_together: bool = False
    page_break_before: bool = False
    widow_control: bool = True
    contextual_spacing: bool = False
    outline_level: int | None = None
    shading: str | None = None
    borders: Borders = field(default_factory=Borders)
    tabs: tuple[Tab, ...] = ()
    numbering: tuple[int, int] | None = None
    right_to_left: bool = False
    suppress_hyphens: bool = False


@dataclass(frozen=True)
class TableCellFormat(Record):
    """Effective formatting of a table cell (``w:tc``), after table styles,
    conditional formats, row exceptions and border conflicts.

    ``fill`` is the shading (hex) or None; ``valign`` "top", "center" or
    "bottom" (``Table.format_cell``'s names); ``borders`` the border Word
    draws on each edge (``top``, ``left``, ``bottom``, ``right``; None for
    no line); ``margins`` the cell margins in points by side.
    ``conditions`` names the table style's conditional formats that apply
    ("firstRow", "band1Horz" ...); ``style`` is the table style's name."""

    fill: str | None = None
    valign: str = "top"
    borders: Borders = field(default_factory=Borders)
    margins: dict[str, float] = field(default_factory=lambda: dict(_DEFAULT_CELL_MARGINS))
    conditions: tuple[str, ...] = ()
    style: str | None = None


# -- which font a character uses -------------------------------------------------------------------
_COMPLEX = ((0x0590, 0x08FF), (0x0900, 0x0DFF), (0x0E00, 0x0FFF), (0x1000, 0x109F),
            (0x1780, 0x17FF), (0xFB1D, 0xFDFF), (0xFE70, 0xFEFF))
_EAST_ASIAN = ((0x1100, 0x11FF), (0x2E80, 0x2FFF), (0x3000, 0x33FF), (0x3400, 0x4DBF),
               (0x4E00, 0x9FFF), (0xA000, 0xA4CF), (0xA960, 0xA97F), (0xAC00, 0xD7FF),
               (0xF900, 0xFAFF), (0xFE30, 0xFE4F), (0xFF00, 0xFFEF), (0x20000, 0x3FFFF))
# Shared between East Asian and Western text: the hint "eastAsia" moves them to
# the East Asian font (Latin-1 and Latin Extended, IPA, Greek, Cyrillic,
# punctuation, letterlike symbols, arrows, maths, box drawing, shapes, dingbats,
# private use).
_SHARED = ((0x00A1, 0x04FF), (0x2000, 0x27BF), (0xE000, 0xF8FF))


def _within(code: int, ranges: tuple[tuple[int, int], ...]) -> bool:
    return any(lo <= code <= hi for lo, hi in ranges)


def script_slot(char: str, hint: str | None = None, complex_run: bool = False) -> str:
    """Which ``RunFormat`` font field Word uses for ``char``: "font",
    "font_high_ansi", "font_east_asia" or "font_complex"."""
    if not char:
        return "font"
    code = ord(char[0])
    if complex_run or _within(code, _COMPLEX):
        return "font_complex"
    if _within(code, _EAST_ASIAN):
        return "font_east_asia"
    if code < 0x80:
        return "font"
    if hint == "eastAsia" and _within(code, _SHARED):
        return "font_east_asia"
    return "font_high_ansi"


# -- reading one w:rPr / w:pPr into a layer ---------------------------------------------------
def _attr(el: OoxElement | None, name: str) -> str | None:
    return el.get(_W + name) if el is not None else None


def _on(el: OoxElement | None) -> bool | None:
    if el is None:
        return None
    return el.get(_W + "val") not in ("0", "false", "off")


def _number(value: str | None, default: float = 0) -> float:
    try:
        return float(value) if value is not None else default
    except ValueError:
        return default


def _color(value: str | None) -> str | None:
    return None if value in (None, "", "auto") else value.upper()


def _shading(shd: OoxElement) -> str | None:
    if _attr(shd, "val") == "solid":
        return _color(_attr(shd, "color"))
    return _color(_attr(shd, "fill"))


def _border(el: OoxElement) -> Border | None:
    style = _attr(el, "val") or "none"
    if style in ("nil", "none"):
        return None
    return Border(style, int(_number(_attr(el, "sz"), 2)) / 8, _color(_attr(el, "color")),
                  _number(_attr(el, "space")))


def _twips(value: str | None) -> float:
    return int(_number(value)) / 20


class _Theme:
    def __init__(self, doc: WordprocessingDocument) -> None:
        self.fonts: dict[str, str | None] = {}
        found = doc.related_parts(doc.main_part, "ThemePart")
        root = found[0].element if found else None
        for kind in ("major", "minor"):
            font = root.find(f".//{{{A}}}{kind}Font") if root is not None else None
            for slot, tag in (("Ascii", "latin"), ("HAnsi", "latin"), ("EastAsia", "ea"),
                              ("Bidi", "cs")):
                el = font.find(f"{{{A}}}{tag}") if font is not None else None
                self.fonts[kind + slot] = (el.get("typeface") or None) if el is not None else None

    def font(self, fonts: OoxElement, slot: str, theme_slot: str) -> tuple[bool, str | None]:
        theme = _attr(fonts, theme_slot)
        if theme is not None:
            return True, self.fonts.get(theme)
        value = _attr(fonts, slot)
        return value is not None, value


def _run_layer(rpr: OoxElement | None, theme: _Theme) -> dict[str, Any]:
    out: dict[str, Any] = {}
    if rpr is None:
        return out
    for key, tag in (*_RUN_TOGGLES.items(), *_RUN_FLAGS.items()):
        v = _on(rpr.first(f"w:{tag}"))
        if v is not None:
            out[key] = v
    u = rpr.first("w:u")
    if u is not None:
        out["underline"] = None if _attr(u, "val") in (None, "none") else _attr(u, "val")
    for key, tag in (("size", "sz"), ("size_complex", "szCs")):
        sz = _attr(rpr.first(f"w:{tag}"), "val")
        if sz is not None:
            out[key] = int(_number(sz, 20)) / 2
    fonts = rpr.first("w:rFonts")
    if fonts is not None:
        for key, slot, theme_slot in _FONT_SLOTS:
            present, value = theme.font(fonts, slot, theme_slot)
            if present:
                out[key] = value
        hint = _attr(fonts, "hint")
        if hint is not None:
            out["font_hint"] = None if hint == "default" else hint
    color = rpr.first("w:color")
    if color is not None:
        out["color"] = _color(_attr(color, "val"))
    hl = _attr(rpr.first("w:highlight"), "val")
    if hl is not None:
        out["highlight"] = None if hl == "none" else hl
    va = _attr(rpr.first("w:vertAlign"), "val")
    if va is not None:
        out["superscript"], out["subscript"] = va == "superscript", va == "subscript"
    shd = rpr.first("w:shd")
    if shd is not None:
        out["shading"] = _shading(shd)
    spacing = _attr(rpr.first("w:spacing"), "val")
    if spacing is not None:
        out["letter_spacing"] = _twips(spacing)
    position = _attr(rpr.first("w:position"), "val")
    if position is not None:
        out["position"] = int(_number(position)) / 2
    scale = _attr(rpr.first("w:w"), "val")
    if scale is not None:
        out["scale"] = int(_number(scale.rstrip("%"), 100))
    kern = _attr(rpr.first("w:kern"), "val")
    if kern is not None:
        out["kerning"] = int(_number(kern)) / 2 or None
    em = _attr(rpr.first("w:em"), "val")
    if em is not None:
        out["emphasis"] = None if em == "none" else em
    lang = _attr(rpr.first("w:lang"), "val")
    if lang is not None:
        out["language"] = lang
    bdr = rpr.first("w:bdr")
    if bdr is not None:
        out["border"] = _border(bdr)
    return out


def _paragraph_layer(ppr: OoxElement | None) -> dict[str, Any]:
    out: dict[str, Any] = {}
    if ppr is None:
        return out
    jc = _attr(ppr.first("w:jc"), "val")
    if jc is not None:
        out["align"] = _ALIGN_BACK.get(jc, jc)
    for key, tag in _PARA_FLAGS.items():
        v = _on(ppr.first(f"w:{tag}"))
        if v is not None:
            out[key] = v
    sp = ppr.first("w:spacing")
    if sp is not None:
        for key in ("before", "after"):
            if _on_attr(sp, f"{key}Autospacing"):
                out[f"space_{key}"] = _AUTO_SPACING
            elif _attr(sp, key) is not None:
                out[f"space_{key}"] = _twips(_attr(sp, key))
        line = _attr(sp, "line")
        if line is not None:
            rule = _RULES_BACK.get(_attr(sp, "lineRule") or "auto", "auto")
            value = int(_number(line, 240))
            out["line_spacing"] = LineSpacing(rule, value / 240 if rule == "auto" else value / 20)
    ind = ppr.first("w:ind")
    if ind is not None:
        for key, names in (("indent_left", ("left", "start")), ("indent_right", ("right", "end"))):
            value = next((_attr(ind, n) for n in names if _attr(ind, n) is not None), None)
            if value is not None:
                out[key] = _twips(value)
        if _attr(ind, "hanging") is not None:
            out["first_line"] = -_twips(_attr(ind, "hanging"))
        elif _attr(ind, "firstLine") is not None:
            out["first_line"] = _twips(_attr(ind, "firstLine"))
    level = _attr(ppr.first("w:outlineLvl"), "val")
    if level is not None:
        n = int(_number(level, 9))
        out["outline_level"] = n if 0 <= n <= 8 else None
    shd = ppr.first("w:shd")
    if shd is not None:
        out["shading"] = _shading(shd)
    pbdr = ppr.first("w:pBdr")
    if pbdr is not None:
        for side in BORDER_SIDES:
            el = pbdr.first(f"w:{side}")
            if el is not None:
                out[f"border_{side}"] = _border(el)
    tabs = ppr.first("w:tabs")
    if tabs is not None:
        stops: dict[int, Tab | None] = {}
        for t in tabs.iterchildren(_W + "tab"):
            pos = int(_number(_attr(t, "pos")))
            align = _attr(t, "val") or "left"
            stops[pos] = None if align == "clear" else Tab(
                pos / 20, {"start": "left", "end": "right"}.get(align, align),
                _attr(t, "leader") or "none")
        out["tabs"] = stops
    return out


# -- table cells: borders, margins, shading ----------------------------------------------------
_EDGE_NAMES = {"top": ("top",), "left": ("left", "start"), "bottom": ("bottom",),
               "right": ("right", "end"), "insideH": ("insideH",), "insideV": ("insideV",)}


def _edge(el: OoxElement) -> Border | str:
    """A border for conflict resolution: a ``Border``, or "nil" / "none"."""
    style = _attr(el, "val") or "none"
    if style in ("nil", "none"):
        return style
    border = _border(el)
    assert border is not None
    return border


def _cell_layer(tcpr: OoxElement | None, prefix: str = "") -> dict[str, Any]:
    """``w:tcPr`` (or with ``prefix`` "tbl_", a ``w:tblPr``/``w:tblPrEx``) as a layer:
    edge_<side> (insideH/insideV too), margin_<side>, fill, valign."""
    out: dict[str, Any] = {}
    if tcpr is None:
        return out
    borders = tcpr.first("w:tblBorders" if prefix else "w:tcBorders")
    if borders is not None:
        for side, names in _EDGE_NAMES.items():
            el = next((e for e in (borders.first(f"w:{n}") for n in names) if e is not None), None)
            if el is not None:
                out[f"{prefix}edge_{side}"] = _edge(el)
    margins = tcpr.first("w:tblCellMar" if prefix else "w:tcMar")
    if margins is not None:
        for side, names in _EDGE_NAMES.items():
            if side.startswith("inside"):
                continue
            el = next((e for e in (margins.first(f"w:{n}") for n in names) if e is not None), None)
            if el is not None and _attr(el, "type") in (None, "dxa", "nil"):
                out[f"{prefix}margin_{side}"] = (0.0 if _attr(el, "type") == "nil"
                                                 else _twips(_attr(el, "w")))
    shd = tcpr.first("w:shd")
    if shd is not None:
        out[f"{prefix}fill"] = _shading(shd)
    if not prefix:
        va = _attr(tcpr.first("w:vAlign"), "val")
        if va is not None:
            out["valign"] = {"both": "center"}.get(va, va)
    spacing = tcpr.first("w:tblCellSpacing") if prefix else None
    if spacing is not None:
        out["tbl_spacing"] = _twips(_attr(spacing, "w")) if _attr(spacing, "type") in (
            None, "dxa") else 0.0
    return out


# Word's border numbers and precedence (MS-OI29500 2.1.169)
_BORDER_NUMBER = {"single": 1, "thick": 2, "double": 3, "dotted": 1, "dashed": 1,
                  "dotDash": 8, "dotDotDash": 9, "triple": 10, "thinThickSmallGap": 11,
                  "thickThinSmallGap": 12, "thinThickThinSmallGap": 13,
                  "thinThickMediumGap": 14, "thickThinMediumGap": 15,
                  "thinThickThinMediumGap": 16, "thinThickLargeGap": 17,
                  "thickThinLargeGap": 18, "thinThickThinLargeGap": 19, "wave": 20,
                  "doubleWave": 21, "dashSmallGap": 22, "dashDotStroked": 23,
                  "threeDEmboss": 24, "threeDEngrave": 25, "outset": 26, "inset": 27}
_PRECEDENCE = ("single", "thick", "double", "dotted", "dashed", "dotDash", "dotDotDash",
               "triple", "thinThickSmallGap", "thickThinSmallGap", "thinThickThinSmallGap",
               "thinThickMediumGap", "thickThinMediumGap", "thinThickThinMediumGap",
               "thinThickLargeGap", "thickThinLargeGap", "thinThickThinLargeGap", "wave",
               "doubleWave", "dashSmallGap", "dashDotStroked", "threeDEmboss",
               "threeDEngrave", "outset", "inset")


def _weight(b: Border) -> float:
    if b.style in ("dotted", "dashed"):
        return 1
    return round(b.width * 8) * _BORDER_NUMBER.get(b.style, 1)


def _brightness(color: str | None) -> tuple[int, int, int]:
    try:
        r, g, b = (int((color or "000000")[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        r = g = b = 0
    return (r + b + 2 * g, b + 2 * g, g)


def resolve_conflict(first: Border | str | None, second: Border | str | None) -> Border | None:
    """The border Word draws where two cells meet with no cell spacing:
    ``first`` is the upper or left cell's, ``second`` the other's ("nil",
    "none" or None for no border)."""
    if first == "nil" or second == "nil":
        return None
    if not isinstance(first, Border):
        return second if isinstance(second, Border) else None
    if not isinstance(second, Border):
        return first

    def rank(b: Border) -> tuple[float, int, tuple[int, int, int]]:
        order = _PRECEDENCE.index(b.style) if b.style in _PRECEDENCE else len(_PRECEDENCE)
        return (-_weight(b), order, _brightness(b.color))

    return second if rank(second) < rank(first) else first


def _on_attr(el: OoxElement, name: str) -> bool:
    return _attr(el, name) in ("1", "true", "on")


def _merge(acc: dict[str, Any], layer: dict[str, Any], toggle: bool = False,
           sources: dict[str, Any] | None = None, label: Any = None) -> None:
    """Lay ``layer`` over ``acc``.  ``toggle``: a toggle property that is on
    flips the value below, one that is off leaves it (§17.7.3).  ``sources``
    records where each value came from: ``label`` is a string, or a dict
    (property -> string)."""
    for key, value in layer.items():
        where = label.get(key) if isinstance(label, dict) else label
        if key == "tabs":  # tab stops add up; a cleared stop (None) removes one below
            merged = dict(acc.get("tabs", {}))
            merged.update(value)
            acc["tabs"] = merged
            if sources is not None:
                sources.setdefault("tabs", []).append(where)
            continue
        if toggle and key in TOGGLES:
            if not value:
                continue
            acc[key] = not acc.get(key, False)
        else:
            acc[key] = value
        if sources is not None:
            sources[key] = where


@dataclass
class _StyleLayers:
    ppr: dict[str, Any]
    rpr: dict[str, Any]
    sources: dict[str, str]  # property -> id of the style in the chain that set it
    cell: dict[str, Any] = field(default_factory=dict)       # tcPr, and tblPr as tbl_*
    cell_sources: dict[str, str] = field(default_factory=dict)


@dataclass
class _Place:
    """Where a cell sits: its table, row, index and the conditions that apply."""

    tbl: OoxElement
    tr: OoxElement
    row: int
    index: int
    n_rows: int
    cells: list[OoxElement]
    starts: list[int]
    style: str | None
    conditions: dict[str, tuple[int, int, int, int]]  # name -> region rows/columns


class Formatting:
    """Resolves effective formatting for one document; see the module docs.
    Make one with ``doc.formatting()``."""

    def __init__(self, doc: WordprocessingDocument) -> None:
        from .numbering import Definitions

        self._doc = doc
        self._theme = _Theme(doc)
        styles = doc.styles
        self._styles: dict[str, OoxElement] = {
            st.get(_W + "styleId"): st
            for st in (styles.iterchildren(_W + "style") if styles is not None else [])}
        self._default: dict[str, str | None] = {}
        for sid, st in self._styles.items():
            kind = st.get(_W + "type") or "paragraph"
            if st.get(_W + "default") in ("1", "true", "on"):
                self._default.setdefault(kind, sid)
        defaults = styles.first("w:docDefaults") if styles is not None else None
        rd = defaults.first("w:rPrDefault") if defaults is not None else None
        pd = defaults.first("w:pPrDefault") if defaults is not None else None
        self._default_run = _run_layer(rd.first("w:rPr") if rd is not None else None, self._theme)
        self._default_para = _paragraph_layer(pd.first("w:pPr") if pd is not None else None)
        self._numbering = Definitions(doc)
        self._override_table = self._compat("overrideTableStyleFontSizeAndJustification")
        self._chains: dict[str, _StyleLayers] = {}
        self._table_styles: dict[str, tuple[dict[str, _StyleLayers], int, int]] = {}
        self._tables: dict[OoxElement, tuple[dict[OoxElement, int], int, dict[str, bool],
                                             str | None]] = {}
        self._rows: dict[OoxElement, tuple[list[OoxElement], list[int]]] = {}
        self._own_edges: dict[OoxElement, dict[str, tuple[Border | str | None, str]]] = {}

    def __repr__(self) -> str:
        return f"<Formatting {len(self._styles)} styles>"

    def _compat(self, name: str) -> bool:
        settings = self._doc.settings
        compat = settings.first("w:compat") if settings is not None else None
        if compat is None:
            return False
        for cs in compat.iterchildren(_W + "compatSetting"):
            if cs.get(_W + "name") == name:
                return cs.get(_W + "val") in ("1", "true", "on")
        return False

    # -- styles ----------------------------------------------------------------------------
    def _style_name(self, sid: str | None) -> str | None:
        st = self._styles.get(sid or "")
        name = st.first("w:name") if st is not None else None
        return name.get(_W + "val") if name is not None else None

    def _style_chain(self, sid: str | None) -> list[OoxElement]:
        chain: list[OoxElement] = []
        while sid and sid in self._styles and len(chain) < 20:
            st = self._styles[sid]
            if any(st is c for c in chain):
                break
            chain.append(st)
            based = st.first("w:basedOn")
            sid = based.get(_W + "val") if based is not None else None
        return chain

    def _chain(self, sid: str | None) -> _StyleLayers:
        """A style's paragraph and run properties with its basedOn chain
        applied (nearest wins, toggles included: §17.7.3 c)."""
        key = sid or ""
        if key in self._chains:
            return self._chains[key]
        layers = _StyleLayers({}, {}, {})
        for st in reversed(self._style_chain(sid)):
            ppr = _paragraph_layer(st.first("w:pPr"))
            rpr = _run_layer(st.first("w:rPr"), self._theme)
            _merge(layers.ppr, ppr)
            _merge(layers.rpr, rpr)
            for k in (*ppr, *rpr):
                layers.sources[k] = st.get(_W + "styleId") or ""
        self._chains[key] = layers
        return layers

    def _table_style(self, sid: str | None) -> tuple[dict[str, _StyleLayers], int, int]:
        """Per condition ("" = the style itself), the merged properties along
        the basedOn chain; plus the row and column band sizes."""
        key = sid or ""
        if key in self._table_styles:
            return self._table_styles[key]
        out: dict[str, _StyleLayers] = {}
        row_band = col_band = 1
        for st in reversed(self._style_chain(sid)):
            st_id = st.get(_W + "styleId") or ""
            base = out.setdefault("", _StyleLayers({}, {}, {}))
            tblpr = st.first("w:tblPr")
            parts = ((base, st.first("w:pPr"), st.first("w:rPr"),
                      [st.first("w:tcPr"), tblpr]),)
            if tblpr is not None:
                rb = _attr(tblpr.first("w:tblStyleRowBandSize"), "val")
                cb = _attr(tblpr.first("w:tblStyleColBandSize"), "val")
                row_band = max(int(_number(rb, 1)), 1) if rb is not None else row_band
                col_band = max(int(_number(cb, 1)), 1) if cb is not None else col_band
            for cond in st.iterchildren(_W + "tblStylePr"):
                layers = out.setdefault(cond.get(_W + "type") or "", _StyleLayers({}, {}, {}))
                parts += ((layers, cond.first("w:pPr"), cond.first("w:rPr"),
                           [cond.first("w:tcPr")]),)
            for layers, ppr_el, rpr_el, cell_els in parts:
                ppr, rpr = _paragraph_layer(ppr_el), _run_layer(rpr_el, self._theme)
                _merge(layers.ppr, ppr)
                _merge(layers.rpr, rpr)
                for k in (*ppr, *rpr):
                    layers.sources[k] = st_id
                for i, el in enumerate(cell_els):
                    cell = _cell_layer(el, "tbl_" if i == 1 else "")
                    _merge(layers.cell, cell)
                    for k in cell:
                        layers.cell_sources[k] = st_id
        self._table_styles[key] = (out, row_band, col_band)
        return self._table_styles[key]

    # -- tables ----------------------------------------------------------------------------
    @staticmethod
    def _children(parent: OoxElement, tag: str) -> list[OoxElement]:
        out: list[OoxElement] = []
        stack = list(reversed(list(parent)))
        while stack:
            el = stack.pop()
            if el.tag == tag:
                out.append(el)
            elif el.tag in _WRAPPERS:
                stack.extend(reversed(list(el)))
        return out

    def _table(self, tbl: OoxElement):
        if tbl not in self._tables:
            rows = self._children(tbl, _W + "tr")
            tblpr = tbl.first("w:tblPr")
            look_el = tblpr.first("w:tblLook") if tblpr is not None else None
            if look_el is None:
                look = dict(_DEFAULT_LOOK)
            elif any(_attr(look_el, k) is not None for k in _LOOK_BITS):
                look = {k: _on_attr(look_el, k) for k in _LOOK_BITS}
            else:
                try:
                    bits = int(_attr(look_el, "val") or "0", 16)
                except ValueError:
                    bits = 0x04A0
                look = {k: bool(bits & b) for k, b in _LOOK_BITS.items()}
            style = _attr(tblpr.first("w:tblStyle") if tblpr is not None else None, "val")
            self._tables[tbl] = ({r: i for i, r in enumerate(rows)}, len(rows), look,
                                 style if style is not None else self._default.get("table"))
        return self._tables[tbl]

    def _cells(self, tr: OoxElement) -> tuple[list[OoxElement], list[int]]:
        if tr not in self._rows:
            cells = self._children(tr, _W + "tc")
            trpr = tr.first("w:trPr")
            col = int(_number(_attr(trpr.first("w:gridBefore") if trpr is not None else None,
                                    "val")))
            starts = []
            for tc in cells:
                starts.append(col)
                col += self._span(tc)
            self._rows[tr] = (cells, starts)
        return self._rows[tr]

    @staticmethod
    def _span(tc: OoxElement) -> int:
        tcpr = tc.first("w:tcPr")
        return max(int(_number(_attr(tcpr.first("w:gridSpan") if tcpr is not None else None,
                                     "val"), 1)), 1)

    def _place(self, tc: OoxElement) -> _Place | None:
        tr = next((a for a in tc.iterancestors(_W + "tr")), None)
        tbl = next((a for a in tc.iterancestors(_W + "tbl")), None)
        if tr is None or tbl is None:
            return None
        row_of, n_rows, look, style = self._table(tbl)
        _, row_band, col_band = self._table_style(style)
        cells, starts = self._cells(tr)
        r = row_of.get(tr, 0)
        c = next((i for i, cell in enumerate(cells) if cell is tc), 0)
        last_c = len(cells) - 1
        first_row, last_row = look["firstRow"] and r == 0, look["lastRow"] and r == n_rows - 1
        first_col, last_col = look["firstColumn"] and c == 0, look["lastColumn"] and c == last_c
        everything = (0, n_rows - 1, 0, last_c)
        found: dict[str, tuple[int, int, int, int]] = {"wholeTable": everything}
        if not look["noVBand"] and not (first_col or last_col):
            offset = 1 if look["firstColumn"] else 0
            band = (starts[c] - offset) // col_band
            lo, hi = offset + band * col_band, offset + (band + 1) * col_band - 1
            members = [i for i, s in enumerate(starts) if lo <= s <= hi] or [c]
            found["band1Vert" if band % 2 == 0 else "band2Vert"] = (
                0, n_rows - 1, members[0], members[-1])
        if not look["noHBand"] and not (first_row or last_row):
            offset = 1 if look["firstRow"] else 0
            band = (r - offset) // row_band
            lo, hi = offset + band * row_band, offset + (band + 1) * row_band - 1
            hi = min(hi, n_rows - (2 if look["lastRow"] else 1))
            found["band1Horz" if band % 2 == 0 else "band2Horz"] = (lo, hi, 0, last_c)
        for name, on, region in (
                ("firstRow", first_row, (0, 0, 0, last_c)),
                ("lastRow", last_row, (n_rows - 1, n_rows - 1, 0, last_c)),
                ("firstCol", first_col, (0, n_rows - 1, 0, 0)),
                ("lastCol", last_col, (0, n_rows - 1, last_c, last_c)),
                ("nwCell", first_row and first_col, (r, r, c, c)),
                ("neCell", first_row and last_col, (r, r, c, c)),
                ("swCell", last_row and first_col, (r, r, c, c)),
                ("seCell", last_row and last_col, (r, r, c, c))):
            if on:
                found[name] = region
        return _Place(tbl, tr, r, c, n_rows, cells, starts, style, found)

    def _table_label(self, sid: str, condition: str) -> str:
        name = self._style_name(sid) or sid
        return f"table style '{name}'" + (f", {_CONDITION_NAMES[condition]}"
                                         if condition not in ("", "wholeTable") else "")

    def _table_layers(self, p: OoxElement) -> tuple[_StyleLayers, dict[str, str]] | None:
        """The table style's formatting for the cell holding ``p`` (and where
        each value came from), or None."""
        tc = None
        for anc in p.iterancestors():
            if anc.tag == _W + "tc":
                tc = anc
                break
            if anc.tag in (_W + "txbxContent", _W + "body"):
                return None
        if tc is None:
            return None
        place = self._place(tc)
        if place is None:
            return None
        conditions, _, _ = self._table_style(place.style)
        if not conditions:
            return None
        out = _StyleLayers({}, {}, {})
        labels: dict[str, str] = {}
        for name in ("", *_CONDITIONS):
            if name in conditions and (name == "" or name in place.conditions):
                layer = conditions[name]
                _merge(out.ppr, layer.ppr)
                _merge(out.rpr, layer.rpr, toggle=name != "")
                for k in (*layer.ppr, *layer.rpr):
                    labels[k] = self._table_label(layer.sources[k], name)
        return out, labels

    def _style_layer(self, styled: _StyleLayers, table: _StyleLayers | None,
                     which: str) -> dict[str, Any]:
        """The paragraph style's properties, minus what Word's table rule keeps
        from the table style (see the module docs)."""
        layer = dict(styled.ppr if which == "ppr" else styled.rpr)
        if table is None or self._override_table:
            return layer
        normal = self._default.get("paragraph")
        table_props = table.ppr if which == "ppr" else table.rpr
        for key, keep in (("size", (11.0, 12.0)), ("align", ("left",))):
            if (key in layer and key in table_props and layer[key] in keep
                    and styled.sources.get(key) == normal):
                del layer[key]
        return layer

    def _style_labels(self, styled: _StyleLayers, kind: str) -> dict[str, str]:
        return {k: f"{kind} style '{self._style_name(sid) or sid}'"
                for k, sid in styled.sources.items()}

    # -- public ----------------------------------------------------------------------------
    @staticmethod
    def _paragraph_of(el: OoxElement) -> OoxElement:
        if el.tag == _W + "p":
            return el
        p = next((a for a in el.iterancestors(_W + "p")), None)
        if p is None:
            raise ValueError("the run is not in a paragraph")
        return p

    def _paragraph_style(self, ppr: OoxElement | None) -> str | None:
        ps = ppr.first("w:pStyle") if ppr is not None else None
        sid = ps.get(_W + "val") if ps is not None else None
        return sid if sid in self._styles else self._default.get("paragraph")

    def paragraph(self, p: OoxElement) -> ParagraphFormat:
        """The paragraph formatting Word applies to ``p``."""
        return self._paragraph(p, None)

    def _paragraph(self, p: OoxElement, sources: dict[str, Any] | None) -> ParagraphFormat:
        ppr = p.first("w:pPr")
        sid = self._paragraph_style(ppr)
        styled = self._chain(sid)
        found = self._table_layers(p)
        table = found[0] if found else None
        numbering = self._numbering.numbering_of(ppr)
        level = None
        if numbering is not None:
            defined = self._numbering.levels(numbering[0])
            level = _paragraph_layer(defined[1][numbering[1]].ppr) if defined else None
        list_label = (f"list {numbering[0]}, level {numbering[1]}"
                      if numbering is not None else None)
        style_labels = self._style_labels(styled, "paragraph") if sources is not None else None
        acc: dict[str, Any] = {}
        _merge(acc, self._default_para, sources=sources, label="document defaults")
        if found is not None:
            _merge(acc, found[0].ppr, sources=sources, label=found[1])
        if level is not None and numbering is not None and numbering[2]:
            # numbering from the style: the style's own values win
            _merge(acc, level, sources=sources, label=list_label)
        _merge(acc, self._style_layer(styled, table, "ppr"), sources=sources, label=style_labels)
        if level is not None and numbering is not None and not numbering[2]:
            _merge(acc, level, sources=sources, label=list_label)
        _merge(acc, _paragraph_layer(ppr), sources=sources, label=DIRECT)
        borders = Borders(**{side: acc.pop(f"border_{side}", None) for side in BORDER_SIDES})
        tabs = tuple(sorted((t for t in acc.pop("tabs", {}).values() if t is not None),
                            key=lambda t: t.position))
        return ParagraphFormat(
            style=self._style_name(sid), borders=borders, tabs=tabs,
            numbering=(numbering[0], numbering[1]) if numbering is not None else None,
            **acc)

    def run(self, r: OoxElement) -> RunFormat:
        """The character formatting Word applies to the run ``r``."""
        return self._run_format(self._paragraph_of(r), r.first("w:rPr"), None)

    def mark(self, p: OoxElement) -> RunFormat:
        """The formatting of ``p``'s paragraph mark (what an empty paragraph,
        and by default a list label, looks like)."""
        ppr = p.first("w:pPr")
        return self._run_format(p, ppr.first("w:rPr") if ppr is not None else None, None)

    def _run_format(self, p: OoxElement, rpr: OoxElement | None,
                    sources: dict[str, Any] | None) -> RunFormat:
        ppr = p.first("w:pPr")
        styled = self._chain(self._paragraph_style(ppr))
        found = self._table_layers(p)
        table = found[0] if found else None
        acc: dict[str, Any] = {}
        _merge(acc, self._default_run, sources=sources, label="document defaults")
        if found is not None:
            _merge(acc, found[0].rpr, toggle=True, sources=sources, label=found[1])
        _merge(acc, self._style_layer(styled, table, "rpr"), sources=sources,  # Word resets
               label=self._style_labels(styled, "paragraph") if sources is not None else None)
        rs = rpr.first("w:rStyle") if rpr is not None else None
        csid = rs.get(_W + "val") if rs is not None else None
        if csid in self._styles:
            chain = self._chain(csid)
            _merge(acc, chain.rpr, toggle=True, sources=sources,
                   label=self._style_labels(chain, "character") if sources is not None else None)
        _merge(acc, _run_layer(rpr, self._theme), sources=sources, label=DIRECT)
        return RunFormat(style=self._style_name(csid) if csid in self._styles else None, **acc)

    # -- table cells -------------------------------------------------------------------------
    def cell(self, tc: OoxElement) -> TableCellFormat:
        """The formatting Word gives the table cell ``tc`` (a ``w:tc``):
        shading, vertical alignment, margins, and the border it draws on each
        edge after border conflicts with its neighbours."""
        return self._cell(tc, None)

    def _cell_layers(self, place: _Place, tc: OoxElement):
        """(table-level layer, its labels, cell-level layer, its labels) before conflicts."""
        conditions, _, _ = self._table_style(place.style)
        table: dict[str, Any] = {}
        table_src: dict[str, str] = {}
        base = conditions.get("")
        if base is not None:  # the style's own w:tblPr
            _merge(table, {k: v for k, v in base.cell.items() if k.startswith("tbl_")},
                   sources=table_src,
                   label={k: self._table_label(s, "") for k, s in base.cell_sources.items()})
        tblpr = place.tbl.first("w:tblPr")
        _merge(table, _cell_layer(tblpr, "tbl_"), sources=table_src, label="the table's own")
        _merge(table, _cell_layer(place.tr.first("w:tblPrEx"), "tbl_"), sources=table_src,
               label="the row's exceptions")
        trpr = place.tr.first("w:trPr")
        spacing = trpr.first("w:tblCellSpacing") if trpr is not None else None
        if spacing is not None:
            table["tbl_spacing"] = _twips(_attr(spacing, "w"))
        cell: dict[str, Any] = {}
        cell_src: dict[str, str] = {}
        for name in ("", *_CONDITIONS):
            layer = conditions.get(name)
            if layer is None or (name != "" and name not in place.conditions):
                continue
            region = place.conditions.get(name, place.conditions["wholeTable"])
            mapped = self._map_region({k: v for k, v in layer.cell.items()
                                       if not k.startswith("tbl_")}, place, region)
            _merge(cell, mapped, sources=cell_src,
                   label={k: self._table_label(layer.cell_sources.get(src, ""), name)
                          for k, src in mapped.pop("_from", {}).items()})
        direct = _cell_layer(tc.first("w:tcPr"))
        own = {k: v for k, v in direct.items()
               if not (k.startswith("edge_") and (k.endswith("insideH") or k.endswith("insideV")))}
        # a cell's own "none" lets the style's border through (MS-OI29500 2.1.169)
        own = {k: v for k, v in own.items() if not (k.startswith("edge_") and v == "none")}
        _merge(cell, own, sources=cell_src, label=DIRECT)
        return table, table_src, cell, cell_src

    @staticmethod
    def _map_region(layer: dict[str, Any], place: _Place,
                    region: tuple[int, int, int, int]) -> dict[str, Any]:
        """A conditional format's tcPr for this cell: its outer edges where the
        cell is on the region's edge, its inside lines elsewhere."""
        r0, r1, c0, c1 = region
        at = {"top": place.row == r0, "bottom": place.row == r1,
              "left": place.index == c0, "right": place.index == c1}
        inside = {"top": "insideH", "bottom": "insideH", "left": "insideV", "right": "insideV"}
        out: dict[str, Any] = {}
        origin: dict[str, str] = {}
        for k, v in layer.items():
            if not k.startswith("edge_"):
                out[k] = v
                origin[k] = k
        for side in ("top", "bottom", "left", "right"):
            key = f"edge_{side}" if at[side] else f"edge_{inside[side]}"
            if key in layer:
                out[f"edge_{side}"] = layer[key]
                origin[f"edge_{side}"] = key
        out["_from"] = origin
        return out

    def _edges(self, tc: OoxElement) -> dict[str, tuple[Border | str | None, str]]:
        """Each edge's border before conflicts, with where it came from."""
        if tc in self._own_edges:
            return self._own_edges[tc]
        place = self._place(tc)
        out: dict[str, tuple[Border | str | None, str]] = {}
        if place is not None:
            table, table_src, cell, cell_src = self._cell_layers(place, tc)
            last_c = len(place.cells) - 1
            at = {"top": place.row == 0, "bottom": place.row == place.n_rows - 1,
                  "left": place.index == 0, "right": place.index == last_c}
            inside = {"top": "insideH", "bottom": "insideH", "left": "insideV",
                      "right": "insideV"}
            for side in ("top", "left", "bottom", "right"):
                key = f"edge_{side}"
                if key in cell:
                    out[side] = (cell[key], cell_src.get(key, DIRECT))
                    continue
                tkey = f"tbl_edge_{side}" if at[side] else f"tbl_edge_{inside[side]}"
                out[side] = (table.get(tkey), table_src.get(tkey, DEFAULT))
        self._own_edges[tc] = out
        return out

    def _neighbour(self, place: _Place, side: str) -> OoxElement | None:
        if side == "left":
            return place.cells[place.index - 1] if place.index > 0 else None
        if side == "right":
            return place.cells[place.index + 1] if place.index + 1 < len(place.cells) else None
        rows = self._children(place.tbl, _W + "tr")
        r = place.row + (-1 if side == "top" else 1)
        if not 0 <= r < len(rows):
            return None
        cells, starts = self._cells(rows[r])
        col = place.starts[place.index]
        for cell, start in zip(cells, starts, strict=True):
            if start <= col < start + self._span(cell):
                return cell
        return None

    def _cell(self, tc: OoxElement, sources: dict[str, Any] | None) -> TableCellFormat:
        place = self._place(tc)
        if place is None:
            raise ValueError("the cell is not in a table row")
        table, table_src, cell, cell_src = self._cell_layers(place, tc)
        src = sources if sources is not None else {}
        margins = dict(_DEFAULT_CELL_MARGINS)
        for side in margins:
            for layer, labels in ((table, table_src), (cell, cell_src)):
                key = (f"tbl_margin_{side}" if layer is table else f"margin_{side}")
                if key in layer:
                    margins[side] = layer[key]
                    src[f"margins.{side}"] = labels.get(key, DIRECT)
        fill = None
        for layer, labels, key in ((table, table_src, "tbl_fill"), (cell, cell_src, "fill")):
            if key in layer:
                fill = layer[key]
                src["fill"] = labels.get(key, DIRECT)
        valign = cell.get("valign", "top")
        if "valign" in cell:
            src["valign"] = cell_src.get("valign", DIRECT)
        own = self._edges(tc)
        spaced = bool(table.get("tbl_spacing"))
        opposite = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
        borders: dict[str, Border | None] = {}
        for side in ("top", "left", "bottom", "right"):
            mine, where = own.get(side, (None, DEFAULT))
            other = None if spaced else self._neighbour(place, side)
            if other is None:
                borders[side] = mine if isinstance(mine, Border) else None
                src[f"borders.{side}"] = where
                continue
            theirs, their_where = self._edges(other).get(opposite[side], (None, DEFAULT))
            first, second = (theirs, mine) if side in ("top", "left") else (mine, theirs)
            won = resolve_conflict(first, second)
            borders[side] = won
            from_them = won is not None and won is theirs and won is not mine
            src[f"borders.{side}"] = (
                f"the cell {'above' if side == 'top' else 'below' if side == 'bottom' else side}"
                f" ({their_where})" if from_them else where)
        return TableCellFormat(
            fill=fill, valign=valign, borders=Borders(**borders), margins=margins,
            conditions=tuple(n for n in _CONDITIONS if n in place.conditions
                             and n != "wholeTable"),
            style=self._style_name(place.style))

    # -- where values come from ----------------------------------------------------------------
    def explain(self, el: OoxElement) -> dict[str, str]:
        """Where each value of a paragraph (``w:p``), run (``w:r``) or table
        cell (``w:tc``) comes from: {field: "paragraph style 'Heading 1'",
        "table style 'Grid', first row", "list 3, level 0", "character style
        'Strong'", "document defaults", "direct formatting" ...}.  Fields not
        listed have Word's default.  Borders and margins are listed by side
        ("borders.top"); "tabs" lists every layer that set tab stops."""
        sources: dict[str, Any] = {}
        if el.tag == _W + "p":
            self._paragraph(el, sources)
            for side in BORDER_SIDES:
                if f"border_{side}" in sources:
                    sources[f"borders.{side}"] = sources.pop(f"border_{side}")
        elif el.tag == _W + "r":
            self._run_format(self._paragraph_of(el), el.first("w:rPr"), sources)
        elif el.tag == _W + "tc":
            self._cell(el, sources)
        else:
            raise ValueError("explain() takes a w:p, w:r or w:tc")
        if "tabs" in sources:
            sources["tabs"] = " + ".join(dict.fromkeys(sources["tabs"]))
        return {k: v for k, v in sources.items() if v is not None}
