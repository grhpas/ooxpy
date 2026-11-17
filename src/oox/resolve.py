"""What Word shows: the effective formatting of paragraphs and runs.

    fmt = doc.formatting()
    pf = fmt.paragraph(p)     # ParagraphFormat(style='Heading 1', align='left', ...)
    rf = fmt.run(r)           # RunFormat(bold=True, size=16.0, font='Calibri Light', ...)
    fmt.mark(p)               # the paragraph mark's RunFormat

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

Toggle properties (bold, italic, caps, small caps, strike, double strike,
hidden): within a style's ``basedOn`` chain the nearest value wins; a table or
character style that turns one on flips the value below it (bold in a
character style applied to bold text gives regular text, as in Word); a
paragraph style sets the value outright (MS-OI29500 2.1.258: Word resets
rather than toggles there); direct formatting sets it outright.

Word's compatibility rule for tables (MS-DOCX ``overrideTableStyleFontSize
AndJustification``, off unless the setting says otherwise): when the
paragraph style's font size comes from the default paragraph style and is 11
or 12 pt, or its alignment comes from there and is left, the table style's
value is kept.

Not modelled: complex-script variants (``w:bCs``, ``w:szCs``), character-unit
indents and ``beforeLines``/``afterLines`` spacing, East Asian layout,
frames, conditional formatting recorded only in ``w:cnfStyle``.

A resolver indexes the styles when it is made and does not see later edits
to them; make another one after changing styles.  Reading never changes the
document.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .formatting import _ALIGN_BACK, BORDER_SIDES, Border, Borders, LineSpacing, Tab
from .grammar import URI

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
A = URI["a"]
_W = f"{{{W}}}"
_RUN_TOGGLES = {"bold": "b", "italic": "i", "caps": "caps", "small_caps": "smallCaps",
                "strike": "strike", "double_strike": "dstrike", "hidden": "vanish"}
TOGGLES = frozenset(_RUN_TOGGLES)
_PARA_FLAGS = {"keep_with_next": "keepNext", "keep_together": "keepLines",
               "page_break_before": "pageBreakBefore", "widow_control": "widowControl",
               "contextual_spacing": "contextualSpacing"}
_RULES_BACK = {"auto": "auto", "exact": "exact", "atLeast": "at_least"}
_WRAPPERS = frozenset({f"{_W}sdt", f"{_W}sdtContent", f"{_W}customXml"})
# The order Word applies conditional table formatting in (§17.7.6); later wins.
_CONDITIONS = ("wholeTable", "band1Vert", "band2Vert", "band1Horz", "band2Horz",
               "firstRow", "lastRow", "firstCol", "lastCol",
               "nwCell", "neCell", "swCell", "seCell")
_AUTO_SPACING = 14.0  # what Word uses for "Auto" spacing before/after (HTML paragraphs)
_DEFAULT_LOOK = {"firstRow": True, "lastRow": False, "firstColumn": True,
                 "lastColumn": False, "noHBand": False, "noVBand": True}  # Word's 0x04A0
_SINGLE = LineSpacing("auto", 1.0)
_LOOK_BITS = {"firstRow": 0x20, "lastRow": 0x40, "firstColumn": 0x80, "lastColumn": 0x100,
              "noHBand": 0x200, "noVBand": 0x400}


@dataclass(frozen=True)
class RunFormat:
    """Effective run (character) formatting.  ``size``, ``letter_spacing`` in
    points; colors as hex ("C00000") or None for automatic; ``underline`` is
    None or Word's underline style ("single", "double" ...); ``style`` is the
    character style's name."""

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


@dataclass(frozen=True)
class ParagraphFormat:
    """Effective paragraph formatting, in points.  ``first_line`` is negative
    for a hanging indent; ``outline_level`` is 0-8 or None for body text;
    ``style`` is the paragraph style's name; ``numbering`` is
    ``(num_id, level)`` for a list paragraph."""

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
    widow_control: bool = False
    contextual_spacing: bool = False
    outline_level: int | None = None
    shading: str | None = None
    borders: Borders = field(default_factory=Borders)
    tabs: tuple[Tab, ...] = ()
    numbering: tuple[int, int] | None = None


# -- reading one w:rPr / w:pPr into a layer ---------------------------------------------------
def _attr(el: OoxElement | None, name: str) -> str | None:
    return el.get(_W + name) if el is not None else None


def _on(el: OoxElement | None) -> bool | None:
    if el is None:
        return None
    return el.get(_W + "val") not in ("0", "false", "off")


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
    return Border(style, int(_attr(el, "sz") or 2) / 8, _color(_attr(el, "color")),
                  float(_attr(el, "space") or 0))


def _twips(value: str | None) -> float:
    return int(float(value or 0)) / 20


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
    for key, tag in _RUN_TOGGLES.items():
        v = _on(rpr.first(f"w:{tag}"))
        if v is not None:
            out[key] = v
    u = rpr.first("w:u")
    if u is not None:
        out["underline"] = None if _attr(u, "val") in (None, "none") else _attr(u, "val")
    sz = _attr(rpr.first("w:sz"), "val")
    if sz is not None:
        out["size"] = int(float(sz)) / 2
    fonts = rpr.first("w:rFonts")
    if fonts is not None:
        for key, slot, theme_slot in (("font", "ascii", "asciiTheme"),
                                      ("font_east_asia", "eastAsia", "eastAsiaTheme"),
                                      ("font_complex", "cs", "cstheme")):
            present, value = theme.font(fonts, slot, theme_slot)
            if present:
                out[key] = value
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
            value = int(float(line))
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
        out["outline_level"] = int(level) if 0 <= int(level) <= 8 else None
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
            pos = int(float(_attr(t, "pos") or 0))
            align = _attr(t, "val") or "left"
            stops[pos] = None if align == "clear" else Tab(
                pos / 20, {"start": "left", "end": "right"}.get(align, align),
                _attr(t, "leader") or "none")
        out["tabs"] = stops
    return out


def _on_attr(el: OoxElement, name: str) -> bool:
    return _attr(el, name) in ("1", "true", "on")


def _merge(acc: dict[str, Any], layer: dict[str, Any], toggle: bool = False) -> None:
    """Lay ``layer`` over ``acc``.  ``toggle``: a toggle property that is on
    flips the value below, one that is off leaves it (§17.7.3)."""
    for key, value in layer.items():
        if key == "tabs":  # tab stops add up; a cleared stop (None) removes one below
            merged = dict(acc.get("tabs", {}))
            merged.update(value)
            acc["tabs"] = merged
        elif toggle and key in TOGGLES:
            if value:
                acc[key] = not acc.get(key, False)
        else:
            acc[key] = value


@dataclass
class _StyleLayers:
    ppr: dict[str, Any]
    rpr: dict[str, Any]
    sources: dict[str, str]  # property -> id of the style in the chain that set it


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

    def _chain(self, sid: str | None) -> _StyleLayers:
        """A style's paragraph and run properties with its basedOn chain
        applied (nearest wins, toggles included: §17.7.3 c)."""
        key = sid or ""
        if key in self._chains:
            return self._chains[key]
        chain: list[OoxElement] = []
        while sid and sid in self._styles and len(chain) < 20:
            st = self._styles[sid]
            if any(st is c for c in chain):
                break
            chain.append(st)
            based = st.first("w:basedOn")
            sid = based.get(_W + "val") if based is not None else None
        layers = _StyleLayers({}, {}, {})
        for st in reversed(chain):
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
        chain: list[OoxElement] = []
        while sid and sid in self._styles and len(chain) < 20:
            st = self._styles[sid]
            if any(st is c for c in chain):
                break
            chain.append(st)
            based = st.first("w:basedOn")
            sid = based.get(_W + "val") if based is not None else None
        out: dict[str, _StyleLayers] = {}
        row_band = col_band = 1
        for st in reversed(chain):
            base = out.setdefault("", _StyleLayers({}, {}, {}))
            _merge(base.ppr, _paragraph_layer(st.first("w:pPr")))
            _merge(base.rpr, _run_layer(st.first("w:rPr"), self._theme))
            tblpr = st.first("w:tblPr")
            if tblpr is not None:
                rb = _attr(tblpr.first("w:tblStyleRowBandSize"), "val")
                cb = _attr(tblpr.first("w:tblStyleColBandSize"), "val")
                row_band = max(int(rb), 1) if rb is not None else row_band
                col_band = max(int(cb), 1) if cb is not None else col_band
            for cond in st.iterchildren(_W + "tblStylePr"):
                layers = out.setdefault(cond.get(_W + "type") or "", _StyleLayers({}, {}, {}))
                _merge(layers.ppr, _paragraph_layer(cond.first("w:pPr")))
                _merge(layers.rpr, _run_layer(cond.first("w:rPr"), self._theme))
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
                bits = int(_attr(look_el, "val") or "0", 16)
                look = {k: bool(bits & b) for k, b in _LOOK_BITS.items()}
            style = _attr(tblpr.first("w:tblStyle") if tblpr is not None else None, "val")
            self._tables[tbl] = ({r: i for i, r in enumerate(rows)}, len(rows), look,
                                 style if style is not None else self._default.get("table"))
        return self._tables[tbl]

    def _cells(self, tr: OoxElement) -> tuple[list[OoxElement], list[int]]:
        if tr not in self._rows:
            cells = self._children(tr, _W + "tc")
            trpr = tr.first("w:trPr")
            col = int(_attr(trpr.first("w:gridBefore") if trpr is not None else None, "val") or 0)
            starts = []
            for tc in cells:
                starts.append(col)
                tcpr = tc.first("w:tcPr")
                col += int(_attr(tcpr.first("w:gridSpan") if tcpr is not None else None,
                                 "val") or 1)
            self._rows[tr] = (cells, starts)
        return self._rows[tr]

    def _table_layers(self, p: OoxElement) -> _StyleLayers | None:
        """The table style's formatting for the cell holding ``p``, or None."""
        tc = None
        for anc in p.iterancestors():
            if anc.tag == _W + "tc":
                tc = anc
                break
            if anc.tag in (_W + "txbxContent", _W + "body"):
                return None
        if tc is None:
            return None
        tr = next((a for a in tc.iterancestors(_W + "tr")), None)
        tbl = next((a for a in tc.iterancestors(_W + "tbl")), None)
        if tr is None or tbl is None:
            return None
        row_of, n_rows, look, style = self._table(tbl)
        conditions, row_band, col_band = self._table_style(style)
        if not conditions:
            return None
        cells, starts = self._cells(tr)
        r = row_of.get(tr, 0)
        c = next((i for i, cell in enumerate(cells) if cell is tc), 0)
        first_row, last_row = look["firstRow"] and r == 0, look["lastRow"] and r == n_rows - 1
        first_col = look["firstColumn"] and c == 0
        last_col = look["lastColumn"] and c == len(cells) - 1
        active = {"wholeTable"}
        if not look["noVBand"] and not (first_col or last_col):
            band = (starts[c] - (1 if look["firstColumn"] else 0)) // col_band
            active.add("band1Vert" if band % 2 == 0 else "band2Vert")
        if not look["noHBand"] and not (first_row or last_row):
            band = (r - (1 if look["firstRow"] else 0)) // row_band
            active.add("band1Horz" if band % 2 == 0 else "band2Horz")
        for name, on in (("firstRow", first_row), ("lastRow", last_row),
                         ("firstCol", first_col), ("lastCol", last_col),
                         ("nwCell", first_row and first_col), ("neCell", first_row and last_col),
                         ("swCell", last_row and first_col), ("seCell", last_row and last_col)):
            if on:
                active.add(name)
        out = _StyleLayers({}, {}, {})
        for name in ("", *_CONDITIONS):
            if name in conditions and (name == "" or name in active):
                _merge(out.ppr, conditions[name].ppr)
                _merge(out.rpr, conditions[name].rpr)
        return out

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
        ppr = p.first("w:pPr")
        sid = self._paragraph_style(ppr)
        styled = self._chain(sid)
        table = self._table_layers(p)
        numbering = self._numbering.numbering_of(ppr)
        level = None
        if numbering is not None:
            defined = self._numbering.levels(numbering[0])
            level = _paragraph_layer(defined[1][numbering[1]].ppr) if defined else None
        acc: dict[str, Any] = {}
        _merge(acc, self._default_para)
        if table is not None:
            _merge(acc, table.ppr)
        if level is not None and numbering is not None and numbering[2]:
            _merge(acc, level)  # numbering from the style: the style's own values win
        _merge(acc, self._style_layer(styled, table, "ppr"))
        if level is not None and numbering is not None and not numbering[2]:
            _merge(acc, level)
        _merge(acc, _paragraph_layer(ppr))
        borders = Borders(**{side: acc.pop(f"border_{side}", None) for side in BORDER_SIDES})
        tabs = tuple(sorted((t for t in acc.pop("tabs", {}).values() if t is not None),
                            key=lambda t: t.position))
        return ParagraphFormat(
            style=self._style_name(sid), borders=borders, tabs=tabs,
            numbering=(int(numbering[0]), numbering[1]) if numbering is not None else None,
            **acc)

    def run(self, r: OoxElement) -> RunFormat:
        """The character formatting Word applies to the run ``r``."""
        return self._run_format(self._paragraph_of(r), r.first("w:rPr"))

    def mark(self, p: OoxElement) -> RunFormat:
        """The formatting of ``p``'s paragraph mark (what an empty paragraph,
        and by default a list label, looks like)."""
        ppr = p.first("w:pPr")
        return self._run_format(p, ppr.first("w:rPr") if ppr is not None else None)

    def _run_format(self, p: OoxElement, rpr: OoxElement | None) -> RunFormat:
        ppr = p.first("w:pPr")
        styled = self._chain(self._paragraph_style(ppr))
        table = self._table_layers(p)
        acc: dict[str, Any] = {}
        _merge(acc, self._default_run)
        if table is not None:
            _merge(acc, table.rpr, toggle=True)
        _merge(acc, self._style_layer(styled, table, "rpr"))  # Word resets: not a toggle
        rs = rpr.first("w:rStyle") if rpr is not None else None
        csid = rs.get(_W + "val") if rs is not None else None
        if csid in self._styles:
            _merge(acc, self._chain(csid).rpr, toggle=True)
        _merge(acc, _run_layer(rpr, self._theme))
        return RunFormat(style=self._style_name(csid) if csid in self._styles else None, **acc)
