"""Character and paragraph formatting, custom styles, effective formatting.

    p = doc.add_paragraph()
    doc.add_run(p, "Warning: ", bold=True, color="C00000")
    doc.add_run(p, "check the figures.", italic=True, size=12, font="Georgia")
    doc.format_paragraph(p, align="center", space_after=12, first_line=Cm(1))

    doc.add_style("Note", base="Normal", italic=True, color="595959", space_before=6)
    doc.add_paragraph("A note.", style="Note")

    doc.effective_format(p, run)   # {"bold": True, "size": 12.0, "font": "Georgia", ...}

Units: font size and paragraph spacing in points (as in Word's UI); indents
in EMU from ``oox.units`` (``Cm(1)``) or twips (values up to 100000 are taken
as twips); line spacing as a multiple (1.15, 1.5, 2) or ``"12pt"`` exactly.

``effective_format`` resolves what Word shows: document defaults, then the
paragraph style chain (``basedOn``), the character style chain, then direct
formatting; theme fonts are looked up in the theme.  Table styles and
numbering formatting are not applied.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import w

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
A = URI["a"]
_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}
_ALIGN_BACK = {"left": "left", "start": "left", "center": "center", "right": "right",
               "end": "right", "both": "justify", "distribute": "justify"}
RUN_OPTIONS = ("bold", "italic", "underline", "strike", "size", "font", "color", "highlight",
               "superscript", "subscript", "caps", "small_caps")
PARAGRAPH_OPTIONS = ("align", "space_before", "space_after", "line_spacing", "indent_left",
                     "indent_right", "first_line", "keep_with_next", "keep_together",
                     "page_break_before")


def _twips(value: float) -> int:
    return round(value / 635) if abs(value) > 100000 else int(value)


def _val(el: OoxElement | None, attr: str = "val") -> str | None:
    return el.get(f"{{{W}}}{attr}") if el is not None else None


def _on(el: OoxElement | None) -> bool | None:
    """ST_OnOff: present without val means on."""
    if el is None:
        return None
    return _val(el) not in ("0", "false", "off")


# -- building ----------------------------------------------------------------------------------
def apply_run_format(rpr: OoxElement, **fmt: Any) -> OoxElement:
    """Set run properties on a ``w:rPr`` (None leaves a property alone)."""
    unknown = set(fmt) - set(RUN_OPTIONS)
    if unknown:
        raise TypeError(f"unknown run format {sorted(unknown)}; known: {', '.join(RUN_OPTIONS)}")
    toggles = {"bold": (w.B, w.BCs), "italic": (w.I, w.ICs), "strike": (w.Strike, None),
               "caps": (w.Caps, None), "small_caps": (w.SmallCaps, None)}
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
        rpr.add(w.Color(val=fmt["color"].lstrip("#").upper()), replace=True)
    if fmt.get("highlight") is not None:
        rpr.add(w.Highlight(val=fmt["highlight"]), replace=True)
    if fmt.get("superscript") or fmt.get("subscript"):
        rpr.add(w.VertAlign(val="superscript" if fmt.get("superscript") else "subscript"),
                replace=True)
    return rpr


def run(text: str, *, style: str | None = None, **fmt: Any) -> OoxElement:
    """A ``w:r`` with ``text`` (tabs and line breaks converted) and formatting:
    bold, italic, underline (True or a Word style such as "double"), strike,
    size (pt), font, color (hex), highlight ("yellow" ...), superscript,
    subscript, caps, small_caps; ``style`` is a character style id."""
    from . import teletype

    rpr = w.RPr()
    if style:
        rpr.add(w.RStyle(val=style))
    apply_run_format(rpr, **fmt)
    r = w.R(rpr) if len(rpr) else w.R()
    teletype.addTextToElement(r, text)
    return r


def apply_paragraph_format(ppr: OoxElement, **fmt: Any) -> OoxElement:
    unknown = set(fmt) - set(PARAGRAPH_OPTIONS)
    if unknown:
        raise TypeError(f"unknown paragraph format {sorted(unknown)}; "
                        f"known: {', '.join(PARAGRAPH_OPTIONS)}")
    if fmt.get("align") is not None:
        ppr.add(w.Jc(val=_ALIGN[fmt["align"]]), replace=True)
    for name, factory in (("keep_with_next", w.KeepNext), ("keep_together", w.KeepLines),
                          ("page_break_before", w.PageBreakBefore)):
        if fmt.get(name) is not None:
            ppr.add(factory() if fmt[name] else factory(val=False), replace=True)
    spacing = ppr.first("w:spacing")
    sp: dict[str, Any] = dict(spacing.attrib) if spacing is not None else {}
    sp = {k.rpartition("}")[2]: v for k, v in sp.items()}
    if fmt.get("space_before") is not None:
        sp["before"] = round(float(fmt["space_before"]) * 20)
    if fmt.get("space_after") is not None:
        sp["after"] = round(float(fmt["space_after"]) * 20)
    ls = fmt.get("line_spacing")
    if ls is not None:
        m = re.fullmatch(r"\s*([0-9.]+)\s*pt\s*", str(ls))
        if m:
            sp["line"], sp["lineRule"] = round(float(m.group(1)) * 20), "exact"
        else:
            sp["line"], sp["lineRule"] = round(float(ls) * 240), "auto"
    if {"space_before", "space_after", "line_spacing"} & {k for k, v in fmt.items() if v is not None}:
        ppr.add(w.Spacing(**sp), replace=True)
    if any(fmt.get(k) is not None for k in ("indent_left", "indent_right", "first_line")):
        ind = ppr.first("w:ind")
        attrs = {k.rpartition("}")[2]: v for k, v in (ind.attrib.items() if ind is not None else [])}
        if fmt.get("indent_left") is not None:
            attrs["left"] = _twips(fmt["indent_left"])
        if fmt.get("indent_right") is not None:
            attrs["right"] = _twips(fmt["indent_right"])
        if fmt.get("first_line") is not None:
            v = _twips(fmt["first_line"])
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
    space_after (pt), line_spacing (multiple or "12pt"), indent_left /
    indent_right / first_line (EMU or twips; negative first_line = hanging),
    keep_with_next, keep_together, page_break_before."""
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
    paragraph options (paragraph styles may use both)."""
    if kind not in ("paragraph", "character"):
        raise ValueError("kind is 'paragraph' or 'character'")
    try:
        doc.style_id(name)
    except KeyError:
        pass
    else:
        raise ValueError(f"a style named {name!r} already exists")
    sid = style_id_for(name)
    existing = {st.get(f"{{{W}}}styleId") for st in doc.styles.iterchildren(f"{{{W}}}style")}
    n, base_sid = 1, sid
    while sid in existing:
        n += 1
        sid = f"{base_sid}{n}"
    run_fmt = {k: v for k, v in fmt.items() if k in RUN_OPTIONS}
    para_fmt = {k: v for k, v in fmt.items() if k in PARAGRAPH_OPTIONS}
    unknown = set(fmt) - set(run_fmt) - set(para_fmt)
    if unknown:
        raise TypeError(f"unknown format {sorted(unknown)}")
    if kind == "character" and para_fmt:
        raise TypeError("character styles take run formatting only")
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
    doc.styles.add(st)
    return sid


# -- resolving -------------------------------------------------------------------------------
def _theme_fonts(doc: WordprocessingDocument) -> dict[str, str]:
    found = doc.related_parts(doc.main_part, "ThemePart")
    out: dict[str, str] = {}
    if not found:
        return out
    root = found[0].element
    for kind in ("major", "minor"):
        el = root.find(f".//{{{A}}}{kind}Font/{{{A}}}latin")
        if el is not None:
            out[kind] = el.get("typeface", "")
    return out


def _merge_rpr(into: dict[str, Any], rpr: OoxElement | None, theme: dict[str, str]) -> None:
    if rpr is None:
        return
    for key, tag in (("bold", "b"), ("italic", "i"), ("strike", "strike"), ("caps", "caps"),
                     ("small_caps", "smallCaps")):
        v = _on(rpr.first(f"w:{tag}"))
        if v is not None:
            into[key] = v
    u = rpr.first("w:u")
    if u is not None:
        into["underline"] = None if _val(u) in (None, "none") else _val(u)
    sz = _val(rpr.first("w:sz"))
    if sz is not None:
        into["size"] = int(sz) / 2
    fonts = rpr.first("w:rFonts")
    if fonts is not None:
        if _val(fonts, "ascii"):
            into["font"] = _val(fonts, "ascii")
        elif _val(fonts, "asciiTheme"):
            theme_name = _val(fonts, "asciiTheme") or ""
            into["font"] = theme.get("major" if theme_name.startswith("major") else "minor")
    color = rpr.first("w:color")
    if color is not None:
        into["color"] = None if _val(color) in (None, "auto") else (_val(color) or "").upper()
    hl = _val(rpr.first("w:highlight"))
    if hl is not None:
        into["highlight"] = None if hl == "none" else hl
    va = _val(rpr.first("w:vertAlign"))
    if va is not None:
        into["superscript"], into["subscript"] = va == "superscript", va == "subscript"


def _merge_ppr(into: dict[str, Any], ppr: OoxElement | None) -> None:
    if ppr is None:
        return
    jc = _val(ppr.first("w:jc"))
    if jc is not None:
        into["align"] = _ALIGN_BACK.get(jc, jc)
    sp = ppr.first("w:spacing")
    if sp is not None:
        if _val(sp, "before") is not None:
            into["space_before"] = int(_val(sp, "before") or 0) / 20
        if _val(sp, "after") is not None:
            into["space_after"] = int(_val(sp, "after") or 0) / 20
        if _val(sp, "line") is not None:
            line, rule = int(_val(sp, "line") or 240), _val(sp, "lineRule") or "auto"
            into["line_spacing"] = line / 240 if rule == "auto" else f"{line / 20:g}pt"
    ind = ppr.first("w:ind")
    if ind is not None:
        left = _val(ind, "left") or _val(ind, "start")
        right = _val(ind, "right") or _val(ind, "end")
        if left is not None:
            into["indent_left"] = int(left)
        if right is not None:
            into["indent_right"] = int(right)
        if _val(ind, "firstLine") is not None:
            into["first_line"] = int(_val(ind, "firstLine") or 0)
        if _val(ind, "hanging") is not None:
            into["first_line"] = -int(_val(ind, "hanging") or 0)
    for key, tag in (("keep_with_next", "keepNext"), ("keep_together", "keepLines"),
                     ("page_break_before", "pageBreakBefore")):
        v = _on(ppr.first(f"w:{tag}"))
        if v is not None:
            into[key] = v


def _style_chain(styles: dict[str, OoxElement], sid: str | None) -> list[OoxElement]:
    chain: list[OoxElement] = []
    while sid and sid in styles and len(chain) < 20:
        st = styles[sid]
        if st in chain:
            break
        chain.append(st)
        sid = _val(st.first("w:basedOn"))
    return list(reversed(chain))  # root first


def effective_format(doc: WordprocessingDocument, paragraph: OoxElement,
                     run: OoxElement | None = None) -> dict[str, Any]:
    """Formatting Word applies to ``run`` (or to the paragraph mark) in
    ``paragraph``: run options plus paragraph options, plus ``style`` (the
    paragraph style's name).  Sizes and spacing in points, indents in twips."""
    theme = _theme_fonts(doc)
    styles_root = None
    found = doc.related_parts(doc.main_part, "StyleDefinitionsPart")
    if found:
        styles_root = found[0].element
    styles = {st.get(f"{{{W}}}styleId"): st for st in
              (styles_root.iterchildren(f"{{{W}}}style") if styles_root is not None else [])}
    out: dict[str, Any] = {"bold": False, "italic": False, "underline": None, "strike": False,
                           "size": 10.0, "font": "Times New Roman", "color": None,
                           "align": "left", "space_before": 0.0, "space_after": 0.0}
    if styles_root is not None:
        defaults = styles_root.first("w:docDefaults")
        if defaults is not None:
            rd = defaults.first("w:rPrDefault")
            pd = defaults.first("w:pPrDefault")
            _merge_rpr(out, rd.first("w:rPr") if rd is not None else None, theme)
            _merge_ppr(out, pd.first("w:pPr") if pd is not None else None)
    ppr = paragraph.first("w:pPr")
    psid = _val(ppr.first("w:pStyle")) if ppr is not None else None
    if psid is None:
        psid = next((sid for sid, st in styles.items() if st.get(f"{{{W}}}type") == "paragraph"
                     and st.get(f"{{{W}}}default") in ("1", "true")), None)
    for st in _style_chain(styles, psid):
        _merge_ppr(out, st.first("w:pPr"))
        _merge_rpr(out, st.first("w:rPr"), theme)
    out["style"] = _val(styles[psid].first("w:name")) if psid in styles else None
    _merge_ppr(out, ppr)
    rpr = run.first("w:rPr") if run is not None else (ppr.first("w:rPr") if ppr is not None else None)
    if rpr is not None:
        for st in _style_chain(styles, _val(rpr.first("w:rStyle"))):
            _merge_rpr(out, st.first("w:rPr"), theme)
        _merge_rpr(out, rpr, theme)
    return out
