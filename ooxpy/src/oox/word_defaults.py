"""The parts a new Word document gets by default: styles, settings, fonts.

Modeled on what Word itself writes, restricted to markup Office 2007 reads
(checked by ``validate(target="Office2007")`` in the tests):

* styles: document defaults (theme fonts, 11 pt, 8 pt after), Normal, the
  table/numbering/character defaults, Title, Heading 1-3, List Paragraph,
  Hyperlink, Table Grid
* settings: zoom, tab stop, theme colour mapping, locale symbols
* fonts: the theme fonts and Times New Roman

The theme comes from ``drawing.default_theme()`` (Calibri / Calibri Light).
"""

from __future__ import annotations

from typing import Any

from .element import OoxElement
from .ns import w

PAGE_SIZES = {  # twips (1/20 pt): width, height
    "A4": (11906, 16838),
    "Letter": (12240, 15840),
}


def _theme_fonts(major: bool) -> OoxElement:
    kind = "major" if major else "minor"
    return w.RFonts(asciiTheme=f"{kind}HAnsi", eastAsiaTheme=f"{kind}EastAsia",
                    hAnsiTheme=f"{kind}HAnsi", cstheme=f"{kind}Bidi")


def _ui(style: OoxElement, priority: int, *, qformat: bool = False, hidden: bool = False) -> OoxElement:
    style.add(w.UiPriority(val=priority))
    if hidden:
        style.add(w.SemiHidden(), w.UnhideWhenUsed())
    if qformat:
        style.add(w.QFormat())
    return style


def _heading(level: int, size_half_pt: int, before: int) -> OoxElement:
    return _ui(w.Style(
        w.Name(val=f"heading {level}"), w.BasedOn(val="Normal"), w.Next(val="Normal"),
        w.PPr(w.KeepNext(), w.KeepLines(), w.Spacing(before=before, after=0),
              w.OutlineLvl(val=level - 1)),
        w.RPr(_theme_fonts(major=True),
              w.Color(val="2F5496", themeColor="accent1", themeShade="BF"),
              w.Sz(val=size_half_pt), w.SzCs(val=size_half_pt)),
        type="paragraph", styleId=f"Heading{level}"), 9, qformat=True)


def styles(lang: str = "en-US") -> OoxElement:
    border: dict[str, Any] = dict(val="single", sz=4, space=0, color="auto")
    return w.Styles(
        w.DocDefaults(
            w.RPrDefault(w.RPr(_theme_fonts(major=False), w.Sz(val=22), w.SzCs(val=22),
                               w.Lang(val=lang, eastAsia=lang, bidi="ar-SA"))),
            w.PPrDefault(w.PPr(w.Spacing(after=160, line=259, lineRule="auto"))),
        ),
        _ui(w.Style(w.Name(val="Normal"), type="paragraph", default=True, styleId="Normal"),
            0, qformat=True),
        _ui(w.Style(w.Name(val="Default Paragraph Font"),
                    type="character", default=True, styleId="DefaultParagraphFont"), 1, hidden=True),
        _ui(w.Style(w.Name(val="Normal Table"),
                    w.TblPr(w.TblInd(w=0, type="dxa"),
                            w.TblCellMar(w.Top(w=0, type="dxa"), w.Left(w=108, type="dxa"),
                                         w.Bottom(w=0, type="dxa"), w.Right(w=108, type="dxa"))),
                    type="table", default=True, styleId="TableNormal"), 99, hidden=True),
        _ui(w.Style(w.Name(val="No List"), type="numbering", default=True, styleId="NoList"),
            99, hidden=True),
        _ui(w.Style(w.Name(val="Title"), w.BasedOn(val="Normal"), w.Next(val="Normal"),
                    w.PPr(w.Spacing(after=0, line=240, lineRule="auto"), w.ContextualSpacing()),
                    w.RPr(_theme_fonts(major=True), w.Spacing(val=-10), w.Kern(val=28),
                          w.Sz(val=56), w.SzCs(val=56)),
                    type="paragraph", styleId="Title"), 10, qformat=True),
        _heading(1, 32, 240),
        _heading(2, 26, 40),
        _heading(3, 24, 40),
        _ui(w.Style(w.Name(val="List Paragraph"), w.BasedOn(val="Normal"),
                    w.PPr(w.Ind(left=720), w.ContextualSpacing()),
                    type="paragraph", styleId="ListParagraph"), 34, qformat=True),
        _ui(w.Style(w.Name(val="Hyperlink"), w.BasedOn(val="DefaultParagraphFont"),
                    w.RPr(w.Color(val="0563C1", themeColor="hyperlink"), w.U(val="single")),
                    type="character", styleId="Hyperlink"), 99, hidden=True),
        _ui(w.Style(w.Name(val="Table Grid"), w.BasedOn(val="TableNormal"),
                    w.PPr(w.Spacing(after=0, line=240, lineRule="auto")),
                    w.TblPr(w.TblBorders(w.Top(**border), w.Left(**border), w.Bottom(**border),
                                         w.Right(**border), w.InsideH(**border),
                                         w.InsideV(**border))),
                    type="table", styleId="TableGrid"), 39),
    )


def settings(lang: str = "en-US") -> OoxElement:
    return w.Settings(
        w.Zoom(percent=100),
        w.DefaultTabStop(val=720),
        w.CharacterSpacingControl(val="doNotCompress"),
        w.ThemeFontLang(val=lang),
        w.ClrSchemeMapping(bg1="light1", t1="dark1", bg2="light2", t2="dark2",
                           accent1="accent1", accent2="accent2", accent3="accent3",
                           accent4="accent4", accent5="accent5", accent6="accent6",
                           hyperlink="hyperlink", followedHyperlink="followedHyperlink"),
        w.DecimalSymbol(val="."),
        w.ListSeparator(val=","),
    )


def fonts() -> OoxElement:
    def font(name: str, panose: str, family: str, pitch: str = "variable") -> OoxElement:
        return w.Font(w.Panose1(val=panose), w.Charset(val="00"), w.Family(val=family),
                      w.Pitch(val=pitch), name=name)

    return w.Fonts(
        font("Calibri", "020F0502020204030204", "swiss"),
        font("Times New Roman", "02020603050405020304", "roman"),
        font("Calibri Light", "020F0302020204030204", "swiss"),
    )


def section(page: str = "A4", landscape: bool = False) -> OoxElement:
    """``w:sectPr`` with page size and 1-inch margins."""
    wdt, hgt = PAGE_SIZES[page]
    size = w.PgSz(w=hgt, h=wdt, orient="landscape") if landscape else w.PgSz(w=wdt, h=hgt)
    return w.SectPr(size, w.PgMar(top=1440, right=1440, bottom=1440, left=1440,
                                  header=708, footer=708, gutter=0),
                    w.Cols(space=708), w.DocGrid(linePitch=360))


# -- styles added on demand (by the feature that needs them) -----------------------------
def _para(name: str, sid: str, *children: OoxElement, based_on: str = "Normal",
          priority: int = 99, hidden: bool = True, next_style: str | None = None) -> OoxElement:
    st = w.Style(w.Name(val=name), w.BasedOn(val=based_on), type="paragraph", styleId=sid)
    if next_style:
        st.add(w.Next(val=next_style))
    st.add(*children)
    return _ui(st, priority, hidden=hidden)


def _char(name: str, sid: str, rpr: OoxElement, priority: int = 99) -> OoxElement:
    return _ui(w.Style(w.Name(val=name), w.BasedOn(val="DefaultParagraphFont"), rpr,
                       type="character", styleId=sid), priority, hidden=True)


def extra_style(style_id: str, text_width: int = 9026) -> OoxElement | None:
    """Definition of a style Word adds when a feature is first used."""
    small = w.RPr(w.Sz(val=20), w.SzCs(val=20))
    tight = w.Spacing(after=0, line=240, lineRule="auto")
    if style_id in ("Header", "Footer"):
        tabs = w.Tabs(w.Tab(val="center", pos=text_width // 2), w.Tab(val="right", pos=text_width))
        return _para(style_id.lower(), style_id, w.PPr(tabs, w.Spacing(after=0, line=240, lineRule="auto")))
    if style_id == "FootnoteText":
        return _para("footnote text", style_id, w.PPr(tight), small)
    if style_id == "EndnoteText":
        return _para("endnote text", style_id, w.PPr(w.Spacing(after=0, line=240, lineRule="auto")),
                     w.RPr(w.Sz(val=20), w.SzCs(val=20)))
    if style_id == "CommentText":
        return _para("annotation text", style_id, w.PPr(w.Spacing(line=240, lineRule="auto")),
                     w.RPr(w.Sz(val=20), w.SzCs(val=20)))
    if style_id in ("FootnoteReference", "EndnoteReference"):
        name = "footnote reference" if style_id.startswith("Foot") else "endnote reference"
        return _char(name, style_id, w.RPr(w.VertAlign(val="superscript")))
    if style_id == "CommentReference":
        return _char("annotation reference", style_id, w.RPr(w.Sz(val=16), w.SzCs(val=16)))
    if style_id == "Hyperlink":
        return _char("Hyperlink", style_id, w.RPr(w.Color(val="0563C1", themeColor="hyperlink"),
                                                  w.U(val="single")))
    if style_id.startswith("Heading") and style_id[7:] in tuple("456789"):
        return _heading(int(style_id[7:]), 22, 40)  # Word adds Heading 4-9 on first use
    if style_id == "Caption":
        return _para("caption", style_id, w.PPr(w.Spacing(after=200, line=240, lineRule="auto")),
                     w.RPr(w.I(), w.ICs(), w.Color(val="44546A", themeColor="text2"),
                           w.Sz(val=18), w.SzCs(val=18)), priority=35, hidden=True,
                     next_style="Normal")
    if style_id == "PlaceholderText":
        return _char("Placeholder Text", style_id, w.RPr(w.Color(val="808080")))
    if style_id == "TOCHeading":
        return _para("TOC Heading", style_id, w.PPr(w.OutlineLvl(val=9)), based_on="Heading1",
                     priority=39, hidden=False, next_style="Normal")
    if style_id.startswith("TOC") and style_id[3:].isdigit():
        level = int(style_id[3:])
        return _para(f"toc {level}", style_id,
                     w.PPr(w.Spacing(after=100), w.Ind(left=220 * (level - 1))),
                     priority=39, next_style="Normal")
    return None
