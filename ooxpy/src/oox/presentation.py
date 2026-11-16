"""PresentationML (.pptx): the ``OpenDocumentPresentation()`` equivalent.

    from oox.presentation import PresentationDocument
    from oox.units import Inches, Pt

    deck = PresentationDocument.new()                       # 16:9, blank layout
    slide = deck.add_slide()
    deck.add_textbox(slide, Inches(1), Inches(1), Inches(8), Inches(1),
                     "Quarterly review", size=40, bold=True)
    deck.add_picture(slide, "chart.png", Inches(1), Inches(2.5), width=Inches(5))
    deck.save("review.pptx")

Placeholders work with any deck -- ours or one made in PowerPoint:

    slide = deck.add_slide("Title and Content")      # layout by name
    deck.set_placeholder_text(slide, "title", "Agenda")
    deck.set_placeholder_text(slide, "body", ["Results", (1, "by region"), "Outlook"])

A presentation can't exist without a slide master, a slide layout and a
theme, so ``new()`` builds that chain: presentation -> master -> Office's
eleven standard layouts, master -> theme.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from . import drawing
from .document import RT_OFFICE_DOCUMENT, OoxDocument
from .element import OoxElement
from .errors import PackageError
from .grammar import URI
from .ns import a, p
from .opc import Package, Part
from .units import Inches

if TYPE_CHECKING:
    from .slidetable import SlideTable

CT_PML_MAIN = {
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    ".pptm": "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    ".potx": "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
    ".ppsx": "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
}
# ST_SlideMasterId / ST_SlideLayoutId start here; ST_SlideId at 256
_MIN_MASTER_ID, _MIN_SLIDE_ID = 2147483648, 256
_DEFAULT_TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"  # "Medium Style 2 - Accent 1"
_ALIGN = {"left": "l", "center": "ctr", "right": "r", "justify": "just"}


def _empty_tree() -> OoxElement:
    return p.SpTree(
        p.NvGrpSpPr(p.CNvPr(id=1, name=""), p.CNvGrpSpPr(), p.NvPr()),
        p.GrpSpPr(a.Xfrm(a.Off(x=0, y=0), a.Ext(cx=0, cy=0),
                         a.ChOff(x=0, y=0), a.ChExt(cx=0, cy=0))),
    )


def _ph_shape(shape_id: int, name: str, ph: OoxElement, geometry: tuple | None = None,
              body_pr: OoxElement | None = None, lst_style: OoxElement | None = None) -> OoxElement:
    """A placeholder ``p:sp``; without ``geometry`` it inherits position from above."""
    sp_pr = p.SpPr(a.Xfrm(a.Off(x=geometry[0], y=geometry[1]),
                          a.Ext(cx=geometry[2], cy=geometry[3]))) if geometry else p.SpPr()
    return p.Sp(
        p.NvSpPr(p.CNvPr(id=shape_id, name=name), p.CNvSpPr(a.SpLocks(noGrp=True)), p.NvPr(ph)),
        sp_pr,
        p.TxBody(body_pr if body_pr is not None else a.BodyPr(),
                 lst_style if lst_style is not None else a.LstStyle(),
                 a.P(a.EndParaRPr(lang="en-US"))),
    )


def _scaled(geom: tuple, sx: float, sy: float) -> tuple:
    x0, y0, cx, cy = geom
    return (round(x0 * sx), round(y0 * sy), round(cx * sx), round(cy * sy))


# Office's 16:9 placeholder geometry (EMU); scaled for other slide sizes
_TITLE = (838200, 365125, 10515600, 1325563)
_BODY = (838200, 1825625, 10515600, 4351338)
_CTR_TITLE = (1524000, 1122363, 9144000, 2387600)
_SUBTITLE = (1524000, 3602038, 9144000, 1655762)


def _level(lvl: int, size: int, mar: int, bullet: bool) -> OoxElement:
    ppr = getattr(a, f"Lvl{lvl}pPr")(a.SpcBef(a.SpcPts(val=1000)))
    if bullet:
        ppr.add(a.BuFont(typeface="Arial"), a.BuChar(char="\u2022"))
        ppr.set_attr("marL", mar)
        ppr.set_attr("indent", -228600)
    ppr.add(a.DefRPr(a.SolidFill(a.SchemeClr(val="tx1")), a.Latin(typeface="+mn-lt"),
                     sz=size, kern=1200))
    return ppr


def _master(sx: float, sy: float) -> OoxElement:
    tree = _empty_tree()
    tree.add(
        _ph_shape(2, "Title Placeholder 1", p.Ph(type="title"), _scaled(_TITLE, sx, sy),
                  a.BodyPr(a.NormAutofit(), anchor="ctr")),
        _ph_shape(3, "Text Placeholder 2", p.Ph(type="body", idx=1), _scaled(_BODY, sx, sy),
                  a.BodyPr(a.NormAutofit())),
    )
    return p.SldMaster(
        p.CSld(p.Bg(p.BgRef(a.SchemeClr(val="bg1"), idx=1001)), tree),
        p.ClrMap(bg1="lt1", tx1="dk1", bg2="lt2", tx2="dk2", accent1="accent1",
                 accent2="accent2", accent3="accent3", accent4="accent4",
                 accent5="accent5", accent6="accent6", hlink="hlink", folHlink="folHlink"),
        p.SldLayoutIdLst(),
        p.TxStyles(
            p.TitleStyle(a.Lvl1pPr(a.DefRPr(a.SolidFill(a.SchemeClr(val="tx1")),
                                            a.Latin(typeface="+mj-lt"), sz=4400, kern=1200),
                                   algn="l")),
            p.BodyStyle(_level(1, 2800, 228600, True), _level(2, 2400, 685800, True),
                        _level(3, 2000, 1143000, True)),
            p.OtherStyle(a.Lvl1pPr(a.DefRPr(sz=1800))),
        ),
    )


def _layout(name: str, type_: str, shapes: list[OoxElement]) -> OoxElement:
    tree = _empty_tree()
    if shapes:
        tree.add(*shapes)
    return p.SldLayout(p.CSld(tree, name=name), p.ClrMapOvr(a.MasterClrMapping()),
                       type=type_, preserve=True)


def _plain(size: int, bold: bool = False) -> OoxElement:
    """List style: no bullets, ``size`` (1/100 pt)."""
    return a.LstStyle(a.Lvl1pPr(a.BuNone(), a.DefRPr(sz=size, b=True if bold else None),
                                marL=0, indent=0))


def _default_layouts(sx: float, sy: float, which: str = "standard") -> list[OoxElement]:
    """Office's layouts with its 16:9 geometry (scaled).  ``which="basic"``:
    Title Slide, Title and Content, Blank (oox before 0.0.2)."""
    def g(*geom):
        return _scaled(geom, sx, sy)

    centred = lambda size: a.LstStyle(a.Lvl1pPr(a.BuNone(), a.DefRPr(sz=size),
                                                marL=0, indent=0, algn="ctr"))
    title_slide = _layout("Title Slide", "title", [
        _ph_shape(2, "Title 1", p.Ph(type="ctrTitle"), _scaled(_CTR_TITLE, sx, sy),
                  a.BodyPr(anchor="b"), centred(6000)),
        _ph_shape(3, "Subtitle 2", p.Ph(type="subTitle", idx=1), _scaled(_SUBTITLE, sx, sy),
                  None, centred(2400)),
    ])
    title_content = _layout("Title and Content", "obj", [
        _ph_shape(2, "Title 1", p.Ph(type="title")),
        _ph_shape(3, "Content Placeholder 2", p.Ph(idx=1)),
    ])
    blank = _layout("Blank", "blank", [])
    if which == "basic":
        return [title_slide, title_content, blank]
    if which != "standard":
        raise ValueError("layouts must be 'standard' or 'basic'")
    caption_title = (839788, 457200, 3932237, 1600200)
    caption_text = (839788, 2057400, 3932237, 3811588)
    caption_box = (5183188, 987425, 6172200, 4873625)
    return [
        title_slide,
        title_content,
        _layout("Section Header", "secHead", [
            _ph_shape(2, "Title 1", p.Ph(type="title"), g(831850, 1709738, 10515600, 2852737),
                      a.BodyPr(anchor="b"), a.LstStyle(a.Lvl1pPr(a.DefRPr(sz=6000)))),
            _ph_shape(3, "Text Placeholder 2", p.Ph(type="body", idx=1),
                      g(831850, 4589463, 10515600, 1500187), None, _plain(2400)),
        ]),
        _layout("Two Content", "twoObj", [
            _ph_shape(2, "Title 1", p.Ph(type="title")),
            _ph_shape(3, "Content Placeholder 2", p.Ph(sz="half", idx=1),
                      g(838200, 1825625, 5181600, 4351338)),
            _ph_shape(4, "Content Placeholder 3", p.Ph(sz="half", idx=2),
                      g(6172200, 1825625, 5181600, 4351338)),
        ]),
        _layout("Comparison", "twoTxTwoObj", [
            _ph_shape(2, "Title 1", p.Ph(type="title"), g(839788, 365125, 10515600, 1325563)),
            _ph_shape(3, "Text Placeholder 2", p.Ph(type="body", idx=1),
                      g(839788, 1681163, 5157787, 823912), a.BodyPr(anchor="b"),
                      _plain(2400, bold=True)),
            _ph_shape(4, "Content Placeholder 3", p.Ph(sz="half", idx=2),
                      g(839788, 2505075, 5157787, 3684588)),
            _ph_shape(5, "Text Placeholder 4", p.Ph(type="body", sz="quarter", idx=3),
                      g(6172200, 1681163, 5183188, 823912), a.BodyPr(anchor="b"),
                      _plain(2400, bold=True)),
            _ph_shape(6, "Content Placeholder 5", p.Ph(sz="quarter", idx=4),
                      g(6172200, 2505075, 5183188, 3684588)),
        ]),
        _layout("Title Only", "titleOnly", [_ph_shape(2, "Title 1", p.Ph(type="title"))]),
        blank,
        _layout("Content with Caption", "objTx", [
            _ph_shape(2, "Title 1", p.Ph(type="title"), g(*caption_title), a.BodyPr(anchor="b"),
                      a.LstStyle(a.Lvl1pPr(a.DefRPr(sz=3200)))),
            _ph_shape(3, "Content Placeholder 2", p.Ph(idx=1), g(*caption_box)),
            _ph_shape(4, "Text Placeholder 3", p.Ph(type="body", sz="half", idx=2),
                      g(*caption_text), None, _plain(1600)),
        ]),
        _layout("Picture with Caption", "picTx", [
            _ph_shape(2, "Title 1", p.Ph(type="title"), g(*caption_title), a.BodyPr(anchor="b"),
                      a.LstStyle(a.Lvl1pPr(a.DefRPr(sz=3200)))),
            _ph_shape(3, "Picture Placeholder 2", p.Ph(type="pic", idx=1), g(*caption_box),
                      None, _plain(3200)),
            _ph_shape(4, "Text Placeholder 3", p.Ph(type="body", sz="half", idx=2),
                      g(*caption_text), None, _plain(1600)),
        ]),
        _layout("Title and Vertical Text", "vertTx", [
            _ph_shape(2, "Title 1", p.Ph(type="title")),
            _ph_shape(3, "Vertical Text Placeholder 2", p.Ph(type="body", orient="vert", idx=1),
                      None, a.BodyPr(vert="eaVert")),
        ]),
        _layout("Vertical Title and Text", "vertTitleAndTx", [
            _ph_shape(2, "Vertical Title 1", p.Ph(type="title", orient="vert"),
                      g(8724900, 365125, 2628900, 5811838), a.BodyPr(vert="eaVert")),
            _ph_shape(3, "Vertical Text Placeholder 2", p.Ph(type="body", orient="vert", idx=1),
                      g(838200, 365125, 7734300, 5811838), a.BodyPr(vert="eaVert")),
        ]),
    ]


# placeholders PowerPoint does not copy onto new slides
_NOT_CLONED = {"dt", "ftr", "sldNum", "hdr"}
_ALIASES = {"title": {"title", "ctrTitle"}, "subtitle": {"subTitle"}, "body": {"body", "obj"}}


def placeholder_info(shape: OoxElement) -> tuple[str, int] | None:
    """``(type, idx)`` of a placeholder shape (defaults ``obj``/0), else ``None``."""
    for nv in shape:
        nvpr = nv.first("p:nvPr") if isinstance(nv, OoxElement) else None
        if nvpr is not None:
            ph = nvpr.first("p:ph")
            if ph is not None:
                return ph.get("type", "obj"), int(ph.get("idx", "0"))
    return None


class PresentationDocument(OoxDocument):
    MAIN_CONTENT_TYPES = CT_PML_MAIN

    @classmethod
    def new(cls, width: int = 12192000, height: int = 6858000,
            kind: str = ".pptx", layouts: str = "standard") -> PresentationDocument:
        """New deck (default 13.333 x 7.5 in, i.e. 16:9) with master, theme and
        Office's eleven standard layouts (Title Slide, Title and Content, Section
        Header, Two Content, Comparison, Title Only, Blank, Content with Caption,
        Picture with Caption, Title and Vertical Text, Vertical Title and Text);
        ``layouts="basic"`` gives only Title Slide, Title and Content, Blank."""
        pkg = Package()
        part = pkg.add_part("/ppt/presentation.xml", CT_PML_MAIN[kind], element=p.Presentation(
            p.SldMasterIdLst(), p.SldSz(cx=width, cy=height), p.NotesSz(cx=6858000, cy=9144000)))
        pkg.relate(None, part, RT_OFFICE_DOCUMENT)
        doc = cls(pkg)

        sx, sy = width / 12192000, height / 6858000
        master, mrid = doc.add_part("SlideMasterPart", element=_master(sx, sy))
        doc.presentation.need("p:sldMasterIdLst").add(p.SldMasterId(id=_MIN_MASTER_ID, r_id=mrid))
        theme, _ = doc.add_part("ThemePart", element=drawing.default_theme())
        doc.relate(master, theme, "ThemePart")
        for layout in _default_layouts(sx, sy, layouts):
            doc.add_layout(layout, master)
        doc.add_part("PresentationPropertiesPart", element=p.PresentationPr())
        doc.add_part("ViewPropertiesPart", element=p.ViewPr())
        doc.add_part("TableStylesPart", element=a.TblStyleLst(def_=_DEFAULT_TABLE_STYLE))
        return doc

    # -- structure ------------------------------------------------------------
    @property
    def presentation(self) -> OoxElement:
        """The root ``p:presentation`` element of the main part."""
        return self.main_part.element

    @property
    def masters(self) -> list[Part]:
        """The slide master parts."""
        return self.related_parts(self.main_part, "SlideMasterPart")

    @property
    def layouts(self) -> list[Part]:
        """Layout parts in the masters' own order (``p:sldLayoutIdLst``)."""
        out = []
        for m in self.masters:
            lst = m.element.first("p:sldLayoutIdLst")
            for entry in (lst if lst is not None else []):
                rel = m.rels.items[entry.get_attr("r:id")]
                out.append(self.package.parts[m.resolve(rel.target)])
        return out

    @property
    def layout_names(self) -> list[str]:
        """Names of the slide layouts, in master order."""
        return [lp.element.need("p:cSld").get("name", "") for lp in self.layouts]

    def layout(self, name: str) -> Part:
        """The slide layout part named ``name`` ("Title and Content"); KeyError lists the
        available names."""
        for lp, n in zip(self.layouts, self.layout_names, strict=True):
            if n == name:
                return lp
        raise KeyError(f"no layout {name!r}; available: {', '.join(self.layout_names)}")

    def _next_master_or_layout_id(self) -> int:
        """Master and layout ids share one number space, starting at 2^31."""
        used = [_MIN_MASTER_ID - 1]
        for el in self.presentation.getElementsByType(p.SldMasterId):
            used.append(int(el.get("id")))
        for m in self.masters:
            for el in m.element.getElementsByType(p.SldLayoutId):
                used.append(int(el.get("id")))
        return max(used) + 1

    def add_layout(self, layout: OoxElement, master: Part | None = None) -> Part:
        """Add a ``p:sldLayout`` under ``master`` (default: the first master)."""
        master = master or self.masters[0]
        part, rid = self.add_part("SlideLayoutPart", master, element=layout)
        self.relate(part, master, "SlideMasterPart")
        lid = self._next_master_or_layout_id()
        lst = master.element.first("p:sldLayoutIdLst")
        if lst is None:  # optional in the schema; masters from other producers may lack it
            lst = p.SldLayoutIdLst()
            master.element.add(lst)
        lst.add(p.SldLayoutId(id=lid, r_id=rid))
        return part

    @property
    def slides(self) -> list[OoxElement]:
        """Slide roots (``p:sld``) in presentation order."""
        lst = self.presentation.first("p:sldIdLst")
        if lst is None:
            return []
        out = []
        for entry in lst:
            rel = self.main_part.rels.items[entry.get_attr("r:id")]
            out.append(self.package.parts[self.main_part.resolve(rel.target)].element)
        return out

    def add_slide(self, layout: Part | str | None = None) -> OoxElement:
        """Append a slide based on ``layout`` (a layout part or its name).

        The default is the layout named "Blank" if there is one, else the
        first.  The layout's placeholders are copied onto the slide (empty,
        inheriting position and style), except date/footer/slide number --
        the same as PowerPoint's "New Slide".
        """
        if isinstance(layout, str):
            layout = self.layout(layout)
        elif layout is None:
            names = self.layout_names
            layout = self.layouts[names.index("Blank")] if "Blank" in names else self.layouts[0]
        tree = _empty_tree()
        n = 1
        for shape in layout.element.need("p:cSld").need("p:spTree"):
            info = placeholder_info(shape) if isinstance(shape, OoxElement) else None
            if info is None or info[0] in _NOT_CLONED:
                continue
            n += 1
            ph = copy.deepcopy(next(shape.iter(f"{{{URI['p']}}}ph")))
            nv = shape.first("p:nvSpPr")
            cnv = nv.first("p:cNvPr") if nv is not None else None
            name = cnv.get("name") if cnv is not None else f"Placeholder {n}"
            tree.add(_ph_shape(n, name, ph))
        sld = p.Sld(p.CSld(tree), p.ClrMapOvr(a.MasterClrMapping()))
        part, rid = self.add_part("SlidePart", element=sld)
        self.relate(part, layout, "SlideLayoutPart")
        lst = self.presentation.first("p:sldIdLst")
        if lst is None:
            lst = p.SldIdLst()
            self.presentation.add(lst)  # inserted in schema order (after sldMasterIdLst)
        sid = max((int(e.get("id")) for e in lst), default=_MIN_SLIDE_ID - 1) + 1
        lst.add(p.SldId(id=sid, r_id=rid))
        from .sections_pml import join_last_section

        join_last_section(self, sid)
        return sld

    def set_sections(self, sections) -> None:
        """Slide sections: ``[(name, first slide), ...]`` (slide elements or
        1-based numbers); ``[]`` removes them.  See ``oox.sections_pml``."""
        from .sections_pml import set_sections

        set_sections(self, sections)

    def sections(self) -> list[tuple[str, list[int]]]:
        """``[(name, [slide numbers]), ...]``."""
        from .sections_pml import sections

        return sections(self)

    # -- placeholders -----------------------------------------------------------
    def placeholders(self, slide: OoxElement) -> list[tuple[str, int, OoxElement]]:
        """``(type, idx, shape)`` for each placeholder on a slide (or layout)."""
        out = []
        for shape in self._tree(slide):
            info = placeholder_info(shape) if isinstance(shape, OoxElement) else None
            if info is not None:
                out.append((info[0], info[1], shape))
        return out

    def placeholder(self, slide: OoxElement, which: str | int) -> OoxElement:
        """Find a placeholder by idx (int), by type ("ctrTitle", "obj", ...) or by
        the aliases "title" (title/ctrTitle), "subtitle", "body" (body/obj)."""
        phs = self.placeholders(slide)
        if isinstance(which, int):
            hits = [sh for _, idx, sh in phs if idx == which]
        else:
            types = _ALIASES.get(which, {which})
            hits = [sh for t, _, sh in sorted(phs, key=lambda e: e[1]) if t in types]
        if not hits:
            have = ", ".join(f"{t}/{i}" for t, i, _ in phs) or "none"
            raise KeyError(f"no placeholder {which!r} on this slide (has: {have})")
        return hits[0]

    def set_placeholder_text(self, slide: OoxElement, which: str | int,
                             text: str | list, *, lang: str = "en-US") -> OoxElement:
        """Replace a placeholder's text.  ``text`` is a string (``\\n`` splits
        paragraphs) or a list of strings / ``(level, string)`` for indented
        bullets (level 0 = top).  Formatting comes from the layout and master.
        """
        shape = self.placeholder(slide, which)
        body = shape.first("p:txBody")
        if body is None:
            raise PackageError(f"placeholder {which!r} holds no text (e.g. a picture placeholder)")
        items = text.split("\n") if isinstance(text, str) else list(text)
        for old in list(body.iterchildren(f"{{{URI['a']}}}p")):
            body.remove(old)
        for item in items:
            level, line = item if isinstance(item, tuple) else (0, item)
            para = a.P(a.R(a.RPr(lang=lang, dirty=False), a.T(str(line))))
            if level:
                para.add(a.PPr(lvl=level))
            body.add(para)
        return shape

    # -- content --------------------------------------------------------------
    @staticmethod
    def _tree(slide: OoxElement) -> OoxElement:
        return slide.need("p:cSld").need("p:spTree")

    def add_textbox(self, slide: OoxElement, x: int, y: int, cx: int, cy: int, text: str, *,
                    size: float | None = None, bold: bool = False, color: str | None = None,
                    align: str | None = None, lang: str = "en-US") -> OoxElement:
        """Add a text box; ``\\n`` starts a new paragraph.  ``size`` in points,
        ``color`` as hex RGB, ``align`` one of left/center/right/justify."""
        n = drawing.max_shape_id(slide) + 1
        paragraphs = []
        for line in text.split("\n"):
            rpr = a.RPr(lang=lang, sz=int(size * 100) if size else None,
                        b=True if bold else None, dirty=False)
            if color:
                rpr.add(a.SolidFill(a.SrgbClr(val=color)))
            para = a.P(a.R(rpr, a.T(line)))
            if align:
                para.add(a.PPr(algn=_ALIGN[align]))
            paragraphs.append(para)
        sp = p.Sp(
            p.NvSpPr(p.CNvPr(id=n, name=f"TextBox {n}"), p.CNvSpPr(txBox=True), p.NvPr()),
            drawing.sp_pr(p, x, y, cx, cy, a.NoFill()),
            p.TxBody(a.BodyPr(a.SpAutoFit(), wrap="square", rtlCol=False), a.LstStyle(),
                     *paragraphs),
        )
        self._tree(slide).add(sp)
        return sp

    def add_chart(self, slide: OoxElement, kind: str, categories, series,
                  x: int = Inches(1), y: int = Inches(1.5), cx: int = Inches(8),
                  cy: int = Inches(4.5), *, name: str | None = None, **options) -> OoxElement:
        """Place a chart on a slide; returns its ``p:graphicFrame``.

        ``series`` is ``{"North": [1, 2, 3], ...}`` (or ``oox.charts.Series``
        objects); the data also goes into an embedded workbook ("Edit Data").
        Options are those of ``oox.charts.build_chart``.
        """
        from .charts import CHART_URI, add_embedded_chart
        from .ns import c

        rid, _ = add_embedded_chart(self, self.part_of(slide), kind, list(categories), series,
                                    **options)
        n = drawing.max_shape_id(slide) + 1
        frame = p.GraphicFrame(
            p.NvGraphicFramePr(p.CNvPr(id=n, name=name or f"Chart {n}"),
                               p.CNvGraphicFramePr(a.GraphicFrameLocks(noGrp=True)), p.NvPr()),
            p.Xfrm(a.Off(x=x, y=y), a.Ext(cx=cx, cy=cy)),
            a.Graphic(a.GraphicData(c.Chart(r_id=rid), uri=CHART_URI)),
        )
        self._tree(slide).add(frame)
        return frame

    def charts(self, slide: OoxElement) -> list[dict]:
        """Charts on a slide, in order: ``oox.charts.chart_info`` plus "rid"."""
        from .charts import charts_in

        return charts_in(self, self.part_of(slide), slide)

    def slide_layout(self, slide: OoxElement) -> str:
        """Name of the layout a slide is based on."""
        found = self.related_parts(self.part_of(slide), "SlideLayoutPart")
        return found[0].element.need("p:cSld").get("name", "") if found else ""

    def add_shape(self, slide: OoxElement, shape: str, x: int, y: int, cx: int, cy: int,
                  **options) -> OoxElement:
        """An auto shape with fill, line and text; see ``oox.slides.add_shape``."""
        from .slides import add_shape

        return add_shape(self, slide, shape, x, y, cx, cy, **options)

    def add_table(self, slide: OoxElement, rows, x: int, y: int, cx: int = Inches(8),
                  cy: int | None = None, **options) -> OoxElement:
        """A table in a built-in PowerPoint style; see ``oox.slides.add_table``."""
        from .slides import add_table

        return add_table(self, slide, rows, x, y, cx, cy, **options)

    def tables(self, slide: OoxElement) -> list[dict]:
        """``{"name", "rows", "merges", ...}`` for each table on the slide
        (``deck.table(slide, i)`` edits one)."""
        from .slides import tables

        return tables(self, slide)

    def table(self, slide: OoxElement, which: int | str = 0) -> SlideTable:
        """A table on the slide, to edit: by position among the slide's
        tables (0-based) or by shape name; see ``oox.slidetable``."""
        from .slidetable import SlideTable

        frames = [f for f in self._tree(slide).iter(f"{{{URI['p']}}}graphicFrame")
                  if f.find(f".//{{{URI['a']}}}tbl") is not None]
        if isinstance(which, str):
            for f in frames:
                cnv = f.find(f".//{{{URI['p']}}}cNvPr")
                if cnv is not None and cnv.get("name") == which:
                    return SlideTable(self, f)
            raise KeyError(f"no table named {which!r} on this slide")
        try:
            return SlideTable(self, frames[which])
        except IndexError:
            raise IndexError(f"table {which}: the slide has {len(frames)}") from None

    def set_link(self, slide: OoxElement, shape, target, *, text: str | None = None,
                 tooltip: str | None = None) -> OoxElement:
        """Link a shape (element or name), or ``text`` in it, to a URL, a slide,
        or "next"/"previous"/"first"/"last"/"end"; ``None`` removes it.  See
        ``oox.slidelinks``."""
        from .slidelinks import set_link

        return set_link(self, slide, shape, target, text=text, tooltip=tooltip)

    def links(self, slide: OoxElement) -> list[dict]:
        """The slide's click links; see ``oox.slidelinks.links``."""
        from .slidelinks import links

        return links(self, slide)

    def set_notes(self, slide: OoxElement, text: str) -> OoxElement:
        """Set a slide's speaker notes (creates the notes master if needed)."""
        from .slides import set_notes

        return set_notes(self, slide, text)

    def notes(self, slide: OoxElement) -> str | None:
        """A slide's speaker notes, or None if it has none."""
        from .slides import notes

        return notes(self, slide)

    def set_placeholder_picture(self, slide: OoxElement, which: str | int, image, *,
                                crop: bool = True, descr: str = "") -> OoxElement:
        """Fill a picture (or content) placeholder, cropped to fill like PowerPoint."""
        from .slides import set_placeholder_picture

        return set_placeholder_picture(self, slide, which, image, crop=crop, descr=descr)

    def shapes(self, slide: OoxElement) -> list[dict]:
        """What a slide holds; see ``oox.slides.shapes``."""
        from .slides import shapes

        return shapes(self, slide)

    def add_picture(self, slide: OoxElement, image, x: int = Inches(1), y: int = Inches(1), *,
                    width: int | None = None, height: int | None = None,
                    descr: str = "", fallback=None) -> OoxElement:
        """Place a picture on a slide.  Size in EMU; one dimension keeps the
        aspect ratio.  An SVG needs ``fallback=`` (a PNG) for Office before 2016."""
        rid, info, svg_rid = self.add_picture_image(image, fallback, self.part_of(slide))
        cx, cy = info.scaled(width, height)
        n = drawing.max_shape_id(slide) + 1
        pic = drawing.picture(p, n, f"Picture {n}", rid, cx, cy, x, y, descr=descr, nv_pr=True,
                              svg_rid=svg_rid)
        self._tree(slide).add(pic)
        return pic
