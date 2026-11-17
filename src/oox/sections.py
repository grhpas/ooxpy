"""Sections: page setup, headers and footers.

    sec = doc.sections[0]
    sec.footer_text("Page {PAGE} of {NUMPAGES}", align="center")
    sec.header_text("Quarterly report", align="right")
    sec.different_first_page = True
    sec.header_text("Cover page header", kind="first")

    wide = doc.add_section(landscape=True)      # new section on a new page
    wide.header_text("Appendix (landscape)")

How OOXML stores it: the *last* section's properties are ``w:body/w:sectPr``;
every earlier section ends with a paragraph whose ``w:pPr`` holds that
section's ``w:sectPr``.  Headers and footers are separate parts, referenced
per section by kind ("default", "first", "even").  A section without its own
reference shows the previous section's header ("Link to Previous"); asking a
section for ``header(...)`` creates its own.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .errors import OoxError
from .grammar import URI
from .ns import w
from .units import PAGE_LIMIT, twips, twips_to_points

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
R = URI["r"]
KINDS = ("default", "first", "even")
_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}
# What a section without w:pgSz / w:pgMar gets: Word's Letter page and the
# margins set_margins starts from.  Reading reports them without writing them.
DEFAULT_PAGE = {"w": 12240, "h": 15840}
DEFAULT_MARGINS = {"top": 1440, "right": 1440, "bottom": 1440, "left": 1440, "header": 708,
                   "footer": 708, "gutter": 0}


class Section:
    """One section of a Word document: page setup, headers and footers.

    Reading never changes the document: a section without page size or
    margins reports Word's defaults, and a document without any section
    (``w:sectPr``) gets a default section that is added to it only when
    something is set on it."""

    def __init__(self, doc: WordprocessingDocument, sect_pr: OoxElement) -> None:
        self.doc, self.sect_pr = doc, sect_pr

    def _attach(self) -> None:
        """Before writing: add the default section to the body if it isn't there."""
        if self.sect_pr.getparent() is None:
            self.doc.body.add(self.sect_pr)

    def __repr__(self) -> str:
        return f"<Section {self.orientation} {self.page_width:g}x{self.page_height:g} pt>"

    # -- page setup ---------------------------------------------------------------
    def _pg_sz(self, create: bool = False) -> OoxElement | None:
        sz = self.sect_pr.first("w:pgSz")
        if sz is None and create:
            self._attach()
            sz = w.PgSz(w=DEFAULT_PAGE["w"], h=DEFAULT_PAGE["h"])  # Word's default: Letter
            self.sect_pr.add(sz)
        return sz

    def _page_twips(self, attr: str) -> int:
        sz = self._pg_sz()
        value = sz.get(f"{{{W}}}{attr}") if sz is not None else None
        return int(value) if value is not None else DEFAULT_PAGE[attr]

    @property
    def page_width(self) -> float:
        """Page width in points (Letter's when the section doesn't say)."""
        return twips_to_points(self._page_twips("w"))

    @property
    def page_height(self) -> float:
        """Page height in points (Letter's when the section doesn't say)."""
        return twips_to_points(self._page_twips("h"))

    def set_page_size(self, width: float, height: float) -> None:
        """Page size in points (or ``Length``); the orientation follows."""
        wd, ht = (twips(v, name, PAGE_LIMIT) for v, name in ((width, "width"),
                                                             (height, "height")))
        sz = self._pg_sz(create=True)
        assert sz is not None
        sz.set_attr("w", wd)
        sz.set_attr("h", ht)
        sz.set_attr("orient", "landscape" if wd > ht else None)

    @property
    def orientation(self) -> str:
        """"portrait" or "landscape" (assignable; assigning swaps width and height)."""
        sz = self._pg_sz()
        orient = sz.get(f"{{{W}}}orient") if sz is not None else None
        if orient:
            return orient
        return "landscape" if self.page_width > self.page_height else "portrait"

    @orientation.setter
    def orientation(self, value: str) -> None:
        if value not in ("portrait", "landscape"):
            raise ValueError("orientation is 'portrait' or 'landscape'")
        if value == self.orientation:
            return
        wd, ht = self._page_twips("w"), self._page_twips("h")
        sz = self._pg_sz(create=True)
        assert sz is not None
        sz.set_attr("w", ht)
        sz.set_attr("h", wd)
        sz.set_attr("orient", "landscape" if value == "landscape" else None)

    def _margins_twips(self) -> dict[str, int]:
        out = dict(DEFAULT_MARGINS)
        mar = self.sect_pr.first("w:pgMar")
        if mar is not None:
            out.update({k.rpartition("}")[2]: int(v) for k, v in mar.attrib.items()
                        if v.lstrip("-").isdigit()})
        return out

    @property
    def margins(self) -> dict[str, float]:
        """Points: top, right, bottom, left, header, footer, gutter; Word's
        defaults where the section doesn't say."""
        return {k: twips_to_points(v) for k, v in self._margins_twips().items()}

    def set_margins(self, **margins: float) -> None:
        """Set margins in points (or ``Length``): ``top``, ``right``, ``bottom``,
        ``left``, ``header``, ``footer``, ``gutter``."""
        unknown = set(margins) - set(DEFAULT_MARGINS)
        if unknown:
            raise TypeError(f"unknown margins {sorted(unknown)}; known: "
                            f"{', '.join(DEFAULT_MARGINS)}")
        values = {k: twips(v, k, PAGE_LIMIT) for k, v in margins.items()}
        merged = {**self._margins_twips(), **values}
        if merged["left"] + merged["right"] + merged["gutter"] >= self._page_twips("w"):
            raise ValueError("the left and right margins (and gutter) leave no room on the page; "
                             "margins are in points (72 = 1 inch)")
        if abs(merged["top"]) + abs(merged["bottom"]) >= self._page_twips("h"):
            raise ValueError("the top and bottom margins leave no room on the page; "
                             "margins are in points (72 = 1 inch)")
        self._attach()
        mar = self.sect_pr.first("w:pgMar")
        if mar is None:
            base: dict[str, Any] = dict(DEFAULT_MARGINS)
            mar = w.PgMar(**base)
            self.sect_pr.add(mar)
        for k, v in values.items():
            mar.set_attr(k, v)

    @property
    def text_width(self) -> float:
        """Width between the margins in points (the space a full-width table gets)."""
        m = self._margins_twips()
        return twips_to_points(self._page_twips("w") - m["left"] - m["right"] - m["gutter"])

    @property
    def start(self) -> str:
        """How the section starts: nextPage (default), continuous, evenPage, oddPage."""
        t = self.sect_pr.first("w:type")
        return t.get(f"{{{W}}}val", "nextPage") if t is not None else "nextPage"

    @start.setter
    def start(self, value: str) -> None:
        self._attach()
        self.sect_pr.add(w.Type(val=value), replace=True)

    @property
    def different_first_page(self) -> bool:
        """Whether the section's first page has its own header and footer (assignable)."""
        t = self.sect_pr.first("w:titlePg")
        return t is not None and t.get(f"{{{W}}}val", "1") not in ("0", "false", "off")

    @different_first_page.setter
    def different_first_page(self, value: bool) -> None:
        self._attach()
        if value:
            self.sect_pr.add(w.TitlePg(), replace=True)
        elif self.sect_pr.first("w:titlePg") is not None:
            self.sect_pr.remove(self.sect_pr.first("w:titlePg"))

    # -- headers and footers --------------------------------------------------------
    def _ref(self, what: str, kind: str) -> OoxElement | None:
        for ref in self.sect_pr.iterchildren(f"{{{W}}}{what}Reference"):
            if ref.get(f"{{{W}}}type", "default") == kind:
                return ref
        return None

    def _hdr_ftr(self, what: str, kind: str, create: bool) -> OoxElement | None:
        if kind not in KINDS:
            raise ValueError(f"kind must be one of {KINDS}")
        ref = self._ref(what, kind)
        part_type = "HeaderPart" if what == "header" else "FooterPart"
        if ref is not None:
            rel = self.doc.main_part.rels.get(ref.get(f"{{{R}}}id") or "")
            if rel is not None:
                part = self.doc.package.parts.get(self.doc.main_part.resolve(rel.target))
                if part is not None:
                    shared = sum(1 for s in self.doc.sections for r in
                                 s.sect_pr.iterchildren(f"{{{W}}}{what}Reference")
                                 if r.get(f"{{{R}}}id") == ref.get(f"{{{R}}}id"))
                    if not (create and shared > 1):
                        return part.element
                    # the part is shared with another section: give this one its own
        if not create:
            return None
        self._attach()
        style = "Header" if what == "header" else "Footer"
        self.doc.ensure_style(style, text_width=self.text_width)
        root_factory = w.Hdr if what == "header" else w.Ftr
        part, rid = self.doc.add_part(part_type, element=root_factory(
            w.P(w.PPr(w.PStyle(val=style)))))
        if ref is not None:
            self.sect_pr.remove(ref)
        factory = w.HeaderReference if what == "header" else w.FooterReference
        self.sect_pr.add(factory(type=kind, r_id=rid))
        if kind == "first":
            self.different_first_page = True
        elif kind == "even":
            self.doc.even_and_odd_headers = True
        return part.element

    def header(self, kind: str = "default") -> OoxElement | None:
        """This section's own header of ``kind`` ("default", "first", "even") as
        a ``w:hdr`` element, or None if it has none (``shown()`` follows "Link to
        Previous"; ``ensure_header()`` creates one)."""
        return self._hdr_ftr("header", kind, create=False)

    def footer(self, kind: str = "default") -> OoxElement | None:
        """This section's own footer of ``kind`` as a ``w:ftr`` element, or None."""
        return self._hdr_ftr("footer", kind, create=False)

    def ensure_header(self, kind: str = "default") -> OoxElement:
        """This section's own header of ``kind``, created if it has none, or split
        off if it shares one with another section.  "first" also turns on a
        different first page, "even" different even and odd pages."""
        root = self._hdr_ftr("header", kind, create=True)
        assert root is not None
        return root

    def ensure_footer(self, kind: str = "default") -> OoxElement:
        """This section's own footer of ``kind``, created (or split off) if needed."""
        root = self._hdr_ftr("footer", kind, create=True)
        assert root is not None
        return root

    def linked_to_previous(self, what: str = "header", kind: str = "default") -> bool:
        """True if this section has no own header/footer of that kind and
        shows the previous section's."""
        return self._ref(what, kind) is None

    def shown(self, what: str = "header", kind: str = "default") -> OoxElement | None:
        """The header/footer this section actually shows, following "Link to
        Previous" back through earlier sections; None if there is none.

        None means Word shows nothing there.  In particular, with a different
        first page (``kind="first"``) and no first-page header in this or an
        earlier section, the first page has no header: not the default one."""
        secs = self.doc.sections
        idx = next((i for i, s in enumerate(secs) if s.sect_pr is self.sect_pr), None)
        if idx is None:
            raise OoxError("this section is no longer part of the document")
        for s in reversed(secs[:idx + 1]):
            if s._ref(what, kind) is not None:
                return s._hdr_ftr(what, kind, create=False)
        return None

    def _set_text(self, root: OoxElement, text: str, align: str, style: str) -> OoxElement:
        from .fields import fill_paragraph

        for child in list(root):
            root.remove(child)
        ppr = w.PPr(w.PStyle(val=style))
        if align != "left":
            ppr.add(w.Jc(val=_ALIGN[align]))
        p = w.P(ppr)
        fill_paragraph(p, text)
        root.add(p)
        return p

    def header_text(self, text: str, kind: str = "default", align: str = "left") -> OoxElement:
        """Replace the header with one paragraph of ``text``; ``{PAGE}``,
        ``{NUMPAGES}``, ``{SECTIONPAGES}`` become fields.  Returns the paragraph."""
        root = self.ensure_header(kind)
        assert root is not None
        return self._set_text(root, text, align, "Header")

    def footer_text(self, text: str, kind: str = "default", align: str = "left") -> OoxElement:
        """Set the footer of ``kind`` to ``text``; ``{PAGE}``, ``{NUMPAGES}`` and
        ``{SECTIONPAGES}`` become fields, ``\\t`` aligns to the centre and right tab stops."""
        root = self.ensure_footer(kind)
        assert root is not None
        return self._set_text(root, text, align, "Footer")


def sections(doc: WordprocessingDocument) -> list[Section]:
    """All sections in order: paragraph-level sectPr (including paragraphs
    inside block-level content controls and custom XML, where LibreOffice puts
    the break after a table of contents), then the body's own."""
    from .walk import iter_blocks

    out = []
    for p in iter_blocks(doc.body, f"{{{W}}}p"):
        ppr = p.first("w:pPr")
        sp = ppr.first("w:sectPr") if ppr is not None else None
        if sp is not None:
            out.append(Section(doc, sp))
    final = doc.body.first("w:sectPr")
    if final is None:  # Word's defaults, added to the body only when something is set
        # an empty w:sectPr: every value reads as Word's default
        final = doc.__dict__.get("_default_sect_pr")
        if final is None:
            final = doc.__dict__["_default_sect_pr"] = w.SectPr()
    out.append(Section(doc, final))
    return out


def add_section(doc: WordprocessingDocument, start: str = "nextPage", *,
                landscape: bool | None = None, page: str | None = None) -> Section:
    """End the current section here and start a new one; returns the new one."""
    current = sections(doc)[-1].sect_pr
    fresh = copy.deepcopy(current)
    # The section that ends keeps its own w:sectPr element, moved into the
    # closing paragraph, so Section objects taken earlier still describe it;
    # the body gets a copy for the new section.
    ppr = w.PPr()
    doc.body.add(w.P(ppr))
    ppr.add(current)
    doc.body.add(fresh)
    new = Section(doc, fresh)
    # The new section links to the previous one's headers/footers ("Link to
    # Previous") and has no special first page: it would otherwise repeat a
    # cover-page header on its own first page.
    for ref in list(fresh.iterchildren(f"{{{W}}}headerReference", f"{{{W}}}footerReference")):
        fresh.remove(ref)
    new.different_first_page = False
    new.start = start
    if page is not None:
        from .word_defaults import PAGE_SIZES

        wd, ht = PAGE_SIZES[page]
        sz = new._pg_sz(create=True)
        assert sz is not None
        sz.set_attr("w", wd)
        sz.set_attr("h", ht)
        sz.set_attr("orient", None)
    if landscape is not None:
        new.orientation = "landscape" if landscape else "portrait"
    return new
