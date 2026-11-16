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
from .grammar import URI
from .ns import w

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
R = URI["r"]
KINDS = ("default", "first", "even")
_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}


class Section:
    def __init__(self, doc: WordprocessingDocument, sect_pr: OoxElement) -> None:
        self.doc, self.sect_pr = doc, sect_pr

    def __repr__(self) -> str:
        return f"<Section {self.orientation} {self.page_width}x{self.page_height} twips>"

    # -- page setup ---------------------------------------------------------------
    def _pg_sz(self) -> OoxElement:
        sz = self.sect_pr.first("w:pgSz")
        if sz is None:
            sz = w.PgSz(w=12240, h=15840)  # Word's default when absent: Letter
            self.sect_pr.add(sz)
        return sz

    @property
    def page_width(self) -> int:
        """Page width in twips."""
        return int(self._pg_sz().get(f"{{{W}}}w", "12240"))

    @property
    def page_height(self) -> int:
        """Page height in twips."""
        return int(self._pg_sz().get(f"{{{W}}}h", "15840"))

    @property
    def orientation(self) -> str:
        """"portrait" or "landscape" (assignable; assigning swaps width and height)."""
        sz = self._pg_sz()
        orient = sz.get(f"{{{W}}}orient")
        if orient:
            return orient
        return "landscape" if self.page_width > self.page_height else "portrait"

    @orientation.setter
    def orientation(self, value: str) -> None:
        if value not in ("portrait", "landscape"):
            raise ValueError("orientation is 'portrait' or 'landscape'")
        if value == self.orientation:
            return
        sz = self._pg_sz()
        wd, ht = self.page_width, self.page_height
        sz.set_attr("w", ht)
        sz.set_attr("h", wd)
        sz.set_attr("orient", "landscape" if value == "landscape" else None)

    @property
    def margins(self) -> dict[str, int]:
        """Twips: top, right, bottom, left, header, footer, gutter."""
        mar = self.sect_pr.first("w:pgMar")
        if mar is None:
            return {}
        return {k.rpartition("}")[2]: int(v) for k, v in mar.attrib.items() if v.lstrip("-").isdigit()}

    def set_margins(self, **twips: int) -> None:
        """Set margins in twips (or EMU from ``oox.units``): ``top``, ``right``, ``bottom``,
        ``left``, ``header``, ``footer``, ``gutter``."""
        mar = self.sect_pr.first("w:pgMar")
        if mar is None:
            base: dict[str, Any] = dict(top=1440, right=1440, bottom=1440, left=1440, header=708, footer=708, gutter=0)
            mar = w.PgMar(**base)
            self.sect_pr.add(mar)
        for k, v in twips.items():
            # EMU from oox.units (Cm(2) = 720000) as Word tables take them; no
            # real margin is 100,000 twips (69 inches)
            mar.set_attr(k, round(v / 635) if abs(v) > 100000 else int(v))

    @property
    def text_width(self) -> int:
        """Width between the margins in twips (the space a full-width table gets)."""
        m = self.margins
        return self.page_width - m.get("left", 1440) - m.get("right", 1440) - m.get("gutter", 0)

    @property
    def start(self) -> str:
        """How the section starts: nextPage (default), continuous, evenPage, oddPage."""
        t = self.sect_pr.first("w:type")
        return t.get(f"{{{W}}}val", "nextPage") if t is not None else "nextPage"

    @start.setter
    def start(self, value: str) -> None:
        self.sect_pr.add(w.Type(val=value), replace=True)

    @property
    def different_first_page(self) -> bool:
        """Whether the section's first page has its own header and footer (assignable)."""
        t = self.sect_pr.first("w:titlePg")
        return t is not None and t.get(f"{{{W}}}val", "1") not in ("0", "false", "off")

    @different_first_page.setter
    def different_first_page(self, value: bool) -> None:
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
            rel = self.doc.main_part.rels.items.get(ref.get(f"{{{R}}}id") or "")
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

    def header(self, kind: str = "default", create: bool = True) -> OoxElement | None:
        """The ``w:hdr`` of this section (creating its own if needed)."""
        return self._hdr_ftr("header", kind, create)

    def footer(self, kind: str = "default", create: bool = True) -> OoxElement | None:
        """The section's footer of ``kind`` ("default", "first", "even") as a ``w:ftr``
        element, created if missing (``create=False``: None instead)."""
        return self._hdr_ftr("footer", kind, create)

    def linked_to_previous(self, what: str = "header", kind: str = "default") -> bool:
        """True if this section has no own header/footer of that kind and
        shows the previous section's."""
        return self._ref(what, kind) is None

    def shown(self, what: str = "header", kind: str = "default") -> OoxElement | None:
        """The header/footer this section actually shows, following "Link to
        Previous" back through earlier sections; None if there is none."""
        secs = self.doc.sections
        idx = next(i for i, s in enumerate(secs) if s.sect_pr is self.sect_pr)
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
        root = self.header(kind)
        assert root is not None
        return self._set_text(root, text, align, "Header")

    def footer_text(self, text: str, kind: str = "default", align: str = "left") -> OoxElement:
        """Set the footer of ``kind`` to ``text``; ``{PAGE}``, ``{NUMPAGES}`` and
        ``{SECTIONPAGES}`` become fields, ``\\t`` aligns to the centre and right tab stops."""
        root = self.footer(kind)
        assert root is not None
        return self._set_text(root, text, align, "Footer")


def sections(doc: WordprocessingDocument) -> list[Section]:
    """All sections in order: paragraph-level sectPr, then the body's own."""
    out = []
    for p in doc.body.iterchildren(f"{{{W}}}p"):
        ppr = p.first("w:pPr")
        sp = ppr.first("w:sectPr") if ppr is not None else None
        if sp is not None:
            out.append(Section(doc, sp))
    final = doc.body.first("w:sectPr")
    if final is None:
        from .word_defaults import section

        final = section()
        doc.body.add(final)
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
        sz = new._pg_sz()
        sz.set_attr("w", wd)
        sz.set_attr("h", ht)
        sz.set_attr("orient", None)
    if landscape is not None:
        new.orientation = "landscape" if landscape else "portrait"
    return new
