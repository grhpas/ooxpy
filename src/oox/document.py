"""Format-level document classes -- the ``OpenDocumentText()`` equivalents.

``OoxDocument`` holds what all three formats share: the package, the main
part, part creation from the SDK's part metadata, media (pictures) and odfpy's
``save``/``addPicture``/``getElementsByType``.  Format classes add their
skeleton (``new``) and conveniences:

* ``WordprocessingDocument`` (here)        .docx .docm .dotx .dotm
* ``SpreadsheetDocument`` (spreadsheet.py)  .xlsx .xlsm .xltx .xltm
* ``PresentationDocument`` (presentation.py) .pptx .pptm .potx .ppsx
"""

from __future__ import annotations

import hashlib
import os
import posixpath
import sys
from collections.abc import Callable
from typing import IO, ClassVar

from . import drawing, images
from ._generated.parts import PARTS
from .element import OoxElement
from .errors import PackageError
from .opc import Package, Part

RT_OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
RT_IMAGE = PARTS["ImagePart"]["relationship_type"]


class OoxDocument:
    """Base class.  Subclasses set ``MAIN_CONTENT_TYPES`` and implement ``new``."""

    MAIN_CONTENT_TYPES: ClassVar[dict[str, str]] = {}

    def __init__(self, package: Package) -> None:
        self.package = package
        part = package.related(None, RT_OFFICE_DOCUMENT)
        if part is None:
            raise PackageError("no officeDocument relationship")
        if self.MAIN_CONTENT_TYPES and part.content_type not in self.MAIN_CONTENT_TYPES.values():
            raise PackageError(f"{type(self).__name__} cannot open a {part.content_type} package")
        self.main_part: Part = part
        self.stripped: dict[str, list[str]] = {}   # what save(target=, strip=True) removed
        self.suffix = next((k for k, v in self.MAIN_CONTENT_TYPES.items()
                            if v == part.content_type), "")
        self._media: dict[str, Part] | None = None  # sha1 -> image part

    # -- io -------------------------------------------------------------------
    @classmethod
    def open(cls, file: str | os.PathLike | IO[bytes]):
        """Open ``file`` (a path or a binary file object) as this class;
        ``oox.open_document`` picks the class from the content."""
        return cls(Package.open(file))

    def save(self, outputfile: str | os.PathLike | IO[bytes], addsuffix: bool = False, *,
             target: str | None = None, strip: bool = False) -> None:
        """Save the package.  odfpy-compatible: ``addsuffix=True`` appends the
        format's suffix (``.docx``, ``.xlsx``, ``.pptx``); ``"-"`` writes to stdout.

        ``target="Office2007"`` (Office2010, ...) refuses to save markup that
        version can't read -- e.g. Office 2010 attributes in a file from
        LibreOffice, which Excel 2007 refuses to open -- raising
        ``TargetError`` with every finding; ``strip=True`` removes that markup
        instead (``self.stripped`` lists what went).
        """
        self._prepare_save()
        if target is not None:
            self._check_target(target, strip)
        if outputfile == "-":
            self.package.save(sys.stdout.buffer)
            return
        if addsuffix and isinstance(outputfile, (str, os.PathLike)):
            outputfile = os.fspath(outputfile) + self.suffix
        self._sync_conformance_attribute()
        self.package.save(outputfile)

    def _check_target(self, target: str, strip: bool) -> None:
        from .errors import TargetError

        findings: dict[str, list[str]] = {}
        for pn in sorted(self.package.parts):
            part = self.package.parts[pn]
            if not part.is_xml:
                continue
            try:
                root = part.element
            except Exception:
                continue  # unparsable parts are written back as they are
            found = root.strip_for(target) if strip else [m for _, _, m in root.target_findings(target)]
            if found:
                findings[pn] = found
        self.stripped = findings if strip else {}
        if findings and not strip:
            count = sum(len(v) for v in findings.values())
            first = next(iter(findings.items()))
            raise TargetError(f"{count} item(s) {target} can't read, e.g. {first[0]}: {first[1][0]}"
                              "; save(..., strip=True) removes them", findings)

    def _prepare_save(self) -> None:
        """Format-specific fix-ups that the format requires (see subclasses)."""

    @property
    def conformance(self) -> str:
        """"transitional" or "strict"; set it to convert on save."""
        return self.package.conformance

    @conformance.setter
    def conformance(self, value: str) -> None:
        if value not in ("transitional", "strict"):
            raise ValueError("conformance must be 'transitional' or 'strict'")
        self.package.conformance = value

    def _sync_conformance_attribute(self) -> None:
        # Office marks Strict documents on the main root (w:document/@w:conformance,
        # x:workbook/@conformance, p:presentation/@conformance).
        root = self.main_part.element
        info = root.info
        spec = info.attr("conformance") if info else None
        if spec is None:
            return
        current = root.get(spec.clark)
        if self.conformance == "strict" and current != "strict":
            root.set(spec.clark, "strict")
        elif self.conformance == "transitional" and current == "strict":
            del root.attrib[spec.clark]

    # -- parts ----------------------------------------------------------------
    def add_part(self, part_type: str, source: Part | None = None, *,
                 element: OoxElement | None = None, blob: bytes | None = None,
                 content_type: str | None = None, ext: str = "xml",
                 folder: str | None = None) -> tuple[Part, str]:
        """Create a part described by the SDK part table and relate it.

        ``part_type`` is an SDK part name ("WorksheetPart", "SlidePart", ...);
        its content type, relationship type and folder come from the table.
        ``folder`` (absolute, e.g. "/xl/charts") overrides the folder, to put
        parts where Office does.  Returns ``(part, relationship id)``.
        """
        meta = PARTS[part_type]
        src = source or self.main_part
        src_dir = posixpath.dirname(src.partname)
        top = "/" + src.partname.strip("/").split("/")[0]
        if folder is not None:
            folder = folder.rstrip("/")
        else:
            folder = posixpath.normpath(posixpath.join(src_dir, meta["path"] or "."))
        if not (folder + "/").startswith(top + "/"):
            # SDK paths like "../media" assume a nested parent (xl/drawings);
            # from a top-level part (word/document.xml) stay inside its folder.
            folder = posixpath.join(src_dir, posixpath.basename(meta["path"]))
        numbered = meta["multiple"] or src is not self.main_part or meta["target"] == "theme"
        n = 1
        while True:
            stem = f"{meta['target']}{n}" if numbered else meta["target"]
            partname = f"{folder}/{stem}.{ext}"
            if partname not in self.package.parts:
                break
            numbered, n = True, n + 1
        ctype = content_type or meta["content_type"]
        if ctype is None:
            raise PackageError(f"{part_type}: content type must be given")
        part = self.package.add_part(partname, ctype, element=element, blob=blob)
        return part, self.package.relate(src, part, meta["relationship_type"])

    def relate(self, source: Part, target: Part, part_type: str) -> str:
        """Add a relationship of ``part_type``'s kind between existing parts."""
        return self.package.relate(source, target, PARTS[part_type]["relationship_type"])

    def related_parts(self, source: Part, part_type: str) -> list[Part]:
        """Parts that ``source`` relates to with ``part_type``'s relationship (an SDK part
        name such as "ImagePart" or "DocumentSettingsPart"), in relationship order."""
        reltype = PARTS[part_type]["relationship_type"]
        return [self.package.parts[source.resolve(r.target)]
                for r in source.rels.of_type(reltype)
                if not r.external and source.resolve(r.target) in self.package.parts]

    def part_of(self, element: OoxElement) -> Part:
        """The part whose XML tree contains ``element``."""
        root = element.getroottree().getroot()
        for part in self.package.parts.values():
            if part._element is root:
                return part
        raise PackageError("element is not attached to a part of this document")

    # -- media ----------------------------------------------------------------
    def add_image(self, image, source: Part | None = None) -> tuple[str, images.ImageInfo]:
        """Store an image (path, bytes or file object) and relate it from ``source``.

        Identical images are stored once; relating the same image twice from
        the same part reuses the relationship.  Returns ``(rId, ImageInfo)``.
        """
        blob = images.read_blob(image)
        info = images.inspect(blob)
        src = source or self.main_part
        if self._media is None:
            self._media = {hashlib.sha1(p.to_bytes()).hexdigest(): p
                           for p in self.package.parts.values()
                           if p.content_type.startswith("image/")}
        digest = hashlib.sha1(blob).hexdigest()
        part = self._media.get(digest)
        if part is None:
            part, rid = self.add_part("ImagePart", src, blob=blob,
                                      content_type=info.content_type, ext=info.ext)
            self.package.defaults.setdefault(info.ext, info.content_type)
            self._media[digest] = part
            return rid, info
        for r in src.rels.of_type(RT_IMAGE):
            if src.resolve(r.target) == part.partname:
                return r.rid, info
        return self.package.relate(src, part, RT_IMAGE), info

    def add_picture_image(self, image, fallback=None, source: Part | None = None
                          ) -> tuple[str, images.ImageInfo, str | None]:
        """Store a picture's image(s): ``(rid, info, svg_rid)``.

        An SVG needs a raster ``fallback`` (PNG recommended): ``rid`` is the
        fallback, ``svg_rid`` the SVG, and the size comes from the SVG.
        """
        blob = images.read_blob(image)
        info = images.inspect(blob)
        if not info.needs_fallback:
            if fallback is not None:
                raise ValueError("fallback= is only for SVG pictures")
            rid, info = self.add_image(blob, source)
            return rid, info, None
        if fallback is None:
            raise images.UnsupportedImage(
                "an SVG needs a raster fallback for Office before 2016: "
                "add_picture(svg, fallback=png)")
        fb_rid, fb_info = self.add_image(fallback, source)
        if fb_info.needs_fallback:
            raise ValueError("the fallback must be a raster image (PNG recommended)")
        svg_rid, _ = self.add_image(blob, source)
        return fb_rid, info, svg_rid

    # odfpy-compatible: returns the relationship id to use in a:blip r:embed
    def addPicture(self, filename: str, mediatype: str | None = None,
                   content: bytes | None = None) -> str:
        """odfpy's ``addPicture``: store an image (from ``filename``, or ``content`` bytes)
        in the package and return its relationship id; ``add_picture`` places a picture."""
        return self.add_image(content if content is not None else filename)[0]

    def addPictureFromFile(self, filename: str, mediatype: str | None = None) -> str:
        """odfpy's ``addPictureFromFile``: store the image file and return its relationship
        id."""
        return self.add_image(filename)[0]

    def addPictureFromString(self, content: bytes, mediatype: str | None = None) -> str:
        """odfpy's ``addPictureFromString``: store image bytes and return their relationship
        id."""
        return self.add_image(content)[0]

    @property
    def extended_properties(self):
        """``docProps/app.xml`` (company, manager, template, ...): ``ExtendedProperties``."""
        from .properties import ExtendedProperties

        return ExtendedProperties(self)

    @property
    def custom_properties(self):
        """``docProps/custom.xml`` as a mapping: ``CustomProperties``."""
        from .properties import CustomProperties

        return CustomProperties(self)

    @property
    def properties(self):
        """Document properties (title, creator, created, ...): ``CoreProperties``."""
        from .properties import CoreProperties

        return CoreProperties(self)

    # -- checking -------------------------------------------------------------
    def validate(self, target: str | None = None, level: str = "basic") -> dict[str, list[str]]:
        """Validate every XML part; ``{partname: [problems]}`` for parts with any.

        ``target="Office2007"`` adds markup that Office version won't understand;
        ``level="full"`` checks whole content models (choices, counts).
        Parts that fail to parse are reported, not raised.
        """
        report: dict[str, list[str]] = {}
        for pn in sorted(self.package.parts):
            part = self.package.parts[pn]
            if not part.is_xml:
                continue
            try:
                problems = part.element.validate(target=target, level=level)
            except Exception as e:
                problems = [f"cannot parse: {e}"]
            if problems:
                report[pn] = problems
        return report

    # -- queries --------------------------------------------------------------
    def getElementsByType(self, element: Callable) -> list[OoxElement]:
        """odfpy-compatible: search the main part and every part already loaded."""
        found: list[OoxElement] = []
        for part in self.package.parts.values():
            if part is self.main_part or part._element is not None:
                found += part.element.getElementsByType(element)
        return found


# The SDK models main-part content types in code (they depend on the suffix),
# so they aren't in the generated part table.
CT_WML_MAIN = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    ".docm": "application/vnd.ms-word.document.macroEnabled.main+xml",
    ".dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    ".dotm": "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
}


class WordprocessingDocument(OoxDocument):
    """A .docx package.

    >>> from oox import WordprocessingDocument
    >>> from oox.ns import w
    >>> doc = WordprocessingDocument.new()
    >>> doc.body.add(w.P(w.R(w.T("Hello"))))  # doctest: +ELLIPSIS
    <OoxElement w:body ...>
    """

    MAIN_CONTENT_TYPES = CT_WML_MAIN

    @classmethod
    def new(cls, kind: str = ".docx", *, minimal: bool = False, page: str = "A4",
            landscape: bool = False, lang: str = "en-US") -> WordprocessingDocument:
        """A new document, set up the way Word sets one up.

        By default it gets styles (Normal, Title, Heading 1-3, List
        Paragraph, Hyperlink, Table Grid), settings, a font table, the
        Office theme and a section with ``page`` size ("A4" or "Letter")
        and 1-inch margins.  ``minimal=True`` gives just ``word/document.xml``
        (smallest valid package; Word fills in its own defaults).
        """
        from . import word_defaults
        from .ns import w

        pkg = Package()
        body = w.Body()
        part = pkg.add_part("/word/document.xml", CT_WML_MAIN[kind], element=w.Document(body))
        pkg.relate(None, part, RT_OFFICE_DOCUMENT)
        doc = cls(pkg)
        if not minimal:
            doc.add_part("StyleDefinitionsPart", element=word_defaults.styles(lang))
            doc.add_part("DocumentSettingsPart", element=word_defaults.settings(lang))
            doc.add_part("FontTablePart", element=word_defaults.fonts())
            doc.add_part("ThemePart", element=drawing.default_theme())
            body.add(word_defaults.section(page, landscape))
        return doc

    @property
    def document(self) -> OoxElement:
        """The root ``w:document`` element of the main part."""
        return self.main_part.element

    @property
    def body(self) -> OoxElement:
        """The ``w:body`` element, where paragraphs and tables go."""
        body = self.document.first("w:body")
        if body is None:
            raise PackageError("w:document has no w:body")
        return body

    # odfpy-flavoured alias: doc.text.addElement(...) in odfpy
    text = body

    @property
    def styles(self) -> OoxElement:
        """Root ``<w:styles>`` of the styles part, created on first use.

        odfpy: ``doc.styles``.  OOXML has no "automatic styles"; direct
        formatting lives in ``w:pPr``/``w:rPr`` instead.
        """
        found = self.related_parts(self.main_part, "StyleDefinitionsPart")
        if found:
            return found[0].element
        from .ns import w

        part, _ = self.add_part("StyleDefinitionsPart", element=w.Styles())
        return part.element

    # -- settings, styles on demand ---------------------------------------------
    @property
    def settings(self) -> OoxElement:
        """Root ``<w:settings>``, created on first use."""
        found = self.related_parts(self.main_part, "DocumentSettingsPart")
        if found:
            return found[0].element
        from .ns import w

        part, _ = self.add_part("DocumentSettingsPart", element=w.Settings())
        return part.element

    @property
    def even_and_odd_headers(self) -> bool:
        """Whether even and odd pages have different headers and footers (a document-wide
        setting)."""
        el = self.settings.first("w:evenAndOddHeaders")
        return el is not None

    @even_and_odd_headers.setter
    def even_and_odd_headers(self, value: bool) -> None:
        from .ns import w

        el = self.settings.first("w:evenAndOddHeaders")
        if value and el is None:
            self.settings.add(w.EvenAndOddHeaders())
        elif not value and el is not None:
            self.settings.remove(el)

    def ensure_style(self, style_id: str, text_width: int = 9026) -> str:
        """Make sure a style exists, adding Word's definition for the known
        feature styles (Header, FootnoteText, TOC1, Hyperlink ...)."""
        try:
            return self.style_id(style_id)
        except KeyError:
            from .word_defaults import extra_style

            definition = extra_style(style_id, text_width)
            if definition is None:
                raise
            self.styles.add(definition)
            return style_id

    def outline_level(self, paragraph: OoxElement) -> int:
        """1-9 for headings (direct w:outlineLvl or via the style chain), 10 for body text."""
        from .grammar import URI

        w_ns = URI["w"]
        styles: dict = {}
        found = self.related_parts(self.main_part, "StyleDefinitionsPart")
        if found:
            styles = {st.get(f"{{{w_ns}}}styleId"): st
                      for st in found[0].element.iterchildren(f"{{{w_ns}}}style")}
        el: OoxElement | None = paragraph
        for _ in range(12):
            if el is None:
                break
            ppr = el.first("w:pPr")
            lvl = ppr.first("w:outlineLvl") if ppr is not None else None
            if lvl is not None:
                return int(lvl.get(f"{{{w_ns}}}val", "9")) + 1
            ref = (ppr.first("w:pStyle") if ppr is not None else None) if el is paragraph \
                else el.first("w:basedOn")
            el = styles.get(ref.get(f"{{{w_ns}}}val")) if ref is not None else None
        return 10

    # -- sections, headers, footers --------------------------------------------
    @property
    def sections(self):
        """``Section`` objects in order (page setup, headers, footers)."""
        from .sections import sections

        return sections(self)

    def add_section(self, start: str = "nextPage", *, landscape: bool | None = None,
                    page: str | None = None):
        """End the current section and start a new one (default: on a new page)."""
        from .sections import add_section

        return add_section(self, start, landscape=landscape, page=page)

    # -- links, notes, comments -------------------------------------------------
    def add_hyperlink(self, paragraph: OoxElement, text: str, *, url: str | None = None,
                      anchor: str | None = None, tooltip: str | None = None) -> OoxElement:
        """A hyperlink with ``text`` at the end of ``paragraph``: to ``url`` (an external
        relationship) or to the bookmark ``anchor``.  Returns the ``w:hyperlink``."""
        from .notes import add_hyperlink

        return add_hyperlink(self, paragraph, text, url=url, anchor=anchor, tooltip=tooltip)

    def add_bookmark(self, name: str, paragraph: OoxElement) -> OoxElement:
        """Bookmark the whole content of ``paragraph`` as ``name`` (the target of
        ``add_hyperlink(..., anchor=name)`` and cross-references); returns the bookmark
        start."""
        from .notes import add_bookmark

        return add_bookmark(self, name, paragraph)

    def hyperlinks(self) -> list[tuple[str, str]]:
        """``(text, target)`` for each hyperlink in the body; the target is the URL or
        ``#bookmark``."""
        from .notes import hyperlinks

        return hyperlinks(self)

    def bookmarks(self) -> list[str]:
        """Bookmark names, without Word's hidden ones (``_GoBack``, ``_Toc...``)."""
        from .notes import bookmarks

        return bookmarks(self)

    def add_footnote(self, paragraph: OoxElement, text: str) -> OoxElement:
        """A footnote with ``text``, referenced at the end of ``paragraph``; the footnotes
        part is created on first use.  Returns the reference run."""
        from .notes import add_footnote

        return add_footnote(self, paragraph, text)

    def add_endnote(self, paragraph: OoxElement, text: str) -> OoxElement:
        """An endnote with ``text``, referenced at the end of ``paragraph``; the endnotes
        part is created on first use.  Returns the reference run."""
        from .notes import add_endnote

        return add_endnote(self, paragraph, text)

    def footnotes(self) -> list[tuple[int, str]]:
        """``(id, text)`` for each footnote (Word's separator notes excluded)."""
        from .notes import notes

        return notes(self, "footnote")

    def endnotes(self) -> list[tuple[int, str]]:
        """``(id, text)`` for each endnote (Word's separator notes excluded)."""
        from .notes import notes

        return notes(self, "endnote")

    def add_comment(self, paragraph: OoxElement, text: str, *, author: str = "oox",
                    initials: str | None = None, date=None) -> OoxElement:
        """Comment on the whole content of ``paragraph``; ``date`` defaults to now.  Returns
        the comment's ``w:comment``."""
        from .notes import add_comment

        return add_comment(self, paragraph, text, author=author, initials=initials, date=date)

    def comments(self) -> list[dict]:
        """``{"id", "author", "text", "anchor"}`` for each comment; ``anchor`` is the
        commented text in the body."""
        from .notes import comments

        return comments(self)

    # -- fields, templates ------------------------------------------------------
    def add_toc(self, levels: tuple[int, int] = (1, 3), title: str | None = "Contents", *,
                before: OoxElement | None = None, update_on_open: bool = False) -> list[OoxElement]:
        """Insert a table of contents (a TOC field pre-filled with the current
        headings; Word adds page numbers on update).  ``before`` places it
        before that body element; default: at the end."""
        from .fields import toc

        paras = toc(self, levels, title, update_on_open)
        for p in paras:
            if before is not None:
                before.addprevious(p)
            else:
                self.body.add(p)
        return paras

    def replace_text(self, mapping: dict[str, str] | str, new: str | None = None) -> dict[str, int]:
        """Replace text across runs in body, headers, footers, notes and comments:
        ``replace_text({"{{name}}": "Ada"})`` or ``replace_text("old", "new")``."""
        from .template import replace_text

        if isinstance(mapping, str):
            if new is None:
                raise TypeError("replace_text(old, new) needs new")
            mapping = {mapping: new}
        return replace_text(self, mapping)

    # -- formatting, styles --------------------------------------------------
    def add_run(self, paragraph: OoxElement, text: str, *, style: str | None = None,
                **fmt) -> OoxElement:
        """Append a run with formatting (bold, italic, underline, strike, size,
        font, color, highlight, superscript, subscript, caps, small_caps);
        ``style`` is a character style (name or id)."""
        from .formatting import run

        r = run(text, style=self.style_id(style) if style else None, **fmt)
        paragraph.add(r)
        return r

    def format_paragraph(self, paragraph: OoxElement, **fmt) -> OoxElement:
        """Alignment, spacing (pt), indents (EMU or twips), line spacing, keeps;
        see ``oox.formatting.format_paragraph``."""
        from .formatting import format_paragraph

        return format_paragraph(paragraph, **fmt)

    def add_style(self, name: str, *, kind: str = "paragraph", base: str | None = None,
                  next_style: str | None = None, quick: bool = True, **fmt) -> str:
        """A custom paragraph or character style; returns its id."""
        from .formatting import add_style

        return add_style(self, name, kind=kind, base=base, next_style=next_style, quick=quick,
                         **fmt)

    def effective_format(self, paragraph: OoxElement, run: OoxElement | None = None) -> dict:
        """The formatting Word applies to ``run`` (or the paragraph mark),
        resolved through defaults, styles and direct formatting."""
        from .formatting import effective_format

        return effective_format(self, paragraph, run)

    # -- captions, cross-references, field results -----------------------------
    def add_caption(self, text: str, label: str = "Figure", *, before: OoxElement | None = None,
                    after: OoxElement | None = None, sep: str = ": "):
        """A numbered caption ("Figure 1: text"); returns a ``Caption`` to
        cross-reference.  See ``oox.fieldcalc``."""
        from .fieldcalc import add_caption

        return add_caption(self, text, label, before=before, after=after, sep=sep)

    def add_cross_reference(self, paragraph: OoxElement, target, *, what: str = "text",
                            hyperlink: bool = True) -> OoxElement:
        """A cross-reference at the end of ``paragraph`` to a caption (from ``add_caption``),
        a bookmark name or a paragraph such as a heading (which gets a hidden bookmark).
        ``what``: "text" ("Figure 1", the heading's text), "page" (estimated until Word
        updates fields) or "position" ("above"/"below")."""
        from .fieldcalc import add_cross_reference

        return add_cross_reference(self, paragraph, target, what=what, hyperlink=hyperlink)

    def add_field(self, paragraph: OoxElement, instr: str, result: str = "") -> OoxElement:
        """Any field (``"TITLE"``, ``'DOCPROPERTY "Company"'``, ``'DATE \\@ "d MMMM yyyy"'``);
        the result is computed now when oox can."""
        from .fieldcalc import add_field

        return add_field(self, paragraph, instr, result)

    def fields(self, all_stories: bool = False) -> list[dict]:
        """``{"name", "code", "result"}`` for the body's fields; ``all_stories=True`` adds
        headers, footers, notes and comments."""
        from .fieldcalc import fields

        return fields(self, all_stories)

    def update_fields(self, *, now=None) -> dict[str, int]:
        """Recompute field results that don't need page layout (SEQ, REF,
        document properties, dates ...); see ``oox.fieldcalc``."""
        from .fieldcalc import update_fields

        return update_fields(self, now=now)

    # -- content controls (forms) ---------------------------------------------
    def add_text_control(self, paragraph: OoxElement, *, tag: str, title: str | None = None,
                         text: str | None = None, placeholder: str = "Click here to enter text.",
                         rich: bool = False, multiline: bool = False) -> OoxElement:
        """A text (or rich text) content control; see ``oox.controls``."""
        from .controls import add_text_control

        return add_text_control(self, paragraph, tag=tag, title=title, text=text,
                                placeholder=placeholder, rich=rich, multiline=multiline)

    def add_dropdown(self, paragraph: OoxElement, *, tag: str, items: list, value=None,
                     title: str | None = None, combo: bool = False,
                     placeholder: str = "Choose an item.") -> OoxElement:
        """A drop-down content control (``combo=True``: a combo box, which also takes typed
        text).  ``items`` are display texts or ``(display, value)`` pairs."""
        from .controls import add_dropdown

        return add_dropdown(self, paragraph, tag=tag, items=items, value=value, title=title,
                            combo=combo, placeholder=placeholder)

    def add_date_control(self, paragraph: OoxElement, *, tag: str, value=None,
                         fmt: str = "d MMMM yyyy", title: str | None = None,
                         lang: str = "en-US",
                         placeholder: str = "Click here to enter a date.") -> OoxElement:
        """A date-picker content control at the end of ``paragraph``; ``fmt`` is Word's date
        picture (``dd/MM/yyyy`` ...), ``value`` a date or None for the placeholder."""
        from .controls import add_date_control

        return add_date_control(self, paragraph, tag=tag, value=value, fmt=fmt, title=title,
                                lang=lang, placeholder=placeholder)

    def add_checkbox(self, paragraph: OoxElement, *, tag: str, checked: bool = False,
                     title: str | None = None) -> OoxElement:
        """A check-box content control at the end of ``paragraph`` (Word 2010+; Word 2007
        shows the ☐/☒ character)."""
        from .controls import add_checkbox

        return add_checkbox(self, paragraph, tag=tag, checked=checked, title=title)

    def content_controls(self) -> list[dict]:
        """Every content control (body, headers, footers, notes) in order: ``{"tag", "title",
        "type", "value", "placeholder", ...}``.  ``value`` is the text (None while the
        placeholder shows), a bool for check boxes; drop-downs add ``items``, dates add
        ``date``."""
        from .controls import content_controls

        return content_controls(self)

    def fill_controls(self, values: dict, *, strict: bool = True) -> int:
        """Fill content controls by tag; see ``oox.controls.fill_controls``."""
        from .controls import fill_controls

        return fill_controls(self, values, strict=strict)

    # -- tables, floating pictures ---------------------------------------------
    def _prepare_save(self) -> None:
        if self.main_part._element is not None:  # parsed, so possibly changed
            from .tables import ensure_paragraph_after_table

            ensure_paragraph_after_table(self)

    def add_table(self, rows: list[list], *, widths: list[int] | None = None,
                  style: str | None = "Table Grid", header_rows: int = 0,
                  bold_header: bool = True, align: str | None = None,
                  before: OoxElement | None = None):
        """A table from rows of values; returns a ``Table`` (merge, widths,
        formatting, reading).  See ``oox.tables``."""
        from .tables import add_table

        return add_table(self, rows, widths=widths, style=style, header_rows=header_rows,
                         bold_header=bold_header, align=align, before=before)

    def tables(self):
        """``Table`` objects for the body's top-level tables."""
        from .tables import tables

        return tables(self)

    def add_floating_picture(self, image, **options) -> OoxElement:
        """A picture positioned on the page with text wrapping; see
        ``oox.floating.add_floating_picture`` for the options."""
        from .floating import add_floating_picture

        return add_floating_picture(self, image, **options)

    def pictures(self) -> list[dict]:
        """Inline and floating pictures in the body, in order."""
        from .floating import pictures

        return pictures(self)

    # -- text boxes ---------------------------------------------------------------
    def add_text_box(self, content="", **options) -> OoxElement:
        """A floating text box (DrawingML, with a VML copy for Word 2007);
        returns its ``w:r``.  See ``oox.textboxes.add_text_box``."""
        from .textboxes import add_text_box

        return add_text_box(self, content, **options)

    def text_boxes(self) -> list[dict]:
        """Text boxes in the body, each once (``oox.textboxes``)."""
        from .textboxes import text_boxes

        return text_boxes(self)

    def set_text_box_text(self, index: int, content) -> None:
        """Replace the content of text box ``index`` (in ``text_boxes()`` order) with text
        (``\\n`` starts a paragraph) or a list of paragraphs, in both copies Word keeps."""
        from .textboxes import set_text_box_text

        set_text_box_text(self, index, content)

    # -- protection ---------------------------------------------------------------
    def protect(self, edit: str = "readOnly", *, password: str | None = None,
                **options) -> None:
        """Restrict editing (readOnly, comments, trackedChanges, forms),
        optionally with a password; see ``oox.protection.protect``."""
        from .protection import protect

        protect(self, edit, password=password, **options)

    def protection(self) -> dict | None:
        """The editing restriction: ``{"edit", "enforced", "password", "algorithm",
        "formatting"}``, or None."""
        from .protection import protection

        return protection(self)

    def check_protection_password(self, password: str) -> bool:
        """Whether ``password`` lifts the editing restriction (True when there is no
        password).  Checks Word 2007's and Word 2010's hash attributes."""
        from .protection import check_password

        return check_password(self, password)

    def unprotect(self) -> None:
        """Remove the editing restriction (a password is not needed: protection is not
        encryption)."""
        from .protection import unprotect

        unprotect(self)

    # -- tracked changes ----------------------------------------------------------
    @property
    def track_changes(self) -> bool:
        """Word's "Track Changes" switch (``w:trackRevisions`` in settings)."""
        return self.settings.first("w:trackRevisions") is not None

    @track_changes.setter
    def track_changes(self, value: bool) -> None:
        from .ns import w

        el = self.settings.first("w:trackRevisions")
        if value and el is None:
            self.settings.add(w.TrackRevisions())
        elif not value and el is not None:
            self.settings.remove(el)

    def suggest_insert(self, paragraph: OoxElement, text: str, *, author: str, date=None,
                       after: str | None = None) -> OoxElement:
        """Insert ``text`` as a tracked change by ``author``: at the end of ``paragraph``, or
        right after the first occurrence of ``after``."""
        from .revisions import suggest_insert

        return suggest_insert(self, paragraph, text, author=author, date=date, after=after)

    def suggest_delete(self, paragraph: OoxElement, text: str, *, author: str, date=None,
                       occurrence: int = 1) -> OoxElement:
        """Delete the ``occurrence``-th ``text`` in ``paragraph`` (across runs) as a tracked
        change by ``author``."""
        from .revisions import suggest_delete

        return suggest_delete(self, paragraph, text, author=author, date=date,
                              occurrence=occurrence)

    def suggest_replace(self, paragraph: OoxElement, old: str, new: str, *, author: str,
                        date=None, occurrence: int = 1):
        """Replace the ``occurrence``-th ``old`` in ``paragraph`` with ``new`` as a tracked
        deletion plus insertion; the new text keeps the old text's formatting."""
        from .revisions import suggest_replace

        return suggest_replace(self, paragraph, old, new, author=author, date=date,
                               occurrence=occurrence)

    def revisions(self) -> list[dict]:
        """Tracked changes in document order (body first, then other stories): ``{"kind",
        "author", "date", "text"}``; ``kind`` is insert, delete, move-from, move-to,
        paragraph-insert, paragraph-delete, row-insert, row-delete or format."""
        from .revisions import revisions

        return revisions(self)

    def accept_all(self) -> int:
        """Accept every tracked change (body, headers, footers, notes, comments)."""
        from .revisions import accept_all

        return accept_all(self)

    def reject_all(self) -> int:
        """Reject every tracked change in every story; returns how many there were."""
        from .revisions import reject_all

        return reject_all(self)

    # -- lists ------------------------------------------------------------------
    @property
    def numbering(self) -> OoxElement:
        """Root ``<w:numbering>`` of the numbering part, created on first use."""
        found = self.related_parts(self.main_part, "NumberingDefinitionsPart")
        if found:
            return found[0].element
        from .ns import w

        part, _ = self.add_part("NumberingDefinitionsPart", element=w.Numbering())
        return part.element

    def new_list(self, kind: str = "bullet", start: int = 1):
        """A new list: "bullet", "decimal" (1. a. i.) or "legal" (1. 1.1. 1.1.1.).
        Returns a ``ListHandle``; ``handle.add(text, level)`` appends items."""
        from .numbering import new_list

        return new_list(self, kind, start)

    def list_labels(self) -> list[tuple[OoxElement, str]]:
        """``(paragraph, label)`` for every list paragraph, as Word numbers them."""
        from .numbering import list_labels

        return list_labels(self)

    # -- styles and paragraphs --------------------------------------------------
    def style_id(self, style: str) -> str:
        """Style id for a style given by id ("Heading1") or name ("Heading 1",
        "heading 1"; Word stores built-in names in lower case)."""
        from .grammar import URI

        w_ns = URI["w"]
        found = self.related_parts(self.main_part, "StyleDefinitionsPart")
        if found:
            wanted = style.lower()
            for st in found[0].element.iterchildren(f"{{{w_ns}}}style"):
                sid = st.get(f"{{{w_ns}}}styleId")
                name = st.first("w:name")
                if sid == style or (name is not None and
                                    (name.get(f"{{{w_ns}}}val") or "").lower() == wanted):
                    return sid
        raise KeyError(f"no style {style!r} in this document (documents made with "
                       "new(minimal=True) have no styles part)")

    def add_paragraph(self, text: str = "", style: str | None = None) -> OoxElement:
        """Append a paragraph (before the final section properties); returns it.
        ``text`` may contain \\t and \\n (tab, line break)."""
        from . import teletype
        from .ns import w

        p = w.P(w.PPr(w.PStyle(val=self.style_id(style)))) if style else w.P()
        if text:
            teletype.addTextToElement(p, text)
        self.body.add(p)
        return p

    def add_heading(self, text: str, level: int = 1) -> OoxElement:
        """Heading paragraph: level 0 is "Title", 1-9 "Heading N"."""
        if not 0 <= level <= 9:
            raise ValueError("level must be 0 (Title) to 9")
        if level > 3 and self.related_parts(self.main_part, "StyleDefinitionsPart"):
            self.ensure_style(f"Heading{level}")
        return self.add_paragraph(text, "Title" if level == 0 else f"Heading {level}")

    def add_page_break(self) -> OoxElement:
        """A paragraph holding a page break, at the end of the body."""
        from .ns import w

        p = w.P(w.R(w.Br(type="page")))
        self.body.add(p)
        return p

    def add_picture(self, image, *, width: int | None = None, height: int | None = None,
                    paragraph: OoxElement | None = None, descr: str = "",
                    fallback=None) -> OoxElement:
        """Insert an inline picture; returns the ``w:r`` holding it.

        Size is in EMU (``oox.units``); give one dimension to keep the aspect
        ratio, none for the image's natural size.  Without ``paragraph`` the
        picture goes into a new paragraph at the end of the body.  An SVG
        needs ``fallback=`` (a PNG) for Office before 2016.
        """
        from .ns import a, pic, w, wp

        rid, info, svg_rid = self.add_picture_image(image, fallback)
        cx, cy = info.scaled(width, height)
        n = drawing.max_shape_id(self.document) + 1
        name = f"Picture {n}"
        inline = wp.Inline(
            wp.Extent(cx=cx, cy=cy),
            wp.EffectExtent(l=0, t=0, r=0, b=0),
            wp.DocPr(id=n, name=name, descr=descr or None),
            wp.CNvGraphicFramePr(a.GraphicFrameLocks(noChangeAspect=True)),
            a.Graphic(a.GraphicData(drawing.picture(pic, 0, name, rid, cx, cy, descr=descr,
                                                    svg_rid=svg_rid),
                                    uri=drawing.PIC_URI)),
            distT=0, distB=0, distL=0, distR=0,
        )
        run = w.R(w.Drawing(inline))
        if paragraph is None:
            self.body.add(w.P(run))
        else:
            paragraph.add(run)
        return run


    def add_chart(self, kind: str, categories, series, *, width: int = 5400000,
                  height: int = 2700000, paragraph: OoxElement | None = None,
                  descr: str = "", **options) -> OoxElement:
        """Insert an inline chart; returns the ``w:r`` holding it.

        ``series`` is ``{"North": [1, 2, 3], ...}`` (or ``oox.charts.Series``
        objects); the data is also stored in an embedded workbook, so Word's
        "Edit Data" works.  Size in EMU (default 15 x 7.5 cm).  Options are
        those of ``oox.charts.build_chart`` (title, legend, labels, ...).
        """
        from .charts import CHART_URI, add_embedded_chart
        from .ns import a, c, w, wp

        rid, _ = add_embedded_chart(self, self.main_part, kind, list(categories), series,
                                    **options)
        n = drawing.max_shape_id(self.document) + 1
        inline = wp.Inline(
            wp.Extent(cx=width, cy=height),
            wp.EffectExtent(l=0, t=0, r=0, b=0),
            wp.DocPr(id=n, name=f"Chart {n}", descr=descr or None),
            wp.CNvGraphicFramePr(),
            a.Graphic(a.GraphicData(c.Chart(r_id=rid), uri=CHART_URI)),
            distT=0, distB=0, distL=0, distR=0,
        )
        run = w.R(w.Drawing(inline))
        if paragraph is None:
            self.body.add(w.P(run))
        else:
            paragraph.add(run)
        return run

    def charts(self) -> list[dict]:
        """Charts in the body, in order: ``oox.charts.chart_info`` plus "rid"."""
        from .charts import charts_in

        return charts_in(self, self.main_part, self.document)


def open_document(file: str | os.PathLike | IO[bytes]) -> OoxDocument:
    """Open any OOXML package with the matching document class."""
    from .presentation import PresentationDocument
    from .spreadsheet import SpreadsheetDocument

    pkg = Package.open(file)
    main = pkg.related(None, RT_OFFICE_DOCUMENT)
    if main is None:
        raise PackageError("no officeDocument relationship")
    for cls in (WordprocessingDocument, SpreadsheetDocument, PresentationDocument):
        if main.content_type in cls.MAIN_CONTENT_TYPES.values():
            return cls(pkg)
    raise PackageError(f"unsupported main part content type: {main.content_type}")

