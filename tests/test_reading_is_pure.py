"""Reading never changes a document (docs/STABILITY.md).

Every public method and property of the document classes and ``Section`` is
classified here as a reader or a writer; an unclassified one fails
``test_every_member_is_classified``, so a new reader can't skip the check.
Each reader is then called on several documents (including bare ones that lack
styles, settings, numbering and sections) and the saved bytes must not change.
Reported from Petals Office (2026-11-17): ``doc.sections``, ``page_width``,
``doc.styles``, ``doc.settings`` and ``header()`` used to add markup or parts.
"""

from __future__ import annotations

import inspect
import io
import zipfile
from collections.abc import Callable

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, open_document
from oox.sections import Section
from oox.units import Cm
from tests.test_docs import png

Reader = Callable[[object], object]


def _first_paragraph(doc):
    from oox.walk import P, iter_story

    return next(iter_story(doc.body, P), None)


def _images(doc):
    blip = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
    return [doc.image_of(b) for part in list(doc.package.parts.values()) if part.is_xml
            for b in part.element.iter(blip)]


def _parts(*names: str) -> dict[str, Reader]:
    return {n: (lambda d, n=n: getattr(d, n)) for n in names}


def _resolve_all(doc):
    from oox.walk import P, W, iter_story

    fmt = doc.formatting()
    out = [(fmt.paragraph(p), fmt.mark(p), fmt.explain(p),
            [(fmt.run(r), fmt.explain(r)) for r in iter_story(p, f"{{{W}}}r")])
           for p in iter_story(doc.body, P)]
    out += [(fmt.cell(tc), fmt.explain(tc)) for tc in iter_story(doc.body, f"{{{W}}}tc")]
    return out


WORD_READERS: dict[str, Reader] = {
    **_parts("styles_part", "settings_part", "numbering_part", "theme_part", "footnotes_part",
             "endnotes_part", "comments_part", "font_table_part"),
    "image_of": _images,
    "body": lambda d: d.body, "text": lambda d: d.text, "document": lambda d: d.document,
    "styles": lambda d: d.styles, "settings": lambda d: d.settings,
    "numbering": lambda d: d.numbering, "sections": lambda d: d.sections,
    "even_and_odd_headers": lambda d: d.even_and_odd_headers,
    "track_changes": lambda d: d.track_changes, "conformance": lambda d: d.conformance,
    "properties": lambda d: (d.properties.as_dict(), d.properties.title),
    "extended_properties": lambda d: d.extended_properties.as_dict(),
    "custom_properties": lambda d: d.custom_properties.as_dict(),
    "bookmarks": lambda d: d.bookmarks(), "charts": lambda d: d.charts(),
    "check_protection_password": lambda d: d.check_protection_password("x"),
    "comments": lambda d: d.comments(), "content_controls": lambda d: d.content_controls(),
    "effective_format": lambda d: _first_paragraph(d) is not None
    and d.effective_format(_first_paragraph(d)),
    "endnotes": lambda d: d.endnotes(), "fields": lambda d: d.fields(all_stories=True),
    "footnotes": lambda d: d.footnotes(),
    "formatting": lambda d: _resolve_all(d), "getElementsByType": lambda d: d.getElementsByType(
        __import__("oox.ns", fromlist=["w"]).w.P),
    "hyperlinks": lambda d: d.hyperlinks(), "list_labels": lambda d: d.list_labels(all_stories=True),
    "outline_level": lambda d: _first_paragraph(d) is not None
    and d.outline_level(_first_paragraph(d)),
    "part_of": lambda d: d.part_of(d.body), "pictures": lambda d: d.pictures(),
    "protection": lambda d: d.protection(),
    "related_parts": lambda d: d.related_parts(d.main_part, "StyleDefinitionsPart"),
    "revisions": lambda d: d.revisions(), "style_id": lambda d: d.style_id("Normal"),
    "tables": lambda d: [t.values() for t in d.tables()], "text_boxes": lambda d: d.text_boxes(),
    "validate": lambda d: d.validate(level="full", target="Office2007"),
}
WORD_WRITERS = {
    "accept_all", "addPicture", "addPictureFromFile", "addPictureFromString", "add_bookmark",
    "add_caption", "add_chart", "add_checkbox", "add_comment", "add_cross_reference",
    "add_date_control", "add_dropdown", "add_endnote", "add_field", "add_floating_picture",
    "add_footnote", "add_heading", "add_hyperlink", "add_image", "add_page_break",
    "add_paragraph", "add_part", "add_picture", "add_picture_image", "add_run", "add_section",
    "add_style", "add_table", "add_text_box", "add_text_control", "add_toc", "ensure_numbering",
    "ensure_settings", "ensure_style", "ensure_styles", "fill_controls", "format_paragraph",
    "new_list", "protect", "reject_all", "relate", "replace_text", "set_text_box_text",
    "suggest_delete", "suggest_insert", "suggest_replace", "unprotect", "update_fields",
}

SECTION_READERS: dict[str, Reader] = {
    name: (lambda n: lambda d: [getattr(s, n) for s in d.sections])(name)
    for name in ("margins", "orientation", "page_height", "page_width", "start", "text_width",
                 "different_first_page")
}
SECTION_READERS.update({
    "header": lambda d: [s.header(k) for s in d.sections for k in ("default", "first", "even")],
    "footer": lambda d: [s.footer(k) for s in d.sections for k in ("default", "first", "even")],
    "linked_to_previous": lambda d: [s.linked_to_previous() for s in d.sections],
    "shown": lambda d: [s.shown(w, k) for s in d.sections for w in ("header", "footer")
                        for k in ("default", "first", "even")],
})
SECTION_WRITERS = {"ensure_footer", "ensure_header", "footer_text", "header_text", "set_margins",
                   "set_page_size"}

EXCEL_READERS: dict[str, Reader] = {
    **_parts("shared_strings_part", "styles_part", "theme_part"), "image_of": _images,
    "workbook": lambda b: b.workbook, "spreadsheet": lambda b: b.spreadsheet,
    "worksheets": lambda b: b.worksheets, "sheet_names": lambda b: b.sheet_names,
    "tab_names": lambda b: b.tab_names, "chart_sheet_names": lambda b: b.chart_sheet_names,
    "conformance": lambda b: b.conformance,
    "properties": lambda b: b.properties.as_dict(),
    "extended_properties": lambda b: b.extended_properties.as_dict(),
    "custom_properties": lambda b: b.custom_properties.as_dict(),
    "sheet": lambda b: b.sheet("S"), "sheet_data": lambda b: b.sheet_data("S"),
    "cell_element": lambda b: b.cell_element("S", "B2"), "value": lambda b: b.value("S", "B2"),
    "values": lambda b: b.values("S"), "cells": lambda b: b.cells("S"),
    "display_text": lambda b: b.display_text("S", "B2"),
    "cell_format": lambda b: b.cell_format("S", "B2"),
    "column_width": lambda b: b.column_width("S", "A"), "row_height": lambda b: b.row_height("S", 1),
    "merged_cells": lambda b: b.merged_cells("S"), "frozen_panes": lambda b: b.frozen_panes("S"),
    "auto_filter_range": lambda b: b.auto_filter_range("S"),
    "defined_names": lambda b: b.defined_names(), "validations": lambda b: b.validations("S"),
    "conditional_formats": lambda b: b.conditional_formats("S"),
    "tables": lambda b: b.tables("S"), "comments": lambda b: b.comments("S"),
    "hyperlinks": lambda b: b.hyperlinks("S"), "page_setup": lambda b: b.page_setup("S"),
    "sheet_options": lambda b: b.sheet_options("S"), "charts": lambda b: b.charts("S"),
    "chart_sheet": lambda b: b.chart_sheet_names and b.chart_sheet(b.chart_sheet_names[0]),
    "workbook_protection": lambda b: b.workbook_protection(),
    "check_workbook_password": lambda b: b.check_workbook_password("x"),
    "getElementsByType": lambda b: b.getElementsByType(
        __import__("oox.ns", fromlist=["x"]).x.C),
    "part_of": lambda b: b.part_of(b.workbook),
    "related_parts": lambda b: b.related_parts(b.main_part, "WorksheetPart"),
    "validate": lambda b: b.validate(level="full", target="Office2007"),
}
EXCEL_WRITERS = {
    "addPicture", "addPictureFromFile", "addPictureFromString", "add_chart", "add_chart_sheet",
    "add_comment", "add_hyperlink", "add_image", "add_part", "add_picture", "add_picture_image",
    "add_sheet", "add_table", "add_validation", "append_row", "auto_filter", "autofit_columns",
    "cell_style", "conditional_format", "define_name", "set_frozen_panes", "merge_cells",
    "set_page_setup", "protect_sheet", "protect_workbook", "relate", "set_column_width",
    "set_row_height", "set_value", "share_strings", "set_sheet_options", "style_range",
    "unmerge_cells", "unprotect_sheet", "unprotect_workbook",
}

PPT_READERS: dict[str, Reader] = {
    "image_of": _images, "slide_size": lambda p: p.slide_size,
    "content_area": lambda p: [p.content_area(s) for s in p.slides],
    "presentation": lambda p: p.presentation, "slides": lambda p: p.slides,
    "masters": lambda p: p.masters, "layouts": lambda p: p.layouts,
    "layout_names": lambda p: p.layout_names, "conformance": lambda p: p.conformance,
    "layout": lambda p: p.layout(p.layout_names[0]),
    "properties": lambda p: p.properties.as_dict(),
    "extended_properties": lambda p: p.extended_properties.as_dict(),
    "custom_properties": lambda p: p.custom_properties.as_dict(),
    "placeholders": lambda p: [p.placeholders(s) for s in p.slides],
    "placeholder": lambda p: [p.placeholder(s, "title") for s in p.slides],
    "shapes": lambda p: [p.shapes(s) for s in p.slides],
    "tables": lambda p: [p.tables(s) for s in p.slides],
    "table": lambda p: [p.table(s).values() for s in p.slides if p.tables(s)],
    "notes": lambda p: [p.notes(s) for s in p.slides], "links": lambda p: [p.links(s) for s in p.slides],
    "sections": lambda p: p.sections(), "charts": lambda p: [p.charts(s) for s in p.slides],
    "slide_layout": lambda p: [p.slide_layout(s) for s in p.slides],
    "getElementsByType": lambda p: p.getElementsByType(__import__("oox.ns", fromlist=["a"]).a.T),
    "part_of": lambda p: p.part_of(p.presentation),
    "related_parts": lambda p: p.related_parts(p.main_part, "SlidePart"),
    "validate": lambda p: p.validate(level="full", target="Office2007"),
}
PPT_WRITERS = {
    "addPicture", "addPictureFromFile", "addPictureFromString", "add_chart", "add_image",
    "add_layout", "add_part", "add_picture", "add_picture_image", "add_shape", "add_slide",
    "add_table", "add_text_box", "relate", "set_link", "set_notes", "set_placeholder_picture",
    "set_placeholder_text", "set_sections",
}
# creating, opening and saving a document aren't reading or writing it
NOT_APPLICABLE = {"MAIN_CONTENT_TYPES", "new", "open", "save"}


def _public(cls) -> set[str]:
    out = set()
    for name in dir(cls):
        if name.startswith("_"):
            continue
        owner = next((k for k in cls.__mro__ if name in vars(k)), None)
        if owner is not None and owner.__module__.startswith("oox"):
            out.add(name)
    return out


@pytest.mark.parametrize("cls, readers, writers", [
    (WordprocessingDocument, WORD_READERS, WORD_WRITERS),
    (Section, SECTION_READERS, SECTION_WRITERS),
    (SpreadsheetDocument, EXCEL_READERS, EXCEL_WRITERS),
    (PresentationDocument, PPT_READERS, PPT_WRITERS),
])
def test_every_member_is_classified(cls, readers, writers):
    members = _public(cls)
    unclassified = members - set(readers) - writers - NOT_APPLICABLE
    assert not unclassified, f"classify {cls.__name__}: {sorted(unclassified)}"
    assert not set(readers) & writers
    assert not (set(readers) | writers) - members, "classified names that don't exist"


# -- documents to read ----------------------------------------------------------------
def _word_bare():
    """No styles, settings, numbering, or section of any kind."""
    doc = WordprocessingDocument.new(minimal=True)
    doc.add_paragraph("bare")
    assert doc.body.first("w:sectPr") is None
    return doc


def _word_rich():
    doc = WordprocessingDocument.new()
    doc.add_heading("Rich", 1)
    lst = doc.new_list("decimal")
    lst.add("one")
    p = doc.add_paragraph("text with a ")
    doc.add_hyperlink(p, "link", target="https://example.org")
    doc.add_footnote(p, "a note")
    doc.add_comment(p, "a comment", author="A")
    doc.add_bookmark("mark", p)
    doc.add_table([["a", "b"], [1, 2]], header_rows=1)
    doc.add_picture(png(b"\x10\x20\x30"), width=Cm(1))
    doc.add_text_box([lst.paragraph("in a box")])
    doc.add_chart("column", ["a"], {"s": [1]})
    doc.add_caption("A caption")
    doc.add_text_control(doc.add_paragraph("Name: "), tag="name")
    doc.sections[0].header_text("Header {PAGE}")
    doc.add_section(landscape=True)
    doc.suggest_insert(doc.add_paragraph("x"), "y", author="A")
    doc.protect("comments")
    return doc


def _book_bare():
    book = SpreadsheetDocument.new("S")
    book.append_row("S", ["a", 1])
    book.append_row("S", ["b", 2])
    return book


def _book_rich():
    book = _book_bare()
    book.style_range("S", "A1:B1", bold=True)
    book.add_table("S", "A1:B2", name="T1x")
    book.conditional_format("S", "B2", "data_bar")
    book.add_comment("S", "A2", "note", author="A")
    book.add_hyperlink("S", "C1", "https://example.org")
    book.add_validation("S", "A2", values=["a", "b"])
    book.merge_cells("S", "D1:E1")
    book.set_frozen_panes("S", "A2")
    book.add_chart("S", "column", "A1:B2")
    book.add_chart_sheet("Chart", "line", "S", "A1:B2")
    book.protect_workbook(password="x")
    return book


def _deck_bare():
    return PresentationDocument.new(layouts="basic")


def _deck_rich():
    deck = PresentationDocument.new()
    s = deck.add_slide("Title and Content")
    deck.set_placeholder_text(s, "title", "T")
    deck.add_table(s, [["a", "b"], [1, 2]], x=Cm(1), y=Cm(4))
    deck.set_notes(s, "notes")
    deck.set_link(s, deck.shapes(s)[0].name, "https://example.org")
    deck.add_chart(s, "column", ["a"], {"s": [1]})
    deck.set_sections([("One", s)])
    return deck


def _bytes(doc) -> dict[str, bytes]:
    buf = io.BytesIO()
    doc.save(buf)
    z = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    return {n: z.read(n) for n in z.namelist()}


CASES = [
    ("word-bare", _word_bare, {**WORD_READERS, **SECTION_READERS}),
    ("word-rich", _word_rich, {**WORD_READERS, **SECTION_READERS}),
    ("excel-bare", _book_bare, EXCEL_READERS),
    ("excel-rich", _book_rich, EXCEL_READERS),
    ("ppt-bare", _deck_bare, PPT_READERS),
    ("ppt-rich", _deck_rich, PPT_READERS),
]


@pytest.mark.parametrize("label, make, readers", CASES, ids=[c[0] for c in CASES])
def test_reading_changes_nothing(label, make, readers):
    raw = io.BytesIO()
    make().save(raw)
    data = raw.getvalue()
    changed = []
    for name, read in readers.items():
        doc = open_document(io.BytesIO(data))
        before = _bytes(doc)
        try:
            read(doc)
        except (KeyError, ValueError, LookupError):
            pass  # a reader may refuse (no style "Normal" in a bare document): it still mustn't write
        after = _bytes(doc)
        if after != before:
            parts = sorted(n for n in set(before) | set(after) if before.get(n) != after.get(n))
            changed.append(f"{name}: {parts}")
    assert not changed, f"readers that changed the {label} document: {changed}"


def test_bare_document_reports_defaults_and_writes_on_demand():
    doc = _word_bare()
    assert doc.styles is None and doc.settings is None and doc.numbering is None
    section = doc.sections[0]
    assert (section.page_width, section.orientation) == (612, "portrait")
    assert section.margins["left"] == 72 and section.header() is None
    assert doc.body.first("w:sectPr") is None                       # nothing written yet
    section.set_margins(left=36)                                     # the first write adds it
    assert doc.body.first("w:sectPr") is section.sect_pr
    assert doc.sections[0].margins["left"] == 36
    assert doc.ensure_styles() is doc.styles is not None
    header = doc.sections[0].ensure_header()
    assert doc.sections[0].header() is header and doc.sections[0].shown() is header
    doc.track_changes = True
    assert doc.settings is not None and doc.track_changes


def test_signature_of_header_has_no_create_flag():
    assert "create" not in inspect.signature(Section.header).parameters
