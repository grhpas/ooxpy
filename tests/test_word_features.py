"""Sections, headers/footers, links, notes, comments, TOC, template replace."""

from __future__ import annotations

import datetime as dt
import io

import pytest

from oox import WordprocessingDocument, open_document, teletype
from oox.fields import field_codes
from oox.ns import w
from oox.records import Hyperlink, Note
from oox.units import Inches


def reopen(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


# -- sections, headers, footers ---------------------------------------------------------
def test_new_document_has_one_a4_portrait_section():
    doc = WordprocessingDocument.new()
    (sec,) = doc.sections
    assert sec.orientation == "portrait"
    assert (sec.page_width, sec.page_height) == (595.3, 841.9)          # A4 in points
    assert sec.text_width == pytest.approx(595.3 - 2 * 72)


def test_orientation_swaps_size_and_round_trips():
    doc = WordprocessingDocument.new()
    sec = doc.sections[0]
    sec.orientation = "landscape"
    assert (sec.page_width, sec.page_height) == (841.9, 595.3)
    assert reopen(doc).sections[0].orientation == "landscape"
    sec.orientation = "portrait"
    assert sec.page_width == 595.3 and sec.sect_pr.first("w:pgSz").get(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}orient") is None
    with pytest.raises(ValueError):
        sec.orientation = "sideways"


def test_header_footer_text_with_page_fields():
    doc = WordprocessingDocument.new()
    sec = doc.sections[0]
    sec.header_text("Report\t\tDraft")
    sec.footer_text("Page {PAGE} of {NUMPAGES}", align="center")
    back = reopen(doc).sections[0]
    assert teletype.extractText(back.header()).replace("\t", "|") == "Report||Draft"
    footer = back.footer()
    assert field_codes(footer) == ["PAGE", "NUMPAGES"]
    assert footer.first("w:p").first("w:pPr").first("w:jc").get(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "center"
    assert not doc.validate()


def test_header_footer_styles_are_added_only_when_used():
    doc = WordprocessingDocument.new()
    with pytest.raises(KeyError):
        doc.style_id("Header")
    doc.sections[0].header_text("x")
    assert doc.style_id("Header") == "Header"
    with pytest.raises(KeyError):
        doc.style_id("Footer")


def test_first_and_even_headers_set_the_switches():
    doc = WordprocessingDocument.new()
    sec = doc.sections[0]
    sec.header_text("first", kind="first")
    sec.header_text("even", kind="even")
    assert sec.different_first_page and doc.even_and_odd_headers
    with pytest.raises(ValueError):
        sec.header(kind="odd")
    sec.different_first_page = False
    doc.even_and_odd_headers = False
    assert not sec.different_first_page and not doc.even_and_odd_headers


def test_add_section_links_headers_and_drops_first_page():
    doc = WordprocessingDocument.new()
    first = doc.sections[0]
    first.header_text("cover", kind="first")
    first.footer_text("Page {PAGE}")
    doc.add_paragraph("body of section 1")
    wide = doc.add_section(landscape=True)
    doc.add_paragraph("body of section 2")
    secs = doc.sections
    assert [s.orientation for s in secs] == ["portrait", "landscape"]
    assert secs[0].different_first_page and not wide.different_first_page
    assert wide.linked_to_previous("footer")
    assert field_codes(wide.shown("footer")) == ["PAGE"]
    assert wide.shown("header") is None
    wide.header_text("appendix")
    assert not wide.linked_to_previous("header")
    # the closing paragraph comes before the new section's content
    texts = [teletype.extractText(p) for p in doc.body.iterchildren(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")]
    assert texts.index("body of section 1") < texts.index("") < texts.index("body of section 2")
    back = reopen(doc)
    assert len(back.sections) == 2
    assert teletype.extractText(back.sections[1].header()) == "appendix"
    assert not doc.validate(target="Office2007")


def test_shared_header_part_is_split_on_write():
    doc = WordprocessingDocument.new()
    doc.sections[0].header_text("one")
    doc.add_section()
    # point section 2 at section 1's header, as Word files sometimes do
    ref = doc.sections[0].sect_pr.first("w:headerReference")
    import copy

    doc.sections[1].sect_pr.add(copy.deepcopy(ref))
    assert teletype.extractText(doc.sections[1].header()) == "one"
    doc.sections[1].header_text("two")
    assert teletype.extractText(doc.sections[0].header()) == "one"
    assert teletype.extractText(doc.sections[1].header()) == "two"


def test_section_page_size_and_margins():
    doc = WordprocessingDocument.new()
    sec = doc.add_section(page="Letter", start="oddPage")
    assert (sec.page_width, sec.page_height) == (612, 792)
    assert sec.start == "oddPage"
    sec.set_margins(left=36, right=36)
    assert sec.margins["left"] == 36 and sec.text_width == 612 - 72
    assert sec.sect_pr.first("w:pgMar").get(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left") == "720"
    with pytest.raises(ValueError, match="no room"):
        sec.set_margins(left=720)            # twips from old code: 10 inches
    sec.set_page_size(Inches(11), Inches(8.5))
    assert (sec.page_width, sec.orientation) == (792, "landscape")
    with pytest.raises(ValueError, match="Emu"):
        sec.set_margins(left=914400)
    with pytest.raises(TypeError):
        sec.set_margins(inside=10)


# -- links and bookmarks ---------------------------------------------------------------------
def test_hyperlinks_and_bookmarks_round_trip():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("See ")
    doc.add_hyperlink(p, "Ecma", target="https://ecma-international.org/", tooltip="tip")
    doc.add_hyperlink(p, "below", anchor="target")
    doc.add_bookmark("target", doc.add_heading("Target", 1))
    back = reopen(doc)
    assert back.hyperlinks() == [
        Hyperlink("Ecma", target="https://ecma-international.org/", tooltip="tip"),
        Hyperlink("below", anchor="target", kind="bookmark")]
    assert back.bookmarks() == ["target"]
    rels = list(back.main_part.rels.values())
    assert any(r.target == "https://ecma-international.org/" and r.external for r in rels)
    assert not doc.validate()


@pytest.mark.parametrize("name", ["", "1abc", "has space", "x" * 41])
def test_bad_bookmark_names(name):
    doc = WordprocessingDocument.new()
    with pytest.raises(ValueError):
        doc.add_bookmark(name, doc.add_paragraph("x"))


def test_duplicate_bookmark_and_link_arguments():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("x")
    doc.add_bookmark("a", p)
    with pytest.raises(ValueError):
        doc.add_bookmark("a", p)
    with pytest.raises(ValueError):
        doc.add_hyperlink(p, "t")
    with pytest.raises(ValueError):
        doc.add_hyperlink(p, "t", target="u", anchor="a")


# -- notes and comments ------------------------------------------------------------------------
def test_footnotes_endnotes_numbering_and_separators():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("Text")
    doc.add_footnote(p, "one")
    doc.add_footnote(p, "two")
    doc.add_endnote(p, "end")
    back = reopen(doc)
    assert back.footnotes() == [Note(1, "one"), Note(2, "two")]
    assert back.endnotes() == [Note(1, "end")]
    fn_root = back.related_parts(back.main_part, "FootnotesPart")[0].element
    types = [n.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type")
             for n in fn_root]
    assert types[:2] == ["separator", "continuationSeparator"]
    assert back.settings.first("w:footnotePr") is not None
    assert not doc.validate(target="Office2007")


def test_no_notes_or_comments_without_parts():
    doc = WordprocessingDocument.new()
    assert doc.footnotes() == [] and doc.endnotes() == [] and doc.comments() == []
    assert doc.related_parts(doc.main_part, "FootnotesPart") == []


def test_comments_round_trip():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("Commented sentence.")
    doc.add_comment(p, "Check this.", author="Ada Lovelace", date=dt.datetime(2026, 11, 14, 9, 30))
    doc.add_comment(doc.add_paragraph("Another."), "Second.")
    (first, second) = reopen(doc).comments()
    assert (first.id, first.author, first.text, first.anchor_text) == (
        0, "Ada Lovelace", "Check this.", "Commented sentence.")
    assert first.date.startswith("2026-11-14T09:30")
    assert second.id == 1 and second.author == "oox"
    c = doc.related_parts(doc.main_part, "WordprocessingCommentsPart")[0].element.first("w:comment")
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    assert c.get(W + "initials") == "AL" and c.get(W + "date") == "2026-11-14T09:30:00Z"
    assert not doc.validate()


# -- TOC ----------------------------------------------------------------------------------------
def test_toc_lists_headings_in_range_and_is_one_field():
    doc = WordprocessingDocument.new()
    doc.add_heading("Title", 0)
    first = doc.add_heading("One", 1)
    doc.add_heading("One.one", 2)
    doc.add_heading("Deep", 4)
    doc.add_heading("Two", 1)
    paras = doc.add_toc(levels=(1, 2), before=first)
    texts = [teletype.extractText(p) for p in paras]
    assert texts == ["Contents", "One", "One.one", "Two"]
    assert field_codes(doc.body) == ["TOC"]
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    chars = [f.get(W + "fldCharType") for f in doc.body.iter(W + "fldChar")]
    assert chars == ["begin", "separate", "end"]
    assert doc.style_id("TOC2") == "TOC2"
    assert doc.settings.first("w:updateFields") is None
    assert not doc.validate()


def test_toc_without_headings_and_update_on_open():
    doc = WordprocessingDocument.new()
    paras = doc.add_toc(title=None, update_on_open=True)
    assert len(paras) == 1 and "F9" in teletype.extractText(paras[0])
    assert doc.settings.first("w:updateFields") is not None


def test_outline_level_follows_style_chain():
    doc = WordprocessingDocument.new()
    assert doc.outline_level(doc.add_heading("H", 3)) == 3
    assert doc.outline_level(doc.add_paragraph("body")) == 10
    p = doc.add_paragraph("direct")
    p.first("w:pPr") if p.first("w:pPr") is not None else p.insert(0, w.PPr())
    p.first("w:pPr").add(w.OutlineLvl(val=1))
    assert doc.outline_level(p) == 2


# -- template replace ------------------------------------------------------------------------------
def test_replace_across_runs_keeps_first_run_formatting():
    doc = WordprocessingDocument.new()
    p = w.P(w.R(w.RPr(w.B()), w.T("Dear {{na")), w.R(w.T("me}}, hi {{name}}")))
    doc.body.add(p)
    assert doc.replace_text("{{name}}", "Ada") == {"{{name}}": 2}
    assert teletype.extractText(p) == "Dear Ada, hi Ada"
    runs = list(p.iterchildren("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"))
    assert teletype.extractText(runs[0]) == "Dear Ada" and runs[0].first("w:rPr") is not None


def test_replace_three_way_split_and_spaces_preserved():
    doc = WordprocessingDocument.new()
    p = w.P(w.R(w.T("x{{")), w.R(w.T("a")), w.R(w.T("}}y")))
    doc.body.add(p)
    doc.replace_text({"{{a}}": "  two spaces  "})
    t = p.first("w:r").first("w:t")
    assert t.text == "x  two spaces  " and t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert teletype.extractText(p) == "x  two spaces  y"


def test_replace_reaches_headers_footers_notes_comments():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("{{x}}")
    doc.sections[0].header_text("H {{x}}")
    doc.sections[0].footer_text("F {{x}}")
    doc.add_footnote(p, "N {{x}}")
    doc.add_comment(p, "C {{x}}")
    assert doc.replace_text({"{{x}}": "ok", "{{missing}}": "-"}) == {"{{x}}": 5, "{{missing}}": 0}
    assert teletype.extractText(doc.sections[0].header()) == "H ok"
    assert doc.footnotes() == [Note(1, "N ok")]
    assert doc.comments()[0].text.endswith("C ok")


def test_replace_text_needs_new():
    with pytest.raises(TypeError):
        WordprocessingDocument.new().replace_text("a")


def test_page_break():
    doc = WordprocessingDocument.new()
    p = doc.add_page_break()
    br = p.first("w:r").first("w:br")
    assert br.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type") == "page"


def test_section_objects_stay_with_their_section_after_add_section():
    """Found while writing the Word guide: a Section taken before add_section()
    went on to describe the new section (the body's w:sectPr was reused)."""
    doc = WordprocessingDocument.new()
    doc.add_paragraph("cover")
    first = doc.sections[0]
    first.set_margins(left=90, right=90)
    wide = doc.add_section(landscape=True)
    assert first.orientation == "portrait" and wide.orientation == "landscape"
    assert first.text_width == doc.sections[0].text_width == pytest.approx(595.3 - 180)
    assert wide.text_width == doc.sections[1].text_width
    first.footer_text("still the cover section's footer")
    assert doc.sections[0].shown("footer") is not None


def test_section_margins_take_points_or_lengths():
    from oox.units import Cm

    doc = WordprocessingDocument.new()
    section = doc.sections[0]
    section.set_margins(left=Cm(2), right=72)
    assert section.margins["left"] == 56.7 and section.margins["right"] == 72
