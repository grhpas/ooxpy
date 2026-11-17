"""Reading a document the way Word does (oox.walk): text boxes count once,
block-level content controls are looked through.  Defects reported from
Petals Office (2026-11-17) and found while checking them."""

from __future__ import annotations

import copy

import pytest

from oox import WordprocessingDocument, teletype
from oox.errors import OoxError
from oox.grammar import URI
from oox.ns import v, w
from oox.walk import chosen_branch, iter_blocks, iter_story, twins

W, MC = URI["w"], URI["mc"]
P = f"{{{W}}}p"


def _box_doc():
    doc = WordprocessingDocument.new()
    anchor = doc.add_paragraph("anchor")
    content = w.P(w.R(w.T("in the box")))
    doc.add_text_box([content], paragraph=anchor)
    return doc


def test_story_walk_takes_one_copy_of_a_text_box():
    doc = _box_doc()
    all_paragraphs = list(doc.body.iter(P))
    shown = list(iter_story(doc.body, P))
    assert len(all_paragraphs) == len(shown) + 1           # the VML copy's paragraph
    assert [teletype.extractText(p, text_boxes=False) for p in shown] == ["anchor", "in the box"]
    assert [teletype.extractText(p) for p in iter_story(doc.body, P, text_boxes=False)] \
        == ["anchorin the box"]
    (shown_box, hidden_box), = twins(doc.document)
    assert shown_box.getparent().tag.endswith("txbx") and hidden_box.getparent().tag.endswith("textbox")


def test_unknown_choice_falls_back():
    alt = w.R()
    from lxml import etree

    ac = etree.SubElement(alt, f"{{{MC}}}AlternateContent", nsmap={"mc": MC, "zz": "urn:unknown"})
    etree.SubElement(ac, f"{{{MC}}}Choice", Requires="zz")
    fb = etree.SubElement(ac, f"{{{MC}}}Fallback")
    assert chosen_branch(ac) is fb


def test_list_item_in_a_text_box_is_numbered_once():
    doc = WordprocessingDocument.new()
    numbers = doc.new_list("decimal")
    numbers.add("one")
    doc.add_text_box([numbers.paragraph("in a text box")])
    numbers.add("two")
    assert [item.label for item in doc.list_labels()] == ["1.", "2.", "3."]


def test_fields_bookmarks_and_captions_in_a_text_box():
    doc = WordprocessingDocument.new()
    caption = doc.add_caption("Inside the box")
    doc.body.remove(caption.paragraph)
    doc.add_text_box([caption.paragraph], paragraph=doc.add_paragraph("anchor"))
    doc.add_caption("After the box")
    doc.update_fields()
    assert [f.result for f in doc.fields()] == ["1", "2"]
    assert len(doc.bookmarks()) == len(set(doc.bookmarks()))
    # both copies of the box show the same number (Word 2007 reads the VML one)
    texts = [teletype.extractText(p, text_boxes=False) for p in doc.body.iter(P)]
    assert texts.count("Figure 1: Inside the box") == 2 and "Figure 2: After the box" in texts
    assert "Figure 3: After the box" not in texts


def _sdt_section_doc():
    doc = WordprocessingDocument.new()
    doc.add_paragraph("cover")
    doc.add_section()
    doc.add_paragraph("contents")
    end_of_toc = w.P(w.PPr(copy.deepcopy(doc.body.first("w:sectPr"))), w.R(w.T("end of toc")))
    doc.body.add(w.Sdt(w.SdtContent(end_of_toc)))
    doc.add_paragraph("Executive summary")
    return doc, end_of_toc


def test_section_break_inside_a_content_control():
    doc, end_of_toc = _sdt_section_doc()
    assert sum(1 for _ in doc.body.iter(f"{{{W}}}sectPr")) == len(doc.sections) == 3
    middle = doc.sections[1]
    assert middle.sect_pr.getparent().getparent() is end_of_toc
    assert middle.shown("header") is None                       # no StopIteration


def test_shown_on_a_removed_section_raises_a_clear_error():
    doc, end_of_toc = _sdt_section_doc()
    middle = doc.sections[1]
    end_of_toc.getparent().remove(end_of_toc)
    with pytest.raises(OoxError, match="no longer part"):
        middle.shown("header")


def test_tables_inside_content_controls_and_blocks():
    doc = WordprocessingDocument.new()
    table = doc.add_table([["a"]])
    doc.body.remove(table.element)
    doc.body.add(w.Sdt(w.SdtContent(table.element)))
    assert len(doc.tables()) == 1
    assert "tbl" in [el.tag.rpartition("}")[2] for el in iter_blocks(doc.body)]


def test_pictures_are_only_pictures():
    from tests.test_docs import png

    doc = WordprocessingDocument.new()
    doc.add_picture(png(b"\x01\x02\x03"))
    doc.add_text_box("not a picture")
    doc.add_chart("column", ["a"], {"s": [1]})
    rid = doc.addPictureFromString(png(b"\x04\x05\x06"), "image/png")
    vml = v.Shape(v.Imagedata(r_id=rid), style="position:absolute;width:72pt;height:36pt;z-index:-5")
    doc.body.add(w.P(w.R(w.Pict(vml))))
    found = doc.pictures()
    assert [(p.kind, p.floating, p.wrap) for p in found] == [
        ("drawing", False, "inline"), ("vml", True, "behind")]
    assert found[1].width == 72 and found[1].embed == rid


def test_command_line_counts_once(capsys, tmp_path):
    from oox.__main__ import main

    doc = _box_doc()
    doc.save(tmp_path / "box.docx")
    assert main(["info", str(tmp_path / "box.docx"), "--json"]) == 0
    import json

    info = json.loads(capsys.readouterr().out)
    assert info["paragraphs"] == 2
