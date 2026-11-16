"""PowerPoint depth: standard layouts, shapes, tables, notes, picture placeholders."""

from __future__ import annotations

import io

import pytest

from oox.document import open_document
from oox.presentation import PresentationDocument
from oox.slides import SHAPES, TABLE_STYLES, table_info
from oox.units import Inches
from tools.officecheck.assets import image


def _reload(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


def _valid(doc):
    problems = []
    for part in doc.package.parts.values():
        if part.content_type.endswith("xml") and part.partname.startswith("/ppt/"):
            problems += [f"{part.partname}: {p}" for p in part.element.validate(target="Office2007")]
    return problems


def test_standard_layouts_have_office_types_and_placeholders():
    deck = PresentationDocument.new()
    types = [lp.element.get("type") for lp in deck.layouts]
    assert types == ["title", "obj", "secHead", "twoObj", "twoTxTwoObj", "titleOnly", "blank",
                     "objTx", "picTx", "vertTx", "vertTitleAndTx"]
    s = deck.add_slide("Comparison")
    assert [(t, i) for t, i, _ in deck.placeholders(s)] == [
        ("title", 0), ("body", 1), ("obj", 2), ("body", 3), ("obj", 4)]
    pic = deck.add_slide("Picture with Caption")
    assert [(t, i) for t, i, _ in deck.placeholders(pic)] == [("title", 0), ("pic", 1), ("body", 2)]
    assert _valid(deck) == []
    # 4:3 decks scale the geometry
    narrow = PresentationDocument.new(width=9144000, height=6858000)
    sec = narrow.layout("Section Header").element
    off = sec.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}sp"
                   "//{http://schemas.openxmlformats.org/drawingml/2006/main}off")
    assert int(off.get("x")) == round(831850 * 9144000 / 12192000)
    with pytest.raises(ValueError, match="layouts"):
        PresentationDocument.new(layouts="all")


def test_slide_layout_reader():
    deck = PresentationDocument.new()
    deck.add_slide("Two Content")
    back = _reload(deck)
    assert back.slide_layout(back.slides[0]) == "Two Content"


def test_shapes_with_fill_line_and_text():
    deck = PresentationDocument.new()
    s = deck.add_slide()
    deck.add_shape(s, "roundRect", 0, 0, Inches(2), Inches(1), text="Step 1")
    deck.add_shape(s, "ellipse", Inches(3), 0, Inches(1), Inches(1), fill=None, line="FF0000",
                   line_width=2, color="000000", text="x\ny")
    assert _valid(deck) == []
    back = _reload(deck)
    first, second = back.shapes(back.slides[0])
    assert (first["kind"], first["geometry"], first["text"]) == ("shape", "roundRect", "Step 1")
    assert second["geometry"] == "ellipse" and second["text"] == "x\ny"
    ln = back.slides[0].findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}ln")[1]
    assert ln.get("w") == "25400"
    with pytest.raises(ValueError, match="shape 'star'"):
        deck.add_shape(s, "star", 0, 0, 1, 1)
    assert SHAPES["roundRect"] == 5 and SHAPES["rightArrow"] == 33


def test_table_with_style_and_merges():
    deck = PresentationDocument.new()
    s = deck.add_slide("Title Only")
    deck.add_table(s, [["Region", "Q1", "Q2"], ["North", 120, 135], ["Total", None, 255]],
                   Inches(1), Inches(2), Inches(6), merges=["A3:B3"], style="Medium Style 2")
    deck.add_table(s, [["a"], ["b"]], 0, 0, Inches(1), header=False, banded=False,
                   style="{5940675A-B579-460E-94D1-54222C63F5DA}")
    assert _valid(deck) == []
    back = _reload(deck)
    t1, t2 = back.tables(back.slides[0])
    assert t1["rows"] == [["Region", "Q1", "Q2"], ["North", "120", "135"], ["Total", "", "255"]]
    assert t1["merges"] == ["A3:B3"] and t1["style_name"] == "Medium Style 2"
    assert (t1["header"], t1["banded"]) == (True, True)
    assert t2["style_name"] == "No Style, Table Grid" and not t2["header"]
    covered = back.slides[0].findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}tc")[7]
    assert covered.get("hMerge") == "1"
    with pytest.raises(ValueError, match="outside"):
        deck.add_table(s, [["a"]], 0, 0, merges=["A1:B1"])
    with pytest.raises(ValueError, match="column widths"):
        deck.add_table(s, [["a", "b"]], 0, 0, col_widths=[100])
    assert "Medium Style 2 - Accent 1" in TABLE_STYLES


def test_ragged_rows_are_padded():
    deck = PresentationDocument.new()
    s = deck.add_slide()
    frame = deck.add_table(s, [["a", "b", "c"], ["d"]], 0, 0)
    assert table_info(frame)["rows"] == [["a", "b", "c"], ["d", "", ""]]


def test_speaker_notes():
    deck = PresentationDocument.new()
    s1 = deck.add_slide()
    deck.add_slide()
    deck.set_notes(s1, "First line\nSecond line")
    deck.set_notes(s1, "Replaced")                      # same notes slide, new text
    assert _valid(deck) == []
    back = _reload(deck)
    a, b = back.slides
    assert back.notes(a) == "Replaced" and back.notes(b) is None
    assert sorted(p for p in back.package.parts if "notes" in p) == [
        "/ppt/notesMasters/notesMaster.xml", "/ppt/notesSlides/notesSlide1.xml"]
    lst = back.presentation.first("p:notesMasterIdLst")
    assert lst is not None and len(lst) == 1
    back.set_notes(b, "Later")                          # reuses the notes master
    assert sum("notesMaster" in p for p in back.package.parts) == 1
    assert back.notes(b) == "Later"


def test_picture_placeholder_crops_to_fill():
    deck = PresentationDocument.new()
    s = deck.add_slide("Picture with Caption")
    pic = deck.set_placeholder_picture(s, "pic", image("PNG"))
    assert _valid(deck) == []
    geom = pic.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
    assert geom is not None and geom.get("prst") == "rect"   # PowerPoint 2007 needs a shape
    src = pic.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect")
    # the caption box is 6172200 x 4873625 EMU (1.27:1); the test PNG is wider
    assert src is not None and src.get("l") == src.get("r") and int(src.get("l")) > 0
    back = _reload(deck)
    kinds = [(x["kind"], x["placeholder"]) for x in back.shapes(back.slides[0])]
    assert kinds == [("placeholder", ("title", 0)), ("picture", ("pic", 1)),
                     ("placeholder", ("body", 2))]
    stretched = deck.set_placeholder_picture(deck.add_slide("Two Content"), 1, image("PNG"),
                                             crop=False)
    assert stretched.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect") is None
