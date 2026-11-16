import pytest

from oox import PresentationDocument
from oox.units import Inches

pptx = pytest.importorskip("pptx")


def test_new_deck_has_complete_skeleton():
    deck = PresentationDocument.new()
    names = set(deck.package.parts)
    assert {"/ppt/presentation.xml", "/ppt/slideMasters/slideMaster1.xml",
            "/ppt/slideLayouts/slideLayout1.xml", "/ppt/theme/theme1.xml",
            "/ppt/presProps.xml", "/ppt/viewProps.xml", "/ppt/tableStyles.xml"} <= names
    for part in deck.package.parts.values():
        assert part.element.validate() == [], part.partname  # incl. mandatory children
    # master -> theme and layout <-> master relationships exist
    master = deck.masters[0]
    assert deck.related_parts(master, "ThemePart")
    assert deck.related_parts(deck.layouts[0], "SlideMasterPart") == [master]


def test_slides_text_and_order(tmp_path):
    deck = PresentationDocument.new()
    s1, s2 = deck.add_slide(), deck.add_slide()
    deck.add_textbox(s1, Inches(1), Inches(1), Inches(8), Inches(1), "Title\nsecond line",
                     size=40, bold=True, color="1F3864", align="center")
    deck.add_textbox(s2, Inches(1), Inches(1), Inches(8), Inches(1), "Two")
    assert deck.slides == [s1, s2]
    path = tmp_path / "d.pptx"
    deck.save(path)

    pr = pptx.Presentation(str(path))
    assert pr.slide_width == 12192000 and len(pr.slides) == 2
    tf = pr.slides[0].shapes[0].text_frame
    assert tf.text == "Title\nsecond line"
    run = tf.paragraphs[0].runs[0]
    assert run.font.size.pt == 40 and run.font.bold is True
    assert pr.slides[1].shapes[0].text_frame.text == "Two"


def test_ids_are_in_their_ranges():
    deck = PresentationDocument.new()
    deck.add_slide()
    deck.add_slide()
    sld_ids = [int(e.get("id")) for e in deck.presentation.first("p:sldIdLst")]
    assert sld_ids == [256, 257]
    master_id = int(deck.presentation.first("p:sldMasterIdLst")[0].get("id"))
    layout_id = int(deck.masters[0].element.first("p:sldLayoutIdLst")[0].get("id"))
    assert master_id == 2**31 and layout_id == 2**31 + 1


def test_reopen_and_add_slide(tmp_path):
    path = tmp_path / "d.pptx"
    deck = PresentationDocument.new()
    deck.add_slide()
    deck.save(path)
    again = PresentationDocument.open(path)
    again.add_textbox(again.add_slide(), 0, 0, Inches(4), Inches(1), "added")
    again.save(path)
    pr = pptx.Presentation(str(path))
    assert len(pr.slides) == 2 and pr.slides[1].shapes[0].text_frame.text == "added"


def test_default_layouts_and_placeholders(tmp_path):
    deck = PresentationDocument.new()
    assert deck.layout_names == ["Title Slide", "Title and Content", "Section Header",
                                 "Two Content", "Comparison", "Title Only", "Blank",
                                 "Content with Caption", "Picture with Caption",
                                 "Title and Vertical Text", "Vertical Title and Text"]
    assert PresentationDocument.new(layouts="basic").layout_names == [
        "Title Slide", "Title and Content", "Blank"]
    s1 = deck.add_slide("Title Slide")
    deck.set_placeholder_text(s1, "title", "Hello")
    deck.set_placeholder_text(s1, "subtitle", "world")
    s2 = deck.add_slide("Title and Content")
    deck.set_placeholder_text(s2, "title", "Agenda")
    deck.set_placeholder_text(s2, "body", ["A", (1, "a.1"), "B"])
    assert deck.placeholders(deck.add_slide()) == []  # default layout: Blank
    for part in deck.package.parts.values():
        assert part.element.validate() == [], part.partname
    path = tmp_path / "ph.pptx"
    deck.save(path)

    pr = pptx.Presentation(str(path))
    t, sub = pr.slides[0].placeholders
    assert (t.placeholder_format.type, t.text_frame.text) == (3, "Hello")  # CENTER_TITLE
    assert sub.text_frame.text == "world"
    body = pr.slides[1].placeholders[1]
    assert [(p.text, p.level) for p in body.text_frame.paragraphs] == [("A", 0), ("a.1", 1), ("B", 0)]


def test_placeholders_on_powerpoint_template(tmp_path):
    from pathlib import Path

    path = Path(pptx.__file__).parent / "templates" / "default.pptx"
    deck = PresentationDocument.open(path)
    assert deck.layout_names[:2] == ["Title Slide", "Title and Content"]  # master order
    slide = deck.add_slide("Two Content")
    # date/footer/slide-number placeholders are not copied, like PowerPoint
    assert [(t, i) for t, i, _ in deck.placeholders(slide)] == [("title", 0), ("obj", 1), ("obj", 2)]
    deck.set_placeholder_text(slide, "title", "Two columns")
    deck.set_placeholder_text(slide, 1, "left")
    deck.set_placeholder_text(slide, 2, "right")
    with pytest.raises(KeyError, match="has: title/0, obj/1, obj/2"):
        deck.set_placeholder_text(slide, "subtitle", "nope")
    out = tmp_path / "t.pptx"
    deck.save(out)
    texts = [sh.text_frame.text for sh in pptx.Presentation(str(out)).slides[-1].placeholders]
    assert texts == ["Two columns", "left", "right"]


def test_non_169_slide_size_scales_placeholders():
    deck = PresentationDocument.new(width=9144000, height=6858000)  # 4:3
    title = next(sh for t, _, sh in deck.placeholders(deck.masters[0].element) if t == "title")
    off = title.first("p:spPr").first("a:xfrm").first("a:ext")
    assert int(off.get("cx")) == round(10515600 * 9144000 / 12192000)
