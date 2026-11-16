"""PowerPoint: editing existing tables (oox.slidetable) and click links (oox.slidelinks)."""

from __future__ import annotations

import io

import pytest

from oox import PresentationDocument, open_document
from oox.units import Inches


def _reload(deck):
    buf = io.BytesIO()
    deck.save(buf)
    buf.seek(0)
    return open_document(buf)


def _deck_with_table(rows, **kw):
    deck = PresentationDocument.new(layouts="basic")
    slide = deck.add_slide("Blank")
    deck.add_table(slide, rows, Inches(1), Inches(1), Inches(6), name="Prices", **kw)
    return deck, slide


# -- tables ---------------------------------------------------------------------------
def test_values_cell_and_lookup_by_name_or_index():
    deck, slide = _deck_with_table([["Region", "Q1"], ["North", 120], ["South", 95]],
                                   merges=["A2:A3"])
    t = deck.table(slide, "Prices")
    assert (t.n_rows, t.n_cols) == (3, 2)
    assert t.values() == [["Region", "Q1"], ["North", "120"], [None, "95"]]
    assert t.text(2, 0) == "North"                  # the merge's top-left cell
    assert t.merges() == [(1, 0, 2, 0)]
    assert deck.table(slide, 0).element is t.element
    with pytest.raises(KeyError):
        deck.table(slide, "Nope")
    with pytest.raises(IndexError):
        deck.table(slide, 1)


def test_set_text_keeps_the_cell_formatting():
    deck, slide = _deck_with_table([["Region", "Q1"], ["North", 120]], size=20)
    t = deck.table(slide)
    t.set_text(1, 1, 125)
    rpr = t.cell(1, 1).find(".//{*}rPr")
    assert t.text(1, 1) == "125" and rpr.get("sz") == "2000"
    t.set_text(1, 0, "Two\nlines", bold=True, color="#c00000", align="center")
    cell = t.cell(1, 0)
    assert t.text(1, 0) == "Two\nlines"
    assert [r.get("b") for r in cell.iter("{*}rPr")] == ["1", "1"]
    assert cell.find(".//{*}srgbClr").get("val") == "C00000"
    assert {p.get("algn") for p in cell.iter("{*}pPr")} == {"ctr"}
    t.set_text(1, 0, "")
    assert t.text(1, 0) == "" and cell.find(".//{*}endParaRPr") is not None
    with pytest.raises(ValueError, match="align"):
        t.set_text(0, 0, "x", align="middle")
    assert _reload(deck).validate(level="full") == {}


def test_rows_grow_and_shrink_the_frame():
    deck, slide = _deck_with_table([["a", "b"], ["1", "2"]], row_height=Inches(0.5))
    t = deck.table(slide)
    ext = t.frame.find(".//{*}xfrm/{*}ext")
    t.add_row(["3", "4"])
    t.add_row(["0", "0"], index=1)
    assert t.values() == [["a", "b"], ["0", "0"], ["1", "2"], ["3", "4"]]
    assert int(ext.get("cy")) == 4 * Inches(0.5)
    t.delete_row(1)
    assert t.values() == [["a", "b"], ["1", "2"], ["3", "4"]]
    assert int(ext.get("cy")) == 3 * Inches(0.5)
    with pytest.raises(IndexError):
        t.delete_row(7)
    assert _reload(deck).validate(level="full") == {}


def test_new_rows_and_columns_copy_formatting():
    deck, slide = _deck_with_table([["H1", "H2"], ["x", "y"]], size=14)
    t = deck.table(slide)
    t.add_row(["p", "q"])
    assert t.cell(2, 0).find(".//{*}rPr").get("sz") == "1400"
    t.add_column(["H3", "z", "r"], width=Inches(1))
    assert t.values()[0] == ["H1", "H2", "H3"]
    assert t.widths[-1] == Inches(1)
    assert int(t.frame.find(".//{*}xfrm/{*}ext").get("cx")) == sum(t.widths)
    t.add_column(["H0", "w", "o"], index=0)
    assert [r[0] for r in t.values()] == ["H0", "w", "o"]
    t.delete_column(0)
    assert t.values() == [["H1", "H2", "H3"], ["x", "y", "z"], ["p", "q", "r"]]


def test_merge_and_unmerge():
    deck, slide = _deck_with_table([["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]])
    t = deck.table(slide)
    t.merge(0, 0, 1, 1)
    assert t.values() == [["a\nb\nd\ne", None, "c"], [None, None, "f"], ["g", "h", "i"]]
    assert t.merges() == [(0, 0, 1, 1)]
    covered = t._tc(1, 0)
    assert covered.get("vMerge") == "1" and covered.get("gridSpan") == "2"   # as PowerPoint
    with pytest.raises(ValueError, match="overlaps"):
        t.merge(1, 1, 2, 2)
    with pytest.raises(ValueError, match="inside a merge"):
        t.add_row(index=1)
    with pytest.raises(ValueError, match="inside a merge"):
        t.add_column(index=1)
    again = _reload(deck)
    assert again.validate(level="full") == {}
    assert again.tables(again.slides[0])[0]["merges"] == ["A1:B2"]
    t.unmerge(1, 1)
    assert t.merges() == [] and t.text(0, 0) == "a\nb\nd\ne" and t.text(1, 1) == ""


def test_deleting_through_a_merge_shrinks_it():
    deck, slide = _deck_with_table([["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]])
    t = deck.table(slide)
    t.merge(0, 0, 2, 1)                       # 3 rows x 2 columns
    t.delete_row(0)                           # the origin row: the merge moves down
    assert t.merges() == [(0, 0, 1, 1)]
    assert t.text(0, 0) == "a\nb\nd\ne\ng\nh"
    t.delete_column(1)                        # a covered column
    assert t.merges() == [(0, 0, 1, 0)]
    t.delete_row(1)                           # a covered row
    assert t.merges() == []
    assert t.values() == [["a\nb\nd\ne\ng\nh", "f"]]
    assert _reload(deck).validate(level="full") == {}


def test_set_style_and_switches():
    deck, slide = _deck_with_table([["a"], ["b"]])
    t = deck.table(slide)
    t.set_style("No Style, Table Grid", header=False, first_column=True)
    info = deck.tables(slide)[0]
    assert info["style_name"] == "No Style, Table Grid"
    assert (info["header"], info["banded"], info["first_column"]) == (False, True, True)


def test_editing_a_python_pptx_table():
    pptx = pytest.importorskip("pptx")
    from pptx.util import Inches as PInches

    prs = pptx.Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    table = s.shapes.add_table(3, 3, PInches(1), PInches(1), PInches(6), PInches(2)).table
    for i in range(3):
        for j in range(3):
            table.cell(i, j).text = f"{i}{j}"
    table.cell(0, 0).merge(table.cell(1, 1))
    buf = io.BytesIO()
    prs.save(buf)
    deck = open_document(io.BytesIO(buf.getvalue()))
    t = deck.table(deck.slides[0])
    assert t.merges() == [(0, 0, 1, 1)]
    t.delete_column(0)
    t.add_row(["x", "y"])
    t.set_text(2, 1, "edited")
    out = io.BytesIO()
    deck.save(out)
    back = pptx.Presentation(io.BytesIO(out.getvalue())).slides[0].shapes[0].table
    assert [[c.text for c in r.cells] for r in back.rows][2:] == [["21", "edited"], ["x", "y"]]
    assert back.cell(0, 0).is_merge_origin and back.cell(0, 0).span_height == 2
    assert back.cell(0, 0).span_width == 1


# -- links ----------------------------------------------------------------------------
def _deck_with_text():
    deck = PresentationDocument.new(layouts="basic")
    s1, s2, s3 = (deck.add_slide("Blank") for _ in range(3))
    box = deck.add_textbox(s1, Inches(1), Inches(1), Inches(6), Inches(1),
                           "See the docs, then the docs index")
    return deck, (s1, s2, s3), box


def test_text_links_split_runs_and_share_one_relationship():
    deck, (s1, _, _), box = _deck_with_text()
    deck.set_link(s1, box, "https://example.org/docs", text="docs", tooltip="Manual")
    runs = [r.findtext("{*}t") for r in box.iter("{*}r")]
    assert runs == ["See the ", "docs", ", then the ", "docs", " index"]
    rids = {h.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            for h in box.iter("{*}hlinkClick")}
    assert len(rids) == 1
    assert deck.links(s1) == [
        {"shape": "TextBox 2", "text": "docs", "kind": "url",
         "target": "https://example.org/docs", "tooltip": "Manual"}] * 2
    assert _reload(deck).validate(level="full") == {}


def test_a_link_across_runs_reads_as_one():
    deck, (s1, _, _), box = _deck_with_text()
    deck.set_link(s1, box, "mailto:team@example.org", text="See the docs")
    assert [(lk["text"], lk["target"]) for lk in deck.links(s1)] == [
        ("See the docs", "mailto:team@example.org")]


def test_shape_links_to_slides_and_show_jumps():
    deck, (s1, s2, s3), _ = _deck_with_text()
    btn = deck.add_shape(s1, "rightArrow", Inches(1), Inches(3), Inches(2), Inches(1),
                         text="Go")
    deck.set_link(s1, btn, s3)
    deck.set_link(s2, deck.add_shape(s2, "rect", 0, 0, Inches(1), Inches(1)), "next")
    deck.set_link(s3, deck.add_shape(s3, "rect", 0, 0, Inches(1), Inches(1)), "first")
    again = _reload(deck)
    assert again.validate(level="full") == {}
    s1b, s2b, s3b = again.slides
    assert [(lk["kind"], lk["target"], lk["text"]) for lk in again.links(s1b)] == [
        ("slide", 3, None)]
    assert again.links(s2b)[0] | {"shape": None} == {"shape": None, "text": None,
                                                     "kind": "jump", "target": "next",
                                                     "tooltip": None}
    assert again.links(s3b)[0]["target"] == "first"
    # by shape name, replacing and removing
    name = btn.find(".//{*}cNvPr").get("name")
    deck.set_link(s1, name, "https://example.org/")
    assert [lk["kind"] for lk in deck.links(s1)] == ["url"]
    deck.set_link(s1, name, None)
    assert deck.links(s1) == []


def test_link_errors():
    deck, (s1, _, _), box = _deck_with_text()
    with pytest.raises(ValueError, match="scheme"):
        deck.set_link(s1, box, "example.org")
    with pytest.raises(ValueError, match="not in the shape"):
        deck.set_link(s1, box, "https://x.org/", text="missing")
    with pytest.raises(ValueError, match="itself"):
        deck.set_link(s1, box, s1)
    with pytest.raises(KeyError):
        deck.set_link(s1, "No such shape", "next")
    with pytest.raises(TypeError):
        deck.set_link(s1, box, 42)


def test_python_pptx_sees_oox_links():
    pptx = pytest.importorskip("pptx")
    deck, (s1, _, _), box = _deck_with_text()
    deck.set_link(s1, box, "https://example.org/docs", text="docs")
    btn = deck.add_shape(s1, "rect", 0, 0, Inches(1), Inches(1), text="b")
    deck.set_link(s1, btn, "https://example.org/button")
    buf = io.BytesIO()
    deck.save(buf)
    shapes = pptx.Presentation(io.BytesIO(buf.getvalue())).slides[0].shapes
    run_links = [r.hyperlink.address for p in shapes[0].text_frame.paragraphs for r in p.runs]
    assert run_links == [None, "https://example.org/docs", None, "https://example.org/docs", None]
    assert shapes[1].click_action.hyperlink.address == "https://example.org/button"


def test_reading_python_pptx_links(tmp_path):
    pytest.importorskip("pptx")
    from tools.writers import _python_pptx

    path = next(f for f in _python_pptx(tmp_path) if "features" in f.name)
    deck = open_document(path)
    assert deck.links(deck.slides[2]) == [
        {"shape": "Rounded Rectangle 2", "text": "rounded", "kind": "url",
         "target": "https://example.org/", "tooltip": None}]


# -- sections ---------------------------------------------------------------------------
def test_sections_partition_the_slides_and_follow_new_slides():
    deck = PresentationDocument.new(layouts="basic")
    slides = [deck.add_slide("Blank") for _ in range(5)]
    deck.set_sections([("Intro", slides[1]), ("Results", 4)])
    assert deck.sections() == [("Default Section", [1]), ("Intro", [2, 3]),
                               ("Results", [4, 5])]
    deck.add_slide("Blank")
    assert deck.sections()[-1] == ("Results", [4, 5, 6])
    again = _reload(deck)
    assert again.validate(level="full") == {}
    assert again.validate(target="Office2007") == {}
    assert again.sections() == deck.sections()
    deck.set_sections([])
    assert deck.sections() == [] and deck.presentation.first("p:extLst") is None


def test_sections_are_deterministic_and_checked():
    def build():
        deck = PresentationDocument.new(layouts="basic")
        s = [deck.add_slide("Blank") for _ in range(3)]
        deck.set_sections([("A", s[0]), ("B", s[2])])
        buf = io.BytesIO()
        deck.save(buf)
        return deck, s, buf.getvalue()

    deck, s, data = build()
    assert build()[2] == data
    with pytest.raises(ValueError, match="slide order"):
        deck.set_sections([("B", s[2]), ("A", s[0])])
    with pytest.raises(IndexError):
        deck.set_sections([("A", 9)])
    with pytest.raises(ValueError, match="name"):
        deck.set_sections([("", 1)])


def test_python_pptx_opens_a_deck_with_sections():
    pptx = pytest.importorskip("pptx")
    deck = PresentationDocument.new(layouts="basic")
    s = [deck.add_slide("Blank") for _ in range(2)]
    deck.set_sections([("One", s[0]), ("Two", s[1])])
    buf = io.BytesIO()
    deck.save(buf)
    assert len(pptx.Presentation(io.BytesIO(buf.getvalue())).slides) == 2
