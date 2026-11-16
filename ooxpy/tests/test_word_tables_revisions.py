"""Tables (merges, widths, header rows), floating pictures, tracked changes."""

from __future__ import annotations

import copy
import datetime as dt
import io

import pytest

from oox import WordprocessingDocument, open_document, teletype
from oox.ns import w
from oox.units import Cm

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def reopen(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


def png() -> bytes:
    from tools.officecheck.assets import image

    return image("PNG")


# -- tables ------------------------------------------------------------------------------------
def test_add_table_defaults_share_text_width():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a", "b"], ["c", 1.5], ["d"]])
    assert (t.n_rows, t.n_cols) == (3, 2)
    assert t.widths == [doc.sections[0].text_width // 2] * 2
    assert t.values() == [["a", "b"], ["c", "1.5"], ["d", ""]]
    assert t.element.first("w:tblPr").first("w:tblLayout") is None
    assert not doc.validate(target="Office2007")


def test_widths_in_emu_or_twips_and_fixed_layout():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a", "b"]], widths=[Cm(5), 1440])
    assert t.widths == [2835, 1440]
    assert t.element.first("w:tblPr").first("w:tblLayout") is not None
    t.set_widths([2000, 3000])
    assert t.widths == [2000, 3000]
    tcw = t.cell(0, 1).first("w:tcPr").first("w:tcW")
    assert tcw.get(W + "w") == "3000"
    with pytest.raises(ValueError):
        t.set_widths([1])
    with pytest.raises(ValueError):
        doc.add_table([["a"]], widths=[1, 2])


def test_header_rows_repeat_and_bold():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["H1", "H2"], ["a", "b"]], header_rows=1)
    assert t.header_rows == 1
    assert t.cell(0, 0).find(f".//{W}b") is not None and t.cell(1, 0).find(f".//{W}b") is None
    t.set_header_rows(2)
    assert t.header_rows == 2
    t.set_header_rows(2)  # again: no duplicate w:tblHeader
    assert len(t.rows[0].first("w:trPr").findall(W + "tblHeader")) == 1
    t.set_header_rows(0)
    assert t.header_rows == 0


def test_horizontal_merge_keeps_text_and_grid_coordinates():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a", "b", "c"], ["d", "e", "f"]], widths=[1000, 2000, 3000])
    merged = t.merge(0, 0, 0, 1)
    assert t.cell(0, 1) is merged and t.cell(0, 2) is not merged
    assert t.values()[0] == ["a\nb", None, "c"]
    assert merged.first("w:tcPr").first("w:tcW").get(W + "w") == "3000"
    assert len(list(t.rows[0].iterchildren(W + "tc"))) == 2
    assert not doc.validate()


def test_vertical_and_rectangular_merge():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["A", "B", "C"], ["", "X", ""], ["", "", ""], ["D", "Y", ""]])
    t.merge(0, 0, 2, 0)
    t.merge(1, 1, 2, 2)
    assert t.values() == [["A", "B", "C"], [None, "X", None], [None, None, None],
                          ["D", "Y", ""]]
    assert t.cell(2, 0) is t.cell(0, 0) and t.cell(2, 2) is t.cell(1, 1)
    assert (t.n_rows, t.word_row_count) == (4, 3)  # row 3 is all continuations
    back = reopen(doc).tables()[0]
    assert back.values() == t.values()
    vms = [tc.first("w:tcPr").first("w:vMerge") for tc in back.element.iter(W + "tc")
           if tc.first("w:tcPr").first("w:vMerge") is not None]
    assert [v.get(W + "val") for v in vms] == ["restart", None, "restart", None, None]
    # every continuation cell still has a paragraph
    assert all(tc.first("w:p") is not None for tc in back.element.iter(W + "tc"))
    assert not doc.validate(target="Office2007")


def test_merge_errors():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a", "b", "c"], ["d", "e", "f"]])
    t.merge(0, 0, 1, 1)
    with pytest.raises(ValueError):
        t.merge(0, 1, 0, 2)  # cuts through the existing merge
    with pytest.raises(ValueError):
        t.merge(1, 1, 0, 0)
    with pytest.raises(IndexError):
        t.merge(0, 2, 5, 2)
    with pytest.raises(IndexError):
        t.cell(9, 0)


def test_format_cell():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a", "b"]])
    tc = t.format_cell(0, 1, fill="FFF2CC", valign="center", align="right", bold=True)
    tcpr = tc.first("w:tcPr")
    assert tcpr.first("w:shd").get(W + "fill") == "FFF2CC"
    assert tcpr.first("w:vAlign").get(W + "val") == "center"
    assert tc.first("w:p").first("w:pPr").first("w:jc").get(W + "val") == "right"
    assert tc.find(f".//{W}b") is not None
    t.format_cell(0, 1, bold=False)
    assert tc.find(f".//{W}b") is None
    with pytest.raises(ValueError):
        t.format_cell(0, 0, valign="middle")
    assert not doc.validate()


def test_set_text_add_row_and_row_height():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a", "b"]])
    t.set_text(0, 0, "new", bold=True)
    t.add_row(["x", "y"])
    t.set_row_height(1, Cm(1))
    assert t.values() == [["new", "b"], ["x", "y"]]
    assert t.rows[1].first("w:trPr").first("w:trHeight").get(W + "val") == "567"
    assert not doc.validate()


def test_consecutive_tables_are_kept_apart_and_body_ends_with_paragraph():
    doc = WordprocessingDocument.new(minimal=True)
    doc.add_table([["a"]], style=None)
    doc.add_table([["b"]], style=None)
    kinds = [el.tag.rpartition("}")[2] for el in doc.body if not el.tag.endswith("sectPr")]
    assert kinds == ["tbl", "p", "tbl"]
    back = reopen(doc)
    kinds = [el.tag.rpartition("}")[2] for el in back.body if not el.tag.endswith("sectPr")]
    assert kinds == ["tbl", "p", "tbl", "p"]
    assert len(back.tables()) == 2


def test_table_reader_on_a_word_style_table_with_grid_before():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a", "b"], ["c", "d"]])
    # Word's "gridBefore": row 2 starts one column in
    tr = t.rows[1]
    tr.remove(next(iter(tr.iterchildren(W + "tc"))))
    tr.insert(0, w.TrPr(w.GridBefore(val=1)))
    assert t.values() == [["a", "b"], [None, "d"]]
    with pytest.raises(IndexError):
        t.cell(1, 0)


def test_unknown_table_style():
    with pytest.raises(KeyError):
        WordprocessingDocument.new().add_table([["a"]], style="No Such Style")


# -- floating pictures ---------------------------------------------------------------------------
@pytest.mark.parametrize("wrap", ["square", "tight", "through", "topAndBottom", "front", "behind"])
def test_floating_picture_wraps_validate_and_read_back(wrap):
    doc = WordprocessingDocument.new()
    doc.add_floating_picture(png(), width=Cm(3), x=Cm(1), y=Cm(2), wrap=wrap, descr=wrap)
    assert not doc.validate(target="Office2007")
    (info,) = reopen(doc).pictures()
    assert info["floating"] and info["wrap"] == wrap and info["descr"] == wrap
    assert info["width"] == Cm(3)


def test_floating_alignment_and_reader_order():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("text")
    doc.add_picture(png(), width=Cm(1))
    doc.add_floating_picture(png(), width=Cm(2), h_align="center", v_align="top",
                             h_relative="page", v_relative="margin", paragraph=p)
    infos = doc.pictures()
    assert [i["floating"] for i in infos] == [True, False]  # document order
    anchor = doc.body.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor")
    pos_h = anchor.first("wp:positionH")
    assert pos_h.get("relativeFrom") == "page" and pos_h.first("wp:align").text == "center"
    ids = [d.get("id") for d in doc.body.iter("{*}docPr")]
    assert len(set(ids)) == 2


def test_floating_bad_options():
    doc = WordprocessingDocument.new()
    with pytest.raises(ValueError):
        doc.add_floating_picture(png(), wrap="around")
    with pytest.raises(ValueError):
        doc.add_floating_picture(png(), h_relative="paragraph")


# -- tracked changes ------------------------------------------------------------------------------
def tracked_doc():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("The quick brown fox jumps over the very lazy dog in 2025.")
    when = dt.datetime(2026, 11, 14, 12, 0)
    doc.suggest_delete(p, "very ", author="Bob", date=when)
    doc.suggest_replace(p, "2025", "2026", author="Ada")
    doc.suggest_insert(p, " Really.", author="Ada")
    return doc, p


def test_suggestions_are_listed_in_order():
    doc, _ = tracked_doc()
    revs = doc.revisions()
    assert [(r["kind"], r["author"], r["text"]) for r in revs] == [
        ("delete", "Bob", "very "), ("delete", "Ada", "2025"),
        ("insert", "Ada", "2026"), ("insert", "Ada", " Really.")]
    assert revs[0]["date"] == "2026-11-14T12:00:00Z" and revs[1]["date"] is None
    ids = [el.get(W + "id") for el in doc.body.iter(W + "ins", W + "del")]
    assert len(set(ids)) == 4
    assert not doc.validate(target="Office2007")


def test_accept_and_reject_all():
    doc, _ = tracked_doc()
    back = reopen(doc)
    assert back.accept_all() == 4
    assert teletype.extractText(back.body.first("w:p")) == \
        "The quick brown fox jumps over the lazy dog in 2026. Really."
    assert back.revisions() == []
    again = reopen(doc)
    assert again.reject_all() == 4
    assert teletype.extractText(again.body.first("w:p")) == \
        "The quick brown fox jumps over the very lazy dog in 2025."
    assert not again.validate() and not back.validate()


def test_replacement_keeps_formatting_and_split_runs():
    doc = WordprocessingDocument.new()
    p = w.P(w.R(w.RPr(w.B()), w.T("bold old")), w.R(w.T("er text")))
    doc.body.add(p)
    dele, ins = doc.suggest_replace(p, "older", "newer", author="Ada")
    assert "".join(t.text for t in dele.iter(W + "delText")) == "older"
    assert ins.first("w:r").first("w:rPr").first("w:b") is not None
    doc.accept_all()
    assert teletype.extractText(p) == "bold newer text"


def test_insert_after_and_occurrence():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("one two one two")
    doc.suggest_insert(p, " and a half", author="A", after="one two")
    doc.suggest_delete(p, "two", author="A", occurrence=2)
    doc.accept_all()
    assert teletype.extractText(p) == "one two and a half one "


def test_cannot_split_mixed_runs_or_missing_text():
    doc = WordprocessingDocument.new()
    p = w.P(w.R(w.T("ab"), w.Tab(), w.T("cd")))
    doc.body.add(p)
    with pytest.raises(ValueError):
        doc.suggest_delete(p, "b", author="A")
    with pytest.raises(ValueError):
        doc.suggest_delete(doc.add_paragraph("xyz"), "nope", author="A")


def test_accept_reject_word_markup_marks_rows_and_formatting():
    """What Word writes: paragraph-mark revisions, row insertions, rPrChange, moves."""
    doc = WordprocessingDocument.new()
    p1 = w.P(w.PPr(w.RPr(w.Ins(id=1, author="A"))), w.R(w.T("first")))
    p2 = w.P(w.R(w.RPr(w.B(), w.RPrChange(w.RPr(w.I()), id=2, author="A")), w.T("second")))
    p3 = w.P(w.MoveFrom(w.R(w.DelText("moved")), id=3, author="A"))
    p4 = w.P(w.MoveTo(w.R(w.T("moved")), id=4, author="A"))
    doc.body.add(p1, p2, p3, p4)
    t = doc.add_table([["kept"], ["new row"]], style=None)
    t.rows[1].insert(0, w.TrPr(w.Ins(id=5, author="A")))
    kinds = sorted(r["kind"] for r in doc.revisions())
    assert kinds == ["format", "move-from", "move-to", "paragraph-insert", "row-insert"]
    snapshot = copy.deepcopy(doc.body)

    doc.accept_all()
    texts = [teletype.extractText(p) for p in doc.body.iterchildren(W + "p")]
    assert texts[:4] == ["first", "second", "", "moved"]
    assert p2.find(f".//{W}b") is not None and p2.find(f".//{W}i") is None
    assert doc.tables()[0].n_rows == 2

    doc.document.replace(doc.body, snapshot)
    doc.reject_all()
    texts = [teletype.extractText(p) for p in doc.body.iterchildren(W + "p")]
    assert texts[0] == "firstsecond"  # the inserted paragraph mark is gone
    assert doc.body.find(f".//{W}i") is not None and doc.body.find(f".//{W}b") is None
    assert "moved" in texts and texts.count("moved") == 1
    assert doc.tables()[0].n_rows == 1
    assert doc.revisions() == []


def test_track_changes_switch():
    doc = WordprocessingDocument.new()
    assert not doc.track_changes
    doc.track_changes = True
    assert doc.track_changes and reopen(doc).track_changes
    doc.track_changes = False
    assert not doc.track_changes
