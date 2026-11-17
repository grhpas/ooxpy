"""Formatting and styles, fields and captions, content controls, extended and
custom properties, TIFF/EMF/WMF/SVG pictures."""

from __future__ import annotations

import datetime as dt
import io

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, images, open_document, teletype
from oox.fieldcalc import format_date, tokens
from oox.formatting import LineSpacing
from oox.ns import w
from oox.units import EMU_PER_INCH, Cm

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def reopen(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


def asset(fmt: str) -> bytes:
    from tools.officecheck.assets import image

    return image(fmt)


# -- images ------------------------------------------------------------------------------------
@pytest.mark.parametrize("fmt,ctype", [("TIFF", "image/tiff"), ("EMF", "image/x-emf"),
                                       ("WMF", "image/x-wmf"), ("SVG", "image/svg+xml")])
def test_more_formats_have_physical_size(fmt, ctype):
    info = images.inspect(asset(fmt))
    assert info.content_type == ctype
    cx, cy = info.native_size()
    assert round(cx / EMU_PER_INCH, 2) == 1.67 and round(cy / EMU_PER_INCH, 2) == 0.83


def test_tiff_big_endian_and_centimetre_resolution():
    import struct

    # 1 IFD: width 200, height 100, 40 px/cm resolution
    entries = [(256, 3, 1, 200), (257, 3, 1, 100), (296, 3, 1, 3)]
    ifd_off = 8
    n = len(entries) + 2
    rational_off = ifd_off + 2 + 12 * n + 4
    body = struct.pack(">H", n)
    for tag, typ, count, val in entries:
        body += struct.pack(">HHIHH", tag, typ, count, val, 0)
    body += struct.pack(">HHII", 282, 5, 1, rational_off)
    body += struct.pack(">HHII", 283, 5, 1, rational_off)
    body += struct.pack(">I", 0) + struct.pack(">II", 40, 1)
    info = images.inspect(b"MM\x00*" + struct.pack(">I", ifd_off) + body)
    assert (info.width_px, info.height_px, info.dpi_x) == (200, 100, 102)


def test_svg_sizes_from_units_and_viewbox():
    def size(svg: bytes):
        i = images.inspect(svg)
        return i.width_px, i.height_px

    assert size(b'<svg xmlns="http://www.w3.org/2000/svg" width="2in" height="1in"/>') == (192, 96)
    assert size(b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 100"/>') == (400, 100)
    assert size(b'<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg" width="200" '
                b'viewBox="0 0 400 100"/>') == (200, 50)
    with pytest.raises(images.UnsupportedImage):
        images.inspect(b"<svg nope")


def test_svg_needs_fallback_and_is_stored_with_it():
    doc = WordprocessingDocument.new()
    with pytest.raises(images.UnsupportedImage):
        doc.add_picture(asset("SVG"))
    with pytest.raises(ValueError):
        doc.add_picture(asset("PNG"), fallback=asset("PNG"))
    with pytest.raises(ValueError):
        doc.add_picture(asset("SVG"), fallback=asset("SVG"))
    run = doc.add_picture(asset("SVG"), fallback=asset("PNG"), width=Cm(4))
    blip = run.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    svg = blip.find(".//{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip")
    assert svg is not None
    rels = doc.main_part.rels
    rid_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    assert rels[blip.get(rid_attr)].target.endswith(".png")
    assert rels[svg.get(rid_attr)].target.endswith(".svg")
    assert not doc.validate(target="Office2007")


def test_svg_in_excel_and_powerpoint():
    book = SpreadsheetDocument.new()
    book.add_picture("Sheet1", asset("SVG"), fallback=asset("PNG"))
    deck = PresentationDocument.new()
    deck.add_picture(deck.add_slide("Blank"), asset("SVG"), fallback=asset("PNG"))
    for d in (book, deck):
        assert not d.validate(target="Office2007")
        assert any(p.content_type == "image/svg+xml" for p in reopen(d).package.parts.values())


# -- formatting --------------------------------------------------------------------------------
def test_runs_paragraph_format_and_resolution():
    doc = WordprocessingDocument.new()
    doc.add_style("Note", italic=True, color="595959", space_before=6, indent_left=36)
    doc.add_style("Code Char", kind="character", font="Courier New", size=10)
    p = doc.add_paragraph("", style="Note")
    bold = doc.add_run(p, "Warning: ", bold=True, color="#c00000")
    code = doc.add_run(p, "x", style="Code Char")
    doc.format_paragraph(p, align="center", space_after=12, first_line=-18,
                         line_spacing=LineSpacing.exactly(14))
    fmt = doc.formatting()
    f = fmt.run(bold)
    assert (f.bold, f.italic, f.color, f.style) == (True, True, "C00000", None)
    pf = fmt.paragraph(p)
    assert (pf.style, pf.align, pf.space_before, pf.space_after) == ("Note", "center", 6.0, 12.0)
    assert (pf.indent_left, pf.first_line) == (36.0, -18.0)
    assert pf.line_spacing == LineSpacing("exact", 14.0)
    g = fmt.run(code)
    assert (g.font, g.size, g.color, g.bold, g.style) == ("Courier New", 10.0, "595959", False,
                                                         "Code Char")
    assert doc.effective_format(p, code) == g
    assert not doc.validate(target="Office2007")


def test_defaults_and_theme_fonts_resolve():
    doc = WordprocessingDocument.new()
    fmt = doc.formatting()
    p = doc.add_paragraph("x")
    body, mark = doc.formatting().paragraph(p), doc.effective_format(p)
    assert (mark.font, mark.size, body.style, body.space_after) == ("Calibri", 11.0, "Normal", 8.0)
    h = doc.add_heading("H", 1)
    fmt = doc.formatting()
    assert fmt.mark(h).font == "Calibri Light" and fmt.paragraph(h).keep_with_next is True


def test_toggle_off_and_errors():
    doc = WordprocessingDocument.new()
    doc.add_style("Loud", bold=True)
    p = doc.add_paragraph("", style="Loud")
    r = doc.add_run(p, "quiet", bold=False)
    assert doc.effective_format(p, r).bold is False
    with pytest.raises(TypeError):
        doc.add_run(p, "x", boldd=True)
    with pytest.raises(ValueError):
        doc.add_style("Loud")
    with pytest.raises(TypeError):
        doc.add_style("Char", kind="character", align="center")
    sid = doc.add_style("Loud 2!")  # id is letters and digits
    assert sid == "Loud2"


# -- properties -----------------------------------------------------------------------------------
@pytest.mark.parametrize("cls", [WordprocessingDocument, SpreadsheetDocument,
                                 PresentationDocument])
def test_extended_and_custom_properties(cls):
    doc = cls.new()
    assert doc.extended_properties.company is None and len(doc.custom_properties) == 0
    doc.extended_properties.company = "Example Ltd"
    doc.extended_properties.doc_security = 0
    c = doc.custom_properties
    c["Project"], c["Approved"], c["Budget"], c["Count"] = "Apollo", True, 1250.5, 7
    c["When"] = dt.datetime(2026, 11, 14, 12, tzinfo=dt.timezone.utc)
    c["Project"] = "Artemis"  # replace keeps one entry
    back = reopen(doc)
    assert back.extended_properties.company == "Example Ltd"
    assert back.extended_properties.doc_security == 0
    assert back.custom_properties.as_dict() == {
        "Project": "Artemis", "Approved": True, "Budget": 1250.5, "Count": 7,
        "When": dt.datetime(2026, 11, 14, 12, tzinfo=dt.timezone.utc)}
    del back.custom_properties["Count"]
    assert "Count" not in back.custom_properties
    assert not doc.validate(target="Office2007")


def test_custom_property_errors():
    c = WordprocessingDocument.new().custom_properties
    with pytest.raises(TypeError):
        c["x"] = [1]
    with pytest.raises(ValueError):
        c["x"] = 2 ** 40
    with pytest.raises(KeyError):
        _ = c["missing"]
    with pytest.raises(AttributeError):
        WordprocessingDocument.new().extended_properties.nonsense = 1


def test_untouched_documents_get_no_property_parts():
    doc = WordprocessingDocument.new()
    _ = doc.extended_properties.company, doc.custom_properties.get("x")
    names = set(reopen(doc).package.parts)
    assert "/docProps/app.xml" not in names and "/docProps/custom.xml" not in names


# -- fields ---------------------------------------------------------------------------------------
def test_field_tokens():
    assert tokens('DOCPROPERTY "My Prop" \\* Upper') == ["DOCPROPERTY", "My Prop", "\\*", "Upper"]
    assert tokens('CREATEDATE \\@"d\\ MMMM\\ yyyy"') == ["CREATEDATE", "\\@", "d MMMM yyyy"]
    assert tokens("SEQ Figure \\r 5 \\h") == ["SEQ", "Figure", "\\r", "5", "\\h"]


def test_date_pictures():
    d = dt.datetime(2026, 11, 4, 15, 7, 9)
    assert format_date(d, "d MMMM yyyy") == "4 November 2026"
    assert format_date(d, "dddd, dd/MM/yy HH:mm") == "Wednesday, 04/11/26 15:07"
    assert format_date(d, "h:mm AM/PM 'on' MMM d") == "3:07 PM on Nov 4"


def test_captions_renumber_and_references_follow():
    doc = WordprocessingDocument.new()
    first = doc.add_caption("One")
    p = doc.add_paragraph("See ")
    doc.add_cross_reference(p, first)
    doc.add_cross_reference(p, first, what="position")
    assert teletype.extractText(p) == "See Figure 1above"
    doc.add_caption("Zero", before=first.paragraph)
    assert first.number == "2" and teletype.extractText(p) == "See Figure 2above"
    table = doc.add_caption("T", label="Table")
    assert table.number == "1"
    assert teletype.extractText(first.paragraph) == "Figure 2: One"
    assert doc.style_id("Caption") == "Caption"
    assert not doc.validate(target="Office2007")


def test_seq_switches_and_ref_errors():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("")
    for code in ("SEQ X \\* ROMAN", "SEQ X \\* alphabetic", "SEQ X \\c", "SEQ X \\r 10",
                 "SEQ X \\h", "REF nowhere", "PAGEREF nowhere"):
        doc.add_field(p, code, result="?")
    doc.update_fields()
    assert [f.result for f in doc.fields()] == [
        "I", "b", "2", "10", "", "Error! Reference source not found.",
        "Error! Bookmark not defined."]


def test_property_fields_and_dates():
    doc = WordprocessingDocument.new()
    doc.properties.title = "T"
    doc.properties.created = dt.datetime(2026, 1, 2, 9, tzinfo=dt.timezone.utc)
    doc.extended_properties.company = "Co"
    doc.custom_properties["Flag"] = False
    doc.custom_properties["Whole"] = 3.0
    p = doc.add_paragraph("")
    for code in ("TITLE \\* Upper", 'DOCPROPERTY "Company" \\* Lower', "DOCPROPERTY Flag",
                 "DOCPROPERTY Whole", 'CREATEDATE \\@ "yyyy-MM-dd"', 'DATE \\@ "d MMM"',
                 "AUTHOR"):
        doc.add_field(p, code, result="old")
    doc.update_fields(now=dt.datetime(2026, 3, 4))
    assert [f.result for f in doc.fields()] == ["T", "co", "N", "3", "2026-01-02",
                                                   "4 Mar", "old"]  # no author set: kept


def test_page_estimates_from_breaks_and_rendered_breaks():
    doc = WordprocessingDocument.new()
    cap = doc.add_caption("On page 1")
    doc.add_page_break()
    doc.add_paragraph("page 2")
    doc.add_section()
    later = doc.add_caption("On page 3")
    p = doc.add_paragraph("")
    doc.add_cross_reference(p, cap, what="page")
    doc.add_cross_reference(p, later, what="page")
    assert [f.result for f in doc.fields() if f.name == "PAGEREF"] == ["1", "3"]
    # a Word-laid-out file: lastRenderedPageBreak markers win
    doc2 = WordprocessingDocument.new()
    doc2.body.add(w.P(w.R(w.LastRenderedPageBreak(), w.T("p2"))))
    c2 = doc2.add_caption("x")
    q = doc2.add_paragraph("")
    doc2.add_cross_reference(q, c2, what="page")
    assert doc2.fields()[-1].result == "2"


def test_complex_results_spanning_paragraphs_are_left_alone():
    doc = WordprocessingDocument.new()
    doc.add_heading("A", 1)
    doc.add_toc()
    before = [f.result for f in doc.fields()]
    doc.update_fields()
    assert [f.result for f in doc.fields()] == before


# -- content controls ---------------------------------------------------------------------------------
def form():
    doc = WordprocessingDocument.new()
    doc.add_text_control(doc.add_paragraph("Name: "), tag="name", placeholder="Your name")
    doc.add_dropdown(doc.add_paragraph("Dept: "), tag="dept", items=["Sales", ("Finance", "FIN")])
    doc.add_dropdown(doc.add_paragraph("Office: "), tag="office", items=["A", "B"], combo=True)
    doc.add_date_control(doc.add_paragraph("Start: "), tag="start", fmt="dd/MM/yyyy")
    doc.add_checkbox(doc.add_paragraph("OK: "), tag="ok")
    return doc


def test_controls_read_fill_and_round_trip():
    doc = form()
    assert [(c.tag, c.type, c.value) for c in doc.content_controls()] == [
        ("name", "text", None), ("dept", "dropdown", None), ("office", "combo", None),
        ("start", "date", None), ("ok", "checkbox", False)]
    assert doc.fill_controls({"name": "Ada", "dept": "FIN", "office": "Lisbon",
                              "start": dt.date(2026, 11, 14), "ok": True}) == 5
    back = {c.tag: c for c in reopen(doc).content_controls()}
    assert back["name"].value == "Ada" and not back["name"].placeholder
    assert back["dept"].value == "Finance" and back["office"].value == "Lisbon"
    assert back["start"].value == "14/11/2026" and back["start"].date == dt.date(2026, 11, 14)
    assert back["ok"].value is True
    assert not doc.validate(target="Office2007")
    # placeholder style is dropped once filled
    sdt = next(doc.document.iter(W + "sdt"))
    assert sdt.find(f".//{W}rStyle") is None


def test_fill_errors_and_block_level_controls():
    doc = form()
    with pytest.raises(ValueError):
        doc.fill_controls({"dept": "Marketing"})
    with pytest.raises(KeyError):
        doc.fill_controls({"missing": 1})
    assert doc.fill_controls({"missing": 1}, strict=False) == 0
    # a block-level control as Word templates have it
    block = w.Sdt(w.SdtPr(w.Tag(val="body"), w.Id(val=5)),
                  w.SdtContent(w.P(w.PPr(w.Jc(val="center")), w.R(w.T("old 1"))),
                               w.P(w.R(w.T("old 2")))))
    doc.body.add(block)
    doc.fill_controls({"body": "new"})
    paras = list(block.first("w:sdtContent").iterchildren(W + "p"))
    assert len(paras) == 1 and teletype.extractText(paras[0]) == "new"
    assert paras[0].first("w:pPr").first("w:jc") is not None


def test_checkbox_declares_w14_ignorable():
    doc = form()
    buf = io.BytesIO()
    doc.save(buf)
    import zipfile

    xml = zipfile.ZipFile(buf).read("word/document.xml").decode()
    assert 'mc:Ignorable="w14"' in xml and "w14:checkbox" in xml
