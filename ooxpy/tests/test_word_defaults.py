"""Word's default parts, headings/paragraph helpers, document properties."""

import datetime as dt
import io
import zipfile

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, open_document
from oox.ns import w

docx = pytest.importorskip("docx")


def test_defaults_are_valid_for_office_2007():
    doc = WordprocessingDocument.new()
    assert doc.validate(target="Office2007") == {}


def test_headings_and_styles(tmp_path):
    doc = WordprocessingDocument.new()
    doc.add_heading("Title text", 0)
    doc.add_heading("Chapter", 1)
    doc.add_paragraph("Body\twith a tab")
    doc.add_heading("Section", 3)
    doc.add_paragraph("item", style="List Paragraph")
    assert doc.style_id("heading 1") == doc.style_id("Heading1") == "Heading1"
    with pytest.raises(KeyError, match="no style"):
        doc.add_paragraph("x", style="Nonexistent")
    with pytest.raises(ValueError):
        doc.add_heading("x", 10)
    # paragraphs go before the final section properties
    assert doc.body[-1].tag_name == "w:sectPr"
    path = tmp_path / "h.docx"
    doc.save(path)
    d = docx.Document(str(path))
    assert [p.style.name for p in d.paragraphs] == \
        ["Title", "Heading 1", "Normal", "Heading 3", "List Paragraph"]
    assert d.styles["Heading 1"].font.size.pt == 16
    assert d.paragraphs[2].text == "Body\twith a tab"


@pytest.mark.parametrize("page,landscape,size_mm", [("A4", False, (210, 297)),
                                                    ("Letter", False, (216, 279)),
                                                    ("A4", True, (297, 210))])
def test_page_setup(page, landscape, size_mm, tmp_path):
    path = tmp_path / "p.docx"
    WordprocessingDocument.new(page=page, landscape=landscape).save(path)
    s = docx.Document(str(path)).sections[0]
    assert (round(s.page_width.mm), round(s.page_height.mm)) == size_mm


def test_minimal_documents_have_no_styles():
    doc = WordprocessingDocument.new(minimal=True)
    with pytest.raises(KeyError, match="minimal=True"):
        doc.add_heading("x")


@pytest.mark.parametrize("make", [WordprocessingDocument.new, SpreadsheetDocument.new,
                                  PresentationDocument.new])
def test_core_properties_all_formats(make):
    doc = make()
    assert doc.properties.title is None  # nothing created yet
    assert all("docProps" not in n for n in doc.package.parts)
    when = dt.datetime(2026, 11, 14, 11, 30, tzinfo=dt.timezone(dt.timedelta(hours=2)))
    doc.properties.title = "Report"
    doc.properties.creator = "Ada"
    doc.properties.keywords = "oox, test"
    doc.properties.created = when
    buf = io.BytesIO()
    doc.save(buf)
    again = open_document(io.BytesIO(buf.getvalue()))
    assert again.properties.as_dict() == {"title": "Report", "creator": "Ada",
                                          "keywords": "oox, test",
                                          "created": when.astimezone(dt.timezone.utc)}
    xml = zipfile.ZipFile(io.BytesIO(buf.getvalue())).read("docProps/core.xml")
    assert b'xsi:type="dcterms:W3CDTF">2026-11-14T09:30:00Z<' in xml
    again.properties.keywords = None
    assert again.properties.keywords is None
    with pytest.raises(AttributeError, match="unknown property"):
        again.properties.author = "x"  # Dublin Core calls it "creator"


def test_properties_visible_to_python_docx(tmp_path):
    doc = WordprocessingDocument.new()
    doc.properties.title = "T"
    doc.properties.creator = "C"
    doc.properties.last_modified_by = "M"
    doc.properties.revision = 3
    path = tmp_path / "p.docx"
    doc.save(path)
    cp = docx.Document(str(path)).core_properties
    assert (cp.title, cp.author, cp.last_modified_by, cp.revision) == ("T", "C", "M", 3)


def test_body_sectpr_replace():
    doc = WordprocessingDocument.new()
    doc.body.add(w.SectPr(w.PgSz(w=12240, h=15840)), replace=True)
    assert len(doc.body.findall("{*}sectPr")) == 1


def test_properties_read_libreoffice_relationship_variant():
    doc = WordprocessingDocument.new()
    doc.properties.title = "via LibreOffice"
    rel = next(r for r in doc.package.root.rels.items.values() if "core-properties" in r.reltype)
    rel.reltype = ("http://schemas.openxmlformats.org/officedocument/2006/relationships/"
                   "metadata/core-properties")  # what LibreOffice writes
    buf = io.BytesIO()
    doc.save(buf)
    assert open_document(io.BytesIO(buf.getvalue())).properties.title == "via LibreOffice"
