"""Package helpers: relationships as a mapping, part lookups, image_of, oox.ns."""

from __future__ import annotations

import io

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, open_document
from oox.document import RT_IMAGE, ImageData
from oox.errors import PackageError
from oox.ns import MC, PROAM, A, R, W, qn, w, x
from oox.opc import Relationship

PNG = (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00"
       b"\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0\x00\x00\x03\x01\x01\x00\xc9\xfe"
       b"\x92\xef\x00\x00\x00\x00IEND\xaeB`\x82")


def reopen(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


def test_namespace_constants_and_qn():
    assert W == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    assert A.endswith("/drawingml/2006/main") and R.endswith("/relationships")
    assert MC.endswith("markup-compatibility/2006") and PROAM.startswith("http")
    assert qn("w:p") == f"{{{W}}}p" and qn("plain") == "plain"
    assert w.P().tag == qn("w:p")
    with pytest.raises(KeyError):
        qn("nope:p")


def test_relationships_are_a_mapping():
    doc = WordprocessingDocument.new()
    rels = doc.main_part.rels
    rid = rels.add(RT_IMAGE, "https://example.com/x.png", external=True)
    assert rid in rels and rels[rid].external and rels.get("rId999") is None
    assert len(rels) == len(list(rels)) == len(dict(rels.items()))
    assert all(isinstance(r, Relationship) for r in rels.values())
    with pytest.raises(ValueError):
        rels["rId77"] = Relationship("rId78", RT_IMAGE, "x.png")
    del rels[rid]
    assert rid not in rels and "Relationships(" in repr(rels)


def test_related_part_blob_and_named_parts():
    doc = WordprocessingDocument.new()
    styles = doc.styles_part
    assert styles is not None and styles.element is doc.styles
    rid = next(r.rid for r in doc.main_part.rels.values() if r.reltype.endswith("/styles"))
    assert doc.main_part.related_part(rid) is styles
    assert styles.blob.startswith(b"<?xml") and styles.package is doc.package
    assert (doc.settings_part, doc.theme_part, doc.font_table_part) != (None, None, None)
    assert doc.numbering_part is doc.footnotes_part is doc.comments_part is None
    doc.ensure_numbering()
    assert doc.numbering_part is not None and doc.endnotes_part is None
    with pytest.raises(KeyError):
        doc.main_part.related_part("rId999")
    ext = doc.main_part.rels.add(RT_IMAGE, "https://example.com/x.png", external=True)
    with pytest.raises(PackageError, match="external"):
        doc.main_part.related_part(ext)
    gone = doc.main_part.rels.add(RT_IMAGE, "media/missing.png")
    with pytest.raises(PackageError, match="missing"):
        doc.main_part.related_part(gone)
    back = reopen(doc)
    assert back.styles_part.package is back.package and back.styles_part.blob == styles.blob

    book = SpreadsheetDocument.new()
    assert book.styles_part is None and book.shared_strings_part is None
    book.set_value("Sheet1", "A1", "text", style={"bold": True})
    assert book.styles_part is not None and book.styles_part.element.tag == qn("x:styleSheet")
    sst = book.shared_strings_part or book.add_part("SharedStringTablePart", element=x.Sst())[0]
    assert book.shared_strings_part is sst
    assert book.theme_part is None or book.theme_part.content_type.endswith("theme+xml")


def test_part_of_follows_replaced_trees():
    doc = WordprocessingDocument.new()
    assert doc.part_of(doc.body) is doc.main_part
    assert doc.part_of(doc.styles) is doc.styles_part
    part = doc.styles_part
    part._element = w.Styles()
    assert doc.part_of(part.element) is part
    with pytest.raises(PackageError):
        doc.part_of(w.P())


def test_image_of_embedded_linked_and_vml():
    doc = WordprocessingDocument.new()
    run = doc.add_picture(io.BytesIO(PNG))
    pic = doc.pictures()[0]
    img = doc.image_of(pic)
    assert img == ImageData(PNG, "image/png", img.partname, None)
    assert img.partname.startswith("/word/media/")
    assert doc.image_of(run) == img and doc.image_of(pic.element) == img

    rid = doc.main_part.rels.add(RT_IMAGE, "https://example.com/logo.png", external=True)
    linked = doc.add_picture(io.BytesIO(PNG))
    blip = next(linked.iter(qn("a:blip")))
    del blip.attrib[qn("r:embed")]
    blip.set(qn("r:link"), rid)
    assert doc.image_of(linked) == ImageData(None, None, None, "https://example.com/logo.png")

    vml_rid, _ = doc.add_image(io.BytesIO(PNG))
    from oox.ns import v

    shape = v.Shape(v.Imagedata(r_id=vml_rid))
    doc.add_paragraph().add(w.R(w.Pict(shape)))
    assert doc.image_of(shape).data == PNG

    with pytest.raises(ValueError):
        doc.image_of(w.P())
    bad = v.Imagedata()
    doc.add_paragraph().add(w.R(w.Pict(v.Shape(bad))))
    with pytest.raises(ValueError):
        doc.image_of(bad)


def test_image_of_in_excel_and_powerpoint():
    book = SpreadsheetDocument.new()
    book.add_picture("Sheet1", io.BytesIO(PNG), "B2")
    deck = PresentationDocument.new()
    deck.add_picture(deck.add_slide("Blank"), io.BytesIO(PNG))
    for doc in (book, deck):
        blips = [b for p in doc.package.parts.values() if p.is_xml
                 for b in p.element.iter(qn("a:blip"))]
        assert [doc.image_of(b).data for b in blips] == [PNG]
