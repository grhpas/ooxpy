"""Image inspection and pictures in all three formats."""

import io
import struct

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument
from oox.images import UnsupportedImage, inspect
from oox.units import Cm, Inches, Pt, Px

PIL = pytest.importorskip("PIL.Image")


def make(fmt: str, size=(200, 100), dpi=None) -> bytes:
    buf = io.BytesIO()
    kw = {"dpi": dpi} if dpi else {}
    PIL.new("RGB", size, (68, 114, 196)).save(buf, fmt, **kw)
    return buf.getvalue()


@pytest.mark.parametrize("fmt,ct", [("PNG", "image/png"), ("JPEG", "image/jpeg"),
                                    ("GIF", "image/gif"), ("BMP", "image/bmp")])
def test_inspect_formats(fmt, ct):
    info = inspect(make(fmt, (321, 123)))
    assert (info.content_type, info.width_px, info.height_px) == (ct, 321, 123)


@pytest.mark.parametrize("fmt", ["PNG", "JPEG"])
def test_inspect_reads_resolution(fmt):
    info = inspect(make(fmt, (300, 150), dpi=(150, 150)))
    assert round(info.dpi_x) == 150
    assert info.native_size() == (Inches(2), Inches(1))


def test_progressive_jpeg_and_unsupported():
    buf = io.BytesIO()
    PIL.new("RGB", (64, 32)).save(buf, "JPEG", progressive=True)
    assert inspect(buf.getvalue()).width_px == 64
    svg = inspect(b'<svg xmlns="http://www.w3.org/2000/svg"/>')  # no size: CSS default 300x150
    assert svg.needs_fallback and (svg.width_px, svg.height_px) == (300, 150)
    with pytest.raises(UnsupportedImage):
        inspect(b"not an image")


def test_scaling_keeps_aspect_ratio():
    info = inspect(make("PNG", (200, 100), dpi=(96, 96)))
    assert info.scaled(Cm(10), None) == (Cm(10), Cm(5))
    assert info.scaled(None, Cm(2)) == (Cm(4), Cm(2))
    assert info.scaled(1, 2) == (1, 2)


def test_units():
    assert Inches(1) == 914400 and Cm(1) == 360000 and Pt(1) == 12700 and Px(96) == 914400


def test_docx_pictures(tmp_path):
    docx = pytest.importorskip("docx")
    doc = WordprocessingDocument.new()
    png = make("PNG")
    doc.add_picture(png, width=Cm(6), descr="box")
    doc.add_picture(png, height=Cm(1))  # same bytes -> one media part
    path = tmp_path / "p.docx"
    doc.save(path)
    assert [n for n in doc.package.parts if "media" in n] == ["/word/media/image1.png"]
    assert doc.document.validate() == []
    shapes = docx.Document(str(path)).inline_shapes
    assert [(s.width, s.height) for s in shapes] == [(Cm(6), Cm(3)), (Cm(2), Cm(1))]
    ids = [int(e.get("id")) for e in doc.document.iter("{*}docPr")]
    assert len(set(ids)) == 2  # unique drawing ids


def test_xlsx_picture(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    book = SpreadsheetDocument.new()
    book.add_picture("Sheet1", make("PNG"), anchor="C3", width=Cm(4))
    book.add_picture("Sheet1", make("JPEG"), anchor="H3")
    path = tmp_path / "p.xlsx"
    book.save(path)
    for part in book.package.parts.values():
        if part.is_xml:
            assert part.element.validate() == [], part.partname
    imgs = openpyxl.load_workbook(path).active._images
    assert len(imgs) == 2
    anchor = imgs[0].anchor._from
    assert (anchor.col, anchor.row) == (2, 2)  # C3, zero-based


def test_pptx_picture_shared_between_slides(tmp_path):
    pptx = pytest.importorskip("pptx")
    deck = PresentationDocument.new()
    png = make("PNG")
    for _ in range(2):
        deck.add_picture(deck.add_slide(), png, x=Inches(1), y=Inches(1), width=Inches(3))
    path = tmp_path / "p.pptx"
    deck.save(path)
    assert [n for n in deck.package.parts if "media" in n] == ["/ppt/media/image1.png"]
    pr = pptx.Presentation(str(path))
    pics = [s.shapes[0] for s in pr.slides]
    assert all(p.shape_type == 13 and p.width == Inches(3) for p in pics)
    assert pics[0].image.blob == png


def test_odfpy_add_picture_returns_rid():
    doc = WordprocessingDocument.new()
    rid = doc.addPictureFromString(make("PNG"), "image/png")
    assert doc.main_part.rels[rid].target == "media/image1.png"
    assert struct  # keep import for readers poking at headers
