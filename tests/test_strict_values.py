"""Strict value forms: DrawingML percentages ("50%") and ST_OnOff as xsd:boolean."""

import io
import re
import zipfile

from oox import PresentationDocument, WordprocessingDocument, open_document
from oox.ns import a, w
from tools.officecheck.assets import image


def _strict_bytes(doc) -> bytes:
    doc.conformance = "strict"
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _part(data: bytes, name: str) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return z.read(name).decode("utf-8")


def test_percentages_written_as_strict_and_read_back():
    deck = PresentationDocument.new()
    s = deck.add_slide("Picture with Caption")
    pic = deck.set_placeholder_picture(s, "pic", image("PNG"))
    src = pic.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect")
    cut = int(src.get("l"))
    sp = deck.add_shape(s, "rect", x=0, y=0, width=100, height=100)
    fill = sp.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
    fill.add(a.LumMod(val=75000))
    data = _strict_bytes(deck)
    slide = _part(data, "ppt/slides/slide1.xml")
    assert f'l="{cut / 1000:g}%"' in slide and 'val="75%"' in slide
    assert "purl.oclc.org/ooxml/drawingml/main" in slide
    # in memory nothing changed
    assert src.get("l") == str(cut)
    back = open_document(io.BytesIO(data))
    assert back.conformance == "strict"
    src2 = back.slides[0].find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect")
    assert src2.get("l") == str(cut)                      # thousandths again
    assert back.validate() == {}
    again = io.BytesIO()
    back.save(again)                                      # read, unchanged: byte-identical parts
    assert _part(again.getvalue(), "ppt/slides/slide1.xml") == slide
    back.conformance = "transitional"                     # convert
    tbuf = io.BytesIO()
    back.save(tbuf)
    trans = _part(tbuf.getvalue(), "ppt/slides/slide1.xml")
    assert f'l="{cut}"' in trans and 'val="75000"' in trans


def test_unknown_percentage_attributes_round_trip_as_read():
    """A percentage on an integer attribute oox doesn't list still converts
    on reading and goes back on writing (the part remembers it)."""
    deck = PresentationDocument.new()
    s = deck.add_slide()
    sp = deck.add_shape(s, "rect", x=0, y=0, width=100, height=100)
    fill = sp.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
    fill.add(a.HueOff(val=600000))                        # an angle, not in KNOWN_PERCENT
    data = _strict_bytes(deck)
    assert 'val="600000"' in _part(data, "ppt/slides/slide1.xml")    # not turned into "%"
    # a producer writing a percentage there: converted on reading, restored on writing
    patched = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(patched, "w") as zout:
        for item in zin.infolist():
            raw = zin.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml":
                raw = raw.replace(b'val="600000"', b'val="12.5%"')
            zout.writestr(item, raw)
    back = open_document(io.BytesIO(patched.getvalue()))
    hue = back.slides[0].find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}hueOff")
    assert hue.get("val") == "12500"
    hue.addprevious(a.LumMod(val=50000))                  # change the part
    out = io.BytesIO()
    back.save(out)
    slide = _part(out.getvalue(), "ppt/slides/slide1.xml")
    assert 'val="12.5%"' in slide and 'val="50%"' in slide


def test_on_off_becomes_boolean_in_strict():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("bold")
    r = p.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
    rpr = w.RPr(w.B())
    r.insert(0, rpr)
    rpr[0].set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "on")
    body = _part(_strict_bytes(doc), "word/document.xml")
    assert re.search(r'<w:b w:val="true"\s*/>', body)
    assert rpr[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "on"
