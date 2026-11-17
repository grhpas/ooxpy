"""A corpus of Word files written by LibreOffice, with ground truth.

    python -m tools.locorpus OUT          # needs LibreOffice (soffice)

The sources are flat OpenDocument files written here, so we know what is in
them: how many sections, which list labels, how many pictures and text boxes.
LibreOffice converts each to .docx, writing its own markup (outline numbering
linked to heading styles, list styles, text boxes as a DrawingML shape plus a
VML copy, page-style changes as section breaks). LibreOffice also exports each
.docx to plain text, which is its own reading of the list labels; the truth
records both, and the test (``tests/test_libreoffice_corpus.py``) checks oox's
readers against them.

The .docx files and ``truth.json`` are committed under
``tests/data/libreoffice/``; rerun this after a LibreOffice upgrade to see if
its output changed.
"""

from __future__ import annotations

import base64
import json
import shutil
import struct
import subprocess
import sys
import tempfile
import zlib
from pathlib import Path

HEAD = """<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
"""


def _png(rgb: bytes, w: int = 32, h: int = 24) -> bytes:
    raw = b"".join(b"\x00" + rgb * w for _ in range(h))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data)))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _image(name: str, rgb: bytes, anchor: str = "as-char") -> str:
    data = base64.b64encode(_png(rgb)).decode()
    return (f'<draw:frame draw:name="{name}" text:anchor-type="{anchor}" svg:width="2cm" '
            f'svg:height="1.5cm"><draw:image><office:binary-data>{data}</office:binary-data>'
            "</draw:image></draw:frame>")


STYLES = """
<office:styles>
 <style:style style:name="Standard" style:family="paragraph"/>
 <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph"
   style:parent-style-name="Standard" style:default-outline-level="1"/>
 <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph"
   style:parent-style-name="Standard" style:default-outline-level="2"/>
 <text:outline-style style:name="Outline">
  <text:outline-level-style text:level="1" style:num-format="1"/>
  <text:outline-level-style text:level="2" style:num-format="1" text:display-levels="2"/>
 </text:outline-style>
 <text:list-style style:name="Steps">
  <text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/>
  <text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"/>
 </text:list-style>
 <text:list-style style:name="Roman">
  <text:list-level-style-number text:level="1" style:num-format="I" style:num-suffix="."
    text:start-value="3"/>
 </text:list-style>
 <text:list-style style:name="Dots">
  <text:list-level-style-bullet text:level="1" text:bullet-char="•"/>
 </text:list-style>
</office:styles>
<office:automatic-styles>
 <style:page-layout style:name="pmPortrait">
  <style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"
    style:print-orientation="portrait"/>
 </style:page-layout>
 <style:page-layout style:name="pmLandscape">
  <style:page-layout-properties fo:page-width="29.7cm" fo:page-height="21cm"
    style:print-orientation="landscape"/>
 </style:page-layout>
 <style:style style:name="PLand" style:family="paragraph" style:parent-style-name="Standard"
   style:master-page-name="Landscape"/>
 <style:style style:name="PPort" style:family="paragraph" style:parent-style-name="Standard"
   style:master-page-name="Standard"/>
</office:automatic-styles>
<office:master-styles>
 <style:master-page style:name="Standard" style:page-layout-name="pmPortrait"/>
 <style:master-page style:name="Landscape" style:page-layout-name="pmLandscape"/>
</office:master-styles>
"""


def _p(text: str, style: str = "Standard") -> str:
    return f'<text:p text:style-name="{style}">{text}</text:p>'


def _items(*texts: str) -> str:
    return "".join(f"<text:list-item>{_p(t)}</text:list-item>" for t in texts)


def _sublist(parent: str, *children: str) -> str:
    return (f"<text:list-item>{_p(parent)}<text:list>{_items(*children)}</text:list>"
            "</text:list-item>")


def _body(*parts: str) -> str:
    return "<office:body><office:text>" + "".join(parts) + "</office:text></office:body>"


DOCUMENTS: dict[str, tuple[str, dict]] = {
    "lists": (_body(
        '<text:h text:style-name="Heading_20_1" text:outline-level="1">Introduction</text:h>',
        '<text:h text:style-name="Heading_20_2" text:outline-level="2">Scope</text:h>',
        '<text:h text:style-name="Heading_20_2" text:outline-level="2">Terms</text:h>',
        '<text:list text:style-name="Steps">',
        _items("Collect"), _sublist("Check", "North", "South"), _items("Report"),
        "</text:list>",
        '<text:list text:style-name="Dots">', _items("Apples", "Pears"), "</text:list>",
        '<text:list text:style-name="Roman">', _items("Third", "Fourth"), "</text:list>",
        '<text:h text:style-name="Heading_20_1" text:outline-level="1">Results</text:h>',
        '<text:list text:style-name="Steps">', _items("Again"), "</text:list>",
    ), {"labels": ["1", "1.1", "1.2", "1.", "2.", "a)", "b)", "3.", "•", "•",
                   "III.", "IV.", "2", "1."]}),
    "sections": (_body(
        _p("Portrait page."),
        _p("Landscape page.", "PLand"),
        _p("Portrait again.", "PPort"),
    ), {"sections": 3, "orientations": ["portrait", "landscape", "portrait"]}),
    "frames": (_body(
        _p("Picture: " + _image("Logo", b"\x2f\x54\x96")),
        _p('<draw:frame draw:name="Box" text:anchor-type="paragraph" svg:width="6cm" '
           'svg:height="2cm"><draw:text-box>'
           + _p("Inside the box.") + "</draw:text-box></draw:frame>After the box."),
        '<table:table table:name="T"><table:table-column table:number-columns-repeated="2"/>'
        "<table:table-row><table:table-cell>" + _p("a") + "</table:table-cell>"
        "<table:table-cell>" + _p("b") + "</table:table-cell></table:table-row></table:table>",
        _p("Second picture: " + _image("Mark", b"\xc0\x00\x00")),
    ), {"pictures": 2, "text_boxes": 1, "tables": 1, "once": ["Inside the box."]}),
}


def find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    raise SystemExit("LibreOffice (soffice) not found")


def _convert(soffice: str, src: Path, fmt: str, outdir: Path, profile: Path) -> Path:
    subprocess.run([soffice, f"-env:UserInstallation=file://{profile}", "--headless",
                    "--convert-to", fmt, "--outdir", str(outdir), str(src)],
                   check=True, capture_output=True, timeout=180)
    return outdir / (src.stem + "." + fmt.split(":")[0])


def build(out: str | Path) -> dict:
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    soffice = find_soffice()
    version = subprocess.run([soffice, "--version"], capture_output=True, text=True, check=False,
                             timeout=60).stdout.strip()
    truth: dict = {"libreoffice": version, "documents": {}}
    with tempfile.TemporaryDirectory() as tmp:
        work, profile = Path(tmp) / "work", Path(tmp) / "profile"
        work.mkdir()
        for name, (body, expected) in DOCUMENTS.items():
            src = work / f"{name}.fodt"
            src.write_text(HEAD + STYLES + body + "</office:document>\n", encoding="utf-8")
            docx = _convert(soffice, src, "docx", work, profile)
            shutil.copy(docx, out / docx.name)
            # LibreOffice's own reading of its .docx, as text: list labels included
            txt = _convert(soffice, out / docx.name, "txt:Text (encoded):UTF8", work, profile)
            lines = [ln for ln in txt.read_text(encoding="utf-8-sig").splitlines() if ln.strip()]
            truth["documents"][docx.name] = {**expected, "libreoffice_text": lines}
    (out / "truth.json").write_text(json.dumps(truth, indent=1, ensure_ascii=False) + "\n",
                                    encoding="utf-8")
    return truth


if __name__ == "__main__":
    result = build(sys.argv[1] if len(sys.argv) > 1 else "tests/data/libreoffice")
    print(json.dumps(result, indent=1, ensure_ascii=False))
