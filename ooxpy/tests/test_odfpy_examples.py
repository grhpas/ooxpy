"""odfpy's own example scripts (github.com/eea/odfpy, examples/), translated to oox.

Each test quotes the odfpy original in its comments and keeps the same
shape line for line.  Where OOXML forces a difference, a comment says why.
Outputs are checked with python-docx as an independent reader.
"""

import zipfile

import pytest

from oox import teletype
from oox.ns import w
from oox.odfcompat import OpenDocumentText, load

docx = pytest.importorskip("docx")


def test_helloworld(tmp_path):
    # from odf.opendocument import OpenDocumentText
    # from odf.text import P
    # textdoc = OpenDocumentText()
    # p = P(text="Hello World!")
    # textdoc.text.addElement(p)
    # textdoc.save("helloworld", True)
    textdoc = OpenDocumentText()
    p = w.P()
    teletype.addTextToElement(p, "Hello World!")  # OOXML: text lives in w:r/w:t, not w:p
    textdoc.text.addElement(p)
    textdoc.save(tmp_path / "helloworld", True)

    assert docx.Document(str(tmp_path / "helloworld.docx")).paragraphs[0].text == "Hello World!"


def test_loadsave(tmp_path):
    # from odf.opendocument import load
    # doc = load(infile)
    # outfile = infile[:-4] + "-bak" + infile[-4:]
    # doc.save(outfile)
    infile = str(tmp_path / "in.docx")
    src = OpenDocumentText()
    src.text.addElement(w.P(w.R(w.T("keep me"))))
    src.save(infile)

    doc = load(infile)
    outfile = infile[:-5] + "-bak" + infile[-5:]  # ".docx" is 5 chars, ".odt" was 4
    doc.save(outfile)

    a, b = zipfile.ZipFile(infile), zipfile.ZipFile(outfile)
    assert a.read("word/document.xml") == b.read("word/document.xml")  # untouched -> identical


def test_manualpagebreak(tmp_path):
    # withbreak = Style(name="WithBreak", parentstylename="Standard", family="paragraph")
    # withbreak.addElement(ParagraphProperties(breakbefore="page"))
    # textdoc.automaticstyles.addElement(withbreak)
    # p = P(text=u'First paragraph')
    # textdoc.text.addElement(p)
    # p = P(stylename=withbreak, text=u'Second paragraph')
    # textdoc.text.addElement(p)
    # textdoc.save("pagebreak_odfpy.odt")
    textdoc = OpenDocumentText()
    withbreak = w.Style(w.Name(val="WithBreak"), type="paragraph", styleId="WithBreak")
    withbreak.addElement(w.PPr(w.PageBreakBefore()))
    textdoc.styles.addElement(withbreak)  # OOXML has no automatic styles: use the styles part
    p = w.P()
    teletype.addTextToElement(p, "First paragraph")
    textdoc.text.addElement(p)
    p = w.P(w.PPr(w.PStyle(val=withbreak)))  # style element as value -> its w:styleId
    teletype.addTextToElement(p, "Second paragraph")
    textdoc.text.addElement(p)
    textdoc.save(tmp_path / "pagebreak.docx")

    d = docx.Document(str(tmp_path / "pagebreak.docx"))
    assert [p.text for p in d.paragraphs] == ["First paragraph", "Second paragraph"]
    assert d.paragraphs[1].style.name == "WithBreak"
    assert d.paragraphs[1].paragraph_format.page_break_before is None  # inherited, not direct
    assert d.styles["WithBreak"].paragraph_format.page_break_before is True


PASSWD = """\
root:x:0:0:root:/root:/bin/bash
daemon:x:1:1:daemon:/usr/sbin:/usr/sbin/nologin
ada:x:1000:1000:Ada Lovelace:/home/ada:/bin/zsh
"""


def test_passwd_as_table(tmp_path):
    # tablecontents = Style(name="Table Contents", family="paragraph")
    # tablecontents.addElement(ParagraphProperties(numberlines="false", linenumber="0"))
    # textdoc.styles.addElement(tablecontents)
    # widthshort = Style(name="Wshort", family="table-column")  ... columnwidth="1.7cm"
    # widthwide  = Style(name="Wwide",  family="table-column")  ... columnwidth="1.5in"
    # table = Table()
    # table.addElement(TableColumn(numbercolumnsrepeated=4, stylename=widthshort))
    # table.addElement(TableColumn(numbercolumnsrepeated=3, stylename=widthwide))
    # for line in f:
    #     rec = line.strip().split(":")
    #     tr = TableRow(); table.addElement(tr)
    #     for val in rec:
    #         tc = TableCell(); tr.addElement(tc)
    #         p = P(stylename=tablecontents, text=val); tc.addElement(p)
    # textdoc.text.addElement(table)
    textdoc = OpenDocumentText()
    tablecontents = w.Style(w.Name(val="Table Contents"), type="paragraph", styleId="TableContents")
    tablecontents.addElement(w.PPr(w.SuppressLineNumbers()))
    textdoc.styles.addElement(tablecontents)
    short, wide = 964, 2160  # column widths go in w:tblGrid, in twips (1.7cm, 1.5in)

    table = w.Tbl(w.TblPr())  # w:tblPr is mandatory in OOXML
    grid = w.TblGrid()  # OOXML: column widths live in w:tblGrid, not column styles
    for width in [short] * 4 + [wide] * 3:  # numbercolumnsrepeated has no OOXML twin
        grid.addElement(w.GridCol(w=width))
    table.addElement(grid)
    for line in PASSWD.splitlines():
        rec = line.strip().split(":")
        tr = w.Tr()
        table.addElement(tr)
        for val in rec:
            tc = w.Tc()
            tr.addElement(tc)
            p = w.P(w.PPr(w.PStyle(val=tablecontents)))
            teletype.addTextToElement(p, val)
            tc.addElement(p)
    textdoc.text.addElement(table)
    textdoc.save(tmp_path / "passwd.docx")

    assert textdoc.document.validate() == []
    assert len(textdoc.getElementsByType(w.Tc)) == 21
    d = docx.Document(str(tmp_path / "passwd.docx"))
    assert [c.text for c in d.tables[0].rows[2].cells] == PASSWD.splitlines()[2].split(":")
