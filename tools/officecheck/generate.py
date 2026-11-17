"""Generate the Office check corpus: fixed files + manifest + manual checklist.

    python -m tools.officecheck.generate OUTDIR

Output (copy the whole folder to each Office machine):

    OUTDIR/manifest.json        cases, file hashes, probes (read by check_office.py)
    OUTDIR/CHECKLIST.md         the same cases as a manual checklist
    OUTDIR/manual-results.csv   blank sheet for recording a manual run
    OUTDIR/check_office.py      the Windows checker (copied in: the folder is self-contained)
    OUTDIR/report.py            merges results into REPORT.md
    OUTDIR/word|excel|powerpoint/*.docx|xlsx|pptx

The corpus is deterministic: the same oox version produces identical bytes,
so a result recorded on one machine is comparable with any other.

Each case carries *probes*: facts that must hold after Office opens the file
(table counts, computed cell values, slide texts...).  Office sometimes
repairs a file silently and drops content; probes catch that.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import shutil
import sys
import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

import oox
from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, open_document, teletype
from oox.ns import w, w14
from oox.spreadsheet import Formula
from oox.units import Cm, Emu, Inches

from .assets import image

GREEK = "Καλημέρα κόσμε"  # non-ASCII text must survive every app


@dataclass
class Case:
    id: str
    app: str  # word | excel | powerpoint
    title: str
    look_for: str  # what a human should see
    build: Callable[[], bytes]
    probes: list[dict] = field(default_factory=list)
    requires: str | None = None  # optional Python package needed to build
    expect_status: tuple[str, ...] = ("ok",)  # controls expect a prompt/error


def _bytes(doc) -> bytes:
    return doc.package.to_bytes()


# -- Word ---------------------------------------------------------------------------
def docx_minimal() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    doc.body.add(w.P(w.R(w.T("Minimal package: one paragraph."))))
    return _bytes(doc)


def docx_formatting() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    doc.body.add(
        w.P(w.PPr(w.Jc(val="center"), w.Spacing(after=240)),
            w.R(w.RPr(w.B(), w.Color(val="1F3864"), w.Sz(val=40)), w.T("Centered bold title"))),
        w.P(w.R(w.T("Normal, ")), w.R(w.RPr(w.I()), w.T("italic, ")),
            w.R(w.RPr(w.U(val="single")), w.T("underlined, ")),
            w.R(w.RPr(w.Color(val="C00000")), w.T("red"))),
        w.P(w.R(w.T("   three leading spaces and a trailing one "))),
        w.P(w.R(w.T(GREEK))),
    )
    p = w.P()
    teletype.addTextToElement(p, "Tab:\tafter tab\nSecond line (line break)")
    doc.body.add(p)
    return _bytes(doc)


def docx_table() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    border = dict(val="single", sz=4, space=0, color="808080")
    widths = [2400, 2400, 2400]
    tbl = w.Tbl(
        w.TblPr(w.TblW(w=sum(widths), type="dxa"),
                w.TblBorders(w.Top(**border), w.Left(**border), w.Bottom(**border),
                             w.Right(**border), w.InsideH(**border), w.InsideV(**border))),
        w.TblGrid(*(w.GridCol(w=x) for x in widths)),
    )
    for r, row in enumerate([["Region", "Units", "Price"], ["North", "120", "9.50"],
                             ["South", "95", "11.00"]]):
        tbl.add(w.Tr(*(
            w.Tc(w.TcPr(w.TcW(w=widths[c], type="dxa"),
                        *([w.Shd(val="clear", color="auto", fill="D9E2F3")] if r == 0 else [])),
                 w.P(w.R(*([w.RPr(w.B())] if r == 0 else []), w.T(text))))
            for c, text in enumerate(row))))
    doc.body.add(w.P(w.R(w.T("A 3x3 table with borders; header row shaded and bold:"))), tbl,
                 w.P())
    return _bytes(doc)


def docx_pictures() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    for fmt, width in (("PNG", Cm(6)), ("JPEG", Cm(6)), ("GIF", Cm(3)), ("BMP", Cm(2))):
        doc.body.add(w.P(w.R(w.T(f"{fmt}:"))))
        doc.add_picture(image(fmt), width=width, descr=f"{fmt} test image")
    para = w.P(w.R(w.T("Inline with text: ")))
    doc.body.add(para)
    doc.add_picture(image("PNG"), height=Cm(0.6), paragraph=para)
    para.add(w.R(w.T(" (same PNG, stored once)")))
    return _bytes(doc)


def docx_styles_pagebreak() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    style = w.Style(w.Name(val="WithBreak"), w.PPr(w.PageBreakBefore()),
                    type="paragraph", styleId="WithBreak")
    doc.ensure_styles().add(style)
    doc.body.add(w.P(w.R(w.T("Page 1"))),
                 w.P(w.PPr(w.PStyle(val=style)), w.R(w.T("Page 2: style 'WithBreak'"))))
    return _bytes(doc)


def docx_w14_extension() -> bytes:
    """Office 2010 markup (w14) -> must be declared mc:Ignorable for Word 2007."""
    doc = WordprocessingDocument.new(minimal=True)
    doc.body.add(w.P(w.R(w.RPr(w14.Ligatures(val="standard")),
                         w.T("Text with a Word 2010 ligature setting (w14)."))))
    return _bytes(doc)


def docx_odfpy_passwd() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    style = w.Style(w.Name(val="Table Contents"), w.PPr(w.SuppressLineNumbers()),
                    type="paragraph", styleId="TableContents")
    doc.ensure_styles().add(style)
    grid = w.TblGrid(*(w.GridCol(w=x) for x in [964] * 4 + [2160] * 3))
    tbl = w.Tbl(w.TblPr(), grid)
    for line in ["root:x:0:0:root:/root:/bin/bash", "ada:x:1000:1000:Ada:/home/ada:/bin/zsh"]:
        tr = w.Tr()
        tbl.add(tr)
        for val in line.split(":"):
            p = w.P(w.PPr(w.PStyle(val=style)))
            teletype.addTextToElement(p, val)
            tr.add(w.Tc(p))
    doc.body.add(tbl, w.P())
    return _bytes(doc)


def docx_edit_word_template() -> bytes:
    import docx  # python-docx's template was authored in Word

    path = Path(docx.__file__).parent / "templates" / "default.docx"
    doc = WordprocessingDocument.open(path)
    doc.body.add(w.P(w.R(w.T("Appended by oox to a Word-authored file."))))
    doc.add_picture(image("PNG"), width=Cm(4))
    return _bytes(doc)


# -- Excel --------------------------------------------------------------------------
def xlsx_types() -> bytes:
    book = SpreadsheetDocument.new("Types")
    book.append_row("Types", ["int", 42])
    book.append_row("Types", ["negative float", -1234.5678])
    book.append_row("Types", ["large", 12345678901])
    book.append_row("Types", ["bool", True])
    book.append_row("Types", ["text", "  padded text  "])
    book.append_row("Types", ["greek", GREEK])
    book.append_row("Types", ["multi-line", "line 1\nline 2"])
    return _bytes(book)


def xlsx_formulas() -> bytes:
    book = SpreadsheetDocument.new("Calc")
    book.append_row("Calc", ["Item", "Qty", "Price", "Total"])
    book.append_row("Calc", ["Pens", 3, 1.5, Formula("B2*C2")])
    book.append_row("Calc", ["Paper", 2, 4.25, Formula("B3*C3")])
    book.append_row("Calc", ["Sum", Formula("SUM(B2:B3)"), None, Formula("SUM(D2:D3)")])
    book.append_row("Calc", ["Text", Formula('"oo"&"x"'), None, Formula("IF(D4>10,\"big\",\"small\")")])
    return _bytes(book)


def xlsx_sheets() -> bytes:
    book = SpreadsheetDocument.new("First")
    for name in ("With spaces", "Ελληνικά"):
        book.add_sheet(name)
        book.append_row(name, [f"This is sheet {name}"])
    book.append_row("First", ["See the other two sheets"])
    return _bytes(book)


def xlsx_pictures() -> bytes:
    book = SpreadsheetDocument.new("Pictures")
    book.append_row("Pictures", ["PNG at C2, JPEG at C10"])
    book.add_picture("Pictures", image("PNG"), anchor="C2", width=Cm(5))
    book.add_picture("Pictures", image("JPEG"), anchor="C10", width=Cm(5))
    return _bytes(book)


def _excel_shaped_workbook() -> bytes:
    """A workbook shaped like Excel's own output: shared strings, styles."""
    ns = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    files = {
        "[Content_Types].xml":
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{ct}.sheet.main+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{ct}.worksheet+xml"/>'
            f'<Override PartName="/xl/sharedStrings.xml" ContentType="{ct}.sharedStrings+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{ct}.styles+xml"/></Types>',
        "_rels/.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/_rels/workbook.xml.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{rel}/sharedStrings" Target="sharedStrings.xml"/>'
            f'<Relationship Id="rId3" Type="{rel}/styles" Target="styles.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook {ns}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/sharedStrings.xml": f'<sst {ns} count="2" uniqueCount="2"><si><t>Name</t></si><si><t>Qty</t></si></sst>',
        "xl/styles.xml": f'<styleSheet {ns}><fonts count="2"><font/><font><b/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs></styleSheet>',
        "xl/worksheets/sheet1.xml": f'<worksheet {ns}><sheetData>'
            '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" t="s" s="1"><v>1</v></c></row>'
            '<row r="2"><c r="A2" t="inlineStr"><is><t>pens</t></is></c><c r="B2"><v>3</v></c></row>'
            '</sheetData></worksheet>',
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            z.writestr(info, text)
    return buf.getvalue()


def xlsx_edit_existing() -> bytes:
    book = SpreadsheetDocument.open(io.BytesIO(_excel_shaped_workbook()))
    book.set_value("Data", "A3", "paper")
    book.set_value("Data", "B3", 2)
    book.set_value("Data", "B1", "Quantity")  # keeps the bold header style
    book.set_value("Data", "C1", "edited by oox")
    return _bytes(book)


# -- PowerPoint ---------------------------------------------------------------------
def pptx_minimal() -> bytes:
    deck = PresentationDocument.new(layouts="basic")
    deck.add_slide()
    return _bytes(deck)


def pptx_placeholders() -> bytes:
    deck = PresentationDocument.new(layouts="basic")
    s1 = deck.add_slide("Title Slide")
    deck.set_placeholder_text(s1, "title", "Placeholder test")
    deck.set_placeholder_text(s1, "subtitle", GREEK)
    s2 = deck.add_slide("Title and Content")
    deck.set_placeholder_text(s2, "title", "Agenda")
    deck.set_placeholder_text(s2, "body", ["Results", (1, "by region"), (2, "north"), "Outlook"])
    return _bytes(deck)


def pptx_textboxes_pictures() -> bytes:
    deck = PresentationDocument.new(layouts="basic")
    s = deck.add_slide()
    deck.add_text_box(s, "Text boxes and pictures", x=Inches(0.7), y=Inches(0.5), width=Inches(12),
                      height=Inches(1), size=36, bold=True, color="1F3864")
    deck.add_text_box(s, "Left aligned\nsecond paragraph", x=Inches(0.7), y=Inches(1.8),
                      width=Inches(5), height=Inches(2), size=20)
    deck.add_text_box(s, "Centered", x=Inches(0.7), y=Inches(4.5), width=Inches(5),
                      height=Inches(1), size=20, align="center")
    for i, fmt in enumerate(("PNG", "JPEG", "GIF", "BMP")):
        deck.add_picture(s, image(fmt), x=Inches(6.5 + (i % 2) * 3.2),
                         y=Inches(1.8 + (i // 2) * 2), width=Inches(3))
    return _bytes(deck)


def pptx_4x3() -> bytes:
    deck = PresentationDocument.new(width=Emu(9144000), height=Emu(6858000),
                                    layouts="basic")
    s = deck.add_slide("Title and Content")
    deck.set_placeholder_text(s, "title", "4:3 slide size")
    deck.set_placeholder_text(s, "body", ["Placeholders scaled to 4:3", "Title fits the width"])
    return _bytes(deck)


def pptx_edit_powerpoint_template() -> bytes:
    import pptx  # python-pptx's template was authored in PowerPoint

    deck = PresentationDocument.open(Path(pptx.__file__).parent / "templates" / "default.pptx")
    s = deck.add_slide("Two Content")
    deck.set_placeholder_text(s, "title", "Two columns (PowerPoint layout)")
    deck.set_placeholder_text(s, 1, ["left 1", "left 2"])
    deck.set_placeholder_text(s, 2, "right")
    deck.add_picture(deck.add_slide("Title Only"), image("PNG"), x=Inches(2), y=Inches(2),
                     width=Inches(4))
    return _bytes(deck)


# -- Word in depth ------------------------------------------------------------------
def docx_defaults() -> bytes:
    """A document set up the way Word sets one up: styles, settings, fonts,
    theme, A4 section, document properties."""
    import datetime as dt

    doc = WordprocessingDocument.new()
    doc.properties.title = "oox default document"
    doc.properties.creator = "oox"
    doc.properties.created = dt.datetime(2026, 11, 14, 12, 0, tzinfo=dt.timezone.utc)
    doc.add_heading("Default document", 0)
    doc.add_heading("Heading level 1", 1)
    doc.add_paragraph("Normal paragraph in the theme's body font (Calibri 11).")
    doc.add_heading("Heading level 2", 2)
    doc.add_paragraph("Indented 'List Paragraph' style.", style="List Paragraph")
    tbl = w.Tbl(w.TblPr(w.TblStyle(val=doc.style_id("Table Grid")), w.TblW(w=0, type="auto")),
                w.TblGrid(w.GridCol(w=4513), w.GridCol(w=4513)))
    for row in (("Style", "Table Grid"), ("Borders", "from the style")):
        tbl.add(w.Tr(*(w.Tc(w.TcPr(w.TcW(w=4513, type="dxa")), w.P(w.R(w.T(t)))) for t in row)))
    doc.body.add(tbl, w.P())
    return _bytes(doc)


def docx_lists() -> bytes:
    """Bullets, multi-level decimal, a restarted list, legal numbering from 3.
    Paragraph numbers (for the probes) are given in the comments."""
    doc = WordprocessingDocument.new()
    doc.add_heading("Lists", 1)                         # 1
    bullets = doc.new_list("bullet")
    bullets.add("Bullet")                               # 2
    bullets.add("Sub-bullet (Courier New o)", 1)        # 3
    bullets.add("Sub-sub-bullet", 2)                    # 4
    steps = doc.new_list("decimal")
    steps.add("First")                                  # 5  -> 1.
    steps.add("First, a", 1)                            # 6  -> a.
    steps.add("First, b", 1)                            # 7  -> b.
    steps.add("First, b, i", 2)                         # 8  -> i.
    steps.add("Second")                                 # 9  -> 2.
    steps.add("Second, a (restarts under 2.)", 1)       # 10 -> a.
    again = steps.restart()
    again.add("Restarted list")                         # 11 -> 1.
    again.add("Continues")                              # 12 -> 2.
    legal = doc.new_list("legal", start=3)
    legal.add("Legal from 3")                           # 13 -> 3.
    legal.add("Legal 3.1", 1)                           # 14 -> 3.1.
    legal.add("Legal 3.1.1", 2)                         # 15 -> 3.1.1.
    legal.add("Legal 4")                                # 16 -> 4.
    return _bytes(doc)


def docx_sections() -> bytes:
    """Two sections: portrait with first/odd/even headers and page-number
    footers, then landscape with its own header and the footer linked."""
    doc = WordprocessingDocument.new()
    doc.add_heading("Cover page", 0)
    doc.add_paragraph("Page 1: the first-page header says 'Cover header'; no footer.")
    doc.add_page_break()
    doc.add_paragraph("Page 2 (even): 'Even header' left, 'Page 2 of 4' on the right.")
    doc.add_page_break()
    doc.add_paragraph("Page 3 (odd): 'Odd header' right, 'Page 3 of 4' centered.")
    first = doc.sections[0]
    first.header_text("Cover header", kind="first")
    first.header_text("Odd header", align="right")
    first.header_text("Even header", kind="even")
    first.footer_text("Page {PAGE} of {NUMPAGES}", align="center")
    first.footer_text("Page {PAGE} of {NUMPAGES}", kind="even", align="right")
    wide = doc.add_section(landscape=True)
    doc.add_heading("Landscape appendix", 1)
    doc.add_paragraph("Page 4: landscape, header 'Appendix header', footer from section 1.")
    wide.header_text("Appendix header")
    wide.header_text("Appendix header", kind="even")
    return _bytes(doc)


def docx_links_notes() -> bytes:
    """External and internal hyperlinks, a bookmark, footnotes, an endnote,
    a comment."""
    import datetime as dt

    doc = WordprocessingDocument.new()
    doc.add_heading("Links and notes", 1)
    p = doc.add_paragraph("External link to ")
    doc.add_hyperlink(p, "Ecma", target="https://ecma-international.org/", tooltip="Ecma site")
    p.add(w.R(w.T(" and an internal link to ", space="preserve")))
    doc.add_hyperlink(p, "the target heading", anchor="target")
    p.add(w.R(w.T(".")))
    q = doc.add_paragraph("A sentence with two footnotes")
    doc.add_footnote(q, "First footnote.")
    q.add(w.R(w.T(" and an endnote", space="preserve")))
    doc.add_endnote(q, "The endnote.")
    q.add(w.R(w.T(".")))
    doc.add_footnote(doc.add_paragraph("Another paragraph"), "Second footnote.")
    c = doc.add_paragraph("This sentence has a comment.")
    doc.add_comment(c, "Comment text from oox.", author="oox reviewer",
                    date=dt.datetime(2026, 11, 14, 12, 0))
    doc.add_page_break()
    doc.add_bookmark("target", doc.add_heading("Target heading", 1))
    doc.add_paragraph("The internal link jumps here.")
    return _bytes(doc)


def docx_toc_template() -> bytes:
    """A table of contents field over three heading levels; placeholders,
    one split across runs with different formatting, replaced everywhere."""
    doc = WordprocessingDocument.new()
    doc.add_heading("{{title}}", 0)
    toc_at = doc.add_heading("Chapter one", 1)
    doc.add_paragraph("Prepared for {{client}} on {{date}}.")
    doc.add_heading("Section 1.1", 2)
    split = w.P(w.R(w.T("Split placeholder: {{cli", space="preserve")),
                w.R(w.RPr(w.B()), w.T("en")), w.R(w.RPr(w.I()), w.T("t}} (bold/italic runs).")))
    doc.body.add(split)
    doc.add_heading("Section 1.1.1", 3)
    doc.add_heading("Chapter two", 1)
    doc.add_footnote(doc.add_paragraph("A footnote mentioning the client"), "For {{client}}.")
    doc.sections[0].header_text("{{client}} -- confidential")
    doc.add_toc(levels=(1, 3), title="Contents", before=toc_at)
    doc.replace_text({"{{title}}": "Template report", "{{client}}": "Example Ltd",
                      "{{date}}": "14 November 2026"})
    return _bytes(doc)


def docx_tables() -> bytes:
    """Table 1: header row (repeats), fixed widths, centered, shaded header, a
    horizontal merge.  Table 2: vertical, rectangular and horizontal merges."""
    doc = WordprocessingDocument.new()
    doc.add_heading("Tables", 1)
    t = doc.add_table([["Region", "Q1", "Q2", "Total"], ["North", 120, 135, 255],
                       ["South", 95, 110, 205], ["Subtotal", None, None, 460]],
                      header_rows=1, widths=[Cm(5), Cm(3), Cm(3), Cm(3)], align="center")
    t.merge(3, 0, 3, 2)
    for c in range(4):
        t.format_cell(0, c, fill="D9E2F3", align="center")
    t.format_cell(3, 0, align="right", bold=True)
    doc.add_paragraph("Merged cells: 'Group A' spans three rows, 'X' two rows and two "
                      "columns, 'Y' two columns.")
    t2 = doc.add_table([["Group A", "B1", "C1"], [None, "X", None], [None, None, None],
                        ["D4", "Y", None]])
    t2.merge(0, 0, 2, 0)
    t2.merge(1, 1, 2, 2)
    t2.merge(3, 1, 3, 2)
    t2.format_cell(0, 0, valign="center", fill="FFF2CC")
    t2.format_cell(1, 1, valign="center", align="center")
    return _bytes(doc)


def docx_floating() -> bytes:
    """Floating pictures: square wrap at the right margin, tight wrap by offset,
    behind the text centered on the page; one inline picture."""
    doc = WordprocessingDocument.new()
    doc.add_heading("Floating pictures", 1)
    p = doc.add_paragraph("Square wrap: the picture sits at the right margin and this text "
                          "flows around it. " * 6)
    doc.add_floating_picture(image("PNG"), width=Cm(4), h_align="right", h_relative="margin",
                             paragraph=p, wrap="square", descr="square")
    q = doc.add_paragraph("Tight wrap, placed 1 cm from the column and paragraph. " * 6)
    doc.add_floating_picture(image("PNG"), width=Cm(3), x=Cm(1), y=Cm(1), paragraph=q,
                             wrap="tight", descr="tight")
    r = doc.add_paragraph("Behind the text: a large picture centered on the page.")
    doc.add_floating_picture(image("PNG"), width=Cm(12), h_align="center", v_align="center",
                             h_relative="page", v_relative="page", paragraph=r, wrap="behind",
                             descr="behind")
    doc.add_paragraph("Inline picture for comparison:")
    doc.add_picture(image("PNG"), width=Cm(2))
    return _bytes(doc)


def docx_tracked() -> bytes:
    """Tracked changes by two authors: an insertion, a deletion, a replacement;
    Track Changes switched on."""
    import datetime as dt

    doc = WordprocessingDocument.new()
    doc.add_heading("Tracked changes", 1)
    p = doc.add_paragraph("The quick brown fox jumps over the very lazy dog in 2025.")
    when = dt.datetime(2026, 11, 14, 12, 0)
    doc.suggest_delete(p, "very ", author="Bob", date=when)
    doc.suggest_replace(p, "2025", "2026", author="Ada", date=when)
    doc.suggest_insert(p, " Really.", author="Ada", date=when)
    doc.add_paragraph("This paragraph has no changes.")
    doc.track_changes = True
    return _bytes(doc)


def docx_formatting_styles() -> bytes:
    """Run formatting, a custom paragraph style and a character style,
    paragraph alignment, spacing and indents.  Paragraph numbers in comments."""
    doc = WordprocessingDocument.new()
    doc.add_heading("Formatting and styles", 1)                          # 1
    doc.add_style("Note", italic=True, color="595959", space_before=6, indent_left=36)
    doc.add_style("Code Char", kind="character", font="Courier New", size=10)
    p = doc.add_paragraph("", style="Note")                               # 2
    doc.add_run(p, "Warning: ", bold=True, color="C00000")                # chars 1-9
    doc.add_run(p, "check ", style="Code Char")                           # chars 10-15
    doc.add_run(p, "this", underline=True, size=14, highlight="yellow")   # chars 16-19
    q = doc.add_paragraph("Justified, 12 pt after, first line indented 18 pt. " * 4)  # 3
    doc.format_paragraph(q, align="justify", space_after=12, first_line=18)
    r = doc.add_paragraph("Right aligned, indented 36 pt on the left.")   # 4
    doc.format_paragraph(r, align="right", indent_left=36)
    s = doc.add_paragraph("")                                             # 5
    doc.add_run(s, "H")
    doc.add_run(s, "2", subscript=True)
    doc.add_run(s, "O and E = mc")
    doc.add_run(s, "2", superscript=True)
    doc.add_run(s, "; small caps", small_caps=True)
    doc.add_run(s, "; struck", strike=True)
    return _bytes(doc)


def docx_fields_captions() -> bytes:
    """Captions with SEQ numbers (one inserted before the others, so all
    renumber), cross-references, document-property fields."""
    import datetime as dt

    doc = WordprocessingDocument.new()
    doc.properties.title = "Fields and captions"
    doc.properties.created = dt.datetime(2026, 11, 14, 10, 0, tzinfo=dt.timezone.utc)
    doc.extended_properties.company = "Example Ltd"
    doc.custom_properties["Project"] = "Apollo"
    doc.custom_properties["Approved"] = True
    intro = doc.add_heading("Introduction", 1)
    pic1 = doc.add_picture(image("PNG"), width=Cm(4))
    cap1 = doc.add_caption("Sales by region", after=pic1.getparent())
    p = doc.add_paragraph("See ")
    doc.add_cross_reference(p, cap1)
    p.add(w.R(w.T(", which is ", space="preserve")))
    doc.add_cross_reference(p, cap1, what="position")
    p.add(w.R(w.T(" this paragraph, on page ", space="preserve")))
    doc.add_cross_reference(p, cap1, what="page")
    p.add(w.R(w.T(". Section: ", space="preserve")))
    doc.add_cross_reference(p, intro)
    doc.add_page_break()
    pic2 = doc.add_picture(image("PNG"), width=Cm(4))
    doc.add_caption("Costs", after=pic2.getparent())
    doc.add_caption("Inserted before the others", before=intro)   # becomes Figure 1
    doc.add_caption("Headcount", label="Table")
    q = doc.add_paragraph("Title: ")
    doc.add_field(q, "TITLE")
    q.add(w.R(w.T("; company: ", space="preserve")))
    doc.add_field(q, 'DOCPROPERTY "Company"')
    q.add(w.R(w.T("; project: ", space="preserve")))
    doc.add_field(q, "DOCPROPERTY Project \\* Upper")
    q.add(w.R(w.T("; approved: ", space="preserve")))
    doc.add_field(q, 'DOCPROPERTY "Approved"')
    q.add(w.R(w.T("; created: ", space="preserve")))
    doc.add_field(q, 'CREATEDATE \\@ "d MMMM yyyy"')
    return _bytes(doc)


def docx_content_controls() -> bytes:
    """A form: text, drop-down, combo, date and check-box controls, mostly
    filled, one still showing its placeholder."""
    import datetime as dt

    doc = WordprocessingDocument.new()
    doc.add_heading("Request form", 1)
    rows = [("Name: ", lambda p: doc.add_text_control(p, tag="name", title="Name",
                                                       placeholder="Your name")),
            ("Department: ", lambda p: doc.add_dropdown(
                p, tag="dept", title="Department", items=["Sales", ("Finance", "FIN"), "IT"])),
            ("Office: ", lambda p: doc.add_dropdown(p, tag="office", items=["Athens", "London"],
                                                     combo=True)),
            ("Start date: ", lambda p: doc.add_date_control(p, tag="start", fmt="d MMMM yyyy")),
            ("I agree: ", lambda p: doc.add_checkbox(p, tag="agree")),
            ("Comments: ", lambda p: doc.add_text_control(p, tag="comments",
                                                           placeholder="Any comments")),
            ]
    for label, make in rows:
        make(doc.add_paragraph(label))
    doc.fill_controls({"name": "Ada Lovelace", "dept": "FIN", "office": "Lisbon",
                       "start": dt.date(2026, 11, 14), "agree": True})
    return _bytes(doc)


def docx_more_images() -> bytes:
    """TIFF, EMF, WMF and SVG (with a PNG fallback), all 4 cm wide."""
    doc = WordprocessingDocument.new()
    doc.add_heading("More image formats", 1)
    for fmt in ("TIFF", "EMF", "WMF"):
        doc.add_paragraph(f"{fmt}:")
        doc.add_picture(image(fmt), width=Cm(4), descr=fmt)
    doc.add_paragraph("SVG (Office 2016+; older versions show the PNG fallback):")
    doc.add_picture(image("SVG"), fallback=image("PNG"), width=Cm(4), descr="SVG")
    return _bytes(doc)


def xlsx_properties_images() -> bytes:
    book = SpreadsheetDocument.new()
    book.extended_properties.company = "Example Ltd"
    book.custom_properties["Project"] = "Apollo"
    book.custom_properties["Budget"] = 1250
    book.set_value("Sheet1", "A1", "TIFF and SVG pictures below")
    book.add_picture("Sheet1", image("TIFF"), "B3", width=Cm(4), descr="TIFF")
    book.add_picture("Sheet1", image("SVG"), "B9", fallback=image("PNG"), width=Cm(4),
                     descr="SVG")
    return _bytes(book)


def pptx_properties_images() -> bytes:
    deck = PresentationDocument.new(layouts="basic")
    deck.extended_properties.company = "Example Ltd"
    deck.custom_properties["Project"] = "Apollo"
    deck.custom_properties["Approved"] = True
    slide = deck.add_slide("Blank")
    deck.add_picture(slide, image("EMF"), x=Inches(1), y=Inches(1), width=Cm(6), descr="EMF")
    deck.add_picture(slide, image("SVG"), x=Inches(5), y=Inches(1), width=Cm(6),
                     fallback=image("PNG"), descr="SVG")
    return _bytes(deck)


# excel-08: (cell, value, number format, text Excel shows)
EXCEL_FORMATS = [
    ("B2", 1234.5, "#,##0.00", "1,234.50"),
    ("B3", -98.25, "#,##0.00;[Red](#,##0.00)", "(98.25)"),
    ("B4", 0.125, "0.0%", "12.5%"),
    ("B5", 1234.5, '"\u20ac"#,##0.00', "\u20ac1,234.50"),
    ("B6", 1234.5, "0.00E+00", "1.23E+03"),
    ("B7", 7, "000", "007"),
    ("B8", 1 / 3, "General", "0.333333333"),
    ("B9", 0, '#,##0;-#,##0;"zero"', "zero"),
    ("B10", "text", '"<"@">"', "<text>"),
]
EXCEL_DATES = [
    ("B12", "date", "yyyy-mm-dd", "2026-11-14"),
    ("B13", "datetime", "d mmmm yyyy hh:mm", "14 November 2026 15:07"),
    ("B14", "time", "h:mm AM/PM", "3:07 PM"),
    ("B15", "date", "dddd", "Saturday"),
    ("B16", "default-date", None, "2026-11-14"),
    ("B17", "default-datetime", None, "2026-11-14 15:07"),
]


def xlsx_styles_formats() -> bytes:
    """Fonts, fills, borders, alignment; number formats as Excel displays them."""
    import datetime as dt

    book = SpreadsheetDocument.new("Formats")
    hdr = book.cell_style(bold=True, fill="D9E2F3", border="thin", align="center",
                          font="Arial", size=12, color="1F3864")
    book.append_row("Formats", ["Format", "Value"], style=hdr)
    for ref, value, code, _ in EXCEL_FORMATS:
        book.set_value("Formats", "A" + ref[1:], code)
        book.set_value("Formats", ref, value, style={"number_format": code})
    when = dt.datetime(2026, 11, 14, 15, 7, 9)
    for ref, kind, code, _ in EXCEL_DATES:
        value = {"date": when.date(), "datetime": when, "time": when.time(),
                 "default-date": when.date(), "default-datetime": when}[kind]
        book.set_value("Formats", "A" + ref[1:], code or f"({kind})")
        book.set_value("Formats", ref, value, style={"number_format": code} if code else None)
    book.set_value("Formats", "A19", "Italic, underlined, wrapped text in a narrow column",
                   style={"italic": True, "underline": True, "wrap": True, "valign": "top"})
    book.set_column_width("Formats", "A", 28)
    book.set_column_width("Formats", "B", 26)
    return _bytes(book)


def xlsx_layout() -> bytes:
    """Column widths, row height, merged cells, frozen header row, autofilter,
    a defined name, a drop-down validation."""
    book = SpreadsheetDocument.new("Sales")
    book.append_row("Sales", ["Region", "Q1", "Q2", "Total", "Checked"],
                    style={"bold": True, "fill": "D9E2F3"})
    for i, (region, q1, q2) in enumerate([("North", 120, 135), ("South", 95, 110),
                                          ("West", 80, 84)], start=2):
        book.append_row("Sales", [region, q1, q2, Formula(f"B{i}+C{i}"), "Yes"])
    book.set_column_width("Sales", "A", 20)
    book.set_column_width("Sales", "B:D", 12)
    book.set_row_height("Sales", 1, 24)
    book.merge_cells("Sales", "A6:D6")
    book.set_value("Sales", "A6", "Merged across A6:D6", style={"align": "center", "bold": True})
    book.set_frozen_panes("Sales", "B2")
    book.auto_filter("Sales", "A1:E4")
    book.define_name("Totals", "Sales", "D2:D4")
    book.add_validation("Sales", "E2:E4", values=["Yes", "No"], prompt="Pick one")
    return _bytes(book)


def xlsx_style_existing() -> bytes:
    """Style cells in an Excel-shaped workbook with its own (minimal) stylesheet."""
    import datetime as dt

    book = open_document(io.BytesIO(_excel_shaped_workbook()))
    book.set_value("Data", "C1", "When", style={"bold": True})   # reuses the existing bold font
    book.set_value("Data", "C2", dt.date(2026, 11, 14))
    book.set_value("Data", "D2", 0.5, style={"number_format": "0%", "fill": "FFF2CC"})
    return _bytes(book)


def xlsx_names_autofit() -> bytes:
    """English month/day names on any regional setting ([$-409]) and columns
    widened to fit their text (autofit)."""
    import datetime as dt

    book = SpreadsheetDocument.new("Names")
    when = dt.datetime(2026, 11, 14, 15, 7, 9)
    rows = [("[$-409]d mmmm yyyy", when), ("[$-409]dddd", when), ("[$-409]h:mm AM/PM", when),
            ("[$-409]mmm d, yyyy", when), ("#,##0.00", 1234567.891), ("yyyy-mm-dd hh:mm:ss", when)]
    for i, (code, value) in enumerate(rows, start=1):
        book.set_value("Names", f"A{i}", code)
        book.set_value("Names", f"B{i}", value, style={"number_format": code})
    book.autofit_columns("Names")
    return _bytes(book)


def xlsx_conditional() -> bytes:
    """Conditional formats of every kind on one block of data."""
    book = SpreadsheetDocument.new("Rules")
    book.append_row("Rules", ["Region", "Sales", "Margin", "Units", "Trend"],
                    style={"bold": True})
    data = [("North", 120, 14, 2, 9), ("South", 30, 21, 4, 1), ("West", 40, 28, 6, 5),
            ("North", 80, 35, 8, 14), ("East", 60, 42, 10, 10), ("South", 70, 49, 12, 6),
            ("North", 90, 6, 14, 2), ("West", 100, 13, 16, 15), ("East", 50, 20, 18, 11),
            ("North", 110, 27, 20, 7)]
    for row in data:
        book.append_row("Rules", list(row))
    book.conditional_format("Rules", "B2:B11", "cell", operator=">", value=95,
                            style={"fill": "FFC7CE", "color": "9C0006"})
    book.conditional_format("Rules", "B2:B11", "cell", operator="between", values=(30, 50),
                            style={"fill": "FFEB9C"})
    book.conditional_format("Rules", "B2:B11", "top", n=2, style={"bold": True})
    book.conditional_format("Rules", "C2:C11", "color_scale")
    book.conditional_format("Rules", "D2:D11", "data_bar", color="638EC6")
    book.conditional_format("Rules", "E2:E11", "icon_set", icons="3Arrows")
    book.conditional_format("Rules", "A2:A11", "text", text="South", style={"italic": True})
    book.conditional_format("Rules", "A2:A11", "duplicates", style={"color": "0070C0"})
    book.conditional_format("Rules", "A2:E11", "formula", formula="$E2>$D2",
                            style={"fill": "C6EFCE"})
    book.autofit_columns("Rules")
    return _bytes(book)


def xlsx_tables() -> bytes:
    """Two tables: one with a totals row (SUBTOTAL), one plain with another
    style; a formula with a structured reference."""
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["Region", "Q1", "Q2", "Notes"])
    for row in [("North", 120, 135, "ok"), ("South", 95, 110, None), ("West", 80, 84.5, "late")]:
        book.append_row("Data", list(row))
    book.add_table("Data", "A1:D4", name="SalesTable",
                   totals={"Region": "Total", "Q1": "sum", "Q2": "average", "Notes": "count"})
    book.set_value("Data", "F1", "Q1 total")
    book.set_value("Data", "G1", Formula("SUM(SalesTable[Q1])", 295))
    book.append_row("Data", [])
    book.set_value("Data", "A8", "Code")
    book.set_value("Data", "B8", "Name")
    for i, (code, name) in enumerate([("N", "North"), ("S", "South")], start=9):
        book.set_value("Data", f"A{i}", code)
        book.set_value("Data", f"B{i}", name)
    book.add_table("Data", "A8:B10", name="Lookup", style="TableStyleLight9")
    book.autofit_columns("Data")
    return _bytes(book)


def xlsx_comments_links() -> bytes:
    """Comments on two sheets (one always visible), external and internal links."""
    book = SpreadsheetDocument.new("Data")
    book.add_sheet("Summary")
    book.append_row("Data", ["Region", "Amount"])
    book.append_row("Data", ["North", 120])
    book.add_comment("Data", "B2", "Checked by finance", author="Ada Lovelace")
    book.add_comment("Data", "A1", "This note stays open", author="oox", visible=True)
    book.add_comment("Summary", "C3", "On the second sheet", author="Ada Lovelace")
    book.add_hyperlink("Data", "A4", "https://ecma-international.org/", text="Ecma",
                       tooltip="The OOXML standard")
    book.add_hyperlink("Data", "A5", "#Summary!A1", text="Go to Summary")
    book.set_value("Summary", "A1", "You arrived here from Data!A5")
    return _bytes(book)


def xlsx_print_options() -> bytes:
    """Print setup, tab colour, zoom, hidden and very hidden sheets, a
    protected sheet, shared strings."""
    book = SpreadsheetDocument.new("Report")
    book.add_sheet("Lookup")
    book.add_sheet("Internal")
    book.append_row("Report", ["Region", "Amount", "Owner"], style={"bold": True})
    for i in range(1, 30):
        book.append_row("Report", ["North" if i % 2 else "South", i * 10, "Ada"])
    book.append_row("Lookup", ["North", "South"])
    book.append_row("Internal", ["not for users"])
    book.set_page_setup("Report", orientation="landscape", paper="A4", fit_width=1,
                    print_area="A1:C30", repeat_rows="1:1", header="Quarterly report",
                    footer={"left": "R&D", "right": "Page {page} of {pages}"}, gridlines=True)
    book.set_sheet_options("Report", tab_color="2F5496", zoom=120)
    book.set_sheet_options("Lookup", hidden=True)
    book.set_sheet_options("Internal", hidden="very")
    book.protect_sheet("Report", password="secret", allow=("sort", "autoFilter"))
    book.share_strings()
    return _bytes(book)


def xlsx_properties() -> bytes:
    book = SpreadsheetDocument.new("Props")
    book.append_row("Props", ["See File > Properties"])
    book.properties.title = "oox workbook"
    book.properties.creator = "oox"
    return _bytes(book)


def xlsx_charts() -> bytes:
    """Five charts from cell ranges on two sheets."""
    book = SpreadsheetDocument.new("Sales")
    book.append_row("Sales", ["Region", "Q1", "Q2", "Q3"], style={"bold": True})
    for row in [("North", 120, 135, 150), ("South", 95, 110, 99), ("West", 80, 84.5, 91),
                ("East", 60, 72, 88)]:
        book.append_row("Sales", list(row))
    book.add_chart("Sales", "column", "A1:D5", anchor="F2", title="Quarterly sales",
                   x_title="Region", y_title="EUR k", legend="bottom")
    book.add_chart("Sales", "pie", "A1:B5", anchor="F18", title="Q1 share", labels="percent")
    book.add_chart("Sales", "line", "A1:D5", anchor="N2", series_in="rows", markers=False,
                   title="Trend", x_title="Quarter")
    book.add_chart("Sales", "bar", "A1:D5", anchor="N18", grouping="stacked",
                   title="Year to date")
    book.add_sheet("Points")
    book.append_row("Points", ["x", "Measured"], style={"bold": True})
    for row in [(1, 2.1), (2, 3.9), (3, 6.2), (4, 7.8), (5, 10.1)]:
        book.append_row("Points", list(row))
    book.add_chart("Points", "scatter", "A1:B6", anchor="D2", lines=True, title="Measurements",
                   x_title="x", y_title="y", legend=None)
    return _bytes(book)


def xlsx_combo_charts() -> bytes:
    """A column + line chart with the line on a secondary axis, and areas behind columns."""
    book = SpreadsheetDocument.new("Combo")
    book.append_row("Combo", ["Month", "Units", "Price", "Target"], style={"bold": True})
    for row in [("Jan", 120, 2.5, 100), ("Feb", 150, 2.25, 130), ("Mar", 90, 2.75, 140),
                ("Apr", 170, 2.4, 150)]:
        book.append_row("Combo", list(row))
    book.add_chart("Combo", "column", "A1:C5", anchor="F2", kinds={"Price": "line"},
                   secondary=["Price"], title="Units and price", y_title="Units",
                   y2_title="EUR", y2_number_format="0.00", legend="bottom")
    book.add_chart("Combo", "column", "A1:D5", anchor="F18",
                   kinds={"Target": "area", "Price": "line"}, secondary=["Price"],
                   title="Against target")
    return _bytes(book)


def xlsx_edit_openpyxl_file() -> bytes:
    """A workbook openpyxl wrote (styles, a table, a chart without cached
    values), edited by oox: a row appended and a second chart added."""
    import datetime as dt

    import openpyxl
    from openpyxl.chart import BarChart, Reference

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Item", "Count"])
    for row in [("a", 3), ("b", 5), ("c", 4)]:
        ws.append(row)
    ws["A1"].font = openpyxl.styles.Font(bold=True)
    chart = BarChart()
    chart.title = "openpyxl chart"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=4), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=4))
    ws.add_chart(chart, "D2")
    buf = io.BytesIO()
    wb.save(buf)
    book = SpreadsheetDocument.open(io.BytesIO(buf.getvalue()))
    when = dt.datetime(2026, 11, 15, 12, 0, tzinfo=dt.timezone.utc)   # openpyxl wrote "now"
    book.properties.created = book.properties.modified = when
    book.append_row("Data", ["d", 6])
    book.add_chart("Data", "line", "A1:B5", anchor="D18", title="oox chart")
    return _bytes(book)


def xlsx_workbook_protection() -> bytes:
    """Two sheets with the workbook structure locked by a password."""
    book = SpreadsheetDocument.new("First")
    book.set_value("First", "A1", "The workbook structure is locked")
    book.add_sheet("Second")
    book.set_value("Second", "A1", "Sheets can't be added, moved or renamed")
    book.protect_workbook(password="secret")
    return _bytes(book)


def xlsx_chart_sheet() -> bytes:
    """A chart sheet between two worksheets; a name scoped to the last sheet."""
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["Region", "Q1", "Q2"], style={"bold": True})
    for row in [("North", 120, 135), ("South", 95, 110), ("West", 80, 84.5)]:
        book.append_row("Data", list(row))
    book.add_chart_sheet("Sales chart", "column", "Data", "A1:C4", title="Sales by region",
                         y_title="EUR k", legend="bottom", labels="value")
    book.add_sheet("Notes")
    book.set_value("Notes", "A1", "After the chart sheet")
    book.define_name("NoteCell", "Notes", "A1", local=True)
    return _bytes(book)


def docx_text_boxes() -> bytes:
    """Three floating text boxes: square wrap beside the text, a yellow box on
    the page with top-and-bottom wrap, a borderless one in front of the text."""
    doc = WordprocessingDocument.new()
    doc.add_heading("Text boxes", 1)
    p = doc.add_paragraph("The first text box sits at the right margin with square "
                          "wrapping, so this text flows to its left. " * 4)
    doc.add_text_box("Note\nSecond paragraph", width=Cm(5), height=Cm(3), h_align="right",
                     h_relative="margin", paragraph=p)
    q = doc.add_paragraph("The second box is placed on the page, 7 cm from the left and "
                          "18 cm from the top, with text above and below it.")
    doc.add_text_box("Page box", width=Cm(7), height=Cm(2.5), x=Cm(7), y=Cm(18),
                     h_relative="page", v_relative="page", wrap="topAndBottom",
                     fill="FFF2CC", line="BF9000", line_width=1.5, paragraph=q)
    r = doc.add_paragraph("The third box has no border or fill and sits in front of the text.")
    doc.add_text_box("Plain box", width=Cm(4), height=Cm(1.5), x=Cm(12), y=Cm(1),
                     wrap="front", fill=None, line=None, paragraph=r)
    return _bytes(doc)


PROTECTION_SALT = bytes(range(16))  # fixed, so the corpus files are reproducible


def _protected_doc(title: str) -> WordprocessingDocument:
    doc = WordprocessingDocument.new()
    doc.add_heading(title, 1)
    doc.add_paragraph("This paragraph can be read; editing is restricted.")
    return doc


def docx_protection_read_only() -> bytes:
    doc = _protected_doc("Read-only protection")
    doc.protect("readOnly", password="secret", salt=PROTECTION_SALT)
    return _bytes(doc)


def docx_protection_tracked_sha512() -> bytes:
    doc = _protected_doc("Tracked changes only (SHA-512 password)")
    doc.protect("trackedChanges", password="secret", algorithm="SHA-512",
                salt=PROTECTION_SALT)
    return _bytes(doc)


def docx_protection_comments() -> bytes:
    doc = _protected_doc("Comments only, no password")
    doc.protect("comments")
    return _bytes(doc)


def docx_charts() -> bytes:
    """Three inline charts, each with an embedded workbook."""
    doc = WordprocessingDocument.new()
    doc.body.add(w.P(w.R(w.T("Charts in Word"))))
    doc.add_chart("column", ["Q1", "Q2", "Q3"], {"North": [120, 135, 150], "South": [95, 110, 99]},
                  title="Sales", legend="bottom", labels="value")
    doc.add_chart("area", ["Jan", "Feb", "Mar", "Apr"],
                  {"Web": [30, 42, 51, 60], "Shop": [20, 18, 25, 22]},
                  grouping="stacked", title="Orders by channel")
    doc.add_chart("pie", ["North", "South", "West"], {"Share": [45, 30, 25]}, title="Share",
                  labels="percent", width=Cm(10), height=Cm(7))
    return _bytes(doc)


def pptx_charts() -> bytes:
    """Charts on two slides, each with an embedded workbook."""
    deck = PresentationDocument.new(layouts="basic")
    s1 = deck.add_slide("Blank")
    deck.add_text_box(s1, "Traffic", x=Inches(0.5), y=Inches(0.3), width=Inches(9),
                      height=Inches(0.8))
    deck.add_chart(s1, "line", ["Jan", "Feb", "Mar", "Apr"],
                   {"Visits": [1200, 1450, 1390, 1600], "Orders": [310, 365, 330, 410]},
                   x=Inches(1), y=Inches(1.5), width=Inches(8), height=Inches(4.5),
                   title="Visits and orders", y_title="count")
    s2 = deck.add_slide("Blank")
    deck.add_text_box(s2, "Mix", x=Inches(0.5), y=Inches(0.3), width=Inches(9), height=Inches(0.8))
    deck.add_chart(s2, "doughnut", ["Web", "Shop", "Phone"], {"Channel": [55, 30, 15]},
                   x=Inches(0.5), y=Inches(1.5), width=Inches(4.5), height=Inches(4.5),
                   labels="percent", title="Channels")
    deck.add_chart(s2, "column", ["2024", "2025"], {"New": [40, 55], "Returning": [60, 45]},
                   x=Inches(5.2), y=Inches(1.5), width=Inches(4.5), height=Inches(4.5),
                   grouping="percent", title="Customers", number_format="0%")
    return _bytes(deck)


def pptx_layouts_tables_notes() -> bytes:
    """Office's standard layouts, a table, speaker notes, a filled picture
    placeholder and auto shapes."""
    deck = PresentationDocument.new()
    s1 = deck.add_slide("Section Header")
    deck.set_placeholder_text(s1, "title", "Part 1")
    deck.set_placeholder_text(s1, "body", "Results")
    s2 = deck.add_slide("Two Content")
    deck.set_placeholder_text(s2, "title", "Compare")
    deck.set_placeholder_text(s2, 1, ["Left a", "Left b"])
    deck.set_placeholder_text(s2, 2, ["Right a", (1, "detail")])
    deck.set_notes(s2, "Speaker note for slide 2")
    s3 = deck.add_slide("Title Only")
    deck.set_placeholder_text(s3, "title", "Table")
    deck.add_table(s3, [["Region", "Q1", "Q2"], ["North", 120, 135], ["South", 95, 110],
                        ["Total", None, 245]],
                   x=Inches(1), y=Inches(2), width=Inches(8), merges=["A4:B4"])
    deck.set_notes(s3, "Totals are rounded.")
    s4 = deck.add_slide("Picture with Caption")
    deck.set_placeholder_text(s4, "title", "Picture")
    deck.set_placeholder_text(s4, 2, "Cropped to fill the placeholder")
    deck.set_placeholder_picture(s4, "pic", image("PNG"))
    s5 = deck.add_slide("Blank")
    deck.add_shape(s5, "roundRect", x=Inches(1), y=Inches(2), width=Inches(3), height=Inches(1.2),
                   text="Step 1")
    deck.add_shape(s5, "chevron", x=Inches(4.5), y=Inches(2), width=Inches(3), height=Inches(1.2),
                   text="Step 2", fill="ED7D31")
    deck.add_shape(s5, "ellipse", x=Inches(8), y=Inches(2), width=Inches(2), height=Inches(2),
                   text="Done", fill="70AD47", line="375623", line_width=2)
    return _bytes(deck)


def pptx_links_table_edit() -> bytes:
    """Text and shape links (URL, mail, slide, show jumps) and a table edited
    after it was made: a value changed, a row and a column added, a row
    deleted, a merged total row."""
    deck = PresentationDocument.new()
    s1 = deck.add_slide("Title Only")
    s2 = deck.add_slide("Title Only")
    s3 = deck.add_slide("Title Only")
    deck.set_placeholder_text(s1, "title", "Links")
    box = deck.add_text_box(s1, "Read the docs or mail the team.", x=Inches(1), y=Inches(2),
                            width=Inches(8), height=Inches(0.8), size=24)
    deck.set_link(s1, box, "https://example.org/docs", text="docs", tooltip="Manual")
    deck.set_link(s1, box, "mailto:team@example.org", text="mail the team")
    nxt = deck.add_shape(s1, "rightArrow", x=Inches(1), y=Inches(4), width=Inches(2.5),
                         height=Inches(1), text="Next", name="Next arrow")
    deck.set_link(s1, nxt, "next")
    jump = deck.add_shape(s1, "roundRect", x=Inches(5), y=Inches(4), width=Inches(3),
                          height=Inches(1), text="Summary", name="To summary")
    deck.set_link(s1, jump, s3)
    deck.set_placeholder_text(s2, "title", "Edited table")
    deck.add_table(s2, [["Region", "Q1", "Q2"], ["North", 120, 135], ["South", 95, 110]],
                   x=Inches(1), y=Inches(2), width=Inches(6), name="Sales")
    t = deck.table(s2, "Sales")
    t.set_text(1, 1, 125)
    t.add_row(["West", 80, 84])
    t.add_column(["Q3", 150, 99, 91], width=Inches(2))
    t.delete_row(2)
    t.add_row(["Total", "see appendix"])
    t.merge(3, 1, 3, 3)
    deck.set_placeholder_text(s3, "title", "Summary")
    back = deck.add_shape(s3, "leftArrow", x=Inches(1), y=Inches(4), width=Inches(2.5),
                          height=Inches(1), text="Start", name="Back arrow")
    deck.set_link(s3, back, "first")
    return _bytes(deck)


def pptx_edit_python_pptx_file() -> bytes:
    """A python-pptx deck (title slide; a table with a merge) edited by oox:
    a column deleted, a row added, a link on the title."""
    import datetime as dt

    import pptx
    from pptx.util import Inches as PInches

    prs = pptx.Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "python-pptx table"
    table = s.shapes.add_table(3, 3, PInches(1), PInches(2), PInches(6), PInches(1.5)).table
    for i, row in enumerate([["Item", "Old", "Count"], ["a", "x", "3"], ["b", "y", "5"]]):
        for j, text in enumerate(row):
            table.cell(i, j).text = text
    table.cell(1, 2).merge(table.cell(2, 2))
    buf = io.BytesIO()
    prs.save(buf)
    deck = PresentationDocument.open(io.BytesIO(buf.getvalue()))
    when = dt.datetime(2026, 11, 15, 12, 0, tzinfo=dt.timezone.utc)   # python-pptx wrote "now"
    deck.properties.created = deck.properties.modified = when
    slide = deck.slides[0]
    t = deck.table(slide)
    t.delete_column(1)
    t.add_row(["c", "4"])
    deck.set_link(slide, "Title 1", "https://example.org/", text="python-pptx")
    return _bytes(deck)


def pptx_sections() -> bytes:
    """Five slides in three sections (PowerPoint 2010+; 2007 ignores them)."""
    deck = PresentationDocument.new()
    titles = ["Welcome", "Agenda", "North", "South", "Questions"]
    slides = []
    for title in titles:
        s = deck.add_slide("Title Only")
        deck.set_placeholder_text(s, "title", title)
        slides.append(s)
    deck.set_sections([("Opening", slides[0]), ("Regions", slides[2]), ("Close", slides[4])])
    return _bytes(deck)


def pptx_properties() -> bytes:
    deck = PresentationDocument.new(layouts="basic")
    s = deck.add_slide("Title Slide")
    deck.set_placeholder_text(s, "title", "See File > Properties")
    deck.properties.title = "oox deck"
    deck.properties.creator = "oox"
    return _bytes(deck)


# -- negative controls -------------------------------------------------------------
# Deliberately broken files.  Office MUST complain about them; if the checker
# reports "ok", it is not seeing Office's dialogs (or Office silently repaired
# a malformed part).  They prove that an all-green column means something.
def _truncate_part(package: bytes, name: str) -> bytes:
    """Rewrite the zip with ``name`` cut in half -> malformed XML."""
    src = zipfile.ZipFile(io.BytesIO(package))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info)
            if info.filename == name:
                data = data[: len(data) // 2]
            out.writestr(info, data)
    return buf.getvalue()


def docx_control_malformed() -> bytes:
    return _truncate_part(docx_formatting(), "word/document.xml")


def xlsx_control_malformed() -> bytes:
    return _truncate_part(xlsx_types(), "xl/worksheets/sheet1.xml")


def pptx_control_malformed() -> bytes:
    return _truncate_part(pptx_placeholders(), "ppt/slides/slide2.xml")


CONTROL = ("prompt", "error", "mismatch")
OBSERVE = ("observe",)  # no expectation: record how strict Office is


# -- observe cases ------------------------------------------------------------------
# Well-formed XML that breaks the schema in one specific way.  There is no
# right answer: the results tell us how much each kind of validate() finding
# matters to real Office versions.  Built with raw lxml calls, which bypass
# oox's ordering and checks on purpose.
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def docx_observe_order() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    p = w.P()
    doc.body.add(p)
    ppr, rpr = w.PPr(), w.RPr()
    p.append(ppr)
    ppr.append(w.Jc(val="center"))
    ppr.append(w.PStyle(val="Normal"))  # pStyle must come first
    run = w.R()
    p.append(run)
    run.append(rpr)
    rpr.append(w.Sz(val=40))
    rpr.append(w.B())  # b must come before sz
    run.append(w.T("Properties in the wrong order"))
    return _bytes(doc)


def docx_observe_sdk_values() -> bytes:
    """Values ISO/IEC 29500 Transitional allows but the SDK (and so oox) rejects,
    as LibreOffice writes them: w:lvlJc val="start" (ST_Jc), table w:jc
    val="start" (ST_JcTable), w:cantSplit and w:tblHeader val="true" (ST_OnOff).
    Found by tools/interop.py."""
    from lxml import etree

    doc = WordprocessingDocument.new()
    steps = doc.new_list("decimal")
    steps.add("First item")                             # 1 -> 1.
    steps.add("Second item")                            # 2 -> 2.
    doc.add_table([["Header", "Value"], ["a", "1"]], header_rows=1)
    W = W_NS
    for part in doc.package.parts.values():
        if part.is_xml and part.partname.endswith("numbering.xml"):
            for jc in part.element.iter(W + "lvlJc"):
                jc.set(W + "val", "start")
    for tblpr in doc.body.iter(W + "tblPr"):
        jc = tblpr.find(W + "jc")
        if jc is None:
            jc = etree.SubElement(tblpr, W + "jc")
        jc.set(W + "val", "start")
    for tr in doc.body.iter(W + "tr"):
        trpr = tr.find(W + "trPr")
        if trpr is None:
            trpr = etree.Element(W + "trPr")
            tr.insert(0, trpr)
        for hdr in trpr.iter(W + "tblHeader"):
            hdr.set(W + "val", "true")
        etree.SubElement(trpr, W + "cantSplit").set(W + "val", "true")
    doc.add_paragraph("lvlJc start, table jc start, cantSplit true, tblHeader true")
    return _bytes(doc)


def docx_observe_libreoffice_structure() -> bytes:
    """Structure LibreOffice writes that the schema forbids (tools/interop.py):
    w:pStyle before w:numFmt in a numbering level, a w:rPr with w:i twice and
    w:sz twice."""
    doc = WordprocessingDocument.new()
    steps = doc.new_list("decimal")
    steps.add("Numbered with pStyle first")
    W = W_NS
    for part in doc.package.parts.values():
        if part.is_xml and part.partname.endswith("numbering.xml"):
            for lvl in part.element.iter(W + "lvl"):
                fmt = lvl.find(W + "numFmt")
                if fmt is not None:
                    lvl.remove(fmt)
                    lvl.append(fmt)          # raw append: after pStyle/lvlText/...
    p = w.P()
    doc.body.add(p)
    run = w.R()
    p.append(run)
    rpr = w.RPr()
    run.append(rpr)
    for el in (w.I(), w.I(), w.Sz(val=28), w.Sz(val=28)):
        rpr.append(el)                        # raw append: duplicates kept
    run.append(w.T("Italic 14 pt with duplicate properties"))
    return _bytes(doc)


def docx_observe_unknown_element() -> bytes:
    from lxml import etree

    doc = WordprocessingDocument.new(minimal=True)
    p = w.P(w.R(w.T("Before the unknown element. ")))
    doc.body.add(p)
    etree.SubElement(p, W_NS + "notAnElement")  # not in the schema, not ignorable
    p.append(w.R(w.T("After it.")))
    return _bytes(doc)


def docx_observe_bad_values() -> bytes:
    doc = WordprocessingDocument.new(minimal=True)
    jc, color = w.Jc(val="center"), w.Color(val="000000")
    doc.body.add(w.P(w.PPr(jc), w.R(w.RPr(color), w.T("Invalid attribute values"))))
    jc.set(W_NS + "val", "middle")  # not an ST_Jc value
    color.set(W_NS + "val", "red")  # not 6 hex digits
    return _bytes(doc)


def xlsx_observe_row_order() -> bytes:
    book = SpreadsheetDocument.new("Order")
    data = book.sheet_data("Order")
    from oox.spreadsheet import row as make_row

    data.append(make_row(3, ["third"]))
    data.append(make_row(1, ["first"]))  # rows must be ascending
    data.append(make_row(2, ["second"]))
    return _bytes(book)


def xlsx_observe_newer_attribute() -> bytes:
    """x:workbookPr/@dateCompatibility is Office 2010 markup in the 2007 namespace
    (LibreOffice writes it).  Office 2007 cannot know it and is not told to ignore it."""
    from oox.ns import x

    book = SpreadsheetDocument.new("Data")
    book.workbook.add(x.WorkbookPr(dateCompatibility=True))
    book.append_row("Data", ["workbookPr/@dateCompatibility is set"])
    return _bytes(book)


# -- observe cases from the XSD cross-check (tools.gen.irdiff, docs/XSD_DIFF.md):
# markup ISO/IEC 29500-4 Transitional allows and the SDK data (so oox) rejects
def xlsx_observe_iso_optional() -> bytes:
    """Chart attributes and elements ISO makes optional and the SDK requires:
    a doughnut without c:holeSize (ISO default 10), a bar chart's c:barDir
    without @val (ISO default "col")."""
    from lxml import etree

    C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["Kind", "Value"])
    for row in (["a", 3], ["b", 5], ["c", 2]):
        book.append_row("Data", row)
    book.add_chart("Data", "doughnut", "A1:B4", anchor="D2", title="No holeSize")
    book.add_chart("Data", "column", "A1:B4", anchor="D18", title="barDir without val")
    for part in book.package.parts.values():
        if not part.content_type.endswith("chart+xml"):
            continue
        root = part.element
        for hole in list(root.iter(C + "holeSize")):
            hole.getparent().remove(hole)
        for bar_dir in root.iter(C + "barDir"):
            etree.strip_attributes(bar_dir, "val")
    return _bytes(book)


def xlsx_observe_iso_paper_size() -> bytes:
    """x:pageSetup/@paperHeight and @paperWidth (ISO/IEC 29500 Transitional;
    not in the SDK data): a custom A5 sheet size."""
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["pageSetup paperWidth=148mm paperHeight=210mm"])
    ws = book.sheet("Data")
    from oox.element import lenient
    from oox.ns import x

    with lenient():
        setup = x.PageSetup(orientation="portrait")
        setup.set("paperWidth", "148mm")
        setup.set("paperHeight", "210mm")
        ws.add(setup)
    return _bytes(book)


def docx_observe_iso_table_band_size() -> bytes:
    """w:tblStyleRowBandSize / w:tblStyleColBandSize in a body table's w:tblPr
    (ISO's CT_TblPr allows them; the SDK only inside a table style)."""
    from oox.element import lenient

    doc = WordprocessingDocument.new()
    doc.add_table([["Header", "Value"], ["a", "1"], ["b", "2"], ["c", "3"]], header_rows=1)
    W = W_NS
    with lenient():
        for tblpr in doc.body.iter(W + "tblPr"):
            # in ISO's order: tblStyle, tblpPr, tblOverlap, bidiVisual, then these
            row, col = w.TblStyleRowBandSize(val=2), w.TblStyleColBandSize(val=1)
            style = tblpr.find(W + "tblStyle")
            if style is not None:
                style.addnext(row)
            else:
                tblpr.insert(0, row)
            row.addnext(col)
    doc.add_paragraph("tblStyleRowBandSize 2, tblStyleColBandSize 1 in the table's tblPr")
    return _bytes(doc)


def pptx_observe_missing_required() -> bytes:
    deck = PresentationDocument.new(layouts="basic")
    s = deck.add_slide()
    deck.add_text_box(s, "spTree without grpSpPr", x=Inches(1), y=Inches(1), width=Inches(6),
                      height=Inches(1))
    tree = s.first("p:cSld").first("p:spTree")
    tree.remove(tree.first("p:grpSpPr"))  # required
    return _bytes(deck)


# -- the corpus ---------------------------------------------------------------------
# word-20: fields compared before and after Word recomputes them (PAGEREF and
# DATE are layout- or clock-dependent, so they aren't compared)
FIELD_CODES = ["SEQ", "REF", "TITLE", "DOCPROPERTY", "CREATEDATE"]
FIELD_RESULTS = ["1", "2", "Figure 2", "above", "Introduction", "3", "1", "Fields and captions",
                 "Example Ltd", "APOLLO", "Y", "14 November 2026"]


def cases() -> list[Case]:
    W, X, P = "word", "excel", "powerpoint"
    return [
        Case("word-01-minimal", W, "Minimal document", "One line of text.", docx_minimal,
             [{"kind": "word.text_contains", "value": "Minimal package"}]),
        Case("word-02-formatting", W, "Character formatting",
             "Centered bold blue title; italic/underlined/red runs; 3 leading spaces kept; "
             "Greek text; a tab and a line break.", docx_formatting,
             [{"kind": "word.text_contains", "value": GREEK},
              {"kind": "word.text_contains", "value": "   three leading spaces"}]),
        Case("word-03-table", W, "Table", "3x3 table, grey borders, shaded bold header.",
             docx_table, [{"kind": "word.tables", "expect": 1},
                          {"kind": "word.text_contains", "value": "South"}]),
        Case("word-04-pictures", W, "Pictures (PNG/JPEG/GIF/BMP)",
             "Four pictures with labels, then a small PNG inline in a sentence.", docx_pictures,
             [{"kind": "word.inline_shapes", "expect": 5}]),
        Case("word-05-styles", W, "Paragraph style with page break",
             "Two pages; page 2 starts with the 'WithBreak' styled paragraph.",
             docx_styles_pagebreak, [{"kind": "word.pages", "expect": 2}]),
        Case("word-06-w14", W, "Office 2010 markup (w14 + mc:Ignorable)",
             "Opens cleanly in Word 2007 too (the w14 setting is ignorable).",
             docx_w14_extension, [{"kind": "word.text_contains", "value": "ligature"}]),
        Case("word-07-odfpy-passwd", W, "odfpy example translation (passwd table)",
             "2-row, 7-column table of passwd fields.", docx_odfpy_passwd,
             [{"kind": "word.tables", "expect": 1},
              {"kind": "word.text_contains", "value": "/home/ada"}]),
        Case("word-08-edit-word-file", W, "Edited Word-authored file",
             "Word's default template with an added paragraph and picture.",
             docx_edit_word_template,
             [{"kind": "word.text_contains", "value": "Appended by oox"},
              {"kind": "word.inline_shapes", "expect": 1}], requires="docx"),
        Case("excel-01-types", X, "Cell value types",
             "Labels in A; 42, -1234.5678, 12345678901, TRUE, padded text, Greek, two-line text in B.",
             xlsx_types,
             [{"kind": "excel.cell", "sheet": "Types", "ref": "B1", "expect": 42},
              {"kind": "excel.cell", "sheet": "Types", "ref": "B2", "expect": -1234.5678},
              {"kind": "excel.cell", "sheet": "Types", "ref": "B4", "expect": True},
              {"kind": "excel.cell", "sheet": "Types", "ref": "B5", "expect": "  padded text  "},
              {"kind": "excel.cell", "sheet": "Types", "ref": "B6", "expect": GREEK}]),
        Case("excel-02-formulas", X, "Formulas",
             "Totals 4.5 and 8.5, sums 5 and 13, 'oox', 'big'. (Excel asks to save on close: "
             "expected, results are calculated on open.)", xlsx_formulas,
             [{"kind": "excel.cell", "sheet": "Calc", "ref": "D2", "expect": 4.5},
              {"kind": "excel.cell", "sheet": "Calc", "ref": "D4", "expect": 13},
              {"kind": "excel.cell", "sheet": "Calc", "ref": "B5", "expect": "oox"},
              {"kind": "excel.cell", "sheet": "Calc", "ref": "D5", "expect": "big"}]),
        Case("excel-03-sheets", X, "Several sheets", "Tabs: First, With spaces, Ελληνικά.",
             xlsx_sheets, [{"kind": "excel.sheets", "expect": ["First", "With spaces", "Ελληνικά"]}]),
        Case("excel-04-pictures", X, "Pictures", "PNG floating at C2, JPEG at C10.", xlsx_pictures,
             [{"kind": "excel.shapes", "sheet": "Pictures", "expect": 2}]),
        Case("excel-05-edit-existing", X, "Edited Excel-shaped file",
             "Bold header 'Name | Quantity | edited by oox'; rows pens/3 and paper/2.",
             xlsx_edit_existing,
             [{"kind": "excel.cell", "sheet": "Data", "ref": "B1", "expect": "Quantity"},
              {"kind": "excel.cell", "sheet": "Data", "ref": "A3", "expect": "paper"},
              {"kind": "excel.cell", "sheet": "Data", "ref": "B3", "expect": 2}]),
        Case("ppt-01-minimal", P, "Minimal deck", "One empty 16:9 slide.", pptx_minimal,
             [{"kind": "ppt.slides", "expect": 1}]),
        Case("ppt-02-placeholders", P, "Layouts and placeholders",
             "Slide 1 centered title + Greek subtitle; slide 2 'Agenda' with 3 bullet levels.",
             pptx_placeholders,
             [{"kind": "ppt.slides", "expect": 2},
              {"kind": "ppt.text_contains", "slide": 2, "value": "by region"},
              {"kind": "ppt.text_contains", "slide": 1, "value": GREEK}]),
        Case("ppt-03-textboxes-pictures", P, "Text boxes and pictures",
             "Title, two text boxes, four pictures in a 2x2 grid.", pptx_textboxes_pictures,
             [{"kind": "ppt.shapes", "slide": 1, "expect": 7},
              {"kind": "ppt.slide_image", "slide": 1, "expect": "picture"}]),
        Case("ppt-04-4x3", P, "4:3 slide size", "Title and two bullets fit a 4:3 slide.",
             pptx_4x3, [{"kind": "ppt.text_contains", "slide": 1, "value": "4:3 slide size"}]),
        Case("ppt-05-edit-powerpoint-file", P, "Edited PowerPoint-authored file",
             "PowerPoint's default template + 'Two Content' slide + slide with a picture.",
             pptx_edit_powerpoint_template,
             [{"kind": "ppt.slides", "expect": 2},
              {"kind": "ppt.text_contains", "slide": 1, "value": "right"}], requires="pptx"),
        # negative controls: must be reported as prompt or error
        # Negative controls.  Passing outcomes: prompt/error (Office complained)
        # or mismatch (Office repaired silently, and content past the cut is
        # gone).  The probes check content that sits after the truncation.
        Case("word-90-control-malformed", W, "CONTROL: malformed document.xml",
             "Word must complain (unreadable content / cannot open).", docx_control_malformed,
             [{"kind": "word.text_contains", "value": "Second line (line break)"}],
             expect_status=CONTROL),
        Case("excel-90-control-malformed", X, "CONTROL: malformed worksheet",
             "Excel must complain (unreadable content / repair).", xlsx_control_malformed,
             [{"kind": "excel.cell", "sheet": "Types", "ref": "B7", "expect": "line 1\nline 2"}],
             expect_status=CONTROL),
        # Word in depth (default parts, styles, properties)
        Case("word-10-defaults", W, "Default parts: styles, settings, fonts, theme, properties",
             "Title, Heading 1 and 2 in Word's heading styles, Normal text in Calibri 11, an "
             "indented List Paragraph, a 2x2 'Table Grid' table with borders; A4 page; "
             "File > Properties: title 'oox default document', author 'oox'.", docx_defaults,
             [{"kind": "word.outline_level", "paragraph": 2, "expect": 1},
              {"kind": "word.outline_level", "paragraph": 4, "expect": 2},
              {"kind": "word.outline_level", "paragraph": 3, "expect": 10},
              {"kind": "word.tables", "expect": 1},
              {"kind": "doc.property", "name": "Title", "expect": "oox default document"},
              {"kind": "doc.property", "name": "Author", "expect": "oox"}]),
        Case("word-11-lists", W, "Bulleted and numbered lists",
             "Bullets on 3 levels; 1. a. b. i. 2. a.; a restarted list 1. 2.; legal numbering "
             "3. 3.1. 3.1.1. 4. -- all indented like Word's own lists.", docx_lists,
             [{"kind": "word.list_string", "paragraph": p, "expect": e} for p, e in
              [(3, "o"), (5, "1."), (6, "a."), (7, "b."), (8, "i."), (9, "2."), (10, "a."),
               (11, "1."), (12, "2."), (13, "3."), (14, "3.1."), (15, "3.1.1."), (16, "4.")]]),
        Case("word-12-sections", W, "Sections, headers and footers",
             "4 pages. Page 1 header 'Cover header', no footer. Even pages: 'Even header', "
             "footer 'Page N of 4' right. Odd pages: 'Odd header' right, footer centered. "
             "Page 4 landscape, header 'Appendix header', footer 'Page 4 of 4'.", docx_sections,
             [{"kind": "word.count", "what": "sections", "expect": 2},
              {"kind": "word.pages", "expect": 4},
              {"kind": "word.orientation", "section": 1, "expect": "portrait"},
              {"kind": "word.orientation", "section": 2, "expect": "landscape"},
              {"kind": "word.header_footer", "part": "header", "section": 1, "type": "first",
               "value": "Cover header"},
              {"kind": "word.header_footer", "part": "header", "section": 1, "type": "default",
               "value": "Odd header"},
              {"kind": "word.header_footer", "part": "header", "section": 1, "type": "even",
               "value": "Even header"},
              {"kind": "word.header_footer", "part": "header", "section": 2, "type": "default",
               "value": "Appendix header"},
              {"kind": "word.fields", "part": "footer", "section": 1, "type": "default",
               "expect": ["PAGE", "NUMPAGES"]},
              {"kind": "word.fields", "part": "footer", "section": 2, "type": "default",
               "expect": ["PAGE", "NUMPAGES"]}]),
        Case("word-13-links-notes", W, "Hyperlinks, bookmark, footnotes, endnote, comment",
             "Blue underlined links: 'Ecma' opens the web site, 'the target heading' jumps to "
             "page 2. Footnotes 1 and 2 at the page bottom, endnote i at the end, a comment by "
             "'oox reviewer' on the last sentence of page 1.", docx_links_notes,
             [{"kind": "word.count", "what": "hyperlinks", "expect": 2},
              {"kind": "word.hyperlink", "index": 1, "expect": "https://ecma-international.org/"},
              {"kind": "word.hyperlink", "index": 2, "expect": "#target"},
              {"kind": "word.count", "what": "bookmarks", "expect": 1},
              {"kind": "word.count", "what": "footnotes", "expect": 2},
              {"kind": "word.count", "what": "endnotes", "expect": 1},
              {"kind": "word.note", "type": "footnote", "index": 2, "value": "Second footnote."},
              {"kind": "word.note", "type": "endnote", "index": 1, "value": "The endnote."},
              {"kind": "word.count", "what": "comments", "expect": 1},
              {"kind": "word.comment", "index": 1, "author": "oox reviewer",
               "value": "Comment text from oox."}]),
        Case("word-14-toc-template", W, "Table of contents and template placeholders",
             "A 'Contents' list (Chapter one, Section 1.1, Section 1.1.1, Chapter two; right-"
             "click > Update Field adds page numbers). No '{{' left anywhere: title 'Template "
             "report', 'Example Ltd' in body, header and footnote.", docx_toc_template,
             [{"kind": "word.count", "what": "toc", "expect": 1},
              {"kind": "word.text_contains", "value": "Split placeholder: Example Ltd"},
              {"kind": "word.text_contains", "value": "Prepared for Example Ltd on 14 November 2026."},
              {"kind": "word.text_lacks", "value": "{{"},
              {"kind": "word.header_footer", "part": "header", "section": 1, "type": "default",
               "value": "Example Ltd -- confidential"},
              {"kind": "word.note", "type": "footnote", "index": 1, "value": "For Example Ltd."}]),
        Case("word-15-tables", W, "Tables: header row, widths, shading, merged cells",
             "Table 1 centered: blue header row, 'Subtotal' spanning three columns (bold, "
             "right). Table 2: 'Group A' spans three rows (yellow), 'X' two rows and two "
             "columns, 'Y' two columns.", docx_tables,
             [{"kind": "word.tables", "expect": 2},
              {"kind": "word.table", "table": 1, "what": "rows", "expect": 4},
              {"kind": "word.table", "table": 1, "what": "cells", "expect": 14},
              {"kind": "word.table_header", "table": 1, "expect": True},
              {"kind": "word.table_cell", "table": 1, "row": 4, "col": 1, "value": "Subtotal"},
              {"kind": "word.table_cell", "table": 1, "row": 4, "col": 2, "value": "460"},
              # 4 w:tr, but the third holds only vertical-merge continuations:
              # Word doesn't count such a row (Office 2007, 2026-11-14)
              {"kind": "word.table", "table": 2, "what": "rows", "expect": 3},
              {"kind": "word.table", "table": 2, "what": "cells", "expect": 6},
              {"kind": "word.table_cell", "table": 2, "row": 1, "col": 1, "value": "Group A"},
              {"kind": "word.table_cell", "table": 2, "row": 4, "col": 2, "value": "Y"}]),
        Case("word-16-floating-pictures", W, "Floating pictures with text wrapping",
             "A picture at the right margin with text around it, a small one with tight "
             "wrapping, a large one behind the text in the middle of the page, and an inline "
             "picture at the end.", docx_floating,
             [{"kind": "word.shapes", "expect": 3},
              {"kind": "word.inline_shapes", "expect": 1},
              {"kind": "word.shape_wraps", "expect": ["behind", "square", "tight"]}]),
        Case("word-17-tracked-changes", W, "Tracked changes (then accepted)",
             "Review > Track Changes is on. 'very ' struck through (Bob), '2025' replaced "
             "by '2026' and ' Really.' inserted (Ada).", docx_tracked,
             [{"kind": "word.count", "what": "revisions", "expect": 4},
              {"kind": "word.revision", "index": 1, "type": "delete", "author": "Bob",
               "value": "very"},
              {"kind": "word.revision", "index": 3, "type": "insert", "author": "Ada",
               "value": "2026"},
              {"kind": "word.revision", "index": 4, "type": "insert", "author": "Ada",
               "value": "Really."},
              # last: accepting changes the open document
              {"kind": "word.text_after", "action": "accept",
               "value": "jumps over the lazy dog in 2026. Really.", "lacks": "very"}]),
        Case("word-18-tracked-changes-reject", W, "Tracked changes (then rejected)",
             "Same file as word-17; the checker rejects all changes.", docx_tracked,
             [{"kind": "word.text_after", "action": "reject",
               "value": "jumps over the very lazy dog in 2025.", "lacks": "Really"}]),
        Case("word-19-formatting-styles", W, "Run and paragraph formatting, custom styles",
             "Paragraph 2 in grey italic 'Note' style, indented: 'Warning:' bold red, 'check' "
             "in Courier New, 'this' underlined, 14 pt, highlighted. Paragraph 3 justified with "
             "a first-line indent; paragraph 4 right aligned; H2O, mc2, small caps, struck.",
             docx_formatting_styles,
             [{"kind": "word.paragraph", "paragraph": 2, "attr": "style", "expect": "Note"},
              {"kind": "word.font", "paragraph": 2, "char": 1, "attr": "bold", "expect": True},
              {"kind": "word.font", "paragraph": 2, "char": 1, "attr": "italic", "expect": True},
              {"kind": "word.font", "paragraph": 2, "char": 1, "attr": "color", "expect": "C00000"},
              {"kind": "word.font", "paragraph": 2, "char": 10, "attr": "font",
               "expect": "Courier New"},
              {"kind": "word.font", "paragraph": 2, "char": 10, "attr": "size", "expect": 10},
              {"kind": "word.font", "paragraph": 2, "char": 10, "attr": "color", "expect": "595959"},
              {"kind": "word.font", "paragraph": 2, "char": 16, "attr": "underline", "expect": True},
              {"kind": "word.font", "paragraph": 2, "char": 16, "attr": "size", "expect": 14},
              {"kind": "word.paragraph", "paragraph": 2, "attr": "space_before", "expect": 6},
              {"kind": "word.paragraph", "paragraph": 2, "attr": "indent_left", "expect": 36},
              {"kind": "word.paragraph", "paragraph": 3, "attr": "align", "expect": "justify"},
              {"kind": "word.paragraph", "paragraph": 3, "attr": "space_after", "expect": 12},
              {"kind": "word.paragraph", "paragraph": 3, "attr": "first_line", "expect": 18},
              {"kind": "word.paragraph", "paragraph": 4, "attr": "align", "expect": "right"},
              {"kind": "word.paragraph", "paragraph": 4, "attr": "indent_left", "expect": 36},
              {"kind": "word.font", "paragraph": 5, "char": 2, "attr": "subscript", "expect": True},
              {"kind": "word.font", "paragraph": 5, "char": 15, "attr": "superscript",
               "expect": True}]),
        Case("word-20-fields-captions", W, "Captions, cross-references, property fields",
             "Captions 'Figure 1: Inserted before the others', 'Figure 2: Sales by region', "
             "'Figure 3: Costs', 'Table 1: Headcount'. 'See Figure 2, which is above this "
             "paragraph...'. Last line: title, Example Ltd, APOLLO, approved, created date. "
             "Select all + F9 should change nothing (except the page number).",
             docx_fields_captions,
             [{"kind": "word.field_results", "codes": FIELD_CODES, "expect": FIELD_RESULTS},
              {"kind": "doc.property", "name": "Company", "expect": "Example Ltd"},
              {"kind": "doc.custom_property", "name": "Project", "expect": "Apollo"},
              {"kind": "doc.custom_property", "name": "Approved", "expect": True},
              # last: Word recomputes every field; results must equal oox's
              {"kind": "word.field_results", "codes": FIELD_CODES, "update": True,
               "expect": FIELD_RESULTS}]),
        Case("word-21-content-controls", W, "Content controls (a form)",
             "Name 'Ada Lovelace', Department 'Finance' (drop-down), Office 'Lisbon' (combo), "
             "Start date 14 November 2026 (date picker), a ticked box, Comments showing grey "
             "'Any comments'. Click each to see the control.", docx_content_controls,
             [{"kind": "word.count", "what": "controls", "expect": 6},
              {"kind": "word.control", "tag": "name", "type": "text", "value": "Ada Lovelace"},
              {"kind": "word.control", "tag": "dept", "type": "dropdown", "value": "Finance"},
              {"kind": "word.control", "tag": "office", "type": "combo", "value": "Lisbon"},
              {"kind": "word.control", "tag": "start", "type": "date",
               "value": "14 November 2026"},
              {"kind": "word.control", "tag": "agree", "value": "\u2612"},
              {"kind": "word.control", "tag": "comments", "type": "text", "value": "Any comments",
               "placeholder": True}]),
        Case("word-22-more-images", W, "TIFF, EMF, WMF and SVG pictures",
             "Four blue pictures with a white ellipse, 4 cm wide (TIFF with a red square, EMF, "
             "WMF; the SVG has a red triangle in Office 2016+, the PNG fallback before).",
             docx_more_images, [{"kind": "word.inline_shapes", "expect": 4}]),
        Case("excel-07-properties-images", X, "Extended and custom properties; TIFF and SVG",
             "Two pictures (TIFF, SVG with PNG fallback). File > Properties: company 'Example "
             "Ltd'; custom: Project 'Apollo', Budget 1250.", xlsx_properties_images,
             [{"kind": "excel.shapes", "sheet": "Sheet1", "expect": 2},
              {"kind": "doc.property", "name": "Company", "expect": "Example Ltd"},
              {"kind": "doc.custom_property", "name": "Project", "expect": "Apollo"},
              {"kind": "doc.custom_property", "name": "Budget", "expect": 1250}]),
        Case("ppt-07-properties-images", P, "Extended and custom properties; EMF and SVG",
             "Two pictures (EMF, SVG with PNG fallback). File > Properties: company 'Example "
             "Ltd'; custom: Project 'Apollo', Approved yes.", pptx_properties_images,
             [{"kind": "ppt.shapes", "slide": 1, "expect": 2},
              {"kind": "doc.property", "name": "Company", "expect": "Example Ltd"},
              {"kind": "doc.custom_property", "name": "Project", "expect": "Apollo"},
              {"kind": "doc.custom_property", "name": "Approved", "expect": True},
              {"kind": "ppt.slide_image", "slide": 1, "expect": "picture"}]),
        Case("excel-08-styles-formats", X, "Cell styles and number formats",
             "Header: bold Arial 12, dark blue on light blue, bordered, centred. Column B shows "
             "1,234.50 / (98.25) in red / 12.5% / \u20ac1,234.50 / 1.23E+03 / 007 / 0.333333333 / "
             "zero / <text>, then dates and times. A19 italic, underlined, wrapped.",
             xlsx_styles_formats,
             [dict({"kind": "excel.text", "sheet": "Formats", "ref": ref, "expect": text},
                   **({"names": True} if code and re.search(r"mmmm|ddd|AM/PM", code) else {}))
              for ref, _, code, text in EXCEL_FORMATS + EXCEL_DATES]
             + [{"kind": "excel.font", "sheet": "Formats", "ref": "A1", "attr": a, "expect": e}
                for a, e in (("bold", True), ("font", "Arial"), ("size", 12),
                             ("color", "1F3864"))]
             + [{"kind": "excel.fill", "sheet": "Formats", "ref": "B1", "expect": "D9E2F3"},
                {"kind": "excel.align", "sheet": "Formats", "ref": "B1", "expect": "center"},
                {"kind": "excel.number_format", "sheet": "Formats", "ref": "B4",
                 "expect": "0.0%"},
                {"kind": "excel.font", "sheet": "Formats", "ref": "A19", "attr": "italic",
                 "expect": True}]),
        Case("excel-09-layout", X, "Column widths, merges, frozen panes, filter, names",
             "Column A 20 wide, B-D 12; row 1 taller; header frozen (scrolling keeps row 1 and "
             "column A); filter buttons on A1:E1; 'Totals' in the Name Box; E2:E4 drop-down "
             "Yes/No; A6:D6 merged.", xlsx_layout,
             [{"kind": "excel.column_width", "sheet": "Sales", "col": "A", "expect": 20},
              {"kind": "excel.column_width", "sheet": "Sales", "col": "C", "expect": 12},
              {"kind": "excel.row_height", "sheet": "Sales", "row": 1, "expect": 24},
              {"kind": "excel.merge_area", "sheet": "Sales", "ref": "B6", "expect": "A6:D6"},
              {"kind": "excel.frozen", "sheet": "Sales", "expect": "B2"},
              {"kind": "excel.autofilter", "sheet": "Sales", "expect": "A1:E4"},
              {"kind": "excel.defined_name", "name": "Totals", "expect": "Sales!$D$2:$D$4"},
              {"kind": "excel.validation", "sheet": "Sales", "ref": "E3", "expect": "Yes,No"},
              {"kind": "excel.cell", "sheet": "Sales", "ref": "D3", "expect": 205}]),
        Case("excel-10-style-existing", X, "Styling an Excel-shaped workbook",
             "C1 'When' bold like A1/B1; C2 2026-11-14; D2 50% on light yellow; A1:B2 unchanged.",
             xlsx_style_existing,
             [{"kind": "excel.font", "sheet": "Data", "ref": "C1", "attr": "bold", "expect": True},
              # the text would be "########": a date is wider than the default column
              {"kind": "excel.number_format", "sheet": "Data", "ref": "C2", "expect": "yyyy-mm-dd"},
              {"kind": "excel.text", "sheet": "Data", "ref": "D2", "expect": "50%"},
              {"kind": "excel.fill", "sheet": "Data", "ref": "D2", "expect": "FFF2CC"},
              {"kind": "excel.text", "sheet": "Data", "ref": "B2", "expect": "3"},
              {"kind": "excel.font", "sheet": "Data", "ref": "A1", "attr": "bold", "expect": True}]),
        Case("excel-11-names-autofit", X, "English names on any locale; autofit columns",
             "Column B: 14 November 2026 / Saturday / 3:07 PM / Nov 14, 2026 / 1,234,567.89 / "
             "2026-11-14 15:07:09, in English whatever the regional settings; no ####.",
             xlsx_names_autofit,
             [{"kind": "excel.text", "sheet": "Names", "ref": ref, "expect": text} for ref, text in
              (("B1", "14 November 2026"), ("B2", "Saturday"), ("B3", "3:07 PM"),
               ("B4", "Nov 14, 2026"), ("B5", "1,234,567.89"), ("B6", "2026-11-14 15:07:09"))]
             + [{"kind": "excel.column_width", "sheet": "Names", "col": "B", "expect": 21}]),
        Case("excel-12-conditional", X, "Conditional formatting",
             "Sales > 95 red, 30-50 yellow, top 2 bold; Margin colour scale; Units data bars; "
             "Trend arrows; 'South' italic, repeated regions blue; rows where Trend > Units green.",
             xlsx_conditional,
             [{"kind": "excel.cf_count", "sheet": "Rules", "expect": 9},
              {"kind": "excel.cf_types", "sheet": "Rules", "ref": "B2",
               "expect": ["cellIs", "cellIs", "expression", "top10"]},
              {"kind": "excel.cf_types", "sheet": "Rules", "ref": "C2",
               "expect": ["colorScale", "expression"]},
              {"kind": "excel.cf_types", "sheet": "Rules", "ref": "D2",
               "expect": ["dataBar", "expression"]},
              {"kind": "excel.cf_types", "sheet": "Rules", "ref": "E2",
               "expect": ["expression", "iconSet"]},
              {"kind": "excel.cf_types", "sheet": "Rules", "ref": "A3",
               "expect": ["containsText", "duplicateValues", "expression"]},
              {"kind": "excel.screenshot", "sheet": "Rules", "ref": "A1:E11",
               "expect": "picture"}]),
        Case("excel-13-tables", X, "Excel tables with a totals row",
             "SalesTable (blue, banded) with filter buttons and a totals row: Total / 295 / "
             "109.8333333 / 2; G1 = SUM(SalesTable[Q1]) = 295; a second table 'Lookup' in "
             "another style.", xlsx_tables,
             [{"kind": "excel.tables", "sheet": "Data", "expect": ["Lookup", "SalesTable"]},
              {"kind": "excel.table", "sheet": "Data", "name": "SalesTable", "attr": "ref",
               "expect": "A1:D5"},
              {"kind": "excel.table", "sheet": "Data", "name": "SalesTable", "attr": "totals",
               "expect": True},
              {"kind": "excel.table", "sheet": "Data", "name": "SalesTable", "attr": "columns",
               "expect": ["Region", "Q1", "Q2", "Notes"]},
              {"kind": "excel.table", "sheet": "Data", "name": "Lookup", "attr": "style",
               "expect": "TableStyleLight9"},
              {"kind": "excel.cell", "sheet": "Data", "ref": "B5", "expect": 295},
              {"kind": "excel.cell", "sheet": "Data", "ref": "C5", "expect": 109.83333333333333},
              {"kind": "excel.cell", "sheet": "Data", "ref": "D5", "expect": 2},
              {"kind": "excel.cell", "sheet": "Data", "ref": "G1", "expect": 295}]),
        Case("excel-14-comments-links", X, "Cell comments and hyperlinks",
             "Red triangles on Data!B2 and Summary!C3 (hover to read); the note on A1 stays "
             "open; A4 opens ecma-international.org, A5 jumps to Summary!A1.",
             xlsx_comments_links,
             [{"kind": "excel.comments", "sheet": "Data", "expect": 2},
              {"kind": "excel.comment", "sheet": "Data", "ref": "B2", "author": "Ada Lovelace",
               "value": "Checked by finance"},
              {"kind": "excel.comment", "sheet": "Summary", "ref": "C3",
               "author": "Ada Lovelace", "value": "On the second sheet"},
              {"kind": "excel.hyperlink", "sheet": "Data", "ref": "A4",
               "expect": "https://ecma-international.org/"},
              {"kind": "excel.hyperlink", "sheet": "Data", "ref": "A5", "expect": "#Summary!A1"},
              {"kind": "excel.cell", "sheet": "Data", "ref": "A5", "expect": "Go to Summary"}]),
        Case("excel-15-print-options", X, "Print setup, sheet options, protection",
             "File > Print preview: landscape A4, one page wide, row 1 repeated, gridlines, "
             "header 'Quarterly report', footer 'R&D' / 'Page 1 of 1'. Blue tab, 120% zoom, "
             "sheet protected (password 'secret'); 'Lookup' hidden (Unhide lists it), "
             "'Internal' very hidden.", xlsx_print_options,
             [{"kind": "excel.page", "sheet": "Report", "attr": "orientation",
               "expect": "landscape"},
              {"kind": "excel.page", "sheet": "Report", "attr": "paper", "expect": "a4"},
              {"kind": "excel.page", "sheet": "Report", "attr": "fit_width", "expect": 1},
              {"kind": "excel.page", "sheet": "Report", "attr": "print_area", "expect": "A1:C30"},
              {"kind": "excel.page", "sheet": "Report", "attr": "print_titles", "expect": "1:1"},
              {"kind": "excel.page", "sheet": "Report", "attr": "header_center",
               "expect": "Quarterly report"},
              {"kind": "excel.page", "sheet": "Report", "attr": "footer_left", "expect": "R&&D"},
              {"kind": "excel.page", "sheet": "Report", "attr": "footer_right",
               "expect": "Page &P of &N"},
              {"kind": "excel.sheet", "sheet": "Report", "attr": "tab_color", "expect": "2F5496"},
              {"kind": "excel.sheet", "sheet": "Report", "attr": "protected", "expect": True},
              {"kind": "excel.sheet", "sheet": "Lookup", "attr": "hidden", "expect": True},
              {"kind": "excel.sheet", "sheet": "Internal", "attr": "hidden", "expect": "very"},
              {"kind": "excel.cell", "sheet": "Report", "ref": "A3", "expect": "South"},
              {"kind": "excel.cell", "sheet": "Lookup", "ref": "B1", "expect": "South"}]),
        Case("excel-16-charts", X, "Charts from cell ranges",
             "Sales: clustered columns (title, axis titles, legend below), Q1 pie with "
             "percentages, a line per region without markers, stacked bars. Points: scatter "
             "with lines. Right-click > Select Data shows the cell ranges.", xlsx_charts,
             [{"kind": "excel.charts", "sheet": "Sales", "expect": 4},
              {"kind": "excel.charts", "sheet": "Points", "expect": 1},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "type",
               "expect": "column"},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "title",
               "expect": "Quarterly sales"},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "series",
               "expect": ["Q1", "Q2", "Q3"]},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "values",
               "series": 2, "expect": [135, 110, 84.5, 72]},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "categories",
               "expect": ["North", "South", "West", "East"]},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "legend",
               "expect": "bottom"},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "x_title",
               "expect": "Region"},
              {"kind": "excel.chart", "sheet": "Sales", "index": 1, "attr": "y_title",
               "expect": "EUR k"},
              {"kind": "excel.chart", "sheet": "Sales", "index": 2, "attr": "type",
               "expect": "pie"},
              {"kind": "excel.chart", "sheet": "Sales", "index": 2, "attr": "values",
               "expect": [120, 95, 80, 60]},
              {"kind": "excel.chart", "sheet": "Sales", "index": 2, "attr": "labels",
               "expect": ["percent"]},
              {"kind": "excel.chart", "sheet": "Sales", "index": 3, "attr": "type",
               "expect": "line"},
              {"kind": "excel.chart", "sheet": "Sales", "index": 3, "attr": "series",
               "expect": ["North", "South", "West", "East"]},
              {"kind": "excel.chart", "sheet": "Sales", "index": 3, "attr": "categories",
               "expect": ["Q1", "Q2", "Q3"]},
              {"kind": "excel.chart", "sheet": "Sales", "index": 4, "attr": "type",
               "expect": "bar_stacked"},
              {"kind": "excel.chart", "sheet": "Points", "index": 1, "attr": "type",
               "expect": "scatter_lines"},
              {"kind": "excel.chart", "sheet": "Points", "index": 1, "attr": "categories",
               "expect": [1, 2, 3, 4, 5]},
              {"kind": "excel.chart", "sheet": "Points", "index": 1, "attr": "values",
               "expect": [2.1, 3.9, 6.2, 7.8, 10.1]},
              {"kind": "excel.chart", "sheet": "Points", "index": 1, "attr": "legend",
               "expect": None}]
             + [{"kind": "excel.chart_image", "sheet": "Sales", "index": i, "expect": "picture"}
                for i in (1, 2, 3, 4)]
             + [{"kind": "excel.chart_image", "sheet": "Points", "index": 1,
                 "expect": "picture"}]),
        Case("excel-17-combo-charts", X, "Combination charts, secondary axis",
             "Columns of units with a price line on its own axis at the right (EUR, two "
             "decimals), legend below; then target as an area behind the columns, price as a "
             "line on the right axis (legend: Target, Units, Price).", xlsx_combo_charts,
             [{"kind": "excel.charts", "sheet": "Combo", "expect": 2},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "series",
               "expect": ["Units", "Price"]},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "series_type",
               "series": 1, "expect": "column"},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "series_type",
               "series": 2, "expect": "line_markers"},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "axis",
               "series": 1, "expect": "primary"},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "axis",
               "series": 2, "expect": "secondary"},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "y_title",
               "expect": "Units"},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "y2_title",
               "expect": "EUR"},
              {"kind": "excel.chart", "sheet": "Combo", "index": 1, "attr": "values",
               "series": 2, "expect": [2.5, 2.25, 2.75, 2.4]},
              {"kind": "excel.chart", "sheet": "Combo", "index": 2, "attr": "series",
               "expect": ["Target", "Units", "Price"]},
              {"kind": "excel.chart", "sheet": "Combo", "index": 2, "attr": "series_type",
               "series": 1, "expect": "area"},
              {"kind": "excel.chart", "sheet": "Combo", "index": 2, "attr": "axis",
               "series": 3, "expect": "secondary"},
              {"kind": "excel.chart_image", "sheet": "Combo", "index": 1, "expect": "picture"},
              {"kind": "excel.chart_image", "sheet": "Combo", "index": 2, "expect": "picture"}]),
        Case("excel-20-workbook-protection", X, "Workbook structure protection",
             "Right-clicking a tab: Insert, Delete, Rename, Move are greyed out. Review > "
             "Protect Workbook > Protect Structure asks for the password: 'secret'.",
             xlsx_workbook_protection,
             [{"kind": "excel.tabs", "expect": ["First", "Second"]},
              {"kind": "excel.workbook_protection", "attr": "structure", "expect": True},
              {"kind": "excel.workbook_protection", "attr": "windows", "expect": False},
              {"kind": "excel.unprotect_workbook", "wrong": "wrong", "password": "secret",
               "expect": [True, True, False]}]),
        Case("excel-19-chart-sheet", X, "Chart sheet",
             "Tabs: Data, 'Sales chart' (a whole-tab column chart, value labels, legend "
             "below), Notes. Formulas > Name Manager: NoteCell scoped to Notes.",
             xlsx_chart_sheet,
             [{"kind": "excel.tabs", "expect": ["Data", "Sales chart", "Notes"]},
              {"kind": "excel.chart_sheet", "name": "Sales chart", "attr": "type",
               "expect": "column"},
              {"kind": "excel.chart_sheet", "name": "Sales chart", "attr": "title",
               "expect": "Sales by region"},
              {"kind": "excel.chart_sheet", "name": "Sales chart", "attr": "series",
               "expect": ["Q1", "Q2"]},
              {"kind": "excel.chart_sheet", "name": "Sales chart", "attr": "values",
               "series": 2, "expect": [135, 110, 84.5]},
              {"kind": "excel.chart_sheet", "name": "Sales chart", "attr": "labels",
               "expect": ["value"]},
              {"kind": "excel.cell", "sheet": "Notes", "ref": "A1",
               "expect": "After the chart sheet"}]),
        Case("excel-18-edit-openpyxl-file", X, "Edited openpyxl-written file",
             "openpyxl's workbook (bold header, a bar chart of a-c) with a row 'd 6' and a "
             "line chart of a-d added by oox.", xlsx_edit_openpyxl_file,
             [{"kind": "excel.cell", "sheet": "Data", "ref": "B5", "expect": 6},
              {"kind": "excel.charts", "sheet": "Data", "expect": 2},
              {"kind": "excel.chart", "sheet": "Data", "index": 1, "attr": "title",
               "expect": "openpyxl chart"},
              {"kind": "excel.chart", "sheet": "Data", "index": 1, "attr": "values",
               "expect": [3, 5, 4]},
              {"kind": "excel.chart", "sheet": "Data", "index": 2, "attr": "values",
               "expect": [3, 5, 4, 6]},
              {"kind": "excel.chart_image", "sheet": "Data", "index": 2, "expect": "picture"}],
             requires="openpyxl"),
        Case("word-24-text-boxes", W, "Text boxes",
             "Three text boxes: 'Note / Second paragraph' with a black border at the right "
             "margin, the first paragraph flowing to its left; a yellow 'Page box' with a "
             "thick gold border lower on the page, text above and below; 'Plain box' "
             "without border or fill right of the last line.",
             docx_text_boxes,
             [{"kind": "word.text_contains", "value": "flows to its left"},
              {"kind": "word.text_boxes", "expect": ["Note\nSecond paragraph", "Page box",
                                                     "Plain box"]}]),
        Case("word-25-protection-read-only", W, "Read-only protection with a password",
             "Typing is refused; Review > Protect Document > Restrict Formatting and Editing "
             "shows 'No changes (Read only)'. Stop Protection asks for a password: 'secret'.",
             docx_protection_read_only,
             [{"kind": "word.text_contains", "value": "editing is restricted"},
              {"kind": "word.protection", "expect": "readOnly"},
              {"kind": "word.unprotect", "wrong": "wrong", "password": "secret",
               "expect": [True, True, "none"]}]),
        Case("word-26-protection-tracked-sha512", W,
             "Tracked-changes protection, SHA-512 password",
             "Edits are tracked and Track Changes can't be turned off; Stop Protection asks "
             "for a password: 'secret' (hashed with SHA-512, as Word 2010 and later do).",
             docx_protection_tracked_sha512,
             [{"kind": "word.protection", "expect": "trackedChanges"},
              {"kind": "word.unprotect", "wrong": "wrong", "password": "secret",
               "expect": [True, True, "none"]}]),
        Case("word-27-protection-comments", W, "Comments-only protection, no password",
             "Only comments can be inserted; Stop Protection doesn't ask for a password.",
             docx_protection_comments,
             [{"kind": "word.protection", "expect": "comments"},
              {"kind": "word.unprotect", "expect": [None, True, "none"]}]),
        Case("word-23-charts", W, "Inline charts with embedded workbooks",
             "Columns with value labels, stacked area, pie with percentages; right-click > "
             "Edit Data opens each chart's workbook.", docx_charts,
             [{"kind": "word.text_contains", "value": "Charts in Word"},
              {"kind": "word.charts", "expect": 3},
              {"kind": "word.chart", "index": 1, "attr": "type", "expect": "column"},
              {"kind": "word.chart", "index": 1, "attr": "title", "expect": "Sales"},
              {"kind": "word.chart", "index": 1, "attr": "series", "expect": ["North", "South"]},
              {"kind": "word.chart", "index": 2, "attr": "type", "expect": "area_stacked"},
              {"kind": "word.chart", "index": 3, "attr": "type", "expect": "pie"}]),
        Case("ppt-08-charts", P, "Charts on slides with embedded workbooks",
             "Slide 1: line chart with markers. Slide 2: doughnut with percentages next to "
             "100% stacked columns.", pptx_charts,
             [{"kind": "ppt.slides", "expect": 2},
              {"kind": "ppt.charts", "slide": 1, "expect": 1},
              {"kind": "ppt.charts", "slide": 2, "expect": 2},
              {"kind": "ppt.chart", "slide": 1, "index": 1, "attr": "type",
               "expect": "line_markers"},
              {"kind": "ppt.chart", "slide": 1, "index": 1, "attr": "title",
               "expect": "Visits and orders"},
              {"kind": "ppt.chart", "slide": 2, "index": 1, "attr": "type", "expect": "doughnut"},
              {"kind": "ppt.chart", "slide": 2, "index": 2, "attr": "type",
               "expect": "column_percent"},
              {"kind": "ppt.slide_image", "slide": 1, "expect": "picture"},
              {"kind": "ppt.slide_image", "slide": 2, "expect": "picture"}]),
        Case("ppt-09-layouts-tables-notes", P, "Standard layouts, table, notes, shapes",
             "Layout gallery shows Office's 11 layouts. Slide 1 Section Header, 2 Two Content "
             "(notes below), 3 a blue table with 'Total' across two columns (notes), 4 a "
             "picture filling its placeholder, 5 blue rounded box, orange chevron, green circle.",
             pptx_layouts_tables_notes,
             [{"kind": "ppt.slides", "expect": 5},
              {"kind": "ppt.layouts", "expect": sorted([
                  "Title Slide", "Title and Content", "Section Header", "Two Content",
                  "Comparison", "Title Only", "Blank", "Content with Caption",
                  "Picture with Caption", "Title and Vertical Text", "Vertical Title and Text"])},
              {"kind": "ppt.slide_layout", "slide": 1, "expect": "Section Header"},
              {"kind": "ppt.slide_layout", "slide": 2, "expect": "Two Content"},
              {"kind": "ppt.slide_layout", "slide": 4, "expect": "Picture with Caption"},
              {"kind": "ppt.text_contains", "slide": 1, "value": "Part 1"},
              {"kind": "ppt.text_contains", "slide": 2, "value": "detail"},
              {"kind": "ppt.notes", "slide": 2, "expect": "Speaker note for slide 2"},
              {"kind": "ppt.notes", "slide": 3, "expect": "Totals are rounded."},
              {"kind": "ppt.table", "slide": 3, "attr": "rows", "expect": 4},
              {"kind": "ppt.table", "slide": 3, "attr": "cols", "expect": 3},
              {"kind": "ppt.table", "slide": 3, "attr": "cell", "row": 2, "col": 2,
               "expect": "120"},
              {"kind": "ppt.table", "slide": 3, "attr": "cell", "row": 4, "col": 1,
               "expect": "Total"},
              {"kind": "ppt.picture_placeholder", "slide": 4, "expect": True},
              {"kind": "ppt.shape_types", "slide": 5, "expect": ["roundRect", "chevron", "ellipse"]},
              {"kind": "ppt.text_contains", "slide": 5, "value": "Step 2"}]
             + [{"kind": "ppt.slide_image", "slide": i, "expect": "picture"}
                for i in (1, 2, 3, 4, 5)]),
        Case("ppt-10-links-table-edit", P, "Links and an edited table",
             "Slide 1: 'docs' and 'mail the team' are links (hover 'docs': tooltip Manual); in "
             "the slide show the arrow goes to slide 2, 'Summary' to slide 3. Slide 2: table "
             "Region/Q1/Q2/Q3 with North 125, West 80 (no South), four equal columns, 'see "
             "appendix' across three columns. Slide 3: the arrow returns to slide 1.",
             pptx_links_table_edit,
             [{"kind": "ppt.links", "slide": 1, "expect": sorted([
                 "https://example.org/docs", "mailto:team@example.org", "slide:3"])},
              {"kind": "ppt.action", "slide": 1, "shape": "Next arrow", "expect": "next"},
              {"kind": "ppt.action", "slide": 1, "shape": "To summary", "expect": "hyperlink"},
              {"kind": "ppt.action", "slide": 3, "shape": "Back arrow", "expect": "first"},
              {"kind": "ppt.table", "slide": 2, "attr": "rows", "expect": 4},
              {"kind": "ppt.table", "slide": 2, "attr": "cols", "expect": 4},
              {"kind": "ppt.table", "slide": 2, "attr": "cell", "row": 2, "col": 2,
               "expect": "125"},
              {"kind": "ppt.table", "slide": 2, "attr": "cell", "row": 3, "col": 1,
               "expect": "West"},
              {"kind": "ppt.table", "slide": 2, "attr": "cell", "row": 1, "col": 4,
               "expect": "Q3"},
              {"kind": "ppt.table", "slide": 2, "attr": "cell", "row": 4, "col": 2,
               "expect": "see appendix"}]
             + [{"kind": "ppt.slide_image", "slide": i, "expect": "picture"}
                for i in (1, 2, 3)]),
        Case("ppt-12-sections", P, "Slide sections",
             "PowerPoint 2010+: the slide pane shows sections Opening (Welcome, Agenda), "
             "Regions (North, South), Close (Questions). PowerPoint 2007: five slides, no "
             "sections, no prompt.", pptx_sections,
             [{"kind": "ppt.slides", "expect": 5},
              {"kind": "ppt.sections", "expect": [["Opening", 1, 2], ["Regions", 3, 2],
                                                  ["Close", 5, 1]]},
              {"kind": "ppt.text_contains", "slide": 3, "value": "North"}]),
        Case("ppt-11-edit-python-pptx-file", P, "Edited python-pptx file",
             "python-pptx's deck: title (python-pptx is a link), table Item/Count with a, b, "
             "c and the count of a-b merged ('3' over '5' in one cell), 'Old' column gone.",
             pptx_edit_python_pptx_file,
             [{"kind": "ppt.table", "slide": 1, "attr": "cols", "expect": 2},
              {"kind": "ppt.table", "slide": 1, "attr": "rows", "expect": 4},
              {"kind": "ppt.table", "slide": 1, "attr": "cell", "row": 4, "col": 1,
               "expect": "c"},
              {"kind": "ppt.links", "slide": 1, "expect": ["https://example.org/"]},
              {"kind": "ppt.slide_image", "slide": 1, "expect": "picture"}],
             requires="pptx"),
        Case("excel-06-properties", X, "Document properties", "File > Properties: title "
             "'oox workbook', author 'oox'.", xlsx_properties,
             [{"kind": "doc.property", "name": "Title", "expect": "oox workbook"},
              {"kind": "doc.property", "name": "Author", "expect": "oox"}]),
        Case("ppt-06-properties", P, "Document properties", "File > Properties: title 'oox "
             "deck', author 'oox'.", pptx_properties,
             [{"kind": "doc.property", "name": "Title", "expect": "oox deck"},
              {"kind": "doc.property", "name": "Author", "expect": "oox"}]),
        # Observe cases: no expectation, results tell us how strict Office is.
        Case("word-91-observe-order", W, "OBSERVE: properties out of schema order",
             "Centered? Bold, large text? Or a repair prompt?", docx_observe_order,
             [{"kind": "word.text_contains", "value": "wrong order"}], expect_status=OBSERVE),
        Case("word-92-observe-unknown-element", W, "OBSERVE: unknown element in w: namespace",
             "Both sentences visible? Or a repair prompt?", docx_observe_unknown_element,
             [{"kind": "word.text_contains", "value": "After it."}], expect_status=OBSERVE),
        Case("word-93-observe-bad-values", W, "OBSERVE: invalid attribute values",
             "jc='middle', color='red': ignored, or a repair prompt?", docx_observe_bad_values,
             [{"kind": "word.text_contains", "value": "Invalid attribute values"}],
             expect_status=OBSERVE),
        Case("word-94-observe-sdk-values", W, "OBSERVE: ISO values the SDK rejects",
             "lvlJc='start', table jc='start', cantSplit/tblHeader val='true' (LibreOffice writes them): "
             "list numbered 1. 2.? table header repeats? Or a repair prompt?",
             docx_observe_sdk_values,
             [{"kind": "word.text_contains", "value": "Second item"},
              {"kind": "word.list_string", "paragraph": 1, "expect": "1."},
              {"kind": "word.list_string", "paragraph": 2, "expect": "2."},
              {"kind": "word.tables", "expect": 1}], expect_status=OBSERVE),
        Case("word-96-observe-iso-band-size", W, "OBSERVE: ISO table band sizes in a table",
             "tblStyleRowBandSize/ColBandSize in a body table's properties (ISO allows, the SDK "
             "only in table styles): opens? table intact?", docx_observe_iso_table_band_size,
             [{"kind": "word.tables", "expect": 1},
              {"kind": "word.text_contains", "value": "tblStyleRowBandSize"}],
             expect_status=OBSERVE),
        Case("word-95-observe-libreoffice-structure", W,
             "OBSERVE: LibreOffice's out-of-order numbering level, duplicate run properties",
             "List numbered '1.'? Italic 14 pt text? Or a repair prompt?",
             docx_observe_libreoffice_structure,
             [{"kind": "word.list_string", "paragraph": 1, "expect": "1."},
              {"kind": "word.text_contains", "value": "duplicate properties"}],
             expect_status=OBSERVE),
        Case("excel-91-observe-row-order", X, "OBSERVE: rows out of order",
             "A1 first, A2 second, A3 third? Or a repair prompt?", xlsx_observe_row_order,
             [{"kind": "excel.cell", "sheet": "Order", "ref": "A1", "expect": "first"},
              {"kind": "excel.cell", "sheet": "Order", "ref": "A3", "expect": "third"}],
             expect_status=OBSERVE),
        Case("excel-93-observe-iso-optional", X, "OBSERVE: chart parts ISO makes optional",
             "Doughnut without holeSize, column chart's barDir without val (ISO defaults 10 and "
             "col; the SDK requires them): both charts draw? doughnut hole size? columns upright?",
             xlsx_observe_iso_optional,
             [{"kind": "excel.charts", "sheet": "Data", "expect": 2},
              {"kind": "excel.chart", "sheet": "Data", "index": 1, "attr": "type",
               "expect": "doughnut"},
              {"kind": "excel.chart", "sheet": "Data", "index": 2, "attr": "type",
               "expect": "column"},
              {"kind": "excel.chart_image", "sheet": "Data", "index": 1, "expect": "picture"},
              {"kind": "excel.chart_image", "sheet": "Data", "index": 2, "expect": "picture"}],
             expect_status=OBSERVE),
        Case("excel-94-observe-iso-paper-size", X, "OBSERVE: ISO custom paper size",
             "pageSetup paperWidth=148mm paperHeight=210mm (ISO Transitional, not in the SDK "
             "data): opens? File > Print shows A5-sized pages?", xlsx_observe_iso_paper_size,
             [{"kind": "excel.cell", "sheet": "Data", "ref": "A1",
               "expect": "pageSetup paperWidth=148mm paperHeight=210mm"}],
             expect_status=OBSERVE),
        Case("excel-92-observe-newer-attribute", X, "OBSERVE: Office 2010 attribute, not ignorable",
             "Does Excel 2007 complain about workbookPr/@dateCompatibility?",
             xlsx_observe_newer_attribute,
             [{"kind": "excel.cell", "sheet": "Data", "ref": "A1",
               "expect": "workbookPr/@dateCompatibility is set"}], expect_status=OBSERVE),
        Case("ppt-91-observe-missing-required", P, "OBSERVE: required element missing",
             "Text box shown? Or a repair prompt?", pptx_observe_missing_required,
             [{"kind": "ppt.text_contains", "slide": 1, "value": "without grpSpPr"}],
             expect_status=OBSERVE),
        Case("ppt-90-control-malformed", P, "CONTROL: malformed slide",
             "PowerPoint must complain (repair prompt / cannot open).", pptx_control_malformed,
             [{"kind": "ppt.slides", "expect": 2},
              {"kind": "ppt.text_contains", "slide": 2, "value": "Outlook"}],
             expect_status=CONTROL),
    ]


_EXT = {"word": "docx", "excel": "xlsx", "powerpoint": "pptx"}


def generate(outdir: str | Path) -> dict:
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    entries = []
    for case in cases():
        rel = f"{case.app}/{case.id}.{_EXT[case.app]}"
        entry = {"id": case.id, "app": case.app, "title": case.title,
                 "look_for": case.look_for, "file": rel, "probes": case.probes}
        if case.expect_status == OBSERVE:
            entry["observe"] = True
        elif case.expect_status != ("ok",):
            entry["control"] = True
            entry["expect_status"] = list(case.expect_status)
        try:
            data = case.build()
        except ImportError:
            entry.update(skipped=f"needs the {case.requires} package to build")
            entries.append(entry)
            continue
        (out / rel).parent.mkdir(parents=True, exist_ok=True)
        (out / rel).write_bytes(data)
        entry["sha256"] = hashlib.sha256(data).hexdigest()
        entries.append(entry)
    manifest = {"oox_version": oox.__version__, "grammar_source": oox.GRAMMAR_SOURCE,
                "cases": entries}
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False),
                                       encoding="utf-8")
    _write_checklist(out, manifest)
    here = Path(__file__).parent
    for script in ("check_office.py", "report.py"):
        shutil.copyfile(here / script, out / script)
    return manifest


def _write_checklist(out: Path, manifest: dict) -> None:
    lines = [
        "# Office check: manual checklist", "",
        f"Corpus built with oox {manifest['oox_version']} ({manifest['grammar_source']}).", "",
        "For each file: open it in the matching app and note any **repair/recovery prompt**, "
        "error, or missing content. Record results in `manual-results.csv` (status: ok, prompt, "
        "error, mismatch), one copy per machine/Office version, then run `report.py`.", "",
        "Files marked **CONTROL** are broken on purpose: Office *should* complain. "
        "Record what it actually does (decline any repair offer).", "",
        "Files marked **OBSERVE** break the schema in one specific way; there is no right "
        "answer. Record exactly what happens (prompt or not, and what the content looks like).", "",
    ]
    for app in ("word", "excel", "powerpoint"):
        lines += [f"## {app.capitalize()}", "", "| ✓ | File | What you should see |", "|---|---|---|"]
        for c in manifest["cases"]:
            if c["app"] == app:
                look = c["look_for"] if "skipped" not in c else f"*(skipped: {c['skipped']})*"
                lines.append(f"| ☐ | `{c['file']}` | {look} |")
        lines.append("")
    (out / "CHECKLIST.md").write_text("\n".join(lines), encoding="utf-8")
    with open(out / "manual-results.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["id", "status", "notes"])
        for c in manifest["cases"]:
            if "skipped" not in c:
                wr.writerow([c["id"], "", ""])


if __name__ == "__main__":
    m = generate(sys.argv[1] if len(sys.argv) > 1 else "officecheck-corpus")
    built = sum("sha256" in c for c in m["cases"])
    print(f"{built} files written, {len(m['cases']) - built} skipped")
