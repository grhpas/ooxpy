"""Effective formatting (oox.resolve) and list labels (oox.numbering)."""

from __future__ import annotations

import pytest
from lxml import etree

from oox import WordprocessingDocument
from oox.element import PARSER
from oox.formatting import Border, Borders, LineSpacing, Tab
from oox.grammar import URI
from oox.ns import w
from oox.resolve import ParagraphFormat, RunFormat
from oox.units import Cm

W = URI["w"]
NS = f'xmlns:w="{W}"'


def xml(text: str):
    return etree.fromstring(text, PARSER)


def add_styles(doc, *snippets: str) -> None:
    styles = doc.ensure_styles()
    for s in snippets:
        styles.append(xml(s))


def style(attrs: str, body: str) -> str:
    return f"<w:style {NS} {attrs}>{body}</w:style>"


def body_append(doc, snippet: str):
    el = xml(snippet.replace("<w:tbl>", f"<w:tbl {NS}>", 1).replace("<w:p>", f"<w:p {NS}>", 1)
             if snippet.startswith(("<w:tbl>", "<w:p>")) else snippet)
    doc.body.insert(len(doc.body) - 1, el)
    return el


def grid(rows: int, cols: int, look: str = "") -> str:
    cells = "".join(f"<w:tc><w:p><w:r><w:t>{{r}}{c}</w:t></w:r></w:p></w:tc>" for c in range(cols))
    trs = "".join(f"<w:tr>{cells.replace('{r}', str(r))}</w:tr>" for r in range(rows))
    cols_xml = '<w:gridCol w:w="1000"/>' * cols
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/>{look}</w:tblPr>'
            f"<w:tblGrid>{cols_xml}</w:tblGrid>{trs}</w:tbl>")


def cell_run(tbl, r: int, c: int):
    tr = tbl.findall(f"{{{W}}}tr")[r]
    return tr.findall(f"{{{W}}}tc")[c].find(f".//{{{W}}}r")


TABLE_STYLE = style(
    'w:type="table" w:styleId="Grid"',
    '<w:name w:val="Grid"/><w:rPr><w:sz w:val="18"/><w:color w:val="111111"/></w:rPr>'
    '<w:pPr><w:jc w:val="center"/></w:pPr>'
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>'
    '<w:pPr><w:jc w:val="left"/></w:pPr></w:tblStylePr>'
    '<w:tblStylePr w:type="lastRow"><w:rPr><w:i/></w:rPr></w:tblStylePr>'
    '<w:tblStylePr w:type="firstCol"><w:rPr><w:color w:val="00FF00"/></w:rPr></w:tblStylePr>'
    '<w:tblStylePr w:type="band1Horz"><w:rPr><w:shd w:val="clear" w:fill="DDDDDD"/></w:rPr>'
    '</w:tblStylePr>'
    '<w:tblStylePr w:type="band2Horz"><w:rPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:rPr>'
    '</w:tblStylePr>'
    '<w:tblStylePr w:type="nwCell"><w:rPr><w:u w:val="double"/></w:rPr></w:tblStylePr>')


# -- the basics -----------------------------------------------------------------------------
def test_types_and_defaults():
    doc = WordprocessingDocument.new(minimal=True)
    p = doc.add_paragraph("x")
    fmt = doc.formatting()
    assert fmt.paragraph(p) == ParagraphFormat()
    assert fmt.run(p.first("w:r")) == RunFormat()  # Word's own: Times New Roman 10 pt
    assert "Formatting" in repr(fmt)
    with pytest.raises(ValueError):
        fmt.run(w.R())


def test_line_spacing_rules_are_distinct_and_round_trip():
    doc = WordprocessingDocument.new()
    exact, least, multi = (doc.add_paragraph(t) for t in "abc")
    doc.format_paragraph(exact, line_spacing=LineSpacing.exactly(18))
    doc.format_paragraph(least, line_spacing=LineSpacing.at_least(18))
    doc.format_paragraph(multi, line_spacing=1.5)
    fmt = doc.formatting()
    got = [fmt.paragraph(p).line_spacing for p in (exact, least, multi)]
    assert got == [("exact", 18.0), ("at_least", 18.0), LineSpacing.multiple(1.5)]
    copy = doc.add_paragraph("d")
    doc.format_paragraph(copy, line_spacing=got[1])
    assert doc.formatting().paragraph(copy).line_spacing == LineSpacing("at_least", 18.0)
    with pytest.raises(TypeError, match="LineSpacing"):
        doc.format_paragraph(copy, line_spacing="12pt")
    with pytest.raises(ValueError):
        doc.format_paragraph(copy, line_spacing=LineSpacing("double", 2))


def test_points_and_lengths():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("x")
    doc.format_paragraph(p, indent_left=Cm(1), indent_right=12, first_line=-18,
                         space_before=Cm(0.5))
    pf = doc.formatting().paragraph(p)
    assert (pf.indent_left, pf.indent_right, pf.first_line) == (28.35, 12.0, -18.0)
    assert pf.space_before == 14.15
    ind = p.first("w:pPr").first("w:ind")
    assert (ind.get(f"{{{W}}}left"), ind.get(f"{{{W}}}hanging")) == ("567", "360")
    assert Cm(1).pt == pytest.approx(28.346, abs=1e-3) and Cm(1).twips == 567


def test_new_properties_round_trip_and_validate():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("Total\t12")
    r = doc.add_run(p, "!", shading="#ffff00", letter_spacing=1.5, hidden=True,
                    language="de-DE", border=Border("double", 1, "ff0000", 2),
                    double_strike=True, superscript=True)
    tabs = [Tab(72, "right", "dot"), Tab(144)]
    doc.format_paragraph(p, tabs=tabs, borders={"bottom": Border()}, shading="eeeeee",
                         outline_level=2, contextual_spacing=True, widow_control=True)
    fmt = doc.formatting()
    pf, rf = fmt.paragraph(p), fmt.run(r)
    assert pf.tabs == tuple(tabs) and pf.borders == Borders(bottom=Border())
    assert (pf.shading, pf.outline_level, pf.contextual_spacing, pf.widow_control) == \
        ("EEEEEE", 2, True, True)
    assert (rf.shading, rf.letter_spacing, rf.hidden, rf.language) == ("FFFF00", 1.5, True, "de-DE")
    assert rf.border == Border("double", 1.0, "FF0000", 2.0)
    assert (rf.double_strike, rf.superscript, rf.subscript) == (True, True, False)
    q = doc.add_paragraph("copy")
    doc.format_paragraph(q, borders=pf.borders, tabs=pf.tabs)
    assert doc.formatting().paragraph(q).borders == pf.borders
    assert not doc.validate(target="Office2007")


def test_tab_stops_add_up_and_clear():
    doc = WordprocessingDocument.new()
    add_styles(doc, style('w:type="paragraph" w:styleId="Tabbed"',
                          '<w:name w:val="Tabbed"/><w:basedOn w:val="Normal"/>'
                          '<w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/>'
                          '<w:tab w:val="center" w:pos="2880"/></w:tabs></w:pPr>'))
    p = body_append(doc, '<w:p><w:pPr><w:pStyle w:val="Tabbed"/><w:tabs>'
                         '<w:tab w:val="clear" w:pos="1440"/>'
                         '<w:tab w:val="decimal" w:pos="4320" w:leader="dot"/>'
                         '</w:tabs></w:pPr></w:p>')
    assert doc.formatting().paragraph(p).tabs == (Tab(144, "center"), Tab(216, "decimal", "dot"))


# -- toggles ------------------------------------------------------------------------------
def test_character_style_toggles_bold_paragraph_style_resets():
    doc = WordprocessingDocument.new()
    add_styles(doc,
               style('w:type="character" w:styleId="Strong"',
                     '<w:name w:val="Strong"/><w:rPr><w:b/></w:rPr>'),
               style('w:type="character" w:styleId="NotBold"',
                     '<w:name w:val="Not Bold"/><w:rPr><w:b w:val="0"/></w:rPr>'),
               style('w:type="paragraph" w:styleId="Loud"',
                     '<w:name w:val="Loud"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr>'),
               style('w:type="paragraph" w:styleId="Calm"',
                     '<w:name w:val="Calm"/><w:basedOn w:val="Loud"/>'
                     '<w:rPr><w:b w:val="0"/></w:rPr>'))
    p = doc.add_paragraph("", style="Loud")
    strong = doc.add_run(p, "a", style="Strong")       # bold style on bold text: regular
    off = doc.add_run(p, "b", style="NotBold")          # off in a style leaves it alone
    plain = doc.add_run(p, "c")
    direct = doc.add_run(p, "d", style="Strong", bold=True)  # direct is absolute
    fmt = doc.formatting()
    assert [fmt.run(r).bold for r in (strong, off, plain, direct)] == [False, True, True, True]
    assert fmt.run(strong).style == "Strong"
    q = doc.add_paragraph("", style="Calm")             # nearest in the chain wins
    assert doc.formatting().run(doc.add_run(q, "e", style="Strong")).bold is True


# -- tables -------------------------------------------------------------------------------
def test_table_style_conditional_formats_and_default_look():
    doc = WordprocessingDocument.new()
    add_styles(doc, TABLE_STYLE)
    tbl = body_append(doc, grid(4, 3))   # no tblLook: Word's 04A0 (first row, first column)
    fmt = doc.formatting()
    head, first_col, body, last = (fmt.run(cell_run(tbl, *rc)) for rc in
                                   ((0, 1), (1, 0), (1, 1), (3, 1)))
    assert (head.bold, head.color, head.shading) == (True, "FFFFFF", None)
    assert fmt.run(cell_run(tbl, 0, 0)).underline == "double"   # nwCell
    assert (first_col.color, first_col.bold) == ("00FF00", False)  # firstCol after band
    assert (body.size, body.color, body.shading) == (9.0, "111111", "DDDDDD")
    assert fmt.run(cell_run(tbl, 2, 1)).shading == "EEEEEE"
    assert (last.italic, last.shading) == (False, "DDDDDD")      # lastRow is off by default
    assert fmt.paragraph(cell_run(tbl, 1, 1).getparent()).align == "center"
    assert fmt.paragraph(cell_run(tbl, 0, 1).getparent()).align == "left"


def test_table_look_attributes_and_hex():
    doc = WordprocessingDocument.new()
    add_styles(doc, TABLE_STYLE)
    attrs = body_append(doc, grid(3, 2, '<w:tblLook w:firstRow="0" w:lastRow="1" '
                                        'w:firstColumn="0" w:lastColumn="0" w:noHBand="1" '
                                        'w:noVBand="1"/>'))
    hexa = body_append(doc, grid(3, 2, '<w:tblLook w:val="0240"/>'))  # lastRow + noHBand
    fmt = doc.formatting()
    for tbl in (attrs, hexa):
        assert fmt.run(cell_run(tbl, 0, 0)).bold is False
        assert fmt.run(cell_run(tbl, 2, 1)).italic is True
        assert fmt.run(cell_run(tbl, 1, 1)).shading is None


def test_table_style_toggles_and_word_table_rule():
    doc = WordprocessingDocument.new()   # Normal: 11 pt from docDefaults, left aligned
    add_styles(doc, TABLE_STYLE,
               style('w:type="paragraph" w:styleId="Small"',
                     '<w:name w:val="Small"/><w:basedOn w:val="Normal"/>'
                     '<w:rPr><w:sz w:val="16"/></w:rPr>'),
               style('w:type="character" w:styleId="Strong"',
                     '<w:name w:val="Strong"/><w:rPr><w:b/></w:rPr>'))
    normal = doc.styles.find(f"{{{W}}}style[@{{{W}}}styleId='Normal']")
    assert normal.first("w:rPr") is None
    normal.add(w.RPr(w.Sz(val=22)))   # Normal states 11 pt itself
    tbl = body_append(doc, grid(2, 1))
    fmt = doc.formatting()
    head = cell_run(tbl, 0, 0)
    assert fmt.run(head).size == 9.0           # Normal's 11 pt doesn't override the table
    head.addprevious(w.R(w.RPr(w.RStyle(val="Strong")), w.T("x")))
    assert doc.formatting().run(head.getprevious()).bold is False  # table bold, toggled off
    p = cell_run(tbl, 1, 0).getparent()
    p.insert(0, w.PPr(w.PStyle(val="Small")))
    assert doc.formatting().run(cell_run(tbl, 1, 0)).size == 8.0  # a real style overrides
    doc.ensure_settings().add(w.Compat(w.CompatSetting(
        name="overrideTableStyleFontSizeAndJustification",
        uri="http://schemas.microsoft.com/office/word", val="1")))
    assert doc.formatting().run(head).size == 11.0


# -- numbering ----------------------------------------------------------------------------
NUMBERING = (
    f'<w:numbering {NS}>'
    '<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="Outline"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/>'
    '<w:lvlText w:val="%1."/><w:suff w:val="space"/>'
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="ordinal"/>'
    '<w:lvlText w:val="%1-%2"/></w:lvl>'
    '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="(%3)"/><w:lvlRestart w:val="1"/></w:lvl></w:abstractNum>'
    '<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="Outline"/></w:abstractNum>'
    '<w:abstractNum w:abstractNumId="3">'
    '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>'
    '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>'
    '<w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/>'
    '<w:pStyle w:val="Step"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
    '</w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
    '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>'
    '<w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>'
    '</w:numbering>')
NUM_STYLES = (
    style('w:type="numbering" w:styleId="Outline"',
          '<w:name w:val="Outline"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>'),
    style('w:type="paragraph" w:styleId="Step"',
          '<w:name w:val="Step"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr>'
          '<w:ind w:left="567" w:hanging="283"/></w:pPr>'))


def numbered_doc():
    doc = WordprocessingDocument.new()
    root = doc.ensure_numbering()
    for child in xml(NUMBERING):
        root.append(child)
    add_styles(doc, *NUM_STYLES)
    return doc


def item(doc, num: int, level: int, text: str = "x"):
    return body_append(doc, f'<w:p><w:pPr><w:numPr><w:ilvl w:val="{level}"/>'
                            f'<w:numId w:val="{num}"/></w:numPr></w:pPr>'
                            f'<w:r><w:t>{text}</w:t></w:r></w:p>')


def test_num_style_link_restart_and_formats():
    doc = numbered_doc()
    for num, level in ((1, 0), (1, 1), (1, 2), (2, 1), (2, 2),
                       (1, 0), (1, 1), (1, 2), (2, 1), (1, 2)):
        item(doc, num, level)
    items = doc.list_labels()
    # numId 2 uses the Outline numbering style, so it shares numId 1's levels and counters;
    # level 3 restarts only after level 1 (lvlRestart 1), not after level 2
    assert [i.label for i in items] == ["I.", "I-1st", "(1)", "I-2nd", "(2)",
                                        "II.", "II-1st", "(1)", "II-2nd", "(2)"]
    assert {i.abstract_id for i in items} == {1}
    first = items[0]
    assert (first.level, first.num_id, first.format, first.suffix, first.is_bullet) == \
        (0, 1, "upperRoman", "space", False)
    assert items[3].num_id == 2 and items[3].format == "ordinal"


def test_bullets_style_levels_and_numbering_indents():
    doc = numbered_doc()
    bullet = item(doc, 3, 0)
    step = body_append(doc, '<w:p><w:pPr><w:pStyle w:val="Step"/></w:pPr></w:p>')
    roman = item(doc, 1, 0)
    items = doc.list_labels()
    assert (items[0].label, items[0].raw_label, items[0].font, items[0].is_bullet) == \
        ("•", "", "Symbol", True)
    assert (items[1].paragraph, items[1].level, items[1].label) == (step, 1, "1.")
    fmt = doc.formatting()
    assert (fmt.paragraph(roman).indent_left, fmt.paragraph(roman).first_line) == (36.0, -18.0)
    assert fmt.paragraph(roman).numbering == (1, 0)
    # numbering from the style: the style's indent wins over the level's (§17.9)
    assert (fmt.paragraph(step).indent_left, fmt.paragraph(step).first_line) == (28.35, -14.15)
    direct = item(doc, 3, 1)       # direct numbering: the level's indent comes after the style
    direct.first("w:pPr").insert(0, w.PStyle(val="Step"))
    assert doc.formatting().paragraph(direct).indent_left == 72.0
    off = item(doc, 0, 0)          # numId 0 switches a style's numbering off
    off.first("w:pPr").insert(0, w.PStyle(val="Step"))
    assert off not in [i.paragraph for i in doc.list_labels()]
    assert bullet is items[0].paragraph


def test_picture_bullets_and_all_stories():
    doc = numbered_doc()
    root = doc.numbering
    pic = xml(f'<w:numPicBullet {NS} xmlns:v="urn:schemas-microsoft-com:vml" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
              'w:numPicBulletId="0"><w:pict><v:shape><v:imagedata r:id="rId9"/></v:shape>'
              '</w:pict></w:numPicBullet>')
    root.insert(0, pic)
    lvl = root.find(f"{{{W}}}abstractNum[@{{{W}}}abstractNumId='3']/{{{W}}}lvl")
    lvl.append(xml(f'<w:lvlPicBulletId {NS} w:val="0"/>'))
    item(doc, 3, 0)
    sect = doc.sections[0]
    sect.ensure_header().append(xml(
        f'<w:p {NS}><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        '</w:p>'))
    assert len(doc.list_labels()) == 1
    body, header = doc.list_labels(all_stories=True)
    assert body.picture is not None and body.picture.tag.endswith("imagedata")
    assert (header.story, header.label) == ("header", "I.")   # its own counters
