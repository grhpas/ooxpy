"""0.6.1 resolver additions: fonts per script, Word's widow control, more run
and paragraph properties, highlight colours, explain(), table cells."""

from __future__ import annotations

import pytest

from oox import WordprocessingDocument
from oox.formatting import HIGHLIGHT_COLORS, Border, Borders
from oox.resolve import RunFormat, TableCellFormat, resolve_conflict, script_slot
from tests.test_resolve import NS, TABLE_STYLE, W, add_styles, body_append, grid, style, xml


# -- fonts per script ----------------------------------------------------------------------
def test_greek_text_uses_the_high_ansi_font():
    doc = WordprocessingDocument.new()
    p = body_append(doc, '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Arial" '
                         'w:eastAsia="MS Mincho" w:cs="Arial Unicode MS"/></w:rPr>'
                         '<w:t>Aα</w:t></w:r></w:p>')
    rf = doc.formatting().run(p.first("w:r"))
    assert (rf.font, rf.font_high_ansi, rf.font_east_asia, rf.font_complex) == \
        ("Calibri", "Arial", "MS Mincho", "Arial Unicode MS")
    assert [rf.font_for(c) for c in "Aαé—漢あשب"] == [
        "Calibri", "Arial", "Arial", "Arial", "MS Mincho", "MS Mincho",
        "Arial Unicode MS", "Arial Unicode MS"]


def test_hint_and_right_to_left_choose_the_font():
    assert script_slot("α") == "font_high_ansi"
    assert script_slot("α", "eastAsia") == "font_east_asia"
    assert script_slot("A", "eastAsia") == "font"
    assert script_slot("A", complex_run=True) == "font_complex"
    doc = WordprocessingDocument.new()
    p = body_append(doc, '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia" w:hAnsi="X" '
                         'w:eastAsia="Y"/><w:rtl/></w:rPr><w:t>“</w:t></w:r></w:p>')
    rf = doc.formatting().run(p.first("w:r"))
    assert (rf.font_hint, rf.right_to_left, rf.font_for("A")) == ("eastAsia", True, None)


def test_theme_high_ansi_font_and_inheritance():
    doc = WordprocessingDocument.new()
    add_styles(doc, style('w:type="paragraph" w:styleId="Greek"',
                          '<w:name w:val="Greek"/><w:rPr><w:rFonts w:ascii="Calibri" '
                          'w:hAnsi="Tahoma"/></w:rPr>'))
    p = body_append(doc, '<w:p><w:pPr><w:pStyle w:val="Greek"/></w:pPr><w:r><w:rPr>'
                         '<w:rFonts w:ascii="Verdana"/></w:rPr><w:t>x</w:t></w:r></w:p>')
    rf = doc.formatting().run(p.first("w:r"))
    assert (rf.font, rf.font_high_ansi) == ("Verdana", "Tahoma")  # each slot on its own


# -- Word's widow control ------------------------------------------------------------------
def test_widow_control_is_on_unless_turned_off():
    doc = WordprocessingDocument.new()
    on = doc.add_paragraph("default")
    off = doc.add_paragraph("off")
    doc.format_paragraph(off, widow_control=False)
    fmt = doc.formatting()
    assert fmt.paragraph(on).widow_control is True
    assert fmt.paragraph(off).widow_control is False


# -- more properties, written and read back ------------------------------------------------
def test_new_run_and_paragraph_properties_round_trip():
    doc = WordprocessingDocument.new()
    p = doc.add_paragraph()
    values = dict(position=3, scale=150, kerning=14, emphasis="dot", outline=True,
                  shadow=True, emboss=True, imprint=False, right_to_left=True,
                  complex_script=True, bold_complex=True, italic_complex=True,
                  size_complex=13, font_high_ansi="Tahoma", font_east_asia="SimSun",
                  font_complex="Arial", font_hint="eastAsia", highlight="00FFFF")
    r = doc.add_run(p, "κείμενο", font="Calibri", **values)
    doc.format_paragraph(p, right_to_left=True, suppress_hyphens=True)
    fmt = doc.formatting()
    rf, pf = fmt.run(r), fmt.paragraph(p)
    for key, value in values.items():
        expected = "cyan" if key == "highlight" else value
        assert getattr(rf, key) == expected, key
    assert (rf.font, rf.highlight_color, rf.font_for("κ")) == ("Calibri", "00FFFF", "Arial")
    assert (pf.right_to_left, pf.suppress_hyphens) == (True, True)
    assert not doc.validate(target="Office2007")
    # as a record: what is read can be written to another run
    copy = {k: v for k, v in rf.as_dict().items() if k not in ("style", "border")}
    r2 = doc.add_run(p, "copy", **copy)
    assert doc.formatting().run(r2) == rf
    off = doc.add_run(p, "x", kerning=0)
    assert doc.formatting().run(off).kerning is None
    with pytest.raises(ValueError, match="highlight"):
        doc.add_run(p, "x", highlight="123456")
    assert HIGHLIGHT_COLORS["darkYellow"] == "808000"


# -- explain -------------------------------------------------------------------------------
def test_explain_names_the_layers():
    doc = WordprocessingDocument.new()
    add_styles(doc, TABLE_STYLE,
               style('w:type="paragraph" w:styleId="Body"',
                     '<w:name w:val="Body"/><w:pPr><w:jc w:val="right"/></w:pPr>'),
               style('w:type="character" w:styleId="Strong"',
                     '<w:name w:val="Strong"/><w:rPr><w:b/></w:rPr>'))
    p = body_append(doc, '<w:p><w:pPr><w:pStyle w:val="Body"/><w:spacing w:after="0"/></w:pPr>'
                         '<w:r><w:rPr><w:rStyle w:val="Strong"/><w:i/></w:rPr><w:t>x</w:t>'
                         '</w:r></w:p>')
    fmt = doc.formatting()
    pe = fmt.explain(p)
    assert pe["align"] == "paragraph style 'Body'" and pe["space_after"] == "direct formatting"
    re_ = fmt.explain(p.first("w:r"))
    assert re_["bold"] == "character style 'Strong'" and re_["italic"] == "direct formatting"
    tbl = body_append(doc, grid(3, 2))
    first = tbl.find(f".//{{{W}}}r")
    assert fmt.explain(first)["bold"] == "table style 'Grid', first row"
    assert "size" in fmt.explain(first) and fmt.explain(first)["size"] == "table style 'Grid'"
    with pytest.raises(ValueError):
        fmt.explain(tbl)


# -- table cells -----------------------------------------------------------------------------
CELL_STYLE = style(
    'w:type="table" w:styleId="Lines"',
    '<w:name w:val="Lines"/><w:tblPr><w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:color="000000"/>'
    '<w:bottom w:val="single" w:sz="4" w:color="000000"/>'
    '<w:insideH w:val="single" w:sz="4" w:color="999999"/>'
    '</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/>'
    '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    '<w:tcPr><w:vAlign w:val="center"/></w:tcPr>'
    '<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders>'
    '<w:bottom w:val="double" w:sz="12" w:color="FF0000"/></w:tcBorders>'
    '<w:shd w:val="clear" w:fill="4472C4"/></w:tcPr></w:tblStylePr>'
    '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/>'
    '</w:tcPr></w:tblStylePr>')


def cells(tbl):
    return [tr.findall(f"{{{W}}}tc") for tr in tbl.findall(f"{{{W}}}tr")]


def test_cell_format_from_the_table_style():
    doc = WordprocessingDocument.new()
    add_styles(doc, CELL_STYLE)
    tbl = body_append(doc, grid(4, 2).replace('w:val="Grid"', 'w:val="Lines"'))
    rows = cells(tbl)
    fmt = doc.formatting()
    head, body1, body2, last = (fmt.cell(r[0]) for r in rows)
    assert isinstance(head, TableCellFormat) and head.style == "Lines"
    assert (head.fill, head.valign) == ("4472C4", "center")
    assert head.conditions == ("firstRow", "firstCol", "nwCell")  # Word's default look
    assert head.borders.top == Border("single", 0.5, "000000", 0.0)
    # the header's double bottom is heavier than the row below's insideH: it wins
    assert head.borders.bottom == Border("double", 1.5, "FF0000", 0.0)
    assert fmt.cell(rows[1][0]).borders.top == head.borders.bottom
    assert (body1.fill, body2.fill) == ("D9E2F3", None)
    assert body1.borders.bottom == Border("single", 0.5, "999999", 0.0)
    assert last.borders.bottom == Border("single", 0.5, "000000", 0.0)
    assert head.borders.left is None and head.borders.right is None
    assert head.margins == {"top": 0.0, "left": 5.4, "bottom": 0.0, "right": 5.4}
    explained = fmt.explain(rows[1][0])
    assert explained["borders.top"].startswith("the cell above (table style 'Lines', first row")
    assert explained["fill"] == "table style 'Lines', odd row band"


def test_direct_cell_formatting_and_conflicts():
    doc = WordprocessingDocument.new()
    add_styles(doc, CELL_STYLE)
    t = doc.add_table([["a", "b"], ["c", "d"]])
    t.element.first("w:tblPr").add(xml(f'<w:tblStyle {NS} w:val="Lines"/>'), replace=True)
    thick = Border("single", 6, "00B050")  # 48 x 1 beats the header's double, 12 x 3
    t.format_cell(1, 0, fill="FFF2CC", valign="bottom", borders={"top": thick, "right": None},
                  margins={"left": 2, "top": 1})
    rows = cells(t.element)
    fmt = doc.formatting()
    c = fmt.cell(rows[1][0])
    assert (c.fill, c.valign) == ("FFF2CC", "bottom")
    assert c.margins == {"top": 1.0, "left": 2.0, "bottom": 0.0, "right": 5.4}
    assert c.borders.top == thick and fmt.cell(rows[0][0]).borders.bottom == thick
    assert c.borders.right is None  # nil wins over the neighbour
    assert fmt.explain(rows[1][0])["fill"] == "direct formatting"
    assert not doc.validate(target="Office2007")
    # written back from what was read
    t.format_cell(1, 1, borders=c.borders.__class__(top=thick), margins=c.margins)
    assert doc.formatting().cell(rows[1][1]).margins == c.margins


def test_border_conflict_rules():
    single = Border("single", 1, "000000")
    assert resolve_conflict(single, Border("double", 1, "000000")).style == "double"  # 8x3 > 8x1
    assert resolve_conflict(Border("single", 3), Border("double", 1)).style == "single"
    # dotted weighs 1 whatever its width
    assert resolve_conflict(Border("dotted", 6), Border("single", 0.25)).style == "single"
    # equal weight: the earlier style in Word's list
    assert resolve_conflict(Border("thick", 0.5), Border("single", 1)).style == "single"
    # equal style: the darker colour, then the first
    assert resolve_conflict(Border("single", 1, "FF0000"), single) == single
    assert resolve_conflict(single, Border("single", 1, "000000")) is single
    assert resolve_conflict("nil", single) is None and resolve_conflict("none", single) == single
    assert resolve_conflict(None, None) is None


def test_cell_spacing_turns_conflicts_off_and_row_exceptions_apply():
    doc = WordprocessingDocument.new()
    add_styles(doc, CELL_STYLE)
    tbl = body_append(doc, grid(2, 2).replace(
        '<w:tblStyle w:val="Grid"/>',
        '<w:tblStyle w:val="Lines"/><w:tblCellSpacing w:w="20" w:type="dxa"/>'
        '<w:tblLook w:val="0620"/>'))  # no row bands: the row's own shading shows
    rows = cells(tbl)
    fmt = doc.formatting()
    assert fmt.cell(rows[1][0]).borders.top == Border("single", 0.5, "999999", 0.0)
    ex = xml(f'<w:tblPrEx {NS}><w:tblBorders><w:insideH w:val="dashed" w:sz="8"/>'
             '</w:tblBorders><w:shd w:val="clear" w:fill="ABCDEF"/></w:tblPrEx>')
    rows[1][0].getparent().insert(0, ex)
    fmt = doc.formatting()
    assert fmt.cell(rows[1][0]).borders.top == Border("dashed", 1.0, None, 0.0)
    assert fmt.cell(rows[1][1]).fill == "ABCDEF"


def test_cell_formats_are_records_and_errors():
    doc = WordprocessingDocument.new()
    t = doc.add_table([["a"]])
    rec = doc.formatting().cell(cells(t.element)[0][0])
    assert isinstance(rec.as_dict()["borders"]["top"], Border) and isinstance(RunFormat().as_dict(), dict)
    with pytest.raises(ValueError):
        t.format_cell(0, 0, borders={"between": Border()})
    with pytest.raises(ValueError):
        t.format_cell(0, 0, margins={"inside": 3})
    assert Borders(top=Border()) and not Borders()
