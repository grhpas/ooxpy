"""Lists: writing (new_list/add/restart) and reading (list_labels)."""

import io

import pytest

from oox import WordprocessingDocument, open_document, teletype
from oox.ns import w
from oox.numbering import format_number

docx = pytest.importorskip("docx")


def labels(doc):
    return [(item.label, teletype.extractText(item.paragraph)) for item in doc.list_labels()]


def test_kinds_levels_restart_and_start():
    doc = WordprocessingDocument.new()
    b = doc.new_list("bullet")
    b.add("b1")
    b.add("b1.1", 1)
    d = doc.new_list("decimal")
    d.add("one")
    d.add("one-a", 1)
    d.add("one-b", 1)
    d.add("one-b-i", 2)
    d.add("two")
    d.add("two-a", 1)  # lower levels restart under a new parent
    d.restart().add("again")
    leg = doc.new_list("legal", start=3)
    leg.add("x")
    leg.add("y", 1)
    leg.add("z", 2)
    assert [lab for lab, _ in labels(doc)] == [
        "\u2022", "o", "1.", "a.", "b.", "i.", "2.", "a.", "1.", "3.", "3.1.", "3.1.1."]
    assert doc.validate(target="Office2007") == {}


def test_list_paragraphs_use_list_style_and_numbering_part(tmp_path):
    doc = WordprocessingDocument.new()
    doc.new_list("decimal").add("item")
    path = tmp_path / "l.docx"
    doc.save(path)
    d = docx.Document(str(path))
    p = d.paragraphs[0]
    assert p.style.name == "List Paragraph"
    num_pr = p._p.pPr.numPr
    assert (num_pr.ilvl.val, num_pr.numId.val) == (0, 1)
    assert "word/numbering.xml" in {pt.partname.lstrip("/") for pt in d.part.package.iter_parts()}


def test_shared_abstract_counter_and_start_override():
    """Word semantics: instances of one abstractNum share a counter; a
    startOverride resets it when that instance is first used."""
    doc = WordprocessingDocument.new()
    first = doc.new_list("decimal")
    first.add("a")
    first.add("b")
    restarted = first.restart()
    restarted.add("c")  # 1.
    continued = doc.numbering  # a second plain instance of the same abstractNum
    continued.add(w.Num(w.AbstractNumId(val=first.abstract_id), numId=99))
    p = w.P(w.PPr(w.NumPr(w.Ilvl(val=0), w.NumId(val=99))), w.R(w.T("d")))
    doc.body.add(p)  # continues the shared counter: 2.
    assert [lab for lab, _ in labels(doc)] == ["1.", "2.", "1.", "2."]


def test_numbering_from_style_and_numid_zero():
    doc = WordprocessingDocument.new()
    lst = doc.new_list("decimal")
    doc.styles.add(w.Style(w.Name(val="Numbered"), w.BasedOn(val="Normal"),
                           w.PPr(w.NumPr(w.NumId(val=lst.num_id))),
                           type="paragraph", styleId="Numbered"))
    doc.add_paragraph("via style", style="Numbered")
    doc.add_paragraph("via style too", style="Numbered")
    p = doc.add_paragraph("style, but numId 0 switches it off", style="Numbered")
    p.first("w:pPr").add(w.NumPr(w.NumId(val=0)))
    assert labels(doc) == [("1.", "via style"), ("2.", "via style too")]


def test_labels_survive_save_and_reload():
    doc = WordprocessingDocument.new()
    lst = doc.new_list("legal")
    lst.add("one")
    lst.add("one.one", 1)
    buf = io.BytesIO()
    doc.save(buf)
    assert [lab for lab, _ in labels(open_document(io.BytesIO(buf.getvalue())))] == ["1.", "1.1."]


def test_format_number():
    assert [format_number(n, "lowerLetter") for n in (1, 26, 27, 28)] == ["a", "z", "aa", "bb"]
    assert [format_number(n, "upperRoman") for n in (4, 9, 14, 1999)] == ["IV", "IX", "XIV", "MCMXCIX"]
    assert format_number(7, "decimalZero") == "07" and format_number(3, "none") == ""


def test_errors():
    doc = WordprocessingDocument.new()
    with pytest.raises(ValueError, match="list kind"):
        doc.new_list("roman")
    with pytest.raises(ValueError, match="0-8"):
        doc.new_list().add("x", level=9)
    assert WordprocessingDocument.new(minimal=True).list_labels() == []


# -- 0.6.1: more number formats and symbol fonts ---------------------------------------------
def test_more_number_formats():
    from oox.numbering import format_number as f

    assert [f(n, "russianLower") for n in (1, 10, 28, 29)] == ["а", "к", "я", "аа"]
    assert f(3, "russianUpper") == "В"
    assert [f(n, "hebrew1") for n in (1, 15, 16, 21, 115, 999)] == \
        ["א", "טו", "טז", "כא", "קטו", "תתקצט"]
    assert (f(1, "hebrew2"), f(23, "hebrew2")) == ("א", "אא")
    assert (f(2, "arabicAlpha"), f(3, "arabicAbjad")) == ("ب", "ج")
    assert (f(12, "hindiNumbers"), f(12, "thaiNumbers")) == ("١٢", "๑๒")
    assert (f(1, "aiueoFullWidth"), f(47, "aiueoFullWidth"), f(2, "aiueo")) == ("ア", "ア", "ｲ")
    assert (f(3, "irohaFullWidth"), f(3, "iroha")) == ("ハ", "ﾊ")
    assert f(105, "ideographDigital") == "一〇五"
    assert [f(n, "japaneseCounting") for n in (10, 11, 20, 101, 1234)] == \
        ["十", "十一", "二十", "百一", "千二百三十四"]
    assert [f(n, "chineseCounting") for n in (10, 100, 101, 110)] == \
        ["十", "一百", "一百零一", "一百一十"]
    assert (f(11, "ideographTraditional"), f(12, "ideographZodiac")) == ("甲", "亥")
    assert (f(2, "ganada"), f(2, "chosung"), f(2, "ideographEnclosedCircle")) == ("나", "ㄴ", "㈡")
    assert (f(7, "decimalFullWidth"), f(7, "thaiLetters"), f(7, "hindiVowels")) == ("７", "7", "7")


def test_symbol_fonts_and_unconverted_characters():
    from oox.numbering import needs_symbol_font, readable_text

    assert readable_text("", "Symbol") == "•"
    assert readable_text("a", "Symbol") == "αβ"  # Greek typed in the Symbol font
    assert readable_text("", "Wingdings 2") == "✓☑●"
    assert readable_text("", "Wingdings 3") == "►▲"
    assert readable_text("", "Webdings") == "⏵"
    assert readable_text("", "Arial") == ""  # not a symbol font
    unknown = readable_text("", "Wingdings 2")
    assert needs_symbol_font(unknown) and not needs_symbol_font("•")
