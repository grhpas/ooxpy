"""0.6.1 fixes from the Petals Office review: list ids, length arithmetic,
the cost of Section.shown(), related_part() outside a package."""

from __future__ import annotations

import pytest

from oox import WordprocessingDocument
from oox.errors import PackageError
from oox.numbering import Definitions
from oox.opc import Part
from oox.units import Cm, Emu, Inches, Length, Pt
from oox.walk import iter_blocks
from tests.test_resolve import body_append, item, numbered_doc


# -- list ids: one type everywhere ---------------------------------------------------------
def test_list_ids_are_ints_and_find_their_levels():
    doc = numbered_doc()
    for num, level in ((1, 0), (2, 1), (3, 0)):
        item(doc, num, level)
    defs = Definitions(doc)
    fmt = doc.formatting()
    for li in doc.list_labels():
        assert isinstance(li.num_id, int) and isinstance(li.abstract_id, int)
        found = defs.levels(li.num_id)
        assert found is not None and found[0] == li.abstract_id
        assert defs.levels(str(li.num_id)) is found          # the XML's form works too
        numbering = fmt.paragraph(li.paragraph).numbering
        assert numbering is not None and defs.levels(numbering[0]) is found
        assert defs.numbering_of(li.paragraph.first("w:pPr"))[0] == li.num_id
    assert defs.levels(99) is None
    with pytest.raises(TypeError):
        defs.levels(1.0)  # type: ignore[arg-type]


# -- lengths stay lengths --------------------------------------------------------------------
def test_length_arithmetic_keeps_the_unit():
    for value in (Cm(2) + Cm(1), Cm(2) - Cm(1), Cm(2) * 2, 2 * Cm(2), Cm(3) / 2, Cm(3) // 2,
                  -Cm(1), +Cm(1), abs(-Cm(1)), sum([Cm(1), Cm(2)]), Cm(1) + 0, 0 + Cm(1),
                  sum([Cm(1), Inches(1)], Length(0)), Cm(5) % Cm(2)):
        assert isinstance(value, Length), value
    assert Cm(2) + Cm(1) == Cm(3) and Cm(3) / Cm(1) == 3.0 and Cm(3) // Cm(1) == 3
    assert (Cm(2) * 1.5).cm == 3.0 and (Pt(10) + Pt(2)).pt == 12.0
    assert repr(Cm(1)) == "Emu(360000)" and str(Cm(1)) == "360000" and f"{Cm(1)}" == "360000"
    for bad in (lambda: Cm(1) + 5, lambda: 5 + Cm(1), lambda: Cm(1) - 5, lambda: 5 - Cm(1),
                lambda: Cm(1) * Cm(1)):
        with pytest.raises(TypeError):
            bad()
    with pytest.raises(TypeError, match="Pt"):
        Cm(1) + 2.5


def test_a_computed_length_reaches_the_setter_intact():
    doc = WordprocessingDocument.new()
    sec = doc.sections[0]
    sec.set_margins(left=Cm(2) + Cm(1), right=sum([Cm(1), Cm(1)]))
    assert sec.margins["left"] == pytest.approx(Cm(3).pt, abs=0.05)
    assert sec.margins["right"] == pytest.approx(Cm(2).pt, abs=0.05)
    doc.add_picture  # noqa: B018  (sizes take the same values)


# -- Section.shown() -------------------------------------------------------------------------
def _brute_shown(sec, what, kind):
    secs = sec.doc.sections
    idx = next(i for i, s in enumerate(secs) if s.sect_pr is sec.sect_pr)
    for s in reversed(secs[:idx + 1]):
        if s._ref(what, kind) is not None:
            return s._hdr_ftr(what, kind, create=False)
    return None


def test_shown_matches_the_section_list_through_wrappers():
    doc = WordprocessingDocument.new()
    for i in range(12):
        doc.add_paragraph(f"p{i}")
        sec = doc.add_section()
        if i % 4 == 0:
            sec.header_text(f"H{i}")
        if i % 5 == 0:
            sec.footer_text(f"F{i}", kind="first")
    # move a section break into a content control, as LibreOffice does after a TOC
    paras = [p for p in iter_blocks(doc.body, doc.body[0].tag)
             if p.first("w:pPr") is not None and p.first("w:pPr").first("w:sectPr") is not None]
    target = paras[5]
    sdt = body_append(doc, '<w:sdt xmlns:w="http://schemas.openxmlformats.org/'
                           'wordprocessingml/2006/main"><w:sdtContent/></w:sdt>')
    target.addprevious(sdt)
    sdt[0].append(target)
    for sec in doc.sections:
        for what in ("header", "footer"):
            for kind in ("default", "first", "even"):
                assert sec.shown(what, kind) is _brute_shown(sec, what, kind)


def test_shown_reads_without_walking_the_body_for_each_section(monkeypatch):
    doc = WordprocessingDocument.new()
    for i in range(50):
        doc.add_paragraph(f"p{i}")
        doc.add_section().header_text(f"H{i}") if i == 0 else doc.add_section()
    secs = doc.sections
    import oox.sections as mod
    from oox import walk

    calls = []
    real = walk.iter_blocks
    monkeypatch.setattr(walk, "iter_blocks", lambda *a: calls.append(1) or real(*a))
    monkeypatch.setattr(mod, "sections", lambda d: calls.append(1) or [])
    assert secs[0].shown("header") is None
    for sec in secs[1:]:
        assert sec.shown("header") is secs[1].header()
        assert sec.header() is None or sec is secs[1]
    assert not calls  # no whole-body walk per call


def test_shown_on_a_removed_section_still_raises():
    from oox.errors import OoxError

    doc = WordprocessingDocument.new()
    doc.add_paragraph("a")
    doc.add_section()
    first = doc.sections[0]
    p = first.sect_pr.getparent().getparent()
    p.getparent().remove(p)
    with pytest.raises(OoxError, match="no longer part"):
        first.shown()


def test_default_section_shown_without_a_body_sect_pr():
    doc = WordprocessingDocument.new()
    final = doc.body.first("w:sectPr")
    doc.body.remove(final)
    assert doc.sections[-1].shown() is None


# -- small fixes -----------------------------------------------------------------------------
def test_related_part_outside_a_package_says_so():
    doc = WordprocessingDocument.new()
    loose = Part.__new__(Part)
    loose.__dict__.update(doc.main_part.__dict__)
    loose.package = None
    rid = next(iter(doc.main_part.rels))
    with pytest.raises(PackageError, match="not in a package"):
        loose.related_part(rid)
    assert Emu(5) == 5
