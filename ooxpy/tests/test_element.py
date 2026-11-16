import pytest

from oox import (
    DuplicateChild,
    InvalidAttribute,
    InvalidAttributeValue,
    InvalidChild,
    MissingAttribute,
    lenient,
)
from oox.ns import w, w14


def child_names(el):
    return [c.tag_name for c in el]


def test_children_are_inserted_in_schema_order():
    # w:pPr has 5 content models; a detached one is ordered once attached
    p = w.P(w.PPr(w.Jc(val="center"), w.Spacing(after=240), w.PStyle(val="Title")))
    ppr = p[0]
    assert child_names(ppr) == ["w:pStyle", "w:spacing", "w:jc"]
    ppr.add(w.KeepNext())
    assert child_names(ppr) == ["w:pStyle", "w:keepNext", "w:spacing", "w:jc"]


def test_ambiguous_qname_resolves_from_parent():
    p = w.P(w.PPr(w.RPr(w.B())), w.R(w.RPr(w.I()), w.T("x")))
    assert p[0][0].schema_key == "w:CT_ParaRPr/w:rPr"  # paragraph-mark run props
    assert p[1][0].schema_key == "w:CT_RPr/w:rPr"  # run props


def test_detached_ambiguous_element_is_reordered_on_attach():
    rpr = w.RPr(w.Sz(val=48), w.B())  # model unknown while detached
    run = w.R(rpr)
    assert child_names(run[0]) == ["w:b", "w:sz"]


def test_sdk_alias_binds_content_model():
    assert w.Paragraph is w.P  # unambiguous qname -> plain alias
    p = w.P(w.PPr(w.ParagraphMarkRunProperties(w.B())))
    assert p[0][0].schema_key == "w:CT_ParaRPr/w:rPr"


def test_invalid_child_is_rejected_with_allowed_list():
    with pytest.raises(InvalidChild, match="not allowed in <w:p>"):
        w.P(w.Tbl())
    with lenient():
        assert w.P(w.Tbl()) is not None


def test_singletons_and_replace():
    p = w.P(w.PPr(w.Jc(val="left")))
    with pytest.raises(DuplicateChild):
        p.add(w.PPr())
    p.add(w.PPr(w.Jc(val="right")), replace=True)
    assert len(p) == 1 and p[0][0].get_attr("val") == "right"


def test_attribute_checks():
    with pytest.raises(InvalidAttributeValue, match="not one of"):
        w.Jc(val="middle")  # w:jc is ambiguous: checked against ST_Jc | ST_JcTable
    with pytest.raises(InvalidChild, match="not one of: left, center, right"):
        # valid ST_Jc, but table justification only allows left/center/right;
        # checked once the chain reaches the unambiguous w:tbl
        w.Tbl(w.TblPr(w.Jc(val="both")), w.TblGrid())
    with pytest.raises(MissingAttribute):
        w.Jc()
    with pytest.raises(InvalidAttribute, match="did you mean"):
        w.Spacing(aftr=240)
    with pytest.raises(InvalidAttributeValue):
        w.OutlineLvl(val="high")  # Int32
    # ST_TwipsMeasure is a union incl. "1in"-style strings: accepted for now
    # (pattern validation from the SDK's StringValidator data is phase 2)
    assert w.Spacing(after="1in").get_attr("after") == "1in"


def test_attribute_naming_forms():
    s = w.Spacing(after=240, w_before=120, **{"w:line": 360})
    assert (s.get_attr("after"), s.get_attr("w:before"), s.get_attr("line")) == ("240", "120", "360")
    assert w.B(val=False).get_attr("val") == "0"
    # r:id and w:id on the same element -> prefixed kwargs disambiguate
    ref = w.FooterReference(type="default", r_id="rId7")
    assert ref.get_attr("r:id") == "rId7"


def test_text_and_whitespace_preservation():
    t = w.T("  padded ")
    assert t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert w.T("plain").get("{http://www.w3.org/XML/1998/namespace}space") is None
    with pytest.raises(InvalidChild, match="cannot contain text"):
        w.P(text="nope")


def test_extension_namespace_elements():
    rpr = w.R(w.RPr(w14.Ligatures(val="standard"), w.B()))[0]
    assert child_names(rpr) == ["w:b", "w14:ligatures"]


def test_validate_reports_problems_on_raw_trees():
    p = w.P()
    ppr = w.PPr()
    p.append(ppr)  # raw lxml append/insert bypass oox's ordering and checks
    ppr.append(w.Jc(val="center"))
    ppr.append(w.PStyle(val="x"))
    ppr.append(w.Jc(val="left"))
    problems = p.validate()
    assert any("out of schema order" in m for m in problems)
    assert any("more than one <w:jc>" in m for m in problems)
    assert w.P(w.PPr(w.PStyle(val="x"), w.Jc(val="center"))).validate() == []


def test_odfpy_compatible_api():
    p = w.P()
    r_ = w.R()
    r_.addElement(w.T("hi"))
    p.addElement(r_)
    p.setAttribute("w14:paraId", "1A2B3C4D")
    assert p.getAttribute("w14:paraId") == "1A2B3C4D"


def test_replace_also_works_for_repeatable_children():
    from oox.ns import w

    trpr = w.Tr(w.TrPr())[0]  # CT_TrPr: an unbounded choice
    trpr.add(w.TblHeader(), replace=True)
    trpr.add(w.TblHeader(), replace=True)
    assert len(trpr) == 1
    trpr.add(w.TblHeader())  # without replace= a repeat is appended, as the schema allows
    assert len(trpr) == 2


def test_python_booleans_on_on_off_enumerations():
    from oox.ns import m

    assert m.DegHide(val=True).get("{http://schemas.openxmlformats.org/officeDocument/2006/math}val") == "1"
    assert m.DegHide(val=False).get("{http://schemas.openxmlformats.org/officeDocument/2006/math}val") == "0"


def test_lenient_accepts_invalid_values_and_validate_still_reports_them():
    """lenient() turns every construction check off, values included (as the
    architecture notes always said); validate() reports regardless of it."""
    import oox
    from oox.errors import InvalidAttributeValue
    from oox.ns import w

    with oox.lenient():
        jc = w.Jc(val="middle")
        sz = w.Sz()
        sz.set_attr("val", "big")
        found = jc.validate()
    assert jc.get_attr("val") == "middle" and sz.get_attr("val") == "big"
    assert found and "middle" in found[0]          # inside the block too
    assert "middle" in jc.validate()[0]
    with pytest.raises(InvalidAttributeValue):
        w.Jc(val="middle")
