"""odfpy API surface on OoxElement and the factories."""

import io

import pytest

from oox import InvalidAttribute, InvalidChild, teletype
from oox.grammar import URI
from oox.ns import w
from oox.odfcompat import OpenDocumentText, Text

W = URI["w"]


def test_factory_keywords_attributes_qattributes_parent_check_grammar():
    assert w.Jc(attributes={"val": "center"}).get_attr("val") == "center"
    assert w.Jc(qattributes={(W, "val"): "right"}).get_attr("val") == "right"
    p = w.P()
    r = w.R(parent=p)
    assert r.getparent() is p
    with pytest.raises(InvalidChild):
        w.P(w.Tbl())
    assert len(w.P(w.Tbl(), check_grammar=False)) == 1


def test_odfpy_attribute_spellings():
    ind = w.Ind(firstline=720, LEFT=360)  # odfpy lowercases, drops dashes
    assert ind.get_attr("firstLine") == "720" and ind.get_attr("left") == "360"
    with pytest.raises(InvalidAttribute):
        w.Ind(firstlin=1)


def test_qname_tagname_nodetype():
    p = w.P()
    assert p.qname == (W, "p")  # same shape as odfpy
    assert p.tagName == p.tag_name == "w:p"
    assert p.nodeType == p.ELEMENT_NODE


def test_dom_navigation_and_text_nodes():
    p = w.P(w.R(w.T("a")), w.R(w.T("b")))
    r1, r2 = p.childNodes
    assert p.firstChild is r1 and p.lastChild is r2
    assert r1.nextSibling is r2 and r2.previousSibling is r1 and r1.parentNode is p
    t = r1[0]
    (node,) = t.childNodes
    assert isinstance(node, Text) and node.nodeType == node.TEXT_NODE and node.data == "a"
    node.data = "A"
    assert t.text == "A" and t.hasChildNodes() and not w.R().hasChildNodes()


def test_raw_dom_mutation_bypasses_grammar_like_odfpy():
    ppr = w.PPr()
    w.P(ppr)
    jc, style = w.Jc(val="center"), w.PStyle(val="x")
    ppr.appendChild(jc)
    ppr.insertBefore(style, jc)
    assert [c.tagName for c in ppr.childNodes] == ["w:pStyle", "w:jc"]
    assert ppr.removeChild(jc) is jc
    new = w.Jc(val="left")
    ppr.appendChild(jc)
    assert ppr.replaceChild(new, jc) is jc and ppr[-1] is new


def test_namespaced_attribute_methods():
    jc = w.Jc(val="left")
    jc.setAttrNS(W, "val", "right")
    assert jc.getAttrNS(W, "val") == "right"
    p = w.P()
    p.setAttribute("w14:paraId", "00AB12CD")
    p.removeAttribute("w14:paraId")
    assert p.getAttribute("w14:paraId") is None
    assert (W, "val") in w.Jc(val="left").allowed_attributes()


def test_get_elements_by_type_and_is_instance_of():
    doc = OpenDocumentText()
    doc.text.addElement(w.P(w.R(w.T("x"))))
    doc.text.addElement(w.P())
    assert len(doc.getElementsByType(w.P)) == 2
    assert doc.text.getElementsByType(w.T)[0].isInstanceOf(w.T)


def test_toxml_writes_to_file():
    buf = io.StringIO()
    w.P().toXml(0, buf)
    assert buf.getvalue().startswith("<w:p ")


def test_teletype_roundtrip():
    p = w.P()
    teletype.addTextToElement(p, "Name:\tAda  Lovelace\nline two")
    assert teletype.extractText(p) == "Name:\tAda  Lovelace\nline two"
    assert [c.tagName for c in p[0]] == ["w:t", "w:tab", "w:t", "w:br", "w:t"]
    assert p.validate() == []


def test_save_to_stdout(capsysbinary):
    OpenDocumentText().save("-")
    assert capsysbinary.readouterr().out[:2] == b"PK"


def test_unqualified_attributes_named_like_keywords():
    """cfRule/@text is unqualified; the documented prefixed form reaches it."""
    from oox.ns import x

    rule = x.CfRule(x.Formula(text="A1>0"), type="containsText", priority=1, x_text="North")
    assert rule.get("text") == "North" and rule.first("x:formula").text == "A1>0"
    assert x.CfRule(type="expression", priority=1, **{"x_text": "y"}).get("text") == "y"
