"""Phase 3: the XSD frontend (tools.gen.xsd_frontend) and the IR diff
(tools.gen.irdiff).  Synthetic schemas cover each construct; the real
comparison runs when OOX_XSD_DIR and OOX_SDK_DATA point at the schemas and
the pinned SDK data."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from tools.gen import irdiff, xsd_frontend
from tools.gen.ir import AttrDef, ChildRef, EnumDef, Namespace, Schema, TypeDef

A_NS = "urn:test:a"
B_NS = "urn:test:b"

A_XSD = f"""<?xml version="1.0"?>
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="{A_NS}" xmlns:b="{B_NS}"
  targetNamespace="{A_NS}" elementFormDefault="qualified" attributeFormDefault="qualified">
  <xsd:import namespace="{B_NS}" schemaLocation="b.xsd"/>
  <xsd:simpleType name="ST_Align">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="left"/><xsd:enumeration value="right"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_Side"><xsd:restriction base="ST_Align"/></xsd:simpleType>
  <xsd:simpleType name="ST_OnOff1">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="on"/><xsd:enumeration value="off"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_OnOff"><xsd:union memberTypes="xsd:boolean ST_OnOff1"/></xsd:simpleType>
  <xsd:attributeGroup name="AG_Common">
    <xsd:attribute name="id" type="xsd:unsignedInt" use="required"/>
  </xsd:attributeGroup>
  <xsd:group name="EG_Runs">
    <xsd:choice><xsd:element name="r" type="CT_Text"/><xsd:element ref="b:note"/></xsd:choice>
  </xsd:group>
  <xsd:complexType name="CT_Text">
    <xsd:simpleContent>
      <xsd:extension base="xsd:string"><xsd:attribute name="lang" type="xsd:string"/></xsd:extension>
    </xsd:simpleContent>
  </xsd:complexType>
  <xsd:complexType name="CT_Base">
    <xsd:sequence><xsd:element name="pPr" type="CT_PPr" minOccurs="0"/></xsd:sequence>
    <xsd:attributeGroup ref="AG_Common"/>
  </xsd:complexType>
  <xsd:complexType name="CT_P">
    <xsd:complexContent>
      <xsd:extension base="CT_Base">
        <xsd:sequence>
          <xsd:group ref="EG_Runs" minOccurs="0" maxOccurs="unbounded"/>
          <xsd:any namespace="##other" minOccurs="0"/>
        </xsd:sequence>
        <xsd:attribute name="align" type="ST_Side"/>
      </xsd:extension>
    </xsd:complexContent>
  </xsd:complexType>
  <xsd:complexType name="CT_PPr">
    <xsd:attribute name="keep" type="ST_OnOff"/>
    <xsd:attribute name="plain" form="unqualified" type="xsd:double"/>
  </xsd:complexType>
  <xsd:element name="p" type="CT_P"/>
  <xsd:element name="box">
    <xsd:complexType><xsd:sequence><xsd:element ref="p" maxOccurs="3"/></xsd:sequence></xsd:complexType>
  </xsd:element>
</xsd:schema>
"""

B_XSD = f"""<?xml version="1.0"?>
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="{B_NS}"
  targetNamespace="{B_NS}" elementFormDefault="qualified">
  <xsd:complexType name="CT_Note"><xsd:attribute name="text" type="xsd:string"/></xsd:complexType>
  <xsd:element name="note" type="CT_Note"/>
</xsd:schema>
"""


@pytest.fixture
def schema(tmp_path):
    (tmp_path / "a.xsd").write_text(A_XSD, encoding="utf-8")
    (tmp_path / "b.xsd").write_text(B_XSD, encoding="utf-8")
    return xsd_frontend.load(tmp_path, prefixes={A_NS: "a", B_NS: "b"})


def test_types_keys_and_content_models(schema):
    p = schema.types["a:CT_P/a:p"]
    # base sequence, then the extension's: (pPr?, (r | b:note)*, any?)
    assert irdiff.render(irdiff.normalize(p.particle)) == "(a:pPr?, (a:r | b:note)*, any?)"
    assert {c.qname: c.key for c in p.children} == {
        "a:pPr": "a:CT_PPr/a:pPr", "a:r": "a:CT_Text/a:r", "b:note": "b:CT_Note/b:note"}
    assert p.any_child
    assert schema.types["a:CT_Text/a:r"].leaf_text
    box = schema.types["a:anon_box/a:box"]
    assert box.particle == ("S", 1, 1, (("E", "a:p", 1, 3),))
    assert schema.warnings == []


def test_attributes_groups_forms_and_enumerations(schema):
    p = {a.qname: a for a in schema.types["a:CT_P/a:p"].attrs}
    assert (p["a:id"].kind, p["a:id"].required) == ("int", True)        # from the group
    assert p["a:align"].kind == "enum" and schema.enums[p["a:align"].enum].values == (
        "left", "right")                                                 # through a restriction
    ppr = {a.qname: a for a in schema.types["a:CT_PPr/a:pPr"].attrs}
    assert set(schema.enums[ppr["a:keep"].enum].values) == {"true", "false", "1", "0", "on", "off"}
    assert ppr["plain"].kind == "float"                                  # form="unqualified"
    assert [a.qname for a in schema.types["b:CT_Note/b:note"].attrs] == ["text"]


def _schema(types: dict[str, TypeDef], enums=None, prefixes=("a",)) -> Schema:
    return Schema("test", [Namespace(p, f"urn:{p}") for p in prefixes], types, enums or {})


def _t(key, particle=None, attrs=(), children=None, any_child=False):
    qname = key.partition("/")[2]
    kids = children if children is not None else [
        ChildRef(f"a:CT_{q.partition(':')[2]}/{q}", q, (), False) for q in sorted(irdiff._qnames(
            particle))]
    return TypeDef(key, qname, "", attrs=list(attrs), children=kids, particle=particle,
                   any_child=any_child)


def test_normalize_treats_equivalent_models_alike():
    n = irdiff.normalize
    e = ("E", "a:x", 0, None)
    assert n(("S", 0, 1, (e,))) == e                                   # (x*)? = x*
    assert n(("S", 0, 1, (("E", "a:x", 1, None),))) == e               # (x+)? = x*
    assert n(("S", 1, 1, (("S", 1, 1, (("E", "a:x", 1, 1), ("E", "a:y", 1, 1))),
                          ("E", "a:z", 0, 1)))) == \
        ("S", 1, 1, (("E", "a:x", 1, 1), ("E", "a:y", 1, 1), ("E", "a:z", 0, 1)))
    assert n(("S", 1, 1, (("E", "w14:x", 0, 1), ("E", "a:y", 1, 1))), {"a"}) == ("E", "a:y", 1, 1)
    assert irdiff._relaxed(("E", "a:x", 1, 1000)) == ("E", "a:x", 1, None)


def test_compare_finds_each_kind_of_difference():
    enum_s = EnumDef("a:ST_Jc", ("left", "right"))
    enum_x = EnumDef("a:ST_Jc2", ("left", "right", "start"))
    sdk = _schema({
        "a:CT_Root/a:root": _t("a:CT_Root/a:root", ("S", 1, 1, (("E", "a:jc", 0, 1),
                                                                ("E", "a:row", 1, 1000)))),
        "a:CT_TblJc/a:jc": _t("a:CT_TblJc/a:jc", None, [AttrDef("a:val", "enum", True, "a:ST_Jc")]),
        "a:CT_Row/a:row": _t("a:CT_Row/a:row", None, [AttrDef("a:h", "int")]),
    }, {"a:ST_Jc": enum_s})
    xsd = _schema({
        "a:CT_Root/a:root": _t("a:CT_Root/a:root", ("S", 1, 1, (("E", "a:jc", 0, 1),
                                                                ("E", "a:row", 1, None)))),
        "a:CT_Jc/a:jc": _t("a:CT_Jc/a:jc", None, [AttrDef("a:val", "enum", False, "a:ST_Jc2")]),
        "a:CT_Row/a:row": _t("a:CT_Row/a:row", None, [AttrDef("a:h", "int"), AttrDef("a:w", "int")]),
        "a:CT_Extra/a:extra": _t("a:CT_Extra/a:extra"),
    }, {"a:ST_Jc2": enum_x})
    xsd.types["a:CT_Root/a:root"].children[0] = ChildRef("a:CT_Jc/a:jc", "a:jc", (0,), True)
    sdk.types["a:CT_Root/a:root"].children[0] = ChildRef("a:CT_TblJc/a:jc", "a:jc", (0,), True)
    result = irdiff.compare(sdk, xsd)
    cats = {(f["category"], f["xsd"]): f for f in result["findings"]}
    assert ("renamed", "a:CT_Jc/a:jc") in cats                  # paired through the parent
    assert cats[("model", "a:CT_Root/a:root")]["kind"] == "limits"
    req = cats[("required", "a:CT_Jc/a:jc")]
    assert req["sdk_required"] and irdiff.stricter(req) == {"sdk"}
    enum = cats[("enum", "a:CT_Jc/a:jc")]
    assert enum["only_xsd"] == ["start"] and irdiff.stricter(enum) == {"sdk"}
    assert cats[("attrs", "a:CT_Row/a:row")]["only_xsd"] == ["a:w"]
    assert ("only_xsd", "a:CT_Extra/a:extra") in cats
    md = irdiff.markdown(result, "sdk", "xsd")
    assert "SDK stricter" in md and "(limits only)" in md


def test_children_under_xsd_any_are_not_differences():
    sdk = _schema({"a:CT_G/a:g": _t("a:CT_G/a:g", ("E", "a:pic", 0, 1))})
    xsd = _schema({"a:CT_G/a:g": _t("a:CT_G/a:g", ("A", 0, None), children=[], any_child=True)})
    result = irdiff.compare(sdk, xsd)
    assert not [f for f in result["findings"] if f["category"] == "children"]
    assert result["stats"]["sdk_children_covered_by_xsd_any"] == 1


XSD_DIR = os.environ.get("OOX_XSD_DIR")
SDK_DATA = os.environ.get("OOX_SDK_DATA")


@pytest.mark.skipif(not (XSD_DIR and SDK_DATA), reason="set OOX_XSD_DIR and OOX_SDK_DATA")
def test_real_schemas_against_the_pinned_sdk_data():
    """The known deviations are found: the ones the Office checks and
    overrides.py already met must show up in the diff."""
    from tools.gen import overrides, sdk_frontend

    sdk = sdk_frontend.load(Path(SDK_DATA or ""))
    overrides.apply(sdk)
    prefixes: dict[str, str] = {}
    for n in sdk.namespaces:
        prefixes.setdefault(n.uri, n.prefix)
    xsd = xsd_frontend.load(Path(XSD_DIR or ""), prefixes=prefixes)
    findings = irdiff.compare(sdk, xsd)["findings"]

    def has(cat, xsd_key, **detail):
        return any(f["category"] == cat and f["xsd"] == xsd_key
                   and all(f.get(k) == v for k, v in detail.items()) for f in findings)

    # word-94 (Office 2007 opened them): ISO allows start/end and true/false
    assert any(f["category"] == "enum" and f["xsd"] == "w:CT_Jc/w:lvlJc"
               and "start" in f["only_xsd"] for f in findings)
    assert any(f["category"] == "enum" and f["xsd"] == "w:CT_OnOff/w:cantSplit"
               and "true" in f["only_xsd"] for f in findings)
    # overrides.py: sml's font is an unordered choice
    assert has("model", "x:CT_Font/x:font", known=True)
