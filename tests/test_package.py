import io
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from oox import Package, WordprocessingDocument
from oox.ns import w

docx = pytest.importorskip("docx")  # python-docx: an independent reader
DEFAULT_TEMPLATE = Path(docx.__file__).parent / "templates" / "default.docx"


def build() -> WordprocessingDocument:
    doc = WordprocessingDocument.new()
    doc.body.add(
        w.P(w.PPr(w.Jc(val="center")), w.R(w.RPr(w.B()), w.T("Title"))),
        w.P(w.R(w.T("Body text"))),
    )
    return doc


def test_new_document_is_readable_by_python_docx(tmp_path):
    path = tmp_path / "x.docx"
    build().save(path)
    d = docx.Document(str(path))
    assert [p.text for p in d.paragraphs] == ["Title", "Body text"]
    assert d.paragraphs[0].runs[0].bold is True


def test_content_types_first_and_minimal_parts():
    doc = WordprocessingDocument.new(minimal=True)
    doc.body.add(w.P(w.R(w.T("x"))))
    data = io.BytesIO()
    doc.save(data)
    names = zipfile.ZipFile(data).namelist()
    assert names[0] == "[Content_Types].xml"
    assert set(names) == {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}


def test_default_document_parts():
    names = set(zipfile.ZipFile(io.BytesIO(build().package.to_bytes())).namelist())
    assert {"word/styles.xml", "word/settings.xml", "word/fontTable.xml",
            "word/theme/theme1.xml"} <= names
    assert "docProps/core.xml" not in names  # only created when a property is set


def test_save_is_deterministic():
    assert build().package.to_bytes() == build().package.to_bytes()


def test_roundtrip_untouched_parts_are_byte_identical():
    src = zipfile.ZipFile(DEFAULT_TEMPLATE)
    out = zipfile.ZipFile(io.BytesIO(Package.open(DEFAULT_TEMPLATE).to_bytes()))
    for name in src.namelist():
        if name.endswith(".rels") or name == "[Content_Types].xml":
            continue  # regenerated, compared semantically below
        assert out.read(name) == src.read(name), name


def test_roundtrip_edit_preserves_everything_else(tmp_path):
    doc = WordprocessingDocument.open(DEFAULT_TEMPLATE)
    doc.body.add(w.P(w.R(w.T("added by oox"))))
    path = tmp_path / "edited.docx"
    doc.save(path)
    d = docx.Document(str(path))
    assert d.paragraphs[-1].text == "added by oox"
    assert len(d.styles) == len(docx.Document(str(DEFAULT_TEMPLATE)).styles)
    # canonical XML of styles part unchanged
    a = etree.tostring(etree.fromstring(zipfile.ZipFile(DEFAULT_TEMPLATE).read("word/styles.xml")), method="c14n")
    b = etree.tostring(etree.fromstring(zipfile.ZipFile(path).read("word/styles.xml")), method="c14n")
    assert a == b


def test_loaded_document_is_schema_aware():
    doc = WordprocessingDocument.open(DEFAULT_TEMPLATE)
    assert doc.body.schema_key == "w:CT_Body/w:body"
    assert doc.document.validate() == []


def _entries(data):
    z = zipfile.ZipFile(io.BytesIO(data) if isinstance(data, bytes) else data)
    return {n: z.read(n) for n in z.namelist()}


@pytest.mark.parametrize("module,template", [("docx", "default.docx"), ("pptx", "default.pptx")])
def test_read_only_roundtrip_is_byte_identical(module, template):
    """Parsing (reading) every part must not change a single byte on save."""
    mod = pytest.importorskip(module)
    path = Path(mod.__file__).parent / "templates" / template
    pkg = Package.open(path)
    for part in pkg.parts.values():
        if part.is_xml:
            part.element.validate()  # touch everything
    assert _entries(pkg.to_bytes()) == _entries(path)


def test_edit_changes_only_what_it_must():
    pptx = pytest.importorskip("pptx")
    from oox import PresentationDocument

    path = Path(pptx.__file__).parent / "templates" / "default.pptx"
    deck = PresentationDocument.open(path)
    deck.add_text_box(deck.add_slide(deck.layouts[6]), "x", x=0, y=0, width=72, height=72)
    before, after = _entries(path), _entries(deck.package.to_bytes())
    changed = {n for n in before if before[n] != after.get(n)}
    assert changed == {"[Content_Types].xml", "ppt/presentation.xml",
                       "ppt/_rels/presentation.xml.rels"}


def test_ignorable_only_for_ordinary_content():
    from oox.ns import p, w, w14

    doc = WordprocessingDocument.new()
    doc.body.add(w.P(w.R(w.RPr(w14.Ligatures(val="standard")), w.T("x"))))
    assert b'mc:Ignorable="w14"' in doc.main_part.to_bytes()

    # p14 inside an extension list must NOT be declared ignorable
    root = p.Sld(p.CSld(p.SpTree(p.NvGrpSpPr(p.CNvPr(id=1, name=""), p.CNvGrpSpPr(), p.NvPr()),
                                 p.GrpSpPr())))
    from lxml import etree

    from oox.grammar import URI
    from oox.opc import normalize_namespaces

    ext = etree.SubElement(etree.SubElement(root, f"{{{URI['p']}}}extLst"),
                           f"{{{URI['p']}}}ext", uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}")
    etree.SubElement(ext, "{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId", val="1")
    normalize_namespaces(root)
    assert root.get("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable") is None
