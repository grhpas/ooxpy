"""Document properties (File > Info): ``docProps/core.xml`` (Dublin Core).

    doc.properties.title = "Quarterly report"
    doc.properties.creator = "Ada"
    doc.properties.created = datetime(2026, 11, 14, 9, 30, tzinfo=timezone.utc)
    print(doc.properties.title)

Works the same for .docx, .xlsx and .pptx.  The part is created on first
write, so documents that never set a property stay byte-identical.  The SDK
data has no schema for core properties (the SDK models them in code), so
this module writes the few well-known elements directly; values are text,
except the three dates.
"""

from __future__ import annotations

import datetime as _dt
from typing import TYPE_CHECKING

from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

from ._generated.parts import PARTS
from .element import PARSER

if TYPE_CHECKING:
    from .document import OoxDocument

CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
XSI = "http://www.w3.org/2001/XMLSchema-instance"

# python name -> (namespace, local name)
FIELDS: dict[str, tuple[str, str]] = {
    "title": (DC, "title"),
    "subject": (DC, "subject"),
    "creator": (DC, "creator"),
    "keywords": (CP, "keywords"),
    "description": (DC, "description"),
    "last_modified_by": (CP, "lastModifiedBy"),
    "revision": (CP, "revision"),
    "category": (CP, "category"),
    "content_status": (CP, "contentStatus"),
    "language": (DC, "language"),
    "identifier": (DC, "identifier"),
    "version": (CP, "version"),
    "created": (DCTERMS, "created"),
    "modified": (DCTERMS, "modified"),
    "last_printed": (CP, "lastPrinted"),
}
# The standard relationship type, plus a variant LibreOffice (and some older
# producers) write: "officedocument" in lower case under the Office path.
# Office accepts both, so we read both (and write the standard one).
CORE_RELTYPES = (
    PARTS["CoreFilePropertiesPart"]["relationship_type"],
    "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties",
)

_DATES = {"created", "modified", "last_printed"}
_W3CDTF = {"created", "modified"}  # carry xsi:type="dcterms:W3CDTF"


def _format_date(value: _dt.datetime) -> str:
    if value.tzinfo is not None:
        value = value.astimezone(_dt.timezone.utc).replace(tzinfo=None)
    return value.replace(microsecond=0).isoformat() + "Z"


def _parse_date(text: str) -> _dt.datetime:
    return _dt.datetime.fromisoformat(text.replace("Z", "+00:00"))


class CoreProperties:
    _doc: OoxDocument

    def __init__(self, doc: OoxDocument) -> None:
        object.__setattr__(self, "_doc", doc)

    # -- part access -------------------------------------------------------------
    def _part(self, create: bool):
        pkg = self._doc.package
        reltype = CORE_RELTYPES[0]
        part = next((pt for rt in CORE_RELTYPES if (pt := pkg.related(None, rt)) is not None), None)
        if part is None and create:
            root = PARSER.makeelement(f"{{{CP}}}coreProperties",
                                      nsmap={"cp": CP, "dc": DC, "dcterms": DCTERMS, "xsi": XSI})
            part = pkg.add_part("/docProps/core.xml",
                                PARTS["CoreFilePropertiesPart"]["content_type"], element=root)
            pkg.relate(None, part, reltype)
        return part

    def __getattr__(self, name: str):
        if name not in FIELDS:
            raise AttributeError(f"unknown property {name!r}; known: {', '.join(FIELDS)}")
        part = self._part(create=False)
        if part is None:
            return None
        ns, local = FIELDS[name]
        el = part.element.find(f"{{{ns}}}{local}")
        if el is None or el.text is None:
            return None
        return _parse_date(el.text) if name in _DATES else el.text

    def __setattr__(self, name: str, value) -> None:
        if name not in FIELDS:
            raise AttributeError(f"unknown property {name!r}; known: {', '.join(FIELDS)}")
        part = self._part(create=True)
        assert part is not None
        root = part.element
        ns, local = FIELDS[name]
        el = root.find(f"{{{ns}}}{local}")
        if value is None:
            if el is not None:
                root.remove(el)
            return
        if el is None:
            el = etree.SubElement(root, f"{{{ns}}}{local}")
        if name in _DATES:
            if not isinstance(value, _dt.datetime):
                raise TypeError(f"{name} must be a datetime")
            el.text = _format_date(value)
            if name in _W3CDTF:
                el.set(f"{{{XSI}}}type", "dcterms:W3CDTF")
        else:
            el.text = str(value)

    def as_dict(self) -> dict[str, object]:
        """Every core property that is set, as ``{name: value}``."""
        return {k: v for k in FIELDS if (v := getattr(self, k)) is not None}



# -- extended properties (docProps/app.xml) ----------------------------------------------------
AP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
OP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

# python name -> element (strings unless listed in _AP_INTS / _AP_BOOLS)
EXTENDED: dict[str, str] = {
    "company": "Company", "manager": "Manager", "application": "Application",
    "app_version": "AppVersion", "template": "Template", "hyperlink_base": "HyperlinkBase",
    "total_time": "TotalTime", "pages": "Pages", "words": "Words", "characters": "Characters",
    "lines": "Lines", "paragraphs": "Paragraphs", "slides": "Slides", "notes": "Notes",
    "doc_security": "DocSecurity", "links_up_to_date": "LinksUpToDate",
    "scale_crop": "ScaleCrop", "shared_doc": "SharedDoc",
}
_AP_INTS = {"total_time", "pages", "words", "characters", "lines", "paragraphs", "slides",
            "notes", "doc_security"}
_AP_BOOLS = {"links_up_to_date", "scale_crop", "shared_doc"}


class ExtendedProperties:
    """``docProps/app.xml``: company, manager, application, template ...

        doc.extended_properties.company = "Example Ltd"

    Counts (pages, words, slides ...) are written by Office when it saves; oox
    reads them but doesn't compute them.  The part is created on first write.
    """

    _doc: OoxDocument

    def __init__(self, doc: OoxDocument) -> None:
        object.__setattr__(self, "_doc", doc)

    def _part(self, create: bool):
        doc = self._doc
        found = doc.package.related(None, PARTS["ExtendedFilePropertiesPart"]["relationship_type"])
        if found is None and create:
            from .ns import ap

            found = doc.package.add_part(
                "/docProps/app.xml", PARTS["ExtendedFilePropertiesPart"]["content_type"],
                element=ap.Properties())
            doc.package.relate(None, found, PARTS["ExtendedFilePropertiesPart"]["relationship_type"])
        return found

    def __getattr__(self, name: str):
        if name not in EXTENDED:
            raise AttributeError(f"unknown property {name!r}; known: {', '.join(EXTENDED)}")
        part = self._part(create=False)
        if part is None:
            return None
        el = part.element.find(f"{{{AP}}}{EXTENDED[name]}")
        if el is None or el.text is None:
            return None
        if name in _AP_INTS:
            return int(el.text)
        if name in _AP_BOOLS:
            return el.text.strip() in ("1", "true")
        return el.text

    def __setattr__(self, name: str, value) -> None:
        if name not in EXTENDED:
            raise AttributeError(f"unknown property {name!r}; known: {', '.join(EXTENDED)}")
        part = self._part(create=True)
        assert part is not None
        root = part.element
        el = root.find(f"{{{AP}}}{EXTENDED[name]}")
        if value is None:
            if el is not None:
                root.remove(el)
            return
        if name in _AP_BOOLS:
            text = "true" if value else "false"
        elif name in _AP_INTS:
            text = str(int(value))
        else:
            text = str(value)
        if el is None:
            from .ns import ap

            el = getattr(ap, EXTENDED[name])(text)
            root.add(el)  # schema order
        else:
            el.text = text

    def as_dict(self) -> dict[str, object]:
        """Every extended (``app.xml``) property that is set, as ``{name: value}``."""
        return {k: v for k in EXTENDED if (v := getattr(self, k)) is not None}


# -- custom properties (docProps/custom.xml) ------------------------------------------------------
_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


class CustomProperties:
    """``docProps/custom.xml`` as a mapping (File > Properties > Custom):

        doc.custom_properties["Project"] = "Apollo"
        doc.custom_properties["Approved"] = True
        doc.custom_properties["Budget"] = 1250.5
        del doc.custom_properties["Draft"]

    Values: str, int, float, bool, datetime.  The part is created on first write.
    """

    def __init__(self, doc: OoxDocument) -> None:
        self._doc = doc

    def _part(self, create: bool):
        rt = PARTS["CustomFilePropertiesPart"]["relationship_type"]
        part = self._doc.package.related(None, rt)
        if part is None and create:
            from .ns import op

            part = self._doc.package.add_part(
                "/docProps/custom.xml", PARTS["CustomFilePropertiesPart"]["content_type"],
                element=op.Properties())
            self._doc.package.relate(None, part, rt)
        return part

    def _items(self):
        part = self._part(create=False)
        return [] if part is None else list(part.element.iterchildren(f"{{{OP}}}property"))

    def _find(self, name: str):
        return next((el for el in self._items() if el.get("name") == name), None)

    @staticmethod
    def _decode(el):
        if len(el) == 0:
            return None
        v = el[0]
        kind, text = v.tag.rpartition("}")[2], v.text or ""
        if kind in ("i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint"):
            return int(text)
        if kind in ("r4", "r8", "decimal"):
            return float(text)
        if kind == "bool":
            return text.strip() in ("1", "true")
        if kind == "filetime":
            return _parse_date(text)
        return text

    def __getitem__(self, name: str):
        el = self._find(name)
        if el is None:
            raise KeyError(name)
        return self._decode(el)

    def get(self, name: str, default=None):
        """The custom property ``name``, or ``default``."""
        el = self._find(name)
        return default if el is None else self._decode(el)

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and self._find(name) is not None

    def __iter__(self):
        return iter([el.get("name") for el in self._items()])

    def __len__(self) -> int:
        return len(self._items())

    def keys(self) -> list[str]:
        """Names of the custom properties, in file order."""
        return [el.get("name") for el in self._items()]

    def as_dict(self) -> dict[str, object]:
        """All custom properties as ``{name: value}`` (typed: str, int, float, bool,
        datetime)."""
        return {el.get("name"): self._decode(el) for el in self._items()}

    def __setitem__(self, name: str, value) -> None:
        from .ns import op, vt

        if isinstance(value, bool):
            v = vt.Bool("true" if value else "false")
        elif isinstance(value, int):
            if not -(2 ** 31) <= value < 2 ** 31:
                raise ValueError("custom integer properties are 32-bit (vt:i4); use a float or str")
            v = vt.I4(str(value))
        elif isinstance(value, float):
            v = vt.R8(repr(value))
        elif isinstance(value, _dt.datetime):
            v = vt.Filetime(_format_date(value))
        elif isinstance(value, str):
            v = vt.Lpwstr(value)
        else:
            raise TypeError("custom property values are str, int, float, bool or datetime")
        el = self._find(name)
        if el is not None:
            for child in list(el):
                el.remove(child)
            el.append(v)
            return
        part = self._part(create=True)
        assert part is not None
        pids = [int(e.get("pid", "1")) for e in self._items()]
        part.element.add(op.Property(v, fmtid=_FMTID, pid=max([1, *pids]) + 1, name=name))

    def __delitem__(self, name: str) -> None:
        el = self._find(name)
        if el is None:
            raise KeyError(name)
        el.getparent().remove(el)
