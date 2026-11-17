"""Open Packaging Conventions (ECMA-376 Part 2): the zip container.

Equivalent of the packaging half of odfpy's ``opendocument.py``.  Shared by
.docx, .xlsx and .pptx -- the format-specific classes only decide which parts
exist and how they relate.

Properties worth keeping as the code grows:

* **Lossless round-trip.**  Parts are parsed lazily; a part that is never
  touched is written back byte-for-byte.  Unknown parts and relationships are
  preserved.
* **Deterministic output.**  Fixed zip timestamps and stable ordering, so the
  same document produces the same bytes (diffable, cacheable, testable).
"""

from __future__ import annotations

import copy
import io
import os
import posixpath
import zipfile
import zlib
from collections.abc import MutableMapping
from dataclasses import dataclass
from typing import IO

from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

from . import grammar
from . import strict as _strict
from .element import OoxElement, parse_bytes
from .errors import PackageError

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)


_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


@dataclass(frozen=True)
class Limits:
    """Bounds applied when opening a package (untrusted input).

    The defaults are far above real Office files (the largest parts are
    usually media) but stop zip bombs and pathological archives.  Use
    ``Limits.unlimited()`` only for trusted files.
    """

    max_parts: int = 10_000
    max_part_size: int = 512 * 2**20  # uncompressed bytes per part
    max_total_size: int = 2 * 2**30  # uncompressed bytes for the whole package
    max_ratio: int = 1_000  # uncompressed/compressed, checked for parts > 1 MiB

    @classmethod
    def unlimited(cls) -> Limits:
        """Limits that never trigger, for trusted files larger than the defaults allow."""
        big = 2**62
        return cls(big, big, big, big)


DEFAULT_LIMITS = Limits()


def _reject_dtd(data: bytes, where: str) -> None:
    # OPC (ECMA-376 Part 2) forbids DTDs in parts; they are also the vehicle
    # for entity-expansion attacks.
    head = data[:4096].lstrip(b"\xef\xbb\xbf").lstrip()
    if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
        raise PackageError(f"{where}: DTD declarations are not allowed in OOXML parts")


def _safe_parse(data: bytes, where: str = "part") -> etree._Element:
    _reject_dtd(data, where)
    try:
        return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))
    except etree.XMLSyntaxError as e:
        raise PackageError(f"{where}: malformed XML: {e}") from e


def _c14n(el: etree._Element) -> bytes:
    return etree.tostring(el, method="c14n")


def is_xml_content_type(ct: str) -> bool:
    return ct.endswith("+xml") or ct in ("application/xml", "text/xml")


@dataclass
class Relationship:
    rid: str
    reltype: str
    target: str  # as written: relative to the source part's folder, or a URL
    external: bool = False


class Relationships(MutableMapping):
    """A part's relationships: a mapping of relationship id to ``Relationship``
    (``rels["rId1"]``, ``rels.get(rid)``, ``rels.items()``, ``len(rels)``)."""

    def __init__(self, items: dict[str, Relationship] | None = None) -> None:
        self._items: dict[str, Relationship] = dict(items or {})
        self._original: tuple[bytes, dict] | None = None
        self._original_strict = False

    def __getitem__(self, rid: str) -> Relationship:
        return self._items[rid]

    def __setitem__(self, rid: str, rel: Relationship) -> None:
        if not isinstance(rel, Relationship) or rel.rid != rid:
            raise ValueError("set a Relationship whose rid is the key")
        self._items[rid] = rel

    def __delitem__(self, rid: str) -> None:
        del self._items[rid]

    def __iter__(self):
        return iter(self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __repr__(self) -> str:
        return f"Relationships({list(self._items.values())!r})"

    def add(self, reltype: str, target: str, external: bool = False) -> str:
        """Add a relationship; returns its new id ("rId3")."""
        n = len(self._items) + 1
        while f"rId{n}" in self._items:
            n += 1
        rid = f"rId{n}"
        self._items[rid] = Relationship(rid, reltype, target, external)
        return rid

    def of_type(self, reltype: str) -> list[Relationship]:
        return [r for r in self._items.values() if r.reltype == reltype]

    def to_bytes(self, strict: bool = False) -> bytes:
        if (self._original is not None and self._original[1] == self._items
                and self._original_strict == strict):
            return self._original[0]  # unchanged since load -> byte-identical
        root = etree.Element(f"{{{PR_NS}}}Relationships", nsmap={None: PR_NS})
        for r in self._items.values():
            rtype = _strict.reltype_to_strict(r.reltype) if strict else r.reltype
            e = etree.SubElement(root, f"{{{PR_NS}}}Relationship",
                                 Id=r.rid, Type=rtype, Target=r.target)
            if r.external:
                e.set("TargetMode", "External")
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    @classmethod
    def from_bytes(cls, data: bytes, strict: bool = False) -> Relationships:
        """Parse a .rels part; Strict relationship types are kept in Transitional
        form in memory (``strict=True`` when the package is Strict)."""
        rels = cls()
        for e in _safe_parse(data, "relationships"):
            if not isinstance(e.tag, str):
                continue  # comments / processing instructions
            rid, rtype, target = e.get("Id"), e.get("Type"), e.get("Target")
            if not (rid and rtype and target is not None):
                raise PackageError(f"relationship without Id/Type/Target: {dict(e.attrib)}")
            if rid in rels._items:
                raise PackageError(f"duplicate relationship id {rid}")
            if strict:
                rtype = _strict.reltype_to_transitional(rtype)
            rels._items[rid] = Relationship(rid, rtype, target, e.get("TargetMode") == "External")
        rels._original = (data, {k: Relationship(**vars(v)) for k, v in rels._items.items()})
        rels._original_strict = strict
        return rels


class Part:
    def __init__(self, partname: str, content_type: str, *,
                 element: OoxElement | None = None, blob: bytes | None = None) -> None:
        if not partname.startswith("/"):
            raise PackageError(f"part names are absolute: {partname!r}")
        self.partname = partname
        self.content_type = content_type
        self._element = element
        self._blob = blob
        self._blob_strict = False  # blob uses Strict namespaces (set by Package.open)
        self._strict_percent: set[tuple[str, str]] = set()  # percentages converted on read
        self.rels = Relationships()
        #: the package holding the part (set when it is added or read)
        self.package: Package | None = None

    @property
    def is_xml(self) -> bool:
        """Whether the part holds XML (by content type)."""
        return is_xml_content_type(self.content_type)

    @property
    def element(self) -> OoxElement:
        """Root element (parsed on first access)."""
        if self._element is None:
            if not self.is_xml or self._blob is None:
                raise PackageError(f"{self.partname} is not an XML part")
            _reject_dtd(self._blob, self.partname)
            data = _strict.xml_to_transitional(self._blob) if self._blob_strict else self._blob
            try:
                self._element = parse_bytes(data)
            except etree.XMLSyntaxError as e:
                raise PackageError(f"{self.partname}: malformed XML: {e}") from e
            if self._blob_strict:
                self._strict_percent = _strict.values_to_transitional(self._element)
        return self._element

    @property
    def rels_name(self) -> str:
        """Name of the part's relationships part ("/word/_rels/document.xml.rels")."""
        d, f = posixpath.split(self.partname)
        return posixpath.join(d, "_rels", f + ".rels")

    @property
    def blob(self) -> bytes:
        """The part's bytes, as they would be saved."""
        strict = self.package is not None and self.package.conformance == "strict"
        return self.to_bytes(strict=strict)

    def related_part(self, rid: str) -> Part:
        """The part that relationship ``rid`` of this part points to.  Raises
        ``KeyError`` for an unknown id and ``PackageError`` for an external
        link or a missing target."""
        rel = self.rels[rid]
        if rel.external:
            raise PackageError(f"{rid} of {self.partname} is an external link to {rel.target}")
        name = self.resolve(rel.target)
        part = self.package.parts.get(name) if self.package is not None else None
        if part is None:
            raise PackageError(f"{rid} of {self.partname} points to {name}, which is missing")
        return part

    def resolve(self, target: str) -> str:
        """The part name a relationship ``target`` points to, relative to this part."""
        if target.startswith("/"):
            return target
        return posixpath.normpath(posixpath.join(posixpath.dirname(self.partname), target))

    def to_bytes(self, strict: bool = False) -> bytes:
        """Serialize; ``strict=True`` writes Strict namespace URIs."""
        if self._element is None:  # never parsed
            if self._blob is None:
                return b""
            if not self.is_xml or self._blob_strict == strict:
                return self._blob  # byte-identical
            return _strict.xml_to_strict(self._blob) if strict else \
                _strict.xml_to_transitional(self._blob)
        tree = self._element
        if strict:  # Strict value forms are written from a copy; memory stays Transitional
            tree = copy.deepcopy(self._element)
            _strict.values_to_strict(tree, self._strict_percent)
        if self._blob is not None and self._blob_strict == strict:
            if strict:
                current = _safe_parse(_strict.xml_to_strict(etree.tostring(tree)))
            else:
                current = tree
            if _c14n(current) == _c14n(_safe_parse(self._blob, self.partname)):
                return self._blob  # parsed (read) but not changed -> byte-identical
        normalize_namespaces(tree)
        out = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        return _strict.xml_to_strict(out) if strict else out

    def __repr__(self) -> str:
        return f"<Part {self.partname} {self.content_type}>"


class _PackageRoot:
    """Pseudo source part for package-level relationships (/_rels/.rels)."""

    partname = "/"
    rels_name = "/_rels/.rels"

    def __init__(self) -> None:
        self.rels = Relationships()

    @staticmethod
    def resolve(target: str) -> str:
        return target if target.startswith("/") else "/" + posixpath.normpath(target)


class Package:
    def __init__(self) -> None:
        self.parts: dict[str, Part] = {}
        self.root = _PackageRoot()
        self.defaults: dict[str, str] = {"rels": CT_RELS, "xml": "application/xml"}
        self._ct_original: tuple[bytes, dict, dict] | None = None
        self._by_root: dict[etree._Element, Part] = {}
        #: "transitional" or "strict" -- how the package is written.  Opening
        #: a Strict file sets "strict"; trees are always Transitional in memory.
        self.conformance = "transitional"

    # -- building -------------------------------------------------------------
    def add_part(self, partname: str, content_type: str, element: OoxElement | None = None,
                 blob: bytes | None = None) -> Part:
        """Add a part named ``partname`` ("/word/foo.xml") with ``content_type``, holding an
        element tree or raw bytes; relate it with ``relate``."""
        if partname in self.parts:
            raise PackageError(f"duplicate part {partname}")
        part = Part(partname, content_type, element=element, blob=blob)
        part.package = self
        self.parts[partname] = part
        return part

    def part_of(self, element: etree._Element) -> Part:
        """The part whose XML tree contains ``element``."""
        root = element.getroottree().getroot()
        part = self._by_root.get(root)
        if part is None or part._element is not root or self.parts.get(part.partname) is not part:
            self._by_root = {p._element: p for p in self.parts.values() if p._element is not None}
            part = self._by_root.get(root)
        if part is None:
            raise PackageError("element is not attached to a part of this document")
        return part

    def relate(self, source: Part | None, target: Part, reltype: str) -> str:
        """Add a relationship of type ``reltype`` from ``source`` (None: the package) to
        ``target``; returns the new relationship id."""
        src = source or self.root
        base = posixpath.dirname(src.partname) if source else "/"
        rel = posixpath.relpath(target.partname, base) if source else target.partname.lstrip("/")
        return src.rels.add(reltype, rel)

    def related(self, source: Part | None, reltype: str) -> Part | None:
        """The first part that ``source`` (None: the package) relates to with ``reltype``, or
        None."""
        src = source or self.root
        for r in src.rels.of_type(reltype):
            if not r.external:
                return self.parts.get(src.resolve(r.target))
        return None

    # -- io ---------------------------------------------------------------------
    @classmethod
    def open(cls, file: str | os.PathLike | IO[bytes], limits: Limits = DEFAULT_LIMITS) -> Package:
        """Open an OPC package.  Any problem with the file raises ``PackageError``."""
        head = _peek(file, 8)
        if head == _OLE_MAGIC:
            raise PackageError("this is an OLE compound file: either a password-protected "
                               "(encrypted) Office file or a legacy .doc/.xls/.ppt; oox reads "
                               "OOXML packages only")
        try:
            with zipfile.ZipFile(file) as z:
                return cls._read(z, limits)
        except PackageError:
            raise
        except (zipfile.BadZipFile, zipfile.LargeZipFile, zlib.error, EOFError,
                NotImplementedError, RuntimeError, ValueError, KeyError, OSError) as e:
            raise PackageError(f"not a readable OOXML package: {type(e).__name__}: {e}") from e

    @classmethod
    def _read(cls, z: zipfile.ZipFile, limits: Limits) -> Package:
        infos = [i for i in z.infolist() if not i.filename.endswith("/")]
        if len(infos) > limits.max_parts:
            raise PackageError(f"too many entries ({len(infos)} > {limits.max_parts})")
        seen: set[str] = set()
        total = 0
        for i in infos:
            key = i.filename.lower()  # OPC part names are case-insensitive
            if key in seen:
                raise PackageError(f"duplicate zip entry {i.filename!r}")
            seen.add(key)
            if i.flag_bits & 0x1:
                raise PackageError(f"encrypted zip entry {i.filename!r}")
            if i.file_size > limits.max_part_size:
                raise PackageError(f"{i.filename}: {i.file_size} bytes exceeds the part size limit")
            if i.file_size > 2**20 and i.file_size > limits.max_ratio * max(i.compress_size, 1):
                raise PackageError(f"{i.filename}: compression ratio too high (zip bomb?)")
            total += i.file_size
        if total > limits.max_total_size:
            raise PackageError(f"package expands to {total} bytes, over the limit")

        def read(info: zipfile.ZipInfo) -> bytes:
            # Declared sizes can lie; never read more than the limit allows.
            with z.open(info) as f:
                data = f.read(limits.max_part_size + 1)
            if len(data) > limits.max_part_size:
                raise PackageError(f"{info.filename}: larger than declared / over the limit")
            return data

        by_name = {"/" + i.filename: i for i in infos}
        if "/[Content_Types].xml" not in by_name:
            raise PackageError("not an OPC package: [Content_Types].xml missing")
        pkg = cls()
        ct_bytes = read(by_name["/[Content_Types].xml"])
        overrides = {}
        pkg.defaults = {}
        for e in _safe_parse(ct_bytes, "[Content_Types].xml"):
            if not isinstance(e.tag, str):
                continue
            local = etree.QName(e).localname
            if local == "Default" and e.get("Extension") and e.get("ContentType"):
                pkg.defaults[e.get("Extension").lower()] = e.get("ContentType")
            elif local == "Override" and e.get("PartName") and e.get("ContentType"):
                overrides[e.get("PartName")] = e.get("ContentType")
        pkg._ct_original = (ct_bytes, dict(pkg.defaults), dict(overrides))
        for pn, info in by_name.items():
            if pn == "/[Content_Types].xml" or ("/_rels/" in pn and pn.endswith(".rels")):
                continue
            ext = pn.rpartition(".")[2].lower()
            ctype = overrides.get(pn) or pkg.defaults.get(ext, "application/octet-stream")
            part = pkg.parts[pn] = Part(pn, ctype, blob=read(info))
            part.package = pkg
        if "/_rels/.rels" in by_name:
            root_rels = read(by_name["/_rels/.rels"])
            strict = any(_strict.is_strict_reltype(e.get("Type") or "")
                         for e in _safe_parse(root_rels, "_rels/.rels") if isinstance(e.tag, str))
            pkg.conformance = "strict" if strict else "transitional"
            pkg.root.rels = Relationships.from_bytes(root_rels, strict)
        strict = pkg.conformance == "strict"
        for part in pkg.parts.values():
            part._blob_strict = strict and part.is_xml
            if part.rels_name in by_name:
                part.rels = Relationships.from_bytes(read(by_name[part.rels_name]), strict)
        return pkg

    def _overrides(self) -> dict[str, str]:
        return {pn: part.content_type for pn, part in self.parts.items()
                if self.defaults.get(pn.rpartition(".")[2].lower()) != part.content_type}

    def _content_types(self) -> bytes:
        if self._ct_original is not None:
            data, defaults, overrides = self._ct_original
            # Same Defaults, every part declared with its current type, and no
            # Override left for a removed part -> keep the original bytes.
            declared = lambda pn: overrides.get(pn) or defaults.get(pn.rpartition(".")[2].lower())
            if (defaults == self.defaults
                    and all(declared(pn) == p.content_type for pn, p in self.parts.items())
                    # LibreOffice also lists the .rels files as Overrides
                    and all(pn in self.parts or pn.endswith(".rels") for pn in overrides)):
                return data
        root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
        for ext, ctype in sorted(self.defaults.items()):
            etree.SubElement(root, f"{{{CT_NS}}}Default", Extension=ext, ContentType=ctype)
        for pn in sorted(self.parts):
            part = self.parts[pn]
            if self.defaults.get(pn.rpartition(".")[2].lower()) != part.content_type:
                etree.SubElement(root, f"{{{CT_NS}}}Override", PartName=pn,
                                 ContentType=part.content_type)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def to_bytes(self) -> bytes:
        """The package as bytes (what ``save`` writes)."""
        buf = io.BytesIO()
        self.save(buf)
        return buf.getvalue()

    def save(self, file: str | os.PathLike | IO[bytes]) -> None:
        """Write the package to ``file`` (a path or a binary file object): deterministic zip,
        untouched parts byte for byte."""
        def write(z: zipfile.ZipFile, name: str, data: bytes) -> None:
            info = zipfile.ZipInfo(name.lstrip("/"), _ZIP_EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)

        if self.conformance not in ("transitional", "strict"):
            raise PackageError(f"conformance must be 'transitional' or 'strict', not {self.conformance!r}")
        strict = self.conformance == "strict"
        with zipfile.ZipFile(file, "w") as z:
            write(z, "[Content_Types].xml", self._content_types())  # first, by convention
            if self.root.rels:
                write(z, self.root.rels_name, self.root.rels.to_bytes(strict))
            for pn in sorted(self.parts):
                part = self.parts[pn]
                write(z, pn, part.to_bytes(strict))
                if part.rels:
                    write(z, part.rels_name, part.rels.to_bytes(strict))


def normalize_namespaces(root: OoxElement) -> None:
    """Hoist namespace declarations to the root and maintain ``mc:Ignorable``.

    Elements from post-2007 namespaces (w14, w15, ...) in ordinary content
    must be listed in ``mc:Ignorable`` or older consumers reject the file.
    Inside ``*:extLst`` and ``mc:AlternateContent`` they need not be.  Prefixes already
    named there are kept declared even if unused, because removing their
    declaration would make the Ignorable value dangle (a corruption Word
    reports as "unreadable content").
    """
    used: dict[str, str] = {}  # every namespace in the tree -> declare on root
    content: set[str] = set()  # used outside extLst / mc:AlternateContent

    def walk(el: etree._Element, in_ext: bool) -> None:
        tag = el.tag
        if not isinstance(tag, str):
            return
        local = tag.rpartition("}")[2]
        in_ext = in_ext or local == "extLst" or tag == f"{{{MC_NS}}}AlternateContent"
        for name in (tag, *el.attrib):
            if name[:1] == "{":
                p = grammar.PREFIX.get(name[1:].partition("}")[0])
                if p and p != "xml":
                    used[p] = grammar.URI[p]
                    if not in_ext:
                        content.add(p)
        for child in el:
            walk(child, in_ext)

    walk(root, False)
    ign_attr = f"{{{MC_NS}}}Ignorable"
    ignorable = (root.get(ign_attr) or "").split()
    # Extension lists and AlternateContent have their own compatibility
    # rules; only namespaces in ordinary content need mc:Ignorable.
    new_ign = [p for p in sorted(content) if grammar.VERSION.get(p) and p not in ignorable]
    keep = set(ignorable)
    if new_ign:
        ignorable += new_ign
        used["mc"] = MC_NS
        keep.add("mc")
        for p in ignorable:
            if p in grammar.URI:
                used.setdefault(p, grammar.URI[p])
    # Declare first, then set mc:Ignorable, so lxml picks the "mc" prefix
    # instead of inventing "ns0".
    etree.cleanup_namespaces(root, top_nsmap=used, keep_ns_prefixes=sorted(keep) or None)
    if new_ign:
        root.set(ign_attr, " ".join(ignorable))


def _peek(file, n: int) -> bytes:
    if isinstance(file, (str, os.PathLike)):
        with open(file, "rb") as f:  # a missing file stays FileNotFoundError
            return f.read(n)
    pos = file.tell()
    data = file.read(n)
    file.seek(pos)
    return data
