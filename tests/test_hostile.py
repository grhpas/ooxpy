"""Untrusted input: opening must either work or raise PackageError -- never
crash with an unrelated exception, hang, or expand without bound."""

import io
import random
import zipfile

import pytest

from oox import (
    Limits,
    OoxError,
    Package,
    PackageError,
    PresentationDocument,
    SpreadsheetDocument,
    WordprocessingDocument,
    open_document,
)
from oox.ns import w


def _zip(entries: dict[str, bytes], compression=zipfile.ZIP_DEFLATED) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    return buf.getvalue()


def _docx_entries() -> dict[str, bytes]:
    doc = WordprocessingDocument.new()
    doc.body.add(w.P(w.R(w.T("hello"))))
    z = zipfile.ZipFile(io.BytesIO(doc.package.to_bytes()))
    return {n: z.read(n) for n in z.namelist()}


def test_legacy_and_encrypted_files_get_a_clear_message():
    ole = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 504
    with pytest.raises(PackageError, match=r"password-protected .* or a legacy"):
        Package.open(io.BytesIO(ole))


def test_not_a_zip():
    with pytest.raises(PackageError, match="not a readable OOXML package"):
        Package.open(io.BytesIO(b"plain text, not a package"))


def test_missing_file_is_still_file_not_found(tmp_path):
    with pytest.raises(FileNotFoundError):
        Package.open(tmp_path / "nope.docx")


def test_zip_bomb_limits():
    bomb = _docx_entries()
    bomb["word/media/huge.bin"] = b"\0" * (64 * 2**20)  # 64 MiB of zeros -> ~64 KiB zipped
    with pytest.raises(PackageError, match="compression ratio too high"):
        Package.open(io.BytesIO(_zip(bomb)))  # caught by the default limits

    big = _docx_entries()
    big["word/media/photo.bin"] = random.Random(1).randbytes(3 * 2**20)  # incompressible
    data = _zip(big)
    assert Package.open(io.BytesIO(data))  # an ordinary big media file is fine
    with pytest.raises(PackageError, match="part size limit"):
        Package.open(io.BytesIO(data), limits=Limits(max_part_size=2**20))
    with pytest.raises(PackageError, match="over the limit"):
        Package.open(io.BytesIO(data), limits=Limits(max_total_size=2**20))
    many = {**_docx_entries(), **{f"x/{i}.bin": b"x" for i in range(50)}}
    with pytest.raises(PackageError, match="too many entries"):
        Package.open(io.BytesIO(_zip(many)), limits=Limits(max_parts=20))


def test_duplicate_entries_are_rejected():
    buf = io.BytesIO()
    entries = _docx_entries()
    with zipfile.ZipFile(buf, "w") as z:
        for name, data in entries.items():
            z.writestr(name, data)
        with pytest.warns(UserWarning):  # zipfile itself warns about duplicates
            z.writestr("word/document.xml", b"<other/>")
    with pytest.raises(PackageError, match="duplicate zip entry"):
        Package.open(io.BytesIO(buf.getvalue()))


def test_dtd_and_entity_expansion_are_refused():
    entries = _docx_entries()
    entries["word/document.xml"] = (
        b'<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 '
        b'"&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body><w:p><w:r><w:t>&lol2;</w:t></w:r></w:p></w:body></w:document>')
    doc = WordprocessingDocument.open(io.BytesIO(_zip(entries)))  # lazy: opens
    with pytest.raises(PackageError, match="DTD declarations are not allowed"):
        doc.document  # noqa: B018


def test_malformed_xml_and_relationships():
    entries = _docx_entries()
    entries["word/document.xml"] = b"<w:document"
    doc = WordprocessingDocument.open(io.BytesIO(_zip(entries)))
    with pytest.raises(PackageError, match="malformed XML"):
        doc.body  # noqa: B018
    assert "cannot parse" in doc.validate()["/word/document.xml"][0]  # reported, not raised
    entries = _docx_entries()
    entries["_rels/.rels"] = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1"/></Relationships>'
    with pytest.raises(PackageError, match="without Id/Type/Target"):
        Package.open(io.BytesIO(_zip(entries)))


def _seeds():
    for build in (lambda: WordprocessingDocument.new(), lambda: SpreadsheetDocument.new(),
                  lambda: PresentationDocument.new()):
        doc = build()
        if isinstance(doc, PresentationDocument):
            doc.add_text_box(doc.add_slide(), "fuzz", x=0, y=0, width=72, height=72)
        yield doc.package.to_bytes()


def _mutate(rng: random.Random, data: bytes) -> bytes:
    b = bytearray(data)
    op = rng.randrange(4)
    if op == 0:  # flip bytes
        for _ in range(rng.randint(1, 20)):
            b[rng.randrange(len(b))] = rng.randrange(256)
    elif op == 1:  # truncate
        del b[rng.randrange(1, len(b)):]
    elif op == 2:  # splice garbage
        pos = rng.randrange(len(b))
        b[pos:pos] = bytes(rng.randrange(256) for _ in range(rng.randint(1, 64)))
    else:  # corrupt one entry's content but keep a valid zip
        z = zipfile.ZipFile(io.BytesIO(data))
        names = z.namelist()
        victim = rng.choice(names)
        entries = {n: z.read(n) for n in names}
        content = bytearray(entries[victim])
        for _ in range(rng.randint(1, 8)):
            if content:
                content[rng.randrange(len(content))] = rng.choice(b"<>&\"'/=x\0")
        entries[victim] = bytes(content)
        return _zip(entries)
    return bytes(b)


def test_fuzz_open_parse_validate_never_crashes():
    rng = random.Random(20261114)  # deterministic
    seeds = list(_seeds())
    outcomes = {"opened": 0, "rejected": 0}
    for _ in range(400):
        data = _mutate(rng, rng.choice(seeds))
        try:
            doc = open_document(io.BytesIO(data))
            doc.validate()  # parses every part; problems are reported, not raised
            doc.package.to_bytes()
            outcomes["opened"] += 1
        except OoxError:
            outcomes["rejected"] += 1
    assert outcomes["opened"] and outcomes["rejected"]  # both paths exercised
