"""Document protection: Word's Review > Restrict Editing and Excel's
Review > Protect Workbook.

    doc.protect("readOnly", password="secret")    # or comments, trackedChanges, forms
    doc.protection()       # Protection(edit="readOnly", enforced=True, password=True, ...)
    doc.check_protection_password("secret")        # True
    doc.unprotect()

    book.protect_workbook(password="secret")       # lock the structure (sheets)
    book.workbook_protection()                     # WorkbookProtection(structure=True, ...)

Word's password hash (ECMA-376 Part 1 §17.15.1.29, Part 4 §14.8.1): the
password is first reduced to Word 97's 32-bit key (a high-order word from an
encryption matrix, a low-order word like Excel's 16-bit hash), its four bytes
are reversed and written as an uppercase hex string, and that string
(UTF-16LE) is hashed with the salt prepended, then again ``spin_count``
times with a 4-byte little-endian counter appended.  Checked against the
standard's worked example ("Example" -> 0x64CEED7E) and against a hash
Word 2019 wrote.  The hash goes into the attributes Word 2007 understands
(``w:cryptAlgorithmSid`` …); SHA-1 by default, which every Word version can
check.  Protection is an editing restriction, not encryption: the text stays
readable in the file.
"""

from __future__ import annotations

import base64
import hashlib
import os
from typing import TYPE_CHECKING

from .grammar import URI

if TYPE_CHECKING:
    from .document import WordprocessingDocument
    from .spreadsheet import SpreadsheetDocument

W = URI["w"]
EDIT_MODES = ("readOnly", "comments", "trackedChanges", "forms")
# CryptoAPI algorithm ids (ALG_SID) and providers, as Word writes them
_SID = {"SHA-1": 4, "SHA-256": 12, "SHA-384": 13, "SHA-512": 14}
_ALGORITHM_OF_SID = {v: k for k, v in _SID.items()}
_HASHLIB = {"SHA-1": "sha1", "SHA-256": "sha256", "SHA-384": "sha384", "SHA-512": "sha512"}

_INITIAL_CODE = (0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872,
                 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3)
_ENCRYPTION_MATRIX = (
    (0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09),
    (0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF),
    (0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0),
    (0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40),
    (0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5),
    (0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A),
    (0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9),
    (0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0),
    (0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC),
    (0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10),
    (0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168),
    (0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C),
    (0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD),
    (0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC),
    (0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4),
)


def word_legacy_key(password: str) -> int:
    """Word 97's 32-bit password key (high-order word << 16 | low-order word)."""
    chars = bytes((ord(c) & 0xFF) or (ord(c) >> 8 & 0xFF) for c in password[:15])
    n = len(chars)
    if n == 0:
        return 0
    high = _INITIAL_CODE[n - 1]
    for i, ch in enumerate(chars):
        row = _ENCRYPTION_MATRIX[15 - n + i]
        for bit in range(7):
            if ch & (1 << bit):
                high ^= row[bit]
    low = 0
    for ch in reversed(chars):
        low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ ch
    low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ n ^ 0xCE4B
    return (high << 16) | low


def word_password_hash(password: str, salt: bytes, algorithm: str = "SHA-1",
                       spin_count: int = 100000) -> str:
    """The base64 hash Word stores for ``password`` (see the module docstring)."""
    if algorithm not in _HASHLIB:
        raise ValueError(f"algorithm must be one of {sorted(_HASHLIB)}")
    key = word_legacy_key(password).to_bytes(4, "big")[::-1].hex().upper()
    name = _HASHLIB[algorithm]
    h = hashlib.new(name, salt + key.encode("utf-16-le")).digest()
    for i in range(spin_count):
        h = hashlib.new(name, h + i.to_bytes(4, "little")).digest()
    return base64.b64encode(h).decode("ascii")


# -- Word ---------------------------------------------------------------------------
def _w(el, local: str) -> str | None:
    return el.get(f"{{{W}}}{local}")


def protect(doc: WordprocessingDocument, edit: str = "readOnly", *,
            password: str | None = None, algorithm: str = "SHA-1",
            spin_count: int = 100000, salt: bytes | None = None,
            formatting: bool | None = None) -> None:
    """Restrict editing to ``edit`` (readOnly, comments, trackedChanges,
    forms), optionally with a password (hashed; ``salt`` for reproducible
    files).  ``formatting=True`` also limits formatting to allowed styles."""
    from .ns import w

    if edit not in EDIT_MODES:
        raise ValueError(f"edit must be one of {EDIT_MODES}")
    attrs: dict = {"edit": edit, "enforcement": True}
    if formatting is not None:
        attrs["formatting"] = formatting
    if password is not None:
        if not password:
            raise ValueError("an empty password: pass password=None for no password")
        if algorithm not in _SID:
            raise ValueError(f"algorithm must be one of {sorted(_SID)}")
        salt = os.urandom(16) if salt is None else salt
        attrs.update(cryptProviderType="rsaFull" if algorithm == "SHA-1" else "rsaAES",
                     cryptAlgorithmClass="hash", cryptAlgorithmType="typeAny",
                     cryptAlgorithmSid=_SID[algorithm], cryptSpinCount=spin_count,
                     hash=word_password_hash(password, salt, algorithm, spin_count),
                     salt=base64.b64encode(salt).decode("ascii"))
    doc.ensure_settings().add(w.DocumentProtection(**attrs), replace=True)


def protection(doc: WordprocessingDocument):
    """A ``Protection`` record (edit, enforced, password, algorithm,
    formatting), or None."""
    from .records import Protection

    found = doc.related_parts(doc.main_part, "DocumentSettingsPart")
    if not found:
        return None
    el = found[0].element.find(f"{{{W}}}documentProtection")
    if el is None:
        return None
    sid = _w(el, "cryptAlgorithmSid")
    algorithm = _w(el, "algorithmName") or (
        _ALGORITHM_OF_SID.get(int(sid)) if sid and sid.isdigit() else None)
    return Protection(edit=_w(el, "edit") or "none",
                      enforced=(_w(el, "enforcement") or "0") in ("1", "true", "on"),
                      password=bool(_w(el, "hash") or _w(el, "hashValue")),
                      algorithm=algorithm,
                      formatting=(_w(el, "formatting") or "0") in ("1", "true", "on"))


def check_password(doc: WordprocessingDocument, password: str) -> bool:
    """Whether ``password`` lifts the protection (True without a password).
    Reads both Word 2007's attributes and Word 2010's ``algorithmName`` set."""
    found = doc.related_parts(doc.main_part, "DocumentSettingsPart")
    el = found[0].element.find(f"{{{W}}}documentProtection") if found else None
    if el is None:
        return True
    stored = _w(el, "hash") or _w(el, "hashValue")
    if not stored:
        return True
    salt = base64.b64decode(_w(el, "salt") or _w(el, "saltValue") or "")
    info = protection(doc)
    algorithm = (info.algorithm if info is not None else None) or "SHA-1"
    spin = int(_w(el, "cryptSpinCount") or _w(el, "spinCount") or "0")
    if algorithm not in _HASHLIB:
        raise ValueError(f"can't check a {algorithm} hash")
    return word_password_hash(password, salt, algorithm, spin) == stored


def unprotect(doc: WordprocessingDocument) -> None:
    found = doc.related_parts(doc.main_part, "DocumentSettingsPart")
    el = found[0].element.find(f"{{{W}}}documentProtection") if found else None
    if el is not None:
        el.getparent().remove(el)


# -- Excel --------------------------------------------------------------------------
def protect_workbook(book: SpreadsheetDocument, *, password: str | None = None,
                     structure: bool = True, windows: bool = False) -> None:
    """Lock the workbook's structure (adding, deleting, moving, renaming
    sheets) and/or windows, optionally with Excel's legacy password hash."""
    from .ns import x
    from .xlprint import excel_password_hash

    attrs: dict = {}
    if structure:
        attrs["lockStructure"] = True
    if windows:
        attrs["lockWindows"] = True
    if password:
        attrs["workbookPassword"] = excel_password_hash(password)
    book.workbook.add(x.WorkbookProtection(**attrs), replace=True)


def workbook_protection(book: SpreadsheetDocument):
    """A ``WorkbookProtection`` record (structure, windows, password), or None."""
    from .records import WorkbookProtection

    el = book.workbook.first("x:workbookProtection")
    if el is None:
        return None

    def on(attr: str) -> bool:
        return (el.get(attr) or "0") in ("1", "true")

    return WorkbookProtection(structure=on("lockStructure"), windows=on("lockWindows"),
                              password=bool(el.get("workbookPassword")
                                            or el.get("workbookHashValue")))


def agile_password_hash(password: str, salt: bytes, algorithm: str = "SHA-512",
                        spin_count: int = 100000) -> str:
    """Office 2010's password hash for spreadsheets (``workbookHashValue``,
    ``hashValue``): the UTF-16LE password hashed with the salt prepended, then
    ``spin_count`` times with a 4-byte little-endian counter appended."""
    if algorithm not in _HASHLIB:
        raise ValueError(f"algorithm must be one of {sorted(_HASHLIB)}")
    name = _HASHLIB[algorithm]
    h = hashlib.new(name, salt + password.encode("utf-16-le")).digest()
    for i in range(spin_count):
        h = hashlib.new(name, h + i.to_bytes(4, "little")).digest()
    return base64.b64encode(h).decode("ascii")


def check_workbook_password(book: SpreadsheetDocument, password: str) -> bool:
    """Whether ``password`` unlocks the workbook (True without a password);
    reads Excel 2007's 16-bit hash and Excel 2010's salted one."""
    from .xlprint import excel_password_hash

    el = book.workbook.first("x:workbookProtection")
    if el is None:
        return True
    if el.get("workbookHashValue"):
        return agile_password_hash(
            password, base64.b64decode(el.get("workbookSaltValue") or ""),
            el.get("workbookAlgorithmName") or "SHA-512",
            int(el.get("workbookSpinCount") or "0")) == el.get("workbookHashValue")
    stored = el.get("workbookPassword")
    if not stored:
        return True
    return excel_password_hash(password) == stored.upper()


def unprotect_workbook(book: SpreadsheetDocument) -> None:
    el = book.workbook.first("x:workbookProtection")
    if el is not None:
        el.getparent().remove(el)
