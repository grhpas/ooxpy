"""Tracked changes: suggest edits, list them, accept or reject them.

    doc.track_changes = True                              # Word keeps tracking
    doc.suggest_insert(p, " (draft)", author="Ada")
    doc.suggest_delete(p, "very ", author="Bob")
    doc.suggest_replace(p, "2025", "2026", author="Ada")
    doc.revisions()       # [Revision(kind="insert", author="Ada", text=" (draft)", ...), ...]
    doc.accept_all()      # or doc.reject_all()

Insertions are ``w:ins`` around runs; deletions are ``w:del`` around runs whose
text moved from ``w:t`` to ``w:delText``.  Accepting and rejecting also
handle what Word writes and oox doesn't: moves, inserted or deleted
paragraph marks and table rows, and formatting changes (``w:rPrChange``,
``w:pPrChange`` ...).  They work on the body, headers, footers, notes and
comments.
"""

from __future__ import annotations

import copy
import datetime as _dt
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import w

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_CHANGE_TAGS = ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange",
                "tcPrChange", "tblPrExChange", "tblGridChange", "numberingChange")
# children of a properties element that a *Change record doesn't cover
_KEEP_ON_RESTORE = {"rPr", "sectPr", "headerReference", "footerReference", "cellIns",
                    "cellDel", "cellMerge", "ins", "del"}


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _local(el: OoxElement) -> str:
    return el.tag.rpartition("}")[2]


def _stories(doc: WordprocessingDocument):
    yield doc.document
    for part_type in ("HeaderPart", "FooterPart", "FootnotesPart", "EndnotesPart",
                      "WordprocessingCommentsPart"):
        for part in doc.related_parts(doc.main_part, part_type):
            yield part.element


def _next_id(doc: WordprocessingDocument) -> int:
    ids = [0]
    for root in _stories(doc):
        for el in root.iter():
            v = el.get(_q("id"))
            if v is not None and v.lstrip("-").isdigit():
                ids.append(int(v))
    return max(ids) + 1


def _attrs(doc: WordprocessingDocument, author: str, date: _dt.datetime | None) -> dict[str, Any]:
    attrs: dict[str, Any] = {"id": _next_id(doc), "author": author}
    if date is not None:
        attrs["date"] = date.replace(microsecond=0, tzinfo=None).isoformat() + "Z"
    return attrs


# -- run splitting -------------------------------------------------------------------------
def _segments(paragraph: OoxElement) -> list[tuple[OoxElement, OoxElement, int]]:
    """(run, w:t, start offset) for text runs that are direct children of the
    paragraph (or of a hyperlink in it); tracked runs are skipped."""
    out, pos = [], 0
    for r in paragraph.iter(_q("r")):
        parent = r.getparent()
        if parent is None or _local(parent) not in ("p", "hyperlink"):
            continue
        for t in r.iterchildren(_q("t")):
            out.append((r, t, pos))
            pos += len(t.text or "")
    return out


def _split_run(r: OoxElement, offset: int) -> OoxElement:
    """Split a single-text run at ``offset``; returns the second half."""
    texts = list(r.iterchildren(_q("t")))
    others = [c for c in r if _local(c) not in ("rPr", "t")]
    if len(texts) != 1 or others:
        raise ValueError("can only split runs that hold just text (no tabs, breaks or fields)")
    t = texts[0]
    second = copy.deepcopy(r)
    t2 = second.first("w:t")
    assert t2 is not None
    text = t.text or ""
    _set_t(t, text[:offset])
    _set_t(t2, text[offset:])
    r.addnext(second)
    return second


def _set_t(t: OoxElement, text: str) -> None:
    t.text = text
    if text != text.strip():
        t.set(XML_SPACE, "preserve")


def _isolate(paragraph: OoxElement, old: str, occurrence: int = 1) -> list[OoxElement]:
    """Runs that hold exactly the ``occurrence``-th match of ``old``."""
    segs = _segments(paragraph)
    joined = "".join(t.text or "" for _, t, _ in segs)
    idx = -1
    for _ in range(occurrence):
        idx = joined.find(old, idx + 1)
        if idx == -1:
            raise ValueError(f"{old!r} (occurrence {occurrence}) not found in the paragraph")
    end = idx + len(old)
    # split at the end first, then at the start, so offsets stay valid
    for cut in (end, idx):
        for r, t, start in _segments(paragraph):
            length = len(t.text or "")
            if start < cut < start + length:
                _split_run(r, cut - start)
                break
    return [r for r, t, start in _segments(paragraph)
            if idx <= start < end and (t.text or "")]


# -- suggesting edits ------------------------------------------------------------------------
def suggest_insert(doc: WordprocessingDocument, paragraph: OoxElement, text: str, *,
                   author: str, date: _dt.datetime | None = None,
                   after: str | None = None) -> OoxElement:
    """Tracked insertion of ``text`` at the end of the paragraph, or right
    after the first occurrence of ``after``."""
    from . import teletype

    r = w.R()
    teletype.addTextToElement(r, text)
    ins = w.Ins(r, **_attrs(doc, author, date))
    if after is None:
        paragraph.add(ins)
    else:
        runs = _isolate(paragraph, after)
        runs[-1].addnext(ins)
    return ins


def _wrap_deleted(doc: WordprocessingDocument, runs: list[OoxElement], author: str,
                  date: _dt.datetime | None) -> OoxElement:
    dele = w.Del(**_attrs(doc, author, date))
    runs[0].addprevious(dele)
    for r in runs:
        for t in r.iterchildren(_q("t")):
            t.tag = _q("delText")
        for t in r.iterchildren(_q("instrText")):
            t.tag = _q("delInstrText")
        dele.append(r)
    return dele


def suggest_delete(doc: WordprocessingDocument, paragraph: OoxElement, text: str, *,
                   author: str, date: _dt.datetime | None = None,
                   occurrence: int = 1) -> OoxElement:
    """Tracked deletion of the ``occurrence``-th ``text`` in the paragraph
    (across runs)."""
    return _wrap_deleted(doc, _isolate(paragraph, text, occurrence), author, date)


def suggest_replace(doc: WordprocessingDocument, paragraph: OoxElement, old: str, new: str, *,
                    author: str, date: _dt.datetime | None = None,
                    occurrence: int = 1) -> tuple[OoxElement, OoxElement]:
    """Tracked deletion of ``old`` followed by insertion of ``new``, with the
    formatting of the deleted text."""
    runs = _isolate(paragraph, old, occurrence)
    rpr = runs[0].first("w:rPr")
    dele = _wrap_deleted(doc, runs, author, date)
    r = w.R(copy.deepcopy(rpr)) if rpr is not None else w.R()
    from . import teletype

    teletype.addTextToElement(r, new)
    ins = w.Ins(r, **_attrs(doc, author, date))
    dele.addnext(ins)
    return dele, ins


# -- reading -----------------------------------------------------------------------------------
def revisions(doc: WordprocessingDocument) -> list:
    """A ``Revision`` (kind, author, date, text) for each tracked change, in
    document order (body first, then other stories); kind is insert, delete,
    move-from, move-to, paragraph-insert, paragraph-delete, row-insert,
    row-delete or format."""
    from .records import Revision

    out = []
    names = {"ins": "insert", "del": "delete", "moveFrom": "move-from", "moveTo": "move-to"}
    from .walk import iter_story

    for root in _stories(doc):
        for el in iter_story(root, _q("ins"), _q("del"), _q("moveFrom"), _q("moveTo"),
                             *(_q(t) for t in _CHANGE_TAGS)):
            local, parent = _local(el), el.getparent()
            kind = names.get(local, "format")
            if local in ("ins", "del") and parent is not None:
                if _local(parent) == "rPr":
                    kind = "paragraph-" + ("insert" if local == "ins" else "delete")
                elif _local(parent) == "trPr":
                    kind = "row-" + ("insert" if local == "ins" else "delete")
            text = "".join(t.text or "" for t in el.iter(_q("t"), _q("delText"))) \
                if kind in ("insert", "delete", "move-from", "move-to") else ""
            out.append(Revision(kind, author=el.get(_q("author")), date=el.get(_q("date")),
                                text=text))
    return out


# -- accepting and rejecting ---------------------------------------------------------------------
def _unwrap(el: OoxElement) -> None:
    parent = el.getparent()
    assert parent is not None
    for child in list(el):
        el.addprevious(child)
    parent.remove(el)


def _undelete_text(el: OoxElement) -> None:
    for t in el.iter(_q("delText")):
        t.tag = _q("t")
    for t in el.iter(_q("delInstrText")):
        t.tag = _q("instrText")


def _merge_with_next(p: OoxElement) -> None:
    """Remove a paragraph mark: the next paragraph's content joins this one."""
    nxt = p.getnext()
    while nxt is not None and _local(nxt) != "p":
        if _local(nxt) in ("tbl", "sectPr"):
            return
        nxt = nxt.getnext()
    if nxt is None:
        return
    for child in list(nxt):
        if _local(child) != "pPr":
            p.append(child)
    nxt.getparent().remove(nxt)


def _paragraph_of(mark: OoxElement) -> OoxElement | None:
    el = mark.getparent()
    while el is not None and _local(el) != "p":
        el = el.getparent()
    return el


def _restore(change: OoxElement) -> None:
    """Reject a formatting change: put the recorded old properties back."""
    props = change.getparent()
    assert props is not None
    old = change[0] if len(change) else None
    for child in list(props):
        if child is not change and _local(child) not in _KEEP_ON_RESTORE:
            props.remove(child)
    props.remove(change)
    if old is not None:
        for child in old:
            props.add(copy.deepcopy(child))


def _resolve(root: OoxElement, accept: bool) -> int:
    count = 0
    # paragraph marks and rows first: they reference whole paragraphs/rows
    for mark in list(root.iter(_q("ins"), _q("del"))):
        parent = mark.getparent()
        if parent is None or _local(parent) not in ("rPr", "trPr"):
            continue
        count += 1
        drop = (_local(mark) == "del") == accept  # accept+del or reject+ins removes the thing
        if _local(parent) == "trPr":
            tr = parent.getparent()
            parent.remove(mark)
            if drop and tr is not None and tr.getparent() is not None:
                tr.getparent().remove(tr)
        else:
            p = _paragraph_of(mark)
            parent.remove(mark)
            if drop and p is not None:
                _merge_with_next(p)
    for el in list(root.iter(_q("ins"), _q("del"), _q("moveFrom"), _q("moveTo"))):
        if el.getparent() is None:
            continue
        count += 1
        keep = (_local(el) in ("ins", "moveTo")) == accept
        if keep:
            _undelete_text(el)
            _unwrap(el)
        else:
            el.getparent().remove(el)
    for rng in list(root.iter(*(_q(t) for t in ("moveFromRangeStart", "moveFromRangeEnd",
                                                 "moveToRangeStart", "moveToRangeEnd")))):
        rng.getparent().remove(rng)
    for change in list(root.iter(*(_q(t) for t in _CHANGE_TAGS))):
        if change.getparent() is None:
            continue
        count += 1
        if accept:
            change.getparent().remove(change)
        else:
            _restore(change)
    for marker in list(root.iter(_q("cellIns"), _q("cellDel"), _q("cellMerge"))):
        marker.getparent().remove(marker)
    return count


def accept_all(doc: WordprocessingDocument) -> int:
    """Accept every tracked change; returns how many there were."""
    return sum(_resolve(root, accept=True) for root in _stories(doc))


def reject_all(doc: WordprocessingDocument) -> int:
    """Reject every tracked change; returns how many there were."""
    return sum(_resolve(root, accept=False) for root in _stories(doc))
