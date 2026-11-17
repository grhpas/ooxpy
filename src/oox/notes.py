"""Hyperlinks, bookmarks, footnotes, endnotes and comments.

    p = doc.add_paragraph("See ")
    doc.add_hyperlink(p, "the spec", target="https://ecma-international.org/")
    doc.add_bookmark("results", doc.add_heading("Results", 1))
    doc.add_hyperlink(p, "the results", anchor="results")      # internal link
    doc.add_footnote(p, "A footnote.")
    doc.add_endnote(p, "An endnote.")
    doc.add_comment(p, "Please check.", author="Ada")

and back:  doc.hyperlinks(), doc.bookmarks(), doc.footnotes(), doc.endnotes(),
doc.comments().

Footnote and endnote parts are created with the two separator notes Word
expects (ids -1 and 0) and registered in settings (``w:footnotePr``), as Word
does; real notes number from 1.
"""

from __future__ import annotations

import datetime as _dt
from collections.abc import Callable
from typing import TYPE_CHECKING, Any, NamedTuple

from .element import OoxElement
from .grammar import URI
from .ns import w
from .records import Comment, Hyperlink, Note

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
R = URI["r"]
# hyperlinks are external relationships, not parts, so they aren't in the part table
RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def _text_run(text: str, rstyle: str | None = None) -> OoxElement:
    r = w.R(w.RPr(w.RStyle(val=rstyle))) if rstyle else w.R()
    from . import teletype

    teletype.addTextToElement(r, text)
    return r


def _source_part(doc: WordprocessingDocument, el: OoxElement):
    """The part a paragraph lives in (headers/footers relate their own links)."""
    from .opc import PackageError

    try:
        return doc.part_of(el)
    except PackageError:  # detached paragraph: it will go in the body
        return doc.main_part


# -- hyperlinks and bookmarks ----------------------------------------------------------
def add_hyperlink(doc: WordprocessingDocument, paragraph: OoxElement, text: str, *,
                  target: str | None = None, anchor: str | None = None,
                  tooltip: str | None = None) -> OoxElement:
    if (target is None) == (anchor is None):
        raise ValueError("give exactly one of target= (a URL) or anchor= (a bookmark name)")
    doc.ensure_style("Hyperlink")
    attrs: dict[str, Any] = {}
    if target is not None:
        part = _source_part(doc, paragraph)
        attrs["r_id"] = part.rels.add(RT_HYPERLINK, target, external=True)
    else:
        attrs["anchor"] = anchor
    if tooltip:
        attrs["tooltip"] = tooltip
    link = w.Hyperlink(_text_run(text, "Hyperlink"), history=True, **attrs)
    paragraph.add(link)
    return link


def _next_bookmark_id(doc: WordprocessingDocument) -> int:
    ids = [int(b.get(f"{{{W}}}id", "0")) for b in doc.document.iter(f"{{{W}}}bookmarkStart")]
    return max(ids, default=-1) + 1


def add_bookmark(doc: WordprocessingDocument, name: str, paragraph: OoxElement) -> OoxElement:
    """Bookmark the whole content of ``paragraph`` (target for anchor links)."""
    if not name or len(name) > 40 or not (name[0].isalpha() or name[0] == "_") or " " in name:
        raise ValueError("bookmark names: up to 40 characters, start with a letter, no spaces")
    if name in bookmarks(doc):
        raise ValueError(f"duplicate bookmark {name!r}")
    bid = _next_bookmark_id(doc)
    start, end = w.BookmarkStart(id=bid, name=name), w.BookmarkEnd(id=bid)
    ppr = paragraph.first("w:pPr")
    if ppr is not None:
        ppr.addnext(start)
    else:
        paragraph.insert(0, start)
    paragraph.append(end)
    return start


def hyperlinks(doc: WordprocessingDocument) -> list[Hyperlink]:
    """A ``Hyperlink`` for each link in the body: ``target`` the URL, or
    ``anchor`` the bookmark (kind "url" or "bookmark")."""
    from . import teletype
    from .walk import iter_story

    out = []
    for link in iter_story(doc.body, f"{{{W}}}hyperlink"):
        rid = link.get(f"{{{R}}}id")
        tip = link.get(f"{{{W}}}tooltip")
        text = teletype.extractText(link)
        if rid:
            rel = doc.main_part.rels.get(rid)
            out.append(Hyperlink(text, target=rel.target if rel is not None else None,
                                 anchor=link.get(f"{{{W}}}anchor"), tooltip=tip))
        else:
            out.append(Hyperlink(text, anchor=link.get(f"{{{W}}}anchor") or "", tooltip=tip,
                                 kind="bookmark"))
    return out


def bookmarks(doc: WordprocessingDocument) -> list[str]:
    """Bookmark names, without Word's hidden ones (``_GoBack``, ``_Toc...``)."""
    from .walk import iter_story

    return [b.get(f"{{{W}}}name") or "" for b in iter_story(doc.document, f"{{{W}}}bookmarkStart")
            if not (b.get(f"{{{W}}}name") or "").startswith("_")]


# -- footnotes and endnotes --------------------------------------------------------------
class _NoteKind(NamedTuple):
    part: str
    root: Callable[..., OoxElement]
    note: Callable[..., OoxElement]
    ref: Callable[..., OoxElement]
    mark: Callable[..., OoxElement]
    props: Callable[..., OoxElement]
    text_style: str
    ref_style: str


_NOTE = {
    "footnote": _NoteKind("FootnotesPart", w.Footnotes, w.Footnote, w.FootnoteReference,
                          w.FootnoteRef, w.FootnotePr, "FootnoteText", "FootnoteReference"),
    "endnote": _NoteKind("EndnotesPart", w.Endnotes, w.Endnote, w.EndnoteReference,
                         w.EndnoteRef, w.EndnotePr, "EndnoteText", "EndnoteReference"),
}


def _notes_root(doc: WordprocessingDocument, kind: str, create: bool) -> OoxElement | None:
    spec = _NOTE[kind]
    found = doc.related_parts(doc.main_part, spec.part)
    if found:
        return found[0].element
    if not create:
        return None
    sep = w.P(w.PPr(w.Spacing(after=0, line=240, lineRule="auto")), w.R(w.Separator()))
    cont = w.P(w.PPr(w.Spacing(after=0, line=240, lineRule="auto")), w.R(w.ContinuationSeparator()))
    root = spec.root(spec.note(sep, type="separator", id=-1),
                        spec.note(cont, type="continuationSeparator", id=0))
    doc.add_part(spec.part, element=root)
    # Word lists the separator notes in settings
    doc.ensure_settings().add(spec.props(spec.note(id=-1), spec.note(id=0)), replace=True)
    return root


def _add_note(doc: WordprocessingDocument, kind: str, paragraph: OoxElement, text: str) -> OoxElement:
    spec = _NOTE[kind]
    doc.ensure_style(spec.text_style)
    doc.ensure_style(spec.ref_style)
    root = _notes_root(doc, kind, create=True)
    assert root is not None
    ids = [int(n.get(f"{{{W}}}id", "0")) for n in root.iterchildren(f"{{{W}}}{kind}")]
    nid = max([0, *ids]) + 1
    body = w.P(w.PPr(w.PStyle(val=spec.text_style)),
               w.R(w.RPr(w.RStyle(val=spec.ref_style)), spec.mark()))
    from . import teletype

    teletype.addTextToElement(body, " " + text)
    note = spec.note(body, id=nid)
    root.add(note)
    paragraph.add(w.R(w.RPr(w.RStyle(val=spec.ref_style)), spec.ref(id=nid)))
    return note


def add_footnote(doc: WordprocessingDocument, paragraph: OoxElement, text: str) -> OoxElement:
    return _add_note(doc, "footnote", paragraph, text)


def add_endnote(doc: WordprocessingDocument, paragraph: OoxElement, text: str) -> OoxElement:
    return _add_note(doc, "endnote", paragraph, text)


def notes(doc: WordprocessingDocument, kind: str) -> list[Note]:
    """A ``Note`` (id, text) for each real note (separators excluded)."""
    from . import teletype

    root = _notes_root(doc, kind, create=False)
    if root is None:
        return []
    return [Note(int(n.get(f"{{{W}}}id", "0")), teletype.extractText(n).strip())
            for n in root.iterchildren(f"{{{W}}}{kind}")
            if n.get(f"{{{W}}}type") in (None, "normal")]


# -- comments ------------------------------------------------------------------------------
def _comments_root(doc: WordprocessingDocument, create: bool) -> OoxElement | None:
    found = doc.related_parts(doc.main_part, "WordprocessingCommentsPart")
    if found:
        return found[0].element
    if not create:
        return None
    root = w.Comments()
    doc.add_part("WordprocessingCommentsPart", element=root)
    return root


def add_comment(doc: WordprocessingDocument, paragraph: OoxElement, text: str, *,
                author: str = "oox", initials: str | None = None,
                date: _dt.datetime | None = None) -> OoxElement:
    """Comment on the whole content of ``paragraph``."""
    from . import teletype

    doc.ensure_style("CommentText")
    doc.ensure_style("CommentReference")
    root = _comments_root(doc, create=True)
    assert root is not None
    ids = [int(c.get(f"{{{W}}}id", "0")) for c in root.iterchildren(f"{{{W}}}comment")]
    ids += [int(c.get(f"{{{W}}}id", "0")) for c in doc.document.iter(f"{{{W}}}commentRangeStart")]
    cid = max(ids, default=-1) + 1
    body = w.P(w.PPr(w.PStyle(val="CommentText")),
               w.R(w.RPr(w.RStyle(val="CommentReference")), w.AnnotationRef()))
    teletype.addTextToElement(body, text)
    attrs: dict[str, Any] = {"id": cid, "author": author,
                                "initials": initials or "".join(p[0] for p in author.split())[:9]}
    if date is not None:
        attrs["date"] = date.replace(microsecond=0, tzinfo=None).isoformat() + "Z"
    comment = w.Comment(body, **attrs)
    root.add(comment)
    start = w.CommentRangeStart(id=cid)
    ppr = paragraph.first("w:pPr")
    if ppr is not None:
        ppr.addnext(start)
    else:
        paragraph.insert(0, start)
    paragraph.append(w.CommentRangeEnd(id=cid))
    paragraph.append(w.R(w.RPr(w.RStyle(val="CommentReference")), w.CommentReference(id=cid)))
    return comment


def comments(doc: WordprocessingDocument) -> list[Comment]:
    """A ``Comment`` for each comment; ``anchor_text`` is the commented text
    in the body."""
    from . import teletype

    root = _comments_root(doc, create=False)
    if root is None:
        return []
    from .walk import iter_story

    anchors: dict[str, str] = {}
    for p in iter_story(doc.body, f"{{{W}}}p"):
        for start in p.iter(f"{{{W}}}commentRangeStart"):
            anchors[start.get(f"{{{W}}}id") or ""] = teletype.extractText(p)
    return [Comment(teletype.extractText(c), author=c.get(f"{{{W}}}author"),
                    date=c.get(f"{{{W}}}date"), initials=c.get(f"{{{W}}}initials"),
                    anchor_text=anchors.get(c.get(f"{{{W}}}id") or "", ""),
                    id=int(c.get(f"{{{W}}}id", "0")))
            for c in root.iterchildren(f"{{{W}}}comment")]
