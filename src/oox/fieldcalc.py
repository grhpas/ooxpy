"""Captions, cross-references, document-information fields, and computing
field results without Word.

    cap = doc.add_caption("Sales by region", label="Figure", after=picture_paragraph)
    p = doc.add_paragraph("As ")
    doc.add_cross_reference(p, cap)                  # "Figure 1"
    doc.add_cross_reference(p, cap, what="page")     # page estimate
    doc.add_field(p, "TITLE")                        # the document title
    doc.update_fields()                              # renumber, refresh references
    doc.fields()                                     # [Field(code="SEQ Figure ...", result="1"), ...]

Word stores a field's last result in the file and only recomputes it on F9 or
printing, so the file has to carry correct results.  ``update_fields``
computes what doesn't need page layout:

* ``SEQ`` (caption numbers; ``\\r n`` restart, ``\\c`` repeat, ``\\h`` hidden,
  ``\\* ARABIC|ROMAN|roman|ALPHABETIC|alphabetic``)
* ``REF`` (bookmark text; ``\\p`` gives "above"/"below"), ``PAGEREF`` (an
  *estimate* from explicit page and section breaks, or from the page breaks
  Word recorded at its last layout)
* ``TITLE``, ``SUBJECT``, ``AUTHOR``, ``KEYWORDS``, ``COMMENTS``,
  ``LASTSAVEDBY``, ``DOCPROPERTY`` (core, extended and custom properties),
  ``CREATEDATE``/``SAVEDATE``/``PRINTDATE`` with a ``\\@`` date picture, and
  ``DATE``/``TIME`` when ``now=`` is given
* case switches ``\\* Upper|Lower|Caps|FirstCap``

``PAGE``, ``NUMPAGES``, ``TOC`` and other layout-dependent fields keep their
cached results.
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import w

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
REF_ERROR = "Error! Reference source not found."
BOOKMARK_ERROR = "Error! Bookmark not defined."


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


# -- scanning ----------------------------------------------------------------------------------
@dataclass
class Field:
    """A field: its ``code`` ("SEQ Figure \\* ARABIC"), ``name`` ("SEQ"),
    cached ``result``, ``kind`` ("simple" or "complex") and elements."""

    code: str
    kind: str                      # "simple" or "complex"
    element: OoxElement            # w:fldSimple, or the begin w:fldChar
    separate: OoxElement | None = None
    end: OoxElement | None = None
    result: str = ""

    @property
    def name(self) -> str:
        """The field type in upper case ("SEQ", "REF", "PAGE" ...)."""
        words = self.code.split()
        return words[0].upper() if words else ""

    def args(self) -> list[str]:
        """The field code's arguments and switches, tokenized as Word does (quotes
        respected)."""
        return tokens(self.code)[1:]


def tokens(code: str) -> list[str]:
    """Split a field code Word's way: a switch is a backslash and one
    character (``\\*``, ``\\@``, ``\\h``), double quotes group, and inside quotes
    a backslash escapes the next character (LibreOffice writes ``\\@"d\\ MMMM"``)."""
    out: list[str] = []
    i, n = 0, len(code)
    while i < n:
        ch = code[i]
        if ch.isspace():
            i += 1
        elif ch == '"':
            i += 1
            buf = []
            while i < n and code[i] != '"':
                if code[i] == "\\" and i + 1 < n:
                    i += 1
                buf.append(code[i])
                i += 1
            out.append("".join(buf))
            i += 1
        elif ch == "\\" and i + 1 < n:
            out.append(code[i:i + 2])
            i += 2
        else:
            j = i
            while j < n and not code[j].isspace() and code[j] != '"':
                j += 1
            out.append(code[i:j])
            i = j
    return out


def scan(root: OoxElement, all_copies: bool = False) -> list[Field]:
    """Fields under ``root`` in document order (nested fields included, outer
    first), with their current results.  Text boxes count once, as Word shows
    them (``oox.walk``); ``all_copies=True`` also scans the copies Word keeps
    for older readers."""
    from .walk import iter_story

    out: list[Field] = []
    stack: list[dict[str, Any]] = []
    tags = (_q("fldSimple"), _q("fldChar"), _q("instrText"), _q("t"), _q("tab"), _q("br"), _q("p"))
    for el in (root.iter(*tags) if all_copies else iter_story(root, *tags)):
        local = el.tag.rpartition("}")[2]
        if local == "fldSimple":
            from . import teletype

            out.append(Field((el.get(_q("instr")) or "").strip(), "simple", el,
                             result=teletype.extractText(el)))
            continue
        if local == "fldChar":
            kind = el.get(_q("fldCharType"))
            if kind == "begin":
                f = Field("", "complex", el)
                out.append(f)
                stack.append({"field": f, "instr": [], "state": "instr", "result": []})
            elif kind == "separate" and stack:
                stack[-1]["state"] = "result"
                stack[-1]["field"].separate = el
            elif kind == "end" and stack:
                top = stack.pop()
                f = top["field"]
                f.code = "".join(top["instr"]).strip()
                f.end = el
                f.result = "".join(top["result"])
                if stack and stack[-1]["state"] == "result":
                    stack[-1]["result"].append(f.result)
            continue
        if not stack:
            continue
        top = stack[-1]
        if local == "instrText" and top["state"] == "instr":
            top["instr"].append(el.text or "")
        elif top["state"] == "result":
            if local == "t" and not _inside_simple(el):
                top["result"].append(el.text or "")
            elif local == "tab":
                top["result"].append("\t")
            elif local == "br":
                top["result"].append("\n")
            elif local == "p" and top["result"]:
                top["result"].append("\n")
    return out


def _inside_simple(el: OoxElement) -> bool:
    parent = el.getparent()
    while parent is not None:
        if parent.tag == _q("fldSimple"):
            return True
        parent = parent.getparent()
    return False


# -- writing results -----------------------------------------------------------------------------
def _result_rpr(runs: list[OoxElement]) -> OoxElement | None:
    import copy

    for r in runs:
        if r.tag == _q("r"):
            rpr = r.first("w:rPr")
            return copy.deepcopy(rpr) if rpr is not None else None
    return None


def _text_run(text: str, rpr: OoxElement | None) -> OoxElement:
    from . import teletype

    r = w.R(rpr) if rpr is not None else w.R()
    if text:
        teletype.addTextToElement(r, text)
    return r


def set_result(f: Field, text: str) -> bool:
    """Replace a field's cached result; False if the result spans paragraphs
    (a TOC, say) and was left alone."""
    if f.kind == "simple":
        kids = [k for k in f.element if k.tag in (_q("r"), _q("hyperlink"))]
        rpr = _result_rpr(kids)
        for k in kids:
            f.element.remove(k)
        f.element.append(_text_run(text, rpr))
        f.result = text
        return True
    if f.separate is None or f.end is None:
        return False
    sep_run, end_run = f.separate.getparent(), f.end.getparent()
    if sep_run is None or end_run is None or sep_run.getparent() is not end_run.getparent():
        return False
    between = []
    nxt = sep_run.getnext()
    while nxt is not None and nxt is not end_run:
        between.append(nxt)
        nxt = nxt.getnext()
    if nxt is None:
        return False
    rpr = _result_rpr(between)
    for el in between:
        el.getparent().remove(el)
    end_run.addprevious(_text_run(text, rpr))
    f.result = text
    return True


# -- computing ----------------------------------------------------------------------------------
def _letters(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out, n = out + s, n - v
    return out


def format_number(n: int, fmt: str | None) -> str:
    """Word's numeric picture switches (``\\* ARABIC`` ...)."""
    if fmt == "ROMAN":
        return _roman(n)
    if fmt == "roman":
        return _roman(n).lower()
    if fmt == "ALPHABETIC":
        return _letters(n)
    if fmt == "alphabetic":
        return _letters(n).lower()
    return str(n)


_DATE_TOKENS = re.compile(r"'[^']*'|dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy|HH|H|hh|h|mm|m|ss|s|"
                          r"AM/PM|am/pm")


def format_date(value: _dt.datetime, picture: str) -> str:
    """Word's date-time picture (``\\@ "d MMMM yyyy"``), English names."""
    def tok(m: re.Match[str]) -> str:
        t = m.group(0)
        if t.startswith("'"):
            return t[1:-1]
        hour12 = value.hour % 12 or 12
        return {
            "dddd": value.strftime("%A"), "ddd": value.strftime("%a"),
            "dd": f"{value.day:02d}", "d": str(value.day),
            "MMMM": value.strftime("%B"), "MMM": value.strftime("%b"),
            "MM": f"{value.month:02d}", "M": str(value.month),
            "yyyy": f"{value.year:04d}", "yy": f"{value.year % 100:02d}",
            "HH": f"{value.hour:02d}", "H": str(value.hour),
            "hh": f"{hour12:02d}", "h": str(hour12),
            "mm": f"{value.minute:02d}", "m": str(value.minute),
            "ss": f"{value.second:02d}", "s": str(value.second),
            "AM/PM": "AM" if value.hour < 12 else "PM",
            "am/pm": "am" if value.hour < 12 else "pm",
        }[t]

    return _DATE_TOKENS.sub(tok, picture)


def _switch(args: list[str], name: str) -> str | None:
    """Value of a switch: ``\\* ROMAN`` -> "ROMAN"; ``\\h`` -> ""."""
    for i, a in enumerate(args):
        if a == f"\\{name}":
            return args[i + 1] if i + 1 < len(args) and not args[i + 1].startswith("\\") else ""
    return None


def _formats(args: list[str]) -> list[str]:
    return [args[i + 1] for i, a in enumerate(args) if a == "\\*" and i + 1 < len(args)]


def _apply_case(text: str, args: list[str]) -> str:
    for f in _formats(args):
        low = f.lower()
        if low == "upper":
            text = text.upper()
        elif low == "lower":
            text = text.lower()
        elif low == "caps":
            text = " ".join(word[:1].upper() + word[1:].lower() for word in text.split(" "))
        elif low == "firstcap":
            text = text[:1].upper() + text[1:]
    return text


def bookmark_texts(root: OoxElement) -> dict[str, str]:
    """Text inside each bookmark (by name); paragraphs joined with newlines.
    Deleted text (``w:delText``) and field codes aren't text, so they don't count."""
    from .walk import iter_story

    out: dict[str, str] = {}
    open_: dict[str, tuple[str, list[str]]] = {}
    for el in iter_story(root, _q("bookmarkStart"), _q("bookmarkEnd"), _q("t"), _q("tab"),
                         _q("p")):
        local = el.tag.rpartition("}")[2]
        if local == "bookmarkStart":
            open_[el.get(_q("id")) or ""] = (el.get(_q("name")) or "", [])
        elif local == "bookmarkEnd":
            found = open_.pop(el.get(_q("id")) or "", None)
            if found is not None:
                out[found[0]] = "".join(found[1])
        else:
            for _, parts in open_.values():
                if local == "p":
                    if parts:  # a later paragraph of a multi-paragraph bookmark
                        parts.append("\n")
                else:
                    parts.append(el.text or "" if local == "t" else "\t")
    return out


def _page_estimates(root: OoxElement) -> dict[str, int]:
    """Bookmark name -> estimated page number."""
    from .walk import iter_story

    rendered = root.find(f".//{_q('lastRenderedPageBreak')}") is not None
    page, out = 1, {}
    section_types = [sp.first("w:type") for sp in iter_story(root, _q("sectPr"))]
    sect_index = 0
    for el in iter_story(root, _q("bookmarkStart"), _q("br"), _q("lastRenderedPageBreak"),
                         _q("pageBreakBefore"), _q("sectPr")):
        local = el.tag.rpartition("}")[2]
        if local == "bookmarkStart":
            out.setdefault(el.get(_q("name")) or "", page)
        elif rendered:
            if local == "lastRenderedPageBreak":
                page += 1
        elif local == "br" and el.get(_q("type")) == "page":
            page += 1
        elif local == "pageBreakBefore" and el.get(_q("val")) not in ("0", "false", "off"):
            page += 1
        elif local == "sectPr":
            sect_index += 1
            parent = el.getparent()
            if parent is not None and parent.tag == _q("pPr") and sect_index < len(section_types):
                nxt = section_types[sect_index]
                if nxt is None or nxt.get(_q("val")) != "continuous":
                    page += 1
    return out


def _property(doc: WordprocessingDocument, name: str) -> Any:
    core = {"title": "title", "subject": "subject", "author": "creator",
            "keywords": "keywords", "comments": "description", "category": "category",
            "lastsavedby": "last_modified_by", "last author": "last_modified_by",
            "revision number": "revision"}
    ext = {"company": "company", "manager": "manager", "template": "template",
           "hyperlinkbase": "hyperlink_base", "hyperlink base": "hyperlink_base"}
    key = name.lower()
    if key in core:
        return getattr(doc.properties, core[key])
    if key in ext:
        return getattr(doc.extended_properties, ext[key])
    custom = doc.custom_properties
    for k in custom.keys():
        if k.lower() == key:
            return custom[k]
    return None


def _render_value(value: Any) -> str:
    if isinstance(value, bool):
        return "Y" if value else "N"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return "" if value is None else str(value)


def update_fields(doc: WordprocessingDocument, *, now: _dt.datetime | None = None,
                  only: set[str] | None = None) -> dict[str, int]:
    """Recompute field results that don't need page layout; returns
    ``{field name: number updated}``.  ``only`` limits it to some names."""
    from .walk import twins

    counts: dict[str, int] = {}
    root = doc.document
    # Numbers and references come from the copy of each text box Word shows;
    # every result is written into the copy kept for older readers too
    # (the n-th field of one copy is the n-th of the other).
    twin_of: dict[int, Field] = {}
    keep_alive: list[Field] = []  # keeps the element proxies (and so their ids) stable
    for shown, hidden in twins(root):
        a, b = scan(shown), scan(hidden)
        if len(a) == len(b):
            keep_alive += a
            twin_of.update({id(x.element): y for x, y in zip(a, b, strict=True)})

    def done(f: Field, text: str) -> None:
        if f.result != text and set_result(f, text):
            counts[f.name] = counts.get(f.name, 0) + 1
        twin = twin_of.get(id(f.element))
        if twin is not None and twin.result != text:
            set_result(twin, text)

    def wanted(name: str) -> bool:
        return only is None or name in only

    # 1. SEQ numbering, in document order
    seq: dict[str, int] = {}
    for f in scan(root):
        if f.name != "SEQ" or not wanted("SEQ"):
            continue
        args = f.args()
        if not args:
            continue
        ident = args[0]
        restart = _switch(args, "r")
        if restart:
            seq[ident] = int(restart)
        elif _switch(args, "c") is None:
            seq[ident] = seq.get(ident, 0) + 1
        n = seq.get(ident, 0)
        fmt = next((x for x in _formats(args) if x in ("ARABIC", "ROMAN", "roman",
                                                         "ALPHABETIC", "alphabetic")), None)
        done(f, "" if _switch(args, "h") is not None else format_number(n, fmt))
    # 2. references, now that numbers are current
    texts = bookmark_texts(root)
    pages = _page_estimates(root) if wanted("PAGEREF") else {}
    order = {el.get(_q("name")): i for i, el in enumerate(root.iter(_q("bookmarkStart")))}
    fields = scan(root)
    # positions in one pass: lxml proxies of unreferenced elements aren't
    # stable, so ids are only compared for the field elements we hold
    held = {id(f.element): f for f in fields}
    positions: dict[int, int] = {}
    starts: dict[str | None, int] = {}
    for i, el in enumerate(root.iter()):
        if el.tag == _q("bookmarkStart"):
            starts.setdefault(el.get(_q("name")), i)
        else:
            f = held.get(id(el))
            if f is not None and f.element is el:
                positions[id(el)] = i
    for f in fields:
        name, args = f.name, f.args()
        if not wanted(name):
            continue
        if name == "REF" and args:
            target = args[0]
            if target not in order:
                done(f, REF_ERROR)
            elif _switch(args, "p") is not None:
                done(f, "above" if starts[target] < positions[id(f.element)] else "below")
            else:
                done(f, _apply_case(texts.get(target, ""), args))
        elif name == "PAGEREF" and args:
            target = args[0]
            done(f, str(pages[target]) if target in pages else BOOKMARK_ERROR)
        elif name in ("TITLE", "SUBJECT", "AUTHOR", "KEYWORDS", "COMMENTS", "LASTSAVEDBY"):
            value = _property(doc, name)
            if value is not None:
                done(f, _apply_case(_render_value(value), args))
        elif name == "DOCPROPERTY" and args:
            value = _property(doc, args[0])
            if value is not None:
                done(f, _apply_case(_render_value(value), args))
        elif name in ("CREATEDATE", "SAVEDATE", "PRINTDATE", "DATE", "TIME"):
            picture = _switch(args, "@")
            if not picture:
                continue
            if name in ("DATE", "TIME"):
                value = now
            else:
                value = {"CREATEDATE": doc.properties.created, "SAVEDATE": doc.properties.modified,
                         "PRINTDATE": doc.properties.last_printed}[name]
            if value is not None:
                done(f, format_date(value, picture))
    return counts


# -- captions and cross-references ----------------------------------------------------------------
@dataclass
class Caption:
    paragraph: OoxElement
    bookmark: str
    label: str

    @property
    def number(self) -> str:
        """The caption's current number, as its SEQ field shows it."""
        f = next(f for f in scan(self.paragraph) if f.name == "SEQ")
        return f.result


def _new_ref_name(doc: WordprocessingDocument) -> str:
    used = {b.get(_q("name")) for b in doc.document.iter(_q("bookmarkStart"))}
    n = 1
    while f"_Ref{n:09d}" in used:
        n += 1
    return f"_Ref{n:09d}"


def _bookmark_around(doc: WordprocessingDocument, paragraph: OoxElement, name: str,
                     first: OoxElement | None = None, last: OoxElement | None = None) -> None:
    from .notes import _next_bookmark_id

    bid = _next_bookmark_id(doc)
    start, end = w.BookmarkStart(id=bid, name=name), w.BookmarkEnd(id=bid)
    if first is not None:
        first.addprevious(start)
    else:
        ppr = paragraph.first("w:pPr")
        if ppr is not None:
            ppr.addnext(start)
        else:
            paragraph.insert(0, start)
    if last is not None:
        last.addnext(end)
    else:
        paragraph.append(end)


def add_caption(doc: WordprocessingDocument, text: str, label: str = "Figure", *,
                before: OoxElement | None = None, after: OoxElement | None = None,
                sep: str = ": ") -> Caption:
    """A caption paragraph "Figure 1: text" (style Caption) with a numbered
    SEQ field.  The label and number are bookmarked, so cross-references read
    "Figure 1".  Goes after/before a body element, or at the end."""
    from .fields import complex_field

    doc.ensure_style("Caption")
    p = w.P(w.PPr(w.PStyle(val="Caption")))
    label_run = w.R(w.T(f"{label} ", space="preserve"))
    runs = complex_field(f"SEQ {label} \\* ARABIC", [w.R(w.T("0"))])
    p.add(label_run, *runs)
    if text:
        p.add(w.R(w.T(f"{sep}{text}", space="preserve")))
    if after is not None:
        after.addnext(p)
    elif before is not None:
        before.addprevious(p)
    else:
        doc.body.add(p)
    name = _new_ref_name(doc)
    _bookmark_around(doc, p, name, first=label_run, last=runs[-1])
    update_fields(doc, only={"SEQ", "REF"})
    return Caption(p, name, label)


def add_cross_reference(doc: WordprocessingDocument, paragraph: OoxElement,
                        target: Caption | OoxElement | str, *, what: str = "text",
                        hyperlink: bool = True) -> OoxElement:
    """A REF/PAGEREF field to a caption, a bookmark name, or a paragraph (a
    heading, say; it gets a hidden bookmark).  ``what``: "text" (the caption's
    "Figure 1", the heading's text), "page" (estimated until Word updates) or
    "position" ("above"/"below")."""
    from .fields import complex_field

    if isinstance(target, Caption):
        name = target.bookmark
    elif isinstance(target, str):
        name = target
    else:
        name = next((b.get(_q("name")) for b in target.iter(_q("bookmarkStart"))
                     if (b.get(_q("name")) or "").startswith("_Ref")), None) or ""
        if not name:
            name = _new_ref_name(doc)
            _bookmark_around(doc, target, name)
    code = {"text": f"REF {name}", "page": f"PAGEREF {name}",
            "position": f"REF {name} \\p"}.get(what)
    if code is None:
        raise ValueError("what is 'text', 'page' or 'position'")
    if hyperlink:
        code += " \\h"
    runs = complex_field(code, [w.R(w.T("?"))])
    paragraph.add(*runs)
    f = next(f for f in scan(paragraph) if f.element is runs[0].first("w:fldChar"))
    update_fields(doc, only={"REF", "PAGEREF"})
    return f.element.getparent()


def add_field(doc: WordprocessingDocument, paragraph: OoxElement, instr: str,
              result: str = "") -> OoxElement:
    """A simple field with ``instr``; its result is computed now when oox can
    (TITLE, DOCPROPERTY "Company", SEQ ...), else ``result`` is cached."""
    from .fields import field

    el = field(instr, result)
    paragraph.add(el)
    words = instr.split()
    if words:
        update_fields(doc, only={words[0].upper()})
    return el


def fields(doc: WordprocessingDocument, all_stories: bool = False) -> list[Field]:
    """A ``Field`` (name, code, result, kind, element) for each field in the
    body (``all_stories`` adds headers, footers, notes and comments)."""
    roots = [doc.document]
    if all_stories:
        for part_type in ("HeaderPart", "FooterPart", "FootnotesPart", "EndnotesPart",
                          "WordprocessingCommentsPart"):
            roots += [p.element for p in doc.related_parts(doc.main_part, part_type)]
    return [f for root in roots for f in scan(root)]
