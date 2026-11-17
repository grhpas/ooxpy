"""Content controls (Word forms): text, rich text, drop-down, combo box,
date and check box; reading and filling them by tag.

    p = doc.add_paragraph("Name: ")
    doc.add_text_control(p, tag="name", title="Name", placeholder="Your name")
    doc.add_dropdown(p, tag="dept", items=["Sales", "Finance", "IT"])
    doc.add_date_control(p, tag="start", fmt="d MMMM yyyy")
    doc.add_checkbox(p, tag="agree", checked=False)

    doc.fill_controls({"name": "Ada", "dept": "Finance",
                       "start": date(2026, 11, 14), "agree": True})
    doc.content_controls()   # [ContentControl(tag="name", type="text", value="Ada", ...)]

Controls are inline (inside a paragraph).  Readers also see block-level ones
(controls holding whole paragraphs, as Word templates often use).  Check
boxes are an Office 2010 feature (``w14:checkbox``): Word 2007 shows them
as a control holding the ☐/☒ character, which ``fill_controls`` keeps in
step.
"""

from __future__ import annotations

import datetime as _dt
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import w

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
W14 = URI["w14"]
CHECKED, UNCHECKED = "☒", "☐"  # ☒ ☐
_TYPES = {"text": "text", "dropDownList": "dropdown", "comboBox": "combo", "date": "date",
          "picture": "picture", "group": "group", "docPartObj": "gallery",
          "docPartList": "gallery", "richText": "rich", "equation": "equation",
          "citation": "citation", "bibliography": "bibliography"}


def _q(local: str, ns: str = W) -> str:
    return f"{{{ns}}}{local}"


def _next_id(doc: WordprocessingDocument) -> int:
    ids = [int(el.get(_q("val"), "0")) for el in doc.document.iter(_q("id"))
           if (el.get(_q("val")) or "").lstrip("-").isdigit()]
    return max([100, *ids]) + 1


def _sdt_pr(doc: WordprocessingDocument, tag: str, title: str | None, *extra: OoxElement,
            placeholder: bool = False) -> OoxElement:
    if not tag:
        raise ValueError("controls need a tag (used to find and fill them)")
    pr = w.SdtPr()
    if title:
        pr.add(w.Alias(val=title))
    pr.add(w.Tag(val=tag), w.Id(val=_next_id(doc)))
    if placeholder:
        pr.add(w.ShowingPlcHdr())
    pr.add(*extra)
    return pr


def _content_run(doc: WordprocessingDocument, text: str, placeholder: bool,
                 rpr: OoxElement | None = None) -> OoxElement:
    from . import teletype

    if placeholder:
        doc.ensure_style("PlaceholderText")
        r = w.R(w.RPr(w.RStyle(val="PlaceholderText")))
    else:
        r = w.R(rpr) if rpr is not None else w.R()
    teletype.addTextToElement(r, text)
    return r


def _inline(paragraph: OoxElement, pr: OoxElement, run: OoxElement) -> OoxElement:
    sdt = w.Sdt(pr, w.SdtContent(run))
    paragraph.add(sdt)
    return sdt


def add_text_control(doc: WordprocessingDocument, paragraph: OoxElement, *, tag: str,
                     title: str | None = None, text: str | None = None,
                     placeholder: str = "Click here to enter text.", rich: bool = False,
                     multiline: bool = False) -> OoxElement:
    """A plain-text control (``rich=True``: rich text, which may hold
    formatting); shows ``placeholder`` until filled."""
    show = text is None or text == ""
    extra = [] if rich else [w.Text(multiLine=True) if multiline else w.Text()]
    pr = _sdt_pr(doc, tag, title, *extra, placeholder=show)
    return _inline(paragraph, pr, _content_run(doc, placeholder if show else str(text), show))


def add_dropdown(doc: WordprocessingDocument, paragraph: OoxElement, *, tag: str,
                 items: list[str | tuple[str, str]], value: str | None = None,
                 title: str | None = None, combo: bool = False,
                 placeholder: str = "Choose an item.") -> OoxElement:
    """A drop-down list (``combo=True``: a combo box, which also takes typed
    text).  ``items``: display texts, or ``(display, value)`` pairs."""
    pairs = [(i, i) if isinstance(i, str) else (i[0], i[1]) for i in items]
    if not pairs:
        raise ValueError("a drop-down needs items")
    factory = w.ComboBox if combo else w.DropDownList
    kind = factory(*(w.ListItem(displayText=d, value=v) for d, v in pairs))
    show = value is None
    display = None if show else _display_for(pairs, value, combo)
    if display is not None and not combo:
        kind.set(_q("lastValue"), dict((d, v) for d, v in pairs).get(display, display))
    pr = _sdt_pr(doc, tag, title, kind, placeholder=show)
    return _inline(paragraph, pr, _content_run(doc, placeholder if show else str(display), show))


def _display_for(pairs: list[tuple[str, str]], value: str, combo: bool) -> str:
    for d, v in pairs:
        if value in (d, v):
            return d
    if combo:
        return value
    raise ValueError(f"{value!r} is not one of the items {[d for d, _ in pairs]}")


def _format_date(value: _dt.date, fmt: str) -> str:
    from .fieldcalc import format_date

    if not isinstance(value, _dt.datetime):
        value = _dt.datetime(value.year, value.month, value.day)
    return format_date(value, fmt)


def add_date_control(doc: WordprocessingDocument, paragraph: OoxElement, *, tag: str,
                     value: _dt.date | None = None, fmt: str = "d MMMM yyyy",
                     title: str | None = None, lang: str = "en-US",
                     placeholder: str = "Click here to enter a date.") -> OoxElement:
    """A date picker; ``fmt`` is Word's date picture (``dd/MM/yyyy`` ...)."""
    attrs: dict[str, Any] = {}
    if value is not None:
        d = value if isinstance(value, _dt.datetime) else _dt.datetime(value.year, value.month,
                                                                         value.day)
        attrs["fullDate"] = d.replace(microsecond=0).isoformat() + "Z"
    kind = w.Date(w.DateFormat(val=fmt), w.Lid(val=lang),
                  w.StoreMappedDataAs(val="dateTime"), w.Calendar(val="gregorian"), **attrs)
    show = value is None
    pr = _sdt_pr(doc, tag, title, kind, placeholder=show)
    return _inline(paragraph, pr, _content_run(
        doc, placeholder if show else _format_date(value, fmt), show))  # type: ignore[arg-type]


def add_checkbox(doc: WordprocessingDocument, paragraph: OoxElement, *, tag: str,
                 checked: bool = False, title: str | None = None) -> OoxElement:
    """A check box (Office 2010+; Word 2007 shows the ☐/☒ character)."""
    from .ns import w14

    box = w14.Checkbox(w14.Checked(val="1" if checked else "0"),
                       w14.CheckedState(val="2612", font="MS Gothic"),
                       w14.UncheckedState(val="2610", font="MS Gothic"))
    pr = _sdt_pr(doc, tag, title, box)
    rpr = w.RPr(w.RFonts(ascii="MS Gothic", hAnsi="MS Gothic", eastAsia="MS Gothic"))
    return _inline(paragraph, pr, _content_run(doc, CHECKED if checked else UNCHECKED, False, rpr))


# -- reading and filling ----------------------------------------------------------------------
def _stories(doc: WordprocessingDocument):
    yield doc.document
    for part_type in ("HeaderPart", "FooterPart", "FootnotesPart", "EndnotesPart"):
        for part in doc.related_parts(doc.main_part, part_type):
            yield part.element


def _kind(pr: OoxElement) -> str:
    for child in pr:
        ns, _, local = child.tag[1:].partition("}")
        if ns == W14 and local == "checkbox":
            return "checkbox"
        if ns == W and local in _TYPES:
            return _TYPES[local]
    return "rich"


def _info(sdt: OoxElement) -> dict[str, Any]:
    from . import teletype

    pr = sdt.first("w:sdtPr")
    content = sdt.first("w:sdtContent")
    tag = pr.first("w:tag") if pr is not None else None
    alias = pr.first("w:alias") if pr is not None else None
    kind = _kind(pr) if pr is not None else "rich"
    text = teletype.extractText(content) if content is not None else ""
    info: dict[str, Any] = {
        "tag": tag.get(_q("val")) if tag is not None else None,
        "title": alias.get(_q("val")) if alias is not None else None,
        "type": kind,
        "placeholder": pr is not None and pr.first("w:showingPlcHdr") is not None,
        "value": text,
    }
    if info["placeholder"]:
        info["value"] = None
    if kind in ("dropdown", "combo") and pr is not None:
        box = pr.first("w:dropDownList") if kind == "dropdown" else pr.first("w:comboBox")
        info["items"] = [li.get(_q("displayText")) or li.get(_q("value"))
                         for li in box.iterchildren(_q("listItem"))] if box is not None else []
    if kind == "checkbox" and pr is not None:
        el = pr.find(f"{_q('checkbox', W14)}/{_q('checked', W14)}")
        info["value"] = el is not None and el.get(_q("val", W14)) in ("1", "true")
    if kind == "date" and pr is not None:
        d = pr.first("w:date")
        full = d.get(_q("fullDate")) if d is not None else None
        info["date"] = _dt.datetime.fromisoformat(full.replace("Z", "")).date() if full else None
    return info


def content_controls(doc: WordprocessingDocument) -> list:
    """A ``ContentControl`` for every content control (body, headers,
    footers, notes), in order: tag, title, type, value, placeholder.
    ``value`` is the text (None while the placeholder shows), a bool for
    check boxes; drop-downs have ``items``, dates a ``date``."""
    from .records import ContentControl
    from .walk import iter_story

    return [ContentControl(**_info(sdt)) for root in _stories(doc)
            for sdt in iter_story(root, _q("sdt"))]


def _set_content(doc: WordprocessingDocument, sdt: OoxElement, text: str) -> None:
    """Replace the control's text, keeping the first run's formatting (minus
    the placeholder style)."""
    import copy

    from . import teletype

    content = sdt.first("w:sdtContent")
    assert content is not None
    first_run = content.find(f".//{_q('r')}")
    rpr = copy.deepcopy(first_run.first("w:rPr")) if first_run is not None and \
        first_run.first("w:rPr") is not None else None
    if rpr is not None:
        style = rpr.first("w:rStyle")
        if style is not None and style.get(_q("val")) == "PlaceholderText":
            rpr.remove(style)
        if len(rpr) == 0:
            rpr = None
    paragraphs = list(content.iterchildren(_q("p")))
    if paragraphs:  # block-level control: keep the first paragraph's properties
        keep = paragraphs[0]
        for p in paragraphs[1:]:
            content.remove(p)
        for child in list(keep):
            if child.tag != _q("pPr"):
                keep.remove(child)
        target = keep
    else:
        for child in list(content):
            content.remove(child)
        target = content
    r = w.R(rpr) if rpr is not None else w.R()
    teletype.addTextToElement(r, text)
    target.append(r)
    pr = sdt.first("w:sdtPr")
    if pr is not None and pr.first("w:showingPlcHdr") is not None:
        pr.remove(pr.first("w:showingPlcHdr"))


def fill_controls(doc: WordprocessingDocument, values: dict[str, Any], *,
                  strict: bool = True) -> int:
    """Fill controls by tag: text for text/rich/combo, an item (display text or
    value) for drop-downs, a date for dates, a bool for check boxes.  Every
    control with the tag is filled.  ``strict``: unknown tags raise
    ``KeyError``.  Returns the number of controls filled."""
    from .ns import w14

    seen: set[str] = set()
    count = 0
    for root in _stories(doc):
        for sdt in list(root.iter(_q("sdt"))):
            pr = sdt.first("w:sdtPr")
            tag_el = pr.first("w:tag") if pr is not None else None
            tag = tag_el.get(_q("val")) if tag_el is not None else None
            if tag not in values or pr is None:
                continue
            value = values[tag]
            seen.add(tag)
            kind = _kind(pr)
            if kind == "checkbox":
                box = pr.find(_q("checkbox", W14))
                assert box is not None
                checked = box.find(_q("checked", W14))
                if checked is None:
                    checked = w14.Checked()
                    box.insert(0, checked)
                checked.set(_q("val", W14), "1" if value else "0")
                _set_content(doc, sdt, CHECKED if value else UNCHECKED)
            elif kind in ("dropdown", "combo"):
                box = pr.first("w:dropDownList") if kind == "dropdown" else pr.first("w:comboBox")
                pairs = [(li.get(_q("displayText")) or li.get(_q("value")) or "",
                          li.get(_q("value")) or li.get(_q("displayText")) or "")
                         for li in box.iterchildren(_q("listItem"))] if box is not None else []
                display = _display_for(pairs, str(value), kind == "combo")
                if box is not None and kind == "dropdown":
                    box.set(_q("lastValue"), dict(pairs).get(display, display))
                _set_content(doc, sdt, display)
            elif kind == "date":
                d = pr.first("w:date")
                fmt_el = d.first("w:dateFormat") if d is not None else None
                fmt = fmt_el.get(_q("val")) if fmt_el is not None else "d MMMM yyyy"
                if isinstance(value, _dt.date):
                    dt = value if isinstance(value, _dt.datetime) else \
                        _dt.datetime(value.year, value.month, value.day)
                    if d is not None:
                        d.set(_q("fullDate"), dt.replace(microsecond=0).isoformat() + "Z")
                    _set_content(doc, sdt, _format_date(dt, fmt or "d MMMM yyyy"))
                else:
                    _set_content(doc, sdt, str(value))
            else:
                _set_content(doc, sdt, str(value))
            count += 1
    missing = set(values) - seen
    if strict and missing:
        raise KeyError(f"no content control tagged {sorted(missing)}")
    return count
