"""Word fields: values Word computes when it lays the document out.

    from oox import fields
    w.P(w.R(w.T("Page ")), fields.page(), w.R(w.T(" of ")), fields.num_pages())

Each helper returns a simple field (``w:fldSimple``) carrying a placeholder
result, which Word replaces with the real value when it renders or updates
fields.  ``field(instr)`` accepts any field instruction ("DATE \\\\@ ...",
"AUTHOR", "FILENAME" ...).
"""

from __future__ import annotations

import re

from .element import OoxElement
from .ns import w


def field(instr: str, placeholder: str = "") -> OoxElement:
    """``<w:fldSimple w:instr=" INSTR ">`` with a cached ``placeholder`` result."""
    return w.FldSimple(w.R(w.T(placeholder or " ")), instr=f" {instr.strip()} ")


def page(placeholder: str = "1") -> OoxElement:
    return field("PAGE", placeholder)


def num_pages(placeholder: str = "1") -> OoxElement:
    return field("NUMPAGES", placeholder)


def section_pages(placeholder: str = "1") -> OoxElement:
    return field("SECTIONPAGES", placeholder)


def date(fmt: str = "d MMMM yyyy", placeholder: str = "") -> OoxElement:
    return field(f'DATE \\@ "{fmt}"', placeholder)


TOKENS = {"PAGE": page, "NUMPAGES": num_pages, "SECTIONPAGES": section_pages}
_TOKEN = re.compile(r"\{(" + "|".join(TOKENS) + r")\}")


def fill_paragraph(p: OoxElement, text: str) -> OoxElement:
    """Append ``text`` to a paragraph, turning ``{PAGE}``, ``{NUMPAGES}`` and
    ``{SECTIONPAGES}`` into fields; ``\\t``/``\\n`` become tab/line break."""
    from . import teletype

    pos = 0
    for m in _TOKEN.finditer(text):
        if m.start() > pos:
            teletype.addTextToElement(p, text[pos:m.start()])
        p.add(TOKENS[m.group(1)]())
        pos = m.end()
    if pos < len(text):
        teletype.addTextToElement(p, text[pos:])
    return p


def complex_field(instr: str, result: list[OoxElement] | None = None) -> list[OoxElement]:
    """Runs for a complex field (begin / instruction / separate / result / end).

    Complex fields can span paragraphs and hold rich results -- the table of
    contents is one.  Returns the runs to place in a paragraph; ``result`` runs
    (the cached value) go between separate and end.
    """
    return [w.R(w.FldChar(fldCharType="begin")),
            w.R(w.InstrText(f" {instr.strip()} ")),
            w.R(w.FldChar(fldCharType="separate")),
            *(result or []),
            w.R(w.FldChar(fldCharType="end"))]


def toc(doc, levels: tuple[int, int] = (1, 3), title: str | None = "Contents",
        update_on_open: bool = False) -> list[OoxElement]:
    """Paragraphs for a table of contents covering heading ``levels``.

    The cached result lists the document's current headings (indented per
    level, without page numbers); Word fills in page numbers when the field
    is updated (right-click > Update Field, or F9).  ``update_on_open=True``
    sets ``w:updateFields``, which makes Word offer to update when the file
    opens -- note that this shows a prompt.
    """
    from .ns import w as _w

    lo, hi = levels
    from .walk import iter_story

    headings = []
    for p in iter_story(doc.body, f"{{{_W}}}p", text_boxes=False):  # Word leaves text boxes out
        lvl = doc.outline_level(p)
        if lo <= lvl <= hi:
            from . import teletype

            text = teletype.extractText(p).strip()
            if text:
                headings.append((lvl, text))
    out: list[OoxElement] = []
    if title:
        doc.ensure_style("TOCHeading")
        out.append(_w.P(_w.PPr(_w.PStyle(val="TOCHeading")), _w.R(_w.T(title))))
    instr = f'TOC \\o "{lo}-{hi}" \\h \\z \\u'
    if not headings:
        headings = [(lo, "Update this field (F9) to build the table of contents.")]
    for i, (lvl, text) in enumerate(headings):
        doc.ensure_style(f"TOC{lvl}")
        p = _w.P(_w.PPr(_w.PStyle(val=f"TOC{lvl}")))
        runs = [_w.R(_w.T(text))]
        if i == 0:
            begin, instr_run, sep = complex_field(instr)[:3]
            p.add(begin, instr_run, sep)
        p.add(*runs)
        if i == len(headings) - 1:
            p.add(_w.R(_w.FldChar(fldCharType="end")))
        out.append(p)
    if update_on_open:
        doc.ensure_settings().add(_w.UpdateFields(), replace=True)
    return out


_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def field_codes(root: OoxElement) -> list[str]:
    """Field names used under ``root`` (``["PAGE", "NUMPAGES"]``), in order;
    simple and complex fields alike."""
    from .walk import iter_story

    out = []
    for el in iter_story(root, f"{{{_W}}}fldSimple", f"{{{_W}}}instrText"):
        code = el.get(f"{{{_W}}}instr") if el.tag.endswith("fldSimple") else el.text
        words = (code or "").split()
        if words:
            out.append(words[0].upper())
    return out
