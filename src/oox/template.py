"""Find and replace text across runs -- the template workflow.

    doc.replace_text({"{{name}}": "Ada", "{{date}}": "14 November 2026"})

Word splits text into runs wherever formatting, spell-check or editing
history changes, so a placeholder typed as ``{{name}}`` is often stored as
``{{``, ``na``, ``me}}`` in three runs.  Matching is done on the paragraph's
joined text; the replacement goes into the run where the match starts (so it
takes that run's formatting) and the matched characters are removed from the
other runs.  Body, headers, footers, footnotes, endnotes and comments are
searched.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from .grammar import URI

if TYPE_CHECKING:
    from .document import WordprocessingDocument
    from .element import OoxElement

W = URI["w"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _set_text(t: OoxElement, text: str) -> None:
    t.text = text
    if text != text.strip() or "  " in text:
        t.set(XML_SPACE, "preserve")


def replace_in_paragraph(p: OoxElement, old: str, new: str) -> int:
    texts = [t for t in p.iter(f"{{{W}}}t")]
    if not texts:
        return 0
    joined = "".join(t.text or "" for t in texts)
    if old not in joined:
        return 0
    # (start offset of each w:t in the joined string)
    starts, pos = [], 0
    for t in texts:
        starts.append(pos)
        pos += len(t.text or "")
    count = 0
    idx = joined.rfind(old)
    while idx != -1:  # right to left keeps earlier offsets valid
        end = idx + len(old)
        first = max(i for i, s in enumerate(starts) if s <= idx)
        for i in range(first, len(texts)):
            s, txt = starts[i], texts[i].text or ""
            e = s + len(txt)
            if s >= end:
                break
            lo, hi = max(idx, s) - s, min(end, e) - s
            if i == first:
                _set_text(texts[i], txt[:lo] + new + txt[hi:])
            else:
                _set_text(texts[i], txt[:lo] + txt[hi:])
        count += 1
        idx = joined.rfind(old, 0, idx)
    return count


def _roots(doc: WordprocessingDocument):
    yield doc.body
    for part_type in ("HeaderPart", "FooterPart", "FootnotesPart", "EndnotesPart",
                      "WordprocessingCommentsPart"):
        for part in doc.related_parts(doc.main_part, part_type):
            yield part.element


def replace_text(doc: WordprocessingDocument, mapping: dict[str, str]) -> dict[str, int]:
    """Replace every key with its value; returns ``{key: replacements}``."""
    counts = dict.fromkeys(mapping, 0)
    for root in _roots(doc):
        for p in root.iter(f"{{{W}}}p"):
            for old, new in mapping.items():
                if old:
                    counts[old] += replace_in_paragraph(p, old, str(new))
    return counts
