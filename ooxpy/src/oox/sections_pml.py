"""Slide sections (PowerPoint 2010's Home > Section), stored as ``p14:sectionLst``
in the presentation's extension list.

    deck.set_sections([("Introduction", s1), ("Results", s4)])
    deck.sections()          # [("Introduction", [1, 2, 3]), ("Results", [4, 5])]
    deck.set_sections([])    # remove them

Each section starts at the given slide and runs to the next one; slides before
the first section go into a "Default Section", as PowerPoint does.  New slides
join the last section.  The list sits in ``p:extLst``, which PowerPoint 2007
and other readers that don't know it skip.  Section ids are GUIDs derived
from the names and positions, so the same sections give the same bytes.
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING

from .element import OoxElement
from .grammar import URI
from .ns import p, p14

if TYPE_CHECKING:
    from .presentation import PresentationDocument

SECTIONS_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_NS = uuid.UUID("6d1b7f7e-3c1a-4c8e-9a51-0f0f6b6b0a14")


def _ext(deck: PresentationDocument, create: bool) -> OoxElement | None:
    pres = deck.presentation
    ext_lst = pres.first("p:extLst")
    if ext_lst is not None:
        for ext in ext_lst.iterchildren(f"{{{URI['p']}}}ext"):
            if ext.get("uri") == SECTIONS_EXT:
                return ext  # type: ignore[return-value]
    if not create:
        return None
    if ext_lst is None:
        ext_lst = p.ExtLst()
        pres.add(ext_lst)
    ext = p.Ext(uri=SECTIONS_EXT)
    ext_lst.add(ext)
    return ext


def _slide_ids(deck: PresentationDocument) -> list[str]:
    lst = deck.presentation.first("p:sldIdLst")
    return [e.get("id") for e in lst] if lst is not None else []


def set_sections(deck: PresentationDocument, sections) -> None:
    """``[(name, first slide), ...]`` in slide order; first slides are slide
    elements or 1-based numbers.  An empty list removes the sections."""
    ext = _ext(deck, create=False)
    if ext is not None:
        parent = ext.getparent()
        parent.remove(ext)
        if parent is not None and len(parent) == 0:
            parent.getparent().remove(parent)
    sections = list(sections)
    if not sections:
        return
    ids = _slide_ids(deck)
    slides = deck.slides
    starts = []
    for name, first in sections:
        if not name:
            raise ValueError("a section needs a name")
        if isinstance(first, int):
            if not 1 <= first <= len(slides):
                raise IndexError(f"slide {first}: the deck has {len(slides)} slides")
            starts.append(first - 1)
        else:
            try:
                starts.append(next(i for i, s in enumerate(slides) if s is first))
            except StopIteration:
                raise ValueError(f"section {name!r}: that slide isn't in this deck") from None
    if starts != sorted(starts) or len(set(starts)) != len(starts):
        raise ValueError("sections must start at different slides, in slide order")
    names = [n for n, _ in sections]
    if starts and starts[0] > 0:
        names.insert(0, "Default Section")
        starts.insert(0, 0)
    bounds = [*starts, len(ids)]
    lst = p14.SectionLst()
    for i, name in enumerate(names):
        guid = "{" + str(uuid.uuid5(_NS, f"{i}:{name}")).upper() + "}"
        members = [p14.SldId(id=sid) for sid in ids[bounds[i]:bounds[i + 1]]]
        lst.add(p14.Section(p14.SldIdLst(*members), name=name, id=guid))
    _ext(deck, create=True).add(lst)  # type: ignore[union-attr]


def sections(deck: PresentationDocument) -> list[tuple[str, list[int]]]:
    """``[(name, [1-based slide numbers]), ...]``; empty without sections."""
    ext = _ext(deck, create=False)
    if ext is None:
        return []
    number = {sid: i + 1 for i, sid in enumerate(_slide_ids(deck))}
    out = []
    for sec in ext.iter(f"{{{P14}}}section"):
        members = [number[s.get("id")] for s in sec.iter(f"{{{P14}}}sldId")
                   if s.get("id") in number]
        out.append((sec.get("name") or "", members))
    return out


def join_last_section(deck: PresentationDocument, slide_id: int) -> None:
    """A new slide joins the last section (called by ``add_slide``)."""
    ext = _ext(deck, create=False)
    if ext is None:
        return
    found = list(ext.iter(f"{{{P14}}}section"))
    if found:
        lst = found[-1].first("p14:sldIdLst")
        if lst is None:
            lst = p14.SldIdLst()
            found[-1].add(lst)
        lst.add(p14.SldId(id=slide_id))
