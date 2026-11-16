"""Hyperlinks and click actions on slides.

    deck.set_link(slide, shape, "https://example.org/")             # click the shape
    deck.set_link(slide, shape, "https://example.org/", text="docs")  # just that text
    deck.set_link(slide, shape, other_slide)                        # jump to a slide
    deck.set_link(slide, shape, "next")                             # next/previous/first/last/end
    deck.set_link(slide, shape, None)                               # remove
    deck.links(slide)   # [{"shape": "TextBox 2", "text": "docs", "kind": "url",
                        #   "target": "https://example.org/", "tooltip": None}, ...]

A whole-shape link is ``a:hlinkClick`` in the shape's ``p:cNvPr``; a text link
is ``a:hlinkClick`` in the runs' ``a:rPr`` (PowerPoint colours and underlines
them from the theme).  URLs (``http``, ``https``, ``mailto``, ``file``) are
external relationships, slide jumps relationships to the slide part with the
``ppaction://hlinksldjump`` action, and show jumps ``ppaction://hlinkshowjump``.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from .element import OoxElement
from .grammar import URI
from .ns import a

if TYPE_CHECKING:
    from .presentation import PresentationDocument

A, P, R = URI["a"], URI["p"], URI["r"]
RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
SLIDE_JUMP = "ppaction://hlinksldjump"
SHOW_JUMPS = {"next": "nextslide", "previous": "previousslide", "first": "firstslide",
              "last": "lastslide", "last_viewed": "lastslideviewed", "end": "endshow"}
_SHOW_JUMP = "ppaction://hlinkshowjump?jump="


def _shape(deck: PresentationDocument, slide: OoxElement, shape) -> OoxElement:
    if isinstance(shape, OoxElement):
        return shape
    for el in deck._tree(slide).iter():
        if not isinstance(el.tag, str) or el.tag.split("}")[1] != "cNvPr":
            continue
        if el.get("name") == shape:
            holder = el.getparent().getparent()
            if holder is not None:
                return holder  # type: ignore[return-value]
    raise KeyError(f"no shape named {shape!r} on this slide")


def _hlink(deck: PresentationDocument, slide: OoxElement, target, tooltip: str | None
           ) -> OoxElement:
    part = deck.part_of(slide)
    attrs: dict = {}
    if tooltip:
        attrs["tooltip"] = tooltip
    if isinstance(target, OoxElement):
        target_part = deck.part_of(target)
        if target_part is part:
            raise ValueError("a slide can't link to itself")
        rid = deck.package.relate(part, target_part, RT_SLIDE)
        return a.HlinkClick(r_id=rid, action=SLIDE_JUMP, **attrs)
    if not isinstance(target, str) or not target:
        raise TypeError("target: a URL, a slide, or one of " + ", ".join(SHOW_JUMPS))
    if target in SHOW_JUMPS:
        return a.HlinkClick(r_id="", action=_SHOW_JUMP + SHOW_JUMPS[target], **attrs)
    if ":" not in target:
        raise ValueError(f"{target!r}: a URL needs a scheme (https://, mailto:, file:///), "
                         f"or use one of {sorted(SHOW_JUMPS)}")
    rid = part.rels.add(RT_HYPERLINK, target, external=True)
    return a.HlinkClick(r_id=rid, **attrs)


def _split_run(run: OoxElement, at: int) -> OoxElement:
    """Split ``a:r`` after ``at`` characters; returns the new right-hand run."""
    t = run.find(f"{{{A}}}t")
    text = (t.text or "") if t is not None else ""
    right = copy.deepcopy(run)
    if t is not None:
        t.text = text[:at]
    rt = right.find(f"{{{A}}}t")
    if rt is not None:
        rt.text = text[at:]
    run.addnext(right)
    return right  # type: ignore[return-value]


def _run_texts(para: OoxElement) -> list[tuple[OoxElement, str]]:
    return [(r, r.findtext(f"{{{A}}}t") or "") for r in para.iterchildren(f"{{{A}}}r")]


def _split_at(para: OoxElement, offset: int) -> None:
    """Make ``offset`` (in the paragraph's run text) fall between two runs."""
    pos = 0
    for r, t in _run_texts(para):
        if pos < offset < pos + len(t):
            _split_run(r, offset - pos)
            return
        pos += len(t)


def _runs_for(body: OoxElement, text: str) -> list[OoxElement]:
    """The runs exactly covering each occurrence of ``text`` within a
    paragraph, splitting runs at the edges."""
    found: list[OoxElement] = []
    for para in body.iter(f"{{{A}}}p"):
        start = 0
        while True:
            joined = "".join(t for _, t in _run_texts(para))
            hit = joined.find(text, start)
            if hit < 0:
                break
            end = hit + len(text)
            _split_at(para, hit)
            _split_at(para, end)
            pos = 0
            for r, t in _run_texts(para):
                if t and hit <= pos and pos + len(t) <= end:
                    found.append(r)
                pos += len(t)
            start = end
    return found


def set_link(deck: PresentationDocument, slide: OoxElement, shape, target, *,
             text: str | None = None, tooltip: str | None = None) -> OoxElement:
    """Link a shape (element or name) or the occurrences of ``text`` in it to
    ``target``: a URL, a slide (element), or "next", "previous", "first",
    "last", "last_viewed", "end"; ``None`` removes the link.  Returns the shape."""
    sp = _shape(deck, slide, shape)
    if text is None:
        cnv = sp.find(f".//{{{P}}}cNvPr")
        if cnv is None:
            raise ValueError("this shape has no non-visual properties to carry a link")
        old = cnv.find(f"{{{A}}}hlinkClick")
        if old is not None:
            cnv.remove(old)
        if target is not None:
            cnv.add(_hlink(deck, slide, target, tooltip))  # type: ignore[union-attr]
        return sp
    if not text:
        raise ValueError("text: give the text to link, or None for the whole shape")
    body = sp.find(f"{{{P}}}txBody")
    if body is None:
        raise ValueError("this shape has no text")
    runs = _runs_for(body, text)
    if not runs:
        raise ValueError(f"{text!r} is not in the shape's text")
    link = _hlink(deck, slide, target, tooltip) if target is not None else None
    for r in runs:
        rpr = r.find(f"{{{A}}}rPr")
        if rpr is None:
            rpr = a.RPr(lang="en-US", dirty=False)
            r.insert(0, rpr)
        old = rpr.find(f"{{{A}}}hlinkClick")
        if old is not None:
            rpr.remove(old)
        if link is not None:
            rpr.add(copy.deepcopy(link))  # type: ignore[union-attr]
    return sp


def _describe(deck: PresentationDocument, part, link: OoxElement) -> dict:
    action = link.get("action") or ""
    rid = link.get(f"{{{R}}}id") or ""
    rel = part.rels.items.get(rid) if rid else None
    if action.startswith(_SHOW_JUMP):
        jump = action[len(_SHOW_JUMP):]
        name = {v: k for k, v in SHOW_JUMPS.items()}.get(jump, jump)
        return {"kind": "jump", "target": name}
    if action == SLIDE_JUMP and rel is not None:
        target = part.resolve(rel.target)
        slides = [deck.part_of(s).partname for s in deck.slides]
        return {"kind": "slide", "target": slides.index(target) + 1 if target in slides else None}
    if rel is not None and rel.reltype == RT_HYPERLINK:
        return {"kind": "url", "target": rel.target}
    return {"kind": "other", "target": action or None}


def links(deck: PresentationDocument, slide: OoxElement) -> list[dict]:
    """Every click link on a slide, in document order: shape name, linked
    text (``None`` for a whole-shape link), kind (url, slide, jump, other),
    target (URL, 1-based slide number, jump name) and tooltip."""
    part = deck.part_of(slide)
    out = []
    for cnv in deck._tree(slide).iter(f"{{{P}}}cNvPr"):
        holder = cnv.getparent().getparent() if cnv.getparent() is not None else None
        link = cnv.find(f"{{{A}}}hlinkClick")
        if link is not None:
            out.append({"shape": cnv.get("name"), "text": None,
                        **_describe(deck, part, link), "tooltip": link.get("tooltip")})
        if holder is None:
            continue
        body = holder.find(f"{{{P}}}txBody")
        if body is None:
            continue
        # consecutive runs with the same link read as one link
        for para in body.iter(f"{{{A}}}p"):
            current, texts = None, []
            for r in [*para.iterchildren(f"{{{A}}}r"), None]:
                lk = r.find(f"{{{A}}}rPr/{{{A}}}hlinkClick") if r is not None else None
                key = (lk.get(f"{{{R}}}id"), lk.get("action")) if lk is not None else None
                if current is not None and key != current[0]:
                    out.append({"shape": cnv.get("name"), "text": "".join(texts),
                                **_describe(deck, part, current[1]),
                                "tooltip": current[1].get("tooltip")})
                    current, texts = None, []
                if lk is not None:
                    if current is None:
                        current = (key, lk)
                    texts.append(r.findtext(f"{{{A}}}t") or "")  # type: ignore[union-attr]
    return out
