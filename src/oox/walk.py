"""Walking a Word story the way Word reads it.

Two things make a plain ``element.iter()`` read a document wrongly:

- **Markup compatibility.**  ``mc:AlternateContent`` holds the same content
  several times: a ``mc:Choice`` for newer readers (Word 2010's text-box
  shape) and a ``mc:Fallback`` for older ones (the VML copy Word 2007 reads).
  A reader takes exactly one branch: the first ``mc:Choice`` whose
  ``Requires`` namespaces it understands, otherwise the ``mc:Fallback``
  (ECMA-376 Part 3, §10.2).  Walking both counts every list item, field,
  bookmark and picture in a text box twice.
- **Block-level containers.**  Content controls (``w:sdt``/``w:sdtContent``)
  and custom XML (``w:customXml``) wrap paragraphs and tables, including
  paragraphs that end a section (LibreOffice puts a table of contents and the
  break after it in one).  Code that looks only at ``w:body``'s children
  misses them.

``iter_story`` and ``iter_blocks`` handle both, and every reader uses them.
Code that *changes* a document (accepting revisions, replacing text, filling
controls, allocating ids) walks every copy instead, so the copies stay alike.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING

from .grammar import URI

if TYPE_CHECKING:
    from .element import OoxElement

W, MC = URI["w"], URI["mc"]
ALTERNATE = f"{{{MC}}}AlternateContent"
CHOICE = f"{{{MC}}}Choice"
FALLBACK = f"{{{MC}}}Fallback"
TXBX = f"{{{W}}}txbxContent"
P, TBL = f"{{{W}}}p", f"{{{W}}}tbl"
# block-level wrappers a reader looks through (CT_SdtBlock, CT_CustomXmlBlock)
_BLOCK_WRAPPERS = frozenset({f"{{{W}}}sdt", f"{{{W}}}sdtContent", f"{{{W}}}customXml"})
_KNOWN = frozenset(URI.values())


def chosen_branch(alternate: OoxElement) -> OoxElement | None:
    """The branch of an ``mc:AlternateContent`` a reader shows: the first
    ``mc:Choice`` whose ``Requires`` prefixes all name namespaces oox knows,
    else the ``mc:Fallback``, else None."""
    fallback = None
    for branch in alternate:
        if branch.tag == CHOICE:
            prefixes = (branch.get("Requires") or "").split()
            if prefixes and all(branch.nsmap.get(p) in _KNOWN for p in prefixes):
                return branch
        elif branch.tag == FALLBACK and fallback is None:
            fallback = branch
    return fallback


def iter_story(root: OoxElement, *tags: str, text_boxes: bool = True) -> Iterator[OoxElement]:
    """Descendants of ``root`` (and ``root`` itself) with one of ``tags``
    (Clark names; all elements if none), in document order, reading each
    ``mc:AlternateContent`` through its chosen branch only.  ``text_boxes=False``
    also leaves out text-box content (``w:txbxContent``) below ``root``."""
    wanted = frozenset(tags)
    if not wanted or root.tag in wanted:
        yield root
    stack = [iter(root)]
    while stack:
        el = next(stack[-1], None)
        if el is None:
            stack.pop()
            continue
        tag = el.tag
        if not isinstance(tag, str):  # comments, processing instructions
            continue
        if tag == ALTERNATE:
            branch = chosen_branch(el)
            if branch is not None:
                stack.append(iter(branch))
            continue
        if tag == TXBX and not text_boxes:
            continue
        if not wanted or tag in wanted:
            yield el
        if len(el):
            stack.append(iter(el))


def iter_blocks(container: OoxElement, *tags: str) -> Iterator[OoxElement]:
    """The block-level content of ``container`` (a body, cell, header,
    footer, note, comment or text box): its paragraphs and tables, looking
    through content controls, custom XML and markup-compatibility blocks but
    not into tables.  ``tags`` narrows it (default: paragraphs and tables)."""
    wanted = frozenset(tags or (P, TBL))
    stack = [iter(container)]
    while stack:
        el = next(stack[-1], None)
        if el is None:
            stack.pop()
            continue
        tag = el.tag
        if tag in wanted:
            yield el
        elif tag in _BLOCK_WRAPPERS:
            stack.append(iter(el))
        elif tag == ALTERNATE:
            branch = chosen_branch(el)
            if branch is not None:
                stack.append(iter(branch))


def in_fallback(el: OoxElement) -> bool:
    """Whether ``el`` sits in a branch of ``mc:AlternateContent`` that a reader
    doesn't show (a fallback, or a choice not taken)."""
    child = el
    for anc in el.iterancestors():
        if anc.tag == ALTERNATE:
            if child is not chosen_branch(anc):
                return True
        child = anc
    return False


def twins(root: OoxElement) -> list[tuple[OoxElement, OoxElement]]:
    """``(shown, hidden)`` pairs of text-box contents: for every
    ``mc:AlternateContent`` under ``root``, the n-th ``w:txbxContent`` of the
    branch a reader shows with the n-th of every other branch.  Word writes
    the same paragraphs in both, so changes made to the shown copy can be
    repeated in the hidden one."""
    out = []
    for alt in root.iter(ALTERNATE):
        shown = chosen_branch(alt)
        if shown is None:
            continue
        shown_boxes = list(shown.iter(TXBX))
        for branch in alt:
            if branch is shown or branch.tag not in (CHOICE, FALLBACK):
                continue
            for a, b in zip(shown_boxes, branch.iter(TXBX), strict=False):
                out.append((a, b))
    return out
