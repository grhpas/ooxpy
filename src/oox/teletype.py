"""Plain text <-> markup, mirroring odfpy's ``odf.teletype``.

    from oox import teletype
    teletype.extractText(paragraph)            # "Name:\\tAda"
    teletype.addTextToElement(paragraph, "a\\tb\\nc")
"""

from __future__ import annotations

import re

from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

from .grammar import URI

W = URI["w"]
_T, _TAB, _BR, _CR, _P, _R = (f"{{{W}}}{n}" for n in ("t", "tab", "br", "cr", "p", "r"))
_PTAB = f"{{{W}}}ptab"
_NOBREAK_HYPHEN = f"{{{W}}}noBreakHyphen"
_TXBX = f"{{{W}}}txbxContent"
_TEXT_TAGS = frozenset((_T, _TAB, _PTAB, _BR, _CR, _NOBREAK_HYPHEN))


def extractText(element, text_boxes: bool = True) -> str:
    """Visible text of an element subtree.

    ``w:t`` contributes its text, ``w:tab``/``w:ptab`` a tab, ``w:br``/``w:cr``
    a newline.  Deleted text (``w:delText``) and field codes (``w:instrText``)
    are not ``w:t`` and are skipped.  Like odfpy, no separator is inserted
    between paragraphs.

    Markup compatibility is read as Word reads it (``oox.walk``): of each
    ``mc:AlternateContent`` only the branch a reader shows counts, so the
    VML copy Word keeps of every text box isn't counted twice.
    ``text_boxes=False`` also leaves out text boxes below ``element``
    (``w:txbxContent``), as Word's ``Paragraph.Range.Text`` does.
    """
    from .walk import ALTERNATE, iter_story

    if not any(True for _ in element.iter(ALTERNATE, *(() if text_boxes else (_TXBX,)))):
        # the common case: nothing to skip, one plain walk
        return "".join(_pieces(element.iter(_T, _TAB, _PTAB, _BR, _CR, _NOBREAK_HYPHEN)))
    return "".join(_pieces(iter_story(element, *_TEXT_TAGS, text_boxes=text_boxes)))


def _pieces(elements):
    for el in elements:
        tag = el.tag
        if tag == _T:
            yield el.text or ""
        elif tag in (_TAB, _PTAB):
            yield "\t"
        elif tag == _NOBREAK_HYPHEN:
            yield "‑"
        else:
            yield "\n"


_SPLIT = re.compile(r"(\t|\n)")


def addTextToElement(element, s: str) -> None:
    """Append ``s`` to a paragraph or run.

    ``\\t`` becomes ``w:tab`` and ``\\n`` becomes ``w:br``; runs of spaces are kept
    (``xml:space="preserve"`` is set automatically).  Given a ``w:p``, a new
    ``w:r`` is created; given a ``w:r``, content is appended to it.
    """
    from .ns import w

    if element.tag == _P:
        run = w.R()
        element.add(run)
    elif element.tag == _R:
        run = element
    else:
        raise TypeError(f"addTextToElement expects w:p or w:r, got {etree.QName(element).localname}")
    for piece in _SPLIT.split(s):
        if piece == "\t":
            run.add(w.Tab())
        elif piece == "\n":
            run.add(w.Br())
        elif piece:
            run.add(w.T(piece))
