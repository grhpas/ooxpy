"""odfpy compatibility: the names an odfpy program reaches for.

    # odfpy                                   # oox
    from odf.opendocument import load    ->   from oox.odfcompat import load
    from odf import teletype             ->   from oox import teletype
    OpenDocumentText()                   ->   OpenDocumentText()          (WordprocessingDocument.new)
    OpenDocumentSpreadsheet()            ->   OpenDocumentSpreadsheet()   (SpreadsheetDocument.new)
    OpenDocumentPresentation()           ->   OpenDocumentPresentation()  (PresentationDocument.new)

Element-level compatibility (``addElement``, ``childNodes``,
``getElementsByType``, ``attributes=``/``parent=``/``check_grammar=`` keywords,
odfpy attribute spellings, element-valued style references) is built into
``OoxElement`` and the factories, so it works everywhere -- not only through
this module.  What cannot be made identical, and why, is listed in
docs/ODFPY_COMPAT.md.
"""

from __future__ import annotations

import os
from typing import IO

from . import teletype
from ._odfdom import TextNode
from .document import OoxDocument, WordprocessingDocument, open_document
from .presentation import PresentationDocument
from .spreadsheet import SpreadsheetDocument

__all__ = [
    "OpenDocumentPresentation",
    "OpenDocumentSpreadsheet",
    "OpenDocumentText",
    "Text",
    "TextNode",
    "load",
    "teletype",
]

Text = TextNode  # odfpy: odf.element.Text


def OpenDocumentText() -> WordprocessingDocument:
    """New WordprocessingML (.docx) document."""
    return WordprocessingDocument.new()


def OpenDocumentSpreadsheet() -> SpreadsheetDocument:
    """New SpreadsheetML (.xlsx) workbook.  Unlike odfpy's, it starts with one
    sheet ("Sheet1"), because Excel refuses a workbook without sheets."""
    return SpreadsheetDocument.new()


def OpenDocumentPresentation() -> PresentationDocument:
    """New PresentationML (.pptx) deck with master, blank layout and theme."""
    return PresentationDocument.new()


def load(odffile: str | os.PathLike | IO[bytes]) -> OoxDocument:
    """Open any OOXML package with the matching document class (docx/xlsx/pptx)."""
    return open_document(odffile)
