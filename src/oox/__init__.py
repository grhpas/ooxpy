"""oox -- a schema-driven, odfpy-style library for Office Open XML.

    from oox import WordprocessingDocument
    from oox.ns import w

    doc = WordprocessingDocument.new()
    doc.body.add(w.P(w.PPr(w.Jc(val="center")), w.R(w.T("Hello"))))
    doc.save("hello.docx")

Also: SpreadsheetDocument (.xlsx), PresentationDocument (.pptx), open_document(path).
"""

from ._generated import SOURCE as GRAMMAR_SOURCE
from .document import OoxDocument, WordprocessingDocument, open_document
from .element import OoxElement, lenient
from .errors import (
                     DuplicateChild,
                     InvalidAttribute,
                     InvalidAttributeValue,
                     InvalidChild,
                     MissingAttribute,
                     OoxError,
                     OoxValidationError,
                     PackageError,
                     TargetError,
)
from .opc import Limits, Package, Part
from .presentation import PresentationDocument
from .spreadsheet import SpreadsheetDocument

__version__ = "0.6.0"

__all__ = [
                     "GRAMMAR_SOURCE",
                     "DuplicateChild",
                     "InvalidAttribute",
                     "InvalidAttributeValue",
                     "InvalidChild",
                     "Limits",
                     "MissingAttribute",
                     "OoxDocument",
                     "OoxElement",
                     "OoxError",
                     "OoxValidationError",
                     "Package",
                     "PackageError",
                     "Part",
                     "PresentationDocument",
                     "SpreadsheetDocument",
                     "TargetError",
                     "WordprocessingDocument",
                     "lenient",
                     "open_document",
]
