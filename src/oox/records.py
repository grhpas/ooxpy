"""What the readers return: small frozen records with named fields.

    for link in doc.hyperlinks():
        print(link.text, link.target)
    book.comments("Sheet1")["B2"].author
    deck.shapes(slide)[0].kind

One concept has one class in all three formats (a ``Hyperlink`` from Word,
Excel or PowerPoint, a ``Comment`` from Word or Excel).  Where a writer
exists, field names are its keyword names, so a record can be written back:
``book.set_page_setup("T", **book.page_setup("S").as_dict())``.  Lengths
are in points.

Fields are only ever added, never renamed or removed, within a major
version; build records with keywords if you build them at all.  ``as_dict()``
gives a plain dict (without the XML elements some records carry), for JSON.
"""

from __future__ import annotations

from dataclasses import dataclass, field, fields
from typing import Any

from .element import OoxElement


def _plain(value: Any) -> Any:
    if isinstance(value, Record):
        return value.as_dict()
    if isinstance(value, tuple) and hasattr(value, "_fields"):  # Border, Tab, LineSpacing
        return type(value)(*(_plain(v) for v in value))
    if isinstance(value, (list, tuple)):
        return type(value)(_plain(v) for v in value)
    if isinstance(value, dict):
        return {k: _plain(v) for k, v in value.items()}
    if hasattr(value, "__dataclass_fields__"):
        return {f.name: _plain(getattr(value, f.name)) for f in fields(value)}
    return value


@dataclass(frozen=True)
class Record:
    """Base class: ``as_dict()`` for JSON and for passing back to writers."""

    def as_dict(self) -> dict[str, Any]:
        """The fields as a dict, without XML elements (``element`` fields)."""
        return {f.name: _plain(getattr(self, f.name)) for f in fields(self)
                if not isinstance(getattr(self, f.name), OoxElement)
                and f.name != "element"}


# -- shared ------------------------------------------------------------------------------------
@dataclass(frozen=True)
class Hyperlink(Record):
    """A link.  ``target`` is the address as the format's writer takes it (a
    URL; in Excel also ``#Sheet!A1``; in PowerPoint also a jump such as
    "next"); ``anchor`` a place inside the document (a Word bookmark, a
    PowerPoint slide number); ``kind`` is "url", "bookmark", "location",
    "slide", "jump" or "other"; ``cell`` (Excel) and ``shape`` (PowerPoint)
    say where it is."""

    text: str | None
    target: str | None = None
    anchor: str | int | None = None
    tooltip: str | None = None
    kind: str = "url"
    cell: str | None = None
    shape: str | None = None


@dataclass(frozen=True)
class Comment(Record):
    """A comment.  ``anchor_text`` is the commented text (Word); ``cell`` the
    commented cell (Excel)."""

    text: str
    author: str | None = None
    date: str | None = None
    initials: str | None = None
    anchor_text: str | None = None
    cell: str | None = None
    id: int | None = None


@dataclass(frozen=True)
class Chart(Record):
    """A chart, in Excel's ``XlChartType`` vocabulary: ``kind`` ("column",
    "pie" ...), ``plot_kinds`` (one per plot of a combination chart);
    ``series`` are ``oox.charts.Series`` (each with its ``kind`` and
    ``axis``).  ``anchor`` is the top-left cell and ``name`` the chart's
    name (Excel); ``rid`` the chart part's relationship id."""

    kind: str | None
    plot_kinds: list = field(default_factory=list)
    title: str | None = None
    legend: str | None = None
    x_title: str | None = None
    y_title: str | None = None
    y2_title: str | None = None
    number_format: str | None = None
    labels: list = field(default_factory=list)
    series: list = field(default_factory=list)
    categories: list | None = None
    anchor: str | None = None
    name: str | None = None
    rid: str | None = None


@dataclass(frozen=True)
class CellFormat(Record):
    """A cell's (or a highlight's) formatting, as ``cell_style`` options;
    None where the format doesn't say."""

    number_format: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    size: float | None = None
    font: str | None = None
    color: str | None = None
    fill: str | None = None
    border: Any = None
    align: str | None = None
    valign: str | None = None
    wrap: bool | None = None


# -- Word --------------------------------------------------------------------------------------
@dataclass(frozen=True)
class Note(Record):
    """A footnote or endnote."""

    id: int
    text: str


@dataclass(frozen=True)
class ContentControl(Record):
    """A content control (form field): ``type`` is "text", "rich_text",
    "dropdown", "combo", "date", "checkbox" ...; ``placeholder`` True while it
    shows its placeholder text."""

    tag: str | None
    title: str | None = None
    type: str | None = None
    placeholder: bool = False
    value: Any = None
    items: list | None = None
    date: Any = None


@dataclass(frozen=True)
class Picture(Record):
    """A picture.  ``kind`` "drawing" (DrawingML) or "vml"; ``width`` and
    ``height`` in points; ``wrap`` "inline" for inline pictures; ``embed`` /
    ``link`` / ``svg`` are relationship ids (``doc.image_of(picture)`` gives
    the bytes)."""

    kind: str
    name: str | None = None
    descr: str | None = None
    width: float | None = None
    height: float | None = None
    floating: bool = False
    wrap: str | None = None
    embed: str | None = None
    link: str | None = None
    svg: str | None = None
    element: OoxElement | None = field(default=None, repr=False, compare=False)


@dataclass(frozen=True)
class Protection(Record):
    """Word's editing restriction: ``edit`` "readOnly", "comments",
    "trackedChanges" or "forms"; ``enforced``; ``password`` whether one is set."""

    edit: str | None
    enforced: bool = False
    password: bool = False
    algorithm: str | None = None
    formatting: bool = False


@dataclass(frozen=True)
class Revision(Record):
    """A tracked change: ``kind`` "insert", "delete", "move-from", "move-to",
    "paragraph-insert", "paragraph-delete", "row-insert", "row-delete" or
    "format"."""

    kind: str
    author: str | None = None
    date: str | None = None
    text: str = ""


@dataclass(frozen=True)
class TextBox(Record):
    """A text box, counted once (``vml_only``: no DrawingML shape, as in
    files from Word 2007 and older)."""

    text: str
    paragraphs: int = 0
    vml_only: bool = False
    name: str | None = None


# -- Excel -------------------------------------------------------------------------------------
@dataclass(frozen=True)
class ConditionalFormat(Record):
    """A conditional formatting rule, as ``conditional_format`` options:
    ``rule`` is oox's name ("cell", "data_bar", "top" ...; None for rules
    oox doesn't write), ``type`` Excel's; ``style`` the highlight."""

    ref: str
    rule: str | None
    type: str | None = None
    priority: int = 0
    operator: str | None = None
    value: Any = None
    values: list | None = None
    formula: str | None = None
    formulas: list = field(default_factory=list)
    colors: list | None = None
    color: str | None = None
    icons: str | None = None
    n: int | None = None
    percent: bool | None = None
    text: str | None = None
    how: str | None = None
    stop: bool = False
    style: CellFormat | None = None


@dataclass(frozen=True)
class Validation(Record):
    """A data validation, as ``add_validation`` options: ``values`` (a
    list), ``source`` (a range), ``whole`` or ``decimal`` (min, max)."""

    ref: str
    type: str
    values: list | None = None
    source: str | None = None
    whole: tuple | None = None
    decimal: tuple | None = None
    allow_blank: bool = True
    prompt: str | None = None
    error: str | None = None


@dataclass(frozen=True)
class ExcelTable(Record):
    """An Excel table (Insert > Table): ``ref``, column names, style and its
    switches as ``add_table`` takes them; ``totals_row`` whether it has one."""

    name: str | None
    ref: str
    columns: list = field(default_factory=list)
    style: str | None = None
    stripes: bool = True
    first_column: bool = False
    last_column: bool = False
    header: bool = True
    totals_row: bool = False


@dataclass(frozen=True)
class PageSetup(Record):
    """A sheet's print setup, as ``set_page_setup`` options (None: not set);
    ``margins`` in points."""

    orientation: str | None = None
    paper: Any = None
    fit_width: int | None = None
    fit_height: int | None = None
    scale: int | None = None
    margins: dict | None = None
    print_area: str | None = None
    repeat_rows: str | None = None
    repeat_cols: str | None = None
    header: Any = None
    footer: Any = None
    gridlines: bool | None = None
    center_horizontally: bool | None = None
    center_vertically: bool | None = None


@dataclass(frozen=True)
class SheetOptions(Record):
    """A sheet's options, as ``set_sheet_options`` takes them, plus
    ``protected``."""

    tab_color: str | None = None
    hidden: bool | str = False
    zoom: int = 100
    gridlines: bool = True
    protected: bool = False


@dataclass(frozen=True)
class WorkbookProtection(Record):
    """Workbook protection, as ``protect_workbook`` takes it; ``password``
    whether one is set."""

    structure: bool = False
    windows: bool = False
    password: bool = False


# -- PowerPoint --------------------------------------------------------------------------------
@dataclass(frozen=True)
class Placeholder(Record):
    """A placeholder on a slide: ``type`` ("title", "body", "obj", "pic"
    ...) and ``idx``."""

    type: str
    idx: int
    name: str | None = None
    element: OoxElement | None = field(default=None, repr=False, compare=False)


@dataclass(frozen=True)
class Shape(Record):
    """Something on a slide: ``kind`` "placeholder", "text_box", "shape",
    "picture", "table", "chart", "group" or "other"; position and size in
    points (None for a placeholder that inherits them from its layout)."""

    kind: str
    name: str | None = None
    text: str | None = None
    placeholder: Placeholder | None = None
    geometry: str | None = None
    x: float | None = None
    y: float | None = None
    width: float | None = None
    height: float | None = None
    element: OoxElement | None = field(default=None, repr=False, compare=False)


@dataclass(frozen=True)
class SlideTableInfo(Record):
    """A table on a slide (``deck.table(slide, i)`` edits it): cell texts,
    merges as "A1:B1" ranges, style and switches as ``add_table`` takes them."""

    name: str | None
    rows: list = field(default_factory=list)
    merges: list = field(default_factory=list)
    style: str | None = None
    style_name: str | None = None
    header: bool = False
    banded: bool = False
    first_column: bool = False


@dataclass(frozen=True)
class SlideSection(Record):
    """A section of slides: its name and slide numbers (1-based)."""

    name: str
    slides: list = field(default_factory=list)


__all__ = [
    "CellFormat",
    "Chart",
    "Comment",
    "ConditionalFormat",
    "ContentControl",
    "ExcelTable",
    "Hyperlink",
    "Note",
    "PageSetup",
    "Picture",
    "Placeholder",
    "Protection",
    "Record",
    "Revision",
    "Shape",
    "SheetOptions",
    "SlideSection",
    "SlideTableInfo",
    "TextBox",
    "Validation",
    "WorkbookProtection",
]
