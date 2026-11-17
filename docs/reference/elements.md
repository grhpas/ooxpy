# Elements

The element class, factories, `lenient()`, plain text and odfpy names.

## `OoxElement`

*Based on* `OdfDomMixin`.

An OOXML element.  Create via factories: `from oox.ns import w; w.P()`.

Also speaks odfpy's DOM dialect (`childNodes`, `appendChild`,
`getElementsByType` ...) -- see `oox._odfdom`.

### `add`

```python
add(*items: OoxElement | str, replace: bool = False) -> OoxElement
```

Add children in schema order; strings are added as text.

`replace=True` swaps out an existing child with the same tag (e.g. a
second `w:pPr`) instead of raising `DuplicateChild`.  It also
applies where the schema allows repeats (`w:trPr` is an unbounded
choice): the first child with that tag is replaced.
Returns `self` so calls can be chained.

### `addElement`

```python
addElement(element: OoxElement, check_grammar: bool = True) -> None
```

odfpy's `addElement`: `add` in schema order; `check_grammar=False` skips
validation, as in odfpy.

### `addText`

```python
addText(text: str) -> None
```

odfpy's `addText`: append text (see `add_text`).

### `add_text`

```python
add_text(text: str) -> OoxElement
```

Append `text` to this element (after its last child, if any); refused for
elements whose schema type holds no text.  Returns the element.

### `first`

```python
first(qname: str) -> OoxElement | None
```

First direct child with the given qname (`"w:pPr"`).

### `getAttribute`

```python
getAttribute(name: str) -> str | None
```

odfpy's `getAttribute`: the attribute's value, or None.

### `get_attr`

```python
get_attr(name: str, default: str | None = None) -> str | None
```

The value of attribute `name` (`"val"`, `"w:val"` or a Clark name), or
`default`.

### `info`

*read-only property*

The schema type of this element (`grammar.TypeInfo`: attributes, children,
order), or None for unknown elements.

### `need`

```python
need(qname: str) -> OoxElement
```

Like `first`, but a missing child raises `PackageError` naming it.

Use where the structure is required: files from other producers can
be well-formed yet lack it, and a clear error beats `'NoneType' has
no attribute ...`.

### `qname`

*read-only property*

`(namespace URI, local name)` -- same shape as odfpy's `Element.qname`.

### `schema_key`

*read-only property*

`TYPE/qname` of the content model in effect, resolved from context.

### `setAttribute`

```python
setAttribute(name: str, value: object) -> None
```

odfpy's `setAttribute`: same as `set_attr`.

### `set_attr`

```python
set_attr(name: str, value: object) -> OoxElement
```

Set by kwarg name (`val`), qname (`w:val`) or Clark name.

`None` removes the attribute.  Values are converted/validated by kind.

### `strip_for`

```python
strip_for(target: str) -> list[str]
```

Make the tree readable by `target` (see `target_findings`):
markup that was only renamed later goes back to its old name (Word
2010's `w:start`/`w:end` -> `w:left`/`w:right`), anything else
is removed.  Returns a message per change.

### `tag_name`

*read-only property*

Prefixed name with canonical prefix, e.g. `"w:p"` (odfpy: `tagName`).

### `target_findings`

```python
target_findings(target: str) -> list[tuple[OoxElement, str | None, str]]
```

Markup `target` (e.g. "Office2007") won't understand, as
`(element, attribute Clark name or None, message)`: the attribute,
or (`None`) the whole element, was added to its namespace later.

### `to_xml`

```python
to_xml(pretty: bool = False) -> str
```

This element and its subtree as an XML string.

### `validate`

```python
validate(
    recursive: bool = True,
    target: str | None = None,
    level: str = 'basic',
) -> list[str]
```

Return a list of problems (empty == valid as far as we model it).

Checks: unknown/missing attributes, attribute values (enums, numbers,
lengths, patterns), allowed children, singleton children, sequence
order, stray text, and mandatory children (those required regardless
of choices).  `level="full"` also matches each element's children
against its whole content model (`oox.contentmodel`): exclusive
choices, required alternatives, occurrence counts, unfinished groups.

`target="Office2007"` (or Office2010, ...) also reports markup that
version would not understand: elements and attributes added later to
namespaces it knows.  Newer *namespaces* (w14, ...) are fine -- they
are declared `mc:Ignorable` -- as is anything inside `extLst` or
`mc:AlternateContent`.

**Details** (from `oox.contentmodel`)

Full content-model checking: does an element's sequence of children match
its schema particle (sequences, choices, `all` groups, min/max occurrences)?

The basic validator checks each child on its own (allowed? singleton? in
order?).  This module checks the whole sequence, which also catches a missing
required child inside a choice, two alternatives of an exclusive choice, too
many repetitions, and a group that is started but not finished.

Each particle tree (generated into `_generated/p_<prefix>.py`) is compiled
once into a Thompson NFA; matching simulates it over the children's qnames.
Children the particle doesn't mention at all are left to the basic validator
(it reports unknown ones; ignorable extension elements are fine), as are
`mc:AlternateContent` blocks.

## `lenient`

```python
lenient() -> Iterator[None]
```

Turn construction-time checks off inside the block: unknown attributes
and children, children out of place and invalid values are all accepted
(`validate()` still reports them).  For repairing messy input, or for
markup Office reads although the schema doesn't allow it.

## `oox.teletype`

Plain text <-> markup, mirroring odfpy's `odf.teletype`.

from oox import teletype
teletype.extractText(paragraph)            # "Name:\tAda"
teletype.addTextToElement(paragraph, "a\tb\nc")

### `addTextToElement`

```python
addTextToElement(element, s: str) -> None
```

Append `s` to a paragraph or run.

`\t` becomes `w:tab` and `\n` becomes `w:br`; runs of spaces are kept
(`xml:space="preserve"` is set automatically).  Given a `w:p`, a new
`w:r` is created; given a `w:r`, content is appended to it.

### `extractText`

```python
extractText(element, text_boxes: bool = True) -> str
```

Visible text of an element subtree.

`w:t` contributes its text, `w:tab`/`w:ptab` a tab, `w:br`/`w:cr`
a newline.  Deleted text (`w:delText`) and field codes (`w:instrText`)
are not `w:t` and are skipped.  Like odfpy, no separator is inserted
between paragraphs.

Markup compatibility is read as Word reads it (`oox.walk`): of each
`mc:AlternateContent` only the branch a reader shows counts, so the
VML copy Word keeps of every text box isn't counted twice.
`text_boxes=False` also leaves out text boxes below `element`
(`w:txbxContent`), as Word's `Paragraph.Range.Text` does.

## `oox.odfcompat`

odfpy compatibility: the names an odfpy program reaches for.

```python
# odfpy                                   # oox
from odf.opendocument import load    ->   from oox.odfcompat import load
from odf import teletype             ->   from oox import teletype
OpenDocumentText()                   ->   OpenDocumentText()          (WordprocessingDocument.new)
OpenDocumentSpreadsheet()            ->   OpenDocumentSpreadsheet()   (SpreadsheetDocument.new)
OpenDocumentPresentation()           ->   OpenDocumentPresentation()  (PresentationDocument.new)
```

Element-level compatibility (`addElement`, `childNodes`,
`getElementsByType`, `attributes=`/`parent=`/`check_grammar=` keywords,
odfpy attribute spellings, element-valued style references) is built into
`OoxElement` and the factories, so it works everywhere -- not only through
this module.  What cannot be made identical, and why, is listed in
docs/ODFPY_COMPAT.md.

### `OpenDocumentPresentation`

```python
OpenDocumentPresentation() -> PresentationDocument
```

New PresentationML (.pptx) deck with master, blank layout and theme.

### `OpenDocumentSpreadsheet`

```python
OpenDocumentSpreadsheet() -> SpreadsheetDocument
```

New SpreadsheetML (.xlsx) workbook.  Unlike odfpy's, it starts with one
sheet ("Sheet1"), because Excel refuses a workbook without sheets.

### `OpenDocumentText`

```python
OpenDocumentText() -> WordprocessingDocument
```

New WordprocessingML (.docx) document.

### `load`

```python
load(odffile: str | os.PathLike | IO[bytes]) -> OoxDocument
```

Open any OOXML package with the matching document class (docx/xlsx/pptx).
