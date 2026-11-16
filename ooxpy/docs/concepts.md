# Concepts

ooxpy has two layers. The **element layer** gives you every element of the format as
a factory function, checked against the schema. The **document layer** adds
conveniences (`add_table`, `add_chart`, `set_value` …) that build those elements for
you. You can mix them freely: every convenience returns the elements it made, and
every element can be passed where a convenience expects one.

## Documents, parts and elements

An Office file is a zip **package** of **parts**: `word/document.xml`, `word/styles.xml`,
images, charts. Parts point at each other through **relationships**. A document object
holds the package and gives you the main part's root element:

```python
from oox import WordprocessingDocument

doc = WordprocessingDocument.new()
para = doc.add_paragraph("Hello")
print(doc.document.tag_name)            # w:document
print(doc.body.tag_name)                # w:body
print(doc.part_of(para).partname)       # /word/document.xml
print(sorted(doc.package.parts))        # the parts new() created
```

`SpreadsheetDocument.workbook` and `PresentationDocument.presentation` are the
equivalents for Excel and PowerPoint; sheets and slides are elements in their own
parts (`book.sheet("Sales")`, `deck.slides`).

## Element factories

`oox.ns` has one module per namespace prefix, and one factory per element in it. The
factory's name is the element's local name with a capital first letter: `w:p` is
`w.P`, `w:rPr` is `w.RPr`, `a:solidFill` is `a.SolidFill`, `x:c` is `x.C`.

```python
from oox.ns import w

p = w.P(
    w.PPr(w.Jc(val="center")),
    w.R(w.RPr(w.B(), w.Sz(val=28)), w.T("Centred and bold")),
)
doc.body.add(p)
```

- **Children** are positional arguments; **attributes** are keywords, named by their
  local name (`val=`, `ascii=`, `firstLine=`). An attribute in another namespace is
  spelled the same way when the name is unique (`id=` for `r:id`), or with the prefix
  (`w14_paraId=`, or `**{"w14:paraId": …}`).
- **Text** is a string argument: `w.T("x")`. Leading or trailing spaces set
  `xml:space="preserve"` for you.
- **Python values** are converted: `True`/`False` become the schema's spelling, numbers
  become text, dates become ISO dates, bytes become hex or base64 as the type says.
- Each factory has a docstring listing its attributes and children, and the package
  ships type stubs, so an IDE completes attribute names and offers the allowed values.

Elements are [lxml](https://lxml.de/) elements, so the whole lxml API (`find`, `iter`,
`xpath`, `getparent` …) works on them too.

## Schema order

OOXML is strict about the order of children: Word rejects a `w:pPr` that comes after
the runs. `add()` and the factories insert every child where the schema puts it,
whatever order you pass them in:

```python
p = w.P(w.R(w.T("text")), w.PPr(w.Jc(val="right")))      # pPr given last...
# ...but placed first: w:pPr
print(p[0].tag_name)
p.add(w.PPr(w.Jc(val="left")), replace=True)               # replace an existing child
```

## Validation as you build

Mistakes are caught when you make them, with the allowed alternatives in the message:

```python
from oox.errors import InvalidAttributeValue, InvalidChild

try:
    w.Jc(val="middle")
except InvalidAttributeValue as e:
    # w:val='middle' is not one of: left, start, center, right, end, both, ...
    print(e)

try:
    w.P(w.Tbl())
except InvalidChild as e:
    print(str(e)[:40])   # <w:tbl> is not allowed in <w:p>; ...
```

What is checked: attribute names (with "did you mean" hints), values (enumerations,
numbers, the lexical forms such as `"1.5in"` or six hex digits), required attributes,
which children are allowed, and children that may occur only once.

`oox.lenient()` turns these checks off inside a block, for repairing other producers'
files or writing markup Office accepts although the schema doesn't:

```python
import oox

with oox.lenient():
    odd = w.Jc(val="middle")
print(odd.validate())       # still reported, but you could build it
```

## Validating whole documents

`doc.validate()` checks every part and returns `{part name: [problems]}`; an empty dict
means nothing to report. `level="full"` also matches each element's children against
its whole content model (exclusive choices, occurrence counts). `target="Office2007"`
adds markup that Office 2007 wouldn't understand. See
[Validation and Office versions](guide/validation.md).

```python
print(doc.validate())                         # {}
print(doc.validate(level="full", target="Office2007"))
```

## Units

DrawingML measures in **EMU** (914,400 per inch), Word's page and table layout in
**twips** (1,440 per inch), font sizes in half-points. `oox.units` converts to EMU, and
the conveniences take EMU wherever they place something:

```python
from oox.units import Cm, Inches, Mm, Pt, Px, to_twips

print(Inches(1), Cm(2.54), Pt(72), Px(96))    # all 914400
print(to_twips(Inches(1)))                     # 1440
```

Word table widths and section margins accept either twips or EMU (values above
100,000 are taken as EMU), so `Cm(3)` works there too. Excel column widths are in
characters and row heights in points, as in Excel's own dialogs.

## Reading what you (or anyone) wrote

Every convenience has a reader that works on files from any producer:
`doc.tables()`, `doc.comments()`, `book.values(sheet)`, `book.charts(sheet)`,
`deck.shapes(slide)`, `deck.notes(slide)` and so on. Plain text comes from
`oox.teletype.extractText(element)`:

```python
from oox import teletype

print(teletype.extractText(doc.body))
```

## Lossless round trips

Opening a file parses parts only when you touch them. When you save, every part whose
XML you didn't change is written back exactly as it was read, byte for byte, and so
are the relationships and the content types. Saving is deterministic: the same
document always gives the same bytes, which makes generated files easy to diff and
test.

## Namespaces and newer markup

When a part uses markup from a namespace newer than Office 2007's (w14, w15, x14 …),
it is declared in `mc:Ignorable` on save, so older readers skip it instead of failing,
except inside extension lists and `mc:AlternateContent`, which carry their own
compatibility rules. You don't need to manage namespace declarations: they are
collected on the part's root when it's saved.
