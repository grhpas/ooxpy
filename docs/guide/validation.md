# Validation and Office versions

Office is unforgiving: a value it doesn't know or an element out of place gives "Word
found unreadable content" or a file it refuses to open. ooxpy checks markup at three
moments: when you build an element, when you ask for a report, and when you save for a
particular Office version.

## While building

Factories and `add()` check attribute names and values, required attributes, allowed
children and how often a child may occur, and put children in schema order. Mistakes
raise an `OoxValidationError` subclass (`InvalidAttribute`, `InvalidAttributeValue`,
`MissingAttribute`, `InvalidChild`, `DuplicateChild`) with the alternatives in the
message. `oox.lenient()` switches this off for a block (see [Concepts](../concepts.md)).

## Reports

```python
from oox import WordprocessingDocument
from oox.ns import a, p, w

doc = WordprocessingDocument.new()
para = doc.add_paragraph("Indented")
para.add(w.PPr(w.Ind(start=720)))

print(doc.validate())                     # {}: valid ISO/IEC 29500 Transitional
```

`doc.validate()` returns `{part name: [problems]}` for every XML part; `{}` means
nothing to report. `element.validate()` checks one subtree and returns a list. It checks:

- attributes: unknown names, values (enumerations, numbers, booleans, the lexical
  forms the schema allows, such as `"1.5in"` or six hex digits), required attributes;
- children: allowed elements, order, duplicates, mandatory children, stray text;
- rules from the specification's prose that the schema can't express, where Office
  enforces them (rows and cells in ascending order in Excel).

`level="full"` also matches each element's children against its complete content
model, which catches exclusive choices and occurrence counts the basic level allows:

```python
shape_props = p.SpPr(a.NoFill(), a.SolidFill(a.SrgbClr(val="FF0000")))
print(shape_props.validate())                       # []
print(shape_props.validate(level="full")[0][:60])
# <p:spPr>: <a:solidFill> is not allowed here ...
```

A shape has one fill: `noFill` *or* `solidFill`. The full level is slower and is meant
for tests and CI rather than every save.

## Older Office versions

Each element and attribute in the grammar knows which Office version introduced it.
`target=` reports markup an older version won't understand:

```python
print(doc.validate(target="Office2007"))
# {'/word/document.xml':
#   ['<w:ind>: attribute w:start needs Office2010 (not understood by Office2007)']}
```

Markup that the older version safely skips isn't reported: newer namespaces declared
in `mc:Ignorable`, extension lists, `mc:AlternateContent` blocks. What remains is markup
it would choke on, such as Word 2010's `w:start` (Word 2007 knows only `w:left`) or
LibreOffice's `x:workbookPr/@dateCompatibility`, which Excel 2007 refuses.

Saving for a target refuses such markup, or removes it:

```python
from oox import TargetError

try:
    doc.save("for-2007.docx", target="Office2007")
except TargetError as e:
    print(str(e)[:50])                   # 1 item(s) Office2007 can't read, ...

doc.save("for-2007.docx", target="Office2007", strip=True)
print(doc.stripped)                      # what was changed, per part
```

`strip=True` changes the document in memory: unknown attributes and elements are
removed, and Word 2010's bidi-neutral names (`w:start`, `w:end`) are renamed to the
2007 ones instead of being dropped. The command line has the same as `oox downgrade`.
Targets are `Office2007`, `Office2010`, `Office2013`, `Office2016`, `Office2019`,
`Office2021` and `Microsoft365`.

## Strict conformance

ISO/IEC 29500 Strict uses other namespaces and a few other value forms. ooxpy reads
Strict files into the usual Transitional form, so the same code works on both, and writes
them back as Strict:

```python
strict = WordprocessingDocument.new()
strict.conformance = "strict"
strict.add_paragraph("Written as ISO/IEC 29500 Strict")
strict.save("strict.docx")

from oox import open_document

print(open_document("strict.docx").conformance)        # strict
```

Namespaces, relationship types, DrawingML percentages (`"50%"` ↔ 50000) and on/off
values are translated. Microsoft documents that Office 2010 reads Strict files, Office
2013 and later also write them, and Office 2007 SP2 needs a separate downloadable
filter. oox's Strict output has so far been checked with LibreOffice only; for Office
2007, save as Transitional (the default).

## How strict is Office, really?

The schema and Office don't always agree. The Office check corpus
(`tools/officecheck/`) records how Office 2007 and 2016 react to specific deviations,
and oox follows the evidence: values Office accepts although the Open XML SDK's data
rejects them (`w:jc="start"` in tables, `"true"` on some on/off elements) are accepted;
out-of-order paragraph properties and unknown elements are tolerated by Office but
still reported, because other readers may not tolerate them; invalid values, rows out
of order and missing required elements make Office refuse the file, so oox treats them
as errors. See [Office compatibility](../office.md).
