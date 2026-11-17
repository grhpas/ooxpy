# Units and errors

Length units and the exception hierarchy.

## `oox.units`

Length units.

**oox's rule: a plain number is points, as in Office's dialogs; anything else
is a `Length`.**  Every size, position, margin, width and indent oox takes
follows it, and every reader reports lengths in points::

```python
from oox.units import Cm, Inches, Pt, Emu
doc.add_picture("logo.png", width=Cm(4))     # or width=113.4 (points)
section.set_margins(left=Inches(1))          # or left=72
deck.add_shape(slide, "rect", x=Emu(914400)) # EMU from another library
```

The helpers return a `Length`: an `int` count of EMU (DrawingML's unit:
914400 per inch, 12700 per point) that remembers it is a length.  Excel column
widths are the exception: they are in characters, as in Excel's dialog.

A plain number too large to be meant as points (a raw EMU or twips value from
older code) raises `ValueError` naming `Emu()`, instead of drawing a
picture a few hundred metres wide.

### `Cm`

```python
Cm(value: float) -> Length
```

### `Emu`

```python
Emu(value: float) -> Length
```

### `Inches`

```python
Inches(value: float) -> Length
```

### `Length`

*class*

A length in EMU that remembers it is a length.

Arithmetic keeps it a `Length`: `Cm(2) + Cm(1)`, `-Cm(1)`,
`abs()`, `Cm(2) * 1.5` and `Cm(3) / 2` are all lengths, and
`sum(lengths)` works.  `Length / Length` is a plain ratio.  Adding a
plain number other than 0 raises `TypeError`, because oox reads a plain
number as points: write `Cm(2) + Pt(6)`.  Anywhere else it behaves as an
`int` count of EMU.

### `Mm`

```python
Mm(value: float) -> Length
```

### `Pt`

```python
Pt(value: float) -> Length
```

### `Px`

```python
Px(value: float, dpi: float = 96) -> Length
```

### `emu`

```python
emu(value: float, name: str = 'length', limit: float | None = None) -> int
```

Points (or a `Length`) -> EMU.  `limit` (points) catches raw EMU or
twips passed by mistake.

### `points`

```python
points(emu_value: float) -> float
```

EMU -> points, as the readers report them (exact enough to write back).

### `to_twips`

```python
to_twips(emu: int) -> int
```

EMU -> twips (for w:pgSz, w:pgMar, w:tblW ...).

### `twips`

```python
twips(value: float, name: str = 'length', limit: float | None = None) -> int
```

Points (or a `Length`) -> twips.

### `twips_to_points`

```python
twips_to_points(value: float) -> float
```

Twips -> points.

## `oox.errors`

Exception hierarchy.  All grammar errors are `OoxValidationError`.

### `DuplicateChild`

*exception class* (OoxValidationError)

The operation would produce markup that violates the schema.

### `InvalidAttribute`

*exception class* (OoxValidationError, AttributeError)

The operation would produce markup that violates the schema.

### `InvalidAttributeValue`

*exception class* (OoxValidationError)

The operation would produce markup that violates the schema.

### `InvalidChild`

*exception class* (OoxValidationError)

The operation would produce markup that violates the schema.

### `MissingAttribute`

*exception class* (OoxValidationError)

The operation would produce markup that violates the schema.

### `OoxError`

*exception class* (Exception)

Base class for all oox errors.

### `OoxValidationError`

*exception class* (OoxError, ValueError)

The operation would produce markup that violates the schema.

### `PackageError`

*exception class* (OoxError)

Malformed or unsupported OPC package.

### `TargetError`

*exception class* (OoxValidationError)

`save(target=...)`: the document uses markup that Office version
can't read.  `findings` maps part names to messages.
