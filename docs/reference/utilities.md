# Units and errors

Length units and the exception hierarchy.

## `oox.units`

Length units.  DrawingML measures in EMU (English Metric Units):
914400 per inch, 360000 per cm, 12700 per point.  WordprocessingML page
geometry uses twips (1/20 pt).  All helpers return `int`.

```python
from oox.units import Inches, Cm, Pt
doc.add_picture("logo.png", width=Cm(4))
```

### `Cm`

```python
Cm(value: float) -> int
```

### `Emu`

```python
Emu(value: float) -> int
```

### `Inches`

```python
Inches(value: float) -> int
```

### `Mm`

```python
Mm(value: float) -> int
```

### `Pt`

```python
Pt(value: float) -> int
```

### `Px`

```python
Px(value: float, dpi: float = 96) -> int
```

### `to_twips`

```python
to_twips(emu: int) -> int
```

EMU -> twips (for w:pgSz, w:pgMar, w:tblW ...).

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
