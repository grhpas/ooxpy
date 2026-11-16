# Package and properties

The zip package, parts, limits and document properties.

## `Package`

### `add_part`

```python
add_part(
    partname: str,
    content_type: str,
    element: OoxElement | None = None,
    blob: bytes | None = None,
) -> Part
```

Add a part named `partname` ("/word/foo.xml") with `content_type`, holding an
element tree or raw bytes; relate it with `relate`.

### `open`

```python
open(
    file: str | os.PathLike | IO[bytes],
    limits: Limits = Limits(max_parts=10000, max_part_size=536870912, max_total_size=2147483648, max_ratio=1000),
) -> Package
```

*classmethod*

Open an OPC package.  Any problem with the file raises `PackageError`.

### `relate`

```python
relate(source: Part | None, target: Part, reltype: str) -> str
```

Add a relationship of type `reltype` from `source` (None: the package) to
`target`; returns the new relationship id.

### `related`

```python
related(source: Part | None, reltype: str) -> Part | None
```

The first part that `source` (None: the package) relates to with `reltype`, or
None.

### `save`

```python
save(file: str | os.PathLike | IO[bytes]) -> None
```

Write the package to `file` (a path or a binary file object): deterministic zip,
untouched parts byte for byte.

### `to_bytes`

```python
to_bytes() -> bytes
```

The package as bytes (what `save` writes).

## `Part`

### `element`

*read-only property*

Root element (parsed on first access).

### `is_xml`

*read-only property*

Whether the part holds XML (by content type).

### `rels_name`

*read-only property*

Name of the part's relationships part ("/word/_rels/document.xml.rels").

### `resolve`

```python
resolve(target: str) -> str
```

The part name a relationship `target` points to, relative to this part.

### `to_bytes`

```python
to_bytes(strict: bool = False) -> bytes
```

Serialize; `strict=True` writes Strict namespace URIs.

## `Limits`

```python
Limits(
    max_parts: int = 10000,
    max_part_size: int = 536870912,
    max_total_size: int = 2147483648,
    max_ratio: int = 1000,
) -> None
```

Bounds applied when opening a package (untrusted input).

The defaults are far above real Office files (the largest parts are
usually media) but stop zip bombs and pathological archives.  Use
`Limits.unlimited()` only for trusted files.

Fields: `max_parts`, `max_part_size`, `max_total_size`, `max_ratio`

### `unlimited`

```python
unlimited() -> Limits
```

*classmethod*

Limits that never trigger, for trusted files larger than the defaults allow.

## `CoreProperties`

### `as_dict`

```python
as_dict() -> dict[str, object]
```

Every core property that is set, as `{name: value}`.

## `ExtendedProperties`

`docProps/app.xml`: company, manager, application, template ...

```python
doc.extended_properties.company = "Example Ltd"
```

Counts (pages, words, slides ...) are written by Office when it saves; oox
reads them but doesn't compute them.  The part is created on first write.

### `as_dict`

```python
as_dict() -> dict[str, object]
```

Every extended (`app.xml`) property that is set, as `{name: value}`.

## `CustomProperties`

`docProps/custom.xml` as a mapping (File > Properties > Custom):

```python
doc.custom_properties["Project"] = "Apollo"
doc.custom_properties["Approved"] = True
doc.custom_properties["Budget"] = 1250.5
del doc.custom_properties["Draft"]
```

Values: str, int, float, bool, datetime.  The part is created on first write.

### `as_dict`

```python
as_dict() -> dict[str, object]
```

All custom properties as `{name: value}` (typed: str, int, float, bool,
datetime).

### `get`

```python
get(name: str, default=None)
```

The custom property `name`, or `default`.

### `keys`

```python
keys() -> list[str]
```

Names of the custom properties, in file order.
