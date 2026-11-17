# Documents (common)

Opening, saving, validating and the parts of any document.

## `open_document`

```python
open_document(file: str | os.PathLike | IO[bytes]) -> OoxDocument
```

Open any OOXML package with the matching document class.

## `OoxDocument`

Base class.  Subclasses set `MAIN_CONTENT_TYPES` and implement `new`.

### `open`

```python
open(file: str | os.PathLike | IO[bytes])
```

*classmethod*

Open `file` (a path or a binary file object) as this class;
`oox.open_document` picks the class from the content.

### `save`

```python
save(
    outputfile: str | os.PathLike | IO[bytes],
    addsuffix: bool = False,
    *,
    target: str | None = None,
    strip: bool = False,
) -> None
```

Save the package.  odfpy-compatible: `addsuffix=True` appends the
format's suffix (`.docx`, `.xlsx`, `.pptx`); `"-"` writes to stdout.

`target="Office2007"` (Office2010, ...) refuses to save markup that
version can't read -- e.g. Office 2010 attributes in a file from
LibreOffice, which Excel 2007 refuses to open -- raising
`TargetError` with every finding; `strip=True` removes that markup
instead (`self.stripped` lists what went).

### `validate`

```python
validate(target: str | None = None, level: str = 'basic') -> dict[str, list[str]]
```

Validate every XML part; `{partname: [problems]}` for parts with any.

`target="Office2007"` adds markup that Office version won't understand;
`level="full"` checks whole content models (choices, counts).
Parts that fail to parse are reported, not raised.

### `properties`

*read-only property*

Document properties (title, creator, created, ...): `CoreProperties`.

### `extended_properties`

*read-only property*

`docProps/app.xml` (company, manager, template, ...): `ExtendedProperties`.

### `custom_properties`

*read-only property*

`docProps/custom.xml` as a mapping: `CustomProperties`.

### `conformance`

*property*

"transitional" or "strict"; set it to convert on save.

### `add_picture_image`

```python
add_picture_image(
    image,
    fallback=None,
    source: Part | None = None,
) -> tuple[str, images.ImageInfo, str | None]
```

Store a picture's image(s): `(rid, info, svg_rid)`.

An SVG needs a raster `fallback` (PNG recommended): `rid` is the
fallback, `svg_rid` the SVG, and the size comes from the SVG.

### `add_image`

```python
add_image(image, source: Part | None = None) -> tuple[str, images.ImageInfo]
```

Store an image (path, bytes or file object) and relate it from `source`.

Identical images are stored once; relating the same image twice from
the same part reuses the relationship.  Returns `(rId, ImageInfo)`.

### `add_part`

```python
add_part(
    part_type: str,
    source: Part | None = None,
    *,
    element: OoxElement | None = None,
    blob: bytes | None = None,
    content_type: str | None = None,
    ext: str = 'xml',
    folder: str | None = None,
) -> tuple[Part, str]
```

Create a part described by the SDK part table and relate it.

`part_type` is an SDK part name ("WorksheetPart", "SlidePart", ...);
its content type, relationship type and folder come from the table.
`folder` (absolute, e.g. "/xl/charts") overrides the folder, to put
parts where Office does.  Returns `(part, relationship id)`.

### `related_parts`

```python
related_parts(source: Part, part_type: str) -> list[Part]
```

Parts that `source` relates to with `part_type`'s relationship (an SDK part
name such as "ImagePart" or "DocumentSettingsPart"), in relationship order.

### `relate`

```python
relate(source: Part, target: Part, part_type: str) -> str
```

Add a relationship of `part_type`'s kind between existing parts.

### `part_of`

```python
part_of(element: OoxElement) -> Part
```

The part whose XML tree contains `element`.

### `image_of`

```python
image_of(element) -> ImageData
```

The image an element shows: `(data, content_type, partname,
external)`.  `element` is an `a:blip` or `v:imagedata`, anything
that holds one (`w:drawing`, `pic:pic`, `p:pic`, `w:pict`), a
picture from `pictures()` or a picture bullet's `ListItem.picture`.
A linked image has `data` None and its address in `external`.

### `getElementsByType`

```python
getElementsByType(element: Callable) -> list[OoxElement]
```

odfpy-compatible: search the main part and every part already loaded.

### `addPicture`

```python
addPicture(
    filename: str,
    mediatype: str | None = None,
    content: bytes | None = None,
) -> str
```

odfpy's `addPicture`: store an image (from `filename`, or `content` bytes)
in the package and return its relationship id; `add_picture` places a picture.

### `addPictureFromFile`

```python
addPictureFromFile(filename: str, mediatype: str | None = None) -> str
```

odfpy's `addPictureFromFile`: store the image file and return its relationship
id.

### `addPictureFromString`

```python
addPictureFromString(content: bytes, mediatype: str | None = None) -> str
```

odfpy's `addPictureFromString`: store image bytes and return their relationship
id.

### `MAIN_CONTENT_TYPES`

*attribute* = `{}`
