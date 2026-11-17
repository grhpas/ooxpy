# API reference

Generated from the docstrings by `python -m tools.apidocs`; it covers exactly the public API described in [API stability](../STABILITY.md). Element factories (`oox.ns.w.P`, `oox.ns.x.C` …) follow the schemas: see [Concepts](../concepts.md).

- [Documents (common)](documents.md): Opening, saving, validating and the parts of any document.
- [Word](word.md): `WordprocessingDocument` and the objects its methods return.
- [Excel](excel.md): `SpreadsheetDocument`, cell values and the stylesheet.
- [PowerPoint](powerpoint.md): `PresentationDocument` and `SlideTable`.
- [Charts](charts.md): Chart data and the builder and reader the document methods use.
- [Elements](elements.md): The element class, factories, `lenient()`, plain text and odfpy names.
- [Package and properties](package.md): The zip package, parts, limits and document properties.
- [Records](records.md): What the readers return: records with named fields, and the formatting values the writers take.
- [Units and errors](utilities.md): Length units and the exception hierarchy.
