# Attributions

oox is built on other people's work. This page lists what we take from whom,
in what form, and under which license. Keep it current: **anything new we
derive from, copy, translate or test against gets an entry here in the same
change.**

## Data and code included or derived

| Source | What we use | How it appears in oox | License |
|---|---|---|---|
| [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) (`data/` folder), © .NET Foundation and Contributors | Schema metadata: types, particles, attributes, enums, namespaces, part definitions | Generated code in `src/oox/_generated/` and `src/oox/ns/`; the pinned commit is in every generated file's header and in `NOTICE` | MIT |
| [odfpy](https://github.com/eea/odfpy) | The API design oox follows (factory per element, `addElement`, `teletype`, `load`, `save(name, addsuffix)`, DOM methods); four example scripts (helloworld, loadsave, manualpagebreak, passwd-as-odt) translated as tests, with the originals quoted in comments | `src/oox/_odfdom.py`, `src/oox/odfcompat.py`, `src/oox/teletype.py`, `tests/test_odfpy_examples.py` | GPL-2.0-or-later OR Apache-2.0 (we use it under Apache-2.0) |

## Specifications

| Specification | Used for |
|---|---|
| [ECMA-376 Office Open XML](https://ecma-international.org/publications-and-standards/standards/ecma-376/), Parts 1–4 (also ISO/IEC 29500) | The vocabulary itself; Strict vs Transitional namespaces (`oox/strict.py`); rules from the prose such as row and cell order (`oox/semantics.py`); OPC packaging and the no-DTD rule (Part 2); Markup Compatibility, `mc:Ignorable` (Part 3); checking SDK deviations (`tools/gen/overrides.py`) |
| ECMA-376 Part 1: §17.6 sections, §17.10 headers and footers, §17.11 footnotes and endnotes, §17.13 annotations (comments, bookmarks), §17.16 fields and hyperlinks (incl. the TOC field switches); §17.4 tables (`gridSpan`, `vMerge`, `tblHeader`); §17.13.5 revisions; §20.4 DrawingML WordprocessingML drawing (`wp:anchor`, wrapping) | `oox/sections.py`, `oox/notes.py`, `oox/fields.py`, `oox/tables.py`, `oox/revisions.py`, `oox/floating.py` |
| ECMA-376 Part 1: §18.8 styles (incl. the built-in number formats of §18.8.30), §18.3.1.13 column width formula, §18.3.1 worksheet (panes, merges, autofilter, data validation), §18.2.5 defined names | `oox/xlstyles.py`, `oox/xlsheet.py` |
| ECMA-376 Part 1: §18.3.1.18 conditional formatting, §18.5 tables, §18.7 comments, §18.3.1.47 hyperlinks, §18.3.1.62–63 page setup, §18.4 shared strings, sheet protection; Part 4: VML (legacy note boxes) and the legacy password hash | `oox/xlcond.py`, `oox/xltables.py`, `oox/xlnotes.py`, `oox/xlprint.py` |
| ECMA-376 Part 1: §21.2 DrawingML Charts (`c:chartSpace`, chart types, axes, series, caches, `externalData`), §20.5 SpreadsheetML Drawing (`xdr:twoCellAnchor`, `xdr:graphicFrame`), §19.3.1.21 PresentationML graphic frames; Part 2: embedded package parts | `oox/charts.py`, `oox/xlcharts.py`, Word and PowerPoint `add_chart` |
| ECMA-376 Part 1: §19.3.1 PresentationML slides, layouts (`ST_SlideLayoutType`), notes master and notes slides, placeholders (`p:ph`); §21.1.3 DrawingML tables (`a:tbl`, `gridSpan`, `rowSpan`, `hMerge`, `vMerge`, `tableStyleId`); §20.1.9 preset geometries; §20.1.8.55 `a:srcRect` cropping | `oox/presentation.py` (standard layouts), `oox/slides.py`, `oox/slidetable.py` |
| ECMA-376 Part 1: DrawingML `a:hlinkClick` (on `p:cNvPr` and in `a:rPr`), the PresentationML actions `ppaction://hlinksldjump` and `ppaction://hlinkshowjump?jump=`; MS-PPTX: `p14:sectionLst` (presentation extension `{521415D9-36F7-43E2-AB2F-B90AF26B5E84}`) | `oox/slidelinks.py`, `oox/sections_pml.py` |
| ECMA-376 Part 1: §17.15.1.29 `w:documentProtection`, §18.2.29 `x:workbookProtection`; Part 4: §14.8.1 the legacy Word password key (initial codes and encryption matrix, worked example "Example" → `0x64CEED7E`) and the CryptoAPI hash attributes; Part 3: Markup Compatibility (`mc:AlternateContent` choice/fallback); MS-DOCX: `wps:wsp` (the Word 2010 shape namespace); Part 4 VML text boxes (`v:textbox`, shape type 202). A password hash written by Word 2019 serves as a test vector only | `oox/protection.py`, `oox/textboxes.py` |
| ECMA-376 Part 1: §21.2 several chart groups in one plot area with a second axis pair (`c:crosses` "max" for the secondary value axis); SpreadsheetML chart sheets (`x:chartsheet`) and `xdr:absoluteAnchor` | `oox/charts.py` (combinations, secondary axes), `oox/xlcharts.py` (chart sheets) |
| ISO/IEC 29500-4:2016 Transitional XML schemas, the MCE and OPC schemas, and Microsoft's Word extension schemas (MS-DOCX: `wml-2010.xsd` … `wml-sdtdatahash-2020.xsd`) | Input to the XSD frontend and IR diff (`tools/gen/xsd_frontend.py`, `tools/gen/irdiff.py`); not in the repository and not shipped; `docs/reports/xsd-diff.md` lists element, attribute and value names only |
| ISO/IEC 29500-1 §17.8.3.2 `charset` (`characterSet` is `ST_String`: IANA character set names) | `lexical.OPEN_ENUMS` |
| [Microsoft Open Specifications](https://learn.microsoft.com/en-us/openspecs/office_standards/) (MS-DOCX, MS-XLSX, MS-PPTX, MS-ODRAWXML, MS-OI29500) | Office extension namespaces (w14, x14, p14, …) and implementer notes |
| [PNG](https://www.w3.org/TR/png/), [JPEG/JFIF](https://www.w3.org/Graphics/JPEG/jfif3.pdf), [GIF89a](https://www.w3.org/Graphics/GIF/spec-gif89a.txt), BMP (Windows bitmap headers), TIFF 6.0, [SVG 1.1](https://www.w3.org/TR/SVG11/) | Image header parsing in `src/oox/images.py` (written from the specs) |
| [MS-EMF](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/), [MS-WMF](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/) | EMF/WMF header parsing (`images.py`) and the hand-built EMF/WMF test pictures (`tools/officecheck/assets.py`) |
| ECMA-376 Part 1: §17.5 structured document tags (content controls), §17.16 fields (SEQ, REF, PAGEREF, DOCPROPERTY, date pictures), §22.2–22.3 extended and custom properties; MS-ODRAWXML (`asvg:svgBlip`) | `oox/controls.py`, `oox/fieldcalc.py`, `oox/properties.py`, `oox/drawing.py` |

## Libraries

| Library | Role | License |
|---|---|---|
| [lxml](https://lxml.de/) | Runtime dependency: XML parsing, serialization, the element class base | BSD-3-Clause |
| [python-docx-ng](https://github.com/toxicphreAK/python-docx-ng) / [python-docx](https://github.com/python-openxml/python-docx), © Steve Canny and contributors | Design reference: the OPC / oxml / API layering and the declarative element approach. python-docx is the independent .docx reader in tests, and its Word-authored template is test input | MIT |
| [python-pptx](https://github.com/scanny/python-pptx) | Independent .pptx reader in tests; its PowerPoint-authored template is test and corpus input; writes interop test files (`tools/writers.py`, corpus `ppt-11`) (not shipped) | MIT |
| [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) | Independent .xlsx reader in tests; writes interop test files (`tools/writers.py`, corpus `excel-18`) | MIT |
| [XlsxWriter](https://github.com/jmcnamara/XlsxWriter), © John McNamara | Writes interop test files (`tools/writers.py`), including combination charts, secondary axes and chart sheets | BSD-2-Clause |
| [Pillow](https://python-pillow.org/) | Draws the example images and the embedded test images (`tools/officecheck/assets.py`); not a runtime dependency | MIT-CMU |
| [pytest](https://pytest.org/) | Test runner | MIT |
| [MkDocs](https://www.mkdocs.org/), [Material for MkDocs](https://squidfunk.github.io/mkdocs-material/), [PyMdown Extensions](https://facelessuser.github.io/pymdown-extensions/) | Build the documentation site from `docs/` (the `docs` extra; not a runtime dependency) | BSD-2-Clause, MIT, MIT |
| [setuptools](https://setuptools.pypa.io/), [build](https://build.pypa.io/), [twine](https://twine.readthedocs.io/), [pypa/gh-action-pypi-publish](https://github.com/pypa/gh-action-pypi-publish) | Building, checking and publishing the distributions (`release.yml`) | MIT, MIT, Apache-2.0, BSD-3-Clause |
| [pywin32](https://github.com/mhammond/pywin32) | Drives Office through COM in `tools/officecheck/check_office.py` (Windows only, not a dependency of oox) | PSF |

## Tools used for verification

| Tool | Role |
|---|---|
| [LibreOffice](https://www.libreoffice.org/) | Rendering checks and `check_libreoffice.py` probes; produces third-party test files |
| Microsoft Office 2007 and 2016 | Manual and scripted compatibility checks (`tools/officecheck/`) |
| [Word VBA object model](https://learn.microsoft.com/en-us/office/vba/api/overview/word/object-model): `Sections`, `HeadersFooters`, `Hyperlinks`, `Footnotes`, `Comments`, `TablesOfContents`, `Range.Fields`, `Tables`/`Cell`/`HeadingFormat`, `Shapes`/`WrapFormat` (`WdWrapType`), `Revisions` (`WdRevisionType`, `AcceptAll`, `RejectAll`), `Font`, `ParagraphFormat`, `Fields.Update`, `ContentControls` / `SelectContentControlsByTag` (`WdContentControlType`), `CustomDocumentProperties`, `ProtectionType` (`WdProtectionType`) / `Unprotect`, `Shapes(i).TextFrame.TextRange`; Excel `Range.Text`/`NumberFormat`/`Font`/`Interior`/`HorizontalAlignment`/`MergeArea`/`Validation`, `Window.FreezePanes`, `Names`, `FormatConditions` (`XlFormatConditionType`), `ListObjects`, `Comments`, `Hyperlinks`, `PageSetup`, `Tab.Color`, `Visible`, `ProtectContents`, `Workbook.ProtectStructure` / `ProtectWindows` / `Unprotect`; the shared [`Chart` object](https://learn.microsoft.com/en-us/office/vba/api/excel.chart(object)) (`ChartType` / [`XlChartType`](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype), `ChartTitle`, `Legend.Position`, `Axes`/`AxisTitle`, `SeriesCollection`, `Values`/`XValues`, `DataLabels`, `Export`), `ChartObjects`, Word `InlineShapes(i).Chart`, PowerPoint `Shape.HasChart` and `Slide.Export`; PowerPoint `SlideMaster.CustomLayouts`, `Slide.CustomLayout`, `Shape.HasTable`/`Table.Cell`, `NotesPage`, `PlaceholderFormat` (`PpPlaceholderType`, `ContainedType`), `AutoShapeType` ([`MsoAutoShapeType`](https://learn.microsoft.com/en-us/office/vba/api/office.msoautoshapetype)) | The Office probes in `check_office.py` |
| [Office VBA reference: PowerPoint `Application.DisplayAlerts`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.displayalerts) | Why `check_office.py` turns PowerPoint alerts on (the default `ppAlertsNone` hides prompts) |
| [IANA character set registry](https://www.iana.org/assignments/character-sets) | Why `windows-1252` is a valid `w:characterSet` |
| [ruff](https://docs.astral.sh/ruff/), [pyright](https://github.com/microsoft/pyright), [GitHub Actions](https://docs.github.com/actions) | Linting, type checking, CI |

## Prior art consulted (no code taken)

| Project | What we learned |
|---|---|
| [ooxmlsdk](https://github.com/KaiserY/ooxmlsdk) (Rust) | Generating from SDK metadata. Their XSD-vs-SDK comparison ([#25](https://github.com/KaiserY/ooxmlsdk/issues/25), [#26](https://github.com/KaiserY/ooxmlsdk/pull/26)) documented gaps (e.g. Visio) and deviations that shaped ADR-001 |
| [pterror/ooxml](https://github.com/pterror/ooxml) (Rust) | Generating from the ECMA-376 RELAX NG schemas: an alternative frontend for Phase 3 |
| [openxml-audit](https://pypi.org/project/openxml-audit/) | A Python port of the SDK validator; planned as an external oracle |
| [docx4j](https://www.docx4java.org/) (Java) | Classes generated from the schemas, at full coverage |
| [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) | Cross-check for Excel's legacy password hash (`secret` → `DAA7`) and the VML layout of comment boxes; oox's code is written from the specifications |
| [python-docx-template (docxtpl)](https://github.com/elapouya/python-docx-template) | The run-splitting problem for placeholders; oox solves it by matching on joined paragraph text instead of templating the XML |
| Office's default theme | The colour and font values in `drawing.default_theme()` (Office 2013+ defaults), also the default chart palette in `charts.PALETTE` |
| Office's default presentation template | Placeholder positions and sizes of the eleven standard layouts (`presentation._default_layouts`), the notes master geometry, and the GUIDs of the built-in table styles (`slides.TABLE_STYLES`) |
