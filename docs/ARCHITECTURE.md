# oox: Architecture & Decision Record

**What this is:** a schema-driven, odfpy-style Python library for Office Open XML
(.docx, .xlsx, .pptx). odfpy gives you one factory per ODF element plus a grammar
table to validate against. oox does the same for OOXML: one factory per element,
generated from schema metadata, with no hand-written element classes.

```python
from oox import WordprocessingDocument
from oox.ns import w

doc = WordprocessingDocument.new()
doc.body.add(w.P(w.PPr(w.Jc(val="center")), w.R(w.RPr(w.B()), w.T("Hello"))))
doc.save("hello.docx")
```

## Goals and non-goals

**Goals**
- Cover the *whole* vocabulary, including Microsoft's extension namespaces (w14, w15, x14, p14, …), not a curated subset.
- Make invalid markup hard to produce: validate on construction, keep children in schema order, report problems clearly.
- Round-trip losslessly: anything not touched is written back unchanged.
- One package/element core shared by all three formats.

**Non-goals (for now)**
- A python-docx-style high-level object model (`paragraph.style = …`). That can live on top later (Phase 5).
- Rendering, layout, or format conversion.

---

## ADR-001: Source of truth for the grammar

**Status:** accepted (prototype), to be revisited at Phase 3.

### Options considered

| Source | Coverage | Effort | Risk |
|---|---|---|---|
| **Open XML SDK metadata** (`dotnet/Open-XML-SDK/data/*.json`, MIT) | ECMA-376 Transitional plus about 150 Microsoft extension namespaces. Has per-element Office version tags, enum facets, lexical validators and part metadata. | Low: the data is already compiled (groups inlined, derivation resolved). | Not a public contract. Its README says the files are produced by a backend processor. No Strict and no Visio. Some known deviations from the XSDs. |
| **ECMA-376 XSD** (Transitional/Strict) | The authoritative standard. No Microsoft extensions. | High: needs an XSD compiler (groups, extension/restriction, substitution, `ref`). | Low: stable, published. |
| **ECMA-376 RELAX NG** | Same as XSD. This is the route odfpy (ODF RNG) and pterror/ooxml (Rust) took. | Medium. | Same gap: no extensions. |
| **MS-DOCX/MS-XLSX/MS-PPTX extension XSDs** | Only the extensions. | High: scattered across Open Specifications documents. | Medium. |
| **Hand-written classes** (python-docx's xmlchemy) | Only what someone has needed so far. | Very high to reach full coverage. | Doesn't scale to 4,000+ types. |

### Decision
Use the **SDK metadata as the primary input, through our own intermediate
representation (IR)**: `tools/gen/ir.py` is the only thing the emitter sees.
Pin the SDK commit (it's embedded in every generated file's header).

Real Word files nearly always contain extension markup (`w14:paraId`,
`mc:AlternateContent`, …), and only the SDK data covers it without a lot of
manual work. That decides it. The data also gives friendly class names, which
we use to disambiguate the 191 qnames that have several content models. Its
risks are handled by keeping it behind the IR. The Rust port `ooxmlsdk` is
heading the same way: it is adding an XSD-derived generator and diffing it
against the SDK JSON (KaiserY/ooxmlsdk#25/#26). Their findings: the SDK backend
doesn't cover every namespace (e.g. Visio), and it deviates from the XSD in a
few places (some per-parent extension-list rewrites; a few attributes whose
required-ness differs).

### Consequences
- Deviations found so far are corrected in `tools/gen/overrides.py`, each with a reason. The first: `x:CT_Font` and `x:CT_RPrElt` are unordered choices in ECMA-376 but fixed sequences in the SDK data, which wrongly flagged openpyxl's valid output.
- Phase 3 adds an **XSD frontend** into the same IR. It serves two purposes: (a) an oracle to diff against the SDK data, and (b) a way to cover Strict and namespaces the SDK lacks. Part (a) runs since 2026-11-15 (`tools/gen/xsd_frontend.py`, `tools/gen/irdiff.py`, `docs/XSD_DIFF.md`); deviations are relaxed only with Office evidence, at run time.
- .NET-specific heuristics in the data (`ClassName`, `BaseClass`, `CompositeType`) are used only as naming hints, never as schema facts.
- Refinement of the original idea: we do **not** generate a class per type the way the .NET SDK does. We generate *data* (grammar tables) plus thin factory functions over one runtime element class. That is odfpy's model, and it avoids importing 4,000 classes.

---

## ADR-002: odfpy names are permanent

**Status:** accepted 2026-11-14.

**Context.** The odfpy method names (`addElement`, `childNodes`,
`getElementsByType`, `setAttrNS`, `teletype.extractText`, `save(name, True)`,
`OpenDocumentText()` …) are camelCase, which isn't PEP 8. We had to decide
whether they are permanent or a migration aid that gets deprecated later.

**Decision.** Keep them permanently, as a supported, tested surface.
1. Being "odfpy for OOXML" is the project's identity. Ports are mechanical only while the names stay.
2. They cost almost nothing. They live in one mixin (`_odfdom.py`) plus `odfcompat.py`, and every alias delegates to the snake_case implementation. Nothing is duplicated.
3. Removing them later would break exactly the users we're aiming at. Deprecation warnings would only add noise for no gain.

**Rules that keep this cheap:**
- snake_case is the primary documented API (`add`, `set_attr`, `tag_name` …). The camelCase names are documented on the compatibility page only.
- Aliases never get behavior of their own. They delegate, except where odfpy's semantics differ on purpose: `appendChild` and `insertBefore` are raw DOM operations, while `addElement` checks the grammar.
- An odfpy name is added only when odfpy has it, with odfpy's signature.
- No `DeprecationWarning`s. The 1.0 API freeze includes the compatibility surface.
- Translated odfpy example scripts in `tests/test_odfpy_examples.py` guard it.

## Runtime design

### 1. Content-model keys and context resolution
An element's content model is identified by a **key**, `TYPE/qname`
(`w:CT_RPr/w:rPr`). In the pinned data, 191 of the 3,702 qnames have more than
one key. `w:rPr` has seven: inside `w:r` it is `CT_RPr`, inside `w:pPr` it is
`CT_ParaRPr`, inside a style it is `CT_RPrStyle`, and so on. odfpy's grammar
assumes one model per qname. That assumption does not hold for OOXML.

oox resolves the key **from the parent's content model**, as XSD does:
- Unique qname: resolved immediately.
- Ambiguous and attached: resolved through the parent chain (cached per `(parent key, qname)`).
- Ambiguous and detached: validated against the *union* of models. Where the models disagree on an attribute's type, enum value sets are merged (`w:jc/@w:val` accepts ST_Jc ∪ ST_JcTable); other conflicts accept the raw value (`x:c/@t` is a boolean in the calc chain, an enum in worksheets). Once the parent chain reaches an unambiguous element, the subtree is reordered and re-validated precisely. Some chains go several levels deep: `x:sheetData → x:row → x:c` resolves only at `x:worksheet`.
- SDK aliases (`w.ParagraphMarkRunProperties`) bind a specific model explicitly.

Nothing is stored on the element. lxml proxies aren't stable holders of Python
state, and the XML stays clean.

### 2. Schema-ordered insertion
Word rejects `<w:pPr>` children that are out of sequence order ("unreadable
content"). odfpy never had to deal with this, because ODF is mostly order-free.
The generator computes a **rank** per child from the particle tree:
- Sequences extend the rank.
- Choices share a rank.
- Inside a *repeating* compositor (`(a,b)*`, unbounded choices) everything shares one rank, so insertion order is preserved instead of being wrongly regrouped.

`add()` inserts at the right position. It scans from the last child (lxml's
`len()` and indexing walk the whole child list), so appending in order and
inserting just before a trailing `w:sectPr` cost one or two steps.
`add(child, replace=True)` swaps out an existing child with the same tag, also
where the schema allows repeats (`w:trPr` is an unbounded choice).

### 3. lxml-backed elements, not an own tree
odfpy has its own minidom-like tree, which is slow and memory-heavy on large
files. oox registers a single `OoxElement(lxml.etree.ElementBase)` for every
tag. That gives C-speed parse and serialize, XPath, exact preservation of
unknown markup, and safe parser defaults (`resolve_entities=False`,
`no_network=True`). python-docx's success with the same pattern is evidence it
works.

### 4. Validation policy
- **Construction / `add()`:** unknown attributes (with "did you mean" hints), enum values, value kinds, required attributes, allowed children, singletons, text in non-text elements. Raises `OoxValidationError` subclasses.
- **`with oox.lenient():`** turns those checks off, for repairing messy input.
- **`element.validate()`:** a non-raising report over a subtree (attributes, children, order, duplicates, stray text). Loading never validates.
- **Mandatory children:** a child is reported missing when the schema requires it no matter which choices are made (e.g. `w:tbl` without `w:tblGrid`, a theme font scheme without `a:ea`). The generator marks a child mandatory when it has minOccurs ≥ 1 and no optional or choice ancestor.
- **Value shapes (lexical):** the SDK's String/Number/Enum validators are compiled by the generator into per-attribute constraints (`oox/lexical.py`). Unions are supported ("an integer >= 0 or a measure like 1.5in"), and hexBinary lengths count bytes. Where rules differ by Office version, the newest rules apply. Checked on construction and in `validate()`.
- **Semantic rules** (`oox/semantics.py`): constraints from the specification's prose that the XSD can't express, e.g. Excel rows and cells in ascending order. They're added when the Office check shows Office enforces them.
- **Version targeting:** `validate(target="Office2007")` reports elements and attributes that a given Office version won't understand, using the per-element and per-attribute versions from the SDK data. Content that version safely ignores is skipped: newer namespaces (declared `mc:Ignorable`), extension lists, `mc:AlternateContent`. `save(target=...)` refuses such markup (`TargetError`, findings per part); `strip=True` removes it, renaming Word 2010's bidi-neutral names (`w:start`/`w:end`) back to `w:left`/`w:right` instead of dropping them.
- **Full content models** (`validate(level="full")`, `oox/contentmodel.py`): the generator also emits each type's particle tree (`_generated/p_<prefix>.py`); at run time it's compiled once into a Thompson NFA (repeat bounds unrolled up to 8, larger ones treated as unbounded; `xsd:all` approximated as a repeated choice) and simulated over the children's qnames. This catches exclusive choices, required alternatives, counts and unfinished groups. Children the particle doesn't mention are left to the basic checks, as are `mc:AlternateContent` blocks. Zero findings beyond the basic level on the corpus, LibreOffice re-saves and the python-docx/python-pptx templates.
- **Open enumerations** (`lexical.OPEN_ENUMS`): where the SDK closes a list that ECMA-376 types as `ST_String` with examples (`w:charset/@w:characterSet`, §17.8.3.2), any value is accepted. Applied at run time so the generated code stays identical to the pinned SDK's output; the Phase 3 XSD frontend will correct the IR.

### 5. Document classes
`OoxDocument` (in `document.py`) owns the package and main part, `save`, media, document properties (`properties.py`, created on first write), `validate` and odfpy's document-level methods. The three format classes only add a skeleton (`new()`) and conveniences:

| | Skeleton built by `new()` | Conveniences |
|---|---|---|
| `WordprocessingDocument` | `word/document.xml`, plus by default styles, settings, font table, theme and a section (`word_defaults.py`; `minimal=True` for the bare package) | `body`/`text`, `styles`, `style_id`, `add_heading`, `add_paragraph`, `add_page_break`, `new_list` / `list_labels` (`numbering.py`), `add_picture` (inline), `sections` / `add_section` (`sections.py`), hyperlinks, bookmarks, notes, comments (`notes.py`), `add_toc` (`fields.py`), `replace_text` (`template.py`), `add_table` / `tables` (`tables.py`), `add_floating_picture` / `pictures` (`floating.py`), `suggest_*`, `revisions`, `accept_all` / `reject_all` (`revisions.py`), `add_run` / `format_paragraph` / `add_style` / `effective_format` (`formatting.py`), `add_caption` / `add_cross_reference` / `add_field` / `fields` / `update_fields` (`fieldcalc.py`), content controls (`controls.py`), `add_chart` / `charts` (inline, with an embedded workbook), `add_text_box` / `text_boxes` / `set_text_box_text` (`textboxes.py`), `protect` / `protection` / `check_protection_password` / `unprotect` (`protection.py`) |
| `SpreadsheetDocument` | `xl/workbook.xml` plus one worksheet (Excel rejects a workbook with no sheets) | `add_sheet`, `sheet`, `append_row`, `value`/`values`/`cells` (reading), `set_value` (editing), `cell`/`row`/`Formula` helpers, `add_picture` (anchored in a drawing part), `cell_style` / `style_range` / `cell_format` / `display_text` (`xlstyles.py`), widths, heights, merges, frozen panes, autofilter, names, validation (`xlsheet.py`), conditional formats (`xlcond.py`), tables (`xltables.py`), comments and hyperlinks (`xlnotes.py`), print setup, sheet options, sheet protection (`xlprint.py`), `protect_workbook` / `workbook_protection` (`protection.py`), `add_chart` / `charts` from cell ranges and `add_chart_sheet` (`xlcharts.py`), `tab_names` / `sheet_names` (worksheets) / `chart_sheet_names`, `share_strings` |
| `PresentationDocument` | presentation → slide master (title/body placeholders, text styles) → Office's eleven standard layouts (`layouts="basic"`: Title Slide, Title and Content, Blank); master → theme; presProps, viewProps, tableStyles; a notes master on first `set_notes` | `add_slide(layout)` (copies placeholders), `layout(name)`, `placeholders`, `set_placeholder_text`, `add_layout`, `add_textbox`, `add_picture`, `add_chart` / `charts` (graphic frame, embedded workbook), `slide_layout`, `add_shape`, `add_table` / `tables`, `set_notes` / `notes`, `set_placeholder_picture`, `shapes` (`slides.py`), `table` → `SlideTable` for editing (`slidetable.py`), `set_link` / `links` (`slidelinks.py`), `set_sections` / `sections` (`sections_pml.py`) |

Word conveniences follow two rules. **Styles are added on first use**:
`ensure_style("FootnoteText")` adds Word's definition only if the document
lacks it (`word_defaults.extra_style`), so a document gets exactly the styles
it uses and existing corpus files stay byte-stable. **Every writer has a
reader** that also works on Word and LibreOffice files (`list_labels`,
`hyperlinks`, `footnotes`, `comments`, `Section.shown`), and the LibreOffice
checker uses those readers as its probes.

Rules the format needs but a single `add()` can't enforce are applied on
save by `_prepare_save()`: in Word, a body or table cell that ends in a table
gets a closing paragraph. It only runs on a parsed main part, so untouched
files stay byte-identical.

**Field results are computed, not left to Word.** Word only recomputes fields
on F9 or printing, so a file's cached results are what readers see.
`fieldcalc.update_fields` recomputes everything that doesn't need page layout,
and every feature that adds or moves a field (captions, cross-references,
property fields) calls it, so files are always consistent. The Office check
compares the results with Word's own recompute.

Excel styles are **shared entries**: `cellXfs` items combine indexes into
fonts, fills, borders and number formats, and each list is de-duplicated by a
structural key (namespace URIs, normalised booleans, default ids dropped), so
an entry oox builds equals the same entry in an Excel-authored file and is
reused. Dates are numbers with a date format; reading consults the cell's xf.

**Legacy VML is edited as text.** Excel shows cell comments through a VML
drawing, and VML written by Excel is often not well-formed XML (unclosed
`<br>`). oox never parses it: adding a comment removes that cell's
`<v:shape>` (if any) and inserts a new one before `</xml>`, leaving
everything else byte-for-byte as it was.

All three document classes have `properties` (core), `extended_properties`
(`app.xml`) and `custom_properties` (`custom.xml`); each part is created on
first write. Pictures go through `add_picture_image`, which pairs an SVG with
its raster fallback (`a:blip` plus an `asvg:svgBlip` extension).

Table coordinates are **grid positions**, not cell indexes, so they don't
shift when cells merge: `Table._grid()` maps each position to the `w:tc` shown
there, following `w:gridSpan`, `w:vMerge` and `w:gridBefore`.

`add_part(part_type, source)` creates any part from the SDK's part table. It
fills in the content type, the relationship type, the folder relative to the
parent, and numbering where a part can occur more than once. SDK paths such as
`../media` assume a nested parent (`xl/drawings/`). From a top-level part they
are kept inside its folder, so Word images land in `word/media/`.

Pictures share one code path. Image headers are parsed without Pillow
(`images.py`: PNG, JPEG, GIF, BMP, including DPI). Identical images are stored
once (sha1) and related from each part that uses them. `drawing.picture()`
builds `pic:pic`, `xdr:pic` or `p:pic` from the same structure.

### 6. Packaging (OPC)
- **Untrusted input.** `Package.open(file, limits=Limits(...))` bounds the part count, the part size, the total expanded size and the compression ratio. Reads never trust declared sizes. Duplicate entries and DTDs are refused, and encrypted or legacy binary files are identified. Every failure is a `PackageError`, never a stray `zipfile`/`lxml` exception (fuzz-tested).
- **Strict conformance.** Strict namespace URIs and relationship types are translated as text before parsing, so trees are always Transitional in memory. On save they're translated back when `conformance == "strict"`. Value forms too: a percentage string on an integer attribute becomes thousandths on reading (the part remembers which attributes), and on writing Strict those, the known DrawingML percentage attributes and ST_OnOff values are converted on a copy of the tree, so memory stays Transitional and unchanged parts stay byte-identical. LibreOffice reads the result with the same values.
- Parts are parsed lazily. **Unchanged parts are written back byte-identical, even if you read them.** At save, a parsed part is compared to its original bytes by canonical XML (C14N), and if they match the original bytes are kept. Relationship files and `[Content_Types].xml` get the same treatment. So opening a file, reading it and saving it reproduces every zip entry exactly. This is tested on Word, PowerPoint, LibreOffice and openpyxl files.
- Saves are **deterministic**: fixed zip timestamps and sorted order, so the same document gives the same bytes.
- For changed parts, namespace declarations are hoisted to the part root on save. Post-2007 namespaces used in *ordinary content* are added to `mc:Ignorable`. Namespaces used only inside `*:extLst` or `mc:AlternateContent` are not, because those mechanisms carry their own compatibility rules (PowerPoint puts `p14` in extension lists). Prefixes already listed in Ignorable stay declared, so the value can't dangle.

## Layout

```
ooxpy/
├─ pyproject.toml
├─ src/oox/
│  ├─ __init__.py          public API
│  ├─ element.py           OoxElement, make(), lenient(), parser       (hand-written)
│  ├─ grammar.py           lazy grammar registry, context resolution  (hand-written)
│  ├─ attrconverters.py    Python value -> XML lexical value          (hand-written)
│  ├─ opc.py               zip / content types / relationships        (hand-written)
│  ├─ document.py          OoxDocument base, WordprocessingDocument, open_document()
│  ├─ spreadsheet.py       SpreadsheetDocument, cell/row/Formula, A1 helpers
│  ├─ presentation.py      PresentationDocument (master/layout/theme skeleton, slides)
│  ├─ drawing.py           DrawingML builders: picture, shape props, default theme
│  ├─ images.py            PNG/JPEG/GIF/BMP/TIFF/EMF/WMF/SVG header parsing (size, DPI)
│  ├─ units.py             EMU helpers: Inches, Cm, Mm, Pt, Px
│  ├─ word_defaults.py     Word's default styles, settings, fonts, section
│  ├─ properties.py        core, extended (app.xml) and custom (custom.xml) properties
│  ├─ numbering.py         Word lists: definitions, list instances, displayed labels
│  ├─ sections.py          Word sections: page setup, headers and footers
│  ├─ notes.py             hyperlinks, bookmarks, footnotes, endnotes, comments
│  ├─ fields.py            simple and complex fields, page numbers, TOC
│  ├─ template.py          find/replace across runs
│  ├─ tables.py            Word tables: merges, widths, header rows, reading
│  ├─ floating.py          anchored pictures with text wrapping; picture reader
│  ├─ revisions.py         tracked changes: suggest, list, accept, reject
│  ├─ formatting.py        run/paragraph formatting, custom styles, effective formatting
│  ├─ fieldcalc.py         captions, cross-references, field scanning and recomputing
│  ├─ controls.py          content controls (forms): add, read, fill
│  ├─ textboxes.py         Word text boxes (DrawingML with a VML fallback): add, read, edit
│  ├─ protection.py        Word editing restrictions and Excel workbook protection, password hashes
│  ├─ xlstyles.py          Excel stylesheet manager, number-format rendering, date detection
│  ├─ xlsheet.py           Excel layout: widths, heights, merges, panes, filter, names, validation
│  ├─ xlcond.py            Excel conditional formatting
│  ├─ xltables.py          Excel tables (ListObjects) with totals rows
│  ├─ xlnotes.py           Excel cell comments (with legacy VML) and hyperlinks
│  ├─ xlprint.py           Excel print setup, sheet options, protection
│  ├─ charts.py            DrawingML charts: builder (combinations, secondary axis), reader
│  │                       (Excel's type names, per-series type and axis), embedded data
│  ├─ xlcharts.py          Excel charts from cell ranges, anchors from column widths, chart sheets
│  ├─ slides.py            PowerPoint shapes, tables, speaker notes, picture placeholders, slide reader
│  ├─ slidetable.py        editing existing PowerPoint tables (rows, columns, merges, text)
│  ├─ slidelinks.py        PowerPoint click links on shapes and text (URLs, slides, show jumps)
│  ├─ sections_pml.py      PowerPoint slide sections (p14:sectionLst)
│  ├─ lexical.py           value-shape checks compiled from SDK validators
│  ├─ contentmodel.py      full content-model matching (particle trees -> NFAs)
│  ├─ __main__.py          command line: validate, info, text, downgrade
│  ├─ semantics.py         rules from the spec's prose (e.g. Excel row order)
│  ├─ strict.py            Strict <-> Transitional URI translation
│  ├─ errors.py
│  ├─ _odfdom.py          odfpy DOM mixin for OoxElement
│  ├─ odfcompat.py        load(), OpenDocumentText(), Text
│  ├─ teletype.py         extractText / addTextToElement
│  ├─ _generated/          GENERATED: namespaces, index, enums, parts, g_<prefix>.py grammar tables,
│  │                       p_<prefix>.py content-model trees
│  └─ ns/                  GENERATED: one factory module per prefix  (from oox.ns import w, a, p, x),
│                          each with a .pyi stub listing attribute keywords and their types
├─ tools/officecheck/      Office/LibreOffice compatibility kit (corpus, checkers, report)
├─ tools/interop.py        interop check over third-party files (and a corpus built from LibreOffice)
├─ tools/writers.py        files from python-docx, openpyxl, XlsxWriter, python-pptx (interop both ways)
├─ tools/bench.py          timings for large documents
├─ tools/apidocs.py        API reference pages (docs/reference) from the docstrings
├─ tools/api_surface.py    the public API, pinned in tests/api/public_api.txt
├─ tools/smoke_install.py  checks an installed wheel from outside the repository
├─ tools/gen/              generator (not shipped)
│  ├─ ir.py                source-agnostic IR
│  ├─ sdk_frontend.py      Open XML SDK JSON -> IR
│  ├─ xsd_frontend.py      ECMA-376 / ISO 29500 XSD -> IR (oracle; not used by the emitter)
│  ├─ irdiff.py            compares two IRs: python -m tools.gen.irdiff (docs/XSD_DIFF.md)
│  ├─ overrides.py         documented corrections where the SDK deviates from ECMA-376
│  ├─ fetch_sdk.py         sparse fetch of the pinned SDK data (SDK_REVISION)
│  ├─ emit.py              IR -> Python
│  └─ __main__.py          python -m tools.gen --sdk-data <Open-XML-SDK>/data
├─ tests/                  generator, element, package (python-docx used as an independent reader)
├─ examples/              runnable scripts, each with a short HTML tutorial next to it:
│                          hello (docx basics), edit_existing (open/read/edit), showcase (3 formats + pictures),
│                          report (sections, headers, notes, TOC, templates),
│                          tables_review (merged cells, floating pictures, tracked changes),
│                          forms_fields (formatting, styles, captions, fields, forms),
│                          excel_report (cell styles, number formats, dates, layout),
│                          excel_dashboard (tables, conditional formats, notes, print setup),
│                          charts (all three formats), slides (layouts, tables, shapes, notes)
├─ docs/                   the documentation site (MkDocs, mkdocs.yml): guides, cookbook, reference,
│                          stability policy, design documents; every code block runs in tests/test_docs.py
├─ CHANGELOG.md, RELEASING.md, MANIFEST.in
└─ .github/workflows/      ci.yml (tests, lint, docs, generated code, LibreOffice), release.yml, office-check.yml
```

Generated code is **committed**, as odfpy commits `grammar.py`. Users never
need the SDK, and diffs show exactly how the vocabulary changed between SDK
versions. `tests/test_generator.py::test_committed_code_matches_regeneration`
guards against drift (set `OOX_SDK_DATA`).

Pinned numbers: 4,315 types, 3,702 element factories, 2,596 SDK-name aliases,
139 namespace modules, about 4 MB of generated Python, imported lazily per
namespace. `import oox` plus the `w` factories takes about 0.1 s.

## odfpy concept map

| odfpy | oox |
|---|---|
| `odf.text`, `odf.table`, `odf.style` … | `oox.ns.w`, `oox.ns.x`, `oox.ns.p`, `oox.ns.a` … (one module per prefix) |
| `text.P(stylename=…, text=…)` | `w.P(w.PPr(w.PStyle(val=…)), w.R(w.T(…)))` |
| `grammar.py` (allowed_children/attributes) | `_generated/g_*.py` (+ order ranks, singletons, keys) |
| `Element.addElement / addText / setAttribute`, DOM (`childNodes`, `appendChild`…) | same names supported (see ODFPY_COMPAT.md); plus `add()`, `set_attr()` |
| `Element.qname` `(ns, local)` / `tagName` | same; oox's string form is `tag_name` |
| `attrconverters.py` | `attrconverters.py` |
| `OpenDocumentText()` / `OpenDocumentSpreadsheet()` / `OpenDocumentPresentation()` | same names in `oox.odfcompat`, or `WordprocessingDocument.new()` / `SpreadsheetDocument.new()` / `PresentationDocument.new()` |
| `addPicture(...)` | `addPicture(...)` returns an `r:id`; `add_picture(...)` places the picture |
| `load(path)` | `oox.odfcompat.load(path)` / `oox.open_document(path)`: any of the three formats |
| `teletype.py` | `oox.teletype` (`extractText`, `addTextToElement`) |
| `doc.styles` | `doc.styles` (OOXML has no automatic styles) |

## Known limitations
- **Microsoft Office testing:** scripted on Office 2007 (complete) and Office 2016 (every case opened cleanly; five Word cases have probes still to run on an activated Office); see `docs/office.md`. Excel's "save changes?" prompt on close is expected for formulas stored without a cached result (`Formula(expr, cached)` avoids it).
- Excel: `display_text` gives what English-locale Excel shows (Excel uses the viewer's regional separators and names unless the code has a `[$-409]`-style locale prefix); it uses "." / "," separators and English month and day names; fraction formats (`# ?/?`) and conditional sections (`[>100]`) aren't rendered. Column widths assume Calibri 11 (7 px per digit).
- PowerPoint: Office's eleven standard layouts are built in; handout masters, transitions and animations aren't modelled (they round-trip untouched). Sections are PowerPoint 2010+ (2007 and LibreOffice ignore them).
- Charts: columns, lines and areas combine (bars only with bars); scatter, pie and doughnut don't combine and have no secondary axis. Values are cached from the cells when the chart is added; charts without caches (openpyxl's) are read from the cells.
- Single-cell reads use caches checked on every lookup; editing the tree in ways that keep the same last row and shared-string count but change content in place (rewriting a shared string's text) isn't detected until the next structural change.
- Equations, SmartArt and mail merge have no helpers yet: the factories (`oox.ns.m`, `dgm`) build them by hand, and files containing them round-trip untouched.
- Protection is an editing restriction, not encryption: the content stays readable in the file, and any tool can remove it. Password-encrypted files (an OLE container around the package) can't be opened or written.
- Counts in `docProps/app.xml` (pages, words …) are read but not computed; Office rewrites them on save.
- Field results that depend on layout (PAGE, NUMPAGES, TOC page numbers) keep their cached values; PAGEREF results are estimates.
- Word: a TOC's page numbers appear only after Word updates the field. Tracked edits split only runs that hold plain text (not runs that mix text with tabs, breaks or fields).
- Strict files: namespaces, relationship types, DrawingML percentages (`"50%"` ↔ thousandths) and ST_OnOff (`on`/`off` → `true`/`false`) are translated. Word's own Strict renames (`w:jc` "left" → "start", `w:ind/@w:left` → `w:start`) are left as they are: Transitional accepts both spellings.
- Lexical checks skip the SDK's NCName/QName/URI/token flags (too loose to be worth the false-positive risk).
- SVG pictures need a raster fallback from the caller (oox doesn't rasterize SVG).
