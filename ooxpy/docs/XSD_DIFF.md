# XSD cross-check (Phase 3)

oox's grammar comes from the Open XML SDK's metadata (ADR-001). This page
records how that metadata compares with the standard's own schemas, and what
to do about the differences. The full, generated list is
[`reports/xsd-diff.md`](reports/xsd-diff.md).

## Tools

- `tools/gen/xsd_frontend.py` reads XSD files into the same IR as the SDK
  frontend. Keys follow the SDK's convention (`w:CT_RPr/w:rPr`), so the two IRs
  can be compared directly.
- `tools/gen/irdiff.py` pairs the two IRs' types and reports every deviation.
  It is **not wired into the emitter**, so the generated code is still the
  SDK's, byte for byte (`test_committed_code_matches_regeneration`).

```bash
python -m tools.gen.irdiff --sdk-data /path/to/Open-XML-SDK/data \
    --xsd /path/to/schemas --out docs/reports/xsd-diff --limit 5000 --label "…"
```

**Schemas.** The XSD files aren't in the repository, because their licence is
the standard's, not ours. The run below used:
- the ISO/IEC 29500-4:2016 Transitional schemas (`wml.xsd`, `sml.xsd`, `pml.xsd`, `dml-*.xsd`, `vml-*.xsd`, `shared-*.xsd`);
- the MCE schema and the OPC schemas;
- Microsoft's Word extension schemas (`wml-2010.xsd` … `wml-sdtdatahash-2020.xsd`).

The ECMA-376 and ISO/IEC 29500 downloads contain the same files. The frontend
reads any directory of `*.xsd`, so the Strict schemas (ISO/IEC 29500-1) can be
compared the same way once they're at hand.

**Tests.** `tests/test_xsd_frontend.py` covers every XSD construct the schemas
use, with synthetic schemas, plus the diff's normalisation, pairing and
categories. With `OOX_XSD_DIR` and `OOX_SDK_DATA` set, it also runs the real
comparison. That run must find the deviations the Office checks have already
met.

## How the comparison works

**Pairing.** Types are paired by key. Where the SDK renames or specialises a
type, they are paired through context instead: two children with the same
name under two paired parents are paired too. For example, the SDK's
`w:CT_OnOffOnly/w:cantSplit` pairs with the XSD's `w:CT_OnOff/w:cantSplit`, and
`a:CT_BlipExtensionList` pairs with `a:CT_OfficeArtExtensionList`. A qname that
has a single type on each side is paired too. In this run, 2,637 pairs were
found; 404 of them were paired through context.

**What isn't counted as a difference:**
- Microsoft's extension children and attributes (w14, c15, p14 …) inside standard types: the SDK folds them in, while the standard's types don't have them.
- SDK children that fall under an `xsd:any` in the XSD.
- Content models that differ only in finite maximums, which are the SDK's implementation limits (`a:tr{1,1000}`, `x:xf{1,65430}`). These are listed, but last.

Both content models are normalised the same way before comparing:
- groups are inlined;
- nested 1..1 compositors are flattened;
- equivalent wrappers are collapsed: `(x*)?` is the same as `x*`.

**Direction.** For each finding, the report says which side is stricter.
- **SDK stricter:** oox's validator may flag a file that ISO/IEC 29500 allows. This is a false-positive risk.
- **XSD stricter:** oox may accept something that ISO/IEC 29500 forbids. This is leniency.

## Result (SDK data @ 431ab05, ISO/IEC 29500-4:2016)

| | SDK stricter | XSD stricter |
|---|---|---|
| Children on one side only | 93 | 59 |
| Attributes on one side only | 52 | 43 |
| Enumeration values | 94 | 65 |
| Required attributes | 88 | 22 |

There are also 66 content-model differences, 35 of them limits only.

The XSD frontend reports one problem in the schemas themselves: `wml.xsd`
references a `w:ST_String` that it doesn't define.

### It finds what the Office checks found

- **`word-94` (Office 2007 opened it):**
  - `w:lvlJc` and table `w:jc` accept `start`/`end` in ISO. The SDK's `ST_LevelJustification` and `ST_TblJc` don't.
  - The SDK's `CT_OnOffOnly` (`w:cantSplit`, `w:tblHeader`, and 17 others) lacks ISO's `true`/`false`/`1`/`0`.
  - All of these are already accepted at run time (`grammar.EXTRA_ENUM_VALUES`). The diff lists every other element that uses `CT_OnOffOnly` too.
- **`overrides.py`:** sml's `x:font` and `x:rPr` are unordered choices in the XSD and fixed sequences in the SDK data. This shows up as "known".

### Worth triaging (SDK stricter, plausible in real files)

| Finding | Risk | Next step |
|---|---|---|
| `c:doughnutChart` requires `c:holeSize`; `c:barDir`, `c:holeSize`, `c:order`, `c:splitType`, `c:radarStyle`, `c:ofPieType` require `@val` (ISO: optional, with defaults) | Charts from other producers flagged | **Observe case `excel-93`** |
| `x:pageSetup/@paperWidth`, `@paperHeight` (ISO custom paper size) | Sheets with custom paper flagged | **Observe case `excel-94`** |
| `w:tblStyleRowBandSize`/`ColBandSize` allowed in a body table's `w:tblPr` (SDK: only in table styles) | Documents from other producers flagged | **Observe case `word-96`** |
| Page borders `w:pgBorders/*`: `r:id`, `r:topLeft` … (art borders) | Word files with art page borders flagged | Check a Word-authored file with an art border |
| `x:cfRule/@type`, `x:filter/@val`, `x:c/@r` (calc chain) required in the SDK only | Workbooks from other producers flagged | Look for such files in the interop corpus |
| ISO-only enumeration values: `c:scatterStyle` "none", `c:splitType` "auto", `w:pos` "docEnd", `w:caption/@pos` "left"/"right", `a:anchor` "dist"/"just" | Values from ISO-following producers flagged | Office check when one turns up |
| `w:ST_Border`: SDK has `tribal1`–`6`, ISO has `custom`, `earth3`; VML `o:bwmode`: SDK `lightGrayScale`, ISO `lightGrayscale` | Spelling and edition differences | Check which spelling Word-authored files use |

The XSD-stricter side includes the seven `w:pgMar` attributes and
`a:latin/@typeface`, which are required in ISO and optional in the SDK. oox's
own output writes all of them, so this only matters for checking other files.

## Policy

The same as for `word-94`: **no relaxing without Office evidence.** A
difference becomes an observe case, and runs on Office 2007 and 2016. Only
when Office opens the file does oox accept the markup, and it does so at run
time (`grammar.EXTRA_ENUM_VALUES`, `lexical.OPEN_ENUMS`, `tools/gen/overrides.py`),
so the generated code stays the pinned SDK's. When Office refuses a file, the
SDK's stricter rule stays, and the finding is recorded as Office's behaviour.

## Next steps

1. Run `word-96`, `excel-93` and `excel-94` on Office 2007 and 2016 (`--only observe`), then decide case by case. **Office 2016 (2026-11-16):** `excel-93` refused to open, so the SDK's required `c:holeSize` and `@val` stay required; `word-96` and `excel-94` opened, so their relaxations wait only for Office 2007 (`--only word-96,excel-93,excel-94`).
2. Compare against the Strict schemas (ISO/IEC 29500-1) when available, which also checks oox's Strict translation.
3. Use the XSD frontend for namespaces the SDK lacks, if a user needs them (Phase 3 item 3).
4. Differential testing against openxml-audit (Phase 3 item 4) remains.
