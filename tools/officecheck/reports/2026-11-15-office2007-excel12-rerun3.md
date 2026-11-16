# Office check report

Corpus: oox 0.0.1, Open XML SDK data @ 431ab05cf160248cc3885a4a766026d4f8243792 2026-08-18.

✅ ok · ⚠️ prompt (repair/recovery dialog) · ❌ error or content mismatch · ⏱️ timeout · · not run

*Control* cases are broken on purpose: ✅ means the checker noticed, either because Office complained or because it repaired silently and content went missing.

*Observe* cases (👁) break the schema in one specific way and are not judged: they record how strict each Office version is.

| Case | mosyxlon3 (Office 2007) |
|---|---|
| `word-01-minimal` | · |
| `word-02-formatting` | · |
| `word-03-table` | · |
| `word-04-pictures` | · |
| `word-05-styles` | · |
| `word-06-w14` | · |
| `word-07-odfpy-passwd` | · |
| `word-08-edit-word-file` | · |
| `excel-01-types` | · |
| `excel-02-formulas` | · |
| `excel-03-sheets` | · |
| `excel-04-pictures` | · |
| `excel-05-edit-existing` | · |
| `ppt-01-minimal` | · |
| `ppt-02-placeholders` | · |
| `ppt-03-textboxes-pictures` | · |
| `ppt-04-4x3` | · |
| `ppt-05-edit-powerpoint-file` | · |
| `word-90-control-malformed` *(control)* | · |
| `excel-90-control-malformed` *(control)* | · |
| `word-10-defaults` | · |
| `word-11-lists` | · |
| `word-12-sections` | · |
| `word-13-links-notes` | · |
| `word-14-toc-template` | · |
| `word-15-tables` | · |
| `word-16-floating-pictures` | · |
| `word-17-tracked-changes` | · |
| `word-18-tracked-changes-reject` | · |
| `word-19-formatting-styles` | · |
| `word-20-fields-captions` | · |
| `word-21-content-controls` | · |
| `word-22-more-images` | · |
| `excel-07-properties-images` | · |
| `ppt-07-properties-images` | · |
| `excel-08-styles-formats` | · |
| `excel-09-layout` | · |
| `excel-10-style-existing` | · |
| `excel-11-names-autofit` | · |
| `excel-12-conditional` | ❌ |
| `excel-13-tables` | · |
| `excel-14-comments-links` | · |
| `excel-15-print-options` | · |
| `excel-06-properties` | · |
| `ppt-06-properties` | · |
| `word-91-observe-order` *(observe)* | · |
| `word-92-observe-unknown-element` *(observe)* | · |
| `word-93-observe-bad-values` *(observe)* | · |
| `excel-91-observe-row-order` *(observe)* | · |
| `excel-92-observe-newer-attribute` *(observe)* | · |
| `ppt-91-observe-missing-required` *(observe)* | · |
| `ppt-90-control-malformed` *(control)* | · |

## Details

- **excel-12-conditional** on mosyxlon3 (Office 2007): mismatch
  - probe `excel.cf_types` sheet=Rules ref=B2: expected `['cellIs', 'cellIs', 'expression', 'top10']`, got `['cellIs', 'cellIs', 'cellIs', 'top10']`
    - sheet Item(1) kind: com_error 0x800A03EC
    - cell Item(1): FormatCondition cellIs priority 1
    - cell Item(2): FormatCondition cellIs priority 2
    - cell Item(3): Top10 top10 priority 3
    - cell Item(4): FormatCondition cellIs priority 9
    - route 2 (cell rules) used
  - probe `excel.cf_types` sheet=Rules ref=C2: expected `['colorScale', 'expression']`, got `['colorScale', 'colorScale']`
    - sheet Item(9) kind: com_error 0x800A03EC
    - cell Item(1): ColorScale colorScale priority 4
    - cell Item(2): FormatCondition colorScale priority 9
    - route 2 (cell rules) used
  - probe `excel.cf_types` sheet=Rules ref=D2: expected `['dataBar', 'expression']`, got `['dataBar', 'dataBar']`
    - sheet Item(9) kind: com_error 0x800A03EC
    - cell Item(1): Databar dataBar priority 5
    - cell Item(2): FormatCondition dataBar priority 9
    - route 2 (cell rules) used
  - probe `excel.cf_types` sheet=Rules ref=E2: expected `['expression', 'iconSet']`, got `['iconSet', 'iconSet']`
    - sheet Item(9) kind: com_error 0x800A03EC
    - cell Item(1): IconSetCondition iconSet priority 6
    - cell Item(2): FormatCondition iconSet priority 9
    - route 2 (cell rules) used
  - probe `excel.cf_types` sheet=Rules ref=A3: expected `['containsText', 'duplicateValues', 'expression']`, got `['containsText', 'containsText', 'duplicateValues']`
    - sheet Item(7) kind: com_error 0x800A03EC
    - cell Item(1): FormatCondition containsText priority 7
    - cell Item(2): UniqueValues duplicateValues priority 8
    - cell Item(3): FormatCondition containsText priority 9
    - route 2 (cell rules) used

Sources: `results-mosyxlon3-excel12-20261115-032315.json`
