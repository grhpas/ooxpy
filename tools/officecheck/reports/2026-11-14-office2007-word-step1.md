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
| `word-10-defaults` | ✅ |
| `excel-06-properties` | ✅ |
| `ppt-06-properties` | ✅ |
| `word-91-observe-order` *(observe)* | · |
| `word-92-observe-unknown-element` *(observe)* | · |
| `word-93-observe-bad-values` *(observe)* | · |
| `excel-91-observe-row-order` *(observe)* | · |
| `excel-92-observe-newer-attribute` *(observe)* | · |
| `ppt-91-observe-missing-required` *(observe)* | · |
| `ppt-90-control-malformed` *(control)* | · |

Sources: `results-mosyxlon3-word12.json`, `results-mosyxlon3-excel12-powerpoint12.json`
