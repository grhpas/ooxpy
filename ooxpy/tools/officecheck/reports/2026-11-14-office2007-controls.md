# Office check report

Corpus: oox 0.0.1, Open XML SDK data @ 431ab05cf160248cc3885a4a766026d4f8243792 2026-08-18.

✅ ok · ⚠️ prompt (repair/recovery dialog) · ❌ error or content mismatch · ⏱️ timeout · · not run

*Control* cases are broken on purpose: ✅ means Office complained and the checker saw it.

| Case | mosyxlon3 (Office 2007) |
|---|---|
| `word-01-minimal` | ✅ |
| `word-02-formatting` | ✅ |
| `word-03-table` | ✅ |
| `word-04-pictures` | ✅ |
| `word-05-styles` | ✅ |
| `word-06-w14` | ✅ |
| `word-07-odfpy-passwd` | ✅ |
| `word-08-edit-word-file` | ✅ |
| `excel-01-types` | ✅ |
| `excel-02-formulas` | ✅ |
| `excel-03-sheets` | ✅ |
| `excel-04-pictures` | ✅ |
| `excel-05-edit-existing` | ✅ |
| `ppt-01-minimal` | ✅ |
| `ppt-02-placeholders` | ✅ |
| `ppt-03-textboxes-pictures` | ✅ |
| `ppt-04-4x3` | ✅ |
| `ppt-05-edit-powerpoint-file` | ✅ |
| `word-90-control-malformed` *(control)* | ✅ detected |
| `excel-90-control-malformed` *(control)* | ✅ detected |
| `ppt-90-control-malformed` *(control)* | ❌ not detected (ok) |

## Details

- **ppt-90-control-malformed** on mosyxlon3 (Office 2007): ok

Sources: `results-mosyxlon3-word12-excel12-powerpoint12.json`
