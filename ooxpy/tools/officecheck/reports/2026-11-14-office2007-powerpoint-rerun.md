# Office check report

Corpus: oox 0.0.1, Open XML SDK data @ 431ab05cf160248cc3885a4a766026d4f8243792 2026-08-18.

✅ ok · ⚠️ prompt (repair/recovery dialog) · ❌ error or content mismatch · ⏱️ timeout · · not run

*Control* cases are broken on purpose: ✅ means the checker noticed, either because Office complained or because it repaired silently and content went missing.

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
| `ppt-01-minimal` | ✅ |
| `ppt-02-placeholders` | ✅ |
| `ppt-03-textboxes-pictures` | ✅ |
| `ppt-04-4x3` | ✅ |
| `ppt-05-edit-powerpoint-file` | ✅ |
| `word-90-control-malformed` *(control)* | · |
| `excel-90-control-malformed` *(control)* | · |
| `ppt-90-control-malformed` *(control)* | ✅ detected |

## Details

- **ppt-90-control-malformed** on mosyxlon3 (Office 2007): prompt
  - dialog `Microsoft Office PowerPoint`: OK / PowerPoint was unable to display some of the text, images, or objects on slides in the file, "ppt-90-control-malformed.pptx", because they have become corrupted. Affected slides have been replaced by blank slides in the presentation and it is not possible to recover the lost information. To ensure that the file can be opened in previous versions of PowerPoint, use the Save As command (File me
  - probe `ppt.text_contains`: expected `'Outlook'`, got `'\n'`

Sources: `results-mosyxlon3-powerpoint12.json`
