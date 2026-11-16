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
| `word-91-observe-order` *(observe)* | 👁 ok |
| `word-92-observe-unknown-element` *(observe)* | 👁 ok |
| `word-93-observe-bad-values` *(observe)* | 👁 prompt |
| `excel-91-observe-row-order` *(observe)* | 👁 error |
| `excel-92-observe-newer-attribute` *(observe)* | 👁 error |
| `ppt-91-observe-missing-required` *(observe)* | 👁 prompt |
| `ppt-90-control-malformed` *(control)* | · |

## Details

- **word-91-observe-order** on mosyxlon3 (Office 2007): ok
- **word-92-observe-unknown-element** on mosyxlon3 (Office 2007): ok
- **word-93-observe-bad-values** on mosyxlon3 (Office 2007): prompt
  - dialog `Microsoft Office Word`: 
  - error: `com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'Command failed', 'C:\\Program Files (x86)\\Microsoft Office\\Office12\\1033\\WDMAIN11.CHM', 36966, -2146824090), None)`
- **excel-91-observe-row-order** on mosyxlon3 (Office 2007): error
  - error: `com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Office Excel', 'Open method of Workbooks class failed', 'C:\\Program Files (x86)\\Microsoft Office\\Office12\\1033\\XLMAIN11.CHM', 0, -2146827284), None)`
- **excel-92-observe-newer-attribute** on mosyxlon3 (Office 2007): error
  - error: `com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Office Excel', 'Open method of Workbooks class failed', 'C:\\Program Files (x86)\\Microsoft Office\\Office12\\1033\\XLMAIN11.CHM', 0, -2146827284), None)`
- **ppt-91-observe-missing-required** on mosyxlon3 (Office 2007): prompt
  - dialog `Microsoft Office PowerPoint`: OK / PowerPoint was unable to display some of the text, images, or objects on slides in the file, "ppt-91-observe-missing-required.pptx", because they have become corrupted. Affected slides have been replaced by blank slides in the presentation and it is not possible to recover the lost information. To ensure that the file can be opened in previous versions of PowerPoint, use the Save As command (
  - probe `ppt.text_contains`: expected `'without grpSpPr'`, got `'\n'`

Sources: `results-mosyxlon3-word12-excel12-powerpoint12.json`
