# Document-family fixtures

Small packages for the validator's Excel and PowerPoint tests (`../document-kinds.test.ts`).

- `valid.xlsx`: a hand-built minimal SpreadsheetML package (one sheet, one inline string, one number). It was not saved by Excel.
- `invalid-schema.xlsx`: the same package with an unknown element after `sheetData`.
- `invalid-schema.pptx`: `../fixtures/minimal.pptx` with an unknown element appended to `ppt/presentation.xml`.

They live outside `../fixtures` because the parity test treats every file in that folder as a golden corpus document.
