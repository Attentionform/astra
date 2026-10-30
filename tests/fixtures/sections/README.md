# Section fixtures

Files with a known section structure, used by the section tests. They exist so the parser is checked against files that ASTra did not write.

| File | Made by | Sections |
|---|---|---|
| `python-docx-three-sections.docx` | `generate_python_docx.py` (python-docx 1.2.0) | 3: portrait Letter; landscape with its own header; continuous portrait with a different first-page header |
| `libreoffice-two-page-styles.docx` | LibreOffice, converted from `libreoffice-two-page-styles.fodt` | 3: Letter portrait, A4 landscape, Letter portrait |
| `characterization.golden.json` | the characterization test | hashes of what opening each document produces. The "frozen" values were recorded before the parser read the file's sections and are unchanged, except one deliberate fix: a stray `w:sectPr` inside a table cell no longer decides the page setup of a one-section document. The "info" values show the sections as read now |

The LibreOffice file is kept exactly as LibreOffice wrote it. It puts `w:sectPr` before `w:pStyle` inside `w:pPr`, which Word accepts and the strict schema does not; ASTra's save puts the elements back in schema order. The fixture test pins both facts.

The other shapes (trailing empty section, a break inside a content control, a tracked section change, a `w:sectPr` inside a table cell, and so on) are built from plain XML in `tests/helpers/section-docs.ts`, so they need no binary file.

## Regenerating

```bash
cd tests/fixtures/sections
python3 generate_python_docx.py
soffice --headless --convert-to docx:"MS Word 2007 XML" --outdir . libreoffice-two-page-styles.fodt
```

Regenerating changes the bytes (creation times), so rewrite the golden file afterwards:

```bash
ASTRA_UPDATE_GOLDEN=1 npx vitest run tests/unit/sections-characterization.test.ts
```

Only do that for a change you meant to make: the golden file is what proves that a change to how sections are read did not alter anything else.
