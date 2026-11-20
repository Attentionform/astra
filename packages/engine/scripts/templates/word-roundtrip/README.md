# Word round trip

`tests/fixtures/word-saved/*.docx` are the built-in templates after a real Microsoft Word opened and saved them. The test
`tests/word-saved-templates.test.ts` composes onto each one and requires that the placeholder body is still recognised as untouched.
That is the check that the shortcut survives what Word does to a file: it moves bookmarks into paragraphs, splits runs, adds `_Toc`
bookmarks, `w:cnfStyle` and `w:lastRenderedPageBreak`, drops `xml:space` and recomputes `w:tblLook`.

Regenerate the fixtures when a template's placeholder body changes (the manifest hash changes with it). This needs Word on Windows:

1. `npm run build:templates --workspace=@attentionform/astra-engine`
2. Make a folder with the sub-folders `in` and `out`. Copy the nine designed Letter templates from `packages/engine/templates` into `in`.
3. `powershell -NoProfile -ExecutionPolicy Bypass -File save-with-word.ps1 -Dir <that folder>` opens and saves each file with Word into `out`.
4. Remove the author, last-modified-by and company text from `docProps/core.xml` and `docProps/app.xml` of each result, then copy
   the files into `tests/fixtures/word-saved`. Word writes the Windows user name into them; do not commit it.
