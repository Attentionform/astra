# Page Layout: Sections, Headers, Footers and Lists

A task guide for laying out a document with ASTra: page orientation and size, separate sections, a different header or footer for each section, columns, page borders and multi-level lists.

Every example is a list of operations. The same list works in all three places:

- **Command line:** save it as `ops.json` and run `astra batch report.docx ops.json --out report.docx --force`.
- **MCP server:** pass it as `operations` to `execute_batch`, then call `save_document`.
- **Library:** pass it to `OperationEngine.executeBatch({ operations })`.

The three give the same file. This was checked by running one scenario through all three (three sections, a header for each, a landscape middle section, a bullet list, then deleting a section break) and comparing the results. The operation fields are listed in the [MCP Specification](mcp-specification.md) and the commands in the [CLI Guide](cli-guide.md).

## Contents

1. [Concepts](#1-concepts)
2. [Orientation and paper size](#2-orientation-and-paper-size)
3. [Sections](#3-sections)
4. [Headers and footers](#4-headers-and-footers)
5. [Columns and page borders](#5-columns-and-page-borders)
6. [Lists](#6-lists)
7. [Deleting a section break](#7-deleting-a-section-break)
8. [A complete example](#8-a-complete-example)
9. [Known limits](#9-known-limits)

---

## 1. Concepts

- **A section is a run of pages with its own page setup.** Orientation, paper size, margins, columns, page border, header and footer belong to a section, not to the whole file.
- **A section break ends a section.** In the file the break is a property of the last paragraph of that section. The last section has no break paragraph; its settings sit at the end of the document.
- **Sections are numbered from 1**, in reading order, in every operation that takes `sectionIndex`. Leave `sectionIndex` out and the operation applies to the whole document (for `pageSetup`) or to the first section (for the others).
- **Sections are read from the file.** If an edit adds, deletes or moves a section break, ASTra rebuilds its section list from the file, so the list never disagrees with what Word will show.
- **Node IDs belong to a session.** A paragraph you insert gets the ID you give it in `id` for the rest of that batch. In a later command or session, read the IDs again with `astra inspect` or `astra search`.

## 2. Orientation and paper size

Landscape for the whole document:

```json
[{ "action": "pageSetup", "orientation": "landscape" }]
```

A4, portrait, with narrower side margins (measured in twentieths of a point; 1440 is one inch):

```json
[{ "action": "pageSetup", "paperSize": "A4", "orientation": "portrait", "marginLeft": 1080, "marginRight": 1080 }]
```

`paperSize` is `A4`, `Letter`, `Legal` or `A3`. Landscape swaps the page width and height and leaves the margins as they are.

One section only:

```json
[{ "action": "pageSetup", "sectionIndex": 2, "orientation": "landscape" }]
```

Other fields: `marginTop`, `marginBottom`, `headerDistance`, `footerDistance` and `columns`.

## 3. Sections

Insert a section break after a paragraph. `breakType` is `nextPage` (start the next section on a new page) or `continuous` (no page change).

```json
[
  { "action": "insert", "type": "paragraph", "anchor": "end", "position": "after", "text": "Cover page", "id": "cover" },
  { "action": "insert", "type": "sectionBreak", "anchor": "cover", "position": "after", "breakType": "nextPage" },
  { "action": "insert", "type": "paragraph", "anchor": "end", "position": "after", "text": "Wide table section", "id": "wide" }
]
```

`anchor: "end"` adds to the end of the document. Set a section's page setup after you have created it, with `sectionIndex`, as in section 2.

## 4. Headers and footers

Set a header or footer for one section:

```json
[
  { "action": "headerFooter", "operation": "set", "targetType": "header", "sectionIndex": 1, "text": "Quarterly report" },
  { "action": "headerFooter", "operation": "set", "targetType": "footer", "sectionIndex": 3, "text": "Page {PAGE} of {NUMPAGES}" }
]
```

- `{PAGE}` is the current page number and `{NUMPAGES}` the total. They are stored as real Word fields, so they update in Word.
- A section without its own header or footer shows the one from the section before it, as in Word. In the example above, sections 2 and 3 show "Quarterly report", and only section 3 has the footer.
- To give a section no header of its own again, use `"operation": "remove"` with the same `targetType` and `sectionIndex`.

## 5. Columns and page borders

```json
[
  { "action": "section", "operation": "setColumns", "sectionIndex": 3, "columns": 2, "separator": true },
  { "action": "pageBorder", "sectionIndex": 1, "style": "double", "sz": 12, "color": "1F4E79" }
]
```

- `columns` sets the number of columns; `spacing` is the gap between them (default 720, half an inch) and `separator` draws a line between them.
- `pageBorder` takes `style` (`single`, `double`, `dashed`, `triple`), `sz` in eighths of a point (12 is 1.5 pt), `color` as a hex value without `#`, and `marginOffset` (`page` or `text`).

## 6. Lists

Make a paragraph a bullet or numbered list item with `list` / `create`. Use `continue` to add the next item, with a `level` from 1 (top) to 9.

```json
[
  { "action": "insert", "type": "paragraph", "anchor": "end", "position": "after", "text": "Prepare", "id": "s1" },
  { "action": "insert", "type": "paragraph", "anchor": "s1", "position": "after", "text": "Check", "id": "s2" },
  { "action": "insert", "type": "paragraph", "anchor": "s2", "position": "after", "text": "Ship", "id": "s3" },
  { "action": "list", "operation": "create", "target": "s1", "listType": "numbered" },
  { "action": "list", "operation": "continue", "target": "s2", "previousNodeId": "s1", "level": 2 },
  { "action": "list", "operation": "continue", "target": "s3", "previousNodeId": "s2", "level": 1 }
]
```

This produces `1. Prepare`, `a. Check`, `2. Ship`.

- **Bullets** use a different mark at each level: • ◦ ▪ and so on, down to level 9.
- **Numbers** run 1. a. i. and repeat the pattern through nine levels.
- `changeLevel` moves an item up or down: `"level": 3`, or `"delta": 1` for one level deeper. Level 0 takes the item out of the list. `remove` also takes an item out of its list.

## 7. Deleting a section break

Delete the break paragraph by its ID. It is the last paragraph of its section, so in the output of `astra inspect` it is the (usually empty) row just before the next section starts.

```json
[{ "action": "delete", "target": "<ID of the break paragraph>" }]
```

What happens:

- The two sections on either side merge. **The text before the deleted break takes the page setup of the section after it**, which is how Word behaves. If a landscape section is merged into a portrait one, it turns portrait.
- The deleted section's header and footer are **removed from the file**, not just hidden. Their text, their relationships and pictures used only by them are gone, so the text of a header you deleted cannot be found by unzipping the file. A header that something else still uses (another section, or a tracked change) is kept.
- ASTra prints a warning saying how many sections the document had and now has.
- The edit is part of the batch like any other: if a later operation in the same batch fails, the whole batch is undone and the file is not written.

A delete with an ID that does not exist is not an error: the response lists it under `skippedIds` and `appliedCount` is `0`. Check `appliedCount` when you script this.

## 8. A complete example

Three sections: a cover page with a double border, a landscape page, and a page with two columns, a list and a page-numbered footer.

```bash
astra create report.docx --template blank --force
astra batch report.docx build.json --out report.docx --force     # operations from sections 3 and 6 of this guide
astra batch report.docx layout.json --out report.docx --force    # operations from sections 2, 4 and 5 of this guide
astra validate report.docx
astra visual report.docx --page 2 --out page2.png
```

`build.json` holds the insert and list operations from sections 3 and 6; `layout.json` holds the header, footer, `pageSetup`, `setColumns` and `pageBorder` operations. They are two files because the second needs sections that the first creates.

## 9. Known limits

- **`.docx` only.** `.doc`, `.odt` and other formats are not edited.
- **Checked automatically with the Open XML SDK and LibreOffice, not with Word.** Fifteen sample documents covering sections, headers and footers, lists, a chart and other content were also opened by hand in desktop Word and opened without a repair prompt; the chart stayed editable. Equations and comments were not confirmed separately.
- **PDF export and page images need LibreOffice and Poppler** installed on the machine. On Windows, `--renderer word` (Word 2010 or later) and `--pdf-to-image windows` (built in) can stand in. Word and LibreOffice lay a page out differently, so page breaks can differ between the two renderers; the wireframe's page count follows the renderer that drew the page.
- **Tracked changes and comments in renders are opt-in** (`--show-markup`, MCP `showMarkup`). The review view uses a different layout mode: pages can be scaled and pagination can change, so page numbers in a markup render refer to that render. Word shows comments in balloons and LibreOffice shows them in the margin. LibreOffice cuts a long comment short on every version that was tried: on 24.8, 25.2 and 25.8 the margin box runs past the page's right edge, and on 7.6, 26.2 and 26.8 it wraps but is shortened when other comments follow close below. Word cuts a very long comment too: its balloon has a maximum height and ends with an ellipsis (a 60-word comment was complete, a 300-word one ended after about 150 words on Word 16.0.14334.20974). A markup render always carries a warning for the program that drew it. A comment anchored in a header, footer or footnote is not drawn by Word, and LibreOffice cannot open such a file. Comment icons and the exact balloon layout are not controlled. ASTra needs LibreOffice 7.6.7 or later for markup, because older versions were not tried and LibreOffice silently ignores an export option it does not know. The check is a guard on the version number, not a guarantee that every later build behaves the same. The evidence is in the table below.
- **Version 0.x.** Interfaces may change before 1.0.
- **Node.js 22.19 or later.**
- **Node IDs last for one session.** Read them again after you reopen a file.
- **A batch takes at most 500 operations and 200 MB** of document data; over that it stops with `E_QUOTA`.
- **Deleting a section break merges the sections** and the text takes the next section's page setup (section 7).
- **Landscape swaps width and height and keeps the margins.** Change margins separately if you want them to follow.
- **Revision history is off until you turn it on** (`astra history <file> --init`).

### Review marks: what was tried

`scripts/word-markup-smoke.mjs` runs these checks with the real programs. It is run by hand before a release, not in the unit tests.

| Program | L1 comment and tracked deletion | L2a three authors | L3 comment across a change, comment on last line | L4 plain document unchanged | L2b long comment (recorded only) |
|---|---|---|---|---|---|
| LibreOffice 7.6.7.2 (Linux) | pass | pass | pass | pass | cut short |
| LibreOffice 24.8.7.2 (Linux) | pass | pass | pass | pass | cut short |
| LibreOffice 25.2.7.2 (Linux) | pass | pass | pass | pass | cut short |
| LibreOffice 25.8.7.3 (Linux) | pass | pass | pass | pass | cut short |
| LibreOffice 26.2.7.3 (Linux) | pass | pass | pass | pass | cut short |
| LibreOffice 26.8.0.3 (Linux) | pass | pass | pass | pass | cut short |
| LibreOffice 26.8.0.3 (Windows) | pass | pass | pass | pass | cut short |

Versions older than 7.6.7 were not tried. Windows was tried with one LibreOffice version only.

Microsoft Word 16.0.14334.20974 on Windows passed the checks for the default view, for Word starting in "No Markup", with one reviewer hidden and with comments off (the view comes back after each), for three reviewers, and for a user's own open Word keeping its process and window. It also showed a 60-word comment complete beside two other comments, and cut a 300-word comment after about 150 words with an ellipsis (checks R17 and R17b). Other Word builds were not tried.

### Review marks: other documents that were tried

The same script (`--cases`) renders these documents with and without markup and prints what it found. These are recorded, not pass or fail: they show what each program does. The documents are made by the script from raw OOXML, not by Word. Tried on Word 16.0.14334.20974, LibreOffice 25.2.7.2 (Linux) and LibreOffice 26.8.0.3 (Linux). Other versions were not tried for these.

| Case | Word | LibreOffice 25.2.7.2 and 26.8.0.3 |
|---|---|---|
| R5 a reply thread and a resolved comment | root, reply and resolved comment all drawn | all three drawn |
| R6a, R6b, R6c comment anchored in a header, a footer, a footnote | the file opens, the comment is not drawn, the header, footer and footnote text is | the file does not open, with or without markup (25.2.7.2: "source file could not be loaded"; 26.8.0.3: the program exits with an error) |
| R6d comment in a table cell | drawn | drawn |
| R7 15 comments on one page | all 15 drawn, one page | all 15 drawn, one page |
| R8 formatting change, move, inserted table row | the "Formatted" note, the moved text in both places and the inserted row | moved text in both places and the inserted row; no "Formatted" note |
| R9 right-to-left paragraph with a comment | comment and Hebrew text drawn | comment and Hebrew text drawn |
| R10 30-page document, comments on pages 3, 12 and 25 | all three drawn, 30 pages with and without markup | all three drawn, 30 pages with and without markup |
| R11 the document asks for a view that hides markup (`w:revisionView`) | comment and deleted text drawn, because ASTra sets the view itself | comment drawn, deleted text not drawn |
| R12 comment in the middle of a sentence | comment drawn, sentence intact | comment drawn, sentence intact |
| R13 plain document | same text and pages with and without markup | same text and pages with and without markup |

Two limits follow from this table. A comment anchored in a header, a footer or a footnote is not drawn by Word, and LibreOffice cannot open such a file at all, with or without markup; this is the same in a render that has no markup. When `auto` is used and LibreOffice cannot open the file, ASTra tries Word where it is installed. With LibreOffice, a document whose `w:revisionView` hides markup also hides the deleted text in a markup render. Word 16.0.14334.20974 was also run with three renders at the same time in separate processes (all three produced the same PDF and no Word process was left), and with another process holding the Word launch lock (the render failed after 30 seconds with `E_RENDER` and started no Word).

