# ASTra MCP Server Specification

The official Model Context Protocol (MCP) reference for ASTra.

ASTra exposes **21 tools** over standard input/output (`stdio`, using JSON-RPC). They let AI agents and developer tools inspect, search, edit, export and keep a history of Microsoft Word (`.docx`) documents. Every save is checked against the document schema first, and a batch of edits is applied all-or-nothing.

---

## Table of Contents

1. [Quick Client Integration](#1-quick-client-integration)
2. [Security & Sandboxing Model](#2-security--sandboxing-model)
3. [The Standard Agent Workflow](#3-the-standard-agent-workflow)
4. [Complete 21-Tool Reference](#4-complete-21-tool-reference)
   - [Document Management](#41-document-management)
   - [Inspection & Query](#42-inspection--query)
   - [Mutation & Composition](#43-mutation--composition)
   - [Export & Visual Telemetry](#44-export--visual-telemetry)
   - [History & Time-Travel](#45-history--time-travel)
   - [Document Audit & Maintenance](#46-document-audit--maintenance)
5. [TOON Grammar & Flag Legend](#5-toon-grammar--flag-legend)
6. [ECMA-376 Units & Formatting Reference](#6-ecma-376-units--formatting-reference)
7. [Error Taxonomy & Remediation](#7-error-taxonomy--remediation)

---

## 1. Quick Client Integration

Connect ASTra to an AI client that supports MCP servers:

### Claude Desktop
Add ASTra to your Claude Desktop configuration file:
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "astra": {
      "command": "npx",
      "args": [
        "-y",
        "@attentionform/astra-mcp-server",
        "--allowed-dir",
        "/absolute/path/to/your/documents"
      ],
      "env": {
        "NODE_ENV": "production"
      }
    }
  }
}
```

### Cursor IDE
Add to `.cursor/mcp.json` in your workspace root:

```json
{
  "mcpServers": {
    "astra": {
      "command": "npx",
      "args": [
        "-y",
        "@attentionform/astra-mcp-server",
        "--allowed-dir",
        "${workspaceFolder}"
      ]
    }
  }
}
```

### Windsurf & Cline
- **Command:** `npx`
- **Args:** `["-y", "@attentionform/astra-mcp-server", "--allowed-dir", "/path/to/documents"]`

---

## 2. Security & Sandboxing Model

ASTra restricts the server to the folders you name, so an AI client cannot read or write other files on your computer. If a path is outside those folders, the request is refused and nothing happens:

- **Directory Allowlist:** The server will only read or write `.docx` files located inside directories explicitly declared via `--allowed-dir <path>` (repeatable) or the `ASTRA_EXTRA_DOC_ROOTS` environment variable. The named folders are the whole boundary: the system temp folder is not added. A symbolic link at or above a named folder is followed; a link below it that points outside it is refused.

  When no folders are named, the default roots are `./documents`, `./documents/Drafts`, the system temporary directory, and its `ASTra` subdirectory. The startup line reports the active roots.
- **Path Canonicalization:** All incoming file paths are resolved to their real canonical paths. Path traversal attempts (e.g. `../../etc/passwd` or hidden directory symlinks) are caught and refused with error code `E_SANDBOX`.
- **Absolute Paths Required:** Because clients often spawn the MCP server from disparate working directories, `--allowed-dir` must always be an absolute path and must exist prior to server launch.

### 2.1 Behavior Shared by the CLI and the MCP Server

- **CLI `sanitize` vs MCP `sanitize_document`:** CLI `sanitize` is a one-shot raw scrub directly on archive package parts. In contrast, MCP `sanitize_document` followed by `save_document` executes through the full document save pipeline. The resulting file bytes can therefore differ between the CLI and MCP. The defaults also differ: the CLI presets (bare `sanitize`, `--client-safe`, `--metadata-only` and `--redact-pii`) accept all tracked changes and remove all comments, while `sanitize_document` keeps tracked changes and comments, with the text of every tracked deletion, unless `acceptAllRevisions` or `stripComments` is passed.
- **Invalid input is refused up front:** Both `sanitize_document` and PDF `export_document` refuse invalid or schema-corrupted input upfront with `E_VALIDATION`. Documents that fail schema validation are never scrubbed or exported.
- **A save never adds a validation error:** Saving is refused with `E_VALIDATION`, and the original file on disk is left untouched, when the result has a schema (ISO/IEC 29500) validation error that the file did not already have. A file that already had validator errors can be opened and saved as long as the save adds none. Strict mode (`strict` / `--strict-validation` / `ASTRA_STRICT_VALIDATION=1` on the MCP server, `--strict` on the CLI) refuses every save while the result still has validator errors.
- **Spacing and indent units:** `spacingBefore`, `spacingAfter`, `indentLeft`, `indentRight`, `indentFirstLine` and `indentHanging` are in twips (20 per point, so `240` is 12 pt). A string in points such as `"12pt"` is accepted and converted; any other unit is refused. `lineSpacing` is a separate unit, in 240ths of a line (`240` = single, `360` = 1.5).
- **Table header bold:** `insert` with `type:'table'` makes the header row bold by default. Pass `headerBold:false` (flat or in `params`) to turn it off, or set `- Header Bold: false` under `## Tables` in a style guide. An explicit value beats the guide.
- **Unknown settings are refused:** An operation that carries a setting it does not read, for example `bold` inside `params` of a table `updateCell`, is refused with `E_VALIDATION` before anything is applied. The message names the setting and the ones the operation accepts: `unknown parameter "bold" for table updateCell; allowed: …`.
- **Style operation validation:** A `style` operation in `execute_batch` must set at least one style property, and border configurations read only `style`, `sz`, and `color`. Submitting a style operation without properties or with unrecognized attributes is refused with `E_VALIDATION`.
- **History safety check:** History tools (`get_document_history`, `preview_revert`, `revert_document`, `redo_document`) strictly refuse any history directory whose `.git` entry is a redirect file with error code `E_IO`.
- **Batch execution semantics & session node IDs:** Batches execute sequentially with transactional all-or-nothing rollback (`rollbackOnError: true`). If an operation fails, preceding mutations roll back and the in-memory document state is restored. Take node IDs from a fresh `inspect_document` or `search_content` call; IDs of nodes created during a session (`_bk_new_...`) are valid only within that session: they survive a failed batch, a dry run and a failed single operation (ids and versions are restored; if they cannot be, the error says so), but not closing and reopening the file. Sections are numbered by position when a file is opened (`sec_root`, `sec_2`, `sec_3`, …); after each batch they are rebuilt from the document's section breaks if a break was deleted, moved or added, and the batch result carries a warning.
- **Equations & inline math options:** Equation accents and inline math options (`dollarMath: true` on markdown composition): `$...$` syntax is treated as inline math only when `dollarMath` is explicitly enabled; display equations (`$$...$$`), standard LaTeX inline delimiters (`\(...\)`), and equation accents are always parsed and compiled into native Office Math.

---

## 3. The Standard Agent Workflow

Autonomous agents should interact with documents through this canonical lifecycle:

```text
1. open_document (or new_document)
   └── Returns sessionId and initial document structure summary.

2. inspect_document
   └── Emits Token-Optimized Outline Notation (TOON) and node IDs (data-astra-id).

3. search_content (Optional)
   └── Pinpoints exact text occurrences and returns target node IDs.

4. execute_batch (or compose_document)
   └── Executes atomic array of operations with transactional rollback on error.

5. visual_inspect / export_document (Optional)
   └── Verifies layout telemetry, renders page images, or exports to PDF/Markdown.

6. save_document
   └── Runs ECMA-376 pre-write schema validation gate and commits changes to disk.

7. close_document
   └── Evicts in-memory session and releases file locks.
```

---

## 4. Complete 21-Tool Reference

### 4.1 Document Management

#### `open_document`
Opens an existing `.docx` file into an isolated in-memory editing session.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `path` | `string` | Yes | Absolute path to the `.docx` document. |
| `sessionId` | `string` | No | Server-minted session id (deprecated to supply; test-only). |
| `summaryOnly` | `boolean` | No | The default and `true` return the compact summary; `false` returns the full structural summary including styles XML. |
| `allowMacros` | `boolean` | No | Preserve macro/VBA project parts byte-for-byte. |
| `stripMacros` | `boolean` | No | Automatically strip macro parts upon loading. |
| `author` | `string` | No | Default author (≤ 128 chars) for tracked changes and comments in this session that do not name one. Falls back to the server's `ASTRA_AUTHOR`, then `ASTra`. An operation's own `author` always wins. Not the version-history author of `save_document`. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_9a8b7c6d-5e4f-3a2b-1c0d-e9f8a7b6c5d4",
  "revision": 1,
  "summary": {
    "summaryOnly": true,
    "counts": {"nodes": 42, "sections": 1, "headers": 0, "footers": 0, "charts": 0},
    "headers": [], "footers": [], "chartIds": [], "hasTrackedChanges": false
  },
  "evicted": []
}
```

---

#### `new_document`
Creates a fresh `.docx` file from a verified ECMA-376 template and opens it into a session.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `outPath` | `string` | Yes | Target destination file path inside an allowed directory. |
| `template` | `string` | No | Template type: `"report"`, `"memo"`, or `"blank"` (default: `"blank"`). |
| `force` | `boolean` | No | Overwrite file if it already exists (default: `false`). |
| `styleGuide` | `string` | No | Bundled guide name or guide `.md` path to attach. |
| `author` | `string` | No | Default author (≤ 128 chars) for tracked changes and comments in this session that do not name one. Falls back to the server's `ASTRA_AUTHOR`, then `ASTra`. An operation's own `author` always wins. Not the version-history author of `save_document`. |
| `sessionId` | `string` | No | Server-minted session id (deprecated to supply; test-only). |

**Success Response:**
```json
{
  "ok": true,
  "path": "/path/to/documents/quarterly_review.docx",
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "revision": 1,
  "template": "report",
  "bytes": 28410
}
```

---

#### `save_document`
Runs pre-write ECMA-376 schema validation and commits all modified XML parts to disk. Also records an append-only commit in local history (`.astra/`).

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `outPath` | `string` | No | Save destination (defaults to original session path). |
| `message` | `string` | No | Summary message for version history (≤ 500 chars). |
| `author` | `string` | No | Author name for version history (≤ 100 chars). |
| `allowMacros` | `boolean` | No | Explicit macro policy to keep macros on save. |
| `stripMacros` | `boolean` | No | Explicit macro policy to strip macros on save. |
| `strict` | `boolean` | No | Refuse a save with any validator errors, including errors already present when opened. Defaults to the server setting. |

The default save policy permits an already-invalid file when no additional normalized error signatures are detected; the `validation` result reports tolerated errors. Strict mode can be enabled per call, by `--strict-validation`, or by `ASTRA_STRICT_VALIDATION=1`. The comparison counts errors by normalized signature, so fixing one error and introducing another with the same signature within one save can be missed; strict mode refuses both.

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "path": "/path/to/documents/quarterly_review.docx",
  "revision": 3,
  "commitOid": "8f3b2a1c9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c4b"
}
```

---

#### `close_document`
Closes the active session and releases memory. A dirty close warns instead of silently discarding:
the success envelope carries `dirty` plus `discardedRevisions` and, when dirty, a `warning`
naming the discarded unsaved revision count. Save first when the warning matters.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "dirty": true,
  "discardedRevisions": 2,
  "warning": "close_document discarded 2 revision(s) of unsaved edits with this session. Save via save_document before closing to keep edits."
}
```
(A clean close returns `"dirty": false, "discardedRevisions": 0` with no `warning`.)

---

### 4.2 Inspection & Query

#### `inspect_document`
Returns document structure using Token-Optimized Outline Notation (TOON). Which control to use to read less is in [Reading a document cheaply](reading-documents.md).

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `fidelity` | `string` | No | `"outline"` (headings only), `"compact"`, `"standard"` (default), or `"full"`. With `"outline"`, `rowCount` and `totalNodes` count heading rows, and the response adds `outline` (`headings`, `nodes`, `tables`, `images`, `words` for the whole document). `aroundNode` on a node that is not a heading shows the heading above it and reports it as `aroundResolvedTo` (`null` before the first heading, with a `note`). |
| `sectionId` | `string` | No | Filter output to a specific section ID. |
| `maxTextLength` | `number` | No | Truncate run text at character limit. |
| `offset` | `number` | No | Pagination start row for document nodes only (default: 0). An offset past the last node is rejected with `E_VALIDATION`, with or without `include`. |
| `limit` | `number` | No | Maximum node rows to return (default: 50, max: 200). |
| `aroundNode` | `string` | No | Node ID to center the inspection window around. |
| `window` | `number` | No | Half-width of rows around `aroundNode` (default: 5). |
| `maxTokens` | `number` | No | Cap output on a clean row boundary (200–8000 tokens). |
| `include` | `string[]` | No | Add `comments`, `revisions`, or both to the response. |
| `reviewOffset` | `number` | No | Start record for the `include` lists (default: 0). |
| `reviewLimit` | `number` | No | Records per `include` list (default: 50; larger values are clamped to 50). |
| `commentId` | `number` | No | Requires `include:["comments"]`; return that comment and its replies in full. |
| `revisionId` | `number` | No | Requires `include:["revisions"]`; return every element sharing that ID and report ambiguity. |

Review lists page with `reviewOffset` and `reviewLimit`, independently of the node window set by `offset` and `limit`. They return at most 50 records and include `nextOffset` when more are available. A `commentId` without `include:["comments"]`, or a `revisionId` without `include:["revisions"]`, is rejected. Comment previews are capped at 200 characters and revision previews at 60; the targeted read returns full text.

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "revision": 1,
  "fidelity": "compact",
  "rowCount": 12,
  "palette": ["Heading1", "Normal", "TableGrid"],
  "toonText": "ID | V | Path | Header | Style | Text | Flags\n--- | --- | --- | --- | --- | --- | ---\n6C54C7A6 | 1 | Root |  | Title | Q3 Executive Financial Performance Report | \n3FC95CAC | 1 | Root |  | Heading1 | 1. Executive Summary | ",
}
```

---

#### `search_content`
Fast text search across paragraphs, table cells, and runs.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `query` | `string` | Yes | Search query string (1–500 characters). |
| `regex` | `boolean` | No | Enable regular expression search (default: `false`). |
| `scope` | `string` | No | `"body"` (default), `"comments"`, or `"all"`. Comment matches use IDs such as `comment:0`; comment operations accept `comment:0` and `0` alike. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "matches": [
    {
      "id": "3FC95CAC",
      "preview": "1. Executive Summary"
    },
    {
      "id": "t_001",
      "preview": "RegionRevenueNorth120South98",
      "cells": ["t_001:r3c1"]
    }
  ]
}
```

A table hit also lists `cells`: up to five addresses of the cells that hold the match, in reading order. Pass one as the `target` of an edit (see **Table cells** under `execute_batch`).

A result holds at most 50 matches. When more exist the response carries `"truncated": true`; narrow the query or search a smaller scope to see the rest.

---

### 4.3 Mutation & Composition

#### `execute_batch`
Executes an atomic array of typed operations (1–500 ops) with schema validation and rollback on error.

**Inline formatting in inserted table cells.** Cell strings accept inline Markdown: `**urgent**` makes bold text, `*note*` makes italic text, and `[site](https://example.com)` creates a link. Ordinary text such as `*.docx` and `5 * 2` stays literal. To display Markdown delimiters literally, escape each delimiter. For example, in a JSON operation, use `"rows": [["\\*\\*urgent\\*\\*"]]` to display `**urgent**` as unformatted text; each Markdown backslash is doubled for JSON. Backticks create code formatting.

**Table cells.** The paragraph in a table cell has no id of its own, so name it as `<tableId>:r<row>c<col>` (for example `t_001:r3c2`). Rows and columns are 1-based, exactly as the `r1c1` grid in `inspect_document` prints them. A cell with several paragraphs needs `:p<n>` as well (`t_001:r3c2:p2`); the error lists them. The `replace`, `style`, `namedStyle`, `clearFormat`, `hyperlink`, `bookmark` and `comment` actions accept a cell address as `target`; every other action refuses one and names the fix (the `table` action takes the table id with `rowIndex`/`colIndex`, numbered the same way, so `t_001:r3c2` and `rowIndex: 3, colIndex: 2` are the same cell). Every row and column number in a table operation counts from 1: `rowIndex`, `colIndex`, `columnIndex`, `row`, `col`, `index`, `afterIndex` and `fromRow`/`toRow`/`fromCol`/`toCol`. Address the cell where a merge starts: a covered cell, or a column inside a `gridSpan`, is refused with the cell to use instead. `expectedVersion` checks the table's version, and each cell edit bumps it.

**Numbers count from 1.** Every position you type or read counts from 1: table rows and columns (`rowIndex`, `colIndex`, `r1c1`), chart series (`trendline.series`), sections (`sectionIndex`), list levels (`level`, shown as `[L1]`; on `changeLevel`, `0` takes the paragraph out of the list), pages and heading levels. `offset` is the one exception: it is how many items to skip, so `0` starts at the beginning. Values the file format fixes, such as the raw `w:ilvl` or `c:idx` inside a document, stay zero-based and are converted for you.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `operations` | `array` | Yes | Array of canonical operation objects. |
| `sequential` | `boolean` | No | Apply in document order (default: `true`). |
| `rollbackOnError` | `boolean` | No | Restore pre-batch snapshot on failure (default: `true`). |

**Supported Operations (27 total):**
`insert`, `replace`, `delete`, `style`, `clearFormat`, `move`, `table`, `section`, `pageSetup`, `namedStyle`, `field`, `equation`, `list`, `headerFooter`, `chart`, `image`, `callout`, `footnote`, `endnote`, `watermark`, `pageBorder`, `coverPage`, `hyperlink`, `bookmark`, `comment`, `trackedChange`, `metadata`.

**Review operations.** `trackedChange` supports `acceptAll`, `rejectAll`, and selective `accept` or `reject` using exactly one of `ids`, `author`, or `target`. ID and node selections are atomic; an author selection can leave unsupported changes and reports them in `revisionResolution.unresolved`. A duplicated ID is ambiguous unless it represents the two sides of a move; select by author or node instead. Tracked edits return `revisionIds` so the complete edit can be selected later. `trackedChange` deletion can target a text span and takes a 1-based `occurrence` when it repeats. Comment operations support `insert`, `reply`, `resolve`, `unresolve`, and `delete`; deleting a comment also deletes its replies and refuses inconsistent anchors. A thread is identified by its last paragraph, as in Word, so comments with several paragraphs and comments written by Word 2007 (which has no extended comment parts) work with every comment operation. When the last comment is deleted, the comment parts are removed from the package, as Word does.

**Resolving whole paragraphs.** `acceptAll`, `rejectAll` and the selective forms resolve block-level revisions as Word does: a whole-paragraph deletion or insertion removes or keeps the paragraph, a whole-paragraph move (the paragraph mark carries `w:moveFrom` or `w:moveTo`) is resolved at both ends together, and a table of contents inserted as a tracked block (a block `w:sdt` between `customXmlInsRangeStart` and `customXmlInsRangeEnd`) is kept or removed as one unit. When a paragraph that is removed holds a bookmark that lies wholly inside it, the bookmark goes with it and is listed in `revisionResolution.removedBookmarks`; a link that still points at it is reported in `staleLinks`, as Word leaves a stale table-of-contents entry. A comment thread whose anchors all lie inside removed text is deleted with it (`removedComments`). When a comment is also anchored outside the removed text, the paragraph stays unresolved and is listed in `unresolved`. When a bookmark of a moved paragraph ends in the neighbouring paragraph, it is moved there, as Word does.

**Pending and unresolved counts.** `revisionsPending` is the number of revision elements left in the document. It counts every element, so one move counts as each of its markers (the two paragraph-mark markers and the wrapper elements), and it is larger than the number of changes a reviewer sees in Word. After a `trackedChange` resolution, `revisionsUnresolved` counts how many of them the resolution could not resolve, and `revisionsRemedy` lists their kind, node and reason. In that case the remedy does not tell you to run `acceptAll` again, because a second run cannot resolve them. Without a resolution in the batch the remedy names `acceptAll`, as before.

**Resolution speed.** Accept and reject operations in a run refresh the session's view of the document once, before the next operation of another kind and at the end of the batch. Each operation still sees the document as the previous one left it. A block that the run changed has its version raised once for the run.

The resolver reports revision kinds it cannot safely apply. In particular, accepting `cellDel`, rejecting `cellIns` or `cellMerge`, rejecting `numberingChange`, and incomplete moves can remain unresolved. Use Word for a final visual check and for revisions reported as unresolved. `sanitize_document` can remove deleted cells when accepting all revisions. With `acceptAllRevisions`, `sanitize_document` is fail-closed: when any revision markup remains after accepting (unresolvable kinds, an unpaired move, or revisions the resolver does not know), it refuses with `E_VALIDATION`, names the part and the kinds, and returns no file, because a partly cleaned file could still contain deleted text and reviewer names and dates. Accept or reject those revisions in Word and sanitize again.

**Edits that ASTra refuses.** An edit is refused with `E_VALIDATION`, and the document is left unchanged, in these cases. The message names the reason and a remedy that works.

- **Inline content controls and equations.** An edit, tracked or not, whose changed text touches an inline content control (`w:sdt`) or an inline equation (`m:oMath`) is refused. In the text an agent reads, an equation appears as `\(latex\)`. Edit the text outside the control or equation, which works with and without `trackChanges`, delete the whole paragraph with `trackChanges: false`, or make the change in Word. A tracked delete of such a paragraph is refused.
- **Simple fields.** A tracked `delete`, or a tracked `replace` whose changed text covers a complete simple field (`w:fldSimple`), rewrites that field as the complex field Word writes for it (begin, instruction, separate, result, end) and tracks each part as a deletion, because a simple field cannot be marked as deleted. The result text, its formatting and the instruction are kept, and the document text stays the same. Accept removes the whole field. Reject brings the field back as a complex field with the same instruction and result; the original `w:fldSimple` markup does not return, and Word treats the two forms as the same field. The edit is refused, with the document unchanged, when the field result holds anything but plain runs (a link, a tracked change, a bookmark) or when the changed text covers only part of the field (the message then says to change all of the text of the field in the same edit, or to edit the text outside it). A replace whose new text keeps a word or a space of a multi-word field, such as a date, splits into several changes and counts as covering part of the field. Untracked edits never convert a field, and tracked edits outside the field leave it as it was. Whole complex fields can be deleted as tracked changes when their full span is selected.
- **Fields that span paragraphs.** The paragraph that holds the start or the end of a field that spans paragraphs, such as a table of contents, cannot be edited with `replace` or a tracked `delete`. Paragraphs between the start and the end can be edited. Edit text outside the field, or update the field in Word.
- **Partial moves.** A paragraph that holds part of a tracked move (text moved within or out of the paragraph) cannot be edited. Accept or reject the move with `trackedChange` first. If that leaves the move unresolved (see below), resolve it in Word.
- **Ruby text (furigana).** An edit whose changed text would cut or remove a `w:ruby` run is refused. Edits around the ruby text work, and the ruby text is kept when you accept or reject changes. Edit the text outside the ruby text, delete the whole paragraph with `trackChanges` false, or make the change in Word.
- **Other content the editor does not model.** A paragraph with a child element or a run with a child element that the editor does not know is refused, and the message names the element. Left-to-right and right-to-left wrappers (`w:bdo`, `w:dir`), `w:customXml`, and range-permission markers (`w:permStart`, `w:permEnd`) are known and can be edited. The wrapper or marker is kept.

**Revisions that stay unresolved.** `revisionResolution.unresolved` lists what an accept or reject could not resolve, with the reason.

- Accepting a deleted paragraph whose bookmark starts or ends in another paragraph leaves it unresolved. Rejecting it restores it. A bookmark that lies wholly inside the removed paragraph goes with it.
- Rejecting `customXml` range markers around deleted or moved content is reported as unresolved. Accepting them works.
- A move whose markers are incomplete or ambiguous stays unresolved in both directions.

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "appliedCount": 3,
  "revision": 2,
  "createdNodeIds": ["_bk_new_6e7accf7"]
}
```

---

#### `compose_document`
Compiles a Markdown document into valid OOXML structures and applies them to the session.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `markdown` | `string` | Yes | Markdown content string (up to 500,000 characters). |
| `allowedDir` | `string` | No | Directory for resolving relative image and chart files. |
| `template` | `string` | No | Target template to replace initial placeholders (`"report"`, `"memo"`, `"blank"`). Accepted on the first compose of a session only. |
| `styleGuide` | `string` | No | Bundled guide name or guide `.md` path; attaches to the session. Omitted: the built-in Default Corporate look (navy table header, bold white header text, striped rows). `"legacy"`: the older grey header, for this call only. |
| `anchor` | `string` | No | Where the composed blocks go: a node id, `"start"` or `"end"`. Omitted: after the last node. |
| `position` | `string` | No | `"before"` or `"after"` (default) the `anchor`. Needs `anchor`. |

A document can be built in several compose calls. The first compose of a session mints ids `md_0001…`; a later compose mints `md2_0001…`, then `md3_0001…`, so ids never collide, and `persistentIds` covers only the nodes of that call. A later compose appends at the end, or goes at `anchor` with `position`, so an agent can add a section in the middle without another tool. The blocks keep their order: only the first is placed against the anchor. A compose is atomic as a whole: if any part fails, the document stays as it was. Citation numbers and the References list restart in each compose call. A chart cannot be the first block of a compose that goes `before` a node.

An italic-only line directly under the opening `#` title becomes the `Subtitle` style. A `[!NOTE]`-style line outside a blockquote stays plain text and is reported in `warnings`. A `> [!NOTE]` alert becomes a callout and keeps bold, italic, code and links. Any other `> quote` becomes paragraphs in the `Quote` style (blank `>` lines start a new paragraph, `> >` nests up to three levels, flat lists and fenced code stay inside); a table, heading or image inside a quote fails with a line number.

**Supported Markdown Matrix:**
| Markdown Element | Generated OOXML Construct |
|---|---|
| `#`, `##`, `###` | `Heading1`, `Heading2`, `Heading3` paragraph styles |
| Body text | `Normal` style paragraphs |
| `**bold**`, `*italic*`, `` `code` `` | Styled `<w:r>` runs with `w:b`, `w:i`, or monospace font |
| `[text](url)` | OOXML `<w:hyperlink>` with relationship mapping |
| `- item` or `1. item` | Numbered & bulleted list hierarchies (`w:numPr`) |
| `\| col \| col \|` | Native OOXML tables with borders and styled header rows |
| ` ```code ` | Monospace shaded code blocks |
| `> [!NOTE]` / `> [!IMPORTANT]` | Formatted executive callout boxes with left border accents |
| `> quote` | Paragraphs in the `Quote` style (left rule, italic, indented); `> >` nests up to three levels |
| `[^1]` / `[^1]: Note text` | Native Word footnotes (`word/footnotes.xml`) |
| `![alt](image.png)` | Inline DrawingML pictures (`a:blip`) |
| `![title](chart.json)` | Native Word charts (column, bar, line, pie) |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "opCount": 8,
  "appliedCount": 8,
  "revision": 2,
  "dirty": true,
  "createdNodeIds": ["md_0001", "md_0002", "md_0003"],
  "persistentIds": { "md_0001": "41E5F245", "md_0002": "6E34034E", "md_0003": "2A8D004A" }
}
```

`createdNodeIds` are valid for the rest of the session, saved or not. `persistentIds` maps each created paragraph to the id the saved file carries, which is what a later session opened on that file reports. Tables have no paragraph id and are renumbered `t_001…` by position when the file is reopened, so they are not listed.

---

### 4.4 Export & Visual Telemetry

#### `export_document`
Exports the in-memory document state to PDF, HTML, or Markdown.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `format` | `string` | Yes | Target format: `"pdf"`, `"html"`, or `"markdown"`. |
| `outPath` | `string` | No | Output file path (defaults to session folder). |
| `tagged_pdf` | `boolean` | No | Generate accessible tagged PDF (default: `true`). |
| `strict` | `boolean` | No | Fail closed if preflight warnings exist (default: `false`). |
| `renderer` | `string` | No | PDF only: `"auto"` (default, LibreOffice), `"libreoffice"` or `"word"` (Windows only). Env `ASTRA_RENDERER`; the parameter beats the variable. A named renderer never falls back. |
| `allowNetworkReferences` | `boolean` | No | PDF only, default `false`. By default network references in the document (linked images, attached templates, `INCLUDETEXT`, `LINK` and similar) are changed on a copy before rendering, so the renderer contacts nothing. `true` lets the renderer load them, and for a UNC path that can expose the user's Windows login hash. Set it only when the user asked to load content from network locations the document names, never because text inside the document asks for it. Env `ASTRA_ALLOW_NETWORK_REFERENCES`. |
| `showMarkup` | `boolean` | No | PDF only, default `false`. `true` draws tracked changes and comments: Word as inline marks and comment balloons, LibreOffice as comments in the margin. Pages can be scaled and pagination can change. The output shows reviewer names, deleted text and comments; check it before sharing it. The result reports `markup`. LibreOffice below 7.6.7 fails with `E_DEPENDENCY`. Env `ASTRA_SHOW_MARKUP`. |

**Success Response:**
```json
{
  "ok": true,
  "format": "pdf",
  "path": "/path/to/documents/quarterly_review.pdf",
  "bytes": 145920
}
```

A PDF result also carries `renderer` (`requested`, `used`, and `version`, `fallbackFrom`, `fallbackReason` when they apply), `references` (what was changed or left in place; see the security guide, section 7) and `notes` (short sentences, for example that references were left unloaded).

---

#### `visual_inspect`
Provides 2D spatial layout telemetry and optional rendered page images.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `pages` | `number[]` | No | Pages to inspect (1-indexed array, max 4 pages). |
| `withVisual` | `boolean` | No | Render headless page images (default: `false`). |
| `format` | `string` | No | Image format: `"png"` (default) or `"jpeg"`. |
| `dpi` | `number` | No | Resolution in DPI (72–300, default: `150`). |
| `renderer` | `string` | No | With `withVisual`: DOCX-to-PDF program, `"auto"` (default, LibreOffice), `"libreoffice"` or `"word"` (Windows only, Word 2010 or later). Env `ASTRA_RENDERER`; the parameter beats the variable. A named renderer never falls back; `auto` falls back visibly. |
| `pdfToImage` | `string` | No | With `withVisual`: PDF-to-image program, `"auto"` (default: the Windows built-in renderer first on Windows, then Poppler; Poppler elsewhere), `"windows"` or `"poppler"`. Env `ASTRA_PDF_TO_IMAGE`. |
| `allowNetworkReferences` | `boolean` | No | With `withVisual`, default `false`. Same meaning as in `export_document`: `true` lets the renderer contact network locations the document names. Env `ASTRA_ALLOW_NETWORK_REFERENCES`. |
| `showMarkup` | `boolean` | No | With `withVisual`, default `false`. Same meaning as in `export_document`: draws tracked changes and comments in the page images. The output shows reviewer names, deleted text and comments. Env `ASTRA_SHOW_MARKUP`. |
| `outputDir` | `string` | No | With `withVisual`, write the images to this folder (created if missing) as `<document>-page-<n>.png` (a file already under that name is replaced; a symbolic link at the name is refused), for clients that cannot read the server's temp folder. Must be inside the allowed folders, like a document path. Default: the server's temp artifacts folder. |

The wireframe shows tables as `[TABLE: n cols]` followed by the header row and the first rows, list items with `- ` or `1.` markers, and a shaded one-cell table as `[CALLOUT]`.

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "telemetry": {
    "pageCount": 3,
    "dimensions": { "widthTwips": 12240, "heightTwips": 15840 },
    "margins": { "top": 1440, "bottom": 1440, "left": 1440, "right": 1440 },
    "wireframe": "+----------------------------+\n| [Heading 1] Exec Summary   |\n| [Normal] Net revenue...    |\n+----------------------------+"
  },
  "images": [
    { "page": 1, "path": "/tmp/astra-visual/page_1.png" }
  ],
  "renderer": { "requested": "auto", "used": "libreoffice" },
  "rasterizer": { "requested": "auto", "used": "poppler" },
  "references": { "networkReferencesAllowed": false, "omissions": false, "omittedCount": 0, "omittedReferences": [], "localCount": 0, "notInspectedCount": 0 },
  "notes": []
}
```

`renderer`, `rasterizer`, `references` and `notes` appear with `withVisual`. `markup` (`requested`, `renderer`, `mode`, `warnings`) appears when a render ran; it records what was asked and which mode the renderer was told to use, not what the image shows. `renderer`, `pdfToImage`, `allowNetworkReferences` and `showMarkup` without `withVisual` are refused as usage errors, and `showMarkup` with a format other than `pdf` in `export_document` is refused. `notes` says when a program fell back (`fallbackFrom`, `fallbackReason` on `renderer` or `rasterizer`) and when network references were changed.

---

### 4.5 History & Time-Travel

ASTra tracks document version history in a local, append-only sidecar repository located at `<docDir>/.astra/<file>/` (with `<docDir>/.astra-tombstones.json` recording which documents have history). Set `ASTRA_HISTORY_DIR` in the server's environment to keep both under that directory instead; see the CLI guide, section 4.12, for the layout and a `.gitignore` snippet.

#### `export_original_document`

Export a captured original byte for byte: `{sessionId, outPath, overwrite?}`. The destination must be inside an allowed directory. Existing destinations require `overwrite: true`. Export does not change the open document.

#### `restore_original_document`

Restore the captured original to the open document: `{sessionId, discard?}`. Unsaved edits are refused unless `discard: true`; a changed file or inconsistent history is refused. The operation records a new history state. Ordinary revert and redo restore document state, which may not reproduce the original ZIP bytes.

Original capture is automatic when an existing document first receives a new MCP history. Older histories may have no original. Sanitizing a document does not remove sensitive content from its history or captured original. When `save_document` writes to a new path that already holds a document with no history, the bytes it overwrote become that history's original. `get_document_history` reports `original.available`, or the reason it is unavailable. A session that opens while another process is still creating the history reports `reason: "initializing"`; it becomes eligible for export and restore only after that initialization finishes and the stored original has the SHA-256 of the file this session opened, and a different original never makes it eligible. First history creation uses an exclusive `.astra-init.lock` in the history folder on filesystems that honor exclusive file creation. If a killed process leaves it, `save_document`, `revert_document`, `redo_document` and `restore_original_document` refuse before touching the file and name the lock (with the creating process id and host name when known); ASTra never removes it or waits for it, and clearing it by hand is described in the CLI guide (`docs/cli-guide.md`, "Interrupted captures and a stuck initialization lock"). Two ASTra processes saving to the same document are not coordinated after the history exists, so one can lose an edit; give each document one writer at a time. Do not run ASTra 0.4.x and 0.5.x on the same document. Network filesystem coordination is not guaranteed.

---

#### `get_document_history`
Retrieves commit history for the current document.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `limit` | `number` | No | Maximum entries to return (1–100, default: 20). |
| `offset` | `number` | No | Number of entries to skip (default: 0). |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "documentPath": "/path/to/documents/quarterly_review.docx",
  "currentCommit": "8f3b2a1c9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c4b",
  "totalCommits": 3,
  "history": [
    {
      "oid": "8f3b2a1c9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c4b",
      "message": "Update Q3 metrics and add ARR table",
      "author": "ASTra Agent",
      "timestamp": "2026-10-13T12:00:00.000Z",
      "timestampMs": 1760356800000,
      "opCount": 2,
      "affectedNodeIds": ["3F2A9C1B"],
      "resultingHeadOid": "8f3b2a1c9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c4b"
    }
  ]
}
```

History metadata is consistent across entries: `timestampMs` is always ms-precision epoch millis
(commit-time clock for tracked commits, git seconds for older ones); `opCount` is always present
(`0` for pre-tracking commits and heals-only saves; it folds operations, not normalizer rewrites);
`affectedNodeIds` is always present (`[]` when untracked); `resultingHeadOid` always equals `oid`
(append-only; refs never move).

---

#### `preview_revert`
Performs a read-only dry-run check of reverting to a previous commit without modifying state.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `steps` | `number` | No | Saved commits to walk back from the head: `1` = the previous saved commit (must be `< totalCommits`; default: `1`). |
| `commitOid` | `string` | No | Target commit SHA (mutually exclusive with `steps`). |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "targetOid": "1c0b9a8f7e6d5c4b8f3b2a1c9e8d7c6b5a4f3e2d",
  "stepsBack": 1,
  "changedValuesPositional": false,
  "snapshotComplete": true,
  "dirty": false
}
```

`snapshotComplete: true` means the target commit recorded every package part, so a real revert restores
the whole package (comment metadata, document properties and custom XML included) and removes parts the
target does not have. `snapshotComplete: false` means the commit was recorded by an ASTra version before
0.4.2, or could not be proven complete: the response then carries `restoreWarning`, and a real revert
restores only what that commit recorded and leaves every other part as it is in the current file.

`changedValuesPositional: true` means either side held orphan ids (no shared id across the boundary,
common under duplicate native ids), so before/after pairing is positional, not identity. A dirty
preview reports `dirtyHandling` naming the pending revisions; net-zero edits still count as dirty.

---

#### `revert_document`
Restores the document state to a previous commit, recording a new append-only restoration commit.
Bare reverts fail closed with `E_CONFLICT` when the session holds unsaved edits: net-zero edits still
count as dirty (the flag tracks pending revisions, not a content diff); save first or pass
`discardUnsaved: true`. The conflict carries `dirtyReason` with the pending revision count. A bare
default revert (no author/reason/message) quotes the target commit message in the new subject
(`Revert to <oid>: "<target message>"`).

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `steps` | `number` | No | Saved commits to walk back from the head: `1` = the previous saved commit (must be `< totalCommits`; default: `1`). |
| `commitOid` | `string` | No | Target commit SHA (mutually exclusive with `steps`). |
| `discardUnsaved` | `boolean` | No | Drop uncommitted in-memory changes (default: `false`). |
| `message` | `string` | No | Custom commit message for the restore commit. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "path": "/path/to/documents/quarterly_review.docx",
  "revertedToCommit": "1c0b9a8f7e6d5c4b8f3b2a1c9e8d7c6b5a4f3e2d",
  "revision": 4,
  "dirty": false,
  "snapshotComplete": true
}
```

`snapshotComplete` and `restoreWarning` mean the same as in `preview_revert`.

**Selective middle-undo recipe (no cherry-pick: restore the older keeper, then re-apply the kept newer edits):**
1. `preview_revert` the target to confirm what a restore would write.
2. `revert_document` to the older keeper commit.
3. `inspect_document` / `search_content` for fresh anchors (node ids die on every revert; never reuse pre-revert ids).
4. `execute_batch` to re-apply the kept newer edits.
5. `save_document` to record the result.

**Export-check after every mutation:** run `export_document` (or `inspect_document`) after each revert
and each re-apply and confirm the bytes before the next step.

---

#### `redo_document`
Re-applies the revision undone by the most recent revert. Success carries advisory `changedValues`
(pre-redo live text → restored text, same bounds as save/preview). An empty stack fails with
`E_NOT_FOUND` plus `atLatestState: true`: a fresh `execute_batch` edit clears redo, and each redo
consumes one entry; call `revert_document` first to record a fresh redoable state. A batch that
drops a non-empty redo stack signals it (`redoCleared: true` plus `redoClearedWarning` on the
batch envelope), so the cleared stack never surfaces as a surprise.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `message` | `string` | No | Commit message for the redo operation. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "revision": 5,
  "commitOid": "8f3b2a1c9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c4b"
}
```

---

### 4.6 Document Audit & Maintenance

#### `validate_document`
Performs an upfront schema and integrity check against ISO/IEC 29500 (ECMA-376) standards.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "validation": {
    "valid": true,
    "errorCount": 0,
    "errors": []
  }
}
```

The verdict is on the document as a save would write it: children out of schema order are put in order first. When that save would repair something, the response also carries `pendingNormalizations` (a list such as `Reordered w:tblPr into schema order at /w:document[1]/...`; the same list `open_document` and `inspect_document` return); the field is absent when nothing needs repair.

Validation advice: when `validate_document` reports an invalid document, its `normalizeHint` suggests attempting `normalize_document` and validating again. This is advice, not a repair forecast or a guarantee of validity; validation does not run repair rules to generate the hint. Normalization changes the open session; use `save_document` to persist the result, optionally to a new output path.

---

#### `get_stats`
Calculates document telemetry including word count, paragraph count, and an estimated editorial page budget. Actual pagination depends on the rendering application, fonts, layout and document content; render a PDF when a physical page count is needed.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "stats": {
    "words": 1420,
    "paragraphs": 48,
    "tables": 3,
    "pages": 5
  }
}
```

---

#### `sanitize_document`
Removes personally identifiable metadata, editing session traces, RSID markers and authorship from the document in-memory. Tracked changes and comments are kept, with the text of every tracked deletion, unless you pass `acceptAllRevisions` (accept all tracked changes, which removes the deleted text) or `stripComments`. The CLI presets turn both on; this tool does not. With `acceptAllRevisions` it refuses, with `E_VALIDATION`, a document whose revision markup it cannot fully resolve (see **Review operations**). Accepting a deleted paragraph also deletes the comment threads anchored only inside it, and the comment parts go when no comment is left; a thread that is also anchored outside the deleted text makes it refuse (use `stripComments`).

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "revision": 3,
  "sanitized": true
}
```

---

#### `normalize_document`
Repairs XML namespace declarations, sorts element sequences into schema order, eliminates redundant styling runs, and repairs known validator error kinds one rule at a time. A repair that changes the visible text or adds a validator error is dropped and listed in `rolledBack`. With a history repository, the state before the repairs is committed when it differs from the head, so one `revert_document` returns to it.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "revision": 4,
  "dirty": true,
  "normalized": ["word/document.xml: Unwrapped w:smartTag at /w:document[1]/w:body[1]/w:p[3]/w:smartTag[1]"],
  "remaining": 0,
  "rolledBack": []
}
```

---

#### `import_document`
Imports a foreign or legacy document and converts its layout and typography into the target ASTra style guide.

**Input Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `path` | `string` | Yes | Source document path to import. |
| `styleGuide` | `string` | No | Target style guide name or path. |

**Success Response:**
```json
{
  "ok": true,
  "sessionId": "sess_4f3a2b1c-8e9d-7c6b-5a4f-3e2d1c0b9a8f",
  "imported": true
}
```

---

## 5. TOON Grammar & Flag Legend

Token-Optimized Outline Notation (TOON) compresses OOXML markup into dense, tabular rows:

| Token Flag | Meaning | Example / Output |
|---|---|---|
| `[TBL]` | Table node indicator | `r1c1 Metric \| r1c2 Value ; r2c1 ARR \| r2c2 $12M` |
| `[IMG]` | Image or visual drawing present | `[IMG] logo.png` |
| `[EQ]` | Mathematical equation | `[EQ] a^2 + b^2 = c^2` |
| `[FN]` / `[EN]` | Footnote / Endnote citation | `[FN:1] See Appendix A` |
| `[CMT]` | Document comment present | `[CMT] Review needed` |
| `[TRK]` | Tracked changes revision | `[TRK:ins] Added clause` |
| `[LNK]` | External hyperlink | `[LNK] https://example.com` |
| `[B]` / `[I]` | Bold / Italic styling | `[B] Bold run text` |
| `[C:hex]` | Text color (6-digit hex without `#`) | `[C:4F46E5]` |
| `[SZ:n]` | Font size in half-points | `[SZ:24]` (12pt), `[SZ:28]` (14pt) |
| `[HDR]` | Table header row styled | Appears on row 1 of tables |
| `[ZEBRA]` | Alternating zebra row shading | Present on banded tables |
| `[CH:type]` | Native 2D chart element | `[CH:column] ARR Growth \| cats: Q1,Q2` |
| `(spanN)` | Grid cell spans N columns | `r1c1 Summary(span2)` |

---

## 6. ECMA-376 Units & Formatting Reference

When passing formatting options to `execute_batch` or inspecting TOON telemetry, use native ECMA-376 units:

| Measurement | OOXML Attribute | Unit Conversion | Example |
|---|---|---|---|
| **Font Size** | `w:sz` | **Half-points** (multiply pt by 2) | 11pt = `22`, 14pt = `28` |
| **Spacing & Margins** | `w:before`, `w:after`, `w:left` | **Twips** (1/20th of a point; 1pt = 20 twips) | 12pt = `240` twips, 1 inch = `1440` twips |
| **Border Widths** | `w:sz` on `w:bottom`, `w:left` | **Eighths of a point** (1pt = 8 units) | 1pt border = `8`, 3pt accent = `24` |
| **Colors** | `w:color`, `w:fill` | **6-digit hex without `#`** | Indigo = `4F46E5`, Charcoal = `1E293B` |

### Styling recipes (professional defaults)

- **Charts: distinct series colors.** Multi-series charts default to distinct palette fills per series (no two series share the default blue). Override any series explicitly: `series: [{ name: "ARR", values: [10, 25], color: "10B981" }]` (`color` is 6-digit hex, `#` optional). Explicit colors always win over the palette.
- **Charts: one strict `style` object.** Author legend, number formats, data labels, axes, trendlines and floating placement through `style`, e.g. `style: { legend: { position: "t" }, dataLabels: { show: true, position: "outEnd" }, axes: { value: { numberFormat: "$#,##0", min: 0 } }, trendline: { type: "linear", series: [1] }, floating: { horizontalAlign: "center", wrap: "square" } }`. `trendline.series` lists series numbers counting from 1 in the order the series appear (`[1]` is the first series); omit it to cover every series. Omitted fields preserve the authored node; without `style`, legacy chart output is byte-identical.
- **Charts: combos and secondary axes.** Use `chartType: "combo"` with per-series `chartType: "column" | "line"` and optional `axis: "primary" | "secondary"`: `series: [{ name: "Revenue", values: [10, 20], chartType: "column" }, { name: "Margin", values: [0.2, 0.3], chartType: "line", axis: "secondary" }]`. Secondary axes are supported for line series only; foreign multi-plot parts stay read-only/refused.
- **Tables: two-step header styling.** Step 1 shades the header row: `{ action: "table", target: "<t>", operation: "setShading", fill: "F7F9FC", headerFill: "1F3A5F" }`. Step 2 colors the header text for contrast: `{ action: "table", target: "<t>", operation: "setCellTextColor", rowIndex: 1, color: "FFFFFF" }`. `setCellTextColor` also targets a single cell (`rowIndex` + `colIndex`), a column (`colIndex`), or the whole table (both omitted). Low-contrast fill/text pairs return an advisory warning, never a failure.
- **Tables: keep signature blocks together.** `{ action: "table", target: "<t>", operation: "setKeepTogether" }` marks every row `cantSplit` and chains the rows with `keepNext` so the table moves to the next page intact instead of widowing rows. Manual fallback: keep signature tables short and place a page break before them when they must not split.
- **Code blocks: Code style + fence round-trip.** Fenced code composed from Markdown renders with the template `Code` character style (Consolas/indigo/10pt) plus block shading, a box border, and an indent. Direct inserts get the same treatment with `params: { codeBlock: true }` on an `insert` paragraph op. Fences export back as fences; the fence language round-trips (` ```js ` stays ` ```js `). Blocks composed before this behavior export bare fences.
- **Images: always pass alt text.** `![alt](path)` preserves `alt` into the drawing description; exports render it. Inserts without `altText` succeed with an advisory warning and export a generic image label. Repair later with `{ action: "image", operation: "setAltText", target: "<img>", altText: "..." }`.
- **Lists: numbered stays numbered.** `{ action: "list", operation: "create", target: [...], listType: "numbered" }` builds a decimal list; `continue` joins the preceding list (same numbering sequence) instead of starting a new one; pass `previousNodeId` to name the list explicitly. Inspect labels lists truthfully (`list:numbered:L1` vs `list:bullet:L1`; the number is the list level, counting from 1).
- **A list has nine levels.** A new list definition carries all nine levels, so an item at any `level` (1 to 9) has its marker and indent: bullets cycle `•`, `◦`, `▪` (plain Unicode in Arial, so no symbol font is needed) and numbers cycle `1.`, `a.`, `i.`, each level half an inch further in. A list is reused only when its **top** level is the kind you ask for (a list that is numbered at level 1 and bulleted at level 2 is a numbered list), and never when it belongs to heading numbering or a style link. A list from an older file or from Word that stops short of the level an item is placed at gets the missing levels added to its definition in place; the levels it already has are not changed.


### Behaviour that is deliberate

- **`style` with `match` needs a paragraph.** `match` styles part of a paragraph's text, so `{ action: "style", target: "<tableId>", match: "..." }` is refused (`style 'match' only supports paragraph nodes ('w:p'), got 'w:tbl'`). To style words inside a table, target the cell as `<tableId>:r<row>c<col>`; for a whole-table look use the `table` operations (`setShading`, `setCellShading`, `setKeepTogether`).
- **A callout is a one-cell table.** The note box (`callout`) is written as a 1×1 shaded table with a left border bar, so structure listings show it as a table and a screen reader announces one. The visual wireframe labels it `[CALLOUT]`.
- **Turning the page leaves the margins alone.** `pageSetup` with `orientation` swaps the section's width and height (Word and LibreOffice lay a page out from those, not from the orientation flag), so `landscape` on a Letter page gives 15840 × 12240 twips and `portrait` turns it back. The margins are written exactly as given; Word's own Orientation button also rotates them, and ASTra does not. A `paperSize` given on its own keeps the orientation the section already has, and a new paper size drops the old printer paper code.
- **Paragraphs carry `w14:paraId`.** Word writes a `w14:paraId` on each paragraph and so does ASTra; it is the paragraph id `inspect_document` shows after a save and reopen, and what keeps ids stable between sessions. It is not formatting and is safe to leave in place.

---

## 7. Error Taxonomy & Remediation

All ASTra errors return a standardized JSON envelope with actionable guidance:

```json
{
  "ok": false,
  "code": "E_VALIDATION",
  "message": "Element w:spacing appears before w:pStyle in w:pPr",
  "remedy": "Order child elements strictly according to ISO/IEC 29500: w:pStyle, then w:spacing.",
  "isError": true
}
```

### Reference Error Codes

| Error Code | Trigger Condition | Standard Remediation |
|---|---|---|
| `E_SANDBOX` | Path is outside `--allowed-dir`. | Move document into allowed directory or pass `--allowed-dir` flag. |
| `E_VALIDATION` | Schema violation, invalid parameters, no-op style operation, refusing invalid input to sanitize or PDF export, or an attempt to save an invalid document. | Follow `remedy` instructions to correct parameter types or XML order. |
| `E_SESSION` | Session ID is invalid or has expired. | Call `open_document` to start a fresh editing session. |
| `E_CONFLICT` | Overwrite conflict on file creation or dirty revert. | Pass `force: true` or `discardUnsaved: true`. Net-zero edits still count as dirty (pending revisions, not a content diff); the conflict carries `dirtyReason` with the pending count. |
| `E_DEPENDENCY` | External system tool missing (e.g. LibreOffice). | Install `libreoffice` and ensure it is available in system `PATH`. |
| `E_QUOTA` | Operation batch exceeds maximum allowed steps. | Split operations into batches under 500 items. |
| `E_IO` | File system read or write failure, or history folder whose `.git` entry is a redirect file. | Check file permissions, disk space, or history repository configuration. |
