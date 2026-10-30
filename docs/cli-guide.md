# ASTra CLI Reference & Guide (`astra`)

The official user manual and command-line reference for ASTra.

The `astra` CLI provides a unified toolchain for headless document automation, CI/CD schema verification, batch editing, privacy sanitization, and multi-format export.

---

## Table of Contents

1. [Installation & Setup](#1-installation--setup)
2. [Global Flags & Pipeline Conventions](#2-global-flags--pipeline-conventions)
3. [Exit Code Contract & Shared Behavior](#3-exit-code-contract--shared-behavior)
4. [Command Reference](#4-command-reference)
   - [Document Inspection (`astra inspect`)](#41-document-inspection-astra-inspect)
   - [Schema Validation (`astra validate`)](#42-schema-validation-astra-validate)
   - [Atomic Batch Editing (`astra batch`)](#43-atomic-batch-editing-astra-batch)
   - [Privacy & Sanitization (`astra sanitize`)](#44-privacy--sanitization-astra-sanitize)
   - [Document Creation (`astra create`)](#45-document-creation-astra-create)
   - [Markdown Import (`astra import`)](#46-markdown-import-astra-import)
   - [Multi-Format Export (`astra export`)](#47-multi-format-export-astra-export)
   - [Text Search (`astra search`)](#48-text-search-astra-search)
   - [Structure Normalization (`astra normalize`)](#49-structure-normalization-astra-normalize)
   - [Document Statistics (`astra stats`)](#410-document-statistics-astra-stats)
   - [Visual Page Rendering (`astra visual`)](#411-visual-page-rendering-astra-visual)
   - [Document History (`astra history`)](#412-document-history-astra-history)
   - [Preview Revert (`astra preview-revert`)](#413-preview-revert-astra-preview-revert)
   - [Revert to Recorded State (`astra revert`)](#414-revert-to-recorded-state-astra-revert)
   - [Redo Reverted State (`astra redo`)](#415-redo-reverted-state-astra-redo)
5. [CI/CD & Scripting Recipes](#5-cicd--scripting-recipes)
6. [Style Guides (house skin)](#6-style-guides-house-skin)

---

## 1. Installation & Setup

### Zero-Install via `npx`
Run any command immediately without global installation:

```bash
npx -y @attentionform/astra <command> [options]
```

### Global Installation
Install the standalone `astra` binary onto your system `PATH`:

```bash
npm install -g @attentionform/astra
astra --version
```

---

## 2. Global Flags & Pipeline Conventions

```text
Usage: astra [--json] [--verbose] [--version] [--help] <command> [args]

Commands (exit 0 ok · 1 operational failure · 2 usage):
  inspect  <docx> [--fidelity compact|standard|full] [--section <id>] [--max-text <n>] [--width <n>] [--wrap] [--visual] [--page <n>] [--format png|jpeg] [--max-ratio <n>] [--max-uncompressed-bytes <n>]
  validate <docx> [--max-ratio <n>] [--max-uncompressed-bytes <n>] [--style-guide <name|path>]
  sanitize <in.docx> <out.docx> [--client-safe] [--redact-pii] [--zero-retention] [--replacement-author <name>] [--accept-all-revisions] [--strip-comments] [--metadata-only]
  batch    <docx> [<ops.json|->] [--target-text <needle> --append <text>] [--append <text>] [--out <docx>] [--force] [--sequential|--no-sequential] [--no-validate] [--schema] [--max-ratio <n>] [--max-uncompressed-bytes <n>] [--dry-run] [--backup[=<path>]] [--matrix <vars.csv>] [--reproducible|--no-reproducible] [--author <name>] [--stream [--on-error abort|continue] [--max-errors <n>]] [--style-guide <name|path>]
  normalize <docx> [--out <docx>] [--dry-run] [--backup[=<path>]]
  export   <docx> [--format html|markdown|md|pdf] [-o <outPath>] [--force]
  search   <docx> <query> [--limit <n>]
  create   <out.docx> [--template report|memo|blank] [--force] [--from-md <file.md>] [--allowed-dir <dir>] [--dollar-math] [--style-guide <name|path>]
  import   <file.md> [-o <out.docx>] [--template report|memo|blank] [--force] [--allowed-dir <dir>] [--dollar-math] [--style-guide <name|path>]
  stats    <docx>
  visual   <docx> [--page <n>] --out <render.png> [--format png|jpeg] [--force] [--max-ratio <n>] [--max-uncompressed-bytes <n>] [--allow-macros] [--strip-macros]
  history  <docx> [--init|--track]
  preview-revert <docx> [--steps <n>]
  revert   <docx> [--steps <n>]
  redo     <docx>

Output carries real dates and ids by default. For byte-identical output pass --reproducible (or set ASTRA_REPRODUCIBLE=1 or SOURCE_DATE_EPOCH): timestamps, ids and docProps are then pinned to SOURCE_DATE_EPOCH else 1980-01-01Z. --author <name> sets the default author for tracked changes and comments that do not name one.

Every subcommand answers `--help` (exit 0); `batch --schema` prints canonical ops JSON.
Ops file for batch: { "operations": [ ... ], "sequential": true } or a bare [ ... ] array.
Batch shortcuts (no ops file, no ids): `--target-text <needle> --append <text>` replaces one text-matched node;
  lone `--append <text>` inserts a final paragraph. Both share the ops-file validation/rollback/exit codes.
Single edit = 1-item batch. Failed batches roll back and never write.
Sanitize is fail-closed: bare `sanitize in out` keeps body text (metadata/rsids/authors/comments/revisions still stripped);
  add --redact-pii (or --client-safe bundle) for body PII redaction. Lone --redact-pii carries the metadata preset.
Sanitize --json carries survivors as warnings:[...] (bare and --metadata-only always warn; absent only when fully redacted and clean).
Create/import refuse to overwrite without --force; batch/export/visual --out targets do too (exit 2 naming the file).
Inspect --visual adds spatial telemetry (<500 tok) + one headless page render (--page, PNG default, JPEG opt-in); missing LO/pdftoppm degrades with E_DEPENDENCY hint.
Visual renders the saved on-disk file (MCP visual_inspect renders session state incl. unsaved edits: save first for matching renders); inspect --visual is the same renderer as visual (alias: same --page/--format yields identical bytes).
```

The following flags apply to every subcommand:

| Flag | Description |
|---|---|
| `--json` | Outputs machine-readable JSON on `stdout`. Human text and banners are suppressed. Diagnostics are routed to `stderr`. |
| `--verbose` | Emits detailed timing, parse phases, and debug information on `stderr`. |
| `--version` | Displays current ASTra CLI version and exits with `0`. |
| `--help`, `-h` | Displays detailed usage and flag explanations for any subcommand. |

### Robust Shell Pipelines
When composing shell pipelines with `astra ... --json | jq`, use `set -o pipefail` in bash scripts so that non-zero exit codes from `astra` are preserved:

```bash
#!/usr/bin/env bash
set -euo pipefail

astra validate report.docx --json | jq '.ok'
```

---

## 3. Exit Code Contract & Shared Behavior

Every command uses the same exit codes, so scripts and CI jobs can rely on them:

| Exit Code | Classification | Trigger Scenario |
|:---:|---|---|
| **`0`** | **Success** | Operation completed successfully, document is valid. |
| **`1`** | **Operational Failure** | Target file missing, corrupted zip archive, schema validation failed, batch rolled back, or external tool timeout. |
| **`2`** | **Usage Error** | Unknown flag, invalid argument type, missing required positional, or refusal to overwrite without `--force`. |

### Behavior Shared by the CLI and the MCP Server

- **Invalid documents are never saved:** In every ASTra interface, saving a document that fails ISO/IEC 29500 schema validation is refused with `E_VALIDATION`, and the original file is left untouched.
- **Fail-closed operations:** Operations that mutate or export files (`batch`, `sanitize`, `export`) validate document integrity before writing to disk. Both `sanitize` and PDF `export` refuse invalid or schema-corrupted input upfront with exit code 1 (`E_VALIDATION`).
- **CLI `sanitize` vs MCP `sanitize_document`:** CLI `sanitize` is a one-shot raw scrub operating directly on archive parts. In contrast, MCP `sanitize_document` followed by `save_document` runs through the full document save pipeline, so the resulting file bytes can differ. Both remove the same categories of information.
- **Spacing and indent units:** `spacingBefore`, `spacingAfter`, `indentLeft`, `indentRight`, `indentFirstLine` and `indentHanging` are in twips (20 per point, so `240` is 12 pt). A string in points such as `"12pt"` is accepted and converted; any other unit is refused. `lineSpacing` is a separate unit, in 240ths of a line (`240` = single, `360` = 1.5).
- **Table header bold:** `insert` with `type:'table'` makes the header row bold by default. Pass `headerBold:false` (flat or in `params`) to turn it off, or set `- Header Bold: false` under `## Tables` in a style guide. An explicit value beats the guide.
- **Unknown settings are refused:** An operation that carries a setting it does not read, for example `bold` inside `params` of a table `updateCell`, is refused with `E_VALIDATION` (exit code 1) before anything is applied. The message names the setting and the ones the operation accepts: `unknown parameter "bold" for table updateCell; allowed: …`.
- **Style operation refusal:** A `style` operation must set at least one style property, and border configurations read only `style`, `sz`, and `color`. Submitting a style operation without properties or with unrecognized attributes is refused with `E_VALIDATION`.
- **History safety checks:** History commands (`history`, `preview-revert`, `revert`, `redo`) strictly refuse any history directory whose `.git` entry is a redirect file with error code `E_IO`.
- **Atomic batch execution:** Batch operations execute sequentially with all-or-nothing rollback semantics (`rollbackOnError: true`). If any operation fails, preceding mutations roll back and the file remains untouched. Take node IDs from a fresh `inspect` or `search`; IDs of nodes created during an editing session (`_bk_new_...`) are valid only within that session.

---

## 4. Command Reference

### 4.1 Document Inspection (`astra inspect`)

Inspects the internal structure of a `.docx` file using Token-Optimized Outline Notation (TOON).

```bash
# Basic inspection with standard fidelity
astra inspect report.docx

# High-density compact view for AI prompts
astra inspect report.docx --fidelity compact

# Deep inspection of a specific section
astra inspect report.docx --fidelity full --section sec_root

# Enable visual layout telemetry and headless PNG rendering
astra inspect report.docx --visual --format png
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--fidelity` | `compact\|standard\|full` | Detail tier (default: `standard`). `full` requires `--section`. |
| `--section <id>` | `string` | Restrict outline to a specific section ID. |
| `--max-text <n>` | `number` | Truncate node text display at `n` characters (1–10,000). |
| `--width <n>` | `number` | Wrap text column to `n` columns with hanging indent. |
| `--wrap` | `flag` | Wrap automatically to terminal width. |
| `--visual` | `flag` | Compute 2D layout telemetry and generate headless page images. |
| `--format` | `png\|jpeg` | Image format for visual rendering (default: `png`). |
| `--json` | `flag` | Emit structured JSON output containing rows, styles, and telemetry. |

---

### 4.2 Schema Validation (`astra validate`)

Validates internal XML parts against the ISO/IEC 29500 / ECMA-376 schema gate using `@attentionform/ooxml-validator`.

```bash
# Human-readable validation output
astra validate contract.docx

# Structured JSON for CI/CD test gates
astra validate contract.docx --json
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--max-ratio <n>` | `number` | Zip-bomb compression ratio limit (default: `100`). |
| `--max-uncompressed-bytes <n>` | `number` | Zip-bomb uncompressed size ceiling (default: `209715200`). |
| `--style-guide <name\|path>` | `string` | Advisory style pass: bundled guide (`report\|technical\|legal`) or path to a style-guide.md file (never blocks the save). |

**JSON Output Example:**
```json
{
  "ok": true,
  "file": "contract.docx",
  "errors": []
}
```

`validate` judges the file as a save would write it: children that are out of schema order (for example `w:pBdr` after `w:spacing` in a paragraph's properties) are put in order first, then the result is checked. When that save would repair something, the command still says `VALID` but adds one `would-be normalize: ...` line per repair (and `pendingNormalizations` in the JSON), the same list `astra inspect` and `astra normalize --dry-run` print. Run `astra normalize` to write the repaired file. `w:tblHeader` accepts no value, `on` or `off`; the validator reports `true`, `1`, `false` and `0` there.

With `--style-guide`, the JSON envelope adds `styleGuide` (compact summary) and extends `styleWarnings` with one `guide-parse` entry per guide-parse note (human `STYLE`/`STYLE-NOTE` lines stay too; see §6 for the diagnostics contract).

---

### 4.3 Atomic Batch Editing (`astra batch`)

Applies atomic operations with full schema validation and rollback protection. If any operation fails, the document reverts to its pre-operation snapshot.

```bash
# 1. Quick paragraph append (no ops file required)
astra batch report.docx --append "New concluding paragraph."

# 2. Targeted paragraph replacement by needle text
astra batch report.docx --target-text "Draft v1" --append "Final Approved"

# 3. Apply an operations JSON file
astra batch report.docx ops.json --out updated_report.docx

# 4. Stream operations via stdin
cat ops.json | astra batch report.docx -

# 5. Continuous streaming mode with error policy
cat stream.jsonl | astra batch report.docx - --stream --on-error continue --max-errors 5
```

**Batch Operations File Syntax (`ops.json`):**
```json
[
  {
    "action": "replace",
    "target": "3FC95CAC",
    "text": "1. Executive Summary (Updated for Q4)"
  },
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "3FC95CAC",
    "position": "after",
    "text": "Net retention grew to 118% in Q4."
  }
]
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--append <text>` | `string` | Inserts a new final paragraph into the document. |
| `--target-text <str>` | `string` | Locates first paragraph containing `<str>` to update with `--append`. |
| `--out <docx>` | `string` | Output target file path (defaults to in-place edit). |
| `--dry-run` | `flag` | Validates operations without modifying files on disk. |
| `--backup[=<path>]` | `string` | Creates a `.bak` backup copy prior to mutation. |
| `--stream` | `flag` | Enables line-delimited JSONL streaming mode. |
| `--on-error` | `abort\|continue` | Error policy for streaming mode (default: `abort`). |
| `--max-errors <n>` | `number` | Consecutive error ceiling before stream terminates (default: `10`). |
| `--matrix <file.csv>` | `string` | Template variable substitution matrix. |
| `--reproducible` | `flag` | Pins timestamps, ids and RSIDs to `SOURCE_DATE_EPOCH` (else 1980-01-01Z) so the same input gives byte-identical output. Off by default. |
| `--no-reproducible` | `flag` | Forces real dates and ids even when `ASTRA_REPRODUCIBLE` or `SOURCE_DATE_EPOCH` is set. This is the default. |
| `--author <name>` | `string` | Default author for tracked changes and comments that do not name one (also `ASTRA_AUTHOR`). An operation's own `author` wins. |
| `--style-guide <name\|path>` | `string` | House style for this invocation (bundled `report\|technical\|legal` or guide `.md` path; file refs jail to `--allowed-dir` else the input folder). Applies table/header-text/chart/callout/code defaults where ops omit params; explicit params still win. `--json` carries its summary plus guide-parse notes as `styleWarnings` (see §6). |

**Style operations & validation:** A `style` operation must set at least one style property, and border configurations read only `style`, `sz`, and `color`. Submitting a style operation without properties or with unrecognized attributes is refused with `E_VALIDATION`.

**Atomic rollback & session scope:** Batches run sequentially with all-or-nothing rollback semantics (`rollbackOnError: true`). If any single operation in a batch fails schema validation or execution, the entire batch rolls back and leaves the file untouched. Take node IDs from a fresh `inspect` or `search`; IDs of nodes created during an editing session (`_bk_new_...`) are valid only within that session.

**Callout op shape** (also in `astra batch --schema` as `callout_insert`):
```json
{ "action": "callout", "anchor": "<nodeId>", "position": "after", "theme": "amber", "title": "Field note", "text": "Callout body" }
```
`theme` is optional (`indigo|emerald|amber|slate|rose`); omitted inherits the session/`--style-guide` default. `title` is optional.

---

### 4.4 Privacy & Sanitization (`astra sanitize`)

Scrubs sensitive metadata, edit history, comments, and PII from `.docx` documents prior to public distribution or legal submission.

```bash
# Standard sanitization: strips author names, RSIDs, comments, and core metadata
astra sanitize internal_draft.docx public_version.docx

# Client-Safe: strips metadata AND redacts body PII (emails, phone numbers)
astra sanitize internal_draft.docx client_ready.docx --client-safe

# Metadata-only: preserves substantive text and tracked revisions while scrubbing file properties
astra sanitize internal_draft.docx legal_submittal.docx --metadata-only
```

**Cross-surface note (CLI vs MCP):** CLI `sanitize` is a one-shot raw scrub operating directly on document archive parts. In contrast, MCP `sanitize_document` followed by `save_document` runs through the full document save pipeline, so the resulting file bytes can differ. Both remove the same categories of information; neither replaces a human review of the final file.

**Input validation:** `sanitize` refuses invalid or schema-corrupted input upfront with exit code 1 (`E_VALIDATION`). It will never scrub or write out an invalid document.

**Sanitization Modes:**
| Mode / Flag | Core Properties | RSID / Edit Session Tags | Comments | Tracked Changes | Body PII (Emails/Phones) |
|---|:---:|:---:|:---:|:---:|:---:|
| Bare `sanitize` | Scrubbed | Scrubbed | Scrubbed | Accepted | **Preserved** |
| `--client-safe` | Scrubbed | Scrubbed | Scrubbed | Accepted | **Redacted** |
| `--metadata-only` | Scrubbed | Scrubbed | Preserved | Preserved | **Preserved** |

**Additional Options:**
| Flag | Description |
|---|---|
| `--zero-retention` | Cleans document properties and sets empty creator tags. |
| `--replacement-author <name>` | Replaces revision author tags with a generic moniker (e.g. `"Reviewer"`). |
| `--accept-all-revisions` | Force-accepts all tracked insertions and deletions. |
| `--strip-comments` | Strips all comments and comment relationship files. |

---

### 4.5 Document Creation (`astra create`)

Creates a new `.docx` document from one of the built-in templates.

```bash
# Create an executive report
astra create quarterly_report.docx --template report

# Create a clean blank document
astra create scratch.docx --template blank

# Compile directly from Markdown during creation
astra create spec.docx --template blank --from-md spec.md
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--template` | `report\|memo\|blank` | Template layout (default: `blank`). New files carry a title, an author (`ASTRA_AUTHOR`, else `ASTra`) and created/modified times in `docProps`, plus settings, a font table and a theme, so Word opens them without a compatibility prompt. |
| `--style-guide <name\|path>` | `string` | House style: bundled guide (`report\|technical\|legal`) or path to a style-guide.md file. Skins `--from-md` tables (`headerFill`), `> [!TAG]` alerts (guide `defaultTheme`), charts, callouts, and code blocks where content omits styling. `--json` carries its summary plus guide-parse notes as `styleWarnings` (see §6). |
| `--force` | `flag` | Overwrite destination file if it already exists. |
| `--from-md <file>` | `string` | Populate newly created document from a Markdown file. |
| `--dollar-math` | `flag` | With `--from-md`, also read `$...$` as inline math (off by default because prices and shell text can look like math). `$$` blocks and `\(...\)` are always math. |

**Math & equation options:** With `--from-md`, `$...$` syntax is treated as inline math only when `--dollar-math` is specified (preventing prices or shell variables from misinterpreting as math). Display equations (`$$...$$`), standard LaTeX inline delimiters (`\(...\)`), and equation accents are always parsed and compiled into native Office Math (OMML).

---

### 4.6 Markdown Import (`astra import`)

Converts a Markdown file into a `.docx` document.

```bash
astra import release_notes.md -o release_notes.docx --template blank
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `-o <out.docx>` | `string` | Destination file path. |
| `--template` | `report\|memo\|blank` | Base template styling (default: `blank`). |
| `--force` | `flag` | Overwrite destination file if it exists. |
| `--allowed-dir <dir>` | `string` | Folder that block image/chart refs resolve against. |
| `--dollar-math` | `flag` | Also read `$...$` as inline math (off by default because prices and shell text can look like math). `$$` blocks and `\(...\)` are always math. |
| `--style-guide <name\|path>` | `string` | House style (bundled `report\|technical\|legal` or guide `.md` path; same resolve/jail semantics as `create --from-md`: file refs jail to `--allowed-dir` else cwd). Same skinning as `create --from-md`. |

**Math & equation options:** `$...$` syntax is treated as inline math only when `--dollar-math` is specified. Display equations (`$$...$$`), standard LaTeX inline delimiters (`\(...\)`), and equation accents are always parsed and compiled into native Office Math (OMML).

---

### 4.7 Multi-Format Export (`astra export`)

Headlessly renders documents to PDF, HTML, or Markdown.

```bash
# Export to PDF (requires LibreOffice in system PATH)
astra export doc.docx --format pdf -o doc.pdf

# Export to clean semantic HTML
astra export doc.docx --format html -o doc.html

# Export to GitHub Flavored Markdown
astra export doc.docx --format markdown -o doc.md
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--format` | `pdf\|html\|markdown\|md` | Output format. |
| `-o <outPath>` | `string` | Output file path (defaults to stdout for markdown/html). |

**Input validation:** PDF `export` (and other format exports) refuses invalid or corrupted input upfront with exit code 1 (`E_VALIDATION`). Validation occurs before invoking conversion pipelines.

**Known limits of export (intended behavior):** Markdown export is skin-blind: table shading, callout themes, code-block colors, and heading colors never survive (the format has no skin channel), and no record of the producing guide is kept, so re-import skin requires re-attaching the guide (`import --style-guide` / `create --from-md --style-guide`). Footnote markers hop over trailing punctuation (`[^1].` → `.[^1]`) and a blank line is injected between consecutive tables. HTML export uses fixed table/callout CSS (not the guide or the OOXML shading).

---

### 4.8 Text Search (`astra search`)

Searches across document paragraphs, table cells, and runs, returning matched stable node IDs.

```bash
# Find every paragraph containing the text
astra search contract.docx "Indemnification" --limit 5

# Searching ignores upper and lower case
astra search contract.docx "Section 4.1" --json
```

A table hit adds the cells that hold the match, such as `t_001 RegionRevenueNorth120South98 [cells: t_001:r3c1]` (`cells` in `--json`; a table's text is its cells run together). Pass the address as the `target` of a `replace`, `style`, `namedStyle`, `clearFormat`, `hyperlink`, `bookmark` or `comment` operation in `astra batch`. Cell addresses are `<tableId>:r<row>c<col>`, 1-based as `astra inspect` prints the grid; add `:p<n>` for a cell with several paragraphs. Table operations take the table id with `rowIndex`/`colIndex` numbered the same way, so `t_001:r3c2` and `rowIndex: 3, colIndex: 2` are the same cell.

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--limit <n>` | `number` | Maximum number of matches to display (default: 20). |

---

### 4.9 Structure Normalization (`astra normalize`)

Normalizes and repairs XML namespace declarations, sorts element sequences into ISO/IEC 29500 order, and eliminates redundant styling runs.

```bash
astra normalize legacy_doc.docx --out clean_doc.docx
```

---

### 4.10 Document Statistics (`astra stats`)

Calculates document telemetry including word count, paragraph count, and estimated printed page count.

```bash
astra stats report.docx
```

```text
STATS report.docx: words=1420 paragraphs=48 tables=3 pages~5
```

---

### 4.11 Visual Page Rendering (`astra visual`)

Renders one saved-file page to PNG or JPEG using headless LibreOffice (same underlying renderer as `inspect --visual`).

```bash
# Render first page of saved document to PNG
astra visual report.docx --out preview.png

# Render specific page to JPEG with custom limits
astra visual report.docx --page 2 -o preview_p2.jpg --format jpeg --force
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--page <n>` | `number` | 1-based page number to render (default: `1`). |
| `--out <path>`, `-o` | `string` | Destination image file path (required). Refuses to overwrite existing file without `--force` (exit code 2). |
| `--format` | `png\|jpeg` | Image format (default: `png`). |
| `--force` | `flag` | Overwrite destination file if it already exists. |
| `--allow-macros` | `flag` | Allow document macros during headless conversion. |
| `--strip-macros` | `flag` | Strip document macros before rendering. |
| `--max-ratio <n>` | `number` | Zip-bomb compression ratio limit (default: `100`). |
| `--max-uncompressed-bytes <n>` | `number` | Zip-bomb uncompressed size ceiling (default: `209715200`). |

`visual` renders the saved on-disk document file. For MCP sessions where unsaved edits exist in memory, call `save_document` first or use `visual_inspect`. Missing LibreOffice or `pdftoppm` dependencies degrade gracefully with exit code 1 (`E_DEPENDENCY`).

---

### 4.12 Document History (`astra history`)

Inspects recorded sidecar version control commits for a document, listing revisions from newest to oldest.

**Where history lives.** By default ASTra keeps each document's history in a hidden folder next to it, `<docDir>/.astra/<file>/`, and records which documents have had history in `<docDir>/.astra-tombstones.json` (so a deleted `.astra/` folder is noticed instead of silently starting a new chain). Both are plain local files; ASTra never uploads them. Add them to your repository's `.gitignore` if you keep documents under version control:

```gitignore
.astra/
.astra-tombstones.json
```

To keep history somewhere else, set `ASTRA_HISTORY_DIR` to a directory (for the CLI and for the MCP server's environment). Each document's history then lives in `$ASTRA_HISTORY_DIR/<key>/<file>/`, where `<key>` is a short hash of the document's folder, so two folders holding a file of the same name keep separate histories, and the tombstone registry sits in `$ASTRA_HISTORY_DIR/<key>/`. Nothing is written beside the document. History made under one setting is not visible under the other: set it before the first save, and keep it set.

```bash
# View commit history (read-only)
astra history report.docx

# Initialize history tracking baseline on an untracked document
astra history report.docx --init
```

```text
commit 8f3b2a1c9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c4b (HEAD -> main)
Author: ASTra <astra@attentionform.com>
Date:   Sun Oct 25 12:00:00 2026 +0000

    batch edit (revision 2)
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--init`, `--track` | `flag` | Initializes a sidecar history repository with the current file state as baseline if no sidecar exists yet. |

History commands strictly refuse a history repository whose `.git` entry is a redirect file with error code `E_IO`.

---

### 4.13 Preview Revert (`astra preview-revert`)

Previews restoring package state `--steps <n>` commits back read-only without modifying the file on disk.

```bash
# Preview rolling back 1 commit
astra preview-revert report.docx --steps 1
```

```text
revert dry-run: 1 commit back
target oid: 3a2b1c0d9e8f7a6b5c4d3e2f1a0b9c8d7e6f5a4b
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--steps <n>` | `number` | Number of commits to roll back (default: `1`). |

---

### 4.14 Revert to Recorded State (`astra revert`)

Restores tracked package state `--steps <n>` commits back and records the restoration as a new append-only commit.

```bash
# Revert document state by 1 commit
astra revert report.docx --steps 1
```

**Options:**
| Flag | Type | Description |
|---|---|---|
| `--steps <n>` | `number` | Number of commits to roll back (default: `1`). |

---

### 4.15 Redo Reverted State (`astra redo`)

Re-applies the package state that was replaced by the most recent revert operation (one redo per revert head).

```bash
# Redo the last reverted state
astra redo report.docx
```

---

## 5. CI/CD & Scripting Recipes

### GitHub Actions: validate every `.docx` in a repository
This job fails if any `.docx` file under `docs/` does not meet the ECMA-376 schema (`validate` checks one file per call, so the job loops over them):

```yaml
name: Document Quality Gate
on: [push, pull_request]

jobs:
  validate-docs:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 22
      - name: Validate every .docx under docs/
        run: find docs -name '*.docx' -print0 | xargs -0 -r -n1 npx -y @attentionform/astra validate
```

### Automated Document Templating Script
```bash
#!/usr/bin/env bash
# Create one report per client
for client in "Acme" "Stark" "Wayne"; do
  echo "Generating report for ${client}..."
  astra create "report_${client}.docx" --template report --force
  astra batch "report_${client}.docx" --target-text "Executive Summary" --append "Prepared specifically for ${client} Corp."
  astra export "report_${client}.docx" --format pdf -o "report_${client}.pdf"
done
```

---

## 6. Style Guides (house skin)

A style guide (`style-guide.md`) skins tables, callouts, code blocks, and charts. Bundled refs: `report` (display title `Corporate`; ref and title differ by design), `technical`, `legal`.

### Guide-file spec

```markdown
# Style Guide: <name>

## Palette
- Primary: #1E3A8A
- Secondary: #0284C7
- Accent: #D97706
- Neutral Dark: #1E293B
- Neutral Light: #F8FAFC
- Chart Palette: #1E3A8A, #0284C7, #10B981, #F59E0B   (CSV, 4+ colors; short lists pad with defaults)

## Typography
- Font Family: Calibri
- Code Font: Consolas
- Heading 1: #1E3A8A, bold
- Heading 2: #0284C7, bold

## Tables
- Header Fill: #1E3A8A
- Header Text: #FFFFFF
- Header Bold: true
- Alternating Fill: #F8FAFC
- Border Color: #CBD5E1
- Keep Together: true

## Callouts
- Default Theme: slate
- Preferred Themes: slate, emerald, amber
- Warning Theme: amber   (optional per-tag override; any of Note/Tip/Important/Warning/Caution Theme; last wins; unknown values fall back to Default Theme)

## Code Blocks
- Background: #F1F5F9
- Border Color: #CBD5E1

## Layout
- Margin: standard
- Prevent Orphan Headings: true
- Widow Control: true
```

Rules: sections are `##`-headed (`table`, `callout`, `code` spellings accepted); entries are `- Key: value` bullets. Colors are `#RRGGBB` (3-digit shorthand expands with a warning; malformed values take the built-in default with a warning). Theme enum: `indigo|emerald|amber|slate|rose`. Margin enum: `standard|narrow|wide`. `Version: 1` (unknown versions parse best-effort with a warning). The parser never throws: every assumption lands in the guide-parse notes.

Rendered today: `tables.headerFill/headerText/headerBold/keepTogether`, `tables.alternatingFill` (composed Markdown tables only: shades the odd data rows), `callouts.defaultTheme` + per-tag `callouts.tagThemes`, `codeBlocks.*`, `palette.chartPalette`. Advisory-only (parsed, linted, never rendered): palette base colors, `typography.*`, `tables.borderColor`, `callouts.preferredThemes`, `layout.*`, marked so in the bundled template comments.

**Compose without a guide.** `create --from-md`, `import` and `compose_document` use the built-in *Default Corporate* guide when you pass none: navy table header with bold white text, shaded alternate rows, a light code-block box. Alerts keep their per-tag colours (NOTE, TIP, IMPORTANT, WARNING, CAUTION). To get the older look (grey header, no header text colour, no stripes), pass `--style-guide legacy` (MCP: `styleGuide: "legacy"`, for that call only, not stored on the session).

**Title and notes in composed Markdown.** An italic-only line directly under the opening `#` title becomes the *Subtitle* style. A paragraph that starts with `[!NOTE]` (or TIP, IMPORTANT, WARNING, CAUTION) outside a blockquote stays plain text and compose reports a warning (CLI `composeWarnings`, MCP `warnings`); write each line with a `> ` prefix to get a callout.

### Precedence (explicit > guide > built-in)

| Content | Explicit | Guide | Built-in (no guide) |
|---|---|---|---|
| Batch/table op colors | op param wins | fills omitted params (dark fills gain guide header text by computed contrast) | legacy defaults |
| Compose Markdown tables | n/a (Markdown has no table styling) | `tables.headerFill` | `D9D9D9` gray |
| Batch callout `theme` | op `theme` wins | `callouts.defaultTheme` | `indigo` |
| Compose `> [!NOTE]`/`[!TIP]`/`[!IMPORTANT]`/`[!WARNING]`/`[!CAUTION]` | n/a (alert tags carry no theme syntax; every tag is generic) | `callouts.tagThemes[tag] ?? defaultTheme` | legacy per-tag map (NOTE indigo, TIP emerald, IMPORTANT slate, WARNING amber, CAUTION rose) |
| Code blocks / charts | explicit colors win | guide background/border; palette fills uncolored series in order | legacy grays / default palette |

Plain `> quote` blockquotes stay out of scope (fail closed); only `> [!TAG]` alerts compose.

### Diagnostics contract (`styleWarnings` vs stderr)

Guide-parse notes (bad hex, unknown theme/version, duplicate keys, missing sections) ride BOTH channels: human `WARNING:`/`STYLE-NOTE` stderr (or stdout `STYLE` lines for validate) AND machine JSON: `styleWarnings` on `create`/`import`/`batch` (string notes, present only when non-empty), `guide-parse` entries on `validate`, and the client-visible `warnings` array on MCP `execute_batch` (plus `styleGuideWarnings` on `new_document`/`compose_document`). Lint findings (`table-header-contrast`, `chart-duplicate-colors`, `table-orphan`) ride the same JSON keys as objects. Programmatic consumers never need to scrape stderr.

### Sandbox matrix (guide file refs)

| Command | File refs jail to | No flag fallback |
|---|---|---|
| `create --style-guide` | `--allowed-dir` | process cwd |
| `import --style-guide` | `--allowed-dir` | process cwd |
| `batch --style-guide` | `--allowed-dir` | input document folder |
| `validate --style-guide` | (no `--allowed-dir`) | process cwd |
| MCP `new/compose/execute_batch` | explicit `allowedDir` | session document folder |

Remote URLs and outside-jail refs always fail closed (`E_VALIDATION`/`E_SANDBOX`); unknown names fail naming the available set.

### Session semantics (MCP)

`new_document` attaches (persists for the session); `compose_document` with a guide **replaces** the session guide (not a one-shot); `execute_batch` with a guide overrides **that batch only**. Guides cannot be detached; open a new session; an empty `styleGuide` fails with `E_VALIDATION` saying so. Guided re-import after Markdown export: `import --style-guide` (export strips all skin; see §4.7).
