# @attentionform/astra-mcp-server

Official Model Context Protocol (MCP) server for the ASTra document engine.

`@attentionform/astra-mcp-server` exposes a comprehensive, 19-tool agentic interface over standard I/O (`stdio`). It enables AI coding assistants, including **Claude Desktop**, **Cursor**, **Windsurf**, **Cline**, and custom LLM agents, to inspect, query, mutate, and export Microsoft Word (`.docx`) documents with transactional integrity and zero schema corruption.

---

## Quickstart

### Run with `npx` (Recommended)

No installation required. Pass your document directory via `--allowed-dir`:

```bash
npx -y @attentionform/astra-mcp-server --allowed-dir /path/to/your/documents
```

> **Security Note:** ASTra enforces strict directory sandboxing. The server will only read and write files within directories explicitly declared via `--allowed-dir` (or the `ASTRA_EXTRA_DOC_ROOTS` environment variable). The path must be absolute and exist on disk.

### Global Installation

```bash
npm install -g @attentionform/astra-mcp-server
astra-mcp-server --allowed-dir /path/to/your/documents
```

---

## Client Integration

Add ASTra to your AI agent with a simple JSON snippet:

### 1. Claude Desktop
Add ASTra to your Claude Desktop configuration file:
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux:** `~/.config/Claude/claude_desktop_config.json`

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

### 2. Cursor IDE
Create or edit `.cursor/mcp.json` in your project root:

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

### 3. Windsurf
Add to your Windsurf MCP configuration (`~/.codeium/windsurf/mcp_config.json`):

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
      ]
    }
  }
}
```

### 4. Cline (VS Code Extension)
In Cline's MCP settings tab, configure:
- **Server Name:** `astra`
- **Command:** `npx`
- **Args:** `["-y", "@attentionform/astra-mcp-server", "--allowed-dir", "${workspaceFolder}"]`

### Schema freshness (when tools change)

Every served tool ships a version-stamped JSON copy under `schemas/mcp/` (`x-generated-from` names the server build). Clients cache those shapes:

- **Memory caches** (Claude Desktop for the app lifespan, Cursor and Zed per session): restart the host app to pick up new tools.
- **Disk caches** (file-based setups that vendor the JSON copies): re-run the dump, then point the client at the fresh dir:

```bash
npx -y @attentionform/astra-mcp-server --dump-schemas ./fresh-schemas
npx -y @attentionform/astra-mcp-server --check-schemas ./fresh-schemas
```

`--check-schemas` exits `0` when every stamp matches the installed server and prints a stale-file warning with exit `1` otherwise. CI regenerates and diffs `schemas/mcp/` on every push, so a schema change without refreshed copies fails the build.

---

## The 19 MCP Tools

ASTra provides a complete set of 19 tools organized by capability. The server returns all 19 tools over `tools/list`:

### `open_document`
Opens a `.docx` file into an isolated in-memory editing session.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `path` | `string` | Yes | Absolute path to `.docx` file inside an allowed directory. |
| `sessionId` | `string` | No | Server-minted session id (deprecated to supply; test-only). |
| `summaryOnly` | `boolean` | No | Return structural and style palette summary only. |
| `allowMacros` | `boolean` | No | Preserve macro parts byte-for-byte. |
| `stripMacros` | `boolean` | No | Strip macro parts upon loading. |
| `author` | `string` | No | Default author for tracked changes and comments in this session that do not name one (falls back to the server's `ASTRA_AUTHOR`, then `ASTra`). An operation's own `author` wins. |

**Returns:** `{ ok: true, sessionId, revision, summary, evicted }`

---

### `inspect_document`
Inspects document structure using Token-Optimized Outline Notation (TOON).

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `fidelity` | `string` | No | `"compact"` (default), `"standard"`, or `"full"`. |
| `sectionId` | `string` | No | Restrict outline to a specific section. |
| `maxTextLength` | `number` | No | Truncate run text at character limit. |
| `offset` | `number` | No | Pagination start row (default: 0). |
| `limit` | `number` | No | Maximum rows to return (default: 50, max: 200). |
| `aroundNode` | `string` | No | Node ID to center the inspection window around. |
| `window` | `number` | No | Half-width of rows around `aroundNode` (default: 5). |
| `maxTokens` | `number` | No | Token budget cutoff on clean row boundary (200–8000). |

**Returns:** `{ ok: true, sessionId, revision, fidelity, palette, toonText, rowCount }`

---

### `search_content`
Fast text search returning matching stable node IDs and text snippets.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `query` | `string` | Yes | Search query string (1–500 characters). |
| `regex` | `boolean` | No | Treat query as regular expression (default: `false`). |

**Returns:** `{ ok: true, sessionId, matches: [{ id, preview, cells? }] }`. A table hit lists `cells`, up to five cell addresses such as `t_001:r3c1`, that you can pass as the `target` of `replace`, `style`, `namedStyle`, `clearFormat`, `hyperlink`, `bookmark` or `comment`. Cell addresses are `<tableId>:r<row>c<col>`, 1-based as `inspect_document` prints the grid (add `:p<n>` for a cell with several paragraphs).

---

### `execute_batch`
Executes an atomic array of typed operations with schema validation and rollback guarantees.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `operations` | `array` | Yes | Array of 1 to 500 typed operation objects. |
| `sequential` | `boolean` | No | Apply operations sequentially in order (default: `true`). |
| `rollbackOnError` | `boolean` | No | Roll back entire batch if any operation fails (default: `true`). |
| `maxSteps` | `number` | No | Maximum execution steps before aborting. |
| `allowedDir` | `string` | No | Jail directory for image `filePath` reads. |
| `styleGuide` | `string` | No | Per-batch style guide override (name or `.md` path). |

**Returns:** `{ ok: true, sessionId, appliedCount, revision, createdNodeIds }`

---

### `save_document`
Validates all modified XML parts against ECMA-376 schema and commits changes to disk.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `outPath` | `string` | No | Target path (defaults to original session path). |
| `message` | `string` | No | Commit log message for document history. |
| `author` | `string` | No | Author name for document history. |
| `allowMacros` | `boolean` | No | Explicit macro policy to keep macros on save. |
| `stripMacros` | `boolean` | No | Explicit macro policy to strip macros on save. |

**Returns:** `{ ok: true, sessionId, path, revision, commitOid }`

---

### `close_document`
Closes the active session and releases all in-memory resources.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Returns:** `{ ok: true, sessionId }`

---

### `new_document`
Creates a fresh `.docx` file from an executive template and opens it into a session.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `outPath` | `string` | Yes | Destination file path inside an allowed directory. |
| `template` | `string` | No | Template type: `"report"`, `"memo"`, or `"blank"` (default: `"blank"`). |
| `force` | `boolean` | No | Overwrite existing file if it exists (default: `false`). |
| `styleGuide` | `string` | No | Bundled guide name or guide `.md` path to attach. |
| `author` | `string` | No | Default author for tracked changes and comments in this session that do not name one (falls back to the server's `ASTRA_AUTHOR`, then `ASTra`). An operation's own `author` wins. |
| `sessionId` | `string` | No | Server-minted session id (deprecated to supply; test-only). |

**Returns:** `{ ok: true, path, sessionId, revision, template, bytes }`

---

### `compose_document`
Compiles a Markdown string into OOXML structures and applies them cleanly to the document.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `markdown` | `string` | Yes | Markdown source text to compile and apply. |
| `allowedDir` | `string` | No | Jail directory for resolving relative images and charts. |
| `template` | `string` | No | Replace seed content for `"report"`, `"memo"`, or `"blank"`. |
| `styleGuide` | `string` | No | Bundled guide name or guide `.md` path to skin composition. |
| `dollarMath` | `boolean` | No | Treat `$...$` syntax as inline math (default: `false`). |

**Returns:** `{ ok: true, sessionId, appliedCount, revision, createdNodeIds, persistentIds? }`. `createdNodeIds` (`md_0001…`) work for the whole session; `persistentIds` maps each created paragraph to the id the saved file carries, for addressing it after the file is reopened (tables are renumbered by position, so they are not listed).

---

### `import_document`
Ingests an on-disk Markdown file into an open session under directory sandboxing.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `path` | `string` | Yes | Path to on-disk Markdown file inside allowed directory. |
| `allowedDir` | `string` | No | Jail directory for resolving relative images and charts. |
| `template` | `string` | No | Verify template seed disposition (`"report"`, `"memo"`, `"blank"`). |
| `styleGuide` | `string` | No | Style guide attach (bundled name or guide `.md` path). |
| `dollarMath` | `boolean` | No | Treat `$...$` syntax as inline math (default: `false`). |

**Returns:** `{ ok: true, sessionId, appliedCount, revision }`

---

### `export_document`
Renders the current in-memory state to PDF, HTML, or Markdown.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `format` | `string` | Yes | Export format: `"pdf"`, `"html"`, or `"markdown"`. |
| `outPath` | `string` | No | Target destination file path inside allowed directory. |
| `tagged_pdf` | `boolean` | No | PDF only: tagged (accessible) render (default: `true`). |
| `tagged` | `boolean` | No | Alias of `tagged_pdf`. |
| `strict` | `boolean` | No | PDF only: fail closed on preflight blockers (default: `false`). |

**Returns:** `{ ok: true, format, path, bytes, content }`

---

### `visual_inspect`
Generates 2D spatial layout telemetry and optional headless page renders.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `pages` | `array` | No | 1-indexed page numbers to inspect/render (max 4). |
| `withVisual` | `boolean` | No | Generate headless rendered images (default: `false`). |
| `format` | `string` | No | Image format: `"png"` (default) or `"jpeg"`. |
| `jpegQuality` | `number` | No | JPEG quality factor (1–100). |
| `dpi` | `number` | No | Render resolution (72–300, default: `150`). |
| `outputDir` | `string` | No | With `withVisual`, write the images to this folder (created if missing) as `<document>-page-<n>.png`, for clients that cannot read the server's temp folder. Must be inside the allowed folders, like a document path. Default: the server's temp artifacts folder. |

**Returns:** `{ ok: true, sessionId, telemetry, images: [{ page, path }], outDir }`

The wireframe in `telemetry` shows `[TABLE: n cols]` with the header row and the first rows, `- ` and `1.` list markers, and `[CALLOUT]` for a shaded one-cell table.

---

### `get_document_history`
Retrieves past committed versions of the document.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `limit` | `number` | No | Max commit entries to retrieve (default: `20`, max: `100`). |
| `offset` | `number` | No | Pagination offset for commit log. |

**Returns:** `{ ok: true, history: [{ oid, message, author, timestamp }] }`

---

### `preview_revert`
Simulates a rollback without applying changes.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `steps` | `number` | No | Number of commits to roll back. |
| `commitOid` | `string` | No | Specific commit hash target to preview reversion. |

**Returns:** `{ ok: true, targetOid, affectedFiles, dirty }`

---

### `revert_document`
Rolls back the document state to a previous commit, creating an append-only restore commit.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `steps` | `number` | No | Number of commits to roll back. |
| `commitOid` | `string` | No | Specific commit hash target to restore. |
| `discardUnsaved` | `boolean` | No | Drop uncommitted changes (default: `false`). |
| `message` | `string` | No | Revert commit message (≤ 500 chars). |
| `author` | `string` | No | Author identifier for revert commit (≤ 100 chars). |
| `reason` | `string` | No | Audit reason string for revert log (≤ 500 chars). |

**Returns:** `{ ok: true, sessionId, path, revertedToCommit, revision }`

---

### `redo_document`
Re-applies the version undone by the last revert.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `message` | `string` | No | Redo commit message (≤ 500 chars). |
| `author` | `string` | No | Author identifier for redo commit (≤ 100 chars). |
| `reason` | `string` | No | Audit reason string for redo log (≤ 500 chars). |

**Returns:** `{ ok: true, sessionId, revision, commitOid }`

---

### `validate_document`
Preflight read-only OOXML schema audit for a session before committing changes.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Returns:** `{ ok: true, sessionId, revision, dirty, validation: { valid, errorCount, errors } }`

---

### `get_stats`
Token-cheap document census returning word, paragraph, table, image, and page counts.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Returns:** `{ ok: true, sessionId, revision, dirty, words, paragraphs, tables, images, pages }`

---

### `sanitize_document`
Scrub metadata, revision IDs, authors, comments, and PII from a document session.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |
| `zeroRetention` | `boolean` | No | Purge all tracking metadata, revision IDs, authors, and EXIF. |
| `purgeRSIDs` | `boolean` | No | Strip revision save IDs (`w:rsidR`, `w:rsidRPr`, etc.). |
| `purgeAuthors` | `boolean` | No | Strip author identity metadata. |
| `purgeExif` | `boolean` | No | Strip metadata and EXIF profiles from embedded images. |
| `purgePII` | `boolean` | No | Redact potential PII patterns from text and comments. |
| `acceptAllRevisions` | `boolean` | No | Accept all tracked changes before scrubbing. |
| `stripComments` | `boolean` | No | Remove all comments and comment references. |
| `replacementAuthor` | `string` | No | Author replacement string (≤ 256 chars). |
| `allowMacros` | `boolean` | No | Preserve macro parts byte-for-byte during scrub. |
| `stripMacros` | `boolean` | No | Strip macro parts during scrub. |

**Returns:** `{ ok: true, sessionId, revision, dirty, removedRSIDs, purgedAuthors, strippedExifBytes, purgedPII, bodyPII, acceptedRevisions, strippedComments, strippedMacros, macroParts, filesTouched, droppedRogue }`

---

### `normalize_document`
Canonical repair for an open session with schema sequence sorts and tag normalizations.

**Parameters:**
| Parameter | Type | Required | Description |
|---|---|:---:|---|
| `sessionId` | `string` | Yes | Active session identifier. |

**Returns:** `{ ok: true, sessionId, revision, dirty, normalized, validation: { valid: true } }`

---

## Recommended Agent Lifecycle Flow

AI agents should follow this standard execution sequence:

```text
open_document (or new_document)
      ↓
inspect_document (get TOON outline & node IDs)
      ↓
search_content (optional: pinpoint exact target text)
      ↓
execute_batch (apply atomic mutations with rollback protection)
      ↓
visual_inspect / export_document (verify visual presentation)
      ↓
save_document (validate ECMA-376 schema & commit to disk)
      ↓
close_document (release session memory)
```

---

## Error Handling & Remediation

When a tool operation cannot be completed, ASTra responds with structured error details:

```json
{
  "ok": false,
  "code": "E_VALIDATION",
  "message": "Invalid child element sequence in w:pPr",
  "remedy": "Ensure paragraph property tags follow ISO/IEC 29500 ordering rules.",
  "isError": true
}
```

### Common Error Codes
- `E_SANDBOX`: Requested file path lies outside declared `--allowed-dir` directories.
- `E_VALIDATION`: Operation parameters failed Zod schema checks or OOXML validation.
- `E_SESSION`: Session ID does not exist or has timed out.
- `E_CONFLICT`: Operation would overwrite an existing file without `force: true`.
- `E_DEPENDENCY`: External dependency (e.g. LibreOffice for PDF export) is not found in `PATH`.

---

## Local Development

To run or debug `@attentionform/astra-mcp-server` locally from this repository:

```bash
# Build package bundles
npm run build:packages

# Run local MCP server directly with node
node packages/mcp-server/dist/index.js --allowed-dir /path/to/docs
```

---

## License

Business Source License 1.1 (`BUSL-1.1`): source-available, free for individuals, education, non-profits, public bodies and small businesses; other production use needs a commercial license. Each version converts to Apache-2.0 four years after release. See [LICENSE](LICENSE) and write to astra@attentionform.com for a commercial license.
