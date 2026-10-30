# @attentionform/astra

The official command-line interface for the ASTra document engine (`astra`).

`@attentionform/astra` provides high-performance terminal inspection, ECMA-376 schema validation, atomic batch editing, privacy sanitization, and multi-format export for Microsoft Word (`.docx`) files. The interactive chat IDE moved to the separate `astra-agent` binary.

---

## Quickstart

### Run with `npx` (Zero Install)

Run commands immediately without installing locally:

```bash
npx -y @attentionform/astra --help
```

### Install Globally

```bash
npm install -g @attentionform/astra
astra --help
```

### 3-Minute Tour

```bash
# 1. Create a document from an executive template
astra create proposal.docx --template report

# 2. Inspect document structure in token-compact notation
astra inspect proposal.docx --fidelity compact

# 3. Search for headings or paragraphs to find stable node IDs
astra search proposal.docx "Executive"

# 4. Append a new paragraph atomically
astra batch proposal.docx --append "Project timeline spans 6 months."

# 5. Export to PDF and Markdown
astra export proposal.docx --format pdf -o proposal.pdf
astra export proposal.docx --format markdown -o proposal.md

# 6. Sanitize metadata before external sharing
astra sanitize proposal.docx proposal_clean.docx
```

---

## Command Overview

| Command | Description | Example |
|---|---|---|
| `inspect` | Displays TOON outline, table grids, and node IDs | `astra inspect doc.docx --fidelity compact` |
| `validate` | Checks ECMA-376 schema conformance and package integrity | `astra validate doc.docx` |
| `batch` | Applies atomic operations or JSON streams with rollback | `astra batch doc.docx ops.json` |
| `search` | Searches text across nodes with literal or regex matching | `astra search doc.docx "Revenue" --limit 5` |
| `create` | Generates a fresh `.docx` from `report`, `memo`, or `blank` | `astra create new.docx --template memo` |
| `import` | Compiles Markdown files directly into `.docx` | `astra import spec.md -o spec.docx` |
| `export` | Headless export to PDF, HTML, or Markdown | `astra export doc.docx --format pdf` |
| `sanitize` | Scrubs metadata, comments, revision history, and PII | `astra sanitize doc.docx clean.docx` |
| `normalize` | Canonicalizes XML namespaces and fixes element ordering | `astra normalize doc.docx` |
| `stats` | Reports word count, paragraph count, and estimated pages | `astra stats doc.docx` |
| `visual` | Renders one saved-file page to PNG/JPEG via headless LibreOffice | `astra visual doc.docx --out render.png` |
| `history` | Lists recorded sidecar commits, newest first (read-only) | `astra history doc.docx` |
| `revert` | Restores the package state recorded N commits back, as a new commit | `astra revert doc.docx --steps 1` |
| `redo` | Re-applies the state replaced by the most recent revert | `astra redo doc.docx` |

---

## Command Reference

Each command lists its job in one line; run `astra <command> --help` for the full flag table and `--json` for scripting.

### 1. Inspect Document (`astra inspect`)
Structure, node IDs, and document stats; `--visual` adds spatial telemetry plus one headless page render.

### 2. Validate Document (`astra validate`)
ECMA-376 schema conformance and package integrity; never writes.

### 3. Sanitize Document (`astra sanitize`)
Scrubs metadata, comments, revision history, and RSIDs; `--redact-pii` also redacts body PII.

### 4. Atomic Batch Operations (`astra batch`)
Applies an ops file with rollback; supports `--stream`, `--matrix`, and shortcut flags.

### 5. Repair Legacy Files (`astra normalize`)
Canonicalizes XML namespaces and fixes element ordering.

### 6. Multi-Format Export (`astra export`)
Exports to PDF, HTML, or Markdown.

### 7. Full-Text Search (`astra search`)
Finds node IDs by literal or regex text match.

### 8. Create Document (`astra create`)
Generates a fresh `.docx` from a bundled template.

### 9. Import Markdown (`astra import`)
Compiles a Markdown file into a `.docx`.

### 10. Document Statistics (`astra stats`)
Word, paragraph, and page counts (deterministic estimate).

### 11. Visual Page Rendering (`astra visual`)
Renders one saved-file page to PNG/JPEG via headless LibreOffice (same renderer as `inspect --visual`).

### 12. Document History (`astra history`)
Lists recorded sidecar commits newest first; read-only; `--init` (alias `--track`) records the current file as the baseline when no sidecar exists yet.

### 13. Revert to Recorded State (`astra revert`)
Restores the tracked package state `--steps <n>` commits back and records the restore as a new commit (append-only); `preview-revert` previews the same change read-only.

### 14. Redo Reverted State (`astra redo`)
Re-applies the state replaced by the most recent revert (one redo per revert head).

---

## Global Flags

These flags work across all subcommands:

| Flag | Purpose |
|---|---|
| `--json` | Emits single-line structured JSON on `stdout` for CI/CD scripting and `jq` pipelines. |
| `--verbose` | Emits detailed operational diagnostics on `stderr`. |
| `--version` | Displays current CLI version and exits with `0`. |
| `--help`, `-h` | Displays help information for any command (`astra <command> --help`). |

---

## Practical CLI Recipes

### 1. Atomic Node Replacement (Without an Ops File)
Quickly replace targeted text without needing to write JSON operations:

```bash
# Finds the paragraph containing "Draft Status" and updates it
astra batch doc.docx --target-text "Draft Status" --append "Final Approved"
```

### 2. High-Throughput Batch Automation
Pipe JSON operations directly via `stdin`:

```bash
cat operations.json | astra batch doc.docx - --out modified.docx
```

### 3. CI/CD Schema Verification
Integrate ASTra into automated build pipelines:

```bash
# Returns exit code 0 if valid, 1 if schema violations are found
astra validate generated_report.docx --json | jq .ok
```

### 4. Privacy Sanitization for Public Distribution
Scrub author names, editing history, and metadata:

```bash
# Default: scrubs metadata, revision tags, and hidden RSIDs while keeping body text
astra sanitize draft.docx public.docx

# Client-Safe: scrubs metadata, redacts body PII, and accepts all tracked changes
astra sanitize draft.docx client_final.docx --client-safe
```

### 5. Multi-Page Visual Telemetry
Inspect layout and generate high-resolution PNG previews:

```bash
astra inspect doc.docx --visual --format png
```

---

## Dates, authors and reproducible output

By default ASTra writes the **current time** into tracked changes, comments and document properties, and attributes them to the author you give: an `author` on the operation, else `--author <name>` (or `ASTRA_AUTHOR`) for the run, else `ASTra`. Any revision or comment operation also takes an explicit ISO 8601 `date`.

For **byte-identical output** (same input twice gives the same bytes), pass `--reproducible`, or set `ASTRA_REPRODUCIBLE=1` or `SOURCE_DATE_EPOCH`. Timestamps, ids and document properties are then pinned to `SOURCE_DATE_EPOCH`, else `1980-01-01Z`. `--no-reproducible` (or `ASTRA_NO_REPRODUCIBLE=1`) forces real values even when one of those is set.

---

## Exit Code Contract

ASTra enforces a strict, predictable exit code contract for headless automation:

| Exit Code | Meaning | Common Causes |
|:---:|---|---|
| `0` | **Success** | Operation completed successfully, document validated. |
| `1` | **Operational Failure** | File not found, corrupted XML archive, schema validation failed, or batch rolled back. |
| `2` | **Usage Error** | Unknown flag, missing required positional arguments, or file overwrite conflict without `--force`. |

---

## Local Development

```bash
# Build CLI bundle
npm run build:packages

# Run local CLI binary
node packages/cli/dist/cli.js --help
```

---

## License

Business Source License 1.1 (`BUSL-1.1`): source-available, free for individuals, education, non-profits, public bodies and small businesses; other production use needs a commercial license. Each version converts to Apache-2.0 four years after release. See [LICENSE](LICENSE) and write to astra@attentionform.com for a commercial license.
