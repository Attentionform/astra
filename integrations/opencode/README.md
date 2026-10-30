# 30-Second Turnkey Guide: Using ASTra with OpenCode

Integrate **ASTra** — the headless, high-fidelity OOXML engine — directly into **OpenCode** (the premier open-source AI coding agent).

With ASTra's `@attentionform/astra-mcp-server`, OpenCode gains native capabilities to inspect, author, redline, chart, format, validate, and export `.docx` documents through a 14-tool Model Context Protocol (MCP) interface.

---

## ⚡ Quickstart in 30 Seconds

### Step 1: Copy Configuration to Your Project

From the ASTra repository root. Prerequisite: OpenCode installed
(`opencode --version` prints a version; if not, install it first, then
return here):

```bash
# Option A: Project-level configuration (recommended)
cp integrations/opencode/opencode.json ./opencode.json

# Option B: OpenCode hidden folder configuration
mkdir -p .opencode
cp integrations/opencode/opencode.json .opencode/opencode.json
```

Merge, do not overwrite: if `./opencode.json` (or
`.opencode/opencode.json`) already exists with your own `mcp` entries,
copy only the `astra` / `astra-dev` blocks under `mcp` into it — a bare
`cp` clobbers your existing servers.

Working-directory trap: the `command` arrays in `opencode.json` are
repo-relative (`./packages/mcp-server/...`), so always launch `opencode`
from the ASTra repo root. Started elsewhere, the server spawn fails
(resolve the trap with absolute paths in `command`). `--allowed-dir`
itself must always be absolute (see step 3 of the local setup below).

### Step 2: Install the Document Engineering Agent

```bash
mkdir -p .opencode/agents
cp integrations/opencode/agent/doc-editor.md .opencode/agents/doc-editor.md
```

### Step 3: Launch OpenCode

Start OpenCode in your terminal:

```bash
opencode
```

Switch to the document editor agent:
```
@doc-editor Please review documents/report.docx and format the executive summary with our corporate style palette.
```

---

## ⚙️ Configuration Reference (`opencode.json`)

The turnkey configuration (`integrations/opencode/opencode.json`) defines local stdio MCP servers within the ASTra monorepo workspace.

Because `@attentionform/astra-mcp-server` is a workspace package (publishable: no `private` field in `packages/mcp-server/package.json`, `publishConfig.access` is `public`), OpenCode runs the server directly from the local workspace:

1. **`astra` (local build, default enabled)**: Pre-compiled production runtime via `node ./packages/mcp-server/dist/index.js --allowed-dir /absolute/path/to/your/documents`. Fast startup and minimal overhead.
2. **`astra-dev` (direct tsx)**: Direct TypeScript execution via `npx tsx ./packages/mcp-server/src/index.ts --allowed-dir /absolute/path/to/your/documents`. Ideal for developing or testing engine modifications without rebuilding.

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "astra": {
      "type": "local",
      "command": [
        "node",
        "./packages/mcp-server/dist/index.js",
        "--allowed-dir",
        "/absolute/path/to/your/documents"
      ],
      "environment": {
        "NODE_ENV": "production"
      },
      "enabled": true
    },
    "astra-dev": {
      "type": "local",
      "command": [
        "npx",
        "tsx",
        "./packages/mcp-server/src/index.ts",
        "--allowed-dir",
        "/absolute/path/to/your/documents"
      ],
      "environment": {
        "NODE_ENV": "development"
      },
      "enabled": false
    }
  }
}
```

### 📁 Local Workspace Setup
To run ASTra with OpenCode in your local repository:
1. Ensure you are in the ASTra workspace root directory.
2. Build the MCP server bundle: `npm run build:packages` (outputs to `./packages/mcp-server/dist/index.js`).
3. Replace `/absolute/path/to/your/documents` with your own absolute documents path (hosts start the server with an unrelated working directory, so a relative path resolves against the wrong root and every document call fails). Create that directory if needed.
4. Launch OpenCode (`opencode`).

### Security & Sandbox Options
- `--allowed-dir <absolute-path>`: Restricts file reads, writes, and exports strictly within the specified directory. Always use an absolute path. Repeat the flag for extra roots.
- `ASTRA_EXTRA_DOC_ROOTS`: Colon-separated (`:` or `;` on Windows) list of additional authorized document roots. The legacy env var is ignored by the server — setting it widens nothing, so do not rely on it (see `packages/mcp-server/README.md` Environment Variables).

---

## 💬 Interactive Prompt Examples

### 1. Corporate Report Generation with Native 3D Charts
```text
@doc-editor Open documents/Q3_Review.docx and add a section titled "Regional Sales Performance". 
Insert a 3D bar chart comparing North America (450k), EMEA (380k), and APAC (520k) using 
our brand color palette. Validate and save.
```

### 2. Legal Contract Redlining & Tracked Changes
```text
@doc-editor Open documents/Vendor_Agreement.docx. Search for the Indemnification clause 
and redline it: mark the current liability cap of $100,000 as a tracked deletion (w:del) 
and add a new cap of $500,000 as a tracked insertion (w:ins) with author "Legal Counsel".
```

### 3. Multi-Format Omnichannel Publishing
```text
@doc-editor Open documents/whitepaper.docx, inspect the outline, and export the document 
into both pristine Markdown (documents/whitepaper.md) and HTML (documents/whitepaper.html). 
Then export a publication-ready PDF to documents/whitepaper.pdf.
```

---

## 🔧 Dependencies & Troubleshooting

### Headless PDF Export (`E_DEPENDENCY`)
ASTra's `export_document` with `format: 'pdf'` relies on headless LibreOffice (`soffice` / `libreoffice`) and `poppler-utils` (`pdftoppm`).

If missing, ASTra fails-closed with `E_DEPENDENCY` and actionable remedy instructions:

#### Debian / Ubuntu / WSL:
```bash
sudo apt-get update && sudo apt-get install -y libreoffice poppler-utils
```

#### macOS (Homebrew):
```bash
brew install --cask libreoffice
brew install poppler
```

#### Fedora / RHEL:
```bash
sudo dnf install -y libreoffice poppler-utils
```

### Error Code Quick Reference
| Error Code | Meaning | Remedy |
| --- | --- | --- |
| `E_SANDBOX` | Target file is outside `--allowed-dir` | Place documents inside allowed roots or set `ASTRA_EXTRA_DOC_ROOTS`. |
| `E_VALIDATION` | Invalid operation shape or corrupt OOXML | Follow the remedy XPath hint, correct half-points / dxa units, and retry. |
| `E_SESSION` | Session ID not found or expired (>30m idle) | Re-run `open_document` to get a fresh session ID. |
| `E_QUOTA` | Operation exceeded caps (>500 ops / >200 MB) | Split large modifications into smaller sequential batches. |
| `E_DEPENDENCY` | Missing host system utility (LibreOffice) | Install `libreoffice` via your OS package manager. |
