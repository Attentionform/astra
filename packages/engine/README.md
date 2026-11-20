# @attentionform/astra-engine

The core headless OOXML engine for ASTra.

`@attentionform/astra-engine` provides transactional, in-memory manipulation of Microsoft Word (`.docx`) documents with strict ISO/IEC 29500 / ECMA-376 schema compliance, persistent DOM node addressing, atomic batch execution with rollback, and Token-Optimized Outline Notation (TOON).

---

## Installation & Quickstart

```bash
npm install @attentionform/astra-engine
```

### Basic Example: Inspecting and Mutating a Document

```typescript
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';
import * as fs from 'node:fs';

async function main() {
  // 1. Load an existing .docx file into an in-memory DOM
  const fileBuffer = fs.readFileSync('document.docx');
  const parser = new DocxParser(fileBuffer);
  
  // 2. Access the typed Document AST and structural summary
  const ast = await parser.getAST();
  console.log(`Document loaded with ${ast.sections.length} section(s).`);

  // 3. Execute an atomic batch mutation with automatic rollback on error
  const engine = new OperationEngine(ast, parser);
  const result = await engine.executeBatch({
    operations: [
      {
        action: 'replace',
        target: ast.sections[0].nodes[0].id,
        text: 'Executive Strategic Review (Q4)'
      },
      {
        action: 'insert',
        type: 'paragraph',
        anchor: ast.sections[0].nodes[0].id,
        position: 'after',
        text: 'Net revenue increased by 28% quarter-over-quarter.'
      }
    ]
  });

  if (result.success) {
    console.log(`Applied ${result.appliedCount} operations successfully.`);
    
    // 4. Serialize back to a .docx buffer, then validate it before you write it.
    //    rezip() runs integrity checks (duplicate paragraph ids, XXE, macros, package integrity).
    //    It does NOT run the Microsoft schema validator: validateDocxBuffer() does.
    const outBuffer = parser.rezip();
    const validation = await validateDocxBuffer(outBuffer);
    if (!validation.ok) {
      throw new Error(`Not written: ${validation.errors[0]?.description}`);
    }
    fs.writeFileSync('document_updated.docx', outBuffer);
  } else {
    console.error('Batch failed:', result.errors);
  }
}

main();
```

### Starting from a template

```typescript
import { DocxParser, TEMPLATE_NAMES } from '@attentionform/astra-engine';
import * as fs from 'node:fs';

console.log(TEMPLATE_NAMES); // blank, report, memo, letter, proposal, invoice, technical-spec, research-paper, novel, newsletter

const parser = DocxParser.createNewDocument('memo', { pageSize: 'a4' }); // Letter is the default
fs.writeFileSync('memo.docx', parser.rezip());
```

`blank` is one empty paragraph. The other nine are designed starting documents with placeholder text in square brackets. Runnable examples are in `packages/engine/examples/`.

---

## What is public

The package has four supported entry points. Everything else is internal.

| Import | What it holds | Promise |
|--------|---------------|---------|
| `@attentionform/astra-engine` | 44 functions, classes and constants, and 163 types: create and open documents, edit with operations, compose Markdown, export and render, validate, sanitize, history and original recovery, style guides, read a document, errors | Documented, semantic versioning |
| `@attentionform/astra-engine/history` | The recorded history of a document: start it, record a state, list states, read a state back (14 functions and constants, 3 types) | Documented, semantic versioning |
| `@attentionform/astra-engine/ooxml` | `normalizePackage` and `previewNormalizations`: fix the structure of a package that another producer wrote | Documented, semantic versioning |
| `@attentionform/astra-engine/search` | `grepNodes`: search a list of nodes that you prepared. It is not a search of a whole document | Documented, semantic versioning |
| `@attentionform/astra-engine/internal` and `/internal/*` | Everything the command-line tool and the MCP server need, and more | **No promise.** For ASTra's own packages. Names and shapes can change in any release |

The exact names are in [`api-surface.json`](./api-surface.json). CI installs the packed package in a clean project and checks that every
entry point exports exactly those runtime values, that every type imports under `NodeNext` and `Bundler` resolution, and that a deep import such
as `@attentionform/astra-engine/git` is refused.

Some types are public only because a public class or function mentions them, for example the chart style types that `ChartOperation` carries. They
are part of the public declarations: changing their shape can break TypeScript code, so they follow the same versioning as everything else.

Test hooks (`testHooks` on the render options) are not part of the public types. They remain in the code for ASTra's own tests.

`@attentionform/astra-engine/internal` is reachable because the package map has to give the command-line tool and the MCP server a way in. That is
not a promise that it is safe to use: if a name you need is only in `/internal`, open an issue so that it can be documented and supported.

### Which renderer is installed

`getVisualBackendStatus()` reports what is installed: `renderers.word` (found, executable, version), `renderers.libreoffice` (found, path),
`rasterizers.poppler` and `rasterizers.windows`, the order a default call would try (`plan`), and what is missing. Its `available` field means that a
renderer **and** a page-image program were found. A machine with LibreOffice but without Poppler reports `available: false`, although a PDF can be made.
Use `hasPdfBackend()` for that question. [`examples/10-render.mjs`](./examples/10-render.mjs) shows both.

### Handling errors

Handle a failure by its code: `getEngineErrorCode(err)` returns one of `ENGINE_ERROR_CODES`, and `getEngineRemedy(err)` returns the text that tells the
caller what to do next. The classes (`EngineQuotaError`, `EngineValidationError`, `OoxmlValidationError`, `PreflightError`, `RenderError`,
`WordRenderError`, `WindowsPdfError`, `DependencyError`) are exported for `instanceof` checks, but the code is the contract.
[`examples/08-validate.mjs`](./examples/08-validate.mjs) shows it.

### Reading an older version

Open a document with `DocumentManager` to get undo, redo, the version list and original recovery. To preview what a revert would change, or to read an
older state as a document, use `getHistory`, `checkoutCommit` and `trackedPredicateFor` from `/history`.
[`examples/09-history-older-version.mjs`](./examples/09-history-older-version.mjs) previews additions, removals and changes of every recorded part,
including an image that a later state added, and reads the older document back and validates it.

### Reading inside a .docx

The engine does not export its zip helper. To look inside a `.docx` yourself, use any zip library, for example `adm-zip`.

---

## Core Engine Capabilities

### 1. In-Memory DOM Level 2 with Stable Addressing
- Parses `.docx` archives directly in memory using an internal ZIP adapter with zero extraction to disk.
- Automatically assigns persistent tracking IDs (`data-astra-id`) to every paragraph (`p_001` or `w14:paraId`), table (`t_001`), and container.
- AI agents and automation scripts can reference exact nodes without fragile string offsets or line numbers.

### 2. 27 Atomic Operations with Transactional Rollback
The `OperationEngine` dispatches across 27 distinct operation types defined as a strict discriminated union:
- **Content:** `insert`, `replace`, `delete`, `move`
- **Formatting:** `style`, `clearFormat`, `namedStyle`
- **Tables & Structure:** `table`, `section`, `pageSetup`, `pageBorder`, `coverPage`
- **Complex Elements:** `chart`, `image`, `equation`, `callout`, `field`, `list`
- **References & Annotations:** `footnote`, `endnote`, `hyperlink`, `bookmark`, `comment`, `headerFooter`, `watermark`
- **Revisions & Compliance:** `trackedChange`, `metadata`

**Transactions:** Before executing a batch, the engine captures an immutable snapshot of both the AST and package zip. If any operation violates schema rules or throws an unhandled exception, operations are automatically rolled back (`rollbackOnError: true`).

### 3. ECMA-376 Schema Validation
ASTra validates documents with `@attentionform/ooxml-validator` (the Microsoft Open XML SDK validator). Where the check runs depends on the path you use:

| Path | What it does |
|------|--------------|
| `parser.rezip()` | Serializes the package and runs its integrity checks: duplicate paragraph ids, XXE, macro policy and package integrity. It does **not** run the schema validator. A well-formed but schema-invalid edit (for example `<w:sz w:val="notanumber"/>`) passes `rezip()`. |
| `validateDocxBuffer(buffer)` and `parser.validateOOXML()` | Run the schema validator and return the result. You decide whether to write the file. Call one of them after `rezip()` and before you write. |
| `DocumentManager` (`createNewDocument`, `saveAsDocument`, `saveDocument`, and every write of an open document) | Validates the exact bytes before writing. Any reported error refuses the write, leaves the file untouched and returns memory to the last saved state. |
| MCP `save_document` | Validates on save. By default a file that already had validator errors can be saved when the edit adds no new error; `--strict-validation` refuses every reported error. |
| CLI `batch` and `normalize` | Validate before writing, with the same default policy. `--strict` refuses every reported error. `batch --no-validate` skips the check. |

Passing means the selected Open XML SDK validation run reported no errors. It does not prove how the document renders or that its content is right.

#### `rezip()` options

`parser.rezip(options?)` writes the package to a `Buffer`. The forms are `rezip()`, `rezip(options)`, `rezip(ast)`, `rezip(ast, options)` and `rezip({ ast, ...options })`.

Whatever the options, `rezip()` drops rogue zip entries, refuses macro parts under the default policy, and refuses an XML part that declares a DTD or entity (the XXE check).

| Option | Effect |
|--------|--------|
| *(none)* | Refuses two paragraphs with the same `w14:paraId` (`E_VALIDATION`, naming the id). Refuses an XML part that does not parse, has no root element, or has text outside the root. It does not catch every malformation: a mismatched closing tag passes, and `validateDocxBuffer()` reports it. |
| `validate: false` | Skips the duplicate-id check. The XML check still runs. |
| `skipValidation: true` | Skips the duplicate-id check **and** the XML check. The XXE and macro checks still run. |
| `duplicateParaIds: 'renumber'` | Repairs a repeated id and records the change, instead of refusing. The default is `'refuse'`. It has no effect with `validate: false` or `skipValidation: true`. |
| `reproducible: true` | Pins the zip entry times, the `modified` date and the revision and comment stamps to `SOURCE_DATE_EPOCH`, or to 1980-01-01Z when that is not set. The same input bytes then give the same output bytes. The `created` date is left as it is, so documents that `createNewDocument` made at different times still differ by that date. |
| `reproducible: false`, `noReproducible: true` | Real dates, even when `ASTRA_REPRODUCIBLE` or `SOURCE_DATE_EPOCH` is set. |

Without a `reproducible` option, `ASTRA_REPRODUCIBLE=1` or `SOURCE_DATE_EPOCH` turns pinning on, `ASTRA_NO_REPRODUCIBLE=1` turns it off, and otherwise the dates are the real ones.

None of these options runs the Microsoft schema validator. Call `validateDocxBuffer()` for that.


### 4. Token-Optimized Outline Notation (TOON)
Converts verbose OpenXML markup into dense, tabular outlines tailored for LLM context windows:
- **`compact`**: Terse pipe-delimited outline displaying node IDs, types, and text (about 89 % median reduction; 45 % of documents reach 90 %).
- **`standard`**: Compact outline plus inline styling flags (`[B]`, `[I]`, `[C:hex]`, `[SZ:n]`, `[HDR]`, `[ZEBRA]`).
- **`full`**: Deep structural dump including margins, borders, and XML property trees (section-scoped).

### 5. Style Guides (Omitted-Param Defaults + Advisory Linting)
Attach a parsed guide (`parseStyleGuide` in `src/style-guide/parser.ts`) as the engine's aesthetic skin:
```typescript
import { OperationEngine, parseStyleGuide } from '@attentionform/astra-engine';

const { config: styleGuide } = parseStyleGuide(guideMarkdown);
const engine = new OperationEngine(ast, parser, { styleGuide });
// or attach later: engine.setStyleGuide(styleGuide);
```
Precedence is global: **explicit op param > guide value > built-in default**. Guides fill only omitted params (dark table fills gain the guide header text by computed WCAG AA contrast, uncolored chart series pull the guide palette in order, callouts fall back to the guide theme, code blocks take the guide background/borders); caller-passed values always win. Composed Markdown has no per-table or per-alert theme syntax, so a guided compose skins tables with `tables.headerFill` (and shades the odd data rows with `tables.alternatingFill`) and `> [!TAG]` alerts with `callouts.defaultTheme` outright. A compose with no guide uses the built-in *Default Corporate* guide (navy header, bold white header text, striped rows, per-tag alert colours); pass `styleGuide: 'legacy'` for the old gray headers. Palette base colors, `typography.*`, `tables.borderColor`, `callouts.preferredThemes`, and `layout.*` are advisory-only: parsed and linted, never rendered. `lintDocumentStyle(ast, guide)` reports `{ nodeId, rule, message }` findings (header contrast, duplicate chart colors, splittable tables) without ever failing; `DocumentManager` save/validate paths carry them as additive `styleWarnings`. Guide-parse notes ride the same machine keys (`guide-parse` entries on validate, string notes on create/import/batch, MCP `warnings`/`styleGuideWarnings`). Resolve file or bundled refs with `resolveStyleGuide(ref, { allowedDir })`; unknown names and missing files fail closed with `E_VALIDATION`. Full file spec, precedence table, sandbox matrix, and session semantics: `docs/cli-guide.md` §6. Bundled ref `report` carries display title `Corporate`.

---

## API Reference

### `DocxParser`
The primary entry point for reading and parsing Word documents:

```typescript
import { DocxParser } from '@attentionform/astra-engine';

// From Buffer
const parser = new DocxParser(buffer);

// Create a fresh document from a template: blank, report, memo, letter, proposal, invoice, technical-spec, research-paper, novel or newsletter
const freshDoc = DocxParser.createNewDocument('report');

// Retrieve typed AST
const ast = await parser.getAST();

// Retrieve structural summary
const summary = parser.getStructuralSummary();

// Validate OOXML schema compliance
const validation = await parser.validateOOXML();

// Rezip into .docx Buffer with canonical namespace normalization
const outputBuffer = parser.rezip();
```

### Exact original recovery

When a history was initialized with the original file bytes, `getOriginalInfo(path)` reports whether those bytes are available. `exportOriginal(path, outputPath)` writes a hash-verified copy without overwriting an existing file by default. `restoreOriginal(path)` replaces the working file with the verified original and records the restore in history. The file-level restore API requires the caller to coordinate concurrent access and reload any live document session. Ordinary revert and redo can reserialize the package and do not guarantee identical ZIP bytes.

Originals retained in history can contain information removed later by sanitization. Older histories do not automatically acquire an original.

### `validateDocxBuffer` & `assertValidDocx`
Direct validation utility functions:

```typescript
import { validateDocxBuffer, assertValidDocx } from '@attentionform/astra-engine';

const result = await validateDocxBuffer(buffer);
if (!result.ok) {
  console.error('Validation errors:', result.errors);
}

// Throws OoxmlValidationError if invalid
await assertValidDocx(buffer);
```

### `sanitizeDocxBuffer`
Privacy and metadata scrubbing:

```typescript
import { sanitizeDocxBuffer } from '@attentionform/astra-engine';

const { buffer: cleanBuffer, report } = sanitizeDocxBuffer(buffer, {
  zeroRetention: true,
  purgePII: true,
  acceptAllRevisions: true,
  stripComments: true
});

console.log(`Scrubbed ${report.removedRSIDs} RSIDs and ${report.strippedComments} comments.`);
```

With `acceptAllRevisions`, `sanitizeDocxBuffer` throws `E_VALIDATION` (naming the part and the revision kinds) and returns no buffer when revision markup remains that cannot be resolved safely, because a partly cleaned file could still contain deleted text and reviewer names.

### TOON Generator
```typescript
import { ToonGenerator } from '@attentionform/astra-engine';

const generator = new ToonGenerator();

// Generate standard outline
const table = generator.generateToon(ast, undefined, { fidelity: 'compact' });
const outline = generator.renderToonText(table);

// Generate budget-capped outline fitting inside 1500 tokens
// (over-budget docs slice on row boundaries; page the rest with
// { offset: nextOffset } until hasMore is false)
const { rendered: budgetOutline, tokenCount, truncated, hasMore, nextOffset } =
  generator.generateTokenBudgetToon(ast, 1500);
```

---

## Building & Testing Locally

```bash
# Build engine package
npm run build -w @attentionform/astra-engine

# Run unit tests
npm test -w @attentionform/astra-engine
```

---

## License

Business Source License 1.1 (`BUSL-1.1`): source-available, free for individuals, education, non-profits, public bodies, small organizations and small products (see LICENSING.md); other production use needs a commercial license. Plans and prices: https://attentionform.com/pricing. Each version converts to Apache-2.0 four years after release. See [LICENSE](LICENSE) and write to astra@attentionform.com for a commercial license.
