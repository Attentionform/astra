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
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
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
    
    // 4. Validate ECMA-376 schema and serialize back to .docx buffer
    const outBuffer = parser.rezip();
    fs.writeFileSync('document_updated.docx', outBuffer);
  } else {
    console.error('Batch failed:', result.errors);
  }
}

main();
```

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

### 3. Pre-Write ECMA-376 Schema Gate
ASTra integrates `@attentionform/ooxml-validator` directly into its serialization pipeline. No mutated file is committed to storage unless every internal XML part passes strict OpenXML schema validation.

### 4. Token-Optimized Outline Notation (TOON)
Converts verbose OpenXML markup into dense, tabular outlines tailored for LLM context windows:
- **`compact`**: Terse pipe-delimited outline displaying node IDs, types, and text (~90%+ token reduction).
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

// Create fresh document from template ('report' | 'memo' | 'blank')
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

Business Source License 1.1 (`BUSL-1.1`): source-available, free for individuals, education, non-profits, public bodies and small businesses; other production use needs a commercial license. Each version converts to Apache-2.0 four years after release. See [LICENSE](LICENSE) and write to astra@attentionform.com for a commercial license.
