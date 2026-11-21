# @attentionform/astra-engine

The engine behind ASTra, as a package for Node.js. It opens a Word (`.docx`) file, gives each paragraph and table a stable id, changes them with typed operations that either all apply or none do,
checks the result against the Office schema, and keeps a history that you can step back through. It runs on your machine and needs Node.js 22.19 or newer.

Full documentation, with an API reference for every exported name, is at https://attentionform.com/docs/library.

---

## Install and first use

```bash
npm install @attentionform/astra-engine
```

### Read and change a document

```typescript
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';
import * as fs from 'node:fs';

async function main() {
  // 1. Open an existing .docx file
  const fileBuffer = fs.readFileSync('document.docx');
  const parser = new DocxParser(fileBuffer);
  
  // 2. Read the document as typed data (the AST)
  const ast = await parser.getAST();
  console.log(`Document loaded with ${ast.sections.length} section(s).`);

  // 3. Apply a batch of operations. If one fails, none of them is applied
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
    
    // 4. Write the package, then validate it before you save it.
    //    rezip() checks the package (repeated paragraph ids, XML external entities, macros, integrity).
    //    It does not run the Office schema validator. validateDocxBuffer() does.
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

const parser = DocxParser.createNewDocument('memo', { page: { size: 'A4' } }); // Letter portrait is the default
// Any paper and margins: { size: '148x210mm', orientation: 'landscape', margins: '15mm' }
fs.writeFileSync('memo.docx', parser.rezip());
```

`blank` is one empty paragraph. The other nine are designed starting documents with placeholder text in square brackets. A template is built for the page you ask for, so its tables, tab stops and cover fit that paper: `page.size` is a paper name (Letter, Legal, Tabloid, Executive, Statement, Folio, A3, A4, A5, A6, B4, B5, B6) or a size such as `"6x9in"` or `{ width: "150mm", height: "200mm" }` (1 in to 22 in a side), `page.orientation` is `portrait` or `landscape`, and `page.margins` is one length or `{ top, bottom, left, right, header, footer, gutter }`. A page that cannot be laid out throws `E_VALIDATION`; `parser.newDocumentWarnings` lists a valid page that suits the template poorly. The old `pageSize` option was replaced by `page.size`. Runnable examples are in `packages/engine/examples/`.

### Typed operations and untyped data

`executeBatch` takes typed operations. An unknown `action`, an unknown field or a missing required field is a compile error. Annotate your own values with `OperationInput` (one
operation) and `BatchPayload` (the whole batch):

```typescript
import type { BatchPayload, OperationInput } from '@attentionform/astra-engine';

const operations: OperationInput[] = [{ action: 'replace', target: 'p_001', text: 'New text' }];
const batch: BatchPayload = { operations };
await engine.executeBatch(batch);
```

Some rules cannot be written as a type: a field that is required only for one sub-operation, or a rule that spans several fields. `executeBatch` checks those when it
runs, and `parseBatch` checks them earlier. For data that arrives untyped (JSON from a file, a request, a model), call `parseBatch` first. It returns
`{ ok: true, batch }` with typed operations, or `{ ok: false, errors }` with every problem and where it is:

```typescript
import { parseBatch } from '@attentionform/astra-engine';

const parsed = parseBatch(JSON.parse(text));
if (!parsed.ok) throw new Error(parsed.errors.map((e) => e.message).join('; '));
await engine.executeBatch(parsed.batch);
```

`BatchInput` is the form after parsing, not the form that you write. Use `BatchPayload` for a batch that you write by hand.

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
Use `hasPdfBackend()` for that question. The [render example](./examples/10-render.mjs) shows both.

### Handling errors

Handle a failure by its code: `getEngineErrorCode(err)` returns one of `ENGINE_ERROR_CODES`, or `undefined`, and `getEngineRemedy(code)` returns the default advice for a code. The classes (`EngineQuotaError`, `EngineValidationError`, `OoxmlValidationError`, `PreflightError`, `RenderError`,
`WordRenderError`, `WindowsPdfError`, `DependencyError`) are exported for `instanceof` checks, but the code is the contract.
The [validation example](./examples/08-validate.mjs) shows it.

### Reading an older version

Open a document with `DocumentManager` to get undo, redo, the version list and original recovery. To preview what a revert would change, or to read an
older state as a document, use `getHistory`, `checkoutCommit` and `trackedPredicateFor` from `/history`.
The [history example](./examples/09-history-older-version.mjs) previews additions, removals and changes of every recorded part,
including an image that a later state added, and reads the older document back and validates it.

### Reading inside a .docx

The engine does not export its zip helper. To look inside a `.docx` yourself, use any zip library, for example `adm-zip`.

---

## How the engine works

### Stable node ids
- The package is read in memory. Nothing is extracted to disk.
- Each paragraph (`p_001`, or its `w14:paraId`), table (`t_001`) and container gets an id that stays the same between edits.
- An operation names a node by its id, so an agent or a script never depends on a text offset or a line number.

### Operations and batches
`OperationEngine` applies 27 kinds of operation. Each has an `action` name, and TypeScript checks the fields of each one.
- **Content:** `insert`, `replace`, `delete`, `move`
- **Formatting:** `style`, `clearFormat`, `namedStyle`
- **Tables and structure:** `table`, `section`, `pageSetup`, `pageBorder`, `coverPage`
- **Complex elements:** `chart`, `image`, `equation`, `callout`, `field`, `list`
- **References and notes:** `footnote`, `endnote`, `hyperlink`, `bookmark`, `comment`, `headerFooter`, `watermark`
- **Revisions and properties:** `trackedChange`, `metadata`

A batch is all or nothing. Before it runs, the engine keeps a copy of the AST and the package. When an operation breaks a schema rule or throws, the batch is rolled back
and `result.rollbackTriggered` is true.

### Validation
ASTra validates documents with `@attentionform/ooxml-validator`, which uses the Microsoft Open XML SDK validator. Where the check runs depends on the path you use:

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


### The outline for agents
An agent reads a compact table in place of the XML: one row per paragraph or table, with its id, style, text and flags. It is much shorter than the markup, so more of a document fits in a model's context.
`ToonGenerator` builds it at four levels of detail:
- **`outline`**: the headings only.
- **`compact`**: every node, with the node id, type and text.
- **`standard`**: the default. It adds style flags such as `[B]`, `[I]`, `[C:hex]`, `[SZ:n]`, `[HDR]` and `[ZEBRA]`.
- **`full`**: a deep dump of one section, with margins, borders and the XML properties.

### Style guides
A style guide is a Markdown file that sets the look of tables, charts, callouts and code blocks. Read it with `parseStyleGuide` and give it to the engine:

```typescript
import { OperationEngine, parseStyleGuide } from '@attentionform/astra-engine';

const { config: styleGuide } = parseStyleGuide(guideMarkdown);
const engine = new OperationEngine(ast, parser, { styleGuide });
// or later: engine.setStyleGuide(styleGuide);
```

- **Precedence:** an explicit parameter wins over the guide, and the guide wins over the built-in value. A guide fills in only the parameters that you leave out.
  When an explicit value clashes with the guide, you get a warning and the explicit value is used.
- **What a guide sets:** table header fills and text colour (chosen for WCAG AA contrast), the colours of chart series that have none, the theme of callouts, and the background and borders of code blocks.
  Colours of the palette, `typography.*`, `tables.borderColor`, `callouts.preferredThemes` and `layout.*` are read and checked, but they are not drawn.
- **Composing from Markdown:** Markdown has no syntax for a table or alert theme, so the guide's `tables.headerFill` and `callouts.defaultTheme` apply to composed tables and `> [!TAG]` alerts.
  Without a guide, the built-in Default Corporate guide applies. Pass `styleGuide: 'legacy'` for the older look with grey table headers.
- **Checking:** `lintDocumentStyle(ast, guide)` returns warnings (`{ nodeId, rule, message }`) for header contrast, repeated chart colours and tables that can split across pages. A warning never fails a save.
  `DocumentManager` returns them as `styleWarnings`.
- **Finding a guide:** `resolveStyleGuide(ref, { allowedDir })` reads a bundled guide by name (`report`, `legal` or `technical`) or a `.md` file. An unknown name or a missing file throws `E_VALIDATION`.

See [Style guides](https://attentionform.com/docs/guides/style-guides) for the file format.

---

## More examples

Every exported name, with its declaration and description, is in the [API reference](https://attentionform.com/docs/library/api). These are the most used parts.

### `DocxParser`
The primary entry point for reading and parsing Word documents:

```typescript
import { DocxParser } from '@attentionform/astra-engine';

// From Buffer
const parser = new DocxParser(buffer);

// Create a fresh document from a template: blank, report, memo, letter, proposal, invoice, technical-spec, research-paper, novel or newsletter
const freshDoc = DocxParser.createNewDocument('report');

// Read the typed AST
const ast = await parser.getAST();

// Read the structural summary
const summary = parser.getStructuralSummary();

// Check the document against the Office schema
const validation = await parser.validateOOXML();

// Write the package as .docx bytes
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

With `acceptAllRevisions`, `sanitizeDocxBuffer` throws `E_VALIDATION` (naming the part and the revision kinds) and returns no buffer when revision markup remains that cannot be resolved safely, because a partly cleaned file could still contain deleted text and author names.

### Outline generator (`ToonGenerator`)
```typescript
import { ToonGenerator } from '@attentionform/astra-engine';

const generator = new ToonGenerator();

// Build an outline
const table = generator.generateToon(ast, undefined, { fidelity: 'compact' });
const outline = generator.renderToonText(table);

// Generate budget-capped outline fitting inside 1500 tokens
// (over-budget docs slice on row boundaries; page the rest with
// { offset: nextOffset } until hasMore is false)
const { rendered: budgetOutline, tokenCount, truncated, hasMore, nextOffset } =
  generator.generateTokenBudgetToon(ast, 1500);
```

---

## License

Business Source License 1.1 (`BUSL-1.1`): source-available. Who can use it for free and who needs a commercial license is set out at https://attentionform.com/licensing. Plans and prices: https://attentionform.com/pricing. Each version converts to Apache-2.0 four years after release. See [LICENSE](LICENSE) and write to astra@attentionform.com for a commercial license.
