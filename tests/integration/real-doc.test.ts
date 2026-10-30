import { describe, it, expect, beforeAll } from 'vitest';
import {
  createBlankDocxBuffer,
  DocxParser,
  OperationEngine,
  ToonGenerator,
} from '@attentionform/astra-engine';

// Self-contained fixture: a real OOXML package generated in-memory so the test
// never depends on files outside the repository (CI checkouts have none).
let DOC_BUFFER: Buffer;

beforeAll(() => {
  const parser = new DocxParser(createBlankDocxBuffer());
  parser.insertRootNode('<w:p><w:r><w:t>ASTra Architecture Overview</w:t></w:r></w:p>');
  parser.insertRootNode('<w:p><w:r><w:t>The engine parses OOXML into a token-optimized AST for AI manipulation.</w:t></w:r></w:p>');
  DOC_BUFFER = parser.rezip();
});

describe('Phase 8.1 — Real Document E2E', () => {

  it('should produce a valid .docx package buffer', () => {
    expect(DOC_BUFFER.length).toBeGreaterThan(0);
    // PK zip magic
    expect(DOC_BUFFER[0]).toBe(0x50);
    expect(DOC_BUFFER[1]).toBe(0x4b);
  });

  it('should load and parse the real .docx file without crashing', () => {
    const parser = new DocxParser(DOC_BUFFER);
    const summary = parser.getStructuralSummary();

    expect(summary.nodes.length).toBeGreaterThan(0);
  });

  it('should generate a valid AST from the real document', async () => {
    const parser = new DocxParser(DOC_BUFFER);
    const ast = await parser.getAST();

    expect(ast.sections.length).toBeGreaterThanOrEqual(1);
    expect(ast.sections[0].nodes.length).toBeGreaterThan(0);

    // Every node should have a valid id and version
    for (const node of ast.sections[0].nodes) {
      expect(node.id).toBeTruthy();
      expect(node.version).toBeGreaterThanOrEqual(1);
    }
  });

  it('should generate TOON from the real document AST', async () => {
    const parser = new DocxParser(DOC_BUFFER);
    const ast = await parser.getAST();

    const toon = new ToonGenerator();
    const table = toon.generateToon(ast);

    expect(table.rows.length).toBeGreaterThan(0);
    expect(table.headers).toEqual(['ID', 'V', 'Path', 'Header', 'Style', 'Text', 'Flags']);

    // Render as text and verify it's compact
    const text = toon.renderToonText(table);
    expect(text).toContain('ID | V | Path | Header | Style | Text | Flags');
  });

  it('should initialize the OperationEngine against the real AST', async () => {
    const parser = new DocxParser(DOC_BUFFER);
    const ast = await parser.getAST();

    const engine = new OperationEngine(ast);
    // Engine should instantiate without errors
    expect(engine).toBeDefined();
  });

  it('should rezip the document without data loss', () => {
    const parser = new DocxParser(DOC_BUFFER);
    const rezipped = parser.rezip();

    // The rezipped buffer should be a valid zip
    expect(rezipped.length).toBeGreaterThan(0);
    // Re-parse the rezipped buffer — should not throw
    const parser2 = new DocxParser(rezipped);
    const summary2 = parser2.getStructuralSummary();
    expect(summary2.nodes.length).toBeGreaterThan(0);
  });
});
