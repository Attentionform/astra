import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, composeDocument, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

/**
 * A table has no spacing of its own, so a block added right under a table sat against its bottom border.
 * Compose has always given such a block 12 pt above it; the operation engine now does the same for every
 * path (batch, MCP execute_batch), unless the block names a style, already has space above it, or is a
 * page or section break. A caller can restyle it.
 */
const TABLE = { action: 'insert', type: 'table', anchor: 'end', headers: ['A', 'B'], rows: [['1', '2']], id: 't1' };

async function build(operations: unknown[]) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const res = await engine.executeBatch({ operations: operations as never, sequential: true });
  expect(res.success).toBe(true);
  return { parser, ast, engine };
}

const node = (b: Awaited<ReturnType<typeof build>>, id: string): Element =>
  b.ast.sections.flatMap((s) => s.nodes).find((n) => n.id === id)!.xml as Element;
const before = (el: Element): string | null =>
  el.getElementsByTagName('w:pPr')[0]?.getElementsByTagName('w:spacing')[0]?.getAttribute('w:before') ?? null;

describe('space under a table', () => {
  it('a paragraph added right under a table gets 12 pt above it', async () => {
    const b = await build([TABLE, { action: 'insert', type: 'paragraph', anchor: 't1', position: 'after', text: 'under', id: 'p1' }]);
    expect(before(node(b, 'p1'))).toBe('240');
  });

  it('a divider added right under a table gets it too', async () => {
    const b = await build([TABLE, { action: 'insert', type: 'divider', anchor: 't1', position: 'after', id: 'd1' }]);
    expect(before(node(b, 'd1'))).toBe('240');
  });

  it('a paragraph that follows a paragraph is left as it was', async () => {
    const b = await build([
      { action: 'insert', type: 'paragraph', anchor: 'end', text: 'first', id: 'p0' },
      { action: 'insert', type: 'paragraph', anchor: 'p0', position: 'after', text: 'second', id: 'p1' },
    ]);
    expect(before(node(b, 'p1'))).toBeNull();
  });

  it('a paragraph added above a table is left as it was', async () => {
    const b = await build([TABLE, { action: 'insert', type: 'paragraph', anchor: 't1', position: 'before', text: 'above', id: 'p1' }]);
    expect(before(node(b, 'p1'))).toBeNull();
  });

  it('a page break under a table is not spaced', async () => {
    const b = await build([TABLE, { action: 'insert', type: 'pageBreak', anchor: 't1', position: 'after', id: 'br1' }]);
    expect(before(node(b, 'br1'))).toBeNull();
  });

  it('a block that a later paragraph moves away from the table keeps what it got', async () => {
    const b = await build([
      TABLE,
      { action: 'insert', type: 'paragraph', anchor: 't1', position: 'after', text: 'under', id: 'p1' },
      { action: 'insert', type: 'paragraph', anchor: 'p1', position: 'after', text: 'next', id: 'p2' },
    ]);
    expect(before(node(b, 'p1'))).toBe('240');
    expect(before(node(b, 'p2'))).toBeNull();
  });

  it('a caller can change the space with a style operation, and it stays changed', async () => {
    const b = await build([
      TABLE,
      { action: 'insert', type: 'paragraph', anchor: 't1', position: 'after', text: 'under', id: 'p1' },
      { action: 'style', target: 'p1', spacingBefore: 600 },
    ]);
    expect(before(node(b, 'p1'))).toBe('600');
  });

  it('a caller can set it to nothing', async () => {
    const b = await build([
      TABLE,
      { action: 'insert', type: 'paragraph', anchor: 't1', position: 'after', text: 'under', id: 'p1' },
      { action: 'style', target: 'p1', spacingBefore: 0 },
    ]);
    expect(before(node(b, 'p1'))).toBe('0');
  });

  it('a heading composed under a table has the space once (compose gives it; the engine adds nothing to a styled block)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const res = await composeDocument(parser, '| A |\n| --- |\n| 1 |\n\n# Heading\n');
    expect(res.success).toBe(true);
    const xml = parser.readDocumentXml();
    const heading = (xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find((p) => p.includes('w:val="Heading1"'))!;
    expect(heading).toBeDefined();
    expect((heading.match(/<w:spacing\b/g) ?? []).length).toBe(1);
  });

  it('composed text under a table has the space once', async () => {
    const parser = DocxParser.createNewDocument('blank');
    await composeDocument(parser, '| A |\n| --- |\n| 1 |\n\nunder the table\n');
    const xml = parser.readDocumentXml();
    expect((xml.match(/w:before="240"/g) ?? []).length).toBe(1);
  });
});
