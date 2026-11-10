import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, ToonGenerator, tableGridPlainText } from '@attentionform/astra-engine';
import { composeDocument } from '../../packages/engine/src/compose/markdown';
import { collectAstIdTexts } from '../../packages/mcp-server/src/lib/mcp/diff';

const MARKDOWN = [
  '| Vendor | Cost | SLA |',
  '| --- | --- | --- |',
  '| Alpha | 10 | 99.9 |',
  '| Beta | 20 | 99.5 |',
  '',
  '> [!IMPORTANT]',
  '> Back up the file first.',
  '',
].join('\n');

async function composed() {
  const parser = DocxParser.createNewDocument('blank');
  await composeDocument(parser, MARKDOWN);
  return { parser, ast: await parser.getAST() };
}

describe('table text keeps its separators', () => {
  it('the saved-values text reads as rows of cells: " | " between cells, a newline between rows', async () => {
    const { ast } = await composed();
    const table = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'table' && /Vendor/.test(n.metadata.text ?? ''));
    expect(table).toBeDefined();
    expect(tableGridPlainText(table!.xml)).toBe('Vendor | Cost | SLA\nAlpha | 10 | 99.9\nBeta | 20 | 99.5');
    const { texts } = collectAstIdTexts(ast as never);
    expect(texts.get(table!.id)).toBe('Vendor | Cost | SLA\nAlpha | 10 | 99.9\nBeta | 20 | 99.5');
    expect(texts.get(table!.id)).not.toMatch(/SLAAlpha|1099\.9/);
  });

  it('a callout title and body are two words in the inspect text', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const out = await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'callout', title: 'Important', text: 'Back up the file first.' }] as never,
    });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    const rows = new ToonGenerator().toon(ast, { fidelity: 'standard' }).rows;
    const callout = rows.find((r) => /Back up the file first/.test(r.text));
    expect(callout).toBeDefined();
    expect(callout!.text).toMatch(/Important Back up the file first\./);
    expect(callout!.text).not.toMatch(/ImportantBack/);
  });

  it('a paragraph that is not in a table keeps its own text', async () => {
    const { ast } = await composed();
    const { texts, order } = collectAstIdTexts(ast as never);
    const paragraph = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'paragraph' && n.metadata.text);
    if (paragraph) expect(texts.get(paragraph.id)).toBe(paragraph.metadata.text);
    expect(order.length).toBeGreaterThan(0);
  });
});
