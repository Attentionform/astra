import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const T2 = path.resolve(__dirname, '../../test-gallery/T2_edit_target.docx');

/** T2 plus a chart node, an image node, and a paragraph that holds a picture next to text. */
async function loadBase() {
  const seed = new DocxParser(fs.readFileSync(T2));
  const seedAst = await seed.getAST();
  const first = seedAst.sections[0].nodes[0].id;
  const engine = new OperationEngine(seedAst, seed);
  const r = await engine.executeBatch({
    rollbackOnError: true,
    operations: [
      { action: 'insert', type: 'chart', anchor: first, position: 'after', chartType: 'column', title: 'T',
        categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }], id: 'chart_1' },
      { action: 'image', operation: 'insert', anchor: 'chart_1', position: 'after',
        fileData: 'data:image/png;base64,' + PNG, width: 914400, height: 914400, altText: 'a', id: 'img_1' },
      { action: 'image', operation: 'insert', anchor: 'img_1', position: 'after',
        fileData: 'data:image/png;base64,' + PNG, width: 914400, height: 914400, altText: 'b', id: 'img_2' },
    ] as never,
  });
  expect(r.success).toBe(true);
  const mixed = seedAst.sections.flatMap((s) => s.nodes).find((n) => n.id === 'img_2')!.xml as Element;
  const run = mixed.ownerDocument.createElement('w:r');
  const t = mixed.ownerDocument.createElement('w:t');
  t.setAttribute('xml:space', 'preserve');
  t.textContent = 'Text beside the picture. ';
  run.appendChild(t);
  const pPr = mixed.getElementsByTagName('w:pPr')[0];
  mixed.insertBefore(run, pPr ? pPr.nextSibling : mixed.firstChild);
  const bytes = seed.rezip();
  const parser = new DocxParser(bytes);
  const ast = await parser.getAST();
  const nodes = ast.sections.flatMap((s) => s.nodes);
  return {
    parser, ast, bytes,
    ids: {
      table: nodes.find((n) => n.type === 'table')!.id,
      chart: nodes.find((n) => n.type === 'chart')!.id,
      image: nodes.find((n) => n.type === 'image')!.id,
      mixed: nodes.find((n) => n.type === 'paragraph' && n.metadata.hasImage === true)!.id,
      paragraph: nodes.find((n) => n.type === 'paragraph' && !n.metadata.hasImage && (n.metadata.text ?? '').length > 0)!.id,
    },
  };
}

async function run(op: Record<string, unknown>) {
  const { parser, ast, bytes, ids } = await loadBase();
  const res: any = await new OperationEngine(ast, parser).executeBatch({ operations: [op] as never, rollbackOnError: true });
  return { res, unchanged: Buffer.compare(parser.rezip(), bytes) === 0, ids };
}

// errors[0] is the operation's own error; the last entry is the batch wrapper, which always carries E_VALIDATION.
const firstError = (res: any): string => String(res.errors?.[0]?.error ?? '');

describe('paragraph-only operations refuse table nodes', () => {
  const cases: Array<[string, (id: string) => Record<string, unknown>, string]> = [
    ['list.create', (id) => ({ action: 'list', operation: 'create', target: [id], listType: 'bullet' }), 'applies to paragraphs'],
    ['namedStyle.apply', (id) => ({ action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: id }), 'applies a paragraph style'],
    ['hyperlink', (id) => ({ action: 'hyperlink', target: id, url: 'https://example.com' }), 'applies to paragraph text'],
  ];
  it.each(cases)('%s on a table gives E_VALIDATION and changes nothing', async (_name, build, wording) => {
    const { ids } = await loadBase();
    const { res, unchanged } = await run(build(ids.table));
    expect(res.success).toBe(false);
    expect(firstError(res)).toContain('[E_VALIDATION]');
    expect(firstError(res)).toContain(wording);
    expect(unchanged).toBe(true);
  });
});

describe('equation replace refuses nodes that hold an object', () => {
  it.each(['image', 'chart', 'mixed', 'table'] as const)('equation.replace on a %s node gives E_VALIDATION and changes nothing', async (kind) => {
    const { ids } = await loadBase();
    const { res, unchanged } = await run({ action: 'equation', operation: 'replace', target: ids[kind], latex: 'x' });
    expect(res.success).toBe(false);
    expect(firstError(res)).toContain('[E_VALIDATION]');
    expect(unchanged).toBe(true);
  });

  it.each([
    ['comment', (id: string) => ({ action: 'comment', operation: 'insert', target: id, text: 'review this' })],
    ['footnote', (id: string) => ({ action: 'footnote', anchor: id, text: 'a note' })],
    ['endnote', (id: string) => ({ action: 'endnote', anchor: id, text: 'a note' })],
  ] as Array<[string, (id: string) => Record<string, unknown>]>)(
    'equation.replace on a paragraph with a %s anchor gives E_VALIDATION and changes nothing', async (_kind, anchorOp) => {
      const { parser, ast, ids } = await loadBase();
      const seeded = await new OperationEngine(ast, parser).executeBatch({ operations: [anchorOp(ids.paragraph)] as never, rollbackOnError: true });
      expect(seeded.success).toBe(true);
      const bytes = parser.rezip();
      const reloaded = new DocxParser(bytes);
      const res: any = await new OperationEngine(await reloaded.getAST(), reloaded)
        .executeBatch({ operations: [{ action: 'equation', operation: 'replace', target: ids.paragraph, latex: 'x' }] as never, rollbackOnError: true });
      expect(res.success).toBe(false);
      expect(firstError(res)).toContain('[E_VALIDATION]');
      expect(firstError(res)).toContain('holds a picture, chart, comment or note anchor');
      expect(Buffer.compare(reloaded.rezip(), bytes)).toBe(0);
    });

  it('equation.replace on a plain paragraph still works', async () => {
    const { ids } = await loadBase();
    const { res } = await run({ action: 'equation', operation: 'replace', target: ids.paragraph, latex: 'x' });
    expect(res.success).toBe(true);
  });
});
