/**
 * List levels count from 1: level 1 is the top level (w:ilvl 0 in the file), the
 * inspect output prints the same numbers ([L1], list:bullet:L1), and on
 * changeLevel 0 takes the paragraph out of the list. Out-of-range levels are
 * refused with the valid range before anything is created.
 */
import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, ToonGenerator, createBlankDocxBuffer, type DocumentAST } from '@attentionform/astra-engine';
import { OperationSchema } from '../../packages/engine/src/operation-schema';
import { useReproducibleOutput } from '../helpers/reproducible';

// The refusal test compares package bytes before and after; without pinned output they carry the clock.
useReproducibleOutput();

interface Fixture {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
}

/** Three paragraphs li_1..li_3, none in a list yet. */
async function paragraphs(): Promise<Fixture> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const seeded = await engine.executeBatch({
    operations: ['li_1', 'li_2', 'li_3'].map((id) => ({ action: 'insert', type: 'paragraph', anchor: 'end', text: `item ${id}`, id })) as never,
    sequential: true,
  });
  expect(seeded.success).toBe(true);
  return { parser, ast, engine };
}

const node = (f: Fixture, id: string) => f.ast.sections.flatMap((s) => s.nodes).find((n) => n.id === id)!;
const ilvl = (f: Fixture, id: string): string | null =>
  (node(f, id).xml as Element).getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val') ?? null;
const run = (f: Fixture, op: Record<string, unknown>) => f.engine.executeBatch({ operations: [op] as never });

describe('list levels count from 1', () => {
  it('create at level 1 is the top level (w:ilvl 0); level 3 is w:ilvl 2', async () => {
    const f = await paragraphs();
    expect((await run(f, { action: 'list', operation: 'create', target: ['li_1'], listType: 'bullet', level: 1 })).success).toBe(true);
    expect((await run(f, { action: 'list', operation: 'create', target: ['li_2'], listType: 'bullet', level: 3 })).success).toBe(true);
    expect(ilvl(f, 'li_1')).toBe('0');
    expect(ilvl(f, 'li_2')).toBe('2');
  });

  it('create without a level is the top level, and inspect prints the same numbers', async () => {
    const f = await paragraphs();
    await run(f, { action: 'list', operation: 'create', target: ['li_1', 'li_2'], listType: 'numbered' });
    await run(f, { action: 'list', operation: 'changeLevel', target: 'li_2', level: 2 });
    const rows = new ToonGenerator().toon(f.ast, { fidelity: 'standard' }).rows;
    const text = JSON.stringify(rows.filter((r) => r.id === 'li_1' || r.id === 'li_2'));
    expect(rows.find((r) => r.id === 'li_1')!.style).toBe('list:numbered:L1');
    expect(rows.find((r) => r.id === 'li_2')!.style).toBe('list:numbered:L2');
    expect(text).toContain('[L1]');
    expect(text).toContain('[L2]');
    expect(text).not.toContain('[L0]');
  });

  it('refuses level 0 and levels past 9 on create, before creating any list', async () => {
    for (const bad of [0, 10, -1, 1.5]) {
      const f = await paragraphs();
      const before = f.parser.rezip().toString('base64');
      const out = await run(f, { action: 'list', operation: 'create', target: ['li_1'], listType: 'bullet', params: { level: bad } });
      expect(out.success).toBe(false);
      expect(JSON.stringify(out.errors)).toMatch(new RegExp(`list\\.create level ${String(bad).replace('.', '\\.')} is out of range`));
      expect(JSON.stringify(out.errors)).toMatch(/pass level in \[1, 9\], counting from 1 \(1 is the top level\)/);
      expect(f.parser.rezip().toString('base64')).toBe(before);
    }
  });

  it('the schema takes 0 to 9 and describes the numbering', () => {
    for (const level of [0, 1, 9]) {
      expect(OperationSchema.safeParse({ action: 'list', operation: 'changeLevel', target: 'li_1', level }).success).toBe(true);
    }
    for (const level of [-1, 10]) {
      expect(OperationSchema.safeParse({ action: 'list', operation: 'changeLevel', target: 'li_1', level }).success).toBe(false);
    }
  });
});

describe('changeLevel', () => {
  it('level N moves the item to level N; level 1 brings it back to the top and keeps it in the list', async () => {
    const f = await paragraphs();
    await run(f, { action: 'list', operation: 'create', target: ['li_1'], listType: 'bullet' });
    expect((await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', level: 3 })).success).toBe(true);
    expect(ilvl(f, 'li_1')).toBe('2');
    expect(node(f, 'li_1').metadata.listLevel).toBe(2);
    expect((await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', level: 1 })).success).toBe(true);
    expect(ilvl(f, 'li_1')).toBe('0');
    expect((node(f, 'li_1').xml as Element).getElementsByTagName('w:numPr').length).toBe(1);
  });

  it('level 0 takes the paragraph out of the list', async () => {
    const f = await paragraphs();
    await run(f, { action: 'list', operation: 'create', target: ['li_1'], listType: 'bullet' });
    expect((await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', level: 0 })).success).toBe(true);
    expect((node(f, 'li_1').xml as Element).getElementsByTagName('w:numPr').length).toBe(0);
    expect(node(f, 'li_1').metadata.listLevel).toBeUndefined();
  });

  it('refuses a level past 9 with the range, 0 included in the remedy', async () => {
    const f = await paragraphs();
    await run(f, { action: 'list', operation: 'create', target: ['li_1'], listType: 'bullet' });
    const out = await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', params: { level: 12 } });
    expect(out.success).toBe(false);
    expect(JSON.stringify(out.errors)).toMatch(/list\.changeLevel level 12 is out of range/);
    expect(JSON.stringify(out.errors)).toMatch(/or 0 to take the paragraph out of the list/);
  });

  it('delta moves by levels: deeper, shallower, and out of the list below level 1', async () => {
    const f = await paragraphs();
    await run(f, { action: 'list', operation: 'create', target: ['li_1'], listType: 'bullet' });
    await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', delta: 1 });
    expect(ilvl(f, 'li_1')).toBe('1'); // level 2
    await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', delta: 1 });
    expect(ilvl(f, 'li_1')).toBe('2'); // level 3
    // Shallower from level 3 lands on level 2, still in the list.
    await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', delta: -1 });
    expect(ilvl(f, 'li_1')).toBe('1');
    // Shallower from level 2 lands on the top level, still in the list.
    await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', delta: -1 });
    expect(ilvl(f, 'li_1')).toBe('0');
    expect((node(f, 'li_1').xml as Element).getElementsByTagName('w:numPr').length).toBe(1);
    // Shallower from the top level leaves the list.
    await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', delta: -1 });
    expect((node(f, 'li_1').xml as Element).getElementsByTagName('w:numPr').length).toBe(0);
  });

  it('delta on a paragraph outside any list makes it a top-level item of a real list', async () => {
    const f = await paragraphs();
    expect((await run(f, { action: 'list', operation: 'changeLevel', target: 'li_3', delta: 1 })).success).toBe(true);
    expect(ilvl(f, 'li_3')).toBe('0');
    expect(node(f, 'li_3').metadata.listLevel).toBe(0);
    // The list id the paragraph points at exists in the numbering part.
    const numId = (node(f, 'li_3').xml as Element).getElementsByTagName('w:numId')[0]?.getAttribute('w:val') ?? '';
    expect(numId).not.toBe('');
    expect(Object.keys(f.ast.numbering?.concreteNum ?? {})).toContain(numId);
    expect(node(f, 'li_3').metadata.listStyle).toBe('bullet');
  });

  it('level N on a paragraph outside any list joins a list of the requested type', async () => {
    const f = await paragraphs();
    expect((await run(f, { action: 'list', operation: 'changeLevel', target: 'li_3', level: 3, listType: 'numbered' })).success).toBe(true);
    expect(ilvl(f, 'li_3')).toBe('2');
    const numId = (node(f, 'li_3').xml as Element).getElementsByTagName('w:numId')[0]?.getAttribute('w:val') ?? '';
    expect(Object.keys(f.ast.numbering?.concreteNum ?? {})).toContain(numId);
    expect(node(f, 'li_3').metadata.listStyle).toBe('numbered');
  });
});
