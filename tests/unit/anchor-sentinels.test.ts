import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, OperationSchema, validateOperationStrict, validateOperationsStrict } from '@attentionform/astra-engine';
import { anchorAliasOf } from '../../packages/engine/src/anchor-alias';

async function open() {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const texts = (ast: { sections: Array<{ nodes: Array<{ xml: Element }> }> }) =>
  ast.sections.flatMap((s) => s.nodes).map((n) => n.xml.textContent ?? '');

describe('start and end anchors', () => {
  it('anchorAliasOf accepts start, end and the @ spellings, and nothing else', () => {
    expect(anchorAliasOf('start')).toBe('start');
    expect(anchorAliasOf('@start')).toBe('start');
    expect(anchorAliasOf('end')).toBe('end');
    expect(anchorAliasOf('@end')).toBe('end');
    for (const v of ['End', 'START', '@End', 'p_1', '', undefined, 5]) expect(anchorAliasOf(v)).toBeNull();
  });

  it("position is optional with anchor 'end' (after) and 'start' (before)", async () => {
    const { ast, engine } = await open();
    const out = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'LAST' },
        { action: 'insert', type: 'paragraph', anchor: 'start', text: 'FIRST' },
      ] as never,
      sequential: true,
    });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    const all = texts(ast);
    expect(all[0]).toBe('FIRST');
    expect(all[all.length - 1]).toBe('LAST');
  });

  it('@start and @end work like start and end', async () => {
    const { ast, engine } = await open();
    const out = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: '@end', position: 'after', text: 'LAST' },
        { action: 'insert', type: 'paragraph', anchor: '@start', position: 'before', text: 'FIRST' },
      ] as never,
      sequential: true,
    });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    const all = texts(ast);
    expect(all[0]).toBe('FIRST');
    expect(all[all.length - 1]).toBe('LAST');
  });

  it('an explicit position still wins: end with before goes before the last block', async () => {
    const { ast, engine } = await open();
    const before = texts(ast);
    await engine.executeBatch({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'before', text: 'X' }] as never });
    const all = texts(ast);
    expect(all[all.length - 2]).toBe('X');
    expect(all[all.length - 1]).toBe(before[before.length - 1]);
  });

  it('other anchor-carrying operations take the default too (callout at the end)', async () => {
    const { ast, engine } = await open();
    const out = await engine.executeBatch({ operations: [{ action: 'callout', anchor: 'end', title: 'Note', text: 'Body' }] as never });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(texts(ast).join('|')).toContain('Body');
  });

  it('a body insert with no anchor is still refused, and the message names start and end', () => {
    expect(() => validateOperationStrict({ action: 'insert', type: 'paragraph', text: 'x', position: 'after' })).toThrow(/'start' \/ 'end'/);
  });

  it('a body insert with a node id still needs a position', () => {
    expect(() => validateOperationStrict({ action: 'insert', type: 'paragraph', text: 'x', anchor: 'p_1' })).toThrow(/position is required/);
  });

  it('the schema validators accept the shorthand', () => {
    expect(() => validateOperationsStrict([{ action: 'insert', type: 'paragraph', text: 'x', anchor: 'end' }])).not.toThrow();
    expect(OperationSchema.safeParse({ action: 'insert', type: 'paragraph', text: 'x', anchor: '@end' }).success).toBe(true);
  });

  it('the insert anchor says what start and end mean', () => {
    const insert = OperationSchema.options.find((o) => (o.shape.action as { value: string }).value === 'insert')!;
    const description: string = (insert.shape as Record<string, any>).anchor.unwrap().description;
    expect(description).toContain("'start'");
    expect(description).toContain("'end'");
    expect(description).toContain('@end');
  });
});
