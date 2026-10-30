/**
 * Word-level match error-surface pins.
 *
 * A `match` miss ("text ... not found in node ...") carries its own
 * inspect-and-retry remedy and must NOT take the stale node-id hint: the
 * hint fires for every other "not found" (exclusion semantics — node misses
 * in any casing, e.g. "Anchor node '...' not found"), and only the
 * "not found in node" word-miss shape is excluded. The schema trio below
 * snapshots the discoverability annotations.
 */
import { describe, it, expect } from 'vitest';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { OperationSchema } from '../src/operation-schema';
import type { Node } from '../src/ast';

async function freshEngine() {
  const parser = new DocxParser(createBlankDocxBuffer('Match Errors'));
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}

function lastNode(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const nodes = ast.sections[0].nodes;
  return nodes[nodes.length - 1];
}

async function writeParagraph(engine: OperationEngine, ast: { sections: Array<{ nodes: Node[] }> }, text: string): Promise<Node> {
  const anchor = lastNode(ast).id;
  await engine.execute({ action: 'insert', type: 'paragraph', anchor, position: 'after', text } as never);
  return lastNode(ast);
}

async function batchErrors(engine: OperationEngine, operations: unknown) {
  const res = await engine.executeBatch({ operations: operations as never, sequential: false, rollbackOnError: false });
  expect(res.success).toBe(false);
  expect(res.errors.length).toBeGreaterThan(0);
  return res.errors[0];
}

function styleShape(): Record<string, { description?: string }> {
  const union = OperationSchema as unknown as {
    options: Array<{ shape: Record<string, { description?: string; value?: unknown }> }>;
  };
  const style = union.options.find((o) => o.shape.action?.value === 'style');
  if (!style) throw new Error('style option missing from OperationSchema');
  return style.shape;
}

function nestedShape(field: { description?: string } & Record<string, unknown>): Record<string, { description?: string }> {
  const inner = (field as { _def: { innerType: { shape: Record<string, { description?: string }> } } })._def.innerType;
  return inner.shape;
}

describe('word miss keeps its own remedy, takes no stale-id hint', () => {
  it('match miss carries E_NOT_FOUND with the exact-substring remedy and no hint', async () => {
    const { engine, ast } = await freshEngine();
    const para = await writeParagraph(engine, ast, 'Hello Riverbend world');
    const first = await batchErrors(engine, [
      { action: 'style', target: para.id, match: 'zz-absent-needle', bold: true },
    ]);
    expect(String(first.error)).toContain('E_NOT_FOUND');
    expect(String(first.error)).toContain('not found in node');
    expect(String(first.error)).toContain('exact matching substring');
    expect(first.hint).toBeUndefined();
  });

  it('genuine node miss still takes the stale-id hint', async () => {
    const { engine } = await freshEngine();
    const first = await batchErrors(engine, [
      { action: 'replace', target: 'p_no_such_node', text: 'x' },
    ]);
    expect(String(first.error)).toContain('Node p_no_such_node not found');
    expect(first.hint).toBeDefined();
    expect(String(first.hint)).toContain('go stale');
  });

  it('lowercase anchor miss still takes the stale-id hint', async () => {
    const { engine } = await freshEngine();
    const first = await batchErrors(engine, [
      { action: 'insert', type: 'paragraph', anchor: 'p_no_such_node', position: 'after', text: 'x' },
    ]);
    expect(String(first.error)).toContain("Anchor node 'p_no_such_node' not found");
    expect(first.hint).toBeDefined();
    expect(String(first.hint)).toContain('go stale');
  });
});

describe('match trio carries describe text', () => {
  it('flat style fields are annotated', () => {
    const shape = styleShape();
    expect(shape.match?.description).toContain('Word-level span');
    expect(shape.matchAll?.description).toContain('first occurrence');
    expect(shape.matchCase?.description).toContain('Case-sensitive');
  });

  it('nested properties bag shares the annotations', () => {
    const shape = styleShape();
    const nested = nestedShape(shape.properties as { description?: string } & Record<string, unknown>);
    expect(nested.match?.description).toContain('Word-level span');
    expect(nested.matchAll?.description).toContain('first occurrence');
    expect(nested.matchCase?.description).toContain('Case-sensitive');
  });
});
