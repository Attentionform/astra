import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, computeDocStats } from '../src';

async function statsAfter(ops: unknown[]) {
  const parser = DocxParser.createNewDocument('blank');
  const engine = new OperationEngine(await parser.getAST(), parser);
  const res = await engine.executeBatch({ operations: ops as never, sequential: true });
  expect(res.success).toBe(true);
  return computeDocStats(await parser.getAST());
}

describe('callouts in the document census', () => {
  it('counts a callout in callouts and not in tables', async () => {
    const stats = await statsAfter([
      { action: 'insert', type: 'table', anchor: 'end', position: 'after', headers: ['A', 'B', 'C'], rows: [['1', '2', '3'], ['4', '5', '6']] },
      { action: 'callout', anchor: 'end', position: 'after', title: 'Heads up', text: 'Body of the note', theme: 'amber' },
    ]);
    expect(stats.tables).toBe(1);
    expect(stats.callouts).toBe(1);
  });

  it('keeps the page estimate the same as when a callout counted as a table', async () => {
    const withCallout = await statsAfter([{ action: 'callout', anchor: 'end', position: 'after', text: 'note', theme: 'indigo' }]);
    expect(withCallout.tables).toBe(0);
    expect(withCallout.callouts).toBe(1);
    expect(withCallout.pages).toBeGreaterThanOrEqual(1);
  });

  it('does not mistake a one-cell table without a fill, or a multi-cell shaded table, for a callout', async () => {
    const stats = await statsAfter([
      { action: 'insert', type: 'table', anchor: 'end', position: 'after', headers: ['Only'], rows: [] },
      { action: 'insert', type: 'table', anchor: 'end', position: 'after', headers: ['A', 'B'], rows: [['1', '2']] },
    ]);
    expect(stats.callouts).toBe(0);
    expect(stats.tables).toBe(2);
  });
});
