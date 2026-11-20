import { describe, expect, it } from 'vitest';
import { DocxParser, OperationEngine, parseBatch } from '../src';

const replace = { action: 'replace', target: 'p_1', text: 'x' };

describe('parseBatch: valid input', () => {
  it('accepts an envelope with every option', () => {
    const res = parseBatch({
      operations: [replace],
      sequential: true,
      rollbackOnError: true,
      maxSteps: 5,
      maxRatio: 100,
      maxUncompressedBytes: 1_000_000,
      allowedDir: '/tmp/images',
    });
    expect(res.ok).toBe(true);
    if (res.ok) {
      expect(res.batch.operations).toHaveLength(1);
      expect(res.batch.allowedDir).toBe('/tmp/images');
      expect(res.batch.sequential).toBe(true);
    }
  });
  it('accepts a bare array of operations', () => {
    const res = parseBatch([replace]);
    expect(res.ok).toBe(true);
  });
  it('returns a batch that executeBatch accepts and applies', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const engine = new OperationEngine(await parser.getAST(), parser);
    const first = parser.getStructuralSummary().nodes[0].id;
    const res = parseBatch({ operations: [{ action: 'replace', target: first, text: 'Parsed and applied' }] });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const result = await engine.executeBatch(res.batch);
    expect(result.success).toBe(true);
  });
  it('accepts the heading recipe: insert with an id, then namedStyle apply', () => {
    const res = parseBatch([
      { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Chapter one', id: 'h_one' },
      { action: 'namedStyle', operation: 'apply', styleName: 'Heading 1', target: 'h_one' },
    ]);
    expect(res.ok).toBe(true);
  });
});

describe('parseBatch: invalid input lists every problem with the operation index', () => {
  it('rejects non-objects without throwing', () => {
    for (const bad of [null, undefined, 42, 'text']) {
      const res = parseBatch(bad);
      expect(res.ok).toBe(false);
    }
  });
  it('rejects an empty operations array', () => {
    const res = parseBatch({ operations: [] });
    expect(res.ok).toBe(false);
  });
  it('rejects an unknown action and names its index', () => {
    const res = parseBatch([replace, { action: 'frobnicate', target: 'p_2' }]);
    expect(res.ok).toBe(false);
    if (!res.ok) expect(res.errors.some((e) => e.operationIndex === 1)).toBe(true);
  });
  it('rejects a missing required field and names its index', () => {
    const res = parseBatch([{ action: 'replace', target: 'p_1' }]);
    expect(res.ok).toBe(false);
    if (!res.ok) expect(res.errors[0]).toMatchObject({ operationIndex: 0 });
  });
  it("rejects an insert 'style' field, which the engine would ignore", () => {
    const res = parseBatch([{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'T', style: 'Heading1' }]);
    expect(res.ok).toBe(false);
  });
  it("rejects insert type 'heading'", () => {
    const res = parseBatch([{ action: 'insert', type: 'heading', anchor: 'end', position: 'after', text: 'T' }]);
    expect(res.ok).toBe(false);
  });
  it('rejects a body insert without anchor and position (a rule across fields)', () => {
    const res = parseBatch([{ action: 'insert', type: 'paragraph', text: 'no anchor' }]);
    expect(res.ok).toBe(false);
    if (!res.ok) expect(res.errors[0].operationIndex).toBe(0);
  });
  it('rejects an unknown envelope key', () => {
    const res = parseBatch({ operations: [replace], retries: 3 });
    expect(res.ok).toBe(false);
  });
  it('collects problems from several operations in one pass', () => {
    const res = parseBatch([{ action: 'replace' }, { action: 'nope' }, { action: 'delete' }]);
    expect(res.ok).toBe(false);
    if (!res.ok) {
      const indexes = new Set(res.errors.map((e) => e.operationIndex));
      expect(indexes.has(0) && indexes.has(1) && indexes.has(2)).toBe(true);
    }
  });
});

describe('F13: bare table and chart pass the schema; the engine checks the sub-operation', () => {
  it('documents the limit: parseBatch accepts them and executeBatch refuses at run time', async () => {
    expect(parseBatch([{ action: 'table' }]).ok).toBe(true);
    const parser = DocxParser.createNewDocument('blank');
    const engine = new OperationEngine(await parser.getAST(), parser);
    const result = await engine.executeBatch({ operations: [{ action: 'table' }], rollbackOnError: true }).catch((e: Error) => ({ success: false, errors: [{ error: e.message }] }));
    expect(result.success).toBe(false);
  });
});
