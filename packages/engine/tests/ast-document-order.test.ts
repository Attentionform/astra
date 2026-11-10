import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, composeDocument } from '../src';

describe('getAST lists blocks in document order', () => {
  it('keeps a block that an operation put in the middle between its neighbours', async () => {
    const parser = DocxParser.createNewDocument('blank');
    await composeDocument(parser, '# Head\n\nAlpha.\n\nOmega.\n');
    const before = await parser.getAST();
    const engine = new OperationEngine(before, parser);
    const res = await engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'md_0002', position: 'after', text: 'Between', id: 'zz_mid' }] as never,
      sequential: true,
    });
    expect(res.success).toBe(true);
    const after = await parser.getAST();
    const order = after.sections.flatMap((s) => s.nodes.map((n) => String(n.metadata?.text ?? '')));
    expect(order).toEqual(['Head', 'Alpha.', 'Between', 'Omega.']);
  });
});
