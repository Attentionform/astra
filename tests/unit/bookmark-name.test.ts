import { describe, expect, it } from 'vitest';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';

async function bookmarkNames(extra: Record<string, unknown>): Promise<string[]> {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const res = await engine.executeBatch({
    operations: [{ action: 'bookmark', target: ast.sections[0].nodes[0].id, ...extra } as never],
  });
  expect(res.success).toBe(true);
  return [...parser.readDocumentXml().matchAll(/<w:bookmarkStart[^>]*w:name="([^"]*)"/g)].map((m) => m[1]);
}

describe('bookmark names', () => {
  it('name creates the bookmark under that name', async () => {
    expect(await bookmarkNames({ name: 'Intro' })).toEqual(['Intro']);
  });
  it('bookmarkName creates the bookmark under that name', async () => {
    expect(await bookmarkNames({ bookmarkName: 'Intro' })).toEqual(['Intro']);
  });
  it('bookmarkName wins when both are given', async () => {
    expect(await bookmarkNames({ name: 'FromName', bookmarkName: 'FromBookmarkName' })).toEqual(['FromBookmarkName']);
  });
  it('without a name the bookmark gets a generated name', async () => {
    const names = await bookmarkNames({});
    expect(names).toHaveLength(1);
    expect(names[0]).toMatch(/^bm_\d+$/);
  });
});
