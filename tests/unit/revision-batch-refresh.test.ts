import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';

const STAMP = 'w:author="Alice" w:date="2026-01-01T00:00:00Z"';
const ins = (id: number, text: string) => `<w:p><w:ins w:id="${id}" ${STAMP}><w:r><w:t>${text}</w:t></w:r></w:ins></w:p>`;
const plain = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const nodes = (ast: { sections: { nodes: { id: string; version: number; metadata: { text?: unknown } }[] }[] }) => ast.sections.flatMap(s => s.nodes);

describe('FX-14: a run of accept/reject operations refreshes the model once, with the same result', () => {
  it('each operation sees the state left by the previous one', async () => {
    const { engine, ast } = await open(`${ins(1, 'one')}${ins(2, 'two')}${ins(3, 'three')}${plain('tail')}`);
    const result = await engine.executeBatch({ operations: [
      { action: 'trackedChange', operation: 'reject', ids: [1] },
      { action: 'trackedChange', operation: 'accept', ids: [2] },
      { action: 'trackedChange', operation: 'reject', ids: [3] },
    ] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionsPending).toBe(0);
    expect(result.revisionResolution?.resolved).toBe(3);
    const texts = nodes(ast).map(n => String(n.metadata.text));
    expect(texts).toContain('two');
    expect(texts).toContain('tail');
    expect(texts).not.toContain('one');
    expect(texts).not.toContain('three');
  });

  it('a later accept of an id an earlier reject already resolved fails the same way as before', async () => {
    const { engine } = await open(`${ins(1, 'one')}${plain('tail')}`);
    const result = await engine.executeBatch({ operations: [
      { action: 'trackedChange', operation: 'reject', ids: [1] },
      { action: 'trackedChange', operation: 'accept', ids: [1] },
    ] } as never);
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/Revision id 1 was not found/);
  });

  it('an edit after a resolution sees the refreshed text, and a resolution after an edit sees the edit', async () => {
    const { engine, ast, parser } = await open(`${ins(1, 'drop me')}${plain('alpha beta')}${ins(2, 'keep me')}`);
    const target = nodes(ast)[1].id;
    const result = await engine.executeBatch({ operations: [
      { action: 'trackedChange', operation: 'reject', ids: [1] },
      { action: 'replace', target, text: 'alpha BETA' },
      { action: 'trackedChange', operation: 'accept', ids: [2] },
    ] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const texts = nodes(ast).map(n => String(n.metadata.text));
    // a rejected insertion leaves its (now empty) paragraph, as Word does
    expect(texts).toEqual(['', 'alpha BETA', 'keep me']);
    expect(parser.getDocumentDom().getElementsByTagName('w:ins').length).toBe(0);
    expect(result.revisionsPending).toBe(0);
  });

  it('versions: blocks the run changed move up, blocks it left alone keep their version', async () => {
    const { engine, ast } = await open(`${ins(1, 'one')}${plain('untouched')}${ins(2, 'two')}`);
    const before = new Map(nodes(ast).map(n => [n.id, n.version]));
    const [first, middle, last] = nodes(ast).map(n => n.id);
    await engine.executeBatch({ operations: [
      { action: 'trackedChange', operation: 'reject', ids: [1] },
      { action: 'trackedChange', operation: 'accept', ids: [2] },
    ] } as never);
    const after = new Map(nodes(ast).map(n => [n.id, n.version]));
    expect(after.get(middle)).toBe(before.get(middle));
    // the paragraph holding each resolved insertion changed, so both versions move up
    expect(after.get(first)).toBeGreaterThan(before.get(first)!);
    expect(after.get(last)).toBeGreaterThan(before.get(last)!);
  });

  it('a lone acceptAll through executeOperation leaves the model current', async () => {
    const { engine, ast } = await open(`${ins(1, 'one')}${plain('tail')}`);
    await engine.execute({ action: 'trackedChange', operation: 'rejectAll' } as never);
    expect(nodes(ast).map(n => String(n.metadata.text))).toEqual(['', 'tail']);
  });

  it('a failing operation after a resolution rolls back to the state before the batch', async () => {
    const { engine, ast, parser } = await open(`${ins(1, 'one')}${plain('tail')}`);
    const before = nodes(ast).map(n => `${n.id}:${n.version}:${String(n.metadata.text)}`);
    const result = await engine.executeBatch({ sequential: true, operations: [
      { action: 'trackedChange', operation: 'accept', ids: [1] },
      { action: 'replace', target: 'no-such-node', text: 'x' },
    ] } as never);
    expect(result.success).toBe(false);
    expect(result.rollbackTriggered).toBe(true);
    expect(nodes(ast).map(n => `${n.id}:${n.version}:${String(n.metadata.text)}`)).toEqual(before);
    expect(parser.getDocumentDom().getElementsByTagName('w:ins').length).toBe(1);
  });

  it('a dry run shows the resolution and leaves the document and model unchanged', async () => {
    const { engine, ast, parser } = await open(`${ins(1, 'one')}${plain('tail')}`);
    const before = nodes(ast).map(n => `${n.id}:${n.version}:${String(n.metadata.text)}`);
    const result = await engine.executeBatch({ dryRun: true, operations: [{ action: 'trackedChange', operation: 'accept', ids: [1] }] } as never);
    expect(result.success).toBe(true);
    expect(result.revisionResolution?.resolved).toBe(1);
    expect(nodes(ast).map(n => `${n.id}:${n.version}:${String(n.metadata.text)}`)).toEqual(before);
    expect(parser.getDocumentDom().getElementsByTagName('w:ins').length).toBe(1);
  });

  it('a failed batch does not leave a refresh owed for the next batch', async () => {
    const { engine, ast } = await open(`${ins(1, 'one')}${plain('tail')}`);
    await engine.executeBatch({ sequential: true, operations: [
      { action: 'trackedChange', operation: 'accept', ids: [1] },
      { action: 'replace', target: 'no-such-node', text: 'x' },
    ] } as never);
    const next = await engine.executeBatch({ operations: [{ action: 'replace', target: nodes(ast)[1].id, text: 'tail 2' }] } as never);
    expect(next.success, JSON.stringify(next.errors)).toBe(true);
    expect(nodes(ast).map(n => String(n.metadata.text))[1]).toBe('tail 2');
  });
});
