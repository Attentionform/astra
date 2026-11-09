import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';

/**
 * An accept or reject with nothing to resolve leaves the document as it is, so the model is not re-read from the
 * file. Re-reading a document of 3,000 paragraphs cost about 210 of the 350 ms of such a call. The model is still
 * refreshed when something was resolved.
 */
const A = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';
const r = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const blocks = (ast: { sections: { nodes: unknown[] }[] }) => ast.sections.flatMap(section => section.nodes) as { id: string; version: number }[];

describe('accept and reject with nothing pending', () => {
  for (const operation of ['acceptAll', 'rejectAll'] as const) {
    it(`${operation} changes no block, id or version and reports 0 resolved`, async () => {
      const doc = await open(`<w:p>${r('one')}</w:p><w:p>${r('two')}</w:p>`);
      const before = blocks(doc.ast).map(node => ({ node, id: node.id, version: node.version }));
      const result = await doc.engine.executeBatch({ operations: [{ action: 'trackedChange', operation }] } as never);
      expect(result.success, JSON.stringify(result.errors)).toBe(true);
      expect(result.revisionResolution?.resolved).toBe(0);
      const after = blocks(doc.ast);
      expect(after.length).toBe(before.length);
      after.forEach((node, i) => {
        expect(node).toBe(before[i].node); // the same block object: the model was not rebuilt
        expect(node.id).toBe(before[i].id);
        expect(node.version).toBe(before[i].version);
      });
    });
  }

  it('the model is still refreshed when something was resolved', async () => {
    const doc = await open(`<w:p>${r('keep ')}<w:ins w:id="1" ${A}>${r('added')}</w:ins></w:p><w:p>${r('two')}</w:p>`);
    const first = blocks(doc.ast)[0];
    const version = first.version;
    const result = await doc.engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionResolution?.resolved).toBe(1);
    const after = blocks(doc.ast)[0];
    expect(after.version).toBe(version + 1);
    expect((after as unknown as { metadata: { hasTrackedChanges?: boolean } }).metadata.hasTrackedChanges).not.toBe(true);
  });

  it('a deleted paragraph is gone from the model after Accept All', async () => {
    const doc = await open(`<w:p><w:pPr><w:rPr><w:del w:id="2" ${A}/></w:rPr></w:pPr><w:del w:id="1" ${A}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p><w:p>${r('stay')}</w:p>`);
    expect(blocks(doc.ast).length).toBe(2);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(blocks(doc.ast).length).toBe(1);
  });

  it('a selective accept that matches nothing is refused and the ids and versions stay as they were', async () => {
    const doc = await open(`<w:p>${r('one')}</w:p>`);
    const before = blocks(doc.ast).map(node => ({ id: node.id, version: node.version }));
    const result = await doc.engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'accept', ids: [999] }] } as never);
    expect(result.success).toBe(false);
    expect(blocks(doc.ast).map(node => ({ id: node.id, version: node.version }))).toEqual(before);
  });
});
