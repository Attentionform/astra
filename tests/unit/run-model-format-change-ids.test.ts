import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

/**
 * A run that carries a tracked formatting change (w:rPrChange) holds one revision id. An edit that splits the run
 * must not copy that id onto every part: ids must stay unique in the file, the validator rejects duplicates, and Word
 * reads each id as one revision. New text must not claim a formatting change that nobody made to it.
 */
const A = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';
const fmtRun = (id: number, text: string, extra = '') => `<w:r><w:rPr><w:b/><w:rPrChange w:id="${id}" ${A}><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">${text}</w:t>${extra}</w:r>`;
const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  return { parser, ast, engine, ids: ast.sections[0].nodes.map(n => n.id) };
}
const documentXml = (parser: DocxParser) => new AdmZip(parser.rezip({ skipValidation: true })).readAsText('word/document.xml');
const revisionIds = (xml: string) => [...xml.matchAll(/<w:(?:ins|del|rPrChange|pPrChange|moveFrom|moveTo)\s[^>]*?w:id="(\d+)"/g)].map(m => m[1]);
const formatChanges = (xml: string) => [...xml.matchAll(/<w:rPrChange\s[^>]*?w:id="(\d+)"/g)].map(m => m[1]);
const validatorErrors = async (parser: DocxParser) => ((await validateDocxBuffer(parser.rezip({ skipValidation: true }))).errors ?? []).map(e => e.id ?? '');

describe('a run with a tracked formatting change keeps unique revision ids when an edit splits it', () => {
  for (const tracked of [true, false]) {
    it(`${tracked ? 'tracked' : 'untracked'} replace of a word in the middle of the run`, async () => {
      const doc = await open(`<w:p>${fmtRun(7, 'alpha beta gamma')}</w:p>${'<w:p><w:r><w:t>second</w:t></w:r></w:p>'}`);
      const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.ids[0], text: 'alpha BETA gamma', trackChanges: tracked }] } as never);
      expect(result.success, JSON.stringify(result.errors)).toBe(true);
      const xml = documentXml(doc.parser);
      const ids = revisionIds(xml);
      expect(new Set(ids).size).toBe(ids.length);
      expect(await validatorErrors(doc.parser)).toEqual([]);
      // The first part keeps the original id.
      expect(formatChanges(xml)).toContain('7');
    });
  }

  it('every part keeps the formatting change: Reject All restores the old formatting on all of them', async () => {
    const doc = await open(`<w:p>${fmtRun(7, 'alpha beta gamma')}</w:p>`);
    await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.ids[0], text: 'alpha BETA gamma', trackChanges: true }] } as never);
    const xml = documentXml(doc.parser);
    // alpha_, beta (deleted), gamma: the kept parts carry a formatting change each; the new text carries none.
    const inserted = xml.match(/<w:ins\s[^>]*>[\s\S]*?<\/w:ins>/g) ?? [];
    expect(inserted.length).toBe(1);
    expect(inserted[0]).not.toContain('rPrChange');
    expect(inserted[0]).toContain('BETA');
    const saved = new DocxParser(doc.parser.rezip({ skipValidation: true }));
    const ast = await saved.getAST();
    const engine = new OperationEngine(ast, saved);
    const reject = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'rejectAll' }] } as never);
    expect(reject.success, JSON.stringify(reject.errors)).toBe(true);
    const after = documentXml(saved);
    expect(after).not.toContain('<w:b/>');
    expect(after).not.toContain('rPrChange');
    expect(after.replace(/<[^>]+>/g, '')).toContain('alpha beta gamma');
  });

  it('Accept All keeps the new formatting and removes every formatting-change record', async () => {
    const doc = await open(`<w:p>${fmtRun(7, 'alpha beta gamma')}</w:p>`);
    await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.ids[0], text: 'alpha BETA gamma', trackChanges: true }] } as never);
    const saved = new DocxParser(doc.parser.rezip({ skipValidation: true }));
    const ast = await saved.getAST();
    const engine = new OperationEngine(ast, saved);
    const accept = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
    expect(accept.success, JSON.stringify(accept.errors)).toBe(true);
    const after = documentXml(saved);
    expect(after).not.toContain('rPrChange');
    expect(after).toContain('<w:b/>');
    expect(after.replace(/<[^>]+>/g, '')).toContain('alpha BETA gamma');
  });

  it('a run split twice and two runs with their own formatting changes all end with distinct ids', async () => {
    const doc = await open(`<w:p>${fmtRun(7, 'one two three')}${fmtRun(8, ' four five six')}</w:p>`);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.ids[0], text: 'one TWO three four FIVE six', trackChanges: true }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const xml = documentXml(doc.parser);
    const ids = revisionIds(xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect(formatChanges(xml)).toEqual(expect.arrayContaining(['7', '8']));
    expect(await validatorErrors(doc.parser)).toEqual([]);
  });

  it('a run that holds a text and a tab is separated into runs with distinct formatting-change ids', async () => {
    const doc = await open(`<w:p>${fmtRun(7, 'alpha', '<w:tab/><w:t>beta</w:t>')}${plain(' gamma')}</w:p>`);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.ids[0], text: 'ALPHA\tbeta gamma', trackChanges: true }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const xml = documentXml(doc.parser);
    const ids = revisionIds(xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect(await validatorErrors(doc.parser)).toEqual([]);
  });

  it('new revision ids stay above the ids of the formatting changes', async () => {
    const doc = await open(`<w:p>${fmtRun(40, 'alpha beta gamma')}</w:p>`);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.ids[0], text: 'alpha BETA gamma', trackChanges: true }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const ids = revisionIds(documentXml(doc.parser)).map(Number);
    expect(ids.filter(id => id !== 40).every(id => id > 40)).toBe(true);
  });
});
