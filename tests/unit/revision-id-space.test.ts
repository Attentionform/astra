import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';

const STAMP = 'w:author="A" w:date="2026-01-01T00:00:00Z"';
const para = (text: string, extra = '') => `<w:p>${extra}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  return { parser, ast, engine, first: ast.sections[0].nodes[0].id };
}

const batch = (engine: OperationEngine, operations: unknown[]) => engine.executeBatch({ operations } as never);
const moveMarkers = (a: string, b: string) =>
  `<w:p><w:moveFromRangeStart w:id="${a}" w:name="m1" ${STAMP}/><w:moveFromRangeEnd w:id="${a}"/><w:moveToRangeStart w:id="${b}" w:name="m1" ${STAMP}/><w:moveToRangeEnd w:id="${b}"/></w:p>`;
const docXml = (parser: DocxParser) => new XMLSerializer().serializeToString(parser.getDocumentDom());
const ids = (xml: string, tag: string) => [...xml.matchAll(new RegExp(`<w:${tag}[ >][^>]*w:id="(\\d+)"`, 'g'))].map(m => m[1]);

describe('revision ids never collide with move range markers (FX-09)', () => {
  it('new ids start above the highest marker id', async () => {
    const { engine, first } = await open(para('alpha beta gamma') + moveMarkers('1', '2'));
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionIds!.length).toBeGreaterThan(0);
    for (const id of result.revisionIds!) expect(id).toBeGreaterThanOrEqual(3);
  });

  it('a marker with a higher id than any revision still reserves its id', async () => {
    const { engine, first } = await open(para('alpha beta gamma') + moveMarkers('40', '41'));
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.revisionIds).toEqual([42, 43]);
  });

  it('counting and resolving ignore the markers: they are not revisions on their own', async () => {
    const { parser } = await open(para('alpha beta gamma') + moveMarkers('1', '2'));
    expect(parser.countPackageRevisions()).toBe(0);
  });
});

describe('the id allocator stays exact (FX-09, decision D-6)', () => {
  const MAX = '9007199254740991';

  it('an id at 2^53 - 1 makes tracked edits fail safely and change nothing', async () => {
    const { parser, ast, engine, first } = await open(para('alpha beta gamma') + moveMarkers('1', MAX));
    const before = docXml(parser);
    const counter = ast.nextChangeId;
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/E_VALIDATION.*revision id/i);
    expect(docXml(parser)).toBe(before);
    expect(ast.nextChangeId).toBe(counter);
  });

  it('untracked edits still work when the id space is exhausted', async () => {
    const { engine, first, parser } = await open(para('alpha beta gamma') + moveMarkers('1', MAX));
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: false }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(docXml(parser)).toContain('BETA');
  });

  it('an id above 2^53 - 1 is not rounded: tracked edits fail safely', async () => {
    const { parser, engine, first } = await open(para('alpha beta gamma') + `<w:p><w:ins w:id="9007199254740993" ${STAMP}><w:r><w:t>x</w:t></w:r></w:ins></w:p>`);
    const before = docXml(parser);
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.success).toBe(false);
    expect(docXml(parser)).toBe(before);
  });

  it('a non-numeric id is ignored', async () => {
    const { engine, first } = await open(para('alpha beta gamma') + `<w:p><w:moveFromRangeStart w:id="abc" w:name="m" ${STAMP}/></w:p>`);
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionIds).toEqual([1, 2]);
  });

  it('a two-id operation checks every id before it changes anything', async () => {
    // highest id 2^53 - 2: the first new id (2^53 - 1) fits, the second does not.
    const { parser, ast, engine, first } = await open(para('alpha beta gamma') + moveMarkers('1', '9007199254740990'));
    const before = docXml(parser);
    const counter = ast.nextChangeId;
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.success).toBe(false);
    expect(docXml(parser)).toBe(before);
    expect(ast.nextChangeId).toBe(counter);
  });

  it('the largest ids that fit are handed out exactly', async () => {
    const { engine, first } = await open(para('alpha beta gamma') + moveMarkers('1', '9007199254740989'));
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionIds!.map(String)).toEqual(['9007199254740990', '9007199254740991']);
  });

  it('an existing id above the signed 32-bit range is not capped (limit not confirmed)', async () => {
    const { engine, first } = await open(para('alpha beta gamma') + `<w:p><w:ins w:id="2147483647" ${STAMP}><w:r><w:t>x</w:t></w:r></w:ins></w:p>`);
    const result = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(result.revisionIds).toEqual([2147483648, 2147483649]);
  });

  it('within one session an id is not handed out twice, even after acceptAll', async () => {
    const { engine, first } = await open(para('alpha beta gamma'));
    const one = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    expect(one.revisionIds).toEqual([1, 2]);
    await batch(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    const two = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA delta', trackChanges: true }]);
    expect(two.revisionIds).toEqual([3, 4]);
  });
});

describe('every tracked operation reports the ids it created (FX-10)', () => {
  it('insert paragraph with trackChanges returns the id of its w:ins', async () => {
    const { parser, engine, first } = await open(para('alpha') + para('omega'));
    const result = await batch(engine, [{ action: 'insert', type: 'paragraph', anchor: first, position: 'after', text: 'brand new', trackChanges: true }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionIds).toBeDefined();
    expect(result.revisionIds!.map(String).sort()).toEqual(ids(docXml(parser), 'ins').sort());
  });

  it('an untracked insert reports no revision ids', async () => {
    const { engine, first } = await open(para('alpha') + para('omega'));
    const result = await batch(engine, [{ action: 'insert', type: 'paragraph', anchor: first, position: 'after', text: 'plain', trackChanges: false }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionIds ?? []).toEqual([]);
  });

  it('delete and replace report exactly the ids present in the document', async () => {
    const { parser, engine, first } = await open(para('alpha beta gamma') + para('omega'));
    const replaced = await batch(engine, [{ action: 'replace', target: first, text: 'alpha BETA gamma', trackChanges: true }]);
    const deleted = await batch(engine, [{ action: 'delete', target: first, trackChanges: true }]);
    expect(replaced.success, JSON.stringify(replaced.errors)).toBe(true);
    expect(deleted.success, JSON.stringify(deleted.errors)).toBe(true);
    const xml = docXml(parser);
    const present = [...ids(xml, 'ins'), ...ids(xml, 'del')].sort();
    const reported = [...replaced.revisionIds!, ...deleted.revisionIds!].map(String).sort();
    expect(reported.filter(id => present.includes(id)).sort()).toEqual(reported);
  });
});
