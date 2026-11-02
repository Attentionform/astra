import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';

const count = (haystack: string, needle: RegExp): number => (haystack.match(needle) ?? []).length;
const SHAPE_TYPE_ID = /<v:shapetype[^>]*id="_x0000_t136"/g;

async function applyOps(input: Buffer | null, operations: unknown[]): Promise<Buffer> {
  const parser = input ? new DocxParser(input) : DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const result = await new OperationEngine(ast, parser).executeBatch({ operations: operations as never });
  expect(result.errors ?? []).toEqual([]);
  expect(result.success).toBe(true);
  return parser.rezip();
}

const headerParts = (docx: Buffer): string[] =>
  new AdmZip(docx)
    .getEntries()
    .map((e) => e.entryName)
    .filter((n) => /^word\/header\d+\.xml$/.test(n));

describe('watermark declares the WordArt shape type once per header part', () => {
  it('a header that already has a Word-made watermark (same shape type, other shape id) keeps one shape type', async () => {
    const first = await applyOps(null, [{ action: 'watermark', text: 'DRAFT' }]);
    // Word names its watermark shapes PowerPlusWaterMarkObject plus a number, so the operation does not treat it as its own.
    const zip = new AdmZip(first);
    const part = headerParts(first)[0];
    zip.updateFile(part, Buffer.from(zip.readAsText(part).replace('id="PowerPlusWaterMarkObject"', 'id="PowerPlusWaterMarkObject357476642"')));
    const wordMade = zip.toBuffer();
    expect(count(new AdmZip(wordMade).readAsText(part), SHAPE_TYPE_ID)).toBe(1);
    expect((await validateDocxBuffer(wordMade)).ok).toBe(true);

    const out = await applyOps(wordMade, [{ action: 'watermark', text: 'CONFIDENTIAL' }]);
    const header = new AdmZip(out).readAsText(headerParts(out)[0]);
    expect(count(header, SHAPE_TYPE_ID)).toBe(1);
    expect(count(header, /<v:shape\s/g)).toBe(2);
    expect(header).toContain('DRAFT');
    expect(header).toContain('CONFIDENTIAL');
    const res = await validateDocxBuffer(out);
    expect(res.errors).toEqual([]);
    expect(res.ok).toBe(true);
  });

  it('a watermark on a document without one declares the shape type once', async () => {
    const out = await applyOps(null, [{ action: 'watermark', text: 'DRAFT' }]);
    const header = new AdmZip(out).readAsText(headerParts(out)[0]);
    expect(count(header, SHAPE_TYPE_ID)).toBe(1);
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('running the operation twice still leaves one watermark and one shape type', async () => {
    const once = await applyOps(null, [{ action: 'watermark', text: 'DRAFT' }]);
    const twice = await applyOps(once, [{ action: 'watermark', text: 'FINAL' }]);
    const header = new AdmZip(twice).readAsText(headerParts(twice)[0]);
    expect(count(header, SHAPE_TYPE_ID)).toBe(1);
    expect(count(header, /<v:shape\s/g)).toBe(1);
    expect(header).toContain('FINAL');
    expect(header).not.toContain('DRAFT');
    expect((await validateDocxBuffer(twice)).ok).toBe(true);
  });

  it('a header with an absolute relationship target keeps its text and gets no word/word/ part', async () => {
    const base = await applyOps(null, [{ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'KEEP THIS HEADER TEXT' }]);
    const zip = new AdmZip(base);
    const rels = zip.readAsText('word/_rels/document.xml.rels');
    expect(rels).toContain('Target="header1.xml"');
    zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels.replace('Target="header1.xml"', 'Target="/word/header1.xml"')));
    const absolute = zip.toBuffer();

    const out = await applyOps(absolute, [{ action: 'watermark', text: 'DRAFT' }]);
    const names = new AdmZip(out).getEntries().map((e) => e.entryName);
    expect(names.filter((n) => n.startsWith('word/word/'))).toEqual([]);
    expect(headerParts(out)).toEqual(['word/header1.xml']);
    const header = new AdmZip(out).readAsText('word/header1.xml');
    expect(header).toContain('KEEP THIS HEADER TEXT');
    expect(header).toContain('DRAFT');
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });
});
