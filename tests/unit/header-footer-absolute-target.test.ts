import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';

/**
 * Word and other writers can store the target of a header or footer relationship as an absolute part name
 * (`/word/footer1.xml`) instead of a name relative to word/. The edit has to land in that same part.
 */
async function docWithFooter(target: string): Promise<Buffer> {
  const blank = DocxParser.createNewDocument('blank');
  const ast = await blank.getAST();
  const first = await new OperationEngine(ast, blank).executeBatch({
    operations: [{ action: 'headerFooter', operation: 'set', targetType: 'footer', text: 'OLD FOOTER' }],
  });
  expect(first.success).toBe(true);
  const zip = new AdmZip(blank.rezip());
  const relsPath = 'word/_rels/document.xml.rels';
  const rels = zip.readAsText(relsPath);
  expect(rels).toContain('Target="footer1.xml"');
  zip.updateFile(relsPath, Buffer.from(rels.replace('Target="footer1.xml"', `Target="${target}"`)));
  return zip.toBuffer();
}

async function setFooterText(input: Buffer, text: string): Promise<Buffer> {
  const parser = new DocxParser(input);
  const ast = await parser.getAST();
  const result = await new OperationEngine(ast, parser).executeBatch({
    operations: [{ action: 'headerFooter', operation: 'set', targetType: 'footer', text }],
  });
  expect(result.success).toBe(true);
  return parser.rezip();
}

describe('headerFooter set edits the existing part when the relationship target is absolute', () => {
  it('writes the new text into word/footer1.xml and creates no word/word/ part', async () => {
    const input = await docWithFooter('/word/footer1.xml');
    expect(new AdmZip(input).readAsText('word/footer1.xml')).toContain('OLD FOOTER');
    const out = await setFooterText(input, 'NEW FOOTER');
    const names = new AdmZip(out).getEntries().map((e) => e.entryName);
    expect(names.filter((n) => n.startsWith('word/word/') || n.includes('//'))).toEqual([]);
    expect(names.filter((n) => /^word\/footer\d+\.xml$/.test(n))).toEqual(['word/footer1.xml']);
    const footer = new AdmZip(out).readAsText('word/footer1.xml');
    expect(footer).toContain('NEW FOOTER');
    expect(footer).not.toContain('OLD FOOTER');
    const res = await validateDocxBuffer(out);
    expect(res.errors).toEqual([]);
    expect(res.ok).toBe(true);
  });

  it('still works when the target is relative, with or without the word/ prefix', async () => {
    for (const target of ['footer1.xml', 'word/footer1.xml']) {
      const input = await docWithFooter(target);
      const out = await setFooterText(input, 'NEW FOOTER');
      const names = new AdmZip(out).getEntries().map((e) => e.entryName);
      expect(names.filter((n) => n.startsWith('word/word/'))).toEqual([]);
      expect(new AdmZip(out).readAsText('word/footer1.xml')).toContain('NEW FOOTER');
      expect((await validateDocxBuffer(out)).ok).toBe(true);
    }
  });
});
