import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser, checkoutCommit, getHistory } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';

const tmpDirs: string[] = [];

/** A blank document with two body paragraphs that share one w14:paraId. */
function duplicateIdDocx(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-mcp-dup-paraid-'));
  tmpDirs.push(dir);
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  let xml = zip.getEntry('word/document.xml')?.getData().toString('utf8') as string;
  if (!xml.includes('xmlns:w14=')) xml = xml.replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ');
  const anchor = xml.indexOf('<w:sectPr') >= 0 ? xml.indexOf('<w:sectPr') : xml.lastIndexOf('</w:body>');
  xml =
    xml.slice(0, anchor) +
    '<w:p w14:paraId="5C5C5C5C"><w:r><w:t>first copy</w:t></w:r></w:p><w:p w14:paraId="5C5C5C5C"><w:r><w:t>second copy</w:t></w:r></w:p>' +
    xml.slice(anchor);
  zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
  const file = path.join(dir, 'in.docx');
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

afterEach(() => {
  for (const d of tmpDirs.splice(0)) fs.rmSync(d, { recursive: true, force: true });
});

const docXml = (file: string) => new AdmZip(fs.readFileSync(file)).getEntry('word/document.xml')?.getData().toString('utf8') as string;
const countId = (xml: string, id: string) => (xml.match(new RegExp(`w14:paraId="${id}"`, 'g')) ?? []).length;

describe('save_document renumbers repeated paragraph ids', () => {
  it('saves, reports each renumbered paragraph and keeps the state before in the history', async () => {
    const file = duplicateIdDocx();
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId });
    if (!saved.ok) throw new Error(JSON.stringify(saved));
    expect(JSON.stringify(saved)).toContain('Renumbered duplicate w14:paraId');

    const xml = docXml(file);
    expect(countId(xml, '5C5C5C5C')).toBe(1);
    expect(xml).toContain('first copy');
    expect(xml).toContain('second copy');
    expect(new DocxParser(fs.readFileSync(file)).getDuplicateParaIds()).toEqual([]);

    const entries = await getHistory(file);
    expect(entries.length).toBeGreaterThanOrEqual(2);
    const head = await checkoutCommit(file, entries[0].oid);
    const before = await checkoutCommit(file, entries[1].oid);
    expect(countId(head.xml, '5C5C5C5C')).toBe(1);
    expect(countId(before.xml, '5C5C5C5C')).toBe(2);
  });

  it('a second save of the same session renumbers nothing more', async () => {
    const file = duplicateIdDocx();
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    expect((await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId })).ok).toBe(true);
    const again = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId });
    expect(again.ok).toBe(true);
    expect(JSON.stringify(again)).not.toContain('Renumbered duplicate');
  });

  it('strict keeps the refusal and leaves the file as it was', async () => {
    const file = duplicateIdDocx();
    const before = fs.readFileSync(file);
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId, strict: true } as never);
    expect(saved.ok).toBe(false);
    expect(JSON.stringify(saved)).toMatch(/duplicate|Duplicate/);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
});
