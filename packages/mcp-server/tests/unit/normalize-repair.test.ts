import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNormalizeDocument } from '../../src/lib/mcp/handlers/parity';

const tmpDirs: string[] = [];

function smartTagDocx(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-mcp-normalize-repair-'));
  tmpDirs.push(dir);
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const xml = zip.getEntry('word/document.xml')?.getData().toString('utf8') as string;
  const anchor = xml.indexOf('<w:sectPr') >= 0 ? xml.indexOf('<w:sectPr') : xml.lastIndexOf('</w:body>');
  const withTag =
    xml.slice(0, anchor) +
    '<w:p><w:smartTag w:uri="urn:x" w:element="place"><w:r><w:t>kept text</w:t></w:r></w:smartTag></w:p>' +
    xml.slice(anchor);
  zip.updateFile('word/document.xml', Buffer.from(withTag, 'utf8'));
  const file = path.join(dir, 'in.docx');
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

afterEach(() => {
  for (const d of tmpDirs.splice(0)) fs.rmSync(d, { recursive: true, force: true });
});

describe('normalize_document repairs known validator error kinds', () => {
  it('unwraps a smart tag, keeps the text in the saved file, reports remaining and rolledBack', async () => {
    const file = smartTagDocx();
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const norm = await handleNormalizeDocument(store, { sessionId: opened.sessionId });
    if (!norm.ok) throw new Error(JSON.stringify(norm));
    expect(norm.normalized.some((n) => n.includes('smartTag'))).toBe(true);
    expect(typeof norm.remaining).toBe('number');
    expect(norm.rolledBack).toEqual([]);

    const out = path.join(path.dirname(file), 'out.docx');
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId, outPath: out });
    expect(saved.ok).toBe(true);
    const xml = new AdmZip(fs.readFileSync(out)).getEntry('word/document.xml')?.getData().toString('utf8') as string;
    expect(xml).not.toContain('smartTag');
    expect(xml).toContain('kept text');
  });

  it('a second normalize_document on the same session repairs nothing more', async () => {
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: smartTagDocx() });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    await handleNormalizeDocument(store, { sessionId: opened.sessionId });
    const again = await handleNormalizeDocument(store, { sessionId: opened.sessionId });
    if (!again.ok) throw new Error(JSON.stringify(again));
    expect(again.normalized.filter((n) => n.includes('smartTag'))).toEqual([]);
  });
});
