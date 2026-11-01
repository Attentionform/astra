import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNormalizeDocument } from '../../src/lib/mcp/handlers/parity';
import { handleRevertDocument } from '../../src/lib/mcp/handlers/history';

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

const docXml = (file: string) => new AdmZip(fs.readFileSync(file)).getEntry('word/document.xml')?.getData().toString('utf8') as string;

describe('normalize_document is reversible with one revert_document', () => {
  it('records the state before the repairs when it differs from the head, then the result on save', async () => {
    const file = smartTagDocx();
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const sessionId = opened.sessionId;
    expect((await LEAN6_HANDLERS.save_document(store, { sessionId })).ok).toBe(true); // creates the history
    const edit = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'added later' }],
    } as never);
    expect(edit.ok).toBe(true);

    const norm = await handleNormalizeDocument(store, { sessionId });
    if (!norm.ok) throw new Error(JSON.stringify(norm));
    expect(norm.normalized.some((n) => n.includes('smartTag'))).toBe(true);
    expect(typeof norm.preStateCommit).toBe('string'); // the unsaved edit made the state differ from the head
    expect((await LEAN6_HANDLERS.save_document(store, { sessionId })).ok).toBe(true);
    expect(docXml(file)).not.toContain('smartTag');

    const reverted = await handleRevertDocument(store, { sessionId, steps: 1 } as never);
    expect(JSON.stringify(reverted).slice(0, 500)).toContain('"ok":true');
    const back = path.join(path.dirname(file), 'back.docx');
    expect((await LEAN6_HANDLERS.save_document(store, { sessionId, outPath: back })).ok).toBe(true);
    const xml = docXml(back);
    expect(xml).toContain('smartTag');
    expect(xml).toContain('added later');
  });

  it('records nothing before the repairs when the state equals the head', async () => {
    const file = smartTagDocx();
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    expect((await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId })).ok).toBe(true);
    const norm = await handleNormalizeDocument(store, { sessionId: opened.sessionId });
    if (!norm.ok) throw new Error(JSON.stringify(norm));
    expect(norm.preStateCommit).toBeUndefined();
  });

  it('records nothing for a document without a history', async () => {
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: smartTagDocx() });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const norm = await handleNormalizeDocument(store, { sessionId: opened.sessionId });
    if (!norm.ok) throw new Error(JSON.stringify(norm));
    expect(norm.preStateCommit).toBeUndefined();
  });
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
