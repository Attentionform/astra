import { describe, it, expect } from 'vitest';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { readParserNormalizations } from '../../src/lib/mcp/normalize-read';
import { handleNormalizeDocument } from '../../src/lib/mcp/handlers/parity';
import { useReproducibleOutput } from '../../../../tests/helpers/reproducible';

useReproducibleOutput();

/** Blank package with non-canonical markup (a shading node missing its default + out-of-order props). */
function nonCanonicalBuffer(): Buffer {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const zip = new AdmZip(buf);
  const xml = zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
  expect(xml).toContain('<w:pPr>');
  const evil = xml.replace('<w:pPr>', '<w:pPr><w:shd w:fill="FF0000"/>');
  const crafted = new AdmZip(buf);
  crafted.updateFile('word/document.xml', Buffer.from(evil, 'utf8'));
  return crafted.toBuffer();
}

async function openBufferSession(buf: Buffer): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(buf, 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

describe('normalize_document handler', () => {
  it('heals non-canonical markup and reports each rewrite', async () => {
    const { store, sessionId } = await openBufferSession(nonCanonicalBuffer());
    const res = await handleNormalizeDocument(store, { sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`normalize failed: ${JSON.stringify(res)}`);
    expect(res.normalized.length).toBeGreaterThan(0);
    expect(res.normalized.some((line) => line.includes('w:shd'))).toBe(true);
    expect(res.validation.valid).toBe(true);
    expect(res.revision).toBe(2);
    expect(res.dirty).toBe(true);
    const xml = new AdmZip(store.get(sessionId).parser.rezip()).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
    expect(xml).toContain('w:val="clear"');
  });

  it('output bytes match a direct engine rezip of the same input', async () => {
    const crafted = nonCanonicalBuffer();
    // Direct path: the exact calls CLI normalize makes (parse + rezip + core log).
    const direct = new DocxParser(crafted);
    const expectedBytes = direct.rezip();
    const expectedLog = readParserNormalizations(direct);
    expect(expectedLog.length).toBeGreaterThan(0);
    // Session path: open the same bytes, normalize, serialize (fixpoint-stable).
    const { store, sessionId } = await openBufferSession(crafted);
    const res = await handleNormalizeDocument(store, { sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`normalize failed: ${JSON.stringify(res)}`);
    expect(res.normalized).toEqual(expectedLog);
    expect(Buffer.compare(store.get(sessionId).parser.rezip(), expectedBytes)).toBe(0);
  });

  it('second run is a fixpoint (empty log, stable bytes)', async () => {
    const { store, sessionId } = await openBufferSession(nonCanonicalBuffer());
    const first = await handleNormalizeDocument(store, { sessionId });
    expect(first.ok).toBe(true);
    if (!first.ok) throw new Error('first normalize failed');
    const bytesAfterFirst = store.get(sessionId).parser.rezip();
    const second = await handleNormalizeDocument(store, { sessionId });
    expect(second.ok).toBe(true);
    if (!second.ok) throw new Error('second normalize failed');
    expect(second.normalized).toEqual([]);
    expect(Buffer.compare(store.get(sessionId).parser.rezip(), bytesAfterFirst)).toBe(0);
  });

  it('restores the session untouched when the package stays invalid', async () => {
    const buf = DocxParser.createNewDocument('blank').rezip();
    const xml = new AdmZip(buf).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
    const ids = [...xml.matchAll(/w14:paraId="([^"]+)"/g)].map((m) => m[1]);
    expect(ids.length).toBeGreaterThanOrEqual(2);
    const evil = xml.replace(`w14:paraId="${ids[1]}"`, `w14:paraId="${ids[0]}"`);
    const zipped = new AdmZip(buf);
    zipped.updateFile('word/document.xml', Buffer.from(evil, 'utf8'));
    const { store, sessionId } = await openBufferSession(zipped.toBuffer());
    const res = await handleNormalizeDocument(store, { sessionId });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('invalid package unexpectedly normalized');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.remedy.length).toBeGreaterThan(0);
    expect(store.get(sessionId).revision).toBe(1);
  });

  it('refuses an unknown session with code and remedy', async () => {
    const store = createMcpSessionStore();
    const res = await handleNormalizeDocument(store, { sessionId: 'sess_nope' });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('unknown session unexpectedly normalized');
    expect(res.code).toBe('E_SESSION');
    expect(res.remedy.length).toBeGreaterThan(0);
  });

  it('rejects wire violations with code and remedy', async () => {
    const { store } = await openBufferSession(nonCanonicalBuffer());
    const res = await handleNormalizeDocument(store, {});
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('missing sessionId unexpectedly normalized');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.remedy.length).toBeGreaterThan(0);
  });
});
