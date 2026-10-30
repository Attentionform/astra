import { describe, it, expect } from 'vitest';
import { AdmZip, DocxParser, listMacroPartsInBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleSanitizeDocument } from '../../src/lib/mcp/handlers/parity';

async function openBlankSession(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const store = createMcpSessionStore();
  const buf = DocxParser.createNewDocument('blank').rezip();
  const handle = await store.openBuffer(buf, 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

/** Session opened with macros preserved (refuse-default bypassed explicitly at open). */
async function openMacroSession(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const zip = new AdmZip(buf);
  zip.addFile('word/vbaProject.bin', Buffer.from('fake-macro-bytes'));
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(zip.toBuffer(), 'memory.docx', undefined, { allowMacros: true });
  return { store, sessionId: handle.sessionId };
}

function sessionBytes(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): Buffer {
  return store.get(sessionId).parser.rezip();
}

describe('sanitize_document handler', () => {
  it('strips macro parts while keeping body text', async () => {
    const { store, sessionId } = await openMacroSession();
    expect(listMacroPartsInBuffer(sessionBytes(store, sessionId))).toEqual(['word/vbaProject.bin']);
    const res = await handleSanitizeDocument(store, { sessionId, stripMacros: true });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`sanitize failed: ${JSON.stringify(res)}`);
    expect(res.strippedMacros).toContain('word/vbaProject.bin');
    expect(listMacroPartsInBuffer(sessionBytes(store, sessionId))).toEqual([]);
    const found = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Blank Document' });
    expect(found.ok).toBe(true);
    if (!found.ok) throw new Error('search failed');
    expect(found.matches.length).toBeGreaterThan(0);
    expect(res.revision).toBe(2);
    expect(res.dirty).toBe(true);
  });

  it('default options scrub without removing body text', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleSanitizeDocument(store, { sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`sanitize failed: ${JSON.stringify(res)}`);
    const found = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Blank Document' });
    expect(found.ok).toBe(true);
    if (!found.ok) throw new Error('search failed');
    expect(found.matches.length).toBeGreaterThan(0);
    expect(res.revision).toBe(2);
    expect(res.dirty).toBe(true);
  });

  it('rejects conflicting macro choices with code and remedy', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleSanitizeDocument(store, { sessionId, allowMacros: true, stripMacros: true });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('conflicting macro choices unexpectedly scrubbed');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.remedy.length).toBeGreaterThan(0);
    expect(store.get(sessionId).revision).toBe(1);
  });

  it('rejects unknown scrub keys with code and remedy (no flat-key drift)', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleSanitizeDocument(store, { sessionId, stripEverything: true });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('unknown scrub key unexpectedly scrubbed');
    expect(res.code).toBe('E_VALIDATION');
    expect(store.get(sessionId).revision).toBe(1);
  });

  it('refuses an unknown session with code and remedy', async () => {
    const store = createMcpSessionStore();
    const res = await handleSanitizeDocument(store, { sessionId: 'sess_nope' });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('unknown session unexpectedly scrubbed');
    expect(res.code).toBe('E_SESSION');
    expect(res.remedy.length).toBeGreaterThan(0);
  });
});
