import { describe, it, expect } from 'vitest';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { handleValidateDocument } from '../../src/lib/mcp/handlers/parity';

async function openBlankSession(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const store = createMcpSessionStore();
  const buf = DocxParser.createNewDocument('blank').rezip();
  const handle = await store.openBuffer(buf, 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

/** Session holding two paragraphs sharing one native paraId (opens cleanly, fails the save gate). */
async function openDuplicatedIdSession(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const xml = new AdmZip(buf).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
  const ids = [...xml.matchAll(/w14:paraId="([^"]+)"/g)].map((m) => m[1]);
  expect(ids.length).toBeGreaterThanOrEqual(2);
  const evil = xml.replace(`w14:paraId="${ids[1]}"`, `w14:paraId="${ids[0]}"`);
  const zipped = new AdmZip(buf);
  zipped.updateFile('word/document.xml', Buffer.from(evil, 'utf8'));
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(zipped.toBuffer(), 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

describe('validate_document handler', () => {
  it('reports a clean document as valid without touching revision or dirty', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleValidateDocument(store, { sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`validate failed: ${JSON.stringify(res)}`);
    expect(res.sessionId).toBe(sessionId);
    expect(res.validation.valid).toBe(true);
    expect(res.validation.errorCount).toBe(0);
    expect(res.validation.errors).toEqual([]);
    expect(res.revision).toBe(1);
    expect(res.dirty).toBe(false);
  });

  it('repeat audits stay identical (read-only: no revision bump, no dirty flip)', async () => {
    const { store, sessionId } = await openBlankSession();
    const first = await handleValidateDocument(store, { sessionId });
    const second = await handleValidateDocument(store, { sessionId });
    expect(first.ok).toBe(true);
    expect(second.ok).toBe(true);
    if (!first.ok || !second.ok) throw new Error('validate failed');
    expect(second).toEqual(first);
    expect(store.get(sessionId).revision).toBe(1);
    expect(store.isDirty(store.get(sessionId))).toBe(false);
  });

  it('detects a duplicated paragraph id before save', async () => {
    const { store, sessionId } = await openDuplicatedIdSession();
    const res = await handleValidateDocument(store, { sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`validate failed: ${JSON.stringify(res)}`);
    expect(res.validation.valid).toBe(false);
    expect(res.validation.errorCount).toBeGreaterThan(0);
    expect(res.validation.errors.length).toBeGreaterThan(0);
    expect(res.validation.errors[0].description).toContain('duplicate');
    expect(res.revision).toBe(1);
    expect(res.dirty).toBe(false);
  });

  it('refuses an unknown session with code and remedy', async () => {
    const store = createMcpSessionStore();
    const res = await handleValidateDocument(store, { sessionId: 'sess_nope' });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('unknown session unexpectedly validated');
    expect(res.code).toBe('E_SESSION');
    expect(res.remedy.length).toBeGreaterThan(0);
  });

  it('rejects wire violations with code and remedy', async () => {
    const { store } = await openBlankSession();
    const res = await handleValidateDocument(store, {});
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('missing sessionId unexpectedly validated');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.remedy.length).toBeGreaterThan(0);
  });
});
