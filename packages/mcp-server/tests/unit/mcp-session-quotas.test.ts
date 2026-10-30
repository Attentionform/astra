import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import {
  createMcpSessionStore,
  SessionRedoStack,
  SESSION_REDO_MAX_DEPTH,
  SESSION_REDO_MAX_BYTES,
  resolveRedoMaxDepth,
  resolveRedoMaxBytes,
} from '../../src/lib/session-store';
import {
  commitChange,
  getRepoDir,
  initAstraRepo,
  HISTORY_MAX_COMMITS_PER_DOC,
  HISTORY_MAX_SIDECAR_BYTES,
  resolveHistoryMaxCommits,
  resolveHistoryMaxSidecarBytes,
} from '@attentionform/astra-engine';

function tinySnap(text: string): Map<string, Buffer> {
  return new Map<string, Buffer>([['word/document.xml', Buffer.from(text, 'utf8')]]);
}

describe('session redo quotas (depth + bytes, loud)', () => {
  it('defaults are high so normal chains never hit them', () => {
    expect(SESSION_REDO_MAX_DEPTH).toBeGreaterThanOrEqual(50);
    expect(SESSION_REDO_MAX_BYTES).toBeGreaterThanOrEqual(10 * 1024 * 1024);
    expect(resolveRedoMaxDepth()).toBe(SESSION_REDO_MAX_DEPTH);
    expect(resolveRedoMaxBytes()).toBe(SESSION_REDO_MAX_BYTES);
    const store = createMcpSessionStore();
    expect(store.getRedoMaxDepth()).toBe(SESSION_REDO_MAX_DEPTH);
    expect(store.getRedoMaxBytes()).toBe(SESSION_REDO_MAX_BYTES);
  });

  it('depth cap refuses loudly with raise-the-limit remedy', async () => {
    const store = createMcpSessionStore({ redoMaxDepth: 2, redoMaxBytes: 10 * 1024 * 1024 });
    const h = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    expect(h.redoStack.getMaxDepth()).toBe(2);
    h.redoStack.push(tinySnap('a'));
    h.redoStack.push(tinySnap('b'));
    expect(h.redoStack.length).toBe(2);
    try {
      h.redoStack.push(tinySnap('c'));
      expect.unreachable('third push must throw E_QUOTA');
    } catch (err: unknown) {
      expect((err as { code?: string }).code).toBe('E_QUOTA');
      expect(String((err as Error).message)).toMatch(/E_QUOTA|quota|cap/i);
      expect(String((err as { remedy?: unknown }).remedy)).toMatch(/Raise the limit/i);
    }
    expect(h.redoStack.length).toBe(2);
  });

  it('bytes cap refuses loudly', () => {
    const stack = new SessionRedoStack({ maxDepth: 100, maxBytes: 10 });
    stack.push(tinySnap('12345'));
    expect(stack.length).toBe(1);
    try {
      stack.push(tinySnap('6789012345'));
      expect.unreachable('bytes push must throw E_QUOTA');
    } catch (err: unknown) {
      expect((err as { code?: string }).code).toBe('E_QUOTA');
      expect(String((err as { remedy?: unknown }).remedy)).toMatch(/Raise the limit/i);
    }
  });

  it('normal sessions push small chains without hitting defaults', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    h.redoStack.push(tinySnap('hello'));
    h.redoStack.push(tinySnap('world'));
    expect(h.redoStack.length).toBe(2);
    expect(h.redoStack.currentBytes()).toBeGreaterThan(0);
  });
});

describe('eviction refuses locked/dirty, all-busy open fails closed', () => {
  it('evictExpired skips locked and dirty sessions', async () => {
    const store = createMcpSessionStore({ idleMs: 60 * 1000 });
    const clean = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    const dirty = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    store.bumpRevision(dirty);
    expect(store.isDirty(dirty)).toBe(true);
    expect(store.isDirty(clean)).toBe(false);
    // Hold the clean session lock so it counts as busy too.
    let release!: () => void;
    const gate = new Promise<void>((r) => { release = () => r(); });
    const held = store.withSession(clean.sessionId, () => gate);
    await new Promise((r) => setTimeout(r, 20));
    expect(clean.mutex.isLocked()).toBe(true);
    // Force idle for both (far past idleMs).
    const past = Date.now() - 120 * 1000;
    clean.lastAccess = past;
    dirty.lastAccess = past;
    const evicted = store.evictExpired(Date.now());
    // Both refused: locked + dirty survive.
    expect(evicted).toEqual([]);
    expect(store.has(clean.sessionId)).toBe(true);
    expect(store.has(dirty.sessionId)).toBe(true);
    release();
    await held;
    // After unlock, dirty still refused but clean becomes evictable.
    clean.lastAccess = past;
    dirty.lastAccess = past;
    const evicted2 = store.evictExpired(Date.now());
    expect(evicted2).toContain(clean.sessionId);
    expect(evicted2).not.toContain(dirty.sessionId);
    expect(store.has(dirty.sessionId)).toBe(true);
  });

  it('all-busy store refuses new open with E_QUOTA and keeps uncommitted edits', async () => {
    const store = createMcpSessionStore({ maxDocs: 2 });
    const a = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    const b = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    store.bumpRevision(a);
    store.bumpRevision(b);
    expect(store.isDirty(a)).toBe(true);
    expect(store.isDirty(b)).toBe(true);
    try {
      await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
      expect.unreachable('third open must throw E_QUOTA');
    } catch (err: unknown) {
      expect((err as { code?: string }).code).toBe('E_QUOTA');
      expect(String((err as { remedy?: unknown }).remedy)).toMatch(/close.*clean|raise the limit/i);
    }
    expect(store.has(a.sessionId)).toBe(true);
    expect(store.has(b.sessionId)).toBe(true);
    expect(store.size()).toBe(2);
  });

  it('clean sessions still evict normally (unaffected path)', async () => {
    const store = createMcpSessionStore({ maxDocs: 2 });
    const a = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    const b = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    expect(store.isDirty(a)).toBe(false);
    const c = await store.openBuffer(createBlankDocxBuffer(), 'memory.docx');
    // Oldest clean session rotated out.
    expect(store.has(a.sessionId)).toBe(false);
    expect(store.has(b.sessionId)).toBe(true);
    expect(store.has(c.sessionId)).toBe(true);
  });
});

describe('history quotas (commits + sidecar bytes, loud)', () => {
  it('defaults are high', () => {
    expect(HISTORY_MAX_COMMITS_PER_DOC).toBeGreaterThanOrEqual(500);
    expect(HISTORY_MAX_SIDECAR_BYTES).toBeGreaterThanOrEqual(100 * 1024 * 1024);
    expect(resolveHistoryMaxCommits()).toBe(HISTORY_MAX_COMMITS_PER_DOC);
    expect(resolveHistoryMaxSidecarBytes()).toBe(HISTORY_MAX_SIDECAR_BYTES);
  });

  it('commit count cap refuses loudly', async () => {
    const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-hist-quota-'));
    try {
      const file = path.join(dir, 'q.docx');
      await fs.promises.writeFile(file, createBlankDocxBuffer('blank'));
      const xml = '<w:doc>q</w:doc>';
      await initAstraRepo(file, xml, {});
      await commitChange(file, xml, 'c1', {}, undefined, undefined, { maxCommits: 3 });
      await commitChange(file, xml, 'c2', {}, undefined, undefined, { maxCommits: 3 });
      try {
        await commitChange(file, xml, 'c3', {}, undefined, undefined, { maxCommits: 3 });
        expect.unreachable('third commit must throw E_QUOTA');
      } catch (err: unknown) {
        expect((err as { code?: string }).code).toBe('E_QUOTA');
        expect(String((err as Error).message)).toMatch(/E_QUOTA|quota|cap/i);
        expect(String((err as { remedy?: unknown }).remedy)).toMatch(/Raise the limit/i);
      }
      expect(getRepoDir(file)).toContain('.astra');
    } finally {
      await fs.promises.rm(dir, { recursive: true, force: true });
    }
  });

  it('sidecar bytes cap refuses loudly', async () => {
    const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-hist-bytes-'));
    try {
      const file = path.join(dir, 'b.docx');
      await fs.promises.writeFile(file, createBlankDocxBuffer('blank'));
      const xml = '<w:doc>b</w:doc>';
      await initAstraRepo(file, xml, {});
      try {
        await commitChange(file, xml, 'big', {}, undefined, undefined, { maxSidecarBytes: 1 });
        expect.unreachable('tiny sidecar cap must throw E_QUOTA');
      } catch (err: unknown) {
        expect((err as { code?: string }).code).toBe('E_QUOTA');
        expect(String((err as { remedy?: unknown }).remedy)).toMatch(/Raise the limit/i);
      }
    } finally {
      await fs.promises.rm(dir, { recursive: true, force: true });
    }
  });
});
