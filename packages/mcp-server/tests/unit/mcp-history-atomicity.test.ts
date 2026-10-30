import { describe, it, expect, vi, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as engine from '@attentionform/astra-engine';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS, GIT_HISTORY_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpDirs: string[] = [];

async function makeSession(
  store: ReturnType<typeof createMcpSessionStore>,
  prefix: string
): Promise<{ sessionId: string; file: string }> {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), `astra-hist-atom-${prefix}-`));
  tmpDirs.push(dir);
  const file = path.join(dir, 'doc.docx');
  await fs.promises.writeFile(file, createBlankDocxBuffer('blank'));
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  expect(opened.ok).toBe(true);
  if (!opened.ok) throw new Error('open failed in test setup');
  return { sessionId: opened.sessionId, file };
}

afterEach(async () => {
  vi.restoreAllMocks();
  while (tmpDirs.length > 0) {
    const d = tmpDirs.pop() as string;
    try {
      await fs.promises.rm(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

/** First whitespace-delimited token of every TOON data row (skips header lines). */
function rowIds(toonText: string): string[] {
  return toonText
    .split('\n')
    .map((l) => l.split(/\s+/)[0])
    .filter((id) => !!id && !['>', 'ID'].includes(id) && !/^---/.test(id));
}

async function firstNodeId(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string
): Promise<string> {
  const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId, offset: 0, limit: 1 });
  expect(first.ok).toBe(true);
  if (!first.ok) throw new Error('inspect failed in test setup');
  const id = rowIds(first.toonText)[0];
  expect(typeof id).toBe('string');
  return id;
}

async function replaceAnchor(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string,
  text: string
): Promise<void> {
  const anchor = await firstNodeId(store, sessionId);
  const batch = await LEAN6_HANDLERS.execute_batch(store, {
    sessionId,
    operations: [{ action: 'replace', target: anchor, text }],
  });
  expect(batch.ok).toBe(true);
}

async function searchTexts(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string,
  query: string
): Promise<number> {
  const res = await LEAN6_HANDLERS.search_content(store, { sessionId, query });
  expect(res.ok).toBe(true);
  if (!res.ok) throw new Error('search failed in test setup');
  return res.matches.length;
}

async function saveExpectOk(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string,
  message: string
): Promise<string> {
  const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, message });
  expect(saved.ok).toBe(true);
  if (!saved.ok) throw new Error('save failed in test setup');
  return saved.commitOid;
}

function revisionMarkers(store: ReturnType<typeof createMcpSessionStore>, sessionId: string) {
  const handle = store.get(sessionId);
  return { revision: handle.revision, savedRevision: handle.savedRevision, redoLength: handle.redoStack.length };
}

describe('document history commit atomicity', () => {
  it('failed revert leaves revision, file bytes and redo stack unchanged', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'revert-fail');

    await replaceAnchor(store, sessionId, 'Atomic alpha text');
    await saveExpectOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Atomic beta text');
    await saveExpectOk(store, sessionId, 'beta');

    const before = revisionMarkers(store, sessionId);
    const bytesBefore = await fs.promises.readFile(file);
    expect(before.redoLength).toBe(0);
    expect(await searchTexts(store, sessionId, 'Atomic beta text')).toBeGreaterThan(0);

    const spy = vi.spyOn(engine, 'commitChange').mockRejectedValueOnce(new Error('injected commit outage'));
    const failed = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(spy).toHaveBeenCalledTimes(1);
    expect(failed.ok).toBe(false);
    if (failed.ok) throw new Error('revert should have failed');
    expect(failed.code).toBe('E_IO');
    expect(typeof failed.remedy).toBe('string');
    expect(failed.remedy.length).toBeGreaterThan(0);

    // Atomic rollback: revision markers, redo depth and disk bytes untouched,
    // and the live session still holds the pre-revert content.
    expect(revisionMarkers(store, sessionId)).toEqual(before);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesBefore)).toBe(0);
    expect(await searchTexts(store, sessionId, 'Atomic beta text')).toBeGreaterThan(0);

    // Retry with a healthy commit path restores the earlier saved state.
    const retried = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(retried.ok).toBe(true);
    if (!retried.ok) throw new Error('retry should have succeeded');
    expect(await searchTexts(store, sessionId, 'Atomic alpha text')).toBeGreaterThan(0);
    expect(await searchTexts(store, sessionId, 'Atomic beta text')).toBe(0);
  });

  it('failed save leaves prior file bytes and revision markers in place', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'save-fail');

    await replaceAnchor(store, sessionId, 'Saved baseline text');
    await saveExpectOk(store, sessionId, 'baseline');
    await replaceAnchor(store, sessionId, 'Unsaved candidate text');

    const before = revisionMarkers(store, sessionId);
    const bytesBefore = await fs.promises.readFile(file);

    const spy = vi.spyOn(engine, 'commitChange').mockRejectedValueOnce(new Error('injected commit outage'));
    const failed = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'candidate' });
    expect(spy).toHaveBeenCalledTimes(1);
    expect(failed.ok).toBe(false);
    if (failed.ok) throw new Error('save should have failed');
    expect(failed.code).toBe('E_IO');
    expect(typeof failed.remedy).toBe('string');

    expect(revisionMarkers(store, sessionId)).toEqual(before);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesBefore)).toBe(0);

    const retried = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'candidate' });
    expect(retried.ok).toBe(true);
  });

  it('failed redo hands the stack entry back and changes nothing', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'redo-fail');

    await replaceAnchor(store, sessionId, 'Redo guard alpha');
    await saveExpectOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Redo guard beta');
    await saveExpectOk(store, sessionId, 'beta');

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    expect(revisionMarkers(store, sessionId).redoLength).toBe(1);

    const before = revisionMarkers(store, sessionId);
    const bytesBefore = await fs.promises.readFile(file);

    const spy = vi.spyOn(engine, 'commitChange').mockRejectedValueOnce(new Error('injected commit outage'));
    const failed = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(spy).toHaveBeenCalledTimes(1);
    expect(failed.ok).toBe(false);
    if (failed.ok) throw new Error('redo should have failed');
    expect(failed.code).toBe('E_IO');
    expect(typeof (failed as { remedy?: unknown }).remedy).toBe('string');

    // The popped entry is handed back; disk, memory and markers are unchanged.
    expect(revisionMarkers(store, sessionId)).toEqual(before);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesBefore)).toBe(0);
    expect(await searchTexts(store, sessionId, 'Redo guard beta')).toBe(0);

    const retried = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(retried.ok).toBe(true);
    if (!retried.ok) throw new Error('retry should have succeeded');
    expect(await searchTexts(store, sessionId, 'Redo guard beta')).toBeGreaterThan(0);
  });

  it('failed redo preserves engine E_STATE verbatim instead of collapsing to E_IO', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'redo-estate');

    await replaceAnchor(store, sessionId, 'Redo estate alpha');
    await saveExpectOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Redo estate beta');
    await saveExpectOk(store, sessionId, 'beta');

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    expect(revisionMarkers(store, sessionId).redoLength).toBe(1);

    const spy = vi.spyOn(engine, 'commitChange').mockRejectedValueOnce(
      Object.assign(new Error('[E_STATE] injected state conflict'), {
        code: 'E_STATE',
        remedy: 're-inspect live vs saved state, then retry the operation',
      })
    );
    const failed = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(spy).toHaveBeenCalledTimes(1);
    expect(failed.ok).toBe(false);
    if (failed.ok) throw new Error('redo should have failed');
    expect(failed.code).toBe('E_STATE');
  });

  it('revert write failure restores staged memory and leaves disk unchanged', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'revert-write');

    await replaceAnchor(store, sessionId, 'Write guard alpha text');
    await saveExpectOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Write guard beta text');
    await saveExpectOk(store, sessionId, 'beta');

    const before = revisionMarkers(store, sessionId);
    const bytesBefore = await fs.promises.readFile(file);
    expect(await searchTexts(store, sessionId, 'Write guard beta text')).toBeGreaterThan(0);

    const spy = vi.spyOn(engine, 'writeFileNoFollow').mockRejectedValueOnce(new Error('injected write outage'));
    const failed = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    // Primary write (injected refusal) + failure-restore write: both route
    // through the hardened surface.
    expect(spy).toHaveBeenCalledTimes(2);
    expect(failed.ok).toBe(false);
    if (failed.ok) throw new Error('revert should have failed');
    expect(failed.code).toBe('E_IO');
    expect(typeof failed.remedy).toBe('string');
    expect(failed.remedy.length).toBeGreaterThan(0);

    expect(revisionMarkers(store, sessionId)).toEqual(before);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesBefore)).toBe(0);
    expect(await searchTexts(store, sessionId, 'Write guard beta text')).toBeGreaterThan(0);
    expect(await searchTexts(store, sessionId, 'Write guard alpha text')).toBe(0);

    const retried = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(retried.ok).toBe(true);
    if (!retried.ok) throw new Error('retry should have succeeded');
    expect(await searchTexts(store, sessionId, 'Write guard alpha text')).toBeGreaterThan(0);
  });

  it('redo write failure hands the entry back and restores memory', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'redo-write');

    await replaceAnchor(store, sessionId, 'Redo write alpha');
    await saveExpectOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Redo write beta');
    await saveExpectOk(store, sessionId, 'beta');

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    expect(revisionMarkers(store, sessionId).redoLength).toBe(1);

    const before = revisionMarkers(store, sessionId);
    const bytesBefore = await fs.promises.readFile(file);
    expect(await searchTexts(store, sessionId, 'Redo write alpha')).toBeGreaterThan(0);

    const spy = vi.spyOn(engine, 'writeFileNoFollow').mockRejectedValueOnce(new Error('injected write outage'));
    const failed = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    // Primary write (injected refusal) + failure-restore write: both route
    // through the hardened surface.
    expect(spy).toHaveBeenCalledTimes(2);
    expect(failed.ok).toBe(false);
    if (failed.ok) throw new Error('redo should have failed');
    expect(failed.code).toBe('E_IO');
    expect(typeof (failed as { remedy?: unknown }).remedy).toBe('string');

    expect(revisionMarkers(store, sessionId)).toEqual(before);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesBefore)).toBe(0);
    expect(await searchTexts(store, sessionId, 'Redo write alpha')).toBeGreaterThan(0);
    expect(await searchTexts(store, sessionId, 'Redo write beta')).toBe(0);

    const retried = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(retried.ok).toBe(true);
    if (!retried.ok) throw new Error('retry should have succeeded');
    expect(await searchTexts(store, sessionId, 'Redo write beta')).toBeGreaterThan(0);
  });

  it('malformed oid is E_VALIDATION; absent well-formed oid is E_NOT_FOUND', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'oid-codes');

    await replaceAnchor(store, sessionId, 'Oid code probe text');
    await saveExpectOk(store, sessionId, 'probe');

    const malformed = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      commitOid: 'not-an-oid',
    });
    expect(malformed.ok).toBe(false);
    if (malformed.ok) throw new Error('malformed oid should have failed');
    expect(malformed.code).toBe('E_VALIDATION');

    const absent = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      commitOid: 'f'.repeat(40),
    });
    expect(absent.ok).toBe(false);
    if (absent.ok) throw new Error('absent oid should have failed');
    expect(absent.code).toBe('E_NOT_FOUND');
  });
});

describe('history reads under contention', () => {
  it('history read during a held lock returns a consistent snapshot', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'busy-snapshot');

    await replaceAnchor(store, sessionId, 'Contention probe text');
    await saveExpectOk(store, sessionId, 'probe');

    const markersBefore = revisionMarkers(store, sessionId);

    // Hold the per-session lock, then queue a history read behind it.
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = () => resolve();
    });
    const holder = store.withSession(sessionId, () => gate);
    const pending = GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    await new Promise((r) => setTimeout(r, 30));
    release();
    const res = await pending;
    await holder;

    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error('history should have succeeded');
    expect(res.sessionId).toBe(sessionId);
    expect(res.documentPath).toBe(file);
    expect(res.totalCommits).toBeGreaterThanOrEqual(2);
    expect(res.history.length).toBeLessThanOrEqual(res.totalCommits);
    if (res.totalCommits > 0) {
      expect(res.history.length).toBeGreaterThan(0);
      expect(res.currentCommit).toBe(res.history[0].oid);
    } else {
      expect(res.currentCommit).toBeNull();
    }
    for (const entry of res.history) {
      expect(entry.oid).toMatch(/^[0-9a-f]{40}$/);
    }
    // Read-only: the read changed no revision state.
    expect(revisionMarkers(store, sessionId)).toEqual(markersBefore);
  });

  it('history read queued while the session is dropped reports a clean conflict', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'busy-dropped');

    await replaceAnchor(store, sessionId, 'Drop probe text');
    await saveExpectOk(store, sessionId, 'probe');

    // Hold the raw per-session lock, queue a history read behind it, then
    // drop the session before the read acquires the lock.
    const handle = store.get(sessionId);
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = () => resolve();
    });
    const holder = handle.mutex.runExclusive(() => gate);
    const pending = GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    await new Promise((r) => setTimeout(r, 30));
    store.clear();
    release();
    const res = await pending;
    await holder;

    // Clean envelope, never a throw or torn payload.
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('history should have failed for a dropped session');
    expect(res.code).toBe('E_CONFLICT');
    expect(typeof res.remedy).toBe('string');
    expect(res.remedy.length).toBeGreaterThan(0);
  });
});

describe('history no-follow: revert/redo/sidecar refuse link targets', () => {
  async function setupTwoSaves(
    store: ReturnType<typeof createMcpSessionStore>,
    prefix: string
  ): Promise<{ sessionId: string; file: string }> {
    const { sessionId, file } = await makeSession(store, prefix);
    await replaceAnchor(store, sessionId, 'No-follow alpha text');
    await saveExpectOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'No-follow beta text');
    await saveExpectOk(store, sessionId, 'beta');
    return { sessionId, file };
  }

  it('revert refuses a dangling-link docx target with E_SANDBOX and writes nothing outside', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await setupTwoSaves(store, 'revert-link');
    const outsideDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-hist-nf-out-'));
    tmpDirs.push(outsideDir);
    const outsideTarget = path.join(outsideDir, 'exfil.docx');
    const bytesBefore = await fs.promises.readFile(file);
    const before = revisionMarkers(store, sessionId);

    await fs.promises.rm(file);
    fs.symlinkSync(outsideTarget, file);

    const refused = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(refused.ok).toBe(false);
    if (refused.ok) throw new Error('revert through a link should have failed');
    expect(refused.code).toBe('E_SANDBOX');
    expect(typeof refused.remedy).toBe('string');
    expect(refused.remedy).toMatch(/symlink/i);

    // Nothing followed the link: outside untouched, the link itself kept
    // (never unlinked nor replaced), markers and live content unchanged.
    expect(fs.existsSync(outsideTarget)).toBe(false);
    expect(await fs.promises.readdir(outsideDir)).toEqual([]);
    expect(fs.lstatSync(file).isSymbolicLink()).toBe(true);
    expect(revisionMarkers(store, sessionId)).toEqual(before);
    expect(await searchTexts(store, sessionId, 'No-follow beta text')).toBeGreaterThan(0);

    // Legit control: the real file back in place reverts cleanly.
    await fs.promises.rm(file);
    await fs.promises.writeFile(file, bytesBefore);
    const retried = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(retried.ok).toBe(true);
    if (!retried.ok) throw new Error('legit revert should have succeeded');
    expect(await searchTexts(store, sessionId, 'No-follow alpha text')).toBeGreaterThan(0);
  });

  it('redo refuses a dangling-link docx target with E_SANDBOX and keeps the stack entry', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await setupTwoSaves(store, 'redo-link');
    const outsideDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-hist-nf-out-'));
    tmpDirs.push(outsideDir);
    const outsideTarget = path.join(outsideDir, 'exfil.docx');

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    expect(revisionMarkers(store, sessionId).redoLength).toBe(1);
    const bytesAlpha = await fs.promises.readFile(file);

    await fs.promises.rm(file);
    fs.symlinkSync(outsideTarget, file);
    const before = revisionMarkers(store, sessionId);

    const refused = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(refused.ok).toBe(false);
    if (refused.ok) throw new Error('redo through a link should have failed');
    expect(refused.code).toBe('E_SANDBOX');
    expect(typeof (refused as { remedy?: unknown }).remedy).toBe('string');

    expect(fs.existsSync(outsideTarget)).toBe(false);
    expect(await fs.promises.readdir(outsideDir)).toEqual([]);
    expect(fs.lstatSync(file).isSymbolicLink()).toBe(true);
    expect(revisionMarkers(store, sessionId)).toEqual(before);
    expect(revisionMarkers(store, sessionId).redoLength).toBe(1);

    // Legit control: the real file back in place redoes cleanly.
    await fs.promises.rm(file);
    await fs.promises.writeFile(file, bytesAlpha);
    const retried = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(retried.ok).toBe(true);
    if (!retried.ok) throw new Error('legit redo should have succeeded');
    expect(await searchTexts(store, sessionId, 'No-follow beta text')).toBeGreaterThan(0);
  });

  it('save refuses with E_SANDBOX when a sidecar working-tree leaf is a link', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'sidecar-link');
    await replaceAnchor(store, sessionId, 'Sidecar baseline text');
    await saveExpectOk(store, sessionId, 'baseline');
    await replaceAnchor(store, sessionId, 'Sidecar candidate text');
    const bytesBefore = await fs.promises.readFile(file);
    const before = revisionMarkers(store, sessionId);

    const outsideDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-hist-nf-out-'));
    tmpDirs.push(outsideDir);
    const outsideTarget = path.join(outsideDir, 'sidecar-exfil.xml');
    const repoDocXml = path.join(engine.getRepoDir(file), 'document.xml');
    const backupDocXml = path.join(outsideDir, 'document.xml.bak');
    await fs.promises.rename(repoDocXml, backupDocXml);
    fs.symlinkSync(outsideTarget, repoDocXml);

    const refused = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'candidate' });
    expect(refused.ok).toBe(false);
    if (refused.ok) throw new Error('save with a linked sidecar leaf should have failed');
    expect(refused.code).toBe('E_SANDBOX');
    expect(typeof refused.remedy).toBe('string');
    expect(refused.remedy).toMatch(/symlink/i);

    // The sidecar write never landed outside; the docx was restored in place.
    expect(fs.existsSync(outsideTarget)).toBe(false);
    expect(fs.lstatSync(repoDocXml).isSymbolicLink()).toBe(true);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesBefore)).toBe(0);
    expect(revisionMarkers(store, sessionId)).toEqual(before);

    // Legit control: sidecar restored, the same save lands cleanly.
    await fs.promises.rm(repoDocXml);
    await fs.promises.rename(backupDocXml, repoDocXml);
    const retried = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'candidate' });
    expect(retried.ok).toBe(true);
  });
});
