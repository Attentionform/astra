import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS, GIT_HISTORY_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpDirs: string[] = [];

async function makeSession(
  store: ReturnType<typeof createMcpSessionStore>,
  prefix: string
): Promise<{ sessionId: string; file: string; dir: string }> {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), `astra-git-hist-${prefix}-`));
  tmpDirs.push(dir);
  const file = path.join(dir, 'doc.docx');
  await fs.promises.writeFile(file, createBlankDocxBuffer('blank'));
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  expect(opened.ok).toBe(true);
  if (!opened.ok) throw new Error('open failed in test setup');
  return { sessionId: opened.sessionId, file, dir };
}

afterEach(async () => {
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

describe('mcp git history + revert + redo', () => {
  it('open records the initial baseline commit', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'open');
    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    expect(hist.documentPath).toBe(file);
    expect(hist.totalCommits).toBe(1);
    expect(hist.history).toHaveLength(1);
    expect(hist.currentCommit).toBe(hist.history[0].oid);
    expect(hist.history[0].oid).toMatch(/^[0-9a-f]{40}$/);
    expect(hist.history[0].message).toBe('Initial state');
    expect(typeof hist.history[0].author).toBe('string');
    expect(hist.history[0].timestamp).toBeTruthy();
    // Read-only: still clean after browsing history.
    const insp = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(insp.ok).toBe(true);
    if (insp.ok) expect(insp.dirty).toBe(false);
  });

  it('save records custom and fallback messages with hygiene', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'save-msg');

    // Clean first save -> initial baseline fallback.
    const first = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(first.ok).toBe(true);
    if (!first.ok) return;
    expect(first.commitMessage).toBe('save_document: rev 1 (initial baseline)');
    expect(first.commitOid).toMatch(/^[0-9a-f]{40}$/);

    // Edited save with a multi-line custom message + author.
    await replaceAnchor(store, sessionId, 'Custom message probe');
    const custom = await LEAN6_HANDLERS.save_document(store, {
      sessionId,
      message: 'Checkpoint\nsecond line',
      author: 'History Tester',
    });
    expect(custom.ok).toBe(true);
    if (!custom.ok) return;
    expect(custom.commitMessage).toBe('Checkpoint second line');
    expect(custom.commitMessage).not.toMatch(/\n/);

    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    expect(hist.totalCommits).toBe(3);
    expect(hist.history[0].message).toBe('Checkpoint second line');
    expect(hist.history[0].author).toBe('History Tester');

    // Edited save without a message -> revision-delta fallback, default author.
    await replaceAnchor(store, sessionId, 'Fallback probe');
    const fallback = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(fallback.ok).toBe(true);
    if (!fallback.ok) return;
    expect(fallback.commitMessage).toBe('save_document: rev 3 (1 ops since last save)');

    const hist2 = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist2.ok).toBe(true);
    if (!hist2.ok) return;
    expect(hist2.history[0].author).toBe('ASTra Agent');
  });

  it('history paginates with limit/offset and rejects unknown sessions', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'page');
    await replaceAnchor(store, sessionId, 'Page v1');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'v1' });
    await replaceAnchor(store, sessionId, 'Page v2');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'v2' });
    await replaceAnchor(store, sessionId, 'Page v3');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'v3' });

    const page1 = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId, limit: 2 });
    expect(page1.ok).toBe(true);
    if (!page1.ok) return;
    expect(page1.totalCommits).toBe(4);
    expect(page1.history).toHaveLength(2);
    expect(page1.history[0].message).toBe('v3');
    expect(page1.history[1].message).toBe('v2');
    expect(page1.currentCommit).toBe(page1.history[0].oid);

    const page2 = await GIT_HISTORY_HANDLERS.get_document_history(store, {
      sessionId,
      limit: 2,
      offset: 2,
    });
    expect(page2.ok).toBe(true);
    if (!page2.ok) return;
    expect(page2.totalCommits).toBe(4);
    expect(page2.history).toHaveLength(2);
    expect(page2.history[0].message).toBe('v1');
    expect(page2.history[1].message).toBe('Initial state');

    const missing = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId: 'sess_nope' });
    expect(missing.ok).toBe(false);
    if (!missing.ok) {
      expect(missing.code).toBe('E_SESSION');
      expect(missing.message).toBe('Unknown session "sess_nope" (expired or never opened)');
      expect(missing.remedy).toBe('Call open_document first, then retry with the returned sessionId.');
    }
  });

  it('single-step revert restores the previously saved bytes', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'revert1');

    await replaceAnchor(store, sessionId, 'Alpha version one');
    const saveA = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'alpha' });
    expect(saveA.ok).toBe(true);
    const bytesA = await fs.promises.readFile(file);

    await replaceAnchor(store, sessionId, 'Beta version two');
    const saveB = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'beta' });
    expect(saveB.ok).toBe(true);
    if (!saveB.ok) return;
    expect(await searchTexts(store, sessionId, 'Beta version two')).toBeGreaterThan(0);

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    expect(reverted.revertedToCommit).toBe(saveA.ok ? saveA.commitOid : '');
    expect(reverted.dirty).toBe(false);
    expect(reverted.path).toBe(file);

    expect(await searchTexts(store, sessionId, 'Alpha version one')).toBeGreaterThan(0);
    expect(await searchTexts(store, sessionId, 'Beta version two')).toBe(0);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesA)).toBe(0);
  });

  it('revert by commit oid jumps to an arbitrary depth as a new head', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'revert-oid');

    await replaceAnchor(store, sessionId, 'Era one text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'era-one' });
    await replaceAnchor(store, sessionId, 'Era two text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'era-two' });
    await replaceAnchor(store, sessionId, 'Era three text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'era-three' });

    const before = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(before.ok).toBe(true);
    if (!before.ok) return;
    expect(before.totalCommits).toBe(4);
    const depthBefore = before.totalCommits;
    const headBefore = before.currentCommit;
    const eraOneOid = before.history[2].oid;
    expect(before.history[2].message).toBe('era-one');

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      commitOid: eraOneOid,
      message: 'backup restore',
    });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    expect(reverted.revertedToCommit).toBe(eraOneOid);

    expect(await searchTexts(store, sessionId, 'Era one text')).toBeGreaterThan(0);
    expect(await searchTexts(store, sessionId, 'Era three text')).toBe(0);

    // Append-only: a NEW head sits on top, refs never rewound.
    const after = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok) return;
    expect(after.totalCommits).toBe(depthBefore + 1);
    expect(after.currentCommit).not.toBe(headBefore);
    expect(after.history[0].message).toMatch(/^Revert to [0-9a-f]{7}: backup restore$/);
  });

  it('redo re-applies the reverted state; empty redo reports atLatestState', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'redo');

    await replaceAnchor(store, sessionId, 'Redo alpha text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'alpha' });
    await replaceAnchor(store, sessionId, 'Redo beta text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'beta' });
    const bytesBeta = await fs.promises.readFile(file);

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    expect(await searchTexts(store, sessionId, 'Redo beta text')).toBe(0);

    const redone = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(redone.ok).toBe(true);
    if (!redone.ok) return;
    expect(redone.dirty).toBe(false);
    expect(redone.commitOid).toMatch(/^[0-9a-f]{40}$/);
    expect(await searchTexts(store, sessionId, 'Redo beta text')).toBeGreaterThan(0);
    expect(Buffer.compare(await fs.promises.readFile(file), bytesBeta)).toBe(0);

    // Stack drained: a second redo has nothing left.
    const empty = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(empty.ok).toBe(false);
    if (!empty.ok) {
      expect(empty.code).toBe('E_NOT_FOUND');
      expect((empty as { atLatestState?: boolean }).atLatestState).toBe(true);
    }
  });

  it('fresh sessions have an empty redo stack', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'redo-empty');
    const res = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(res.ok).toBe(false);
    if (!res.ok) {
      expect(res.code).toBe('E_NOT_FOUND');
      expect((res as { atLatestState?: boolean }).atLatestState).toBe(true);
      expect(typeof res.remedy).toBe('string');
    }
  });

  it('dirty sessions refuse revert without discard consent', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'dirty');
    await replaceAnchor(store, sessionId, 'Dirty unsaved text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'baseline' });
    await replaceAnchor(store, sessionId, 'Dirty unsaved text');

    const refused = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      discardUnsaved: false,
    });
    expect(refused.ok).toBe(false);
    if (!refused.ok) {
      expect(refused.code).toBe('E_CONFLICT');
      expect(typeof refused.remedy).toBe('string');
    }
    // Nothing was discarded: the unsaved text is still live.
    expect(await searchTexts(store, sessionId, 'Dirty unsaved text')).toBeGreaterThan(0);

    // Bare reverts fail closed when dirty.
    const bare = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(bare.ok).toBe(false);
    if (!bare.ok) expect(bare.code).toBe('E_CONFLICT');
    expect(await searchTexts(store, sessionId, 'Dirty unsaved text')).toBeGreaterThan(0);

    // Explicit consent discards and restores the baseline.
    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1, discardUnsaved: true });
    expect(reverted.ok).toBe(true);
  });

  it('unknown oids fail closed; steps and oid are mutually exclusive', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'bad-oid');
    await replaceAnchor(store, sessionId, 'Oid probe text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'probe' });

    const unknown = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      commitOid: 'f'.repeat(40),
    });
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) expect(unknown.code).toBe('E_NOT_FOUND');

    const malformed = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      commitOid: 'not-an-oid',
    });
    expect(malformed.ok).toBe(false);
    if (!malformed.ok) expect(malformed.code).toBe('E_VALIDATION');

    const both = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      commitOid: 'e'.repeat(40),
    });
    expect(both.ok).toBe(false);
    if (!both.ok) expect(both.code).toBe('E_VALIDATION');

    const tooDeep = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 99 });
    expect(tooDeep.ok).toBe(false);
    if (!tooDeep.ok) expect(tooDeep.code).toBe('E_NOT_FOUND');
  });
});

describe('history/ux signaling pins', () => {
  it('net-zero edits still count as dirty: conflict names it + dirtyReason', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'netzero');
    await replaceAnchor(store, sessionId, 'Net-zero baseline text');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'baseline' });

    // Net-zero round trip: insert a temp node, then delete it again.
    const anchor = await firstNodeId(store, sessionId);
    const ins = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'temp scratch', id: 'temp_scratch_node' },
      ],
    });
    expect(ins.ok).toBe(true);
    const del = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'delete', target: 'temp_scratch_node' }],
    });
    expect(del.ok).toBe(true);

    // Content matches the save, but revisions are pending: still dirty.
    const insp = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(insp.ok).toBe(true);
    if (insp.ok) expect(insp.dirty).toBe(true);

    const refused = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(refused.ok).toBe(false);
    if (!refused.ok) {
      expect(refused.code).toBe('E_CONFLICT');
      expect(refused.remedy).toMatch(/net-zero/);
      expect(refused.remedy).toMatch(/discardUnsaved:true/);
      expect((refused as { dirtyReason?: unknown }).dirtyReason).toMatch(/pending revision/);
    }
  });

  it('empty redo remedy names the cleared-by-edit rule', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'redo-remedy');
    const res = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(res.ok).toBe(false);
    if (!res.ok) {
      expect(res.code).toBe('E_NOT_FOUND');
      expect(res.remedy).toMatch(/clears redo/);
      expect((res as { atLatestState?: boolean }).atLatestState).toBe(true);
    }
  });

  it('batch signals a cleared redo stack and stays silent otherwise', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'redo-signal');
    await replaceAnchor(store, sessionId, 'Signal alpha');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'alpha' });
    await replaceAnchor(store, sessionId, 'Signal beta');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'beta' });

    // No redo pending: the batch envelope carries no redo signal.
    const anchor0 = await firstNodeId(store, sessionId);
    const quiet = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'replace', target: anchor0, text: 'Signal beta touched' }],
    });
    expect(quiet.ok).toBe(true);
    if (!quiet.ok) return;
    expect(quiet.redoCleared).toBeUndefined();
    expect(quiet.redoClearedWarning).toBeUndefined();
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'beta touched' });

    // Revert to alpha: beta-touched becomes the redo source.
    const rev = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(rev.ok).toBe(true);

    // A fresh edit clears that redo entry and says so.
    const anchor = await firstNodeId(store, sessionId);
    const batched = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'replace', target: anchor, text: 'Signal edited after revert' }],
    });
    expect(batched.ok).toBe(true);
    if (!batched.ok) return;
    expect(batched.redoCleared).toBe(true);
    expect(batched.redoClearedWarning).toMatch(/cleared 1 redo/);

    // The cleared entry is really gone.
    const empty = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(empty.ok).toBe(false);
    if (!empty.ok) expect(empty.code).toBe('E_NOT_FOUND');
  });

  it('history metadata is consistent even for pre-tracking commits', async () => {
    const storeA = createMcpSessionStore();
    const { sessionId, file } = await makeSession(storeA, 'meta-fallback');
    await replaceAnchor(storeA, sessionId, 'Meta fallback text');
    await LEAN6_HANDLERS.save_document(storeA, { sessionId, message: 'fallback save' });

    // A fresh store loses the in-memory audit bridge: every entry falls back.
    const storeB = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(storeB, { path: file });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const hist = await GIT_HISTORY_HANDLERS.get_document_history(storeB, { sessionId: opened.sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    expect(hist.totalCommits).toBeGreaterThanOrEqual(2);
    for (const entry of hist.history) {
      expect(typeof entry.timestampMs).toBe('number');
      expect(typeof entry.opCount).toBe('number');
      expect(Array.isArray(entry.affectedNodeIds)).toBe(true);
      expect(entry.resultingHeadOid).toBe(entry.oid);
    }
  });

  it('preview carries the positional flag; dirty preview names net-zero', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'preview-flag');
    await replaceAnchor(store, sessionId, 'Flag alpha');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'alpha' });
    await replaceAnchor(store, sessionId, 'Flag beta');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'beta' });

    const preview = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, steps: 1 });
    expect(preview.ok).toBe(true);
    if (!preview.ok) return;
    expect(typeof preview.changedValuesPositional).toBe('boolean');

    await replaceAnchor(store, sessionId, 'Flag unsaved marker');
    const dirtyPreview = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, steps: 1 });
    expect(dirtyPreview.ok).toBe(true);
    if (!dirtyPreview.ok) return;
    expect(dirtyPreview.dirty).toBe(true);
    expect(dirtyPreview.dirtyHandling).toMatch(/et-zero/);
  });

  it('default revert quotes the target message', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'quote-default');
    await replaceAnchor(store, sessionId, 'Quote alpha');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'quoted keeper' });
    await replaceAnchor(store, sessionId, 'Quote beta');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'throwaway' });

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    expect(reverted.commitMessage).toMatch(/^Revert to [0-9a-f]{7}: /);
    expect(reverted.commitMessage).toContain('"quoted keeper"');
  });

  it('redo success carries changedValues', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'redo-values');
    await replaceAnchor(store, sessionId, 'Values alpha');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'alpha' });
    await replaceAnchor(store, sessionId, 'Values beta');
    await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'beta' });

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    const redone = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(redone.ok).toBe(true);
    if (!redone.ok) return;
    expect(Array.isArray(redone.changedValues)).toBe(true);
    expect(typeof redone.changedValuesTruncated).toBe('boolean');
  });

  it('dirty close warns; clean close stays quiet', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'close-dirty');
    await replaceAnchor(store, sessionId, 'Close unsaved text');

    const closed = await LEAN6_HANDLERS.close_document(store, { sessionId });
    expect(closed.ok).toBe(true);
    if (!closed.ok) return;
    expect(closed.dirty).toBe(true);
    expect(closed.discardedRevisions).toBeGreaterThan(0);
    expect(closed.warning).toMatch(/discarded/);

    const clean = await makeSession(store, 'close-clean');
    const closedClean = await LEAN6_HANDLERS.close_document(store, { sessionId: clean.sessionId });
    expect(closedClean.ok).toBe(true);
    if (!closedClean.ok) return;
    expect(closedClean.dirty).toBe(false);
    expect(closedClean.discardedRevisions).toBe(0);
    expect(closedClean.warning).toBeUndefined();
  });
});
