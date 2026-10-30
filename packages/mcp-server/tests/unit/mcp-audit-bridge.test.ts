import { describe, it, expect, afterEach } from 'vitest';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createBlankDocxBuffer, checkoutCommit } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import {
  AUDIT_RETENTION_NOTE,
  EXPORT_AUDIT_TRAIL_DESCRIPTION,
  GIT_HISTORY_DESCRIPTIONS,
  GIT_HISTORY_HANDLERS,
  LEAN6_HANDLERS,
  handleExportAuditTrail,
} from '../../src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpDirs: string[] = [];

async function makeSession(
  store: ReturnType<typeof createMcpSessionStore>,
  prefix: string
): Promise<{ sessionId: string; file: string }> {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), `astra-audit-bridge-${prefix}-`));
  tmpDirs.push(dir);
  const file = path.join(dir, 'doc.docx');
  await fs.promises.writeFile(file, createBlankDocxBuffer('blank'));
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  expect(opened.ok).toBe(true);
  if (!opened.ok) throw new Error('open failed in test setup');
  return { sessionId: opened.sessionId, file };
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

async function searchCount(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string,
  query: string
): Promise<number> {
  const res = await LEAN6_HANDLERS.search_content(store, { sessionId, query });
  expect(res.ok).toBe(true);
  if (!res.ok) throw new Error('search failed in test setup');
  return res.matches.length;
}

async function saveOk(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string,
  message: string
): Promise<string> {
  const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, message });
  expect(saved.ok).toBe(true);
  if (!saved.ok) throw new Error('save failed in test setup');
  return saved.commitOid;
}

function markers(store: ReturnType<typeof createMcpSessionStore>, sessionId: string) {
  const h = store.get(sessionId);
  return { revision: h.revision, savedRevision: h.savedRevision, redoLength: h.redoStack.length };
}

describe('mcp audit bridge', () => {
  it('revert with dirty state echoes counts and node ids without text', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'dirty-echo');

    await replaceAnchor(store, sessionId, 'Echo baseline alpha');
    await saveOk(store, sessionId, 'baseline');
    const discardedText = 'Unsaved discard me XYZ-12345';
    await replaceAnchor(store, sessionId, discardedText);

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1, discardUnsaved: true });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    const echo = reverted.discardedUnsaved;
    expect(echo.wasDirty).toBe(true);
    expect(echo.discardedNodeCount).toBeGreaterThan(0);
    expect(echo.discardedNodeIds.length).toBeGreaterThan(0);
    // Counts + ids only: no text echo.
    expect(Object.keys(echo as unknown as Record<string, unknown>).sort()).toEqual(['discardedNodeCount', 'discardedNodeIds', 'truncated', 'wasDirty']);
    expect(echo.discardedNodeIds.length).toBeLessThanOrEqual(20);
    // And the drop really happened: baseline is back, unsaved text is gone.
    expect(await searchCount(store, sessionId, 'Echo baseline alpha')).toBe(0);
    expect(await searchCount(store, sessionId, discardedText)).toBe(0);
  });

  it('clean revert echoes the zero summary and baselines carry audit meta', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'clean-echo');

    await replaceAnchor(store, sessionId, 'Clean alpha one');
    await saveOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Clean beta two');
    await saveOk(store, sessionId, 'beta');

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    expect(reverted.discardedUnsaved).toEqual({
      wasDirty: false,
      discardedNodeCount: 0,
      discardedNodeIds: [],
      truncated: false,
    });

    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    const baseline = hist.history[hist.history.length - 1];
    expect(baseline.message).toBe('Initial state');
    expect(typeof baseline.timestampMs).toBe('number');
    expect(baseline.opCount).toBe(0);
    expect(baseline.affectedNodeIds).toEqual([]);
    expect(baseline.resultingHeadOid).toBe(baseline.oid);
  });

  it('revert with author+reason uses the T-5 subject; history exposes audit depth', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 't5');

    await replaceAnchor(store, sessionId, 'Audit alpha text');
    await saveOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Audit beta text');
    await saveOk(store, sessionId, 'beta');

    const before = Date.now();
    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      author: 'Audit Author',
      reason: 'needs rollback',
    });
    const after = Date.now();
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    const short = reverted.revertedToCommit.slice(0, 7);
    expect(reverted.commitMessage).toBe(`Revert [1 steps] to ${short} by Audit Author: needs rollback`);
    expect(reverted.commitOid).toMatch(/^[0-9a-f]{40}$/);

    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    const head = hist.history[0];
    expect(head.message).toBe(reverted.commitMessage);
    expect(head.author).toBe('Audit Author');
    expect(head.oid).toBe(reverted.commitOid);
    expect(head.resultingHeadOid).toBe(head.oid);
    // Ms-precision timestamp recorded at commit time.
    expect(head.timestamp).toBeInstanceOf(Date);
    expect(head.timestampMs).toBeGreaterThanOrEqual(before);
    expect(head.timestampMs).toBeLessThanOrEqual(after);
    expect(head.timestamp.getTime()).toBe(head.timestampMs);
    // Op count + affected nodes from the bridge.
    expect(head.opCount).toBe(1);
    expect(Array.isArray(head.affectedNodeIds)).toBe(true);
    expect((head.affectedNodeIds as string[]).length).toBeGreaterThan(0);
  });

  it('revert by oid counts steps-back in the T-5 subject', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 't5-oid');

    await replaceAnchor(store, sessionId, 'Era one text');
    await saveOk(store, sessionId, 'era-one');
    await replaceAnchor(store, sessionId, 'Era two text');
    await saveOk(store, sessionId, 'era-two');
    await replaceAnchor(store, sessionId, 'Era three text');
    await saveOk(store, sessionId, 'era-three');

    const before = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(before.ok).toBe(true);
    if (!before.ok) return;
    const eraOneOid = before.history[2].oid;

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      commitOid: eraOneOid,
      author: 'Oid Author',
      reason: 'jump back',
    });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    expect(reverted.revertedToCommit).toBe(eraOneOid);
    expect(reverted.commitMessage).toBe(
      `Revert [2 steps] to ${eraOneOid.slice(0, 7)} by Oid Author: jump back`
    );

    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    expect(hist.history[0].opCount).toBe(2);
    expect(hist.history[0].author).toBe('Oid Author');
  });

  it('legacy message-only revert keeps the old subject shape', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'legacy-msg');

    await replaceAnchor(store, sessionId, 'Legacy alpha');
    await saveOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Legacy beta');
    await saveOk(store, sessionId, 'beta');

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      message: 'backup restore',
    });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    expect(reverted.commitMessage).toMatch(/^Revert to [0-9a-f]{7}: backup restore$/);
  });

  it('redo accepts author+reason with sanitizer hygiene', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'redo-hygiene');

    await replaceAnchor(store, sessionId, 'Redo alpha text');
    await saveOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Redo beta text');
    await saveOk(store, sessionId, 'beta');
    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(reverted.ok).toBe(true);

    const before = Date.now();
    const redone = await GIT_HISTORY_HANDLERS.redo_document(store, {
      sessionId,
      author: 'Redo\nAuthor',
      reason: 'line1\nline2',
    });
    const after = Date.now();
    expect(redone.ok).toBe(true);
    if (!redone.ok) return;
    expect(redone.commitMessage).toBe('Redo: line1 line2');

    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    expect(hist.history[0].author).toBe('Redo Author');
    expect(hist.history[0].timestampMs).toBeGreaterThanOrEqual(before);
    expect(hist.history[0].timestampMs).toBeLessThanOrEqual(after);
    expect(hist.history[0].opCount).toBe(1);
    expect(await searchCount(store, sessionId, 'Redo beta text')).toBeGreaterThan(0);
  });

  it('execute_batch honors client op ids and returns them per op', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'batch-ids');
    const anchor = await firstNodeId(store, sessionId);

    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'chained anchor', id: 'audit_chain_anchor' },
      ],
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.createdNodeIds).toContain('audit_chain_anchor');
    expect(res.results).toHaveLength(1);
    expect(res.results[0]).toEqual({ index: 0, action: 'insert', createdNodeIds: ['audit_chain_anchor'] });
    expect(await searchCount(store, sessionId, 'chained anchor')).toBeGreaterThan(0);
  });

  it('footnote id with no created node is never echoed; insert/table/equation/chart ids echo exactly', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'phantom-id');
    const anchor = await firstNodeId(store, sessionId);

    // Footnote honors no id and creates zero AST nodes: the ghost id must
    // appear in neither the flat list nor the per-op attribution.
    const ghost = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'footnote', anchor, text: 'ghost note', id: 'ghost' }],
    });
    expect(ghost.ok).toBe(true);
    if (!ghost.ok) return;
    expect(ghost.createdNodeIds).not.toContain('ghost');
    expect(ghost.results).toHaveLength(1);
    expect(ghost.results[0]).toEqual({ index: 0, action: 'footnote', createdNodeIds: [] });

    // Insert paragraph honors op.id verbatim.
    const para = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'phantom para', id: 'phantom_para_1' },
      ],
    });
    expect(para.ok).toBe(true);
    if (!para.ok) return;
    expect(para.createdNodeIds).toContain('phantom_para_1');
    expect(para.results[0]).toEqual({ index: 0, action: 'insert', createdNodeIds: ['phantom_para_1'] });

    // Insert table honors op.id verbatim.
    const tbl = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        {
          action: 'insert',
          type: 'table',
          anchor,
          position: 'after',
          headers: ['A', 'B'],
          rows: [['1', '2']],
          id: 'phantom_table_1',
        },
      ],
    });
    expect(tbl.ok).toBe(true);
    if (!tbl.ok) return;
    expect(tbl.createdNodeIds).toContain('phantom_table_1');
    expect(tbl.results[0]).toEqual({ index: 0, action: 'insert', createdNodeIds: ['phantom_table_1'] });

    // Equation insert honors op.id verbatim.
    const eq = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'equation', operation: 'insert', anchor, position: 'after', latex: 'x^2', id: 'phantom_eq_1' },
      ],
    });
    expect(eq.ok).toBe(true);
    if (!eq.ok) return;
    expect(eq.createdNodeIds).toContain('phantom_eq_1');
    expect(eq.results[0]).toEqual({ index: 0, action: 'equation', createdNodeIds: ['phantom_eq_1'] });

    // Chart insert honors op.id verbatim.
    const ch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        {
          action: 'chart',
          operation: 'insert',
          anchor,
          chartType: 'column',
          title: 'Phantom chart',
          categories: ['A'],
          series: [{ name: 'S', values: [1] }],
          id: 'phantom_chart_1',
        },
      ],
    });
    expect(ch.ok).toBe(true);
    if (!ch.ok) return;
    expect(ch.createdNodeIds).toContain('phantom_chart_1');
    expect(ch.results[0]).toEqual({ index: 0, action: 'chart', createdNodeIds: ['phantom_chart_1'] });
  });

  it('returned ids chain across batches without re-inspect (custom + generated)', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'batch-chain');
    const anchor = await firstNodeId(store, sessionId);

    // Chain on a client-supplied id.
    const first = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'chain root', id: 'chain_root_1' },
      ],
    });
    expect(first.ok).toBe(true);
    if (!first.ok) return;
    const second = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'chain_root_1', position: 'after', text: 'chain leaf two' },
      ],
    });
    expect(second.ok).toBe(true);
    if (!second.ok) return;
    const genId = second.createdNodeIds[0];
    expect(genId).toMatch(/^_bk_new_/);

    // Chain on the generated id (no inspect between batches).
    const third = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'replace', target: genId, text: 'chain leaf two edited' }],
    });
    expect(third.ok).toBe(true);
    if (!third.ok) return;
    expect(third.createdNodeIds).toEqual([]);
    expect(await searchCount(store, sessionId, 'chain leaf two edited')).toBeGreaterThan(0);
    expect(await searchCount(store, sessionId, 'chain root')).toBeGreaterThan(0);
  });

  it('replace with an id is rejected at the strict schema gate', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'batch-id-reject');
    const anchor = await firstNodeId(store, sessionId);

    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'replace', target: anchor, text: 'x', id: 'bogus' } as never],
    });
    expect(res.ok).toBe(false);
    if (!res.ok) expect(res.code).toBe('E_VALIDATION');
  });

  it('export_audit_trail returns full history with reproducible package hashes', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'audit-trail');

    await replaceAnchor(store, sessionId, 'Trail alpha text');
    await saveOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Trail beta text');
    await saveOk(store, sessionId, 'beta');
    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      author: 'Trail Author',
      reason: 'trail check',
    });
    expect(reverted.ok).toBe(true);

    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId, limit: 100 });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;

    const trail = await handleExportAuditTrail(store, { sessionId });
    expect(trail.ok).toBe(true);
    if (!trail.ok) return;
    expect(trail.totalCommits).toBe(hist.totalCommits);
    expect(trail.currentCommit).toBe(hist.currentCommit);
    expect(trail.returnedCount).toBe(hist.totalCommits);
    expect(trail.truncated).toBe(false);
    expect(trail.trail).toHaveLength(hist.totalCommits);
    // Message/author parity with history reads.
    expect(trail.trail.map((r) => r.oid)).toEqual(hist.history.map((h) => h.oid));
    expect(trail.trail[0].author).toBe('Trail Author');
    // Every package hash reproduces independently via read-only checkout.
    for (const row of trail.trail) {
      const target = await checkoutCommit(file, row.oid);
      const expected = crypto.createHash('sha256').update(Buffer.from(target.xml, 'utf8')).digest('hex');
      expect(row.packageHash).toBe(expected);
      expect(row.packageHash).toMatch(/^[0-9a-f]{64}$/);
      expect(row.resultingHeadOid).toBe(row.oid);
      expect(typeof row.timestampMs).toBe('number');
    }
    // Retention note stays on the handler-only audit export; the registered
    // get_document_history tool description carries no retention note.
    expect(trail.retentionNote).toBe(AUDIT_RETENTION_NOTE);
    expect(trail.retentionNote).toMatch(/deletable/);
    expect(trail.retentionNote).toMatch(/OUT of scope/);
    expect(trail.retentionNote).toMatch(/NOT WORM/);
    expect(EXPORT_AUDIT_TRAIL_DESCRIPTION).toMatch(/HOLD/);
    expect(EXPORT_AUDIT_TRAIL_DESCRIPTION).toMatch(/deletable/);
    expect(GIT_HISTORY_DESCRIPTIONS.get_document_history).not.toMatch(/HOLD/);
  });

  it('preview_revert returns a WOULD-restore summary without mutating anything', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'preview');

    await replaceAnchor(store, sessionId, 'Preview alpha text');
    await saveOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Preview beta text');
    await saveOk(store, sessionId, 'beta');

    const beforeMarkers = markers(store, sessionId);
    const beforeBytes = await fs.promises.readFile(file);
    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;

    const preview = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, steps: 1 });
    expect(preview.ok).toBe(true);
    if (!preview.ok) return;
    expect(preview.targetOid).toBe(hist.history[1].oid);
    expect(preview.targetMessage).toBe(hist.history[1].message);
    expect(preview.stepsBack).toBe(1);
    expect(preview.affectedFiles).toContain('word/document.xml');
    expect(preview.dirty).toBe(false);
    expect(preview.currentHeadOid).toBe(hist.currentCommit);
    expect(preview.totalCommits).toBe(hist.totalCommits);

    // Non-mutation proof: revisions, redo depth, disk bytes, and head all identical.
    expect(markers(store, sessionId)).toEqual(beforeMarkers);
    expect(Buffer.compare(await fs.promises.readFile(file), beforeBytes)).toBe(0);
    const afterHist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(afterHist.ok).toBe(true);
    if (!afterHist.ok) return;
    expect(afterHist.totalCommits).toBe(hist.totalCommits);
    expect(afterHist.currentCommit).toBe(hist.currentCommit);
    expect(await searchCount(store, sessionId, 'Preview beta text')).toBeGreaterThan(0);
  });

  it('preview_revert reports dirty handling without discarding', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store, 'preview-dirty');

    await replaceAnchor(store, sessionId, 'Preview dirty baseline');
    await saveOk(store, sessionId, 'baseline');
    await replaceAnchor(store, sessionId, 'Preview dirty unsaved marker');
    const beforeBytes = await fs.promises.readFile(file);
    const beforeMarkers = markers(store, sessionId);

    const preview = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, steps: 1 });
    expect(preview.ok).toBe(true);
    if (!preview.ok) return;
    expect(preview.dirty).toBe(true);
    expect(preview.dirtyHandling).toMatch(/discard/);
    // Nothing discarded: unsaved text still live, revisions and bytes untouched.
    expect(await searchCount(store, sessionId, 'Preview dirty unsaved marker')).toBeGreaterThan(0);
    expect(markers(store, sessionId)).toEqual(beforeMarkers);
    expect(Buffer.compare(await fs.promises.readFile(file), beforeBytes)).toBe(0);
  });

  it('preview_revert rejects bad shapes with consistent codes', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'preview-bad');
    await replaceAnchor(store, sessionId, 'Preview probe text');
    await saveOk(store, sessionId, 'probe');

    const both = await GIT_HISTORY_HANDLERS.preview_revert(store, {
      sessionId,
      steps: 1,
      commitOid: 'e'.repeat(40),
    });
    expect(both.ok).toBe(false);
    if (!both.ok) expect(both.code).toBe('E_VALIDATION');

    const malformed = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, commitOid: 'not-an-oid' });
    expect(malformed.ok).toBe(false);
    if (!malformed.ok) expect(malformed.code).toBe('E_VALIDATION');

    const unknown = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, commitOid: 'f'.repeat(40) });
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) expect(unknown.code).toBe('E_NOT_FOUND');

    const tooDeep = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, steps: 99 });
    expect(tooDeep.ok).toBe(false);
    if (!tooDeep.ok) expect(tooDeep.code).toBe('E_NOT_FOUND');

    const negative = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, steps: -2 });
    expect(negative.ok).toBe(false);
    if (!negative.ok) expect(negative.code).toBe('E_VALIDATION');

    const noSession = await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId: 'sess_nope' });
    expect(noSession.ok).toBe(false);
    if (!noSession.ok) {
      expect(noSession.code).toBe('E_SESSION');
      expect(noSession.message).toBe('Unknown session "sess_nope" (expired or never opened)');
      expect(noSession.remedy).toBe('Call open_document first, then retry with the returned sessionId.');
    }
  });

  it('breaker set: negative steps, unknown sessions, empty redo, dirty+no-consent', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'breakers');
    await replaceAnchor(store, sessionId, 'Breaker baseline text');
    await saveOk(store, sessionId, 'baseline');

    const negative = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: -1 });
    expect(negative.ok).toBe(false);
    if (!negative.ok) expect(negative.code).toBe('E_VALIDATION');

    const unknownRevert = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId: 'sess_nope', steps: 1 });
    expect(unknownRevert.ok).toBe(false);
    if (!unknownRevert.ok) {
      expect(unknownRevert.code).toBe('E_SESSION');
      expect(unknownRevert.message).toBe('Unknown session "sess_nope" (expired or never opened)');
      expect(unknownRevert.remedy).toBe('Call open_document first, then retry with the returned sessionId.');
    }

    const unknownRedo = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId: 'sess_nope' });
    expect(unknownRedo.ok).toBe(false);
    if (!unknownRedo.ok) {
      expect(unknownRedo.code).toBe('E_SESSION');
      expect(unknownRedo.message).toBe('Unknown session "sess_nope" (expired or never opened)');
      expect(unknownRedo.remedy).toBe('Call open_document first, then retry with the returned sessionId.');
    }

    const emptyRedo = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(emptyRedo.ok).toBe(false);
    if (!emptyRedo.ok) {
      expect(emptyRedo.code).toBe('E_NOT_FOUND');
      expect((emptyRedo as { atLatestState?: boolean }).atLatestState).toBe(true);
    }

    await replaceAnchor(store, sessionId, 'Breaker unsaved text');
    const refused = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      discardUnsaved: false,
    });
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.code).toBe('E_CONFLICT');
    expect(await searchCount(store, sessionId, 'Breaker unsaved text')).toBeGreaterThan(0);
  });

  it('author/reason hygiene matches the existing sanitizers', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store, 'hygiene');

    await replaceAnchor(store, sessionId, 'Hygiene alpha');
    await saveOk(store, sessionId, 'alpha');
    await replaceAnchor(store, sessionId, 'Hygiene beta');
    await saveOk(store, sessionId, 'beta');

    // Newlines collapse to spaces (same sanitizers as the save path).
    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      author: 'Hygiene\nAuthor Name',
      reason: 'First line\nSecond line',
    });
    expect(reverted.ok).toBe(true);
    if (!reverted.ok) return;
    expect(reverted.commitMessage).toBe(
      `Revert [1 steps] to ${reverted.revertedToCommit.slice(0, 7)} by Hygiene Author Name: First line Second line`
    );
    expect(reverted.commitMessage).not.toMatch(/[\r\n]/);
    expect(reverted.commitMessage.length).toBeLessThanOrEqual(500);

    const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(hist.ok).toBe(true);
    if (!hist.ok) return;
    expect(hist.history[0].author).toBe('Hygiene Author Name');
    expect(hist.history[0].message).toBe(reverted.commitMessage);

    // Over the caps fails closed with E_VALIDATION (same as the save schema).
    const longAuthor = await GIT_HISTORY_HANDLERS.revert_document(store, {
      sessionId,
      steps: 1,
      author: 'A'.repeat(101),
    });
    expect(longAuthor.ok).toBe(false);
    if (!longAuthor.ok) expect(longAuthor.code).toBe('E_VALIDATION');

    const longReason = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId, reason: 'R'.repeat(501) });
    expect(longReason.ok).toBe(false);
    if (!longReason.ok) expect(longReason.code).toBe('E_VALIDATION');
  });
});
