import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import {
  InspectDocumentSchema,
  LEAN6_HANDLERS,
  collectTableNodeIdsInOrder,
  formatValidationLocation,
  remedyForValidationError,
} from '../../src/lib/mcp-tool-adapter';
import { handleExportDocument } from '../../src/tools/export';
import { handleVisualInspect } from '../../src/tools/visual';

function blankBuf(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

/** Mirrors ToonGenerator.estimateTokens (chars/3.8 heuristic). */
function estimateTokens(text: string): number {
  return Math.ceil(text.length / 3.8);
}

/** First whitespace-delimited token of every TOON data row (skips header lines). */
function rowIds(toonText: string): string[] {
  return toonText
    .split('\n')
    .map((l) => l.split(/\s+/)[0])
    .filter((id) => !!id && !['>', 'ID'].includes(id) && !/^---/.test(id));
}

const tmpDocxFiles: string[] = [];

function writeTmpDocx(prefix: string): string {
  const buf = blankBuf();
  const file = path.join(os.tmpdir(), `astra-phase-c-${prefix}-${process.pid}-${Date.now()}-${tmpDocxFiles.length}.docx`);
  fs.writeFileSync(file, buf);
  tmpDocxFiles.push(file);
  return file;
}

afterEach(() => {
  while (tmpDocxFiles.length > 0) {
    const f = tmpDocxFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
});

async function firstNodeId(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): Promise<string> {
  const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId, offset: 0, limit: 1 });
  expect(first.ok).toBe(true);
  if (!first.ok) throw new Error('inspect failed in test setup');
  expect(first.returnedNodes).toBe(1);
  const id = rowIds(first.toonText)[0];
  expect(typeof id).toBe('string');
  return id;
}

describe('phase C MCP hardening', () => {
  it('C-1a: export failures use message (never error)', async () => {
    const store = createMcpSessionStore();
    const missing = await handleExportDocument(store, { sessionId: 'sess_nope', format: 'markdown' });
    expect(missing.ok).toBe(false);
    expect((missing as Record<string, unknown>).error).toBeUndefined();
    expect(typeof (missing as Record<string, unknown>).message).toBe('string');

    const badArgs = await handleExportDocument(store, { sessionId: 'x', format: 'bogus' });
    expect(badArgs.ok).toBe(false);
    expect((badArgs as Record<string, unknown>).error).toBeUndefined();
    expect(typeof (badArgs as Record<string, unknown>).message).toBe('string');
  });

  it('C-1b: dirty/revision on export + visual; save clears dirty', async () => {
    const store = createMcpSessionStore();
    const out = path.join(os.tmpdir(), `astra-phase-c-dirty-${process.pid}-${Date.now()}.docx`);
    tmpDocxFiles.push(out);
    const h = await store.openBuffer(blankBuf(), 'dirty.docx');
    const fresh = await handleExportDocument(store, { sessionId: h.sessionId, format: 'markdown' });
    expect(fresh.ok).toBe(true);
    expect(fresh.dirty).toBe(false);
    expect(fresh.revision).toBe(1);

    const insp = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId });
    expect(insp.ok).toBe(true);
    if (!insp.ok) return;
    expect(insp.dirty).toBe(false);
    expect(insp.revision).toBe(1);
    expect(insp.totalNodes).toBeGreaterThanOrEqual(1);

    const anchorId = await firstNodeId(store, h.sessionId);
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: h.sessionId,
      operations: [{ action: 'replace', target: anchorId, text: 'Edited unsaved content' }],
    });
    expect(batch.ok).toBe(true);

    const unsaved = await handleExportDocument(store, { sessionId: h.sessionId, format: 'markdown' });
    expect(unsaved.ok).toBe(true);
    expect(unsaved.dirty).toBe(true);
    expect(unsaved.revision).toBe(2);
    expect(typeof unsaved.content).toBe('string');
    expect(unsaved.content).toContain('Edited unsaved content');

    const vis = await handleVisualInspect(store, { sessionId: h.sessionId });
    expect(vis.ok).toBe(true);
    if (vis.ok) {
      expect(vis.dirty).toBe(true);
      expect(vis.revision).toBe(2);
    }

    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: h.sessionId, outPath: out });
    expect(saved.ok).toBe(true);

    const afterSave = await handleExportDocument(store, { sessionId: h.sessionId, format: 'markdown' });
    expect(afterSave.ok).toBe(true);
    expect(afterSave.dirty).toBe(false);
    expect(afterSave.revision).toBe(2);
  });

  it('C-1c: save remedies are xpath-specific', () => {
    expect(remedyForValidationError('/w:document/w:body/w:p/w:r/wp:docPr@id', 'x')).toMatch(/docPr/);
    expect(remedyForValidationError('/w:document/w:body/w:tbl/w:tr/w:tc', 'x')).toMatch(/tblGrid|Table/);
    expect(remedyForValidationError('/w:document', 'relationship rId9 target missing')).toMatch(/rId|relationships/i);
    expect(remedyForValidationError('/w:document/w:body/w:p', 'some generic schema complaint')).toMatch(
      /hex\/half-point\/twips/
    );
  });

  it('save remedies name the header/footer part and tracking-attribute leaks', () => {
    expect(
      remedyForValidationError('/w:ftr[1]/w:p[1]', "The 'data-astra-id' attribute is not declared.")
    ).toMatch(/footer part/);
    expect(
      remedyForValidationError('/w:ftr[1]/w:p[1]', "The 'data-astra-id' attribute is not declared.")
    ).toMatch(/insertPageNumber/);
    expect(
      remedyForValidationError('/w:hdr[1]/w:p[2]', "The 'data-astra-id' attribute is not declared.")
    ).toMatch(/header part/);
    expect(remedyForValidationError('/w:hdr[1]/w:p[1]', 'unexpected child element')).toMatch(/header/);
    expect(remedyForValidationError('/w:ftr[1]/w:p[1]', 'unexpected child element')).toMatch(/footer/);
    expect(remedyForValidationError('/w:document/w:body/w:p', 'some generic schema complaint')).toMatch(
      /hex\/half-point\/twips/
    );
  });

  it('C-2: 1500-token budget slice on a large doc stays within budget', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'large.docx');
    const anchorId = await firstNodeId(store, h.sessionId);

    const ops = Array.from({ length: 250 }, (_, i) => ({
      action: 'insert' as const,
      type: 'paragraph' as const,
      anchor: anchorId,
      position: 'after' as const,
      text: `Synthetic window probe paragraph ${String(i).padStart(3, '0')} carrying filler words to pad every TOON row well beyond the compact budget`,
    }));
    const batch = await LEAN6_HANDLERS.execute_batch(store, { sessionId: h.sessionId, operations: ops });
    expect(batch.ok).toBe(true);

    // Default inspect on a large doc: first window, not a full dump.
    const def = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId });
    expect(def.ok).toBe(true);
    if (!def.ok) return;
    expect(def.totalNodes).toBeGreaterThanOrEqual(200);
    expect(def.returnedNodes).toBeLessThanOrEqual(50);
    expect(def.hasMore).toBe(true);
    expect(def.nextOffset).toBe(def.returnedNodes);
    expect(typeof def.remedy).toBe('string');

    // Explicit 1500-token budget slice.
    const budgeted = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      maxTokens: 1500,
    });
    expect(budgeted.ok).toBe(true);
    if (!budgeted.ok) return;
    expect(estimateTokens(budgeted.toonText)).toBeLessThanOrEqual(1600);
    expect(budgeted.truncated).toBe(true);
    expect(budgeted.hasMore).toBe(true);
  }, 120000);

  it('C-2: aroundNode returns the exact window; offset paging is stable', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'window.docx');
    const anchorId = await firstNodeId(store, h.sessionId);
    const ops = Array.from({ length: 30 }, (_, i) => ({
      action: 'insert' as const,
      type: 'paragraph' as const,
      anchor: anchorId,
      position: 'after' as const,
      text: `Window stability paragraph ${String(i).padStart(2, '0')}`,
    }));
    const batch = await LEAN6_HANDLERS.execute_batch(store, { sessionId: h.sessionId, operations: ops });
    expect(batch.ok).toBe(true);

    // aroundNode exact window [node-w, node+w].
    const all = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId, offset: 0, limit: 200 });
    expect(all.ok).toBe(true);
    if (!all.ok) return;
    const ids = rowIds(all.toonText);
    expect(ids.length).toBe(all.totalNodes);
    const midIdx = Math.floor(ids.length / 2);
    const midId = ids[midIdx];
    const zoom = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      aroundNode: midId,
      window: 2,
    });
    expect(zoom.ok).toBe(true);
    if (!zoom.ok) return;
    expect(rowIds(zoom.toonText)).toEqual(ids.slice(midIdx - 2, midIdx + 3));

    // Offset paging partitions the rows with no overlap or gaps.
    const p1 = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId, offset: 0, limit: 10 });
    const p2 = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId, offset: 10, limit: 10 });
    const both = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId, offset: 0, limit: 20 });
    expect(p1.ok && p2.ok && both.ok).toBe(true);
    if (!p1.ok || !p2.ok || !both.ok) return;
    const i1 = rowIds(p1.toonText);
    const i2 = rowIds(p2.toonText);
    expect(new Set([...i1, ...i2]).size).toBe(i1.length + i2.length);
    expect([...i1, ...i2]).toEqual(rowIds(both.toonText));
    expect(p1.nextOffset).toBe(10);
    // Unknown aroundNode fails closed with a search-first remedy.
    const badZoom = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      aroundNode: 'p_does_not_exist',
    });
    expect(badZoom.ok).toBe(false);
    if (!badZoom.ok) {
      expect(badZoom.code).toBe('E_VALIDATION');
      expect(badZoom.remedy).toMatch(/search_content/);
    }
  }, 120000);

  it('C-3: regex:true works; regex:false stays literal with a hint', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'search.docx');
    const anchorId = await firstNodeId(store, h.sessionId);
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: h.sessionId,
      operations: [
        { action: 'insert' as const, type: 'paragraph' as const, anchor: anchorId, position: 'after' as const, text: 'Alpha red team milestone' },
        { action: 'insert' as const, type: 'paragraph' as const, anchor: anchorId, position: 'after' as const, text: 'Beta blue team milestone' },
        { action: 'insert' as const, type: 'paragraph' as const, anchor: anchorId, position: 'after' as const, text: 'Gamma green team retrospective' },
      ],
    });
    expect(batch.ok).toBe(true);

    // Literal default: metacharacters are NOT interpreted + hint points at regex:true.
    const literal = await LEAN6_HANDLERS.search_content(store, { sessionId: h.sessionId, query: 'Alpha|Beta' });
    expect(literal.ok).toBe(true);
    if (!literal.ok) return;
    expect(literal.matches).toEqual([]);
    expect(typeof literal.hint).toBe('string');
    expect(literal.hint as string).toMatch(/regex/);

    // Plain literal without metacharacters: no hint, current behavior.
    const plain = await LEAN6_HANDLERS.search_content(store, { sessionId: h.sessionId, query: 'milestone' });
    expect(plain.ok).toBe(true);
    if (!plain.ok) return;
    expect(plain.matches.length).toBe(2);
    expect((plain as { hint?: string }).hint).toBeUndefined();

    // Opt-in regex interprets the alternation.
    const re = await LEAN6_HANDLERS.search_content(store, {
      sessionId: h.sessionId,
      query: 'Alpha|Beta',
      regex: true,
    });
    expect(re.ok).toBe(true);
    if (!re.ok) return;
    expect(re.matches.length).toBe(2);

    // Invalid regex fails closed; over-long regex patterns are rejected.
    // Both use generic messages/remedies (no pattern echo, no hit hint).
    const badRe = await LEAN6_HANDLERS.search_content(store, { sessionId: h.sessionId, query: '(', regex: true });
    expect(badRe.ok).toBe(false);
    if (!badRe.ok) {
      expect(badRe.code).toBe('E_VALIDATION');
      expect(badRe.message).not.toContain('(');
      expect(badRe.remedy).toMatch(/omit regex:true/);
    }
    const longRe = await LEAN6_HANDLERS.search_content(store, {
      sessionId: h.sessionId,
      query: `ab(${'x|y'.repeat(80)})`,
      regex: true,
    });
    expect(longRe.ok).toBe(false);
    if (!longRe.ok) {
      expect(longRe.code).toBe('E_VALIDATION');
      expect(longRe.message).not.toContain('x|y');
      expect(longRe.remedy).toMatch(/omit regex:true/);
    }
  });

  it('C-4: overflow open reports evicted[]; evicted ids get a capacity E_SESSION remedy', async () => {
    const store = createMcpSessionStore({ maxDocs: 2, allowClientSessionIds: true });
    await store.openBuffer(blankBuf(), 'a.docx', 'sess_a');
    await store.openBuffer(blankBuf(), 'b.docx', 'sess_b');
    const cPath = writeTmpDocx('c2');
    const res = await LEAN6_HANDLERS.open_document(store, { path: cPath, sessionId: 'sess_c' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.evicted).toContain('sess_a');

    const gone = await LEAN6_HANDLERS.inspect_document(store, { sessionId: 'sess_a' });
    expect(gone.ok).toBe(false);
    if (!gone.ok) {
      expect(gone.code).toBe('E_SESSION');
      expect(gone.remedy).toMatch(/2-doc cap/);
      expect(gone.remedy).toMatch(/close_document/);
    }
  });

  it('C-4: open responses always carry evicted[] (empty when nothing evicted)', async () => {
    const store = createMcpSessionStore();
    const p = writeTmpDocx('solo');
    const res = await LEAN6_HANDLERS.open_document(store, { path: p });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.evicted).toEqual([]);
  });

  it('C-2 loop-back: single-row over-budget keeps the row and flags truncated', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'single-row-budget.docx');
    const anchorId = await firstNodeId(store, h.sessionId);
    const longText = `Over-budget probe ${'filler words '.repeat(200)}`.slice(0, 2600);
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: h.sessionId,
      operations: [
        { action: 'insert' as const, type: 'paragraph' as const, anchor: anchorId, position: 'after' as const, text: longText },
      ],
    });
    expect(batch.ok).toBe(true);
    // The probe row sits at index 1 (inserted right after the anchor row).
    const res = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      offset: 1,
      limit: 1,
      maxTokens: 200,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.returnedNodes).toBe(1);
    expect(res.truncated).toBe(true);
  });

  it('C-2 loop-back: small-doc envelope limit stays within 1-200', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'small.docx');
    const res = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.limit).toBeGreaterThanOrEqual(1);
    expect(res.limit).toBeLessThanOrEqual(200);
    expect(res.limit).toBe(Math.min(res.totalNodes, 200));
  });

  it('C-2 loop-back: aroundNode with explicit non-default offset/limit is rejected E_VALIDATION', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'mutual.docx');
    const anchorId = await firstNodeId(store, h.sessionId);
    // Non-default paging values alongside aroundNode are genuinely ambiguous
    // (the zoom would silently discard them) -> fail closed.
    const withOffset = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      aroundNode: anchorId,
      offset: 5,
    });
    expect(withOffset.ok).toBe(false);
    if (!withOffset.ok) {
      expect(withOffset.code).toBe('E_VALIDATION');
      expect(withOffset.message).toMatch(/mutually exclusive/);
    }
    const withLimit = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      aroundNode: anchorId,
      limit: 10,
    });
    expect(withLimit.ok).toBe(false);
    if (!withLimit.ok) expect(withLimit.code).toBe('E_VALIDATION');
    // aroundNode alone (with window) still zooms.
    const zoom = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      aroundNode: anchorId,
      window: 2,
    });
    expect(zoom.ok).toBe(true);
  });

  it('C-2 wire loop-back: SDK-parsed shape (defaults injected) still zooms; small-doc envelope exact', async () => {
    // The old wire schema carried .default(0/50/5), so McpServer.registerTool
    // zod-parsed BEFORE the handler and every real-wire aroundNode-only call
    // arrived as {aroundNode, offset:0, limit:50, window:5} — rejected as
    // "mutually exclusive" while unit tests (direct handler calls, no SDK)
    // stayed green. The wire schema no longer carries defaults, AND the
    // handler treats default-valued keys as implicit, so this exact shape
    // must zoom. This is the wire-breaker regression shape.
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'wire.docx');
    const anchorId = await firstNodeId(store, h.sessionId);
    const ops = Array.from({ length: 30 }, (_, i) => ({
      action: 'insert' as const,
      type: 'paragraph' as const,
      anchor: anchorId,
      position: 'after' as const,
      text: `Wire probe paragraph ${String(i).padStart(2, '0')}`,
    }));
    const batch = await LEAN6_HANDLERS.execute_batch(store, { sessionId: h.sessionId, operations: ops });
    expect(batch.ok).toBe(true);

    // Lock the wire fix: the schema itself must not inject defaults.
    const wireParsed = InspectDocumentSchema.safeParse({ sessionId: h.sessionId });
    expect(wireParsed.success).toBe(true);
    if (wireParsed.success) {
      expect('offset' in wireParsed.data).toBe(false);
      expect('limit' in wireParsed.data).toBe(false);
      expect('window' in wireParsed.data).toBe(false);
    }

    const all = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId, offset: 0, limit: 200 });
    expect(all.ok).toBe(true);
    if (!all.ok) return;
    const ids = rowIds(all.toonText);
    const midIdx = Math.floor(ids.length / 2);
    const midId = ids[midIdx];

    // SDK-shape aroundNode zoom (defaults injected) succeeds with the exact window.
    const sdkZoom = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      aroundNode: midId,
      offset: 0,
      limit: 50,
      window: 5,
    });
    expect(sdkZoom.ok).toBe(true);
    if (!sdkZoom.ok) return;
    expect(rowIds(sdkZoom.toonText)).toEqual(ids.slice(midIdx - 5, midIdx + 6));

    // SDK-shape unknown aroundNode reaches the lookup (search_content remedy),
    // not the mutual-exclusion error.
    const sdkBadZoom = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: h.sessionId,
      aroundNode: 'p_does_not_exist',
      offset: 0,
      limit: 50,
      window: 5,
    });
    expect(sdkBadZoom.ok).toBe(false);
    if (!sdkBadZoom.ok) {
      expect(sdkBadZoom.code).toBe('E_VALIDATION');
      expect(sdkBadZoom.remedy).toMatch(/search_content/);
      expect(sdkBadZoom.message).not.toMatch(/mutually exclusive/);
    }

    // Small-doc envelope over the SDK shape: all-default windowing keys count
    // as implicit, so a small doc still returns limit == min(total,200).
    const small = await store.openBuffer(blankBuf(), 'wire-small.docx');
    const sdkSmall = await LEAN6_HANDLERS.inspect_document(store, {
      sessionId: small.sessionId,
      offset: 0,
      limit: 50,
      window: 5,
    });
    expect(sdkSmall.ok).toBe(true);
    if (!sdkSmall.ok) return;
    expect(sdkSmall.limit).toBe(Math.min(sdkSmall.totalNodes, 200));
    expect(sdkSmall.hasMore).toBe(false);
  }, 120000);

  it('C-1c loop-back: bare w:tbl xpath maps to the table remedy', () => {
    expect(remedyForValidationError('w:tbl', 'x')).toMatch(/tblGrid|Table/);
  });
});

describe('save-blocked validation location', () => {
  it('indexed table xpath resolves to table ordinal plus node id', () => {
    const loc = formatValidationLocation('/w:document[1]/w:body[1]/w:tbl[2]/w:tr[1]/w:tc[1]', ['t_001', 't_002']);
    expect(loc).toContain('/w:document[1]/w:body[1]/w:tbl[2]');
    expect(loc).toMatch(/table 2/);
    expect(loc).toContain('t_002');
  });

  it('table xpath without a resolvable node still names the table', () => {
    const loc = formatValidationLocation('/w:document[1]/w:body[1]/w:tbl[3]', ['t_001']);
    expect(loc).toContain('w:tbl');
    expect(loc).toMatch(/table 3/);
  });

  it('missing xpath fails honest with an explicit location fallback', () => {
    const loc = formatValidationLocation(undefined, ['t_001']);
    expect(loc).toMatch(/no xPath/);
    expect(loc).toContain('t_001');
    const empty = formatValidationLocation(undefined, []);
    expect(empty).toMatch(/no xPath/);
  });

  it('collects table ids in document order', () => {
    const ast = { sections: [{ nodes: [{ id: 'A1B2C3D4', type: 'paragraph' }, { id: 't_002', type: 'table' }, { id: 't_001', type: 'table' }] }] };
    expect(collectTableNodeIdsInOrder(ast)).toEqual(['t_002', 't_001']);
    expect(collectTableNodeIdsInOrder(null)).toEqual([]);
  });

  it('crafted invalid table blocks save with the tbl index named', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'invalid-table.docx');
    const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId: h.sessionId, offset: 0, limit: 1 });
    expect(first.ok).toBe(true);
    if (!first.ok) return;
    const anchorId = rowIds(first.toonText)[0];
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: h.sessionId,
      operations: [
        { action: 'insert', type: 'table', anchor: anchorId, position: 'after', headers: ['A', 'B'], rows: [['1', '2']] } as never,
      ],
    });
    expect(batch.ok).toBe(true);
    if (!batch.ok) return;
    const tableId = (batch as { createdNodeIds: string[] }).createdNodeIds[0];
    await store.withSession(h.sessionId, async (handle) => {
      const dom = handle.parser.getDocumentDom();
      const tbl = dom.getElementsByTagName('w:tbl')[0] as unknown as Element;
      tbl.appendChild(dom.createElement('w:bogus'));
      return null as never;
    });
    const outPath = path.join(os.tmpdir(), `astra-save-blocked-${process.pid}-${Date.now()}.docx`);
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: h.sessionId, outPath });
    expect(saved.ok).toBe(false);
    if (saved.ok) return;
    expect(saved.code).toBe('E_VALIDATION');
    expect(saved.message).toContain('w:tbl');
    expect(saved.message).toMatch(/table 1/);
    expect(saved.message).toContain(tableId);
    expect(saved.remedy).toMatch(/tblGrid|Table/);
    expect(fs.existsSync(outPath)).toBe(false);
  }, 120000);
});
