/**
 * D2 fix verification (BRUTAL fix agent):
 * - pie per-slice colors match legend exactly (no series-color smear)
 * - multi-series pie renders explicit placeholder (not silent drop)
 * - skippedIds dedupe, appliedCount only-applied, rollback clears skips
 * - insert/trackedChange loud errors carry value + valid set + hint
 */
import { describe, it, expect, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  convertAstToHtml,
  createBlankDocxBuffer,
  DocxParser,
  executeChart,
  executeDelete,
  executeInsert,
  executeTrackedChange,
  OperationEngine,
} from '@attentionform/astra-engine';
import {
  handleExecuteBatch,
  handleSaveDocument,
} from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';

function meta(text: string, extra: any = {}): any {
  return {
    text, hasCitation: false, hasComment: false, hasBookmark: false,
    hasTrackedChanges: false, hasEquation: false, hasField: false,
    hasImage: false, hasHyperlink: false, ...extra,
  };
}
function node(id: string, type: string, xml: Element | null, metadata: any): any {
  return { id, type, xml, metadata, version: 1, checksum: `ck-${id}` };
}
function chartAst(series: any[], categories = ['A', 'B', 'C']): any {
  return {
    metadata: {}, settings: {}, nextChangeId: 1, styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
    sections: [{
      id: 'sec_1', heading: null, pageSetup: null, version: 1, checksum: 's1',
      nodes: [node('ch_pie', 'chart', null, {
        ...meta(''), title: 'Share', categories, series, chartType: 'pie',
      })],
    }],
  };
}
function sliceFills(html: string): string[] {
  // Slice paths carry stroke="#FFFFFF"; legend rects do not.
  const out: string[] = [];
  const re = /<path[^>]*fill="([^"]+)"[^>]*stroke="#FFFFFF"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html)) !== null) out.push(m[1].toLowerCase());
  return out;
}
function legendFills(html: string): string[] {
  const out: string[] = [];
  const re = /<rect[^>]*fill="([^"]+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html)) !== null) out.push(m[1].toLowerCase());
  return out;
}

describe('D2 pie color: slices use palette, legend matches exactly', () => {
  it('series.color does NOT smear across all slices; legend == slices', () => {
    const html = convertAstToHtml(chartAst(
      [{ name: 'S', values: [30, 30, 40], color: 'FF0000' }]
    ));
    const slices = sliceFills(html);
    const legends = legendFills(html);
    expect(slices.length).toBe(3);
    expect(legends.length).toBe(3);
    // Before fix: slices all #ff0000, legends [#ff0000, palette1, palette2] → diverge.
    // After fix: slices are palette-distinct, legends match slices 1:1.
    expect(new Set(slices).size).toBe(3);
    expect(slices).toEqual(legends);
    expect(slices).not.toEqual(['#ff0000', '#ff0000', '#ff0000']);
  });

  it('single-slice pie may honor series color (both slice + legend agree)', () => {
    const html = convertAstToHtml(chartAst(
      [{ name: 'S', values: [100], color: 'FF0000' }], ['Only']
    ));
    const slices = sliceFills(html);
    const legends = legendFills(html);
    expect(slices.length).toBe(1);
    expect(legends.length).toBe(1);
    expect(slices[0]).toBe('#ff0000');
    expect(legends[0]).toBe('#ff0000');
  });

  it('explicit per-slice colors win and still match legend', () => {
    const html = convertAstToHtml(chartAst(
      [{ name: 'S', values: [10, 20, 30], color: 'FF0000', colors: ['111111', '222222', '333333'] }]
    ));
    const slices = sliceFills(html);
    const legends = legendFills(html);
    expect(slices).toEqual(['#111111', '#222222', '#333333']);
    expect(legends).toEqual(slices);
  });

  it('multi-series pie is explicit, not silent drop', () => {
    const html = convertAstToHtml(chartAst([
      { name: 'S1', values: [30, 70, 10] },
      { name: 'S2', values: [5, 5, 90] },
    ]));
    // First series still renders (pie semantics) + explicit note names series 2.
    expect(html).toContain('Share');
    expect(html).toContain('S1');
    expect(html).toContain('S2');
    expect(html).toMatch(/shows.*only|additional series|2 series/i);
  });
});

describe('D2 skippedIds: dedupe + appliedCount + rollback clear', () => {
  it('same ghost twice in one op reports once', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const res = executeDelete({ action: 'delete', target: ['ghost_x', 'ghost_x'] } as any, ast, parser);
    expect(res.skipped).toEqual(['ghost_x']);
    warn.mockRestore();
  });

  it('engine dedupes across ops and counts only applied', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const realId = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [
        { action: 'delete', target: ['ghost_dup', 'ghost_dup'] } as any,
        { action: 'delete', target: ['ghost_dup'] } as any,
        { action: 'delete', target: [realId, 'ghost_other'] } as any,
      ],
    });
    // ghost_dup once despite 3 mentions; ghost_other once.
    expect(res.skippedIds).toEqual(['ghost_dup', 'ghost_other']);
    // Two all-miss deletes apply nothing; the mixed delete applies → count reflects applies.
    // (all-miss ops contribute 0, mixed contributes 1)
    expect(res.appliedCount).toBe(1);
    warn.mockRestore();
  });

  it('all-miss batch has appliedCount 0', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const res = await engine.executeBatch({
      operations: [{ action: 'delete', target: ['nope_a', 'nope_b'] } as any],
    });
    expect(res.success).toBe(true);
    expect(res.appliedCount).toBe(0);
    expect(res.skippedIds).toEqual(['nope_a', 'nope_b']);
    warn.mockRestore();
  });

  it('skippedIds cleared on rollback', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const res = await engine.executeBatch({
      operations: [
        { action: 'delete', target: ['ghost_before_fail'] } as any,
        { action: 'delete', target: 'ghost_boom', strict: true } as any,
      ],
      sequential: true,
    });
    expect(res.success).toBe(false);
    expect(res.rollbackTriggered).toBe(true);
    expect(res.skippedIds).toEqual([]);
    warn.mockRestore();
  });
});

describe('insert + trackedChange loud errors', () => {
  it('insert unknown type carries value + valid set + hint', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await expect(
      executeInsert({ action: 'insert', type: 'frobnicate' as any, anchor, position: 'after' } as any, ast, parser)
    ).rejects.toThrow(/Unsupported insert type 'frobnicate'/);
    try {
      await executeInsert({ action: 'insert', type: 'frobnicate' as any, anchor, position: 'after' } as any, ast, parser);
      expect.unreachable();
    } catch (e: any) {
      expect(e.message).toContain('Supported:');
      expect(e.message).toContain('inspect_document');
    }
  });

  it('trackedChange unknown op carries value + valid set + hint', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    return parser.getAST().then((ast) => {
      expect(() =>
        executeTrackedChange({ action: 'trackedChange', operation: 'bogus' as any } as any, ast, parser)
      ).toThrow(/Unsupported trackedChange operation: bogus/);
      try {
        executeTrackedChange({ action: 'trackedChange', operation: 'bogus' as any } as any, ast, parser);
        expect.unreachable();
      } catch (e: any) {
        expect(e.message).toContain('Supported:');
        expect(e.message).toContain('acceptAll');
        expect(e.message).toContain('inspect_document');
      }
    });
  });
});

describe('R-1 revisionsPending: batch surfaces pending tracked revisions', () => {
  it('tracked replace batch reports revisionsPending >= 1 with accept remedy (no auto-accept)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    (ast as any).settings.trackChanges = true;
    const engine = new OperationEngine(ast, parser);
    const target = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [{ action: 'replace', target, text: 'Hello R-1' } as any],
    });
    expect(res.success).toBe(true);
    expect(res.appliedCount).toBe(1);
    expect(res.revisionsPending).toBeGreaterThanOrEqual(1);
    expect(res.revisionsRemedy).toContain('acceptAll');
    expect(res.revisionsRemedy).toContain('trackedChange');
    expect(res.revisionsRemedy).toContain('--accept-all-revisions');
    // No auto-accept: the revisions are still in the live DOM.
    const dom = parser.getDocumentDom();
    const live = dom.getElementsByTagName('w:ins').length + dom.getElementsByTagName('w:del').length;
    expect(live).toBeGreaterThanOrEqual(1);
    expect(live).toBe(res.revisionsPending);
    // No default flip: tracking is still on.
    expect((ast.settings as any)?.trackChanges).toBe(true);
  });

  it('clean metadata batch reports revisionsPending 0 with no remedy', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [{ action: 'metadata', operation: 'setProperties', title: 'R-1 clean' } as any],
    });
    expect(res.success).toBe(true);
    expect(res.revisionsPending).toBe(0);
    expect(res.revisionsRemedy).toBeUndefined();
  });

  it('acceptAll clears pending (remedy works, defaults untouched)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    (ast as any).settings.trackChanges = true;
    const engine = new OperationEngine(ast, parser);
    const target = ast.sections[0].nodes[0].id;
    const dirty = await engine.executeBatch({
      operations: [{ action: 'replace', target, text: 'Hello R-1' } as any],
    });
    expect(dirty.revisionsPending).toBeGreaterThanOrEqual(1);
    const cleared = await engine.executeBatch({
      operations: [{ action: 'trackedChange', operation: 'acceptAll' } as any],
    });
    expect(cleared.success).toBe(true);
    expect(cleared.revisionsPending).toBe(0);
    expect(cleared.revisionsRemedy).toBeUndefined();
    expect((ast.settings as any)?.trackChanges).toBe(true);
  });

  it('failed batch reports the rewound live pending count', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [
        { action: 'metadata', operation: 'setProperties', title: 'R-1' } as any,
        { action: 'replace', target: 'ghost_missing_xyz', text: 'boom' } as any,
      ],
      sequential: true,
    });
    expect(res.success).toBe(false);
    expect(res.rollbackTriggered).toBe(true);
    // Nothing applied, nothing tracked: rewound doc holds zero revisions.
    expect(res.revisionsPending).toBe(0);
    expect(res.revisionsRemedy).toBeUndefined();
  });
});

describe('R-1 MCP envelopes carry revisionsPending', () => {
  it('execute_batch envelope carries pending count + remedy on tracked edits', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'r1-tracked.docx');
    const target = await store.withSession(handle.sessionId, async (h: any) => h.ast.sections[0].nodes[0].id);
    const out: any = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'replace', target, text: 'Hello R-1', trackChanges: true }],
    });
    expect(out.ok).toBe(true);
    expect(out.appliedCount).toBe(1);
    expect(out.revisionsPending).toBeGreaterThanOrEqual(1);
    expect(String(out.revisionsRemedy)).toContain('acceptAll');
    expect(String(out.revisionsRemedy)).toContain('--accept-all-revisions');
  });

  it('execute_batch envelope reports 0 with no remedy on clean edits', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'r1-clean.docx');
    const out: any = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'metadata', operation: 'setProperties', title: 'R-1 clean' }],
    });
    expect(out.ok).toBe(true);
    expect(out.revisionsPending).toBe(0);
    expect(out.revisionsRemedy).toBeUndefined();
  });

  it('save_document envelope carries pending count yet still succeeds (advisory only)', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'r1-save.docx');
    const target = await store.withSession(handle.sessionId, async (h: any) => h.ast.sections[0].nodes[0].id);
    const batch: any = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'replace', target, text: 'Hello R-1', trackChanges: true }],
    });
    expect(batch.ok).toBe(true);
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-r1-save-'));
    try {
      const saved: any = await handleSaveDocument(store, {
        sessionId: handle.sessionId,
        outPath: path.join(tmpDir, 'tracked.docx'),
      });
      expect(saved.ok).toBe(true);
      expect(saved.revisionsPending).toBeGreaterThanOrEqual(1);
      expect(String(saved.revisionsRemedy)).toContain('acceptAll');
      expect(String(saved.revisionsRemedy)).toContain('--accept-all-revisions');
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it('save_document envelope reports 0 on a clean document', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'r1-save-clean.docx');
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-r1-save-clean-'));
    try {
      const saved: any = await handleSaveDocument(store, {
        sessionId: handle.sessionId,
        outPath: path.join(tmpDir, 'clean.docx'),
      });
      expect(saved.ok).toBe(true);
      expect(saved.revisionsPending).toBe(0);
      expect(saved.revisionsRemedy).toBeUndefined();
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});

describe('R-14 batch stale refs: engine batch + MCP envelope carry chart advisory', () => {
  const OLD_TITLE = 'H1 Revenue January\u2013June';
  const OLD_CATS = ['January', 'February', 'March', 'April', 'May', 'June'];
  const OLD_VALS = [1200, 2400, 3100, 4200, 5300, 6000];
  const CAPTION = 'Figure 1: H1 Revenue January\u2013June sales trend (see chart).';
  const CTA = 'Act now \u2014 6,000 early-bird bonus ends Friday.';
  const UNRELATED = 'Staffing plan for Q3 covers onboarding and mentoring.';

  async function setupEngineDoc() {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const anchor: string = ast.sections[0].nodes[0].id;
    executeChart(
      {
        action: 'chart',
        chartType: 'column',
        title: OLD_TITLE,
        categories: OLD_CATS,
        series: [{ name: 'Revenue', values: OLD_VALS, color: '0066CC' }],
        anchor,
        id: 'chart-r14-batch',
      } as any,
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: CAPTION, id: 'cap-r14-batch' } as any,
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: CTA, id: 'cta-r14-batch' } as any,
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: UNRELATED, id: 'unrelated-r14-batch' } as any,
      ast,
      parser
    );
    return { parser, ast };
  }

  const UPDATE_OP = {
    action: 'chart',
    operation: 'update',
    target: 'chart-r14-batch',
    title: 'July Revenue Update',
    categories: ['Jul'],
    series: [{ name: 'Revenue', values: [7500], color: '0066CC' }],
  } as any;

  it('engine batch chart update carries stale refs (advisory, body untouched)', async () => {
    const { parser, ast } = await setupEngineDoc();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({ operations: [UPDATE_OP] });
    expect(res.success).toBe(true);
    expect(res.appliedCount).toBe(1);
    expect(Array.isArray((res as any).possiblyStaleBodyRefs)).toBe(true);
    const refs = (res as any).possiblyStaleBodyRefs as Array<{ nodeId: string; preview: string }>;
    const ids = refs.map((r) => r.nodeId);
    expect(ids).toContain('cap-r14-batch');
    expect(ids).toContain('cta-r14-batch');
    expect(ids).not.toContain('unrelated-r14-batch');
    expect(ids).not.toContain('chart-r14-batch');
    for (const r of refs) {
      expect(typeof r.nodeId).toBe('string');
      expect(typeof r.preview).toBe('string');
      expect(r.preview.length).toBeGreaterThan(0);
    }
    const byId: Record<string, string> = {};
    for (const r of refs) byId[r.nodeId] = r.preview;
    expect(byId['cap-r14-batch']).toContain('January');
    expect(byId['cta-r14-batch']).toContain('6,000');
    // Advisory only: body text left intact.
    const texts = new Map(ast.sections.flatMap((s: any) => s.nodes).map((n: any) => [n.id, n.metadata?.text]));
    expect(texts.get('cap-r14-batch')).toBe(CAPTION);
    expect(texts.get('cta-r14-batch')).toBe(CTA);
  });

  it('engine batch returns [] when body has no echoes', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const anchor: string = ast.sections[0].nodes[0].id;
    executeChart(
      {
        action: 'chart',
        chartType: 'column',
        title: 'Alpha Summary',
        categories: ['Alpha', 'Beta'],
        series: [{ name: 'S', values: [150, 250], color: '0066CC' }],
        anchor,
        id: 'chart-clean-batch',
      } as any,
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: UNRELATED, id: 'plain-batch' } as any,
      ast,
      parser
    );
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [
        {
          action: 'chart',
          operation: 'update',
          target: 'chart-clean-batch',
          title: 'Gamma Summary',
          categories: ['Gamma'],
          series: [{ name: 'S', values: [350], color: '0066CC' }],
        } as any,
      ],
    });
    expect(res.success).toBe(true);
    expect((res as any).possiblyStaleBodyRefs).toEqual([]);
  });

  it('engine rollback clears stale refs from the rewound attempt', async () => {
    const { parser, ast } = await setupEngineDoc();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [UPDATE_OP, { action: 'replace', target: 'ghost_missing_xyz', text: 'boom' } as any],
      sequential: true,
    });
    expect(res.success).toBe(false);
    expect(res.rollbackTriggered).toBe(true);
    expect((res as any).possiblyStaleBodyRefs).toEqual([]);
  });

  it('MCP execute_batch envelope carries stale refs end-to-end', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'r14-batch.docx');
    const ids = await store.withSession(handle.sessionId, async (h: any) => {
      const anchor: string = h.ast.sections[0].nodes[0].id;
      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: OLD_TITLE,
          categories: OLD_CATS,
          series: [{ name: 'Revenue', values: OLD_VALS, color: '0066CC' }],
          anchor,
          id: 'chart-r14-mcp',
        } as any,
        h.ast,
        h.parser
      );
      await executeInsert(
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: CAPTION, id: 'cap-r14-mcp' } as any,
        h.ast,
        h.parser
      );
      await executeInsert(
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: CTA, id: 'cta-r14-mcp' } as any,
        h.ast,
        h.parser
      );
      await executeInsert(
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: UNRELATED, id: 'unrelated-r14-mcp' } as any,
        h.ast,
        h.parser
      );
      return { anchor };
    });
    expect(ids).toBeDefined();
    const out: any = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [
        {
          action: 'chart',
          operation: 'update',
          target: 'chart-r14-mcp',
          title: 'July Revenue Update',
          categories: ['Jul'],
          series: [{ name: 'Revenue', values: [7500], color: '0066CC' }],
        },
      ],
    });
    expect(out.ok).toBe(true);
    expect(out.appliedCount).toBe(1);
    expect(Array.isArray(out.possiblyStaleBodyRefs)).toBe(true);
    const mcpIds = (out.possiblyStaleBodyRefs as Array<{ nodeId: string; preview: string }>).map((r) => r.nodeId);
    expect(mcpIds).toContain('cap-r14-mcp');
    expect(mcpIds).toContain('cta-r14-mcp');
    expect(mcpIds).not.toContain('unrelated-r14-mcp');
    for (const r of out.possiblyStaleBodyRefs) {
      expect(typeof r.nodeId).toBe('string');
      expect(typeof r.preview).toBe('string');
    }
  });
});
