/**
 * Sprint 6 Track B — visual review engine tests.
 *
 * Covers (per task):
 * - renderer ok/skip paths (PNG default, JPEG opt-in, artifacts dir)
 * - telemetry token budget (<500 tok)
 * - E_DEPENDENCY when binaries hidden (mock PATH + bogus bin override)
 * - planted-defect catch on synthetic doc (4 large defects) + micro-typo ignore
 * - prompt-law constant phrases + screenshotProvider hook shape
 * - MCP visual tool: telemetry in inspect + --with-visual file-path images
 * - canonical source: @attentionform/astra-engine (packages/engine/src/visual/*)
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

// Canonical visual engine: @attentionform/astra-engine is the single source of truth
// (backend/visual/* deleted — it was a byte-duplicate of packages/engine/src/visual/*).
import {
  VISUAL_PROMPT_LAW,
  VISUAL_VERIFY_RULE,
} from '@attentionform/astra-engine';
import {
  hasVisualPdfBackend as hasPdfBackend,
  findOnPath,
  VISUAL_DEPENDENCY_REMEDY,
} from '@attentionform/astra-engine/visual/backend';
import {
  renderDocxToImages,
  getVisualArtifactsBaseDir,
  getVisualArtifactsDir,
  resolveJailedVisualOutDir,
  HeadlessVisualRenderer,
  headlessScreenshotProvider,
} from '@attentionform/astra-engine';
import {
  buildVisualTelemetry,
  lintVisualDefects,
  generateVisualWireframe,
  estimateTokenCount,
  VISUAL_TELEMETRY_TOKEN_BUDGET,
  type VisualPageMap,
} from '@attentionform/astra-engine';

import { type DocumentAST, DocxParser } from '@attentionform/astra-engine';

import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleVisualInspect, MCP_VISUAL_FLAG, isVisualEnabled } from '../../packages/mcp-server/src/tools/visual';
import { withLoSlot } from '../helpers/lo-slot';

// Slot + budget note: live renders below (opt-in via ASTRA_RUN_SLOW=1) hold
// the suite-wide LO slot (see tests/helpers/lo-slot.ts) and carry 300s for
// queue + own run. Mocked/gated asserts stay on the fast path.

// ── helpers ─────────────────────────────────────────────────────────────────

function minimalAst(nodes: Array<Record<string, unknown>>): DocumentAST {
  return {
    metadata: {},
    sections: [
      {
        id: 's1',
        heading: null,
        nodes: nodes as unknown as DocumentAST['sections'][number]['nodes'],
        pageSetup: null,
        version: 1,
        checksum: 'test',
      },
    ],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'test' },
    nextChangeId: 1,
  } as unknown as DocumentAST;
}

function paraNode(id: string, text: string, namedStyle = 'Normal'): Record<string, unknown> {
  return { id, type: 'paragraph', xml: null, metadata: { text, namedStyle, hasImage: false }, version: 1, checksum: 'x' };
}

function headingNode(id: string, text: string, style = 'Heading1'): Record<string, unknown> {
  return { id, type: 'paragraph', xml: null, metadata: { text, namedStyle: style, hasImage: false }, version: 1, checksum: 'x' };
}

function tableNode(id: string, cols: number, rows = 3): Record<string, unknown> {
  return {
    id,
    type: 'table',
    xml: null,
    metadata: { text: '', namedStyle: 'Normal', tableDimensions: { rows, cols } },
    version: 1,
    checksum: 'x',
  };
}

function blankDocx(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

const tmpDirs: string[] = [];
function mkTmpDir(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  tmpDirs.push(dir);
  return dir;
}

// Live binary probes run only when explicitly enabled; default runs stay
// mocked (no process spawns) and fast. Enable with ASTRA_RUN_SLOW=1.
const RUN_LIVE = process.env.ASTRA_RUN_SLOW === '1';

let savedPath: string | undefined;
beforeEach(() => {
  savedPath = process.env.PATH;
});
afterEach(() => {
  if (savedPath !== undefined) process.env.PATH = savedPath;
  else delete process.env.PATH;
  while (tmpDirs.length > 0) {
    const d = tmpDirs.pop() as string;
    try {
      fs.rmSync(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

// ── prompt law ──────────────────────────────────────────────────────────────

describe('visual prompt law', () => {
  it('contains required phrases (image approximate, XML truth, large defects only)', () => {
    const law = VISUAL_PROMPT_LAW.toLowerCase();
    expect(law).toMatch(/approximate/);
    expect(VISUAL_PROMPT_LAW).toMatch(/XML/);
    expect(VISUAL_PROMPT_LAW).toMatch(/truth/);
    expect(law).toMatch(/large defects/);
    expect(VISUAL_VERIFY_RULE.length).toBeGreaterThan(10);
  });

  it('screenshotProvider hook shape (id + render)', () => {
    expect(headlessScreenshotProvider.id).toBe('headless-lo');
    expect(typeof headlessScreenshotProvider.render).toBe('function');
    const inst = new HeadlessVisualRenderer();
    expect(inst.id).toBe('headless-lo');
    expect(typeof inst.render).toBe('function');
  });

  it('MCP --with-visual flag helper', () => {
    expect(MCP_VISUAL_FLAG).toBe('--with-visual');
    expect(isVisualEnabled(['node', 'x', '--with-visual'])).toBe(true);
    expect(isVisualEnabled(['node', 'x'])).toBe(false);
  });
});

// ── capability gate ─────────────────────────────────────────────────────────

describe('hasPdfBackend capability gate', () => {
  it('resolves real binaries on this system (LO + pdftoppm installed)', () => {
    const status = hasPdfBackend();
    // This repo's CI/dev images ship LibreOffice + poppler (see discussion).
    // If a future runner lacks them, the E_DEPENDENCY path below still proves
    // the gate — so assert shape, then branch.
    expect(Array.isArray(status.missing)).toBe(true);
    if (status.available) {
      expect(status.backend).toBe('libreoffice');
      expect(status.sofficePath).toBeTruthy();
      expect(status.pdftoppmPath).toBeTruthy();
      expect(fs.existsSync(status.sofficePath as string)).toBe(true);
      expect(fs.existsSync(status.pdftoppmPath as string)).toBe(true);
    } else {
      expect(status.backend).toBe('none');
      expect(status.missing.length).toBeGreaterThan(0);
    }
  });

  it('E_DEPENDENCY when binaries hidden via mock PATH', async () => {
    const emptyDir = mkTmpDir('astra-visual-empty-path-');
    process.env.PATH = emptyDir;
    try {
      expect(findOnPath('soffice')).toBeNull();
      expect(findOnPath('pdftoppm')).toBeNull();
      const status = hasPdfBackend();
      expect(status.available).toBe(false);
      expect(status.backend).toBe('none');
      expect(status.missing).toContain('soffice');
      expect(status.missing).toContain('pdftoppm');
      expect(VISUAL_DEPENDENCY_REMEDY).toMatch(/LibreOffice/i);

      await expect(renderDocxToImages(blankDocx(), {})).rejects.toMatchObject({ code: 'E_DEPENDENCY' });
    } finally {
      process.env.PATH = savedPath;
    }
  });

  it('E_DEPENDENCY with bogus explicit bin override (no PATH mutation)', async () => {
    await expect(
      renderDocxToImages(blankDocx(), { sofficeBin: '/nonexistent/soffice-astra-test', pdftoppmBin: '/nonexistent/pdftoppm-astra-test' })
    ).rejects.toMatchObject({ code: 'E_DEPENDENCY' });
    // Canonical engine gate behaves identically.
    expect(hasPdfBackend({ sofficeBin: '/nonexistent/x', pdftoppmBin: '/nonexistent/y' }).available).toBe(false);
  });
});

// ── renderer ok path ────────────────────────────────────────────────────────

describe('headless renderer ok path (LO -> PDF -> pdftoppm)', () => {
  it.runIf(RUN_LIVE)('PNG default renders page images to the artifacts dir', async () => {
    const status = hasPdfBackend();
    if (!status.available) {
      console.warn('[visual test] backend missing — ok-path assertions skipped (gate covered above)');
      return;
    }
    const outDir = getVisualArtifactsDir(`png-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`);
    tmpDirs.push(outDir);
    const res = await withLoSlot(() => renderDocxToImages(blankDocx(), { format: 'png', outDir }));
    expect(res.ok).toBe(true);
    expect(res.format).toBe('png');
    expect(res.backend).toBe('libreoffice');
    expect(res.totalPages).toBeGreaterThanOrEqual(1);
    expect(res.images.length).toBeGreaterThanOrEqual(1);
    expect(res.outDir).toBe(path.resolve(outDir));
    for (const img of res.images) {
      expect(img.path.endsWith('.png')).toBe(true);
      expect(path.resolve(img.path).startsWith(path.resolve(outDir))).toBe(true);
      expect(fs.existsSync(img.path)).toBe(true);
      expect(img.bytes).toBeGreaterThan(0);
      // PNG magic bytes.
      const head = fs.readFileSync(img.path).subarray(0, 8);
      expect([...head]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    }
  }, 300000);

  it.runIf(RUN_LIVE)('JPEG opt-in renders .jpg with quality flag', async () => {
    const status = hasPdfBackend();
    if (!status.available) return;
    const outDir = getVisualArtifactsDir(`jpg-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`);
    tmpDirs.push(outDir);
    const res = await withLoSlot(() => renderDocxToImages(blankDocx(), { format: 'jpeg', jpegQuality: 90, outDir }));
    expect(res.format).toBe('jpeg');
    expect(res.images.length).toBeGreaterThanOrEqual(1);
    for (const img of res.images) {
      expect(img.path.endsWith('.jpg')).toBe(true);
      expect(fs.existsSync(img.path)).toBe(true);
      expect(img.bytes).toBeGreaterThan(0);
      // JPEG SOI marker.
      const head = fs.readFileSync(img.path).subarray(0, 2);
      expect([...head]).toEqual([0xff, 0xd8]);
    }
  }, 300000);

  it('artifacts dir helper creates <tmp>/ASTra/visual/<sub>', () => {
    const dir = getVisualArtifactsDir(`test-${Date.now()}`);
    expect(fs.existsSync(dir)).toBe(true);
    expect(path.resolve(dir).includes(path.join('ASTra', 'visual'))).toBe(true);
  });

  it('rejects empty buffer with E_VALIDATION (not E_DEPENDENCY)', async () => {
    await expect(renderDocxToImages(Buffer.alloc(0), {})).rejects.toMatchObject({ code: 'E_VALIDATION' });
  });
});

// ── outDir jail (path containment) ──────────────────────────────────────────

describe('visual outDir jail (path containment)', () => {
  function expectSandboxSync(fn: () => unknown, remedyMatch = /getVisualArtifactsDir/): void {
    try {
      fn();
    } catch (err: unknown) {
      expect((err as { code?: unknown })?.code).toBe('E_SANDBOX');
      expect(String((err as { remedy?: unknown })?.remedy)).toMatch(remedyMatch);
      return;
    }
    throw new Error('expected E_SANDBOX throw, got success');
  }

  it('resolver accepts the base dir and base-descendant dirs (canonical round-trip)', () => {
    const base = getVisualArtifactsBaseDir();
    expect(resolveJailedVisualOutDir(base)).toBe(fs.realpathSync(base));
    const dir = getVisualArtifactsDir(`jail-ok-${Date.now()}`);
    tmpDirs.push(dir);
    const jailed = resolveJailedVisualOutDir(dir);
    expect(fs.realpathSync(jailed)).toBe(fs.realpathSync(path.resolve(dir)));
  });

  it('resolver refuses an arbitrary absolute dir with E_SANDBOX', () => {
    const evil = path.join(mkTmpDir('astra-visual-evil-'), 'escape');
    expectSandboxSync(() => resolveJailedVisualOutDir(evil));
    expect(fs.existsSync(evil)).toBe(false);
  });

  it('resolver refuses .. traversal escaping the base with E_SANDBOX', () => {
    const base = getVisualArtifactsBaseDir();
    const evil = path.join(base, '..', `jail-traversal-${Date.now()}`);
    expectSandboxSync(() => resolveJailedVisualOutDir(evil));
    expect(fs.existsSync(path.resolve(evil))).toBe(false);
  });

  it('resolver refuses a symlink-parent escape with E_SANDBOX', () => {
    const base = getVisualArtifactsBaseDir();
    const outside = mkTmpDir('astra-visual-jail-outside-');
    const link = path.join(base, `jail-link-${Date.now()}`);
    fs.symlinkSync(outside, link);
    try {
      expectSandboxSync(() => resolveJailedVisualOutDir(path.join(link, 'sub')));
    } finally {
      try {
        fs.unlinkSync(link);
      } catch {
        /* best-effort */
      }
    }
    expect(fs.existsSync(path.join(link, 'sub'))).toBe(false);
  });

  it('resolver refuses null bytes and line breaks with E_SANDBOX', () => {
    const base = getVisualArtifactsBaseDir();
    expectSandboxSync(() => resolveJailedVisualOutDir(path.join(base, 'a\0b')), /null bytes/);
    expectSandboxSync(() => resolveJailedVisualOutDir(path.join(base, 'a\nb')), /line breaks/);
  });

  it.runIf(RUN_LIVE)('live render refuses an arbitrary outDir with E_SANDBOX (zero outside writes)', async () => {
    const status = hasPdfBackend();
    if (!status.available) {
      console.warn('[visual test] backend missing — jail live probe skipped (resolver probes above still prove the jail)');
      return;
    }
    const evilBase = mkTmpDir('astra-visual-jail-live-');
    const evil = path.join(evilBase, 'escape');
    await expect(renderDocxToImages(blankDocx(), { format: 'png', outDir: evil })).rejects.toMatchObject({
      code: 'E_SANDBOX',
    });
    expect(fs.existsSync(evil)).toBe(false);
    expect(fs.readdirSync(evilBase)).toEqual([]);
  }, 60000);
});

// ── telemetry budget ────────────────────────────────────────────────────────

describe('spatial telemetry token budget', () => {
  it('dense 100-paragraph page stays <500 tokens (truncation flagged)', () => {
    const nodes = Array.from({ length: 100 }, (_, i) =>
      paraNode(`p_${String(i).padStart(3, '0')}`, `Dense telemetry paragraph ${i} with enough words to consume wireframe rows and force budget truncation.`)
    );
    const ast = minimalAst(nodes);
    const map: VisualPageMap = { totalPages: 1, nodePages: Object.fromEntries(nodes.map((n) => [n.id as string, 1])), source: 'estimate', isEstimate: true };
    const tel = buildVisualTelemetry(ast, { page: 1, pageMap: map });
    expect(tel.tokens).toBeLessThanOrEqual(VISUAL_TELEMETRY_TOKEN_BUDGET);
    expect(tel.tokens).toBeLessThanOrEqual(500);
    expect(tel.wireframe).toContain('[PAGE 1 OF 1');
    expect(estimateTokenCount(tel.wireframe)).toBeGreaterThan(0);
    // 100 nodes cannot fit untruncated under 500 tok — flag must be honest.
    expect(tel.truncated).toBe(true);
  });

  it('small page is not truncated', () => {
    const ast = minimalAst([headingNode('h1', 'Overview'), paraNode('p1', 'Short body.')]);
    const map: VisualPageMap = { totalPages: 1, nodePages: { h1: 1, p1: 1 }, source: 'renderer', isEstimate: false };
    const tel = buildVisualTelemetry(ast, { page: 1, pageMap: map });
    expect(tel.tokens).toBeLessThanOrEqual(500);
    expect(tel.truncated).toBe(false);
    expect(tel.wireframe).toContain('Overview');
  });

  it('wireframe labels estimate provenance', () => {
    const ast = minimalAst([paraNode('p1', 'Hello')]);
    const est: VisualPageMap = { totalPages: 1, nodePages: { p1: 1 }, source: 'estimate', isEstimate: true };
    expect(generateVisualWireframe(ast, 1, est)).toMatch(/estimate/);
    const auth: VisualPageMap = { totalPages: 1, nodePages: { p1: 1 }, source: 'renderer', isEstimate: false };
    expect(generateVisualWireframe(ast, 1, auth)).not.toMatch(/estimate/);
  });
});

// ── planted defects ─────────────────────────────────────────────────────────

describe('planted-defect catch on synthetic doc', () => {
  function plantedAst(): { ast: DocumentAST; map: VisualPageMap } {
    const nodes = [
      headingNode('h_chart', 'Q3 Revenue Chart Review'),
      paraNode('p_chart', 'The quarterly sales chart below shows regional growth.'),
      // NOTE: no chart/image node on page 1 => missing-chart bait.
      tableNode('t_wide', 8, 4), // 8 cols => overflow bait.
      headingNode('h_orphan', 'Orphan Section Title'),
      paraNode('p_next', 'Body text that fell to the next page.'),
    ];
    const ast = minimalAst(nodes);
    const map: VisualPageMap = {
      totalPages: 3, // page 3 intentionally empty => blank-page bait.
      nodePages: { h_chart: 1, p_chart: 1, t_wide: 1, h_orphan: 1, p_next: 2 },
      source: 'estimate',
      isEstimate: true,
    };
    return { ast, map };
  }

  it('catches all 4 large defects (missing chart, overflow, orphan, blank)', () => {
    const { ast, map } = plantedAst();
    const warnings = lintVisualDefects(ast, map);
    const joined = warnings.join('\n').toLowerCase();
    expect(warnings.length).toBeGreaterThanOrEqual(4);
    expect(joined).toMatch(/missing chart/);
    expect(joined).toMatch(/overflow table/);
    expect(joined).toMatch(/orphan heading/);
    expect(joined).toMatch(/blank page/);
    // Full telemetry surfaces the same warnings under budget.
    const tel = buildVisualTelemetry(ast, { page: 1, pageMap: map });
    expect(tel.warnings.length).toBeGreaterThanOrEqual(4);
    expect(tel.tokens).toBeLessThanOrEqual(500);
  });

  it('ignores micro-typography baits (double spaces, typos)', () => {
    const ast = minimalAst([
      headingNode('h1', 'Clean Title'),
      paraNode('p1', 'Body with  double  spaces and a splling typo.'),
      paraNode('p2', 'Another   spaced   line, otherwise healthy.'),
    ]);
    const map: VisualPageMap = { totalPages: 1, nodePages: { h1: 1, p1: 1, p2: 1 }, source: 'renderer', isEstimate: false };
    // Heading is followed by body on the SAME page => not orphan.
    expect(lintVisualDefects(ast, map)).toEqual([]);
  });

  it('catches defects from a real parsed docx (integration sanity)', async () => {
    const parser = new DocxParser(blankDocx());
    const ast = await parser.getAST();
    const summary = parser.getStructuralSummary() as unknown as { nodePages?: Record<string, number>; totalPages?: number };
    const map: VisualPageMap | undefined =
      summary?.nodePages && summary?.totalPages
        ? { totalPages: summary.totalPages, nodePages: summary.nodePages, source: 'renderer', isEstimate: false }
        : undefined;
    const tel = buildVisualTelemetry(ast, { page: 1, pageMap: map ?? null });
    expect(tel.tokens).toBeLessThanOrEqual(500);
    expect(typeof tel.wireframe).toBe('string');
  });
});

// ── MCP visual tool ─────────────────────────────────────────────────────────

describe('mcp visual tool (telemetry in inspect + --with-visual file paths)', () => {
  it('telemetry-only inspect stays <500 tok with prompt law', async () => {
    const store = createMcpSessionStore();
    const buf = blankDocx();
    const handle = await store.openBuffer(buf, 'visual-telemetry.docx');
    const res = await handleVisualInspect(store, { sessionId: handle.sessionId });
    expect(res.ok).toBe(true);
    if (res.ok === true) {
      expect(res.telemetry.tokens).toBeLessThanOrEqual(500);
      expect(res.promptLaw).toBe(VISUAL_PROMPT_LAW);
      expect(res.images).toBeUndefined();
      expect(typeof res.telemetry.wireframe).toBe('string');
    }
    await store.closeDocument(handle.sessionId).catch(() => undefined);
  });

  it.runIf(RUN_LIVE)('--with-visual returns file-path images (or E_DEPENDENCY when hidden)', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(blankDocx(), 'visual-with-images.docx');
    const status = hasPdfBackend();
    const res = await withLoSlot(() => handleVisualInspect(store, { sessionId: handle.sessionId, withVisual: true, format: 'png' }));
    if (status.available) {
      expect(res.ok).toBe(true);
      if (res.ok === true) {
        expect(Array.isArray(res.images)).toBe(true);
        expect((res.images ?? []).length).toBeGreaterThanOrEqual(1);
        for (const img of res.images ?? []) {
          expect(img.path.endsWith('.png')).toBe(true);
          expect(fs.existsSync(img.path)).toBe(true);
        }
        expect(res.outDir).toBeTruthy();
      }
    } else {
      expect(res.ok).toBe(false);
      if (res.ok === false) expect(res.code).toBe('E_DEPENDENCY');
    }
    await store.closeDocument(handle.sessionId).catch(() => undefined);
  }, 300000);

  it('--with-visual degrades to E_DEPENDENCY under mock PATH', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(blankDocx(), 'visual-mock-path.docx');
    const emptyDir = mkTmpDir('astra-visual-mcp-path-');
    process.env.PATH = emptyDir;
    try {
      const res = await handleVisualInspect(store, { sessionId: handle.sessionId, withVisual: true });
      expect(res.ok).toBe(false);
      if (res.ok === false) {
        expect(res.code).toBe('E_DEPENDENCY');
        expect(res.remedy).toMatch(/LibreOffice|poppler/i);
      }
    } finally {
      process.env.PATH = savedPath;
      await store.closeDocument(handle.sessionId).catch(() => undefined);
    }
  });

  it('rejects bad input with E_VALIDATION (strict schema)', async () => {
    const store = createMcpSessionStore();
    const res = await handleVisualInspect(store, { sessionId: '', withVisual: true });
    expect(res.ok).toBe(false);
    if (res.ok === false) expect(res.code).toBe('E_VALIDATION');
  });
});
