/**
 * Session style guide attachment: attach-at-creation (new_document /
 * compose_document), per-batch override (execute_batch), persistence,
 * inspect summary, save warnings, and fail-closed resolution.
 *
 * Attachment model under test: the guide attaches at session creation and
 * persists for the session; a per-batch styleGuide overrides for that batch
 * only; omitting it inherits the session guide. No detach API.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser, OperationEngine, createBlankDocxBuffer, resolveStyleGuide } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNewDocument } from '../../src/tools/new-document';
import { handleComposeDocument } from '../../src/tools/compose-document';

const SLATE_GUIDE = [
  '# Style Guide: Probe Slate',
  '',
  'Slate house style for session attachment probes.',
  '',
  '## Palette',
  '- Primary: #1E3A8A',
  '- Chart Palette: #1E3A8A, #0284C7, #10B981, #F59E0B',
  '',
  '## Tables',
  '- Header Fill: #1E3A8A',
  '- Header Text: #FFFFFF',
  '- Keep Together: true',
  '',
  '## Typography',
  '- Font Family: Calibri',
  '',
  '## Callouts',
  '- Default Theme: slate',
  '',
  '## Code Blocks',
  '- Background: #F1F5F9',
  '- Border Color: #CBD5E1',
  '',
  '## Layout',
  '- Margin: standard',
  '',
].join('\n');

const EMERALD_GUIDE = SLATE_GUIDE.replace('Probe Slate', 'Probe Emerald').replace(
  '- Default Theme: slate',
  '- Default Theme: emerald'
);

const BADHEX_GUIDE = SLATE_GUIDE.replace('Probe Slate', 'Probe Badhex').replace(
  '- Primary: #1E3A8A',
  '- Primary: #ZZZZZZ'
);

let scratchDir: string;
const extraDirs: string[] = [];

beforeEach(() => {
  scratchDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-guide-session-'));
  extraDirs.push(scratchDir);
});

afterEach(() => {
  while (extraDirs.length > 0) {
    const dir = extraDirs.pop() as string;
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

function writeGuide(name: string, markdown: string): string {
  const file = path.join(scratchDir, name);
  fs.writeFileSync(file, markdown);
  return file;
}

function documentXmlOf(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): string {
  const handle = store.get(sessionId);
  const buf = handle.parser.rezip();
  return new AdmZip(buf).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

async function newGuidedSession(
  store: ReturnType<typeof createMcpSessionStore>,
  guideRef: string,
  tag: string
): Promise<string> {
  const outPath = path.join(scratchDir, `${tag}.docx`);
  const res = await handleNewDocument(store, { outPath, template: 'blank', styleGuide: guideRef });
  expect(res.ok).toBe(true);
  if (!res.ok) throw new Error(`guided new_document failed: ${JSON.stringify(res)}`);
  return res.sessionId;
}

describe('new_document guide attach-at-creation', () => {
  it('attaches a file guide, persists it, and reports summary + diagnostics', async () => {
    const store = createMcpSessionStore();
    const guideFile = writeGuide('slate.md', SLATE_GUIDE);
    const outPath = path.join(scratchDir, 'guided.docx');
    const res = await handleNewDocument(store, {
      outPath,
      template: 'blank',
      styleGuide: './slate.md',
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`new_document failed: ${JSON.stringify(res)}`);
    expect(res.styleGuide?.name).toBe('Probe Slate');
    expect(res.styleGuide?.headerFill).toBe('1E3A8A');
    expect(res.styleGuide?.headerText).toBe('FFFFFF');
    expect(res.styleGuide?.calloutTheme).toBe('slate');
    expect(res.styleGuide?.chartPalette).toEqual(['1E3A8A', '0284C7', '10B981', 'F59E0B']);
    expect(Array.isArray(res.styleGuideWarnings)).toBe(true);
    // Session persistence: the live handle carries the guide on the engine.
    const handle = store.get(res.sessionId);
    expect(handle.styleGuide?.callouts?.defaultTheme).toBe('slate');
    expect(handle.engine.getStyleGuide()?.callouts?.defaultTheme).toBe('slate');
    expect(handle.styleGuideRef).toBe('./slate.md');
    expect(guideFile.endsWith('slate.md')).toBe(true);
  });

  it('fails closed on unknown names, naming the available set', async () => {
    const store = createMcpSessionStore();
    const outPath = path.join(scratchDir, 'unknown.docx');
    const res = await handleNewDocument(store, { outPath, styleGuide: 'definitely-not-a-guide' });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('expected failure');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toMatch(/unknown style guide|available|no bundled guides/);
    expect(fs.existsSync(outPath)).toBe(false);
  });

  it('refuses guide files outside the document folder jail', async () => {
    const store = createMcpSessionStore();
    const elsewhere = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-guide-elsewhere-'));
    extraDirs.push(elsewhere);
    const outside = path.join(elsewhere, 'outside.md');
    fs.writeFileSync(outside, SLATE_GUIDE);
    const outPath = path.join(scratchDir, 'jailed.docx');
    const res = await handleNewDocument(store, { outPath, styleGuide: outside });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('expected failure');
    expect(res.code).toBe('E_SANDBOX');
    expect(res.message).toMatch(/outside the allowed directory/);
    expect(fs.existsSync(outPath)).toBe(false);
  });

  it('fails closed on remote guide URLs', async () => {
    const store = createMcpSessionStore();
    const outPath = path.join(scratchDir, 'remote.docx');
    const res = await handleNewDocument(store, {
      outPath,
      styleGuide: 'https://example.com/brand.md',
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('expected failure');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toMatch(/remote URL/);
    expect(fs.existsSync(outPath)).toBe(false);
  });
});

describe('session inheritance and per-batch override', () => {
  it('batches inherit the session guide; an override applies to that batch only', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    writeGuide('emerald.md', EMERALD_GUIDE);
    const sessionId = await newGuidedSession(store, './slate.md', 'inherit');

    // Inherited: no theme passed, guide default (slate border 475569) wins.
    const first = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'House note' }],
    });
    expect(first.ok).toBe(true);
    if (!first.ok) throw new Error(`batch failed: ${JSON.stringify(first)}`);
    expect(documentXmlOf(store, sessionId)).toContain('w:color="475569"');

    // Override: emerald guide for this batch only (border 059669).
    const second = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      styleGuide: './emerald.md',
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Override note' }],
    });
    expect(second.ok).toBe(true);
    if (!second.ok) throw new Error(`override batch failed: ${JSON.stringify(second)}`);
    expect(documentXmlOf(store, sessionId)).toContain('w:color="059669"');

    // Session guide survives the override: engine + handle still slate.
    const handle = store.get(sessionId);
    expect(handle.styleGuide?.callouts?.defaultTheme).toBe('slate');
    expect(handle.engine.getStyleGuide()?.callouts?.defaultTheme).toBe('slate');

    // Next batch without override inherits slate again.
    const third = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Back to house' }],
    });
    expect(third.ok).toBe(true);
    if (!third.ok) throw new Error(`third batch failed: ${JSON.stringify(third)}`);
    const xml = documentXmlOf(store, sessionId);
    expect(xml).toContain('w:color="475569"');
  });

  it('explicit params still beat the inherited guide (no silent override)', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    const sessionId = await newGuidedSession(store, './slate.md', 'explicit');
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Rose pick', theme: 'rose' }],
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`batch failed: ${JSON.stringify(res)}`);
    // Rose border wins over the slate guide default.
    expect(documentXmlOf(store, sessionId)).toContain('w:color="E11D48"');
  });

  it('batch-level unknown names fail closed without touching the session guide', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    const sessionId = await newGuidedSession(store, './slate.md', 'batch-unknown');
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      styleGuide: 'definitely-not-a-guide',
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Never applied' }],
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('expected failure');
    expect(res.code).toBe('E_VALIDATION');
    expect(store.get(sessionId).styleGuide?.callouts?.defaultTheme).toBe('slate');
  });
});

describe('inspect summary and save warnings', () => {
  it('inspect carries the compact guide summary when active, absent otherwise', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    const guided = await newGuidedSession(store, './slate.md', 'inspect-guided');
    const seen = await LEAN6_HANDLERS.inspect_document(store, { sessionId: guided });
    expect(seen.ok).toBe(true);
    if (!seen.ok) throw new Error(`inspect failed: ${JSON.stringify(seen)}`);
    expect(seen.styleGuide?.name).toBe('Probe Slate');
    expect(seen.styleGuide?.headerFill).toBe('1E3A8A');
    expect(seen.styleGuide?.calloutTheme).toBe('slate');

    const plainPath = path.join(scratchDir, 'plain.docx');
    const plain = await handleNewDocument(store, { outPath: plainPath });
    expect(plain.ok).toBe(true);
    if (!plain.ok) throw new Error('plain new_document failed');
    const plainSeen = await LEAN6_HANDLERS.inspect_document(store, { sessionId: plain.sessionId });
    expect(plainSeen.ok).toBe(true);
    if (!plainSeen.ok) throw new Error('plain inspect failed');
    expect('styleGuide' in (plainSeen as Record<string, unknown>)).toBe(false);
  });

  it('save carries advisory styleWarnings without blocking', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    const sessionId = await newGuidedSession(store, './slate.md', 'save-warn');
    const inserted = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        {
          action: 'insert',
          type: 'table',
          anchor: 'end',
          position: 'after',
          headers: ['Metric'],
          rows: [['1']],
        },
      ],
    });
    expect(inserted.ok).toBe(true);
    if (!inserted.ok) throw new Error(`insert failed: ${JSON.stringify(inserted)}`);
    // The fresh table carries a minted id (document order: last created).
    const tableId = inserted.createdNodeIds[inserted.createdNodeIds.length - 1];
    expect(tableId).toBeDefined();
    // Explicit dark-on-dark conflict: guide default is refused by the
    // explicit param, so the linter must flag it (advisory, never a failure).
    const shaded = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        {
          action: 'table',
          operation: 'setShading',
          target: tableId as string,
          headerFill: '1E3A8A',
          fill: 'F7F9FC',
          textColor: '000000',
        },
      ],
    });
    expect(shaded.ok).toBe(true);
    if (!shaded.ok) throw new Error(`shade failed: ${JSON.stringify(shaded)}`);
    expect(shaded.warnings.length).toBeGreaterThan(0);

    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(true);
    if (!saved.ok) throw new Error(`save failed: ${JSON.stringify(saved)}`);
    expect(saved.styleWarnings?.some((w) => w.rule === 'table-header-contrast')).toBe(true);
    const contrast = (saved.styleWarnings ?? []).find((w) => w.rule === 'table-header-contrast');
    expect(contrast?.message).toContain('FFFFFF');
  });

  it('save without a guide carries no styleWarnings field', async () => {
    const store = createMcpSessionStore();
    const outPath = path.join(scratchDir, 'noguide.docx');
    const created = await handleNewDocument(store, { outPath });
    expect(created.ok).toBe(true);
    if (!created.ok) throw new Error('new_document failed');
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: created.sessionId });
    expect(saved.ok).toBe(true);
    if (!saved.ok) throw new Error('save failed');
    expect('styleWarnings' in (saved as Record<string, unknown>)).toBe(false);
  });

  it('save flags persisted twin chart colors like validate', async () => {
    // Persisted twin: stored chart XML carries twin colors while the live
    // insert detail is gone after reopen, so only the stored-XML reader
    // shared with validate can flag it. Open carries no guide, so the pin
    // attaches the resolved file guide to isolate the save path.
    writeGuide('slate.md', SLATE_GUIDE);
    const store = createMcpSessionStore();
    const cases = [
      { tag: 'twin', colors: ['FF0000', 'FF0000'], want: 1 },
      { tag: 'distinct', colors: ['FF0000', '00FF00'], want: 0 },
    ];
    for (const { tag, colors, want } of cases) {
      const builder = new DocxParser(createBlankDocxBuffer('Twin Pin'));
      const ast = await builder.getAST();
      const engine = new OperationEngine(ast, builder);
      const anchor = ast.sections[0].nodes[ast.sections[0].nodes.length - 1].id;
      await engine.execute({
        action: 'chart',
        operation: 'insert',
        chartType: 'column',
        title: 'Twins',
        categories: ['A', 'B'],
        series: [
          { name: 'S1', values: [1, 2], color: colors[0] },
          { name: 'S2', values: [3, 4], color: colors[1] },
        ],
        anchor,
      } as never);
      const file = path.join(scratchDir, `${tag}.docx`);
      fs.writeFileSync(file, Buffer.from(builder.rezip()));
      const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
      expect(opened.ok).toBe(true);
      if (!opened.ok) throw new Error(`open failed: ${JSON.stringify(opened)}`);
      const handle = store.get(opened.sessionId);
      handle.styleGuide = resolveStyleGuide('./slate.md', { allowedDir: scratchDir }).config;
      const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId });
      expect(saved.ok).toBe(true);
      if (!saved.ok) throw new Error(`save failed: ${JSON.stringify(saved)}`);
      const dups = (saved.styleWarnings ?? []).filter((w) => w.rule === 'chart-duplicate-colors');
      expect(dups).toHaveLength(want);
      if (want > 0) expect(dups[0].message).toContain('FF0000');
    }
  });
});

describe('compose_document guide attach', () => {
  it('attaches a file guide, skins the compose, and persists', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    const outPath = path.join(scratchDir, 'compose-base.docx');
    const created = await handleNewDocument(store, { outPath });
    expect(created.ok).toBe(true);
    if (!created.ok) throw new Error('new_document failed');
    const composed = await handleComposeDocument(store, {
      sessionId: created.sessionId,
      markdown: '# Composed\n\nBody text.\n',
      allowedDir: scratchDir,
      styleGuide: './slate.md',
    });
    expect(composed.ok).toBe(true);
    if (!composed.ok) throw new Error(`compose failed: ${JSON.stringify(composed)}`);
    expect(composed.styleGuide?.name).toBe('Probe Slate');
    expect(Array.isArray(composed.styleGuideWarnings)).toBe(true);
    const handle = store.get(created.sessionId);
    expect(handle.styleGuide?.callouts?.defaultTheme).toBe('slate');

    // Inherited afterwards: a themeless callout batch takes slate.
    const batched = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: created.sessionId,
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'After compose' }],
    });
    expect(batched.ok).toBe(true);
    if (!batched.ok) throw new Error(`batch failed: ${JSON.stringify(batched)}`);
    expect(documentXmlOf(store, created.sessionId)).toContain('w:color="475569"');
  });

  it('skins composed tables and alerts with the session guide', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    const outPath = path.join(scratchDir, 'compose-skin.docx');
    const created = await handleNewDocument(store, { outPath, styleGuide: './slate.md' });
    expect(created.ok).toBe(true);
    if (!created.ok) throw new Error('new_document failed');
    const composed = await handleComposeDocument(store, {
      sessionId: created.sessionId,
      markdown: '# T\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n> [!NOTE]\n> Body text.\n',
      allowedDir: scratchDir,
    });
    expect(composed.ok).toBe(true);
    if (!composed.ok) throw new Error(`compose failed: ${JSON.stringify(composed)}`);
    const xml = documentXmlOf(store, created.sessionId);
    // Navy headers + slate callout (legacy gray/indigo absent).
    expect(xml).toContain('w:fill="1E3A8A"');
    expect(xml).not.toContain('w:fill="D9D9D9"');
    expect(xml).toContain('w:color="475569"');
    expect(xml).not.toContain('w:color="4F46E5"');
  });

  it("keeps the legacy compose fills under styleGuide: 'legacy', and skins a plain compose with the default guide", async () => {
    const store = createMcpSessionStore();
    const outPath = path.join(scratchDir, 'compose-legacy.docx');
    const created = await handleNewDocument(store, { outPath });
    expect(created.ok).toBe(true);
    if (!created.ok) throw new Error('new_document failed');
    const composed = await handleComposeDocument(store, {
      sessionId: created.sessionId,
      markdown: '# T\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n> [!NOTE]\n> Body text.\n',
      allowedDir: scratchDir,
      styleGuide: 'legacy',
    });
    expect(composed.ok).toBe(true);
    if (!composed.ok) throw new Error(`compose failed: ${JSON.stringify(composed)}`);
    const xml = documentXmlOf(store, created.sessionId);
    expect(xml).toContain('w:fill="D9D9D9"');
    expect(xml).toContain('w:color="4F46E5"');

    const plain = createMcpSessionStore();
    const made = await handleNewDocument(plain, { outPath: path.join(scratchDir, 'compose-default.docx') });
    if (!made.ok) throw new Error('new_document failed');
    const guided = await handleComposeDocument(plain, {
      sessionId: made.sessionId,
      markdown: '# T\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n> [!NOTE]\n> Body text.\n',
      allowedDir: scratchDir,
    });
    expect(guided.ok).toBe(true);
    const guidedXml = documentXmlOf(plain, made.sessionId);
    expect(guidedXml).toContain('w:fill="1E3A8A"');
    expect(guidedXml).not.toContain('w:fill="D9D9D9"');
    // Alerts keep their per-tag colour under the default guide.
    expect(guidedXml).toContain('w:color="4F46E5"');
  });
});

describe('guide diagnostics on machine channels', () => {
  it('execute_batch carries guide-parse notes in client-visible warnings', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    writeGuide('badhex.md', BADHEX_GUIDE);
    const sessionId = await newGuidedSession(store, './slate.md', 'batch-notes');
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      styleGuide: './badhex.md',
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Override note' }],
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`batch failed: ${JSON.stringify(res)}`);
    expect(res.warnings.some((w) => /malformed hex/.test(w))).toBe(true);
  });

  it('empty styleGuide hints detach instead of a bare schema error', async () => {
    const store = createMcpSessionStore();
    writeGuide('slate.md', SLATE_GUIDE);
    const sessionId = await newGuidedSession(store, './slate.md', 'detach');
    const batched = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      styleGuide: '',
      operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Never applied' }],
    });
    expect(batched.ok).toBe(false);
    if (batched.ok) throw new Error('expected failure');
    expect(batched.message).toMatch(/cannot be detached — open a new session/);
    const composed = await handleComposeDocument(store, {
      sessionId,
      markdown: '# T\n\nBody.\n',
      styleGuide: '',
    });
    expect(composed.ok).toBe(false);
    if (composed.ok) throw new Error('expected failure');
    expect(composed.message).toMatch(/cannot be detached — open a new session/);
    const outPath = path.join(scratchDir, 'detach.docx');
    const created = await handleNewDocument(store, { outPath, styleGuide: '' });
    expect(created.ok).toBe(false);
    if (created.ok) throw new Error('expected failure');
    expect(created.message).toMatch(/cannot be detached — open a new session/);
  });
});
