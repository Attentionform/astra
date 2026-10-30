import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { printCliVisualTelemetry } from '../src/cli/visual-cmd';
import { AdmZip, DocxParser, buildVisualTelemetry, type DocumentAST } from '@attentionform/astra-engine';
import { withLoSlot } from '../../../tests/helpers/lo-slot';
import { createMcpSessionStore } from '../../mcp-server/src/lib/session-store';
import { handleVisualInspect } from '../../mcp-server/src/tools/visual';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

// First-class `visual` command: saved-file renders via the same builder as
// `inspect --visual` (alias parity), with MCP post-save render parity and a
// fail-closed LO-absent remedy. LO-backed cases hold the suite-wide slot.
const tmpFiles: string[] = [];
const tmpDirs: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function tmpDir(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), `${prefix}-`));
  tmpDirs.push(dir);
  return dir;
}

function writeBlankDocx(prefix = 'astra-visual'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
  return file;
}

function writeMultiPageDocx(prefix = 'astra-visual-multi', paras = 80): string {
  const parser = DocxParser.createNewDocument('blank');
  for (let i = 0; i < paras; i++) {
    parser.insertRootNode(
      `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i} ` + 'lorem ipsum dolor sit amet '.repeat(10) + `</w:t></w:r></w:p>`,
      undefined,
      'append',
    );
  }
  const file = tmpPath(prefix);
  fs.writeFileSync(file, parser.rezip());
  return file;
}

function writeMacroDocx(prefix = 'astra-visual-macro'): string {
  const blank = DocxParser.createNewDocument('blank').rezip();
  const zip = new AdmZip(blank);
  zip.addFile('word/vbaProject.bin', Buffer.from('fake-macro-bytes'));
  const file = tmpPath(prefix);
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

/**
 * LO-count ≠ word-layout-count fixture: paragraphs that are one very long
 * token each. The word-layout estimate counts ~12 words/line so each
 * paragraph costs one line, while LibreOffice wraps the token character by
 * character and paginates far past the estimate — a deterministic,
 * environment-robust pagination mismatch (the same defect shape as the real
 * docs that reported `[PAGE 31 OF 21]`).
 */
function writeLongTokenDocx(prefix = 'astra-visual-mismatch', paras = 60, tokenLen = 1500): string {
  const parser = DocxParser.createNewDocument('blank');
  const token = 'x'.repeat(tokenLen);
  for (let i = 0; i < paras; i++) {
    parser.insertRootNode(
      `<w:p><w:r><w:t xml:space="preserve">p${i} ${token}</w:t></w:r></w:p>`,
      undefined,
      'append',
    );
  }
  const file = tmpPath(prefix);
  fs.writeFileSync(file, parser.rezip());
  return file;
}

/** Parse a paragraph-built doc and return its AST plus the root node ids. */
async function buildAstWithParagraphs(texts: string[]): Promise<{ ast: DocumentAST; ids: string[] }> {
  const built = DocxParser.createNewDocument('blank');
  for (const t of texts) {
    built.insertRootNode(`<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`, undefined, 'append');
  }
  const parser = new DocxParser(built.rezip());
  const ast = await parser.getAST();
  const ids = (ast.sections ?? []).flatMap((s) => (s.nodes ?? []).map((n) => n.id));
  return { ast, ids };
}

function stdoutText(): string {
  return vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
}

function stderrText(): string {
  return vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
}

async function runCli(args: string[]): Promise<{ status: number; stdout: string; stderr: string }> {
  vi.mocked(console.log).mockClear();
  vi.mocked(console.error).mockClear();
  const status = await run(args);
  return { status, stdout: stdoutText(), stderr: stderrText() };
}

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  while (tmpFiles.length > 0) {
    const f = tmpFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
  while (tmpDirs.length > 0) {
    const d = tmpDirs.pop() as string;
    try {
      fs.rmSync(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

describe('visual usage contract (exit 2, no LO needed)', () => {
  it('bare visual exits 2 with a usage error', async () => {
    const r = await runCli(['visual']);
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/usage error/i);
  });
  it('bare visual --json carries the E_USAGE envelope with quiet stderr', async () => {
    const r = await runCli(['--json', 'visual']);
    expect(r.status).toBe(2);
    const payload = JSON.parse(r.stdout);
    expect(payload.ok).toBe(false);
    expect(payload.code).toBe('E_USAGE');
    expect(typeof payload.remedy).toBe('string');
    expect(payload.remedy.length).toBeGreaterThan(0);
    expect(r.stderr).toBe('');
  });
  it('missing --out, bad page, bad format, unknown flag, extra positional all exit 2', async () => {
    const file = writeBlankDocx('astra-visual-usage');
    const png = tmpPath('astra-visual-usage', '.png');
    for (const args of [
      ['visual', file],
      ['visual', file, '--out', png, '--page', '0'],
      ['visual', file, '--out', png, '--page', 'abc'],
      ['visual', file, '--out', png, '--format', 'gif'],
      ['visual', file, '--out', png, '--bogus'],
      ['visual', file, 'extra.docx', '--out', png],
    ] as string[][]) {
      const r = await runCli(args);
      expect(r.status).toBe(2);
      expect(r.stderr).toMatch(/usage error/i);
    }
    expect(fs.existsSync(png)).toBe(false);
  });
  it('--help names the saved-file input and the MCP distinction', async () => {
    const r = await runCli(['visual', '--help']);
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('Usage: astra visual');
    expect(r.stdout).toContain('--out');
    expect(r.stdout).toContain('saved');
    expect(r.stdout).toContain('visual_inspect');
    expect(r.stdout).toContain('--force');
    expect(r.stdout).toContain('--allow-macros');
    expect(r.stdout).toContain('without --force');
  });
  it('new parser/ceiling/force flags misuse all exit 2 without touching --out', async () => {
    const file = writeBlankDocx('astra-visual-flags');
    const png = tmpPath('astra-visual-flags', '.png');
    fs.writeFileSync(png, 'sentinel');
    for (const args of [
      ['visual', file, '--out', png],
      ['visual', file, '--out', tmpPath('astra-visual-flags', '.png'), '--allow-macros', '--strip-macros'],
      ['visual', file, '--out', tmpPath('astra-visual-flags', '.png'), '--max-ratio', '0'],
      ['visual', file, '--out', tmpPath('astra-visual-flags', '.png'), '--max-ratio', 'abc'],
      ['visual', file, '--out', tmpPath('astra-visual-flags', '.png'), '--max-uncompressed-bytes', '-5'],
      ['visual', file, '--out', tmpPath('astra-visual-flags', '.png'), '--force', '--bogus-combo'],
    ] as string[][]) {
      const r = await runCli(args);
      expect(r.status).toBe(2);
      expect(r.stderr).toMatch(/usage error/i);
    }
    expect(fs.readFileSync(png, 'utf8')).toBe('sentinel');
  });
  it('inspect --visual --page misuse exits 2', async () => {
    const file = writeBlankDocx('astra-visual-inspect-page');
    for (const args of [
      ['inspect', file, '--visual', '--page', '0'],
      ['inspect', file, '--visual', '--page', 'abc'],
      ['inspect', file, '--visual', '--format', 'gif'],
      ['inspect', file, '--visual', '--allow-macros', '--strip-macros'],
    ] as string[][]) {
      const r = await runCli(args);
      expect(r.status).toBe(2);
      expect(r.stderr).toMatch(/usage error/i);
    }
  });
  it('missing value at end for --page/--format/--max-ratio exits 2 naming the flag (visual + inspect)', async () => {
    // Script typos must not silently fall back to the default page/format/
    // ceiling (a wrong page would render and exit 0).
    const file = writeBlankDocx('astra-visual-missing-value');
    const png = tmpPath('astra-visual-missing-value', '.png');
    const cases: Array<[string[], string]> = [
      [['visual', file, '--out', png, '--page'], '--page'],
      [['visual', file, '--out', png, '--format'], '--format'],
      [['visual', file, '--out', png, '--max-ratio'], '--max-ratio'],
      [['inspect', file, '--page'], '--page'],
      [['inspect', file, '--format'], '--format'],
      [['inspect', file, '--max-ratio'], '--max-ratio'],
    ];
    for (const [args, flag] of cases) {
      const r = await runCli(args);
      expect(r.status).toBe(2);
      expect(r.stderr).toMatch(/usage error/i);
      expect(r.stderr).toContain(`${flag} requires a value`);
    }
    // None of the typo shapes rendered or wrote the --out target.
    expect(fs.existsSync(png)).toBe(false);
  });
});

describe('visual read failures (exit 1, single-shape wording)', () => {
  it('missing input is E_IO with the shared remedy', async () => {
    const ghost = `ghost-${process.pid}-${Date.now()}.docx`;
    const png = tmpPath('astra-visual-ghost', '.png');
    const r = await runCli(['visual', ghost, '--out', png]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('[E_IO]');
    expect(r.stderr).toContain(`cannot read "${ghost}": no such file or directory`);
    expect(r.stderr).toContain('Remedy: check the path exists and is readable, then retry');
    expect(fs.existsSync(png)).toBe(false);
  });
  it('corrupt input is the clean package wording, nothing written', async () => {
    const bad = tmpPath('astra-visual-corrupt');
    fs.writeFileSync(bad, 'this is not a zip package');
    const png = tmpPath('astra-visual-corrupt', '.png');
    const r = await runCli(['visual', bad, '--out', png]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('[E_VALIDATION]');
    expect(r.stderr).toMatch(/is not a valid \.docx package/);
    expect(r.stderr).not.toMatch(/ADM-ZIP/i);
    expect(fs.existsSync(png)).toBe(false);
  });
});

describe('honest pagination marker (renderer is the single [PAGE n OF m] source of truth)', () => {
  const out = { json: false, verbose: false };

  it(
    'mismatch on an unmapped page: marker uses the renderer count, wireframe never claims an empty page, both warnings hit stderr',
    async () => {
      const { ast, ids } = await buildAstWithParagraphs(['Alpha paragraph', 'Beta paragraph']);
      const summary = { nodePages: { [ids[0]]: 1, [ids[1]]: 2 }, totalPages: 2 };
      const telemetry = printCliVisualTelemetry(out, ast, summary, 4, 4);
      // Single source of truth: `OF m` is the renderer's count, not the map's.
      expect(telemetry.wireframe).toContain('[PAGE 4 OF 4]');
      expect(telemetry.wireframe).not.toContain('[PAGE 4 OF 2]');
      // The rendered page exists in LO but has no word-layout nodes: say so,
      // never the false "(Empty Page / Section Break)" over full-content pages.
      expect(telemetry.wireframe).toContain('no word-layout nodes mapped');
      expect(telemetry.wireframe).not.toContain('(Empty Page / Section Break)');
      expect(telemetry.totalPages).toBe(4);
      expect(telemetry.wordLayoutTotalPages).toBe(2);
      expect(stdoutText()).toContain('visual: Pagination mismatch');
      // Renderer paginated BEYOND the map, so the "nodes beyond the map"
      // clause has content: the engine warning (stdout) and the stderr note
      // both keep it.
      expect(stdoutText()).toContain('nodes past the word-layout map are unmapped');
      // stderr, not silent — including under --json (printStderr is unconditional).
      expect(stderrText()).toContain('pagination mismatch');
      expect(stderrText()).toContain('renderer (LibreOffice) produced 4 pages');
      expect(stderrText()).toContain('word-layout page map totals 2');
      expect(stderrText()).toContain('wireframe nodes beyond page 2 are unmapped');
      expect(stderrText()).toContain('--page 4 exceeds the word-layout page map (2 pages)');
    },
  );

  it(
    'mismatch on a mapped page: node content still comes from the word-layout map, marker reconciled, no false unmapped label',
    async () => {
      const { ast, ids } = await buildAstWithParagraphs(['Alpha paragraph', 'Beta paragraph']);
      // createNewDocument seeds paragraphs first: the last two ids are ours.
      const summary = { nodePages: { [ids[ids.length - 2]]: 1, [ids[ids.length - 1]]: 2 }, totalPages: 2 };
      const telemetry = printCliVisualTelemetry(out, ast, summary, 2, 4);
      expect(telemetry.wireframe).toContain('[PAGE 2 OF 4]');
      expect(telemetry.wireframe).toContain('Beta paragraph');
      expect(telemetry.wireframe).not.toContain('no word-layout nodes mapped');
      expect(telemetry.wireframe).not.toContain('(Empty Page / Section Break)');
      // Mismatch warns; --page 2 is inside the map so the exceeds warning does not.
      expect(stderrText()).toContain('pagination mismatch');
      expect(stderrText()).not.toContain('exceeds the word-layout page map');
    },
  );

  it(
    'reverse mismatch (renderer fewer pages than the map) also labels the renderer count',
    async () => {
      const { ast, ids } = await buildAstWithParagraphs(['Alpha paragraph', 'Beta paragraph']);
      const summary = { nodePages: { [ids[0]]: 1, [ids[1]]: 3 }, totalPages: 3 };
      const telemetry = printCliVisualTelemetry(out, ast, summary, 1, 2);
      expect(telemetry.wireframe).toContain('[PAGE 1 OF 2]');
      expect(telemetry.totalPages).toBe(2);
      expect(telemetry.wordLayoutTotalPages).toBe(3);
      expect(stderrText()).toContain('pagination mismatch');
      expect(stderrText()).toContain('renderer (LibreOffice) produced 2 pages');
      // Reverse mismatch: nothing exists beyond the map, so the vacuous
      // "nodes beyond page N are unmapped" clause is dropped from both the
      // engine warning (stdout) and the CLI note (stderr).
      expect(stderrText()).not.toContain('nodes beyond page');
      expect(stdoutText()).not.toContain('nodes past the word-layout map');
    },
  );

  it(
    'counts of one render singular ("1 page"), never "1 pages"',
    async () => {
      const { ast, ids } = await buildAstWithParagraphs(['Alpha paragraph', 'Beta paragraph']);
      // Case A: the renderer produced a single page while the map totals 3.
      const summaryA = { nodePages: { [ids[0]]: 1 }, totalPages: 3 };
      const a = printCliVisualTelemetry(out, ast, summaryA, 1, 1);
      expect(a.totalPages).toBe(1);
      expect(a.wordLayoutTotalPages).toBe(3);
      expect(stderrText()).toContain('renderer (LibreOffice) produced 1 page;');
      expect(stdoutText()).toContain('produced 1 page;');
      // Case B: an over-range --page against a one-page word-layout map.
      const summaryB = { nodePages: { [ids[0]]: 1 }, totalPages: 1 };
      printCliVisualTelemetry(out, ast, summaryB, 2);
      expect(stderrText()).toContain('--page 2 exceeds the word-layout page map (1 page)');
      expect(stderrText()).not.toContain('1 pages');
      expect(stdoutText()).not.toContain('1 pages');
    },
  );

  it(
    'compat: without a renderer count the wireframe stays byte-identical to the word-layout output (no warnings)',
    async () => {
      const { ast, ids } = await buildAstWithParagraphs(['Alpha paragraph', 'Beta paragraph']);
      const summary = { nodePages: { [ids[ids.length - 2]]: 1, [ids[ids.length - 1]]: 2 }, totalPages: 2 };
      const telemetry = printCliVisualTelemetry(out, ast, summary, 2);
      expect(telemetry.wireframe).toContain('[PAGE 2 OF 2]');
      expect(telemetry.wireframe).toContain('Beta paragraph');
      expect(telemetry.totalPages).toBe(2);
      expect(telemetry.wordLayoutTotalPages).toBeUndefined();
      expect(stderrText()).toBe('');
    },
  );

  it(
    'MCP visual_inspect without withVisual: telemetry stays word-layout-only (no renderer relabel, no mismatch warning)',
    async () => {
      const file = writeLongTokenDocx('astra-visual-mcp-telemetry-only');
      const store = createMcpSessionStore();
      const h = await store.openBuffer(fs.readFileSync(file), 'mismatch-telemetry.docx');
      const mcp = await handleVisualInspect(store, { sessionId: h.sessionId });
      expect(mcp.ok).toBe(true);
      if (!mcp.ok) return;
      expect(mcp.images).toBeUndefined();
      expect(mcp.telemetry.wordLayoutTotalPages).toBeUndefined();
      expect(mcp.totalPages).toBe(mcp.telemetry.totalPages);
      expect(mcp.telemetry.warnings.join('\n')).not.toContain('Pagination mismatch');
    },
  );

  it(
    'renderer-reported zero pages never zeroes the telemetry total (guard: no totalPages 0 payload)',
    async () => {
      const { ast, ids } = await buildAstWithParagraphs(['Alpha paragraph']);
      const pageMap = { nodePages: { [ids[0]]: 5 }, totalPages: 5, source: 'renderer' as const, isEstimate: false };
      const t = buildVisualTelemetry(ast, { page: 1, pageMap, renderTotalPages: 0 });
      expect(t.totalPages).toBe(5);
      expect(t.wordLayoutTotalPages).toBeUndefined();
      expect(t.warnings.join('\n')).not.toContain('Pagination mismatch');
    },
  );
});

describe('visual LO-absent degrades fail-closed with a remedy', () => {
  it('hidden binaries exit 1 with E_DEPENDENCY naming LibreOffice (human + json)', async () => {
    const file = writeBlankDocx('astra-visual-nodeps');
    const prevPath = process.env.PATH;
    const emptyDir = tmpDir('astra-visual-empty-path');
    try {
      process.env.PATH = emptyDir;
      const png = tmpPath('astra-visual-nodeps', '.png');
      const human = await runCli(['visual', file, '--out', png]);
      expect(human.status).toBe(1);
      expect(human.stderr).toContain('[E_DEPENDENCY]');
      expect(human.stderr).toMatch(/LibreOffice/i);
      expect(human.stderr).toContain('Remedy:');
      expect(fs.existsSync(png)).toBe(false);
      const png2 = tmpPath('astra-visual-nodeps-json', '.png');
      const machined = await runCli(['--json', 'visual', file, '--out', png2]);
      expect(machined.status).toBe(1);
      const payload = JSON.parse(machined.stdout);
      expect(payload.ok).toBe(false);
      expect(payload.code).toBe('E_DEPENDENCY');
      expect(String(payload.error)).toMatch(/LibreOffice/i);
      expect(machined.stderr).toContain('[E_DEPENDENCY]');
      expect(fs.existsSync(png2)).toBe(false);
    } finally {
      if (prevPath === undefined) delete process.env.PATH;
      else process.env.PATH = prevPath;
    }
  });
});

describe('single telemetry builder (one definition, two entry points)', () => {
  it('engine builder and render path are defined once (visual-cmd) and shared by inspect --visual', () => {
    const repoCli = path.resolve(__dirname, '..', 'src', 'cli');
    const visualSrc = fs.readFileSync(path.join(repoCli, 'visual-cmd.ts'), 'utf8');
    const inspectSrc = fs.readFileSync(path.join(repoCli, 'inspect-cmd.ts'), 'utf8');
    expect(visualSrc).toContain('export function buildCliVisualTelemetry');
    expect(visualSrc).toContain('export function printCliVisualTelemetry');
    expect(inspectSrc).toContain('printCliVisualTelemetry');
    expect(inspectSrc).not.toContain('buildVisualTelemetry');
    expect(visualSrc).toContain('export async function renderCliVisualPage');
    expect(inspectSrc).toContain('renderCliVisualPage');
    expect(inspectSrc).not.toContain('renderDocxToImages');
    expect(visualSrc).toContain('macroOptsFromFlags');
    expect(visualSrc).toContain('withZipCeilingEnv');
    expect(inspectSrc).toContain('macroOptsFromFlags');
    expect(inspectSrc).toContain('withZipCeilingEnv');
  });
});

describe('visual parser gates (no LO needed: parse refuses before any render)', () => {
  it('macro file refuses E_VALIDATION on both entry points by default, naming the part', async () => {
    const file = writeMacroDocx('astra-visual-macro-refuse');
    const png = tmpPath('astra-visual-macro-refuse', '.png');
    const vr = await runCli(['visual', file, '--out', png]);
    expect(vr.status).toBe(1);
    expect(vr.stderr).toContain('[E_VALIDATION]');
    expect(vr.stderr).toMatch(/vbaProject\.bin/);
    expect(fs.existsSync(png)).toBe(false);
    const ir = await runCli(['inspect', file, '--visual']);
    expect(ir.status).toBe(1);
    expect(ir.stderr).toContain('[E_VALIDATION]');
    expect(ir.stderr).toMatch(/vbaProject\.bin/);
  });
});

describe('visual render parity (LO-backed)', () => {
  it(
    'visual renders page 1 to --out with telemetry parity to inspect --visual (byte-identical images)',
    async () => {
      const file = writeBlankDocx('astra-visual-parity');
      const png = tmpPath('astra-visual-parity', '.png');
      await withLoSlot(async () => {
        const vr = await runCli(['--json', 'visual', file, '--out', png]);
        expect(vr.status).toBe(0);
        const payload = JSON.parse(vr.stdout);
        expect(payload.ok).toBe(true);
        expect(payload.telemetry).toBeDefined();
        expect(payload.telemetry.wireframe).toContain('[PAGE 1 OF');
        expect(fs.existsSync(png)).toBe(true);
        const magic = fs.readFileSync(png).subarray(0, 8);
        expect(magic.equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))).toBe(true);

        const ir = await runCli(['--json', 'inspect', file, '--visual']);
        expect(ir.status).toBe(0);
        const inspectPayload = JSON.parse(ir.stdout);
        expect(inspectPayload.visual).toBeDefined();
        expect(inspectPayload.visual.telemetry.wireframe).toBe(payload.telemetry.wireframe);
        const inspectImg = inspectPayload.visual.images?.[0]?.path as string | undefined;
        expect(typeof inspectImg).toBe('string');
        expect(fs.existsSync(inspectImg as string)).toBe(true);
        expect(fs.readFileSync(png).equals(fs.readFileSync(inspectImg as string))).toBe(true);
      });
    },
    300000,
  );

  it(
    'CLI visual and MCP visual_inspect agree post-save (same bytes, same wireframe)',
    async () => {
      const file = writeBlankDocx('astra-visual-mcp');
      const png = tmpPath('astra-visual-mcp', '.png');
      await withLoSlot(async () => {
        const vr = await runCli(['visual', file, '--out', png]);
        expect(vr.status).toBe(0);
        expect(fs.existsSync(png)).toBe(true);

        const store = createMcpSessionStore();
        const buf = fs.readFileSync(file);
        const h = await store.openBuffer(buf, 'parity.docx');
        const mcp = await handleVisualInspect(store, { sessionId: h.sessionId, pages: [1], withVisual: true, format: 'png' });
        expect(mcp.ok).toBe(true);
        if (!mcp.ok) return;
        expect(mcp.images).toBeDefined();
        expect(mcp.images?.length).toBeGreaterThan(0);
        const mcpPath = mcp.images?.[0]?.path as string;
        expect(fs.existsSync(mcpPath)).toBe(true);
        expect(fs.readFileSync(png).equals(fs.readFileSync(mcpPath))).toBe(true);
        expect(mcp.telemetry.wireframe).toContain('[PAGE 1 OF');
      });
    },
    300000,
  );

  it(
    'MCP visual_inspect on a mismatch doc: marker names the renderer count returned in totalPages, renderer-only pages say unmapped, warnings carry the mismatch',
    async () => {
      const file = writeLongTokenDocx('astra-visual-mcp-mismatch');
      await withLoSlot(async () => {
        const store = createMcpSessionStore();
        const h = await store.openBuffer(fs.readFileSync(file), 'mismatch.docx');
        const mcp = await handleVisualInspect(store, { sessionId: h.sessionId, pages: [999], withVisual: true, format: 'png' });
        expect(mcp.ok).toBe(true);
        if (!mcp.ok) return;
        // Renderer count is the single source of truth on both the envelope
        // and the telemetry marker.
        expect(mcp.totalPages).toBeGreaterThanOrEqual(1);
        expect(mcp.telemetry.totalPages).toBe(mcp.totalPages);
        const wordTotal = mcp.telemetry.wordLayoutTotalPages as number | undefined;
        expect(wordTotal).toBeDefined();
        expect(mcp.totalPages).toBeGreaterThan(wordTotal as number);
        // pages:[999] clamps to the renderer's last page: the marker names
        // exactly the image returned (never the raw request, never the map).
        const img = mcp.images?.[0];
        expect(img).toBeDefined();
        expect(mcp.images?.length).toBe(1);
        expect(mcp.telemetry.wireframe).toContain(`[PAGE ${img?.page} OF ${mcp.totalPages}]`);
        expect(mcp.telemetry.wireframe).not.toContain('[PAGE 999 OF');
        expect(mcp.telemetry.wireframe).not.toContain(`OF ${wordTotal}]`);
        // The rendered page exists in LO but has no word-layout nodes:
        // honest note, never a false empty-page claim.
        expect(mcp.telemetry.wireframe).toContain('no word-layout nodes mapped');
        expect(mcp.telemetry.wireframe).not.toContain('(Empty Page / Section Break)');
        expect(mcp.telemetry.warnings.join('\n')).toContain('Pagination mismatch');
      });
    },
    300000,
  );

  it(
    'visual --out refuses an existing target without --force (exit 2, bytes preserved; --force clobbers)',
    async () => {
      const file = writeBlankDocx('astra-visual-force');
      const png = tmpPath('astra-visual-force', '.png');
      await withLoSlot(async () => {
        const first = await runCli(['--json', 'visual', file, '--out', png]);
        expect(first.status).toBe(0);
        const before = fs.readFileSync(png);
        expect(before.length).toBeGreaterThan(0);

        const refused = await runCli(['visual', file, '--out', png]);
        expect(refused.status).toBe(2);
        expect(refused.stderr).toMatch(/already exists.*--force/);
        expect(refused.stderr).toContain('visual:');
        expect(fs.readFileSync(png).equals(before)).toBe(true);

        const forced = await runCli(['--json', 'visual', file, '--out', png, '--force']);
        expect(forced.status).toBe(0);
        const payload = JSON.parse(forced.stdout);
        expect(payload.ok).toBe(true);
        expect(fs.readFileSync(png).equals(before)).toBe(true);
      });
    },
    300000,
  );

  it(
    'visual --format jpeg writes byte-pinned JPEG magic (FF D8 FF)',
    async () => {
      const file = writeBlankDocx('astra-visual-jpeg');
      const jpg = tmpPath('astra-visual-jpeg', '.jpg');
      await withLoSlot(async () => {
        const r = await runCli(['--json', 'visual', file, '--out', jpg, '--format', 'jpeg']);
        expect(r.status).toBe(0);
        const payload = JSON.parse(r.stdout);
        expect(payload.ok).toBe(true);
        expect(payload.format).toBe('jpeg');
        const bytes = fs.readFileSync(jpg);
        expect(bytes.length).toBeGreaterThan(0);
        expect(bytes[0]).toBe(0xff);
        expect(bytes[1]).toBe(0xd8);
        expect(bytes[2]).toBe(0xff);
        expect(payload.bytes).toBe(bytes.length);
      });
    },
    300000,
  );

  it(
    'visual --page past the end clamps to the last rendered page (byte-pinned PNG)',
    async () => {
      const file = writeBlankDocx('astra-visual-clamp');
      const png = tmpPath('astra-visual-clamp', '.png');
      await withLoSlot(async () => {
        const r = await runCli(['--json', 'visual', file, '--out', png, '--page', '999']);
        expect(r.status).toBe(0);
        const payload = JSON.parse(r.stdout);
        expect(payload.ok).toBe(true);
        expect(payload.totalPages).toBe(1);
        expect(payload.page).toBe(1);
        expect(payload.telemetry.wireframe).toContain('[PAGE 1 OF');
        const bytes = fs.readFileSync(png);
        expect(bytes.length).toBeGreaterThan(0);
        expect(bytes.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))).toBe(true);
      });
    },
    300000,
  );

  it(
    'multi-page alias identity: visual --page 2 and inspect --visual --page 2 agree byte-identically (and differ from page 1)',
    async () => {
      const file = writeMultiPageDocx('astra-visual-multi');
      const p2 = tmpPath('astra-visual-multi-p2', '.png');
      const p1 = tmpPath('astra-visual-multi-p1', '.png');
      await withLoSlot(async () => {
        const vr = await runCli(['--json', 'visual', file, '--out', p2, '--page', '2']);
        expect(vr.status).toBe(0);
        const payload = JSON.parse(vr.stdout);
        expect(payload.ok).toBe(true);
        expect(payload.page).toBe(2);
        expect(payload.totalPages).toBeGreaterThanOrEqual(2);
        expect(payload.telemetry.wireframe).toContain('[PAGE 2 OF');
        const p2Bytes = fs.readFileSync(p2);
        expect(p2Bytes.length).toBeGreaterThan(0);

        const ir = await runCli(['--json', 'inspect', file, '--visual', '--page', '2']);
        expect(ir.status).toBe(0);
        const inspectPayload = JSON.parse(ir.stdout);
        expect(inspectPayload.visual.telemetry.wireframe).toBe(payload.telemetry.wireframe);
        expect(inspectPayload.visual.images?.length).toBe(1);
        const inspectImg = inspectPayload.visual.images?.[0]?.path as string;
        expect(fs.existsSync(inspectImg)).toBe(true);
        expect(fs.readFileSync(inspectImg).equals(p2Bytes)).toBe(true);

        const r1 = await runCli(['--json', 'visual', file, '--out', p1, '--page', '1']);
        expect(r1.status).toBe(0);
        const p1Bytes = fs.readFileSync(p1);
        expect(p1Bytes.equals(p2Bytes)).toBe(false);
      });
    },
    300000,
  );

  it(
    'macro file with --strip-macros renders identically on both entry points (divergence closed)',
    async () => {
      const file = writeMacroDocx('astra-visual-macro-strip');
      const png = tmpPath('astra-visual-macro-strip', '.png');
      await withLoSlot(async () => {
        const vr = await runCli(['--json', 'visual', file, '--out', png, '--strip-macros']);
        expect(vr.status).toBe(0);
        const payload = JSON.parse(vr.stdout);
        expect(payload.ok).toBe(true);
        const pngBytes = fs.readFileSync(png);
        expect(pngBytes.length).toBeGreaterThan(0);

        const ir = await runCli(['--json', 'inspect', file, '--visual', '--strip-macros']);
        expect(ir.status).toBe(0);
        const inspectPayload = JSON.parse(ir.stdout);
        expect(inspectPayload.visual.telemetry.wireframe).toBe(payload.telemetry.wireframe);
        const inspectImg = inspectPayload.visual.images?.[0]?.path as string;
        expect(fs.readFileSync(inspectImg).equals(pngBytes)).toBe(true);
      });
    },
    300000,
  );

  it(
    'regression: on a doc where LO count != word-layout count, rendered page + wireframe + [PAGE n OF m] marker stay mutually consistent (both entry points)',
    async () => {
      const file = writeLongTokenDocx('astra-visual-mismatch');
      const png = tmpPath('astra-visual-mismatch', '.png');
      await withLoSlot(async () => {
        const r = await runCli(['--json', 'visual', file, '--out', png, '--page', '999']);
        expect(r.status).toBe(0);
        const payload = JSON.parse(r.stdout);
        expect(payload.ok).toBe(true);
        const wordTotal = payload.telemetry.wordLayoutTotalPages as number | undefined;
        // Fixture precondition: the doc really mismatches (LO wraps the long
        // tokens far past the word-layout estimate). If this fails, the
        // environment or estimator changed and no longer proves the fix.
        expect(wordTotal).toBeDefined();
        expect(payload.totalPages).toBeGreaterThan(wordTotal as number);
        // --page 999 clamps to LO's last page, and the marker names exactly
        // that image: [PAGE n OF m] with BOTH numbers from the renderer.
        expect(payload.page).toBe(payload.totalPages);
        expect(payload.telemetry.wireframe).toContain(`[PAGE ${payload.page} OF ${payload.totalPages}]`);
        expect(payload.telemetry.wireframe).not.toContain(`OF ${wordTotal}]`);
        expect(payload.telemetry.totalPages).toBe(payload.totalPages);
        // The rendered page has real content and no word-layout mapping:
        // the wireframe says unmapped — never a false "(Empty Page ...)".
        expect(payload.telemetry.wireframe).toContain('no word-layout nodes mapped');
        expect(payload.telemetry.wireframe).not.toContain('(Empty Page / Section Break)');
        expect(fs.existsSync(png)).toBe(true);
        expect(fs.readFileSync(png).length).toBeGreaterThan(0);
        // stderr (not silent, survives --json): mismatch + --page past the map.
        expect(r.stderr).toContain('pagination mismatch');
        expect(r.stderr).toContain('exceeds the word-layout page map');

        // Alias parity: inspect --visual --page 999 agrees on the wireframe,
        // the totals, and the rendered image bytes.
        const ir = await runCli(['--json', 'inspect', file, '--visual', '--page', '999']);
        expect(ir.status).toBe(0);
        const inspectPayload = JSON.parse(ir.stdout);
        expect(inspectPayload.visual.telemetry.wireframe).toBe(payload.telemetry.wireframe);
        expect(inspectPayload.visual.telemetry.wordLayoutTotalPages).toBe(wordTotal);
        expect(inspectPayload.visual.totalPages).toBe(payload.totalPages);
        expect(inspectPayload.visual.images?.length).toBe(1);
        const inspectImg = inspectPayload.visual.images?.[0]?.path as string;
        expect(fs.existsSync(inspectImg)).toBe(true);
        expect(fs.readFileSync(png).equals(fs.readFileSync(inspectImg))).toBe(true);
      });
    },
    300000,
  );
});
