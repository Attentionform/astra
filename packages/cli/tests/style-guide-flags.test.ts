/**
 * CLI --style-guide plumbing: create / batch (single + stream) / validate.
 *
 * Per-invocation scope (CLI is stateless): one guide per call, resolved
 * fail-closed before any read/apply; --stream applies one guide to the
 * whole resident session. validate surfaces advisory styleWarnings with
 * exit 0. Guide-less runs stay byte-identical (covered inline).
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run, runStreamSession, type StreamSessionOpts } from '../src/cli';
import { AdmZip, DocxParser, resolveStyleGuide } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const GUIDE_MD = [
  '# Style Guide: Probe Slate',
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

const BAD_GUIDE_MD = [
  '# Style Guide: Probe Badhex',
  '',
  '## Palette',
  '- Primary: #ZZZZZZ',
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
  '- Default Theme: neon-pink',
  '',
  '## Code Blocks',
  '- Background: #F1F5F9',
  '- Border Color: #CBD5E1',
  '',
  '## Layout',
  '- Margin: standard',
  '',
].join('\n');

let scratchDir: string;
let savedCwd: string;
const tmpFiles: string[] = [];

function scratchPath(name: string): string {
  const file = path.join(scratchDir, name);
  tmpFiles.push(file);
  return file;
}

function writeBlankDocx(name: string): string {
  const file = scratchPath(name);
  fs.writeFileSync(file, DocxParser.createNewDocument('blank').rezip());
  return file;
}

function writeOps(payload: unknown): string {
  const file = scratchPath(`ops-${tmpFiles.length}.json`);
  fs.writeFileSync(file, JSON.stringify(payload));
  return file;
}

function documentXml(file: string): string {
  const zip = new AdmZip(fs.readFileSync(file));
  return zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

function stdoutJson(): Array<Record<string, unknown>> {
  const out: Array<Record<string, unknown>> = [];
  for (const call of vi.mocked(console.log).mock.calls) {
    try {
      out.push(JSON.parse(String(call[0])) as Record<string, unknown>);
    } catch {
      /* human lines are not JSON */
    }
  }
  return out;
}

// Validate/create file refs are cwd-anchored (no --allowed-dir on
// validate; create defaults there too), so the suite runs with cwd parked
// on the scratch dir and restores it afterwards.
beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
  savedCwd = process.cwd();
  scratchDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-guide-'));
  fs.writeFileSync(path.join(scratchDir, 'brand.md'), GUIDE_MD);
  process.chdir(scratchDir);
});

afterEach(() => {
  vi.restoreAllMocks();
  try {
    process.chdir(savedCwd);
  } catch {
    /* best-effort */
  }
  while (tmpFiles.length > 0) {
    const f = tmpFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
  try {
    fs.rmSync(scratchDir, { recursive: true, force: true });
  } catch {
    /* best-effort */
  }
});

describe('create --style-guide', () => {
  it('skins --from-md code blocks with the guide', async () => {
    const md = scratchPath('doc.md');
    fs.writeFileSync(md, '```ts\nconst x = 1;\n```\n');
    const out = scratchPath('guided.docx');
    const code = await run(['create', out, '--from-md', md, '--style-guide', './brand.md']);
    expect(code).toBe(0);
    const xml = documentXml(out);
    expect(xml).toContain('w:fill="F1F5F9"');
    expect(xml).not.toContain('w:fill="F2F2F2"');
  });

  it('gives a plain compose the default code background, and --style-guide legacy the old gray', async () => {
    const md = scratchPath('plain.md');
    fs.writeFileSync(md, '```ts\nconst x = 1;\n```\n');
    const out = scratchPath('plain.docx');
    expect(await run(['create', out, '--from-md', md])).toBe(0);
    expect(documentXml(out)).toContain('w:fill="F1F5F9"');
    const legacy = scratchPath('plain-legacy.docx');
    expect(await run(['create', legacy, '--from-md', md, '--style-guide', 'legacy'])).toBe(0);
    expect(documentXml(legacy)).toContain('w:fill="F2F2F2"');
  });

  it('fails closed on unknown names, naming the available set', async () => {
    const out = scratchPath('unknown.docx');
    const code = await run(['create', out, '--style-guide', 'definitely-not-a-guide']);
    expect(code).toBe(1);
    expect(fs.existsSync(out)).toBe(false);
    const stderr = vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
    expect(stderr).toMatch(/unknown style guide|available|no bundled guides/);
  });

  it('skins --from-md tables with the guide header fill', async () => {
    const md = scratchPath('tab.md');
    fs.writeFileSync(md, '# T\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n');
    const out = scratchPath('tab-guided.docx');
    const code = await run(['create', out, '--from-md', md, '--style-guide', './brand.md']);
    expect(code).toBe(0);
    const xml = documentXml(out);
    expect(xml).toContain('w:fill="1E3A8A"');
    expect(xml).not.toContain('w:fill="D9D9D9"');
  });

  it('skins --from-md alerts with the guide default theme', async () => {
    const md = scratchPath('alert.md');
    fs.writeFileSync(md, '# T\n\n> [!NOTE]\n> **Heads up**\n> Body text.\n');
    const out = scratchPath('alert-guided.docx');
    const code = await run(['create', out, '--from-md', md, '--style-guide', './brand.md']);
    expect(code).toBe(0);
    // Slate guide default (border 475569), not the legacy NOTE indigo.
    const xml = documentXml(out);
    expect(xml).toContain('w:color="475569"');
    expect(xml).not.toContain('w:color="4F46E5"');
  });

  it('keeps per-tag alert themes by default and the old gray table under --style-guide legacy', async () => {
    const md = scratchPath('legacy.md');
    fs.writeFileSync(md, '# T\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n> [!NOTE]\n> Body text.\n');
    const out = scratchPath('legacy.docx');
    const code = await run(['create', out, '--from-md', md]);
    expect(code).toBe(0);
    const xml = documentXml(out);
    expect(xml).toContain('w:fill="1E3A8A"');
    expect(xml).not.toContain('w:fill="D9D9D9"');
    expect(xml).toContain('w:color="4F46E5"');
    const legacy = scratchPath('legacy-old.docx');
    expect(await run(['create', legacy, '--from-md', md, '--style-guide', 'legacy'])).toBe(0);
    const old = documentXml(legacy);
    expect(old).toContain('w:fill="D9D9D9"');
    expect(old).not.toContain('w:fill="1E3A8A"');
    expect(old).toContain('w:color="4F46E5"');
  });

  it('carries guide-parse notes in --json styleWarnings', async () => {
    fs.writeFileSync(path.join(scratchDir, 'badhex.md'), BAD_GUIDE_MD);
    const md = scratchPath('warn.md');
    fs.writeFileSync(md, '# T\n\nBody.\n');
    const out = scratchPath('warn.docx');
    const code = await run(['--json', 'create', out, '--from-md', md, '--style-guide', './badhex.md']);
    expect(code).toBe(0);
    const payloads = stdoutJson();
    const ok = payloads.find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    const notes = ok?.['styleWarnings'] as string[] | undefined;
    expect(Array.isArray(notes)).toBe(true);
    expect(notes?.some((n) => /malformed hex|unknown callout theme/.test(n))).toBe(true);
  });
});

describe('compose warnings', () => {
  it('reports a [!NOTE] outside a blockquote in --json composeWarnings, for create and import', async () => {
    const md = scratchPath('stray.md');
    fs.writeFileSync(md, '# T\n\n[!NOTE] Not a callout.\n');
    const created = scratchPath('stray-create.docx');
    expect(await run(['--json', 'create', created, '--from-md', md])).toBe(0);
    const okCreate = stdoutJson().find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect((okCreate?.['composeWarnings'] as string[] | undefined)?.[0]).toMatch(/line 3.*\[!NOTE\].*blockquote/);
    const imported = scratchPath('stray-import.docx');
    expect(await run(['--json', 'import', md, '-o', imported])).toBe(0);
    const okImport = stdoutJson().filter((p) => p['ok'] === true).pop() as Record<string, unknown> | undefined;
    expect((okImport?.['composeWarnings'] as string[] | undefined)?.[0]).toMatch(/\[!NOTE\]/);
  });

  it('omits composeWarnings when there is nothing to report', async () => {
    const md = scratchPath('clean.md');
    fs.writeFileSync(md, '# T\n\nBody.\n');
    const out = scratchPath('clean.docx');
    expect(await run(['--json', 'create', out, '--from-md', md])).toBe(0);
    const ok = stdoutJson().find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    expect('composeWarnings' in (ok ?? {})).toBe(false);
  });
});

describe('batch --style-guide', () => {
  it('applies guide defaults where ops omit params', async () => {
    const file = writeBlankDocx('batch-guided.docx');
    const ops = writeOps([
      { action: 'callout', anchor: 'end', position: 'after', text: 'House note' },
    ]);
    const code = await run(['batch', file, ops, '--style-guide', './brand.md']);
    expect(code).toBe(0);
    const xml = documentXml(file);
    expect(xml).toContain('w:color="475569"');
    expect(xml).not.toContain('w:color="4F46E5"');
  });

  it('keeps legacy indigo without a guide', async () => {
    const file = writeBlankDocx('batch-plain.docx');
    const ops = writeOps([
      { action: 'callout', anchor: 'end', position: 'after', text: 'Legacy note' },
    ]);
    const code = await run(['batch', file, ops]);
    expect(code).toBe(0);
    expect(documentXml(file)).toContain('w:color="4F46E5"');
  });

  it('fails closed on unknown names and writes nothing', async () => {
    const file = writeBlankDocx('batch-unknown.docx');
    const before = fs.readFileSync(file);
    const ops = writeOps([{ action: 'replace', target: 'nope', text: 'x' }]);
    const code = await run(['batch', file, ops, '--style-guide', 'definitely-not-a-guide']);
    expect(code).toBe(1);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('refuses guide files outside the input folder jail', async () => {
    const file = writeBlankDocx('batch-jail.docx');
    const elsewhere = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-guide-out-'));
    tmpFiles.push(elsewhere);
    const outside = path.join(elsewhere, 'outside.md');
    fs.writeFileSync(outside, GUIDE_MD);
    const ops = writeOps([{ action: 'replace', target: 'nope', text: 'x' }]);
    const code = await run(['batch', file, ops, '--style-guide', outside]);
    expect(code).toBe(1);
    try {
      fs.rmSync(elsewhere, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  });
});

describe('batch --stream --style-guide', () => {
  it('applies one guide to the whole resident stream', async () => {
    const file = writeBlankDocx('stream-guided.docx');
    const docx = fs.readFileSync(file);
    const resolved = resolveStyleGuide('./brand.md', { allowedDir: scratchDir });
    const opts: StreamSessionOpts = {
      file,
      target: file,
      sequential: undefined,
      skipValidate: false,
      flagCeilings: {},
      backupState: { path: undefined, done: false },
      backupSource: docx,
      reproducible: undefined,
      macroOpts: undefined,
      onError: 'abort',
      maxErrors: 10,
      allowedDir: scratchDir,
      styleGuide: resolved.config,
    };
    const lines = [
      JSON.stringify({
        operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Stream one' }],
      }),
      JSON.stringify({
        operations: [{ action: 'callout', anchor: 'end', position: 'after', text: 'Stream two' }],
      }),
    ];
    const acks: Record<string, unknown>[] = [];
    const code = await runStreamSession({ json: false, verbose: false }, docx, opts, lines, (a) => {
      acks.push(a);
    });
    expect(code).toBe(0);
    expect(acks.filter((a) => a['ok'] === true).length).toBeGreaterThanOrEqual(2);
    const xml = documentXml(file);
    expect(xml).toContain('w:color="475569"');
    expect(xml).not.toContain('w:color="4F46E5"');
  });
});

describe('validate --style-guide', () => {
  it('reports advisory styleWarnings with exit 0', async () => {
    const file = writeBlankDocx('validate-warn.docx');
    // Build dark-on-dark bytes: explicit black text on the navy header.
    const parser = new DocxParser(fs.readFileSync(file));
    const ast = await parser.getAST();
    const lastId = ast.sections[0].nodes[ast.sections[0].nodes.length - 1].id;
    const { OperationEngine } = await import('@attentionform/astra-engine');
    const engine = new OperationEngine(ast, parser);
    await engine.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'table',
          anchor: lastId,
          position: 'after',
          headers: ['Metric'],
          rows: [['1']],
        } as never,
      ],
    });
    const table = ast.sections[0].nodes.find((n) => n.type === 'table');
    expect(table).toBeDefined();
    await engine.execute({
      action: 'table',
      target: (table as { id: string }).id,
      operation: 'setShading',
      headerFill: '1E3A8A',
      fill: 'F7F9FC',
      textColor: '000000',
    } as never);
    fs.writeFileSync(file, parser.rezip());

    const code = await run(['--json', 'validate', file, '--style-guide', './brand.md']);
    expect(code).toBe(0);
    const payloads = stdoutJson();
    const ok = payloads.find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    const warnings = ok?.['styleWarnings'] as Array<{ rule?: string }> | undefined;
    expect(Array.isArray(warnings)).toBe(true);
    expect(warnings?.some((w) => w.rule === 'table-header-contrast')).toBe(true);
    const guide = ok?.['styleGuide'] as Record<string, unknown> | undefined;
    expect(guide?.['name']).toBe('Probe Slate');
  });

  it('stays warning-free without a guide (envelope unchanged)', async () => {
    const file = writeBlankDocx('validate-plain.docx');
    const code = await run(['--json', 'validate', file]);
    expect(code).toBe(0);
    const payloads = stdoutJson();
    const ok = payloads.find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    expect('styleWarnings' in (ok as Record<string, unknown>)).toBe(false);
    expect('styleGuide' in (ok as Record<string, unknown>)).toBe(false);
  });

  it('carries guide-parse notes as guide-parse entries', async () => {
    fs.writeFileSync(path.join(scratchDir, 'badhex.md'), BAD_GUIDE_MD);
    const file = writeBlankDocx('validate-parse.docx');
    const code = await run(['--json', 'validate', file, '--style-guide', './badhex.md']);
    expect(code).toBe(0);
    const payloads = stdoutJson();
    const ok = payloads.find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    const warnings = ok?.['styleWarnings'] as Array<{ rule?: string; message?: string }> | undefined;
    expect(Array.isArray(warnings)).toBe(true);
    expect(warnings?.some((w) => w.rule === 'guide-parse' && /malformed hex|unknown callout theme/.test(w.message ?? ''))).toBe(true);
  });
});

describe('import --style-guide', () => {
  it('skins tables and alerts like create --from-md', async () => {
    const md = scratchPath('imp.md');
    fs.writeFileSync(md, '# T\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n> [!WARNING]\n> Watch out.\n');
    const out = scratchPath('imp-guided.docx');
    const code = await run(['import', md, '-o', out, '--style-guide', './brand.md']);
    expect(code).toBe(0);
    const xml = documentXml(out);
    // Slate guide: navy headers + slate callout (not legacy gray/indigo).
    expect(xml).toContain('w:fill="1E3A8A"');
    expect(xml).not.toContain('w:fill="D9D9D9"');
    expect(xml).toContain('w:color="475569"');
    expect(xml).not.toContain('w:color="4F46E5"');
  });

  it('uses the default navy header without a guide and the old gray under --style-guide legacy', async () => {
    const md = scratchPath('imp-plain.md');
    fs.writeFileSync(md, '# T\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n');
    const out = scratchPath('imp-plain.docx');
    const code = await run(['import', md, '-o', out]);
    expect(code).toBe(0);
    expect(documentXml(out)).toContain('w:fill="1E3A8A"');
    const legacy = scratchPath('imp-legacy.docx');
    expect(await run(['import', md, '-o', legacy, '--style-guide', 'legacy'])).toBe(0);
    expect(documentXml(legacy)).toContain('w:fill="D9D9D9"');
  });

  it('fails closed on unknown names and writes nothing', async () => {
    const md = scratchPath('imp-unknown.md');
    fs.writeFileSync(md, '# T\n\nBody.\n');
    const out = scratchPath('imp-unknown.docx');
    const code = await run(['import', md, '-o', out, '--style-guide', 'definitely-not-a-guide']);
    expect(code).toBe(1);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('carries guide-parse notes in --json styleWarnings', async () => {
    fs.writeFileSync(path.join(scratchDir, 'badhex.md'), BAD_GUIDE_MD);
    const md = scratchPath('imp-warn.md');
    fs.writeFileSync(md, '# T\n\nBody.\n');
    const out = scratchPath('imp-warn.docx');
    const code = await run(['--json', 'import', md, '-o', out, '--style-guide', './badhex.md']);
    expect(code).toBe(0);
    const payloads = stdoutJson();
    const ok = payloads.find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    expect((ok?.['styleGuide'] as Record<string, unknown> | undefined)?.['name']).toBe('Probe Badhex');
    const notes = ok?.['styleWarnings'] as string[] | undefined;
    expect(Array.isArray(notes)).toBe(true);
    expect(notes?.some((n) => /malformed hex|unknown callout theme/.test(n))).toBe(true);
  });
});

describe('batch --style-guide machine channel', () => {
  it('carries guide-parse notes in --json styleWarnings', async () => {
    fs.writeFileSync(path.join(scratchDir, 'badhex.md'), BAD_GUIDE_MD);
    const file = writeBlankDocx('batch-warn.docx');
    const ops = writeOps([
      { action: 'callout', anchor: 'end', position: 'after', text: 'Warn note' },
    ]);
    const code = await run(['--json', 'batch', file, ops, '--style-guide', './badhex.md']);
    expect(code).toBe(0);
    const payloads = stdoutJson();
    const ok = payloads.find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    const notes = ok?.['styleWarnings'] as string[] | undefined;
    expect(Array.isArray(notes)).toBe(true);
    expect(notes?.some((n) => /malformed hex|unknown callout theme/.test(n))).toBe(true);
  });

  it('omits styleWarnings for a clean guide (envelope unchanged)', async () => {
    const file = writeBlankDocx('batch-clean.docx');
    const ops = writeOps([
      { action: 'callout', anchor: 'end', position: 'after', text: 'Clean note' },
    ]);
    const code = await run(['--json', 'batch', file, ops, '--style-guide', './brand.md']);
    expect(code).toBe(0);
    const payloads = stdoutJson();
    const ok = payloads.find((p) => p['ok'] === true) as Record<string, unknown> | undefined;
    expect(ok).toBeDefined();
    expect('styleWarnings' in (ok as Record<string, unknown>)).toBe(false);
  });
});
