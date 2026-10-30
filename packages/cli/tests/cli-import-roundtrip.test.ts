import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

// Import pipeline plus Markdown round-trip fidelity: in-process via
// run(argv). Fixtures live in tmpdirs; nothing writes outside.
const dirs: string[] = [];

function tmpDir(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), `${prefix}-`));
  dirs.push(dir);
  return dir;
}

async function runCli(args: string[]): Promise<{ status: number; stdout: string; stderr: string }> {
  vi.mocked(console.log).mockClear();
  vi.mocked(console.error).mockClear();
  const status = await run(args);
  return {
    status,
    stdout: vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n'),
    stderr: vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n'),
  };
}

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  for (const dir of dirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

// Author-style fixture covering every round-trip construct: headings,
// bullet + ordered lists, a table, a footnote, a divider, fenced code.
const MANUSCRIPT = [
  '# Field Notes on Endurance Training',
  '',
  'Opening paragraph sets the scene for the training block.',
  '',
  '## Weekly Structure',
  '',
  '### Day One Focus',
  '',
  '- warm up easy',
  '- main set steady',
  '- cool down slow',
  '',
  '1. first step',
  '2. second step',
  '',
  '| Metric | Value |',
  '| --- | --- |',
  '| Endurance | 42 |',
  '| Recovery | 7 |',
  '',
  '---',
  '',
  '```text',
  'log line one',
  'log line two',
  '```',
  '',
  'A note reference[^fn] in running text.',
  '',
  '[^fn]: The footnote body.',
  '',
].join('\n');

function writeManuscript(dir: string, name = 'manuscript.md'): string {
  const file = path.join(dir, name);
  fs.writeFileSync(file, MANUSCRIPT);
  return file;
}

// Template seed paragraphs round-trip as body text on re-import (compose is
// append-only), so stability is measured with seed lines stripped. Remaining
// comparison is exact after trailing-whitespace cleanup: stable modulo
// whitespace.
const SEED_LINES = new Set(['# Blank Document', 'Start typing your content here...']);

function normalizeMd(md: string): string {
  const kept: string[] = [];
  for (const raw of md.replace(/\r\n|\r/g, '\n').split('\n')) {
    const line = raw.replace(/[ \t]+$/, '');
    if (SEED_LINES.has(line.trim())) continue;
    kept.push(line);
  }
  while (kept.length > 0 && kept[0].trim() === '') kept.shift();
  while (kept.length > 0 && kept[kept.length - 1].trim() === '') kept.pop();
  return kept.join('\n');
}

describe('import manuscript', () => {
  it('composes a manuscript into a fresh document that validates', async () => {
    const dir = tmpDir('import-basic');
    const md = writeManuscript(dir);
    const out = path.join(dir, 'out.docx');
    const r = await runCli(['import', md, '-o', out]);
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('IMPORTED');
    expect(r.stdout).toContain('template=blank');
    expect(r.stdout).toMatch(/applied=[1-9]\d*/);
    expect((await runCli(['validate', out])).status).toBe(0);
    for (const query of ['Field Notes', 'warm up easy', 'first step', 'Endurance', 'footnote body']) {
      const found = await runCli(['search', out, query]);
      expect(found.status).toBe(0);
      expect(found.stdout).toMatch(/[1-9][0-9]* match/);
    }
  });

  it('defaults the output next to the input when -o is omitted', async () => {
    const dir = tmpDir('import-default-out');
    const md = writeManuscript(dir);
    const r = await runCli(['import', md]);
    expect(r.status).toBe(0);
    const expected = path.join(dir, 'manuscript.docx');
    expect(fs.existsSync(expected)).toBe(true);
    expect(r.stdout).toContain('manuscript.docx');
    expect((await runCli(['validate', expected])).status).toBe(0);
  });

  it('accepts --out= equals form and --template report', async () => {
    const dir = tmpDir('import-forms');
    const md = writeManuscript(dir);
    const out = path.join(dir, 'report.docx');
    const r = await runCli(['import', md, `--out=${out}`, '--template', 'report']);
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('template=report');
    expect((await runCli(['validate', out])).status).toBe(0);
  });

  it('rejects an unknown template with exit 2 and writes nothing', async () => {
    const dir = tmpDir('import-bad-template');
    const md = writeManuscript(dir);
    const out = path.join(dir, 'out.docx');
    const r = await runCli(['import', md, '-o', out, '--template', 'fancy']);
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/--template/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('fails closed with a line number on out-of-scope input and writes nothing', async () => {
    const dir = tmpDir('import-bad-block');
    const md = path.join(dir, 'bad.md');
    fs.writeFileSync(md, '# Ok\n\n> quoted aside\n');
    const out = path.join(dir, 'bad.docx');
    const r = await runCli(['import', md, '-o', out]);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_VALIDATION/);
    expect(r.stderr).toMatch(/line 3/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('rejects a missing manuscript with exit 1 and writes nothing', async () => {
    const dir = tmpDir('import-missing');
    const out = path.join(dir, 'missing.docx');
    const r = await runCli(['import', path.join(dir, 'nope.md'), '-o', out]);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_IO/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('refuses .html input with dedicated Markdown guidance and writes nothing', async () => {
    const dir = tmpDir('import-html-precheck');
    const html = path.join(dir, 'page.html');
    fs.writeFileSync(html, '<!DOCTYPE html><html><body><p>hi</p></body></html>');
    const out = path.join(dir, 'page.docx');
    const r = await runCli(['import', html, '-o', out]);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_VALIDATION/);
    expect(r.stderr).toMatch(/Markdown/);
    expect(r.stderr).toMatch(/\.md/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('refuses .htm input the same way; .md control still imports', async () => {
    const dir = tmpDir('import-htm-precheck');
    const htm = path.join(dir, 'page.htm');
    fs.writeFileSync(htm, '<html><body><p>hi</p></body></html>');
    const out = path.join(dir, 'page.docx');
    const r = await runCli(['import', htm, '-o', out]);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_VALIDATION/);
    expect(r.stderr).toMatch(/Markdown/);
    expect(fs.existsSync(out)).toBe(false);

    const md = writeManuscript(dir, 'control.md');
    const controlOut = path.join(dir, 'control.docx');
    const control = await runCli(['import', md, '-o', controlOut]);
    expect(control.status).toBe(0);
    expect(control.stdout).toContain('IMPORTED');
    expect(fs.existsSync(controlOut)).toBe(true);
  });

  it('refuses to overwrite without --force, clobbers with --force', async () => {
    const dir = tmpDir('import-force');
    const md = writeManuscript(dir);
    const out = path.join(dir, 'doc.docx');
    expect((await runCli(['import', md, '-o', out])).status).toBe(0);
    const before = fs.readFileSync(out);
    const refuse = await runCli(['import', md, '-o', out]);
    expect(refuse.status).toBe(2);
    expect(refuse.stderr).toContain(path.basename(out));
    expect(refuse.stderr).toMatch(/--force/);
    expect(fs.readFileSync(out).equals(before)).toBe(true);
    const clobber = await runCli(['import', md, '-o', out, '--force']);
    expect(clobber.status).toBe(0);
    expect(clobber.stdout).toContain('IMPORTED');
    expect((await runCli(['validate', out])).status).toBe(0);
  });

  it('import --help exits 0 with usage text', async () => {
    tmpDir('import-help');
    const r = await runCli(['import', '--help']);
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('Usage:');
    expect(r.stdout).toMatch(/-o/);
  });

  it('--json carries the import report payload', async () => {
    const dir = tmpDir('import-json');
    const md = writeManuscript(dir);
    const out = path.join(dir, 'out.docx');
    const r = await runCli(['--json', 'import', md, '-o', out]);
    expect(r.status).toBe(0);
    const payload = JSON.parse(r.stdout);
    expect(payload.ok).toBe(true);
    expect(payload.command).toBe('import');
    expect(payload.file).toBe(out);
    expect(payload.valid).toBe(true);
    expect(payload.appliedCount).toBeGreaterThan(0);
  });
});

describe('markdown round-trip fidelity', () => {
  it('docx exports carry every construct and re-import to a stable document', async () => {
    const dir = tmpDir('import-roundtrip');
    const md = writeManuscript(dir);
    const docx1 = path.join(dir, 'pass1.docx');
    expect((await runCli(['import', md, '-o', docx1])).status).toBe(0);
    expect((await runCli(['validate', docx1])).status).toBe(0);

    const md1Path = path.join(dir, 'pass1.md');
    expect((await runCli(['export', docx1, '--format', 'markdown', '-o', md1Path])).status).toBe(0);
    const md1 = fs.readFileSync(md1Path, 'utf8');

    // Fidelity: every construct survives the docx leg.
    expect(md1).toContain('# Field Notes on Endurance Training');
    expect(md1).toContain('## Weekly Structure');
    expect(md1).toContain('### Day One Focus');
    expect(md1).toContain('- warm up easy');
    expect(md1).toContain('1. first step');
    expect(md1).toContain('| Metric | Value |');
    expect(md1).toContain('| Endurance | 42 |');
    expect(md1).toMatch(/^---$/m);
    expect(md1).toContain('```');
    expect(md1).toContain('log line one');
    expect(md1).toContain('[^1]');
    expect(md1).toContain('[^1]: The footnote body.');

    // Stability: the exported Markdown re-imports clean and exports
    // identically (modulo whitespace and template seed duplication).
    const docx2 = path.join(dir, 'pass2.docx');
    const reimport = await runCli(['import', md1Path, '-o', docx2]);
    expect(reimport.status).toBe(0);
    expect((await runCli(['validate', docx2])).status).toBe(0);
    const md2Path = path.join(dir, 'pass2.md');
    expect((await runCli(['export', docx2, '--format', 'markdown', '-o', md2Path])).status).toBe(0);
    const md2 = fs.readFileSync(md2Path, 'utf8');
    expect(normalizeMd(md2)).toBe(normalizeMd(md1));
  });
});
