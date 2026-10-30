import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function writeBlankDocx(prefix = 'astra-cli-normalize'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
  return file;
}

/**
 * Minimal legacy file: a body paragraph carrying `w:shd` without the
 * required `w:val` (allowlist infill target). Built by textual surgery on a
 * blank template — no core knowledge beyond the broken input shape.
 */
function writeLegacyDocx(prefix = 'astra-cli-normalize-legacy'): string {
  const blank = DocxParser.createNewDocument('blank').rezip();
  const zip = new AdmZip(blank);
  const xml = zip.getEntry('word/document.xml')?.getData().toString('utf8');
  if (!xml) throw new Error('blank template has no document.xml');
  const sect = xml.indexOf('<w:sectPr');
  const anchor = sect >= 0 ? sect : xml.lastIndexOf('</w:body>');
  if (anchor < 0) throw new Error('blank template has no body anchor');
  const legacy =
    xml.slice(0, anchor) +
    '<w:p><w:pPr><w:shd w:fill="FF0000"/></w:pPr><w:r><w:t>legacy</w:t></w:r></w:p>' +
    xml.slice(anchor);
  zip.updateFile('word/document.xml', Buffer.from(legacy, 'utf8'));
  const file = tmpPath(prefix);
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

async function nodeIdForText(file: string, text: string): Promise<string> {
  const parser = new DocxParser(fs.readFileSync(file));
  const ast = await parser.getAST();
  const hits = parser.searchNodesInAST(ast, text);
  if (hits.length === 0) throw new Error(`no node contains "${text}"`);
  return (hits[0] as { id: string }).id;
}

function writeOps(payload: unknown): string {
  const file = tmpPath('astra-cli-normalize-ops', '.json');
  fs.writeFileSync(file, JSON.stringify(payload));
  return file;
}

async function firstParagraphId(file: string): Promise<string> {
  const parser = new DocxParser(fs.readFileSync(file));
  const ast = await parser.getAST();
  const nodes = ast.sections.flatMap((s) => s.nodes);
  const para = nodes.find((n) => n.type === 'paragraph') ?? nodes[0];
  if (!para) throw new Error('blank template has no nodes');
  return para.id;
}

function stdoutText(): string {
  return vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
}

function stdoutJson(): unknown[] {
  return vi.mocked(console.log).mock.calls.map((c) => {
    try {
      return JSON.parse(String(c[0]));
    } catch {
      return undefined;
    }
  });
}

function stderrText(): string {
  return vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
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
});

describe('batch result renders normalized[]', () => {
  it('BATCH OK human line carries normalized= and --json carries normalized:[...]', async () => {
    const file = writeBlankDocx('astra-norm-batch');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Normalized proof' }], sequential: true });
    const dest = tmpPath('astra-norm-batch-dest');
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--out', dest])).resolves.toBe(0);
    // Clean file: the core reports no rewrites — shape present, honest zero.
    expect(stdoutText()).toContain('BATCH OK');
    expect(stdoutText()).toContain('normalized=0');
    await expect(run(['--json', 'batch', file, ops, '--out', dest, '--force'])).resolves.toBe(0);
    const payloads = stdoutJson().filter((p) => p !== undefined) as Array<Record<string, unknown>>;
    const body = payloads.find((p) => p['ok'] === true && !p['dryRun'] && !p['matrix']);
    expect(body).toBeDefined();
    expect(Array.isArray(body?.['normalized'])).toBe(true);
  });

  it('dry-run lists would-be rewrites and carries normalized in --json', async () => {
    const file = writeBlankDocx('astra-norm-dry');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Dry normalize proof' }], sequential: true });
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--dry-run'])).resolves.toBe(0);
    expect(stdoutText()).toContain('DRY-RUN OK');
    expect(stdoutText()).toContain('would-be normalized=');
    await expect(run(['--json', 'batch', file, ops, '--dry-run'])).resolves.toBe(0);
    const payloads = stdoutJson().filter((p) => p !== undefined) as Array<Record<string, unknown>>;
    const body = payloads.find((p) => p['dryRun'] === true);
    expect(body).toBeDefined();
    expect(Array.isArray(body?.['normalized'])).toBe(true);
  });
});

describe('normalize command (repair-as-action)', () => {
  it('clean file reports normalized=0 and leaves bytes unchanged', async () => {
    const file = writeBlankDocx('astra-norm-clean');
    const before = fs.readFileSync(file);
    const dest = tmpPath('astra-norm-clean-dest');
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', file, '--out', dest])).resolves.toBe(0);
    expect(stdoutText()).toContain('NORMALIZED');
    expect(stdoutText()).toContain('normalized=0');
    expect(fs.readFileSync(dest).equals(before)).toBe(true);
  });

  it('--json carries normalized:[...] on a real run', async () => {
    const file = writeBlankDocx('astra-norm-json');
    const dest = tmpPath('astra-norm-json-dest');
    await expect(run(['--json', 'normalize', file, '--out', dest])).resolves.toBe(0);
    const payloads = stdoutJson().filter((p) => p !== undefined) as Array<Record<string, unknown>>;
    expect(payloads.length).toBeGreaterThan(0);
    const body = payloads[payloads.length - 1] as Record<string, unknown>;
    expect(body['ok']).toBe(true);
    expect(Array.isArray(body['normalized'])).toBe(true);
  });

  it('--dry-run lists would-be rewrites and writes nothing', async () => {
    const file = writeBlankDocx('astra-norm-dryrun');
    const before = fs.readFileSync(file);
    const dest = tmpPath('astra-norm-dryrun-dest');
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', file, '--out', dest, '--dry-run'])).resolves.toBe(0);
    expect(stdoutText()).toContain('DRY-RUN OK');
    expect(stdoutText()).toContain('would-be normalized=');
    expect(stdoutText()).toContain('writes nothing');
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('--help documents the command; top-level help lists it', async () => {
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', '--help'])).resolves.toBe(0);
    expect(stdoutText()).toContain('normalize');
    expect(stdoutText()).toContain('--dry-run');
    vi.mocked(console.log).mockClear();
    await expect(run(['help', 'normalize'])).resolves.toBe(0);
    expect(stdoutText()).toContain('NORMALIZED');
    vi.mocked(console.log).mockClear();
    await expect(run(['--help'])).resolves.toBe(0);
    expect(stdoutText()).toContain('normalize');
  });

  it('validate stays read-only: --fix is rejected as usage', async () => {
    const file = writeBlankDocx('astra-norm-nofix');
    vi.mocked(console.error).mockClear();
    await expect(run(['validate', file, '--fix'])).resolves.toBe(2);
    expect(stderrText()).toContain('unknown flag');
  });
});

describe('normalize command on a legacy file (non-empty normalized[])', () => {
  it('dry-run lists the would-be rewrite and writes nothing', async () => {
    const file = writeLegacyDocx();
    const before = fs.readFileSync(file);
    const dest = tmpPath('astra-norm-legacy-dest');
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', file, '--out', dest, '--dry-run'])).resolves.toBe(0);
    const out = stdoutText();
    expect(out).toContain('would-be normalized=1');
    expect(out).toContain('would-be normalize:');
    expect(out).toContain('w:shd');
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('real run heals, reports normalized=1, output validates', async () => {
    const file = writeLegacyDocx();
    const dest = tmpPath('astra-norm-legacy-healed');
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', file, '--out', dest])).resolves.toBe(0);
    const out = stdoutText();
    expect(out).toContain('NORMALIZED');
    expect(out).toContain('normalized=1');
    expect(out).toContain('w:shd');
    vi.mocked(console.log).mockClear();
    await expect(run(['validate', dest])).resolves.toBe(0);
    expect(stdoutText()).toContain('VALID');
  });

  it('batch on a legacy file heals and reports the rewrite on BATCH OK', async () => {
    const file = writeLegacyDocx();
    const target = await nodeIdForText(file, 'legacy');
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Healed proof' }], sequential: true });
    const dest = tmpPath('astra-norm-batch-legacy-dest');
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--out', dest])).resolves.toBe(0);
    const out = stdoutText();
    expect(out).toContain('BATCH OK');
    expect(out).toContain('normalized=1');
    expect(out).toContain('w:shd');
  });

  it('batch dry-run on a legacy file lists the would-be rewrite', async () => {
    const file = writeLegacyDocx();
    const before = fs.readFileSync(file);
    const target = await nodeIdForText(file, 'legacy');
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Healed dry proof' }], sequential: true });
    const dest = tmpPath('astra-norm-batch-legacy-dry-dest');
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--out', dest, '--dry-run'])).resolves.toBe(0);
    const out = stdoutText();
    expect(out).toContain('would-be normalized=1');
    expect(out).toContain('would-be normalize:');
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
});

describe('normalize backup snapshots the input before writing', () => {
  it('explicit backup path keeps the original bytes and restores them', async () => {
    const file = writeLegacyDocx('astra-norm-bak-legacy');
    const before = fs.readFileSync(file);
    const backup = tmpPath('astra-norm-bak-explicit', '.docx');
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', file, '--backup', backup])).resolves.toBe(0);
    expect(stdoutText()).toContain('BACKUP');
    expect(stdoutText()).toContain('NORMALIZED');
    expect(fs.existsSync(backup)).toBe(true);
    expect(fs.readFileSync(backup).equals(before)).toBe(true);
    expect(fs.readFileSync(file).equals(before)).toBe(false);
    fs.copyFileSync(backup, file);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('bare backup flag writes the default input backup copy', async () => {
    const file = writeLegacyDocx('astra-norm-bak-default');
    const before = fs.readFileSync(file);
    const backup = `${file}.bak.docx`;
    tmpFiles.push(backup);
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', file, '--backup'])).resolves.toBe(0);
    expect(fs.existsSync(backup)).toBe(true);
    expect(fs.readFileSync(backup).equals(before)).toBe(true);
  });

  it('backup with output target leaves the input alone and snapshots it', async () => {
    const file = writeLegacyDocx('astra-norm-bak-out');
    const before = fs.readFileSync(file);
    const dest = tmpPath('astra-norm-bak-out-dest');
    const backup = tmpPath('astra-norm-bak-out-copy', '.docx');
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', file, '--out', dest, '--backup', backup])).resolves.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    expect(fs.readFileSync(backup).equals(before)).toBe(true);
    expect(fs.existsSync(dest)).toBe(true);
    expect(fs.readFileSync(dest).equals(before)).toBe(false);
  });

  it('backup path equal to the output target fails closed with nothing written', async () => {
    const file = writeLegacyDocx('astra-norm-bak-self');
    const before = fs.readFileSync(file);
    const dest = tmpPath('astra-norm-bak-self-dest');
    vi.mocked(console.error).mockClear();
    await expect(run(['normalize', file, '--out', dest, `--backup=${dest}`])).resolves.toBe(1);
    expect(stderrText()).toContain('E_VALIDATION');
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('backup is skipped on dry run with a note and no backup file', async () => {
    const file = writeLegacyDocx('astra-norm-bak-dry');
    const before = fs.readFileSync(file);
    const backup = tmpPath('astra-norm-bak-dry-copy', '.docx');
    vi.mocked(console.error).mockClear();
    await expect(run(['normalize', file, '--dry-run', '--backup', backup])).resolves.toBe(0);
    expect(stderrText()).toContain('--backup skipped');
    expect(fs.existsSync(backup)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('json output carries the backup path on a real run', async () => {
    const file = writeLegacyDocx('astra-norm-bak-json');
    const backup = tmpPath('astra-norm-bak-json-copy', '.docx');
    await expect(run(['--json', 'normalize', file, '--backup', backup])).resolves.toBe(0);
    const payloads = stdoutJson().filter((p) => p !== undefined) as Array<Record<string, unknown>>;
    const body = payloads[payloads.length - 1] as Record<string, unknown>;
    expect(body['ok']).toBe(true);
    expect(body['backup']).toBe(backup);
  });

  it('help text documents the backup flag', async () => {
    vi.mocked(console.log).mockClear();
    await expect(run(['normalize', '--help'])).resolves.toBe(0);
    expect(stdoutText()).toContain('--backup');
  });
});

describe('normalize dry-run on syntax-corrupt input fails closed', () => {
  function writeSyntaxCorruptDocx(prefix = 'astra-norm-corrupt'): string {
    const blank = DocxParser.createNewDocument('blank').rezip();
    const zip = new AdmZip(blank);
    const xml = zip.getEntry('word/document.xml')?.getData().toString('utf8');
    if (!xml) throw new Error('blank template has no document.xml');
    const broken = xml.replace('</w:body>', '<w:p><w:r><w:t>broken');
    zip.updateFile('word/document.xml', Buffer.from(broken, 'utf8'));
    const file = tmpPath(prefix);
    fs.writeFileSync(file, zip.toBuffer());
    return file;
  }

  it('dry-run exits 1 with validation error and writes nothing', async () => {
    const file = writeSyntaxCorruptDocx();
    const before = fs.readFileSync(file);
    const dest = tmpPath('astra-norm-corrupt-dry-dest');
    vi.mocked(console.error).mockClear();
    await expect(run(['normalize', file, '--out', dest, '--dry-run'])).resolves.toBe(1);
    expect(stderrText()).toContain('E_VALIDATION');
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('real run on the same corrupt input also exits 1 (dry-run parity)', async () => {
    const file = writeSyntaxCorruptDocx('astra-norm-corrupt-real');
    const before = fs.readFileSync(file);
    const dest = tmpPath('astra-norm-corrupt-real-dest');
    vi.mocked(console.error).mockClear();
    await expect(run(['normalize', file, '--out', dest])).resolves.toBe(1);
    expect(stderrText()).toContain('E_VALIDATION');
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
});
