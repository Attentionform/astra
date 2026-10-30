import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { DocxParser } from '@attentionform/astra-engine';
import { remedyForPostBatchValidation } from '../src/cli/errors';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

// CLI exit-code / envelope / wording contract pins: bare-arg parity,
// --json usage envelopes, single-shape read failures, --out --force
// gating, and location-aware post-batch remedies.
const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function track(file: string): string {
  tmpFiles.push(file);
  return file;
}

function writeBlankDocx(prefix = 'astra-exit-contract'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
  return file;
}

function writeOps(payload: unknown): string {
  const file = tmpPath('astra-exit-contract-ops', '.json');
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

describe('bare subcommand parity (exit 2 like siblings)', () => {
  for (const cmd of ['inspect', 'validate', 'sanitize', 'batch', 'normalize', 'export', 'search', 'create', 'import', 'stats']) {
    it(`bare ${cmd} exits 2 with a usage error`, async () => {
      vi.mocked(console.error).mockClear();
      await expect(run([cmd])).resolves.toBe(2);
      expect(stderrText()).toMatch(/usage error/i);
    });
  }
  it('bare export --json exits 2 with a coded envelope and quiet stderr', async () => {
    vi.mocked(console.log).mockClear();
    vi.mocked(console.error).mockClear();
    await expect(run(['--json', 'export'])).resolves.toBe(2);
    const payload = JSON.parse(stdoutText());
    expect(payload.ok).toBe(false);
    expect(payload.code).toBe('E_USAGE');
    expect(typeof payload.remedy).toBe('string');
    expect(vi.mocked(console.error).mock.calls.length).toBe(0);
  });
});

describe('--json usage envelope', () => {
  it('usage errors carry code E_USAGE plus a remedy on stdout', async () => {
    vi.mocked(console.log).mockClear();
    vi.mocked(console.error).mockClear();
    await expect(run(['--json', 'batch'])).resolves.toBe(2);
    const payload = JSON.parse(stdoutText());
    expect(payload.ok).toBe(false);
    expect(payload.code).toBe('E_USAGE');
    expect(typeof payload.error).toBe('string');
    expect(typeof payload.remedy).toBe('string');
    expect(payload.remedy.length).toBeGreaterThan(0);
  });
  it('stderr stays quiet under --json', async () => {
    vi.mocked(console.error).mockClear();
    await expect(run(['--json', 'batch'])).resolves.toBe(2);
    expect(vi.mocked(console.error).mock.calls.length).toBe(0);
  });
  it('--verbose keeps the human dump under --json', async () => {
    vi.mocked(console.log).mockClear();
    vi.mocked(console.error).mockClear();
    await expect(run(['--json', '--verbose', 'batch'])).resolves.toBe(2);
    expect(JSON.parse(stdoutText()).code).toBe('E_USAGE');
    expect(stderrText()).toMatch(/usage error/i);
  });
  it('human usage output is unchanged (one-liner plus dump)', async () => {
    vi.mocked(console.error).mockClear();
    await expect(run(['batch'])).resolves.toBe(2);
    expect(stderrText()).toMatch(/usage error: batch requires/);
    expect(stderrText()).toContain('Usage: astra');
  });
});

describe('single-shape read failures', () => {
  it('missing input is one wording across validate, batch, and export', async () => {
    // Relative names: the one-liner scrubs absolute paths to <path> by
    // hygiene design, so relative names prove the file is still named.
    const ghost = `ghost-${process.pid}-${Date.now()}.docx`;
    const ops = writeOps({ operations: [] });
    for (const args of [['validate', ghost], ['batch', ghost, ops], ['export', ghost], ['normalize', ghost]] as string[][]) {
      vi.mocked(console.error).mockClear();
      const code = await run(args);
      expect(code).toBe(1);
      const line = stderrText().trim();
      expect(line).toContain('[E_IO]');
      expect(line).toContain(`cannot read "${ghost}": no such file or directory`);
      expect(line).toContain('Remedy: check the path exists and is readable, then retry');
      expect(line).not.toContain('re-inspect');
      expect(line.trim().split('\n')).toHaveLength(1);
    }
  });
  it('missing ops file names the ops path in the same shape', async () => {
    const file = writeBlankDocx('astra-exit-missing-ops');
    const missing = `missing-ops-${process.pid}-${Date.now()}.json`;
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, missing])).resolves.toBe(1);
    const line = stderrText().trim();
    expect(line).toContain('[E_IO]');
    expect(line).toContain(`cannot read "${missing}": no such file or directory`);
    expect(line).not.toMatch(/ENOENT/);
  });
  it('directory input is E_IO with a directory remedy, never a raw EISDIR', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-exit-isdir-'));
    const ops = writeOps({ operations: [] });
    for (const args of [['validate', dir], ['batch', dir, ops], ['export', dir]] as string[][]) {
      vi.mocked(console.error).mockClear();
      const code = await run(args);
      expect(code).toBe(1);
      const line = stderrText().trim();
      expect(line).toContain('[E_IO]');
      // Wording shape is uniform; absolute names may scrub to <path> on
      // paths that re-clean (hygiene layer), relative names echo literally.
      expect(line).toMatch(/cannot read ".+": is a directory/);
      expect(line).toContain('not a directory');
      expect(line).not.toMatch(/EISDIR/);
      expect(line).not.toContain('re-inspect');
    }
    fs.rmdirSync(dir);
  });
  it('corrupt input is one clean package wording on every path, no raw zip text', async () => {
    const bad = tmpPath('astra-exit-corrupt');
    fs.writeFileSync(bad, 'this is not a zip package');
    const ops = writeOps({ operations: [{ action: 'delete', target: 'nope' }] });
    for (const args of [['validate', bad], ['batch', bad, ops], ['export', bad], ['normalize', bad]] as string[][]) {
      vi.mocked(console.error).mockClear();
      const code = await run(args);
      expect(code).toBe(1);
      const line = stderrText().trim();
      expect(line).toContain('[E_VALIDATION]');
      expect(line).toMatch(/".+" is not a valid \.docx package/);
      expect(line).toContain('Remedy: re-export the file from Word');
      expect(line).not.toMatch(/ADM-ZIP/i);
      expect(line).not.toMatch(/No END header/);
      expect(line).not.toContain('re-inspect');
    }
  });
  it('absolute missing paths redact to <path> on every command, never the real path', async () => {
    const ghost = tmpPath('astra-exit-abs-ghost');
    const ops = writeOps({ operations: [] });
    for (const args of [['validate', ghost], ['batch', ghost, ops], ['export', ghost], ['normalize', ghost]] as string[][]) {
      vi.mocked(console.error).mockClear();
      const code = await run(args);
      expect(code).toBe(1);
      const line = stderrText().trim();
      expect(line).toContain('[E_IO]');
      expect(line).toContain('cannot read "<path>": no such file or directory');
      expect(line).not.toContain(ghost);
      expect(line).toContain('Remedy: check the path exists and is readable, then retry');
      expect(line.trim().split('\n')).toHaveLength(1);
    }
  });
  it('normalize dry-run on a non-zip input fails closed with the same clean package wording', async () => {
    const bad = tmpPath('astra-exit-normalize-dry-corrupt');
    fs.writeFileSync(bad, 'this is not a zip package');
    const before = fs.readFileSync(bad);
    vi.mocked(console.error).mockClear();
    await expect(run(['normalize', bad, '--dry-run'])).resolves.toBe(1);
    const line = stderrText().trim();
    expect(line).toContain('[E_VALIDATION]');
    expect(line).toMatch(/".+" is not a valid \.docx package/);
    expect(line).toContain('Remedy: re-export the file from Word');
    expect(line).not.toMatch(/ADM-ZIP/i);
    expect(line).not.toMatch(/No END header/);
    expect(line.trim().split('\n')).toHaveLength(1);
    expect(fs.readFileSync(bad).equals(before)).toBe(true);
  });
});

describe('--out refuses without --force', () => {
  it('batch --out on an existing file exits 2 naming the file and the flag', async () => {
    const file = writeBlankDocx('astra-exit-batch-gate');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Gate proof' }] });
    const dest = tmpPath('astra-exit-batch-dest');
    fs.writeFileSync(dest, before);
    const destBefore = fs.readFileSync(dest);
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--out', dest])).resolves.toBe(2);
    expect(stderrText()).toContain(path.basename(dest));
    expect(stderrText()).toMatch(/--force/);
    expect(fs.readFileSync(dest).equals(destBefore)).toBe(true);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
  it('batch --out with --force overwrites', async () => {
    const file = writeBlankDocx('astra-exit-batch-force');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Force proof' }] });
    const dest = tmpPath('astra-exit-batch-force-dest');
    fs.writeFileSync(dest, fs.readFileSync(file));
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--out', dest, '--force'])).resolves.toBe(0);
    expect(stdoutText()).toContain('BATCH OK');
  });
  it('batch dry-run with an existing --out writes nothing and exits 0', async () => {
    const file = writeBlankDocx('astra-exit-batch-dry-gate');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Dry gate proof' }] });
    const dest = tmpPath('astra-exit-batch-dry-dest');
    fs.writeFileSync(dest, fs.readFileSync(file));
    await expect(run(['batch', file, ops, '--out', dest, '--dry-run'])).resolves.toBe(0);
  });
  it('matrix pre-existing row target refuses before any write', async () => {
    const file = writeBlankDocx('astra-exit-matrix-gate');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Hi {{name}}' }] });
    const csv = tmpPath('astra-exit-matrix-vars', '.csv');
    fs.writeFileSync(csv, 'name\nAda\nBob\n');
    const template = tmpPath('astra-exit-matrix-out-{{name}}');
    const adaOut = template.replace('{{name}}', 'Ada');
    track(adaOut);
    fs.writeFileSync(adaOut, before);
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--matrix', csv, '--out', template])).resolves.toBe(2);
    expect(stderrText()).toMatch(/--force/);
    expect(fs.existsSync(template.replace('{{name}}', 'Bob'))).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
  it('export -o on an existing file exits 2, then succeeds with --force', async () => {
    const file = writeBlankDocx('astra-exit-export-gate');
    const dest = tmpPath('astra-exit-export-dest', '.html');
    fs.writeFileSync(dest, 'stale');
    vi.mocked(console.error).mockClear();
    await expect(run(['export', file, '-o', dest])).resolves.toBe(2);
    expect(stderrText()).toContain(path.basename(dest));
    expect(stderrText()).toMatch(/--force/);
    expect(fs.readFileSync(dest, 'utf8')).toBe('stale');
    vi.mocked(console.log).mockClear();
    await expect(run(['export', file, '-o', dest, '--force'])).resolves.toBe(0);
    expect(fs.readFileSync(dest, 'utf8')).not.toBe('stale');
  });
  it('batch and export --help disclose --force', async () => {
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', '--help'])).resolves.toBe(0);
    expect(stdoutText()).toContain('--force');
    vi.mocked(console.log).mockClear();
    await expect(run(['export', '--help'])).resolves.toBe(0);
    expect(stdoutText()).toContain('--force');
  });
});

describe('location-aware post-batch remedies', () => {
  it('surfaces the validator wording for an unknown element', async () => {
    const file = writeBlankDocx('astra-exit-remedy');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({
      operations: [{ action: 'insert', type: 'xml', anchor: target, position: 'after', xml: '<w:p><w:bogus/></w:p>' }],
    });
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    const line = stderrText().trim();
    expect(line).toContain('[E_VALIDATION]');
    expect(line).toContain('invalid child element');
    expect(line).toContain('bogus');
    expect(line).toContain('Element not valid at this position');
    expect(line).toContain('astra batch');
    expect(line).not.toMatch(/hex\/half-point\/twips/);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
  it('dry-run gate uses the same contextual remedy', async () => {
    const file = writeBlankDocx('astra-exit-remedy-dry');
    const target = await firstParagraphId(file);
    const ops = writeOps({
      operations: [{ action: 'insert', type: 'xml', anchor: target, position: 'after', xml: '<w:tbl><w:tr><w:tc><w:p><w:t>x</w:t></w:p></w:tc></w:tr></w:tbl>' }],
    });
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--dry-run'])).resolves.toBe(1);
    expect(stderrText()).toContain('Table structure invalid');
    expect(stderrText()).toContain('astra batch');
    expect(stderrText()).not.toMatch(/hex\/half-point\/twips/);
  });
  it('remedy branches mirror the save-gate wording per failure shape', () => {
    expect(remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:p[3]', 'unknown element <w:bogus> in word/document.xml'))
      .toContain('Non-OOXML markup');
    expect(remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:tbl[2]/w:tblPr', 'invalid child'))
      .toContain('Table structure invalid');
    expect(remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 'duplicate w14:paraId 3A86481C'))
      .toContain('Duplicate paragraph id');
    expect(remedyForPostBatchValidation('/w:ftr[1]/w:p[2]', "'data-astra-id' not declared"))
      .toContain('footer part');
    expect(remedyForPostBatchValidation('/w:hdr[1]/w:p[1]', "'data-astra-id' not declared"))
      .toContain('header part');
    expect(remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:p[4]/w:r[1]/w:drawing[1]/wp:docPr', 'duplicate docPr id'))
      .toContain('wp:docPr');
    expect(remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:p[1]', 'relationship rId9 missing'))
      .toContain('relationships invalid');
    expect(remedyForPostBatchValidation('/w:ftr[1]/w:p[1]', 'invalid child in footer'))
      .toContain('Header/footer part invalid');
    // The validator's own wording for an element that is not allowed here.
    expect(
      remedyForPostBatchValidation(
        '/w:document[1]/w:body[1]/w:p[3]/w:r[1]',
        "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:bogus'."
      )
    ).toContain('Element not valid at this position (/w:document[1]/w:body[1]/w:p[3]/w:r[1])');
    // Table and header/footer locations keep their own, more specific remedies.
    expect(remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]', 'The element has invalid child element'))
      .toContain('Table structure invalid');
    expect(remedyForPostBatchValidation('/w:ftr[1]/w:p[1]', 'The element has invalid child element'))
      .toContain('Header/footer part invalid');
    // Unknown shapes keep the long-standing generic values hint byte-identical.
    expect(remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:sz', 'bad sz value'))
      .toBe('fix hex/half-point/twips values and retry');
    // Every contextual branch routes the fix through the CLI.
    for (const remedy of [
      remedyForPostBatchValidation('/w:document[1]/w:body[1]/w:tbl[1]', 'invalid child'),
      remedyForPostBatchValidation('x', 'unknown element <w:bogus>'),
    ]) {
      expect(remedy).toContain('astra batch');
    }
  });
});
