import { describe, it, expect, vi, beforeAll, beforeEach, afterAll, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip } from '@attentionform/astra-engine';
import { DOMParser } from '@xmldom/xmldom';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

// CLI DX coverage: in-process via run(argv). Gallery fixtures
// are copied into tmpdirs (read-only); nothing writes outside the tmpdir.
const REPO_ROOT = path.resolve(__dirname, '..', '..', '..');
const GALLERY_T2 = path.join(REPO_ROOT, 'test-gallery', 'T2_edit_target.docx');

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

function copyGallery(dir: string, name = 'doc.docx'): string {
  const dest = path.join(dir, name);
  fs.copyFileSync(GALLERY_T2, dest);
  return dest;
}

function docXml(file: string): string {
  const zip = new AdmZip(fs.readFileSync(file));
  return zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

/** Dirty but schema-valid doc: run-level ins/del + a real comment part. */
async function writeDirtyDocx(dest: string): Promise<void> {
  const blank = path.join(path.dirname(dest), 'blank-src.docx');
  const created = await runCli(['create', blank, '--template', 'blank']);
  expect(created.status).toBe(0);
  const zip = new AdmZip(fs.readFileSync(blank));
  let d = zip.getEntry('word/document.xml')!.getData().toString('utf8');
  d = d.replace(
    '<w:t>Blank Document</w:t>',
    '<w:t>Blank Document</w:t></w:r><w:ins w:id="0" w:author="EvilAuthor" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve"> sneaky insert</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> tail</w:t>'
  );
  d = d.replace(
    '<w:t>Start typing your content here...</w:t>',
    '<w:t>Reviewed body</w:t></w:r><w:del w:id="1" w:author="EvilAuthor" w:date="2026-01-01T00:00:00Z"><w:r><w:t> removed words</w:t></w:r></w:del><w:r><w:t xml:space="preserve"> end</w:t>'
  );
  d = d.replace(
    '<w:sectPr>',
    '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented line</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p><w:sectPr>'
  );
  d = d.replace(/<w:p /g, '<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3" ');
  zip.updateFile('word/document.xml', Buffer.from(d));
  zip.addFile(
    'word/comments.xml',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Spy" w:initials="S" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>leak me</w:t></w:r></w:p></w:comment></w:comments>'
    )
  );
  let ct = zip.getEntry('[Content_Types].xml')!.getData().toString('utf8');
  ct = ct.replace(
    '</Types>',
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'
  );
  zip.updateFile('[Content_Types].xml', Buffer.from(ct));
  let rels = zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
  rels = rels.replace(
    '</Relationships>',
    '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'
  );
  zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels));
  fs.writeFileSync(dest, zip.toBuffer());
}

async function firstNodeId(file: string): Promise<string> {
  const inspected = await runCli(['--json', 'inspect', file]);
  expect(inspected.status).toBe(0);
  const payload = JSON.parse(inspected.stdout);
  const m = /^([0-9A-F]+)\s*\|/m.exec(String(payload.toon));
  expect(m, 'inspect --json toon yields an anchor id').toBeTruthy();
  return m![1];
}

beforeAll(() => {
  expect(fs.existsSync(GALLERY_T2), `gallery fixture missing at ${GALLERY_T2}`).toBe(true);
});

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
});

afterAll(() => {
  for (const dir of dirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe('per-command help exits 0', () => {
  const commands = ['inspect', 'validate', 'sanitize', 'batch', 'export', 'search', 'create', 'stats'];
  for (const cmd of commands) {
    it(`${cmd} --help exits 0 with usage text`, async () => {
      tmpDir('dx-help');
      const r = (await runCli([cmd, '--help']));
      expect(r.status).toBe(0);
      expect(r.stdout.toLowerCase()).toContain('usage:');
    });
  }
  it('help topic routing answers per command', async () => {
    tmpDir('dx-help-topic');
    const r = (await runCli(['help', 'search']));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('search');
  });
  it('batch --schema prints canonical JSON per action', async () => {
    tmpDir('dx-schema');
    const r = (await runCli(['batch', '--schema']));
    expect(r.status).toBe(0);
    const payload = JSON.parse(r.stdout);
    for (const key of ['replace', 'insert_paragraph', 'insert_heading_cli', 'namedStyle_apply', 'delete', 'clearDoc_cli', 'table_updateCell', 'table_deleteRow', 'chart_insert', 'chart_update', 'comment_insert', 'comment_reply']) {
      expect(payload[key], `schema example for ${key}`).toBeTruthy();
    }
    expect(payload.comment_insert.operation).toBe('insert');
    expect(payload.chart_update.operation).toBe('update');
    expect(payload.chart_update.target).toBe('<chartNodeId>');
  });
});

describe('sanitize metadata preset default', () => {
  it('bare sanitize strips revisions+comments (zero w:ins/w:del/comment) and validates', async () => {
    const dir = tmpDir('dx-san-bare');
    const dirty = path.join(dir, 'dirty.docx');
    await writeDirtyDocx(dirty);
    expect((await runCli(['validate', dirty])).status).toBe(0);
    expect(docXml(dirty)).toMatch(/w:rsid/);
    const out = path.join(dir, 'clean.docx');
    const r = (await runCli(['sanitize', dirty, out]));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('SANITIZED');
    expect(r.stdout).toMatch(/rsids=[1-9]\d*/);
    expect(r.stdout).toContain('bodyRedactions=0');
    expect(r.stderr).toMatch(/WARNING/);
    expect(r.stderr).toMatch(/body PII survives/);
    expect(r.stderr).toMatch(/--redact-pii/);
    const xml = docXml(out);
    expect(xml).not.toContain('w:ins');
    expect(xml).not.toContain('<w:del');
    expect(xml).not.toContain('commentReference');
    expect(xml).not.toContain('EvilAuthor');
    expect(xml).not.toMatch(/w:rsid/);
    const zip = new AdmZip(fs.readFileSync(out));
    expect(zip.getEntry('word/comments.xml')).toBeNull();
    expect((await runCli(['validate', out])).status).toBe(0);
  });
  it('client-safe bundle shares metadata cleaning with bare on PII-free input', async () => {
    const dir = tmpDir('dx-san-preset');
    const dirty = path.join(dir, 'dirty.docx');
    await writeDirtyDocx(dirty);
    const bare = path.join(dir, 'bare.docx');
    const preset = path.join(dir, 'preset.docx');
    const b = (await runCli(['sanitize', dirty, bare]));
    expect(b.status).toBe(0);
    const r = (await runCli(['sanitize', dirty, preset, '--client-safe']));
    expect(r.status).toBe(0);
    const rsidsOf = (line: string): number => {
      const m = /rsids=(\d+)/.exec(line);
      expect(m, 'SANITIZED line carries rsids count').toBeTruthy();
      return Number(m![1]);
    };
    expect(rsidsOf(b.stdout)).toBeGreaterThan(0);
    expect(rsidsOf(r.stdout)).toBe(rsidsOf(b.stdout));
    const countRevisions = async (f: string): Promise<number> => {
      const m = await runCli(['sanitize', f, path.join(dir, 'probe.docx'), '--client-safe']);
      return m.status;
    };
    expect(await countRevisions(bare)).toBe(0);
    expect(await countRevisions(preset)).toBe(0);
    for (const f of [bare, preset]) {
      const xml = docXml(f);
      expect(xml).not.toContain('w:ins');
      expect(xml).not.toContain('commentReference');
      expect(xml).not.toMatch(/w:rsid/);
      expect((await runCli(['validate', f])).status).toBe(0);
    }
    expect(r.stdout).toContain('SANITIZED');
  });
  it('partial flags warn loudly naming survivors and keep them', async () => {
    const dir = tmpDir('dx-san-partial');
    const dirty = path.join(dir, 'dirty.docx');
    await writeDirtyDocx(dirty);
    const out = path.join(dir, 'partial.docx');
    const r = (await runCli(['sanitize', dirty, out, '--zero-retention']));
    expect(r.status).toBe(0);
    expect(r.stderr).toMatch(/WARNING/);
    expect(r.stderr).toMatch(/revision/i);
    expect(r.stderr).toMatch(/comment/i);
    expect(docXml(out)).toContain('w:ins');
  });
  it('--json partial carries warnings array, full-clean omits it', async () => {
    const dir = tmpDir('dx-san-json-warn');
    const dirty = path.join(dir, 'dirty.docx');
    await writeDirtyDocx(dirty);
    const partialOut = path.join(dir, 'partial.docx');
    const p = (await runCli(['--json', 'sanitize', dirty, partialOut, '--zero-retention']));
    expect(p.status).toBe(0);
    const pp = JSON.parse(p.stdout);
    expect(pp.ok).toBe(true);
    expect(Array.isArray(pp.warnings)).toBe(true);
    expect(pp.warnings.join(' ')).toMatch(/revision/i);
    expect(pp.warnings.join(' ')).toMatch(/comment/i);
    expect(docXml(partialOut)).toContain('w:ins');
    const cleanOut = path.join(dir, 'clean.docx');
    const c = (await runCli(['--json', 'sanitize', dirty, cleanOut, '--client-safe']));
    expect(c.status).toBe(0);
    const cp = JSON.parse(c.stdout);
    expect(cp.ok).toBe(true);
    expect(cp.warnings ?? []).toHaveLength(0);
    expect(cp.warnings).toBeUndefined();
    expect(docXml(cleanOut)).not.toContain('w:ins');
  });
  it('dangling --replacement-author exits 2 (never silent preset)', async () => {
    const dir = tmpDir('dx-san-author-dangle');
    const dirty = path.join(dir, 'dirty.docx');
    await writeDirtyDocx(dirty);
    const out1 = path.join(dir, 'out1.docx');
    const r1 = (await runCli(['sanitize', dirty, out1, '--replacement-author']));
    expect(r1.status).toBe(2);
    expect(r1.stderr).toMatch(/replacement-author/i);
    const out2 = path.join(dir, 'out2.docx');
    const r2 = (await runCli(['sanitize', dirty, out2, '--replacement-author', '--client-safe']));
    expect(r2.status).toBe(2);
    expect(r2.stderr).toMatch(/replacement-author/i);
  });
  it('sanitize --help documents --json warnings', async () => {
    tmpDir('dx-san-help-json');
    const r = (await runCli(['sanitize', '--help']));
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/warnings/);
  });
});

describe('sanitize body text kept by default, redacted on demand', () => {
  const INVOICE_EMAIL = 'billing@example.com';
  async function writeInvoiceDocx(dest: string): Promise<void> {
    await writeDirtyDocx(dest);
    const zip = new AdmZip(fs.readFileSync(dest));
    const d = zip.getEntry('word/document.xml')!.getData().toString('utf8');
    expect(d).toContain('Blank Document');
    zip.updateFile('word/document.xml', Buffer.from(d.replace('Blank Document', `Invoice INV-001 Contact ${INVOICE_EMAIL}`)));
    fs.writeFileSync(dest, zip.toBuffer());
  }
  it('bare keeps body email but strips authors revisions comments', async () => {
    const dir = tmpDir('dx-san-invoice-bare');
    const invoice = path.join(dir, 'invoice.docx');
    await writeInvoiceDocx(invoice);
    const out = path.join(dir, 'out.docx');
    const r = (await runCli(['sanitize', invoice, out]));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('SANITIZED');
    expect(r.stdout).toContain('bodyRedactions=0');
    expect(r.stderr).toMatch(/WARNING/);
    expect(r.stderr).toMatch(/body PII survives/);
    const xml = docXml(out);
    expect(xml).toContain(INVOICE_EMAIL);
    expect(xml).not.toContain('EvilAuthor');
    expect(xml).not.toContain('w:ins');
    expect(xml).not.toContain('commentReference');
    expect(xml).not.toMatch(/w:rsid/);
    expect((await runCli(['validate', out])).status).toBe(0);
  });
  it('client-safe bundle redacts body email and discloses counts', async () => {
    const dir = tmpDir('dx-san-invoice-safe');
    const invoice = path.join(dir, 'invoice.docx');
    await writeInvoiceDocx(invoice);
    const out = path.join(dir, 'out.docx');
    const r = (await runCli(['sanitize', invoice, out, '--client-safe']));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('SANITIZED');
    expect(r.stdout).toMatch(/bodyRedactions=[1-9]\d*/);
    expect(r.stdout).toContain('bodyLocations=');
    expect(r.stderr).not.toMatch(/WARNING/);
    const xml = docXml(out);
    expect(xml).not.toContain(INVOICE_EMAIL);
    expect(xml).not.toContain('EvilAuthor');
    expect(xml).not.toMatch(/w:rsid/);
    expect((await runCli(['validate', out])).status).toBe(0);
  });
  it('lone redact flag redacts body email with the metadata preset', async () => {
    const dir = tmpDir('dx-san-invoice-redact');
    const invoice = path.join(dir, 'invoice.docx');
    await writeInvoiceDocx(invoice);
    const out = path.join(dir, 'out.docx');
    const r = (await runCli(['sanitize', invoice, out, '--redact-pii']));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('SANITIZED');
    expect(r.stdout).toMatch(/bodyRedactions=[1-9]\d*/);
    expect(r.stdout).toContain('bodyLocations=');
    expect(r.stderr).not.toMatch(/WARNING/);
    const xml = docXml(out);
    expect(xml).not.toContain(INVOICE_EMAIL);
    expect(xml).not.toContain('EvilAuthor');
    expect(xml).not.toContain('w:ins');
    expect(xml).not.toContain('<w:del');
    expect(xml).not.toContain('commentReference');
    expect(xml).not.toMatch(/w:rsid/);
    expect((await runCli(['validate', out])).status).toBe(0);
    const jsonOut = path.join(dir, 'out-json.docx');
    const j = (await runCli(['--json', 'sanitize', invoice, jsonOut, '--redact-pii']));
    expect(j.status).toBe(0);
    const payload = JSON.parse(j.stdout);
    expect(payload.ok).toBe(true);
    expect(payload.warnings).toBeUndefined();
    expect(payload.report.bodyPII).toBeGreaterThan(0);
    expect(docXml(jsonOut)).not.toContain(INVOICE_EMAIL);
  });
  it('redact plus metadata-only keeps body text and names the remedy', async () => {
    const dir = tmpDir('dx-san-invoice-redact-meta');
    const invoice = path.join(dir, 'invoice.docx');
    await writeInvoiceDocx(invoice);
    const out = path.join(dir, 'out.docx');
    const r = (await runCli(['sanitize', invoice, out, '--redact-pii', '--metadata-only']));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('SANITIZED');
    expect(r.stdout).toContain('bodyRedactions=0');
    expect(docXml(out)).toContain(INVOICE_EMAIL);
    expect(r.stderr).toMatch(/WARNING/);
    expect(r.stderr).toMatch(/body PII survives/);
    expect(r.stderr).toMatch(/--metadata-only/);
    const j = (await runCli(['--json', 'sanitize', invoice, path.join(dir, 'j.docx'), '--redact-pii', '--metadata-only']));
    expect(j.status).toBe(0);
    const payload = JSON.parse(j.stdout);
    expect(payload.warnings.join(' ')).toMatch(/body PII survives/);
    expect(payload.warnings.join(' ')).toMatch(/--metadata-only/);
  });
  it('metadata-only alias matches bare and keeps body email', async () => {
    const dir = tmpDir('dx-san-invoice-alias');
    const invoice = path.join(dir, 'invoice.docx');
    await writeInvoiceDocx(invoice);
    const bareOut = path.join(dir, 'bare.docx');
    const aliasOut = path.join(dir, 'alias.docx');
    expect((await runCli(['sanitize', invoice, bareOut])).status).toBe(0);
    const r = (await runCli(['sanitize', invoice, aliasOut, '--metadata-only']));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('SANITIZED');
    expect(r.stdout).toContain('bodyRedactions=0');
    expect(docXml(aliasOut)).toContain(INVOICE_EMAIL);
    expect(docXml(bareOut)).toContain(INVOICE_EMAIL);
    expect((await runCli(['validate', aliasOut])).status).toBe(0);
  });
  it('sanitize help documents bare kept body plus redact modes', async () => {
    tmpDir('dx-san-help-modes');
    const r = (await runCli(['sanitize', '--help']));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('--redact-pii');
    expect(r.stdout).toContain('--client-safe');
    expect(r.stdout).toContain('--metadata-only');
    expect(r.stdout.toLowerCase()).toContain('body text kept');
    expect(r.stdout).toMatch(/always warn/);
    expect(r.stdout).toMatch(/fully redacted and clean/);
    expect(r.stdout).toContain('plus the metadata preset');
  });
});

describe('error hygiene', () => {
  it('ENOENT is one line [E_IO] + Remedy, no frames/paths', async () => {
    const dir = tmpDir('dx-enoent');
    const r = (await runCli(['validate', path.join(dir, 'ghost.docx')]));
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/\[E_IO\]/);
    expect(r.stderr).toMatch(/Remedy:/);
    expect(r.stderr.trim().split('\n')).toHaveLength(1);
    expect(r.stderr).not.toMatch(/^\s*at\s/m);
    expect(r.stderr).not.toContain('dist/');
  });
  it('corrupt zip is [E_VALIDATION] with a remedy', async () => {
    const dir = tmpDir('dx-corrupt');
    const bad = path.join(dir, 'corrupt.docx');
    fs.writeFileSync(bad, 'this is not a zip package');
    const r = (await runCli(['validate', bad]));
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/\[E_VALIDATION\]/);
    expect(r.stderr).toMatch(/Remedy:/);
    expect(r.stderr.trim().split('\n')).toHaveLength(1);
    expect(r.stderr).not.toMatch(/^\s*at\s/m);
  });
  it('--json errors are clean strings with a code field', async () => {
    const dir = tmpDir('dx-json-err');
    const r = (await runCli(['--json', 'validate', path.join(dir, 'ghost.docx')]));
    expect(r.status).toBe(1);
    const payload = JSON.parse(r.stdout);
    expect(payload.ok).toBe(false);
    expect(payload.code).toBe('E_IO');
    expect(String(payload.error)).not.toMatch(/^\s*at\s/m);
    expect(String(payload.error)).not.toContain('dist/');
  });
  it('batch failure rewrites MCP remedies to astra inspect, stays one line', async () => {
    const dir = tmpDir('dx-remedy');
    const file = copyGallery(dir);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'replace', target: 'ghost_node_xyz', text: 'x' }] }));
    const r = (await runCli(['batch', file, ops]));
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('astra inspect');
    expect(r.stderr).not.toContain('get_document_toon');
    expect(r.stderr).not.toContain('toon({');
    expect(r.stderr.trim().split('\n')).toHaveLength(1);
    expect(r.stderr).not.toMatch(/^\s*at\s/m);
    expect(r.stderr).not.toContain('dist/');
    expect(r.stderr).not.toContain('🚨');
  });
  it('--verbose keeps raw stacks on stderr', async () => {
    const dir = tmpDir('dx-verbose');
    const file = copyGallery(dir);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'replace', target: 'ghost_node_xyz', text: 'x' }] }));
    const r = (await runCli(['--verbose', 'batch', file, ops]));
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/\[E_VALIDATION\]/);
    expect(r.stderr).toMatch(/at\s/);
  });
  it('search --limit rejects garbage with usage exit 2', async () => {
    const dir = tmpDir('dx-limit');
    const file = copyGallery(dir);
    expect((await runCli(['search', file, 'q', '--limit', 'bogus'])).status).toBe(2);
  });
  it('search trailing --limit exits 2', async () => {
    const dir = tmpDir('dx-limit-dangle');
    const file = copyGallery(dir);
    const r = (await runCli(['search', file, 'q', '--limit']));
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/--limit/i);
  });
  it('search --limit with flag-like value exits 2', async () => {
    const dir = tmpDir('dx-limit-flag');
    const file = copyGallery(dir);
    const r = (await runCli(['search', file, 'q', '--limit', '--foo']));
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/--limit/i);
    const r2 = (await runCli(['search', file, 'q', '--limit', '-5']));
    expect(r2.status).toBe(2);
    expect(r2.stderr).toMatch(/--limit/i);
  });
});

describe('search', () => {
  it('finds snippet + node id lines', async () => {
    const dir = tmpDir('dx-search');
    const file = copyGallery(dir);
    const r = (await runCli(['search', file, 'accountability', '--limit', '3']));
    expect(r.status).toBe(0);
    const lines = r.stdout.trim().split('\n');
    expect(lines.length).toBeGreaterThan(1);
    expect(lines[0]).toMatch(/[0-9A-F]{4,} .*accountability/i);
    expect(r.stdout).toMatch(/match(es)?/);
  });
  it('--json carries hits with id + snippet', async () => {
    const dir = tmpDir('dx-search-json');
    const file = copyGallery(dir);
    const r = (await runCli(['--json', 'search', file, 'accountability']));
    expect(r.status).toBe(0);
    const payload = JSON.parse(r.stdout);
    expect(payload.count).toBeGreaterThan(0);
    expect(payload.hits[0].id).toMatch(/[0-9A-F]+/);
    expect(String(payload.hits[0].snippet).toLowerCase()).toContain('accountability');
  });
  it('zero matches still exits 0', async () => {
    const dir = tmpDir('dx-search-empty');
    const file = copyGallery(dir);
    const r = (await runCli(['search', file, 'zzz-no-such-token-zzz']));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('0 matches');
  });
});

describe('create', () => {
  for (const template of ['report', 'memo', 'blank'] as const) {
    it(`create --template ${template} validates`, async () => {
      const dir = tmpDir(`dx-create-${template}`);
      const out = path.join(dir, `${template}.docx`);
      const r = (await runCli(['create', out, '--template', template]));
      expect(r.status).toBe(0);
      expect(r.stdout).toContain(`template=${template}`);
      expect((await runCli(['validate', out])).status).toBe(0);
    });
  }
  it('create rejects unknown templates with usage exit 2', async () => {
    const dir = tmpDir('dx-create-bad');
    expect((await runCli(['create', path.join(dir, 'x.docx'), '--template', 'bogus'])).status).toBe(2);
  });
  it('create trailing --template exits 2 with no file written', async () => {
    const dir = tmpDir('dx-create-dangle');
    const out = path.join(dir, 'dangling.docx');
    const r = (await runCli(['create', out, '--template']));
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/--template/i);
    expect(fs.existsSync(out)).toBe(false);
  });
  it('create --template with flag-like value exits 2 with no file written', async () => {
    const dir = tmpDir('dx-create-dangle-flag');
    const out = path.join(dir, 'dangling2.docx');
    const r = (await runCli(['create', out, '--template', '--force']));
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/--template/i);
    expect(fs.existsSync(out)).toBe(false);
  });
  it('create refuses to overwrite without --force, clobbers with --force', async () => {
    const dir = tmpDir('dx-create-force');
    const out = path.join(dir, 'doc.docx');
    expect((await runCli(['create', out, '--template', 'blank'])).status).toBe(0);
    const before = fs.readFileSync(out);
    const refuse = (await runCli(['create', out, '--template', 'blank']));
    expect(refuse.status).toBe(2);
    expect(refuse.stderr).toContain(path.basename(out));
    expect(refuse.stderr).toMatch(/--force/);
    expect(fs.readFileSync(out).equals(before)).toBe(true);
    const clobber = (await runCli(['create', out, '--template', 'blank', '--force']));
    expect(clobber.status).toBe(0);
    expect(clobber.stdout).toContain('CREATED');
    expect((await runCli(['validate', out])).status).toBe(0);
  });
  it('create --help discloses overwrite behavior', async () => {
    tmpDir('dx-create-help');
    const r = (await runCli(['create', '--help']));
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/--force/);
    expect(r.stdout).toMatch(/overwrite/i);
  });
});

describe('create --from-md', () => {
  const RICH_MD = [
    '# Composition Proof',
    '',
    'Opening paragraph for the composed document.',
    '',
    '## Sections',
    '',
    '- first bullet',
    '- second bullet',
    '',
    '1. first step',
    '2. second step',
    '',
    '| Name | Value |',
    '| --- | --- |',
    '| alpha | 1 |',
    '| beta | 2 |',
    '',
    '---',
    '',
    '```ts',
    'const proof = true;',
    '```',
    '',
    'A note reference[^fn] in running text.',
    '',
    '[^fn]: The footnote body.',
    '',
  ].join('\n');

  function writeMd(dir: string, name = 'doc.md'): string {
    const file = path.join(dir, name);
    fs.writeFileSync(file, RICH_MD);
    return file;
  }

  it('composes rich Markdown into a fresh document that validates', async () => {
    const dir = tmpDir('dx-from-md-rich');
    const md = writeMd(dir);
    const out = path.join(dir, 'composed.docx');
    const r = (await runCli(['create', out, '--template', 'report', '--from-md', md]));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('CREATED');
    expect(r.stdout).toContain('template=report');
    expect(r.stdout).toContain('applied=');
    expect((await runCli(['validate', out])).status).toBe(0);
    for (const query of ['Composition Proof', 'first bullet', 'second step', 'alpha']) {
      const found = (await runCli(['search', out, query]));
      expect(found.status).toBe(0);
      expect(found.stdout).toMatch(/[1-9][0-9]* match/);
    }
  });

  it('composes with --from-md=<path> equals form', async () => {
    const dir = tmpDir('dx-from-md-equals');
    const md = writeMd(dir);
    const out = path.join(dir, 'composed.docx');
    const r = (await runCli(['create', out, `--from-md=${md}`]));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('CREATED');
    expect((await runCli(['validate', out])).status).toBe(0);
  });

  it('rejects out-of-scope Markdown with exit 1 and writes nothing', async () => {
    const dir = tmpDir('dx-from-md-bad');
    const md = path.join(dir, 'bad.md');
    fs.writeFileSync(md, '# Ok\n\n![](image.png)\n');
    const out = path.join(dir, 'bad.docx');
    const r = (await runCli(['create', out, '--from-md', md]));
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_VALIDATION/);
    expect(r.stderr).toMatch(/line 3/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('rejects a missing Markdown file with exit 1 and writes nothing', async () => {
    const dir = tmpDir('dx-from-md-missing');
    const out = path.join(dir, 'missing.docx');
    const r = (await runCli(['create', out, '--from-md', path.join(dir, 'nope.md')]));
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_IO/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('create trailing --from-md exits 2 with no file written', async () => {
    const dir = tmpDir('dx-from-md-dangle');
    const out = path.join(dir, 'dangling.docx');
    const r = (await runCli(['create', out, '--from-md']));
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/--from-md/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('create --from-md refuses to overwrite without --force, clobbers with --force', async () => {
    const dir = tmpDir('dx-from-md-force');
    const md = writeMd(dir);
    const out = path.join(dir, 'doc.docx');
    expect((await runCli(['create', out, '--from-md', md])).status).toBe(0);
    const before = fs.readFileSync(out);
    const refuse = (await runCli(['create', out, '--from-md', md]));
    expect(refuse.status).toBe(2);
    expect(refuse.stderr).toContain(path.basename(out));
    expect(refuse.stderr).toMatch(/--force/);
    expect(fs.readFileSync(out).equals(before)).toBe(true);
    const clobber = (await runCli(['create', out, '--from-md', md, '--force']));
    expect(clobber.status).toBe(0);
    expect(clobber.stdout).toContain('CREATED');
    expect((await runCli(['validate', out])).status).toBe(0);
  });

  it('create --help documents --from-md scope', async () => {
    tmpDir('dx-from-md-help');
    const r = (await runCli(['create', '--help']));
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/--from-md/);
    expect(r.stdout).toMatch(/line number/);
  });

  it('splits inline spans into runs with real hyperlink targets', async () => {
    const dir = tmpDir('dx-from-md-spans');
    const md = path.join(dir, 'spans.md');
    fs.writeFileSync(md, 'Hello **bold** and *italic* and `code` and [docs](https://example.com) end\n');
    const out = path.join(dir, 'spans.docx');
    const r = (await runCli(['create', out, '--from-md', md]));
    expect(r.status).toBe(0);
    expect((await runCli(['validate', out])).status).toBe(0);
    const xml = docXml(out);
    expect(xml).toContain('w:b');
    expect(xml).toContain('w:i');
    expect(xml).toContain('w:hyperlink');
    expect(xml).toContain('>bold<');
    expect(xml).toContain('>docs<');
    expect(xml).not.toContain('**bold**');
    expect(xml).not.toContain('[docs](https://example.com)');
  });

  it('embeds block images with --allowed-dir and keeps the template seed', async () => {
    const dir = tmpDir('dx-from-md-image');
    const pngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    fs.writeFileSync(path.join(dir, 'tiny.png'), Buffer.from(pngBase64, 'base64'));
    const md = path.join(dir, 'pic.md');
    fs.writeFileSync(md, 'Before\n\n![Tiny](tiny.png)\n\nAfter\n');
    const out = path.join(dir, 'pic.docx');
    const r = (await runCli(['create', out, '--template', 'report', '--from-md', md, '--allowed-dir', dir]));
    expect(r.status).toBe(0);
    expect((await runCli(['validate', out])).status).toBe(0);
    const xml = docXml(out);
    expect(xml).toContain('w:drawing');
    expect(xml).toContain('Q3 Executive');
    expect(xml.indexOf('Q3 Executive')).toBeLessThan(xml.indexOf('<w:t>After</w:t>'));
  });

  it('rejects image refs outside --allowed-dir with exit 1 and writes nothing', async () => {
    const dir = tmpDir('dx-from-md-jail');
    const md = path.join(dir, 'evil.md');
    fs.writeFileSync(md, '![Evil](../../etc/passwd)\n');
    const out = path.join(dir, 'evil.docx');
    const r = (await runCli(['create', out, '--from-md', md, '--allowed-dir', dir]));
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_VALIDATION/);
    expect(r.stderr).toMatch(/outside the allowed directory/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('create trailing --allowed-dir exits 2 with no file written', async () => {
    const dir = tmpDir('dx-from-md-dir-dangle');
    const md = writeMd(dir);
    const out = path.join(dir, 'dangling.docx');
    const r = (await runCli(['create', out, '--from-md', md, '--allowed-dir']));
    expect(r.status).toBe(2);
    expect(r.stderr).toMatch(/--allowed-dir/);
    expect(fs.existsSync(out)).toBe(false);
  });
});

describe('heading, stats and clear round-trips', () => {
  it('insert type:heading lands a Heading2 style and validates', async () => {
    const dir = tmpDir('dx-heading');
    const file = copyGallery(dir);
    const anchor = await firstNodeId(file);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(
      ops,
      JSON.stringify({
        operations: [{ action: 'insert', type: 'heading', anchor, position: 'after', level: 2, text: 'Wave Heading Probe' }],
      })
    );
    expect((await runCli(['batch', file, ops])).status).toBe(0);
    expect((await runCli(['validate', file])).status).toBe(0);
    expect(docXml(file)).toContain('w:val="Heading2"');
    const found = (await runCli(['search', file, 'Wave Heading Probe']));
    expect(found.status).toBe(0);
    expect(found.stdout).toMatch(/[0-9A-F]{4,} .*Wave Heading Probe/);
  });
  it('insert type:heading rejects level 9 with E_VALIDATION', async () => {
    const dir = tmpDir('dx-heading-bad');
    const file = copyGallery(dir);
    const anchor = await firstNodeId(file);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(
      ops,
      JSON.stringify({ operations: [{ action: 'insert', type: 'heading', anchor, position: 'after', level: 9, text: 'x' }] })
    );
    const before = fs.readFileSync(file);
    const r = (await runCli(['batch', file, ops]));
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/\[E_VALIDATION\]/);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
  it('stats prints words + estimated pages, --json matches', async () => {
    const dir = tmpDir('dx-stats');
    const file = copyGallery(dir);
    const r = (await runCli(['stats', file]));
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/words=\d+ paragraphs=\d+ pages~\d+/);
    const j = (await runCli(['--json', 'stats', file]));
    expect(j.status).toBe(0);
    const payload = JSON.parse(j.stdout);
    expect(payload.words).toBeGreaterThan(100);
    expect(payload.pages).toBeGreaterThanOrEqual(1);
  });
  it('inspect footer carries words + estimated pages', async () => {
    const dir = tmpDir('dx-inspect-stats');
    const file = copyGallery(dir);
    const r = (await runCli(['inspect', file]));
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/words=\d+ paragraphs=\d+ pages~\d+/);
  });
  it('clearDoc resets the body, keeps styles/props, validates', async () => {
    const dir = tmpDir('dx-clear');
    const file = copyGallery(dir);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'clearDoc' }] }));
    const r = (await runCli(['batch', file, ops]));
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('cleared=1');
    expect((await runCli(['validate', file])).status).toBe(0);
    const zip = new AdmZip(fs.readFileSync(file));
    expect(zip.getEntry('word/styles.xml')).not.toBeNull();
    const xml = docXml(file);
    expect((xml.match(/<w:p[\s>]/g) ?? []).length).toBe(1);
    expect(xml).toContain('w:sectPr');
    const stats = (await runCli(['stats', file]));
    expect(stats.stdout).toContain('words=0');
  });
});

describe('stats wordcount recipe', () => {
  it('stats --help documents body+tables recipe and rough 500-wpp estimate', async () => {
    tmpDir('dx-stats-help');
    const r = (await runCli(['stats', '--help']));
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/table cell/i);
    expect(r.stdout).toMatch(/Headers.*footers.*footnotes/i);
    expect(r.stdout).toMatch(/500/);
    expect(r.stdout).toMatch(/rough estimate/i);
  });
  it('stats counts table cells with spacing, excludes footers, labels pages estimate', async () => {
    const dir = tmpDir('dx-stats-recipe');
    const file = copyGallery(dir);
    const r = (await runCli(['stats', file]));
    expect(r.status).toBe(0);
    // pages labeled as rough estimate with method
    expect(r.stdout).toMatch(/pages~\d+ \(est, 500 wpp\)/);
    expect(r.stdout).toMatch(/recipe: body paragraphs \+ table cells/);
    expect(r.stdout).toMatch(/rough estimate = max\(1, ceil\(words\/500\)\)/);
    const m = /words=(\d+) paragraphs=(\d+) pages~(\d+)/.exec(r.stdout);
    expect(m, 'stats line parses').toBeTruthy();
    const words = Number(m![1]);
    const paragraphs = Number(m![2]);
    const pages = Number(m![3]);
    expect(pages).toBe(Math.max(1, Math.ceil(words / 500)));

    // Independent recipe straight from word/document.xml: root w:p text plus
    // per-cell spaced table text; headers/footers/footnotes excluded by construction.
    const xml = docXml(file);
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const body = doc.getElementsByTagName('w:body')[0] as unknown as Element;
    expect(body, 'w:body present').toBeTruthy();
    let expectWords = 0;
    let expectParagraphs = 0;
    const rootKids = body.childNodes;
    let tableSmashed = 0;
    let tableSpaced = 0;
    for (let i = 0; i < rootKids.length; i++) {
      const el = rootKids[i] as unknown as Element;
      if (!el || el.nodeType !== 1) continue;
      if (el.tagName === 'w:p') {
        expectParagraphs += 1;
        const ts = el.getElementsByTagName('w:t');
        let t = '';
        for (let k = 0; k < ts.length; k++) t += (ts[k] as unknown as Element).textContent || '';
        const hits = t.match(/\S+/g);
        if (hits) expectWords += hits.length;
      } else if (el.tagName === 'w:tbl') {
        const tcs = el.getElementsByTagName('w:tc');
        const cells: string[] = [];
        let smashed = '';
        for (let c = 0; c < tcs.length; c++) {
          const tc = tcs[c] as unknown as Element;
          const ts = tc.getElementsByTagName('w:t');
          let cell = '';
          for (let k = 0; k < ts.length; k++) cell += (ts[k] as unknown as Element).textContent || '';
          smashed += cell;
          if (cell.trim().length > 0) cells.push(cell.trim());
        }
        const smashedHits = smashed.match(/\S+/g);
        if (smashedHits) tableSmashed += smashedHits.length;
        const joined = cells.join(' ');
        const spacedHits = joined.match(/\S+/g);
        if (spacedHits) tableSpaced += spacedHits.length;
        expectWords += spacedHits ? spacedHits.length : 0;
      }
    }
    // Fixture has exactly one table; spacing at cell edges must recover words.
    expect(tableSpaced).toBeGreaterThan(0);
    expect(tableSpaced).toBeGreaterThan(tableSmashed);
    expect(words).toBe(expectWords);
    expect(paragraphs).toBe(expectParagraphs);

    // Verify against the markdown export: table headers counted by stats are present in md.
    const mdOut = path.join(dir, 'out.md');
    const ex = (await runCli(['export', file, '--format', 'md', '--output', mdOut]));
    expect(ex.status).toBe(0);
    const md = fs.readFileSync(mdOut, 'utf8');
    expect(md).toContain('Deontological Perspective');
    expect(md).toContain('Consequentialist Perspective');
  });
  it('inspect footer labels pages estimate', async () => {
    const dir = tmpDir('dx-inspect-est');
    const file = copyGallery(dir);
    const r = (await runCli(['inspect', file]));
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/pages~\d+ \(est, 500 wpp\)/);
  });
});
