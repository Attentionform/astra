import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

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

function writeBlankDocx(prefix = 'astra-cli-safety'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
  return file;
}

function documentXml(file: string): string {
  const zip = new AdmZip(fs.readFileSync(file));
  return zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

/** Text-aware check: replace ops land as tracked w:ins/w:del split runs, so
 *  raw XML rarely holds the contiguous phrase — search the AST instead. */
async function docContains(file: string, text: string): Promise<boolean> {
  const parser = new DocxParser(fs.readFileSync(file));
  const ast = await parser.getAST();
  if (parser.searchNodesInAST(ast, text).length > 0) return true;
  return documentXml(file).includes(text);
}

async function firstParagraphId(file: string): Promise<string> {
  const parser = new DocxParser(fs.readFileSync(file));
  const ast = await parser.getAST();
  const nodes = ast.sections.flatMap((s) => s.nodes);
  const para = nodes.find((n) => n.type === 'paragraph') ?? nodes[0];
  if (!para) throw new Error('blank template has no nodes');
  return para.id;
}

function writeOps(payload: unknown): string {
  const file = tmpPath('astra-cli-safety-ops', '.json');
  fs.writeFileSync(file, JSON.stringify(payload));
  return file;
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

describe('batch --dry-run', () => {
  it('valid replace exits 0, lists the node id + diff, writes nothing', async () => {
    const file = writeBlankDocx('astra-dry-ok');
    const before = fs.readFileSync(file);
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Dry run proof' }], sequential: true });
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--dry-run'])).resolves.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    const out = stdoutText();
    expect(out).toContain('DRY-RUN OK');
    expect(out).toContain(target);
    expect(out).toContain('--- a/');
    expect(out).toContain('+++ b/');
    expect(out).toContain('+ Dry run proof');
    expect(out).toContain('writes nothing');
  });

  it('dry run under --json carries affected + diff, stdout stays pure JSON lines', async () => {
    const file = writeBlankDocx('astra-dry-json');
    const before = fs.readFileSync(file);
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Json dry proof' }], sequential: true });
    vi.mocked(console.log).mockClear();
    await expect(run(['--json', 'batch', file, ops, '--dry-run'])).resolves.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    const payload = JSON.parse(stdoutText());
    expect(payload.ok).toBe(true);
    expect(payload.dryRun).toBe(true);
    expect(Array.isArray(payload.affected)).toBe(true);
    expect(payload.affected.length).toBeGreaterThan(0);
    expect(payload.affected[0].nodeId).toBe(target);
    expect(payload.affected[0].after).toContain('Json dry proof');
    expect(typeof payload.diff).toBe('string');
    expect(payload.diff).toContain('+ Json dry proof');
  });

  it('invalid op dry run exits 1 and writes nothing', async () => {
    const file = writeBlankDocx('astra-dry-bad');
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'definitely_not_a_real_op' }] });
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--dry-run'])).resolves.toBe(1);
    expect(stderrText()).toContain('E_VALIDATION');
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('dry run with --out creates no output file', async () => {
    const file = writeBlankDocx('astra-dry-out');
    const before = fs.readFileSync(file);
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Nowhere proof' }], sequential: true });
    const dest = tmpPath('astra-dry-out-dest');
    await expect(run(['batch', file, ops, '--out', dest, '--dry-run'])).resolves.toBe(0);
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('dry run reports style-only targets as affected with no text change', async () => {
    const file = writeBlankDocx('astra-dry-style');
    const before = fs.readFileSync(file);
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'style', target, bold: true }], sequential: true });
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--dry-run'])).resolves.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    expect(stdoutText()).toContain(target);
    expect(stdoutText()).toContain('(no text change)');
  });

  it('engine dryRun core: success preview + full rewind (text, history, bytes)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const target = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'paragraph')?.id;
    if (!target) throw new Error('blank template has no paragraph');
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [{ action: 'replace', target, text: 'Engine dry proof' } as never],
      sequential: true,
      dryRun: true,
    });
    expect(res.success).toBe(true);
    expect(res.dryRun).toBe(true);
    expect(res.appliedCount).toBe(1);
    expect(res.preview?.length).toBeGreaterThan(0);
    const entry = res.preview?.find((p) => p.nodeId === target);
    expect(entry?.after).toContain('Engine dry proof');
    expect(res.dryRunBytes).toBeInstanceOf(Buffer);
    // Rewound: live AST text unchanged, history empty, rezip has no new text.
    const fresh = await parser.getAST();
    const hits = parser.searchNodesInAST(fresh, 'Engine dry proof');
    expect(hits.length).toBe(0);
    expect(engine.getHistory().length).toBe(0);
    expect(parser.rezip().toString('utf8')).not.toContain('Engine dry proof');
  });
});

describe('batch --backup', () => {
  it('explicit --backup snapshots the input before an in-place write', async () => {
    const file = writeBlankDocx('astra-bak-inplace');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Backup proof' }], sequential: true });
    const backup = track(path.join(os.tmpdir(), `astra-bak-explicit-${process.pid}-${Date.now()}.docx`));
    await expect(run(['batch', file, ops, '--backup', backup])).resolves.toBe(0);
    expect(fs.existsSync(backup)).toBe(true);
    expect(fs.readFileSync(backup).equals(before)).toBe(true);
    expect(await docContains(file, 'Backup proof')).toBe(true);
  });

  it('bare --backup uses the default <input>.bak.docx', async () => {
    const file = writeBlankDocx('astra-bak-default');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Default bak proof' }], sequential: true });
    const backup = `${file}.bak.docx`;
    track(backup);
    await expect(run(['batch', file, ops, '--backup'])).resolves.toBe(0);
    expect(fs.readFileSync(backup).equals(before)).toBe(true);
    expect(await docContains(file, 'Default bak proof')).toBe(true);
  });

  it('--backup with --out snapshots the input, writes the output', async () => {
    const file = writeBlankDocx('astra-bak-out');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Out bak proof' }], sequential: true });
    const dest = tmpPath('astra-bak-out-dest');
    const backup = track(path.join(os.tmpdir(), `astra-bak-out-${process.pid}-${Date.now()}.docx`));
    await expect(run(['batch', file, ops, '--out', dest, '--backup', backup])).resolves.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    expect(fs.readFileSync(backup).equals(before)).toBe(true);
    expect(await docContains(dest, 'Out bak proof')).toBe(true);
  });

  it('--backup equal to the target fails closed with nothing written', async () => {
    const file = writeBlankDocx('astra-bak-self');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'must not land' }], sequential: true });
    const dest = tmpPath('astra-bak-self-dest');
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--out', dest, `--backup=${dest}`])).resolves.toBe(1);
    expect(stderrText()).toContain('E_VALIDATION');
    expect(fs.existsSync(dest)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('--backup is skipped under --dry-run (note on stderr, nothing written)', async () => {
    const file = writeBlankDocx('astra-bak-dry');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'no bak proof' }], sequential: true });
    const backup = path.join(os.tmpdir(), `astra-bak-dry-${process.pid}-${Date.now()}.docx`);
    track(backup);
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--dry-run', '--backup', backup])).resolves.toBe(0);
    expect(stderrText()).toContain('--backup skipped');
    expect(fs.existsSync(backup)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
});

describe('batch --matrix', () => {
  function writeCsv(content: string): string {
    const file = tmpPath('astra-matrix-vars', '.csv');
    fs.writeFileSync(file, content);
    return file;
  }

  it('one row = one --out with literal substitution, input untouched', async () => {
    const file = writeBlankDocx('astra-matrix-ok');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({
      operations: [{ action: 'replace', target, text: 'Hello {{name}} of {{company}}' }],
      sequential: true,
    });
    const csv = writeCsv('name,company\nAda,Acme\nBob,Initech\n');
    const outTemplate = track(path.join(os.tmpdir(), `astra-matrix-out-{{name}}-${process.pid}.docx`));
    await expect(run(['batch', file, ops, '--matrix', csv, '--out', outTemplate])).resolves.toBe(0);
    const adaOut = outTemplate.replace('{{name}}', 'Ada');
    const bobOut = outTemplate.replace('{{name}}', 'Bob');
    track(adaOut);
    track(bobOut);
    expect(await docContains(adaOut, 'Hello Ada of Acme')).toBe(true);
    expect(await docContains(bobOut, 'Hello Bob of Initech')).toBe(true);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('unknown {{var}} fails closed with nothing written', async () => {
    const file = writeBlankDocx('astra-matrix-unknown');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Hi {{nope}}' }], sequential: true });
    const csv = writeCsv('name\nAda\n');
    const outTemplate = path.join(os.tmpdir(), `astra-matrix-unk-{{name}}-${process.pid}-${Date.now()}.docx`);
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--matrix', csv, '--out', outTemplate])).resolves.toBe(1);
    expect(stderrText()).toContain('{{nope}}');
    expect(fs.existsSync(outTemplate.replace('{{name}}', 'Ada'))).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('missing --out with --matrix is a usage error (exit 2)', async () => {
    const file = writeBlankDocx('astra-matrix-noout');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'x' }], sequential: true });
    const csv = writeCsv('name\nAda\n');
    await expect(run(['batch', file, ops, '--matrix', csv])).resolves.toBe(2);
  });

  it('duplicate row outputs fail closed before any write', async () => {
    const file = writeBlankDocx('astra-matrix-dupe');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Hi {{name}}' }], sequential: true });
    const csv = writeCsv('name,tag\nAda,x\nBob,x\n');
    const outTemplate = path.join(os.tmpdir(), `astra-matrix-dupe-{{tag}}-${process.pid}-${Date.now()}.docx`);
    track(outTemplate.replace('{{tag}}', 'x'));
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops, '--matrix', csv, '--out', outTemplate])).resolves.toBe(1);
    expect(stderrText()).toContain('same --out');
    expect(fs.existsSync(outTemplate.replace('{{tag}}', 'x'))).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('matrix + --dry-run writes nothing and exits 0', async () => {
    const file = writeBlankDocx('astra-matrix-dry');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Hi {{name}}' }], sequential: true });
    const csv = writeCsv('name\nAda\nBob\n');
    const outTemplate = path.join(os.tmpdir(), `astra-matrix-dry-{{name}}-${process.pid}-${Date.now()}.docx`);
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--matrix', csv, '--out', outTemplate, '--dry-run'])).resolves.toBe(0);
    expect(stdoutText()).toContain('MATRIX OK');
    expect(stdoutText()).toContain('DRY-RUN OK');
    expect(fs.existsSync(outTemplate.replace('{{name}}', 'Ada'))).toBe(false);
    expect(fs.existsSync(outTemplate.replace('{{name}}', 'Bob'))).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('matrix values are literal: no re-expansion, {{ name }} spacing tolerated', async () => {
    const file = writeBlankDocx('astra-matrix-literal');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'v={{a}}' }], sequential: true });
    const csv = writeCsv('a\n{{b}}\n');
    const outTemplate = track(path.join(os.tmpdir(), `astra-matrix-lit-{{ a }}-${process.pid}.docx`));
    await expect(run(['batch', file, ops, '--matrix', csv, '--out', outTemplate])).resolves.toBe(0);
    const literalOut = outTemplate.replace('{{ a }}', '{{b}}');
    track(literalOut);
    expect(await docContains(literalOut, 'v={{b}}')).toBe(true);
  });

  it('batch --help documents --dry-run, --backup, and --matrix', async () => {
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', '--help'])).resolves.toBe(0);
    const text = stdoutText();
    expect(text).toContain('--dry-run');
    expect(text).toContain('--backup');
    expect(text).toContain('--matrix');
    expect(text).toContain('{{column}}');
  });

  it('batch --help documents row delete with a copy-paste shape', async () => {
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', '--help'])).resolves.toBe(0);
    const text = stdoutText();
    expect(text).toContain('deleteRow');
    expect(text).toContain('batch --schema');
  });
});

describe('batch chart stale body refs', () => {
  const oldTitle = 'H1 Revenue January\u2013June';
  const oldCats = ['January', 'February', 'March', 'April', 'May', 'June'];
  const oldVals = [1200, 2400, 3100, 4200, 5300, 6000];
  const caption = 'Figure 1: H1 Revenue January\u2013June sales trend (see chart).';
  const cta = 'Act now \u2014 6,000 early-bird bonus ends Friday.';

  async function setupStaleDoc(prefix: string): Promise<{ file: string; chartId: string }> {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    const engine = new OperationEngine(ast, parser);
    const setup = await engine.executeBatch({
      operations: [
        {
          action: 'chart',
          chartType: 'column',
          title: oldTitle,
          categories: oldCats,
          series: [{ name: 'Revenue', values: oldVals }],
          anchor,
          id: 'chart-stale',
        } as never,
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: caption, id: 'cap-stale' } as never,
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: cta, id: 'cta-stale' } as never,
      ],
      sequential: true,
    });
    if (!setup.success) throw new Error(`stale fixture setup failed: ${JSON.stringify(setup.errors).slice(0, 300)}`);
    const file = tmpPath(prefix);
    fs.writeFileSync(file, parser.rezip());
    const probe = new DocxParser(fs.readFileSync(file));
    const probeAst = await probe.getAST();
    const chart = probeAst.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart' || n.type === 'chart3d');
    if (!chart) throw new Error('stale fixture has no chart node');
    return { file, chartId: chart.id };
  }

  function updateOps(chartId: string): string {
    return writeOps({
      operations: [
        {
          action: 'chart',
          operation: 'update',
          target: chartId,
          title: 'July Revenue Update',
          categories: ['Jul'],
          series: [{ name: 'Revenue', values: [7500] }],
        },
      ],
      sequential: true,
    });
  }

  it('batch success prints stale refs on human output', async () => {
    const { file, chartId } = await setupStaleDoc('astra-stale-human');
    const ops = updateOps(chartId);
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops])).resolves.toBe(0);
    const out = stdoutText();
    expect(out).toContain('BATCH OK');
    expect(out).toContain('possibly-stale');
    expect(out).toContain('January');
    expect(out).toContain('6,000');
  });

  it('batch success includes stale refs in --json when non-empty', async () => {
    const { file, chartId } = await setupStaleDoc('astra-stale-json');
    const ops = updateOps(chartId);
    vi.mocked(console.log).mockClear();
    await expect(run(['--json', 'batch', file, ops])).resolves.toBe(0);
    const payload = JSON.parse(stdoutText());
    expect(payload.ok).toBe(true);
    expect(Array.isArray(payload.possiblyStaleBodyRefs)).toBe(true);
    expect(payload.possiblyStaleBodyRefs.length).toBeGreaterThanOrEqual(2);
    const previews = payload.possiblyStaleBodyRefs.map((r: { preview: string }) => r.preview).join('\n');
    expect(previews).toContain('January');
    expect(previews).toContain('6,000');
    for (const r of payload.possiblyStaleBodyRefs) {
      expect(typeof r.nodeId).toBe('string');
      expect(typeof r.preview).toBe('string');
    }
  });

  it('dry-run prints stale refs on human output and --json', async () => {
    const { file, chartId } = await setupStaleDoc('astra-stale-dry-human');
    const ops = updateOps(chartId);
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--dry-run'])).resolves.toBe(0);
    const human = stdoutText();
    expect(human).toContain('DRY-RUN OK');
    expect(human).toContain('possibly-stale');
    expect(human).toContain('January');

    const { file: file2, chartId: chartId2 } = await setupStaleDoc('astra-stale-dry-json');
    const ops2 = updateOps(chartId2);
    vi.mocked(console.log).mockClear();
    await expect(run(['--json', 'batch', file2, ops2, '--dry-run'])).resolves.toBe(0);
    const payload = JSON.parse(stdoutText());
    expect(payload.ok).toBe(true);
    expect(payload.dryRun).toBe(true);
    expect(Array.isArray(payload.possiblyStaleBodyRefs)).toBe(true);
    expect(payload.possiblyStaleBodyRefs.length).toBeGreaterThanOrEqual(2);
  });

  it('dry-run stays quiet when there are no stale refs', async () => {
    const file = writeBlankDocx('astra-stale-dry-quiet');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Quiet proof' }], sequential: true });
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', file, ops, '--dry-run'])).resolves.toBe(0);
    expect(stdoutText()).not.toContain('possibly-stale');

    const file2 = writeBlankDocx('astra-stale-dry-quiet-json');
    const target2 = await firstParagraphId(file2);
    const ops2 = writeOps({ operations: [{ action: 'replace', target: target2, text: 'Quiet proof' }], sequential: true });
    vi.mocked(console.log).mockClear();
    await expect(run(['--json', 'batch', file2, ops2, '--dry-run'])).resolves.toBe(0);
    const payload = JSON.parse(stdoutText());
    expect(payload.ok).toBe(true);
    expect(payload.dryRun).toBe(true);
    expect('possiblyStaleBodyRefs' in payload).toBe(false);
  });

  it('dry-run json stale refs carry nodeId and preview content', async () => {
    const { file, chartId } = await setupStaleDoc('astra-stale-dry-content');
    const ops = updateOps(chartId);
    vi.mocked(console.log).mockClear();
    await expect(run(['--json', 'batch', file, ops, '--dry-run'])).resolves.toBe(0);
    const payload = JSON.parse(stdoutText());
    expect(payload.ok).toBe(true);
    expect(payload.dryRun).toBe(true);
    expect(Array.isArray(payload.possiblyStaleBodyRefs)).toBe(true);
    expect(payload.possiblyStaleBodyRefs.length).toBeGreaterThanOrEqual(2);
    const previews = payload.possiblyStaleBodyRefs.map((r: { preview: string }) => r.preview).join('\n');
    expect(previews).toContain('January');
    expect(previews).toContain('6,000');
    for (const r of payload.possiblyStaleBodyRefs) {
      expect(typeof r.nodeId).toBe('string');
      expect(typeof r.preview).toBe('string');
    }
  });
});
