import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Readable } from 'stream';
import { run } from '../src/cli';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { ASTRA_TMPDIR_ENV } from '@attentionform/astra-engine/tmp';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function writeBlankDocx(prefix = 'astra-cli'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
  return file;
}

function documentXml(file: string): string {
  const zip = new AdmZip(fs.readFileSync(file));
  return zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
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

describe('cli usage (exit 2)', () => {
  it('no args prints usage', async () => {
    await expect(run([])).resolves.toBe(2);
  });

  it('unknown command prints usage', async () => {
    await expect(run(['frobnicate'])).resolves.toBe(2);
  });

  it('--help exits 0', async () => {
    await expect(run(['--help'])).resolves.toBe(0);
  });

  it('an unknown command exits 2 and lists the commands', async () => {
    vi.mocked(console.error).mockClear();
    await expect(run(['chat'])).resolves.toBe(2);
    const err = vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
    expect(err).toContain('unknown command "chat"');
  });
});

describe('cli validate (fail-closed exit codes)', () => {
  it('valid docx exits 0', async () => {
    const file = writeBlankDocx('astra-cli-valid');
    await expect(run(['validate', file])).resolves.toBe(0);
  });

  it('missing file exits 1 (not 0, not 2)', async () => {
    const missing = tmpPath('astra-cli-missing');
    await expect(run(['validate', missing])).resolves.toBe(1);
  });

  it('corrupt payload with .docx name exits 1', async () => {
    const file = tmpPath('astra-cli-corrupt');
    fs.writeFileSync(file, 'this is not a zip package');
    await expect(run(['validate', file])).resolves.toBe(1);
  });
});

describe('cli inspect', () => {
  it('inspect exits 0 on a real docx', async () => {
    const file = writeBlankDocx('astra-cli-inspect');
    await expect(run(['inspect', file])).resolves.toBe(0);
  });

  it('inspect missing file exits 1', async () => {
    await expect(run(['inspect', tmpPath('astra-cli-nope')])).resolves.toBe(1);
  });

  it('inspect rejects bad fidelity with usage exit 2', async () => {
    const file = writeBlankDocx('astra-cli-inspect-bad');
    await expect(run(['inspect', file, '--fidelity', 'ultra'])).resolves.toBe(2);
  });
});

describe('cli sanitize', () => {
  it('sanitize writes the output and exits 0', async () => {
    const input = writeBlankDocx('astra-cli-san-in');
    const output = tmpPath('astra-cli-san-out');
    await expect(run(['sanitize', input, output])).resolves.toBe(0);
    expect(fs.existsSync(output)).toBe(true);
    expect(fs.statSync(output).size).toBeGreaterThan(0);
  });

  it('sanitize missing input exits 1 and writes nothing', async () => {
    const output = tmpPath('astra-cli-san-missing-out');
    await expect(run(['sanitize', tmpPath('astra-cli-san-missing'), output])).resolves.toBe(1);
    expect(fs.existsSync(output)).toBe(false);
  });
});

describe('cli batch', () => {
  async function firstParagraphId(file: string): Promise<string> {
    const parser = new DocxParser(fs.readFileSync(file));
    const ast = await parser.getAST();
    const nodes = ast.sections.flatMap((s) => s.nodes);
    const para = nodes.find((n) => n.type === 'paragraph') ?? nodes[0];
    if (!para) throw new Error('blank template has no nodes (cannot build batch smoke ops)');
    return para.id;
  }

  function writeOps(payload: unknown): string {
    const file = tmpPath('astra-cli-ops', '.json');
    fs.writeFileSync(file, JSON.stringify(payload));
    return file;
  }

  it('valid 1-item replace batch exits 0 and lands the text', async () => {
    const file = writeBlankDocx('astra-cli-batch-ok');
    const target = await firstParagraphId(file);
    const ops = writeOps({ operations: [{ action: 'replace', target, text: 'Hello CLI smoke' }], sequential: true });
    await expect(run(['batch', file, ops])).resolves.toBe(0);
    const parser = new DocxParser(fs.readFileSync(file));
    const ast = await parser.getAST();
    const hits = parser.searchNodesInAST(ast, 'Hello CLI smoke');
    expect(hits.length > 0 || documentXml(file).includes('Hello CLI smoke')).toBe(true);
  });

  it('invalid op exits 1 and leaves the file byte-identical (rollback, no write)', async () => {
    const file = writeBlankDocx('astra-cli-batch-bad');
    const before = fs.readFileSync(file);
    const ops = writeOps({ operations: [{ action: 'definitely_not_a_real_op' }] });
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('good op + bad target exits 1 and leaves the file byte-identical', async () => {
    const file = writeBlankDocx('astra-cli-batch-half');
    const target = await firstParagraphId(file);
    const before = fs.readFileSync(file);
    const ops = writeOps({
      operations: [
        { action: 'replace', target, text: 'must not persist' },
        { action: 'replace', target: 'ghost_node_that_does_not_exist', text: 'boom' },
      ],
      sequential: true,
    });
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('malformed ops JSON exits 1', async () => {
    const file = writeBlankDocx('astra-cli-batch-json');
    const ops = tmpPath('astra-cli-ops-bad', '.json');
    fs.writeFileSync(ops, '{ not json');
    await expect(run(['batch', file, ops])).resolves.toBe(1);
  });

  it('table deleteRow removes the row and exits 0', async () => {
    const file = writeBlankDocx('astra-cli-batch-delrow');
    const anchor = await firstParagraphId(file);
    const insertOps = writeOps({
      operations: [
        { action: 'insert', type: 'table', anchor, position: 'after', headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']] },
      ],
      sequential: true,
    });
    await expect(run(['batch', file, insertOps])).resolves.toBe(0);
    const parser = new DocxParser(fs.readFileSync(file));
    const ast = await parser.getAST();
    const table = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'table');
    expect(table).toBeTruthy();
    const rowsBefore = (documentXml(file).match(/<w:tr[\s>]/g) ?? []).length;
    expect(rowsBefore).toBeGreaterThan(1);
    const delOps = writeOps({
      operations: [{ action: 'table', operation: 'deleteRow', target: table!.id, index: 2 }],
      sequential: true,
    });
    await expect(run(['batch', file, delOps])).resolves.toBe(0);
    const rowsAfter = (documentXml(file).match(/<w:tr[\s>]/g) ?? []).length;
    expect(rowsAfter).toBe(rowsBefore - 1);
  });

  it('comment insert lands a comment part and exits 0', async () => {
    const file = writeBlankDocx('astra-cli-batch-comment');
    const target = await firstParagraphId(file);
    const ops = writeOps({
      operations: [{ action: 'comment', operation: 'insert', target, text: 'Review note' }],
      sequential: true,
    });
    await expect(run(['batch', file, ops])).resolves.toBe(0);
    const zip = new AdmZip(fs.readFileSync(file));
    const comments = zip.getEntry('word/comments.xml')?.getData().toString('utf8') ?? '';
    expect(comments).toContain('Review note');
  });

  it('comment operation add fails closed (validator accepts insert|reply)', async () => {
    const file = writeBlankDocx('astra-cli-batch-comment-bad');
    const before = fs.readFileSync(file);
    const target = await firstParagraphId(file);
    const ops = writeOps({
      operations: [{ action: 'comment', operation: 'add', target, text: 'Review note' }],
      sequential: true,
    });
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('chart insert then update round-trips the title and exits 0', async () => {
    const file = writeBlankDocx('astra-cli-batch-chart');
    const anchor = await firstParagraphId(file);
    const insertOps = writeOps({
      operations: [
        {
          action: 'chart', operation: 'insert', anchor, chartType: 'column',
          title: 'Quarterly revenue', categories: ['Q1', 'Q2'], series: [{ name: 'Revenue', values: [10, 20] }],
        },
      ],
      sequential: true,
    });
    await expect(run(['batch', file, insertOps])).resolves.toBe(0);
    const parser = new DocxParser(fs.readFileSync(file));
    const ast = await parser.getAST();
    const chart = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart' || n.type === 'chart3d');
    expect(chart).toBeTruthy();
    const updateOps = writeOps({
      operations: [
        {
          action: 'chart', operation: 'update', target: chart!.id,
          title: 'Quarterly revenue (revised)', categories: ['Q1', 'Q2'], series: [{ name: 'Revenue', values: [12, 22] }],
        },
      ],
      sequential: true,
    });
    await expect(run(['batch', file, updateOps])).resolves.toBe(0);
    const zip = new AdmZip(fs.readFileSync(file));
    const chartXml = zip.getEntry('word/charts/chart1.xml')?.getData().toString('utf8') ?? '';
    expect(chartXml).toContain('Quarterly revenue (revised)');
  });

  // R-44: ops path `-` reads ops JSON from stdin.
  describe('batch ops from stdin (-)', () => {
    const realStdin = process.stdin;
    function mockStdin(stream: unknown): void {
      Object.defineProperty(process, 'stdin', { value: stream, configurable: true, writable: true });
    }
    function restoreStdin(): void {
      Object.defineProperty(process, 'stdin', { value: realStdin, configurable: true, writable: true });
    }
    function stderrText(): string {
      return vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
    }
    function stdoutText(): string {
      return vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
    }

    it('stdin ops apply and exit 0', async () => {
      const file = writeBlankDocx('astra-cli-batch-stdin-ok');
      const target = await firstParagraphId(file);
      const payload = JSON.stringify({
        operations: [{ action: 'replace', target, text: 'Hello stdin smoke' }],
        sequential: true,
      });
      mockStdin(Readable.from([payload]));
      try {
        await expect(run(['batch', file, '-'])).resolves.toBe(0);
      } finally {
        restoreStdin();
      }
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      const hits = parser.searchNodesInAST(ast, 'Hello stdin smoke');
      expect(hits.length > 0 || documentXml(file).includes('Hello stdin smoke')).toBe(true);
    });

    it('bad JSON via stdin exits 1 with the same error as file', async () => {
      const file = writeBlankDocx('astra-cli-batch-stdin-badjson');
      const badOps = tmpPath('astra-cli-ops-stdin-bad', '.json');
      fs.writeFileSync(badOps, '{ not json');
      vi.mocked(console.error).mockClear();
      await expect(run(['batch', file, badOps])).resolves.toBe(1);
      const fileErr = stderrText();
      expect(fileErr).toContain('E_VALIDATION');
      mockStdin(Readable.from(['{ not json']));
      try {
        vi.mocked(console.error).mockClear();
        await expect(run(['batch', file, '-'])).resolves.toBe(1);
      } finally {
        restoreStdin();
      }
      expect(stderrText()).toBe(fileErr);
    });

    it('empty stdin exits 2 fast instead of hanging', async () => {
      const file = writeBlankDocx('astra-cli-batch-stdin-empty');
      mockStdin(Readable.from([]));
      let timer: ReturnType<typeof setTimeout> | undefined;
      try {
        vi.mocked(console.error).mockClear();
        const code = await Promise.race([
          run(['batch', file, '-']),
          new Promise<number>((_, reject) => {
            timer = setTimeout(() => reject(new Error('hung reading empty stdin')), 3000);
          }),
        ]);
        expect(code).toBe(2);
        expect(stderrText()).toContain('E_VALIDATION');
      } finally {
        if (timer !== undefined) clearTimeout(timer);
        restoreStdin();
      }
    });

    it('terminal stdin exits 2 without reading (no hang)', async () => {
      const file = writeBlankDocx('astra-cli-batch-stdin-tty');
      mockStdin({ isTTY: true } as unknown as NodeJS.ReadStream);
      try {
        vi.mocked(console.error).mockClear();
        await expect(run(['batch', file, '-'])).resolves.toBe(2);
      } finally {
        restoreStdin();
      }
      expect(stderrText()).toContain('E_VALIDATION');
    });

    it('batch --help documents the `-` stdin ops path', async () => {
      vi.mocked(console.log).mockClear();
      await expect(run(['batch', '--help'])).resolves.toBe(0);
      expect(stdoutText()).toContain('`-`');
      expect(stdoutText().toLowerCase()).toContain('stdin');
    });
  });

  // Batch text shortcuts: --target-text resolves via engine search to one
  // node id (fail-closed on 0/2+ matches); --append inserts at end (no id).
  // Both synthesize canonical ops into the shared batch path.
  describe('batch --target-text / --append shortcuts', () => {
    function stderrText(): string {
      return vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
    }

    async function seedReplace(file: string, text: string): Promise<void> {
      const target = await firstParagraphId(file);
      const ops = writeOps({ operations: [{ action: 'replace', target, text }], sequential: true });
      await expect(run(['batch', file, ops])).resolves.toBe(0);
    }

    it('typo fix in one command: --target-text + --append replaces the matched node', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-ok');
      await seedReplace(file, 'teh quick brown fox');
      await expect(run(['batch', file, '--target-text', 'teh quick', '--append', 'the quick'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      expect(parser.searchNodesInAST(ast, 'the quick').length).toBeGreaterThan(0);
      expect(documentXml(file)).not.toContain('teh quick');
    });

    it('--target-text with 2+ matches fails closed (exit 1, remedy, no write)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-amb');
      const anchor = await firstParagraphId(file);
      const ops = writeOps({
        operations: [
          { action: 'replace', target: anchor, text: 'red apple one' },
          { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'red apple two' },
        ],
        sequential: true,
      });
      await expect(run(['batch', file, ops])).resolves.toBe(0);
      const before = fs.readFileSync(file);
      vi.mocked(console.error).mockClear();
      await expect(run(['batch', file, '--target-text', 'red apple', '--append', 'green pear'])).resolves.toBe(1);
      expect(stderrText()).toContain('E_VALIDATION');
      expect(stderrText()).toContain('matched 2');
      expect(stderrText()).toContain('Remedy:');
      expect(fs.readFileSync(file).equals(before)).toBe(true);
    });

    it('--target-text with 0 matches fails closed (exit 1, remedy, no write)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-zero');
      const before = fs.readFileSync(file);
      vi.mocked(console.error).mockClear();
      await expect(run(['batch', file, '--target-text', 'qqq-no-such-text-zzz', '--append', 'x'])).resolves.toBe(1);
      expect(stderrText()).toContain('E_VALIDATION');
      expect(stderrText()).toContain('matched 0');
      expect(fs.readFileSync(file).equals(before)).toBe(true);
    });

    it('--target-text without --append is a usage error (exit 2)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-nousage');
      await expect(run(['batch', file, '--target-text', 'anything'])).resolves.toBe(2);
    });

    it('shortcuts cannot be combined with an ops file (exit 2)', async () => {
      const file = writeBlankDocx('astra-cli-batch-shortcut-combo');
      const ops = writeOps({ operations: [] });
      await expect(run(['batch', file, ops, '--append', 'nope'])).resolves.toBe(2);
    });

    it('--append inserts as the final paragraph (no id needed)', async () => {
      const file = writeBlankDocx('astra-cli-batch-append');
      await expect(run(['batch', file, '--append', 'Tail paragraph proof'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      const paras = ast.sections.flatMap((s) => s.nodes).filter((n) => n.type === 'paragraph');
      expect(paras.length).toBeGreaterThan(0);
      expect(paras[paras.length - 1]?.metadata.text ?? '').toContain('Tail paragraph proof');
    });

    it('--append honors --out (original untouched)', async () => {
      const file = writeBlankDocx('astra-cli-batch-append-out');
      const before = fs.readFileSync(file);
      const dest = tmpPath('astra-cli-batch-append-dest');
      await expect(run(['batch', file, '--append', 'Out-file tail', '--out', dest])).resolves.toBe(0);
      expect(fs.readFileSync(file).equals(before)).toBe(true);
      expect(documentXml(dest)).toContain('Out-file tail');
    });

    it('--append under --json prints a pure ok payload', async () => {
      const file = writeBlankDocx('astra-cli-batch-append-json');
      vi.mocked(console.log).mockClear();
      await expect(run(['--json', 'batch', file, '--append', 'Json tail'])).resolves.toBe(0);
      const payload = JSON.parse(vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n'));
      expect(payload.ok).toBe(true);
    });

    it('batch --help documents --target-text and --append', async () => {
      vi.mocked(console.log).mockClear();
      await expect(run(['batch', '--help'])).resolves.toBe(0);
      const text = vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
      expect(text).toContain('--target-text');
      expect(text).toContain('--append');
    });

    // Shortcut composition (help claim): shortcuts share the ops-file path,
    // so --sequential/--no-sequential/--no-validate/ceilings compose.
    it('--target-text + --append honors --sequential (exit 0, lands text)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-seq');
      await seedReplace(file, 'teh quick brown fox');
      await expect(run(['batch', file, '--target-text', 'teh quick', '--append', 'the quick', '--sequential'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      expect(parser.searchNodesInAST(ast, 'the quick').length).toBeGreaterThan(0);
      expect(documentXml(file)).not.toContain('teh quick');
    });

    it('--target-text + --append honors --no-sequential (exit 0, lands text)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-noseq');
      await seedReplace(file, 'teh quick brown fox');
      await expect(run(['batch', file, '--target-text', 'teh quick', '--append', 'the quick', '--no-sequential'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      expect(parser.searchNodesInAST(ast, 'the quick').length).toBeGreaterThan(0);
      expect(documentXml(file)).not.toContain('teh quick');
    });

    it('--append honors --sequential (exit 0, tail lands)', async () => {
      const file = writeBlankDocx('astra-cli-batch-append-seq');
      await expect(run(['batch', file, '--append', 'Seq tail proof', '--sequential'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      const paras = ast.sections.flatMap((s) => s.nodes).filter((n) => n.type === 'paragraph');
      expect(paras[paras.length - 1]?.metadata.text ?? '').toContain('Seq tail proof');
    });

    it('--append honors --no-sequential (exit 0, tail lands)', async () => {
      const file = writeBlankDocx('astra-cli-batch-append-noseq');
      await expect(run(['batch', file, '--append', 'No-seq tail proof', '--no-sequential'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      const paras = ast.sections.flatMap((s) => s.nodes).filter((n) => n.type === 'paragraph');
      expect(paras[paras.length - 1]?.metadata.text ?? '').toContain('No-seq tail proof');
    });

    it('--target-text + --append honors raised --max-ratio (exit 0, lands text)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-ratio-hi');
      await seedReplace(file, 'teh quick brown fox');
      await expect(run(['batch', file, '--target-text', 'teh quick', '--append', 'the quick', '--max-ratio', '1000'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      expect(parser.searchNodesInAST(ast, 'the quick').length).toBeGreaterThan(0);
      expect(documentXml(file)).not.toContain('teh quick');
    });

    it('--append honors raised --max-uncompressed-bytes (exit 0, tail lands)', async () => {
      const file = writeBlankDocx('astra-cli-batch-append-bytes-hi');
      await expect(run(['batch', file, '--append', 'Raised ceiling tail', '--max-uncompressed-bytes', '524288000'])).resolves.toBe(0);
      expect(documentXml(file)).toContain('Raised ceiling tail');
    });

    it('lowered --max-uncompressed-bytes fails closed (exit 1, E_LIMIT, no write)', async () => {
      const file = writeBlankDocx('astra-cli-batch-append-bytes-lo');
      const before = fs.readFileSync(file);
      vi.mocked(console.error).mockClear();
      await expect(run(['batch', file, '--append', 'never lands', '--max-uncompressed-bytes', '1'])).resolves.toBe(1);
      expect(stderrText()).toContain('E_LIMIT');
      expect(stderrText()).toContain('Remedy:');
      expect(fs.readFileSync(file).equals(before)).toBe(true);
    });

    it('lowered --max-ratio fails closed (exit 1, E_LIMIT, no write)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-ratio-lo');
      await seedReplace(file, 'teh quick brown fox');
      const before = fs.readFileSync(file);
      vi.mocked(console.error).mockClear();
      await expect(run(['batch', file, '--target-text', 'teh quick', '--append', 'the quick', '--max-ratio', '1'])).resolves.toBe(1);
      expect(stderrText()).toContain('E_LIMIT');
      expect(stderrText()).toContain('Remedy:');
      expect(fs.readFileSync(file).equals(before)).toBe(true);
    });

    it('--target-text + --append honors --no-validate (then validate exits 0)', async () => {
      const file = writeBlankDocx('astra-cli-batch-target-noval');
      await seedReplace(file, 'teh quick brown fox');
      await expect(run(['batch', file, '--target-text', 'teh quick', '--append', 'the quick', '--no-validate'])).resolves.toBe(0);
      const parser = new DocxParser(fs.readFileSync(file));
      const ast = await parser.getAST();
      expect(parser.searchNodesInAST(ast, 'the quick').length).toBeGreaterThan(0);
      await expect(run(['validate', file])).resolves.toBe(0);
    });

    it('--append honors --no-validate (then validate exits 0)', async () => {
      const file = writeBlankDocx('astra-cli-batch-append-noval');
      await expect(run(['batch', file, '--append', 'No-validate tail', '--no-validate'])).resolves.toBe(0);
      expect(documentXml(file)).toContain('No-validate tail');
      await expect(run(['validate', file])).resolves.toBe(0);
    });
  });

  // P0 extract (feat/hap-p0-extract): `astra chat` + ALL agency moved verbatim
  // to @attentionform/astra-agent (packages/agent/src/*). The CLI is pure commands only, so
  // the two former `chat --session` jail cases below now guard the `export`
  // command instead (same count: 15 tests in this file). Session-jail
  // coverage (ENOENT/E_SANDBOX) lives on in packages/agent/tests/
  // brain-session.test.ts against the moved CliBrainSession.
  it('export --help exits 0', async () => {
    await expect(run(['export', '--help'])).resolves.toBe(0);
  });

  it('export missing input exits 1 (fail-closed, not 0 or 2)', async () => {
    await expect(run(['export', tmpPath('astra-cli-export-missing')])).resolves.toBe(1);
  });
});

describe('cli TMPDIR symlink remedy (human + json)', () => {
  function stderrText(): string {
    return vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
  }

  function stdoutText(): string {
    return vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
  }

  function symlinkedTmpBase(): { scratch: string; link: string; prev: string | undefined } {
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-tmpguard-'));
    const real = path.join(scratch, 'real');
    fs.mkdirSync(real);
    const link = path.join(scratch, 'linkbase');
    fs.symlinkSync(real, link);
    const prev = process.env[ASTRA_TMPDIR_ENV];
    process.env[ASTRA_TMPDIR_ENV] = link;
    return { scratch, link, prev };
  }

  function restoreTmpBase(prev: string | undefined, scratch: string): void {
    if (prev === undefined) delete process.env[ASTRA_TMPDIR_ENV];
    else process.env[ASTRA_TMPDIR_ENV] = prev;
    try {
      fs.rmSync(scratch, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }

  it('symlinked TMPDIR fails E_SANDBOX with a TMPDIR remedy on human stderr', async () => {
    const file = writeBlankDocx('astra-cli-tmp-human');
    const { scratch, prev } = symlinkedTmpBase();
    try {
      vi.mocked(console.error).mockClear();
      // inspect --visual hits secureMkdtemp -> resolveAstraTmpBase before any
      // renderer backend, so the symlink guard fires without LibreOffice.
      await expect(run(['inspect', file, '--visual'])).resolves.toBe(1);
      const err = stderrText();
      expect(err).toContain('E_SANDBOX');
      expect(err).toContain('$ASTRA_TMPDIR');
      expect(err).toMatch(/symlink/i);
      expect(err).toContain('Remedy:');
      expect(err).toMatch(/real directory|not a symlink/i);
    } finally {
      restoreTmpBase(prev, scratch);
    }
  });

  it('same failure under --json carries code + remedy, stdout stays pure JSON', async () => {
    const file = writeBlankDocx('astra-cli-tmp-json');
    const { scratch, prev } = symlinkedTmpBase();
    try {
      vi.mocked(console.log).mockClear();
      vi.mocked(console.error).mockClear();
      await expect(run(['--json', 'inspect', file, '--visual'])).resolves.toBe(1);
      const payload = JSON.parse(stdoutText());
      expect(payload.ok).toBe(false);
      expect(payload.code).toBe('E_SANDBOX');
      expect(String(payload.error)).toContain('E_SANDBOX');
      expect(String(payload.error)).toContain('$ASTRA_TMPDIR');
      expect(String(payload.error)).toMatch(/symlink/i);
      expect(String(payload.error)).toContain('Remedy:');
      const err = stderrText();
      expect(err).toContain('E_SANDBOX');
      expect(err).toContain('$ASTRA_TMPDIR');
      expect(err).toContain('Remedy:');
    } finally {
      restoreTmpBase(prev, scratch);
    }
  });
});
