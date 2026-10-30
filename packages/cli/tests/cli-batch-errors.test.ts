/**
 * CLI batch error-surface pins: engine remedies are rewritten to CLI verbs
 * (no MCP tool names leak), and gate misclassification is pinned in both
 * directions (array-level overflow = E_QUOTA, field-length = E_VALIDATION)
 * on BOTH gates (single batch + --stream line gate, which must agree).
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run, runStreamSession, type StreamSessionOpts } from '../src/cli';
import { DocxParser, executeChart } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function writeBlankDocx(prefix = 'astra-cli-errors'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
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

function writeOps(payload: unknown): string {
  const file = tmpPath('astra-cli-errors-ops', '.json');
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

describe('batch engine-remedy hygiene', () => {
  it('E_NOT_FOUND remedy names the CLI inspect command, never MCP tool names', async () => {
    const file = writeBlankDocx('astra-err-remedy');
    const target = await firstParagraphId(file);
    const ops = writeOps({
      operations: [
        { action: 'replace', target, text: 'Hello Riverbend world' },
        { action: 'style', target, match: 'zz-absent-needle', bold: true },
      ],
    });
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    const err = stderrText();
    expect(err).toContain('E_NOT_FOUND');
    expect(err).toContain('astra inspect');
    expect(err).not.toContain('inspect_document');
    expect(err).not.toContain('search_content');
    expect(err).not.toContain('get_document_toon');
  });

  it('--json E_NOT_FOUND remedy is equally clean', async () => {
    const file = writeBlankDocx('astra-err-remedy-json');
    const target = await firstParagraphId(file);
    const ops = writeOps({
      operations: [
        { action: 'replace', target, text: 'Hello Riverbend world' },
        { action: 'style', target, match: 'zz-absent-needle', bold: true },
      ],
    });
    vi.mocked(console.log).mockClear();
    vi.mocked(console.error).mockClear();
    await expect(run(['--json', 'batch', file, ops])).resolves.toBe(1);
    const payload = JSON.parse(stdoutText());
    expect(payload.code).toBe('E_NOT_FOUND');
    expect(String(payload.error)).toContain('astra inspect');
    expect(String(payload.error)).not.toContain('inspect_document');
    expect(stderrText()).not.toContain('inspect_document');
  });
});

describe('batch quota classification', () => {
  it('501 operations is E_QUOTA', async () => {
    const file = writeBlankDocx('astra-err-quota');
    const ops = writeOps({
      operations: Array.from({ length: 501 }, () => ({ action: 'delete', target: 'p_no_such_node' })),
    });
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    const err = stderrText();
    expect(err).toContain('E_QUOTA');
  });

  it('over-long match field is E_VALIDATION, never E_QUOTA', async () => {
    const file = writeBlankDocx('astra-err-fieldlen');
    const target = await firstParagraphId(file);
    const ops = writeOps({
      operations: [{ action: 'style', target, match: 'x'.repeat(1001), bold: true }],
    });
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    const err = stderrText();
    expect(err).toContain('E_VALIDATION');
    expect(err).not.toContain('E_QUOTA');
  });
});

const OUT = { json: false, verbose: false };

function streamOptsFor(file: string, docx: Buffer): StreamSessionOpts {
  return {
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
  };
}

async function streamAcks(file: string, lines: string[]): Promise<Record<string, unknown>[]> {
  const docx = fs.readFileSync(file);
  const acks: Record<string, unknown>[] = [];
  await runStreamSession(OUT, docx, streamOptsFor(file, docx), lines, (a) => {
    acks.push(a);
  });
  return acks;
}

async function unparseableChartFixture(prefix: string): Promise<{ file: string; chartId: string }> {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  executeChart(
    {
      action: 'chart',
      chartType: 'column',
      title: 'Refusal Probe',
      categories: ['Q1', 'Q2'],
      series: [
        { name: 'ARR', values: [10, 20] },
        { name: 'Costs', values: [5, 7] },
      ],
      anchor,
      id: 'chart_refusal_probe',
    },
    ast,
    parser
  );
  const stripped = parser
    .getZipAdapter()
    .readText('word/charts/chart1.xml')
    .replace(/<c:numCache>[\s\S]*?<\/c:numCache>/g, '');
  parser.getZipAdapter().writeText('word/charts/chart1.xml', stripped);
  const file = tmpPath(prefix);
  fs.writeFileSync(file, parser.rezip());
  const probe = new DocxParser(fs.readFileSync(file));
  const probeAst = await probe.getAST();
  const chart = probeAst.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart' || n.type === 'chart3d');
  if (!chart) throw new Error('refusal fixture has no chart node');
  return { file, chartId: chart.id };
}

function chartRefusalOps(chartId: string): string {
  return writeOps({
    operations: [
      {
        action: 'chart',
        operation: 'update',
        target: chartId,
        series: [
          { name: 'ARR', values: [99, 20] },
          { name: 'Costs', values: [5, 7] },
        ],
      },
    ],
    sequential: true,
  });
}

describe('batch chart refusal remedy', () => {
  it('human output keeps the authored remedy, not the generic id hint', async () => {
    const { file, chartId } = await unparseableChartFixture('astra-err-chart-refusal');
    const ops = chartRefusalOps(chartId);
    const before = fs.readFileSync(file);
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    const err = stderrText();
    expect(err).toContain('has no cached numeric values');
    expect(err).toContain('Remedy: re-read the chart via astra inspect');
    expect(err).toContain("delete it and re-insert it with action:'chart'");
    expect(err).not.toContain('re-inspect ids via `astra inspect`');
    expect(err.match(/Remedy:/g)).toHaveLength(1);
    expect(err).not.toContain('Remedy: Remedy:');
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('--json envelope carries one Remedy prefix and the authored remedy', async () => {
    const { file, chartId } = await unparseableChartFixture('astra-err-chart-refusal-json');
    const ops = chartRefusalOps(chartId);
    const before = fs.readFileSync(file);
    vi.mocked(console.error).mockClear();
    await expect(run(['--json', 'batch', file, ops])).resolves.toBe(1);
    const payload = JSON.parse(stdoutText()) as {
      error: string;
      code: string;
      errors: Array<{ operationIndex: number; error: string }>;
    };
    expect(payload.code).toBe('E_VALIDATION');
    expect(payload.error).toContain('Remedy: re-read the chart via astra inspect');
    expect(payload.error.match(/Remedy:/g)).toHaveLength(1);
    expect(payload.error).not.toContain('Remedy: Remedy:');
    expect(payload.errors.every((e) => !e.error.includes('Remedy: Remedy:'))).toBe(true);
    const rollback = payload.errors.find((e) => e.operationIndex === -1);
    expect(rollback?.error).toContain('Remedy: re-read the chart via astra inspect');
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
});

describe('stream quota classification (must agree with the single-batch gate)', () => {
  it('501 operations on one stream line is E_QUOTA', async () => {
    const file = writeBlankDocx('astra-err-stream-quota');
    const line = JSON.stringify({
      operations: Array.from({ length: 501 }, () => ({ action: 'delete', target: 'p_no_such_node' })),
    });
    const acks = await streamAcks(file, [line]);
    expect(acks.length).toBeGreaterThan(0);
    expect(acks[0].code).toBe('E_QUOTA');
  });

  it('over-long match field on one stream line is E_VALIDATION, never E_QUOTA', async () => {
    const file = writeBlankDocx('astra-err-stream-fieldlen');
    const target = await firstParagraphId(file);
    const line = JSON.stringify({
      operations: [{ action: 'style', target, match: 'x'.repeat(1001), bold: true }],
    });
    const acks = await streamAcks(file, [line]);
    expect(acks.length).toBeGreaterThan(0);
    expect(acks[0].code).toBe('E_VALIDATION');
  });
});
