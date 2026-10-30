import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run, runStreamSession, ANSI_REGEX, type StreamSessionOpts } from '../src/cli';
import { DocxParser } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function writeBlankDocx(prefix = 'astra-cli-stream'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
  return file;
}

function writeOpsFile(prefix: string, payload: unknown): string {
  const file = tmpPath(prefix, '.json');
  fs.writeFileSync(file, JSON.stringify(payload));
  return file;
}

async function docContains(file: string, text: string): Promise<boolean> {
  const parser = new DocxParser(fs.readFileSync(file));
  const ast = await parser.getAST();
  return parser.searchNodesInAST(ast, text).length > 0;
}

function stdoutText(): string {
  return vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
}

function streamOpts(
  file: string,
  docx: Buffer,
  extra?: Partial<StreamSessionOpts>
): StreamSessionOpts {
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
    ...extra,
  };
}

const OUT = { json: false, verbose: false };

function appendLine(text: string): string {
  return JSON.stringify({
    operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text }],
  });
}

const BAD_LINE = JSON.stringify({ operations: [{ action: 'definitely_not_a_real_op' }] });

async function runLines(
  file: string,
  lines: string[],
  extra?: Partial<StreamSessionOpts>
): Promise<{ code: number; acks: Record<string, unknown>[] }> {
  const docx = fs.readFileSync(file);
  const acks: Record<string, unknown>[] = [];
  const code = await runStreamSession(OUT, docx, streamOpts(file, docx, extra), lines, (a) => {
    acks.push(a);
  });
  return { code, acks };
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

describe('batch --stream flags', () => {
  it('help documents the stream flags', async () => {
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', '--help'])).resolves.toBe(0);
    const text = stdoutText();
    expect(text).toContain('--stream');
    expect(text).toContain('--on-error');
    expect(text).toContain('--max-errors');
    expect(text).toContain('{"control":"save"}');
  });

  it('stream requires the stdin ops path', async () => {
    const file = writeBlankDocx('astra-stream-needs-dash');
    const ops = writeOpsFile('astra-stream-needs-dash-ops', { operations: [] });
    await expect(run(['batch', file, ops, '--stream'])).resolves.toBe(2);
  });

  it('bad on-error value is a usage error', async () => {
    const file = writeBlankDocx('astra-stream-onerror');
    await expect(run(['batch', file, '-', '--stream', '--on-error=bogus'])).resolves.toBe(2);
  });

  it('bad max-errors value is a usage error', async () => {
    const file = writeBlankDocx('astra-stream-maxerr');
    await expect(run(['batch', file, '-', '--stream', '--max-errors=-1'])).resolves.toBe(2);
  });

  it('on-error without stream is a usage error', async () => {
    const file = writeBlankDocx('astra-stream-onerror-nostream');
    const ops = writeOpsFile('astra-stream-onerror-nostream-ops', { operations: [] });
    await expect(run(['batch', file, ops, '--on-error=continue'])).resolves.toBe(2);
  });

  it('max-errors without stream is a usage error', async () => {
    const file = writeBlankDocx('astra-stream-maxerr-nostream');
    const ops = writeOpsFile('astra-stream-maxerr-nostream-ops', { operations: [] });
    await expect(run(['batch', file, ops, '--max-errors=5'])).resolves.toBe(2);
  });

  it('stream refuses matrix input', async () => {
    const file = writeBlankDocx('astra-stream-matrix');
    const csv = tmpPath('astra-stream-matrix-vars', '.csv');
    fs.writeFileSync(csv, 'name\nAda\n');
    await expect(run(['batch', file, '-', '--stream', '--matrix', csv, '--out', tmpPath('o')])).resolves.toBe(2);
  });

  it('stream refuses dry-run input', async () => {
    const file = writeBlankDocx('astra-stream-dry');
    await expect(run(['batch', file, '-', '--stream', '--dry-run'])).resolves.toBe(2);
  });

  it('stream refuses shortcut input', async () => {
    const file = writeBlankDocx('astra-stream-short');
    await expect(run(['batch', file, '-', '--stream', '--append', 'hi'])).resolves.toBe(2);
  });
});

describe('batch --stream session', () => {
  it('applies three batches, saves once at EOF, exits 0', async () => {
    const file = writeBlankDocx('astra-stream-ok');
    const { code, acks } = await runLines(file, [appendLine('first proof'), appendLine('second proof'), appendLine('third proof')]);
    expect(code).toBe(0);
    expect(acks.length).toBe(4);
    for (const ack of acks.slice(0, 3)) {
      expect(ack.ok).toBe(true);
      expect(typeof (ack as { appliedCount: number }).appliedCount).toBe('number');
      expect((ack as { appliedCount: number }).appliedCount).toBe(1);
    }
    const summary = acks[3] as Record<string, unknown>;
    expect(summary).toMatchObject({ ok: true, control: 'eof', batchesApplied: 3, batchesFailed: 0, revision: 1, exitCode: 0 });
    expect(await docContains(file, 'first proof')).toBe(true);
    expect(await docContains(file, 'second proof')).toBe(true);
    expect(await docContains(file, 'third proof')).toBe(true);
  });

  it('every ack serializes as one clean JSON line', async () => {
    const file = writeBlankDocx('astra-stream-pure');
    const { acks } = await runLines(file, [
      appendLine('pure proof'),
      BAD_LINE,
      '{"control":"save"}',
    ], { onError: 'continue' });
    expect(acks.length).toBeGreaterThan(0);
    for (const ack of acks) {
      const line = JSON.stringify(ack);
      expect(line).not.toContain('\n');
      expect(() => JSON.parse(line)).not.toThrow();
      expect(new RegExp(ANSI_REGEX.source).test(line)).toBe(false);
      expect(line).not.toContain('BATCH OK');
    }
  });

  it('checkpoint flushes to disk and stays resident', async () => {
    const file = writeBlankDocx('astra-stream-checkpoint');
    const { code, acks } = await runLines(file, [
      appendLine('before checkpoint'),
      '{"control":"save"}',
      appendLine('after checkpoint'),
    ]);
    expect(code).toBe(0);
    expect(acks[1]).toMatchObject({ ok: true, control: 'save', revision: 1 });
    expect(acks[3]).toMatchObject({ ok: true, control: 'eof', batchesApplied: 2, batchesFailed: 0, revision: 2, exitCode: 0 });
    expect(await docContains(file, 'before checkpoint')).toBe(true);
    expect(await docContains(file, 'after checkpoint')).toBe(true);
  });

  it('abort discards uncheckpointed state, disk untouched', async () => {
    const file = writeBlankDocx('astra-stream-abort-clean');
    const before = fs.readFileSync(file);
    const { code, acks } = await runLines(file, [appendLine('doomed proof'), BAD_LINE]);
    expect(code).toBe(1);
    expect(acks.length).toBe(3);
    expect(acks[0]).toMatchObject({ ok: true, index: 0 });
    expect(acks[1]).toMatchObject({ ok: false, index: 1 });
    expect(String((acks[1] as { code: string }).code)).toMatch(/E_/);
    expect(acks[2]).toMatchObject({ ok: false, control: 'eof', batchesApplied: 1, batchesFailed: 1, revision: 0, exitCode: 1, reason: 'aborted' });
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('disk after abort equals the last checkpoint bytes', async () => {
    const solo = writeBlankDocx('astra-stream-solo');
    const { code: soloCode } = await runLines(solo, [appendLine('checkpoint proof')]);
    expect(soloCode).toBe(0);
    const checkpointBytes = fs.readFileSync(solo);

    const file = writeBlankDocx('astra-stream-abort-checkpoint');
    const { code, acks } = await runLines(file, [
      appendLine('checkpoint proof'),
      '{"control":"save"}',
      appendLine('lost proof'),
      BAD_LINE,
    ]);
    expect(code).toBe(1);
    expect(acks[acks.length - 1]).toMatchObject({ reason: 'aborted', revision: 1 });
    // Disk holds the checkpoint (first batch only): identical to the solo run.
    expect(fs.readFileSync(file).equals(checkpointBytes)).toBe(true);
    expect(await docContains(file, 'lost proof')).toBe(false);
  });

  it('action save is an op failure, never a checkpoint', async () => {
    const file = writeBlankDocx('astra-stream-actionsave');
    const before = fs.readFileSync(file);
    const { code, acks } = await runLines(file, [JSON.stringify({ action: 'save' })]);
    expect(code).toBe(1);
    expect(acks[0]).toMatchObject({ ok: false, index: 0 });
    expect('control' in (acks[0] as Record<string, unknown>)).toBe(false);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('continue rolls back only the failed batch and stays resident', async () => {
    const file = writeBlankDocx('astra-stream-continue');
    const { code, acks } = await runLines(file, [BAD_LINE, appendLine('survivor proof')], { onError: 'continue' });
    expect(code).toBe(1);
    expect(acks.length).toBe(3);
    expect(acks[0]).toMatchObject({ ok: false, index: 0 });
    expect(acks[1]).toMatchObject({ ok: true, index: 1 });
    expect(acks[2]).toMatchObject({ ok: false, control: 'eof', batchesApplied: 1, batchesFailed: 1, exitCode: 1 });
    expect(await docContains(file, 'survivor proof')).toBe(true);
  });

  it('breaker trips on consecutive failures and stops reading', async () => {
    const file = writeBlankDocx('astra-stream-breaker');
    const { code, acks } = await runLines(file, [BAD_LINE, BAD_LINE, appendLine('never proof')], {
      onError: 'continue',
      maxErrors: 2,
    });
    expect(code).toBe(1);
    expect(acks.length).toBe(3);
    expect(acks[1]).toMatchObject({ ok: false, index: 1, reason: 'max_errors_exceeded' });
    expect(acks[2]).toMatchObject({ ok: false, control: 'eof', reason: 'max_errors_exceeded', batchesApplied: 0, batchesFailed: 2, exitCode: 1 });
    expect(await docContains(file, 'never proof')).toBe(false);
  });

  it('breaker resets on success', async () => {
    const file = writeBlankDocx('astra-stream-reset');
    const { code, acks } = await runLines(file, [BAD_LINE, appendLine('reset proof'), BAD_LINE], {
      onError: 'continue',
      maxErrors: 2,
    });
    expect(code).toBe(1);
    expect(acks.length).toBe(4);
    for (const ack of acks.slice(0, 3)) {
      expect('reason' in ack).toBe(false);
    }
    expect(acks[3]).toMatchObject({ control: 'eof', batchesApplied: 1, batchesFailed: 2, exitCode: 1 });
    expect('reason' in (acks[3] as Record<string, unknown>)).toBe(false);
    expect(await docContains(file, 'reset proof')).toBe(true);
  });

  it('max-errors zero disables the breaker', async () => {
    const file = writeBlankDocx('astra-stream-nobreaker');
    const { code, acks } = await runLines(file, [BAD_LINE, BAD_LINE, BAD_LINE], {
      onError: 'continue',
      maxErrors: 0,
    });
    expect(code).toBe(1);
    expect(acks.length).toBe(4);
    for (const ack of acks.slice(0, 3)) {
      expect(ack).toMatchObject({ ok: false });
      expect('reason' in ack).toBe(false);
    }
    expect(acks[3]).toMatchObject({ control: 'eof', batchesFailed: 3, exitCode: 1 });
  });

  it('unknown control fails the line', async () => {
    const file = writeBlankDocx('astra-stream-control');
    const before = fs.readFileSync(file);
    const { code, acks } = await runLines(file, ['{"control":"revert"}']);
    expect(code).toBe(1);
    expect(acks[0]).toMatchObject({ ok: false, index: 0 });
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('malformed JSON line fails under abort', async () => {
    const file = writeBlankDocx('astra-stream-malformed');
    const before = fs.readFileSync(file);
    const { code, acks } = await runLines(file, ['{oops']);
    expect(code).toBe(1);
    expect(acks[0]).toMatchObject({ ok: false, index: 0, code: 'E_VALIDATION' });
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('blank lines are skipped without an ack', async () => {
    const file = writeBlankDocx('astra-stream-blank');
    const { code, acks } = await runLines(file, ['', '   ', appendLine('blank proof')]);
    expect(code).toBe(0);
    expect(acks.length).toBe(2);
    expect(acks[0]).toMatchObject({ ok: true, index: 0, appliedCount: 1 });
    expect(acks[1]).toMatchObject({ control: 'eof', batchesApplied: 1, exitCode: 0 });
    expect(await docContains(file, 'blank proof')).toBe(true);
  });

  it('clearDoc batch works inside the resident session', async () => {
    const file = writeBlankDocx('astra-stream-clear');
    const { code, acks } = await runLines(file, [
      JSON.stringify({ operations: [{ action: 'clearDoc' }] }),
      appendLine('fresh proof'),
    ]);
    expect(code).toBe(0);
    expect(acks[0]).toMatchObject({ ok: true, appliedCount: 1 });
    expect(await docContains(file, 'fresh proof')).toBe(true);
    expect(await docContains(file, 'Blank Document')).toBe(false);
  });

  it('empty stream saves once and exits 0', async () => {
    const file = writeBlankDocx('astra-stream-empty');
    const { code, acks } = await runLines(file, []);
    expect(code).toBe(0);
    expect(acks.length).toBe(1);
    expect(acks[0]).toMatchObject({ ok: true, control: 'eof', batchesApplied: 0, batchesFailed: 0, revision: 1, exitCode: 0 });
  });
});
