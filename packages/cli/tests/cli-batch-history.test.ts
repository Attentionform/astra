/**
 * `astra batch` records a history revision when the document already has a history repository, and
 * only then: no folder is created for a document without history, `--no-history` and `--out` record
 * nothing, and a history failure does not undo the saved edit.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run, runStreamSession, type StreamSessionOpts } from '../src/cli';
import { AdmZip, createBlankDocxBuffer, getRepoDir, getHistoryCommitCount } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpDirs: string[] = [];

function tmpDir(prefix: string): string {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), `${prefix}-`));
  tmpDirs.push(d);
  return d;
}

function setup(prefix: string): { dir: string; file: string; ops: string } {
  const dir = tmpDir(prefix);
  const file = path.join(dir, 'doc.docx');
  fs.writeFileSync(file, createBlankDocxBuffer('blank'));
  return { dir, file, ops: writeOps(dir, 'first edit') };
}

function writeOps(dir: string, text: string): string {
  const ops = path.join(dir, `ops-${text.replace(/\W+/g, '-')}.json`);
  fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text }] }));
  return ops;
}

const documentText = (file: string): string =>
  new AdmZip(fs.readFileSync(file)).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
const commits = (file: string): Promise<number> => getHistoryCommitCount(getRepoDir(file));
const stdoutText = (): string => vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
const stderrText = (): string => vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
const lastJson = (): Record<string, any> => JSON.parse(stdoutText().split('\n').filter((l) => l.startsWith('{')).pop() as string);

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  while (tmpDirs.length > 0) {
    fs.rmSync(tmpDirs.pop() as string, { recursive: true, force: true });
  }
});

describe('batch and the history repository', () => {
  it('adds one revision, which revert undoes and redo applies again', async () => {
    const { file, ops } = setup('astra-bh-record');
    await run(['history', file, '--init']);
    const before = await commits(file);
    await expect(run(['--json', 'batch', file, ops])).resolves.toBe(0);
    expect(await commits(file)).toBe(before + 1);
    expect(lastJson().history).toMatchObject({ recorded: true });
    expect(documentText(file)).toContain('first edit');

    await expect(run(['revert', file, '--steps', '1'])).resolves.toBe(0);
    expect(documentText(file)).not.toContain('first edit');
    await expect(run(['redo', file])).resolves.toBe(0);
    expect(documentText(file)).toContain('first edit');
  });

  it('records the number of operations and the action names in the message', async () => {
    const { file, ops } = setup('astra-bh-message');
    await run(['history', file, '--init']);
    await run(['batch', file, ops]);
    vi.mocked(console.log).mockClear();
    await run(['--json', 'history', file]);
    expect(stdoutText()).toContain('batch: 1 operation(s) (insert)');
  });

  it('records nothing with --no-history', async () => {
    const { file, ops } = setup('astra-bh-off');
    await run(['history', file, '--init']);
    const before = await commits(file);
    await expect(run(['batch', file, ops, '--no-history'])).resolves.toBe(0);
    expect(await commits(file)).toBe(before);
  });

  it('creates no history folder for a document without history', async () => {
    const { dir, file, ops } = setup('astra-bh-none');
    await expect(run(['--json', 'batch', file, ops])).resolves.toBe(0);
    expect(fs.existsSync(path.join(dir, '.astra'))).toBe(false);
    expect(lastJson().history).toBeUndefined();
  });

  it('records nothing when the result is written to another file', async () => {
    const { dir, file, ops } = setup('astra-bh-out');
    await run(['history', file, '--init']);
    const before = await commits(file);
    await expect(run(['batch', file, ops, '--out', path.join(dir, 'copy.docx')])).resolves.toBe(0);
    expect(await commits(file)).toBe(before);
    expect(fs.existsSync(path.join(dir, '.astra', 'copy'))).toBe(false);
  });

  it('keeps the saved edit and exit code 0 when the revision cannot be recorded', async () => {
    const { file, ops } = setup('astra-bh-fail');
    await run(['history', file, '--init']);
    // A sidecar whose objects folder is a file: the commit cannot be written.
    const repo = getRepoDir(file);
    const objects = path.join(repo, '.git', 'objects');
    fs.rmSync(objects, { recursive: true, force: true });
    fs.writeFileSync(objects, 'not a folder');
    vi.mocked(console.error).mockClear();
    await expect(run(['--json', 'batch', file, ops])).resolves.toBe(0);
    expect(documentText(file)).toContain('first edit');
    expect(lastJson().history).toMatchObject({ recorded: false });
    expect(stderrText()).toContain('the edit was saved but the history revision was not recorded');
  });
});

describe('stream mode and the history repository', () => {
  it('records one revision for a flush after two batches', async () => {
    const { dir, file } = setup('astra-bh-stream');
    await run(['history', file, '--init']);
    const before = await commits(file);
    const lines = [writeOps(dir, 'one'), writeOps(dir, 'two')].map((p) => fs.readFileSync(p, 'utf8').trim());
    const docx = fs.readFileSync(file);
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
    };
    const code = await runStreamSession({ json: false, verbose: false }, docx, opts, lines, () => undefined);
    expect(code).toBe(0);
    expect(documentText(file)).toContain('two');
    expect(await commits(file)).toBe(before + 1);
  });
});
