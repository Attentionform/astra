/**
 * `astra revert`, `redo` and `preview-revert` against comment metadata and old history stores.
 * Commits come from the real MCP save path; the revert is run from the command line, as a user would.
 * Every assertion reads decompressed zip entries.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import git from 'isomorphic-git';
import { run } from '../src/cli';
import {
  AdmZip,
  createBlankDocxBuffer,
  getHistory,
  getRepoDir,
  isLegacyTrackedPackageFile,
} from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../mcp-server/src/lib/session-store';
import { LEAN6_HANDLERS } from '../../mcp-server/src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const MARKER = '.astra-snapshot-format';
const tmpDirs: string[] = [];
type Store = ReturnType<typeof createMcpSessionStore>;

const stdoutText = (): string => vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
const stderrText = (): string => vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');

async function runCli(args: string[]): Promise<{ status: number; stdout: string; stderr: string }> {
  vi.mocked(console.log).mockClear();
  vi.mocked(console.error).mockClear();
  const status = await run(args);
  return { status, stdout: stdoutText(), stderr: stderrText() };
}

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  while (tmpDirs.length > 0) fs.rmSync(tmpDirs.pop() as string, { recursive: true, force: true });
});

async function session(): Promise<{ store: Store; sessionId: string; file: string }> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-parts-'));
  tmpDirs.push(dir);
  const file = path.join(dir, 'doc.docx');
  fs.writeFileSync(file, createBlankDocxBuffer('blank'));
  const store = createMcpSessionStore();
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  if (!opened.ok) throw new Error('open failed in test setup');
  return { store, sessionId: opened.sessionId, file };
}

async function firstNodeId(store: Store, sessionId: string): Promise<string> {
  const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId, offset: 0, limit: 1 });
  if (!first.ok) throw new Error('inspect failed in test setup');
  const id = first.toonText
    .split('\n')
    .map((l) => l.split(/\s+/)[0])
    .filter((x) => !!x && !['>', 'ID'].includes(x) && !/^---/.test(x))[0];
  if (!id) throw new Error('no node id');
  return id;
}

async function save(store: Store, sessionId: string, message: string): Promise<void> {
  const res = await LEAN6_HANDLERS.save_document(store, { sessionId, message });
  if (!res.ok) throw new Error(`save failed: ${JSON.stringify(res)}`);
}

async function addComment(store: Store, sessionId: string, author: string): Promise<void> {
  const target = await firstNodeId(store, sessionId);
  const res = await LEAN6_HANDLERS.execute_batch(store, {
    sessionId,
    operations: [{ action: 'comment', operation: 'insert', target, text: 'please review', author }],
  } as never);
  if (!res.ok) throw new Error(`batch failed: ${JSON.stringify(res)}`);
}

function parts(file: string): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  for (const e of new AdmZip(fs.readFileSync(file)).getEntries()) if (!e.isDirectory) out.set(e.entryName, e.getData());
  return out;
}

const anywhere = (file: string, needle: string): string[] =>
  [...parts(file)].filter(([, b]) => b.toString('latin1').includes(needle)).map(([n]) => n);

describe('astra revert past a comment', () => {
  it('leaves no comment parts and no author name, and redo brings them back', async () => {
    const { store, sessionId, file } = await session();
    await save(store, sessionId, 'start');
    const before = parts(file);
    await addComment(store, sessionId, 'Jane Secret Reviewer');
    await save(store, sessionId, 'with comment');
    const withComment = parts(file);
    expect(anywhere(file, 'Jane Secret Reviewer')).toContain('word/people.xml');

    const preview = await runCli(['--json', 'preview-revert', file, '--steps', '1']);
    expect(preview.status).toBe(0);
    const previewJson = JSON.parse(preview.stdout.split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(previewJson.snapshotComplete).toBe(true);
    expect(previewJson.restoreWarning).toBeUndefined();
    expect(previewJson.affectedFiles).toEqual(expect.arrayContaining(['word/people.xml', 'word/commentsExtended.xml']));

    const reverted = await runCli(['revert', file, '--steps', '1']);
    expect(reverted.status).toBe(0);
    expect(reverted.stderr).not.toContain('warning');
    expect([...parts(file).keys()].sort()).toEqual([...before.keys()].sort());
    expect(anywhere(file, 'Jane Secret Reviewer')).toEqual([]);

    const redone = await runCli(['redo', file]);
    expect(redone.status).toBe(0);
    expect([...parts(file).keys()].sort()).toEqual([...withComment.keys()].sort());
    for (const [name, buf] of withComment) expect(parts(file).get(name)!.equals(buf), name).toBe(true);
  });
});

describe('astra revert against an old history store', () => {
  async function rewriteHeadAsLegacy(file: string): Promise<string> {
    const repoDir = getRepoDir(file);
    for (const rel of await git.listFiles({ fs, dir: repoDir })) {
      if (rel === 'document.xml') continue;
      if (rel === MARKER || !isLegacyTrackedPackageFile(rel)) {
        await git.remove({ fs, dir: repoDir, filepath: rel });
        await fs.promises.rm(path.join(repoDir, rel), { force: true });
      }
    }
    return git.commit({
      fs,
      dir: repoDir,
      author: { name: 'ASTra Agent', email: 'astra@attentionform.com' },
      message: 'legacy snapshot',
    });
  }

  it('warns, keeps the parts the old commit never recorded, and does not report a false conflict', async () => {
    const { store, sessionId, file } = await session();
    await save(store, sessionId, 'start');
    await rewriteHeadAsLegacy(file);
    // The working file is unchanged since the last (old) commit: a revert must not call that a conflict.
    const unrecorded = [...parts(file).keys()].filter((n) => !isLegacyTrackedPackageFile(n) && n !== 'word/document.xml');
    expect(unrecorded.length).toBeGreaterThan(0);
    await addComment(store, sessionId, 'Alice');
    await save(store, sessionId, 'new commit after the old one');
    const history = await getHistory(file);
    expect(history.length).toBeGreaterThanOrEqual(3);

    const preview = await runCli(['--json', 'preview-revert', file, '--steps', '1']);
    expect(preview.status).toBe(0);
    const previewJson = JSON.parse(preview.stdout.split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(previewJson.snapshotComplete).toBe(false);
    expect(previewJson.restoreWarning).toMatch(/older ASTra version/i);
    expect(preview.stderr).toMatch(/not exact/i);

    const reverted = await runCli(['--json', 'revert', file, '--steps', '1']);
    expect(reverted.status).toBe(0);
    expect(reverted.stderr).toMatch(/older ASTra version/i);
    const revertedJson = JSON.parse(reverted.stdout.split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(revertedJson.snapshotComplete).toBe(false);
    for (const name of unrecorded) expect(parts(file).has(name), `${name} was deleted`).toBe(true);
  });

  it('a hand edit to an unrecorded part is not a conflict when the head commit is an old one', async () => {
    const { store, sessionId, file } = await session();
    await save(store, sessionId, 'start');
    await rewriteHeadAsLegacy(file);
    const zip = new AdmZip(fs.readFileSync(file));
    const core = zip.getEntry('docProps/core.xml');
    expect(core).not.toBeNull();
    zip.updateFile('docProps/core.xml', Buffer.concat([core!.getData(), Buffer.from('\n<!-- edit -->\n')]));
    fs.writeFileSync(file, zip.toBuffer());
    const history = await getHistory(file);
    expect(history.length).toBeGreaterThanOrEqual(2);
    const res = await runCli(['revert', file, '--steps', '1']);
    expect(res.status).toBe(0);
  });
});
