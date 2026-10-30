import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import git from 'isomorphic-git';
import { run } from '../src/cli';
import { AdmZip, DocxParser, checkoutCommit, createBlankDocxBuffer, getHistory, getHistoryCommitCount } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../mcp-server/src/lib/session-store';
import { GIT_HISTORY_HANDLERS, LEAN6_HANDLERS } from '../../mcp-server/src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

// Interception seam for the vanished-sidecar race pin: the CLI's file writer
// is mocked (pass-through unless armed) so one test can delete the .astra
// sidecar exactly between the restore write and the commit-path re-check.
const vanishHook = vi.hoisted(() => ({ sidecarRoot: '', armed: false }));

vi.mock('../src/cli/io', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/cli/io')>();
  return {
    ...actual,
    writeDocx: async (file: string, buf: Buffer): Promise<void> => {
      if (vanishHook.armed) {
        vanishHook.armed = false;
        const { rmSync } = await import('fs');
        rmSync(vanishHook.sidecarRoot, { recursive: true, force: true });
      }
      return actual.writeDocx(file, buf);
    },
  };
});

// Passive Git sidecar CLI pins: read-only history/preview, fail-closed
// revert/redo (sidecar required, diverged working files refuse), explicit
// history --init opt-in, and the listing-proof that batch never creates
// .astra/. Commits are produced through the real MCP save path so the
// history read is proven against MCP-created sidecars.

const tmpDirs: string[] = [];

function tmpDir(prefix: string): string {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), `${prefix}-`));
  tmpDirs.push(d);
  return d;
}

function writeBlank(prefix: string): { dir: string; file: string } {
  const dir = tmpDir(prefix);
  const file = path.join(dir, 'doc.docx');
  fs.writeFileSync(file, createBlankDocxBuffer('blank'));
  return { dir, file };
}

/** First whitespace-delimited token of every TOON data row (skips header lines). */
function rowIds(toonText: string): string[] {
  return toonText
    .split('\n')
    .map((l) => l.split(/\s+/)[0])
    .filter((id) => !!id && !['>', 'ID'].includes(id) && !/^---/.test(id));
}

async function firstNodeId(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string,
): Promise<string> {
  const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId, offset: 0, limit: 1 });
  expect(first.ok).toBe(true);
  if (!first.ok) throw new Error('inspect failed in test setup');
  const id = rowIds(first.toonText)[0];
  expect(typeof id).toBe('string');
  return id;
}

async function replaceAnchor(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string,
  text: string,
): Promise<void> {
  const anchor = await firstNodeId(store, sessionId);
  const batch = await LEAN6_HANDLERS.execute_batch(store, {
    sessionId,
    operations: [{ action: 'replace', target: anchor, text }],
  });
  expect(batch.ok).toBe(true);
}

interface TwoStateDoc {
  file: string;
  dir: string;
  store: ReturnType<typeof createMcpSessionStore>;
  sessionId: string;
  state1Bytes: Buffer;
  state2Bytes: Buffer;
}

/** Open (baseline commit) then two recorded edits: checkpoint one, checkpoint two. */
async function setupTwoSaves(prefix: string): Promise<TwoStateDoc> {
  const { dir, file } = writeBlank(prefix);
  const store = createMcpSessionStore();
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  expect(opened.ok).toBe(true);
  if (!opened.ok) throw new Error('open failed in test setup');
  const sessionId = opened.sessionId;
  await replaceAnchor(store, sessionId, 'Alpha state one');
  const save1 = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'checkpoint one' });
  expect(save1.ok).toBe(true);
  const state1Bytes = fs.readFileSync(file);
  await replaceAnchor(store, sessionId, 'Beta state two');
  const save2 = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'checkpoint two' });
  expect(save2.ok).toBe(true);
  const state2Bytes = fs.readFileSync(file);
  return { file, dir, store, sessionId, state1Bytes, state2Bytes };
}

function stdoutText(): string {
  return vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
}

function stderrText(): string {
  return vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
}

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
  while (tmpDirs.length > 0) {
    const d = tmpDirs.pop() as string;
    try {
      fs.rmSync(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

describe('history reads MCP-created sidecar commits', () => {
  it('--json lists every recorded commit newest-first and matches the MCP view', async () => {
    const { file, store, sessionId } = await setupTwoSaves('astra-hist-list');
    const mcp = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    expect(mcp.ok).toBe(true);
    if (!mcp.ok) return;
    expect(mcp.currentCommit).not.toBeNull();
    const mcpHead = mcp.currentCommit ?? '';

    const r = await runCli(['--json', 'history', file]);
    expect(r.status).toBe(0);
    const payload = JSON.parse(r.stdout);
    expect(payload.ok).toBe(true);
    expect(payload.command).toBe('history');
    expect(payload.totalCommits).toBe(3);
    expect(payload.history.map((c: { message: string }) => c.message)).toEqual([
      'checkpoint two',
      'checkpoint one',
      'Initial state',
    ]);
    expect(payload.currentCommit).toBe(mcpHead);
    expect(payload.history.map((c: { oid: string }) => c.oid)).toEqual(mcp.history.map((c) => c.oid));
    for (const c of payload.history) expect(c.oid).toMatch(/^[0-9a-f]{40}$/);

    const human = await runCli(['history', file]);
    expect(human.status).toBe(0);
    expect(human.stdout).toContain(`head ${mcpHead.slice(0, 7)}`);
    expect(human.stdout).toContain('checkpoint two');
  });
});

describe('revert restores the recorded state and redo round-trips', () => {
  it('revert --steps 1 byte-restores the prior state and records a revert commit', async () => {
    const { file, state1Bytes } = await setupTwoSaves('astra-hist-revert');
    const commits = await getHistory(file);
    const target = commits[1];
    const targetPkg = await checkoutCommit(file, target.oid);

    const r = await runCli(['--json', 'revert', file, '--steps', '1']);
    expect(r.status).toBe(0);
    const payload = JSON.parse(r.stdout);
    expect(payload.ok).toBe(true);
    expect(payload.revertedToCommit).toBe(target.oid);

    // Byte-level restore: whole file back to the recorded state-1 bytes,
    // and the restored document.xml byte-equals the commit's stored part.
    const reverted = fs.readFileSync(file);
    expect(reverted.equals(state1Bytes)).toBe(true);
    const docPart = new AdmZip(reverted).getEntry('word/document.xml')?.getData();
    expect(docPart).toBeDefined();
    expect(docPart!.equals(Buffer.from(targetPkg.xml, 'utf8'))).toBe(true);

    const after = await getHistory(file);
    expect(after.length).toBe(4);
    expect(after[0].message).toMatch(/^Revert to [0-9a-f]{7}: /);
    expect(after[0].message).toContain('checkpoint one');
  });

  it('redo re-applies the state the revert replaced, then refuses a second redo', async () => {
    const { file, state2Bytes } = await setupTwoSaves('astra-hist-redo');
    const commits = await getHistory(file);
    const preRevertOid = commits[0].oid;

    const rev = await runCli(['--json', 'revert', file, '--steps', '1']);
    expect(rev.status).toBe(0);
    expect(fs.readFileSync(file).equals(state2Bytes)).toBe(false);

    const r = await runCli(['--json', 'redo', file]);
    expect(r.status).toBe(0);
    const payload = JSON.parse(r.stdout);
    expect(payload.ok).toBe(true);
    expect(payload.restoredCommit).toBe(preRevertOid);
    expect(fs.readFileSync(file).equals(state2Bytes)).toBe(true);
    const after = await getHistory(file);
    expect(after[0].message).toMatch(/^Redo: /);

    const again = await runCli(['redo', file]);
    expect(again.status).toBe(1);
    expect(again.stderr).toContain('[E_NOT_FOUND]');
    expect(again.stderr).toContain('nothing to redo');
    expect(fs.readFileSync(file).equals(state2Bytes)).toBe(true);
    expect((await getHistory(file)).length).toBe(5);
  });

  it('MCP revert followed by CLI redo restores the pre-revert state (cross-surface)', async () => {
    const { file, store, sessionId, state2Bytes } = await setupTwoSaves('astra-hist-redo-cross');
    const rev = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, steps: 1 });
    expect(rev.ok).toBe(true);
    if (!rev.ok) return;

    const r = await runCli(['--json', 'redo', file]);
    expect(r.status).toBe(0);
    expect(JSON.parse(r.stdout).ok).toBe(true);
    expect(fs.readFileSync(file).equals(state2Bytes)).toBe(true);
  });
});

describe('preview-revert is read-only', () => {
  it('names the target commit and affected files, writes nothing', async () => {
    const { file } = await setupTwoSaves('astra-hist-preview');
    const commits = await getHistory(file);
    const before = fs.readFileSync(file);

    const r = await runCli(['--json', 'preview-revert', file, '--steps', '1']);
    expect(r.status).toBe(0);
    const payload = JSON.parse(r.stdout);
    expect(payload.ok).toBe(true);
    expect(payload.command).toBe('preview-revert');
    expect(payload.targetOid).toBe(commits[1].oid);
    expect(payload.currentHeadOid).toBe(commits[0].oid);
    expect(payload.totalCommits).toBe(3);
    expect(payload.affectedFiles).toContain('word/document.xml');
    expect(payload.diverged).toBe(false);
    expect(payload.divergedFiles).toEqual([]);

    expect(fs.readFileSync(file).equals(before)).toBe(true);
    const after = await getHistory(file);
    expect(after.length).toBe(3);
    expect(after[0].oid).toBe(commits[0].oid);
  });
});

describe('sidecar-absent refusal and explicit opt-in', () => {
  it('history/preview-revert/revert/redo refuse with the --init remedy and create nothing', async () => {
    const { dir, file } = writeBlank('astra-hist-absent');
    for (const cmd of ['history', 'preview-revert', 'revert', 'redo']) {
      const r = await runCli([cmd, file]);
      expect(r.status).toBe(2);
      expect(r.stderr).toMatch(/usage error/i);
      expect(r.stderr).toContain('--init');
      expect(fs.existsSync(path.join(dir, '.astra'))).toBe(false);
    }
    const json = await runCli(['--json', 'revert', file]);
    expect(json.status).toBe(2);
    const payload = JSON.parse(json.stdout);
    expect(payload.ok).toBe(false);
    expect(payload.code).toBe('E_USAGE');
    expect(payload.remedy).toContain('--init');
    expect(json.stderr).toBe('');
    expect(fs.existsSync(path.join(dir, '.astra'))).toBe(false);
  });

  it('history --init / --track create the baseline explicitly and stay idempotent', async () => {
    const { dir, file } = writeBlank('astra-hist-init');
    const init = await runCli(['--json', 'history', file, '--init']);
    expect(init.status).toBe(0);
    const payload = JSON.parse(init.stdout);
    expect(payload.ok).toBe(true);
    expect(payload.initialized).toBe(true);
    expect(payload.totalCommits).toBe(1);
    expect(payload.commitOid).toMatch(/^[0-9a-f]{40}$/);
    expect(fs.existsSync(path.join(dir, '.astra', 'doc.docx', '.git'))).toBe(true);

    const again = await runCli(['--json', 'history', file, '--init']);
    expect(again.status).toBe(0);
    const againPayload = JSON.parse(again.stdout);
    expect(againPayload.initialized).toBe(false);
    expect(againPayload.totalCommits).toBe(1);

    const track = await runCli(['--json', 'history', file, '--track']);
    expect(track.status).toBe(0);
    expect(JSON.parse(track.stdout).totalCommits).toBe(1);
  });
});

describe('diverged working file fails closed', () => {
  it('revert refuses E_CONFLICT and writes nothing; preview reports the divergence', async () => {
    const { file } = await setupTwoSaves('astra-hist-diverged');
    const parser = new DocxParser(fs.readFileSync(file));
    parser.insertRootNode(
      '<w:p><w:r><w:t xml:space="preserve">Unrecorded local edit</w:t></w:r></w:p>',
      undefined,
      'append',
    );
    fs.writeFileSync(file, parser.rezip());
    const before = fs.readFileSync(file);
    const commitsBefore = await getHistory(file);

    const r = await runCli(['revert', file]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('[E_CONFLICT]');
    // Remedy names both recording paths: MCP save and the standard-git sidecar.
    expect(r.stderr).toContain('save_document');
    expect(r.stderr).toContain('standard git repo');
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    const commitsAfter = await getHistory(file);
    expect(commitsAfter.length).toBe(commitsBefore.length);
    expect(commitsAfter[0].oid).toBe(commitsBefore[0].oid);

    const preview = await runCli(['--json', 'preview-revert', file]);
    expect(preview.status).toBe(0);
    const payload = JSON.parse(preview.stdout);
    expect(payload.diverged).toBe(true);
    expect(payload.divergedFiles).toContain('word/document.xml');
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
});

describe('ordinary batch never touches .astra/', () => {
  it('batch --append leaves no sidecar behind and does not create .astra/', async () => {
    const { dir, file } = writeBlank('astra-hist-batch');
    const r = await runCli(['batch', file, '--append', 'Batch probe']);
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('BATCH OK');
    expect(fs.existsSync(path.join(dir, '.astra'))).toBe(false);
  });
});

describe('commit failure rolls the working file back', () => {
  it('a sidecar commit write failure restores the pre-command bytes and records nothing', async () => {
    const { dir, file } = await setupTwoSaves('astra-hist-commit-fail');
    const before = fs.readFileSync(file);
    const commitsBefore = await getHistory(file);

    // Nearest reachable commit failure: plant a symlink at the sidecar's
    // working-tree document.xml so the commit path's write guard refuses
    // (E_SANDBOX) after the restore write already landed. The rollback must
    // put the pre-command bytes back and append no commit.
    const sidecarDocXml = path.join(dir, '.astra', 'doc.docx', 'document.xml');
    fs.rmSync(sidecarDocXml);
    fs.symlinkSync(path.join(dir, 'missing-sidecar-target'), sidecarDocXml);

    const r = await runCli(['revert', file, '--steps', '1']);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('[E_SANDBOX]');
    expect(r.stderr).toContain('history commit failed');
    expect(r.stderr).toContain('restored to the pre-command bytes');
    expect(fs.readFileSync(file).equals(before)).toBe(true);

    const commitsAfter = await getHistory(file);
    expect(commitsAfter.length).toBe(commitsBefore.length);
    expect(commitsAfter[0].oid).toBe(commitsBefore[0].oid);
  });
});

describe('vanished sidecar between read and commit', () => {
  it('deleting .astra mid-command refuses E_NOT_FOUND and restores the pre-command bytes', async () => {
    const { dir, file } = await setupTwoSaves('astra-hist-vanish');
    const before = fs.readFileSync(file);
    const repoRoot = path.join(dir, '.astra');

    // Arm the write seam: the sidecar disappears right after the restore
    // write, before the commit path re-checks it.
    vanishHook.sidecarRoot = repoRoot;
    vanishHook.armed = true;
    const r = await runCli(['revert', file, '--steps', '1']);
    vanishHook.armed = false;
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('[E_NOT_FOUND]');
    expect(r.stderr).toContain('sidecar disappeared');
    expect(r.stderr).toContain('--init');
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    expect(fs.existsSync(repoRoot)).toBe(false);
  });
});

describe('killed-mid-init recovery, init refusals, and sidecar entry semantics', () => {
  it('history --init repairs a 0-commit (unborn) sidecar; the read names the repair', async () => {
    const { dir, file } = writeBlank('astra-hist-unborn');
    // Kill-mid-init shape: the sidecar's `git init` landed but the baseline
    // commit never did (unborn HEAD, git.log unreadable).
    const sidecarDir = path.join(dir, '.astra', 'doc.docx');
    fs.mkdirSync(sidecarDir, { recursive: true });
    await git.init({ fs, dir: sidecarDir });
    expect(await getHistoryCommitCount(sidecarDir)).toBe(0);

    const read = await runCli(['history', file]);
    expect(read.status).toBe(1);
    expect(read.stderr).toContain('[E_IO]');
    expect(read.stderr).toContain('no commits reachable from HEAD');
    expect(read.stderr).toContain('--init');

    const init = await runCli(['--json', 'history', file, '--init']);
    expect(init.status).toBe(0);
    const payload = JSON.parse(init.stdout);
    expect(payload.initialized).toBe(true);
    expect(payload.totalCommits).toBe(1);
    expect(payload.commitOid).toMatch(/^[0-9a-f]{40}$/);

    const after = await runCli(['history', file]);
    expect(after.status).toBe(0);
    expect(after.stdout).toContain('1 commit');
    expect((await getHistoryCommitCount(sidecarDir))).toBe(1);
  });

  it.skipIf(typeof process.getuid === 'function' && process.getuid() === 0)(
    'a read-only (0555) document directory: history --init fails E_IO with a writable-dir remedy',
    async () => {
      const { dir, file } = writeBlank('astra-hist-eacces');
      fs.chmodSync(dir, 0o555);
      try {
        const r = await runCli(['history', file, '--init']);
        expect(r.status).toBe(1);
        expect(r.stderr).toContain('[E_IO]');
        expect(r.stderr).toContain('writable');
        expect(r.stderr).not.toContain('re-inspect');
        expect(fs.existsSync(path.join(dir, '.astra'))).toBe(false);
      } finally {
        fs.chmodSync(dir, 0o755);
      }
    },
  );

  it('.git-as-file sidecar refuses reads and --init (stray entry) and writes nothing', async () => {
    const { dir, file } = writeBlank('astra-hist-stray');
    const sidecarDir = path.join(dir, '.astra', 'doc.docx');
    fs.mkdirSync(sidecarDir, { recursive: true });
    fs.writeFileSync(path.join(sidecarDir, '.git'), 'gitdir: ../elsewhere\n');

    // Reads: an operational refusal (exit 1), not the "no sidecar" usage gate.
    for (const cmd of ['history', 'preview-revert', 'revert', 'redo']) {
      const r = await runCli([cmd, file]);
      expect(r.status).toBe(1);
      expect(r.stderr).toContain('[E_IO]');
      expect(r.stderr).toContain('stray');
    }
    // --init must refuse too: never follow the stray gitdir, never write a
    // baseline into the sidecar root.
    const init = await runCli(['history', file, '--init']);
    expect(init.status).toBe(1);
    expect(init.stderr).toContain('[E_IO]');
    expect(init.stderr).toContain('stray');
    expect(init.stderr).toContain('--init');

    expect(fs.readdirSync(sidecarDir).sort()).toEqual(['.git']);
    expect(fs.readFileSync(path.join(sidecarDir, '.git'), 'utf8')).toBe('gitdir: ../elsewhere\n');
    expect(fs.existsSync(path.join(dir, '.astra', 'elsewhere'))).toBe(false);
  });

  it('a broken .git symlink is refused as a stray entry, not treated as absent', async () => {
    const { dir, file } = writeBlank('astra-hist-stray-link');
    const sidecarDir = path.join(dir, '.astra', 'doc.docx');
    fs.mkdirSync(sidecarDir, { recursive: true });
    fs.symlinkSync(path.join(sidecarDir, 'missing-target'), path.join(sidecarDir, '.git'));

    const init = await runCli(['history', file, '--init']);
    expect(init.status).toBe(1);
    expect(init.stderr).toContain('[E_IO]');
    expect(init.stderr).toContain('stray');
    expect(fs.readdirSync(sidecarDir).sort()).toEqual(['.git']);
  });

  it('a document.xml symlink planted in the sidecar: --init keeps the engine E_SANDBOX code and remedy (H2 pin)', async () => {
    const { dir, file } = writeBlank('astra-hist-docxml-link');
    const sidecarDir = path.join(dir, '.astra', 'doc.docx');
    fs.mkdirSync(sidecarDir, { recursive: true });
    const planted = path.join(sidecarDir, 'document.xml');
    const outside = path.join(dir, 'escaped-document.xml');
    fs.symlinkSync(outside, planted);

    // Run 1 exercises the absent-.git create path; run 2 the 0-commit repair
    // path (run 1's refusal leaves an unborn `.git` behind). Both funnel into
    // initHistorySidecar, and both must surface the engine refusal verbatim:
    // [E_SANDBOX] + the sandbox remedy, never the wrap's generic
    // writable-directory E_IO fallback.
    for (const shape of ['create', 'repair']) {
      const r = await runCli(['history', file, '--init']);
      expect(r.status, shape).toBe(1);
      expect(r.stderr, shape).toContain('[E_SANDBOX]');
      expect(r.stderr, shape).toContain('cannot record the baseline');
      expect(r.stderr, shape).toContain('Remove the symlink');
      expect(r.stderr, shape).not.toContain('writable');
    }

    // The write guard refused before anything was written through the link;
    // the sidecar stays as planted, with no baseline commit.
    expect(fs.existsSync(outside)).toBe(false);
    expect(fs.lstatSync(planted).isSymbolicLink()).toBe(true);
    expect(await getHistoryCommitCount(sidecarDir)).toBe(0);
  });

  it('sidecar .git symlinked to a real gitdir: honest empty-chain read, --init records through the link (H3 pin)', async () => {
    const { dir, file } = writeBlank('astra-hist-gitlink');
    // Followed-symlink semantics pinned here: the CLI classifier and
    // isomorphic-git (discoverGitdir, via stat — not lstat) both follow a
    // `.git` symlink to a real directory, so the sidecar is a usable — if
    // unborn — repo whose chain physically lives in the linked gitdir. A
    // link to a file or a broken link stays a stray refusal (covered above);
    // only the directory case is followed.
    const gitHost = path.join(dir, 'git-host');
    fs.mkdirSync(gitHost);
    await git.init({ fs, dir: gitHost });
    const sidecarDir = path.join(dir, '.astra', 'doc.docx');
    fs.mkdirSync(sidecarDir, { recursive: true });
    fs.symlinkSync(path.join(gitHost, '.git'), path.join(sidecarDir, '.git'));

    // Unborn HEAD reads as the honest no-commits-reachable state — not as a
    // missing sidecar and not as a generic readability problem.
    const read = await runCli(['history', file]);
    expect(read.status).toBe(1);
    expect(read.stderr).toContain('[E_IO]');
    expect(read.stderr).toContain('no commits reachable from HEAD');

    const init = await runCli(['--json', 'history', file, '--init']);
    expect(init.status).toBe(0);
    const payload = JSON.parse(init.stdout);
    expect(payload.initialized).toBe(true);
    expect(payload.totalCommits).toBe(1);

    // The baseline commit landed in the linked gitdir (through the symlink),
    // and a re-read through the sidecar shows it.
    expect(await getHistoryCommitCount(gitHost)).toBe(1);
    const after = await runCli(['history', file]);
    expect(after.status).toBe(0);
    expect(after.stdout).toContain('1 commit');
    expect(after.stdout).toContain(String(payload.commitOid).slice(0, 7));
    expect(fs.lstatSync(path.join(sidecarDir, '.git')).isSymbolicLink()).toBe(true);
  });
});

describe('revert preserves untracked package parts', () => {
  it('docProps bytes edited after the last commit survive the revert byte-identically', async () => {
    const { file } = writeBlank('astra-hist-untracked');
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    expect(opened.ok).toBe(true);
    if (!opened.ok) throw new Error('open failed in test setup');
    const sessionId = opened.sessionId;
    // Synthesize docProps/core.xml + its [Content_Types].xml override before
    // any recorded commit: the override is tracked, core.xml never is.
    const meta = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'metadata', operation: 'setProperties', title: 'Untracked probe' }],
    });
    expect(meta.ok).toBe(true);
    await replaceAnchor(store, sessionId, 'Alpha state one');
    const save1 = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'checkpoint one' });
    expect(save1.ok).toBe(true);
    await replaceAnchor(store, sessionId, 'Beta state two');
    const save2 = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'checkpoint two' });
    expect(save2.ok).toBe(true);

    const commits = await getHistory(file);
    const target = await checkoutCommit(file, commits[1].oid);

    // Untracked-part divergence is not a conflict: docProps/ is outside the
    // tracked set, so the revert proceeds and must keep these live bytes.
    const zip = new AdmZip(fs.readFileSync(file));
    const core = zip.getEntry('docProps/core.xml');
    expect(core).not.toBeNull();
    const originalBytes = core!.getData();
    const modifiedBytes = Buffer.concat([
      originalBytes,
      Buffer.from('\n<!-- untracked-part-preservation-probe -->\n'),
    ]);
    expect(modifiedBytes.equals(originalBytes)).toBe(false);
    zip.updateFile('docProps/core.xml', modifiedBytes);
    fs.writeFileSync(file, zip.toBuffer());

    const r = await runCli(['revert', file, '--steps', '1']);
    expect(r.status).toBe(0);

    const rezipped = new AdmZip(fs.readFileSync(file));
    const afterBytes = rezipped.getEntry('docProps/core.xml')?.getData();
    expect(afterBytes).toBeDefined();
    expect(afterBytes!.equals(modifiedBytes)).toBe(true);
    const docPart = rezipped.getEntry('word/document.xml')?.getData();
    expect(docPart).toBeDefined();
    expect(docPart!.equals(Buffer.from(target.xml, 'utf8'))).toBe(true);
  });
});

describe('tombstone warning after external sidecar removal', () => {
  it('rm -rf .astra then history --init re-records a baseline and surfaces the warning', async () => {
    const { dir, file } = writeBlank('astra-hist-tombstone');
    const first = await runCli(['--json', 'history', file, '--init']);
    expect(first.status).toBe(0);
    const firstPayload = JSON.parse(first.stdout);
    expect(firstPayload.initialized).toBe(true);
    expect(firstPayload.warning).toBeNull();
    expect(fs.existsSync(path.join(dir, '.astra-tombstones.json'))).toBe(true);

    fs.rmSync(path.join(dir, '.astra'), { recursive: true, force: true });

    const again = await runCli(['--json', 'history', file, '--init']);
    expect(again.status).toBe(0);
    const payload = JSON.parse(again.stdout);
    expect(payload.initialized).toBe(true);
    expect(payload.totalCommits).toBe(1);
    expect(typeof payload.warning).toBe('string');
    expect(payload.warning).toContain('possible history removal');
    expect(payload.warning).toContain('doc.docx');
    expect(again.stderr).toContain('warning:');
    expect(again.stderr).toContain('possible history removal');
  });
});

describe('usage and depth gating', () => {
  it('bare commands exit 2 and --help exits 0 with per-command usage', async () => {
    for (const cmd of ['history', 'preview-revert', 'revert', 'redo']) {
      const bare = await runCli([cmd]);
      expect(bare.status).toBe(2);
      expect(bare.stderr).toMatch(/usage error/i);

      const help = await runCli([cmd, '--help']);
      expect(help.status).toBe(0);
      expect(help.stdout).toContain(`Usage: astra ${cmd}`);
    }

    // revert --help tracked-set list matches isTrackedPackageFile truth
    // (previously omitted [Content_Types].xml and the media/charts/embeddings/theme dirs).
    const revHelp = await runCli(['revert', '--help']);
    expect(revHelp.status).toBe(0);
    expect(revHelp.stdout).toContain('[Content_Types].xml');
    expect(revHelp.stdout).toContain('word/{media,charts,embeddings,theme}/');
    expect(revHelp.stdout).toContain('docProps/');
  });

  it('bad flags and steps exit 2; steps past the history depth exits 1 with nothing written', async () => {
    const { file } = writeBlank('astra-hist-steps');
    for (const args of [
      ['revert', file, '--steps', '0'],
      ['revert', file, '--steps', 'abc'],
      ['revert', file, '--steps'],
      ['preview-revert', file, '--steps', '-2'],
      ['history', file, '--bogus'],
      ['redo', file, 'extra.docx'],
    ] as string[][]) {
      const r = await runCli(args);
      expect(r.status).toBe(2);
      expect(r.stderr).toMatch(/usage error/i);
    }
    expect(fs.existsSync(path.join(path.dirname(file), '.astra'))).toBe(false);

    const { file: tracked } = await setupTwoSaves('astra-hist-depth');
    const before = fs.readFileSync(tracked);
    const r = await runCli(['revert', tracked, '--steps', '9']);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('[E_NOT_FOUND]');
    expect(fs.readFileSync(tracked).equals(before)).toBe(true);
  });

  it('git-ref-style positionals name the depth contract (--steps, history, MCP-only oid targeting)', async () => {
    const { file } = writeBlank('astra-hist-gitref');
    for (const ref of ['HEAD', 'HEAD~1', 'HEAD@{1}', '@{1}', '~', 'a'.repeat(40)]) {
      const res = await runCli(['revert', file, ref]);
      expect(res.status).toBe(2);
      expect(res.stderr).toContain('--steps');
      expect(res.stderr).toContain('astra history');
      expect(res.stderr).toContain('MCP');
      expect(res.stderr).not.toContain('takes exactly one');
    }
    // preview-revert shares the parser and gets the same remedy.
    const preview = await runCli(['preview-revert', file, 'HEAD~1']);
    expect(preview.status).toBe(2);
    expect(preview.stderr).toContain('--steps');
    // Each command names the MCP tool that really takes a commit oid.
    expect(preview.stderr).toContain('preview_revert');
    expect(preview.stderr).not.toContain('revert_document');
    const reverted = await runCli(['revert', file, 'HEAD@{1}']);
    expect(reverted.stderr).toContain('revert_document');
    // redo only moves forward by depth: no commit-oid claim.
    const redone = await runCli(['redo', file, 'HEAD~1']);
    expect(redone.status).toBe(2);
    expect(redone.stderr).toContain('--steps');
    expect(redone.stderr).not.toContain('MCP');
    // A non-git second positional keeps the generic positional error.
    const generic = await runCli(['revert', file, 'other.docx']);
    expect(generic.status).toBe(2);
    expect(generic.stderr).toContain('takes exactly one <docx> argument');
    // --json carries the remedy in the unified usage envelope.
    const json = await runCli(['--json', 'revert', file, 'HEAD~1']);
    expect(json.status).toBe(2);
    const payload = JSON.parse(json.stdout);
    expect(payload.code).toBe('E_USAGE');
    expect(payload.remedy).toContain('--steps');
    expect(payload.remedy).toContain('astra history');
    expect(payload.remedy).toContain('MCP');
    // Refusals create no sidecar and write nothing.
    expect(fs.existsSync(path.join(path.dirname(file), '.astra'))).toBe(false);
  });
});
