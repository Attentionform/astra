import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import git from 'isomorphic-git';
import os from 'node:os';
import path from 'node:path';
import {
  captureOriginal, commitChange, exportOriginal, getOriginalInfo, getRepoDir,
  initAstraRepo, originalSha256, readOriginal, replaceFileAtomic, writeFileExclusive, checkoutCommit, getHistory,
  assertGitHistoryReady, assertHistoryReadyForWrite,
  createVerifiedBackup,
} from '@attentionform/astra-engine';

let dir: string;
let doc: string;
beforeEach(async () => { dir = await fsp.realpath(await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-original-'))); doc = path.join(dir, 'test.docx'); });
afterEach(async () => { vi.restoreAllMocks(); vi.unstubAllEnvs(); await fsp.rm(dir, { recursive: true, force: true }); });
const data = Buffer.from('PK original bytes: sensitive reviewer');
describe('exact original capture and export', () => {
  it('refuses exporting over the working file or into its history store', async () => {
    await fsp.writeFile(doc, data);
    await initAstraRepo(doc, '<a/>', {}, data);
    await expect(exportOriginal(doc, doc, { overwrite: true })).rejects.toMatchObject({ code: 'E_CONFLICT' });
    const metadata = path.join(getRepoDir(doc), '.git', 'astra-original', 'original.snapshot');
    await expect(exportOriginal(doc, metadata, { overwrite: true })).rejects.toMatchObject({ code: 'E_CONFLICT' });
    expect(await readOriginal(doc)).toEqual(data);
  });
  it('exports with ordinary permissions while the retained original stays private', async () => {
    if (process.platform === 'win32') return;
    await fsp.writeFile(doc, data);
    await initAstraRepo(doc, '<a/>', {}, data);
    const previous = process.umask(0o022);
    try {
      const fresh = path.join(dir, 'fresh.docx');
      await exportOriginal(doc, fresh);
      expect((await fsp.stat(fresh)).mode & 0o777).toBe(0o644);
      const viaReplace = path.join(dir, 'replace.docx');
      await exportOriginal(doc, viaReplace, { overwrite: true });
      expect((await fsp.stat(viaReplace)).mode & 0o777).toBe(0o644);
    } finally { process.umask(previous); }
    const stored = path.join(getRepoDir(doc), '.git', 'astra-original', 'original.snapshot');
    expect((await fsp.stat(stored)).mode & 0o077).toBe(0);
  });
  it('reports an unreadable history as unreadable, not as an interrupted initialization', async () => {
    await fsp.writeFile(doc, data);
    await initAstraRepo(doc, '<a/>', {}, data);
    await expect(assertHistoryReadyForWrite(doc)).resolves.toBeUndefined();
    const refs = path.join(getRepoDir(doc), '.git', 'refs', 'heads');
    const branch = (await fsp.readdir(refs))[0];
    await fsp.writeFile(path.join(refs, branch), 'not-a-valid-object-id\n');
    const error = await assertHistoryReadyForWrite(doc).then(() => undefined, (e: unknown) => e as Error);
    expect(error).toBeInstanceOf(Error);
    expect(error!.message).not.toMatch(/initialization is incomplete/i);
  });
  it('records exact input, retains it across saves and keeps it out of package parts', async () => {
    await initAstraRepo(doc, '<a/>', {}, data);
    const info = await getOriginalInfo(doc);
    expect(info).toMatchObject({ available: true, sha256: originalSha256(data), bytes: data.length });
    for (let i = 0; i < 5; i++) await commitChange(doc, `<a>${i}</a>`, 'edit', {});
    expect(await readOriginal(doc)).toEqual(data);
    const target = await checkoutCommit(doc, (await getHistory(doc))[0].oid);
    expect([...target.extra.keys()].some((p) => p.includes('astra-original'))).toBe(false);
    const out = path.join(dir, 'out.docx');
    await exportOriginal(doc, out);
    expect(await fsp.readFile(out)).toEqual(data);
    await expect(exportOriginal(doc, out)).rejects.toMatchObject({ code: 'E_CONFLICT' });
    await exportOriginal(doc, out, { overwrite: true });
    expect(originalSha256(await fsp.readFile(out))).toBe(originalSha256(data));
  });
  it('never replaces an original when later input differs', async () => {
    await initAstraRepo(doc, '<a/>', {}, data);
    expect(await captureOriginal(doc, Buffer.from('later'))).toMatchObject({ available: false, reason: 'capture-failed' });
    expect(await readOriginal(doc)).toEqual(data);
  });
  it('five concurrent captures elect one complete original without deleting shared bytes', async () => {
    await initAstraRepo(doc, '<a/>', {});
    const inputs = [data, Buffer.from('b'), data, Buffer.from('c'), Buffer.from('b')];
    const infos = await Promise.all(inputs.map((bytes) => captureOriginal(doc, bytes)));
    const result = await readOriginal(doc);
    expect(inputs.some((v) => v.equals(result))).toBe(true);
    expect(await getOriginalInfo(doc)).toMatchObject({ available: true, sha256: originalSha256(result) });
    for (let i = 0; i < inputs.length; i++) expect(infos[i].available).toBe(inputs[i].equals(result));
    const originalDir = path.join(getRepoDir(doc), '.git', 'astra-original');
    expect(await fsp.readdir(originalDir)).toContain('original.snapshot');
    expect((await fsp.readdir(originalDir)).some((name) => /^original-[a-f0-9]{64}\.docx$/.test(name))).toBe(false);
  });
  it('five same-byte captures all verify one original', async () => {
    await initAstraRepo(doc, '<a/>', {});
    expect((await Promise.all(Array.from({ length: 5 }, () => captureOriginal(doc, data)))).every((x) => x.available)).toBe(true);
  });
  it('reports the committed original when temp cleanup fails after atomic publication', async () => {
    await initAstraRepo(doc, '<a/>', {});
    const unlink = fsp.unlink.bind(fsp);
    vi.spyOn(fsp, 'unlink').mockImplementation(async (file) => {
      if (String(file).includes('.original.snapshot.astra-tmp-')) throw Object.assign(new Error('injected temp cleanup failure'), { code: 'EACCES' });
      return unlink(file);
    });
    const captured = await captureOriginal(doc, data);
    expect(captured).toMatchObject({ available: true, sha256: originalSha256(data) });
    expect(await readOriginal(doc)).toEqual(data);
    expect((await fsp.readdir(path.join(getRepoDir(doc), '.git/astra-original'))).some((name) => name.includes('.astra-tmp-'))).toBe(true);
  });
  it('serializes competing first history initializers and keeps document state matched to its original', async () => {
    const inputs = [Buffer.from('first complete input'), Buffer.from('second complete input')];
    const xml = ['<a><first/></a>', '<a><second/></a>'];
    const attempts = await Promise.allSettled(inputs.map((bytes, index) => initAstraRepo(doc, xml[index], {}, bytes)));
    const winner = attempts.findIndex((attempt) => attempt.status === 'fulfilled');
    expect(attempts.filter((attempt) => attempt.status === 'fulfilled')).toHaveLength(1);
    expect(attempts.filter((attempt) => attempt.status === 'rejected')).toHaveLength(1);
    const history = await getHistory(doc);
    expect(history).toHaveLength(1);
    const head = await checkoutCommit(doc, history[0].oid);
    expect(head.xml).toBe(xml[winner]);
    expect(await readOriginal(doc)).toEqual(inputs[winner]);
    const originalInfo = await getOriginalInfo(doc);
    expect(originalInfo.available).toBe(true);
    if (!originalInfo.available) throw new Error(`Original unavailable: ${originalInfo.reason}`);
    expect(originalInfo.sha256).toBe(originalSha256(inputs[winner]));
    await expect(fsp.stat(path.join(getRepoDir(doc), '.astra-init.lock'))).rejects.toMatchObject({ code: 'ENOENT' });
  });
  it('elects one correct initializer across independent processes (2 and 5 writers)', async () => {
    const worker = path.resolve(__dirname, '../fixtures/original-init-worker.cjs');
    for (const count of [2, 5]) {
      const target = path.join(dir, `process-race-${count}.docx`);
      const barrier = path.join(dir, `barrier-${count}`);
      await fsp.mkdir(barrier);
      const inputs = Array.from({ length: count }, (_, index) => Buffer.from(`writer-${count}-${index}-original-bytes`));
      const children = inputs.map((bytes, index) => {
        const child = spawn(process.execPath, [worker, target, `<a><writer>${index}</writer></a>`, bytes.toString('base64'), String(index), barrier], { cwd: process.cwd() });
        let stdout = '';
        let stderr = '';
        child.stdout.setEncoding('utf8').on('data', (chunk: string) => { stdout += chunk; });
        child.stderr.setEncoding('utf8').on('data', (chunk: string) => { stderr += chunk; });
        const result = new Promise<{ code: number | null; stdout: string; stderr: string }>((resolve, reject) => {
          child.once('error', reject);
          child.once('close', (code) => resolve({ code, stdout, stderr }));
        });
        return { child, result };
      });
      try {
        for (let index = 0; index < count; index++) {
          const ready = path.join(barrier, `${index}.ready`);
          const deadline = Date.now() + 15_000;
          while (Date.now() < deadline) {
            if (await fsp.stat(ready).then(() => true, () => false)) break;
            await new Promise((resolve) => setTimeout(resolve, 5));
          }
          expect(await fsp.stat(ready).then(() => true, () => false), `writer ${index} did not reach the start barrier`).toBe(true);
        }
        await fsp.writeFile(path.join(barrier, 'go'), 'start');
        const results = await Promise.all(children.map((child) => child.result));
        const records = results.map((result) => JSON.parse(result.stdout) as { ok: boolean; id: string; sha256?: string; code?: string });
        expect(records.filter((record) => record.ok)).toHaveLength(1);
        expect(results.filter((result) => result.code === 0)).toHaveLength(1);
        expect(results.filter((result) => result.code === 1)).toHaveLength(count - 1);
        expect(records.filter((record) => !record.ok).every((record) => record.code === 'E_CONFLICT')).toBe(true);
        const winner = Number(records.find((record) => record.ok)!.id);
        expect(await readOriginal(target)).toEqual(inputs[winner]);
        const head = await checkoutCommit(target, (await getHistory(target))[0].oid);
        expect(head.xml).toBe(`<a><writer>${winner}</writer></a>`);
        expect(await fsp.stat(path.join(getRepoDir(target), '.astra-init.lock')).then(() => true, () => false)).toBe(false);
      } finally {
        await fsp.writeFile(path.join(barrier, 'go'), 'release').catch(() => undefined);
        for (const child of children) if (child.child.exitCode === null) child.child.kill();
      }
    }
  });
  it('survives a killed capture process before snapshot publication and retries after the leftover is removed by hand', async () => {
    await initAstraRepo(doc, '<a/>', {});
    const worker = path.resolve(__dirname, '../fixtures/original-capture-hold-worker.cjs');
    const readyFile = path.join(dir, 'capture-ready');
    const child = spawn(process.execPath, [worker, doc, data.toString('base64'), readyFile], { cwd: process.cwd() });
    const closed = new Promise<{ code: number | null; signal: NodeJS.Signals | null }>((resolve, reject) => {
      child.once('error', reject);
      child.once('close', (code, signal) => resolve({ code, signal }));
    });
    try {
      const deadline = Date.now() + 15_000;
      while (Date.now() < deadline && !await fsp.stat(readyFile).then(() => true, () => false)) await new Promise((resolve) => setTimeout(resolve, 5));
      expect(await fsp.stat(readyFile).then(() => true, () => false), 'capture worker did not reach the pre-publication barrier').toBe(true);
      const od = path.join(getRepoDir(doc), '.git/astra-original');
      const temps = async (): Promise<string[]> => (await fsp.readdir(od)).filter((name) => name.includes('.astra-tmp-'));
      expect(await temps()).toHaveLength(1);
      expect((await getOriginalInfo(doc)).available).toBe(false);
      child.kill('SIGKILL');
      expect(await closed).toMatchObject({ code: null, signal: 'SIGKILL' });
      const leftover = await temps();
      expect(leftover).toHaveLength(1);
      await fsp.unlink(path.join(od, leftover[0]));
      expect((await captureOriginal(doc, data)).available).toBe(true);
      expect(await readOriginal(doc)).toEqual(data);
    } finally {
      if (child.exitCode === null) child.kill('SIGKILL');
      await closed.catch(() => undefined);
    }
  });
  it('fails closed on an interrupted initializer lock and permits explicit recovery after it is removed', async () => {
    const repoDir = path.join(getRepoDir(doc));
    await fsp.mkdir(repoDir, { recursive: true });
    const lock = path.join(repoDir, '.astra-init.lock');
    await fsp.writeFile(lock, JSON.stringify({ version: 1, pid: 999999, token: 'interrupted' }));
    await expect(initAstraRepo(doc, '<a/>', {}, data)).rejects.toMatchObject({ code: 'E_CONFLICT' });
    await expect(assertGitHistoryReady(repoDir)).rejects.toMatchObject({ code: 'E_CONFLICT' });
    expect(await fsp.stat(lock)).toBeTruthy();
    expect(await fsp.stat(path.join(repoDir, '.git')).then(() => true, () => false)).toBe(false);
    await fsp.unlink(lock);
    await initAstraRepo(doc, '<a/>', {}, data);
    expect(await readOriginal(doc)).toEqual(data);
  });
  it('does not commit package parts staged by a failed initialization when the zero-commit history is retried', async () => {
    await expect(initAstraRepo(doc, '<a><partial/></a>', {
      'word/styles.xml': Buffer.from('<styles>from failed attempt</styles>'),
      '../outside.xml': Buffer.from('must fail'),
    }, data)).rejects.toThrow('Security Violation');

    await initAstraRepo(doc, '<a><recovered/></a>', {}, data);
    const history = await getHistory(doc);
    expect(history).toHaveLength(1);
    const recovered = await checkoutCommit(doc, history[0].oid);
    expect(recovered.xml).toBe('<a><recovered/></a>');
    expect(recovered.extra.has('word/styles.xml')).toBe(false);
  });
  it('does not destructively repair a history whose existing commit log is unreadable', async () => {
    const style = Buffer.from('<styles>keep existing data</styles>');
    await initAstraRepo(doc, '<a/>', { 'word/styles.xml': style }, data);
    const repoDir = getRepoDir(doc);
    const [head] = await getHistory(doc);
    const commitObject = path.join(repoDir, '.git', 'objects', head.oid.slice(0, 2), head.oid.slice(2));
    await fsp.unlink(commitObject);

    await expect(initAstraRepo(doc, '<b/>', {}, data)).rejects.toThrow();
    expect(await fsp.readFile(path.join(repoDir, 'word/styles.xml'))).toEqual(style);
  });
  it('refuses to reinitialize a committed history without changing its package snapshot', async () => {
    const style = Buffer.from('<styles>committed data</styles>');
    await initAstraRepo(doc, '<a/>', { 'word/styles.xml': style }, data);
    const before = await getHistory(doc);

    await expect(initAstraRepo(doc, '<b/>', {}, data)).rejects.toMatchObject({ code: 'E_CONFLICT' });
    expect(await getHistory(doc)).toEqual(before);
    expect(await fsp.readFile(path.join(getRepoDir(doc), 'word/styles.xml'))).toEqual(style);
  });
  it('reports absence, incomplete publication and corrupt bytes without returning a substitute', async () => {
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'predates-feature' });
    await initAstraRepo(doc, '<a/>', {}, data);
    const od = path.join(getRepoDir(doc), '.git/astra-original');
    const snapshot = path.join(od, 'original.snapshot');
    const saved = await fsp.readFile(snapshot);
    await fsp.unlink(snapshot);
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'incomplete' });
    await fsp.writeFile(snapshot, Buffer.from('corrupted'));
    expect(await getOriginalInfo(doc)).toMatchObject({ available: false, reason: 'corrupt' });
    await expect(readOriginal(doc)).rejects.toMatchObject({ code: 'E_STATE' });
    await fsp.writeFile(snapshot, saved);
  });
  it('new documents have no imported original; relative memory labels are not applicable', async () => {
    await initAstraRepo(doc, '<a/>', {});
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'not-applicable' });
    expect(await getOriginalInfo('memory.docx')).toEqual({ available: false, reason: 'not-applicable' });
  });
  it('counts original bytes against quota', async () => {
    vi.stubEnv('ASTRA_HISTORY_MAX_SIDECAR_BYTES', '1');
    await initAstraRepo(doc, '<a/>', {}, data);
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'too-large' });
  });
  it('keeps capture inside a relocated history root', async () => {
    vi.stubEnv('ASTRA_HISTORY_DIR', path.join(dir, 'histories'));
    await initAstraRepo(doc, '<a/>', {}, data);
    expect(getRepoDir(doc)).toContain('histories');
    expect(await readOriginal(doc)).toEqual(data);
  });
  it('a leftover temporary file does not affect status, read or export', async () => {
    await initAstraRepo(doc, '<a/>', {}, data);
    const od = path.join(getRepoDir(doc), '.git/astra-original');
    await fsp.writeFile(path.join(od, '.original.snapshot.astra-tmp-123e4567-e89b-12d3-a456-426614174000'), 'private temp content');
    expect(await getOriginalInfo(doc)).toMatchObject({ available: true, sha256: originalSha256(data) });
    expect(await readOriginal(doc)).toEqual(data);
    const out = path.join(dir, 'exported.docx');
    await exportOriginal(doc, out);
    expect(await fsp.readFile(out)).toEqual(data);
  });
  it.skipIf(process.platform === 'win32')('refuses a symlinked original snapshot', async () => {
    await initAstraRepo(doc, '<a/>', {}, data);
    const snapshot = path.join(getRepoDir(doc), '.git/astra-original/original.snapshot');
    const outside = path.join(dir, 'outside.snapshot');
    await fsp.copyFile(snapshot, outside);
    await fsp.unlink(snapshot);
    await fsp.symlink(outside, snapshot);
    expect(await getOriginalInfo(doc)).toMatchObject({ available: false, reason: 'corrupt' });
    await expect(readOriginal(doc)).rejects.toMatchObject({ code: 'E_STATE' });
  });
});
describe('first initialization: interrupted initializer, provenance and lock diagnostics', () => {
  const lockPath = (): string => path.join(getRepoDir(doc), '.astra-init.lock');
  /** A real initializer process, killed after it published original.snapshot and before its first commit. */
  async function killAfterSnapshot(bytes: Buffer, xml: string): Promise<void> {
    const worker = path.resolve(__dirname, '../fixtures/original-init-hold-worker.cjs');
    const readyFile = path.join(dir, 'init-published');
    const child = spawn(process.execPath, [worker, doc, xml, bytes.toString('base64'), readyFile], { cwd: process.cwd() });
    let stderr = '';
    child.stderr.setEncoding('utf8').on('data', (chunk: string) => { stderr += chunk; });
    const closed = new Promise<void>((resolve) => child.once('close', () => resolve()));
    try {
      const deadline = Date.now() + 15_000;
      while (Date.now() < deadline && !await fsp.stat(readyFile).then(() => true, () => false)) await new Promise((resolve) => setTimeout(resolve, 5));
      expect(await fsp.stat(readyFile).then(() => true, () => false), `initializer did not reach the pre-commit barrier: ${stderr}`).toBe(true);
      child.kill('SIGKILL');
      await closed;
    } finally {
      if (child.exitCode === null) child.kill('SIGKILL');
    }
  }
  const earlier = Buffer.from('PK earlier document bytes');
  const later = Buffer.from('PK changed document bytes, later text');
  it('a changed-input retry after a kill between publishing the snapshot and the first commit is refused and keeps both copies', async () => {
    await killAfterSnapshot(earlier, '<a><earlier/></a>');
    expect(await fsp.stat(lockPath()).then(() => true, () => false)).toBe(true); // the killed process left its lock
    expect(await getHistory(doc).catch(() => 'unreadable')).toBe('unreadable');
    expect(await getOriginalInfo(doc)).toMatchObject({ available: true, sha256: originalSha256(earlier) });
    await fsp.unlink(lockPath()); // the operator confirmed no ASTra process is running

    const refused = await initAstraRepo(doc, '<a><later/></a>', {}, later).then(() => null, (error: unknown) => error as { code?: string; reason?: string; message: string; remedy?: string });
    expect(refused).toMatchObject({ code: 'E_CONFLICT', reason: 'original-mismatch' });
    expect(refused!.remedy).toContain('do not delete it');
    // Nothing was recorded and nothing was removed: still zero commits, still the earlier original.
    expect(await getHistory(doc).catch(() => 'unreadable')).toBe('unreadable');
    expect((await git.log({ fs, dir: getRepoDir(doc) }).catch(() => [])).length).toBe(0);
    expect(await readOriginal(doc)).toEqual(earlier);
    // The earlier bytes are still retrievable without handling the snapshot by hand.
    const out = path.join(dir, 'earlier-recovered.docx');
    await exportOriginal(doc, out);
    expect(await fsp.readFile(out)).toEqual(earlier);
    // Once the operator keeps a copy and moves the snapshot aside, initialization works again.
    const snapshot = path.join(getRepoDir(doc), '.git/astra-original/original.snapshot');
    await fsp.rename(snapshot, `${snapshot}.kept`);
    await initAstraRepo(doc, '<a><later/></a>', {}, later);
    expect(await readOriginal(doc)).toEqual(later);
    expect(await fsp.readFile(`${snapshot}.kept`)).toBeTruthy();
  });
  it('a same-bytes retry after the same kill resumes, verifies the original and records one commit', async () => {
    await killAfterSnapshot(earlier, '<a><earlier/></a>');
    await fsp.unlink(lockPath());
    await initAstraRepo(doc, '<a><earlier/></a>', {}, earlier);
    expect(await getHistory(doc)).toHaveLength(1);
    expect(await readOriginal(doc)).toEqual(earlier);
    expect(await getOriginalInfo(doc)).toMatchObject({ available: true, sha256: originalSha256(earlier) });
  });
  it('describes the lock owner as running, not found or unknown, and never says deletion is safe', async () => {
    await fsp.mkdir(getRepoDir(doc), { recursive: true });
    const write = (record: unknown): Promise<void> => fsp.writeFile(lockPath(), typeof record === 'string' ? record : JSON.stringify(record));
    const message = async (): Promise<string> => (await assertHistoryReadyForWrite(doc).then(() => ({ message: 'no error' }), (error: Error) => error)).message;
    const here = os.hostname();
    await write({ version: 2, pid: process.pid, hostname: here, token: 't', startedAt: '2026-11-13T00:00:00.000Z' });
    expect(await message()).toMatch(/that process is running/);
    const dead = spawn(process.execPath, ['-e', '0']);
    await new Promise((resolve) => dead.once('close', resolve));
    await write({ version: 2, pid: dead.pid, hostname: here, token: 't', startedAt: '2026-11-13T00:00:00.000Z' });
    expect(await message()).toMatch(/was not found \(this alone does not prove the lock is stale\)/);
    await write({ version: 2, pid: process.pid, hostname: `${here}-elsewhere`, token: 't' });
    expect(await message()).toMatch(/cannot tell whether it is still running/);
    await write({ version: 1, pid: process.pid, token: 'old record without a host' });
    expect(await message()).toMatch(/host not recorded; cannot tell/);
    await write('not json');
    expect(await message()).toMatch(/does not record its owner/);
    const error = await assertHistoryReadyForWrite(doc).then(() => undefined, (e: { code?: string; remedy?: string }) => e);
    expect(error).toMatchObject({ code: 'E_CONFLICT' });
    expect(error!.remedy).toContain(lockPath());
    expect(error!.remedy).toContain('never removes the lock itself');
    expect(await fsp.stat(lockPath())).toBeTruthy(); // diagnosing never deletes
  });
  it('the write guard passes for absent and complete histories and refuses an interrupted one', async () => {
    await expect(assertHistoryReadyForWrite(doc)).resolves.toBeUndefined(); // no history yet: the save may create it
    await initAstraRepo(doc, '<a/>', {}, data);
    await expect(assertHistoryReadyForWrite(doc)).resolves.toBeUndefined();
    const other = path.join(dir, 'zero.docx');
    await fsp.mkdir(getRepoDir(other), { recursive: true });
    await git.init({ fs, dir: getRepoDir(other) });
    await expect(assertHistoryReadyForWrite(other)).rejects.toMatchObject({ code: 'E_STATE' });
  });
});
describe('publication helpers', () => {
  it('creates a byte-verified restrictive backup and never removes a preexisting collision', async () => {
    const backup = path.join(dir, 'restore-backup.docx');
    await createVerifiedBackup(backup, data);
    expect(await fsp.readFile(backup)).toEqual(data);
    if (process.platform !== 'win32') expect((await fsp.stat(backup)).mode & 0o777).toBe(0o600);
    await fsp.writeFile(backup, 'preexisting user file');
    await expect(createVerifiedBackup(backup, data)).rejects.toMatchObject({ code: 'E_CONFLICT' });
    expect(await fsp.readFile(backup, 'utf8')).toBe('preexisting user file');
  });
  it('removes a backup whose verification fails before any restore can use it', async () => {
    const backup = path.join(dir, 'restore-backup-corrupt.docx');
    const readFile = fsp.readFile.bind(fsp);
    vi.spyOn(fsp, 'readFile').mockImplementation(async (...args) => {
      if (String(args[0]) === backup) return Buffer.from('corrupt backup bytes');
      return readFile(...args);
    });
    await expect(createVerifiedBackup(backup, data)).rejects.toMatchObject({ code: 'E_IO' });
    await expect(fsp.stat(backup)).rejects.toMatchObject({ code: 'ENOENT' });
  });
  it('exclusively publishes one complete output in a five-writer race', async () => {
    const out = path.join(dir, 'out');
    const results = await Promise.allSettled(Array.from({ length: 5 }, () => writeFileExclusive(out, data)));
    expect(results.filter((x) => x.status === 'fulfilled')).toHaveLength(1);
    expect(await fsp.readFile(out)).toEqual(data);
    expect((await fsp.readdir(dir)).filter((n) => n.includes('astra-tmp'))).toEqual([]);
  });
  it('atomically replaces the file and preserves its bytes on a failed rename', async () => {
    const out = path.join(dir, 'out'); await fsp.writeFile(out, 'old');
    await replaceFileAtomic(out, data); expect(await fsp.readFile(out)).toEqual(data);
    vi.spyOn(fsp, 'rename').mockRejectedValueOnce(Object.assign(new Error('injected rename'), { code: 'EIO' }));
    await expect(replaceFileAtomic(out, Buffer.from('new'))).rejects.toMatchObject({ outcome: 'not-replaced' });
    expect(await fsp.readFile(out)).toEqual(data);
  });
  it.skipIf(process.platform === 'win32')('refuses read-only and symlink targets', async () => {
    const out = path.join(dir, 'out'); await fsp.writeFile(out, 'old'); await fsp.chmod(out, 0o400);
    await expect(replaceFileAtomic(out, data)).rejects.toMatchObject({ outcome: 'not-replaced' });
    await fsp.chmod(out, 0o600);
    const link = path.join(dir, 'link'); await fsp.symlink(out, link);
    await expect(replaceFileAtomic(link, data)).rejects.toMatchObject({ outcome: 'not-replaced' });
    expect(await fsp.readFile(out, 'utf8')).toBe('old');
  });
  it('a partial temporary write never changes or creates the destination', async () => {
    const out = path.join(dir, 'out');
    await fsp.writeFile(out, 'old');
    const open = fsp.open.bind(fsp);
    const openSpy = vi.spyOn(fsp, 'open');
    openSpy.mockImplementation(async (...args: Parameters<typeof fsp.open>) => {
      const h = await open(...args);
      if (String(args[0]).includes('.astra-tmp-')) {
        const write = h.write.bind(h);
        vi.spyOn(h, 'write').mockImplementationOnce(async () => {
          await write(Buffer.from('partial'));
          throw Object.assign(new Error('injected partial write'), { code: 'EIO' });
        });
      }
      return h;
    });
    await expect(replaceFileAtomic(out, data)).rejects.toMatchObject({ outcome: 'not-replaced' });
    expect(await fsp.readFile(out, 'utf8')).toBe('old');
    const newOut = path.join(dir, 'new');
    await expect(writeFileExclusive(newOut, data)).rejects.toThrow('injected partial write');
    await expect(fsp.stat(newOut)).rejects.toMatchObject({ code: 'ENOENT' });
  });
  it('cleans the candidate when closing the verified temp fails', async () => {
    const out = path.join(dir, 'close-failure-output');
    const open = fsp.open.bind(fsp);
    vi.spyOn(fsp, 'open').mockImplementation(async (...args) => {
      const handle = await open(...args);
      if (String(args[0]).includes('.astra-tmp-')) {
        const close = handle.close.bind(handle);
        vi.spyOn(handle, 'close').mockImplementationOnce(async () => {
          await close();
          throw Object.assign(new Error('injected file close failure'), { code: 'EIO' });
        });
      }
      return handle;
    });
    let caught: unknown;
    try { await writeFileExclusive(out, data); } catch (error) { caught = error; }
    expect(caught).toMatchObject({ code: 'EIO' });
    expect(caught).not.toHaveProperty('tempPath');
    expect(String(caught)).toContain('injected file close failure');
    await expect(fsp.stat(out)).rejects.toMatchObject({ code: 'ENOENT' });
    expect((await fsp.readdir(dir)).some((name) => name.includes('.astra-tmp-'))).toBe(false);
  });
  it('a failure reported after rename is reported as replaced when complete new bytes are verified', async () => {
    const out = path.join(dir, 'out'); await fsp.writeFile(out, 'old');
    const rename = fsp.rename.bind(fsp);
    vi.spyOn(fsp, 'rename').mockImplementationOnce(async (...args) => {
      await rename(...args);
      throw Object.assign(new Error('after rename'), { code: 'EIO' });
    });
    await expect(replaceFileAtomic(out, data)).rejects.toMatchObject({ outcome: 'replaced' });
    expect(await fsp.readFile(out)).toEqual(data);
  });
  it('has no unsafe fallback where hard links are unsupported', async () => {
    const out = path.join(dir, 'out');
    vi.spyOn(fsp, 'link').mockRejectedValueOnce(Object.assign(new Error('unsupported'), { code: 'ENOTSUP' }));
    await expect(writeFileExclusive(out, data)).rejects.toThrow('does not support safe exclusive publication');
    await expect(fsp.stat(out)).rejects.toMatchObject({ code: 'ENOENT' });
  });
  it('reports both publication and temporary cleanup failures with the retained path', async () => {
    const out = path.join(dir, 'out');
    vi.spyOn(fsp, 'link').mockRejectedValueOnce(Object.assign(new Error('injected publish failure'), { code: 'EIO' }));
    vi.spyOn(fsp, 'unlink').mockRejectedValueOnce(Object.assign(new Error('injected cleanup failure'), { code: 'EACCES' }));
    let caught: unknown;
    try { await writeFileExclusive(out, data); } catch (error) { caught = error; }
    expect(caught).toMatchObject({ tempPath: expect.stringContaining('.astra-tmp-') });
    expect(String(caught)).toContain('injected publish failure');
    expect(String(caught)).toContain('injected cleanup failure');
    await expect(fsp.stat(out)).rejects.toMatchObject({ code: 'ENOENT' });
  });
  it('reports that a destination exists when cleanup fails after publication', async () => {
    const out = path.join(dir, 'out');
    vi.spyOn(fsp, 'unlink').mockRejectedValueOnce(Object.assign(new Error('injected cleanup failure'), { code: 'EACCES' }));
    await expect(writeFileExclusive(out, data)).rejects.toMatchObject({ published: true, tempPath: expect.any(String) });
    expect(await fsp.readFile(out)).toEqual(data);
  });
  it('a temp-file fsync failure leaves the old target intact and removes the candidate', async () => {
    const out = path.join(dir, 'sync-failure-target');
    await fsp.writeFile(out, 'old bytes');
    const open = fsp.open.bind(fsp);
    vi.spyOn(fsp, 'open').mockImplementation(async (...args) => {
      const handle = await open(...args);
      if (String(args[0]).includes('.astra-tmp-')) vi.spyOn(handle, 'sync').mockRejectedValueOnce(Object.assign(new Error('injected file sync failure'), { code: 'EIO' }));
      return handle;
    });
    await expect(replaceFileAtomic(out, data)).rejects.toMatchObject({ outcome: 'not-replaced' });
    expect(await fsp.readFile(out, 'utf8')).toBe('old bytes');
    expect((await fsp.readdir(dir)).some((name) => name.includes('.astra-tmp-'))).toBe(false);
  });
  it('a pre-publication capture link failure returns unavailable and can be retried', async () => {
    await initAstraRepo(doc, '<a/>', {});
    vi.spyOn(fsp, 'link').mockRejectedValueOnce(Object.assign(new Error('injected pre-publish link failure'), { code: 'EIO' }));
    expect(await captureOriginal(doc, data)).toMatchObject({ available: false, reason: 'capture-failed' });
    const originalDir = path.join(getRepoDir(doc), '.git', 'astra-original');
    expect(await fsp.stat(path.join(originalDir, 'original.snapshot')).then(() => true, () => false)).toBe(false);
    expect((await fsp.readdir(originalDir)).some((name) => name.includes('.astra-tmp-'))).toBe(false);
    vi.restoreAllMocks();
    expect((await captureOriginal(doc, data)).available).toBe(true);
    expect(await readOriginal(doc)).toEqual(data);
  });
  it('directory-sync failures leave the already published file complete and readable', async () => {
    await initAstraRepo(doc, '<a/>', {});
    const target = path.join(getRepoDir(doc), '.git', 'astra-original', 'exclusive-output.docx');
    const directory = path.dirname(target);
    for (const failure of ['directory-open', 'directory-sync', 'directory-close']) {
      const out = `${target}.${failure}`;
      const open = fsp.open.bind(fsp);
      vi.spyOn(fsp, 'open').mockImplementation(async (...args) => {
        if (failure === 'directory-open' && String(args[0]) === directory) throw Object.assign(new Error('injected directory open failure'), { code: 'EACCES' });
        const handle = await open(...args);
        if (failure === 'directory-sync' && String(args[0]) === directory) vi.spyOn(handle, 'sync').mockRejectedValueOnce(Object.assign(new Error('injected directory sync failure'), { code: 'EIO' }));
        if (failure === 'directory-close' && String(args[0]) === directory) {
          const close = handle.close.bind(handle);
          vi.spyOn(handle, 'close').mockImplementationOnce(async () => {
            await close();
            throw Object.assign(new Error('injected directory close failure'), { code: 'EIO' });
          });
        }
        return handle;
      });
      await expect(writeFileExclusive(out, data)).resolves.toBeUndefined();
      expect(await fsp.readFile(out)).toEqual(data);
      vi.restoreAllMocks();
    }
  });
  it('reports unconfirmed directory durability without undoing a successful atomic replacement', async () => {
    const out = path.join(dir, 'directory-close-replace');
    await fsp.writeFile(out, 'old bytes');
    const open = fsp.open.bind(fsp);
    vi.spyOn(fsp, 'open').mockImplementation(async (...args) => {
      const handle = await open(...args);
      if (String(args[0]) === dir) {
        const close = handle.close.bind(handle);
        vi.spyOn(handle, 'close').mockImplementationOnce(async () => {
          await close();
          throw Object.assign(new Error('injected directory close failure'), { code: 'EIO' });
        });
      }
      return handle;
    });
    await expect(replaceFileAtomic(out, data)).resolves.toMatchObject({ outcome: 'replaced', durabilityUnconfirmed: true });
    expect(await fsp.readFile(out)).toEqual(data);
  });
});
