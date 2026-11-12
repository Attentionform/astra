import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createHash } from 'crypto';
import * as engine from '@attentionform/astra-engine';
import { createBlankDocxBuffer, getHistory, getRepoDir } from '@attentionform/astra-engine';
import { run } from '../src/cli';

// Restore runs inside the engine, so faults are injected where the engine calls them (its own module), not on
// the package namespace the CLI imports from.
const fault = vi.hoisted(() => ({ durabilityUnconfirmed: 0, backupCreate: false, replaceCalls: 0 }));
vi.mock('../../engine/src/original', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../engine/src/original')>();
  return {
    ...actual,
    replaceFileAtomic: async (...args: Parameters<typeof actual.replaceFileAtomic>) => {
      fault.replaceCalls++;
      const result = await actual.replaceFileAtomic(...args);
      if (fault.durabilityUnconfirmed > 0) { fault.durabilityUnconfirmed--; return { ...result, durabilityUnconfirmed: true }; }
      return result;
    },
    createVerifiedBackup: async (...args: Parameters<typeof actual.createVerifiedBackup>) => {
      if (fault.backupCreate) throw Object.assign(new Error('injected backup creation failure'), { code: 'EACCES' });
      return actual.createVerifiedBackup(...args);
    },
  };
});

let dir: string;
let file: string;
const hash = (data: Buffer): string => createHash('sha256').update(data).digest('hex');
async function cli(args: string[]): Promise<{ status: number; body: Record<string, unknown>; stderr: string }> {
  vi.mocked(console.log).mockClear(); vi.mocked(console.error).mockClear();
  const status = await run(['--json', ...args]);
  const text = vi.mocked(console.log).mock.calls.map(c => String(c[0])).join('\n');
  return { status, body: text ? JSON.parse(text) : {}, stderr: vi.mocked(console.error).mock.calls.map(c => String(c[0])).join('\n') };
}
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-original-')); file = path.join(dir, 'report.docx');
  fs.writeFileSync(file, createBlankDocxBuffer('blank'));
  vi.spyOn(console, 'log').mockImplementation(() => undefined); vi.spyOn(console, 'error').mockImplementation(() => undefined);
});
afterEach(() => { fault.durabilityUnconfirmed = 0; fault.backupCreate = false; fault.replaceCalls = 0; vi.restoreAllMocks(); fs.rmSync(dir, { recursive: true, force: true }); });

describe('CLI exact original recovery', () => {
  it('captures the raw input at explicit init and exports byte-identically without changing the working file', async () => {
    const original = fs.readFileSync(file);
    const init = await cli(['history', file, '--init']);
    expect(init.status).toBe(0); expect(init.body.original).toMatchObject({ available: true, sha256: hash(original) });
    const out = path.join(dir, 'original.docx');
    expect((await cli(['history', file, '--export-original', out])).status).toBe(0);
    expect(fs.readFileSync(out).equals(original)).toBe(true); expect(fs.readFileSync(file).equals(original)).toBe(true);
    const commits = await getHistory(file);
    expect((await cli(['history', file, '--export-original', out])).status).toBe(1);
    expect((await cli(['history', file, '--export-original', out, '--force'])).status).toBe(0);
    expect((await getHistory(file)).length).toBe(commits.length);
  });
  it('restores raw original bytes and appends one recorded state', async () => {
    const original = fs.readFileSync(file);
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    const batch = await cli(['batch', file, '--append', 'Saved change']);
    expect(batch.status, batch.stderr).toBe(0);
    const edited = fs.readFileSync(file);
    const depth = (await getHistory(file)).length;
    const restored = await cli(['history', file, '--restore-original']);
    expect(restored.status).toBe(0); expect(restored.body.restoredOriginal).toBe(true);
    expect(fs.readFileSync(file).equals(original)).toBe(true);
    expect((await getHistory(file)).length).toBe(depth + 1);
    expect((await cli(['revert', file, '--steps', '1'])).status).toBe(0);
    const { DocxParser } = await import('@attentionform/astra-engine');
    expect(new DocxParser(fs.readFileSync(file)).readDocumentXml()).toBe(new DocxParser(edited).readDocumentXml());
    expect((await cli(['redo', file])).status).toBe(0);
    expect(new DocxParser(fs.readFileSync(file)).readDocumentXml()).toBe(new DocxParser(original).readDocumentXml());
    expect(fs.readdirSync(dir).some(n => n.includes('restore-backup'))).toBe(false);
  });
  it('uses a short backup name for a valid near-limit document filename', async () => {
    const longFile = path.join(dir, `${'r'.repeat(220)}.docx`);
    fs.renameSync(file, longFile);
    file = longFile;
    const original = fs.readFileSync(file);
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    const changedZip = new engine.AdmZip(original);
    changedZip.addFile('customXml/item1.xml', Buffer.from('<root>saved</root>'));
    const working = changedZip.toBuffer();
    fs.writeFileSync(file, working);
    const parser = new engine.DocxParser(working);
    await parser.getAST();
    const extras = Object.fromEntries(
      [...parser.capturePackageSnapshot()].filter(([name]) => engine.isTrackedPackageFile(name)),
    );
    await engine.commitChange(file, parser.readDocumentXml(), 'Record a change', extras);
    expect((await cli(['history', file, '--restore-original'])).status).toBe(0);
    expect(fs.readFileSync(file)).toEqual(original);
    expect(fs.readdirSync(dir).some(name => name.includes('astra-restore-backup'))).toBe(false);
  });
  it('reports unconfirmed directory durability after a successful restore', async () => {
    const original = fs.readFileSync(file);
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    expect((await cli(['batch', file, '--append', 'Saved change'])).status).toBe(0);
    fault.durabilityUnconfirmed = 1;

    const restored = await cli(['history', file, '--restore-original']);
    expect(restored.status).toBe(0);
    expect(restored.body.durabilityUnconfirmed).toBe(true);
    expect(fs.readFileSync(file)).toEqual(original);

    expect((await cli(['batch', file, '--append', 'Another saved change'])).status).toBe(0);
    fault.durabilityUnconfirmed = 1;
    vi.mocked(console.log).mockClear();
    const humanStatus = await run(['history', file, '--restore-original']);
    expect(humanStatus).toBe(0);
    expect(vi.mocked(console.log).mock.calls.map(call => String(call[0])).join('\n'))
      .toContain('directory durability unconfirmed');
    expect(fs.readFileSync(file)).toEqual(original);
  });
  it('preflights auxiliary XML before replacing the working file', async () => {
    const originalZip = new engine.AdmZip(fs.readFileSync(file));
    originalZip.updateFile('docProps/core.xml', Buffer.from('<cp:coreProperties>'));
    const original = originalZip.toBuffer();
    fs.writeFileSync(file, original);
    expect((await cli(['history', file, '--init'])).status).toBe(0);

    const workingZip = new engine.AdmZip(original);
    workingZip.deleteFile('docProps/core.xml');
    const working = workingZip.toBuffer();
    fs.writeFileSync(file, working);
    const parser = new engine.DocxParser(working);
    await parser.getAST();
    const extras = Object.fromEntries(
      [...parser.capturePackageSnapshot()].filter(([name]) => engine.isTrackedPackageFile(name)),
    );
    await engine.commitChange(file, parser.readDocumentXml(), 'Remove malformed auxiliary XML', extras);
    const historyBefore = (await getHistory(file)).map(entry => entry.oid);

    const restored = await cli(['history', file, '--restore-original']);
    expect(restored.status).toBe(1);
    expect(fs.readFileSync(file)).toEqual(working);
    expect((await getHistory(file)).map(entry => entry.oid)).toEqual(historyBefore);
  });
  it('leaves bytes and history unchanged when backup creation fails', async () => {
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    expect((await cli(['batch', file, '--append', 'Saved change'])).status).toBe(0);
    const before = fs.readFileSync(file);
    const beforeHistory = (await getHistory(file)).map((entry) => entry.oid);
    fault.backupCreate = true;
    const result = await cli(['history', file, '--restore-original']);
    expect(result.status).toBe(1);
    expect(fs.readFileSync(file)).toEqual(before);
    expect((await getHistory(file)).map((entry) => entry.oid)).toEqual(beforeHistory);
    expect(fs.readdirSync(dir).some((name) => name.includes('astra-restore-backup'))).toBe(false);
  });
  it('refuses an unrecorded disk edit without overwriting it', async () => {
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    const { AdmZip } = await import('@attentionform/astra-engine');
    const zip = new AdmZip(fs.readFileSync(file));
    zip.addFile('customXml/unrecorded.xml', Buffer.from('<changed/>'));
    const edited = zip.toBuffer(); fs.writeFileSync(file, edited);
    const res = await cli(['history', file, '--restore-original', '--force']);
    expect(res.status).toBe(1); expect(fs.readFileSync(file).equals(edited)).toBe(true);
    expect((await getHistory(file)).length).toBe(1);
  });
  it('corrupt retained bytes refuse export and restore without touching the working file', async () => {
    const before = fs.readFileSync(file);
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    const originalDir = path.join(getRepoDir(file), '.git', 'astra-original');
    fs.writeFileSync(path.join(originalDir, 'original.snapshot'), 'corrupt');
    const out = path.join(dir, 'bad.docx');
    expect((await cli(['history', file, '--export-original', out])).status).toBe(1);
    expect(fs.existsSync(out)).toBe(false);
    expect((await cli(['history', file, '--restore-original'])).status).toBe(1);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    expect((await getHistory(file)).length).toBe(1);
  });
  it('history remains opt-in and new recovery flags validate arguments', async () => {
    expect((await cli(['history', file, '--export-original', path.join(dir, 'out.docx')])).status).toBe(2);
    expect(fs.existsSync(getRepoDir(file))).toBe(false);
    for (const extra of [['--export-original'], ['--init', '--restore-original'], ['--force'], ['--restore-original', '--strip-macros']]) {
      expect((await cli(['history', file, ...extra])).status).toBe(2);
    }
  });
  it('the removed cleanup flags are unknown flags', async () => {
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    for (const flag of ['--inspect-recovery-artifacts', '--clean-recovery-artifacts', '--assume-quiescent']) {
      const result = await cli(['history', file, flag]);
      expect(result.status).toBe(2);
    }
  });
  it('a leftover temporary file does not affect history, export or restore', async () => {
    const original = fs.readFileSync(file);
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    const originalDir = path.join(getRepoDir(file), '.git', 'astra-original');
    fs.writeFileSync(path.join(originalDir, '.original.snapshot.astra-tmp-123e4567-e89b-12d3-a456-426614174000'), 'sensitive interrupted temporary content');
    const listing = await cli(['history', file]);
    expect(listing.status).toBe(0);
    expect(listing.body.original).toMatchObject({ available: true });
    const out = path.join(dir, 'leftover-export.docx');
    expect((await cli(['history', file, '--export-original', out])).status).toBe(0);
    expect(fs.readFileSync(out)).toEqual(original);
  });
  it('refuses a changed-input retry after an interrupted initialization, keeps both copies and still exports the earlier bytes', async () => {
    const earlier = fs.readFileSync(file);
    // The state a killed initializer leaves: the snapshot is published, no commit was recorded, the lock was removed by the operator.
    await engine.ensureAstraGitRepo(file);
    const { default: git } = await import('isomorphic-git');
    await git.init({ fs, dir: getRepoDir(file) });
    expect((await engine.captureOriginal(file, earlier)).available).toBe(true);
    const changed = Buffer.concat([earlier, Buffer.from('later text')]);
    fs.writeFileSync(file, changed);

    const refused = await cli(['history', file, '--init']);
    expect(refused.status).toBe(1);
    expect(JSON.stringify(refused.body) + refused.stderr).toContain('do not delete it');
    expect(fs.readFileSync(file).equals(changed)).toBe(true);
    expect(fs.existsSync(path.join(getRepoDir(file), '.git', 'astra-original', 'original.snapshot'))).toBe(true);
    // The history has no commits, so restore has nothing to restore over the changed document.
    expect((await cli(['history', file, '--restore-original'])).status).not.toBe(0);
    expect(fs.readFileSync(file).equals(changed)).toBe(true);
    // The earlier document is retrievable through the CLI.
    const out = path.join(dir, 'earlier.docx');
    expect((await cli(['history', file, '--export-original', out])).status).toBe(0);
    expect(fs.readFileSync(out).equals(earlier)).toBe(true);
  });
  it('restores a macro-enabled original only when --allow-macros says the macro parts are accepted', async () => {
    const base = new engine.AdmZip(fs.readFileSync(file));
    base.addFile('word/vbaProject.bin', Buffer.from('fake-vba-bytes', 'utf8'));
    base.updateFile('word/_rels/document.xml.rels', Buffer.from(base.readAsText('word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rMacro" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'), 'utf8'));
    base.updateFile('[Content_Types].xml', Buffer.from(base.readAsText('[Content_Types].xml').replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default '), 'utf8'));
    const macroBytes = base.toBuffer();
    fs.writeFileSync(file, macroBytes);
    expect((await cli(['history', file, '--init', '--allow-macros'])).status).toBe(0);
    const before = (await getHistory(file)).length;

    const refused = await cli(['history', file, '--restore-original']);
    expect(refused.status).toBe(1);
    expect(fs.readFileSync(file).equals(macroBytes)).toBe(true);
    expect((await getHistory(file)).length).toBe(before);

    const allowed = await cli(['history', file, '--restore-original', '--allow-macros']);
    expect(allowed.status, JSON.stringify(allowed.body) + allowed.stderr).toBe(0);
    expect(fs.readFileSync(file).equals(macroBytes)).toBe(true);
    expect((await getHistory(file)).length).toBe(before + 1);
    expect((await cli(['history', file, '--export-original', path.join(dir, 'm.docx'), '--allow-macros'])).status).toBe(2);
  });
  it('refuses revert, redo and restore before touching the file while an initialization lock is held', async () => {
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    const before = fs.readFileSync(file);
    const lock = path.join(getRepoDir(file), '.astra-init.lock');
    fs.writeFileSync(lock, JSON.stringify({ version: 1, pid: 999999, token: 'old' }));
    fault.replaceCalls = 0;
    for (const args of [['revert', file], ['redo', file], ['history', file, '--restore-original'], ['batch', file, '--append', 'Blocked by lock'], ['batch', file, '--append', 'Blocked by lock', '--sequential']]) {
      const result = await cli(args);
      expect(result.status, args.join(' ')).not.toBe(0);
      expect(JSON.stringify(result.body) + result.stderr).toContain('.astra-init.lock');
      expect(fs.readFileSync(file).equals(before)).toBe(true);
    }
    expect(fault.replaceCalls).toBe(0);
    expect(fs.existsSync(lock)).toBe(true);
  });
  it('refuses batch while an initialization lock is held before .git exists, and writes nothing', async () => {
    const before = fs.readFileSync(file);
    const repo = getRepoDir(file); fs.mkdirSync(repo, { recursive: true });
    const lock = path.join(repo, '.astra-init.lock');
    fs.writeFileSync(lock, JSON.stringify({ version: 2, pid: process.pid, hostname: os.hostname(), token: 'test', startedAt: new Date().toISOString() }));
    expect(fs.existsSync(path.join(repo, '.git'))).toBe(false);
    for (const args of [['batch', file, '--append', 'Blocked by lock'], ['batch', file, '--append', 'Blocked by lock', '--sequential']]) {
      const result = await cli(args);
      expect(result.status, args.join(' ')).not.toBe(0);
      expect(JSON.stringify(result.body) + result.stderr).toContain('.astra-init.lock');
      expect(fs.readFileSync(file).equals(before), args.join(' ')).toBe(true);
    }
    fs.rmSync(lock);
    const ok = await cli(['batch', file, '--append', 'Allowed without a lock', '--no-history']);
    expect(ok.status).toBe(0);
  });
});
