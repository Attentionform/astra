import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createHash } from 'crypto';
import * as engine from '@attentionform/astra-engine';
import { createBlankDocxBuffer, getHistory, getRepoDir } from '@attentionform/astra-engine';
import { run } from '../src/cli';

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
afterEach(() => { vi.restoreAllMocks(); fs.rmSync(dir, { recursive: true, force: true }); });

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
    expect((await cli(['batch', file, '--append', 'Saved change'])).status).toBe(0);
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
  it('leaves bytes and history unchanged when backup creation fails', async () => {
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    expect((await cli(['batch', file, '--append', 'Saved change'])).status).toBe(0);
    const before = fs.readFileSync(file);
    const beforeHistory = (await getHistory(file)).map((entry) => entry.oid);
    vi.spyOn(engine, 'createVerifiedBackup').mockRejectedValueOnce(Object.assign(new Error('injected backup creation failure'), { code: 'EACCES' }));
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
  it('inspects and cleans only offline-classified artifacts with explicit flags', async () => {
    const original = fs.readFileSync(file);
    expect((await cli(['history', file, '--init'])).status).toBe(0);
    const originalDir = path.join(getRepoDir(file), '.git', 'astra-original');
    const temp = '.original.snapshot.astra-tmp-123e4567-e89b-12d3-a456-426614174000';
    const secret = 'sensitive interrupted temporary content';
    fs.writeFileSync(path.join(originalDir, temp), secret);
    fs.writeFileSync(path.join(originalDir, 'unknown.txt'), 'preserve me');

    const inspected = await cli(['history', file, '--inspect-recovery-artifacts']);
    expect(inspected.status).toBe(0);
    expect(inspected.body.directory).toBe(originalDir);
    expect(JSON.stringify(inspected.body)).not.toContain(secret);
    expect(inspected.body.totalBytesIsEstimate).toBe(true);
    const artifacts = inspected.body.artifacts as Array<{ name: string; eligibleForCleanup: boolean }>;
    expect(artifacts.find((item) => item.name === temp)?.eligibleForCleanup).toBe(true);

    expect((await cli(['history', file, '--clean-recovery-artifacts'])).status).toBe(2);
    expect(fs.existsSync(path.join(originalDir, temp))).toBe(true);
    const cleaned = await cli(['history', file, '--clean-recovery-artifacts', '--assume-quiescent', '--force']);
    expect(cleaned.status).toBe(0);
    expect(cleaned.body.ok).toBe(true);
    expect(cleaned.body.removed).toEqual([temp]);
    expect(fs.existsSync(path.join(originalDir, temp))).toBe(false);
    expect(fs.readFileSync(path.join(originalDir, 'original.snapshot'))).toBeTruthy();
    expect(fs.readFileSync(path.join(originalDir, 'unknown.txt'), 'utf8')).toBe('preserve me');
    expect(fs.readFileSync(file)).toEqual(original);
  });
  it('inspects and preserves an interrupted initialization lock even before .git exists', async () => {
    const repoDir = getRepoDir(file);
    fs.mkdirSync(repoDir, { recursive: true });
    const lock = path.join(repoDir, '.astra-init.lock');
    fs.writeFileSync(lock, JSON.stringify({ version: 1, pid: 999999, token: 'interrupted' }));
    expect(fs.existsSync(path.join(repoDir, '.git'))).toBe(false);
    const inspection = await cli(['history', file, '--inspect-recovery-artifacts']);
    expect(inspection.status).toBe(0);
    expect(inspection.body.initializationLock).toMatchObject({ path: lock, state: 'present' });
    const cleanup = await cli(['history', file, '--clean-recovery-artifacts', '--assume-quiescent', '--force']);
    expect(cleanup.status).toBe(1);
    expect(fs.existsSync(lock)).toBe(true);
    expect(fs.existsSync(path.join(repoDir, '.git'))).toBe(false);
  });
});
