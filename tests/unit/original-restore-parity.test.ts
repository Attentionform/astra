/**
 * Failure-injection parity for in-place original restore. The engine API, the CLI (`history --restore-original`)
 * and the MCP tool are driven through the same faults and must leave the same file bytes, the same history and
 * the same recovery backup behind, and must report the failure the same way.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';
import * as fsp from 'node:fs/promises';
import * as fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import git from 'isomorphic-git';
import * as engine from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { LEAN6_HANDLERS } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { handleRestoreOriginalDocument } from '../../packages/mcp-server/src/lib/mcp/handlers/original';
import { run } from '../../packages/cli/src/cli';

const fault = vi.hoisted(() => ({ commit: false, stageThenFail: false, backupCreate: false, failRollback: false, uncertainAfterWrite: false, outsideEdit: undefined as { path: string; bytes: Buffer } | undefined, replaceCalls: 0 }));
vi.mock('../../packages/engine/src/git', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../packages/engine/src/git')>();
  return {
    ...actual,
    commitChange: async (...args: Parameters<typeof actual.commitChange>) => {
      if (fault.commit) throw new Error('injected commit failure');
      if (fault.stageThenFail) {
        const repo = actual.getRepoDir(args[0]);
        await fsp.writeFile(path.join(repo, 'injected-stage'), 'unexpected');
        await git.add({ fs, dir: repo, filepath: 'injected-stage' });
        throw new Error('injected staged failure');
      }
      return actual.commitChange(...args);
    },
  };
});
vi.mock('../../packages/engine/src/original', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../packages/engine/src/original')>();
  return {
    ...actual,
    replaceFileAtomic: async (...args: Parameters<typeof actual.replaceFileAtomic>) => {
      fault.replaceCalls++;
      if (fault.failRollback && fault.replaceCalls === 2) throw new Error('injected rollback failure');
      const result = await actual.replaceFileAtomic(...args);
      // The write landed, but the caller is told its outcome is unknown (a crash after the rename, a lost reply).
      if (fault.uncertainAfterWrite && fault.replaceCalls === 1) throw Object.assign(new Error('injected uncertain replacement'), { outcome: 'uncertain' });
      return result;
    },
    createVerifiedBackup: async (...args: Parameters<typeof actual.createVerifiedBackup>) => {
      if (fault.backupCreate) throw Object.assign(new Error('injected backup creation failure'), { code: 'EACCES' });
      const result = await actual.createVerifiedBackup(...args);
      // Another program edits the working file after the backup is verified and before the restore replaces it.
      if (fault.outsideEdit) await fsp.writeFile(fault.outsideEdit.path, fault.outsideEdit.bytes);
      return result;
    },
  };
});

type Surface = 'engine' | 'cli' | 'mcp';
const SURFACES: Surface[] = ['engine', 'cli', 'mcp'];
const dirs: string[] = [];

afterEach(async () => {
  Object.assign(fault, { commit: false, stageThenFail: false, backupCreate: false, failRollback: false, uncertainAfterWrite: false, outsideEdit: undefined, replaceCalls: 0 });
  vi.restoreAllMocks();
  await Promise.all(dirs.splice(0).map(d => fsp.rm(d, { recursive: true, force: true })));
});

/** A saved, edited document with a two-commit history; its first commit holds the retained original. */
async function setup() {
  const dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-restore-parity-'));
  dirs.push(dir);
  const file = path.join(dir, 'report.docx');
  const original = engine.createBlankDocxBuffer('blank');
  await fsp.writeFile(file, original);
  const store = createMcpSessionStore();
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file, validationTimeoutMs: 0 });
  if (!opened.ok) throw new Error(JSON.stringify(opened));
  const handle = store.get(opened.sessionId);
  const paragraph = handle.ast.sections.flatMap(section => section.nodes).find(node => node.type === 'paragraph')!.id;
  const edited = await LEAN6_HANDLERS.execute_batch(store, { sessionId: opened.sessionId, operations: [{ action: 'replace', target: paragraph, text: 'Changed after opening' }] });
  if (!edited.ok) throw new Error(JSON.stringify(edited));
  const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId });
  if (!saved.ok) throw new Error(JSON.stringify(saved));
  const bytes = await fsp.readFile(file);
  expect(bytes.equals(original)).toBe(false);
  const history = (await engine.getHistory(file)).map(entry => entry.oid);
  expect(history.length).toBe(2);
  fault.replaceCalls = 0;
  return { dir, file, original, edited: bytes, history, store, sessionId: opened.sessionId };
}
type Setup = Awaited<ReturnType<typeof setup>>;

async function restore(surface: Surface, s: Setup): Promise<{ failed: boolean; text: string }> {
  if (surface === 'engine') {
    try { await engine.restoreOriginal(s.file); return { failed: false, text: '' }; }
    catch (err) { return { failed: true, text: `${(err as Error).message} ${(err as { remedy?: string }).remedy ?? ''}` }; }
  }
  if (surface === 'cli') {
    vi.spyOn(console, 'log').mockImplementation(() => undefined);
    const errors: string[] = [];
    vi.spyOn(console, 'error').mockImplementation((...args: unknown[]) => { errors.push(args.map(String).join(' ')); });
    const status = await run(['--json', 'history', s.file, '--restore-original']);
    const out = vi.mocked(console.log).mock.calls.map(call => String(call[0])).join('\n');
    return { failed: status !== 0, text: `${out} ${errors.join(' ')}` };
  }
  const result = await handleRestoreOriginalDocument(s.store, { sessionId: s.sessionId });
  return { failed: !result.ok, text: JSON.stringify(result) };
}

const backups = (dir: string) => fs.readdirSync(dir).filter(name => name.includes('astra-restore-backup'));
const historyOids = async (file: string) => (await engine.getHistory(file)).map(entry => entry.oid);

describe.each(SURFACES)('original restore failure injection: %s', (surface) => {
  it('restores the exact original and appends one commit', async () => {
    const s = await setup();
    const result = await restore(surface, s);
    expect(result.failed, result.text).toBe(false);
    expect((await fsp.readFile(s.file)).equals(s.original)).toBe(true);
    const after = await historyOids(s.file);
    expect(after.length).toBe(3);
    expect(after.slice(1)).toEqual(s.history);
    expect(backups(s.dir)).toEqual([]);
  });

  it('commit failure restores the edited file, leaves history alone and removes the backup', async () => {
    const s = await setup();
    fault.commit = true;
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
    expect(await historyOids(s.file)).toEqual(s.history);
    expect(backups(s.dir)).toEqual([]);
  });

  it('a commit that staged files before failing is reported as a history change', async () => {
    const s = await setup();
    fault.stageThenFail = true;
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    expect(result.text).toMatch(/history (may have )?changed despite/i);
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
    expect(await historyOids(s.file)).toEqual(s.history);
  });

  it('backup creation failure changes nothing', async () => {
    const s = await setup();
    fault.backupCreate = true;
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
    expect(await historyOids(s.file)).toEqual(s.history);
    expect(backups(s.dir)).toEqual([]);
  });

  it('a failed rollback keeps the recovery backup and names it', async () => {
    const s = await setup();
    fault.commit = true;
    fault.failRollback = true;
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    const kept = backups(s.dir);
    expect(kept.length).toBe(1);
    expect((await fsp.readFile(path.join(s.dir, kept[0]))).equals(s.edited)).toBe(true);
    expect(result.text).toContain(kept[0]);
    expect(await historyOids(s.file)).toEqual(s.history);
  });

  it('a replacement whose outcome is unknown is rolled back to the edited file and no commit is attempted', async () => {
    const s = await setup();
    fault.uncertainAfterWrite = true;
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
    expect(await historyOids(s.file)).toEqual(s.history);
    expect(result.text).not.toMatch(/history (may have )?changed despite/i);
    expect(fault.replaceCalls).toBe(2);
    expect(backups(s.dir)).toEqual([]);
  });

  it('an unknown replacement outcome with a failed rollback keeps the backup and names it', async () => {
    const s = await setup();
    fault.uncertainAfterWrite = true;
    fault.failRollback = true;
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    const kept = backups(s.dir);
    expect(kept.length).toBe(1);
    expect((await fsp.readFile(path.join(s.dir, kept[0]))).equals(s.edited)).toBe(true);
    expect(result.text).toContain(kept[0]);
    expect(await historyOids(s.file)).toEqual(s.history);
  });

  it('an outside edit made after the backup and before the replacement is never overwritten', async () => {
    const s = await setup();
    const outside = Buffer.concat([s.edited, Buffer.from('edited by another program')]);
    fault.outsideEdit = { path: s.file, bytes: outside };
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    expect(result.text).toMatch(/changed|conflict/i);
    expect((await fsp.readFile(s.file)).equals(outside)).toBe(true);
    expect(await historyOids(s.file)).toEqual(s.history);
    expect(fault.replaceCalls).toBe(0);
    expect(backups(s.dir)).toEqual([]);
  });

  it('a corrupt retained original is refused before the file is touched', async () => {
    const s = await setup();
    await fsp.writeFile(path.join(engine.getRepoDir(s.file), '.git', 'astra-original', 'original.snapshot'), 'corrupt');
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
    expect(await historyOids(s.file)).toEqual(s.history);
    expect(fault.replaceCalls).toBe(0);
  });

  it('a held initialization lock is refused before anything is written', async () => {
    const s = await setup();
    await fsp.writeFile(path.join(engine.getRepoDir(s.file), '.astra-init.lock'), JSON.stringify({ version: 2, pid: 999999, hostname: 'other-host', token: 'held', startedAt: new Date().toISOString() }));
    const result = await restore(surface, s);
    expect(result.failed).toBe(true);
    expect(result.text).toContain('.astra-init.lock');
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
    expect(fault.replaceCalls).toBe(0);
    expect(backups(s.dir)).toEqual([]);
    expect(fs.existsSync(path.join(engine.getRepoDir(s.file), '.astra-init.lock'))).toBe(true);
  });
});

describe('original restore through MCP: session state', () => {
  const snapshotOfSession = (s: Setup) => {
    const h = s.store.get(s.sessionId);
    return { parser: h.parser, ast: h.ast, engine: h.engine, revision: h.revision, savedRevision: h.savedRevision, byteSize: h.byteSize, lastWrittenSha256: h.lastWrittenSha256, baseline: h.validationBaseline, redo: [...h.redoStack] };
  };

  it('a failed restore leaves the live session, its redo stack and its validation baseline exactly as they were', async () => {
    const s = await setup();
    const h = s.store.get(s.sessionId);
    h.redoStack.push(h.parser.capturePackageSnapshot());
    const before = snapshotOfSession(s);
    fault.commit = true;
    const result = await restore('mcp', s);
    expect(result.failed).toBe(true);
    const after = snapshotOfSession(s);
    expect(after.parser).toBe(before.parser);
    expect(after.ast).toBe(before.ast);
    expect(after.engine).toBe(before.engine);
    expect(after.revision).toBe(before.revision);
    expect(after.savedRevision).toBe(before.savedRevision);
    expect(after.byteSize).toBe(before.byteSize);
    expect(after.lastWrittenSha256).toBe(before.lastWrittenSha256);
    expect(after.baseline).toBe(before.baseline);
    expect(after.redo.length).toBe(before.redo.length);
    expect(after.redo[0]).toBe(before.redo[0]);
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
  });

  it('a failure while preparing the candidate leaves the session and the file untouched', async () => {
    const s = await setup();
    const before = snapshotOfSession(s);
    vi.spyOn(engine.DocxParser.prototype, 'getAST').mockRejectedValueOnce(new Error('injected candidate parse failure'));
    const result = await restore('mcp', s);
    expect(result.failed).toBe(true);
    const after = snapshotOfSession(s);
    expect(after.parser).toBe(before.parser);
    expect(after.revision).toBe(before.revision);
    expect(after.redo.length).toBe(before.redo.length);
    expect((await fsp.readFile(s.file)).equals(s.edited)).toBe(true);
    expect(await historyOids(s.file)).toEqual(s.history);
    expect(backups(s.dir)).toEqual([]);
  });

  it('a successful restore adopts the original as the saved session state and keeps the old state redoable', async () => {
    const s = await setup();
    const h = s.store.get(s.sessionId);
    h.redoStack.push(h.parser.capturePackageSnapshot(), h.parser.capturePackageSnapshot());
    const before = snapshotOfSession(s);
    const result = await restore('mcp', s);
    expect(result.failed, result.text).toBe(false);
    const after = snapshotOfSession(s);
    expect(after.parser).not.toBe(before.parser);
    expect(after.revision).toBe(before.revision + 1);
    expect(after.savedRevision).toBe(after.revision);
    expect(after.byteSize).toBe(s.original.length);
    expect(after.lastWrittenSha256).toBe(engine.originalSha256(s.original));
    // Earlier redo entries belong to the lineage before the restore; the pre-restore state is the one redo entry.
    expect(after.redo.length).toBe(1);
    expect((await fsp.readFile(s.file)).equals(s.original)).toBe(true);
    expect((await historyOids(s.file)).length).toBe(3);
  });
});
