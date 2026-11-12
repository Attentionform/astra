/** Restore a captured original to its working path without reserializing its bytes. */
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import git from 'isomorphic-git';
import { DocxParser } from './parser';
import type { DocxParserMacroOptions } from './parser';
import { assertHistoryReadyForWrite, checkoutCommit, commitChange, getHistory, getRepoDir, isTrackedPackageFile } from './git';
import { createVerifiedBackup, newRestoreBackupPath, originalSha256, readOriginal, replaceFileAtomic } from './original';
import { assertWriteTargetNoSymlink } from './sandbox/paths';

export interface RestoreOriginalOptions {
  /** Optional expected hash of the current file, for clients with a saved baseline. */
  expectedSha256?: string;
  /** The original must satisfy the same macro policy as any other opened document. */
  macroPolicy?: DocxParserMacroOptions;
}
export interface RestoreOriginalResult {
  path: string;
  sha256: string;
  commitOid: string;
  durabilityUnconfirmed?: boolean;
  warning?: string;
}

/** What a restore compares after a failure to know whether history moved: the head commit and the index. */
export interface HistoryStateBefore {
  head: string | undefined;
  /** Serialized index status; undefined when it could not be read, which counts as "may have changed". */
  index: string | undefined;
}

/** The history state a restore must be able to compare against after a failed commit. Never throws. */
export async function captureHistoryState(docPath: string): Promise<HistoryStateBefore> {
  const target = path.resolve(docPath);
  const [head, index] = await Promise.all([
    getHistory(target).then(entries => entries[0]?.oid, () => undefined),
    git.statusMatrix({ fs, dir: getRepoDir(target) }).then(matrix => JSON.stringify(matrix), () => undefined),
  ]);
  return { head, index };
}

/**
 * True when the head or the index differs from `before`, and also when that cannot be told: a head or an index
 * that could not be read, before or now, is an unknown outcome and reported as a change. Never throws.
 */
export async function historyChangedSince(docPath: string, before: HistoryStateBefore): Promise<boolean> {
  const after = await captureHistoryState(docPath);
  if (before.head === undefined || after.head === undefined || before.index === undefined || after.index === undefined) return true;
  return after.head !== before.head || after.index !== before.index;
}

/**
 * Put `previous` back after a restore failed. Only needed when the replacement may have happened: it did
 * (`replaced`), or the failure says its outcome is `replaced` or `uncertain`. Verifies by hash; never throws.
 */
export async function rollbackReplacedFile(
  docPath: string,
  previous: Buffer,
  failure: { replaced: boolean; error: unknown }
): Promise<{ recovered: boolean; failure?: string }> {
  const outcome = (failure.error as { outcome?: string } | null)?.outcome;
  if (!failure.replaced && outcome !== 'replaced' && outcome !== 'uncertain') return { recovered: true };
  try {
    await replaceFileAtomic(docPath, previous);
    if (originalSha256(await fsp.readFile(docPath)) === originalSha256(previous)) return { recovered: true };
    return { recovered: false, failure: 'the file does not match the previous bytes after rollback' };
  } catch (error) {
    return { recovered: false, failure: error instanceof Error ? error.message : String(error) };
  }
}

/**
 * Restore the exact originally captured DOCX bytes and append a history commit.
 * The caller must serialize concurrent access to this path, including any in-memory session.
 * The current on-disk document must match a complete history head. A hash check immediately
 * before replacement detects most outside edits, but cannot exclude a writer racing the rename.
 * This low-level file API does not update any live DocxParser, AST, or MCP session.
 */
export async function restoreOriginal(docPath: string, options: RestoreOriginalOptions = {}): Promise<RestoreOriginalResult> {
  const target = path.resolve(docPath);
  assertWriteTargetNoSymlink(target);
  // Refuse before anything is touched while an initialization lock is held or the history is not complete.
  await assertHistoryReadyForWrite(target);
  const before = await fsp.readFile(target);
  const beforeHash = originalSha256(before);
  if (options.expectedSha256 && beforeHash !== options.expectedSha256) {
    throw Object.assign(new Error('Original restore refused: the working file changed since the expected state.'), { code: 'E_CONFLICT' });
  }
  const entries = await getHistory(target);
  const head = entries[0]?.oid;
  if (!head) throw Object.assign(new Error('Original restore refused: history has no recorded state.'), { code: 'E_CONFLICT' });
  const recorded = await checkoutCommit(target, head);
  if (!recorded.complete) throw Object.assign(new Error('Original restore refused: history head is not a complete snapshot.'), { code: 'E_CONFLICT' });
  const currentParser = new DocxParser(before, options.macroPolicy);
  await currentParser.getAST();
  const current = currentParser.capturePackageSnapshot();
  const names = [...current.keys()].filter(isTrackedPackageFile).sort();
  const recordedNames = [...recorded.extra.keys()].sort();
  if (currentParser.readDocumentXml() !== recorded.xml || JSON.stringify(names) !== JSON.stringify(recordedNames) || names.some(n => !current.get(n)!.equals(recorded.extra.get(n)!))) {
    throw Object.assign(new Error('Original restore refused: history differs from the working file. Save or revert before retrying.'), { code: 'E_CONFLICT' });
  }
  const original = await readOriginal(target);
  const originalParser = new DocxParser(original, options.macroPolicy);
  await originalParser.getAST();
  const snapshot = originalParser.capturePackageSnapshot();
  const extras: Record<string, Buffer> = {};
  for (const [name, data] of snapshot) if (isTrackedPackageFile(name)) extras[name] = data;
  const historyBefore = await captureHistoryState(target);
  const backup = newRestoreBackupPath(target);
  await createVerifiedBackup(backup, before);
  let replaced = false;
  let attemptedCommit = false;
  try {
    if (originalSha256(await fsp.readFile(target)) !== beforeHash) throw Object.assign(new Error('Original restore refused: the working file changed during preparation.'), { code: 'E_CONFLICT' });
    const result = await replaceFileAtomic(target, original);
    replaced = true;
    attemptedCommit = true;
    const commitOid = await commitChange(target, originalParser.readDocumentXml(), `Restore original file (sha256 ${originalSha256(original)})`, extras);
    const cleanupFailed = await fsp.unlink(backup).then(() => false, () => true);
    return { path: target, sha256: originalSha256(original), commitOid, ...(result.durabilityUnconfirmed ? { durabilityUnconfirmed: true } : {}), ...(cleanupFailed ? { warning: `Restore succeeded; recovery backup remains at ${backup}.` } : {}) };
  } catch (err) {
    const { recovered } = await rollbackReplacedFile(target, before, { replaced, error: err });
    const historyChanged = attemptedCommit && await historyChangedSince(target, historyBefore);
    if (recovered) await fsp.unlink(backup).catch(() => undefined);
    const cause = err instanceof Error ? err.message : String(err);
    throw Object.assign(new Error(`${cause}${historyChanged ? ' History changed despite failure; inspect it before retrying.' : ''}${recovered ? '' : ` Rollback failed; recovery copy: ${backup}.`}`), { code: (err as { code?: string })?.code ?? 'E_IO', ...(recovered ? {} : { backupPath: backup }) });
  }
}
