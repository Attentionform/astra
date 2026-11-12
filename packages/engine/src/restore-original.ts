/** Restore a captured original to its working path without reserializing its bytes. */
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import git from 'isomorphic-git';
import { DocxParser } from './parser';
import type { DocxParserMacroOptions } from './parser';
import { checkoutCommit, commitChange, getHistory, getRepoDir, isTrackedPackageFile } from './git';
import { createVerifiedBackup, originalSha256, readOriginal, replaceFileAtomic } from './original';
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
  const repoDir = getRepoDir(target);
  const indexBefore = JSON.stringify(await git.statusMatrix({ fs, dir: repoDir }));
  const backup = path.join(path.dirname(target), `.${path.basename(target)}.astra-restore-backup-${randomUUID()}`);
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
    const outcome = (err as { outcome?: string })?.outcome;
    let recovered = true;
    if (replaced || outcome === 'replaced' || outcome === 'uncertain') {
      try { await replaceFileAtomic(target, before); recovered = originalSha256(await fsp.readFile(target)) === beforeHash; } catch { recovered = false; }
    }
    let historyChanged = false;
    if (attemptedCommit) {
      try { historyChanged = (await getHistory(target))[0]?.oid !== head || JSON.stringify(await git.statusMatrix({ fs, dir: repoDir })) !== indexBefore; }
      catch { historyChanged = true; }
    }
    if (recovered) await fsp.unlink(backup).catch(() => undefined);
    const cause = err instanceof Error ? err.message : String(err);
    throw Object.assign(new Error(`${cause}${historyChanged ? ' History changed despite failure; inspect it before retrying.' : ''}${recovered ? '' : ` Rollback failed; recovery copy: ${backup}.`}`), { code: (err as { code?: string })?.code ?? 'E_IO', ...(recovered ? {} : { backupPath: backup }) });
  }
}
