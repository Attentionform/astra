/** Exact-original recovery. Original objects are immutable and never enter the package-part tree. */
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { getRepoDir, getSidecarBytes, resolveHistoryMaxSidecarBytes } from './git';
import { assertWriteTargetNoSymlink } from './sandbox/paths';

export type OriginalUnavailableReason = 'predates-feature' | 'capture-failed' | 'incomplete' | 'corrupt' | 'too-large' | 'not-applicable' | 'initializing';
export type OriginalInfo =
  | { available: true; sha256: string; bytes: number; capturedAt: string }
  | { available: false; reason: OriginalUnavailableReason; detail?: string };
interface SnapshotMeta { magic: 'ASTRA-ORIGINAL'; version: 1; sha256: string; bytes: number; capturedAt: string; sourceName: string }
const SNAPSHOT_NAME = 'original.snapshot';
const SNAPSHOT_MAGIC = Buffer.from('ASTRA-ORIGINAL\n');
const MAX_SNAPSHOT_HEADER = 4096;
export interface AtomicReplaceResult { outcome: 'replaced'; durabilityUnconfirmed?: boolean }
export function originalSha256(bytes: Buffer): string { return createHash('sha256').update(bytes).digest('hex'); }
function coded(message: string, code = 'E_IO'): Error & { code: string } { return Object.assign(new Error(message), { code }); }
function errno(e: unknown): string { return typeof e === 'object' && e !== null && 'code' in e ? String(e.code) : ''; }
function originalDir(docPath: string): string { return path.join(getRepoDir(docPath), '.git', 'astra-original'); }
async function readSafe(file: string, maxBytes = resolveHistoryMaxSidecarBytes()): Promise<Buffer> {
  assertWriteTargetNoSymlink(file);
  const h = await fsp.open(file, fs.constants.O_RDONLY | fs.constants.O_NOFOLLOW);
  try { const stat = await h.stat(); if (!stat.isFile() || stat.size > maxBytes) throw coded('Original recovery refused: expected a regular file within the size limit.'); return await h.readFile(); }
  finally { await h.close(); }
}
async function exists(file: string): Promise<boolean> {
  try { await fsp.lstat(file); return true; } catch (e) { if (errno(e) === 'ENOENT') return false; throw e; }
}
async function syncDirectory(dir: string): Promise<boolean> {
  if (process.platform === 'win32') return false;
  let h: fsp.FileHandle | undefined;
  let synced = false;
  try {
    h = await fsp.open(dir, fs.constants.O_RDONLY);
    await h.sync();
    synced = true;
  } catch {
    /* directory durability is reported separately from publication success */
  }
  try { await h?.close(); } catch { synced = false; }
  return synced;
}
async function cleanupTemp(file: string): Promise<void> {
  try { await fsp.unlink(file); } catch (e) {
    if (errno(e) !== 'ENOENT') throw Object.assign(coded(`Temporary file cleanup failed: ${file}. ${e instanceof Error ? e.message : String(e)}`), { tempPath: file });
  }
}
async function preparedTempParts(target: string, parts: readonly Buffer[], mode = 0o600): Promise<string> {
  assertWriteTargetNoSymlink(target);
  const dir = path.dirname(target);
  await fsp.mkdir(dir, { recursive: true, mode: 0o700 });
  assertWriteTargetNoSymlink(target);
  const tmp = path.join(dir, `.${path.basename(target).slice(0, 60)}.astra-tmp-${randomUUID()}`);
  const h = await fsp.open(tmp, fs.constants.O_WRONLY | fs.constants.O_CREAT | fs.constants.O_EXCL | fs.constants.O_NOFOLLOW, mode);
  let writeError: unknown;
  try {
    for (const part of parts) {
      let offset = 0;
      while (offset < part.length) {
        const { bytesWritten } = await h.write(part, offset, part.length - offset, null);
        if (bytesWritten === 0) throw coded('Original recovery refused: temporary file write made no progress.');
        offset += bytesWritten;
      }
    }
    await h.sync();
  }
  catch (error) { writeError = error; }
  let closeError: unknown;
  try { await h.close(); } catch (error) { closeError = error; }
  if (writeError || closeError) {
    const primary = writeError ?? closeError;
    let message = primary instanceof Error ? primary.message : String(primary);
    if (writeError && closeError) message += `; temporary file close also failed: ${closeError instanceof Error ? closeError.message : String(closeError)}`;
    let cleanupFailed = false;
    try { await cleanupTemp(tmp); }
    catch (cleanup) { cleanupFailed = true; message += `; ${cleanup instanceof Error ? cleanup.message : String(cleanup)}`; }
    throw Object.assign(coded(message, errno(primary) || 'E_IO'), { ...(cleanupFailed ? { tempPath: tmp } : {}), cause: primary });
  }
  try {
    const expectedBytes = parts.reduce((total, part) => total + part.length, 0);
    const expectedHash = createHash('sha256');
    for (const part of parts) expectedHash.update(part);
    const check = await fsp.open(tmp, fs.constants.O_RDONLY | fs.constants.O_NOFOLLOW);
    try {
      const stat = await check.stat();
      if (!stat.isFile() || stat.size !== expectedBytes) throw coded('Original recovery refused: temporary file integrity check failed.');
      const actualHash = createHash('sha256');
      const chunk = Buffer.allocUnsafe(64 * 1024);
      let position = 0;
      while (position < expectedBytes) {
        const { bytesRead } = await check.read(chunk, 0, Math.min(chunk.length, expectedBytes - position), position);
        if (bytesRead === 0) throw coded('Original recovery refused: temporary file integrity check failed.');
        actualHash.update(chunk.subarray(0, bytesRead));
        position += bytesRead;
      }
      if (actualHash.digest('hex') !== expectedHash.digest('hex')) throw coded('Original recovery refused: temporary file integrity check failed.');
    } finally { await check.close(); }
  } catch (e) { try { await cleanupTemp(tmp); } catch (cleanup) { throw coded(`${e instanceof Error ? e.message : String(e)}; ${cleanup instanceof Error ? cleanup.message : String(cleanup)}`); } throw e; }
  return tmp;
}
async function preparedTemp(target: string, data: Buffer, mode = 0o600): Promise<string> {
  return preparedTempParts(target, [data], mode);
}
/** Publish a complete file exclusively. No partial destination and no unsafe fallback on unsupported filesystems. */
async function writePartsExclusive(target: string, parts: readonly Buffer[], mode = 0o600): Promise<void> {
  const tmp = await preparedTempParts(target, parts, mode);
  let failure: unknown;
  try {
    assertWriteTargetNoSymlink(target);
    try { await fsp.link(tmp, target); }
    catch (e) {
      if (errno(e) === 'EEXIST') throw coded(`Destination exists: ${target}`, 'E_CONFLICT');
      if (['ENOTSUP', 'EOPNOTSUPP', 'EPERM', 'EXDEV', 'ENOSYS'].includes(errno(e))) {
        throw coded('This location does not support safe exclusive publication; export to another directory.');
      }
      throw e;
    }
    await syncDirectory(path.dirname(target));
  } catch (error) {
    failure = error;
  }
  try { await cleanupTemp(tmp); }
  catch (cleanup) {
    const detail = cleanup instanceof Error ? cleanup.message : String(cleanup);
    if (failure) {
      throw Object.assign(coded(`${failure instanceof Error ? failure.message : String(failure)}; ${detail}`, errno(failure) || 'E_IO'), { tempPath: tmp });
    }
    throw Object.assign(coded(`Destination was published at ${target}, but ${detail}`), { tempPath: tmp, published: true });
  }
  if (failure) throw failure;
}
/** Publish a complete file exclusively. No partial destination and no unsafe fallback on unsupported filesystems. */
/** `mode` is the new file's permission bits before the process umask: private (0o600) unless the caller asks otherwise. */
export async function writeFileExclusive(target: string, data: Buffer, mode = 0o600): Promise<void> {
  return writePartsExclusive(target, [data], mode);
}
/** Create and verify a private backup before an in-place restore begins. */
export async function createVerifiedBackup(target: string, bytes: Buffer): Promise<void> {
  let ownsTarget = false;
  try {
    await writeFileExclusive(target, bytes);
    ownsTarget = true;
    assertWriteTargetNoSymlink(target);
    const written = await fsp.readFile(target);
    if (written.length !== bytes.length || originalSha256(written) !== originalSha256(bytes)) {
      throw coded(`Restore backup verification failed: ${target}`, 'E_IO');
    }
  } catch (error) {
    if ((error as { published?: unknown } | null)?.published === true) ownsTarget = true;
    if (ownsTarget) {
      try { await fsp.unlink(target); }
      catch (cleanupError) {
        if (errno(cleanupError) !== 'ENOENT') {
          throw Object.assign(coded(`${error instanceof Error ? error.message : String(error)}; invalid backup could not be removed: ${target} (${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)})`, 'E_IO'), { backupPath: target });
        }
      }
    }
    throw error;
  }
}
/** Short, collision-resistant sibling path that also works for maximum-length document names. */
export function newRestoreBackupPath(target: string): string {
  return path.join(path.dirname(target), `.astra-restore-backup-${randomUUID()}`);
}
/** Same-directory replacement. Errors carry an outcome so the caller can recover honestly. */
/** An existing target keeps its mode; a target that does not exist yet is created with `newFileMode` (before the umask). */
export async function replaceFileAtomic(target: string, data: Buffer, newFileMode = 0o600): Promise<AtomicReplaceResult> {
  let tmp: string | undefined;
  let replaced = false;
  let renameStarted = false;
  try {
    assertWriteTargetNoSymlink(target);
    let mode = newFileMode;
    try {
      const st = await fsp.lstat(target);
      if (!st.isFile()) throw coded('Atomic replace requires a regular target file.');
      mode = st.mode & 0o777;
      if ((mode & 0o222) === 0) throw coded('Atomic replace refused: target is read-only.');
      await fsp.access(target, fs.constants.W_OK);
    } catch (e) { if (errno(e) !== 'ENOENT') throw e; }
    tmp = await preparedTemp(target, data, mode);
    assertWriteTargetNoSymlink(target);
    renameStarted = true;
    for (let attempt = 0; ; attempt++) {
      try { await fsp.rename(tmp, target); replaced = true; break; }
      catch (e) {
        if (process.platform !== 'win32' || !['EBUSY', 'EPERM', 'EACCES'].includes(errno(e)) || attempt >= 3) throw e;
        await new Promise((resolve) => setTimeout(resolve, 25 * (attempt + 1)));
      }
    }
    const durable = await syncDirectory(path.dirname(target));
    await cleanupTemp(tmp);
    tmp = undefined;
    return { outcome: 'replaced', ...(!durable ? { durabilityUnconfirmed: true } : {}) };
  } catch (e) {
    let message = e instanceof Error ? e.message : String(e);
    let outcome: 'replaced' | 'uncertain' | 'not-replaced' = replaced ? 'replaced' : renameStarted ? 'uncertain' : 'not-replaced';
    if (!replaced && renameStarted && tmp) {
      try {
        // A failed rename that leaves its source temp in place did not replace
        // the target. This is important for Windows sharing violations: keep
        // the verified backup lifecycle truthful instead of treating a known
        // EPERM/EACCES refusal as an ambiguous write.
        const tempStat = await fsp.lstat(tmp);
        if (tempStat.isFile() && !tempStat.isSymbolicLink()) outcome = 'not-replaced';
      } catch (tempError) {
        if (errno(tempError) === 'ENOENT') {
          try {
            assertWriteTargetNoSymlink(target);
            const targetStat = await fsp.lstat(target);
            if (targetStat.isFile() && targetStat.size === data.length && (await fsp.readFile(target)).equals(data)) {
              replaced = true;
              outcome = 'replaced';
            }
          } catch { /* source vanished and the target cannot prove the outcome */ }
        }
      }
    }
    if (tmp) {
      try { await cleanupTemp(tmp); }
      catch (cleanup) { message += `; ${cleanup instanceof Error ? cleanup.message : String(cleanup)}`; }
    }
    throw Object.assign(coded(message, errno(e) || 'E_IO'), {
      outcome,
    });
  }
}

function encodeSnapshot(bytes: Buffer, sourceName: string): Buffer[] {
  const meta: SnapshotMeta = { magic: 'ASTRA-ORIGINAL', version: 1, sha256: originalSha256(bytes), bytes: bytes.length, capturedAt: new Date().toISOString(), sourceName };
  const json = Buffer.from(JSON.stringify(meta));
  if (json.length > MAX_SNAPSHOT_HEADER) throw coded('Original snapshot metadata exceeds its format limit.', 'E_LIMIT');
  const prefix = Buffer.alloc(SNAPSHOT_MAGIC.length + 4);
  SNAPSHOT_MAGIC.copy(prefix);
  prefix.writeUInt32BE(json.length, SNAPSHOT_MAGIC.length);
  return [prefix, json, bytes];
}
function decodeSnapshot(all: Buffer): { meta: SnapshotMeta; bytes: Buffer } {
  const offset = SNAPSHOT_MAGIC.length;
  if (all.length < offset + 4 || !all.subarray(0, offset).equals(SNAPSHOT_MAGIC)) throw coded('Original snapshot has an invalid format header.', 'E_VALIDATION');
  const headerBytes = all.readUInt32BE(offset);
  if (headerBytes < 2 || headerBytes > MAX_SNAPSHOT_HEADER || all.length < offset + 4 + headerBytes) throw coded('Original snapshot has an invalid metadata length.', 'E_VALIDATION');
  let meta: SnapshotMeta;
  try { meta = JSON.parse(all.subarray(offset + 4, offset + 4 + headerBytes).toString('utf8')) as SnapshotMeta; }
  catch { throw coded('Original snapshot metadata is not valid JSON.', 'E_VALIDATION'); }
  const bytes = all.subarray(offset + 4 + headerBytes);
  if (meta.magic !== 'ASTRA-ORIGINAL' || meta.version !== 1 || !/^[a-f0-9]{64}$/.test(meta.sha256) || !Number.isSafeInteger(meta.bytes) || meta.bytes < 1 || meta.bytes !== bytes.length || typeof meta.capturedAt !== 'string' || typeof meta.sourceName !== 'string') {
    throw coded('Original snapshot metadata is invalid.', 'E_VALIDATION');
  }
  if (originalSha256(bytes) !== meta.sha256) throw coded('Original snapshot integrity verification failed.', 'E_VALIDATION');
  return { meta, bytes };
}
async function verifiedSnapshot(docPath: string): Promise<{ meta: SnapshotMeta; bytes: Buffer }> {
  const file = path.join(originalDir(docPath), SNAPSHOT_NAME);
  const raw = await readSafe(file);
  return decodeSnapshot(raw);
}
export async function getOriginalInfo(docPath: string): Promise<OriginalInfo> {
  if (!path.isAbsolute(docPath)) return { available: false, reason: 'not-applicable' };
  const dir = originalDir(docPath);
  try {
    if (!(await exists(dir))) return { available: false, reason: 'predates-feature' };
    const snapshot = path.join(dir, SNAPSHOT_NAME);
    assertWriteTargetNoSymlink(snapshot);
    if (await exists(snapshot)) {
      const { meta } = await verifiedSnapshot(docPath);
      return { available: true, sha256: meta.sha256, bytes: meta.bytes, capturedAt: meta.capturedAt };
    }
    try {
      const status = JSON.parse((await readSafe(path.join(dir, 'status.json'), 16 * 1024)).toString('utf8')) as { reason: OriginalUnavailableReason; detail?: string };
      if (['capture-failed', 'too-large', 'not-applicable'].includes(status.reason)) return { available: false, reason: status.reason, ...(typeof status.detail === 'string' ? { detail: status.detail.slice(0, 1024) } : {}) };
    } catch { /* directory without a committed marker is incomplete */ }
    return { available: false, reason: 'incomplete' };
  } catch (e) { return { available: false, reason: 'corrupt', detail: e instanceof Error ? e.message : String(e) }; }
}
/** Called ONLY when creating a new history store; existing stores are never backfilled with a later document. */
export async function captureOriginal(docPath: string, bytes?: Buffer): Promise<OriginalInfo> {
  const dir = originalDir(docPath);
  let result: OriginalInfo;
  try {
    assertWriteTargetNoSymlink(path.join(dir, SNAPSHOT_NAME));
    await fsp.mkdir(dir, { recursive: true, mode: 0o700 });
    assertWriteTargetNoSymlink(path.join(dir, SNAPSHOT_NAME));
    if (await exists(path.join(dir, SNAPSHOT_NAME))) {
      const info = await getOriginalInfo(docPath);
      if (info.available && bytes && originalSha256(bytes) !== info.sha256) return { available: false, reason: 'capture-failed', detail: 'An original is already recorded for this history.' };
      return info;
    }
    if (!bytes) result = { available: false, reason: 'not-applicable' };
    else if (!Buffer.isBuffer(bytes) || bytes.length === 0) result = { available: false, reason: 'capture-failed', detail: 'Original input is empty.' };
    else if ((await getSidecarBytes(getRepoDir(docPath))) + bytes.length + 4096 + SNAPSHOT_MAGIC.length > resolveHistoryMaxSidecarBytes()) result = { available: false, reason: 'too-large' };
    else {
      const digest = originalSha256(bytes);
      const snapshot = path.join(dir, SNAPSHOT_NAME);
      try { await writePartsExclusive(snapshot, encodeSnapshot(bytes, path.basename(docPath))); }
      catch (e) {
        if ((e as { published?: unknown } | null)?.published === true) {
          const published = await getOriginalInfo(docPath);
          if (published.available && published.sha256 === digest) return published;
          throw e;
        }
        if (errno(e) !== 'E_CONFLICT') throw e;
        const winner = await verifiedSnapshot(docPath);
        if (!winner.bytes.equals(bytes)) return { available: false, reason: 'capture-failed', detail: 'An original is already recorded for this history.' };
      }
      const info = await getOriginalInfo(docPath);
      if (info.available && info.sha256 !== digest) return { available: false, reason: 'capture-failed', detail: 'An original is already recorded for this history.' };
      return info;
    }
  } catch (e) { result = { available: false, reason: 'capture-failed', detail: e instanceof Error ? e.message : String(e) }; }
  try { await writeFileExclusive(path.join(dir, 'status.json'), Buffer.from(JSON.stringify(result))); } catch { /* best effort; incomplete still reports unavailable */ }
  return result;
}
export async function readOriginal(docPath: string): Promise<Buffer> {
  const info = await getOriginalInfo(docPath);
  if (!info.available) {
    throw Object.assign(coded(`Original unavailable: ${info.reason}${info.detail ? ` (${info.detail})` : ''}`, 'E_STATE'), {
      remedy: 'The retained original cannot be trusted, so no substitute is returned. Keep the working file and any separate backup, and check the history status for the reason.',
    });
  }
  return (await verifiedSnapshot(docPath)).bytes;
}
export async function exportOriginal(docPath: string, outPath: string, opts?: { overwrite?: boolean }): Promise<OriginalInfo> {
  const bytes = await readOriginal(docPath);
  const source = await fsp.realpath(docPath).catch((e: unknown) => { if (errno(e) === 'ENOENT') return path.resolve(docPath); throw e; });
  const parent = path.dirname(path.resolve(outPath));
  const destination = path.join(await fsp.realpath(parent).catch((e: unknown) => { if (errno(e) === 'ENOENT') return parent; throw e; }), path.basename(outPath));
  const samePath = (a: string, b: string): boolean => process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b;
  let sameFile = samePath(source, destination);
  try {
    const [a, b] = await Promise.all([fsp.stat(source), fsp.stat(destination)]);
    sameFile ||= a.ino !== 0 && b.ino !== 0 && a.dev === b.dev && a.ino === b.ino;
  } catch (e) { if (errno(e) !== 'ENOENT') throw e; }
  if (sameFile) throw coded('Use restore_original_document to replace the working file.', 'E_CONFLICT');
  const repo = await fsp.realpath(getRepoDir(docPath));
  const normalizedDest = process.platform === 'win32' ? destination.toLowerCase() : destination;
  const normalizedRepo = process.platform === 'win32' ? repo.toLowerCase() : repo;
  if (normalizedDest === normalizedRepo || normalizedDest.startsWith(`${normalizedRepo}${path.sep}`)) {
    throw coded('Original export cannot write inside the document history store.', 'E_CONFLICT');
  }
  // An export is an ordinary user file: it gets the normal permissions (the umask applies), not the private mode of the history store.
  if (opts?.overwrite) await replaceFileAtomic(outPath, bytes, 0o666); else await writeFileExclusive(outPath, bytes, 0o666);
  return getOriginalInfo(docPath);
}
