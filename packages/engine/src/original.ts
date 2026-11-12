/** Exact-original recovery. Original objects are immutable and never enter the package-part tree. */
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { getRepoDir, getSidecarBytes, resolveHistoryMaxSidecarBytes } from './git';
import { assertWriteTargetNoSymlink } from './sandbox/paths';

export type OriginalUnavailableReason = 'predates-feature' | 'capture-failed' | 'incomplete' | 'corrupt' | 'too-large' | 'not-applicable';
export type OriginalInfo =
  | { available: true; sha256: string; bytes: number; capturedAt: string }
  | { available: false; reason: OriginalUnavailableReason; detail?: string };
interface OriginalMeta { version: 1; sha256: string; bytes: number; capturedAt: string; sourceName: string; state: 'complete' }
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
  try { h = await fsp.open(dir, fs.constants.O_RDONLY); await h.sync(); return true; }
  catch { return false; } finally { await h?.close(); }
}
async function cleanupTemp(file: string): Promise<void> {
  try { await fsp.unlink(file); } catch (e) {
    if (errno(e) !== 'ENOENT') throw Object.assign(coded(`Temporary file cleanup failed: ${file}. ${e instanceof Error ? e.message : String(e)}`), { tempPath: file });
  }
}
async function preparedTemp(target: string, data: Buffer, mode = 0o600): Promise<string> {
  assertWriteTargetNoSymlink(target);
  const dir = path.dirname(target);
  await fsp.mkdir(dir, { recursive: true, mode: 0o700 });
  assertWriteTargetNoSymlink(target);
  const tmp = path.join(dir, `.${path.basename(target).slice(0, 60)}.astra-tmp-${randomUUID()}`);
  const h = await fsp.open(tmp, fs.constants.O_WRONLY | fs.constants.O_CREAT | fs.constants.O_EXCL | fs.constants.O_NOFOLLOW, mode);
  try { await h.writeFile(data); await h.sync(); }
  catch (e) { await h.close(); try { await cleanupTemp(tmp); } catch (cleanup) { throw coded(`${e instanceof Error ? e.message : String(e)}; ${cleanup instanceof Error ? cleanup.message : String(cleanup)}`); } throw e; }
  await h.close();
  try {
    const check = await readSafe(tmp, data.length);
    if (!check.equals(data)) throw coded('Original recovery refused: temporary file integrity check failed.');
  } catch (e) { try { await cleanupTemp(tmp); } catch (cleanup) { throw coded(`${e instanceof Error ? e.message : String(e)}; ${cleanup instanceof Error ? cleanup.message : String(cleanup)}`); } throw e; }
  return tmp;
}
/** Publish a complete file exclusively. No partial destination and no unsafe fallback on unsupported filesystems. */
export async function writeFileExclusive(target: string, data: Buffer): Promise<void> {
  const tmp = await preparedTemp(target, data);
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
/** Same-directory replacement. Errors carry an outcome so the caller can recover honestly. */
export async function replaceFileAtomic(target: string, data: Buffer): Promise<AtomicReplaceResult> {
  let tmp: string | undefined;
  let replaced = false;
  let renameStarted = false;
  try {
    assertWriteTargetNoSymlink(target);
    let mode = 0o600;
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
    if (tmp) {
      try { await cleanupTemp(tmp); }
      catch (cleanup) { message += `; ${cleanup instanceof Error ? cleanup.message : String(cleanup)}`; }
    }
    throw Object.assign(coded(message, errno(e) || 'E_IO'), {
      outcome: replaced ? 'replaced' : renameStarted ? 'uncertain' : 'not-replaced',
    });
  }
}

async function verified(docPath: string): Promise<{ meta: OriginalMeta; bytes: Buffer }> {
  const dir = originalDir(docPath);
  const raw = await readSafe(path.join(dir, 'meta.json'), 16 * 1024);
  const meta = JSON.parse(raw.toString('utf8')) as OriginalMeta;
  if (meta.version !== 1 || meta.state !== 'complete' || !/^[a-f0-9]{64}$/.test(meta.sha256) || !Number.isSafeInteger(meta.bytes) || meta.bytes < 1 || typeof meta.capturedAt !== 'string') {
    throw coded('Original metadata is invalid.', 'E_VALIDATION');
  }
  const stat = await fsp.lstat(path.join(dir, `original-${meta.sha256}.docx`));
  if (!stat.isFile() || stat.size !== meta.bytes || stat.size > resolveHistoryMaxSidecarBytes()) throw coded('Original size does not match its metadata.', 'E_VALIDATION');
  const bytes = await readSafe(path.join(dir, `original-${meta.sha256}.docx`));
  if (bytes.length !== meta.bytes || originalSha256(bytes) !== meta.sha256) throw coded('Original integrity verification failed.', 'E_VALIDATION');
  return { meta, bytes };
}
export async function getOriginalInfo(docPath: string): Promise<OriginalInfo> {
  if (!path.isAbsolute(docPath)) return { available: false, reason: 'not-applicable' };
  const dir = originalDir(docPath);
  try {
    if (!(await exists(dir))) return { available: false, reason: 'predates-feature' };
    assertWriteTargetNoSymlink(path.join(dir, 'meta.json'));
    if (!(await exists(path.join(dir, 'meta.json')))) {
      try {
        const status = JSON.parse((await readSafe(path.join(dir, 'status.json'), 16 * 1024)).toString('utf8')) as { reason: OriginalUnavailableReason; detail?: string };
        if (['capture-failed', 'too-large', 'not-applicable'].includes(status.reason)) return { available: false, reason: status.reason, ...(typeof status.detail === 'string' ? { detail: status.detail.slice(0, 1024) } : {}) };
      } catch { /* directory without a committed marker is incomplete */ }
      return { available: false, reason: 'incomplete' };
    }
    const { meta } = await verified(docPath);
    return { available: true, sha256: meta.sha256, bytes: meta.bytes, capturedAt: meta.capturedAt };
  } catch (e) { return { available: false, reason: 'corrupt', detail: e instanceof Error ? e.message : String(e) }; }
}
/** Called ONLY when creating a new history store; existing stores are never backfilled with a later document. */
export async function captureOriginal(docPath: string, bytes?: Buffer): Promise<OriginalInfo> {
  const dir = originalDir(docPath);
  let result: OriginalInfo;
  try {
    assertWriteTargetNoSymlink(path.join(dir, 'meta.json'));
    await fsp.mkdir(dir, { recursive: true, mode: 0o700 });
    assertWriteTargetNoSymlink(path.join(dir, 'meta.json'));
    if (await exists(path.join(dir, 'meta.json'))) {
      const info = await getOriginalInfo(docPath);
      if (info.available && bytes && originalSha256(bytes) !== info.sha256) return { available: false, reason: 'capture-failed', detail: 'An original is already recorded for this history.' };
      return info;
    }
    if (!bytes) result = { available: false, reason: 'not-applicable' };
    else if (!Buffer.isBuffer(bytes) || bytes.length === 0) result = { available: false, reason: 'capture-failed', detail: 'Original input is empty.' };
    else if ((await getSidecarBytes(getRepoDir(docPath))) + bytes.length + 4096 > resolveHistoryMaxSidecarBytes()) result = { available: false, reason: 'too-large' };
    else {
      const digest = originalSha256(bytes);
      const object = path.join(dir, `original-${digest}.docx`);
      try { await writeFileExclusive(object, bytes); }
      catch (e) { if (errno(e) !== 'E_CONFLICT') throw e; if (!(await readSafe(object)).equals(bytes)) throw coded('Existing original object failed integrity verification.'); }
      if ((await getSidecarBytes(getRepoDir(docPath))) + 4096 > resolveHistoryMaxSidecarBytes()) {
        return { available: false, reason: 'too-large', detail: 'Original publication exceeds the current history quota.' };
      }
      const meta: OriginalMeta = { version: 1, state: 'complete', sha256: digest, bytes: bytes.length, capturedAt: new Date().toISOString(), sourceName: path.basename(docPath) };
      try { await writeFileExclusive(path.join(dir, 'meta.json'), Buffer.from(JSON.stringify(meta))); }
      catch (e) { if (errno(e) !== 'E_CONFLICT') throw e; }
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
  if (!info.available) throw coded(`Original unavailable: ${info.reason}${info.detail ? ` (${info.detail})` : ''}`, 'E_STATE');
  return (await verified(docPath)).bytes;
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
  if (opts?.overwrite) await replaceFileAtomic(outPath, bytes); else await writeFileExclusive(outPath, bytes);
  return getOriginalInfo(docPath);
}
