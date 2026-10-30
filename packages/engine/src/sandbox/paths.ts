/**
 * Q0a central sandbox path jail (single source of truth).
 *
 * Consolidates the 7 pre-existing local copies:
 * - packages/engine/src/document-manager.ts:isPathInsideDir + canonicalizeForSandbox
 * - packages/engine/src/compose/markdown.ts:isPathInsideDir + canonicalizeForAsset
 * - packages/engine/src/recovery.ts:isPathInsideDir
 * - packages/cli/src/export.ts:isPathInsideDir + canonicalizeForAsset
 * - packages/mcp-server/src/lib/session-store.ts:isMcpPathInsideDir + canonicalizeForMcp
 *
 * Semantics are byte-identical to the majority (normalize + win32-aware
 * prefix) implementation. `isPathInsideDir` is lexical only (no I/O);
 * callers canonicalize with `canonicalizeForAsset` first when symlink
 * closure is required. `assertInside` composes both and throws E_SANDBOX
 * with a remedy.
 */
import fs from 'fs';
import fsp from 'fs/promises';
import path from 'path';

/**
 * Lexical containment: true when `candidate` equals `dir` or lives under it.
 * Win32-aware (case-insensitive); POSIX case-sensitive. No filesystem I/O.
 */
export function isPathInsideDir(
  candidate: string,
  dir: string,
  platform: string = process.platform
): boolean {
  const normC = path.normalize(candidate);
  const normD = path.normalize(dir);
  if (platform === 'win32') {
    const lc = normC.toLowerCase();
    const ld = normD.toLowerCase();
    return lc === ld || lc.startsWith(ld + path.sep.toLowerCase());
  }
  return normC === normD || normC.startsWith(normD + path.sep);
}

/**
 * Canonicalizes a resolved path for sandbox decisions: resolves symlinks /
 * junctions via realpathSync when the target exists, otherwise realpaths the
 * nearest existing ancestor and re-appends the remainder (save targets that
 * don't exist yet still get their PARENT canonicalized, closing the
 * symlink-parent escape).
 */
export function canonicalizeForAsset(resolved: string): string {
  try {
    return fs.realpathSync(resolved);
  } catch {
    /* not existing — walk up */
  }
  let cur = resolved;
  const stack: string[] = [];
  for (;;) {
    const parent = path.dirname(cur);
    if (parent === cur) return resolved; // filesystem root reached
    stack.unshift(path.basename(cur));
    cur = parent;
    try {
      return path.join(fs.realpathSync(cur), ...stack);
    } catch {
      /* keep walking up */
    }
  }
}

export interface AssertInsideOptions {
  /** Override the human message body (without the E_SANDBOX prefix). */
  message?: string;
  /** Override the remedy attached as `.remedy` and in the message. */
  remedy?: string;
}

/**
 * Canonicalize both sides and throw E_SANDBOX (with `code` + `remedy`) when
 * the candidate escapes `dir`. Returns the canonical candidate on success.
 */
export function assertInside(
  candidate: string,
  dir: string,
  opts?: AssertInsideOptions
): string {
  const canonicalCandidate = canonicalizeForAsset(path.normalize(candidate));
  const canonicalDir = canonicalizeForAsset(path.normalize(dir));
  if (
    !isPathInsideDir(canonicalCandidate, canonicalDir) &&
    canonicalCandidate !== canonicalDir
  ) {
    const remedy =
      opts?.remedy ??
      'Keep the path inside the allowed directory (paths are canonicalized with realpath).';
    const body =
      opts?.message ?? `Path "${candidate}" resolves outside the allowed directory.`;
    throw Object.assign(new Error(`E_SANDBOX: ${body} Remedy: ${remedy}`), {
      code: 'E_SANDBOX',
      remedy,
    });
  }
  return canonicalCandidate;
}

/**
 * M1 dangling-symlink save guard (single source of truth for all file-writing
 * paths: MCP save/export/new-document/history-revert/redo — primary plus
 * failure-restore writes — engine PDF copy-out and git sidecar writes).
 *
 * `canonicalizeForAsset` resolves symlinks via realpath, but a DANGLING
 * symlink (leaf points at a nonexistent target) throws in realpathSync, falls
 * back to the nearest existing ancestor, and validates the PARENT — while a
 * later `writeFile`/`copyFile` follows the link outside the jail. Live-target
 * symlinks are already refused by canonical containment; this closes the
 * nonexistent-target hole specifically.
 *
 * Refuses (E_SANDBOX) when ANY existing path component from the filesystem
 * root down to the leaf is a symlink (lstat, never follows — so dangling
 * links are seen). Missing components (not-yet-created save targets) are
 * skipped; the post-mkdir re-check covers races. Regular files and real dirs
 * pass untouched, so live targets stay writable.
 */
export function assertWriteTargetNoSymlink(resolvedTarget: string): void {
  const abs = path.resolve(resolvedTarget);
  let cur = abs;
  for (;;) {
    let st: fs.Stats;
    try {
      st = fs.lstatSync(cur);
    } catch {
      /* missing component — nothing to refuse yet */
      const parent = path.dirname(cur);
      if (parent === cur) return;
      cur = parent;
      continue;
    }
    if (st.isSymbolicLink()) {
      const remedy =
        'Remove the symlink at the output path and write inside the allowed directory (paths are canonicalized with realpath).';
      throw Object.assign(
        new Error(`E_SANDBOX: output path "${path.basename(abs)}" traverses a symlink at "${cur}" and was refused. Remedy: ${remedy}`),
        { code: 'E_SANDBOX', remedy }
      );
    }
    const parent = path.dirname(cur);
    if (parent === cur) return;
    cur = parent;
  }
}

/**
 * O_NOFOLLOW hardened file write (closes the lstat→write TOCTOU window).
 * Pre-checks `assertWriteTargetNoSymlink`, then opens with
 * `O_WRONLY|O_CREAT|O_TRUNC|O_NOFOLLOW` so a symlink swapped in between the
 * check and the open fails with ELOOP (mapped to E_SANDBOX) instead of
 * following outside the jail. fsyncs before close. Post-write lstat refuses
 * a symlink planted at the destination during the write.
 */
export async function writeFileNoFollow(
  target: string,
  data: Buffer | string,
  encoding: BufferEncoding = 'utf8'
): Promise<void> {
  assertWriteTargetNoSymlink(path.resolve(target));
  const flags =
    fs.constants.O_WRONLY | fs.constants.O_CREAT | fs.constants.O_TRUNC | fs.constants.O_NOFOLLOW;
  let fh: fsp.FileHandle | undefined;
  try {
    fh = await fsp.open(target, flags, 0o600);
  } catch (err: unknown) {
    if ((err as { code?: string } | undefined)?.code === 'ELOOP') {
      const remedy =
        'Remove the symlink at the output path and write inside the allowed directory (paths are canonicalized with realpath).';
      throw Object.assign(
        new Error(`E_SANDBOX: output path "${path.basename(target)}" is a symlink and was refused. Remedy: ${remedy}`),
        { code: 'E_SANDBOX', remedy }
      );
    }
    throw err;
  }
  try {
    if (typeof data === 'string') await fh.writeFile(data, encoding);
    else await fh.writeFile(data);
    await fh.sync();
  } finally {
    try {
      await fh.close();
    } catch {
      /* close must never mask the write verdict */
    }
  }
  assertWriteTargetNoSymlink(path.resolve(target));
}

/** Sync variant of `writeFileNoFollow` (engine PDF copy-out path). */
export function writeFileNoFollowSync(
  target: string,
  data: Buffer | string,
  encoding: BufferEncoding = 'utf8'
): void {
  assertWriteTargetNoSymlink(path.resolve(target));
  const flags =
    fs.constants.O_WRONLY | fs.constants.O_CREAT | fs.constants.O_TRUNC | fs.constants.O_NOFOLLOW;
  let fd: number | undefined;
  try {
    fd = fs.openSync(target, flags, 0o600);
  } catch (err: unknown) {
    if ((err as { code?: string } | undefined)?.code === 'ELOOP') {
      const remedy =
        'Remove the symlink at the output path and write inside the allowed directory (paths are canonicalized with realpath).';
      throw Object.assign(
        new Error(`E_SANDBOX: output path "${path.basename(target)}" is a symlink and was refused. Remedy: ${remedy}`),
        { code: 'E_SANDBOX', remedy }
      );
    }
    throw err;
  }
  try {
    if (typeof data === 'string') fs.writeFileSync(fd, data, encoding);
    else fs.writeFileSync(fd, data);
    fs.fsyncSync(fd);
  } finally {
    try {
      if (fd !== undefined) fs.closeSync(fd);
    } catch {
      /* close must never mask the write verdict */
    }
  }
  assertWriteTargetNoSymlink(path.resolve(target));
}
