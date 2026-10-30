import fs from 'fs';
import path from 'path';
import { toError } from '../errors';
import { canonicalizeForAsset, isPathInsideDir } from '../sandbox/paths';

/**
 * Decoded-byte cap for batch `filePath` image reads. Matches the inline
 * `fileData` schema cap (20MB) so both input shapes refuse at the same size.
 */
export const IMAGE_FILE_MAX_BYTES = 20_000_000;

/**
 * Validates a configured `allowedDir` for batch image reads (engine side of
 * the compose asset jail). Returns the canonical directory. Fail-closed:
 * a missing/unusable jail refuses with E_SANDBOX/E_VALIDATION, never a
 * silent grant. Mirrors the compose `resolveAllowedDir` decisions; raises
 * the F42 envelope (`.code`/`.remedy`) instead of markdown-line errors.
 */
export function assertImageAllowedDir(allowedDir: string | undefined): string {
  if (typeof allowedDir !== 'string' || allowedDir.trim().length === 0) {
    throw toError(
      'E_SANDBOX',
      'image filePath needs an allowed directory but none is configured',
      'pass allowedDir pointing at the folder holding the image files ' +
        '(engine OperationEngineOptions or per-batch allowedDir; MCP execute_batch allowedDir; CLI batch --allowed-dir)'
    );
  }
  const trimmed = allowedDir.trim();
  if (trimmed.includes('\0') || /%00/i.test(trimmed) || trimmed.includes('\n') || trimmed.includes('\r')) {
    throw toError(
      'E_VALIDATION',
      'image allowed directory path is not usable',
      'pass a plain local directory path without null bytes or line breaks'
    );
  }
  let resolved = path.resolve(trimmed);
  let stat: fs.Stats;
  try {
    stat = fs.statSync(resolved);
  } catch (err: unknown) {
    const code = (err as { code?: string } | undefined)?.code;
    if (code !== 'ENOENT' && code !== 'ENOTDIR') throw err;
    throw toError(
      'E_VALIDATION',
      'image allowed directory does not exist',
      'pass an existing local directory that holds the image files'
    );
  }
  if (!stat.isDirectory()) {
    throw toError(
      'E_VALIDATION',
      'image allowed directory is not a directory',
      'pass an existing local directory that holds the image files'
    );
  }
  try {
    resolved = fs.realpathSync(resolved);
  } catch { /* keep lexical form */ }
  return resolved;
}

/**
 * Resolves a batch image `filePath` against the session `allowedDir` jail.
 * Refusal decisions mirror the compose `resolveAssetPath` verdicts (outside,
 * traversal, hidden, remote); codes follow the batch contract (outside →
 * E_SANDBOX, missing → E_IO, oversize → E_LIMIT, unusable → E_VALIDATION)
 * in the F42 envelope shape. Returns the canonical absolute path.
 */
export function resolveImageFilePath(allowedDir: string, filePath: string): string {
  const trimmed = filePath.trim();
  if (trimmed.includes('\0') || /%00/i.test(trimmed) || trimmed.includes('\n') || trimmed.includes('\r')) {
    throw toError(
      'E_VALIDATION',
      'image filePath is not usable',
      'use a plain relative file path inside the allowed directory'
    );
  }
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(trimmed)) {
    throw toError(
      'E_SANDBOX',
      'image filePath points outside the allowed directory',
      `keep image files inside the allowed directory (${allowedDir}) and use a relative path`
    );
  }
  const candidate = path.isAbsolute(trimmed) ? path.resolve(trimmed) : path.resolve(allowedDir, trimmed);
  const canonicalAllowed = canonicalizeForAsset(allowedDir);
  const canonicalCandidate = canonicalizeForAsset(candidate);
  if (!isPathInsideDir(canonicalCandidate, canonicalAllowed) && canonicalCandidate !== canonicalAllowed) {
    throw toError(
      'E_SANDBOX',
      'image filePath points outside the allowed directory',
      `keep image files inside the allowed directory (${canonicalAllowed}) and use a relative path`
    );
  }
  const rel = path.relative(canonicalAllowed, canonicalCandidate);
  const segs = rel.split(path.sep).filter((s) => s && s !== '.' && s !== '..');
  for (const seg of segs) {
    if (seg.startsWith('.')) {
      throw toError(
        'E_SANDBOX',
        'image filePath traverses a hidden path',
        'keep image files in a visible folder inside the allowed directory'
      );
    }
  }
  const base = path.basename(canonicalCandidate);
  if (base.startsWith('.')) {
    throw toError(
      'E_SANDBOX',
      'image filePath uses a hidden file name',
      'rename the file without a leading dot'
    );
  }
  return canonicalCandidate;
}

/**
 * Engine-side hardened read for batch image `filePath` inserts. Extracts the
 * fd-hardened core of the compose asset read (pre-open canonicalization,
 * `O_NOFOLLOW` open, `fstat`-on-fd sizing, post-read containment re-check)
 * with F42-envelope errors instead of markdown-line failures. Returns the
 * raw bytes; the caller decodes them through the shared image decode path
 * exactly as if `fileData` had been passed.
 */
export function readImageFileBytes(filePath: string, allowedDir: string | undefined): Buffer {
  if (typeof filePath !== 'string' || filePath.trim().length === 0) {
    throw toError(
      'E_VALIDATION',
      'image insert needs an image source but filePath is empty',
      'pass filePath pointing at a file inside allowedDir, or fileData with inline base64 image bytes'
    );
  }
  const jail = assertImageAllowedDir(allowedDir);
  const absPath = resolveImageFilePath(jail, filePath.trim());
  const base = path.basename(absPath);
  const canonicalAllowed = canonicalizeForAsset(jail);
  const pre = canonicalizeForAsset(absPath);
  if (pre !== canonicalAllowed && !isPathInsideDir(pre, canonicalAllowed)) {
    throw toError(
      'E_SANDBOX',
      'image filePath points outside the allowed directory',
      `keep image files inside the allowed directory (${canonicalAllowed}) and use a relative path`
    );
  }
  const nofollow = fs.constants.O_NOFOLLOW as number | undefined;
  const flags =
    typeof nofollow === 'number' && nofollow !== 0
      ? (fs.constants.O_RDONLY as number) | nofollow
      : 'r';
  let fd: number | undefined;
  try {
    fd = fs.openSync(pre, flags);
  } catch (err: unknown) {
    const code = (err as { code?: string } | undefined)?.code;
    if (code === 'ENOENT') {
      throw toError(
        'E_IO',
        `image file is missing (${base})`,
        'add the file inside the allowed directory or fix the reference path'
      );
    }
    if (code === 'ELOOP') {
      throw toError(
        'E_SANDBOX',
        `image filePath traverses a symlink (${base})`,
        'point the reference at a regular image file inside the allowed directory'
      );
    }
    throw toError(
      'E_IO',
      `image file cannot be read (${base})`,
      'check the file exists and is readable, then retry'
    );
  }
  try {
    const st = fs.fstatSync(fd as number);
    if (!st.isFile()) {
      throw toError(
        'E_IO',
        `image reference is not a file (${base})`,
        'point the reference at a regular image file'
      );
    }
    if (st.size <= 0) {
      throw toError(
        'E_VALIDATION',
        `image file is empty (${base})`,
        'use a non-empty image file (png, jpg, gif, webp, bmp, tiff, or svg)'
      );
    }
    if (st.size > IMAGE_FILE_MAX_BYTES) {
      throw toError(
        'E_LIMIT',
        `image file is too large (${base}, ${st.size} bytes > ${IMAGE_FILE_MAX_BYTES} cap)`,
        `use an image file under ${IMAGE_FILE_MAX_BYTES} bytes`
      );
    }
    const buf = fs.readFileSync(fd as number);
    if (buf.length > IMAGE_FILE_MAX_BYTES) {
      throw toError(
        'E_LIMIT',
        `image file is too large (${base}, ${buf.length} bytes > ${IMAGE_FILE_MAX_BYTES} cap)`,
        `use an image file under ${IMAGE_FILE_MAX_BYTES} bytes`
      );
    }
    let post = pre;
    try {
      post = fs.realpathSync(absPath);
    } catch {
      /* target vanished mid-read — pre-open verdict stands */
    }
    if (post !== canonicalAllowed && !isPathInsideDir(post, canonicalAllowed)) {
      throw toError(
        'E_SANDBOX',
        'image filePath points outside the allowed directory',
        `keep image files inside the allowed directory (${canonicalAllowed}) and use a relative path`
      );
    }
    return buf;
  } finally {
    try {
      if (fd !== undefined) fs.closeSync(fd);
    } catch {
      /* close must never mask the read verdict */
    }
  }
}
