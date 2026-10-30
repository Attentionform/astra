import fs from 'fs';
import { resolveMaxUncompressedBytes } from '../zip-hardening';

/**
 * L-13: sandboxed file read for the DocxParser string-path overload.
 * Gates (fail-closed, in order):
 * 1. non-empty string, no null bytes (sandbox validation);
 * 2. stat must succeed or the exact legacy `File not found: <path>` is thrown;
 * 3. directories throw E_VALIDATION (never a raw EISDIR), non-regular files
 *    (devices, sockets, fifos) throw E_VALIDATION;
 * 4. size pre-check against the zip-bomb uncompressed ceiling throws E_LIMIT
 *    before any bytes are allocated.
 * No cwd confinement by design: caller-supplied absolute/relative docx paths
 * are allowed; the sandbox is type + regular-file + size, not location.
 */
export function readDocxFileBuffer(filePath: string): Buffer {
  if (typeof filePath !== 'string' || filePath.length === 0) {
    throw new Error(
      'E_VALIDATION: docx path must be a non-empty string. Remedy: pass a valid .docx file path or a Buffer.'
    );
  }
  if (filePath.includes('\0')) {
    throw new Error(
      'E_VALIDATION: docx path contains a null byte. Remedy: pass a valid .docx file path or a Buffer.'
    );
  }
  let stat: import('fs').Stats;
  try {
    stat = fs.statSync(filePath);
  } catch {
    throw new Error('File not found: ' + filePath);
  }
  if (stat.isDirectory()) {
    throw new Error(
      `E_VALIDATION: path is a directory, not a .docx file: ${filePath}. Remedy: pass a valid .docx file path.`
    );
  }
  if (!stat.isFile()) {
    throw new Error(
      `E_VALIDATION: not a regular file: ${filePath}. Remedy: pass a valid .docx file path.`
    );
  }
  const preReadCeiling = resolveMaxUncompressedBytes();
  if (stat.size > preReadCeiling) {
    throw new Error(
      `[E_LIMIT] docx file ${stat.size} bytes exceeds pre-read ceiling ${preReadCeiling} bytes. ` +
        `Remedy: raise ASTRA_MAX_UNCOMPRESSED_BYTES / --max-uncompressed-bytes above ${stat.size} (current ceiling ${preReadCeiling}) or use a smaller document`
    );
  }
  try {
    return fs.readFileSync(filePath);
  } catch (err: unknown) {
    const code = (err as { code?: unknown } | null | undefined)?.code;
    if (code === 'ENOENT') throw new Error('File not found: ' + filePath);
    if (code === 'EISDIR') {
      throw new Error(
        `E_VALIDATION: path is a directory, not a .docx file: ${filePath}. Remedy: pass a valid .docx file path.`
      );
    }
    throw err instanceof Error ? err : new Error(String(err));
  }
}

export function getMimeType(fileName: string): string {
  const ext = (fileName.split('.').pop() || '').toLowerCase();
  switch (ext) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    case 'bmp':
      return 'image/bmp';
    case 'svg':
      return 'image/svg+xml';
    case 'png':
    default:
      return 'image/png';
  }
}
