import * as fsp from 'fs/promises';
import * as path from 'path';
import { AdmZip, clampZipEntryTimes, ZIP_CEILING_MAX_RATIO_ENV, ZIP_CEILING_MAX_UNCOMPRESSED_ENV } from '@attentionform/astra-engine';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

export async function readDocx(file: string): Promise<Buffer> {
  try {
    return await fsp.readFile(file);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`cannot read "${file}": ${msg}`);
  }
}

/**
 * R-44: read piped stdin fully as UTF-8 (for `batch <docx> -`).
 * Callers check `process.stdin.isTTY` first and reject a terminal with a
 * usage error, so this never blocks on interactive input. A closed pipe
 * with no data (`< /dev/null`) resolves to '' and the caller rejects that
 * as a usage error too — empty check, not a timeout, so a slow-but-live
 * pipe still completes while empty stdin fails fast with exit 2.
 */
export function readStdinText(): Promise<string> {
  return (async () => {
    const chunks: Buffer[] = [];
    for await (const chunk of process.stdin as unknown as AsyncIterable<string | Buffer>) {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk, 'utf8') : chunk);
    }
    return Buffer.concat(chunks).toString('utf8');
  })();
}

// ── R-8 zip-bomb ceiling tunables (env + flags) ─────────────────────────────
// Defaults stay 100:1 / 200 MiB (see zip-hardening.ts). Env
// (ASTRA_MAX_RATIO / ASTRA_MAX_UNCOMPRESSED_BYTES) is honored by every engine
// ceiling gate: `readDocxFileBuffer` pre-read gate, `DocxParser` central-
// directory gate, and the sanitizer gate (all resolve from env at check
// time); explicit `checkZipBombCeilings` opts still win per-call. Explicit
// CLI flags (`--max-ratio` / `--max-uncompressed-bytes` on
// inspect/validate/batch/visual, plus ops-file maxRatio/maxUncompressedBytes on
// batch) win for that invocation by scoping the env around the DocxParser
// load (flag wins over file wins over env). Other commands
// (search/stats/sanitize/export/create) honor the env vars but take no flags.
// No ceiling gate stays on a fixed constant.

export interface ZipCeilingOverrides {
  maxRatio?: number;
  maxUncompressedBytes?: number;
}

/** Validate a `--max-ratio <n>` value: positive finite number (cap guards typos). */
export function parseMaxRatioFlag(raw: string | undefined): number {
  if (raw === undefined || raw === '') {
    throw new Error('inspect/validate/batch/visual --max-ratio requires a value (positive number, e.g. --max-ratio 200)');
  }
  const n = Number(raw);
  if (!Number.isFinite(n) || n <= 0 || n > 1_000_000) {
    throw new Error(`--max-ratio must be a positive number 0..1000000 (got "${raw}")`);
  }
  return n;
}

/** Validate a `--max-uncompressed-bytes <n>` value: positive integer bytes. */
export function parseMaxUncompressedBytesFlag(raw: string | undefined): number {
  if (raw === undefined || raw === '') {
    throw new Error('inspect/validate/batch/visual --max-uncompressed-bytes requires a value in bytes (e.g. --max-uncompressed-bytes 314572800)');
  }
  const n = Number(raw);
  if (!Number.isFinite(n) || !Number.isInteger(n) || n <= 0) {
    throw new Error(`--max-uncompressed-bytes must be a positive integer byte count (got "${raw}")`);
  }
  return n;
}

// ── Macro/active-content disposition (refuse-default) ───────────────────────
// A file carrying macro parts (vbaProject.bin, oleObject/activeX bins with a
// macro rel, word/activeX/*, macro rels) refuses with E_VALIDATION naming the
// parts. Zero bytes change without an explicit choice: --allow-macros
// preserves them byte-identical, --strip-macros removes them. Both together
// is a usage error (exit 2).
export interface MacroFlagChoice {
  allowMacros: boolean;
  stripMacros: boolean;
}

export function macroOptsFromFlags(f: MacroFlagChoice): { allowMacros: true } | { stripMacros: true } | undefined {
  if (f.allowMacros) return { allowMacros: true as const };
  if (f.stripMacros) return { stripMacros: true as const };
  return undefined;
}

export function macroFlagsUsageError(): string {
  return '--allow-macros and --strip-macros are mutually exclusive. Remedy: pass exactly one of --allow-macros (preserve macro parts) or --strip-macros (remove macro parts)';
}

/**
 * Scope ceiling overrides around a DocxParser load. The engine resolves
 * ceilings from env at check time (pre-read + central-directory gates), so
 * temporarily setting the env is the non-invasive plumb for CLI flags.
 * Restores prior values (including unset) afterwards.
 */
export async function withZipCeilingEnv<T>(
  ceilings: ZipCeilingOverrides,
  fn: () => T | Promise<T>
): Promise<T> {
  const prevRatio = process.env[ZIP_CEILING_MAX_RATIO_ENV];
  const prevBytes = process.env[ZIP_CEILING_MAX_UNCOMPRESSED_ENV];
  try {
    if (ceilings.maxRatio !== undefined) process.env[ZIP_CEILING_MAX_RATIO_ENV] = String(ceilings.maxRatio);
    if (ceilings.maxUncompressedBytes !== undefined) {
      process.env[ZIP_CEILING_MAX_UNCOMPRESSED_ENV] = String(ceilings.maxUncompressedBytes);
    }
    return await fn();
  } finally {
    if (prevRatio === undefined) delete process.env[ZIP_CEILING_MAX_RATIO_ENV];
    else process.env[ZIP_CEILING_MAX_RATIO_ENV] = prevRatio;
    if (prevBytes === undefined) delete process.env[ZIP_CEILING_MAX_UNCOMPRESSED_ENV];
    else process.env[ZIP_CEILING_MAX_UNCOMPRESSED_ENV] = prevBytes;
  }
}

/**
 * Reproducible scoping: explicit --no-reproducible sets ASTRA_NO_REPRODUCIBLE=1
 * around engine + save so revision/comment/cover dates see wall-clock;
 * explicit --reproducible clears the opt-out and pins ASTRA_REPRODUCIBLE=1;
 * undefined leaves env untouched (default deterministic unless env opts out).
 */
export async function withReproducibleEnv<T>(
  reproducible: boolean | undefined,
  fn: () => T | Promise<T>
): Promise<T> {
  const prevNo = process.env['ASTRA_NO_REPRODUCIBLE'];
  const prevRep = process.env['ASTRA_REPRODUCIBLE'];
  try {
    if (reproducible === false) {
      process.env['ASTRA_NO_REPRODUCIBLE'] = '1';
      if (prevRep !== undefined) delete process.env['ASTRA_REPRODUCIBLE'];
    } else if (reproducible === true) {
      if (prevNo !== undefined) delete process.env['ASTRA_NO_REPRODUCIBLE'];
      process.env['ASTRA_REPRODUCIBLE'] = '1';
    }
    return await fn();
  } finally {
    if (prevNo === undefined) delete process.env['ASTRA_NO_REPRODUCIBLE'];
    else process.env['ASTRA_NO_REPRODUCIBLE'] = prevNo;
    if (prevRep === undefined) delete process.env['ASTRA_REPRODUCIBLE'];
    else process.env['ASTRA_REPRODUCIBLE'] = prevRep;
  }
}

/**
 * Author scoping: --author <name> sets ASTRA_AUTHOR around the engine run, so the document's
 * default author (used by every tracked change and comment that does not name one) is the
 * caller's. undefined leaves the environment untouched.
 */
export async function withAuthorEnv<T>(author: string | undefined, fn: () => T | Promise<T>): Promise<T> {
  const prev = process.env['ASTRA_AUTHOR'];
  try {
    if (author !== undefined) process.env['ASTRA_AUTHOR'] = author;
    return await fn();
  } finally {
    if (prev === undefined) delete process.env['ASTRA_AUTHOR'];
    else process.env['ASTRA_AUTHOR'] = prev;
  }
}

// ── Shared DX helpers ───────────────────────────────────────────────────────

// Document census counters live in @attentionform/astra-engine (single source of truth
// shared with MCP `get_stats`). Re-exported here so existing CLI importers
// (`stats-cmd`, `inspect-cmd`) keep working unchanged.
export { computeDocStats, tableCellTextSpaced } from '@attentionform/astra-engine';
export type { DocStats } from '@attentionform/astra-engine';

export function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function buildSnippet(text: string, query: string, radius = 60, maxLen = 160): string {
  const flat = text.replace(/\s+/g, ' ').trim();
  if (flat.length <= maxLen) return flat;
  const idx = flat.toLowerCase().indexOf(query.toLowerCase());
  if (idx < 0) return `${flat.slice(0, maxLen - 3)}...`;
  const start = Math.max(0, idx - radius);
  const end = Math.min(flat.length, idx + query.length + radius);
  const prefix = start > 0 ? '...' : '';
  const suffix = end < flat.length ? '...' : '';
  return `${prefix}${flat.slice(start, end).trim()}${suffix}`;
}

/**
 * Reset the document body to a single empty paragraph, keeping styles,
 * numbering, settings and all package props byte-identical. CLI-level only:
 * the engine schema has no clear op, so `batch` expands `{action:'clearDoc'}`
 * through this function before the remaining ops run.
 */
export function clearDocumentBody(buf: Buffer): Buffer {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry('word/document.xml');
  if (!entry) throw new Error('E_VALIDATION: package has no word/document.xml. Remedy: pick a valid .docx and retry');
  let doc;
  try {
    doc = new DOMParser().parseFromString(entry.getData().toString('utf8'), 'text/xml');
  } catch (err: unknown) {
    throw new Error(
      `E_VALIDATION: document.xml is malformed (${err instanceof Error ? err.message : String(err)}). Remedy: repair the document in Word and retry`
    );
  }
  const bodies = doc.getElementsByTagName('w:body');
  if (bodies.length === 0) throw new Error('E_VALIDATION: document.xml has no w:body. Remedy: pick a valid .docx and retry');
  const body = bodies[0];
  const keepSectPr: Element[] = [];
  const remove: Element[] = [];
  for (let i = 0; i < body.childNodes.length; i++) {
    const child = body.childNodes[i] as unknown as Element;
    if (!child || child.nodeType !== 1) continue;
    if (child.tagName === 'w:sectPr') keepSectPr.push(child);
    else remove.push(child);
  }
  for (const el of remove) body.removeChild(el);
  const empty = doc.createElement('w:p');
  body.insertBefore(empty, keepSectPr[0] ?? null);
  zip.updateFile('word/document.xml', Buffer.from(new XMLSerializer().serializeToString(doc), 'utf8'));
  try {
    clampZipEntryTimes(zip);
  } catch {
    /* deterministic stamp best effort; the engine save re-clamps anyway */
  }
  return zip.toBuffer();
}

export async function writeDocx(file: string, buf: Buffer): Promise<void> {
  try {
    const abs = path.resolve(file);
    await fsp.mkdir(path.dirname(abs), { recursive: true });
    // Atomic write: temp file in the TARGET dir + rename (same filesystem,
    // POSIX-atomic). A crash mid-write never truncates the original to 0 bytes.
    const tmp = `${abs}.tmp.${process.pid}.${Date.now()}.${Math.floor(Math.random() * 1e6)}`;
    try {
      await fsp.writeFile(tmp, buf);
      await fsp.rename(tmp, abs);
    } catch (err: unknown) {
      try {
        await fsp.unlink(tmp);
      } catch {
        /* best-effort temp cleanup */
      }
      throw err;
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`cannot write "${file}": ${msg}`);
  }
}
