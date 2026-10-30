/**
 * Zip determinism + extraction-ceiling helpers for the engine package.
 *
 * Scope: package serialization (`DocxParser.rezip`) and package loading
 * (`new DocxParser(buffer)`). No other engine area depends on this module.
 *
 * 1) Determinism: adm-zip stamps every new entry header with the current
 * wall-clock time, so two identical saves produce different bytes. Clamping
 * every entry header to a single fixed epoch immediately before `toBuffer()`
 * makes identical content yield identical bytes. The epoch honors
 * `SOURCE_DATE_EPOCH` (seconds since unix epoch) when set to a usable value
 * and falls back to 1980-01-01T00:00:00Z (the earliest date the DOS time
 * field can represent) otherwise; pre-1980 values clamp up to the fallback.
 *
 * 2) Bomb ceiling: central-directory `size` / `compressedSize` headers are
 * inspected BEFORE any entry bytes are inflated, so an over-limit archive
 * fails closed without allocating the expanded payload. Two ceilings are
 * enforced: total declared uncompressed bytes (default 200 MiB) and
 * compression ratio (default 100:1) per entry plus overall (declared total
 * vs on-wire bytes, which an attacker cannot understate).
 *
 * Both ceilings are tunable without code changes: `ASTRA_MAX_RATIO` and
 * `ASTRA_MAX_UNCOMPRESSED_BYTES` env vars raise/lower the defaults
 * (`checkZipBombCeilings` explicit opts still win over env). Invalid,
 * non-positive, or non-finite env values fall back to the defaults
 * (fail-closed, never fail-open).
 * Honored everywhere the engine checks ceilings: `readDocxFileBuffer`
 * pre-read gate (via `resolveMaxUncompressedBytes`), `new
 * DocxParser(buffer)` central-directory gate, and `sanitizeDocxBuffer` gate
 * (all resolve from env at check time). The CLI `inspect`/`validate`/`batch`
 * commands expose the same knobs as `--max-ratio` /
 * `--max-uncompressed-bytes` flags plus ops-file `maxRatio` /
 * `maxUncompressedBytes` (flag wins over file wins over env for that
 * invocation); other CLI commands (`search`/`stats`/`sanitize`/`export`/
 * `create`) honor the env vars but take no flags. No ceiling gate stays on
 * a fixed constant: the only untunable threshold is the 1 KiB
 * `ZIP_BOMB_RATIO_MIN_ENTRY_BYTES` noise floor, which skips ratio checks
 * for tiny entries and is intentionally not a ceiling.
 */

export const ZIP_FIXED_EPOCH_MS = Date.UTC(1980, 0, 1, 0, 0, 0);

/** Env var selecting deterministic (byte-identical) output. Off unless set. */
export const REPRODUCIBLE_ENV = 'ASTRA_REPRODUCIBLE';

/** Env var forcing wall-clock output even when another switch asks for pinning. */
export const NO_REPRODUCIBLE_ENV = 'ASTRA_NO_REPRODUCIBLE';

function envFlag(raw: unknown): boolean | undefined {
  if (raw === undefined || raw === null || String(raw).trim() === '') return undefined;
  const v = String(raw).trim().toLowerCase();
  if (v === '1' || v === 'true' || v === 'yes' || v === 'on') return true;
  if (v === '0' || v === 'false' || v === 'no' || v === 'off') return false;
  return undefined;
}

/**
 * Deterministic output is opt-in: dates, ids and zip timestamps are real unless
 * asked otherwise. Pinning turns on with ASTRA_REPRODUCIBLE=1/true/yes/on or
 * with a valid SOURCE_DATE_EPOCH (the reproducible-builds convention). An
 * explicit ASTRA_NO_REPRODUCIBLE=1 or ASTRA_REPRODUCIBLE=0 always wins.
 */
export function isReproducibleEnv(env: NodeJS.ProcessEnv = process.env): boolean {
  try {
    if (envFlag(env?.[NO_REPRODUCIBLE_ENV]) === true) return false;
    const explicit = envFlag(env?.[REPRODUCIBLE_ENV]);
    if (explicit !== undefined) return explicit;
    const epoch = env?.SOURCE_DATE_EPOCH;
    if (epoch !== undefined && epoch !== null && String(epoch).trim() !== '') {
      const secs = Number(String(epoch).trim());
      return Number.isFinite(secs) && secs >= 0;
    }
    return false;
  } catch {
    return false;
  }
}

/**
 * Resolve reproducible mode with an explicit tri-state override.
 * explicit true forces true; explicit false forces false (opt-out);
 * undefined falls through to env (default false: real dates and ids).
 */
export function resolveReproducibleMode(
  explicit?: boolean,
  env: NodeJS.ProcessEnv = process.env
): boolean {
  if (explicit === true) return true;
  if (explicit === false) return false;
  return isReproducibleEnv(env);
}

/** Fail-closed ceiling for total declared uncompressed bytes (200 MiB). */
export const ZIP_BOMB_MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024;

/** Fail-closed ceiling for compression ratio (100:1). */
export const ZIP_BOMB_MAX_RATIO = 100;

/** Entries smaller than this skip ratio checks (avoids tiny-file noise). */
export const ZIP_BOMB_RATIO_MIN_ENTRY_BYTES = 1024;

/** Env var overriding the compression-ratio ceiling (e.g. `ASTRA_MAX_RATIO=200`). */
export const ZIP_CEILING_MAX_RATIO_ENV = 'ASTRA_MAX_RATIO';

/** Env var overriding the total-uncompressed-bytes ceiling (e.g. `ASTRA_MAX_UNCOMPRESSED_BYTES=314572800`). */
export const ZIP_CEILING_MAX_UNCOMPRESSED_ENV = 'ASTRA_MAX_UNCOMPRESSED_BYTES';

/**
 * Resolve the effective compression-ratio ceiling. Overridden via env
 * `ASTRA_MAX_RATIO`; invalid/empty/non-positive/non-finite values fall back
 * to `ZIP_BOMB_MAX_RATIO` (fail-closed default 100:1, never fail-open).
 */
export function resolveMaxRatio(env: NodeJS.ProcessEnv = process.env): number {
  try {
    const raw = env?.[ZIP_CEILING_MAX_RATIO_ENV];
    if (raw === undefined || raw === null || String(raw).trim() === '') {
      return ZIP_BOMB_MAX_RATIO;
    }
    const n = Number(String(raw).trim());
    if (!Number.isFinite(n) || n <= 0) return ZIP_BOMB_MAX_RATIO;
    return n;
  } catch {
    return ZIP_BOMB_MAX_RATIO;
  }
}

/**
 * Resolve the effective total-uncompressed-bytes ceiling. Overridden via env
 * `ASTRA_MAX_UNCOMPRESSED_BYTES` (plain bytes); invalid/empty/non-positive/
 * non-finite values fall back to `ZIP_BOMB_MAX_UNCOMPRESSED_BYTES`
 * (fail-closed default 200 MiB, never fail-open). Fractional values floor.
 */
export function resolveMaxUncompressedBytes(env: NodeJS.ProcessEnv = process.env): number {
  try {
    const raw = env?.[ZIP_CEILING_MAX_UNCOMPRESSED_ENV];
    if (raw === undefined || raw === null || String(raw).trim() === '') {
      return ZIP_BOMB_MAX_UNCOMPRESSED_BYTES;
    }
    const n = Number(String(raw).trim());
    if (!Number.isFinite(n) || n <= 0) return ZIP_BOMB_MAX_UNCOMPRESSED_BYTES;
    return Math.floor(n);
  } catch {
    return ZIP_BOMB_MAX_UNCOMPRESSED_BYTES;
  }
}

/** Resolve both ceilings at once (env-aware; explicit opts still win per-call). */
export function resolveZipBombCeilings(env: NodeJS.ProcessEnv = process.env): {
  maxRatio: number;
  maxUncompressedBytes: number;
} {
  return { maxRatio: resolveMaxRatio(env), maxUncompressedBytes: resolveMaxUncompressedBytes(env) };
}

export interface ZipEntryHeaderLike {
  size: number;
  compressedSize: number;
  time?: Date;
  timeval?: number;
}

export interface ZipEntryLike {
  entryName: string;
  isDirectory?: unknown;
  header: ZipEntryHeaderLike;
}

export interface ZipLike {
  getEntries(): ZipEntryLike[];
}

export interface ZipBombCeilingOptions {
  maxUncompressedBytes?: number;
  maxRatio?: number;
  /** On-wire compressed length (e.g. the source Buffer length). */
  sourceCompressedBytes?: number;
}

export interface ZipBombCeilingSummary {
  entryCount: number;
  totalUncompressedBytes: number;
  maxEntryRatio: number;
  overallRatio: number | null;
}

/**
 * Convert Date to 32-bit MS-DOS format in UTC.
 */
export function toDosTimeUtc(d: Date): number {
  const year = d.getUTCFullYear();
  if (year < 1980) return ((1 << 21) | (1 << 16)) >>> 0; // 1980-01-01 00:00:00 DOS
  const date = (((year - 1980) & 0x7f) << 9) | ((d.getUTCMonth() + 1) << 5) | d.getUTCDate();
  const time = (d.getUTCHours() << 11) | (d.getUTCMinutes() << 5) | (d.getUTCSeconds() >> 1);
  return (((date << 16) | time) >>> 0);
}

/**
 * Resolve the deterministic stamp for zip entries. Honors SOURCE_DATE_EPOCH
 * (seconds) when it denotes a representable date, else 1980-01-01Z.
 */
export function resolveDeterministicZipDate(
  env: NodeJS.ProcessEnv = process.env
): Date {
  const fallback = new Date(ZIP_FIXED_EPOCH_MS);
  try {
    const raw = env?.SOURCE_DATE_EPOCH;
    if (raw === undefined || raw === null || String(raw).trim() === '') {
      return fallback;
    }
    const secs = Number(String(raw).trim());
    if (!Number.isFinite(secs) || secs < 0) return fallback;
    const candidate = new Date(Math.floor(secs) * 1000);
    if (Number.isNaN(candidate.getTime())) return fallback;
    if (candidate.getTime() < ZIP_FIXED_EPOCH_MS) return fallback;
    return candidate;
  } catch {
    return fallback;
  }
}

/**
 * Clamp every entry header time to a single stamp (default: resolved epoch).
 * Returns the stamp applied. Best-effort per entry; never throws, so a clamp
 * problem can never block a save whose XML already validated.
 */
export function clampZipEntryTimes(zip: ZipLike, at?: Date): Date {
  const stamp = at ? new Date(at.getTime()) : resolveDeterministicZipDate();
  try {
    const entries = zip.getEntries();
    if (!Array.isArray(entries)) return stamp;
    const dosTime = toDosTimeUtc(stamp);
    for (const entry of entries) {
      try {
        if (!entry || !entry.header) continue;
        entry.header.timeval = dosTime;
      } catch {
        /* per-entry best-effort */
      }
    }
  } catch {
    /* never block save on clamp */
  }
  return stamp;
}

/** Minimal structural surface needed to rewrite one package part. */
export interface CorePropsZipLike extends ZipLike {
  getEntry(name: string): { getData(): Buffer } | null | undefined;
  updateFile(name: string, content: Buffer): void;
}

/**
 * Freeze `dcterms:modified` in `docProps/core.xml` to a single stamp
 * (default: resolved epoch), leaving source `dcterms:created` untouched.
 * Only rewrites the inner text of tags already present; element attributes (e.g. `xsi:type`) and absent tags
 * are left untouched, so a part without stamps stays byte-stable on its own.
 * Best-effort; never throws. Returns true when the part was rewritten.
 */
export function clampCorePropertiesDates(zip: CorePropsZipLike, at?: Date): boolean {
  const stamp = at ? new Date(at.getTime()) : resolveDeterministicZipDate();
  try {
    const entry = zip.getEntry('docProps/core.xml');
    if (!entry) return false;
    const xml = entry.getData().toString('utf8');
    const iso = stamp.toISOString();
    let out = xml;
    out = out.replace(
      /(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified\s*>)/g,
      (_m, open: string, close: string) => `${open}${iso}${close}`
    );
    if (out === xml) return false;
    zip.updateFile('docProps/core.xml', Buffer.from(out, 'utf8'));
    return true;
  } catch {
    return false;
  }
}

/**
 * Inspect central-directory headers BEFORE inflation and throw fail-closed
 * (`[E_LIMIT]`) when either ceiling is exceeded. Never calls `getData()`,
 * so no entry payload is inflated by this check.
 */
export function checkZipBombCeilings(
  zip: ZipLike,
  opts?: ZipBombCeilingOptions
): ZipBombCeilingSummary {
  const maxUncompressed =
    opts?.maxUncompressedBytes ?? resolveMaxUncompressedBytes();
  const maxRatio = opts?.maxRatio ?? resolveMaxRatio();
  const entries = zip.getEntries() ?? [];
  if (!Array.isArray(entries)) {
    throw new Error(
      '[E_LIMIT] zip bomb ceiling: unreadable central directory. ' +
        'Remedy: re-export the file from Word (or pick a valid .docx) and retry'
    );
  }
  let total = 0;
  let maxEntryRatio = 0;
  for (const entry of entries) {
    const name =
      entry && typeof entry.entryName === 'string' ? entry.entryName : '?';
    if (!entry || !entry.header) continue;
    if (entry.isDirectory === true) continue;
    const rawSize = entry.header.size;
    const rawCompressed = entry.header.compressedSize;
    if (
      typeof rawSize !== 'number' ||
      !Number.isFinite(rawSize) ||
      rawSize < 0 ||
      typeof rawCompressed !== 'number' ||
      !Number.isFinite(rawCompressed) ||
      rawCompressed < 0
    ) {
      throw new Error(
        `[E_LIMIT] zip bomb ceiling: entry '${name}' has invalid declared size. ` +
          `Remedy: re-export the file from Word (or pick a valid .docx) and retry`
      );
    }
    const size = rawSize;
    const compressed = rawCompressed;
    if (size === 0) {
      if (compressed > 0) {
        throw new Error(
          `[E_LIMIT] zip bomb ceiling: entry '${name}' declared size 0 with non-zero compressed payload (${compressed} bytes). ` +
            `Remedy: re-export the file from Word (or pick a valid .docx) and retry`
        );
      }
      continue;
    }
    total += size;
    if (total > maxUncompressed) {
      throw new Error(
        `[E_LIMIT] zip bomb ceiling: total uncompressed ${total} bytes exceeds ${maxUncompressed} bytes. ` +
          `Remedy: raise ASTRA_MAX_UNCOMPRESSED_BYTES / --max-uncompressed-bytes above ${total} (current ceiling ${maxUncompressed}) or use a smaller document`
      );
    }
    if (size >= ZIP_BOMB_RATIO_MIN_ENTRY_BYTES) {
      const ratio = size / Math.max(1, compressed);
      if (ratio > maxEntryRatio) maxEntryRatio = ratio;
      if (ratio > maxRatio) {
        throw new Error(
          `[E_LIMIT] zip bomb ceiling: entry '${name}' ratio ${ratio.toFixed(1)}:1 exceeds ${maxRatio}:1 (size ${size}, compressed ${compressed}). ` +
            `Remedy: raise ASTRA_MAX_RATIO / --max-ratio above ${ratio.toFixed(1)} (current ceiling ${maxRatio}) or use a less compressible document`
        );
      }
    }
  }
  let overallRatio: number | null = null;
  const wire = opts?.sourceCompressedBytes;
  if (typeof wire === 'number' && Number.isFinite(wire) && wire > 0) {
    overallRatio = total / wire;
    if (overallRatio > maxRatio) {
      throw new Error(
        `[E_LIMIT] zip bomb ceiling: overall ratio ${overallRatio.toFixed(1)}:1 exceeds ${maxRatio}:1 (total ${total}, wire ${wire}). ` +
          `Remedy: raise ASTRA_MAX_RATIO / --max-ratio above ${overallRatio.toFixed(1)} (current ceiling ${maxRatio}) or use a less compressible document`
      );
    }
  }
  return {
    entryCount: entries.length,
    totalUncompressedBytes: total,
    maxEntryRatio,
    overallRatio,
  };
}

/** Alias kept for call-site readability at load/extract gates. */
export const enforceZipBombCeilings = checkZipBombCeilings;
