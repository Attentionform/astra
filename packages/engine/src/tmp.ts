/**
 * Secure temporary-file hygiene for artifacts carrying client document bytes.
 *
 * - Base dir: `$ASTRA_TMPDIR` when set (created 0700), else `os.tmpdir()`.
 * - Override bases are lstat-checked (symlink refused) and ownership-checked
 *   before mkdir/chmod, and re-checked after mkdir before chmod.
 * - The `os.tmpdir()` fallback (which honors `$TMPDIR`) clears the same
 *   symlink/owner jail when the env injected the value; the system default
 *   itself stays trusted and is never chmodded.
 * - Files 0600, dirs 0700 (explicit chmod after create: umask defense + repair).
 * - Callers delete on success, keep + log identity hash on failure (forensics).
 * - Forensic retention is bounded (count + age + bytes, oldest-first).
 * - Unique names per call (pid + timestamp + random): parallel-run safe.
 */

import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';

export const ASTRA_TMPDIR_ENV = 'ASTRA_TMPDIR';

const TMP_SYMLINK_REMEDY =
  'Point $ASTRA_TMPDIR at a real directory (not a symlink), or unset it and retry.';
const TMP_OWNER_REMEDY =
  'Point $ASTRA_TMPDIR at a directory you own, or unset it and retry.';
const SYSTEM_TMP_SYMLINK_REMEDY =
  'Point $TMPDIR at a real directory (not a symlink), or unset it and retry.';
const SYSTEM_TMP_OWNER_REMEDY =
  'Point $TMPDIR at a directory you own, or unset it and retry.';

function tmpSandboxError(message: string, remedy: string): Error & { code: string; remedy: string } {
  return Object.assign(new Error(message), { code: 'E_SANDBOX', remedy });
}

/** Refuse a symlink at `dir` via lstat (never follows). Throws E_SANDBOX. */
function refuseSymlinkBase(
  dir: string,
  describe = '$ASTRA_TMPDIR override',
  remedy = TMP_SYMLINK_REMEDY
): void {
  let st: fs.Stats;
  try {
    st = fs.lstatSync(dir);
  } catch {
    return; // missing path: nothing to refuse yet; post-mkdir check covers races
  }
  if (st.isSymbolicLink()) {
    throw tmpSandboxError(
      `E_SANDBOX: ${describe} is a symlink and was refused.`,
      remedy
    );
  }
}

/**
 * Refuse a pre-existing base not owned by the caller. Best-effort: skipped
 * on platforms without `process.getuid` or when ownership is unreadable.
 */
function refuseForeignOwner(
  dir: string,
  describe = '$ASTRA_TMPDIR override',
  remedy = TMP_OWNER_REMEDY
): void {
  const getuid = (process as unknown as { getuid?: () => number }).getuid;
  if (typeof getuid !== 'function') return;
  let st: fs.Stats;
  try {
    st = fs.statSync(dir);
  } catch {
    return; // missing path: owned-by-creator after mkdir; post-mkdir check applies
  }
  let me: number;
  try {
    me = getuid.call(process);
  } catch {
    return;
  }
  if (typeof me !== 'number' || Number.isNaN(me)) return;
  if (st.uid !== me) {
    throw tmpSandboxError(
      `E_SANDBOX: ${describe} is not owned by the current user and was refused.`,
      remedy
    );
  }
}

/** Resolve the base dir for transient artifacts. Creates + repairs 0700 when overridden. */
export function resolveAstraTmpBase(): string {
  const override = (process.env[ASTRA_TMPDIR_ENV] ?? '').trim();
  if (override !== '') {
    const dir = path.resolve(override);
    // Pre-mkdir guards: refuse symlink / foreign-owned targets before touching them.
    refuseSymlinkBase(dir);
    refuseForeignOwner(dir);
    try {
      fs.mkdirSync(dir, { recursive: true });
    } catch {
      return os.tmpdir();
    }
    // Post-mkdir re-checks (planted-symlink race): never chmod a refused target.
    refuseSymlinkBase(dir);
    refuseForeignOwner(dir);
    try {
      fs.chmodSync(dir, 0o700);
    } catch {
      /* best-effort hardening */
    }
    return dir;
  }
  // H3: os.tmpdir() honors $TMPDIR ($TEMP/$TMP on some platforms) — an
  // attacker-influenced value clears the same symlink/owner jail as an
  // explicit override. Only a value the env actually injected is checked:
  // the compiled-in system default (no env override in effect) stays
  // trusted, so legit roots like root-owned /tmp keep working untouched.
  // The system base is never mkdir/chmodded here (unlike the override
  // above): it already exists and is not ours to repair.
  const fallback = os.tmpdir();
  const injected = [process.env.TMPDIR, process.env.TEMP, process.env.TMP]
    .filter((v): v is string => typeof v === 'string' && v.trim() !== '')
    .map((v) => path.resolve(v.trim()));
  if (injected.includes(path.resolve(fallback))) {
    const dir = path.resolve(fallback);
    refuseSymlinkBase(dir, '$TMPDIR fallback base', SYSTEM_TMP_SYMLINK_REMEDY);
    refuseForeignOwner(dir, '$TMPDIR fallback base', SYSTEM_TMP_OWNER_REMEDY);
  }
  return fallback;
}

/** mkdtemp under the Astra base with enforced 0700. */
export function secureMkdtemp(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(resolveAstraTmpBase(), prefix));
  try {
    fs.chmodSync(dir, 0o700);
  } catch {
    /* best-effort hardening */
  }
  return dir;
}

function uniqueTmpPath(prefix: string, ext: string): string {
  const base = resolveAstraTmpBase();
  const rand = Math.random().toString(36).slice(2, 10);
  return path.join(base, `${prefix}${process.pid}-${Date.now()}-${rand}${ext}`);
}

/** Create an empty 0600 tmp file and return its path. Retries on name collision. */
export function secureTmpFile(prefix: string, ext: string): string {
  for (let attempt = 0; attempt < 10; attempt++) {
    const p = uniqueTmpPath(prefix, ext);
    try {
      const fd = fs.openSync(p, 'wx', 0o600);
      fs.closeSync(fd);
    } catch (err: unknown) {
      if ((err as NodeJS.ErrnoException)?.code === 'EEXIST') continue;
      throw err;
    }
    try {
      fs.chmodSync(p, 0o600);
    } catch {
      /* best-effort hardening */
    }
    return p;
  }
  throw new Error(`cannot allocate tmp file (prefix ${prefix})`);
}

/** Create a 0600 tmp file carrying `data` and return its path. */
export function secureWriteTmpFile(prefix: string, ext: string, data: Buffer | string): string {
  const p = secureTmpFile(prefix, ext);
  fs.writeFileSync(p, data);
  try {
    fs.chmodSync(p, 0o600);
  } catch {
    /* best-effort hardening */
  }
  return p;
}

/** Enforce 0600 on an existing file (umask defense). */
export function secureFile(p: string): void {
  try {
    fs.chmodSync(p, 0o600);
  } catch {
    /* best-effort hardening */
  }
}

/** Enforce 0700 on an existing dir (umask defense + repair of pre-existing dirs). */
export function secureDir(p: string): void {
  try {
    fs.chmodSync(p, 0o700);
  } catch {
    /* best-effort hardening */
  }
}

/** Best-effort unlink; keeps + logs identity hash when removal fails (possible leak). */
export function removeFileBestEffort(p: string): void {
  try {
    if (!fs.existsSync(p)) return;
    fs.unlinkSync(p);
  } catch {
    logForensics(p, 'could not remove tmp file');
  }
}

/** Best-effort recursive rmdir; keeps + logs identity hash when removal fails (possible leak). */
export function removeDirBestEffort(p: string): void {
  try {
    fs.rmSync(p, { recursive: true, force: true });
  } catch {
    logForensics(p, 'could not remove tmp dir');
  }
}

/**
 * Log a retained artifact without ever printing its absolute path: the line
 * carries only the relative file name plus a short identity hash of the full
 * path, so failure logs stay free of machine-specific directories.
 */
export function logForensics(p: string, reason: string): void {
  try {
    const rawName = path.basename(p);
    const name = rawName !== '' ? rawName : 'artifact';
    let hash = 'unknown';
    try {
      hash = crypto.createHash('sha256').update(p).digest('hex').slice(0, 16);
    } catch {
      /* keep placeholder */
    }
    process.stderr.write(`[ASTra] tmp artifact kept for forensics: name="${name}" sha256="${hash}" (${reason}). Base honors $ASTRA_TMPDIR; remove manually after inspection.\n`);
  } catch {
    /* logging must never throw */
  }
}

/** Keep `p` for failure forensics and log its identity hash. */
export function keepForForensics(p: string, reason: string): void {
  logForensics(p, reason);
  try {
    pruneForensicArtifacts(undefined, p);
  } catch {
    /* pruning is best-effort and never fails the keep */
  }
}

/** Forensic retention bounds: count, age, and total bytes (oldest-first). */
export const FORENSIC_MAX_COUNT = 20;
export const FORENSIC_TTL_MS = 24 * 60 * 60 * 1000;
export const FORENSIC_MAX_BYTES = 50 * 1024 * 1024;
export const FORENSIC_LEDGER_NAME = '.astra-forensics.json';

function forensicLedgerPath(baseDir: string): string {
  return path.join(baseDir, FORENSIC_LEDGER_NAME);
}

function isInsideBase(candidate: string, baseDir: string): boolean {
  const normC = path.normalize(candidate);
  const normB = path.normalize(baseDir);
  return normC === normB || normC.startsWith(normB + path.sep);
}

function forensicEntrySize(target: string): number {
  let lst: fs.Stats;
  try {
    lst = fs.lstatSync(target);
  } catch {
    return 0;
  }
  if (lst.isSymbolicLink()) return 0;
  if (lst.isFile()) return lst.size;
  if (!lst.isDirectory()) return 0;
  let total = 0;
  let entries: string[];
  try {
    entries = fs.readdirSync(target);
  } catch {
    return 0;
  }
  for (const name of entries) {
    if (total > FORENSIC_MAX_BYTES) break;
    try {
      total += forensicEntrySize(path.join(target, name));
    } catch {
      /* best-effort */
    }
  }
  return total;
}

function readForensicLedger(ledgerPath: string): string[] {
  try {
    const raw = fs.readFileSync(ledgerPath, 'utf8');
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((e): e is string => typeof e === 'string' && e.length > 0);
  } catch {
    return [];
  }
}

function logForensicPrune(target: string, reason: string): void {
  try {
    const rawName = path.basename(target);
    const name = rawName !== '' ? rawName : 'artifact';
    let hash = 'unknown';
    try {
      hash = crypto.createHash('sha256').update(target).digest('hex').slice(0, 16);
    } catch {
      /* keep placeholder */
    }
    process.stderr.write(`[ASTra] tmp forensics prune: name="${name}" sha256="${hash}" (${reason}).\n`);
  } catch {
    /* logging must never throw */
  }
}

/**
 * Enforce forensic retention: drop ledger entries for missing/out-of-base
 * paths, delete entries older than the TTL, then delete oldest-first while
 * over the count or byte ceiling. Never throws. When `keepPath` is given it
 * is recorded before pruning so a burst keeps the newest.
 */
export function pruneForensicArtifacts(baseOverride?: string, keepPath?: string): string[] {
  let baseDir: string;
  try {
    baseDir = baseOverride ?? resolveAstraTmpBase();
  } catch {
    return [];
  }
  let ledgerPath: string;
  try {
    ledgerPath = forensicLedgerPath(baseDir);
  } catch {
    return [];
  }
  const now = Date.now();
  let ledger: string[] = [];
  try {
    ledger = readForensicLedger(ledgerPath);
  } catch {
    ledger = [];
  }
  if (keepPath) {
    try {
      const resolvedKeep = path.resolve(keepPath);
      if (isInsideBase(resolvedKeep, path.resolve(baseDir)) && !ledger.includes(resolvedKeep)) {
        ledger.push(resolvedKeep);
      } else if (isInsideBase(resolvedKeep, path.resolve(baseDir))) {
        const idx = ledger.indexOf(resolvedKeep);
        if (idx >= 0) ledger.splice(idx, 1);
        ledger.push(resolvedKeep);
      }
    } catch {
      /* best-effort */
    }
  }
  interface LiveEntry { target: string; mtimeMs: number; sizeBytes: number }
  const live: LiveEntry[] = [];
  for (const entry of ledger) {
    let resolved: string;
    try {
      resolved = path.resolve(entry);
    } catch {
      continue;
    }
    if (!isInsideBase(resolved, path.resolve(baseDir))) continue;
    let lst: fs.Stats;
    try {
      lst = fs.lstatSync(resolved);
    } catch {
      continue;
    }
    if (lst.isSymbolicLink()) continue;
    const mtimeMs = lst.mtimeMs;
    if (!Number.isFinite(mtimeMs)) continue;
    if (now - mtimeMs > FORENSIC_TTL_MS) {
      try {
        if (lst.isDirectory()) fs.rmSync(resolved, { recursive: true, force: true });
        else fs.unlinkSync(resolved);
      } catch {
        /* best-effort */
      }
      logForensicPrune(resolved, 'expired after retention window');
      continue;
    }
    live.push({ target: resolved, mtimeMs, sizeBytes: forensicEntrySize(resolved) });
  }
  live.sort((a, b) => a.mtimeMs - b.mtimeMs);
  let totalBytes = live.reduce((s, e) => s + e.sizeBytes, 0);
  const survivors: LiveEntry[] = [...live];
  while (survivors.length > FORENSIC_MAX_COUNT || totalBytes > FORENSIC_MAX_BYTES) {
    const oldest = survivors.shift();
    if (!oldest) break;
    totalBytes -= oldest.sizeBytes;
    try {
      let lst: fs.Stats;
      try {
        lst = fs.lstatSync(oldest.target);
      } catch {
        logForensicPrune(oldest.target, 'over retention ceiling');
        continue;
      }
      if (lst.isSymbolicLink()) {
        logForensicPrune(oldest.target, 'over retention ceiling');
        continue;
      }
      if (lst.isDirectory()) fs.rmSync(oldest.target, { recursive: true, force: true });
      else fs.unlinkSync(oldest.target);
    } catch {
      /* best-effort */
    }
    logForensicPrune(oldest.target, 'over retention ceiling');
  }
  try {
    fs.mkdirSync(baseDir, { recursive: true });
    const tmpLedger = `${ledgerPath}.tmp-${process.pid}-${Date.now()}`;
    try {
      fs.writeFileSync(tmpLedger, JSON.stringify(survivors.map((s) => s.target)), 'utf8');
      try {
        fs.renameSync(tmpLedger, ledgerPath);
      } catch {
        try {
          fs.unlinkSync(ledgerPath);
        } catch {
          /* ignore */
        }
        fs.renameSync(tmpLedger, ledgerPath);
      }
    } catch {
      try {
        fs.unlinkSync(tmpLedger);
      } catch {
        /* ignore */
      }
    }
  } catch {
    /* best-effort */
  }
  return survivors.map((s) => s.target);
}
