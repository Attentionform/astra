/**
 * UX-C-recovery backend (UX-01 P1-4 crash banner + P2-5 orphan GC).
 *
 * REGISTER CORRECTIONS (read first):
 *  - P1-4 crash banner is scoped to LAST-COMMITTED Drafts snapshots: git
 *    history length > 1 (beyond "Initial state") AND head newer than the
 *    clean-exit marker AND full resolved path not in the marker's discarded
 *    set AND not expired. Unflushed renderer refs / chat / recents are NOT
 *    recoverable. No SQLite exists in package.json — file marker only.
 *  - P2-5 orphan GC is liveness/age-gated: deletes Drafts-dir orphans older
 *    than ORPHAN_GC_MAX_AGE_MS, or oldest-first while the Drafts total
 *    exceeds ORPHAN_GC_MAX_TOTAL_BYTES. NEVER deletes the current-session
 *    recoverable path, NEVER touches titled docs (only operates inside the
 *    Drafts dir), NEVER deletes young live drafts except under size pressure
 *    (oldest-first). Crash-mid-write `*.tmp-<pid>-*` litter older than the
 *    age gate is owned here too (see atomicWriteFile in document-manager.ts).
 *  - Single-instance lock (electron/main.ts) stops MULTI-PROCESS corruption
 *    only — it does not prevent two-window races in one process (file-guard
 *    lease is future work, documented there).
 *  - No 50MB hard cap anywhere (would break legit media docs with no chunk
 *    fallback — rejected by the register).
 *
 * UX-C reviewer follow-up item 2 DOCUMENTED POLICY (discarded matching):
 *  - Discard entries are RESOLVED FULL PATHS (path.resolve), never bare
 *    basenames. A discard in one Drafts dir must not suppress an unrelated
 *    file that merely shares its basename, and a reused basename after GC
 *    must become banner-eligible again.
 *  - Legacy markers (pre-fix) stored bare basenames. They are still honored
 *    as basename fallbacks so deleted intent is never resurrected by the
 *    upgrade itself; new discards always store full paths.
 *  - Comparison is PLATFORM-AWARE: case-insensitive on win32, case-sensitive
 *    on Linux/POSIX (a.docx vs A.docx are distinct on Linux). Helpers take an
 *    explicit platform param for unit testing.
 *  - GC OWNS ITS LITTER: sweepOrphanDrafts prunes the discarded entries for
 *    every file it deletes, so a reused basename is not permanently
 *    suppressed.
 *
 * GIT DISCIPLINE: this module imports ONLY read paths from ./git
 * (getHistory, getRepoDir). It never commits, inits, or rewinds history.
 */
import fsp from 'fs/promises';
import path from 'path';
import { getHistory, getRepoDir } from './git';
import { isPathInsideDir } from './sandbox/paths';

// ── Mirrored constants (keep in sync with renderer/utils/recovery.ts) ───────
// Backend must not import renderer (layering) and renderer must not import fs;
// tests pin equality between the two copies.
export const RECOVERY_BANNER_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
export const ORPHAN_GC_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
export const ORPHAN_GC_MAX_TOTAL_BYTES = 100 * 1024 * 1024;
export const CLEAN_EXIT_FILENAME = '.astra-clean-exit.json';
/** Session filename at the userData root (sibling of Drafts/). Main-owned. */
export const SESSION_FILENAME = 'astra-session.json';

export interface CleanExitMarker {
  cleanExitAt: number | null;
  discarded: string[];
  version: 1;
}

export interface RecoveryCandidate {
  path: string;
  basename: string;
  commitCount: number;
  headMessage: string | null;
  headTimeMs: number | null;
  mtimeMs: number;
  sizeBytes: number;
}

export interface OrphanEntry {
  path: string;
  sizeBytes: number;
  mtimeMs: number;
}

// ── Pure banner predicate (mirrors renderer/utils/recovery.ts) ──────────────

export interface BannerPredicateInput {
  commitCount: number;
  headTimeMs: number | null;
  cleanExitAtMs: number | null;
  isDiscarded: boolean;
  nowMs: number;
  maxAgeMs?: number;
}

export function shouldShowRecoveryBannerBackend(input: BannerPredicateInput): boolean {
  const maxAge = input.maxAgeMs ?? RECOVERY_BANNER_MAX_AGE_MS;
  if (input.isDiscarded) return false;
  if (!Number.isFinite(input.commitCount) || input.commitCount <= 1) return false;
  if (input.headTimeMs == null || !Number.isFinite(input.headTimeMs)) return false;
  if (!Number.isFinite(input.nowMs)) return false;
  if (input.cleanExitAtMs != null && Number.isFinite(input.cleanExitAtMs)) {
    if ((input.headTimeMs as number) <= (input.cleanExitAtMs as number)) return false;
  }
  if ((input.headTimeMs as number) < input.nowMs - maxAge) return false;
  return true;
}

// ── Paths ───────────────────────────────────────────────────────────────────

export function getCleanExitMarkerPath(draftsDir: string): string {
  return path.join(draftsDir, CLEAN_EXIT_FILENAME);
}

export function getSessionFilePath(userDataDir: string): string {
  return path.join(userDataDir, SESSION_FILENAME);
}

// Q0a: isPathInsideDir lives in ./sandbox/paths (imported above).

// ── Clean-exit marker IO (best-effort, never throws) ────────────────────────

export function emptyMarker(): CleanExitMarker {
  return { cleanExitAt: null, discarded: [], version: 1 };
}

/**
 * UX-C item 2: platform-aware path equality. Case-insensitive on win32,
 * case-sensitive on Linux/POSIX. Exported with an explicit platform param
 * so unit tests can pin both behaviors regardless of the host OS.
 */
export function areRecoveryPathsEqual(
  a: string,
  b: string,
  platform: string = process.platform
): boolean {
  const normA = path.normalize(a);
  const normB = path.normalize(b);
  if (platform === 'win32') {
    return normA.toLowerCase() === normB.toLowerCase();
  }
  return normA === normB;
}

/** True for legacy bare-basename entries (no separator, not absolute). */
export function isLegacyDiscardedBasenameEntry(entry: string): boolean {
  if (!entry) return false;
  if (path.isAbsolute(entry)) return false;
  if (entry.includes('/') || entry.includes('\\')) return false;
  return true;
}

export async function readCleanExitMarker(markerPath: string): Promise<CleanExitMarker> {
  try {
    const raw = await fsp.readFile(markerPath, 'utf8');
    const parsed = JSON.parse(raw) as Partial<CleanExitMarker>;
    let cleanExitAt: number | null = null;
    if (typeof parsed.cleanExitAt === 'number' && Number.isFinite(parsed.cleanExitAt)) {
      cleanExitAt = parsed.cleanExitAt;
    }
    // UX-C item 2: preserve full resolved paths (no basename collapsing).
    // Legacy bare basenames are kept verbatim as fallbacks. Dedupe exactly;
    // platform-aware dedupe happens at match time, not here.
    const discarded = Array.isArray(parsed.discarded)
      ? [...new Set(
          parsed.discarded
            .filter((d) => typeof d === 'string' && (d as string).trim() !== '')
            .map((d) => (d as string).trim())
        )]
      : [];
    return { cleanExitAt, discarded, version: 1 };
  } catch {
    return emptyMarker();
  }
}

async function writeMarkerAtomic(markerPath: string, marker: CleanExitMarker): Promise<void> {
  const dir = path.dirname(markerPath);
  await fsp.mkdir(dir, { recursive: true });
  const tmp = `${markerPath}.tmp-${process.pid}-${Date.now()}`;
  try {
    await fsp.writeFile(tmp, JSON.stringify(marker), 'utf8');
    try {
      await fsp.rename(tmp, markerPath);
    } catch {
      // Windows rename-onto-existing fallback (bounded, best-effort).
      try {
        await fsp.unlink(markerPath);
      } catch { /* ignore */ }
      await fsp.rename(tmp, markerPath);
    }
  } catch (err) {
    try {
      await fsp.unlink(tmp);
    } catch { /* ignore */ }
    throw err;
  }
}

/** Records a clean exit (close-gate 'close' outcome). Preserves discards. */
export async function writeCleanExit(markerPath: string, nowMs?: number): Promise<CleanExitMarker> {
  const prev = await readCleanExitMarker(markerPath);
  const next: CleanExitMarker = {
    cleanExitAt: typeof nowMs === 'number' && Number.isFinite(nowMs) ? nowMs : Date.now(),
    discarded: prev.discarded,
    version: 1,
  };
  try {
    await writeMarkerAtomic(markerPath, next);
  } catch { /* best-effort: a failed marker write must never block close */ }
  return next;
}

/** Marks a draft as explicitly discarded (never resurrected). Stores the
 * resolved FULL path (UX-C item 2). Bare basenames are preserved verbatim as
 * legacy fallbacks (old callers) but new callers MUST pass full paths. */
export async function markDraftDiscarded(
  markerPath: string,
  draftPathOrBasename: string,
  nowMs?: number
): Promise<CleanExitMarker> {
  const prev = await readCleanExitMarker(markerPath);
  const raw = String(draftPathOrBasename ?? '').trim();
  let stored = raw;
  if (raw && !isLegacyDiscardedBasenameEntry(raw)) {
    try {
      stored = path.resolve(raw);
    } catch { stored = raw; }
  }
  const set = new Set(prev.discarded);
  if (stored) {
    // Platform-aware dedupe: don't add a duplicate that already matches.
    const dup = prev.discarded.some((d) => {
      if (isLegacyDiscardedBasenameEntry(d) && isLegacyDiscardedBasenameEntry(stored)) {
        return process.platform === 'win32'
          ? d.toLowerCase() === stored.toLowerCase()
          : d === stored;
      }
      if (!isLegacyDiscardedBasenameEntry(d) && !isLegacyDiscardedBasenameEntry(stored)) {
        return areRecoveryPathsEqual(path.resolve(d), path.resolve(stored));
      }
      return false;
    });
    if (!dup) set.add(stored);
  }
  const next: CleanExitMarker = {
    cleanExitAt: typeof nowMs === 'number' && Number.isFinite(nowMs) ? nowMs : Date.now(),
    discarded: [...set],
    version: 1,
  };
  try {
    await writeMarkerAtomic(markerPath, next);
  } catch { /* best-effort */ }
  return next;
}

/**
 * UX-C item 2: full-path discard match (platform-aware). A target is
 * discarded when its resolved full path equals a full-path entry, OR when
 * its basename equals a LEGACY bare-basename entry (backwards compat).
 * Same basename in a different directory does NOT match a full-path entry,
 * and a.docx vs A.docx collide ONLY on win32.
 */
export function isDiscardedPath(
  targetPath: string,
  marker: CleanExitMarker,
  platform: string = process.platform
): boolean {
  let resolvedTarget: string;
  try {
    resolvedTarget = path.resolve(targetPath);
  } catch {
    return false;
  }
  const targetBase = path.basename(resolvedTarget);
  for (const entry of marker.discarded) {
    if (!entry) continue;
    if (isLegacyDiscardedBasenameEntry(entry)) {
      const want = platform === 'win32' ? targetBase.toLowerCase() : targetBase;
      const have = platform === 'win32' ? entry.toLowerCase() : entry;
      if (want === have) return true;
    } else {
      let resolvedEntry: string;
      try {
        resolvedEntry = path.resolve(entry);
      } catch { continue; }
      if (areRecoveryPathsEqual(resolvedTarget, resolvedEntry, platform)) return true;
    }
  }
  return false;
}

/**
 * Legacy alias (pre-fix name). Now platform-aware: case-insensitive on
 * win32, case-sensitive on Linux. Prefer isDiscardedPath for full paths.
 */
export function isDiscardedBasename(
  basename: string,
  marker: CleanExitMarker,
  platform: string = process.platform
): boolean {
  const want = platform === 'win32' ? String(basename).toLowerCase() : String(basename);
  return marker.discarded.some((d) => {
    const have = platform === 'win32' ? String(d).toLowerCase() : String(d);
    if (isLegacyDiscardedBasenameEntry(d)) return have === want;
    // Full-path entries match by their basename only for legacy callers that
    // pass a bare basename — keeps old suppression working post-upgrade.
    try {
      const entryBase = path.basename(path.resolve(d));
      const eb = platform === 'win32' ? entryBase.toLowerCase() : entryBase;
      return eb === want;
    } catch { return have === want; }
  });
}

/**
 * UX-C item 2: prune discarded entries for files that no longer exist
 * (GC-owned). Full-path entries whose resolved path is in `deletedPaths`
 * (platform-aware) are dropped; legacy basename entries are dropped only
 * when their basename matches a deleted file's basename AND no live file
 * with that basename remains (caller passes `liveBasenames` when known;
 * otherwise basename entries for deleted basenames are dropped).
 * Never throws; returns the pruned marker (persisted best-effort by caller).
 */
export function pruneDiscardedEntries(
  marker: CleanExitMarker,
  deletedPaths: string[],
  opts: { platform?: string; liveBasenames?: string[] } = {}
): CleanExitMarker {
  const platform = opts.platform ?? process.platform;
  const resolvedDeleted = new Set<string>();
  const deletedBasenames = new Set<string>();
  for (const p of deletedPaths) {
    try {
      const r = path.resolve(p);
      resolvedDeleted.add(platform === 'win32' ? r.toLowerCase() : r);
      const b = path.basename(r);
      deletedBasenames.add(platform === 'win32' ? b.toLowerCase() : b);
    } catch { /* ignore */ }
  }
  const liveSet = opts.liveBasenames
    ? new Set(opts.liveBasenames.map((b) => (platform === 'win32' ? String(b).toLowerCase() : String(b))))
    : null;
  const kept = marker.discarded.filter((entry) => {
    if (!entry) return false;
    if (isLegacyDiscardedBasenameEntry(entry)) {
      const key = platform === 'win32' ? entry.toLowerCase() : entry;
      if (!deletedBasenames.has(key)) return true;
      // A legacy basename dies with the file unless a live file reuses it.
      if (liveSet && liveSet.has(key)) return true;
      return false;
    }
    try {
      const r = path.resolve(entry);
      const key = platform === 'win32' ? r.toLowerCase() : r;
      return !resolvedDeleted.has(key);
    } catch {
      return true;
    }
  });
  return { cleanExitAt: marker.cleanExitAt, discarded: kept, version: 1 };
}

/** Persists a pruned marker (best-effort, never throws). */
export async function writePrunedMarker(markerPath: string, marker: CleanExitMarker): Promise<CleanExitMarker> {
  try {
    await writeMarkerAtomic(markerPath, marker);
  } catch { /* best-effort */ }
  return marker;
}

// ── Drafts scan (reads only) ────────────────────────────────────────────────

async function isSymlinkEntry(target: string): Promise<boolean> {
  try {
    const lst = await fsp.lstat(target);
    return lst.isSymbolicLink();
  } catch {
    return false;
  }
}

export async function listDraftDocxFiles(draftsDir: string): Promise<string[]> {
  let entries: string[];
  try {
    entries = await fsp.readdir(draftsDir);
  } catch {
    return [];
  }
  const out: string[] = [];
  for (const name of entries) {
    if (!name || name.startsWith('.')) continue;
    if (name.includes('.tmp-')) continue;
    if (!name.toLowerCase().endsWith('.docx')) continue;
    const full = path.join(draftsDir, name);
    try {
      if (await isSymlinkEntry(full)) continue;
      const st = await fsp.lstat(full);
      if (!st.isFile()) continue;
    } catch {
      continue;
    }
    out.push(full);
  }
  out.sort();
  return out;
}

export interface DraftRecoveryInfo {
  commitCount: number;
  headMessage: string | null;
  headTimeMs: number | null;
  mtimeMs: number;
  sizeBytes: number;
}

/** Reads journal head (getHistory) + file stat. Never throws. */
export async function getDraftRecoveryInfo(docPath: string): Promise<DraftRecoveryInfo> {
  let mtimeMs = 0;
  let sizeBytes = 0;
  try {
    const lst = await fsp.lstat(docPath);
    if (lst.isSymbolicLink()) {
      return { commitCount: 0, headMessage: null, headTimeMs: null, mtimeMs: 0, sizeBytes: 0 };
    }
    if (!lst.isFile()) {
      return { commitCount: 0, headMessage: null, headTimeMs: null, mtimeMs: 0, sizeBytes: 0 };
    }
    mtimeMs = lst.mtimeMs;
    sizeBytes = lst.size;
  } catch {
    return { commitCount: 0, headMessage: null, headTimeMs: null, mtimeMs: 0, sizeBytes: 0 };
  }
  try {
    const history = await getHistory(docPath);
    if (!Array.isArray(history) || history.length === 0) {
      return { commitCount: 0, headMessage: null, headTimeMs: null, mtimeMs, sizeBytes };
    }
    const head = history[0] as { message?: string; timestamp?: Date };
    const headTime = head.timestamp instanceof Date ? head.timestamp.getTime() : NaN;
    return {
      commitCount: history.length,
      headMessage: typeof head.message === 'string' ? head.message : null,
      headTimeMs: Number.isFinite(headTime) ? headTime : null,
      mtimeMs,
      sizeBytes,
    };
  } catch {
    return { commitCount: 0, headMessage: null, headTimeMs: null, mtimeMs, sizeBytes };
  }
}

export interface FindCandidatesOptions {
  markerPath?: string;
  nowMs?: number;
  maxAgeMs?: number;
}

/**
 * Startup scan: every Drafts *.docx whose journal head is beyond Initial,
 * newer than the clean-exit marker, not discarded, and not expired.
 * Sorted newest-head first. Includes the current-session path when it
 * qualifies — the renderer filters the just-restored path from the banner.
 * Discard matching is full-path + platform-aware (UX-C item 2).
 */
export async function findRecoveryCandidates(
  draftsDir: string,
  opts: FindCandidatesOptions = {}
): Promise<RecoveryCandidate[]> {
  const now = opts.nowMs ?? Date.now();
  const maxAge = opts.maxAgeMs ?? RECOVERY_BANNER_MAX_AGE_MS;
  const markerPath = opts.markerPath ?? getCleanExitMarkerPath(draftsDir);
  const marker = await readCleanExitMarker(markerPath);
  const files = await listDraftDocxFiles(draftsDir);
  const out: RecoveryCandidate[] = [];
  for (const docPath of files) {
    if (await isSymlinkEntry(docPath)) continue;
    const info = await getDraftRecoveryInfo(docPath);
    const basename = path.basename(docPath);
    const show = shouldShowRecoveryBannerBackend({
      commitCount: info.commitCount,
      headTimeMs: info.headTimeMs,
      cleanExitAtMs: marker.cleanExitAt,
      isDiscarded: isDiscardedPath(docPath, marker),
      nowMs: now,
      maxAgeMs: maxAge,
    });
    if (!show) continue;
    out.push({
      path: docPath,
      basename,
      commitCount: info.commitCount,
      headMessage: info.headMessage,
      headTimeMs: info.headTimeMs,
      mtimeMs: info.mtimeMs,
      sizeBytes: info.sizeBytes,
    });
  }
  out.sort((a, b) => (b.headTimeMs ?? 0) - (a.headTimeMs ?? 0));
  return out;
}

// ── Discard (banner [Discard]) ──────────────────────────────────────────────

/**
 * Explicit discard: deletes the Drafts .docx + its sibling
 * `.astra/<basename>/` git dir, then marks the resolved FULL path discarded
 * (UX-C item 2) so the banner never resurrects deleted intent. Titled docs
 * are refused (only paths inside draftsDir are accepted). Missing files are
 * tolerated.
 */
export async function discardRecoveredDraft(
  draftsDir: string,
  draftPath: string,
  opts: { markerPath?: string; nowMs?: number } = {}
): Promise<{ discarded: boolean; path: string }> {
  const resolvedDrafts = path.resolve(draftsDir);
  const resolvedTarget = path.resolve(draftPath);
  if (!isPathInsideDir(resolvedTarget, resolvedDrafts) && resolvedTarget !== resolvedDrafts) {
    return { discarded: false, path: draftPath };
  }
  if (!resolvedTarget.toLowerCase().endsWith('.docx')) {
    return { discarded: false, path: draftPath };
  }
  if (await isSymlinkEntry(resolvedTarget)) {
    return { discarded: false, path: draftPath };
  }
  try {
    if (await isSymlinkEntry(resolvedTarget)) {
      return { discarded: false, path: draftPath };
    }
    await fsp.rm(resolvedTarget, { force: true });
  } catch { /* best-effort */ }
  try {
    const repoDir = getRepoDir(resolvedTarget);
    if (await isSymlinkEntry(repoDir)) {
      try {
        await fsp.unlink(repoDir);
      } catch { /* best-effort */ }
    } else {
      await fsp.rm(repoDir, { recursive: true, force: true });
    }
  } catch { /* best-effort */ }
  const markerPath = opts.markerPath ?? getCleanExitMarkerPath(resolvedDrafts);
  try {
    await markDraftDiscarded(markerPath, resolvedTarget, opts.nowMs);
  } catch { /* best-effort */ }
  return { discarded: true, path: resolvedTarget };
}

// ── Orphan GC (pure decision + sweeping IO) ─────────────────────────────────

export interface GcOptions {
  nowMs?: number;
  maxAgeMs?: number;
  maxTotalBytes?: number;
  /** Never delete this path even when old/large (session-restorable). */
  currentSessionPath?: string | null;
  /** Marker to prune when GC deletes files (UX-C item 2). Defaults to Drafts marker. */
  markerPath?: string;
  /** Explicit platform for unit tests (defaults to process.platform). */
  platform?: string;
}

export interface GcDecision {
  toDelete: string[];
  kept: Array<{ path: string; reason: 'current-session' | 'fresh' | 'under-cap' }>;
}

/**
 * Pure liveness/age/size gate. Deletes entries older than maxAge, plus
 * oldest-first while the non-current total exceeds maxTotalBytes. The
 * current-session path is always kept. Callers must only pass Drafts-dir
 * entries — titled docs never enter this function.
 */
export function decideOrphanGC(entries: OrphanEntry[], opts: GcOptions = {}): GcDecision {
  const now = opts.nowMs ?? Date.now();
  const maxAge = opts.maxAgeMs ?? ORPHAN_GC_MAX_AGE_MS;
  const maxTotal = opts.maxTotalBytes ?? ORPHAN_GC_MAX_TOTAL_BYTES;
  const platform = opts.platform ?? process.platform;
  const normForCur = (p: string) => {
    try {
      const r = path.resolve(p);
      return platform === 'win32' ? r.toLowerCase() : r;
    } catch { return p; }
  };
  const curNorm = opts.currentSessionPath ? normForCur(opts.currentSessionPath) : null;

  const isCurrent = (p: string) => curNorm !== null && normForCur(p) === curNorm;
  const toDelete: string[] = [];
  const kept: GcDecision['kept'] = [];
  const deleteSet = new Set<string>();

  // Pass 1 — age gate (current-session exempt).
  for (const e of entries) {
    if (isCurrent(e.path)) continue;
    const age = now - e.mtimeMs;
    if (age > maxAge) {
      deleteSet.add(e.path);
    }
  }

  // Pass 2 — size gate over survivors (oldest-first, current exempt).
  const survivors = entries.filter((e) => !isCurrent(e.path) && !deleteSet.has(e.path));
  let total = entries
    .filter((e) => !isCurrent(e.path))
    .reduce((sum, e) => sum + (Number.isFinite(e.sizeBytes) ? e.sizeBytes : 0), 0);
  // Account for age-deleted bytes leaving the total.
  for (const e of entries) {
    if (deleteSet.has(e.path)) total -= Number.isFinite(e.sizeBytes) ? e.sizeBytes : 0;
  }
  if (total > maxTotal) {
    const byAge = [...survivors].sort((a, b) => a.mtimeMs - b.mtimeMs);
    for (const e of byAge) {
      if (total <= maxTotal) break;
      deleteSet.add(e.path);
      total -= Number.isFinite(e.sizeBytes) ? e.sizeBytes : 0;
    }
  }

  for (const e of entries) {
    if (deleteSet.has(e.path)) {
      toDelete.push(e.path);
    } else if (isCurrent(e.path)) {
      kept.push({ path: e.path, reason: 'current-session' });
    } else if (now - e.mtimeMs <= maxAge) {
      kept.push({ path: e.path, reason: total <= maxTotal ? 'under-cap' : 'fresh' });
    } else {
      kept.push({ path: e.path, reason: 'fresh' });
    }
  }
  toDelete.sort();
  return { toDelete, kept };
}

export interface SweepResult {
  deleted: string[];
  kept: Array<{ path: string; reason: string }>;
  totalBytesBefore: number;
  totalBytesAfter: number;
}

/**
 * Sweeping IO: lists Drafts *.docx + stats, applies decideOrphanGC, deletes
 * the losers (.docx + sibling .astra/<basename>/) best-effort. Also removes
 * crash-mid-write `*.tmp-<pid>-*` litter older than the age gate. NEVER
 * operates outside draftsDir — titled docs are unreachable by construction.
 * UX-C item 2: every deleted .docx is pruned from the marker's discarded set
 * (via opts.markerPath, defaulting to the Drafts marker) so a reused
 * basename is not permanently suppressed.
 */
export async function sweepOrphanDrafts(draftsDir: string, opts: GcOptions = {}): Promise<SweepResult> {
  const resolvedDrafts = path.resolve(draftsDir);
  let names: string[] = [];
  try {
    names = await fsp.readdir(resolvedDrafts);
  } catch {
    return { deleted: [], kept: [], totalBytesBefore: 0, totalBytesAfter: 0 };
  }

  const entries: OrphanEntry[] = [];
  const tmpLitter: string[] = [];
  for (const name of names) {
    if (!name || name.startsWith('.')) continue;
    const full = path.join(resolvedDrafts, name);
    if (name.includes('.tmp-')) {
      try {
        if (await isSymlinkEntry(full)) continue;
      } catch { continue; }
      tmpLitter.push(full);
      continue;
    }
    if (!name.toLowerCase().endsWith('.docx')) continue;
    try {
      const lst = await fsp.lstat(full);
      if (lst.isSymbolicLink()) continue;
      if (!lst.isFile()) continue;
      entries.push({ path: full, sizeBytes: lst.size, mtimeMs: lst.mtimeMs });
    } catch { /* vanished mid-scan — skip */ }
  }

  const now = opts.nowMs ?? Date.now();
  const maxAge = opts.maxAgeMs ?? ORPHAN_GC_MAX_AGE_MS;
  const decision = decideOrphanGC(entries, opts);
  const deleted: string[] = [];

  for (const target of decision.toDelete) {
    const resolved = path.resolve(target);
    if (!isPathInsideDir(resolved, resolvedDrafts)) continue;
    try {
      if (await isSymlinkEntry(resolved)) continue;
    } catch { continue; }
    try {
      await fsp.rm(resolved, { force: true });
    } catch { /* best-effort */ }
    try {
      const repoDir = getRepoDir(resolved);
      if (await isSymlinkEntry(repoDir)) {
        try {
          await fsp.unlink(repoDir);
        } catch { /* best-effort */ }
      } else {
        await fsp.rm(repoDir, { recursive: true, force: true });
      }
    } catch { /* best-effort */ }
    deleted.push(resolved);
  }

  // Tmp-litter pass: age-gated only, never size-gated, never current.
  for (const tmp of tmpLitter) {
    const resolved = path.resolve(tmp);
    if (!isPathInsideDir(resolved, resolvedDrafts)) continue;
    try {
      if (await isSymlinkEntry(resolved)) continue;
      const st = await fsp.lstat(resolved);
      if (now - st.mtimeMs > maxAge) {
        if (await isSymlinkEntry(resolved)) continue;
        await fsp.rm(resolved, { force: true });
        deleted.push(resolved);
      }
    } catch { /* best-effort */ }
  }

  // UX-C item 2: GC owns its litter in the marker too — prune discarded
  // entries for every deleted .docx so a reused basename is not permanently
  // suppressed. Tmp litter never enters the discarded set; only real .docx
  // deletions prune. Best-effort, never fails the sweep.
  if (deleted.length > 0) {
    try {
      const markerPath = opts.markerPath ?? getCleanExitMarkerPath(resolvedDrafts);
      const marker = await readCleanExitMarker(markerPath);
      if (marker.discarded.length > 0) {
        // Live basenames after deletion: legacy basename entries survive only
        // if a live file still reuses that basename.
        let liveNames: string[] = [];
        try {
          liveNames = await fsp.readdir(resolvedDrafts);
        } catch { liveNames = []; }
        const liveBasenames = liveNames.filter((n) => n && !n.startsWith('.') && n.toLowerCase().endsWith('.docx'));
        const pruned = pruneDiscardedEntries(marker, deleted, {
          platform: opts.platform ?? process.platform,
          liveBasenames,
        });
        if (pruned.discarded.length !== marker.discarded.length) {
          await writePrunedMarker(markerPath, pruned);
        }
      }
    } catch { /* best-effort: sweep result stands even if prune fails */ }
  }

  const totalBytesBefore = entries.reduce((s, e) => s + e.sizeBytes, 0);
  const deletedBytes = entries
    .filter((e) => deleted.includes(path.resolve(e.path)))
    .reduce((s, e) => s + e.sizeBytes, 0);
  return {
    deleted: deleted.sort(),
    kept: decision.kept.map((k) => ({ path: k.path, reason: k.reason })),
    totalBytesBefore,
    totalBytesAfter: Math.max(0, totalBytesBefore - deletedBytes),
  };
}
