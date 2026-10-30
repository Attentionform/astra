import fs from 'fs';
import type { Dirent } from 'fs';
import fsp from 'fs/promises';
import path from 'path';
import crypto from 'crypto';
import git from 'isomorphic-git';
import { assertWriteTargetNoSymlink, writeFileNoFollow } from './sandbox/paths';
import { toError } from './errors';

/**
 * `ASTRA_HISTORY_DIR`: keep every document's history sidecar (and its tombstone
 * registry) under this directory instead of next to the document. Unset or
 * empty keeps the default `<docDir>/.astra/`.
 */
export const HISTORY_DIR_ENV = 'ASTRA_HISTORY_DIR';

function historyRoot(): string | null {
  const raw = process.env[HISTORY_DIR_ENV];
  return typeof raw === 'string' && raw.trim() !== '' ? path.resolve(raw.trim()) : null;
}

/**
 * Folder name under the history root for one document directory (two folders can hold the same file name).
 * Keyed on the real folder, so reaching it through a symlink finds the same history.
 */
function historyKey(docDir: string): string {
  let dir = path.resolve(docDir);
  try {
    dir = fs.realpathSync(dir);
  } catch {
    /* the folder does not exist yet: key on the path as given */
  }
  return crypto.createHash('sha256').update(dir).digest('hex').slice(0, 16);
}

/** Everything the history root holds sits under `<root>/<16 hex characters>/`; a document elsewhere under the root is just a document. */
const HISTORY_SIDECAR_REL = /^[0-9a-f]{16}[\\/]/;

/** True when `target` is a sidecar repository directory rather than a document path. */
function isRepoDirTarget(target: string): boolean {
  const normalized = target.replace(/\\/g, '/');
  if (normalized.includes('/.astra/') || path.basename(target) === '.astra') return true;
  const root = historyRoot();
  if (root === null) return false;
  const resolved = path.resolve(target);
  if (!HISTORY_SIDECAR_REL.test(path.relative(root, resolved))) return false;
  // A document that happens to sit at such a path is still a document.
  try {
    return !fs.statSync(resolved).isFile();
  } catch {
    return true;
  }
}

// Helper to get the hidden .astra repository directory for a specific docx file
export function getRepoDir(docPath: string): string {
  // e.g. /home/user/thesis.astra.docx -> /home/user/.astra/thesis.astra.docx/
  // With ASTRA_HISTORY_DIR=/h: /h/<16-hex key of /home/user>/thesis.astra.docx/
  const root = historyRoot();
  if (root !== null) return path.join(root, historyKey(path.dirname(docPath)), path.basename(docPath));
  return path.join(path.dirname(docPath), '.astra', path.basename(docPath));
}

/**
 * OP-14: package files beyond document.xml that must be versioned so
 * media/rels/styles/headers survive commit/undo. document.xml is always
 * handled separately; everything matching this predicate is tracked.
 */
export function isTrackedPackageFile(entryName: string): boolean {
  const name = entryName.replace(/\\/g, '/');
  if (name === 'word/document.xml') return false;
  if (name === 'word/styles.xml') return true;
  if (/^word\/_rels\/(document|header[^/]*|footer[^/]*)\.xml\.rels$/.test(name)) return true;
  if (name === '[Content_Types].xml') return true;
  if (name === 'word/numbering.xml') return true;
  if (name === 'word/comments.xml') return true;
  if (name === 'word/footnotes.xml') return true;
  if (name === 'word/endnotes.xml') return true;
  if (name === 'word/settings.xml') return true;
  if (name.startsWith('word/media/')) return true;
  if (name.startsWith('word/charts/')) return true;
  if (name.startsWith('word/embeddings/')) return true;
  if (name.startsWith('word/theme/')) return true;
  if (/^word\/header[^/]*\.xml$/.test(name)) return true;
  if (/^word\/footer[^/]*\.xml$/.test(name)) return true;
  return false;
}

function toRepoRelative(entryName: string): string {
  return entryName.replace(/\\/g, '/');
}

async function writeAndAdd(
  repoDir: string,
  relPath: string,
  content: string | Buffer
): Promise<void> {
  const safeRel = toRepoRelative(relPath);
  // Block path traversal out of the repo dir.
  if (safeRel.includes('..') || path.isAbsolute(safeRel) || /^[a-zA-Z]:/.test(safeRel)) {
    throw new Error(`Security Violation: refusing to track package path '${relPath}'`);
  }
  const fullPath = path.join(repoDir, safeRel);
  const dirName = path.dirname(fullPath);
  // M1: refuse symlink-planted sidecar targets (same guard as the MCP
  // save/export/new-document/history paths: lstat the chain before mkdir so
  // mkdir -p cannot build through a symlink parent, re-check before the
  // O_NOFOLLOW write to close the race window). Refusals throw E_SANDBOX
  // with a remedy and leave history untouched.
  assertWriteTargetNoSymlink(fullPath);
  await fsp.mkdir(dirName, { recursive: true, mode: 0o700 });
  // mkdir mode only applies to newly created dirs: chmod the whole chain
  // (POSIX only) so pre-existing 0755 subdirs also land on 0700.
  await secureDirChain0700(repoDir, dirName);
  assertWriteTargetNoSymlink(fullPath);
  await writeFileNoFollow(fullPath, content);
  await git.add({ fs, dir: repoDir, filepath: safeRel });
}

/** POSIX-only 0700 chmod (no-op on Windows, best-effort on chmod-less filesystems). */
async function chmod0700BestEffort(dir: string): Promise<void> {
  if (process.platform === 'win32') return;
  try {
    await fsp.chmod(dir, 0o700);
  } catch {
    /* ignore on filesystems not supporting chmod like exFAT/NFS */
  }
}

/**
 * Chmod every directory from repoDir down to `from` (inclusive) to 0700.
 * Guards both freshly created chains (mkdir mode is masked by umask and
 * skipped for pre-existing dirs) and pre-existing 0755 trees. POSIX only;
 * never throws.
 */
async function secureDirChain0700(repoDir: string, from: string): Promise<void> {
  if (process.platform === 'win32') return;
  try {
    const root = path.resolve(repoDir);
    let cur = path.resolve(from);
    if (cur !== root && !cur.startsWith(root + path.sep)) return;
    const stack: string[] = [];
    for (;;) {
      stack.push(cur);
      if (cur === root) break;
      const parent = path.dirname(cur);
      if (parent === cur) break;
      cur = parent;
    }
    for (const dir of stack) {
      await chmod0700BestEffort(dir);
    }
  } catch {
    /* best-effort — permissions must never fail a commit */
  }
}

/**
 * Best-effort recursive migration of a pre-existing repo tree to 0700 dirs
 * (POSIX only). Fresh mkdir already covers new dirs; this catches trees
 * created before the 0700 policy (0755 default). Skips files by design
 * (only dirs get chmod); symlinked dirs are skipped without following to
 * avoid escape outside the repo (chmod would hit the target); unreadable
 * entries are skipped; never throws.
 * Note: readdir never yields '.'/'..', so no extra name guard is needed.
 */
async function migrateExistingTree0700(repoDir: string): Promise<void> {
  if (process.platform === 'win32') return;
  try {
    const stack: string[] = [repoDir];
    while (stack.length > 0) {
      const dir = stack.pop() as string;
      await chmod0700BestEffort(dir);
      let entries: Dirent[] = [];
      try {
        entries = await fsp.readdir(dir, { withFileTypes: true });
      } catch {
        continue;
      }
      for (const entry of entries) {
        // Skip symlinks without following: a symlink-to-dir must never be
        // entered or chmodded (avoids escape outside the repo).
        if (typeof entry?.isSymbolicLink === 'function' && entry.isSymbolicLink()) continue;
        if (typeof entry?.isDirectory !== 'function' || !entry.isDirectory()) continue;
        stack.push(path.join(dir, entry.name));
      }
    }
  } catch {
    /* best-effort */
  }
}

/**
 * Default author recorded on git commits. Callers may override per commit;
 * blank values fall back here. Single-line, capped by the MCP hygiene layer.
 */
export const DEFAULT_COMMIT_AUTHOR = 'ASTra Agent';

function resolveCommitAuthor(author?: string): string {
  const name = typeof author === 'string' ? author.replace(/[\r\n]+/g, ' ').trim() : '';
  return name ? name.slice(0, 100) : DEFAULT_COMMIT_AUTHOR;
}

/**
 * History quotas (per document sidecar). High defaults so normal documents
 * (tens of commits, MBs) never hit them; only runaway rapid-save loops do.
 * Configurable via env or per-call limits. Loud E_QUOTA refusal names the
 * raise-the-limit remedy and never silently drops history.
 */
export const HISTORY_MAX_COMMITS_PER_DOC = 1000;
export const HISTORY_MAX_SIDECAR_BYTES = 500 * 1024 * 1024;
/** Env overrides for the history quotas (positive integers, bytes for the latter). */
export const HISTORY_MAX_COMMITS_ENV = 'ASTRA_HISTORY_MAX_COMMITS';
export const HISTORY_MAX_SIDECAR_BYTES_ENV = 'ASTRA_HISTORY_MAX_SIDECAR_BYTES';
/** L1 upper clamps: env/explicit values above these are capped (env-driven memory/disk exhaustion). */
export const HISTORY_MAX_COMMITS_CLAMP = 100_000;
export const HISTORY_MAX_SIDECAR_BYTES_CLAMP = 5 * 1024 * 1024 * 1024;
/** Error code for history quota refusals. */
export const HISTORY_QUOTA_CODE = 'E_QUOTA' as const;

export interface HistoryQuotaLimits {
  maxCommits?: number;
  maxSidecarBytes?: number;
}

/** Resolves the per-document commit cap (explicit > env > default). Never below 1. L1: clamped to HISTORY_MAX_COMMITS_CLAMP. */
export function resolveHistoryMaxCommits(
  explicit?: number,
  env: NodeJS.ProcessEnv = process.env
): number {
  if (typeof explicit === 'number' && Number.isFinite(explicit) && explicit >= 1) {
    return Math.min(Math.floor(explicit), HISTORY_MAX_COMMITS_CLAMP);
  }
  const raw = (env[HISTORY_MAX_COMMITS_ENV] ?? '').trim();
  const parsed = raw ? Number.parseInt(raw, 10) : NaN;
  if (Number.isFinite(parsed) && parsed >= 1) return Math.min(Math.floor(parsed), HISTORY_MAX_COMMITS_CLAMP);
  return HISTORY_MAX_COMMITS_PER_DOC;
}

/** Resolves the per-document sidecar bytes cap (explicit > env > default). Never below 1. L1: clamped to HISTORY_MAX_SIDECAR_BYTES_CLAMP. */
export function resolveHistoryMaxSidecarBytes(
  explicit?: number,
  env: NodeJS.ProcessEnv = process.env
): number {
  if (typeof explicit === 'number' && Number.isFinite(explicit) && explicit >= 1) {
    return Math.min(Math.floor(explicit), HISTORY_MAX_SIDECAR_BYTES_CLAMP);
  }
  const raw = (env[HISTORY_MAX_SIDECAR_BYTES_ENV] ?? '').trim();
  const parsed = raw ? Number.parseInt(raw, 10) : NaN;
  if (Number.isFinite(parsed) && parsed >= 1) return Math.min(Math.floor(parsed), HISTORY_MAX_SIDECAR_BYTES_CLAMP);
  return HISTORY_MAX_SIDECAR_BYTES;
}

/** Best-effort recursive size of a sidecar repo dir (bytes). Never throws. */
export async function getSidecarBytes(repoDir: string): Promise<number> {
  let total = 0;
  const stack: string[] = [repoDir];
  while (stack.length > 0) {
    const dir = stack.pop() as string;
    let entries: Dirent[] = [];
    try {
      entries = await fsp.readdir(dir, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const entry of entries) {
      const full = path.join(dir, entry.name);
      try {
        if (typeof entry?.isSymbolicLink === 'function' && entry.isSymbolicLink()) continue;
        if (typeof entry?.isDirectory === 'function' && entry.isDirectory()) {
          stack.push(full);
          continue;
        }
        if (typeof entry?.isFile === 'function' && !entry.isFile()) continue;
        const st = await fsp.stat(full);
        if (st.isDirectory()) stack.push(full);
        else if (st.isFile()) total += st.size;
      } catch {
        /* skip unreadable */
      }
    }
  }
  return total;
}

/** Best-effort commit count for a repo dir (0 when history is unreadable). Never throws. */
export async function getHistoryCommitCount(repoDir: string): Promise<number> {
  try {
    await assertUsableSidecarGit(repoDir);
    const commits = await git.log({ fs, dir: repoDir });
    return Array.isArray(commits) ? commits.length : 0;
  } catch {
    return 0;
  }
}

/**
 * R-6 tombstone warning (decision Q3 option B): deleting the `.astra/`
 * sidecar used to silently regenerate a lookalike fresh chain, hiding the
 * loss of prior history. Full append-only protection is deferred; this
 * layer only warns, never blocks opening.
 *
 * SIGNAL (documented): a small registry file
 * `<docDir>/.astra-tombstones.json` records every basename that has ever
 * reached a valid git chain (written after a successful init, refreshed on
 * every intact ensure). The registry lives OUTSIDE `.astra/` so it survives
 * both per-document sidecar removal (`.astra/<name>/`) and whole-directory
 * removal (`.astra/`). When ensureAstraGitRepo finds no valid `.git/` chain
 * for a basename the registry knows, the old chain was almost certainly
 * removed and a warning is recorded on the typed channel
 * (getSidecarHistoryWarning and the `warning` field on getHistory results;
 * library code stays silent, CLI remains the logging surface).
 *
 * WHY NOT mtime (reason for registry-outside-sidecar): document mtime
 * cannot serve as the removal signal because an old-but-never-opened docx
 * (no sidecar yet, old mtime) is indistinguishable from a docx whose
 * sidecar was deleted (no sidecar now, old mtime). mtime also moves on
 * every save/export, so it cannot name which basename once had history.
 * The registry instead records explicit first-seen proof that a basename
 * once reached a valid chain, plus its last-known head oid for
 * restart-durable mismatch detection.
 *
 * Honest gaps: removing the registry file alongside the sidecar is
 * indistinguishable from a first open (no warning); a corrupt registry is
 * treated as no signal; reuse of a basename after an intentional discard
 * (discardUntitledDraft removes the sidecar on purpose) can warn
 * spuriously — the text says "possible" for exactly this reason.
 * Detection is advisory and best-effort: all registry IO is wrapped so it
 * can never fail an open, commit, or ensure.
 */
export const SIDECAR_TOMBSTONE_FILENAME = '.astra-tombstones.json';

/** Pending regeneration warnings, keyed by resolved repo dir. */
const sidecarRegenWarnings = new Map<string, string>();

function resolveDocIdentity(target: string): { registryDir: string; base: string } | null {
  try {
    const root = historyRoot();
    if (!isRepoDirTarget(target)) {
      const docDir = path.dirname(target);
      return { registryDir: root !== null ? path.join(root, historyKey(docDir)) : docDir, base: path.basename(target) };
    }
    // Repo-dir form: only the exact `<docDir>/.astra/<base>` shape (or
    // `<history root>/<key>/<base>`) maps back to a document; deeper in-repo
    // paths carry no basename signal.
    const parent = path.dirname(target);
    if (root !== null && path.resolve(path.dirname(parent)) === root) {
      return { registryDir: parent, base: path.basename(target) };
    }
    if (path.basename(parent) !== '.astra') return null;
    return { registryDir: path.dirname(parent), base: path.basename(target) };
  } catch {
    return null;
  }
}

type TombstoneRegistry = Record<
  string,
  { firstSeenAt: number; lastSeenAt: number; head: string | null; warned?: boolean }
>;

function tombstoneRegistryPath(registryDir: string): string {
  return path.join(registryDir, SIDECAR_TOMBSTONE_FILENAME);
}

/** Sync write (see recordSidecarSeen); the folder may not exist yet under a history root. */
function writeTombstoneRegistry(registryDir: string, reg: TombstoneRegistry): void {
  fs.mkdirSync(registryDir, { recursive: true, mode: 0o700 });
  fs.writeFileSync(tombstoneRegistryPath(registryDir), JSON.stringify(reg), 'utf8');
  if (process.platform !== 'win32') {
    try {
      fs.chmodSync(tombstoneRegistryPath(registryDir), 0o600);
    } catch {
      /* best-effort */
    }
  }
}

/** Best-effort current HEAD oid for durability comparison. Never throws. */
async function getRepoHeadOid(repoDir: string): Promise<string | null> {
  try {
    const commits = await git.log({ fs, dir: repoDir, depth: 1 });
    const oid = (commits as Array<{ oid?: unknown }>)[0]?.oid;
    if (typeof oid === 'string' && oid.length > 0) return oid;
  } catch {
    /* fall through to resolveRef */
  }
  try {
    const oid = await git.resolveRef({ fs, dir: repoDir, ref: 'HEAD' });
    return typeof oid === 'string' && oid.length > 0 ? oid : null;
  } catch {
    return null;
  }
}

/**
 * Checks whether a stored head oid is still reachable in the current chain.
 * Returns true (continuity / normal appends), false (replacement: old head
 * gone), or null when the log cannot be read (unknown — caller must stay
 * silent and never block). Never throws.
 */
async function isOidPresent(repoDir: string, oid: string): Promise<boolean | null> {
  try {
    const commits = await git.log({ fs, dir: repoDir });
    for (const c of commits as Array<{ oid?: unknown }>) {
      if (c.oid === oid) return true;
    }
    return false;
  } catch {
    try {
      await git.readCommit({ fs, dir: repoDir, oid });
      return true;
    } catch {
      return null;
    }
  }
}

async function readTombstoneRegistry(registryDir: string): Promise<TombstoneRegistry | null> {
  try {
    // Sync read by design (see recordSidecarSeen): tiny payload, and it
    // keeps registry IO off the async fault-injection surface.
    const raw = fs.readFileSync(tombstoneRegistryPath(registryDir), 'utf8');
    const parsed: unknown = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null;
    return parsed as TombstoneRegistry;
  } catch {
    return null;
  }
}

async function recordSidecarSeen(registryDir: string, base: string, head?: string | null): Promise<void> {
  try {
    if (!base || base === '.' || base === '..' || base.includes('/') || base.includes('\\')) return;
    const existing = await readTombstoneRegistry(registryDir);
    const reg: TombstoneRegistry = existing ?? {};
    const now = Date.now();
    const prev = reg[base];
    const firstSeenAt = prev && typeof prev.firstSeenAt === 'number' ? prev.firstSeenAt : now;
    const prevHead = prev && typeof prev.head === 'string' && prev.head.length > 0 ? prev.head : null;
    const prevWarned = prev && prev.warned === true;
    const nextHead =
      typeof head === 'string' && head.length > 0
        ? head
        : prevHead;
    reg[base] = { firstSeenAt, lastSeenAt: now, head: nextHead, warned: prevWarned || undefined };
    // Sync write by design: the payload is under 1KB, sync read-modify-write
    // is atomic in-process (no interleave between concurrent ensures), and it
    // keeps registry IO off the async `fs.promises.writeFile` surface that
    // fault-injection suites spy on — a registry write must never absorb an
    // injected outage aimed at the document write nor inflate its call count.
    writeTombstoneRegistry(registryDir, reg);
  } catch {
    /* advisory only — never fail the caller */
  }
}

/**
 * Persists the R-6 regeneration signal so a post-restart ensure still
 * re-warns even when the regenerated initial oid collides deterministically
 * with the stored head (same XML, same second). Preserves the stored old
 * head; only flips the warned flag. Never throws.
 */
async function markTombstoneWarned(registryDir: string, base: string): Promise<void> {
  try {
    if (!base || base === '.' || base === '..' || base.includes('/') || base.includes('\\')) return;
    const existing = await readTombstoneRegistry(registryDir);
    const reg: TombstoneRegistry = existing ?? {};
    const now = Date.now();
    const prev = reg[base];
    const head = prev && typeof prev.head === 'string' && prev.head.length > 0 ? prev.head : null;
    const firstSeenAt = prev && typeof prev.firstSeenAt === 'number' ? prev.firstSeenAt : now;
    reg[base] = { firstSeenAt, lastSeenAt: now, head, warned: true };
    writeTombstoneRegistry(registryDir, reg);
  } catch {
    /* advisory only */
  }
}

function buildSidecarRemovedWarning(base: string): string {
  return (
    `[ASTra] WARNING: possible history removal for '${base}' — the .astra/ ` +
    `sidecar was missing despite previously recorded history, so a fresh ` +
    `history chain was started. Earlier commits may have been removed with ` +
    `the deleted sidecar. Opening continues normally; full append-only ` +
    `protection is deferred.`
  );
}

/**
 * Response-field accessor for the R-6 tombstone warning. Open/history
 * callers (openDocument, getVersionHistory, MCP history tools) should
 * include this string in their responses alongside the normal payload:
 * it carries the warning recorded at regeneration time.
 * Returns the warning recorded by the most recent regenerating ensure,
 * or null when the chain is intact. Never throws.
 */
export function getSidecarHistoryWarning(target: string): string | null {
  try {
    const repoDir = isRepoDirTarget(target) ? target : getRepoDir(target);
    return sidecarRegenWarnings.get(path.resolve(repoDir)) ?? null;
  } catch {
    return null;
  }
}

/** Test hook: clears the pending R-6 warning for one document. Never throws. */
export function clearSidecarHistoryWarning(target: string): void {
  try {
    const repoDir = isRepoDirTarget(target) ? target : getRepoDir(target);
    sidecarRegenWarnings.delete(path.resolve(repoDir));
  } catch {
    /* ignore */
  }
}

/**
 * R-6 check run by ensureAstraGitRepo. Intact chains (`.git/` present)
 * refresh the registry and drop any stale warning; a missing chain for a
 * registry-known basename records a warning on the typed channel.
 * Intact chains compare the current HEAD against the registry's last-known
 * head: normal appends retain the old oid (adopt the new head, stay
 * silent) while a regenerated chain drops it (preserve the old head so a
 * post-restart ensure still mismatches, and re-record). Never throws; never
 * blocks the caller.
 */
async function checkSidecarTombstone(target: string, repoDir: string): Promise<void> {
  const ident = resolveDocIdentity(target);
  if (!ident) return;
  const key = path.resolve(repoDir);
  let hasGit = false;
  try {
    hasGit = (await fsp.stat(path.join(repoDir, '.git'))).isDirectory();
  } catch {
    hasGit = false;
  }
  if (hasGit) {
    // A warning already pending in this process (set by the leading
    // missing-chain ensure before this intact re-init/commit) must survive:
    // deterministic re-init oids can equal the stored head, so head-equality
    // below must not clear a just-recorded warning. Persist the flag for
    // post-restart durability, refresh lastSeen only, keep the stored head.
    if (sidecarRegenWarnings.has(key)) {
      try {
        await markTombstoneWarned(ident.registryDir, ident.base);
      } catch {
        /* ignore */
      }
      try {
        await recordSidecarSeen(ident.registryDir, ident.base);
      } catch {
        /* ignore */
      }
      return;
    }
    let currentHead: string | null = null;
    try {
      currentHead = await getRepoHeadOid(repoDir);
    } catch {
      currentHead = null;
    }
    let entry: { firstSeenAt: number; lastSeenAt: number; head: string | null; warned?: boolean } | undefined;
    try {
      const reg = await readTombstoneRegistry(ident.registryDir);
      entry = reg?.[ident.base];
    } catch {
      entry = undefined;
    }
    if (!entry) {
      // First intact sighting (also silently adopts pre-tombstone repos):
      // record the current head and drop any stale in-memory warning.
      await recordSidecarSeen(ident.registryDir, ident.base, currentHead);
      sidecarRegenWarnings.delete(key);
      return;
    }
    // Persisted regeneration flag (set before a restart): re-warn even when
    // heads collide deterministically. Never auto-clears; only the test hook
    // clears in-memory state, and the flag keeps the next ensure loud.
    if (entry.warned === true) {
      await recordSidecarSeen(ident.registryDir, ident.base);
      if (!sidecarRegenWarnings.has(key)) {
        const msg = buildSidecarRemovedWarning(ident.base);
        sidecarRegenWarnings.set(key, msg);
      }
      return;
    }
    const storedHead =
      entry && typeof entry.head === 'string' && entry.head.length > 0 ? entry.head : null;
    if (!storedHead || !currentHead || storedHead === currentHead) {
      // Continuity (or unknown head on either side): adopt the current head
      // when known, refresh timestamps, clear any stale warning.
      if (currentHead) await recordSidecarSeen(ident.registryDir, ident.base, currentHead);
      else await recordSidecarSeen(ident.registryDir, ident.base);
      sidecarRegenWarnings.delete(key);
      return;
    }
    // Heads differ: distinguish normal appends (old oid still reachable)
    // from a regenerated chain (old oid gone).
    let oldPresent: boolean | null = null;
    try {
      oldPresent = await isOidPresent(repoDir, storedHead);
    } catch {
      oldPresent = null;
    }
    if (oldPresent === true) {
      await recordSidecarSeen(ident.registryDir, ident.base, currentHead);
      sidecarRegenWarnings.delete(key);
      return;
    }
    if (oldPresent === false) {
      // Replacement: preserve the stored old head (durability across
      // restarts — overwriting it would make the next ensure look intact),
      // refresh lastSeen only, persist the warned flag, and re-warn when
      // in-memory state is empty (e.g. after a restart).
      await recordSidecarSeen(ident.registryDir, ident.base);
      try {
        await markTombstoneWarned(ident.registryDir, ident.base);
      } catch {
        /* ignore */
      }
      if (!sidecarRegenWarnings.has(key)) {
        const msg = buildSidecarRemovedWarning(ident.base);
        sidecarRegenWarnings.set(key, msg);
      }
      return;
    }
    // Unknown (log unreadable): leave the stored head untouched and keep the
    // current warning state; never block the caller.
    try {
      await recordSidecarSeen(ident.registryDir, ident.base);
    } catch {
      /* ignore */
    }
    return;
  }
  // No valid chain. This is proof of removal only when the registry knows
  // this basename: a crashed first init leaves no entry, so it stays silent.
  const reg = await readTombstoneRegistry(ident.registryDir);
  if (reg && Object.prototype.hasOwnProperty.call(reg, ident.base)) {
    if (!sidecarRegenWarnings.has(key)) {
      const msg = buildSidecarRemovedWarning(ident.base);
      sidecarRegenWarnings.set(key, msg);
    }
    try {
      await markTombstoneWarned(ident.registryDir, ident.base);
    } catch {
      /* advisory only */
    }
  }
}

/**
 * Refuse a sidecar whose `.git` entry is not a repository directory (reads and
 * writes alike). Mirrors
 * the CLI history classifier (sidecarGitState in packages/cli/src/cli/
 * history-cmd.ts): a real directory is a usable repo; a symlink is followed
 * like isomorphic-git follows it — a link to a real directory is usable, a
 * link to a file or a broken link is stray; any other entry kind (e.g. a
 * crafted `gitdir:` redirect file) is stray. Absent stays absent (the
 * auto-init path is unchanged). A stray entry is never followed: isomorphic-
 * git would resolve a `gitdir:` redirect and write objects/refs into the
 * target outside the sidecar.
 */
async function assertUsableSidecarGit(repoDir: string): Promise<void> {
  const gitPath = path.join(repoDir, '.git');
  let entry: fs.Stats;
  try {
    entry = await fsp.lstat(gitPath);
  } catch {
    return;
  }
  let usable = false;
  if (entry.isSymbolicLink()) {
    try {
      usable = (await fsp.stat(gitPath)).isDirectory();
    } catch {
      usable = false;
    }
  } else {
    usable = entry.isDirectory();
  }
  if (!usable) {
    throw toError(
      'E_IO',
      'the history sidecar for this document has a ".git" entry that is not a repository directory — ' +
        'refusing to open or write it (a stray .git file can redirect the repository outside the sidecar).',
      'Remove the stray ".git" entry inside the ".astra/<basename>/" sidecar directory, then retry — ' +
        'the sidecar repository is recreated on the next history init or save.'
    );
  }
}

/**
 * T-4: Ensure .astra repository directory exists with restrictive 0o700 POSIX permissions.
 * Pre-existing 0755 trees (created before this policy) are migrated: the repo
 * dir itself is always chmodded and the existing subtree walked best-effort.
 * All chmod work is POSIX-only (skipped on Windows) and never throws.
 */
export async function ensureAstraGitRepo(target: string): Promise<string> {
  const repoDir = isRepoDirTarget(target) ? target : getRepoDir(target);
  // Fail closed before any write when `.git` is a stray entry (redirect file,
  // symlink to a file, broken symlink): never follow it into another target.
  await assertUsableSidecarGit(repoDir);
  await fsp.mkdir(repoDir, { recursive: true, mode: 0o700 });
  if (process.platform !== 'win32') {
    await chmod0700BestEffort(repoDir);
    const parent = path.dirname(repoDir);
    const root = historyRoot();
    if (path.basename(parent) === '.astra' || (root !== null && path.resolve(path.dirname(parent)) === root)) {
      await chmod0700BestEffort(parent);
    }
    await migrateExistingTree0700(repoDir);
  }
  // R-6 tombstone check: advisory only, never throws, never blocks opening.
  try {
    await checkSidecarTombstone(target, repoDir);
  } catch {
    /* warning path must never fail the ensure */
  }
  return repoDir;
}

export async function initAstraRepo(
  docPath: string,
  initialXml: string,
  extraFiles?: Record<string, string | Buffer>
): Promise<string> {
  const repoDir = await ensureAstraGitRepo(docPath);

  await git.init({ fs, dir: repoDir });

  await writeAndAdd(repoDir, 'document.xml', initialXml);
  if (extraFiles) {
    for (const [name, content] of Object.entries(extraFiles)) {
      await writeAndAdd(repoDir, name, content as string | Buffer);
    }
  }

  await git.commit({
    fs,
    dir: repoDir,
    author: { name: 'ASTra', email: 'astra@attentionform.com' },
    message: 'Initial state'
  });

  // R-6: record the new head durably. Re-runs the tombstone check so a
  // regenerated chain preserves the stored old head (post-restart mismatch
  // still warns) instead of overwriting it; first chains are recorded with
  // their head. Never clears a warning set by the leading ensure.
  try {
    await checkSidecarTombstone(docPath, repoDir);
  } catch {
    /* advisory only — never fail the open */
  }

  return repoDir;
}

export async function commitChange(
  docPath: string,
  xmlContent: string,
  message: string,
  extraFiles?: Record<string, string | Buffer>,
  deletedFiles?: string[],
  author?: string,
  limits?: HistoryQuotaLimits
): Promise<string> {
  const maxCommits = resolveHistoryMaxCommits(limits?.maxCommits);
  const maxSidecarBytes = resolveHistoryMaxSidecarBytes(limits?.maxSidecarBytes);
  const repoDir = await ensureAstraGitRepo(docPath);

  // Quota gate BEFORE any write: commit count first (cheap), then sidecar
  // bytes (current + incoming). Loud E_QUOTA names the raise-the-limit
  // remedy and leaves history untouched.
  const existingCommits = await getHistoryCommitCount(repoDir);
  if (existingCommits + 1 > maxCommits) {
    throw Object.assign(
      new Error(
        `${HISTORY_QUOTA_CODE}: save_document refused: history for '${docPath}' would reach ` +
          `${existingCommits + 1} commits, exceeding cap ${maxCommits} commits. ` +
          `Raise the limit via ${HISTORY_MAX_COMMITS_ENV} or commitChange limits.maxCommits.`
      ),
      {
        code: HISTORY_QUOTA_CODE,
        remedy:
          `History is capped at ${HISTORY_MAX_COMMITS_PER_DOC} commits per document by default. ` +
          `Raise the limit via ${HISTORY_MAX_COMMITS_ENV} or commitChange limits.maxCommits, ` +
          `or archive/prune history before saving again.`,
      }
    );
  }
  try {
    const currentBytes = await getSidecarBytes(repoDir);
    let incoming = Buffer.byteLength(xmlContent, 'utf8');
    if (extraFiles) {
      for (const content of Object.values(extraFiles)) {
        if (typeof content === 'string') incoming += Buffer.byteLength(content, 'utf8');
        else if (Buffer.isBuffer(content)) incoming += (content as Buffer).length;
        else {
          const maybe = content as unknown as { length?: unknown };
          if (maybe && typeof maybe.length === 'number') incoming += maybe.length as number;
        }
      }
    }
    if (currentBytes + incoming > maxSidecarBytes) {
      throw Object.assign(
        new Error(
          `${HISTORY_QUOTA_CODE}: save_document refused: sidecar for '${docPath}' would reach ` +
            `${currentBytes + incoming} bytes, exceeding cap ${maxSidecarBytes} bytes. ` +
            `Raise the limit via ${HISTORY_MAX_SIDECAR_BYTES_ENV} or commitChange limits.maxSidecarBytes.`
        ),
        {
          code: HISTORY_QUOTA_CODE,
          remedy:
            `History sidecar is capped at ${HISTORY_MAX_SIDECAR_BYTES} bytes per document by default. ` +
            `Raise the limit via ${HISTORY_MAX_SIDECAR_BYTES_ENV} or commitChange limits.maxSidecarBytes, ` +
            `or reduce document media before saving again.`,
        }
      );
    }
  } catch (err: unknown) {
    if ((err as { code?: unknown } | null)?.code === HISTORY_QUOTA_CODE) throw err;
    /* size probe is best-effort: unreadable sidecar never blocks a commit */
  }

  await writeAndAdd(repoDir, 'document.xml', xmlContent);
  if (extraFiles) {
    for (const [name, content] of Object.entries(extraFiles)) {
      await writeAndAdd(repoDir, name, content as string | Buffer);
    }
  }
  // OP-14: remove package files deleted from the live package so stale media
  // does not linger in git history. Explicit deletedFiles win; otherwise, when
  // the caller supplied a complete extraFiles snapshot, auto-detect tracked
  // files present in the repo but absent from the new package.
  try {
    const toDelete = new Set<string>();
    if (deletedFiles) {
      for (const name of deletedFiles) {
        const rel = toRepoRelative(name);
        if (rel === 'document.xml') continue;
        toDelete.add(rel);
      }
    }
    if (extraFiles) {
      const newTracked = new Set(
        Object.keys(extraFiles).map(k => toRepoRelative(k))
      );
      // Auto-detect via git index (authoritative tracked set).
      try {
        const tracked: string[] = await git.listFiles({ fs, dir: repoDir });
        for (const t of tracked) {
          const rel = toRepoRelative(t);
          if (rel === 'document.xml') continue;
          if (isTrackedPackageFile(rel) && !newTracked.has(rel)) {
            toDelete.add(rel);
          }
        }
      } catch {
        /* fall through to disk scan */
      }
      // Auto-detect via working-tree scan (catches stray on-disk files).
      try {
        const onDisk = await collectRepoFiles(repoDir);
        for (const rel of onDisk.keys()) {
          if (rel === 'document.xml') continue;
          if (isTrackedPackageFile(rel) && !newTracked.has(rel)) {
            toDelete.add(rel);
          }
        }
      } catch {
        /* best-effort */
      }
    }
    for (const rel of toDelete) {
      if (rel.includes('..') || path.isAbsolute(rel) || /^[a-zA-Z]:/.test(rel)) continue;
      const fullPath = path.join(repoDir, rel);
      try {
        await fsp.unlink(fullPath);
      } catch {
        /* already gone — still remove from index */
      }
      try {
        await git.remove({ fs, dir: repoDir, filepath: rel });
      } catch {
        /* not tracked — nothing to unstage */
      }
    }
  } catch {
    /* deletion cleanup is best-effort; never fail the commit */
  }
  const sha = await git.commit({
    fs,
    dir: repoDir,
    author: { name: resolveCommitAuthor(author), email: 'astra@attentionform.com' },
    message
  });

  // R-6: adopt the new head so the registry tracks the latest oid (a stale
  // initial head can collide deterministically with a regenerated initial
  // commit). Warning-aware: when a regeneration warning is pending the check
  // preserves the old head and keeps the warning instead of overwriting.
  try {
    await checkSidecarTombstone(docPath, repoDir);
  } catch {
    /* advisory only — never fail the commit */
  }

  return sha;
}

export async function getHistory(docPath: string) {
  const repoDir = getRepoDir(docPath);
  // Reads obey the same rule as writes: a stray `.git` entry is never followed.
  await assertUsableSidecarGit(repoDir);
  let commits;
  try {
    commits = await git.log({ fs, dir: repoDir });
  } catch {
    throw new Error(
      `Cannot read history for '${docPath}': no history exists yet (open or save the document first).`
    );
  }

  const entries = commits.map(c => ({
    oid: c.oid,
    message: c.commit.message,
    author: c.commit.author.name,
    timestamp: new Date(c.commit.author.timestamp * 1000)
  }));
  // R-6 history response field: the tombstone warning (if any) rides along
  // as an enumerable `warning` so spread/Object.keys/IPC projections carry
  // it. Array extras never survive JSON.stringify even when enumerable, so
  // an explicit non-enumerable `toJSON` projection carries the warning for
  // JSON/IPC consumers that stringify the history value directly.
  // Never throws: a missing warning surfaces as null, never as a failure.
  let warning: string | null = null;
  try {
    warning = getSidecarHistoryWarning(docPath);
  } catch {
    warning = null;
  }
  Object.defineProperty(entries, 'warning', {
    value: warning,
    enumerable: true,
    writable: false,
    configurable: true,
  });
  Object.defineProperty(entries, 'toJSON', {
    value: function (this: Array<unknown> & { warning?: string | null }) {
      return { entries: Array.from(this), warning: this.warning ?? null };
    },
    enumerable: false,
    writable: false,
    configurable: true,
  });
  return entries;
}

async function collectRepoFiles(repoDir: string): Promise<Map<string, Buffer>> {
  const out = new Map<string, Buffer>();
  async function walk(dir: string, relBase: string): Promise<void> {
    let entries: Dirent[] = [];
    try {
      entries = await fsp.readdir(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      if (entry.name === '.git') continue;
      const full = path.join(dir, entry.name);
      const rel = relBase ? `${relBase}/${entry.name}` : entry.name;
      if (entry.isDirectory()) {
        await walk(full, rel);
      } else if (entry.isFile()) {
        try {
          out.set(rel, await fsp.readFile(full));
        } catch {
          /* skip unreadable */
        }
      }
    }
  }
  await walk(repoDir, '');
  return out;
}

/**
 * OP-14: undo that also restores media/rels/styles/headers. Performs the same
 * branch rewind + forced checkout as undo(), then reads every tracked working-
 * tree file back so callers can re-inject them into the live OOXML package.
 */
export async function undoPackage(
  docPath: string
): Promise<{ xml: string; extra: Map<string, Buffer> }> {
  const repoDir = getRepoDir(docPath);
  // Same stray-`.git` refusal as the init/commit path: undo rewrites refs
  // and checks out the working tree, so a redirect must never be followed.
  await assertUsableSidecarGit(repoDir);
  let commits;
  try {
    commits = await git.log({ fs, dir: repoDir });
  } catch {
    throw new Error(
      `Cannot read history for '${docPath}': no history exists yet (open or save the document first).`
    );
  }

  if (commits.length <= 1) {
    throw new Error("Cannot undo: already at the initial state.");
  }

  const prevCommitOid = commits[1].oid;
  const currentBranch = await git.currentBranch({ fs, dir: repoDir }) || 'master';

  // Move the branch pointer back 1 commit
  await git.writeRef({
    fs,
    dir: repoDir,
    ref: `refs/heads/${currentBranch}`,
    value: prevCommitOid,
    force: true
  });

  // Checkout to sync the working directory
  await git.checkout({ fs, dir: repoDir, ref: currentBranch, force: true });

  const files = await collectRepoFiles(repoDir);
  const xmlBuf = files.get('document.xml');
  if (!xmlBuf) throw new Error('Undo failed: document.xml missing after checkout.');
  // OP-14: only tracked package files belong in extra — never leak
  // document.xml, .git internals, or stray working-tree files.
  const extra = new Map<string, Buffer>();
  for (const [rel, buf] of files) {
    if (rel === 'document.xml') continue;
    if (isTrackedPackageFile(rel)) extra.set(rel, buf);
  }
  return { xml: xmlBuf.toString('utf8'), extra };
}

export async function undo(docPath: string): Promise<string> {
  const { xml } = await undoPackage(docPath);
  return xml;
}

/**
 * Read-only checkout of an arbitrary commit by oid.
 * Never moves refs, never touches the working tree or HEAD — the caller
 * decides whether to restore the result via a new commit.
 * Reads the commit tree directly via readTree/readBlob so concurrent
 * checkouts are safe.
 */
export async function checkoutCommit(
  docPath: string,
  commitOid: string
): Promise<{ xml: string; extra: Map<string, Buffer> }> {
  if (typeof commitOid !== 'string' || commitOid.length === 0) {
    throw new Error(`Cannot checkout: commit ${commitOid} not found in history.`);
  }
  const repoDir = getRepoDir(docPath);
  await assertUsableSidecarGit(repoDir);
  let commits;
  try {
    commits = await git.log({ fs, dir: repoDir });
  } catch {
    throw new Error(
      `Cannot read history for '${docPath}': no history exists yet (open or save the document first).`
    );
  }
  const found = commits.some(c => c.oid === commitOid);
  if (!found) {
    throw new Error(`Cannot checkout: commit ${commitOid} not found in history.`);
  }

  // Walk the commit tree recursively, collecting blob paths.
  const blobs = new Map<string, string>();
  async function walkTree(treeOid: string, prefix: string): Promise<void> {
    let result: { tree: Array<{ path: string; oid: string; type: string }> };
    try {
      result = await git.readTree({ fs, dir: repoDir, oid: treeOid });
    } catch (err) {
      const cause = err instanceof Error ? err.message : String(err);
      throw new Error(
        `Checkout failed: cannot read tree '${treeOid}' at commit ${commitOid}: ${cause}`
      );
    }
    for (const entry of result.tree) {
      const fullPath = prefix ? `${prefix}/${entry.path}` : entry.path;
      if (entry.type === 'blob') {
        blobs.set(fullPath, entry.oid);
      } else if (entry.type === 'tree') {
        await walkTree(entry.oid, fullPath);
      }
    }
  }

  let rootTreeOid: string;
  try {
    const { commit } = await git.readCommit({ fs, dir: repoDir, oid: commitOid });
    rootTreeOid = (commit as { tree: string }).tree;
  } catch (err) {
    const cause = err instanceof Error ? err.message : String(err);
    throw new Error(
      `Checkout failed: cannot read commit ${commitOid}: ${cause}`
    );
  }
  await walkTree(rootTreeOid, '');

  const docBlobOid = blobs.get('document.xml');
  if (!docBlobOid) {
    throw new Error(`Checkout failed: document.xml missing at commit ${commitOid}.`);
  }
  let docBlob: { blob: Uint8Array };
  try {
    docBlob = await git.readBlob({ fs, dir: repoDir, oid: docBlobOid });
  } catch (err) {
    const cause = err instanceof Error ? err.message : String(err);
    throw new Error(
      `Checkout failed: cannot read blob 'document.xml' (${docBlobOid}) at commit ${commitOid}: ${cause}`
    );
  }
  const xml = Buffer.from(docBlob.blob).toString('utf8');

  const trackedEntries = [...blobs.entries()].filter(
    ([rel]) => rel !== 'document.xml' && isTrackedPackageFile(rel)
  );
  const extraBlobs = await Promise.all(
    trackedEntries.map(async ([rel, blobOid]) => {
      let blobResult: { blob: Uint8Array };
      try {
        blobResult = await git.readBlob({ fs, dir: repoDir, oid: blobOid });
      } catch (err) {
        const cause = err instanceof Error ? err.message : String(err);
        throw new Error(
          `Checkout failed: cannot read blob '${rel}' (${blobOid}) at commit ${commitOid}: ${cause}`
        );
      }
      return [rel, Buffer.from(blobResult.blob)] as const;
    })
  );
  const extra = new Map<string, Buffer>(extraBlobs);
  return { xml, extra };
}
