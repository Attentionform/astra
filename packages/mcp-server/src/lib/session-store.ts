import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as os from 'os';
import * as path from 'path';
import { randomUUID, createHash } from 'crypto';
import {
  type DocumentAST,
  type DocxParserMacroOptions,
  DocxParser,
  OperationEngine,
  ToonGenerator,
  canonicalizeForAsset as canonicalizeForMcp,
  getRepoDir,
  initAstraRepo,
  isPathInsideDir as isMcpPathInsideDir,
  isTrackedPackageFile,
  type StyleGuideConfig,
} from '@attentionform/astra-engine';

/**
 * Sprint-2 Muse session store (MCP Lean-6 backing).
 *
 * Sessions: Map<sessionId,DocHandle> (multi-doc; replaces the single active-document holder).
 * - LRU 10 docs, 200MB total cap, idle 30min evict.
 * - Per-session revision (bumped on every successful execute_batch).
 * - Per-session mutex (each DocHandle owns its lock; no shared global lock).
 *
 * There is intentionally no shared module-level instance here:
 * callers construct their own store via createMcpSessionStore().
 */

export const MCP_SESSION_MAX_DOCS = 10;
export const MCP_SESSION_MAX_BYTES = 200 * 1024 * 1024;
export const MCP_SESSION_IDLE_MS = 30 * 60 * 1000;

/**
 * Redo-stack quotas (per session, in-memory). High defaults so normal
 * revert/redo chains (1-5 entries, KBs) never hit them; only runaway
 * rapid-revert loops do. Configurable via store options or env.
 */
export const SESSION_REDO_MAX_DEPTH = 100;
export const SESSION_REDO_MAX_BYTES = 50 * 1024 * 1024;
/** Env overrides for the redo quotas (positive integers, bytes for the latter). */
export const SESSION_REDO_MAX_DEPTH_ENV = 'ASTRA_REDO_MAX_DEPTH';
export const SESSION_REDO_MAX_BYTES_ENV = 'ASTRA_REDO_MAX_BYTES';
/** L1 upper clamps: env/explicit values above these are capped (env-driven memory exhaustion). */
export const SESSION_REDO_MAX_DEPTH_CLAMP = 10_000;
export const SESSION_REDO_MAX_BYTES_CLAMP = 1024 * 1024 * 1024;
/** Error code for quota refusals (shared with history sidecar caps). */
export const SESSION_QUOTA_CODE = 'E_QUOTA' as const;

/** Prefix for every server-minted session id (`sess_<uuid>`). */
export const MCP_SESSION_ID_PREFIX = 'sess_';
/** CLI flag enabling deterministic client-supplied ids (tests only, default off). */
export const CLIENT_SESSION_ID_FLAG = '--allow-client-session-ids';
/** Env var mirroring the flag (accepted values: 1/true, case-insensitive). */
export const CLIENT_SESSION_ID_ENV = 'ASTRA_ALLOW_CLIENT_SESSION_IDS';

export interface McpSessionStoreOptions {
  maxDocs?: number;
  maxBytes?: number;
  idleMs?: number;
  /** Redo-stack depth cap (entries). Defaults to SESSION_REDO_MAX_DEPTH / env. */
  redoMaxDepth?: number;
  /** Redo-stack bytes cap (sum of snapshot bytes). Defaults to SESSION_REDO_MAX_BYTES / env. */
  redoMaxBytes?: number;
  /**
   * Hatch for deterministic ids in tests. Default off: client-supplied
   * sessionIds are rejected with E_VALIDATION and every open mints
   * `sess_<uuid>`. Also enabled via --allow-client-session-ids /
   * ASTRA_ALLOW_CLIENT_SESSION_IDS.
   */
  allowClientSessionIds?: boolean;
}

/** True when the deterministic-id hatch is enabled via env. */
export function isClientSessionIdHatchEnabledByEnv(env: NodeJS.ProcessEnv = process.env): boolean {
  const raw = (env[CLIENT_SESSION_ID_ENV] ?? '').trim().toLowerCase();
  return raw === '1' || raw === 'true';
}

/** True when argv carries --allow-client-session-ids before any `--` marker. */
export function hasClientSessionIdHatch(argv: readonly string[] = process.argv): boolean {
  for (const arg of argv) {
    if (arg === '--') break;
    if (arg === CLIENT_SESSION_ID_FLAG) return true;
  }
  return false;
}

/** Client-facing rejection for a squatted sessionId (E_VALIDATION + remedy). */
export function buildClientSessionIdRejection(
  tool: string,
  raw: unknown
): { code: 'E_VALIDATION'; message: string; remedy: string } {
  const shown = typeof raw === 'string' && raw.trim() ? raw.trim().slice(0, 64) : String(raw).slice(0, 64);
  return {
    code: 'E_VALIDATION',
    message: `${tool}: client-supplied sessionId "${shown}" rejected: session ids are server-minted (${MCP_SESSION_ID_PREFIX}<uuid>). Omit sessionId and use the returned id.`,
    remedy: `Omit sessionId and retry with the returned id; deterministic ids are test-only via ${CLIENT_SESSION_ID_FLAG} (default off).`,
  };
}

/** Resolves the redo depth cap (explicit > env > default). Never below 1. L1: clamped to SESSION_REDO_MAX_DEPTH_CLAMP. */
export function resolveRedoMaxDepth(
  explicit?: number,
  env: NodeJS.ProcessEnv = process.env
): number {
  if (typeof explicit === 'number' && Number.isFinite(explicit) && explicit >= 1) {
    return Math.min(Math.floor(explicit), SESSION_REDO_MAX_DEPTH_CLAMP);
  }
  const raw = (env[SESSION_REDO_MAX_DEPTH_ENV] ?? '').trim();
  const parsed = raw ? Number.parseInt(raw, 10) : NaN;
  if (Number.isFinite(parsed) && parsed >= 1) return Math.min(Math.floor(parsed), SESSION_REDO_MAX_DEPTH_CLAMP);
  return SESSION_REDO_MAX_DEPTH;
}

/** Resolves the redo bytes cap (explicit > env > default). Never below 1. L1: clamped to SESSION_REDO_MAX_BYTES_CLAMP. */
export function resolveRedoMaxBytes(
  explicit?: number,
  env: NodeJS.ProcessEnv = process.env
): number {
  if (typeof explicit === 'number' && Number.isFinite(explicit) && explicit >= 1) {
    return Math.min(Math.floor(explicit), SESSION_REDO_MAX_BYTES_CLAMP);
  }
  const raw = (env[SESSION_REDO_MAX_BYTES_ENV] ?? '').trim();
  const parsed = raw ? Number.parseInt(raw, 10) : NaN;
  if (Number.isFinite(parsed) && parsed >= 1) return Math.min(Math.floor(parsed), SESSION_REDO_MAX_BYTES_CLAMP);
  return SESSION_REDO_MAX_BYTES;
}

/** Byte size of one package snapshot (sum of part lengths). */
export function redoSnapshotByteSize(snap: Map<string, Buffer>): number {
  let total = 0;
  for (const buf of snap.values()) total += (buf as Buffer).length;
  return total;
}

/** Loud quota refusal naming the raise-the-limit remedy. */
function buildRedoQuotaError(kind: 'depth' | 'bytes', current: number, incoming: number, max: number): Error & { code: string; remedy: string } {
  const detail =
    kind === 'depth'
      ? `redo stack depth ${current}+${incoming} would exceed cap ${max} entries`
      : `redo stack bytes ${current}+${incoming} would exceed cap ${max} bytes`;
  return Object.assign(
    new Error(
      `${SESSION_QUOTA_CODE}: revert_document refused: ${detail}. ` +
        `Raise the limit via McpSessionStoreOptions redoMaxDepth/redoMaxBytes or ` +
        `${SESSION_REDO_MAX_DEPTH_ENV}/${SESSION_REDO_MAX_BYTES_ENV}.`
    ),
    {
      code: SESSION_QUOTA_CODE,
      remedy:
        `Redo history is capped at ${SESSION_REDO_MAX_DEPTH} entries / ${SESSION_REDO_MAX_BYTES} bytes by default. ` +
        `Raise the limit via McpSessionStoreOptions redoMaxDepth/redoMaxBytes or ` +
        `${SESSION_REDO_MAX_DEPTH_ENV}/${SESSION_REDO_MAX_BYTES_ENV}, or make a fresh edit to clear redo and retry.`,
    }
  );
}

/** Per-session async mutex (promise-chain, no global lock). */
export class SessionMutex {
  private tail: Promise<void> = Promise.resolve();
  private inflight = 0;

  public async runExclusive<T>(fn: () => Promise<T>): Promise<T> {
    this.inflight += 1;
    const prev = this.tail;
    let release: () => void = () => undefined;
    const next = new Promise<void>((resolve) => {
      release = () => resolve();
    });
    this.tail = next;
    await prev;
    try {
      return await fn();
    } finally {
      this.inflight -= 1;
      release();
    }
  }

  /**
   * M1: real locked state (queued + running).
   * True while any runExclusive section is queued or holding the lock.
   * Used by closeDocument to fail closed with E_CONFLICT instead of racing
   * a concurrent batch.
   */
  public isLocked(): boolean {
    return this.inflight > 0;
  }
}

export interface DocHandle {
  sessionId: string;
  filePath: string;
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
  toon: ToonGenerator;
  revision: number;
  /**
   * Phase C: revision at the last successful save_document. dirty =
   * revision !== savedRevision (in-memory session holds unsaved
   * execute_batch edits vs the last file written to disk). Set to 1 on
   * open, re-synced on every successful save.
   */
  savedRevision: number;
  byteSize: number;
  lastAccess: number;
  mutex: SessionMutex;
  /**
   * History redo stack: full package snapshots captured before each
   * revert_document restore (newest last). redo_document pops one entry,
   * restores it, and commits it as a new head. Cleared by the next
   * successful execute_batch (a fresh edit invalidates the redo chain) and
   * dropped with the session on close/evict.
   *
   * R-2 exactness: entries are immutable pristine copies (the stack
   * deep-copies every pushed map, so later live-package mutations can never
   * corrupt a stored source), each bound to the sha256 fingerprint recorded
   * at push time. A redo must restore EXACTLY the stored pre-revert bytes or
   * fail with E_STATE (see assertRedoSnapshotExact) — never a third mixed
   * state. The stack stays Array-compatible (push/pop/index/length) so the
   * existing session-boundary semantics (cleared on fresh edit, dies with
   * the session) are unchanged.
   */
  redoStack: SessionRedoStack;
  /**
   * Audit depth (in-memory, per-session): operations + touched node ids
   * accumulated since the last successful save/revert/redo. Consumed at the
   * next history commit to fill the entry's op count + affected node ids,
   * then reset by markSaved. Best-effort and bounded (never fails a batch);
   * it dies with the session and is NOT durable history (see the HOLD note
   * on export_audit_trail: .astra/ itself is deletable local state).
   */
  pendingAudit: { ops: number; nodeIds: string[] };
  /**
   * Attached style guide ("aesthetic skin"): set at session creation via
   * new_document/compose_document `styleGuide`, persists for the session,
   * mirrored onto `engine` so batches inherit it when they omit their own.
   * Per-batch `styleGuide` overrides apply to that batch only (restored
   * after). No detach API — open a new session to change guides.
   */
  styleGuide?: StyleGuideConfig;
  /** Raw guide ref the session guide was resolved from (name or file path). */
  styleGuideRef?: string;
}

/**
 * Per-commit audit metadata recorded alongside each history commit.
 * Keyed by commit oid in a store-level map (in-memory only: re-opens in a
 * fresh store lose pre-existing entries, and deleting .astra/ wipes the
 * underlying history itself — see the HOLD retention note).
 */
export interface CommitAuditMeta {
  author: string;
  /** Wall-clock epoch millis at commit time (git stamps seconds only). */
  timestampMs: number;
  /** Operations folded into this commit (saves: since last save; reverts: steps back; redos: 1). */
  opCount: number;
  /** Node ids touched by this commit (bounded, document order, best-effort). */
  affectedNodeIds: string[];
}

/** Cap on affected node ids retained per commit (bounded envelope). */
export const COMMIT_AUDIT_NODE_IDS_MAX = 100;
/** Cap on commit-meta entries retained per store (bounded ring, oldest dropped). */
export const COMMIT_AUDIT_META_MAX = 2000;

/** Error code for exact-state violations (redo/restore source diverged). */
export const REDO_EXACTNESS_CODE = 'E_STATE' as const;

/** sha256 hex over one package-part buffer (redo integrity binding). */
export function hashRedoBytes(buf: Buffer): string {
  return createHash('sha256').update(Buffer.from(buf)).digest('hex');
}

/** Integrity fingerprint of one full package snapshot (stable key order). */
export interface RedoSnapshotIntegrity {
  /** sha256 over word/document.xml bytes ('' when the part is absent). */
  docXmlHash: string;
  /** sha256 over every entry (name + bytes, sorted by name). */
  packageHash: string;
  /** Sorted entry names captured in the snapshot. */
  entryNames: string[];
  /** Total snapshot bytes across all entries. */
  byteSize: number;
}

/** Deep-copies a package snapshot (new map, copied buffers). */
export function clonePackageSnapshot(snap: Map<string, Buffer>): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  for (const [name, buf] of snap) out.set(name, Buffer.from(buf));
  return out;
}

/** Fingerprints a package snapshot without mutating it. */
export function fingerprintPackageSnapshot(snap: Map<string, Buffer>): RedoSnapshotIntegrity {
  const names = [...snap.keys()].sort();
  const h = createHash('sha256');
  let byteSize = 0;
  for (const name of names) {
    const buf = snap.get(name) as Buffer;
    byteSize += buf.length;
    h.update(Buffer.from(name, 'utf8'));
    h.update(Buffer.from([0]));
    h.update(buf);
    h.update(Buffer.from([1]));
  }
  const docXml = snap.get('word/document.xml');
  return {
    docXmlHash: docXml ? hashRedoBytes(docXml) : '',
    packageHash: h.digest('hex'),
    entryNames: names,
    byteSize,
  };
}

/**
 * R-2 session-bound redo stack.
 *
 * Array-compatible (push/pop/index/length behave like the plain array the
 * callers already use, including `length = 0` truncation to invalidate the
 * chain), with two hardening guarantees on top:
 * - push() stores an immutable pristine copy (later live-package mutations
 *   can never corrupt the redo source by aliasing);
 * - every stored copy is bound to the fingerprint recorded at push time,
 *   retrievable via integrityOf() for exactness proofs.
 *
 * Integrity metadata lives in a WeakMap keyed by the stored copy, so
 * `length = 0` truncation or close/evict simply orphans the metadata (no
 * parallel-array bookkeeping to go stale) and entries are never shared
 * back out for mutation.
 */
export class SessionRedoStack extends Array<Map<string, Buffer>> {
  private readonly integrity = new WeakMap<Map<string, Buffer>, RedoSnapshotIntegrity>();
  private readonly maxDepth: number;
  private readonly maxBytes: number;

  constructor(opts?: { maxDepth?: number; maxBytes?: number }) {
    super();
    this.maxDepth = resolveRedoMaxDepth(opts?.maxDepth);
    this.maxBytes = resolveRedoMaxBytes(opts?.maxBytes);
  }

  public getMaxDepth(): number {
    return this.maxDepth;
  }

  public getMaxBytes(): number {
    return this.maxBytes;
  }

  /** Sum of stored snapshot bytes (recomputed so pop/length-truncation stay exact). */
  public currentBytes(): number {
    let total = 0;
    for (const entry of this) {
      try {
        const known = this.integrity.get(entry);
        if (known && typeof known.byteSize === 'number') {
          total += known.byteSize;
          continue;
        }
      } catch {
        /* fall through to direct measure */
      }
      total += redoSnapshotByteSize(entry);
    }
    return total;
  }

  public override push(...items: Array<Map<string, Buffer>>): number {
    if (items.length === 0) return this.length;
    if (this.length + items.length > this.maxDepth) {
      throw buildRedoQuotaError('depth', this.length, items.length, this.maxDepth);
    }
    let incoming = 0;
    for (const item of items) incoming += redoSnapshotByteSize(item);
    const current = this.currentBytes();
    if (current + incoming > this.maxBytes) {
      throw buildRedoQuotaError('bytes', current, incoming, this.maxBytes);
    }
    for (const item of items) {
      const copy = clonePackageSnapshot(item);
      this.integrity.set(copy, fingerprintPackageSnapshot(copy));
      super.push(copy);
    }
    return this.length;
  }

  /** Newest entry without consuming (undefined when the stack is empty). */
  public peek(): Map<string, Buffer> | undefined {
    return this.length > 0 ? this[this.length - 1] : undefined;
  }

  /**
   * Push-time fingerprint of a stored entry (undefined for foreign maps
   * that bypassed push, e.g. entries built by hand in tests).
   */
  public integrityOf(entry: Map<string, Buffer>): RedoSnapshotIntegrity | undefined {
    try {
      return this.integrity.get(entry);
    } catch {
      return undefined;
    }
  }
}

/**
 * R-2 exact-or-error check: asserts that a restored live snapshot is
 * byte-identical to the redo source it must equal (pre-revert on success,
 * post-revert when the restore was rolled back). Throws E_STATE with a
 * remedy naming the divergent parts when exact restore is impossible, so
 * callers fail loudly instead of serving a third mixed state.
 */
export function assertRedoSnapshotExact(
  expected: Map<string, Buffer>,
  actual: Map<string, Buffer>,
  context: string
): void {
  const diffs: string[] = [];
  const names = new Set<string>([...expected.keys(), ...actual.keys()]);
  for (const name of [...names].sort()) {
    const a = expected.get(name);
    const b = actual.get(name);
    if (!a || !b) {
      diffs.push(`${name} (present ${!a ? 'only after restore' : 'only before restore'})`);
    } else if (Buffer.compare(a, b) !== 0) {
      diffs.push(name);
    }
    if (diffs.length >= 20) break;
  }
  if (diffs.length === 0) return;
  const shown = diffs.slice(0, 20).join(', ');
  const extra = names.size > diffs.length ? '' : ' (list truncated at 20)';
  throw Object.assign(
    new Error(
      `${REDO_EXACTNESS_CODE}: ${context}: restored state diverges from the redo source in ${diffs.length} part(s): ${shown}${extra}`
    ),
    {
      code: REDO_EXACTNESS_CODE,
      remedy:
        'The redo source no longer matches live state: run inspect_document and get_document_history to confirm ' +
        'live vs saved state, then retry revert_document to record a fresh redoable state instead of reusing the stale entry.',
    }
  );
}

export interface McpSessionStats {
  sessions: number;
  totalBytes: number;
  maxDocs: number;
  maxBytes: number;
  idleMs: number;
}

/**
 * M3: workspace jail for MCP document paths.
 * Mirrors the narrow sandbox (cwd/documents, tmpdir, app Drafts, extra roots)
 * without importing the template-palette owner's modules (zero overlap).
 * Every candidate is canonicalized (symlinks/junctions resolved, with
 * nearest-existing-ancestor fallback for not-yet-created save targets) and
 * containment is checked on the CANONICAL path.
 *
 * H2/D3: `ASTRA_ALLOWED_DIR` is NOT honored here (dropped: roots come only
 * from explicit `--allowed-dir` flags, which the server entrypoint folds into
 * `ASTRA_EXTRA_DOC_ROOTS`, plus the fixed base roots below). Any
 * `ASTRA_ALLOWED_DIR` value in the environment is ignored.
 * H2: delimiter-split `ASTRA_EXTRA_DOC_ROOTS` entries are existence-checked
 * (must resolve to an existing directory) before joining the jail — phantom
 * entries never widen anything. Primitives (`isMcpPathInsideDir`,
 * `canonicalizeForMcp`) are the Q0a shared module (`@attentionform/astra-engine`
 * sandbox/paths), so MCP/engine verdicts can never disagree on semantics.
 */
export function getMcpAllowedRoots(): string[] {
  const roots: string[] = [];
  try {
    roots.push(path.resolve(process.cwd(), 'documents'));
  } catch {
    /* ignore */
  }
  try {
    roots.push(path.resolve(os.tmpdir()));
  } catch {
    /* ignore */
  }
  try {
    roots.push(path.resolve(process.cwd(), 'documents', 'Drafts'));
  } catch {
    /* ignore */
  }
  try {
    roots.push(path.resolve(os.tmpdir(), 'ASTra'));
  } catch {
    /* ignore */
  }
  for (const p of getExistingExtraRoots()) roots.push(p);
  return [...new Set(roots)];
}

/**
 * H2: explicit-flag/env extra roots that actually exist as directories.
 * Non-existent entries (typos, stale mounts, phantom widening) are dropped
 * instead of joining the jail lexically.
 */
function getExistingExtraRoots(): string[] {
  const raw = (process.env.ASTRA_EXTRA_DOC_ROOTS || '').trim();
  if (!raw) return [];
  const out: string[] = [];
  for (const part of raw.split(path.delimiter)) {
    const p = part.trim();
    if (!p) continue;
    try {
      const resolved = path.resolve(p);
      const stat = fs.statSync(resolved);
      if (!stat.isDirectory()) continue;
      out.push(resolved);
    } catch {
      /* missing or unstatable — never joins the jail */
    }
  }
  return out;
}

// Q0a: jail primitives live in @attentionform/astra-engine sandbox/paths (imported above).
// Local isMcpPathInsideDir + canonicalizeForMcp removed; the imported
// isMcpPathInsideDir (alias of isPathInsideDir) + canonicalizeForMcp (alias
// of canonicalizeForAsset) are the single source of truth. Re-exported below
// for existing importers (tools/export.ts, tools/compose-document.ts).
export { canonicalizeForMcp };

function findMcpContainingRoot(canonical: string): string | undefined {
  for (const root of getMcpAllowedRoots()) {
    let canonicalRoot = root;
    try {
      canonicalRoot = fs.realpathSync(root);
    } catch {
      /* root may not exist (cwd/documents) — use lexical form */
    }
    if (isMcpPathInsideDir(canonical, canonicalRoot) || isMcpPathInsideDir(canonical, root)) {
      return root;
    }
  }
  return undefined;
}

function assertNoDotSegments(canonical: string, root: string): void {
  const rel = path.relative(root, canonical);
  const segs = rel.split(path.sep).filter((s) => s && s !== '.' && s !== '..');
  for (const seg of segs) {
    if (seg.startsWith('.')) {
      throw Object.assign(new Error(`Document path traverses a hidden directory "${seg}"`), {
        code: 'E_SANDBOX',
        remedy: 'Use a non-hidden path inside documents/ or the system temp dir.',
      });
    }
  }
  const base = path.basename(canonical);
  if (base.startsWith('.')) {
    throw Object.assign(new Error('Document path must not use a hidden (dot-prefixed) file name'), {
      code: 'E_SANDBOX',
      remedy: 'Rename the file without a leading dot.',
    });
  }
}

/**
 * M3: jailed document-path validator.
 * Returns the resolved absolute path. Throws E_SANDBOX (with remedy) on:
 * non-string/empty, null byte, line break, `..` traversal, hidden segments,
 * or canonical location outside the workspace roots. Throws E_VALIDATION on
 * non-.docx extension. Symlinks/junctions are resolved BEFORE containment.
 */
export function validateDocPath(rawPath: unknown): string {
  if (typeof rawPath !== 'string' || !rawPath.trim()) {
    throw Object.assign(new Error('Document path must be a non-empty string'), {
      code: 'E_SANDBOX',
      remedy: 'Pass a .docx path such as documents/report.docx.',
    });
  }
  if (rawPath.includes('\0') || /%00/i.test(rawPath)) {
    throw Object.assign(new Error('Document path contains a null byte'), {
      code: 'E_SANDBOX',
      remedy: 'Remove null bytes from the path and retry.',
    });
  }
  if (rawPath.includes('\n') || rawPath.includes('\r')) {
    throw Object.assign(new Error('Document path contains forbidden line break'), {
      code: 'E_SANDBOX',
      remedy: 'Remove line breaks from the path and retry.',
    });
  }
  const trimmed = rawPath.trim();
  if (trimmed.includes('\0') || /%00/i.test(trimmed)) {
    throw Object.assign(new Error('Document path contains a null byte'), {
      code: 'E_SANDBOX',
      remedy: 'Remove null bytes from the path and retry.',
    });
  }
  const segments = trimmed.split(/[\\/]/);
  if (segments.includes('..')) {
    throw Object.assign(new Error('Document path contains forbidden ".." traversal'), {
      code: 'E_SANDBOX',
      remedy: 'Use a path inside the workspace documents directory.',
    });
  }
  const resolved = path.resolve(trimmed);
  const resolvedSegments = resolved.split(path.sep);
  if (resolvedSegments.includes('..')) {
    throw Object.assign(new Error('Document path resolves outside its directory'), {
      code: 'E_SANDBOX',
      remedy: 'Use a path inside the workspace documents directory.',
    });
  }
  // Symlink escape closure: canonicalize (target if it exists, else nearest
  // existing ancestor) and check containment on the CANONICAL path.
  const canonical = canonicalizeForMcp(resolved);
  const containingRoot = findMcpContainingRoot(canonical);
  if (!containingRoot) {
    throw Object.assign(
      new Error(`Document path is outside the allowed directories (${getMcpAllowedRoots().join(', ')})`),
      {
        code: 'E_SANDBOX',
        remedy: 'Place the .docx under documents/ or the system temp dir, or set ASTRA_EXTRA_DOC_ROOTS.',
      }
    );
  }
  assertNoDotSegments(canonical, containingRoot);
  const ext = path.extname(trimmed).toLowerCase();
  if (ext !== '.docx') {
    throw Object.assign(new Error(`Document path must use a .docx extension (got "${ext || '(none)'}")`), {
      code: 'E_VALIDATION',
      remedy: 'Pass a .docx file path.',
    });
  }
  return resolved;
}

/**
 * M3: save-target validator (same jail as open).
 * Save targets that don't exist yet are canonicalized via their nearest
 * existing ancestor, so a symlinked parent dir pointing outside still fails.
 */
export function validateSaveTarget(rawPath: unknown): string {
  return validateDocPath(rawPath);
}

/**
 * L4: openBuffer label validator (jails the `.astra` git sidecar).
 * Absolute labels go through the full `validateDocPath` jail (E_SANDBOX on
 * escape, E_VALIDATION on non-.docx) so `/etc/x.docx` can never become a
 * sidecar write at `/etc/.astra/x.docx`. Relative memory labels are
 * sanitized to a bare `.docx` basename (`memory.docx` fallback) and never
 * produce an absolute sidecar path (git init skips non-absolute handles).
 */
export function validateOpenBufferLabel(rawLabel: unknown): string {
  const fallback = 'memory.docx';
  if (typeof rawLabel !== 'string' || !rawLabel.trim()) return fallback;
  const trimmed = rawLabel.trim();
  if (trimmed.includes('\0') || /%00/i.test(trimmed)) {
    throw Object.assign(new Error('openBuffer label contains a null byte'), {
      code: 'E_VALIDATION',
      remedy: 'Pass a plain memory label such as memory.docx.',
    });
  }
  if (path.isAbsolute(trimmed)) return validateDocPath(trimmed);
  if (trimmed.includes('\n') || trimmed.includes('\r')) return fallback;
  const base = path.basename(trimmed.replace(/\\/g, '/'));
  if (!base || base === '.' || base === '..') return fallback;
  const clean = base.slice(0, 128);
  return clean.toLowerCase().endsWith('.docx') ? clean : fallback;
}

/**
 * Collects the versioned package extras (media/rels/styles/headers) for git.
 * document.xml is committed separately; everything matching
 * isTrackedPackageFile is snapshotted alongside it.
 */
export function collectTrackedExtrasForGit(parser: DocxParser): Record<string, Buffer> {
  try {
    const snap: Map<string, Buffer> = parser.capturePackageSnapshot();
    const extras: Record<string, Buffer> = {};
    for (const [name, buf] of snap) {
      if (isTrackedPackageFile(name)) extras[name] = buf;
    }
    return extras;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error(`[collectTrackedExtrasForGit] snapshot failed: ${msg}`);
    return {};
  }
}

/**
 * Ensures the sidecar git repo exists for `docPath`, creating it with the
 * current package state as the initial baseline when absent.
 * Returns true when the repo already existed, false when freshly created.
 * Best-effort at open (callers swallow failures so open semantics never
 * change); hard requirement before commitChange at save/revert/redo.
 */
export async function ensureGitRepo(
  docPath: string,
  xml: string,
  extras: Record<string, Buffer>
): Promise<boolean> {
  const repoDir = getRepoDir(docPath);
  let hasGit = false;
  try {
    hasGit = fs.existsSync(path.join(repoDir, '.git'));
  } catch {
    hasGit = false;
  }
  if (hasGit) return true;
  await initAstraRepo(docPath, xml, extras);
  return false;
}

/** Best-effort git init for a freshly opened handle (never throws). */
async function initGitForHandleBestEffort(handle: DocHandle): Promise<void> {
  try {
    if (!path.isAbsolute(handle.filePath)) return;
    const xml = handle.parser.readDocumentXml();
    await ensureGitRepo(handle.filePath, xml, collectTrackedExtrasForGit(handle.parser));
  } catch {
    /* open succeeds regardless — history tools report the gap */
  }
}

/** Outcome of McpSessionStore.closeDocument (dirty state read under the delete lock). */
export interface CloseDocumentOutcome {
  closed: boolean;
  /** True when the dropped session held unsaved edits (discarded with the close). */
  wasDirty: boolean;
  /** Pending revisions discarded with the close (0 when clean). */
  discardedRevisions: number;
}

export class McpSessionStore {
  private readonly sessions = new Map<string, DocHandle>();
  private readonly maxDocs: number;
  private readonly maxBytes: number;
  private readonly idleMs: number;
  private readonly allowClientSessionIds: boolean;
  private readonly redoMaxDepth: number;
  private readonly redoMaxBytes: number;
  /**
   * Phase C-4: ring of sessionIds dropped by eviction (LRU/byte-cap/idle),
   * so later unknown-session failures can report a capacity-specific
   * E_SESSION remedy instead of a generic unknown-id error. Explicit
   * close_document deletes are NOT recorded here. Bounded (last 200).
   */
  private readonly evictionHistory: string[] = [];
  /** Ids evicted by the most recent open (drained by handleOpenDocument for `evicted[]`). */
  private lastOpenEvicted: string[] = [];
  /**
   * Store-level commit audit metadata (oid -> meta). In-memory only: fresh
   * stores start empty and .astra/ deletion wipes the history it annotates.
   */
  private readonly commitMeta = new Map<string, CommitAuditMeta>();

  constructor(opts?: McpSessionStoreOptions) {
    this.maxDocs = opts?.maxDocs ?? MCP_SESSION_MAX_DOCS;
    this.maxBytes = opts?.maxBytes ?? MCP_SESSION_MAX_BYTES;
    this.idleMs = opts?.idleMs ?? MCP_SESSION_IDLE_MS;
    this.allowClientSessionIds = opts?.allowClientSessionIds ?? isClientSessionIdHatchEnabledByEnv();
    this.redoMaxDepth = resolveRedoMaxDepth(opts?.redoMaxDepth);
    this.redoMaxBytes = resolveRedoMaxBytes(opts?.redoMaxBytes);
  }

  public getAllowClientSessionIds(): boolean {
    return this.allowClientSessionIds;
  }

  public getMaxDocs(): number {
    return this.maxDocs;
  }

  public getMaxBytes(): number {
    return this.maxBytes;
  }

  public getIdleMs(): number {
    return this.idleMs;
  }

  public getRedoMaxDepth(): number {
    return this.redoMaxDepth;
  }

  public getRedoMaxBytes(): number {
    return this.redoMaxBytes;
  }

  /** Fresh redo stack bound to this store's quotas. */
  public createRedoStack(): SessionRedoStack {
    return new SessionRedoStack({ maxDepth: this.redoMaxDepth, maxBytes: this.redoMaxBytes });
  }

  public size(): number {
    return this.sessions.size;
  }

  public totalBytes(): number {
    let total = 0;
    for (const h of this.sessions.values()) total += h.byteSize;
    return total;
  }

  public getStats(): McpSessionStats {
    return {
      sessions: this.sessions.size,
      totalBytes: this.totalBytes(),
      maxDocs: this.maxDocs,
      maxBytes: this.maxBytes,
      idleMs: this.idleMs,
    };
  }

  public listSessionIds(): string[] {
    return [...this.sessions.keys()];
  }

  /** Evicts sessions idle longer than idleMs. Locked/dirty sessions are refused (never dropped). Returns evicted ids. */
  public evictExpired(now: number = Date.now()): string[] {
    const evicted: string[] = [];
    for (const [id, h] of [...this.sessions.entries()]) {
      if (now - h.lastAccess > this.idleMs) {
        if (h.mutex.isLocked()) continue;
        if (this.isDirty(h)) continue;
        this.sessions.delete(id);
        evicted.push(id);
      }
    }
    this.recordEvicted(evicted);
    return evicted;
  }

  /** Phase C-4: record evicted ids (bounded ring) for capacity-aware remedies. */
  private recordEvicted(ids: string[]): void {
    for (const id of ids) {
      if (!this.evictionHistory.includes(id)) this.evictionHistory.push(id);
    }
    while (this.evictionHistory.length > 200) this.evictionHistory.shift();
  }

  /** Phase C-4: true when `id` was dropped by eviction (not explicit close). */
  public wasEvicted(sessionId: string): boolean {
    return this.evictionHistory.includes(sessionId);
  }

  /**
   * Phase C-4: ids evicted by the most recent openDocument/openBuffer call.
   * Drained (cleared) on read so each open reports only its own evictions.
   */
  public drainLastOpenEvicted(): string[] {
    const out = [...this.lastOpenEvicted];
    this.lastOpenEvicted = [];
    return out;
  }

  /** Phase C-1b: true when the in-memory session holds unsaved edits. */
  public isDirty(handle: DocHandle): boolean {
    return handle.revision !== handle.savedRevision;
  }

  /** Phase C-1b: re-syncs the saved marker after a successful save_document. */
  public markSaved(handle: DocHandle): void {
    handle.savedRevision = handle.revision;
    handle.lastAccess = Date.now();
    this.resetPendingAudit(handle);
  }

  /**
   * Audit depth: appends one batch's footprint to the since-save accumulator.
   * Dedupes ids, preserves first-seen order, caps the list (best-effort:
   * never throws, so a batch can never fail on audit bookkeeping).
   */
  public accumulatePendingAudit(handle: DocHandle, appliedCount: number, nodeIds: readonly string[]): void {
    try {
      handle.pendingAudit.ops += Math.max(0, Math.floor(appliedCount));
      for (const id of nodeIds) {
        if (typeof id !== 'string' || !id) continue;
        if (handle.pendingAudit.nodeIds.length >= COMMIT_AUDIT_NODE_IDS_MAX) break;
        if (!handle.pendingAudit.nodeIds.includes(id)) handle.pendingAudit.nodeIds.push(id);
      }
    } catch {
      /* audit bookkeeping must never fail a mutation */
    }
  }

  /** Clears the since-save audit accumulator (called by markSaved). */
  public resetPendingAudit(handle: DocHandle): void {
    try {
      handle.pendingAudit.ops = 0;
      handle.pendingAudit.nodeIds = [];
    } catch {
      /* never throws */
    }
  }

  /** Records per-commit audit metadata (bounded ring; oldest entries dropped). */
  public recordCommitMeta(oid: string, meta: CommitAuditMeta): void {
    try {
      if (typeof oid !== 'string' || !oid) return;
      this.commitMeta.delete(oid);
      this.commitMeta.set(oid, {
        author: String(meta.author ?? ''),
        timestampMs: Number.isFinite(meta.timestampMs) ? meta.timestampMs : Date.now(),
        opCount: Math.max(0, Math.floor(meta.opCount ?? 0)),
        affectedNodeIds: Array.isArray(meta.affectedNodeIds)
          ? meta.affectedNodeIds.filter((v): v is string => typeof v === 'string' && !!v).slice(0, COMMIT_AUDIT_NODE_IDS_MAX)
          : [],
      });
      while (this.commitMeta.size > COMMIT_AUDIT_META_MAX) {
        const oldest = this.commitMeta.keys().next();
        if (oldest.done) break;
        this.commitMeta.delete(oldest.value);
      }
    } catch {
      /* never throws */
    }
  }

  /** Reads per-commit audit metadata (undefined for commits made before tracking). */
  public getCommitMeta(oid: string): CommitAuditMeta | undefined {
    try {
      return this.commitMeta.get(oid);
    } catch {
      return undefined;
    }
  }

  /** LRU + byte-cap eviction (oldest insertion first). Never evicts `protectId`, locked, or dirty sessions. Returns evicted ids. */
  private evictToFit(protectId?: string): string[] {
    const evicted: string[] = [];
    while (
      this.sessions.size > this.maxDocs ||
      this.totalBytes() > this.maxBytes
    ) {
      const oldest = [...this.sessions.keys()].find((k) => {
        if (k === protectId) return false;
        const h = this.sessions.get(k);
        if (!h) return false;
        if (h.mutex.isLocked()) return false;
        if (this.isDirty(h)) return false;
        return true;
      });
      if (!oldest) break;
      // If only the protected entry remains and we still overflow, the single
      // document itself exceeds the cap — leave it (open fails closed below
      // for that case instead of evicting the just-opened doc).
      if (this.sessions.size === 1 && oldest === protectId) break;
      this.sessions.delete(oldest);
      evicted.push(oldest);
    }
    this.recordEvicted(evicted);
    return evicted;
  }

  private touch(handle: DocHandle): void {
    handle.lastAccess = Date.now();
    // Refresh LRU order: re-insert as most-recently-used.
    this.sessions.delete(handle.sessionId);
    this.sessions.set(handle.sessionId, handle);
  }

  private newSessionId(): string {
    try {
      return `${MCP_SESSION_ID_PREFIX}${randomUUID()}`;
    } catch {
      return `${MCP_SESSION_ID_PREFIX}${Date.now().toString(36)}_${Math.floor(Math.random() * 1e6)}`;
    }
  }

  /** Mints an id no live session holds (retries on the near-impossible clash). */
  private mintUniqueSessionId(): string {
    for (let attempt = 0; attempt < 10; attempt++) {
      const id = this.newSessionId();
      if (!this.sessions.has(id)) return id;
    }
    throw Object.assign(new Error('open_document: failed to mint a unique session id (store saturated)'), {
      code: 'E_CONFLICT',
      remedy: 'Close idle sessions via close_document, then retry open_document.',
    });
  }

  /**
   * Resolves the session id for an open: omitted/blank mints a fresh
   * `sess_<uuid>`; a supplied id is rejected with E_VALIDATION unless the
   * deterministic-id hatch is on, and an id already held by a live session
   * fails with E_CONFLICT — the live entry is never replaced.
   */
  private resolveOpenSessionId(rawSessionId: unknown, tool: string): string {
    const supplied =
      typeof rawSessionId === 'string' && rawSessionId.trim() ? rawSessionId.trim().slice(0, 128) : undefined;
    if (supplied === undefined) return this.mintUniqueSessionId();
    if (!this.allowClientSessionIds) {
      const rej = buildClientSessionIdRejection(tool, rawSessionId);
      throw Object.assign(new Error(rej.message), { code: rej.code, remedy: rej.remedy });
    }
    const existing = this.sessions.get(supplied);
    if (existing && existing.mutex.isLocked()) {
      throw Object.assign(
        new Error(`open_document: session "${supplied}" is busy (concurrent batch in progress)`),
        {
          code: 'E_CONFLICT',
          remedy: 'Wait for the in-flight batch to finish, then retry open_document.',
        }
      );
    }
    if (existing) {
      throw Object.assign(new Error(`open_document: session "${supplied}" is already in use`), {
        code: 'E_CONFLICT',
        remedy:
          'That session id is held by a live session (never replaced): omit sessionId to mint a fresh id, or close_document the live session first.',
      });
    }
    return supplied;
  }

  /**
   * Opens a .docx file into a new (or reused) session.
   * Revision starts at 1. LRU/byte/idle eviction runs before insert.
   * Macro/active content refuses by default (E_VALIDATION naming the parts);
   * pass {allowMacros:true} to preserve or {stripMacros:true} to strip.
   */
  public async openDocument(
    filePath: string,
    sessionId?: string,
    macroOpts?: DocxParserMacroOptions
  ): Promise<DocHandle> {
    const cleanPath = validateDocPath(filePath);
    const expiredEvicted = this.evictExpired();
    let buf: Buffer;
    try {
      buf = await fsp.readFile(cleanPath);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      throw Object.assign(new Error(`open_document: cannot read "${cleanPath}": ${msg}`), {
        code: 'E_IO',
        remedy: 'Check the path exists and is readable, then retry open_document.',
      });
    }
    if (buf.length > this.maxBytes) {
      throw Object.assign(
        new Error(`open_document: document ${buf.length} bytes exceeds session cap ${this.maxBytes} bytes`),
        {
          code: 'E_QUOTA',
          remedy: 'Split the document or raise the store cap; single documents above 200MB are rejected.',
        }
      );
    }
    const parser = new DocxParser(buf, macroOpts);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const toon = new ToonGenerator();
    const id = this.resolveOpenSessionId(sessionId, 'open_document');
    const handle: DocHandle = {
      sessionId: id,
      filePath: cleanPath,
      parser,
      ast,
      engine,
      toon,
      revision: 1,
      savedRevision: 1,
      byteSize: buf.length,
      lastAccess: Date.now(),
      mutex: new SessionMutex(),
      redoStack: this.createRedoStack(),
      pendingAudit: { ops: 0, nodeIds: [] },
    };
    this.sessions.set(id, handle);
    // Git backing: initial baseline commit for history/revert/redo tools.
    await initGitForHandleBestEffort(handle);
    // Phase C-4: surface LRU/byte-cap (+ pre-open idle) evictions on the
    // open response so callers learn which sessions died for capacity.
    this.lastOpenEvicted = [...expiredEvicted, ...this.evictToFit(id)];
    // If the just-inserted entry did not survive (should not happen given the
    // single-doc cap check), fail closed instead of returning a dead handle.
    const live = this.sessions.get(id);
    if (!live) {
      throw Object.assign(new Error('open_document: session evicted immediately (store over capacity)'), {
        code: 'E_SESSION',
        remedy: `Session store is at capacity (${this.maxDocs} documents / ${this.maxBytes} bytes max). Close idle sessions via close_document, then retry open_document.`,
      });
    }
    // Fail-closed when eviction could not make room because every other
    // session is locked or holds unsaved edits: refuse the new open rather
    // than silently dropping uncommitted work.
    if (this.sessions.size > this.maxDocs || this.totalBytes() > this.maxBytes) {
      this.sessions.delete(id);
      throw Object.assign(
        new Error(
          `${SESSION_QUOTA_CODE}: open_document refused: session store at capacity ` +
            `(${this.sessions.size}/${this.maxDocs} documents, ${this.totalBytes()}/${this.maxBytes} bytes) ` +
            `and all other sessions are busy or hold unsaved edits. ` +
            `Raise the limit via McpSessionStoreOptions maxDocs/maxBytes or close a clean session first.`
        ),
        {
          code: SESSION_QUOTA_CODE,
          remedy:
            `Session store is at capacity and no evictable session exists (all others are busy or dirty). ` +
            `Close an idle clean session via close_document and retry open_document, or ` +
            `raise the limit via McpSessionStoreOptions maxDocs/maxBytes.`,
        }
      );
    }
    return live;
  }

  /**
   * Opens a document from an in-memory buffer (tests / headless flows).
   * `labelPath` is the display path recorded on the handle. L4: absolute
   * labels are jailed through validateDocPath (same jail as open) so a
   * future production caller cannot turn the `.astra` git sidecar into a
   * write primitive outside the jail (e.g. label `/etc/x.docx` → sidecar
   * `/etc/.astra/x.docx`). Relative memory labels (e.g. `memory.docx`) are
   * sanitized to a bare basename and never touch git (init skips non-absolute
   * paths). Macro/active content refuses by default; pass {allowMacros:true}
   * or {stripMacros:true}.
   */
  public async openBuffer(
    buf: Buffer,
    labelPath = 'memory.docx',
    sessionId?: string,
    macroOpts?: DocxParserMacroOptions
  ): Promise<DocHandle> {
    if (!Buffer.isBuffer(buf) || buf.length === 0) {
      throw Object.assign(new Error('openBuffer requires a non-empty Buffer'), {
        code: 'E_VALIDATION',
        remedy: 'Pass a valid .docx buffer.',
      });
    }
    if (buf.length > this.maxBytes) {
      throw Object.assign(new Error(`openBuffer: ${buf.length} bytes exceeds cap ${this.maxBytes}`), {
        code: 'E_QUOTA',
        remedy: 'Use a smaller document.',
      });
    }
    // L4: jail the label before it can become a sidecar write path.
    const safeLabel = validateOpenBufferLabel(labelPath);
    const expiredEvicted = this.evictExpired();
    const parser = new DocxParser(Buffer.from(buf), macroOpts);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const toon = new ToonGenerator();
    const id = this.resolveOpenSessionId(sessionId, 'open_document');
    const handle: DocHandle = {
      sessionId: id,
      filePath: safeLabel,
      parser,
      ast,
      engine,
      toon,
      revision: 1,
      savedRevision: 1,
      byteSize: buf.length,
      lastAccess: Date.now(),
      mutex: new SessionMutex(),
      redoStack: this.createRedoStack(),
      pendingAudit: { ops: 0, nodeIds: [] },
    };
    this.sessions.set(id, handle);
    // Git backing when the label is a real absolute path (memory labels skip).
    await initGitForHandleBestEffort(handle);
    this.lastOpenEvicted = [...expiredEvicted, ...this.evictToFit(id)];
    const live = this.sessions.get(id);
    if (!live) {
      throw Object.assign(new Error('openBuffer: session evicted immediately (store over capacity)'), {
        code: 'E_SESSION',
        remedy: `Session store is at capacity (${this.maxDocs} documents / ${this.maxBytes} bytes max). Close idle sessions via close_document and retry.`,
      });
    }
    if (this.sessions.size > this.maxDocs || this.totalBytes() > this.maxBytes) {
      this.sessions.delete(id);
      throw Object.assign(
        new Error(
          `${SESSION_QUOTA_CODE}: open_document refused: session store at capacity ` +
            `(${this.sessions.size}/${this.maxDocs} documents, ${this.totalBytes()}/${this.maxBytes} bytes) ` +
            `and all other sessions are busy or hold unsaved edits. ` +
            `Raise the limit via McpSessionStoreOptions maxDocs/maxBytes or close a clean session first.`
        ),
        {
          code: SESSION_QUOTA_CODE,
          remedy:
            `Session store is at capacity and no evictable session exists (all others are busy or dirty). ` +
            `Close an idle clean session via close_document and retry open_document, or ` +
            `raise the limit via McpSessionStoreOptions maxDocs/maxBytes.`,
        }
      );
    }
    return live;
  }

  /** Returns the live handle (refreshes LRU). Throws E_SESSION with remedy when missing/expired. */
  public get(sessionId: string): DocHandle {
    const h = this.sessions.get(sessionId);
    if (!h) {
      // Phase C-4: capacity-specific remedy when the id died to eviction
      // (mentions the doc cap + the action to take), generic otherwise.
      if (this.wasEvicted(sessionId)) {
        throw Object.assign(
          new Error(`Session "${String(sessionId)}" was evicted (store holds at most ${this.maxDocs} documents / ${this.maxBytes} bytes)`),
          {
            code: 'E_SESSION',
            remedy: `The session was dropped by LRU/idle eviction (${this.maxDocs}-doc cap). Call open_document again to re-open the file, and close idle sessions via close_document to free capacity.`,
          }
        );
      }
      throw Object.assign(new Error(`Unknown session "${String(sessionId)}" (expired or never opened)`), {
        code: 'E_SESSION',
        remedy: 'Call open_document first, then retry with the returned sessionId.',
      });
    }
    if (Date.now() - h.lastAccess > this.idleMs) {
      // Never silently drop uncommitted work or race a live batch: locked or
      // dirty sessions survive idle expiry (access refreshes them).
      if (h.mutex.isLocked() || this.isDirty(h)) {
        this.touch(h);
        return h;
      }
      this.sessions.delete(sessionId);
      this.recordEvicted([sessionId]);
      throw Object.assign(new Error(`Session "${String(sessionId)}" expired after 30min idle and was evicted`), {
        code: 'E_SESSION',
        remedy: 'Call open_document again to re-open the file.',
      });
    }
    this.touch(h);
    return h;
  }

  public has(sessionId: string): boolean {
    return this.sessions.has(sessionId);
  }

  /**
   * Runs `fn` under the handle's own mutex (per-session, never global).
   * Serializes concurrent execute_batch/save races on the same session while
   * leaving other sessions fully parallel.
   * H2: re-validates liveness AFTER acquiring the lock — a concurrent
   * closeDocument that won the race deletes first, and this call fails with
   * E_CONFLICT (not use-after-close).
   */
  public async withSession<T>(sessionId: string, fn: (handle: DocHandle) => Promise<T>): Promise<T> {
    const handle = this.get(sessionId);
    return handle.mutex.runExclusive(async () => {
      const live = this.sessions.get(sessionId);
      if (live !== handle) {
        throw Object.assign(new Error(`Session "${String(sessionId)}" was closed concurrently`), {
          code: 'E_CONFLICT',
          remedy: 'Call open_document again to re-open the file, then retry.',
        });
      }
      const out = await fn(handle);
      handle.lastAccess = Date.now();
      return out;
    });
  }

  /** Bumps the per-session revision after a successful mutation. */
  public bumpRevision(handle: DocHandle): number {
    handle.revision += 1;
    handle.lastAccess = Date.now();
    return handle.revision;
  }

  /** Best-effort byte-size refresh after mutations (drives the 200MB cap). */
  public refreshByteSize(handle: DocHandle): number {
    try {
      const buf = handle.parser.rezip();
      handle.byteSize = buf.length;
    } catch {
      // Size refresh must never fail a batch — eviction uses last known size.
    }
    return handle.byteSize;
  }

  /**
   * H2: lock-aware close. Acquires the SAME per-session SessionMutex instead
   * of deleting without a lock (which raced withSession's critical section).
   * - Unknown id -> { closed: false } (already closed/expired; adapter maps to E_SESSION).
   * - Session busy (batch/save/inspect holding the lock) -> throws E_CONFLICT
   *   with remedy (caller retries after the in-flight op finishes).
   * - Otherwise the delete itself runs under the mutex, so a withSession that
   *   queued between the busy-check and the delete either sees the live handle
   *   (if close lost) or fails its post-lock liveness check (if close won).
   * The outcome names the dropped dirty state (wasDirty + discardedRevisions,
   * read under the same lock as the delete) so callers never silently lose
   * unsaved edits: a dirty close warns instead of vanishing quietly.
   */
  public async closeDocument(sessionId: string): Promise<CloseDocumentOutcome> {
    const handle = this.sessions.get(sessionId);
    if (!handle) return { closed: false, wasDirty: false, discardedRevisions: 0 };
    if (handle.mutex.isLocked()) {
      throw Object.assign(
        new Error(`close_document: session "${String(sessionId)}" is busy (concurrent batch in progress)`),
        {
          code: 'E_CONFLICT',
          remedy: 'Wait for the in-flight batch to finish, then retry close_document.',
        }
      );
    }
    return handle.mutex.runExclusive(async () => {
      const live = this.sessions.get(sessionId);
      if (live !== handle) return { closed: false, wasDirty: false, discardedRevisions: 0 };
      const wasDirty = this.isDirty(handle);
      const discardedRevisions = wasDirty ? Math.max(0, handle.revision - handle.savedRevision) : 0;
      this.sessions.delete(sessionId);
      return { closed: true, wasDirty, discardedRevisions };
    });
  }

  public clear(): void {
    this.sessions.clear();
  }
}

/** Factory (no singleton): each server/test constructs its own store. */
export function createMcpSessionStore(opts?: McpSessionStoreOptions): McpSessionStore {
  return new McpSessionStore(opts);
}
