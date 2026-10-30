/**
 * Cross-process LibreOffice slot gate for tests.
 *
 * Only this: at most SLOT_COUNT headless LibreOffice conversions run at once
 * across every vitest worker on this box. LibreOffice is a heavyweight
 * single-shot process (measured here: 13-18s quiet for one small document,
 * 20-109s with several running at once, plus multi-GB RAM pressure that
 * starves the rest of the suite). Letting every worker spawn its own copy
 * turns per-test timeouts into lottery tickets, so LO-backed tests take a
 * slot around the call that spawns LibreOffice and hand it back right after.
 *
 * What this is NOT: it never retries the conversion, never swallows errors,
 * and never extends a timeout. A genuinely broken render still fails loudly
 * inside the test's own timeout. It only removes the contention variable.
 *
 * Locking: one lock dir per slot under os.tmpdir(), taken with an atomic
 * mkdir. A holder refreshes its heartbeat (see below) while it works; a
 * lock whose recorded holder PIDs are all dead is reclaimed immediately,
 * and a lock with a stale heartbeat and no holder progress is reclaimed, so
 * a killed holder delays the suite instead of deadlocking it. There is no
 * overall acquire deadline on purpose: the test timeout stays the single
 * budget, so a stuck queue surfaces as the usual timeout failure, not a new
 * error kind.
 *
 * Ownership / fencing model (read before touching this file):
 *
 * - Each acquisition mints a unique token (randomUUID) stored in the
 *   canonical holder record `holder.json` inside the lock dir, written
 *   ATOMICALLY via tmp-file + rename so concurrent readers only ever see
 *   the old or the new bytes, never a torn middle (D1: a torn write used
 *   to parse as "malformed" and trigger a steal of a live lock).
 * - Release is best-effort compare-and-delete (NOT atomic): it reads the
 *   canonical holder record and removes the lock dir ONLY if that read
 *   still carries our token. When no reclaimer lands between that read
 *   and the rm, a successor's fresh lock survives (D2: the old
 *   unconditional rm -rf let a stale victim wipe its successor, after
 *   which two holders shared one slot). The read-then-rm window that
 *   remains is residual 6 below.
 * - Heartbeat is a best-effort compare-and-swap (NOT atomic): read the
 *   canonical bytes, compute the refreshed record, then re-read and write
 *   ONLY if the bytes are unchanged since the first read; any change seen
 *   at the re-read aborts the beat. The re-read and the write are NOT one
 *   atomic step, and the post-write re-read only catches a successor write
 *   that landed after ours -- it does NOT catch a reclaimer that lands
 *   between our verify-read and our write (residual 5 below). Outside that
 *   window a reclaimed holder stops beating instead of resurrecting its
 *   old token over a successor (which used to yield two holders until the
 *   next beat). The dir mtime is refreshed only after a successful swap.
 * - Holder self-check: withLoSlot confirms the lock token is still ours
 *   BEFORE invoking the guarded operation, and every beat re-verifies
 *   before writing. Any fencing loss aborts the hold LOUDLY with an
 *   `LO slot fencing violation` error -- the guarded call never runs
 *   unlocked, and a mid-hold fencing loss rejects instead of returning
 *   the call's result.
 * - Reclaim semantics (D3): EITHER every recorded holder PID is dead
 *   (kill(pid,0) -> ESRCH), in which case reclaim immediately regardless
 *   of age; OR the heartbeat is older than STALE_MS AND the canonical
 *   holder bytes are identical across two reads straddling the decision
 *   (holder made no progress), in which case reclaim EVEN IF a recorded
 *   PID looks alive. Rationale for the second branch: PIDs recycle, so a
 *   dead holder's PID can be reused by an unrelated live process; gating
 *   on PID liveness alone would then suppress reclaim until that
 *   unrelated process exits. A genuinely live holder rewrites its
 *   heartbeat every HEARTBEAT_MS (30s) << STALE_MS (10min), so a stale,
 *   byte-unchanged heartbeat proves no progress; worst case we reclaim
 *   from a live-but-wedged (>10min without a beat) holder whose own test
 *   timeout (30s) fired long ago. Conversely a live holder that beats is
 *   never stolen: any beat within STALE_MS refreshes the timestamp and
 *   aborts the slow path, and any ownership change flips the bytes and
 *   aborts the unchanged check.
 * - Malformed canonical holder records are NEVER fast-path reclaimed:
 *   without a valid PID no liveness check is possible, and a torn-looking
 *   file with a fresh dir may still have a live owner. Only the stale +
 *   unchanged backstop may reclaim them.
 *
 * Known-incomplete test gate -- NEVER copy to production. Accepted
 * residuals (explicitly out of scope):
 * 1. No true atomic hold across check+spawn: mkdir + read + write is not
 *    kernel-arbitrated (needs flock -- unavailable). Consequence is bounded
 *    to extra LO concurrency in TESTS only, never wrong results: the
 *    post-fn ownership re-check makes a fenced holder throw instead of
 *    reporting success-while-fenced.
 * 2. Detection latency ~HEARTBEAT_MS (default 30s): a fencing loss between
 *    beats is noticed only at the next beat or the post-fn re-check.
 *    Heartbeat-interval trade-off, accepted for test load.
 * 3. Two-waiter same-stale race (narrowed, not closed): two waiters can
 *    both pass the unchanged check on the same stale lock and both run
 *    rm+mkdir+write; the last writer wins and the loser still believes it
 *    holds the slot for one hold cycle. Window is milliseconds wide,
 *    requires a 10-minute-stale lock plus two simultaneous waiters, and the
 *    post-write verify in reclaimAndTake narrows it further (a loser whose
 *    write was already wiped backs off). Closing it fully needs a
 *    kernel-arbitrated primitive, not mkdir/rm.
 * 4. PID-reuse backstop stall: a dead holder's PID recycled by an unrelated
 *    live process looks alive, so the dead-PID fast path stays shut and
 *    reclaim waits for the stale + unchanged backstop (up to STALE_MS).
 *    Accepted stall in TESTS only; liveness is a hint, the fencing token
 *    is the authority.
 * 5. Heartbeat verify-then-write resurrection (NOT closed): a reclaimer
 *    that lands between the heartbeat verify-read and the holder write
 *    gets clobbered -- the stale holder's write overwrites the successor's
 *    fresh record, and the post-write check then reads back the stale
 *    holder's own bytes so it keeps beating. Consequence is bounded to
 *    extra concurrent LibreOffice in TESTS only, never wrong results:
 *    the post-fn ownership re-check still makes a fenced holder throw
 *    instead of reporting success-while-fenced.
 * 6. Release read-then-rm successor delete (NOT closed): a reclaimer that
 *    lands between the release token read and the rm deletes the
 *    successor's fresh lock dir out from under it. Consequence is bounded
 *    the same way: extra concurrent LibreOffice in TESTS only, never
 *    wrong results, via the same post-fn fencing check.
 */
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

/** How many LibreOffice conversions may run at once suite-wide. */
export const LO_SLOT_COUNT = 2;
/** Poll interval while waiting for a free slot (with jitter added). */
const POLL_MS = 200;
/** Holder heartbeat: refresh lock heartbeat this often while working.
 * Test-only override (e.g. 100) for fencing stress; default unchanged. */
const HEARTBEAT_MS =
  Number(process.env.ASTRA_LO_SLOT_HEARTBEAT_MS ?? 30_000) || 30_000;
/**
 * A lock with no heartbeat for this long is treated as orphaned and
 * reclaimed. Must stay above the longest legitimate hold: one
 * convertDocxToPdf call can legally spend ~2 render runs x (3 attempts x
 * 30s spawn budget + backoff) ~= 200s, plus queue-side overhead.
 */
const STALE_MS = 10 * 60 * 1000;

/** Canonical per-acquisition ownership record. `token` is the fencing epoch. */
interface HolderRecord {
  pid: number;
  token: string;
  since: number;
  heartbeat: number;
}

function baseDir(): string {
  const dir = path.join(os.tmpdir(), 'astra-lo-slot');
  try {
    fs.mkdirSync(dir, { recursive: true });
  } catch {
    /* another worker created it first */
  }
  return dir;
}

function slotDir(i: number): string {
  return path.join(baseDir(), `slot-${i}.lock`);
}

/** Canonical holder record path. Single well-known name (not per-PID) so a
 *  stale holder's file can be fenced by token instead of guessed by name. */
function holderPath(i: number): string {
  return path.join(slotDir(i), 'holder.json');
}

let tmpSeq = 0;

/**
 * D1 fix: atomic holder write. Readers either see the full old bytes or the
 * full new bytes; rename(2) over an existing name is atomic on POSIX, and
 * the tmp file lives in the same dir (same filesystem) so no copy fallback.
 */
function atomicWriteFileSync(target: string, content: string): void {
  const tmp = `${target}.tmp.${process.pid}.${tmpSeq++}`;
  try {
    fs.writeFileSync(tmp, content);
    fs.renameSync(tmp, target);
  } catch (err) {
    try {
      fs.rmSync(tmp, { force: true });
    } catch {
      /* best-effort tmp cleanup */
    }
    throw err;
  }
}

function isHolderRecord(v: unknown): v is HolderRecord {
  if (typeof v !== 'object' || v === null) return false;
  const r = v as Record<string, unknown>;
  return (
    typeof r['pid'] === 'number' &&
    Number.isInteger(r['pid']) &&
    (r['pid'] as number) > 0 &&
    typeof r['token'] === 'string' &&
    (r['token'] as string).length > 0 &&
    typeof r['heartbeat'] === 'number' &&
    Number.isFinite(r['heartbeat'])
  );
}

interface HolderSnapshot {
  dirExists: boolean;
  /** Exact canonical bytes; null when holder.json is absent/unreadable. */
  rawCanonical: string | null;
  record: HolderRecord | null;
  canonicalMalformed: boolean;
  /** Canonical PID (if valid) plus legacy holder-<pid>.json PIDs. */
  pids: number[];
  legacyFiles: number;
  dirMtimeMs: number | null;
}

/**
 * Read the current ownership state of slot i. Never throws. Legacy
 * holder-<pid>.json files (pre-token format) contribute PIDs only; the
 * canonical holder.json is the single source of truth for fencing.
 */
function readHolderSnapshot(i: number): HolderSnapshot {
  const dir = slotDir(i);
  const snap: HolderSnapshot = {
    dirExists: true,
    rawCanonical: null,
    record: null,
    canonicalMalformed: false,
    pids: [],
    legacyFiles: 0,
    dirMtimeMs: null,
  };
  try {
    snap.dirMtimeMs = fs.statSync(dir).mtimeMs;
  } catch {
    snap.dirExists = false;
    return snap;
  }
  try {
    snap.rawCanonical = fs.readFileSync(holderPath(i), 'utf8');
  } catch {
    snap.rawCanonical = null;
  }
  if (snap.rawCanonical !== null) {
    try {
      const parsed: unknown = JSON.parse(snap.rawCanonical);
      if (isHolderRecord(parsed)) {
        snap.record = parsed;
        snap.pids.push(parsed.pid);
      } else {
        snap.canonicalMalformed = true;
      }
    } catch {
      snap.canonicalMalformed = true;
    }
  }
  let names: string[];
  try {
    names = fs.readdirSync(dir);
  } catch {
    return snap;
  }
  for (const n of names) {
    if (!n.startsWith('holder-') || !n.endsWith('.json')) continue;
    snap.legacyFiles++;
    try {
      const raw = fs.readFileSync(path.join(dir, n), 'utf8');
      const pid = (JSON.parse(raw) as { pid?: unknown })?.pid;
      if (typeof pid === 'number' && Number.isInteger(pid) && pid > 0) {
        snap.pids.push(pid);
      }
    } catch {
      /* torn legacy write: contributes no PID, never triggers fast path */
    }
  }
  return snap;
}

/**
 * Liveness check for a recorded holder PID. ESRCH means dead; EPERM means
 * alive-but-unowned. Any other error is treated conservatively as alive so
 * an ambiguous holder never triggers a reclaim on its own (the stale-timer
 * backstop still applies).
 */
function isPidAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (err: unknown) {
    if ((err as { code?: string })?.code === 'ESRCH') return false;
    return true;
  }
}

function newToken(): string {
  try {
    return crypto.randomUUID();
  } catch {
    return `${process.pid}-${Date.now()}-${Math.floor(Math.random() * 2 ** 31)}`;
  }
}

function freshRecord(): HolderRecord {
  const now = Date.now();
  return { pid: process.pid, token: newToken(), since: now, heartbeat: now };
}

/** Refresh the lock dir mtime (legacy/fallback freshness signal).
 * Returns true on success, false when the dir is gone/unreachable so the
 * caller (heartbeat) reports loss instead of false-held. */
function touchDir(i: number): boolean {
  try {
    const now = new Date();
    fs.utimesSync(slotDir(i), now, now);
    return true;
  } catch {
    return false; /* lock released or reclaimed under us: ownership unconfirmable */
  }
}

/**
 * Remove the lock dir and atomically establish fresh ownership under a new
 * fencing token. Returns the new token, or null on losing the reclaim race.
 * The post-write verify narrows the stale-reclaim race: if a racing
 * reclaimer already wiped our fresh write, back off instead of double-holding.
 */
function reclaimAndTake(i: number): string | null {
  try {
    fs.rmSync(slotDir(i), { recursive: true, force: true });
  } catch {
    return null;
  }
  try {
    fs.mkdirSync(slotDir(i));
  } catch {
    return null; // lost the reclaim race; keep polling
  }
  const rec = freshRecord();
  try {
    atomicWriteFileSync(holderPath(i), JSON.stringify(rec));
    touchDir(i);
  } catch {
    return null;
  }
  try {
    const cur = JSON.parse(fs.readFileSync(holderPath(i), 'utf8') as string) as Partial<HolderRecord>;
    if (cur?.token !== rec.token) return null; // wiped under us; back off
  } catch {
    return null;
  }
  return rec.token;
}

/**
 * EEXIST slow path: decide whether the existing lock can be reclaimed now.
 * Returns the fresh fencing token when this process now holds a fresh lock
 * dir, else null. See the header comment for the full reclaim semantics.
 */
function tryReclaim(i: number): string | null {
  const snap = readHolderSnapshot(i);
  if (!snap.dirExists) return null;
  // Fast path: every recorded holder PID is dead -> reclaim immediately,
  // whatever the age. Dead means dead; no waiting out STALE_MS.
  if (snap.pids.length > 0 && !snap.pids.some((pid) => isPidAlive(pid))) {
    return reclaimAndTake(i);
  }
  // NOTE: a malformed canonical record deliberately does NOT reclaim here.
  // Without a valid PID no liveness check is possible, and a torn-looking
  // file next to a fresh lock may still have a live owner.
  // Backstop: heartbeat staleness + provable lack of holder progress.
  const hb = snap.record?.heartbeat ?? snap.dirMtimeMs;
  if (hb === null || hb === undefined || !(Date.now() - hb > STALE_MS)) {
    return null;
  }
  // D2a: re-check CURRENT content at stale-fire time. Decide on what is in
  // the lock dir NOW, not on the earlier snapshot: a beat or an ownership
  // change that landed between the two reads aborts the reclaim.
  const cur = readHolderSnapshot(i);
  if (cur.rawCanonical !== snap.rawCanonical) return null;
  const hb2 = cur.record?.heartbeat ?? cur.dirMtimeMs;
  if (hb2 === null || hb2 === undefined || !(Date.now() - hb2 > STALE_MS)) {
    return null;
  }
  // Stale + byte-unchanged across the decision window: the holder made no
  // progress for >STALE_MS. Reclaim even if a recorded PID looks alive --
  // PIDs recycle (D3), and a genuinely live holder beats every
  // HEARTBEAT_MS << STALE_MS, so it can never trip this branch.
  return reclaimAndTake(i);
}

/**
 * Try to take slot i. Returns the fencing token when this process now holds
 * it, else null.
 */
function tryTake(i: number): string | null {
  try {
    fs.mkdirSync(slotDir(i));
  } catch (err: unknown) {
    if ((err as { code?: string })?.code !== 'EEXIST') throw err;
    return tryReclaim(i);
  }
  const rec = freshRecord();
  try {
    atomicWriteFileSync(holderPath(i), JSON.stringify(rec));
  } catch {
    // No fencing record means no provable ownership: never operate
    // unlocked. Retry via the poll loop (the test timeout stays the single
    // budget) instead of holding a record-less lock.
    return null;
  }
  touchDir(i);
  return rec.token;
}

/**
 * Confirm the lock is still ours: the canonical holder record must parse
 * and carry our fencing token. Anything else (absent, torn, foreign)
 * means we cannot prove ownership -- the caller aborts the hold loudly.
 * Confirm free; abort on collision.
 */
function isStillHolder(i: number, token: string): boolean {
  try {
    const cur = JSON.parse(
      fs.readFileSync(holderPath(i), 'utf8'),
    ) as Partial<HolderRecord>;
    return cur?.token === token;
  } catch {
    return false;
  }
}

function fencingError(i: number, phase: string): Error {
  return new Error(
    `LO slot fencing violation: lost slot ${i} (pid ${process.pid}) ${phase}; ` +
      `refusing to run unlocked`,
  );
}

/**
 * Fencing-aware best-effort compare-and-swap heartbeat (NOT atomic): read
 * the canonical bytes, compute the refreshed record, re-read and rewrite
 * (fresh `heartbeat` timestamp) ONLY while the bytes are unchanged since
 * the first read and still carry our token, then touch the dir mtime.
 * Returns false when ownership loss is observed at a read -- the caller
 * then stops beating and aborts loudly. It does NOT guarantee a reclaimed
 * holder always stops: the verify-read/write window (header residual 5)
 * can still revive a stale token over a successor.
 */
function heartbeat(i: number, token: string): boolean {
  let firstRaw: string;
  try {
    firstRaw = fs.readFileSync(holderPath(i), 'utf8');
  } catch {
    return false; // gone/unreadable: ownership unconfirmable -- stop
  }
  let first: Partial<HolderRecord>;
  try {
    first = JSON.parse(firstRaw) as Partial<HolderRecord>;
  } catch {
    return false; // torn where our record should be: unconfirmable -- stop
  }
  if (first?.token !== token) return false; // reclaimed under us -- stop
  if (typeof first?.heartbeat !== 'number') {
    // Our token but no heartbeat field to refresh (never happens for
    // records we mint; kept for shape-tolerance): re-verify before
    // touching so a successor flip aborts instead of extending them.
    try {
      if (fs.readFileSync(holderPath(i), 'utf8') !== firstRaw) return false;
    } catch {
      return false;
    }
    return touchDir(i);
  }
  const next = JSON.stringify({
    ...first,
    pid: process.pid,
    heartbeat: Date.now(),
  });
  // CAS verify: a reclaim (or any write) that landed after our first read
  // changed the bytes -- abort the beat instead of clobbering it.
  try {
    if (fs.readFileSync(holderPath(i), 'utf8') !== firstRaw) return false;
  } catch {
    return false;
  }
  try {
    atomicWriteFileSync(holderPath(i), next);
  } catch {
    return false; // transient write failure: ownership unconfirmable, never report held
  }
  // Post-write verify: a successor write landing after ours owns the lock
  // now -- report the loss promptly instead of beating on next cycle.
  try {
    if (fs.readFileSync(holderPath(i), 'utf8') !== next) return false;
  } catch {
    return false;
  }
  return touchDir(i);
}

/**
 * D2b fix: best-effort compare-and-delete release (NOT atomic). Reads the
 * canonical holder record and removes the lock dir ONLY if that read still
 * carries our fencing token. Outside the read-then-rm window (header
 * residual 6) a reclaimed holder leaves its successor's fresh lock
 * untouched instead of rm -rf'ing it (which used to enable two holders on
 * one slot).
 */
function release(i: number, token: string): void {
  try {
    const raw = fs.readFileSync(holderPath(i), 'utf8');
    const cur = JSON.parse(raw) as Partial<HolderRecord>;
    if (cur?.token !== token) return;
  } catch {
    return; // holder gone or foreign: nothing provably ours -- hands off
  }
  try {
    fs.rmSync(slotDir(i), { recursive: true, force: true });
  } catch {
    /* best-effort; a stale lock is reclaimed by the next waiter */
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

/**
 * Run fn while holding one suite-wide LibreOffice slot.
 * Calls in the same process are already sequential (one `it` at a time),
 * so no in-process bookkeeping is needed — the lock is purely cross-worker.
 */
export async function withLoSlot<T>(fn: () => Promise<T>): Promise<T> {
  for (;;) {
    for (let i = 0; i < LO_SLOT_COUNT; i++) {
      const token = tryTake(i);
      if (token !== null) {
        // Holder self-check BEFORE the guarded operation: never operate
        // unlocked. A fencing loss here fails loudly instead of running
        // LibreOffice without a slot.
        if (!isStillHolder(i, token)) {
          release(i, token);
          throw fencingError(i, 'before guarded operation');
        }
        let onFenced: ((err: Error) => void) | undefined;
        const fenced = new Promise<never>((_, reject) => {
          onFenced = reject;
        });
        const beat = setInterval(() => {
          // Self-check before each beat lives inside heartbeat (CAS);
          // a lost beat stops the interval and rejects loudly via the race.
          if (!heartbeat(i, token)) {
            clearInterval(beat);
            onFenced?.(fencingError(i, 'during hold'));
          }
        }, HEARTBEAT_MS);
        try {
          const result = await Promise.race([fn(), fenced]);
          // G1: post-fn ownership re-check. fn() may have resolved while a
          // reclaim landed underneath it (check-then-spawn TOCTOU); a fenced
          // holder must throw instead of reporting success-while-fenced.
          if (!isStillHolder(i, token)) {
            throw fencingError(i, 'after guarded operation');
          }
          return result;
        } finally {
          clearInterval(beat);
          release(i, token);
        }
      }
    }
    await sleep(POLL_MS + Math.floor(Math.random() * POLL_MS));
  }
}
