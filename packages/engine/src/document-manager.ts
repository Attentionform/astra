import fs from 'fs';
import fsp from 'fs/promises';
import path from 'path';
import os from 'os';
import { createHash } from 'crypto';
import { DocxParser, DocumentSnapshot, StructuralSummary } from './parser';
import { initAstraRepo, commitChange, undo, getHistory, getRepoDir } from './git';
import * as gitMod from './git';
import { validateDocxBuffer, OoxmlValidationError, ValidationResult } from './ooxml/validator-gate';
import AdmZip from './adm-zip-safe';
import { lintAndFixTypography, TypographicFix, TypographyOptions } from './ooxml/typography';
import type { StyleGuideConfig } from './style-guide/types';
import { lintDocumentStyle, chartXmlGetterForAst, type StyleWarning } from './style-guide/lint';
import {
  canonicalizeForAsset as canonicalizeForSandbox,
  isPathInsideDir,
} from './sandbox/paths';

// ── UX-B autosave + atomic-write (UX-01 P0-1; OX-08 torn-zip) ─────────────────
// Autosave data path (no rerouting anywhere): renderer idle timer (UX-B,
// renderer/utils/autosave.ts) → UX-A flush → executeOperations IPC →
// commitTransaction below, which writes activeDocPath. For untitled docs the
// active path IS the silent Drafts fallback, so Drafts autosave is automatic;
// titled docs write in place. Autosave MUST NEVER clear flags (it never calls
// markClean / never flips isUntitled) — explicit Save/Save-As remain the only
// flag-clearing paths.
//
// Git-coalescing DOCUMENTED POLICY (no history flood + no snapshot eviction):
// renderer debounce (idle 2–3s, reset on input) + flush no-op-when-clean +
// huge-doc backoff bound the commit rate to ≤1 checkpoint batch per idle
// window with real edits — never per keystroke. The time-window coalescer
// below (shouldCoalesceAutosaveCommit + commitTransaction {autosave:true}) is
// the backend half for direct-manager/headless autosave callers: file is
// ALWAYS written + isDirty set; the git commit AND the snapshot push are BOTH
// skipped inside the window (same predicate) so snapshotStack stays aligned
// with the git log and the 50-cap undo ring is not evicted (~2min of idle
// autosaves would otherwise shift it out). Coalesced autosaves are durability-
// only — they are not separate undo steps and must not clear the redoStack.
// The window tracks AUTOSAVE commits only — manual commits (updateNode,
// explicit commitTransaction, Save, AI-stream) ALWAYS commit + push snapshot
// and never touch the autosave clock (integration suites pin exact history
// counts — see tests/integration/document-manipulation.test.ts).
//
// UX-C item 1 NO-OP DOCUMENTED POLICY (empty-autosave false banner):
// an AUTOSAVE whose document.xml hashes identically to the last committed
// hash is a FULL no-op — no snapshot, no file write, no git commit, NO dirty
// flip, no coalesce-clock update. Pristine drafts therefore stay at
// commitCount==1 (no noise banner) and stay clean (close gate stays silent).
// Manual commits bypass the gate by design (explicit intent always commits).
// Coalescing applies only to content-changed autosaves, so its semantics are
// untouched; coalesced durability writes deliberately do NOT advance the
// committed hash.
export const AUTOSAVE_GIT_COALESCE_MS = 30_000;

/** Pure predicate: skip the git commit when an autosave lands inside the window. */
export function shouldCoalesceAutosaveCommit(
  lastAutosaveCommitAt: number,
  now: number,
  windowMs: number = AUTOSAVE_GIT_COALESCE_MS
): boolean {
  if (!Number.isFinite(lastAutosaveCommitAt) || !Number.isFinite(now) || !Number.isFinite(windowMs)) return false;
  return now - lastAutosaveCommitAt < windowMs;
}

/**
 * UX-C empty-autosave no-op gate (UX-01 reviewer follow-up item 1).
 * Cheap content identity for document.xml (sha256). Autosave with zero edits
 * must not mint a git commit (commitCount>1 noise banner), push a snapshot,
 * rewrite the file, flip isDirty, or touch the coalesce clock.
 * Manual commitTransaction calls intentionally BYPASS this gate (explicit user
 * intent always commits — history-count suites pin that).
 */
export function hashDocumentXml(xml: string): string {
  return createHash('sha256').update(xml, 'utf8').digest('hex');
}

let atomicWriteCounter = 0;

/**
 * UX-B single-owner tmp+rename atomic docx write (OX-08 torn-zip fix).
 * Writes to a uniquely-named tmp file in the SAME directory (same filesystem
 * so rename is atomic) then renames over the target. Windows rename-onto-
 * existing / locked-target (EPERM/EEXIST/ENOTEMPTY plus EBUSY/EACCES from
 * AV scanners / Word locks) falls back to unlink+rename with bounded retries,
 * then fails closed (throws). Tmp litter is removed best-effort on failure;
 * crash-mid-write orphans (`*.tmp-<pid>-*`) are owned by UX-C orphan GC.
 * D2/export paths MUST reuse this helper — no ad-hoc docx writers.
 */
export const ATOMIC_WRITE_RETRY_CODES = new Set(['EPERM', 'EEXIST', 'ENOTEMPTY', 'EBUSY', 'EACCES']);
/** Bounded unlink+rename attempts after the fast-path rename fails. */
export const ATOMIC_WRITE_MAX_RETRIES = 5;
/** Base backoff between unlink-retry attempts (linear: delay * (attempt+1)). */
export const ATOMIC_WRITE_RETRY_DELAY_MS = 25;

function sleepMs(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Guard narrowing for unknown catch values (replaces `catch (...: any)`). */
function asRecord(value: unknown): Record<string, unknown> | null {
  return typeof value === 'object' && value !== null ? (value as Record<string, unknown>) : null;
}

function errorCodeOf(err: unknown): string | undefined {
  const rec = asRecord(err);
  const code = rec?.['code'];
  return typeof code === 'string' ? code : undefined;
}

function errorMessageOf(err: unknown): string {
  if (err instanceof Error) return err.message !== '' ? err.message : String(err);
  const rec = asRecord(err);
  const msg = rec?.['message'];
  return typeof msg === 'string' && msg !== '' ? msg : String(err);
}

function hasValidationCode(err: unknown): boolean {
  return asRecord(err)?.['code'] === 'E_VALIDATION';
}

/** Narrow optional git-module fns without `as any` (tolerates partial vitest mocks). */
interface GitUndoPackage {
  undoPackage?: (p: string) => Promise<{ xml: string; extra: Map<string, Buffer> }>;
  isTrackedPackageFile?: (name: string) => boolean;
}

function asGitUndoPackage(mod: unknown): GitUndoPackage {
  const rec = asRecord(mod);
  if (!rec) return {};
  const out: GitUndoPackage = {};
  if (typeof rec['undoPackage'] === 'function') {
    out.undoPackage = rec['undoPackage'] as GitUndoPackage['undoPackage'];
  }
  if (typeof rec['isTrackedPackageFile'] === 'function') {
    out.isTrackedPackageFile = rec['isTrackedPackageFile'] as GitUndoPackage['isTrackedPackageFile'];
  }
  return out;
}

/** Narrow parser snapshot capability without `as any` duck-typing. */
interface PackageSnapshotCapable {
  capturePackageSnapshot?: () => Map<string, Buffer>;
  restorePackageSnapshot?: (snap: Map<string, Buffer>) => void;
}

function asPackageSnapshotCapable(parser: unknown): PackageSnapshotCapable {
  const rec = asRecord(parser);
  if (!rec) return {};
  const out: PackageSnapshotCapable = {};
  if (typeof rec['capturePackageSnapshot'] === 'function') {
    out.capturePackageSnapshot = (rec['capturePackageSnapshot'] as () => Map<string, Buffer>).bind(parser);
  }
  if (typeof rec['restorePackageSnapshot'] === 'function') {
    out.restorePackageSnapshot = (rec['restorePackageSnapshot'] as (snap: Map<string, Buffer>) => void).bind(parser);
  }
  return out;
}

export async function atomicWriteFile(
  targetPath: string,
  data: Buffer | Uint8Array | string
): Promise<void> {
  const dir = path.dirname(targetPath);
  await fsp.mkdir(dir, { recursive: true });
  atomicWriteCounter = (atomicWriteCounter + 1) % 1_000_000;
  const tmpPath = `${targetPath}.tmp-${process.pid}-${Date.now()}-${atomicWriteCounter}`;
  try {
    await fsp.writeFile(tmpPath, data);
    try {
      await fsp.rename(tmpPath, targetPath);
      return;
    } catch (renameErr: unknown) {
      if (!ATOMIC_WRITE_RETRY_CODES.has(errorCodeOf(renameErr) ?? '')) throw renameErr;
      let lastErr: unknown = renameErr;
      for (let attempt = 0; attempt < ATOMIC_WRITE_MAX_RETRIES; attempt++) {
        try {
          try {
            await fsp.unlink(targetPath);
          } catch (unlinkErr: unknown) {
            if (errorCodeOf(unlinkErr) === 'ENOENT') {
              // Target already gone — proceed to the rename below.
            } else if (ATOMIC_WRITE_RETRY_CODES.has(errorCodeOf(unlinkErr) ?? '')) {
              lastErr = unlinkErr;
              await sleepMs(ATOMIC_WRITE_RETRY_DELAY_MS * (attempt + 1));
              continue;
            } else {
              throw unlinkErr;
            }
          }
          await fsp.rename(tmpPath, targetPath);
          return;
        } catch (retryErr: unknown) {
          if (ATOMIC_WRITE_RETRY_CODES.has(errorCodeOf(retryErr) ?? '')) {
            lastErr = retryErr;
            await sleepMs(ATOMIC_WRITE_RETRY_DELAY_MS * (attempt + 1));
            continue;
          }
          throw retryErr;
        }
      }
      // Bounded retries exhausted — fail closed (throw). Tmp survivor (if any)
      // is owned by UX-C orphan GC; the outer catch makes a best-effort removal.
      throw lastErr;
    }
  } catch (err) {
    try {
      await fsp.unlink(tmpPath);
    } catch { /* best-effort tmp cleanup; survivors owned by UX-C orphan GC */ }
    throw err;
  }
}

export class DocumentManager {
  private activeDocPath: string | null = null;
  private parser: DocxParser | null = null;
  private snapshotStack: DocumentSnapshot[] = [];
  private redoStack: DocumentSnapshot[] = [];
  private transactionSnapshot: DocumentSnapshot | null = null;
  /** Package bytes as last opened or persisted: what memory returns to when a write is refused. */
  private lastPersistedBuf: Buffer | null = null;
  private isDirty: boolean = false;
  private isUntitled: boolean = false;
  // UX-B: wall-clock of the last AUTOSAVE git commit from THIS manager.
  // Manual commits never read or write it (they always commit git).
  private lastAutosaveCommitAt: number = 0;
  // UX-C item 1: hash of the last COMMITTED document.xml (init/open/save/
  // committed transaction/undo/redo). Autosave compares the live XML against
  // this; identical content is a no-op (no commit, no snapshot, no file write,
  // no dirty flip, no clock update). Null until the first commit is recorded.
  private lastCommittedXmlHash: string | null = null;
  // Active style guide for advisory linting on save/validate. Style is
  // advisory only: findings ride along as styleWarnings and never block a
  // save (E_VALIDATION stays reserved for schema/structural failures).
  private activeGuide?: StyleGuideConfig;

  /**
   * Attaches (or clears, when omitted) the style guide used for advisory
   * linting on save/validate paths.
   */
  public setStyleGuide(guide?: StyleGuideConfig): void {
    this.activeGuide = guide;
  }

  /** Active style guide (undefined = no advisory linting). */
  public getStyleGuide(): StyleGuideConfig | undefined {
    return this.activeGuide;
  }

  /**
   * Advisory style pass over the live AST. Best-effort and read-only:
   * any failure yields no warnings, never a failed save. Empty when no
   * guide is active.
   */
  private async collectStyleWarnings(): Promise<StyleWarning[]> {
    try {
      if (!this.activeGuide || !this.parser) return [];
      const ast = await this.parser.getAST();
      const zip = this.parser.getZipAdapter();
      const getChartXml = chartXmlGetterForAst(ast, (part) => zip.readTextOrNull(part));
      return lintDocumentStyle(ast, this.activeGuide, getChartXml);
    } catch {
      return [];
    }
  }

  /**
   * Opens a DOCX file. Automatically creates a shadow copy (.astra.docx) 
   * if not already opening one, initializes the Git undo stack, and loads the DOM.
   */
  public async openDocument(sourcePath: string): Promise<{ astraDocPath: string; summary: StructuralSummary }> {
    // S1b follow-up item 1: re-validate at USE time (TOCTOU narrowing) — the
    // on-disk target is re-canonicalized now, catching symlink swaps that
    // happened after IPC-entry validation. Dialog-trusted paths pass via the
    // user-intent registry.
    const trustedSource = assertSandboxPathAtUse(sourcePath);
    this.activeDocPath = trustedSource;
    this.isUntitled = false;
    this.isDirty = false;

    // Load the document into memory parser (use the validated path, not the
    // raw caller string, so normalization drift can't bypass the gate above).
    const buf = await fsp.readFile(trustedSource);
    this.parser = new DocxParser(buf);
    this.lastPersistedBuf = buf;

    // Initialize Git repo if missing
    const repoDir = getRepoDir(trustedSource);
    if (!fs.existsSync(repoDir)) {
      const initialXml = this.parser.readDocumentXml();
      await initAstraRepo(trustedSource, initialXml, this.collectPackageExtras());
    }

    // Push initial snapshot into atomic history stack
    this.snapshotStack = [this.parser.createAtomicSnapshot('Initial state')];
    this.redoStack = [];
    // UX-C item 1: baseline for the empty-autosave no-op gate.
    try {
      this.lastCommittedXmlHash = hashDocumentXml(this.parser.readDocumentXml());
    } catch { this.lastCommittedXmlHash = null; }

    // Return structural summary
    return {
      astraDocPath: trustedSource,
      summary: this.parser.getStructuralSummary()
    };
  }

  public async createNewDocument(templateType: string = 'blank', targetPath?: string, opts?: { untitled?: boolean }): Promise<{ astraDocPath: string; summary: StructuralSummary; warnings: string[] }> {
    const docDir = path.join(process.cwd(), 'documents');
    if (!fs.existsSync(docDir)) {
      fs.mkdirSync(docDir, { recursive: true });
    }

    const filename = `ASTra_New_Document_${Date.now()}.docx`;
    let docPath = targetPath || path.join(docDir, filename);
    // S1b follow-up item 1: at-use re-validation for the create target.
    // App-generated fallbacks (silent Drafts) are registered as trusted by
    // electron/main.ts; direct API targets must pass the sandbox here too.
    try {
      docPath = assertSandboxPathAtUse(docPath);
    } catch (err) {
      // Fall back to the workspace documents tree (always inside the sandbox)
      // rather than failing the create outright on a hostile customPath that
      // slipped past an older caller.
      if (targetPath) throw err;
      docPath = path.join(docDir, filename);
    }

    // Ensure target dir exists (Drafts fallback from electron/main.ts may not exist yet)
    let targetDir = path.dirname(docPath);
    try {
      if (!fs.existsSync(targetDir)) {
        fs.mkdirSync(targetDir, { recursive: true });
      }
    } catch {
      docPath = path.join(docDir, filename);
      targetDir = docDir;
      if (!fs.existsSync(targetDir)) {
        fs.mkdirSync(targetDir, { recursive: true });
      }
    }

    this.parser = DocxParser.createNewDocument(templateType);
    const newDocxBuf = this.parser.rezip();
    // Validate with OOXML validator before committing to disk
    try {
      const validation = await validateDocxBuffer(newDocxBuf);
      if (!validation.ok) {
        const errorDescs = validation.errors.map(e => e.description || e.message || String(e));
        throw Object.assign(new Error(errorDescs.join('; ')), { code: 'E_VALIDATION', retryable: false });
      }
    } catch (valErr: unknown) {
      if (hasValidationCode(valErr)) throw valErr;
      throw Object.assign(new Error(`OOXML validation failure: ${errorMessageOf(valErr)}`), {
        code: 'E_VALIDATION',
        retryable: false
      });
    }

    this.lastPersistedBuf = newDocxBuf;
    let finalPath = docPath;
    const createWarnings: string[] = [];
    try {
      await atomicWriteFile(finalPath, newDocxBuf);
    } catch (writeErr: unknown) {
      // Typed carrier for the dropped primary-write diagnostic: the returned
      // astraDocPath names the location actually used; warnings names WHY the
      // fallback happened (same text as the removed console.warn).
      createWarnings.push(`[DocumentManager] Primary write failed for ${finalPath}: ${errorMessageOf(writeErr)}. Trying workspace documents fallback...`);
      const fallbackDir = path.join(process.cwd(), 'documents');
      if (!fs.existsSync(fallbackDir)) {
        fs.mkdirSync(fallbackDir, { recursive: true });
      }
      const workspaceFallback = path.join(fallbackDir, path.basename(docPath));
      try {
        await atomicWriteFile(workspaceFallback, newDocxBuf);
        finalPath = workspaceFallback;
      } catch (wsErr: unknown) {
        createWarnings.push(`[DocumentManager] Workspace fallback failed: ${errorMessageOf(wsErr)}. Trying temp fallback...`);
        const tempDir = path.join(os.tmpdir(), 'ASTra');
        if (!fs.existsSync(tempDir)) {
          fs.mkdirSync(tempDir, { recursive: true });
        }
        const tempFallback = path.join(tempDir, path.basename(docPath));
        await atomicWriteFile(tempFallback, newDocxBuf);
        finalPath = tempFallback;
      }
    }

    this.activeDocPath = finalPath;
    // UX-C-flags (OX-09): explicit opts.untitled wins over the heuristic both
    // ways. The Drafts-substring + finalPath checks below are kept as a
    // fallback safety net for direct-manager callers/tests that pass no opts.
    if (typeof opts?.untitled === 'boolean') {
      this.isUntitled = opts.untitled;
    } else {
      // ELEC-07: createNewDocument is always called with an explicit targetPath from
      // electron/main.ts — including the silent Drafts fallback. Detect that fallback
      // so the doc stays untitled and document:save still routes to the Save-As dialog.
      this.isUntitled = !targetPath ||
        finalPath !== targetPath ||
        (targetPath.includes('Drafts') && path.basename(targetPath).startsWith('ASTra_Document_'));
    }
    this.isDirty = false;
    const repoDir = getRepoDir(finalPath);
    if (!fs.existsSync(repoDir)) {
      const initialXml = this.parser.readDocumentXml();
      await initAstraRepo(finalPath, initialXml, this.collectPackageExtras());
    } else {
      const initialXml = this.parser.readDocumentXml();
      await commitChange(finalPath, initialXml, `Created document from ${templateType} template`, this.collectPackageExtras());
    }
    this.snapshotStack = [this.parser.createAtomicSnapshot('Initial state')];
    this.redoStack = [];
    // UX-C item 1: baseline for the empty-autosave no-op gate.
    try {
      this.lastCommittedXmlHash = hashDocumentXml(this.parser.readDocumentXml());
    } catch { this.lastCommittedXmlHash = null; }

    return {
      astraDocPath: finalPath,
      summary: this.parser.getStructuralSummary(),
      warnings: createWarnings
    };
  }

  public async saveAsDocument(
    targetPath: string,
    options?: { applyTypography?: boolean; typographyOptions?: TypographyOptions }
  ): Promise<{ astraDocPath: string; summary: StructuralSummary; typographyFixes: TypographicFix[]; styleWarnings: StyleWarning[] }> {
    if (!this.parser) {
      throw new Error("No active document loaded.");
    }
    // S1b follow-up item 1: at-use re-validation for the write target.
    const trustedTarget = assertSandboxPathAtUse(targetPath);
    const targetDir = path.dirname(trustedTarget);
    if (!fs.existsSync(targetDir)) {
      fs.mkdirSync(targetDir, { recursive: true });
    }

    const preSaveSnapshot = this.parser.createAtomicSnapshot('PreSaveTypographyBackup');
    let typographyFixes: TypographicFix[] = [];
    if (options?.applyTypography !== false) {
      const lintRes = lintAndFixTypography(this.parser.getDom(), options?.typographyOptions);
      typographyFixes = lintRes.fixes;
    }

    const docxBuf = this.parser.rezip();
    // Validate with OOXML validator before committing to disk
    try {
      const validation = await validateDocxBuffer(docxBuf);
      if (!validation.ok) {
        this.parser.restoreAtomicSnapshot(preSaveSnapshot);
        const errorDescs = validation.errors.map(e => e.description || e.message || String(e));
        throw Object.assign(new Error(errorDescs.join('; ')), { code: 'E_VALIDATION', retryable: false });
      }
    } catch (valErr: unknown) {
      this.parser.restoreAtomicSnapshot(preSaveSnapshot);
      if (hasValidationCode(valErr)) throw valErr;
      throw Object.assign(new Error(`OOXML validation failure: ${errorMessageOf(valErr)}`), {
        code: 'E_VALIDATION',
        retryable: false
      });
    }

    await atomicWriteFile(trustedTarget, docxBuf);
    this.lastPersistedBuf = docxBuf;

    this.activeDocPath = trustedTarget;
    this.isUntitled = false;
    this.isDirty = false;
    const repoDir = getRepoDir(trustedTarget);
    if (!fs.existsSync(repoDir)) {
      const initialXml = this.parser.readDocumentXml();
      await initAstraRepo(trustedTarget, initialXml, this.collectPackageExtras());
    } else {
      const currentXml = this.parser.readDocumentXml();
      await commitChange(trustedTarget, currentXml, `Saved document to ${path.basename(trustedTarget)}`, this.collectPackageExtras());
    }
    // UX-C item 1: explicit save re-baselines the no-op gate.
    try {
      this.lastCommittedXmlHash = hashDocumentXml(this.parser.readDocumentXml());
    } catch { this.lastCommittedXmlHash = null; }

    return {
      astraDocPath: trustedTarget,
      summary: this.parser.getStructuralSummary(),
      typographyFixes,
      styleWarnings: await this.collectStyleWarnings(),
    };
  }

  /**
   * Validates the active document against official ECMA-376 XSD schemas.
   * The advisory style pass rides along as styleWarnings (never a failure).
   */
  public async validateCurrentDocument(): Promise<ValidationResult & { styleWarnings: StyleWarning[] }> {
    if (!this.parser) throw new Error("No active document loaded.");
    const result = await this.parser.validateOOXML();
    return { ...result, styleWarnings: await this.collectStyleWarnings() };
  }

  /**
   * Creates an atomic snapshot of the document's current state.
   */
  public createSnapshot(label: string = 'User Checkpoint'): DocumentSnapshot {
    if (!this.parser) throw new Error("No active document.");
    const snapshot = this.parser.createAtomicSnapshot(label);
    this.pushSnapshot(snapshot);
    return snapshot;
  }

  private pushSnapshot(snapshot: DocumentSnapshot): void {
    this.snapshotStack.push(snapshot);
    if (this.snapshotStack.length > 50) {
      this.snapshotStack.shift();
    }
    this.redoStack = [];
  }

  /**
   * The only write path for edits made while a document is open: the exact
   * bytes about to be written are validated first. A refused write leaves the
   * file untouched and returns memory to the last persisted package.
   */
  private async persistValidated(target: string, buf: Buffer): Promise<void> {
    const validation = await validateDocxBuffer(buf);
    if (!validation.ok) {
      this.restoreLastPersisted();
      throw new OoxmlValidationError(validation.errors, 'Document save');
    }
    await atomicWriteFile(target, buf);
    this.lastPersistedBuf = buf;
  }

  private restoreLastPersisted(): void {
    if (!this.parser || !this.lastPersistedBuf) return;
    const entries = new Map<string, Buffer>();
    for (const entry of new AdmZip(this.lastPersistedBuf).getEntries()) {
      if (!entry.isDirectory) entries.set(entry.entryName, entry.getData());
    }
    this.parser.restorePackageSnapshot(entries);
  }

  /**
   * Begins an atomic transaction, capturing a pre-transaction snapshot.
   */
  public beginTransaction(label: string = 'Batch Transaction'): void {
    if (!this.parser) throw new Error("No active document.");
    this.transactionSnapshot = this.parser.createAtomicSnapshot(label);
  }

  /**
   * Rolls back the current atomic transaction to pre-transaction state.
   */
  public rollbackTransaction(): void {
    if (!this.parser || !this.transactionSnapshot) return;
    this.parser.restoreAtomicSnapshot(this.transactionSnapshot);
    this.transactionSnapshot = null;
  }

  /**
   * Commits the current atomic transaction, flushing file changes to disk and Git.
   *
   * UX-B autosave (`opts.autosave`): the file is ALWAYS atomically written and
   * flags ALWAYS preserved dirty/untitled (autosave never clears
   * `isUntitled`/`isDirty` — only explicit Save/Save-As do). Inside
   * AUTOSAVE_GIT_COALESCE_MS of the previous autosave commit, BOTH the git
   * commit AND the snapshot push are skipped with the SAME predicate
   * (shouldCoalesceAutosaveCommit) so snapshotStack stays aligned with the git
   * log and coalesced flushes (durability-only, ~2.5s cadence) cannot evict
   * the 50-cap undo ring. A skipped snapshot also preserves redoStack (no new
   * undo step). Manual calls (no opts) ALWAYS commit + push snapshot — never
   * add default coalescing here (history-count suites depend on it).
   *
   * UX-C item 1 NO-OP GATE (documented decision): an AUTOSAVE whose live
   * document.xml hashes identically to the last committed hash is a full
   * no-op — no snapshot, no file write, no git commit, NO dirty flip, no
   * coalesce-clock update. Rationale: an empty draft must stay at
   * commitCount==1 (no noise banner) and stay clean so the close gate does
   * not prompt on a pristine doc. Manual commits bypass the gate (explicit
   * intent always commits). Coalescing is evaluated only for content-changed
   * autosaves, so its window semantics are untouched.
   */
  public async commitTransaction(commitMessage: string = 'Commit Transaction', opts?: { autosave?: boolean }): Promise<void> {
    if (!this.parser || !this.activeDocPath) {
      throw new Error("No active document.");
    }
    this.transactionSnapshot = null;
    // UX-C item 1: empty-autosave no-op gate (autosave only).
    if (opts?.autosave === true && this.lastCommittedXmlHash !== null) {
      try {
        const liveXml = this.parser.readDocumentXml();
        if (hashDocumentXml(liveXml) === this.lastCommittedXmlHash) {
          return;
        }
      } catch { /* hashing must never block a save — fall through */ }
    }
    // Coalesce snapshots WITH git commits (same predicate): a coalesced
    // autosave writes the file for durability but creates no new undo step,
    // keeping snapshotStack.length aligned with the git log.
    const coalesced =
      opts?.autosave === true &&
      shouldCoalesceAutosaveCommit(this.lastAutosaveCommitAt, Date.now(), AUTOSAVE_GIT_COALESCE_MS);
    const snapshot = coalesced ? null : this.parser.createAtomicSnapshot(commitMessage);

    const newDocxBuf = this.parser.rezip();
    await this.persistValidated(this.activeDocPath, newDocxBuf);
    if (snapshot) this.pushSnapshot(snapshot);

    if (!coalesced) {
      const updatedXml = this.parser.readDocumentXml();
      await commitChange(this.activeDocPath, updatedXml, commitMessage, this.collectPackageExtras());
      try {
        this.lastCommittedXmlHash = hashDocumentXml(updatedXml);
      } catch { /* ignore */ }
      if (opts?.autosave === true) this.lastAutosaveCommitAt = Date.now();
    } else {
      // Coalesced durability write: hash stays at the last COMMIT (not the
      // live bytes) so the next out-of-window autosave still commits them.
    }
    this.isDirty = true;
  }

  /**
   * Test hook (UX-C item 1): current no-op gate baseline. Null until the
   * first commit baseline is recorded.
   */
  public getLastCommittedXmlHashForTest(): string | null {
    return this.lastCommittedXmlHash;
  }

  /**
   * Applies an XML update to a node, saves to disk, and commits to undo history.
   */
  public async updateNode(nodeId: string, newXml: string, commitMessage: string = 'Update node'): Promise<void> {
    if (!this.parser || !this.activeDocPath) {
      throw new Error("No active document.");
    }

    // Update in-memory DOM
    this.parser.updateNode(nodeId, newXml);

    const snapshot = this.parser.createAtomicSnapshot(commitMessage);

    // Re-zip and save out the docx file (validated; a refusal reverts the node edit)
    const newDocxBuf = this.parser.rezip();
    await this.persistValidated(this.activeDocPath, newDocxBuf);
    this.pushSnapshot(snapshot);

    // Add to Git undo stack
    const updatedXml = this.parser.readDocumentXml();
    await commitChange(this.activeDocPath, updatedXml, commitMessage, this.collectPackageExtras());
    try {
      this.lastCommittedXmlHash = hashDocumentXml(updatedXml);
    } catch { /* ignore */ }
    this.isDirty = true;
  }

  /**
   * Saves the document in place to its current active path.
   */
  public async saveDocument(
    options?: { applyTypography?: boolean; typographyOptions?: TypographyOptions }
  ): Promise<{ astraDocPath: string; summary: StructuralSummary; typographyFixes: TypographicFix[]; styleWarnings: StyleWarning[] }> {
    if (!this.activeDocPath) {
      throw new Error("No active document path.");
    }
    return this.saveAsDocument(this.activeDocPath, options);
  }

  /**
   * UX-B close-gate Don't-Save path (UX-01 P0-2): deletes the untitled Drafts
   * `.docx` AND its sibling `.astra/` git dir (privacy — no pristine/partial
   * orphans left behind), then resets manager state so a dead path can never
   * be flushed/saved afterwards.
   *
   * SAFETY: only fires when `isUntitled && activeDocPath` are both set.
   * Titled user files are NEVER deleted here — Don't-Save on a titled doc
   * just closes (its last explicit save is already on disk). Missing files
   * are tolerated (force-remove); deletion itself is best-effort but the
   * state reset is unconditional.
   */
  public async discardUntitledDraft(): Promise<{ discarded: boolean; path: string | null }> {
    const doomed = this.activeDocPath;
    if (!this.isUntitled || !doomed) {
      return { discarded: false, path: doomed };
    }
    try {
      await fsp.rm(doomed, { force: true });
    } catch { /* best-effort: orphan GC owns leftovers */ }
    try {
      await fsp.rm(getRepoDir(doomed), { recursive: true, force: true });
    } catch { /* best-effort */ }
    this.activeDocPath = null;
    this.parser = null;
    this.snapshotStack = [];
    this.redoStack = [];
    this.transactionSnapshot = null;
    this.isDirty = false;
    this.isUntitled = false;
    this.lastCommittedXmlHash = null;
    this.lastAutosaveCommitAt = 0;
    return { discarded: true, path: doomed };
  }

  public getIsDirty(): boolean {
    return this.isDirty;
  }

  public getIsUntitled(): boolean {
    return this.isUntitled;
  }

  public markClean(): void {
    this.isDirty = false;
  }

  /**
   * Undoes the last operation using Git and atomic snapshot history.
   */
  public async undoLastAction(): Promise<{ success: boolean; atInitialState?: boolean }> {
    if (!this.parser || !this.activeDocPath) {
      throw new Error("No active document.");
    }

    try {
      // 1. Revert the git tree (document.xml + media/rels/styles/headers) and
      //    get the previous package state. Prefer undoPackage; fall back to
      //    legacy undo() when the git module is partially mocked (property
      //    access on a vitest mock throws, so resolve defensively).
      let undoPackageFn:
        | ((p: string) => Promise<{ xml: string; extra: Map<string, Buffer> }>)
        | undefined;
      try {
        undoPackageFn = asGitUndoPackage(gitMod).undoPackage;
      } catch {
        undoPackageFn = undefined;
      }
      let isTracked: ((name: string) => boolean) | undefined;
      try {
        isTracked = asGitUndoPackage(gitMod).isTrackedPackageFile;
      } catch {
        isTracked = undefined;
      }
      let reversedXml: string;
      let extra: Map<string, Buffer> = new Map();
      let usedPackage = false;
      if (typeof undoPackageFn === 'function') {
        const res = await undoPackageFn(this.activeDocPath);
        reversedXml = res.xml;
        extra = res.extra ?? new Map();
        usedPackage = true;
      } else {
        reversedXml = await undo(this.activeDocPath);
      }

      // 2. Rebuild the live OOXML package: tracked extras are replaced wholesale
      //    so deleted media vanishes and restored headers/rels reappear, then the
      //    DOMs (document + styles) are reloaded with fresh stable ids.
      //    OP-14: always restore the package snapshot when the package undo path
      //    was used — even when extras is empty — otherwise a deletion undo
      //    falls back to doc-only reload and stale media survives in the zip.
      //    Falls back to document-only reload for minimal parser doubles or the
      //    legacy undo() path.
      const parserCaps = asPackageSnapshotCapable(this.parser);
      if (
        usedPackage &&
        typeof parserCaps.capturePackageSnapshot === 'function' &&
        typeof parserCaps.restorePackageSnapshot === 'function' &&
        typeof isTracked === 'function'
      ) {
        const merged: Map<string, Buffer> = parserCaps.capturePackageSnapshot();
        for (const key of [...merged.keys()]) {
          if (key !== 'word/document.xml' && isTracked(key)) merged.delete(key);
        }
        for (const [k, v] of extra) merged.set(k, v);
        merged.set('word/document.xml', Buffer.from(reversedXml, 'utf8'));
        const restoreFn = parserCaps.restorePackageSnapshot;
        restoreFn(merged);
      } else {
        this.parser.reloadDocumentXml(reversedXml);
      }
      const revertedZipBuf = this.parser.rezip();

      // 3. Save buffer to disk
      await this.persistValidated(this.activeDocPath, revertedZipBuf);

      // 4. Update snapshot stack
      if (this.snapshotStack.length > 1) {
        const popped = this.snapshotStack.pop();
        if (popped) this.redoStack.push(popped);
      }
      // UX-C item 1: undo rewinds git head — re-baseline the no-op gate.
      try {
        this.lastCommittedXmlHash = hashDocumentXml(this.parser.readDocumentXml());
      } catch { /* ignore */ }
      return { success: true };
    } catch (err: unknown) {
      const rec = asRecord(err);
      const msg = typeof rec?.['message'] === 'string' ? (rec['message'] as string) : '';
      if (msg && msg.includes('already at the initial state')) {
        return { success: false, atInitialState: true };
      }
      throw err;
    }
  }

  /**
   * Accepts or rejects every tracked change and persists the result through
   * the validated write. Without an active file path the change stays in memory.
   */
  public async applyTrackedChanges(mode: 'accept' | 'reject'): Promise<void> {
    if (!this.parser) throw new Error("No active document.");
    this.parser.applyTrackedChanges(mode);
    if (!this.activeDocPath) return;
    await this.persistValidated(this.activeDocPath, this.parser.rezip());
    await commitChange(this.activeDocPath, this.parser.readDocumentXml(), `Tracked changes ${mode}`, this.collectPackageExtras());
    this.markClean();
  }

  public getActiveDocPath(): string | null {
    return this.activeDocPath;
  }

  public getParser(): DocxParser {
    if (!this.parser) throw new Error("No active document.");
    return this.parser;
  }

  public async getVersionHistory() {
    if (!this.activeDocPath) throw new Error("No active document.");
    return getHistory(this.activeDocPath);
  }

  public getSnapshotHistory(): DocumentSnapshot[] {
    return [...this.snapshotStack];
  }

  public getRedoHistory(): DocumentSnapshot[] {
    return [...this.redoStack];
  }

  public canUndo(): boolean {
    return this.snapshotStack.length > 1;
  }

  public canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * OP-15: re-applies the most recently undone snapshot. Pops the redoStack,
   * restores the parser to that snapshot, pushes it back onto the snapshot
   * stack, persists to disk, and re-commits to git so undo/redo stay symmetric.
   * (Git history cannot move forward after undo rewound the branch pointer, so
   * redo re-commits the restored package as a new head commit.)
   */
  public async redoLastAction(): Promise<{ success: boolean; atLatestState?: boolean }> {
    if (!this.parser || !this.activeDocPath) {
      throw new Error("No active document.");
    }
    const redoSnapshot = this.redoStack.pop();
    if (!redoSnapshot) {
      return { success: false, atLatestState: true };
    }
    this.parser.restoreAtomicSnapshot(redoSnapshot);

    const revertedZipBuf = this.parser.rezip();
    try {
      await this.persistValidated(this.activeDocPath, revertedZipBuf);
    } catch (err: unknown) {
      this.redoStack.push(redoSnapshot);
      throw err;
    }
    this.snapshotStack.push(redoSnapshot);

    const updatedXml = this.parser.readDocumentXml();
    await commitChange(this.activeDocPath, updatedXml, `Redo: ${redoSnapshot.label}`, this.collectPackageExtras());
    try {
      this.lastCommittedXmlHash = hashDocumentXml(updatedXml);
    } catch { /* ignore */ }
    this.isDirty = true;
    return { success: true };
  }

  /**
   * OP-14: collect versioned package extras (media/rels/styles/headers) for git.
   * document.xml is committed separately; everything matching
   * isTrackedPackageFile is snapshotted alongside it.
   */
  private collectPackageExtras(): Record<string, Buffer> {
    if (!this.parser) return {};
    try {
      let isTracked: ((name: string) => boolean) | undefined;
      try {
        isTracked = asGitUndoPackage(gitMod).isTrackedPackageFile;
      } catch {
        return {};
      }
      const parserCaps = asPackageSnapshotCapable(this.parser);
      if (typeof parserCaps.capturePackageSnapshot !== 'function') return {};
      if (typeof isTracked !== 'function') return {};
      const snap: Map<string, Buffer> = parserCaps.capturePackageSnapshot();
      const extras: Record<string, Buffer> = {};
      for (const [name, buf] of snap) {
        if (isTracked(name)) extras[name] = buf;
      }
      return extras;
    } catch {
      return {};
    }
  }
}

// ── S1b / EX-01+EX-02: sandbox path helper ────────────────────────────────────
// Renderer-supplied customPath / open-document-path values previously reached
// fs.writeFile/readFile with only an existsSync check: arbitrary read/write,
// traversal (../), null-byte truncation, non-docx payloads. Interactive callers
// (no customPath) are forced through the native save dialog in electron/main.ts
// (user picks the destination — registered via registerTrustedDocPath below).
// API callers (customPath present) MUST pass validateSandboxDocumentPath below:
// .docx extension + traversal/null-byte rejection + allowed-roots sandbox +
// symlink canonicalization + dot-dir rejection.
//
// S1b follow-up hardening (brutal-review survivors):
//   1. SYMLINK ESCAPE (was CRITICAL): path.resolve() is lexical — a symlink
//      inside the sandbox pointing outside (evil.docx -> /etc/secret.docx)
//      PASSED the old check. Now every candidate is canonicalized with
//      fs.realpathSync (symlinks/junctions resolved) — falling back to
//      realpath-of-nearest-existing-ancestor for not-yet-created save targets
//      — and containment is checked on the CANONICAL path.
//   2. HOME BREADTH (was HIGH): os.homedir() as a root made ~/.ssh/*.docx,
//      ~/.config/*.docx, ~/.aws/*.docx writable. REMOVED. Roots are now
//      NARROW (S1b follow-up): cwd/documents, os.tmpdir(), app-managed Drafts
//      dirs, ASTRA_EXTRA_DOC_ROOTS, plus per-file user-intent (dialog) trust.
//      The bare cwd() root was deliberately dropped: it made the entire repo
//      tree (.git, node_modules, electron/, backend/, …) writable via direct
//      IPC. Legit repo-adjacent locations (OneDrive-redirected Documents,
//      D:\Shared, corp-NAS, /tmp) are reached via dialog trust (item 5), not
//      via a blanket root. Direct IPC paths outside the narrow roots stay
//      sandboxed.
//   3. DOT-DIRS: any dot-prefixed path segment (.*) in the root-relative
//      remainder (.ssh, .config, .git, .aws) is rejected even inside an
//      allowed root — EXCEPT via explicit dialog trust (user intent wins).
//   4. TOCTOU: validate-at-IPC-entry is necessary but not sufficient — a
//      symlink can be swapped between validation and open/write. There is no
//      lock-based fix available here (plain fs calls, no O_NOFOLLOW in Win32
//      Node API surface used), so DocumentManager.openDocument /
//      createNewDocument / saveAsDocument / saveDocument RE-VALIDATE at use
//      time via assertSandboxPathAtUse(). This narrows the race window to the
//      validated-open itself; the residual (privileged local attacker racing
//      the open) is documented, not closable, without file locking.
//   5. USER INTENT IS THE ALLOWLIST: any location the user picked via the
//      native dialog bypasses the sandbox by design (D:\Shared, OneDrive-
//      redirected Documents, corp-NAS mounts, /tmp on POSIX — all reachable
//      interactively). electron/main.ts MUST register every dialog result
//      (and the app-generated silent-Drafts fallback) with
//      registerTrustedDocPath(). Direct IPC paths (open-document-path /
//      customPath) with no dialog intent stay sandboxed.
export const S1B_ALLOWED_DOC_EXTENSIONS = new Set(['.docx']);

/** Delimiter-separated extra sandbox roots (corp-NAS, redirected Documents).
 * H2: existence-checked — entries must resolve to an existing directory or
 * they never join the jail (phantom entries widen nothing). Explicit
 * `--allowed-dir`-style configuration flows in here; `ASTRA_ALLOWED_DIR`
 * is dropped per D3 and never read. */
function getSandboxExtraRoots(): string[] {
  const raw = (process.env.ASTRA_EXTRA_DOC_ROOTS || '').trim();
  if (!raw) return [];
  const out: string[] = [];
  for (const part of raw.split(path.delimiter)) {
    const p = part.trim();
    if (!p) continue;
    try {
      const resolved = path.resolve(p);
      if (!fs.statSync(resolved).isDirectory()) continue;
      out.push(resolved);
    } catch { /* missing or unstatable — never joins the jail */ }
  }
  return out;
}

/** App-managed Drafts candidates (trusted by construction, not renderer input). */
function getAppManagedDraftRoots(): string[] {
  const roots: string[] = [];
  try {
    roots.push(path.resolve(process.cwd(), 'documents', 'Drafts'));
  } catch { /* ignore */ }
  try {
    roots.push(path.resolve(os.tmpdir(), 'ASTra'));
  } catch { /* ignore */ }
  // Electron userData Drafts (e.g. %APPDATA%/ASTra/Drafts, ~/.config/ASTra/Drafts).
  try {
    if (process.env.APPDATA) roots.push(path.resolve(process.env.APPDATA, 'ASTra', 'Drafts'));
  } catch { /* ignore */ }
  try {
    const xdg = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config');
    roots.push(path.resolve(xdg, 'ASTra', 'Drafts'));
  } catch { /* ignore */ }
  return roots;
}

export function getSandboxAllowedRoots(): string[] {
  const roots: string[] = [];
  // NOTE: bare process.cwd() deliberately NOT included (S1b follow-up item 2:
  // the whole repo tree must not be IPC-writable). The workspace documents
  // tree is the narrow root; everything else interactive goes via dialog
  // trust or ASTRA_EXTRA_DOC_ROOTS.
  try {
    roots.push(path.resolve(process.cwd(), 'documents'));
  } catch { /* ignore */ }
  try {
    roots.push(path.resolve(os.tmpdir()));
  } catch { /* ignore */ }
  for (const r of getAppManagedDraftRoots()) roots.push(r);
  for (const r of getSandboxExtraRoots()) roots.push(r);
  // NOTE: os.homedir() deliberately NOT included (S1b follow-up item 2:
  // ~/.ssh, ~/.config, ~/.aws became writable). Home-adjacent legit locations
  // (Desktop/Documents/Downloads incl. OneDrive redirects) are reached via
  // dialog trust, not via a blanket home root.
  return [...new Set(roots)];
}

// ── User-intent trust registry ───────────────────────────────────────────────
// Resolved absolute path -> source. Populated by electron/main.ts for every
// native-dialog result and for the app-generated silent-Drafts fallback.
// Consulted by BOTH the IPC-entry validator and the at-use re-validator, so a
// dialog-picked file outside the sandbox (D:\Shared, corp-NAS) opens/saves
// without tripping the gate, while direct IPC paths stay sandboxed.
// Trust is per-operation with expiry: entries live for TRUST_TTL_MS then
// require re-confirmation. Entries covering hidden dot-segments or symlinked
// locations are single-use (consumed on first successful validation) so the
// next operation re-confirms explicitly. listTrustedDocPaths exposes the
// visible set; registration is logged.
export type TrustedDocPathSource = 'dialog' | 'app-internal';
/** Default lifetime for a dialog trust entry (re-confirm after expiry). */
export const TRUST_TTL_MS = 10 * 60 * 1000;
interface TrustedEntry {
  source: TrustedDocPathSource;
  expiresAt: number;
}
const trustedDocPaths = new Map<string, TrustedEntry>();

function trustNow(): number {
  return Date.now();
}

function purgeExpiredTrust(now: number = trustNow()): void {
  for (const [k, v] of trustedDocPaths) {
    if (v.expiresAt <= now) trustedDocPaths.delete(k);
  }
}

/** True when the resolved path needs re-confirmation per use (dot/symlink). */
function isSensitiveTrustPath(resolved: string, canonical: string): boolean {
  const segs = resolved.split(path.sep).filter((s) => s && s !== '.' && s !== '..');
  for (const seg of segs) {
    if (seg.startsWith('.')) return true;
  }
  if (canonical !== resolved) return true;
  const base = path.basename(resolved);
  if (base.startsWith('.')) return true;
  return false;
}

function removeTrustKeys(keys: string[]): void {
  for (const k of keys) trustedDocPaths.delete(k);
}

export function registerTrustedDocPath(rawPath: string, source: TrustedDocPathSource): string {
  const resolved = path.resolve(String(rawPath));
  const expiresAt = trustNow() + TRUST_TTL_MS;
  const entry: TrustedEntry = { source, expiresAt };
  trustedDocPaths.set(resolved, entry);
  // Also trust the canonical form so a symlinked dialog pick matches at use.
  try {
    trustedDocPaths.set(fs.realpathSync(resolved), entry);
  } catch { /* target may not exist yet (save dialog) — resolved form suffices */ }
  // Parent-dir realpath form: covers save targets whose file doesn't exist yet
  // but whose directory does (canonicalized below in canonicalizeForSandbox).
  try {
    const parentReal = fs.realpathSync(path.dirname(resolved));
    trustedDocPaths.set(path.join(parentReal, path.basename(resolved)), entry);
  } catch { /* ignore */ }
  // Silent by design: the trust registry (listTrustedDocPaths) is the typed
  // channel; CLI remains the logging surface.
  return resolved;
}

/** Visible trust set (expired entries pruned first). */
export function listTrustedDocPaths(): Array<{ path: string; source: TrustedDocPathSource; expiresAt: number }> {
  purgeExpiredTrust();
  const out: Array<{ path: string; source: TrustedDocPathSource; expiresAt: number }> = [];
  for (const [p, e] of trustedDocPaths) {
    out.push({ path: p, source: e.source, expiresAt: e.expiresAt });
  }
  out.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
  return out;
}

export function isTrustedDocPath(candidate: string): boolean {
  let resolved: string;
  try {
    resolved = path.resolve(candidate);
  } catch {
    return false;
  }
  purgeExpiredTrust();
  if (trustedDocPaths.has(resolved)) return true;
  const canonical = canonicalizeForSandbox(resolved);
  if (trustedDocPaths.has(canonical)) return true;
  // A trusted DIRECTORY pick trusts its contents (defensive; dialog picks are
  // files today, but containment keeps this tight either way).
  for (const trusted of trustedDocPaths.keys()) {
    try {
      if (isPathInsideDir(canonical, trusted) || isPathInsideDir(resolved, trusted)) return true;
    } catch { /* ignore */ }
  }
  return false;
}

/** Consumes single-use trust for sensitive locations after a successful check. */
function consumeSensitiveTrust(resolved: string): void {
  let canonical = resolved;
  try {
    canonical = canonicalizeForSandbox(resolved);
  } catch { /* ignore */ }
  if (!isSensitiveTrustPath(resolved, canonical)) return;
  const keys = [resolved, canonical];
  try {
    const parentReal = fs.realpathSync(path.dirname(resolved));
    keys.push(path.join(parentReal, path.basename(resolved)));
  } catch { /* ignore */ }
  removeTrustKeys(keys);
}

/** Test hook: clears the user-intent registry. */
export function clearTrustedDocPaths(): void {
  trustedDocPaths.clear();
}

// Q0a: jail primitives live in ./sandbox/paths (re-exported here for compat).
export { canonicalizeForSandbox };

/** Finds the allowed root containing `canonical`, if any. */
function findContainingRoot(canonical: string): string | undefined {
  for (const root of getSandboxAllowedRoots()) {
    let canonicalRoot = root;
    try {
      canonicalRoot = fs.realpathSync(root);
    } catch { /* root may not exist (cwd/documents) — use lexical form */ }
    if (isPathInsideDir(canonical, canonicalRoot) || isPathInsideDir(canonical, root)) {
      return root;
    }
  }
  return undefined;
}

/** Rejects dot-prefixed segments (.*) in the root-relative remainder. */
function assertNoDotSegments(canonical: string, root: string): void {
  const rel = path.relative(root, canonical);
  const segs = rel.split(path.sep).filter((s) => s && s !== '.' && s !== '..');
  for (const seg of segs) {
    if (seg.startsWith('.')) {
      throw new Error(
        `Document path traverses a hidden directory "${seg}" — pick the file via the open/save dialog for explicit consent`
      );
    }
  }
  // The basename itself must not be dot-prefixed either (.evil.docx).
  const base = path.basename(canonical);
  if (base.startsWith('.')) {
    throw new Error('Document path must not use a hidden (dot-prefixed) file name');
  }
}

export function isSandboxDocumentPathAllowed(candidate: string): boolean {
  try {
    validateSandboxDocumentPath(candidate);
    return true;
  } catch {
    return false;
  }
}

/**
 * Validates a renderer-supplied document path for customPath/open-document-path.
 * Returns the resolved absolute path. Throws on: non-string/empty, embedded
 * null byte (or %00), `..` path segment (lexical traversal), non-.docx
 * extension, hidden (dot-prefixed) segments, or canonical location outside
 * the sandbox roots. Symlinks/junctions are resolved BEFORE the containment
 * check (S1b follow-up item 1). Paths registered via registerTrustedDocPath
 * (native-dialog picks — user intent is the allowlist) bypass the sandbox.
 */
export function validateSandboxDocumentPath(rawPath: unknown): string {
  if (typeof rawPath !== 'string' || !rawPath.trim()) {
    throw new Error('Document path must be a non-empty string');
  }
  // Check the RAW value first: trim() would erase a trailing \n\r smuggling
  // attempt before the check below ever sees it.
  if (rawPath.includes('\0') || /%00/i.test(rawPath)) {
    throw new Error('Document path contains a null byte');
  }
  if (rawPath.includes('\n') || rawPath.includes('\r')) {
    throw new Error('Document path contains forbidden line break');
  }
  const trimmed = rawPath.trim();
  if (trimmed.includes('\0') || /%00/i.test(trimmed)) {
    throw new Error('Document path contains a null byte');
  }
  const segments = trimmed.split(/[\\/]/);
  if (segments.includes('..')) {
    throw new Error('Document path contains forbidden ".." traversal');
  }
  const ext = path.extname(trimmed).toLowerCase();
  if (!S1B_ALLOWED_DOC_EXTENSIONS.has(ext)) {
    throw new Error(`Document path must use a .docx extension (got "${ext || '(none)'}")`);
  }
  const resolved = path.resolve(trimmed);
  const resolvedSegments = resolved.split(path.sep);
  if (resolvedSegments.includes('..')) {
    throw new Error('Document path resolves outside its directory');
  }
  // User-intent bypass: dialog-picked (or app-internal fallback) paths are
  // trusted at ANY location (D:\Shared, OneDrive redirects, corp-NAS).
  // Trust expires; sensitive locations (dot-segments, symlinks) are
  // single-use and require re-confirmation on the next operation.
  if (isTrustedDocPath(resolved)) {
    consumeSensitiveTrust(resolved);
    return resolved;
  }
  // Symlink escape closure: canonicalize (target if it exists, else nearest
  // existing ancestor) and check containment on the CANONICAL path.
  const canonical = canonicalizeForSandbox(resolved);
  const containingRoot = findContainingRoot(canonical);
  if (!containingRoot) {
    throw new Error('Document path is outside the allowed sandbox directories');
  }
  // Dot-dir gate on the root-relative remainder (blocks ~/.ssh-style escapes
  // even when they sit inside an allowed root, e.g. cwd/.ssh/x.docx).
  assertNoDotSegments(canonical, containingRoot);
  return resolved;
}

/**
 * At-use re-validation (S1b follow-up item 1, TOCTOU narrowing). Called by
 * DocumentManager.openDocument/createNewDocument/saveAsDocument immediately
 * before touching the filesystem: re-canonicalizes the CURRENT on-disk state
 * (catching symlink swaps that happened after IPC-entry validation) and
 * re-applies sandbox + dot-dir + trust checks. Returns the resolved path.
 * NOT a full TOCTOU close (no file locking) — documented residual.
 */
export function assertSandboxPathAtUse(rawPath: string): string {
  return validateSandboxDocumentPath(rawPath);
}