// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 2275-2290 + 2300-3029).
// Zero logic edits (sidecar/COMMIT_OID_RE/refresh/restore helpers live in ../diff and ../errors).
import * as fsp from 'fs/promises';
import {
  assertWriteTargetNoSymlink,
  checkoutCommit,
  commitChange,
  getEngineErrorCode,
  getHistory,
  getSidecarHistoryWarning,
  isTrackedPackageFile,
  validateDocxBuffer,
  writeFileNoFollow,
} from '@attentionform/astra-engine';
import type { McpSessionStore } from '../../session-store';
import { collectTrackedExtrasForGit, ensureGitRepo } from '../../session-store';
import {
  mcpError,
  mcpErrorFrom,
  sanitizeCommitMessage,
  sanitizeCommitAuthor,
  COMMIT_MESSAGE_MAX,
  invalidInput,
  unknownSessionError,
} from '../errors';
import type { McpErrorShape } from '../errors';
import {
  AUDIT_FILES_MAX,
  COMMIT_OID_RE,
  diffChangedValues,
  diffChangedValuesAcrossParse,
  parseTargetIdTextsForPreview,
  collectAstIdTexts,
  parseSavedIdTexts,
  summarizeDiscardedUnsaved,
  diffIdTexts,
  formatRevertCommitMessage,
  refreshHandleAstInPlace,
  collectTableNodeIdsInOrder,
  invalidRestoreError,
  logRestoreFailure,
  honestRestoreRemedy,
  pathDirname,
  sidecarRepoExists,
} from '../diff';
import type { ChangedValue, DiscardedUnsavedSummary } from '../diff';
import {
  GetDocumentHistorySchema,
  RevertDocumentSchema,
  RedoDocumentSchema,
  PreviewRevertSchema,
} from '../schemas';

// ── Git-history handlers (history browsing + preview + revert + redo) ────────
export interface HistoryEntry {
  oid: string;
  message: string;
  author: string;
  timestamp: Date;
  /** Ms-precision epoch millis (audit-bridge meta when tracked, else git seconds). */
  timestampMs: number;
  /** Operations folded into this commit (0 for pre-tracking commits and heals-only saves). */
  opCount?: number;
  /** Node ids touched by this commit, bounded ([] for pre-tracking commits). */
  affectedNodeIds?: string[];
  /** Head oid resulting from this commit (== oid; append-only, refs never move). */
  resultingHeadOid: string;
}


export async function handleGetDocumentHistory(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      documentPath: string;
      currentCommit: string | null;
      totalCommits: number;
      history: HistoryEntry[];
      /** R-6 sidecar-removal warning (string when a regenerated chain replaced recorded history, else null). Additive, never a failure. */
      warning: string | null;
    }
  | McpErrorShape
> {
  const parsed = GetDocumentHistorySchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`get_document_history: invalid input. ${detail}`);
  }
  const limit = parsed.data.limit ?? 20;
  const offset = parsed.data.offset ?? 0;
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      // Serialized with save/revert/redo commits on this session: the path is
      // snapshotted inside the per-handle mutex so readers never observe a
      // half-moved file target. Read-only: revision markers, dirty state and
      // the redo stack are left untouched.
      const filePath = handle.filePath;
      const sessionId = handle.sessionId;
      // A session whose sidecar repo is absent (buffer-only handle, never saved)
      // simply has no history yet.
      if (!sidecarRepoExists(filePath)) {
        let emptyWarning: string | null = null;
        try {
          emptyWarning = getSidecarHistoryWarning(filePath);
        } catch {
          emptyWarning = null;
        }
        return {
          ok: true as const,
          sessionId,
          documentPath: filePath,
          currentCommit: null,
          totalCommits: 0,
          history: [],
          warning: emptyWarning,
        };
      }
      const full = await getHistory(filePath);
      // R-6: carry the engine tombstone warning into the envelope. The
      // engine exposes it on the history array itself; read it there first
      // and fall back to the direct accessor (same source, no duplication).
      let warning: string | null = null;
      try {
        const carried = (full as unknown as { warning?: unknown }).warning;
        warning = typeof carried === 'string' ? carried : getSidecarHistoryWarning(filePath);
      } catch {
        warning = null;
      }
      const slice = full.slice(offset, offset + limit).map((c) => {
        const meta = store.getCommitMeta(c.oid);
        const entry: HistoryEntry = {
          oid: c.oid,
          message: typeof c.message === 'string' ? c.message.replace(/\s+$/, '') : String(c.message ?? ''),
          author: typeof c.author === 'string' ? c.author : String(c.author ?? ''),
          timestamp: meta ? new Date(meta.timestampMs) : c.timestamp,
          timestampMs: meta ? meta.timestampMs : c.timestamp.getTime(),
          // Always present: 0/[] covers pre-tracking commits and heals-only
          // saves (opCount folds operations, not normalizer rewrites).
          opCount: meta ? meta.opCount : 0,
          affectedNodeIds: meta ? [...meta.affectedNodeIds] : [],
          resultingHeadOid: c.oid,
        };
        return entry;
      });
      return {
        ok: true as const,
        sessionId,
        documentPath: filePath,
        currentCommit: full.length > 0 ? full[0].oid : null,
        totalCommits: full.length,
        history: slice,
        warning,
      };
    });
  } catch (err: unknown) {
    const code = (err as { code?: unknown } | null)?.code;
    // A liveness failure means the session was dropped while this read was
    // queued behind an in-flight operation: report it cleanly as a conflict.
    if (code === 'E_CONFLICT') {
      return mcpErrorFrom(err, 'E_CONFLICT', 'The session changed while history was reading. Re-open if needed, then retry get_document_history.');
    }
    if (code === 'E_SESSION' || /Unknown session|was closed concurrently|expired/.test(err instanceof Error ? err.message : String(err))) {
      return unknownSessionError(parsed.data.sessionId);
    }
    return mcpErrorFrom(err, 'E_IO', 'Retry get_document_history; re-open the document if the session expired.');
  }
}

export async function handleRevertDocument(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      path: string;
      revertedToCommit: string;
      revision: number;
      dirty: false;
      /** Bounded echo of the unsaved edits discarded by this revert (zero summary when clean). */
      discardedUnsaved: DiscardedUnsavedSummary;
      commitOid: string;
      commitMessage: string;
    }
  | (McpErrorShape & { dirtyReason?: string })
> {
  const parsed = RevertDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`revert_document: invalid input. ${detail}`);
  }
  const rawObj = (typeof input === 'object' && input !== null ? input : {}) as Record<string, unknown>;
  const stepsProvided = 'steps' in rawObj && rawObj['steps'] !== undefined;
  const oidProvided = parsed.data.commitOid !== undefined;
  if (stepsProvided && oidProvided) {
    return invalidInput(
      'revert_document: steps and commitOid are mutually exclusive. Pass {steps:N} to walk back OR {commitOid:"<40-hex>"} to jump, not both.'
    );
  }
  if (oidProvided && !COMMIT_OID_RE.test(parsed.data.commitOid as string)) {
    return mcpError(
      'E_VALIDATION',
      `revert_document: commit "${parsed.data.commitOid}" is not a valid 40-hex-char object id`,
      'Call get_document_history to list valid commit oids, then retry revert_document with one of them.'
    );
  }
  const steps = parsed.data.steps ?? 1;
  const discardUnsaved = parsed.data.discardUnsaved ?? false;
  // Author/reason hygiene matches the save path sanitizers (single-line,
  // capped; blank author falls back to the default). Presence (not value)
  // selects the T-5 subject shape so legacy message-only callers keep the
  // legacy `Revert to <oid>` subject byte-identical.
  const authorName = sanitizeCommitAuthor(parsed.data.author);
  const authorProvided = 'author' in rawObj && rawObj['author'] !== undefined;
  const reasonProvided = 'reason' in rawObj && parsed.data.reason !== undefined;
  const reasonText = reasonProvided ? sanitizeCommitMessage(parsed.data.reason, '') : '';
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      if (handle.revision !== handle.savedRevision && discardUnsaved === false) {
        const pending = handle.revision - handle.savedRevision;
        throw Object.assign(
          new Error(
            `revert_document: session holds unsaved edits (revision ${handle.revision} vs saved ${handle.savedRevision}); refusing to discard them`
          ),
          {
            code: 'E_CONFLICT',
            remedy:
              'Bare reverts fail closed when dirty: net-zero edits still count as dirty ' +
              '(the flag tracks pending revisions, not a content diff) — save the edits via save_document ' +
              'or retry with discardUnsaved:true to drop them.',
            dirtyReason:
              `${pending} pending revision(s) since the last save (revision ${handle.revision} vs saved ` +
              `${handle.savedRevision}); net-zero content still counts as dirty until saved or discarded.`,
          }
        );
      }
      if (!sidecarRepoExists(handle.filePath)) {
        throw Object.assign(new Error('revert_document: no saved history for this session yet'), {
          code: 'E_NOT_FOUND',
          remedy: 'Call save_document first to record a history baseline, then retry revert_document.',
        });
      }
      let full: Awaited<ReturnType<typeof getHistory>>;
      try {
        full = await getHistory(handle.filePath);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw Object.assign(new Error(`revert_document: cannot read history: ${msg}`), {
          code: 'E_IO',
          remedy: 'Retry revert_document; re-open the document if the session expired.',
        });
      }
      let targetOid: string;
      if (oidProvided) {
        const wanted = parsed.data.commitOid as string;
        if (!full.some((c) => c.oid === wanted)) {
          throw Object.assign(new Error(`revert_document: commit "${wanted}" not found in history`), {
            code: 'E_NOT_FOUND',
            remedy: 'Call get_document_history to list valid commit oids, then retry revert_document with one of them.',
          });
        }
        targetOid = wanted;
      } else {
        if (steps >= full.length) {
          throw Object.assign(
            new Error(`revert_document: steps ${steps} exceeds saved history depth ${full.length}`),
            {
              code: 'E_NOT_FOUND',
              remedy: `The session has ${full.length} saved commit(s). Retry with steps < ${full.length} or pick a commitOid from get_document_history.`,
            }
          );
        }
        targetOid = full[steps].oid;
      }
      // Steps-back from the pre-revert head (drives the T-5 subject + op count).
      const stepsBack = oidProvided ? full.findIndex((c) => c.oid === targetOid) : steps;
      // Snapshot the live package BEFORE touching it (redo source), plus the
      // revision markers, redo-stack depth and on-disk bytes, so any failure
      // below can leave no trace behind.
      const preRevertSnap = handle.parser.capturePackageSnapshot();
      const prevRevision = handle.revision;
      const prevSavedRevision = handle.savedRevision;
      const prevRedoLen = handle.redoStack.length;
      const prevByteSize = handle.byteSize;
      let prevFileBytes: Buffer | null = null;
      try {
        prevFileBytes = await fsp.readFile(handle.filePath);
      } catch {
        prevFileBytes = null;
      }
      // Discard echo: capture the dirty diff BEFORE the restore drops it.
      // wasDirty mirrors the inspect/export dirty bit (revision vs saved).
      // Read-only probes only — they never throw into the revert path.
      const wasDirtyBefore = handle.revision !== handle.savedRevision;
      const liveBefore = collectAstIdTexts(handle.ast);
      let discardedUnsaved: DiscardedUnsavedSummary;
      try {
        const savedTexts = wasDirtyBefore
          ? await parseSavedIdTexts(prevFileBytes)
          : new Map<string, string>();
        discardedUnsaved = wasDirtyBefore
          ? summarizeDiscardedUnsaved(liveBefore, savedTexts, handle.revision - handle.savedRevision)
          : {
              wasDirty: false,
              discardedNodeCount: 0,
              discardedNodeIds: [],
              truncated: false,
            };
      } catch {
        discardedUnsaved = {
          wasDirty: wasDirtyBefore,
          discardedNodeCount: wasDirtyBefore ? Math.max(0, handle.revision - handle.savedRevision) : 0,
          discardedNodeIds: [],
          truncated: false,
        };
      }
      let target: Awaited<ReturnType<typeof checkoutCommit>>;
      try {
        target = await checkoutCommit(handle.filePath, targetOid);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw Object.assign(new Error(`revert_document: cannot restore commit "${targetOid}": ${msg}`), {
          code: /not found in history/.test(msg) ? 'E_NOT_FOUND' : 'E_IO',
          remedy: 'Call get_document_history to list valid commit oids, then retry revert_document with one of them.',
        });
      }
      const merged = new Map<string, Buffer>(preRevertSnap);
      for (const key of [...merged.keys()]) {
        if (key !== 'word/document.xml' && isTrackedPackageFile(key)) merged.delete(key);
      }
      for (const [k, v] of target.extra) merged.set(k, v);
      merged.set('word/document.xml', Buffer.from(target.xml, 'utf8'));
      // Unified atomic section: ANY throw after staging restores markers,
      // stack depth, memory and disk bytes. Only validation failures keep
      // E_VALIDATION; everything else surfaces as E_IO with a retry remedy.
      let writeAttempted = false;
      let commitPhase = false;
      try {
        handle.parser.restorePackageSnapshot(merged);
        await refreshHandleAstInPlace(handle);
        const buf = handle.parser.rezip();
        const candidate = await validateDocxBuffer(buf);
        if (!candidate.ok) {
          throw invalidRestoreError('revert_document', candidate.errors, 'restore candidate', collectTableNodeIdsInOrder(handle.ast));
        }
        // M1: lstat the leaf (refuse symlinks incl. dangling) before mkdir
        // (mkdir -p through a symlink parent would create outside dirs) and
        // rely on the O_NOFOLLOW write below for the check-then-use window.
        assertWriteTargetNoSymlink(handle.filePath);
        writeAttempted = true;
        await fsp.mkdir(pathDirname(handle.filePath), { recursive: true });
        assertWriteTargetNoSymlink(handle.filePath);
        await writeFileNoFollow(handle.filePath, buf);
        handle.byteSize = buf.length;
        const revision = store.bumpRevision(handle);
        // The pre-revert package is the redo source (single-step reverts push
        // too, so redo works after the common case).
        handle.redoStack.push(preRevertSnap);
        const restoredXml = handle.parser.readDocumentXml();
        const extras = collectTrackedExtrasForGit(handle.parser);
        const shortOid = targetOid.slice(0, 7);
        // T-5 subject when author/reason are explicitly given; otherwise the
        // legacy shape stays byte-identical for existing message-only callers.
        // A bare default revert quotes the target commit message so the new
        // head names what it restored to.
        let commitMessage: string;
        if (authorProvided || reasonProvided) {
          commitMessage = formatRevertCommitMessage(stepsBack, targetOid, authorName, reasonText);
        } else {
          const customMessage =
            parsed.data.message !== undefined ? sanitizeCommitMessage(parsed.data.message, '') : '';
          const targetEntry = full.find((c) => c.oid === targetOid);
          const targetMsg =
            typeof targetEntry?.message === 'string'
              ? targetEntry.message.replace(/\s+$/, '')
              : '';
          commitMessage = (
            customMessage
              ? `Revert to ${shortOid}: ${customMessage}`
              : targetMsg
                ? `Revert to ${shortOid}: "${targetMsg}"`
                : `Revert to ${shortOid}`
          ).slice(0, COMMIT_MESSAGE_MAX);
        }
        commitPhase = true;
        await ensureGitRepo(handle.filePath, restoredXml, extras);
        const commitOid = await commitChange(handle.filePath, restoredXml, commitMessage, extras, undefined, authorName);
        // Affected-node footprint: pre vs post restore diff (honest footprint).
        const liveAfter = collectAstIdTexts(handle.ast);
        store.recordCommitMeta(commitOid, {
          author: authorName,
          timestampMs: Date.now(),
          opCount: Math.max(0, stepsBack),
          affectedNodeIds: diffIdTexts(liveBefore.texts, liveAfter.order, liveAfter.texts),
        });
        store.markSaved(handle);
        return {
          ok: true as const,
          sessionId: handle.sessionId,
          path: handle.filePath,
          revertedToCommit: targetOid,
          revision,
          dirty: false as const,
          discardedUnsaved,
          commitOid,
          commitMessage,
        };
      } catch (stageErr: unknown) {
        // S3c: engine codes pass through verbatim (E_QUOTA from history quotas
        // stays loud); uncoded failures stay E_IO. Sibling pattern:
        // getEngineErrorCode (visual/export/compose) — reads .code first,
        // then [E_*]/E_* in the message.
        const preserved = getEngineErrorCode(stageErr);
        const isValidation = preserved === 'E_VALIDATION';
        handle.redoStack.length = prevRedoLen;
        handle.revision = prevRevision;
        handle.savedRevision = prevSavedRevision;
        handle.byteSize = prevByteSize;
        let memoryOk = true;
        try {
          handle.parser.restorePackageSnapshot(preRevertSnap);
          await refreshHandleAstInPlace(handle);
        } catch (restoreErr: unknown) {
          memoryOk = false;
          logRestoreFailure('revert_document', handle.filePath, 'memory', restoreErr);
        }
        let diskOk = true;
        // M1: primary-write symlink refusal (pre-commit) leaves disk
        // untouched — skip the disk restore (it would refuse again, and the
        // null-bytes branch would unlink the user's link) and preserve the
        // E_SANDBOX verdict verbatim below (save sibling pattern).
        if (writeAttempted && !(preserved === 'E_SANDBOX' && !commitPhase)) {
          try {
            if (prevFileBytes !== null) {
              await writeFileNoFollow(handle.filePath, prevFileBytes);
            } else {
              await fsp.unlink(handle.filePath);
            }
          } catch (restoreErr: unknown) {
            diskOk = false;
            logRestoreFailure('revert_document', handle.filePath, 'disk', restoreErr);
          }
        }
        if (preserved === 'E_SANDBOX' && !commitPhase) throw stageErr;
        if (isValidation) {
          const baseRemedy = typeof (stageErr as { remedy?: unknown }).remedy === 'string'
            ? String((stageErr as { remedy?: unknown }).remedy)
            : 'Fix the flagged content and retry revert_document.';
          if (!memoryOk) {
            throw Object.assign(stageErr instanceof Error ? stageErr : new Error(String(stageErr)), {
              code: 'E_VALIDATION',
              remedy: `${baseRemedy} restore attempted, VERIFY via inspect/history before continuing. Run inspect_document and get_document_history to confirm state before retrying.`,
            });
          }
          throw stageErr;
        }
        const rawMsg = stageErr instanceof Error ? stageErr.message : String(stageErr);
        const baseRemedy = typeof (stageErr as { remedy?: unknown }).remedy === 'string'
          ? String((stageErr as { remedy?: unknown }).remedy)
          : writeAttempted
            ? 'The document was left unchanged on disk and in memory. Retry revert_document to re-record history.'
            : 'Check the document directory is writable and retry revert_document.';
        throw Object.assign(new Error(rawMsg.startsWith('revert_document') ? rawMsg : `revert_document: operation failed, state restored: ${rawMsg}`), {
          code: preserved ?? 'E_IO',
          remedy: honestRestoreRemedy('revert_document', memoryOk, diskOk, baseRemedy, rawMsg),
        });
      }
    });
  } catch (err: unknown) {
    const code = (err as { code?: unknown } | null)?.code;
    if (code === 'E_SESSION' || /Unknown session|was closed concurrently|expired/.test(err instanceof Error ? err.message : String(err))) {
      return unknownSessionError(parsed.data.sessionId);
    }
    const out = mcpErrorFrom(err, 'E_IO', 'Retry revert_document; re-open the document if the session expired.');
    const dirtyReason = (err as { dirtyReason?: unknown } | null)?.dirtyReason;
    if (typeof dirtyReason === 'string' && dirtyReason) return { ...out, dirtyReason };
    return out;
  }
}

export async function handleRedoDocument(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      path: string;
      revision: number;
      dirty: false;
      commitOid: string;
      commitMessage: string;
      /** Advisory old→new value summaries (pre-redo live -> restored redo text; bounded). */
      changedValues: ChangedValue[];
      /** True when the count or per-string cap truncated changedValues. */
      changedValuesTruncated: boolean;
      /** Present when the value diff was not computable (empty changedValues). */
      changedValuesReason?: string;
    }
  | (McpErrorShape & { atLatestState?: true })
> {
  const parsed = RedoDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`redo_document: invalid input. ${detail}`);
  }
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      const peeked = handle.redoStack[handle.redoStack.length - 1];
      if (!peeked) {
        throw Object.assign(new Error('redo_document: nothing to redo (already at the latest state)'), {
          code: 'E_NOT_FOUND',
          atLatestState: true as const,
          remedy:
            'The redo stack is empty (a fresh execute_batch edit clears redo, and each redo consumes one entry): ' +
            'call revert_document first to create a redoable state, then retry redo_document.',
        });
      }
      // Author/reason hygiene matches the save/revert path (single-line,
      // capped; blank author falls back to the default). reason is preferred
      // over the legacy message when both are given.
      const rawRedo = (typeof input === 'object' && input !== null ? input : {}) as Record<string, unknown>;
      const redoReasonProvided = 'reason' in rawRedo && parsed.data.reason !== undefined;
      const redoReasonText = redoReasonProvided ? sanitizeCommitMessage(parsed.data.reason, '') : '';
      const authorName = sanitizeCommitAuthor(parsed.data.author);
      const redoBefore = collectAstIdTexts(handle.ast);
      // Snapshots for atomic rollback: revision markers, stack depth, the
      // live package and the on-disk bytes.
      const prevRevision = handle.revision;
      const prevSavedRevision = handle.savedRevision;
      const prevByteSize = handle.byteSize;
      const preRedoSnap = handle.parser.capturePackageSnapshot();
      let prevFileBytes: Buffer | null = null;
      try {
        prevFileBytes = await fsp.readFile(handle.filePath);
      } catch {
        prevFileBytes = null;
      }
      // Stage the redo in memory, then validate the exact bytes that would
      // land on disk. Unified atomic section: any throw after staging
      // restores markers, stack depth, memory and disk bytes. Only
      // validation failures keep E_VALIDATION; everything else is E_IO.
      const snap = peeked;
      let didPop = false;
      let writeAttempted = false;
      let commitPhase = false;
      try {
        handle.parser.restorePackageSnapshot(new Map<string, Buffer>(snap));
        await refreshHandleAstInPlace(handle);
        const buf = handle.parser.rezip();
        const candidate = await validateDocxBuffer(buf);
        if (!candidate.ok) {
          throw invalidRestoreError('redo_document', candidate.errors, 'redo candidate', collectTableNodeIdsInOrder(handle.ast));
        }
        // The candidate is valid: consume the stack entry (depth changes only
        // once validation has passed).
        handle.redoStack.pop();
        didPop = true;
        // M1: same dangling-symlink guard as revert (lstat chain +
        // O_NOFOLLOW write; the post-mkdir re-check closes the race window).
        assertWriteTargetNoSymlink(handle.filePath);
        writeAttempted = true;
        await fsp.mkdir(pathDirname(handle.filePath), { recursive: true });
        assertWriteTargetNoSymlink(handle.filePath);
        await writeFileNoFollow(handle.filePath, buf);
        handle.byteSize = buf.length;
        const revision = store.bumpRevision(handle);
        const restoredXml = handle.parser.readDocumentXml();
        const extras = collectTrackedExtrasForGit(handle.parser);
        const legacyMessage =
          parsed.data.message !== undefined ? sanitizeCommitMessage(parsed.data.message, '') : '';
        const customMessage = redoReasonText || legacyMessage;
        const commitMessage = (
          customMessage ? `Redo: ${customMessage}` : `redo_document: rev ${revision}`
        ).slice(0, COMMIT_MESSAGE_MAX);
        commitPhase = true;
        await ensureGitRepo(handle.filePath, restoredXml, extras);
        const commitOid = await commitChange(handle.filePath, restoredXml, commitMessage, extras, undefined, authorName);
        const redoAfter = collectAstIdTexts(handle.ast);
        // Advisory value diff (pre-redo live -> restored text): best-effort,
        // bounded, never fails the redo. Both sides share one session lineage,
        // so the id-keyed diff applies; a reparse surprise degrades to empty.
        let redoChangedValues: ChangedValue[] = [];
        let redoChangedValuesTruncated = false;
        let redoChangedValuesReason: string | undefined;
        try {
          const diff = diffChangedValues(redoBefore.texts, redoAfter.texts, redoAfter.order, redoBefore.order);
          redoChangedValues = diff.values;
          redoChangedValuesTruncated = diff.truncated;
        } catch {
          redoChangedValues = [];
          redoChangedValuesTruncated = false;
          redoChangedValuesReason = 'Value diff failed; treated as not computable.';
        }
        store.recordCommitMeta(commitOid, {
          author: authorName,
          timestampMs: Date.now(),
          opCount: 1,
          affectedNodeIds: diffIdTexts(redoBefore.texts, redoAfter.order, redoAfter.texts),
        });
        store.markSaved(handle);
        return {
          ok: true as const,
          sessionId: handle.sessionId,
          path: handle.filePath,
          revision,
          dirty: false as const,
          commitOid,
          commitMessage,
          changedValues: redoChangedValues,
          changedValuesTruncated: redoChangedValuesTruncated,
          ...(redoChangedValuesReason ? { changedValuesReason: redoChangedValuesReason } : {}),
        };
      } catch (stageErr: unknown) {
        // S3c: engine codes pass through verbatim (E_QUOTA from history quotas
        // stays loud); uncoded failures stay E_IO. Sibling pattern:
        // getEngineErrorCode (visual/export/compose) — reads .code first,
        // then [E_*]/E_* in the message.
        const preserved = getEngineErrorCode(stageErr);
        const isValidation = preserved === 'E_VALIDATION';
        if (didPop) handle.redoStack.push(snap);
        handle.revision = prevRevision;
        handle.savedRevision = prevSavedRevision;
        handle.byteSize = prevByteSize;
        let memoryOk = true;
        try {
          handle.parser.restorePackageSnapshot(preRedoSnap);
          await refreshHandleAstInPlace(handle);
        } catch (restoreErr: unknown) {
          memoryOk = false;
          logRestoreFailure('redo_document', handle.filePath, 'memory', restoreErr);
        }
        let diskOk = true;
        // M1: primary-write symlink refusal (pre-commit) leaves disk
        // untouched — skip the disk restore (it would refuse again, and the
        // null-bytes branch would unlink the user's link) and preserve the
        // E_SANDBOX verdict verbatim below (save sibling pattern).
        if (writeAttempted && !(preserved === 'E_SANDBOX' && !commitPhase)) {
          try {
            if (prevFileBytes !== null) {
              await writeFileNoFollow(handle.filePath, prevFileBytes);
            } else {
              await fsp.unlink(handle.filePath);
            }
          } catch (restoreErr: unknown) {
            diskOk = false;
            logRestoreFailure('redo_document', handle.filePath, 'disk', restoreErr);
          }
        }
        if (preserved === 'E_SANDBOX' && !commitPhase) throw stageErr;
        if (isValidation) {
          const baseRemedy = typeof (stageErr as { remedy?: unknown }).remedy === 'string'
            ? String((stageErr as { remedy?: unknown }).remedy)
            : 'Fix the flagged content and retry redo_document.';
          if (!memoryOk) {
            throw Object.assign(stageErr instanceof Error ? stageErr : new Error(String(stageErr)), {
              code: 'E_VALIDATION',
              remedy: `${baseRemedy} restore attempted, VERIFY via inspect/history before continuing. Run inspect_document and get_document_history to confirm state before retrying.`,
            });
          }
          throw stageErr;
        }
        const rawMsg = stageErr instanceof Error ? stageErr.message : String(stageErr);
        const baseRemedy = typeof (stageErr as { remedy?: unknown }).remedy === 'string'
          ? String((stageErr as { remedy?: unknown }).remedy)
          : writeAttempted
            ? 'The document was left unchanged on disk and in memory. Retry redo_document to re-record history.'
            : 'Check the document directory is writable and retry redo_document.';
        throw Object.assign(new Error(rawMsg.startsWith('redo_document') ? rawMsg : `redo_document: operation failed, state restored: ${rawMsg}`), {
          code: preserved ?? 'E_IO',
          remedy: honestRestoreRemedy('redo_document', memoryOk, diskOk, baseRemedy, rawMsg),
        });
      }
    });
  } catch (err: unknown) {
    const code = (err as { code?: unknown } | null)?.code;
    if (code === 'E_SESSION' || /Unknown session|was closed concurrently|expired/.test(err instanceof Error ? err.message : String(err))) {
      return unknownSessionError(parsed.data.sessionId);
    }
    if (code === 'E_NOT_FOUND' && (err as { atLatestState?: unknown }).atLatestState === true) {
      const base = mcpErrorFrom(
        err,
        'E_NOT_FOUND',
        'The redo stack is empty (a fresh execute_batch edit clears redo): call revert_document first to create a redoable state, then retry redo_document.'
      );
      return { ...base, atLatestState: true as const };
    }
    return mcpErrorFrom(err, 'E_IO', 'Retry redo_document; re-open the document if the session expired.');
  }
}

/**
 * Read-only WOULD-restore preview for revert_document. Resolves the same
 * target (steps XOR commitOid, same E_ codes), reads the target package via
 * the ref-safe checkoutCommit, and diffs package bytes — but performs zero
 * writes: no revision bump, no markSaved, no redo-stack touch, no disk write.
 * Dirty sessions are reported (dirty + dirtyHandling), never discarded here.
 */
export async function handlePreviewRevert(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      targetOid: string;
      targetMessage: string;
      targetAuthor: string;
      stepsBack: number;
      affectedFiles: string[];
      /** Advisory old→new value summaries (live current -> target restore text; bounded). */
      changedValues: ChangedValue[];
      /** True when the count or per-string cap truncated changedValues. */
      changedValuesTruncated: boolean;
      /** True when either side held orphan ids (no shared id across the boundary): paired entries are positional, not identity. */
      changedValuesPositional: boolean;
      /** Present when the target range is not computable (empty changedValues). */
      changedValuesReason?: string;
      dirty: boolean;
      dirtyHandling: string;
      currentHeadOid: string;
      totalCommits: number;
    }
  | McpErrorShape
> {
  const parsed = PreviewRevertSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`preview_revert: invalid input. ${detail}`);
  }
  const rawObj = (typeof input === 'object' && input !== null ? input : {}) as Record<string, unknown>;
  const stepsProvided = 'steps' in rawObj && rawObj['steps'] !== undefined;
  const oidProvided = parsed.data.commitOid !== undefined;
  if (stepsProvided && oidProvided) {
    return invalidInput(
      'preview_revert: steps and commitOid are mutually exclusive. Pass {steps:N} to walk back OR {commitOid:"<40-hex>"} to jump, not both.'
    );
  }
  if (oidProvided && !COMMIT_OID_RE.test(parsed.data.commitOid as string)) {
    return mcpError(
      'E_VALIDATION',
      `preview_revert: commit "${parsed.data.commitOid}" is not a valid 40-hex-char object id`,
      'Call get_document_history to list valid commit oids, then retry preview_revert with one of them.'
    );
  }
  const steps = parsed.data.steps ?? 1;
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      if (!sidecarRepoExists(handle.filePath)) {
        throw Object.assign(new Error('preview_revert: no saved history for this session yet'), {
          code: 'E_NOT_FOUND',
          remedy: 'Call save_document first to record a history baseline, then retry preview_revert.',
        });
      }
      let full: Awaited<ReturnType<typeof getHistory>>;
      try {
        full = await getHistory(handle.filePath);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw Object.assign(new Error(`preview_revert: cannot read history: ${msg}`), {
          code: 'E_IO',
          remedy: 'Retry preview_revert; re-open the document if the session expired.',
        });
      }
      let targetOid: string;
      let stepsBack: number;
      if (oidProvided) {
        const wanted = parsed.data.commitOid as string;
        const at = full.findIndex((c) => c.oid === wanted);
        if (at < 0) {
          throw Object.assign(new Error(`preview_revert: commit "${wanted}" not found in history`), {
            code: 'E_NOT_FOUND',
            remedy: 'Call get_document_history to list valid commit oids, then retry preview_revert with one of them.',
          });
        }
        targetOid = wanted;
        stepsBack = at;
      } else {
        if (steps >= full.length) {
          throw Object.assign(
            new Error(`preview_revert: steps ${steps} exceeds saved history depth ${full.length}`),
            {
              code: 'E_NOT_FOUND',
              remedy: `The session has ${full.length} saved commit(s). Retry with steps < ${full.length} or pick a commitOid from get_document_history.`,
            }
          );
        }
        targetOid = full[steps].oid;
        stepsBack = steps;
      }
      const targetEntry = full.find((c) => c.oid === targetOid);
      // Read-only target read (never moves refs, never touches the worktree).
      let target: Awaited<ReturnType<typeof checkoutCommit>>;
      try {
        target = await checkoutCommit(handle.filePath, targetOid);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw Object.assign(new Error(`preview_revert: cannot read commit "${targetOid}": ${msg}`), {
          code: /not found in history/.test(msg) ? 'E_NOT_FOUND' : 'E_IO',
          remedy: 'Call get_document_history to list valid commit oids, then retry preview_revert with one of them.',
        });
      }
      // Byte-level affected files: live snapshot vs target blobs.
      const affectedFiles: string[] = [];
      let liveSnapForValues: Map<string, Buffer> | null = null;
      try {
        const liveSnap = handle.parser.capturePackageSnapshot();
        liveSnapForValues = liveSnap;
        const liveXml = liveSnap.get('word/document.xml');
        const targetXmlBytes = Buffer.from(target.xml, 'utf8');
        if (!liveXml || Buffer.compare(liveXml, targetXmlBytes) !== 0) affectedFiles.push('word/document.xml');
        const names = new Set<string>();
        for (const k of liveSnap.keys()) if (k !== 'word/document.xml' && isTrackedPackageFile(k)) names.add(k);
        for (const k of target.extra.keys()) names.add(k);
        for (const name of [...names].sort()) {
          if (affectedFiles.length >= AUDIT_FILES_MAX) break;
          const a = liveSnap.get(name);
          const b = target.extra.get(name);
          if (!a && !b) continue;
          if (!a || !b || Buffer.compare(a, b) !== 0) affectedFiles.push(name);
        }
      } catch {
        /* byte diff is best-effort; target identity below still holds */
        if (affectedFiles.length === 0) affectedFiles.push('word/document.xml');
      }
      // Advisory value diff: live current -> target restore text (what a real
      // revert WOULD write). Bounded + best-effort; never fails the preview.
      // Unparsable target packages yield empty + reason (honest, not silent).
      let changedValues: ChangedValue[] = [];
      let changedValuesTruncated = false;
      let changedValuesPositional = false;
      let changedValuesReason: string | undefined;
      try {
        const live = collectAstIdTexts(handle.ast);
        const snap = liveSnapForValues ?? handle.parser.capturePackageSnapshot();
        const targetParsed = await parseTargetIdTextsForPreview(snap, target.xml);
        if (!targetParsed) {
          changedValuesReason = 'Target commit package unreadable; value diff not computable.';
        } else {
          // Across-parse diff (live current -> target restore text): shared
          // ids by id, orphans zipped positionally, nodeId prefers live.
          const diff = diffChangedValuesAcrossParse(live, targetParsed, false);
          changedValues = diff.values;
          changedValuesTruncated = diff.truncated;
          changedValuesPositional = diff.positional;
        }
      } catch {
        changedValues = [];
        changedValuesTruncated = false;
        changedValuesPositional = false;
        changedValuesReason = 'Value diff failed; treated as not computable.';
      }
      const dirty = store.isDirty(handle);
      const unsavedOps = handle.revision - handle.savedRevision;
      return {
        ok: true as const,
        sessionId: handle.sessionId,
        targetOid,
        targetMessage:
          typeof targetEntry?.message === 'string'
            ? targetEntry.message.replace(/\s+$/, '')
            : String(targetEntry?.message ?? ''),
        targetAuthor: typeof targetEntry?.author === 'string' ? targetEntry.author : String(targetEntry?.author ?? ''),
        stepsBack,
        affectedFiles,
        changedValues,
        changedValuesTruncated,
        changedValuesPositional,
        ...(changedValuesReason ? { changedValuesReason } : {}),
        dirty,
        dirtyHandling: dirty
          ? `Session holds ${unsavedOps} revision(s) of unsaved edits: a bare real revert would fail closed with E_CONFLICT; a real revert with discardUnsaved:true WOULD discard them (echoed back bounded). Net-zero edits still count as dirty (revisions pending, not a content diff). This preview discarded nothing.`
          : 'Session is clean: a real revert would discard nothing.',
        currentHeadOid: full[0].oid,
        totalCommits: full.length,
      };
    });
  } catch (err: unknown) {
    const code = (err as { code?: unknown } | null)?.code;
    if (code === 'E_SESSION' || /Unknown session|was closed concurrently|expired/.test(err instanceof Error ? err.message : String(err))) {
      return unknownSessionError(parsed.data.sessionId);
    }
    return mcpErrorFrom(err, 'E_IO', 'Retry preview_revert; re-open the document if the session expired.');
  }
}

/** Git-history dispatch table (registered alongside lean-6 in index.ts). */
export const GIT_HISTORY_HANDLERS = {
  get_document_history: handleGetDocumentHistory,
  preview_revert: handlePreviewRevert,
  revert_document: handleRevertDocument,
  redo_document: handleRedoDocument,
} as const;
