// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 1926-2196 + 2203-2273).
// Zero logic edits (pathDirname + restore helpers live in ../diff and are imported).
import * as fsp from 'fs/promises';
import {
  assertWriteTargetNoSymlink,
  chartXmlGetterForAst,
  commitChange,
  getEngineErrorCode,
  lintDocumentStyle,
  macroChoiceError,
  revisionsRemedy,
  writeFileNoFollow,
  type StyleWarning,
} from '@attentionform/astra-engine';
import type { McpSessionStore } from '../../session-store';
import { collectTrackedExtrasForGit, ensureGitRepo, validateSaveTarget } from '../../session-store';
import { readParserNormalizations, mergeNormalized } from '../normalize-read';
import {
  mcpError,
  mcpErrorFrom,
  sanitizeCommitMessage,
  sanitizeCommitAuthor,
  COMMIT_MESSAGE_MAX,
  COMMIT_AUTHOR_MAX,
  COMMIT_AUTHOR_DEFAULT,
  invalidInput,
} from '../errors';
import type { McpErrorShape } from '../errors';
import {
  AUDIT_NODE_IDS_MAX,
  diffChangedValuesAcrossParse,
  collectAstIdTexts,
  parseSavedIdTextsOrdered,
  countPendingRevisionsInParser,
  remedyForValidationError,
  collectTableNodeIdsInOrder,
  formatValidationLocation,
  logRestoreFailure,
  honestRestoreRemedy,
  pathDirname,
} from '../diff';
import type { ChangedValue } from '../diff';
import { SaveDocumentSchema, CloseDocumentSchema } from '../schemas';

/**
 * Validating save — throwing core.
 * M2 contract (consistent, tested):
 * - executeSaveDocument THROWS on every failure (invalid input E_VALIDATION,
 *   unknown/expired session E_SESSION, sandbox escape E_SANDBOX, invalid
 *   package E_VALIDATION, write failure E_IO). All throws carry {code,remedy}.
 * - handleSaveDocument NEVER throws: it wraps this core and maps every throw
 *   to the {ok:false,code,remedy} return contract. Tool transport MUST go
 *   through handleSaveDocument (single source of truth).
 * Never writes corruption: OOXML validation gates the write, and the save
 * target is jailed (M3) before any mkdir/write.
 */
export async function executeSaveDocument(
  store: McpSessionStore,
  input: unknown
): Promise<{
  ok: true;
  sessionId: string;
  path: string;
  revision: number;
  commitOid: string;
  commitMessage: string;
  /** R-1: pending tracked revisions (w:ins + w:del) in the saved package. */
  revisionsPending: number;
  /** R-1: accept remedy, present only when revisionsPending > 0. */
  revisionsRemedy?: string;
  /** Advisory old→new value summaries (saved baseline -> live just saved; bounded, see CHANGED_VALUES_MAX). */
  changedValues: ChangedValue[];
  /** True when the count or per-string cap truncated changedValues. */
  changedValuesTruncated: boolean;
  /** Present when no baseline was computable (new target / unreadable bytes). */
  changedValuesReason?: string;
  /** Normalizer rewrites applied by this save (display-only core log). */
  normalized: string[];
  /**
   * Advisory style findings under the session guide (never blocking).
   * Absent when no guide is attached (guide-less saves stay byte-identical).
   */
  styleWarnings?: StyleWarning[];
}> {
  const parsed = SaveDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    throw Object.assign(new Error(`save_document: invalid input. ${detail}`), {
      code: 'E_VALIDATION',
      remedy: 'Retry save_document with {sessionId} (optional outPath .docx).',
    });
  }
  // Handler-level cap enforcement (wire schema is unbounded on purpose so
  // over-long input reaches here with a remedy instead of a bare MCP -32602).
  if (typeof parsed.data.message === 'string' && parsed.data.message.length > COMMIT_MESSAGE_MAX) {
    throw Object.assign(
      new Error(
        `save_document: message exceeds ${COMMIT_MESSAGE_MAX} chars (got ${parsed.data.message.length})`
      ),
      {
        code: 'E_VALIDATION',
        remedy: `Shorten message to <=${COMMIT_MESSAGE_MAX} chars (single-line commit subject; newlines become spaces) and retry save_document.`,
      }
    );
  }
  if (typeof parsed.data.author === 'string' && parsed.data.author.length > COMMIT_AUTHOR_MAX) {
    throw Object.assign(
      new Error(`save_document: author exceeds ${COMMIT_AUTHOR_MAX} chars (got ${parsed.data.author.length})`),
      {
        code: 'E_VALIDATION',
        remedy: `Shorten author to <=${COMMIT_AUTHOR_MAX} chars (blank falls back to ${COMMIT_AUTHOR_DEFAULT}) and retry save_document.`,
      }
    );
  }
  if (parsed.data.allowMacros === true && parsed.data.stripMacros === true) throw macroChoiceError();
  const saveMacroOpts =
    parsed.data.allowMacros === true
      ? { allowMacros: true as const }
      : parsed.data.stripMacros === true
        ? { stripMacros: true as const }
        : undefined;
  return store.withSession(parsed.data.sessionId, async (handle) => {
    // Macro disposition override for this save (refuse-default): strip
    // removes macro parts now, allow preserves them byte-identical. Omitted
    // keeps the session disposition; the gates below still refuse by default.
    if (saveMacroOpts) handle.parser.applyMacroChoice(saveMacroOpts, 'save');
    // M3: jail the save target BEFORE validation/write (covers both the
    // default session path and an explicit outPath, including symlink-parent
    // escapes via nearest-ancestor canonicalization). Throws E_SANDBOX.
    const rawTarget = parsed.data.outPath ?? handle.filePath;
    const target = validateSaveTarget(rawTarget);
    const validation = await handle.parser.validateOOXML();
    if (!validation.ok) {
      const first = validation.errors[0] as unknown as
        | { description?: unknown; message?: unknown; xPath?: unknown; path?: unknown }
        | undefined;
      const desc =
        (typeof first?.description === 'string' && first.description) ||
        (typeof first?.message === 'string' && first.message) ||
        'OOXML validation failed';
      const xp = typeof first?.xPath === 'string' ? first.xPath : (typeof first?.path === 'string' ? first.path : undefined);
      const at = formatValidationLocation(xp, collectTableNodeIdsInOrder(handle.ast));
      const count = Array.isArray(validation.errors) && validation.errors.length > 1 ? ` (${validation.errors.length} errors, showing first)` : '';
      throw Object.assign(new Error(`E_VALIDATION: save blocked: package invalid${at}${count}: ${String(desc).slice(0, 2000)}`), {
        code: 'E_VALIDATION',
        // Phase C-1c: xpath-specific remedy (docPr/@id, tbl children,
        // relationships each name their own fix).
        remedy: remedyForValidationError(xp, desc),
      });
    }
    // Display-only: the gate above rezips through the core (healing +
    // logging) — capture its log before the save rezip below replaces the
    // field, then merge both so every rewrite is reported once.
    const saveGateNormalized = readParserNormalizations(handle.parser);
    const buf = handle.parser.rezip();
    const saveNormalized = mergeNormalized(saveGateNormalized, readParserNormalizations(handle.parser));
    // Commit-failure atomicity: snapshot the on-disk bytes plus the revision
    // markers BEFORE the write lands, so a later commit throw can put
    // everything back and advance nothing.
    const prevRevision = handle.revision;
    const prevSavedRevision = handle.savedRevision;
    const prevByteSize = handle.byteSize;
    const prevFilePath = handle.filePath;
    let prevBytes: Buffer | null = null;
    try {
      prevBytes = await fsp.readFile(target);
    } catch {
      prevBytes = null;
    }
    try {
      // M1: lstat the leaf (refuse symlinks incl. dangling) before mkdir
      // (mkdir -p through a symlink parent would create outside dirs) and
      // again before the O_NOFOLLOW write (race window).
      assertWriteTargetNoSymlink(target);
      await fsp.mkdir(pathDirname(target), { recursive: true });
      assertWriteTargetNoSymlink(target);
      await writeFileNoFollow(target, buf);
    } catch (err: unknown) {
      // M1: symlink refusals (E_SANDBOX) pass through verbatim — never
      // collapsed to E_IO — so the jail verdict stays loud.
      if (getEngineErrorCode(err) === 'E_SANDBOX' || (err as { code?: string })?.code === 'E_SANDBOX') throw err;
      const msg = err instanceof Error ? err.message : String(err);
      // A failed write can leave a torn target behind: best-effort restore
      // of the prior bytes keeps the commit path invariant (no partial file).
      let diskOk = true;
      try {
        if (prevBytes !== null) {
          await writeFileNoFollow(target, prevBytes);
        } else {
          await fsp.unlink(target);
        }
      } catch (restoreErr: unknown) {
        diskOk = false;
        logRestoreFailure('save_document', target, 'disk', restoreErr);
      }
      const remedy = diskOk
        ? 'Check the outPath directory is writable and retry save_document.'
        : honestRestoreRemedy('save_document', true, false, '', `write failed for "${target}": ${msg}`);
      throw Object.assign(new Error(`save_document: write failed for "${target}": ${msg}`), {
        code: 'E_IO',
        remedy,
      });
    }
    // History commit AFTER the file lands: the sidecar repo records the
    // exact saved package (document.xml + tracked extras). Omitted messages
    // fall back deterministically from the revision delta (k ops since the
    // last save; a clean first save is the initial baseline).
    const savedXml = handle.parser.readDocumentXml();
    const extras = collectTrackedExtrasForGit(handle.parser);
    const opsSinceSave = handle.revision - handle.savedRevision;
    const fallbackMessage =
      opsSinceSave <= 0
        ? `save_document: rev ${handle.revision} (initial baseline)`
        : `save_document: rev ${handle.revision} (${opsSinceSave} ops since last save)`;
    const commitMessage = sanitizeCommitMessage(parsed.data.message, fallbackMessage);
    const authorName = sanitizeCommitAuthor(parsed.data.author);
    let commitOid: string;
    try {
      await ensureGitRepo(target, savedXml, extras);
      commitOid = await commitChange(target, savedXml, commitMessage, extras, undefined, authorName);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      // The write already landed but history did not: restore the prior disk
      // bytes and revision markers so no disk-ahead-of-history state remains.
      let diskOk = true;
      try {
        if (prevBytes !== null) {
          await writeFileNoFollow(target, prevBytes);
        } else {
          await fsp.unlink(target);
        }
      } catch (restoreErr: unknown) {
        diskOk = false;
        logRestoreFailure('save_document', target, 'disk', restoreErr);
      }
      handle.revision = prevRevision;
      handle.savedRevision = prevSavedRevision;
      handle.byteSize = prevByteSize;
      handle.filePath = prevFilePath;
      const okRemedy = 'The .docx file was left unchanged. Retry save_document to record history.';
      // S3c: engine codes pass through verbatim (E_QUOTA from history quotas
      // stays loud); uncoded failures stay E_IO. Sibling pattern:
      // getEngineErrorCode (visual/export/compose). Prefer the engine's own
      // remedy when present so quota guidance survives the envelope.
      const preserved = getEngineErrorCode(err);
      const engineRemedy =
        typeof (err as { remedy?: unknown }).remedy === 'string'
          ? String((err as { remedy?: unknown }).remedy)
          : undefined;
      const baseRemedy = engineRemedy ?? okRemedy;
      throw Object.assign(new Error(`save_document: history commit failed, file restored: ${msg}`), {
        code: preserved ?? 'E_IO',
        remedy: diskOk
          ? baseRemedy
          : honestRestoreRemedy('save_document', true, diskOk, baseRemedy, `history commit failed: ${msg}`),
      });
    }
    try {
      handle.filePath = target;
      handle.byteSize = buf.length;
      // Audit bridge: snapshot the since-save footprint BEFORE markSaved
      // resets the accumulator, then annotate the new head commit.
      const auditNodeIds = [...handle.pendingAudit.nodeIds];
      const auditOps = handle.pendingAudit.ops;
      // Successful save re-syncs the saved marker (clears dirty).
      store.markSaved(handle);
      store.recordCommitMeta(commitOid, {
        author: authorName,
        timestampMs: Date.now(),
        opCount: Math.max(auditOps, opsSinceSave),
        affectedNodeIds: auditNodeIds.slice(0, AUDIT_NODE_IDS_MAX),
      });
    } catch {
      // Bookkeeping must never fail a validated save.
    }
    // R-1: surface pending tracked revisions on the saved package so
    // consumers stop shipping redline PDFs unknowingly. Advisory only —
    // the save still succeeds, nothing is auto-accepted.
    const saveRevisionsPending = countPendingRevisionsInParser(handle.parser);
    const saveRevisionsRemedy = revisionsRemedy(saveRevisionsPending);
    // Advisory value diff: saved baseline (pre-write bytes) -> live just
    // saved. Bounded + best-effort; never fails the save. prevBytes null
    // (new outPath target) or unparsable bytes yield empty + reason.
    const liveIdTexts = collectAstIdTexts(handle.ast);
    let changedValues: ChangedValue[] = [];
    let changedValuesTruncated = false;
    let changedValuesReason: string | undefined;
    try {
      const savedBefore = await parseSavedIdTextsOrdered(prevBytes);
      if (!savedBefore) {
        changedValuesReason = prevBytes
          ? 'Saved baseline unreadable; value diff not computable.'
          : 'No prior saved bytes at target; value diff has no baseline.';
      } else {
        // Across-parse diff: the reparse mints positional ids, so match
        // shared ids by id and zip orphans positionally (silent on equal).
        const diff = diffChangedValuesAcrossParse(liveIdTexts, savedBefore, true);
        changedValues = diff.values;
        changedValuesTruncated = diff.truncated;
      }
    } catch {
      changedValues = [];
      changedValuesTruncated = false;
      changedValuesReason = 'Value diff failed; treated as not computable.';
    }
    return {
      ok: true as const,
      sessionId: handle.sessionId,
      path: target,
      revision: handle.revision,
      commitOid,
      commitMessage,
      revisionsPending: saveRevisionsPending,
      ...(saveRevisionsRemedy ? { revisionsRemedy: saveRevisionsRemedy } : {}),
      changedValues,
      changedValuesTruncated,
      ...(changedValuesReason ? { changedValuesReason } : {}),
      // Display-only: the core save-path hook healed during the rezip above.
      normalized: saveNormalized,
      // Advisory style pass under the session guide (read-only, best-effort,
      // never blocks the save). Absent without a guide. The stored chart
      // reader mirrors validate/document-manager so persisted twin colors
      // flag the same as the CLI.
      ...(handle.styleGuide !== undefined
        ? {
            styleWarnings: lintDocumentStyle(
              handle.ast,
              handle.styleGuide,
              chartXmlGetterForAst(handle.ast, (part) => {
                try {
                  return handle.parser.getZipAdapter().readTextOrNull(part);
                } catch {
                  return undefined;
                }
              }),
            ),
          }
        : {}),
    };
  });
}

/**
 * M2: return-contract wrapper — NEVER throws.
 * Maps every executeSaveDocument throw to {ok:false,code,remedy} so tool
 * transport sees a single consistent shape. Tested: invalid input, sandbox
 * escape, unknown session, and validation failure all return (never throw).
 */
export async function handleSaveDocument(
  store: McpSessionStore,
  input: unknown
): Promise<{
  ok: true;
  sessionId: string;
  path: string;
  revision: number;
  commitOid: string;
  commitMessage: string;
  /** R-1: pending tracked revisions (w:ins + w:del) in the saved package. */
  revisionsPending: number;
  /** R-1: accept remedy, present only when revisionsPending > 0. */
  revisionsRemedy?: string;
  /** Advisory old→new value summaries (saved baseline -> live just saved; bounded). */
  changedValues: ChangedValue[];
  /** True when the count or per-string cap truncated changedValues. */
  changedValuesTruncated: boolean;
  /** Present when no baseline was computable. */
  changedValuesReason?: string;
  /** Normalizer rewrites applied by this save (display-only core log). */
  normalized: string[];
  /** Advisory style findings under the session guide (absent without a guide). */
  styleWarnings?: StyleWarning[];
} | McpErrorShape> {
  try {
    return await executeSaveDocument(store, input);
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_VALIDATION', 'Fix the flagged content and retry save_document.');
  }
}

export async function handleCloseDocument(
  store: McpSessionStore,
  input: unknown
): Promise<
  | { ok: true; sessionId: string; dirty: boolean; discardedRevisions: number; warning?: string }
  | McpErrorShape
> {
  const parsed = CloseDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`close_document: invalid input. ${detail}`);
  }
  // H2: closeDocument throws E_CONFLICT when a batch holds the session lock.
  // Map it to the return contract (do not let it escape as a throw).
  try {
    const outcome = await store.closeDocument(parsed.data.sessionId);
    if (!outcome.closed) {
      return mcpError(
        'E_SESSION',
        `close_document: unknown session "${parsed.data.sessionId}"`,
        'The session is already closed or expired; no action needed.'
      );
    }
    // A dirty close warns instead of silently discarding: dirty + the
    // discarded revision count ride on the success envelope.
    return {
      ok: true as const,
      sessionId: parsed.data.sessionId,
      dirty: outcome.wasDirty,
      discardedRevisions: outcome.discardedRevisions,
      ...(outcome.wasDirty
        ? {
            warning:
              `close_document discarded ${outcome.discardedRevisions} revision(s) of unsaved edits with this session. ` +
              'Save via save_document before closing to keep edits.',
          }
        : {}),
    };
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_CONFLICT', 'Wait for the in-flight batch to finish, then retry close_document.');
  }
}
