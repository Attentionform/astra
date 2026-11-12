import * as fsp from 'fs/promises';
import * as path from 'path';
import { createHash } from 'crypto';
import {
  DocxParser, OperationEngine, ToonGenerator, readOriginal, getOriginalInfo,
  exportOriginal, replaceFileAtomic, createVerifiedBackup, newRestoreBackupPath, getHistory, checkoutCommit,
  commitChange, isTrackedPackageFile, assertHistoryReadyForWrite,
  captureHistoryState, historyChangedSince, rollbackReplacedFile,
} from '@attentionform/astra-engine';
import type { DocHandle, McpSessionStore } from '../../session-store';
import { collectTrackedExtrasForGit, resolveSessionOriginal, validateSaveTarget } from '../../session-store';
import { ExportOriginalDocumentSchema, RestoreOriginalDocumentSchema } from '../schemas';
import { invalidInput, mcpErrorFrom } from '../errors';
import { validateForBaseline, adoptBaseline } from '../baseline-gate';

const hash = (buf: Buffer) => createHash('sha256').update(buf).digest('hex');
function refuse(code: string, message: string, remedy = 'Inspect the file and its history before retrying. Export the original to a new path for recovery without changing this document.'): never {
  throw Object.assign(new Error(message), { code, remedy });
}
/** Refuse when this session's retained original is not verifiably available now; the remedy names the recovery step. */
async function requireAvailableOriginal(handle: DocHandle): Promise<void> {
  const info = await resolveSessionOriginal(handle);
  if (info.available) return;
  const detail = info.detail ?? 'Close and reopen after checking history.';
  const remedy = info.reason === 'initializing'
    ? 'History initialization for this document has not finished. Call get_document_history again later; the original becomes available once it completes for this same file.'
    : 'The retained original cannot be trusted, so no substitute is returned. Keep the working file and any separate backup, and call get_document_history to see the reason.';
  refuse('E_CONFLICT', `Exact-original recovery is unavailable to this session (${info.reason}). ${detail}`, remedy);
}

/** Read-only export with respect to the live session. Overwriting its own path is deliberately refused. */
export async function handleExportOriginalDocument(store: McpSessionStore, input: unknown) {
  const parsed = ExportOriginalDocumentSchema.safeParse(input);
  if (!parsed.success) return invalidInput(`export_original_document: ${parsed.error.message}`);
  try {
    return await store.withSession(parsed.data.sessionId, async handle => {
      if (!path.isAbsolute(handle.filePath)) refuse('E_VALIDATION', 'This memory session has no retained original.');
      await requireAvailableOriginal(handle);
      const outPath = validateSaveTarget(parsed.data.outPath);
      if (outPath === handle.filePath) refuse('E_CONFLICT', 'Use restore_original_document to replace the active document.');
      const original = await exportOriginal(handle.filePath, outPath, { overwrite: parsed.data.overwrite === true });
      return { ok: true as const, sessionId: handle.sessionId, path: outPath, original };
    });
  } catch (err) {
    return mcpErrorFrom(err, 'E_IO', 'Choose a writable output path inside the allowed directory.');
  }
}

/** Compare the persisted state against HEAD without writing or normalizing the working file. */
async function confirmHistoryCurrent(handle: DocHandle, bytes: Buffer): Promise<{ oid: string; parser: DocxParser }> {
  const history = await getHistory(handle.filePath);
  if (!history[0]) refuse('E_CONFLICT', 'Cannot confirm history is current: there is no recorded head.');
  const head = await checkoutCommit(handle.filePath, history[0].oid);
  if (!head.complete) refuse('E_CONFLICT', 'Cannot confirm history is current: the head is not a complete package snapshot. Save the document first.');
  const parser = new DocxParser(bytes, handle.macroPolicy);
  // getAST is part of the open/save pipeline and can hydrate parser caches. It never writes the file.
  await parser.getAST();
  const snapshot = parser.capturePackageSnapshot();
  if (parser.readDocumentXml() !== head.xml) refuse('E_CONFLICT', 'History differs from the current file. Save or revert before restoring the original.');
  const names = [...snapshot.keys()].filter(isTrackedPackageFile).sort();
  const recorded = [...head.extra.keys()].sort();
  if (JSON.stringify(names) !== JSON.stringify(recorded) || names.some(name => !snapshot.get(name)!.equals(head.extra.get(name)!))) {
    refuse('E_CONFLICT', 'History is behind the file or the file was modified outside ASTra. Save or revert before restoring the original.');
  }
  return { oid: history[0].oid, parser };
}

/** Exact-byte restore. Candidate session is staged separately; normal serialization never overwrites the result. */
export async function handleRestoreOriginalDocument(store: McpSessionStore, input: unknown) {
  const parsed = RestoreOriginalDocumentSchema.safeParse(input);
  if (!parsed.success) return invalidInput(`restore_original_document: ${parsed.error.message}`);
  try {
    return await store.withSession(parsed.data.sessionId, async handle => {
      if (!path.isAbsolute(handle.filePath)) refuse('E_VALIDATION', 'This memory session has no retained original.');
      await requireAvailableOriginal(handle);
      if (handle.revision !== handle.savedRevision && !parsed.data.discard) refuse('E_CONFLICT', 'The session has unsaved edits. Save first, or explicitly pass discard:true.');
      const source = validateSaveTarget(handle.filePath);
      await assertHistoryReadyForWrite(source);
      const original = await getOriginalInfo(source);
      if (!original.available) refuse('E_NOT_FOUND', `The original is unavailable: ${original.reason}.`);
      const bytes = await readOriginal(source);
      const previous = await fsp.readFile(source);
      if (!handle.lastWrittenSha256 || hash(previous) !== handle.lastWrittenSha256) refuse('E_CONFLICT', 'The file changed outside ASTra. Reopen it before restoring.');
      const checked = await confirmHistoryCurrent(handle, previous);
      // Preflight opening policy before changing disk. Candidate never replaces the old session on failure.
      const parser = new DocxParser(bytes, handle.macroPolicy);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);
      if (handle.styleGuide) engine.setStyleGuide(handle.styleGuide);
      const validation = await validateForBaseline(bytes);
      const backup = newRestoreBackupPath(source);
      await createVerifiedBackup(backup, previous);
      const historyBefore = await captureHistoryState(source);
      // Earlier redo entries belong to the lineage before this restore, so they must not be applied on top of it
      // (an ordinary edit clears them the same way). They are put back if the restore fails.
      const oldRedo = Array.from(handle.redoStack);
      const previousSession = {
        parser: handle.parser, ast: handle.ast, engine: handle.engine, toon: handle.toon,
        revision: handle.revision, savedRevision: handle.savedRevision, byteSize: handle.byteSize,
        lastWrittenSha256: handle.lastWrittenSha256, validationBaseline: handle.validationBaseline,
        pendingAudit: { ops: handle.pendingAudit.ops, nodeIds: [...handle.pendingAudit.nodeIds] },
      };
      let replaced = false;
      let commitAttempted = false;
      let commitOid: string | undefined;
      let durabilityUnconfirmed = false;
      try {
        // Keep the pre-restore state redoable. A redo quota never blocks a recovery: when the snapshot does not
        // fit, the restore proceeds without a redo entry.
        handle.redoStack.length = 0;
        try {
          handle.redoStack.push((parsed.data.discard ? checked.parser : handle.parser).capturePackageSnapshot());
        } catch (quota) {
          if ((quota as { code?: string })?.code !== 'E_QUOTA') throw quota;
        }
        // Recheck after staging the candidate and backup, immediately before replacement.
        if (hash(await fsp.readFile(source)) !== hash(previous)) refuse('E_CONFLICT', 'The file changed while restore was being prepared.');
        const write = await replaceFileAtomic(source, bytes);
        replaced = true;
        durabilityUnconfirmed = write.durabilityUnconfirmed === true;
        commitAttempted = true;
        commitOid = await commitChange(source, parser.readDocumentXml(), `Restore original file (sha256 ${hash(bytes)})`, collectTrackedExtrasForGit(parser));
        handle.parser = parser;
        handle.ast = ast;
        handle.engine = engine;
        handle.toon = new ToonGenerator();
        handle.byteSize = bytes.length;
        store.bumpRevision(handle);
        store.markSaved(handle);
        handle.lastWrittenSha256 = hash(bytes);
        adoptBaseline(handle, validation);
      } catch (err) {
        Object.assign(handle, previousSession);
        handle.redoStack.length = 0;
        Array.prototype.push.apply(handle.redoStack, oldRedo);
        const { recovered: diskRecovered, failure: rollbackFailure } = await rollbackReplacedFile(source, previous, { replaced, error: err });
        const rollbackError = rollbackFailure ?? '';
        const historyChanged = commitAttempted && await historyChangedSince(source, historyBefore);
        if (diskRecovered) await fsp.unlink(backup).catch(() => undefined);
        const detail = `${String(err instanceof Error ? err.message : err)}${historyChanged ? ' History may have changed despite the failed commit; inspect history before continuing.' : ''}${diskRecovered ? '' : ` Rollback failed (${rollbackError}); recovery bytes are retained at ${backup}.`}`;
        throw Object.assign(new Error(detail), {
          code: (err as { code?: string })?.code ?? 'E_IO',
          remedy: diskRecovered ? 'Inspect the file and history before retrying.' : `Recover from ${backup} before editing further.`,
        });
      }
      const cleanupFailed = await fsp.unlink(backup).then(() => false, () => true);
      return {
        ok: true as const, sessionId: handle.sessionId, path: source,
        revision: handle.revision, dirty: false as const, commitOid, original,
        ...(durabilityUnconfirmed ? { durabilityUnconfirmed: true } : {}),
        ...(cleanupFailed ? { warning: `Restore succeeded but the temporary recovery backup could not be removed: ${backup}` } : {}),
      };
    });
  } catch (err) {
    return mcpErrorFrom(err, 'E_IO', 'Inspect the document and history; export the original to a new path if in-place restore is unavailable.');
  }
}
