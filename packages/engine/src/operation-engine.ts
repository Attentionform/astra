import { DocumentAST, cloneAST } from './ast';
import { refreshAstInPlace } from './ast-refresh';
import { reconcileSections } from './parser/sections';
import { DocxParser } from './parser';
import { getEngineErrorCode, toError } from './errors';
import { BaseOperation, assertNoCellAddress, parseCellAddress, syncCellStandIns } from './operations/base';
import { ReplaceOperation, executeReplace } from './operations/replace';
import { StyleOperation, executeStyle } from './operations/style';
import { ClearFormatOperation, executeClearFormat } from './operations/clearFormat';
import { DeleteOperation, executeDelete } from './operations/delete';
import { InsertOperation, executeInsert } from './operations/insert';
import { MoveOperation, executeMove } from './operations/move';
import { TableOperation, executeTable } from './operations/table';
import { SectionOperation, executeSection } from './operations/section';
import { PageSetupOperation, executePageSetup } from './operations/pageSetup';
import { NamedStyleOperation, executeNamedStyle } from './operations/namedStyle';
import { FieldOperation, executeField } from './operations/field';
import { EquationOperation, executeEquation } from './operations/equation';
import { ListOperation, executeList } from './operations/list';
import { HeaderFooterOperation, executeHeaderFooter } from './operations/headerFooter';
import { ChartOperation, executeChart, type ChartStaleBodyRef } from './operations/chart';
import { ImageOperation, executeImage } from './operations/image';
import { CalloutOperation, executeCallout } from './operations/callout';
import { CommentOperation, executeComment } from './operations/comment';
import { TrackedChangeOperation, executeTrackedChange } from './operations/trackedChange';
import { FootnoteOperation, EndnoteOperation, executeFootnote, executeEndnote } from './operations/footnote';
import { WatermarkOperation, executeWatermark } from './operations/watermark';
import { PageBorderOperation, executePageBorder } from './operations/pageBorder';
import { CoverPageOperation, executeCoverPage } from './operations/coverPage';
import { HyperlinkOperation, executeHyperlink } from './operations/hyperlink';
import { MetadataOperation, executeMetadata } from './operations/metadata';
import type { StyleGuideConfig } from './style-guide/types';

/**
 * Track-B engine quotas (single source of truth).
 * - maxOps: hard cap per batch (stall/OOM guard, mirrors S1B_MAX_OPERATIONS).
 * - maxBytes: hard cap on the JSON-serialized batch payload (OOM guard).
 * - maxSteps: optional per-request step budget (positive integer, NO ceiling;
 *   omitted = unbounded, loops end by natural completion/abort). Only an
 *   explicitly-passed maxSteps is validated (>= 1 integer, E_QUOTA otherwise).
 *   ENGINE_MAX_STEPS (25) remains as a DOCUMENTED DEFAULT HINT ONLY for
 *   callers that want a suggestion — it enforces nothing. Violations throw
 *   EngineQuotaError ([E_QUOTA] with remedy) — never a silent partial apply,
 *   never null.
 */
export const ENGINE_MAX_OPS = 500;
export const ENGINE_MAX_BYTES = 200 * 1024 * 1024;
/** Documented default hint only (25) — NOT an enforced ceiling. Omit maxSteps for unbounded. */
export const ENGINE_MAX_STEPS = 25;

export class EngineQuotaError extends Error {
  readonly code = 'E_QUOTA' as const;
  readonly remedy: string;
  readonly retryable = false as const;
  constructor(message: string, remedy: string) {
    super(`[E_QUOTA] ${message} Remedy: ${remedy}`);
    this.name = 'EngineQuotaError';
    this.remedy = remedy;
  }
}

export class EngineValidationError extends Error {
  readonly code = 'E_VALIDATION' as const;
  readonly remedy: string;
  readonly retryable = false as const;
  constructor(message: string, remedy: string) {
    super(`[E_VALIDATION] ${message} Remedy: ${remedy}`);
    this.name = 'EngineValidationError';
    this.remedy = remedy;
  }
}

/**
 * Single-envelope throw for structural refusals: the message text stays
 * byte-identical (callers match on it), while `.code`/`.remedy` ride along
 * as machine props so serializers and HTTP mappers stop message-parsing.
 */
function structuralRefusal(message: string, remedy: string): Error {
  return Object.assign(new Error(message), { code: 'E_VALIDATION' as const, remedy, retryable: false as const });
}

/** Measures the JSON wire bytes of a batch payload (throws E_VALIDATION when unserializable). */
export function measureBatchBytes(operations: readonly BaseOperation[]): number {
  let serialized: string;
  try {
    serialized = JSON.stringify(operations) ?? '';
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new EngineValidationError(
      `Batch operations are not JSON-serializable (${msg}).`,
      'pass plain JSON operations only (no functions, circular refs, or BigInt); re-read ids via inspect_document and retry.'
    );
  }
  try {
    return Buffer.byteLength(serialized, 'utf8');
  } catch {
    return serialized.length;
  }
}

/** Enforces maxOps/maxBytes/maxSteps. Returns ok:true or throws E_QUOTA/E_VALIDATION with remedy. */
export function assertEngineQuotas(
  operations: readonly BaseOperation[],
  maxSteps?: number
): { ok: true } {
  if (!Array.isArray(operations) || operations.length === 0) {
    throw new EngineValidationError(
      '"operations" must be a non-empty array.',
      "pass { operations: [{ action: 'replace', target: '<nodeId>', text: '...' }], sequential: true }; single edit = 1-item batch."
    );
  }
  if (operations.length > ENGINE_MAX_OPS) {
    throw new EngineQuotaError(
      `"operations" length ${operations.length} exceeds cap ${ENGINE_MAX_OPS}.`,
      `split into batches of <= ${ENGINE_MAX_OPS} operations (sequential:true, in document order) and retry; single edit = 1-item batch.`
    );
  }
  const bytes = measureBatchBytes(operations);
  if (bytes > ENGINE_MAX_BYTES) {
    throw new EngineQuotaError(
      `Batch payload ${bytes} bytes exceeds cap ${ENGINE_MAX_BYTES} bytes.`,
      'split the batch (fewer ops, smaller text/fileData per op), or attach large media by file path instead of inline base64, and retry.'
    );
  }
  if (maxSteps !== undefined) {
    if (!Number.isInteger(maxSteps) || maxSteps < 1) {
      throw new EngineQuotaError(
        `maxSteps ${String(maxSteps)} must be a positive integer.`,
        'pass maxSteps as a positive integer (>= 1) or omit it for unbounded; loops end by natural completion/abort.'
      );
    }
  }
  return { ok: true };
}

export interface BatchPayload {
  operations: BaseOperation[];
  /**
   * Track-B: sequential defaults to TRUE (document order preserved).
   * Pass sequential:false explicitly to opt into the legacy sort-order path.
   */
  sequential?: boolean;
  rollbackOnError?: boolean;
  /**
   * Jail for batch image `filePath` reads. Wins over the constructor
   * `OperationEngineOptions.allowedDir` when both are set. Omitted (and no
   * constructor jail) + a relative `filePath` fails closed with E_SANDBOX;
   * batches without `filePath` ops never touch the filesystem.
   */
  allowedDir?: string;
  /**
   * Optional per-request step budget (positive integer, no ceiling).
   * Omitted = unbounded (loops end by natural completion/abort).
   * ENGINE_MAX_STEPS (25) is a documented default hint only for callers
   * that want a suggestion — it enforces nothing.
   */
  maxSteps?: number;
  /**
   * Preview mode: validate and apply fully in memory, then rewind
   * everything (package bytes, AST, history) and persist nothing.
   * success/appliedCount describe the rewound attempt; preview lists the
   * affected nodes; dryRunBytes carries the post-batch package so callers
   * can run out-of-band validation (present only when a parser is
   * attached). Failures behave exactly like a normal batch failure
   * (success=false, rolled back, nothing persisted).
   */
  dryRun?: boolean;
  /** Emitted per applied operation plus a terminal 'completed'/'error' event. */
  onProgress?: (processed: number, total: number, status: string) => void;
}

/**
 * One affected node from a dryRun batch. before=null means the node was
 * created by the batch; after=null means it was deleted. before===after
 * means the op addressed the node without changing its text (for example
 * a style-only op) — the node is still reported as affected.
 */
export interface DryRunPreview {
  nodeId: string;
  before: string | null;
  after: string | null;
}

export interface ExecutionResult {
  success: boolean;
  appliedCount: number;
  errors: { operationIndex: number; error: string; hint?: string; remedy?: string }[];
  rollbackTriggered: boolean;
  /** OX-04: ids a delete op skipped (idempotent miss, non-strict) inside this batch. */
  skippedIds: string[];
  /** Advisory table warnings (e.g. low-contrast fills). Additive, never a failure. */
  warnings: string[];
  /** R-1: pending tracked revisions (w:ins + w:del) remaining in the doc after the batch. */
  revisionsPending: number;
  /** R-1: accept remedy, present only when revisionsPending > 0. */
  revisionsRemedy?: string;
  /** R-14: advisory chart stale body refs (same shape [{nodeId, preview}]). Additive, never a failure. */
  possiblyStaleBodyRefs: ChartStaleBodyRef[];
  /**
   * True when this result comes from a dryRun batch. The attempt was fully
   * validated and applied in memory, then rewound — nothing was persisted.
   */
  dryRun?: boolean;
  /** Affected nodes, populated only for a successful dryRun batch. */
  preview?: DryRunPreview[];
  /**
   * Post-batch package bytes captured before the dryRun rewind, for
   * out-of-band validation by the caller. Present only when a parser is
   * attached and the dryRun attempt succeeded.
   */
  dryRunBytes?: Buffer;
}

/**
 * R-1: builds the pending-revisions remedy. Present only when count > 0.
 * Purely advisory — never flips trackChanges defaults, never auto-accepts.
 */
export function revisionsRemedy(count: number): string | undefined {
  if (!Number.isFinite(count) || count <= 0) return undefined;
  const n = Math.floor(count);
  return (
    `Document has ${n} pending tracked revision(s) (w:ins/w:del) — exports/PDFs will show redline markup. ` +
    `Remedy: run { action: 'trackedChange', operation: 'acceptAll' } to accept, ` +
    `or sanitize --accept-all-revisions before shipping a redline-free file.`
  );
}

/** Extracts advisory warnings from a routed op result (table/callout ops). */
function collectRouteWarnings(routeOut: unknown): string[] {
  const raw: unknown = (routeOut as Record<string, unknown> | null | undefined)?.['warnings'];
  if (!Array.isArray(raw)) return [];
  return (raw as unknown[]).filter((v: unknown): v is string => typeof v === 'string' && v.length > 0);
}

/** Extracts advisory chart stale-refs from a routed op result (chart updates only). */
function collectStaleBodyRefs(routeOut: unknown): ChartStaleBodyRef[] {
  const raw: unknown = (routeOut as Record<string, unknown> | null | undefined)?.['possiblyStaleBodyRefs'];
  if (!Array.isArray(raw)) return [];
  return (raw as unknown[]).filter((v: unknown): v is ChartStaleBodyRef => {
    if (typeof v !== 'object' || v === null) return false;
    const r = v as Record<string, unknown>;
    return typeof r['nodeId'] === 'string' && typeof r['preview'] === 'string';
  });
}

/** Snapshots node id -> text for the whole AST (dryRun before/after image). */
function snapshotNodeText(ast: DocumentAST): Map<string, string> {
  const out = new Map<string, string>();
  for (const section of ast.sections ?? []) {
    for (const n of section.nodes ?? []) {
      if (typeof n.id === 'string' && n.id && !out.has(n.id)) {
        out.set(n.id, n.metadata?.text ?? '');
      }
    }
  }
  return out;
}

/**
 * Ids a batch addresses, in document order. Ops declare their nodes via
 * target/anchor/id (top level or under params); the start/end aliases
 * resolve to the first/last live id at capture time. A node addressed by
 * an op counts as affected even when its text ends up unchanged (a
 * style-only op touches no text), so dryRun previews stay honest about
 * which nodes the batch reached.
 */
function collectBatchTouchedIds(ops: BaseOperation[], ast: DocumentAST): string[] {
  const order: string[] = [];
  for (const section of ast.sections ?? []) {
    for (const n of section.nodes ?? []) {
      if (typeof n.id === 'string' && n.id) order.push(n.id);
    }
  }
  const seen = new Set<string>();
  const touch = (raw: unknown): void => {
    if (typeof raw !== 'string' || raw.length === 0) return;
    let resolved = raw;
    if (raw === 'start') resolved = order[0] ?? raw;
    else if (raw === 'end') resolved = order[order.length - 1] ?? raw;
    else {
      // A cell address (t_001:r2c3) touches its table.
      const cell = parseCellAddress(raw);
      if (cell && order.includes(cell.tableId)) resolved = cell.tableId;
    }
    if (!seen.has(resolved)) seen.add(resolved);
  };
  for (const op of ops) {
    const fields = op as unknown as Record<string, unknown>;
    touch(fields['target']);
    touch(fields['anchor']);
    touch(fields['id']);
    const params: unknown = fields['params'];
    if (typeof params === 'object' && params !== null) {
      const rec = params as Record<string, unknown>;
      touch(rec['target']);
      touch(rec['anchor']);
      touch(rec['id']);
    }
  }
  return [...seen];
}

/**
 * Builds the dryRun preview: created nodes (before=null), deleted nodes
 * (after=null), text-changed nodes, plus touched-but-unchanged nodes with
 * before===after. Ordering follows the post-batch document, with deletions
 * appended in pre-batch order.
 */
function buildDryRunPreview(
  before: Map<string, string>,
  touched: string[],
  ast: DocumentAST
): DryRunPreview[] {
  const after = snapshotNodeText(ast);
  const touchedSet = new Set(touched);
  const out: DryRunPreview[] = [];
  for (const [id, afterText] of after) {
    if (!before.has(id)) {
      out.push({ nodeId: id, before: null, after: afterText });
    } else {
      const beforeText = before.get(id) as string;
      if (beforeText !== afterText || touchedSet.has(id)) {
        out.push({ nodeId: id, before: beforeText, after: afterText });
      }
    }
  }
  for (const [id, beforeText] of before) {
    if (!after.has(id)) out.push({ nodeId: id, before: beforeText, after: null });
  }
  return out;
}

/**
 * Construction-time options for the operation engine.
 */
export interface OperationEngineOptions {
  /**
   * Jail for batch image `filePath` reads (same trust as the compose
   * `allowedDir`). Read by `executeImage`; a per-batch `BatchPayload`
   * `allowedDir` wins over this when both are set. No silent grants: a
   * `filePath` with no jail configured fails closed with E_SANDBOX.
   */
  allowedDir?: string;
  /**
   * Active style guide ("aesthetic skin"). Applies as omitted-param
   * defaults inside table/chart/callout/code-block ops under a global
   * precedence invariant: explicit op param > guide value > built-in
   * default. A caller-passed value always wins (even when it clashes with
   * the guide — clashes surface as advisory warnings, never silent
   * overrides); guide values fill only omitted params; chart palettes fill
   * only uncolored series in order without shifting explicit indices.
   * Omitted (no guide) = legacy built-in behavior, byte-identical.
   * Replace via `setStyleGuide` (session guides persist; per-batch
   * overrides belong to the caller layer).
   */
  styleGuide?: StyleGuideConfig;
}

export class OperationEngine {
  private ast: DocumentAST;
  private parser?: DocxParser;
  private engineAllowedDir?: string;
  private activeGuide?: StyleGuideConfig;

  // Sorting order for operations to prevent conflicts during batch execution
  private sortOrder: Record<string, number> = {
    // E2: metadata touches only docProps/core.xml (never body XML), so it
    // runs first and cannot conflict with content ops.
    'metadata': 0,
    'style': 1,
    'clearFormat': 1,
    'namedStyle': 2,
    'pageSetup': 3,
    'headerFooter': 4,
    'replace': 5,
    'insert': 6,
    'chart': 6,
    'delete': 7,
    'move': 8,
    'section': 9,
    'table': 10,
    'comment': 11,
    'trackedChange': 12
  };

  private history: BaseOperation[] = [];

  constructor(ast: DocumentAST, parser?: DocxParser, options?: OperationEngineOptions) {
    this.ast = ast;
    this.parser = parser;
    this.engineAllowedDir = options?.allowedDir;
    this.activeGuide = options?.styleGuide;
  }

  /** Active style guide (undefined = legacy built-ins, no defaulting). */
  public getStyleGuide(): StyleGuideConfig | undefined {
    return this.activeGuide;
  }

  /** Attaches (or clears, when omitted) the active style guide. */
  public setStyleGuide(guide?: StyleGuideConfig): void {
    this.activeGuide = guide;
  }

  public getHistory(): BaseOperation[] {
    return this.history;
  }

  public async execute(op: BaseOperation): Promise<{ warnings: string[] }> {
    this.validateOperation(op);
    // Anchor 'start'/'end' resolves against the live AST at execution time.
    this.resolveAnchorAlias(op);
    // R-4: single-op live collision gate. A client `id` becomes the node id
    // verbatim on creator ops; accepting an id already live would mint two
    // nodes sharing one id (rerun doubling, second reports no created ids).
    // Q1 holds ids session-scoped, so the live AST is the authority.
    this.assertClientIdNotLive(op);
    // Track-B: single ops share the batch byte quota (fail-closed, no OOM).
    measureBatchBytes([op]);
    const snapshot = cloneAST(this.ast);
    // OP-01: mirror the batch path — capture the full package so a failed
    // single op rewinds media/rels/styles, not just the AST.
    const pkgSnap = this.parser?.capturePackageSnapshot();
    const headerFooterBefore = this.parser?.headerFooterReferenceIds();
    // OX-06: failed op must not stay pushed — capture the base so the catch
    // path can truncate even if a future edit pushes before routing.
    const historyBase = this.history.length;
    try {
      const routeOut: unknown = await this.routeOperation(op, this.engineAllowedDir);
      const warnings = collectRouteWarnings(routeOut);
      const sectionNote = this.reconcileSections();
      if (sectionNote) warnings.push(sectionNote);
      this.pruneOrphanedHeaderFooters(headerFooterBefore);
      this.validateASTIntegrity();
      this.history.push(op);
      return { warnings };
    } catch (err: unknown) {
      this.history.splice(historyBase);
      await this.restoreAfterFailedBatch(snapshot, pkgSnap);
      throw err;
    }
  }

  // 4.3 Batch Execution (Track-B: sequential defaults true; quotas throw E_QUOTA)
  public async executeBatch(batch: BatchPayload): Promise<ExecutionResult> {
    const result: ExecutionResult = {
      success: true,
      appliedCount: 0,
      errors: [],
      rollbackTriggered: false,
      skippedIds: [],
      warnings: [],
      revisionsPending: 0,
      possiblyStaleBodyRefs: [],
      dryRun: batch.dryRun === true,
    };

    // Track-B quotas first (fail-closed before any mutation or snapshot work).
    assertEngineQuotas(batch.operations, batch.maxSteps);
    // Track-B: sequential defaults to TRUE — document order is preserved
    // unless the caller explicitly opts out with sequential:false.
    const sequential = batch.sequential !== false;

    // R-4: duplicate client ids rejected up-front (no mutation before gate).
    // A client-supplied `id` becomes the new node id verbatim on creator
    // ops (insert / equation-insert / chart-insert / image-insert). Accepting
    // it twice mints two live nodes sharing one id; the rerun then reports
    // createdNodeIds:[] (id already in the pre-batch set) with no warning.
    // Q1 holds ids session-scoped, so the live AST is the authority. Second
    // use fails with E_VALIDATION naming the id plus a remedy. Covers both
    // duplicate `id` within one batch and collision against live-session ids.
    // `params.id` is an accepted alias (equation/callout pick(), chart
    // resolveChartOp, plus engine dispatch normalization onto `op.id` so
    // insert/image/chart-insert honor it), so it is checked identically to
    // top-level `id`. Within-batch check stays for
    // every action (prior contract); live check applies to id-minting ops.
    const liveIdSet = new Set<string>();
    for (const section of this.ast.sections ?? []) {
      for (const n of section.nodes ?? []) {
        if (typeof n.id === 'string' && n.id) liveIdSet.add(n.id);
      }
    }
    const pendingCustomIds = new Set<string>();
    for (let i = 0; i < batch.operations.length; i++) {
      const customId = OperationEngine.requestedClientId(batch.operations[i]);
      if (!customId) continue;
      if (pendingCustomIds.has(customId)) {
        result.success = false;
        result.errors.push({
          operationIndex: i,
          error: `[E_VALIDATION] Duplicate operation id '${customId}' within the same batch.`,
          hint: `Operation ids must be unique inside one batch. Omit 'id' to auto-generate unique _bk_new_* ids, or supply a distinct id per operation. Re-read ids via inspect_document and retry with a fresh id for '${customId}'.`
        });
        // R-1: surface pre-batch pending revisions even on this early return
        // (no mutation happened, so this is the live remaining count).
        result.revisionsPending = this.countPendingRevisions();
        const earlyRemedy = revisionsRemedy(result.revisionsPending);
        if (earlyRemedy) result.revisionsRemedy = earlyRemedy;
        return result;
      }
      if (OperationEngine.mintsNodeId(batch.operations[i]) && liveIdSet.has(customId)) {
        result.success = false;
        result.errors.push({
          operationIndex: i,
          error: `[E_VALIDATION] Operation id '${customId}' collides with a live node id in this session.`,
          hint: `Ids are session-scoped and must be unique among live nodes. Omit 'id' to auto-generate a unique _bk_new_* id, or re-read ids via inspect_document and retry with a fresh id distinct from '${customId}'.`
        });
        result.revisionsPending = this.countPendingRevisions();
        const liveRemedy = revisionsRemedy(result.revisionsPending);
        if (liveRemedy) result.revisionsRemedy = liveRemedy;
        return result;
      }
      pendingCustomIds.add(customId);
      liveIdSet.add(customId);
    }

    // Sort operations safely (stage 4.3) unless sequential (the default).
    // NOTE: `?? 99` (not `|| 99`) — rank 0 (metadata) is valid and must be
    // preserved; `||` would coerce 0 to 99 and run metadata last.
    const sortedOps = sequential ? [...batch.operations] : [...batch.operations].sort((a, b) => {
      const actA = this.resolveAction(a);
      const actB = this.resolveAction(b);
      const orderA = this.sortOrder[actA] ?? 99;
      const orderB = this.sortOrder[actB] ?? 99;
      return orderA - orderB;
    });

    // 4.4 Atomic Rollback: Deep clone the AST and its DOM nodes
    const snapshot = cloneAST(this.ast);
    // Transactional safety net: clone every zip entry so a failed batch can be
    // rewound byte-for-byte, not just at the AST level.
    const pkgSnap = this.parser?.capturePackageSnapshot();
    const headerFooterBefore = this.parser?.headerFooterReferenceIds();
    // OX-06: rewind history[] on rollback — ops pushed before the failure
    // must not survive the rewind (they describe DOM state that no longer
    // exists). Manager snapshot/redo stacks need no engine-side pop: the
    // engine never pushes them (only document-manager.commitTransaction does,
    // and the IPC layer in electron/main.ts calls it strictly AFTER a
    // successful batch). Restoring the package keeps any pre-batch manager
    // snapshot valid against the live DOM by construction.
    const historyBase = this.history.length;
    // Dry-run before-image: captured after the quota/id gates (no mutation
    // happened yet) so the preview diffs exactly what this batch would do.
    const dryRun = batch.dryRun === true;
    const dryRunBefore = dryRun ? snapshotNodeText(this.ast) : null;
    const dryRunTouched = dryRun ? collectBatchTouchedIds(sortedOps, this.ast) : null;

    try {
      for (let i = 0; i < sortedOps.length; i++) {
        const op = sortedOps[i];
        try {
          // 4.1 Layer 8 Validation
          this.validateOperation(op);
          // Anchor 'start'/'end' resolves here — execution time, after prior
          // ops in this batch applied (order-sensitive by design).
          this.resolveAnchorAlias(op);

          // Route operation
          const routeOut: unknown = await this.routeOperation(op, batch.allowedDir ?? this.engineAllowedDir);
          for (const w of collectRouteWarnings(routeOut)) {
            if (!result.warnings.includes(w)) result.warnings.push(w);
          }
          // OX-04: surface delete skips at the batch level too (deduped —
          // same ghost twice reports once, across ops as well as within).
          const skippedRaw: unknown = (routeOut as Record<string, unknown> | null | undefined)?.['skipped'];
          const skipped: string[] = Array.isArray(skippedRaw)
            ? (skippedRaw as unknown[]).filter((v: unknown): v is string => typeof v === 'string')
            : [];
          if (skipped.length > 0) {
            for (const id of skipped) {
              if (!result.skippedIds.includes(id)) result.skippedIds.push(id);
            }
          }
          // R-14: carry chart stale-refs (advisory, additive, deduped by nodeId).
          for (const r of collectStaleBodyRefs(routeOut)) {
            if (!result.possiblyStaleBodyRefs.some((e) => e.nodeId === r.nodeId)) result.possiblyStaleBodyRefs.push(r);
          }
          // D2: appliedCount counts only applied work. An all-miss delete
          // (deleted 0, skipped >0) is idempotent success with zero applies.
          const deletedRaw: unknown = (routeOut as Record<string, unknown> | null | undefined)?.['deleted'];
          const deleted: unknown[] | null = Array.isArray(deletedRaw) ? (deletedRaw as unknown[]) : null;
          const isAllMissDelete =
            deleted !== null && deleted.length === 0 && skipped.length > 0;
          if (!isAllMissDelete) {
            result.appliedCount++;
          }
          this.history.push(op);
          batch.onProgress?.(result.appliedCount, sortedOps.length, 'processing');
        } catch (err: unknown) {
          const msg = err instanceof Error ? err.message : String(err);
          const stack = err instanceof Error && err.stack ? err.stack : msg;
          result.success = false;
          // 4.2 Self-Correction Hint Generation
          const hint = this.generateHint(op, msg);
          // Authored remedies ride on the op error (e.g. chart refusals); keep
          // them on the recorded failure so consumers do not have to parse the
          // rollback envelope or fall back to a generic hint.
          const rawRemedy = (err as { remedy?: unknown } | null | undefined)?.remedy;
          const propRemedy =
            typeof rawRemedy === 'string' && rawRemedy.trim() !== '' ? rawRemedy : undefined;
          result.errors.push({
            operationIndex: i,
            error: stack,
            hint,
            ...(propRemedy !== undefined ? { remedy: propRemedy } : {}),
          });

          if (sequential || batch.rollbackOnError === true) {
            // Single envelope: the thrown error keeps the op failure's
            // machine code and remedy (thrown batch errors used to drop
            // both, forcing message-parsing for HTTP mapping). The batch's
            // self-correction hint rides as the remedy fallback when the op
            // error carried none of its own.
            throw toError(
              getEngineErrorCode(err) ?? 'E_VALIDATION',
              `Batch operation ${i} failed: ${msg}`,
              propRemedy ?? hint ?? 're-inspect ids via inspect_document and retry with a valid node id'
            );
          }
        }
      }

      // Post-Batch structural validation
      // B-1 validate-on-batch gate: when a batch touches table/chart/drawing
      // content, run a cheap O(N) structural check BEFORE commit so
      // success=true + invalid bytes becomes impossible (loud error here
      // flows into the catch below: rollback + success=false). Full XSD
      // (@attentionform/ooxml-validator via worker, ~4ms/doc dominating a ~150ms
      // rezip per the stage-B measurement ledger) is deliberately NOT run
      // per batch — it would double batch cost for interactive paths.
      if (this.batchTouchesStructuredContent(sortedOps)) {
        this.assertStructuredContentValid();
      }
      // Operations that add, delete or move a section break paragraph change
      // where the file's sections end: bring the model's sections in line.
      const sectionNote = this.reconcileSections();
      if (sectionNote && !result.warnings.includes(sectionNote)) result.warnings.push(sectionNote);
      this.pruneOrphanedHeaderFooters(headerFooterBefore);
      this.validateASTIntegrity();
      batch.onProgress?.(result.appliedCount, sortedOps.length, 'completed');
      // R-1: remaining w:ins + w:del in the doc (batch-created + pre-existing).
      result.revisionsPending = this.countPendingRevisions();
      const okRemedy = revisionsRemedy(result.revisionsPending);
      if (okRemedy) result.revisionsRemedy = okRemedy;
      if (dryRun) {
        // Preview mode: the attempt above is the validation. Capture the
        // post-batch package for out-of-band checks, publish the preview,
        // then rewind package + AST + history so nothing persists.
        // restoreAfterFailedBatch is reused as the rewind primitive (its
        // contract is "restore pre-batch state", independent of why).
        if (this.parser) {
          try {
            result.dryRunBytes = this.parser.rezip();
          } catch {
            /* preview still served without out-of-band bytes */
          }
        }
        result.preview = buildDryRunPreview(dryRunBefore as Map<string, string>, dryRunTouched as string[], this.ast);
        await this.restoreAfterFailedBatch(snapshot, pkgSnap);
        this.history.splice(historyBase);
        result.revisionsPending = this.countPendingRevisions();
        const dryRemedy = revisionsRemedy(result.revisionsPending);
        if (dryRemedy) result.revisionsRemedy = dryRemedy;
        else delete result.revisionsRemedy;
      }

    } catch (criticalErr: unknown) {
      const criticalMsg = criticalErr instanceof Error ? criticalErr.message : String(criticalErr);
      // Typed ExecutionResult (errors + rollbackTriggered + onProgress 'error')
      // is the channel; CLI remains the logging surface. Full diagnostic text
      // (stack, rollback sentence, legacy-path detail) rides errors[] — same
      // text as the removed console.error/warns, behavior otherwise identical.
      const criticalStack = criticalErr instanceof Error && criticalErr.stack ? criticalErr.stack : criticalMsg;
      // 4.4 Atomic Rollback
      const legacyDetail = await this.restoreAfterFailedBatch(snapshot, pkgSnap);
      // OX-06: truncate ops pushed before the failure so history[] describes
      // the rewound DOM, not the discarded attempt.
      this.history.splice(historyBase);
      result.rollbackTriggered = true;
      result.success = false;
      // D2: skippedIds describe the rewound attempt, not live state — clear
      // so callers cannot mistake rolled-back skips for live outcome.
      result.skippedIds = [];
      result.warnings = [];
      result.possiblyStaleBodyRefs = [];
      // R-1: rewound state — report what remains after rollback, not the attempt.
      result.revisionsPending = this.countPendingRevisions();
      const rbRemedy = revisionsRemedy(result.revisionsPending);
      if (rbRemedy) result.revisionsRemedy = rbRemedy;
      else delete result.revisionsRemedy;
      // Pass the critical error into the result with the full diagnostic text:
      // stack (was message-only), the rollback sentence (was a console.warn),
      // and the no-parser legacy detail when that path was taken.
      const rollbackNote = 'Rollback triggered. AST restored to previous snapshot.';
      const criticalParts = legacyDetail
        ? [criticalStack, rollbackNote, legacyDetail]
        : [criticalStack, rollbackNote];
      result.errors.push({ operationIndex: -1, error: criticalParts.join(' '), hint: "Rollback triggered due to batch failure." });
      batch.onProgress?.(result.appliedCount, sortedOps.length, 'error');
    }

    return result;
  }

  /**
   * R-1: counts pending tracked revisions (w:ins + w:del) in the live doc.
   * Prefers the parser document DOM (complete); falls back to an AST XML
   * scan for parser-less engines. Best-effort, never throws. Does NOT
   * mutate, accept, or change any trackChanges default.
   */
  public countPendingRevisions(): number {
    try {
      const dom = this.parser?.getDocumentDom?.() ?? this.parser?.getDom?.();
      if (dom && typeof (dom as Document).getElementsByTagName === 'function') {
        return (
          (dom as Document).getElementsByTagName('w:ins').length +
          (dom as Document).getElementsByTagName('w:del').length
        );
      }
    } catch {
      /* fall through to AST scan */
    }
    let n = 0;
    try {
      for (const section of this.ast.sections ?? []) {
        for (const node of section.nodes ?? []) {
          const xml = (node as { xml?: unknown }).xml as
            | { getElementsByTagName?: (tag: string) => { length: number } }
            | undefined;
          if (!xml || typeof xml.getElementsByTagName !== 'function') continue;
          n += xml.getElementsByTagName('w:ins').length + xml.getElementsByTagName('w:del').length;
        }
      }
    } catch {
      /* best-effort: return what was counted */
    }
    return n;
  }

  /**
   * Rewinds state after a failed batch. With a parser attached, the whole OOXML
   * package is restored from the pre-batch snapshot and THIS ast object is
   * rebuilt IN PLACE (sections spliced, maps re-keyed) so external references
   * to the AST remain valid. The blocks and sections get back the ids and
   * versions they had before the batch, so ids handed out earlier still
   * resolve (see refreshAstInPlace). Without a parser, falls back to the
   * legacy cloned-AST rollback (documented best-effort).
   * Returns a diagnostic for batch callers to carry in typed errors[]: on the
   * parser path, what could not be kept (null when nothing was lost); on the
   * legacy path, the same text as the removed console.warn.
   */
  private async restoreAfterFailedBatch(snapshot: DocumentAST, pkgSnap?: Map<string, Buffer>): Promise<string | null> {
    if (this.parser && pkgSnap) {
      this.parser.restorePackageSnapshot(pkgSnap);
      // Also re-points the parser's bare-rezip numbering cache at this.ast
      // (D1R2-3): without it, numbering defs minted after a rollback would be
      // invisible to bare rezip().
      const notes = await refreshAstInPlace(this.ast, this.parser, { carryIdentityFrom: snapshot });
      return notes.length > 0 ? notes.join(' ') : null;
    } else {
      // No parser attached: legacy AST-only rollback (best-effort — the live
      // document DOM cannot be rewound without a parser). The typed
      // ExecutionResult (rollbackTriggered + errors) is the channel; the
      // returned string carries the dropped warn text (same text preserved).
      this.rollback(snapshot);
      return 'OperationEngine: no parser attached — using legacy AST-only rollback; the live document DOM cannot be rewound.';
    }
  }

  /**
   * Keeps the model's sections in step with the section breaks in the
   * document after operations ran. Returns a sentence when it had to rebuild
   * them. Needs a parser: its document is what the file will contain.
   */
  private reconcileSections(): string | null {
    if (!this.parser) return null;
    const body = this.parser.getDocumentDom().getElementsByTagName('w:body')[0] ?? null;
    return reconcileSections(this.ast.sections, body);
  }

  /**
   * Removes header and footer parts the operations left without a reference
   * (a deleted section break), so their text does not stay in the saved file.
   * Best effort: when anything looks unsafe nothing is removed.
   */
  private pruneOrphanedHeaderFooters(before: ReadonlySet<string> | undefined): void {
    if (!this.parser || !before || before.size === 0) return;
    this.parser.pruneOrphanedHeaderFooters(this.ast, before);
  }

  private validateASTIntegrity() {
    // We don't import DOMParser globally to avoid circular deps, just use the xml node's ownerDocument
    for (const section of this.ast.sections) {
      for (const node of section.nodes) {
        if (!node.xml) {
          throw structuralRefusal(
            `AST corruption: Node ${node.id} lost its XML reference.`,
            're-read the document via inspect_document and retry the batch'
          );
        }
        // If it's a real DOM element (has cloneNode), ensure it still has a tagName
        if (typeof node.xml.cloneNode === 'function' && !node.xml.tagName) {
          throw structuralRefusal(
            `AST corruption: Node ${node.id} lost its XML reference.`,
            're-read the document via inspect_document and retry the batch'
          );
        }
      }
    }
  }

  /**
   * B-1: true when any op in the batch can reshape table/chart/drawing XML
   * (direct table/chart/image ops, or inserts carrying those payloads —
   * including raw 'xml', which may embed w:tbl/w:drawing verbatim).
   */
  private batchTouchesStructuredContent(ops: BaseOperation[]): boolean {
    return ops.some(op => {
      const action = this.resolveAction(op);
      if (action === 'table' || action === 'chart' || action === 'image') return true;
      if (action === 'insert') {
        const fields = op as unknown as Record<string, unknown>;
        const t = String(fields['type'] ?? fields['insertType'] ?? '');
        if (t === 'table' || t === 'chart' || t === 'image' || t === 'xml') return true;
      }
      return false;
    });
  }

  /**
   * B-1 validate-on-batch gate (cheap O(N) DOM scan, no XSD, no rezip).
   * Throws loudly on the first structural violation so the batch catch block
   * rolls back instead of reporting success=true over invalid bytes:
   * - every w:tbl has >= 1 w:tr, every w:tr has >= 1 w:tc, every w:tc has
   *   >= 1 w:p (direct children only — nested tables must not satisfy the
   *   outer cell), and gridSpan values stay >= 1;
   * - every w:drawing carries wp:inline|wp:anchor with wp:extent + wp:docPr
   *   + a:graphic (a drawing without extent/docPr fails Word validation);
   * - every c:chart reference carries its r:id (unresolvable chart refs fail
   *   validation at save).
   */
  private assertStructuredContentValid(): void {
    const direct = (parent: Element, tag: string): Element[] => {
      const out: Element[] = [];
      const kids = parent.childNodes;
      for (let i = 0; i < kids.length; i++) {
        const k = kids[i] as Element;
        if ((k as unknown as { nodeType?: number }).nodeType === 1 && (k as Element).tagName === tag) out.push(k as Element);
      }
      return out;
    };
    for (const section of this.ast.sections) {
      for (const node of section.nodes) {
        const xml = node.xml as Element | undefined;
        if (!xml || typeof (xml as Element).getElementsByTagName !== 'function') continue;
        const el = xml as Element;
        const tables = el.tagName === 'w:tbl' ? [el] : Array.from(el.getElementsByTagName('w:tbl')) as Element[];
        for (const tbl of tables) {
          const rows = direct(tbl, 'w:tr');
          if (rows.length === 0) {
            throw structuralRefusal(
              `[B-1 structural gate] Table node '${node.id}' has no w:tr rows — refusing to commit an invalid table. Remedy: delete the empty table or add a row, then retry.`,
              'delete the empty table or add a row, then retry'
            );
          }
          for (const tr of rows) {
            const cells = direct(tr, 'w:tc');
            if (cells.length === 0) {
              throw structuralRefusal(
                `[B-1 structural gate] Table node '${node.id}' has a w:tr with no w:tc cells — refusing to commit an invalid table. Remedy: re-read via inspect_document and repair the row, then retry.`,
                're-read via inspect_document and repair the row, then retry'
              );
            }
            for (const tc of cells) {
              if (direct(tc, 'w:p').length === 0) {
                throw structuralRefusal(
                  `[B-1 structural gate] Table node '${node.id}' has a w:tc with no w:p paragraph — refusing to commit an invalid table. Remedy: write at least one paragraph into every cell, then retry.`,
                  'write at least one paragraph into every cell, then retry'
                );
              }
              const tcPr = direct(tc, 'w:tcPr')[0];
              const gs = tcPr?.getElementsByTagName('w:gridSpan')[0] as Element | undefined;
              if (gs) {
                const v = parseInt(gs.getAttribute('w:val') || '1', 10);
                if (!Number.isFinite(v) || v < 1) {
                  throw structuralRefusal(
                    `[B-1 structural gate] Table node '${node.id}' has w:gridSpan w:val='${gs.getAttribute('w:val')}' (must be >= 1) — refusing to commit an invalid table.`,
                    `set w:gridSpan w:val to an integer >= 1 on table node '${node.id}', then retry`
                  );
                }
              }
            }
          }
        }
        const drawings = Array.from(el.getElementsByTagName('w:drawing')) as Element[];
        for (const drawing of drawings) {
          const container = (drawing.getElementsByTagName('wp:inline')[0] || drawing.getElementsByTagName('wp:anchor')[0]) as Element | undefined;
          if (!container) {
            throw structuralRefusal(
              `[B-1 structural gate] Node '${node.id}' has a w:drawing with no wp:inline/wp:anchor wrapper — refusing to commit an invalid drawing.`,
              `wrap the drawing in wp:inline or wp:anchor with wp:extent, wp:docPr, and a:graphic on node '${node.id}', then retry`
            );
          }
          if (container.getElementsByTagName('wp:extent').length === 0 || container.getElementsByTagName('wp:docPr').length === 0) {
            throw structuralRefusal(
              `[B-1 structural gate] Node '${node.id}' has a drawing missing wp:extent or wp:docPr — refusing to commit an invalid drawing.`,
              `add wp:extent and wp:docPr to the drawing wrapper on node '${node.id}', then retry`
            );
          }
          if (container.getElementsByTagName('a:graphic').length === 0) {
            throw structuralRefusal(
              `[B-1 structural gate] Node '${node.id}' has a drawing missing a:graphic — refusing to commit an invalid drawing.`,
              `add a:graphic to the drawing wrapper on node '${node.id}', then retry`
            );
          }
        }
        const charts = Array.from(el.getElementsByTagName('c:chart')) as Element[];
        for (const chart of charts) {
          if (!chart.getAttribute('r:id')) {
            throw structuralRefusal(
              `[B-1 structural gate] Node '${node.id}' has a c:chart with no r:id relationship — refusing to commit an unresolvable chart.`,
              `relink the chart part (r:id) on node '${node.id}' or remove the chart reference, then retry`
            );
          }
        }
      }
    }
  }

  /**
   * R-4 helper: client-supplied id for an op. Top-level `id` is canonical;
   * `params.id` is an accepted alias by equation/callout pick() resolvers and
   * chart resolveChartOp, and the engine promotes it onto `op.id` at dispatch
   * (normalizeClientIdAlias) so insert/image/chart-insert executors — which
   * read `op.id` only — honor it too. The gate therefore checks both
   * identically, and the alias cannot bypass duplicate detection.
   */
  private static requestedClientId(op: BaseOperation): string | undefined {
    const fields = op as unknown as Record<string, unknown>;
    const rawTop: unknown = fields['id'];
    if (typeof rawTop === 'string' && rawTop) return rawTop;
    const params: unknown = fields['params'];
    if (typeof params === 'object' && params !== null) {
      const rawNested: unknown = (params as Record<string, unknown>)['id'];
      if (typeof rawNested === 'string' && rawNested) return rawNested;
    }
    return undefined;
  }

  /**
   * R-4 helper: true when the op mints a live node id from the client id.
   * Only these routes use `op.id` verbatim as the node id; other actions
   * either ignore `id` (footnote/callout/field generate their own) or are
   * rejected at the schema gate, so live collision applies here only.
   */
  private static mintsNodeId(op: BaseOperation): boolean {
    const fields = op as unknown as Record<string, unknown>;
    const action = typeof fields['action'] === 'string' ? String(fields['action']) : '';
    if (action === 'insert') return true;
    if (action === 'equation') {
      const sub: unknown = fields['operation'] ?? (fields['params'] as Record<string, unknown> | undefined)?.['operation'];
      return sub === 'insert';
    }
    if (action === 'chart') {
      const sub: unknown = fields['operation'] ?? (fields['params'] as Record<string, unknown> | undefined)?.['operation'];
      return sub === undefined || sub === 'insert';
    }
    if (action === 'image') {
      const sub: unknown = fields['operation'] ?? (fields['params'] as Record<string, unknown> | undefined)?.['operation'];
      return sub === 'insert';
    }
    return false;
  }

  /**
   * R-4 alias normalization (dispatch-time, single point for both execute()
   * and executeBatch() via routeOperation): when top-level `id` is absent
   * and `params.id` is a non-empty string, promote it onto `op.id` so every
   * downstream reader sees the effective id. Equation/callout pick() and
   * chart resolveChartOp already accept the alias; insert/image/chart-insert
   * executors read `op.id` only, so without this the gate would reject ids
   * the executors then ignore. When both are present the top-level id stays
   * canonical (matches chart `op.id ?? params.id`); equation's own pick()
   * prefers params, so one op carrying two DIFFERENT ids is contradictory
   * input with per-executor precedence — avoided by callers, not merged here.
   */
  private static normalizeClientIdAlias(op: BaseOperation): void {
    const fields = op as unknown as Record<string, unknown>;
    const top: unknown = fields['id'];
    if (typeof top === 'string' && top) return;
    const params: unknown = fields['params'];
    if (typeof params === 'object' && params !== null) {
      const nested: unknown = (params as Record<string, unknown>)['id'];
      if (typeof nested === 'string' && nested) fields['id'] = nested;
    }
  }

  /** R-4 single-op gate: throws E_VALIDATION when the client id is already live. */
  private assertClientIdNotLive(op: BaseOperation): void {
    const customId = OperationEngine.requestedClientId(op);
    if (!customId) return;
    if (!OperationEngine.mintsNodeId(op)) return;
    for (const section of this.ast.sections ?? []) {
      for (const n of section.nodes ?? []) {
        if (n.id === customId) {
          throw new EngineValidationError(
            `Operation id '${customId}' collides with a live node id in this session.`,
            `Ids are session-scoped and must be unique among live nodes. Omit 'id' to auto-generate a unique _bk_new_* id, or re-read ids via inspect_document and retry with a fresh id distinct from '${customId}'.`
          );
        }
      }
    }
  }

  private validateOperation(op: BaseOperation): void {
    if (!op || typeof op !== 'object') {
      throw new EngineValidationError(
        'Operation must be an object.',
        "pass a canonical operation such as { action: 'replace', target: '<nodeId>', text: '...' }."
      );
    }
    if (!op.action || typeof op.action !== 'string') {
      throw new EngineValidationError(
        "Operation missing 'action' field.",
        'add action, e.g. { action: \'replace\', target, text } or { action: \'insert\', type, anchor, position }. Re-read ids via inspect_document.'
      );
    }
  }

  private generateHint(op: BaseOperation, errorMessage: string): string | undefined {
    // Stale node-id hint for every "not found" EXCEPT the word-level match
    // miss ("text ... not found in node ...", operations/style.ts), which
    // carries its own inspect-and-retry remedy that must not be overwritten.
    if (errorMessage.includes("not found") && !errorMessage.includes("not found in node")) {
      return `Check if the target node ID is correct. Node ids are session-scoped and go stale after reopen/edit; re-read ids via inspect_document; locate current text with search_content, or append without an id via anchor:end.`;
    }
    if (errorMessage.includes("Version mismatch")) {
      return `The document has been updated since you last read it. Call inspect_document to get the latest versions.`;
    }
    return undefined;
  }

  private resolveAction(op: BaseOperation): string {
    const fields = op as unknown as Record<string, unknown>;
    let action = op.action;
    if (action === 'execute_operation' || action === 'execute_operations' || action === 'execute_batch' || action === 'operation') {
      const fallback: unknown = fields['type'] ?? fields['opType'] ?? fields['targetType'] ?? fields['op'] ?? 'insert';
      action = String(fallback);
    }
    return action;
  }

  /**
   * Anchor aliases 'start'/'end' for anchor-carrying ops (insert, move, and
   * equation/chart/image/field/callout/footnote inserts).
   *
   * Execution-time, order-sensitive: called per op inside the batch loop
   * AFTER prior ops in the batch applied, so 'end' means the CURRENT last
   * node — e.g. a clearDoc-style delete of the pre-batch tail followed by
   * { anchor:'end' } rebuilds without knowing survivor ids. 'start' is the
   * current first node. `position` applies relative to the resolved node
   * unchanged (end+after = append, start+before = prepend). Exact lowercase
   * match only — real node ids (including 'End'/'START') pass through
   * untouched. An empty document (zero live nodes) fails closed with
   * E_VALIDATION + remedy.
   */
  private static isAnchorAlias(value: unknown): value is 'start' | 'end' {
    return value === 'start' || value === 'end';
  }

  private resolveAnchorAlias(op: BaseOperation): void {
    const fields = op as unknown as Record<string, unknown>;
    const holders: Record<string, unknown>[] = [fields];
    const params: unknown = fields['params'];
    if (typeof params === 'object' && params !== null) {
      holders.push(params as Record<string, unknown>);
    }
    let aliasSeen: string | null = null;
    for (const h of holders) {
      if (OperationEngine.isAnchorAlias(h['anchor'])) {
        aliasSeen = h['anchor'] as string;
      }
    }
    if (aliasSeen === null) return;
    let first: string | undefined;
    let last: string | undefined;
    for (const section of this.ast.sections ?? []) {
      for (const n of section.nodes ?? []) {
        if (typeof n.id !== 'string' || !n.id) continue;
        if (first === undefined) first = n.id;
        last = n.id;
      }
    }
    if (first === undefined || last === undefined) {
      throw new EngineValidationError(
        `Cannot resolve anchor '${aliasSeen}': the document has no live nodes.`,
        `The 'start'/'end' anchors resolve to the first/last live node at execution time; an empty document has neither. Start from a non-empty document (a blank template ships one paragraph) and retry, or re-read ids via inspect_document.`
      );
    }
    for (const h of holders) {
      if (h['anchor'] === 'start') h['anchor'] = first;
      else if (h['anchor'] === 'end') h['anchor'] = last;
    }
  }

  private async routeOperation(op: BaseOperation, allowedDir?: string): Promise<unknown> {
    try {
      return await this.routeOperationCore(op, allowedDir);
    } finally {
      // Edits made through a table-cell address belong to the table node: fold them in.
      await syncCellStandIns(this.ast);
    }
  }

  private async routeOperationCore(op: BaseOperation, allowedDir?: string): Promise<unknown> {
    // R-4: promote the params.id alias before any executor reads op.id.
    OperationEngine.normalizeClientIdAlias(op);
    const action = this.resolveAction(op);
    const fields = op as unknown as Record<string, unknown> & { action: string; operation?: string };
    assertNoCellAddress(action, fields, this.ast);
    switch (action) {
      case 'replace': return executeReplace(op as ReplaceOperation, this.ast);
      case 'style': return executeStyle(op as StyleOperation, this.ast);
      case 'clearFormat': return executeClearFormat(op as ClearFormatOperation, this.ast);
      case 'insert': {
        // M3 backward-compat: insert+type:image/chart rewrites to the
        // dedicated image/chart ops (explicit, documented in
        // operation-schema.ts — never a silent drop). Prefer
        // {action:'image',operation:'insert',...} /
        // {action:'chart',...} for new code. Missing media args throw
        // with remedy from executeImage/executeChart (no silent no-op).
        const insertType: unknown = fields['type'];
        if (insertType === 'image') {
          fields['action'] = 'image';
          fields['operation'] = 'insert';
          // Legacy EMU names are aliases of width and height.
          for (const [legacy, canonical] of [['widthEmu', 'width'], ['heightEmu', 'height']] as const) {
            const legacyValue = fields[legacy];
            if (legacyValue === undefined) continue;
            if (fields[canonical] !== undefined && fields[canonical] !== legacyValue) {
              throw toError(
                'E_VALIDATION',
                `insert type:image received both ${canonical} and ${legacy} with different values`,
                `pass only ${canonical} (EMU)`
              );
            }
            fields[canonical] = legacyValue;
            delete fields[legacy];
          }
          return await executeImage(op as unknown as ImageOperation, this.ast, this.parser, allowedDir);
        }
        if (insertType === 'chart') {
          fields['action'] = 'chart';
          fields['operation'] = 'insert';
          return await executeChart(op as unknown as ChartOperation, this.ast, this.parser, this.activeGuide);
        }
        return await executeInsert(op as InsertOperation, this.ast, this.parser, this.activeGuide);
      }
      case 'delete': return executeDelete(op as DeleteOperation, this.ast, this.parser);
      case 'move': return executeMove(op as MoveOperation, this.ast, this.parser);
      case 'table': return executeTable(op as TableOperation, this.ast, this.activeGuide);
      case 'section': return executeSection(op as SectionOperation, this.ast, this.parser);
      case 'pageSetup': return executePageSetup(op as PageSetupOperation, this.ast, this.parser);
      case 'namedStyle': return executeNamedStyle(op as NamedStyleOperation, this.ast, this.parser);
      case 'field': return executeField(op as FieldOperation, this.ast, this.parser);
      case 'equation': return executeEquation(op as EquationOperation, this.ast, this.parser);
      case 'list': return executeList(op as ListOperation, this.ast);
      case 'headerFooter': return executeHeaderFooter(op as HeaderFooterOperation, this.ast, this.parser);
      case 'chart': return executeChart(op as ChartOperation, this.ast, this.parser, this.activeGuide);
      case 'image': return executeImage(op as ImageOperation, this.ast, this.parser, allowedDir);
      case 'callout': return executeCallout(op as CalloutOperation, this.ast, this.parser, this.activeGuide);
      case 'footnote': return executeFootnote(op as FootnoteOperation, this.ast, this.parser);
      case 'endnote': return executeEndnote(op as EndnoteOperation, this.ast, this.parser);
      case 'watermark': return executeWatermark(op as WatermarkOperation, this.ast, this.parser);
      case 'pageBorder': return executePageBorder(op as PageBorderOperation, this.ast, this.parser);
      case 'coverPage': return executeCoverPage(op as CoverPageOperation, this.ast, this.parser);
      case 'hyperlink':
      case 'bookmark': return executeHyperlink(op as HyperlinkOperation, this.ast, this.parser);
      case 'comment':
        if (!this.parser) throw new Error("Parser is required for comment operations");
        return executeComment(op as CommentOperation, this.ast, this.parser);
      case 'trackedChange':
        if (!this.parser) throw new Error("Parser is required for trackedChange operations");
        return executeTrackedChange(op as TrackedChangeOperation, this.ast, this.parser);
      case 'metadata': return executeMetadata(op as MetadataOperation, this.ast, this.parser);
      default:
        // BX-10: loud unknown-op error — carry the received value, the valid
        // set, and the self-correction hint. The old bare message
        // (`Unsupported operation: ${op.action}`) gave the model nothing to
        // correct with (insert.ts:203 and trackedChange.ts:16 remain with
        // their owners — NOT touched here).
        throw new Error(
          `Unknown operation '${String(fields['action'])}' (resolved '${action}'). ` +
            `Supported: replace, style, clearFormat, insert, delete, move, table, section, pageSetup, ` +
          `namedStyle, field, equation, list, headerFooter, chart, image, callout, ` +
          `footnote, endnote, watermark, pageBorder, coverPage, hyperlink, bookmark, comment, trackedChange, metadata. ` +
          `Re-read the current structure with inspect_document and retry with a supported action.`
        );
    }
  }

  private rollback(snapshot: DocumentAST) {
    // OX-06 symmetry: the headless no-parser path must rewind EVERYTHING the
    // parser-attached path rewinds (sections + all shared maps + metadata +
    // nextChangeId). Restoring sections alone leaked styles/numbering/rels/
    // headers/footers/media mutations back into the live AST via the shared
    // references that cloneAST (BX-02) now deep-copies. Best-effort legacy
    // path — the live document DOM still cannot be rewound without a parser.
    // Silent by design; the typed ExecutionResult is the channel.
    this.ast.sections = snapshot.sections;
    this.ast.styles = snapshot.styles;
    this.ast.numbering = snapshot.numbering;
    this.ast.relationships = snapshot.relationships;
    this.ast.headers = snapshot.headers;
    this.ast.footers = snapshot.footers;
    this.ast.mediaFiles = snapshot.mediaFiles;
    this.ast.metadata = snapshot.metadata;
    this.ast.settings = snapshot.settings;
    this.ast.nextChangeId = snapshot.nextChangeId;
  }
}
