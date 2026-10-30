// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 1476-1776).
// Zero logic edits.
import * as path from 'path';
import {
  assertEngineQuotas,
  type BaseOperation,
  BatchOperationsSchema,
  cellAddressesMatching,
  chartHoleSizeSchemaRemedy,
  chartTypeSchemaRemedy,
  getEngineErrorCode,
  revisionsRemedy,
  validateOperationsStrict,
} from '@attentionform/astra-engine';
import { canonicalizeForMcp, getMcpAllowedRoots, type McpSessionStore } from '../../session-store';
import { assertPathAllowed, resolveMcpStyleGuide } from '../../../tools/compose-document';
import { readEngineNormalized, readParserNormalizations, mergeNormalized } from '../normalize-read';
import {
  mcpErrorFrom,
  sanitizeMcpClientMessage,
  invalidInput,
} from '../errors';
import type { McpErrorShape } from '../errors';
import {
  collectAstIdTexts,
  attributeCreatedIds,
  countPendingRevisionsInParser,
} from '../diff';
import type { BatchOpCreatedIds } from '../diff';
import { SearchContentSchema, ExecuteBatchSchema } from '../schemas';

/**
 * Phase C-3 bounds: regex patterns longer than this are rejected
 * fail-closed (the base query cap is 500; regex gets a tighter bound).
 */
export const SEARCH_REGEX_MAX_PATTERN = 200;
/** Phase C-3: wall-clock budget for one regex evaluation pass over the AST. */
export const SEARCH_REGEX_TIME_BUDGET_MS = 500;
/** Hint rule: only queries that look like a pattern trigger the regex hint.
 * A plain decimal dot is a `.` with an ASCII digit immediately before OR
 * after it (covers `847.2`, `3.14`, `.5`, `5.`): literal number noise, not
 * pattern intent, so it never triggers the hint by itself. Any other
 * regex metachar (`*+?^${}()|[]\`) or a dot with no digit neighbor
 * (as in `foo.bar` or `foo.*` via `*`) does trigger the hint.
 * The `regex:true` path is unaffected. */
const SEARCH_HINT_METACHAR_RE = /[*+?^${}()|[\]\\]/;
function queryLooksLikeRegexPattern(query: string): boolean {
  if (SEARCH_HINT_METACHAR_RE.test(query)) return true;
  for (let i = 0; i < query.length; i++) {
    if (query[i] !== '.') continue;
    const before = i > 0 ? query[i - 1] : '';
    const after = i + 1 < query.length ? query[i + 1] : '';
    const beforeIsDigit = before >= '0' && before <= '9';
    const afterIsDigit = after >= '0' && after <= '9';
    if (!beforeIsDigit && !afterIsDigit) return true;
  }
  return false;
}

/** Generic remedy for regex search failures (no pattern echo, no hit hint). */
export const SEARCH_REGEX_GENERIC_REMEDY =
  'Fix the regular-expression syntax and retry, or omit regex:true for a literal substring search.';
/** For a table hit: the cell addresses (`t_001:r2c3`) that hold the match, usable as edit targets. */
function cellsOf(node: Parameters<typeof cellAddressesMatching>[0], test: (cellText: string) => boolean): { cells?: string[] } {
  if (node.type !== 'table') return {};
  const cells = cellAddressesMatching(node, test);
  return cells.length > 0 ? { cells } : {};
}

export async function handleSearchContent(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      matches: Array<{ id: string; preview: string; cells?: string[] }>;
      results: Array<{ id: string; preview: string; cells?: string[] }>;
      hint?: string;
    }
  | McpErrorShape
> {
  const parsed = SearchContentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`search_content: invalid input. ${detail}`);
  }
  const useRegex = parsed.data.regex ?? false;
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      if (useRegex) {
        // Bounded pattern length (fail-closed before compile).
        if (parsed.data.query.length > SEARCH_REGEX_MAX_PATTERN) {
          throw Object.assign(new Error('search_content: invalid regex query.'), {
            code: 'E_VALIDATION',
            remedy: SEARCH_REGEX_GENERIC_REMEDY,
          });
        }
        let re: RegExp;
        try {
          // Case-insensitive, matching the literal path's semantics.
          re = new RegExp(parsed.data.query, 'i');
        } catch {
          throw Object.assign(new Error('search_content: invalid regex query.'), {
            code: 'E_VALIDATION',
            remedy: SEARCH_REGEX_GENERIC_REMEDY,
          });
        }
        // Execution timeout guard: poll a wall-clock deadline while
        // scanning (ReDoS-prone patterns abort fail-closed, never hang).
        const deadline = Date.now() + SEARCH_REGEX_TIME_BUDGET_MS;
        const matches: Array<{ id: string; preview: string; cells?: string[] }> = [];
        let scanned = 0;
        for (const sec of handle.ast.sections) {
          for (const n of sec.nodes) {
            const text = String(n.metadata.text ?? '');
            let hit = false;
            try {
              hit = re.test(text);
            } catch {
              hit = false;
            }
            if (hit) {
              matches.push({ id: n.id, preview: text.slice(0, 120), ...cellsOf(n, (cell) => Date.now() < deadline && re.test(cell)) });
              if (matches.length >= 50) {
                return { ok: true as const, sessionId: handle.sessionId, matches, results: matches };
              }
            }
            scanned += 1;
            if ((scanned & 31) === 0 && Date.now() > deadline) {
              throw Object.assign(new Error('search_content: regex query could not be completed.'), {
                code: 'E_QUOTA',
                remedy: SEARCH_REGEX_GENERIC_REMEDY,
              });
            }
          }
        }
        return { ok: true as const, sessionId: handle.sessionId, matches, results: matches };
      }
      // Literal path: current behavior, unchanged.
      const hits = handle.parser.searchNodesInAST(handle.ast, parsed.data.query);
      const needle = parsed.data.query.toLowerCase();
      const matches = hits.slice(0, 50).map((n) => ({
        id: n.id,
        preview: String(n.metadata.text ?? '').slice(0, 120),
        ...cellsOf(n, (cell) => cell.toLowerCase().includes(needle)),
      }));
      const out: {
        ok: true;
        sessionId: string;
        matches: Array<{ id: string; preview: string; cells?: string[] }>;
        results: Array<{ id: string; preview: string; cells?: string[] }>;
        hint?: string;
      } = {
        ok: true as const,
        sessionId: handle.sessionId,
        matches,
        results: matches,
      };
      if (queryLooksLikeRegexPattern(parsed.data.query)) {
        out.hint =
          'Query contains regex metacharacters but search ran as a literal substring (default). ' +
          'Pass {"regex":true} to interpret the query as a regular expression ' +
          `(bounded: pattern <=${SEARCH_REGEX_MAX_PATTERN} chars, ~${SEARCH_REGEX_TIME_BUDGET_MS}ms budget, case-insensitive).`;
      }
      return out;
    });
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_SESSION', 'Call open_document first, then retry search_content.');
  }
}

export async function handleExecuteBatch(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      appliedCount: number;
      revision: number;
      skippedIds: string[];
      /** Advisory table warnings (additive, never a failure). */
      warnings: string[];
      /** New node ids created by this batch, in post-batch document order. */
      createdNodeIds: string[];
      /** Per-op attribution (client `id` claims are exact; remainder best-effort). */
      results: BatchOpCreatedIds[];
      /** R-1: pending tracked revisions (w:ins + w:del) remaining after the batch. */
      revisionsPending: number;
      /** R-1: accept remedy, present only when revisionsPending > 0. */
      revisionsRemedy?: string;
      /** R-14: advisory chart stale body refs (same shape [{nodeId, preview}]). Additive, never a failure. */
      possiblyStaleBodyRefs: { nodeId: string; preview: string }[];
      /** Normalizer rewrites reported for this batch (display-only core log). */
      normalized: string[];
      /** True only when this batch dropped a non-empty redo stack (absent otherwise). */
      redoCleared?: boolean;
      /** Names the dropped redo depth (present only when redoCleared). */
      redoClearedWarning?: string;
    }
  | McpErrorShape
> {
  const parsed = ExecuteBatchSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    const remedy = chartTypeSchemaRemedy(detail) ?? chartHoleSizeSchemaRemedy(detail) ?? 'single edit = 1-item batch with canonical shapes';
    return invalidInput(`execute_batch: invalid operation batch. ${detail}. Remedy: ${remedy}.`);
  }
  // Second gate: shared batch schema (caps + optional maxSteps) — same source the
  // engine enforces, surfaced here with remedy before any mutation.
  const batchCheck = BatchOperationsSchema.safeParse({
    operations: parsed.data.operations,
    sequential: parsed.data.sequential,
    rollbackOnError: parsed.data.rollbackOnError,
    maxSteps: parsed.data.maxSteps,
  });
  if (!batchCheck.success) {
    const detail = batchCheck.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`execute_batch: batch rejected. ${detail}`);
  }
  // Third gate: explicit strict per-op validation with remedy.
  try {
    validateOperationsStrict(parsed.data.operations);
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_VALIDATION', 'Fix the listed operations to canonical shapes and retry the batch.');
  }
  // Quota gate (E_QUOTA with remedy, fail-closed before mutation).
  try {
    assertEngineQuotas(parsed.data.operations as unknown as BaseOperation[], parsed.data.maxSteps);
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_QUOTA', 'Split into smaller sequential batches and retry.');
  }
  // sequential defaults to true (document order preserved unless opted out).
  const sequential = parsed.data.sequential !== false;
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      // Jail resolution for image `filePath` reads (explicit arg wins, else
      // the session anchor: dirname of the open document, same trust as the
      // open itself). No global-roots fallback — batch performs zero
      // filesystem reads unless an op carries `filePath`, and must not
      // silently gain them. Consulted only when needed; other batches pass
      // nothing and behave exactly as before.
      const opsList = parsed.data.operations as unknown as Array<Record<string, unknown>>;
      const needsFiles = opsList.some(
        (o) => typeof o?.['filePath'] === 'string' && (o['filePath'] as string).length > 0
      );
      let batchAllowedDir: string | undefined;
      if (needsFiles) {
        const explicit = parsed.data.allowedDir;
        if (explicit !== undefined) {
          batchAllowedDir = assertPathAllowed(explicit, getMcpAllowedRoots());
        } else if (path.isAbsolute(handle.filePath)) {
          batchAllowedDir = canonicalizeForMcp(path.resolve(path.dirname(handle.filePath)));
        } else {
          throw Object.assign(
            new Error(
              'execute_batch: image filePath needs an allowed directory but the session has no filesystem anchor (in-memory document).'
            ),
            {
              code: 'E_SANDBOX',
              remedy: 'pass allowedDir pointing at the folder holding the image files and retry.',
            }
          );
        }
      }
      // Snapshot pre-batch identity BEFORE the engine runs: the engine honors
      // client-supplied op ids verbatim and mints _bk_new_*/type-prefixed ids
      // otherwise, but returns neither — the set difference is the exact
      // created set. Action labels are captured up front because the engine
      // rewrites action in place for image/chart compat ops.
      const beforeIds = new Set<string>(collectAstIdTexts(handle.ast).order);
      const opActions: string[] = (parsed.data.operations as unknown as Array<Record<string, unknown>>).map((o) =>
        String(o?.['action'] ?? '')
      );
      const opShadows: Array<Record<string, unknown>> = (parsed.data.operations as unknown as Array<Record<string, unknown>>).map(
        (o) => ({ ...(o as Record<string, unknown>) })
      );
      // Explicit target footprint for the since-save audit accumulator.
      const opTargets: string[] = [];
      for (const o of opShadows) {
        const t = o['target'];
        if (typeof t === 'string' && t) opTargets.push(t);
        else if (Array.isArray(t)) {
          for (const v of t) if (typeof v === 'string' && v) opTargets.push(v);
        }
      }
      // Per-batch style guide override (that batch only): an explicit
      // styleGuide resolves under the batch jail (explicit allowedDir else
      // the session anchor) and replaces the engine guide for this call;
      // omitted inherits the session guide untouched. Restored in `finally`
      // so the session attachment survives regardless of batch outcome.
      // Unknown names, jail escapes, and remote URLs fail closed before any
      // mutation. Guide parse notes ride BOTH stderr and the client-visible
      // envelope `warnings` (never stderr-only); op-scoped warnings join
      // them below.
      const guideRef = parsed.data.styleGuide;
      const previousGuide = handle.engine.getStyleGuide();
      let guideOverridden = false;
      const guideParseNotes: string[] = [];
      if (guideRef !== undefined) {
        if (guideRef.trim() === '') {
          throw Object.assign(
            new Error(
              'execute_batch: styleGuide is empty: guides cannot be detached — open a new session to work without a guide.'
            ),
            {
              code: 'E_VALIDATION',
              remedy: 'Omit styleGuide to inherit the session guide, or open a new session for guide-free edits.',
            }
          );
        }
        // Explicit allowedDir is validated even when no op carries
        // filePath (the image gate above only jails on demand).
        const explicitGuideJail =
          parsed.data.allowedDir !== undefined
            ? assertPathAllowed(parsed.data.allowedDir, getMcpAllowedRoots())
            : undefined;
        const resolved = resolveMcpStyleGuide(guideRef, {
          ...(explicitGuideJail !== undefined ? { allowedDir: explicitGuideJail } : {}),
          sessionFilePath: handle.filePath,
        });
        for (const w of resolved.warnings) {
          guideParseNotes.push(w);
          try {
            console.error(`[mcp execute_batch] style guide note: ${w}`);
          } catch {
            /* stderr-only, never throws */
          }
        }
        handle.engine.setStyleGuide(resolved.config);
        guideOverridden = true;
      }
      let res: Awaited<ReturnType<typeof handle.engine.executeBatch>>;
      try {
        res = await handle.engine.executeBatch({
          operations: parsed.data.operations as unknown as BaseOperation[],
          sequential,
          rollbackOnError: parsed.data.rollbackOnError,
          maxSteps: parsed.data.maxSteps,
          ...(batchAllowedDir !== undefined ? { allowedDir: batchAllowedDir } : {}),
        });
      } finally {
        if (guideOverridden) handle.engine.setStyleGuide(previousGuide);
      }
      if (!res.success) {
        const first = res.errors[0];
        // P2-2: res.errors[0].error carries the engine stack (see
        // operation-engine.ts routeOperation failure path). Keep the full
        // value stderr-only; the client envelope gets the sanitized message.
        const rawMessage = first ? String(first.error).slice(0, 2000) : 'Batch failed.';
        try {
          console.error(`[mcp execute_batch] E_VALIDATION (client-sanitized): ${rawMessage.slice(0, 500)}`);
        } catch {
          /* stderr-only, never throws */
        }
        const message = sanitizeMcpClientMessage(rawMessage);
        const hint = first && first.hint ? String(first.hint) : '';
        // Prefer the engine's own remedy when the failure text carries one
        // (e.g. jail names on sandbox refusals fall back here when no
        // self-correction hint applies); the generic text stays last.
        // Extracted from the sanitized message so path hygiene matches it.
        // First line only: later lines are rollback/context notes, not remedy.
        const firstLine = message.split('\n')[0];
        const remedyIdx = firstLine.indexOf('Remedy:');
        const engineRemedy =
          remedyIdx >= 0
            ? firstLine.slice(remedyIdx + 'Remedy:'.length).replace(/\s+/g, ' ').trim().replace(/[.]+$/, '')
            : '';
        // Preserve the engine failure code (E_SANDBOX/E_IO/E_LIMIT/...) from
        // the error text instead of collapsing everything to E_VALIDATION.
        const engineCode = getEngineErrorCode(rawMessage) ?? 'E_VALIDATION';
        throw Object.assign(new Error(`execute_batch failed: ${message}`), {
          code: engineCode,
          remedy: hint !== '' ? hint : engineRemedy !== '' ? engineRemedy : 'Re-read ids via inspect_document and retry the batch sequentially.',
        });
      }
      const revision = store.bumpRevision(handle);
      store.refreshByteSize(handle);
      // Display-only: bookkeeping rezipped through the core (healing +
      // logging) — capture its log. Heals that land here are reported on
      // this envelope; a later save then honestly reports only what it
      // heals itself. (A failed bookkeeping rezip leaves the previous log;
      // it cannot fail after a successful batch short of resource
      // exhaustion, in which case the batch itself would have failed.)
      const batchBookkeepingNormalized = readParserNormalizations(handle.parser);
      // A fresh edit invalidates the redo chain (standard undo/redo semantics).
      // The drop is signaled on the envelope (redoCleared + warning) so no
      // caller discovers the cleared stack via a surprise empty redo.
      const redoClearedDepth = handle.redoStack.length;
      handle.redoStack.length = 0;
      // Created-id envelope: exact flat list (post-batch document order) plus
      // per-op attribution, then feed the since-save audit accumulator.
      const afterOrder = collectAstIdTexts(handle.ast).order;
      const createdNodeIds = afterOrder.filter((id) => !beforeIds.has(id));
      const results = attributeCreatedIds(opActions, opShadows, createdNodeIds);
      store.accumulatePendingAudit(handle, res.appliedCount, [...opTargets, ...createdNodeIds]);
      // R-1: carry the engine pending-revisions count through (same remedy
      // wording as the engine; never auto-accepts, never flips defaults).
      const batchRevisionsPending =
        typeof (res as { revisionsPending?: unknown }).revisionsPending === 'number'
          ? (res as { revisionsPending: number }).revisionsPending
          : countPendingRevisionsInParser(handle.parser);
      const batchRevisionsRemedy =
        typeof (res as { revisionsRemedy?: unknown }).revisionsRemedy === 'string'
          ? String((res as { revisionsRemedy: string }).revisionsRemedy)
          : revisionsRemedy(batchRevisionsPending);
      // R-14: carry the engine stale-refs array through (advisory, additive, same shape).
      const batchStaleRefs = Array.isArray((res as { possiblyStaleBodyRefs?: unknown }).possiblyStaleBodyRefs)
        ? ((res as { possiblyStaleBodyRefs: unknown[] }).possiblyStaleBodyRefs.filter(
            (v): v is { nodeId: string; preview: string } =>
              typeof v === 'object' &&
              v !== null &&
              typeof (v as Record<string, unknown>)['nodeId'] === 'string' &&
              typeof (v as Record<string, unknown>)['preview'] === 'string'
          ))
        : [];
      return {
        ok: true as const,
        sessionId: handle.sessionId,
        appliedCount: res.appliedCount,
        revision,
        skippedIds: res.skippedIds,
        warnings: [
          ...guideParseNotes,
          ...(Array.isArray((res as { warnings?: unknown }).warnings)
            ? ((res as { warnings: unknown[] }).warnings.filter((v): v is string => typeof v === 'string'))
            : []),
        ],
        createdNodeIds,
        results,
        revisionsPending: batchRevisionsPending,
        ...(batchRevisionsRemedy ? { revisionsRemedy: batchRevisionsRemedy } : {}),
        possiblyStaleBodyRefs: batchStaleRefs,
        // Display-only: engine result (no field yet) merged with the core
        // log from this batch's bookkeeping rezip.
        normalized: mergeNormalized(readEngineNormalized(res), batchBookkeepingNormalized),
        ...(redoClearedDepth > 0
          ? {
              redoCleared: true as const,
              redoClearedWarning:
                `This batch cleared ${redoClearedDepth} redo entr${redoClearedDepth === 1 ? 'y' : 'ies'}: ` +
                'a fresh edit invalidates redo, so redo_document now reports an empty stack. ' +
                'Call revert_document to record a fresh redoable state.',
            }
          : {}),
      };
    });
  } catch (err: unknown) {
    return mcpErrorFrom(
      err,
      'E_VALIDATION',
      'Re-read node ids via inspect_document({fidelity:"standard"}) and retry sequentially (single edit = 1-item batch).'
    );
  }
}
