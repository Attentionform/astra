// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 1078-1118 + 1129-1474).
// Zero logic edits.
import {
  DocxParser,
  ensureAstraGitRepo,
  getHistory,
  getTemplatePalette,
  summarizeStyleGuide,
  type StyleGuideSummary,
} from '@attentionform/astra-engine';
import type { McpSessionStore } from '../../session-store';
import {
  mcpErrorFrom,
  invalidInput,
} from '../errors';
import type { McpErrorShape } from '../errors';
import { OpenDocumentSchema, InspectDocumentSchema } from '../schemas';

// ── Style palette passthrough (wires to backend/ooxml/styles.ts) ────────────
export interface StylePalette {
  availableStyles: string[];
  brandColors: string[];
  primaryFont: string;
  headingStyles?: string[];
}

export function extractStylePalette(parser: DocxParser): StylePalette {
  try {
    const rich = getTemplatePalette(parser);
    const styleIds = (rich.availableStyles || []).map((s) => s.id);
    let fallbackIds: string[] = [];
    try {
      fallbackIds = parser.getStyleIds ? parser.getStyleIds() : [];
    } catch {
      fallbackIds = [];
    }
    const allStyles = Array.from(new Set([...styleIds, ...fallbackIds]));
    return {
      availableStyles: allStyles.length > 0 ? allStyles : ['Normal'],
      brandColors: rich.brandColors && rich.brandColors.length > 0 ? rich.brandColors : [],
      primaryFont: rich.primaryFont || 'Calibri',
      headingStyles: rich.headingStyles,
    };
  } catch {
    // Graceful fallback if getTemplatePalette encounters unparseable XML
    let availableStyles: string[] = [];
    try {
      availableStyles = parser.getStyleIds ? parser.getStyleIds() : [];
    } catch {
      availableStyles = ['Normal'];
    }
    return {
      availableStyles: availableStyles.length > 0 ? availableStyles : ['Normal'],
      brandColors: [],
      primaryFont: 'Calibri',
      headingStyles: [],
    };
  }
}

export async function handleOpenDocument(
  store: McpSessionStore,
  input: unknown
): Promise<{ ok: true; sessionId: string; revision: number; summary: unknown; evicted: string[]; pendingNormalizations: string[] } | McpErrorShape> {
  const parsed = OpenDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`open_document: invalid input. ${detail}`);
  }
  if (parsed.data.allowMacros === true && parsed.data.stripMacros === true) {
    return invalidInput(
      'open_document: allowMacros and stripMacros are mutually exclusive. Retry with exactly one of {allowMacros:true} (preserve) or {stripMacros:true} (strip).'
    );
  }
  try {
    const macroOpts =
      parsed.data.allowMacros === true
        ? { allowMacros: true as const }
        : parsed.data.stripMacros === true
          ? { stripMacros: true as const }
          : undefined;
    const handle = await store.openDocument(parsed.data.path, parsed.data.sessionId, macroOpts);
    if (parsed.data.author !== undefined) handle.ast.settings.defaultAuthor = parsed.data.author;
    // R-6: re-run the sidecar tombstone check on the open path. The store's
    // ensureGitRepo short-circuits when .git/ exists, so a post-restart open
    // of a regenerated chain would otherwise stay silent until the next save.
    // ensureAstraGitRepo runs the check internally; advisory only, never throws.
    try {
      await ensureAstraGitRepo(handle.filePath);
    } catch {
      /* warning path must never fail the open */
    }
    let summary: unknown = null;
    try {
      // Additive summaryOnly flag (default false): true returns the style
      // palette summary only (no stylesXml dump); omitted/false keeps the
      // full structural summary byte-identical for existing clients.
      if (parsed.data.summaryOnly === true) {
        summary = { palette: extractStylePalette(handle.parser), summaryOnly: true as const };
      } else {
        summary = handle.parser.getStructuralSummary();
      }
    } catch {
      summary = null;
    }
    // Audit bridge: seed the baseline commit's meta (author + ms timestamp)
    // best-effort so history entries are complete from the first read.
    try {
      const hist = await getHistory(handle.filePath);
      const head = hist[0];
      if (head && !store.getCommitMeta(head.oid)) {
        store.recordCommitMeta(head.oid, {
          author: typeof head.author === 'string' ? head.author : String(head.author ?? ''),
          timestampMs: Date.now(),
          opCount: 0,
          affectedNodeIds: [],
        });
      }
    } catch {
      /* buffer-style or repo-less opens simply have no baseline meta */
    }
    // Phase C-4 LRU signal: ids dropped by capacity eviction to make room.
    const evicted = store.drainLastOpenEvicted();
    // Would-be normalizations: the allowlist healing a save would apply,
    // previewed here (dry-run, nothing written) so migration operators see
    // ordering repairs at open instead of discovering them at save.
    const pendingNormalizations = readPendingNormalizations(handle);
    return { ok: true, sessionId: handle.sessionId, revision: handle.revision, summary, evicted, pendingNormalizations };
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_IO', 'Check the .docx path is readable and retry open_document.');
  }
}

/**
 * Phase C-2: docs rendering under ~2000 tokens return the full TOON
 * (pre-C behavior, zero friction). Larger docs default to the first
 * ~50-node window. Behavior-change note: large-doc default changed from
 * full dump to first window — the navigation remedy teaches paging.
 */
export const INSPECT_FULL_TOON_TOKEN_THRESHOLD = 2000;

export interface InspectWindowEnvelope {
  totalNodes: number;
  returnedNodes: number;
  offset: number;
  limit: number;
  hasMore: boolean;
  nextOffset: number | null;
  truncated: boolean;
}

/**
 * Best-effort dry-run preview of pending allowlist normalizations for an
 * inspect envelope (clone-based, never mutates, never fails the inspect).
 */
export function readPendingNormalizations(handle: { parser?: { getPendingNormalizations?: () => string[] } }): string[] {
  try {
    const list = handle.parser?.getPendingNormalizations?.();
    return Array.isArray(list) ? list.filter((v): v is string => typeof v === 'string') : [];
  } catch {
    return [];
  }
}

export async function handleInspectDocument(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      revision: number;
      /** Phase C-1b: true when in-memory state holds unsaved execute_batch edits. */
      dirty: boolean;
      fidelity: string;
      palette: StylePalette;
      /** E2 (F-4): docProps/core.xml core properties (Dates as ISO strings; {} when absent). */
      metadata: Record<string, unknown>;
      /** Would-be allowlist normalizations a save would apply (dry-run preview, [] when canonical). */
      pendingNormalizations: string[];
      toonText: string;
      rowCount: number;
      totalNodes: number;
      returnedNodes: number;
      offset: number;
      limit: number;
      hasMore: boolean;
      nextOffset: number | null;
      truncated: boolean;
      /** Present when windowed: teaches search -> aroundNode zoom -> offset paging. */
      remedy?: string;
      /** Compact summary of the session style guide (absent when none attached). */
      styleGuide?: StyleGuideSummary;
    }
  | McpErrorShape
> {
  const parsed = InspectDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`inspect_document: invalid input. ${detail}`);
  }
  const { sessionId, sectionId, maxTextLength, aroundNode } = parsed.data;
  const fidelity = parsed.data.fidelity ?? 'standard';
  const offset = parsed.data.offset ?? 0;
  const limit = parsed.data.limit ?? 50;
  const window = parsed.data.window ?? 5;
  const maxTokens = parsed.data.maxTokens;
  // Explicit windowing = a C-2 key the CALLER actually sent. Two layers:
  // (1) the wire schema above carries no .default(), so a real-wire
  // aroundNode-only call arrives without offset/limit/window; (2) a key
  // present with exactly its default value (offset 0, limit 50, window 5)
  // is treated as implicit — that is the exact SDK-parsed shape
  // ({offset:0,limit:50,window:5,...} all present + aroundNode) that the
  // old .default() schema produced over wire and that must still zoom.
  // Only a non-default offset/limit alongside aroundNode is rejected.
  const rawObj = (typeof input === 'object' && input !== null ? input : {}) as Record<string, unknown>;
  const rawOffset = rawObj['offset'];
  const rawLimit = rawObj['limit'];
  const rawWindow = rawObj['window'];
  const offsetExplicit = 'offset' in rawObj && rawOffset !== undefined && rawOffset !== 0;
  const limitExplicit = 'limit' in rawObj && rawLimit !== undefined && rawLimit !== 50;
  const windowExplicit = 'window' in rawObj && rawWindow !== undefined && rawWindow !== 5;
  const windowingRequested =
    offsetExplicit || limitExplicit || 'aroundNode' in rawObj || windowExplicit || 'maxTokens' in rawObj;
  // aroundNode zooms to [node-w, node+w] and sets its own window — paging
  // keys would be silently discarded, so reject the combo fail-closed
  // (explicit non-default offset/limit only; default-valued keys zoom).
  if (aroundNode !== undefined && (offsetExplicit || limitExplicit)) {
    return invalidInput(
      'inspect_document: aroundNode and offset/limit are mutually exclusive. ' +
        'Use {aroundNode:"<id>", window:5} to zoom OR {offset,limit} to page, not both.'
    );
  }
  try {
    return await store.withSession(sessionId, async (handle) => {
      let table;
      try {
        table = handle.toon.toon(handle.ast, { fidelity, sectionId, maxTextLength });
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw Object.assign(new Error(msg), {
          code: 'E_VALIDATION',
          remedy:
            fidelity === 'full' && !sectionId
              ? 'Full fidelity is section-scoped: retry with {fidelity:"full",sectionId:"<sectionId>"} or use fidelity standard.'
              : 'Re-read ids via inspect_document({fidelity:"standard"}) and retry.',
        });
      }
      const palette = extractStylePalette(handle.parser);
      const dirty = store.isDirty(handle);
      // E2 (F-4): surface ast.metadata (core properties) alongside TOON —
      // Dates serialize as ISO strings; absent core part yields {}.
      const metadata: Record<string, unknown> = {};
      try {
        for (const [k, v] of Object.entries(handle.ast.metadata ?? {})) {
          metadata[k] = v instanceof Date ? v.toISOString() : v;
        }
      } catch {
        /* metadata stays {} — inspect must not fail on it */
      }
      const totalNodes = table.rows.length;
      let fullText = '';
      try {
        fullText = handle.toon.renderToonText(table);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw Object.assign(new Error(`inspect_document: TOON render failed: ${msg}`), {
          code: 'E_VALIDATION',
          remedy: 'Retry inspect_document with fidelity standard and a smaller maxTextLength.',
        });
      }

      // Small docs, no explicit windowing: full TOON, unchanged shape +
      // the C-2 envelope fields describing the (single) full window.
      if (!windowingRequested && handle.toon.estimateTokens(fullText) <= INSPECT_FULL_TOON_TOKEN_THRESHOLD) {
        return {
          ok: true as const,
          sessionId: handle.sessionId,
          revision: handle.revision,
          dirty,
          fidelity,
          palette,
          metadata,
          pendingNormalizations: readPendingNormalizations(handle),
          ...(handle.styleGuide !== undefined ? { styleGuide: summarizeStyleGuide(handle.styleGuide) } : {}),
          toonText: fullText,
          rowCount: totalNodes,
          totalNodes,
          returnedNodes: totalNodes,
          offset: 0,
          limit: Math.min(totalNodes, 200),
          hasMore: false,
          nextOffset: null,
          truncated: false,
        };
      }

      // ── Windowed path: offset/limit first, then maxTokens cut ──────────
      let start: number;
      let end: number;
      let effectiveLimit = limit;
      if (aroundNode !== undefined) {
        const idx = table.rows.findIndex((r) => r.id === aroundNode);
        if (idx < 0) {
          throw Object.assign(
            new Error(`inspect_document: aroundNode "${aroundNode}" not found (${totalNodes} nodes in scope).`),
            {
              code: 'E_VALIDATION',
              remedy:
                'Call search_content to find a valid target node id, then retry inspect_document with {aroundNode:"<id>", window:5}.',
            }
          );
        }
        start = Math.max(0, idx - window);
        end = Math.min(totalNodes, idx + window + 1);
        effectiveLimit = end - start;
      } else {
        if (offset >= totalNodes && totalNodes > 0) {
          throw Object.assign(
            new Error(`inspect_document: offset ${offset} is beyond totalNodes ${totalNodes}.`),
            {
              code: 'E_VALIDATION',
              remedy: `Page within range: retry with {offset:<${totalNodes}, limit:${limit}} or follow nextOffset from the previous window.`,
            }
          );
        }
        start = Math.min(offset, totalNodes);
        end = Math.min(totalNodes, start + limit);
      }
      let sliced = table.rows.slice(start, end);

      // maxTokens: shrink on a clean row boundary (whole rows only — the
      // TOON text is line-oriented, so dropping trailing rows never
      // mid-row-truncates). Always reported via truncated/hasMore.
      let truncated = false;
      if (maxTokens !== undefined && sliced.length > 0) {
        const renderRows = (n: number): string =>
          handle.toon.renderToonText({
            headers: table.headers,
            rows: sliced.slice(0, n),
            fidelity: 'lossy',
            tier: table.tier,
          });
        if (handle.toon.estimateTokens(renderRows(sliced.length)) > maxTokens) {
          // Binary-search the largest fitting prefix; keep ≥1 row so the
          // caller always sees the window head + truncation flags.
          let lo = 1;
          let hi = sliced.length - 1;
          let best = 0;
          while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            if (handle.toon.estimateTokens(renderRows(mid)) <= maxTokens) {
              best = mid;
              lo = mid + 1;
            } else {
              hi = mid - 1;
            }
          }
          if (best < 1) best = 1;
          if (best < sliced.length) {
            sliced = sliced.slice(0, best);
            truncated = true;
          }
        }
        // Single-row over-budget: the binary search above cannot run when
        // sliced.length === 1 (hi = 0), so best stays 1 with no signal.
        // Keep the 1 row but flag truncation when it still exceeds budget.
        if (sliced.length >= 1 && handle.toon.estimateTokens(renderRows(sliced.length)) > maxTokens) {
          truncated = true;
        }
      }

      const returnedNodes = sliced.length;
      const hasMore = start + returnedNodes < totalNodes;
      const toonText = handle.toon.renderToonText({
        headers: table.headers,
        rows: sliced,
        fidelity: 'lossy',
        tier: table.tier,
      });
      const out: {
        ok: true;
        sessionId: string;
        revision: number;
        dirty: boolean;
        fidelity: string;
        palette: StylePalette;
        metadata: Record<string, unknown>;
        pendingNormalizations: string[];
        toonText: string;
        rowCount: number;
        totalNodes: number;
        returnedNodes: number;
        offset: number;
        limit: number;
        hasMore: boolean;
        nextOffset: number | null;
        truncated: boolean;
        remedy?: string;
        styleGuide?: StyleGuideSummary;
      } = {
        ok: true as const,
        sessionId: handle.sessionId,
        revision: handle.revision,
        dirty,
        fidelity,
        palette,
        metadata,
        pendingNormalizations: readPendingNormalizations(handle),
        ...(handle.styleGuide !== undefined ? { styleGuide: summarizeStyleGuide(handle.styleGuide) } : {}),
        toonText,
        rowCount: returnedNodes,
        totalNodes,
        returnedNodes,
        offset: start,
        limit: effectiveLimit,
        hasMore,
        nextOffset: hasMore ? start + returnedNodes : null,
        truncated,
      };
      if (hasMore || truncated) {
        out.remedy =
          `Large document: showing ${returnedNodes}/${totalNodes} TOON nodes` +
          `${truncated ? ` (cut to fit maxTokens ${maxTokens})` : ''}. ` +
          'Navigate: 1) search_content to find a target node id, ' +
          '2) inspect_document {aroundNode:"<id>", window:5} to zoom, ' +
          `3) page with {offset:${hasMore ? start + returnedNodes : start}, limit:${effectiveLimit}} to continue.`;
      }
      return out;
    });
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_SESSION', 'Call open_document first, then retry inspect_document with the returned sessionId.');
  }
}
