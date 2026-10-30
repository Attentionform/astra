/**
 * Page mapping utilities for OOXML Document AST
 * Headless, pure Node/TS without Electron dependencies.
 */

import { DocumentAST } from './ast';

export interface PageMap {
  totalPages: number;
  nodePages: Record<string, number>;
  /**
   * IN-07: provenance of the map. `renderer` = paginator/DOM-derived
   * (authoritative); `estimate` = word/line heuristic fallback (may screenshot
   * the wrong page for node_ids — callers must label it as an estimate).
   */
  source: 'renderer' | 'estimate';
  /** True when this map is the heuristic fallback (== source === 'estimate'). */
  isEstimate: boolean;
}

/**
 * OP-25: detects explicit page breaks. Break nodes carry empty text, so the
 * metadata-text check alone misses them — inspect the live XML for
 * <w:br w:type="page"/> as well (mirrors toon.ts placeholder detection).
 */
export function nodeHasPageBreak(node: { metadata?: { text?: string }; xml?: Element }): boolean {
  const text = node.metadata?.text || '';
  if (text === '[Page Break]') return true;
  try {
    const xml = node.xml as Element | undefined;
    if (xml && typeof xml.getElementsByTagName === 'function') {
      const brs = xml.getElementsByTagName('w:br');
      for (let i = 0; i < brs.length; i++) {
        if ((brs[i] as Element).getAttribute('w:type') === 'page') return true;
      }
    }
  } catch { /* best-effort */ }
  return false;
}

export function buildPageMap(ast: DocumentAST, rendererNodePages?: Record<string, number>): PageMap {
  // IN-07: prefer the renderer-supplied map (paginator PageFrames / DOM
  // data-page-number) when present and non-empty. It is authoritative.
  if (rendererNodePages && typeof rendererNodePages === 'object') {
    try {
      const entries = Object.entries(rendererNodePages).filter(
        ([id, p]) => typeof id === 'string' && id.length > 0 && typeof p === 'number' && Number.isFinite(p) && p >= 1
      );
      if (entries.length > 0) {
        const nodePages: Record<string, number> = {};
        let totalPages = 1;
        for (const [id, p] of entries) {
          const page = Math.floor(p);
          nodePages[id] = page;
          if (page > totalPages) totalPages = page;
        }
        return { totalPages, nodePages, source: 'renderer', isEstimate: false };
      }
    } catch { /* fall through to heuristic */ }
  }
  const nodePages: Record<string, number> = {};
  let currentPage = 1;
  let estimatedLineCount = 0;
  const LINES_PER_PAGE = 45; // rough estimate for A4 with standard margins

  for (const section of ast.sections) {
    for (const node of section.nodes) {
      // Check for explicit page breaks (OP-25: XML w:br, not just '[Page Break]' text)
      const text = node.metadata?.text || '';
      if (nodeHasPageBreak(node)) {
        currentPage++;
        estimatedLineCount = 0;
        nodePages[node.id] = currentPage;
        continue;
      }

      nodePages[node.id] = currentPage;

      // Estimate lines consumed by this node
      if (node.type === 'table') {
        const rows = node.metadata?.tableDimensions?.rows || 3;
        estimatedLineCount += rows + 2; // header + rows + spacing
      } else if (node.type === 'image') {
        estimatedLineCount += 15; // images typically consume ~1/3 page
      } else {
        const wordCount = text.split(/\s+/).length;
        estimatedLineCount += Math.max(1, Math.ceil(wordCount / 12)); // ~12 words per line
      }

      if (estimatedLineCount >= LINES_PER_PAGE) {
        currentPage++;
        estimatedLineCount = 0;
      }
    }
  }

  // IN-07: heuristic fallback — always labeled as an estimate.
  return { totalPages: currentPage, nodePages, source: 'estimate', isEstimate: true };
}

/**
 * IN-07: renderer node→page map from paginator PageFrames
 * (same derivation as renderer/engine/pagePaginator buildNodePagesFromFrames;
 * duplicated here so backend never imports renderer code).
 */
export function nodePagesFromPageFrames(
  pages: { pageNumber: number; nodes: { id?: string }[] }[]
): { nodePages: Record<string, number>; totalPages: number } {
  const nodePages: Record<string, number> = {};
  let totalPages = 1;
  try {
    if (!Array.isArray(pages) || pages.length === 0) return { nodePages, totalPages };
    for (const frame of pages) {
      const p = typeof frame?.pageNumber === 'number' && Number.isFinite(frame.pageNumber) && frame.pageNumber >= 1
        ? Math.floor(frame.pageNumber)
        : totalPages;
      totalPages = Math.max(totalPages, p);
      const frameNodes: { id?: string }[] = Array.isArray(frame?.nodes) ? frame.nodes : [];
      for (const n of frameNodes) {
        const id = n?.id;
        if (typeof id === 'string' && id.length > 0 && nodePages[id] === undefined) {
          nodePages[id] = p;
        }
      }
    }
    totalPages = Math.max(totalPages, pages.length);
  } catch { /* best-effort */ }
  return { nodePages, totalPages };
}

/**
 * IN-07: resolve the effective map, preferring renderer-embedded nodePages
 * (e.g. `getStructuralSummary().nodePages`) over the heuristic.
 */
export function resolveEffectivePageMap(
  ast: DocumentAST,
  rendererNodePages?: Record<string, number> | null
): PageMap {
  return buildPageMap(ast, rendererNodePages ?? undefined);
}

/**
 * G2-7 stale-agent-map helper: true when the embedded nodePages captured at
 * loop start must be re-captured before this visual_review call. Revisioned
 * parsers (DocxParser.getNodePagesRevision) go stale iff the revision
 * advanced (AST mutated mid-loop); unrevisioned doubles (null) refresh every
 * call so a mid-loop mutation can never leave node_ids resolving stale.
 * Pure + exported for unit tests.
 */
export function isEmbeddedMapStale(
  capturedRev: number | null | undefined,
  currentRev: number | null | undefined
): boolean {
  if (capturedRev === null || capturedRev === undefined) return true;
  if (currentRev === null || currentRev === undefined) return true;
  return capturedRev !== currentRev;
}
