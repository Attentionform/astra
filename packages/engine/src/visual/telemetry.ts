/**
 * Spatial telemetry: deterministic ASCII wireframe + large-defect lint.
 *
 * Default, always-on visual signal (<500 tokens). Tables and list markers
 * come from the same readers `inspect` uses (the table XML grid, the list
 * numbering helpers), so the wireframe and the outline never disagree.
 */

import type { DocumentAST } from '../ast';
import { extractTableGrid, type ToonGrid } from '../toon';
import { hasListNumbering, isOrderedListItem } from '../export/markdown';

export interface VisualPageMap {
  totalPages: number;
  nodePages: Record<string, number>;
  source: 'renderer' | 'estimate';
  isEstimate: boolean;
}

export interface VisualPageBreakdown {
  page: number;
  nodeCount: number;
  headings: string[];
  hasTables: boolean;
  hasCharts: boolean;
  hasImages: boolean;
}

export interface VisualTelemetry {
  totalPages: number;
  /**
   * Set only when the caller supplied the renderer's actual page count
   * (`renderTotalPages`) and it differs from the word-layout page map:
   * `totalPages` then follows the renderer (single source of truth for the
   * `[PAGE n OF m]` marker) and this preserves the word-layout total so
   * consumers can label both paginations honestly. Absent on agreement or
   * when no render count was supplied (compat: shape unchanged for all
   * existing callers).
   */
  wordLayoutTotalPages?: number;
  warnings: string[];
  pageBreakdown: VisualPageBreakdown[];
  /** ASCII wireframe for the requested page (budget-enforced). */
  wireframe: string;
  /** Estimated token cost of `wireframe + warnings` (chars/4, ceil). */
  tokens: number;
  /** True when nodes were dropped to respect the token budget. */
  truncated: boolean;
  /** Page-map provenance (renderer = authoritative, estimate = heuristic). */
  mapSource: 'renderer' | 'estimate';
}

/** Token budget: spatial telemetry must stay under 500 tokens. */
export const VISUAL_TELEMETRY_TOKEN_BUDGET = 500;
/** Wireframe ASCII width (matches legacy visual-review for familiarity). */
export const VISUAL_WIREFRAME_WIDTH = 68;
/** Wide-table threshold (cols): margin-clipping risk. */
export const VISUAL_WIDE_TABLE_COLS = 6;

export function estimateTokenCount(text: string): number {
  if (!text) return 0;
  return Math.max(0, Math.ceil(text.length / 4));
}

function asRecord(value: unknown): Record<string, number> {
  if (!value || typeof value !== 'object') return {};
  const out: Record<string, number> = {};
  for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
    if (typeof k === 'string' && k.length > 0 && typeof v === 'number' && Number.isFinite(v) && v >= 1) {
      out[k] = Math.floor(v);
    }
  }
  return out;
}

/**
 * Normalize a renderer-reported total page count. Returns `undefined` when
 * it is unusable (non-numeric/non-finite/<1) OR equal to the word-layout
 * map total — a defined value therefore means "genuine mismatch", which is
 * the only case that may relabel the marker or replace the empty-page body.
 */
function normalizeRenderTotal(renderTotal: unknown, mapTotal: number): number | undefined {
  if (typeof renderTotal !== 'number' || !Number.isFinite(renderTotal) || renderTotal < 1) return undefined;
  const total = Math.max(1, Math.floor(renderTotal));
  return total === mapTotal ? undefined : total;
}

/** Singular/plural label for page counts in user-facing warnings/messages. */
function pageCountLabel(count: number): string {
  return count === 1 ? '1 page' : `${count} pages`;
}

function resolveMap(ast: DocumentAST, pageMap?: VisualPageMap | null): VisualPageMap {
  if (pageMap && typeof pageMap.totalPages === 'number' && Number.isFinite(pageMap.totalPages) && pageMap.totalPages >= 1) {
    const nodePages = asRecord(pageMap.nodePages);
    const source = pageMap.source === 'renderer' ? 'renderer' : 'estimate';
    return {
      totalPages: Math.max(1, Math.floor(pageMap.totalPages)),
      nodePages,
      source,
      isEstimate: source === 'estimate',
    };
  }
  // Fallback: single-page estimate (caller should pass the real map from
  // buildPageMap/resolveEffectivePageMap for multi-page docs).
  const nodePages: Record<string, number> = {};
  try {
    for (const section of ast.sections ?? []) {
      for (const node of section.nodes ?? []) {
        if (node && typeof node.id === 'string' && node.id.length > 0 && nodePages[node.id] === undefined) {
          nodePages[node.id] = 1;
        }
      }
    }
  } catch {
    /* best-effort */
  }
  return { totalPages: 1, nodePages, source: 'estimate', isEstimate: true };
}

function nodeText(node: { metadata?: { text?: unknown } }): string {
  const t = node?.metadata?.text;
  return typeof t === 'string' ? t : '';
}

function nodeStyle(node: { metadata?: { namedStyle?: unknown } }): string {
  const s = node?.metadata?.namedStyle;
  return typeof s === 'string' ? s.toLowerCase() : '';
}

function isHeadingNode(node: { type?: unknown; metadata?: { namedStyle?: unknown; text?: unknown } }): boolean {
  const style = nodeStyle(node as { metadata?: { namedStyle?: unknown } });
  if (style.startsWith('heading') || style === 'title') return true;
  if (style.startsWith('h1') || style.startsWith('h2') || style.startsWith('h3')) return true;
  return false;
}

/** The table's real grid, read from its XML (the model's `tableDimensions` is never filled). */
function tableGridOf(node: unknown): ToonGrid | null {
  try {
    const xml = (node as { xml?: { getElementsByTagName?: unknown } } | null)?.xml;
    if (xml && typeof xml.getElementsByTagName === 'function') {
      const grid = extractTableGrid(xml as unknown as Element);
      if (grid.rows > 0 && grid.cols > 0) return grid;
    }
  } catch {
    /* fall back to metadata */
  }
  return null;
}

function tableCols(node: { type?: unknown; metadata?: { tableDimensions?: unknown; headers?: unknown; rows?: unknown } }): number {
  const grid = tableGridOf(node);
  if (grid) return grid.cols;
  try {
    const meta = (node as { metadata?: Record<string, unknown> }).metadata ?? {};
    const dims = meta.tableDimensions as { cols?: unknown } | undefined;
    if (dims && typeof dims.cols === 'number' && Number.isFinite(dims.cols) && dims.cols > 0) {
      return Math.floor(dims.cols);
    }
    const headers = meta.headers;
    if (Array.isArray(headers) && headers.length > 0) return headers.length;
    const rows = meta.rows;
    if (Array.isArray(rows) && rows.length > 0 && Array.isArray(rows[0])) return (rows[0] as unknown[]).length;
  } catch {
    /* best-effort */
  }
  return 0;
}

const WIREFRAME_TABLE_ROWS = 3;

function flatText(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

/** A one-cell table with a fill is the callout (note box) shape, not a data table. */
function isCalloutTable(node: unknown, grid: ToonGrid | null): boolean {
  if (!grid || grid.rows !== 1 || grid.cols !== 1) return false;
  try {
    const xml = (node as { xml?: Element }).xml;
    const shd = xml?.getElementsByTagName('w:shd')[0];
    const fill = shd?.getAttribute('w:fill') ?? '';
    return /^[0-9A-Fa-f]{6}$/.test(fill);
  } catch {
    return false;
  }
}

/** Header row plus the first rows of a table, each cut to the wireframe width. */
function tableRowLines(grid: ToonGrid, width: number): string[] {
  const byRow = new Map<number, string[]>();
  for (const cell of grid.cells) {
    const list = byRow.get(cell.row) ?? [];
    if (!cell.cover) list.push(flatText(cell.text));
    byRow.set(cell.row, list);
  }
  const rowNumbers = [...byRow.keys()].sort((a, b) => a - b);
  const room = width - 6;
  const perCell = Math.max(4, Math.floor(room / Math.max(1, grid.cols)) - 3);
  const lines: string[] = [];
  for (const r of rowNumbers.slice(0, WIREFRAME_TABLE_ROWS)) {
    const cells = (byRow.get(r) ?? []).map((t) => (t.length > perCell ? t.slice(0, Math.max(1, perCell - 1)) + '~' : t));
    lines.push(`|   ${cells.join(' | ')}`.slice(0, width - 2).padEnd(width - 1) + '|');
  }
  if (rowNumbers.length > WIREFRAME_TABLE_ROWS) {
    lines.push(`|   (+${rowNumbers.length - WIREFRAME_TABLE_ROWS} more rows)`.padEnd(width - 1) + '|');
  }
  return lines;
}

function listNumIdOf(node: unknown): string {
  try {
    const xml = (node as { xml?: Element }).xml;
    return xml?.getElementsByTagName('w:numPr')[0]?.getElementsByTagName('w:numId')[0]?.getAttribute('w:val') ?? '';
  } catch {
    return '';
  }
}

function listLevelOf(node: unknown): number {
  try {
    const xml = (node as { xml?: Element }).xml;
    const v = xml?.getElementsByTagName('w:numPr')[0]?.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val');
    const n = v === null || v === undefined || v === '' ? 0 : parseInt(v, 10);
    return Number.isFinite(n) ? n : 0;
  } catch {
    return 0;
  }
}

function hasChartSignal(node: { type?: unknown; metadata?: Record<string, unknown> }): boolean {
  try {
    if ((node as { type?: unknown }).type === 'chart' || (node as { type?: unknown }).type === 'chart3d') return true;
    const meta = (node as { metadata?: Record<string, unknown> }).metadata ?? {};
    // Series is the honest chart signal. Bare `chartType` alone is NOT
    // sufficient: blank-template paragraphs carry a default `bar3D` filler
    // (no series, type paragraph) that must not render as [CHART/IMAGE].
    if (Array.isArray(meta.series) && (meta.series as unknown[]).length > 0) return true;
    const text = nodeText(node as { metadata?: { text?: unknown } });
    // TOON chart tags: [CH:bar3D] etc. Never count [CH:?]/[CH:-] placeholders.
    const m = text.match(/\[CH:([^\]]*)\]/);
    if (m && m[1].trim() !== '' && m[1].trim() !== '?' && m[1].trim() !== '-') return true;
  } catch {
    /* best-effort */
  }
  return false;
}

function hasImageSignal(node: { type?: unknown; metadata?: Record<string, unknown> }): boolean {
  try {
    if ((node as { type?: unknown }).type === 'image') return true;
    const meta = (node as { metadata?: Record<string, unknown> }).metadata ?? {};
    if (meta.hasImage === true) return true;
  } catch {
    /* best-effort */
  }
  return false;
}

function mentionsChart(text: string): boolean {
  // Whole words only: "paragraph" must not read as "graph".
  return /\b(charts?|figures?|graphs?|plots?)\b/i.test(text);
}

/**
 * Lint large, user-visible layout defects only (prompt law).
 * Catches: orphan heading, wide/overflow table, blank page, missing chart.
 * Ignores micro-typography by construction (no spacing/font/typo checks).
 */
export function lintVisualDefects(ast: DocumentAST, pageMap?: VisualPageMap | null): string[] {
  const warnings: string[] = [];
  try {
    const map = resolveMap(ast, pageMap);
    const byPage = new Map<number, Array<{ id: string; idx: number }>>();
    const order: Array<{ id: string; node: unknown }> = [];
    let idx = 0;
    for (const section of ast.sections ?? []) {
      for (const node of section.nodes ?? []) {
        const n = node as unknown as { id?: unknown; type?: unknown; metadata?: Record<string, unknown> };
        const id = typeof n?.id === 'string' ? n.id : '';
        if (!id) {
          idx++;
          continue;
        }
        const page = map.nodePages[id] ?? 1;
        if (!byPage.has(page)) byPage.set(page, []);
        byPage.get(page)?.push({ id, idx });
        order.push({ id, node: n });
        idx++;
      }
    }

    // Blank pages (large defect): numbered pages with zero nodes.
    for (let p = 1; p <= map.totalPages; p++) {
      if ((byPage.get(p) ?? []).length === 0) {
        warnings.push(`Blank page: page ${p} of ${map.totalPages} has no content nodes.`);
      }
    }

    for (const { id, node } of order) {
      const n = node as { type?: unknown; metadata?: Record<string, unknown> };
      const page = map.nodePages[id] ?? 1;
      if (isHeadingNode(n as { type?: unknown; metadata?: { namedStyle?: unknown; text?: unknown } })) {
        // Orphan heading: heading is the last node on its page but the doc
        // continues on the next page (title stranded without body text).
        const onPage = (byPage.get(page) ?? []).map((e) => e.id);
        const isLastOnPage = onPage.length > 0 && onPage[onPage.length - 1] === id;
        const hasNextPage = page < map.totalPages && ((byPage.get(page + 1) ?? []).length > 0 || order.some((e) => (map.nodePages[e.id] ?? 1) > page));
        if (isLastOnPage && hasNextPage) {
          const label = nodeText(n as { metadata?: { text?: unknown } }).slice(0, 80) || id;
          warnings.push(`Orphan heading: "${label}" ends page ${page} without its body text.`);
        }
      }
      if ((n as { type?: unknown }).type === 'table') {
        const cols = tableCols(n as { type?: unknown; metadata?: { tableDimensions?: unknown; headers?: unknown; rows?: unknown } });
        if (cols > VISUAL_WIDE_TABLE_COLS) {
          warnings.push(`Overflow table: page ${page} table (${cols} columns) risks margin clipping.`);
        }
        try {
          const dims = (n as { metadata?: { tableDimensions?: { rows?: unknown } } }).metadata?.tableDimensions;
          const rows = typeof dims?.rows === 'number' ? dims.rows : 0;
          if (rows > 25) {
            warnings.push(`Overflow table: page ${page} table (${rows} rows) likely splits across pages.`);
          }
        } catch {
          /* best-effort */
        }
      }
    }

    // Missing chart: a page whose text promises a chart/figure but carries
    // no chart/image node (planted-defect bait: heading "Q3 chart" + prose).
    for (let p = 1; p <= map.totalPages; p++) {
      const ids = new Set((byPage.get(p) ?? []).map((e) => e.id));
      let promisesChart = false;
      let hasVisual = false;
      for (const section of ast.sections ?? []) {
        for (const node of section.nodes ?? []) {
          const n = node as unknown as { id?: unknown };
          const id = typeof n?.id === 'string' ? (n.id as string) : '';
          if (!id || !ids.has(id)) continue;
          const nn = node as unknown as { type?: unknown; metadata?: Record<string, unknown> };
          if (hasChartSignal(nn) || hasImageSignal(nn)) hasVisual = true;
          if (mentionsChart(nodeText(nn as { metadata?: { text?: unknown } }))) promisesChart = true;
        }
      }
      if (promisesChart && !hasVisual) {
        warnings.push(`Missing chart: page ${p} mentions a chart/figure but has no chart or image node.`);
      }
    }
  } catch {
    /* lint never throws */
  }
  return warnings;
}

function pageBreakdown(ast: DocumentAST, map: VisualPageMap, totalPages = map.totalPages): VisualPageBreakdown[] {
  const rows = new Map<number, VisualPageBreakdown>();
  // When the renderer paginated beyond the word-layout map, cover the
  // renderer's pages too: extra rows stay honestly empty (nodeCount 0) —
  // no word-layout nodes are mapped past the map's own total.
  for (let p = 1; p <= Math.max(map.totalPages, totalPages); p++) {
    rows.set(p, { page: p, nodeCount: 0, headings: [], hasTables: false, hasCharts: false, hasImages: false });
  }
  try {
    for (const section of ast.sections ?? []) {
      for (const node of section.nodes ?? []) {
        const n = node as unknown as { id?: unknown; type?: unknown; metadata?: Record<string, unknown> };
        const id = typeof n?.id === 'string' ? (n.id as string) : '';
        const page = (id && map.nodePages[id]) || 1;
        if (!rows.has(page)) {
          rows.set(page, { page, nodeCount: 0, headings: [], hasTables: false, hasCharts: false, hasImages: false });
        }
        const bucket = rows.get(page) as VisualPageBreakdown;
        bucket.nodeCount++;
        if (isHeadingNode(n as { type?: unknown; metadata?: { namedStyle?: unknown; text?: unknown } })) {
          bucket.headings.push(nodeText(n as { metadata?: { text?: unknown } }).slice(0, 60) || (id ?? ''));
        }
        if ((n as { type?: unknown }).type === 'table') bucket.hasTables = true;
        if (hasChartSignal(n as { type?: unknown; metadata?: Record<string, unknown> })) bucket.hasCharts = true;
        if (hasImageSignal(n as { type?: unknown; metadata?: Record<string, unknown> })) bucket.hasImages = true;
      }
    }
  } catch {
    /* best-effort */
  }
  return [...rows.values()].sort((a, b) => a.page - b.page);
}

/**
 * ASCII wireframe for one page. Bounded: at most `maxNodes` content rows so
 * the telemetry budget holds even for dense pages.
 */
export function generateVisualWireframe(
  ast: DocumentAST,
  targetPage = 1,
  pageMap?: VisualPageMap | null,
  maxNodes = 24,
  opts: { renderTotalPages?: number } = {}
): string {
  const map = resolveMap(ast, pageMap);
  // Renderer single source of truth: when the caller passes the renderer's
  // (LibreOffice) page count and it disagrees with the word-layout map, the
  // `[PAGE n OF m]` marker reports the RENDERER's `m` — the image on disk
  // came from the renderer, so the marker must describe it.
  const renderTotal = normalizeRenderTotal(opts.renderTotalPages, map.totalPages);
  const markerTotal = renderTotal ?? map.totalPages;
  const page = Number.isFinite(targetPage) && targetPage >= 1 ? Math.floor(targetPage) : 1;
  const W = VISUAL_WIREFRAME_WIDTH;
  const border = '+' + '-'.repeat(W - 2) + '+';
  const emptyLine = '|' + ' '.repeat(W - 2) + '|';
  const lines: string[] = [border];
  const tag = map.isEstimate ? ' (estimate)' : '';
  lines.push(`| [PAGE ${page} OF ${markerTotal}${tag}]`.padEnd(W - 1) + '|');
  lines.push(border);

  const pageNodes: unknown[] = [];
  try {
    for (const section of ast.sections ?? []) {
      for (const node of section.nodes ?? []) {
        const n = node as unknown as { id?: unknown };
        const id = typeof n?.id === 'string' ? (n.id as string) : '';
        if ((map.nodePages[id] ?? 1) === page) pageNodes.push(node);
      }
    }
  } catch {
    /* best-effort */
  }
  if (pageNodes.length === 0) {
    lines.push(emptyLine);
    if (renderTotal !== undefined) {
      // Mismatch regime with zero mapped nodes: the renderer DID paginate to
      // this page (it exists in the rendered image), so claiming
      // "(Empty Page / Section Break)" would describe the word-layout map,
      // not the rendered page. Say what is (and is not) known instead.
      lines.push(`| (renderer page ${page}/${markerTotal}: no word-layout nodes mapped)`.padEnd(W - 1) + '|');
    } else {
      lines.push('| (Empty Page / Section Break)'.padEnd(W - 1) + '|');
    }
    lines.push(emptyLine);
    lines.push(border);
    return lines.join('\n');
  }
  const shown = pageNodes.slice(0, Math.max(1, maxNodes));
  // Running item numbers per list level; any non-list node ends the run.
  let counters: number[] = [];
  let kinds: string[] = [];
  for (const raw of shown) {
    const node = raw as { type?: unknown; metadata?: Record<string, unknown> };
    const isListItem =
      (node as { type?: unknown }).type !== 'table' &&
      !isHeadingNode(node as { type?: unknown; metadata?: { namedStyle?: unknown; text?: unknown } }) &&
      (node.metadata?.['listLevel'] !== undefined || hasListNumbering(node as never));
    if (!isListItem) {
      counters = [];
      kinds = [];
    }
    const style = nodeStyle(node as { metadata?: { namedStyle?: unknown } });
    const text = nodeText(node as { metadata?: { text?: unknown } }).trim();
    if (isHeadingNode(node as { type?: unknown; metadata?: { namedStyle?: unknown; text?: unknown } })) {
      lines.push(emptyLine);
      const prefix = style.includes('1') ? '# ' : style.includes('2') ? '## ' : '### ';
      lines.push(`| ${prefix}${text}`.slice(0, W - 2).padEnd(W - 1) + '|');
    } else if ((node as { type?: unknown }).type === 'table') {
      lines.push(emptyLine);
      const grid = tableGridOf(node);
      if (isCalloutTable(node, grid)) {
        const note = flatText(grid?.cells[0]?.text ?? '');
        const body = note.length > 48 ? note.slice(0, 45) + '...' : note;
        lines.push(`| [CALLOUT] ${body}`.slice(0, W - 2).padEnd(W - 1) + '|');
      } else {
        const cols = tableCols(node as { type?: unknown; metadata?: { tableDimensions?: unknown; headers?: unknown; rows?: unknown } });
        lines.push(`| [TABLE: ${cols || '?'} cols]`.padEnd(W - 1) + '|');
        if (grid) lines.push(...tableRowLines(grid, W));
        else lines.push(('| ' + '+---'.repeat(Math.min(Math.max(cols, 1), 4)).padEnd(W - 4) + ' + |').slice(0, W - 1).padEnd(W - 1) + '|');
      }
    } else if (isListItem) {
      const level = Math.max(0, Math.min(8, Number(node.metadata?.['listLevel'] ?? listLevelOf(node)) || 0));
      const ordered = isOrderedListItem(node as never, ast);
      const kind = `${ordered ? 'numbered' : 'bullet'}:${listNumIdOf(node)}`;
      counters.length = level + 1;
      kinds.length = level + 1;
      // A different list (kind or numbering id) at the same level starts its own count.
      counters[level] = kinds[level] === kind ? (counters[level] ?? 0) + 1 : 1;
      kinds[level] = kind;
      const marker = ordered ? `${counters[level]}.` : '-';
      const body = text.length > 56 ? text.slice(0, 53) + '...' : text;
      lines.push(`| ${'  '.repeat(level)}${marker} ${body}`.slice(0, W - 2).padEnd(W - 1) + '|');
    } else if (hasChartSignal(node) || hasImageSignal(node)) {
      lines.push(emptyLine);
      lines.push('| [CHART / IMAGE]'.padEnd(W - 1) + '|');
    } else if (text) {
      const truncated = text.length > 60 ? text.slice(0, 57) + '...' : text;
      lines.push(`| ${truncated}`.padEnd(W - 1) + '|');
    }
  }
  if (pageNodes.length > shown.length) {
    lines.push(`| ... (${pageNodes.length - shown.length} more nodes, budget-truncated)`.padEnd(W - 1) + '|');
  }
  lines.push(emptyLine);
  lines.push(border);
  return lines.join('\n');
}

/**
 * Full telemetry for one page: wireframe + large-defect warnings, with the
 * token budget enforced. Shrinks `maxNodes` until
 * `tokens <= maxTokens` (default 500) or one node remains.
 */
export function buildVisualTelemetry(
  ast: DocumentAST,
  opts: { page?: number; pageMap?: VisualPageMap | null; maxTokens?: number; renderTotalPages?: number } = {}
): VisualTelemetry {
  const map = resolveMap(ast, opts.pageMap ?? null);
  const page = opts.page ?? 1;
  // Mismatch-only value (see normalizeRenderTotal): defined iff the
  // renderer's page count disagrees with the word-layout map.
  const renderTotal = normalizeRenderTotal(opts.renderTotalPages, map.totalPages);
  const markerTotal = renderTotal ?? map.totalPages;
  const maxTokens = opts.maxTokens ?? VISUAL_TELEMETRY_TOKEN_BUDGET;
  const warnings = lintVisualDefects(ast, map);
  if (renderTotal !== undefined) {
    // The trailing clause only has content when the renderer paginated
    // beyond the word-layout map; on a reverse mismatch (renderer stopped
    // before the map's last page) nothing exists past the map, so the
    // clause would be vacuous and is dropped.
    const pastMapClause = markerTotal > map.totalPages ? '; nodes past the word-layout map are unmapped.' : '.';
    warnings.push(
      `Pagination mismatch: renderer produced ${pageCountLabel(markerTotal)}; word-layout page map totals ${map.totalPages} — [PAGE n OF m] uses the renderer count${pastMapClause}`
    );
  }
  const breakdown = pageBreakdown(ast, map, markerTotal);
  let maxNodes = 24;
  let wireframe = generateVisualWireframe(ast, page, map, maxNodes, { renderTotalPages: renderTotal });
  let truncated = false;
  const warningText = warnings.length > 0 ? warnings.join('\n') : '';
  let tokens = estimateTokenCount(wireframe + '\n' + warningText);
  while (tokens > maxTokens && maxNodes > 2) {
    maxNodes = Math.max(2, Math.floor(maxNodes / 2));
    wireframe = generateVisualWireframe(ast, page, map, maxNodes, { renderTotalPages: renderTotal });
    tokens = estimateTokenCount(wireframe + '\n' + warningText);
    truncated = true;
  }
  if (tokens > maxTokens) {
    // Last resort: drop the wireframe body, keep header + warnings head.
    wireframe = generateVisualWireframe(ast, page, map, 1, { renderTotalPages: renderTotal });
    tokens = estimateTokenCount(wireframe + '\n' + warningText);
    truncated = true;
  }
  return {
    totalPages: markerTotal,
    ...(renderTotal !== undefined ? { wordLayoutTotalPages: map.totalPages } : {}),
    warnings,
    pageBreakdown: breakdown,
    wireframe,
    tokens,
    truncated,
    mapSource: map.source,
  };
}
