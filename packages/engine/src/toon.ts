import { XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST, Section, Node, HeadingNode } from './ast';
import { hasListNumbering, isOrderedListItem } from './export/markdown';

/** Track-B tiered TOON fidelity. compact=current outline, standard=default enriched, full=section-scoped dump. */
export type ToonFidelity = 'compact' | 'standard' | 'full';

export const TOON_DEFAULT_FIDELITY: ToonFidelity = 'standard';
/** Standard-tier token overhead cap vs compact (measured on rendered text). */
export const TOON_STANDARD_OVERHEAD_CAP = 1.5;

export function isToonFidelity(value: unknown): value is ToonFidelity {
  return value === 'compact' || value === 'standard' || value === 'full';
}

export function assertToonFidelity(value: unknown): asserts value is ToonFidelity {
  if (!isToonFidelity(value)) {
    throw new Error(
      `Unknown TOON fidelity '${String(value)}'. Supported: compact|standard|full. ` +
        `Remedy: call toon({ fidelity: 'standard' }) for the default enriched outline, ` +
        `'compact' for the minimal outline, or 'full' with a sectionId for rPr/pPr detail.`
    );
  }
}

export interface ToonRequest {
  fidelity?: ToonFidelity;
  sectionId?: string;
  maxTextLength?: number;
  useCache?: boolean;
}

export interface ToonRow {
  id: string;
  version: number;
  path: string;        // Section breadcrumb (e.g. "Intro > Background")
  header: string;      // Section heading (e.g. "Background")
  style: string;       // Style info (e.g. "heading1:L0", "body:L-1", "list:bullet:L2")
  text: string;
  flags: string;
  /** Full-tier only: serialized w:pPr + w:rPr dump. Absent for compact/standard. */
  detail?: string;
}

export interface ToonTable {
  headers: string[];
  rows: ToonRow[];
  /**
   * ST-11: TOON is lossy by design — fidelity flag surfaced so AI consumers
   * know style detail was collapsed. Always 'lossy' (never 'full').
   * Docs/test only: no renderer change.
   */
  fidelity?: 'lossy';
  /** Track-B tier that produced this table. Always set by generateToon/toon. */
  tier: ToonFidelity;
}

/**
 * ST-11: documented TOON lossy allowlist.
 * TOON preserves: node id/version, section path/header breadcrumb, collapsed
 * style token (heading/body/table/chart/list/image/equation + namedStyle/listLevel),
 * plain text (truncated), table grid cells as rNcM coordinates with (spanN) /
 * hidden-cover markers, chart title + categories + series names/values,
 * structural flags ([TBL]/[IMG]/[EQ]/[CITE]/[CMT]/
 * [TRK]/[LNK]/[BKMK]/[L#]/[+ADDED]/[~MODIFIED]).
 * TOON drops (lossy): bold/italic/underline/color/size/font, paragraph
 * borders/shading, table tblBorders/tblCellMar/banding, theme colors,
 * numbering prefixes/indents, image float/wrap, header/footer zones,
 * watermarks, tracked-change ins/del styling. AI must not assume cascade
 * fidelity from TOON alone.
 *
 * Track-B standard tier recovers a terse, sparse projection of the dropped
 * run/paragraph signals ([B]/[I]/[C:hex]/[SZ:val]/[RS:name]/[N:id]/[J:val] +
 * [HDR]/[ZEBRA] + [CH:type]/[RID:id]) at <1.5x compact cost. Full tier dumps
 * raw w:pPr/w:rPr for one focused section only.
 */
export const TOON_FIDELITY_ALLOWLIST: Readonly<{
  fidelity: 'lossy';
  preserved: readonly string[];
  dropped: readonly string[];
}> = {
  fidelity: 'lossy',
  preserved: [
    'id', 'version', 'path', 'header', 'style-token', 'text', 'flags',
    'heading/body/table/chart/list/image/equation', 'namedStyle', 'listLevel',
    'grid-coordinates', 'span-markers', 'hidden-cover', 'chart-title-categories-series',
    '[TBL]', '[IMG]', '[EQ]', '[CITE]', '[CMT]', '[TRK]', '[LNK]', '[BKMK]'
  ],
  dropped: [
    'bold', 'italic', 'underline', 'color', 'fontSize', 'fontFamily',
    'pBdr', 'shading', 'tblBorders', 'tblCellMar', 'banding',
    'themeColor', 'asciiTheme', 'numbering-prefix', 'numbering-indent',
    'image-float', 'header-footer-zones', 'watermark', 'ins-del-styling'
  ]
};

/**
 * TOON flag legend: every bracket code (plus the grid markers) the generator
 * can emit, with its meaning. Docs tables mirror this const; tests assert
 * programmatically that every emitted token has an entry here.
 * Keys: exact tokens (`TBL`, `CH:?`, `Image`) or prefixes (`C:` matches `[C:FF0000]`,
 * `L#` matches `[L1]`, `Chart` matches `[Chart column]`, `Table` matches
 * `[Table 3x3]`); `span` matches `(span2)`; `hidden-cover` is literal.
 */
export const TOON_FLAG_LEGEND: Readonly<Record<string, string>> = {
  TBL: 'table node (grid cells follow as rNcM coordinates, 1-based; edit one with target "<tableId>:r<row>c<col>")',
  IMG: 'image/drawing present (chart drawings included)',
  EQ: 'equation present',
  CITE: 'citation field present',
  CMT: 'comment present',
  TRK: 'tracked changes present',
  LNK: 'hyperlink present',
  BKMK: 'bookmark present',
  'L#': 'list level, counting from 1: 1 is the top level (e.g. [L1], [L3])',
  '+ADDED': 'diff row: added since the base table',
  '~MODIFIED': 'diff row: changed since the base table',
  B: 'bold run present (standard tier)',
  I: 'italic run present (standard tier)',
  'C:': 'text color hex without # (e.g. [C:FF0000])',
  'SZ:': 'font size in w:sz half-points (e.g. [SZ:28] = 14pt)',
  'RS:': 'run style name (e.g. [RS:Emphasis])',
  'J:': 'paragraph justification (e.g. [J:center])',
  'N:': 'numbering id/level, level counting from 1 (e.g. [N:5/1]) or list level fallback ([N:L2])',
  HDR: 'table header row styled (tblLook firstRow)',
  ZEBRA: 'table banded rows (tblLook banding on)',
  'CH:': 'chart type (e.g. [CH:column], [CH:bar3D])',
  'CH:?': 'chart type unknown (explicit; never fabricated)',
  'RID:': 'chart relationship id (e.g. [RID:rId5])',
  span: 'grid cell spans N logical columns (e.g. r2c1 Alpha(span2))',
  'hidden-cover': 'merge-cover cell, not directly editable (address the restart cell)',
  Image: 'image placeholder (drawing with no readable text)',
  Equation: 'equation placeholder (Office Math with no w:t text)',
  Field: 'field placeholder (no readable text)',
  'Page Break': 'page-break placeholder',
  Table: 'empty-table placeholder [Table RxC] (populated tables render grid cells instead)',
  Chart: '2D chart head [Chart type] title | cats | series (read from the chart part)',
  '3D Chart': '3D chart head [3D Chart type] title | cats | series (read from the chart part)',
};

/** Normalizes one emitted token to its legend key (`[C:FF00]` -> `C:`, `[L2]` -> `L#`). */
export function toonFlagLegendKey(token: string): string {
  const t = token.trim();
  if (t === 'span' || /^span\d+$/.test(t)) return 'span';
  if (t === 'hidden-cover') return 'hidden-cover';
  // Text-column placeholders (bracketed codes agents also see, not flags).
  if (t === 'Image' || t === 'Equation' || t === 'Field' || t === 'Page Break') return t;
  if (t.startsWith('3D Chart')) return '3D Chart';
  if (t.startsWith('Chart')) return 'Chart';
  if (t.startsWith('Table')) return 'Table';
  if (
    t === 'TBL' || t === 'IMG' || t === 'EQ' || t === 'CITE' || t === 'CMT' ||
    t === 'TRK' || t === 'LNK' || t === 'BKMK' || t === 'B' || t === 'I' ||
    t === 'HDR' || t === 'ZEBRA' || t === 'CH:?' || t === '+ADDED' || t === '~MODIFIED'
  ) return t;
  if (/^L\d+$/.test(t)) return 'L#';
  const colon = t.indexOf(':');
  if (colon > 0) {
    const prefix = `${t.slice(0, colon)}:`;
    if (prefix === 'C:' || prefix === 'SZ:' || prefix === 'RS:' || prefix === 'J:' || prefix === 'N:' || prefix === 'CH:' || prefix === 'RID:') {
      return prefix;
    }
  }
  return t;
}

/** Extracts every `[...]` token, `(spanN)` marker, and `hidden-cover` tag from rendered TOON text. */
export function collectToonTokens(rendered: string): string[] {
  const out: string[] = [];
  try {
    const bracket = rendered.match(/\[[^\]\n[\]]+\]/g) ?? [];
    for (const b of bracket) {
      const inner = b.slice(1, -1).trim();
      if (inner) out.push(inner);
    }
    const spans = rendered.match(/\(span\d+\)/g) ?? [];
    for (const s of spans) out.push(s.slice(1, -1));
    if (/(^|[\s|;])hidden-cover([\s|;]|$)/m.test(rendered)) out.push('hidden-cover');
  } catch {
    return out;
  }
  return out;
}

export interface ToonRenderOptions {
  compact?: boolean;             // Strip redundant spacing, trim empty flags
  maxTextLength?: number;        // Max characters for text column before truncation
  omitColumns?: Array<'path' | 'header' | 'style' | 'flags' | 'detail'>;
  includeHeaderRow?: boolean;
  /** Full-tier only: include the Detail column (default true when rows carry detail). */
  includeDetail?: boolean;
  /**
   * Human-presentation text-column wrap: when set, the Text column is
   * word-wrapped so no emitted row line exceeds `width` display columns.
   * Continuation lines carry a hanging indent of spaces aligning under the
   * row's own Text cell start. Every other column (and the fidelity/header/
   * separator rows) is never wrapped. Omitted by default: default rendering
   * stays strict 1-line-per-node for machine consumers; unwrapToonText
   * recovers the canonical single-line form.
   */
  width?: number;
}

export interface ToonGenerateOptions {
  sectionId?: string;
  maxTextLength?: number;
  useCache?: boolean;
  fidelity?: ToonFidelity;
}

export interface ToonDiffOptions {
  compact?: boolean;
  showUnchangedStatus?: boolean;
}

interface RunFmt {
  bold: boolean;
  italic: boolean;
  color: string | undefined;
  sz: string | undefined;
  rStyle: string | undefined;
}

interface ParaFmt {
  jc: string | undefined;
  numId: string | undefined;
  ilvl: string | undefined;
}

function firstDescendant(parent: Element, tag: string): Element | null {
  try {
    const list = parent.getElementsByTagName(tag);
    return list.length > 0 ? (list[0] as Element) : null;
  } catch {
    return null;
  }
}

function attrOf(el: Element | null, name: string): string | undefined {
  if (!el) return undefined;
  try {
    const v = el.getAttribute(name);
    return v === null || v === '' ? undefined : v;
  } catch {
    return undefined;
  }
}

/** Aggregate run-level formatting across all w:r children (any-run-wins for B/I, first-seen for valued). */
function extractRunFmt(xml: Element): RunFmt {
  let bold = false;
  let italic = false;
  let color: string | undefined;
  let sz: string | undefined;
  let rStyle: string | undefined;
  let runs: { length: number; item(i: number): Element | null } | null = null;
  try {
    const raw = xml.getElementsByTagName('w:r');
    runs = raw as unknown as { length: number; item(i: number): Element | null };
  } catch {
    return { bold: false, italic: false, color: undefined, sz: undefined, rStyle: undefined };
  }
  if (!runs) return { bold: false, italic: false, color: undefined, sz: undefined, rStyle: undefined };
  const count = runs.length;
  for (let i = 0; i < count; i++) {
    let run: Element | null = null;
    try {
      const maybe = (runs as unknown as { [k: number]: Element })[i];
      run = (maybe ?? null) as Element | null;
    } catch {
      continue;
    }
    if (!run || typeof run.getElementsByTagName !== 'function') continue;
    let rPr: Element | null = null;
    try {
      const prs = run.getElementsByTagName('w:rPr');
      rPr = prs.length > 0 ? (prs[0] as Element) : null;
    } catch {
      continue;
    }
    if (!rPr) continue;
    try {
      if (!bold && rPr.getElementsByTagName('w:b').length > 0) bold = true;
      if (!italic && rPr.getElementsByTagName('w:i').length > 0) italic = true;
      if (color === undefined) {
        const c = attrOf(firstDescendant(rPr, 'w:color'), 'w:val');
        if (c) color = c.replace(/^#/, '').toUpperCase().slice(0, 6);
      }
      if (sz === undefined) {
        const s = attrOf(firstDescendant(rPr, 'w:sz'), 'w:val');
        if (s) sz = s;
      }
      if (rStyle === undefined) {
        const rs = attrOf(firstDescendant(rPr, 'w:rStyle'), 'w:val');
        if (rs) rStyle = rs;
      }
    } catch {
      continue;
    }
    if (bold && italic && color !== undefined && sz !== undefined && rStyle !== undefined) break;
  }
  return { bold, italic, color, sz, rStyle };
}

/** Note-reference styles present across all runs (footnote/endnote markers). */
function collectNoteReferenceStyles(xml: Element): { footnote: boolean; endnote: boolean } {
  let footnote = false;
  let endnote = false;
  try {
    const runs = xml.getElementsByTagName('w:r');
    for (let i = 0; i < runs.length; i++) {
      const run = runs[i] as Element;
      if (!run || typeof run.getElementsByTagName !== 'function') continue;
      const prs = run.getElementsByTagName('w:rPr');
      for (let p = 0; p < prs.length; p++) {
        const styles = (prs[p] as Element).getElementsByTagName('w:rStyle');
        for (let s = 0; s < styles.length; s++) {
          let v: string | null = null;
          try {
            v = (styles[s] as Element).getAttribute('w:val');
          } catch {
            continue;
          }
          if (v === 'FootnoteReference') footnote = true;
          else if (v === 'EndnoteReference') endnote = true;
          if (footnote && endnote) return { footnote, endnote };
        }
      }
    }
  } catch {
    /* best-effort: the single-style aggregate above stands */
  }
  return { footnote, endnote };
}

function extractParaFmt(xml: Element): ParaFmt {
  let jc: string | undefined;
  let numId: string | undefined;
  let ilvl: string | undefined;
  try {
    const pPrs = xml.getElementsByTagName('w:pPr');
    if (pPrs.length === 0) return { jc: undefined, numId: undefined, ilvl: undefined };
    const pPr = pPrs[0] as Element;
    jc = attrOf(firstDescendant(pPr, 'w:jc'), 'w:val');
    const numPr = firstDescendant(pPr, 'w:numPr');
    if (numPr) {
      numId = attrOf(firstDescendant(numPr, 'w:numId'), 'w:val');
      ilvl = attrOf(firstDescendant(numPr, 'w:ilvl'), 'w:val');
    }
  } catch {
    return { jc: undefined, numId: undefined, ilvl: undefined };
  }
  return { jc, numId, ilvl };
}

function extractTableLook(xml: Element): { firstRow: boolean; zebra: boolean } {
  try {
    const tblPrs = xml.getElementsByTagName('w:tblPr');
    if (tblPrs.length === 0) return { firstRow: false, zebra: false };
    const look = firstDescendant(tblPrs[0] as Element, 'w:tblLook');
    if (!look) return { firstRow: false, zebra: false };
    const firstRow = attrOf(look, 'w:firstRow') === '1' || attrOf(look, 'w:firstRow') === 'true';
    const noHBand = attrOf(look, 'w:noHBand');
    const zebra = noHBand === '0' || noHBand === 'false';
    return { firstRow, zebra };
  } catch {
    return { firstRow: false, zebra: false };
  }
}

function extractChartRId(xml: Element): string | undefined {
  try {
    const charts = xml.getElementsByTagName('c:chart');
    if (charts.length === 0) return undefined;
    return attrOf(charts[0] as Element, 'r:id');
  } catch {
    return undefined;
  }
}

/** Direct element children with a tag (nested w:tbl rows must not pollute the outer grid). */
function directKids(parent: Element, tag: string): Element[] {
  const out: Element[] = [];
  try {
    const kids = parent.childNodes;
    for (let i = 0; i < kids.length; i++) {
      const child = kids[i] as Element;
      if (child.nodeType === 1 && child.tagName === tag) out.push(child);
    }
  } catch {
    return [];
  }
  return out;
}

/** Concatenated w:t text of one table cell, scoped to direct w:p kids so a
 * nested w:tbl inside the cell never pollutes the outer text (best-effort). */
function cellTextOf(tc: Element): string {
  let text = '';
  try {
    const paras = directKids(tc, 'w:p');
    for (const p of paras) {
      const ts = p.getElementsByTagName('w:t');
      for (let i = 0; i < ts.length; i++) text += (ts[i] as Element).textContent || '';
    }
  } catch {
    return '';
  }
  return text;
}

/** One addressable grid cell: 1-based row/col, span width, merge-cover marker. */
export interface ToonGridCell {
  row: number;
  col: number;
  text: string;
  span: number;
  cover: boolean;
}

export interface ToonGrid {
  rows: number;
  cols: number;
  cells: ToonGridCell[];
}

/** Monospace display length in code points (no wide-char dependency). */
function displayLen(value: string): number {
  return [...value].length;
}

/**
 * Greedy word wrap of one Text cell to a per-line budget. Breaks only at
 * whitespace runs (each run is a wrap point consumed as the break), so
 * rejoining the segments with single spaces recovers text whose internal
 * separators are single spaces. A single word longer than `budget` occupies
 * its own line whole (overflows rather than splitting mid-word) so the
 * rejoin stays exact. Returns [text] when budget <= 0 or nothing breaks.
 */
function wrapTextCell(text: string, budget: number): string[] {
  if (budget <= 0) return [text];
  const words = text.split(/\s+/).filter((w) => w.length > 0);
  if (words.length === 0) return [text];
  const lines: string[] = [];
  let cur = '';
  let curLen = 0;
  for (const w of words) {
    const wLen = displayLen(w);
    if (cur === '') {
      cur = w;
      curLen = wLen;
    } else if (curLen + 1 + wLen <= budget) {
      cur += ` ${w}`;
      curLen += 1 + wLen;
    } else {
      lines.push(cur);
      cur = w;
      curLen = wLen;
    }
  }
  if (cur !== '') lines.push(cur);
  return lines.length > 0 ? lines : [text];
}

/**
 * Inverse of width wrapping: folds hanging-indent continuation lines back
 * into their row's Text cell, recovering the canonical 1-line-per-node
 * form. A continuation line is any line starting with whitespace while a
 * logical row is open; its indent is stripped and it rejoins with a single
 * space (the wrap-consumed separator). Fidelity/header/separator rows never
 * start with whitespace and pass through untouched. Round-trips
 * byte-identically when wrapped Text cells use single-space separators.
 */
export function unwrapToonText(rendered: string): string {
  const endsNewline = rendered.endsWith('\n');
  const lines = rendered.split('\n');
  if (endsNewline && lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
  const logical: string[] = [];
  for (const line of lines) {
    if (line.length > 0 && /^\s/.test(line) && logical.length > 0) {
      logical[logical.length - 1] += ` ${line.replace(/^\s+/, '')}`;
    } else {
      logical.push(line);
    }
  }
  return logical.join('\n') + (endsNewline ? '\n' : '');
}

/**
 * Read-only grid extraction from a w:tbl element (surfacing only, never mutates).
 * Coordinates are logical: a gridSpan=N cell occupies N logical columns;
 * each hMerge/vMerge continue cell occupies one (tagged cover).
 */
export function extractTableGrid(tbl: Element): ToonGrid {
  const cells: ToonGridCell[] = [];
  let maxCols = 0;
  let trs: Element[];
  try {
    trs = directKids(tbl, 'w:tr');
  } catch {
    return { rows: 0, cols: 0, cells };
  }
  for (let r = 0; r < trs.length; r++) {
    let tcs: Element[];
    try {
      tcs = directKids(trs[r], 'w:tc');
    } catch {
      continue;
    }
    let logical = 0;
    for (let c = 0; c < tcs.length; c++) {
      logical += 1;
      const tc = tcs[c];
      let span = 1;
      let cover = false;
      try {
        // Scoped to direct kids: a nested w:tbl's tcPr/gridSpan/hMerge must
        // never leak into the outer cell's span/cover markers.
        const tcPr = directKids(tc, 'w:tcPr')[0] ?? null;
        if (tcPr) {
          const gs = directKids(tcPr, 'w:gridSpan')[0] as Element | undefined;
          const gv = gs ? attrOf(gs, 'w:val') : undefined;
          const n = gv !== undefined ? parseInt(gv, 10) : 1;
          if (Number.isFinite(n) && n > 1) span = n;
          const hm = directKids(tcPr, 'w:hMerge')[0] as Element | undefined;
          if (hm) {
            const hv = attrOf(hm, 'w:val');
            // ECMA-376: bare <w:hMerge/> (no val) means continue (cover);
            // val="restart" marks the merge origin.
            if (hv === undefined || hv === '' || hv.toLowerCase() === 'continue') cover = true;
          }
          const vm = directKids(tcPr, 'w:vMerge')[0] as Element | undefined;
          if (vm) {
            const vv = attrOf(vm, 'w:val');
            if (vv === undefined || vv === '' || vv.toLowerCase() === 'continue') cover = true;
          }
        }
      } catch {
        /* markers stay default */
      }
      const end = logical + span - 1;
      if (end > maxCols) maxCols = end;
      cells.push({ row: r + 1, col: logical, text: cover ? '' : cellTextOf(tc), span, cover });
      logical = end;
    }
  }
  return { rows: trs.length, cols: maxCols, cells };
}

/** Lossy cell-text cleanup so delimiters stay unambiguous (TOON is lossy by design). */
function sanitizeGridText(value: string): string {
  return value.replace(/\s+/g, ' ').replace(/\|/g, '/').replace(/;/g, ',').trim();
}

function renderGridCell(c: ToonGridCell): string {
  const coord = `r${c.row}c${c.col}`;
  if (c.cover) return `${coord} hidden-cover`;
  const t = sanitizeGridText(c.text);
  if (c.span > 1) return t ? `${coord} ${t}(span${c.span})` : `${coord} (span${c.span})`;
  return t ? `${coord} ${t}` : coord;
}

/**
 * Terse grid rendering: cells pipe-delimited within a row, rows
 * semicolon-delimited (`r1c1 Name | r1c2 Q1 ; r2c1 Alpha(span2) | ...`).
 * Agents address exact cells from this text alone.
 */
export function renderTableGrid(grid: ToonGrid): string {
  const byRow = new Map<number, ToonGridCell[]>();
  for (const c of grid.cells) {
    const list = byRow.get(c.row);
    if (list) list.push(c);
    else byRow.set(c.row, [c]);
  }
  const rows = [...byRow.keys()].sort((a, b) => a - b);
  return rows.map((r) => (byRow.get(r) ?? []).map(renderGridCell).join(' | ')).join(' ; ');
}

/**
 * Read-only chart surfacing from parser-hydrated metadata (title/categories/
 * series read from the existing chart part at load; never fabricated here).
 * Terse: `[Chart column] Title | cats: Q1,Q2 | ARR: 10,25`.
 */
export function renderChartText(node: Node): string {
  const kind = node.type === 'chart3d' ? '3D Chart' : 'Chart';
  const data = node.metadata.chartData as {
    title?: unknown; categories?: unknown; series?: unknown; synthetic?: unknown;
  } | null | undefined;
  // Provenance: synthetic entries (placeholder-free hydration) render
  // explicit-unknown, never a fabricated title/type.
  const isSynthetic = !!data && typeof data === 'object' && (data as { synthetic?: unknown }).synthetic === true;
  if (isSynthetic) return `[${kind} ?]`;
  // 24 keeps every supported chartType intact (columnPercentStacked = 20);
  // a truncated token cannot round-trip into the update enum.
  const rawCt: unknown = node.metadata.chartType;
  const ct = typeof rawCt === 'string' && rawCt.length > 0 ? sanitizeToken(rawCt, 24) : '?';
  if (!data || typeof data !== 'object') return `[${kind} ${ct}]`;
  const title = typeof data.title === 'string' ? sanitizeGridText(data.title) : '';
  const cats = Array.isArray(data.categories) ? (data.categories as unknown[]).map((c) => sanitizeGridText(String(c))) : [];
  const series = Array.isArray(data.series)
    ? (data.series as Array<{ name?: unknown; values?: unknown; xValues?: unknown; unpairedX?: unknown; unpairedY?: unknown }>)
    : [];
  // A legal scatter carries an x list per series. Print per-point coordinates
  // instead of a shared cats row; points without an x value are reported as
  // unknown rather than filled with a point index.
  const hasPerSeriesX = series.some((s) => s && Array.isArray(s.xValues) && s.xValues.length > 0);
  const head = title ? `[${kind} ${ct}] ${title}` : `[${kind} ${ct}]`;
  const parts = [head];
  if (!hasPerSeriesX && cats.length > 0) parts.push(`cats: ${cats.join(',')}`);
  for (const s of series) {
    const name = s && typeof s.name === 'string' && s.name ? sanitizeGridText(s.name) : 'Series';
    const rawValues = s && Array.isArray(s.values) ? (s.values as unknown[]) : [];
    if (hasPerSeriesX) {
      // When any series carries xValues, a series without one has an unknown
      // x list from a malformed part: do not substitute another series' x.
      const xs = Array.isArray(s.xValues) ? (s.xValues as unknown[]) : [];
      const paired = Math.min(rawValues.length, xs.length);
      const pairs = rawValues.slice(0, paired).map((v, i) => `(${String(xs[i])}, ${String(v)})`).join(',');
      const missingY = rawValues.length - paired + (typeof s.unpairedY === 'number' && s.unpairedY > 0 ? s.unpairedY : 0);
      const missingX = typeof s.unpairedX === 'number' && s.unpairedX > 0 ? s.unpairedX : 0;
      const notes: string[] = [];
      if (missingY > 0) notes.push(`${missingY} y value${missingY === 1 ? '' : 's'} without x`);
      if (missingX > 0) notes.push(`${missingX} x value${missingX === 1 ? '' : 's'} without y`);
      const body = [pairs, notes.join('; ')].filter(Boolean).join('; ');
      parts.push(body ? `${name}: ${body}` : name);
      continue;
    }
    const vals = rawValues.map((v) => String(v)).join(',');
    parts.push(vals ? `${name}: ${vals}` : name);
  }
  return parts.join(' | ');
}

function sanitizeToken(value: string, maxLen: number): string {
  const clean = value.replace(/[\s|[\]]/g, '_').slice(0, maxLen);
  return clean.length > 0 ? clean : '';
}

/** Terse sparse standard-tier suffix appended to the base flags column. */
function buildStandardSuffix(node: Node): string {
  const parts: string[] = [];
  let xml: Element | null = null;
  try {
    xml = node.xml ?? null;
  } catch {
    return '';
  }
  if (!xml || typeof xml.getElementsByTagName !== 'function') return '';

  const run = extractRunFmt(xml);
  if (run.bold) parts.push('[B]');
  if (run.italic) parts.push('[I]');
  if (run.color) {
    const c = sanitizeToken(run.color, 6);
    if (c) parts.push(`[C:${c}]`);
  }
  if (run.sz) {
    const s = sanitizeToken(run.sz, 4);
    if (s) parts.push(`[SZ:${s}]`);
  }
  if (run.rStyle) {
    const rs = sanitizeToken(run.rStyle, 24);
    if (rs) parts.push(`[RS:${rs}]`);
  }
  // Co-located footnote + endnote markers share one paragraph, but the
  // aggregate above keeps the first-seen run style only, which hid the
  // endnote marker. Emit the missing partner so inspect never reads
  // footnote-only when both markers are present.
  {
    const notes = collectNoteReferenceStyles(xml);
    if (notes.footnote && notes.endnote) {
      for (const name of ['FootnoteReference', 'EndnoteReference'] as const) {
        if (run.rStyle === name) continue;
        const clean = sanitizeToken(name, 24);
        if (clean) parts.push(`[RS:${clean}]`);
      }
    }
  }

  const para = extractParaFmt(xml);
  if (para.jc) {
    const j = sanitizeToken(para.jc, 8);
    if (j) parts.push(`[J:${j}]`);
  }
  if (para.numId !== undefined || para.ilvl !== undefined) {
    const nid = para.numId !== undefined ? sanitizeToken(para.numId, 8) : '';
    // w:ilvl counts from 0; the level shown counts from 1, as `level` does when typed.
    const ilvlNumber = para.ilvl !== undefined ? Number.parseInt(para.ilvl, 10) : Number.NaN;
    const lvl = para.ilvl === undefined
      ? ''
      : (Number.isInteger(ilvlNumber) && ilvlNumber >= 0 ? String(ilvlNumber + 1) : sanitizeToken(para.ilvl, 2));
    const token = nid && lvl ? `${nid}/${lvl}` : (nid || `L${lvl}`);
    if (token) parts.push(`[N:${token}]`);
  } else if (node.metadata.listLevel !== undefined) {
    parts.push(`[N:L${node.metadata.listLevel + 1}]`);
  }

  if (node.type === 'table') {
    const look = extractTableLook(xml);
    if (look.firstRow) parts.push('[HDR]');
    if (look.zebra) parts.push('[ZEBRA]');
  }

  const isChartNode = node.type === 'chart' || node.type === 'chart3d';
  if (isChartNode) {
    // H2: never fabricate a type. Missing/empty chartType — or a synthetic
    // (placeholder-free) hydration — emits [CH:?] (explicit unknown).
    const cd = node.metadata.chartData as { synthetic?: unknown } | null | undefined;
    const isSynthetic = !!cd && typeof cd === 'object' && cd.synthetic === true;
    const rawCt: unknown = node.metadata.chartType;
    const clean = !isSynthetic && typeof rawCt === 'string' && rawCt.length > 0
      ? sanitizeToken(rawCt, 24)
      : '';
    parts.push(clean ? `[CH:${clean}]` : '[CH:?]');
    const rId = extractChartRId(xml);
    if (rId) {
      const clean = sanitizeToken(rId, 16);
      if (clean) parts.push(`[RID:${clean}]`);
    }
  }

  return parts.length > 0 ? ' ' + parts.join(' ') : '';
}

/** Full-tier detail: raw w:pPr + w:rPr serialization for one focused node. Throws with remedy on serializer failure. */
function buildFullDetail(xml: Element): string {
  const serializer = new XMLSerializer();
  const chunks: string[] = [];
  try {
    const pPrs = xml.getElementsByTagName('w:pPr');
    for (let i = 0; i < pPrs.length; i++) {
      try {
        chunks.push(`pPr:${serializer.serializeToString(pPrs[i] as Element)}`);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw new Error(
          `TOON full-tier pPr serialization failed: ${msg}. ` +
            `Remedy: call toon({ fidelity: 'standard' }) for the enriched outline instead, ` +
            `or verify the node XML is well-formed via inspect_document.`
        );
      }
    }
    const rPrs = xml.getElementsByTagName('w:rPr');
    for (let i = 0; i < rPrs.length; i++) {
      try {
        chunks.push(`rPr${i}:${serializer.serializeToString(rPrs[i] as Element)}`);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw new Error(
          `TOON full-tier rPr serialization failed: ${msg}. ` +
            `Remedy: call toon({ fidelity: 'standard' }) for the enriched outline instead, ` +
            `or verify the node XML is well-formed via inspect_document.`
        );
      }
    }
  } catch (err: unknown) {
    if (err instanceof Error && err.message.startsWith('TOON full-tier')) throw err;
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(
      `TOON full-tier detail build failed: ${msg}. ` +
        `Remedy: retry with toon({ fidelity: 'standard' }) or re-read the section via inspect_document.`
    );
  }
  return chunks.join('|');
}

/** Short fingerprint of the standard-tier fmt signals so style-only edits invalidate the cache. */
function fingerprintFmt(node: Node): string {
  try {
    const xml = node.xml;
    if (!xml || typeof xml.getElementsByTagName !== 'function') return 'fmt:-';
    const run = extractRunFmt(xml);
    const para = extractParaFmt(xml);
    const tbl = node.type === 'table' ? extractTableLook(xml) : { firstRow: false, zebra: false };
    const chartR = node.type === 'chart' || node.type === 'chart3d' ? (extractChartRId(xml) ?? '-') : '-';
    // H2: mirror buildStandardSuffix — empty chartType or synthetic hydration
    // normalizes to '-', never a fabricated default.
    const rawChartT: unknown = node.metadata.chartType;
    const cdSyn = (node.metadata.chartData as { synthetic?: unknown } | null | undefined);
    const isSyn = !!cdSyn && typeof cdSyn === 'object' && cdSyn.synthetic === true;
    const chartT = !isSyn && typeof rawChartT === 'string' && rawChartT.length > 0 ? rawChartT : '-';
    // Mirror the co-located note-partner emission in buildStandardSuffix so
    // footnote-only and footnote+endnote paragraphs never share a fingerprint.
    let noteSig = '';
    try {
      const notes = collectNoteReferenceStyles(xml);
      if (notes.footnote && notes.endnote) noteSig = ':fn+en';
    } catch {
      noteSig = '';
    }
    return (
      `fmt:${run.bold ? 1 : 0}${run.italic ? 1 : 0}:` +
      `${run.color ?? '-'}/${run.sz ?? '-'}/${run.rStyle ?? '-'}${noteSig}:` +
      `${para.jc ?? '-'}/${para.numId ?? '-'}/${para.ilvl ?? '-'}:` +
      `tbl${tbl.firstRow ? 1 : 0}${tbl.zebra ? 1 : 0}:ch${chartT}/${chartR}`
    );
  } catch {
    return 'fmt:?';
  }
}

export class ToonGenerator {
  private cache: Map<string, ToonTable> = new Map();
  private checksumCache: Map<string, string> = new Map();

  /**
   * BX-08: cache discipline.
   * - `revision` bumps on every recompute; consumers can poll getRevision()
   *   instead of rebuilding O(N) checksum strings to detect change.
   * - The map is bounded (LRU eviction, oldest insertion first) so long
   *   section-scoped sessions cannot grow it without limit. The 'latest'
   *   alias entry is refreshed in place and never evicted.
   */
  private revision = 0;
  private static readonly MAX_CACHE_ENTRIES = 20;

  /** Monotonic recompute counter (BX-08 revision). */
  public getRevision(): number {
    return this.revision;
  }

  /** Drops all cached tables and checksums (forces recompute on next call). */
  public invalidateCache(sectionId?: string): void {
    if (sectionId) {
      const prefixMatch: string[] = [];
      for (const key of this.cache.keys()) {
        if (key === `sec_${sectionId}` || key.startsWith(`sec_${sectionId}:`) || key.startsWith(`sec_${sectionId}|`)) {
          prefixMatch.push(key);
        }
      }
      if (prefixMatch.length > 0) {
        for (const key of prefixMatch) {
          this.cache.delete(key);
          this.checksumCache.delete(key);
        }
        return;
      }
      this.cache.delete(`sec_${sectionId}`);
      this.checksumCache.delete(`sec_${sectionId}`);
      return;
    }
    this.cache.clear();
    this.checksumCache.clear();
  }

  /** Current cache footprint for tests/telemetry. */
  public getCacheStats(): { entries: number; revision: number; maxEntries: number } {
    return { entries: this.cache.size, revision: this.revision, maxEntries: ToonGenerator.MAX_CACHE_ENTRIES };
  }

  /**
   * H1: mirror generateToon's sectionsToProcess resolution.
   * generateToon falls back to node-scoped filtering
   * (n.id === sectionId || n.id.startsWith(sectionId)) when no section id
   * matches. The checksum must see the same rows or a node-scoped call
   * computes an empty fingerprint and serves stale TOON after p_001 edits.
   */
  private resolveSectionsForChecksum(ast: DocumentAST, sectionId?: string): Section[] {
    if (!sectionId) return ast.sections;
    if (ast.sections.some((s) => s.id === sectionId)) {
      return ast.sections.filter((s) => s.id === sectionId);
    }
    return ast.sections
      .map((s) => ({
        ...s,
        nodes: s.nodes.filter((n) => n.id === sectionId || n.id.startsWith(sectionId)),
      }))
      .filter((s) => s.nodes.length > 0);
  }

  private storeCached(cacheKey: string, table: ToonTable, checksum: string): void {
    this.cache.set(cacheKey, table);
    this.checksumCache.set(cacheKey, checksum);
    this.cache.set('latest', table);
    this.revision += 1;
    // Bounded map: evict oldest section/full-doc entries first ('latest' is
    // an alias and is never a victim). Map preserves insertion order, so the
    // first non-'latest' key is the LRU victim; re-setting an existing key
    // does not refresh its age — acceptable for a render cache.
    while (this.cache.size > ToonGenerator.MAX_CACHE_ENTRIES + 1) {
      for (const key of this.cache.keys()) {
        if (key !== 'latest') {
          this.cache.delete(key);
          this.checksumCache.delete(key);
          break;
        }
      }
    }
  }

  /**
   * Track-B entry point: toon({ fidelity, sectionId, maxTextLength, useCache }).
   * Default fidelity is 'standard'. Full tier requires a sectionId (throws with
   * remedy otherwise) and throws with remedy when the sectionId matches nothing.
   */
  public toon(ast: DocumentAST, req: ToonRequest = {}): ToonTable {
    const fidelity: ToonFidelity = req.fidelity ?? TOON_DEFAULT_FIDELITY;
    assertToonFidelity(fidelity);
    return this.generateToon(ast, req.sectionId, {
      sectionId: req.sectionId,
      maxTextLength: req.maxTextLength,
      useCache: req.useCache,
      fidelity,
    });
  }

  // 2.1 & 2.2 - Generate TOON and Extract Metadata Flags
  public generateToon(ast: DocumentAST, sectionId?: string, options?: ToonGenerateOptions): ToonTable {
    const fidelity: ToonFidelity = options?.fidelity ?? TOON_DEFAULT_FIDELITY;
    assertToonFidelity(fidelity);
    const resolvedSectionId = options?.sectionId ?? sectionId;

    if (fidelity === 'full' && !resolvedSectionId) {
      throw new Error(
        `TOON full fidelity requires a focused sectionId (full rPr/pPr dumps are section-scoped by design). ` +
          `Remedy: call toon({ fidelity: 'full', sectionId: '<sectionId>' }) with a section id from ` +
          `toon({ fidelity: 'standard' }), or use fidelity 'standard' for the whole document.`
      );
    }

    const cacheKey = resolvedSectionId ? `sec_${resolvedSectionId}` : 'full_doc';
    // Render-affecting options participate in the key: a prior call with a
    // different maxTextLength or fidelity must not serve its rows here.
    const keySuffix = options?.maxTextLength ? `:len${options.maxTextLength}` : '';
    const effectiveKey = `${cacheKey}${keySuffix}|tier:${fidelity}`;
    const currentChecksum = this.calculateAstChecksum(ast, resolvedSectionId, fidelity);

    if (options?.useCache !== false && this.checksumCache.get(effectiveKey) === currentChecksum && this.cache.has(effectiveKey)) {
      const hit = this.cache.get(effectiveKey);
      if (hit) return hit;
      throw new Error(
        `TOON cache inconsistency for key '${effectiveKey}'. ` +
          `Remedy: call invalidateCache(${resolvedSectionId ? `'${resolvedSectionId}'` : 'no args'}) and retry toon().`
      );
    }

    const rows: ToonRow[] = [];

    // 2.3 Section-Scoped Generation
    const sectionsToProcess = resolvedSectionId
      ? (ast.sections.some(s => s.id === resolvedSectionId)
          ? ast.sections.filter(s => s.id === resolvedSectionId)
          : ast.sections
              .map(s => ({ ...s, nodes: s.nodes.filter(n => n.id === resolvedSectionId || n.id.startsWith(resolvedSectionId)) }))
              .filter(s => s.nodes.length > 0))
      : ast.sections;

    if (fidelity === 'full' && resolvedSectionId && sectionsToProcess.length === 0) {
      const available = ast.sections.map(s => s.id).join(', ') || '(no sections)';
      throw new Error(
        `TOON full fidelity: sectionId '${resolvedSectionId}' matched no section or node. ` +
          `Available section ids: ${available}. ` +
          `Remedy: call toon({ fidelity: 'standard' }) to list sections, then retry full with a valid sectionId.`
      );
    }

    // Build section path map for breadcrumbs
    const sectionPaths = this.buildSectionPaths(ast.sections);

    const maxLen = options?.maxTextLength;

    for (const section of sectionsToProcess) {
      const sectionPath = sectionPaths.get(section.id) || 'Root';
      const sectionHeader = section.heading?.text || '';

      for (const node of section.nodes) {
        let rawText = node.metadata.text.trim() || this.generatePlaceholderText(node);
        // Grid readability: table rows render delimited cells with coordinates
        // (r1c1 value | r1c2 value), spans marked (span2), merge covers tagged
        // hidden-cover — every tier, compact included, stays terse by design.
        if (node.type === 'table') {
          try {
            const xml = node.xml;
            if (xml && typeof xml.getElementsByTagName === 'function') {
              const grid = extractTableGrid(xml);
              if (grid.cells.length > 0) rawText = renderTableGrid(grid);
            }
          } catch {
            /* keep the metadata/placeholder fallback */
          }
        } else if (node.type === 'chart' || node.type === 'chart3d') {
          // Chart visibility: title + categories + series surface read-only
          // from the parser-hydrated chart part (never fabricated).
          try {
            rawText = renderChartText(node);
          } catch {
            /* keep the metadata/placeholder fallback */
          }
        }
        if (maxLen && maxLen > 0 && rawText.length > maxLen) {
          rawText = rawText.substring(0, maxLen) + '...';
        }

        const baseFlags = this.generateFlags(node);
        const suffix = fidelity === 'compact' ? '' : buildStandardSuffix(node);
        const flags = `${baseFlags}${suffix}`.trim();
        const detail = fidelity === 'full' ? buildFullDetail(node.xml) : undefined;

        const row: ToonRow = {
          id: node.id,
          version: node.version,
          path: sectionPath,
          header: sectionHeader,
          style: this.generateStyleInfo(node, section, ast),
          text: rawText,
          flags,
        };
        if (detail !== undefined) row.detail = detail;
        rows.push(row);
      }
    }

    const headers = fidelity === 'full'
      ? ['ID', 'V', 'Path', 'Header', 'Style', 'Text', 'Flags', 'Detail']
      : ['ID', 'V', 'Path', 'Header', 'Style', 'Text', 'Flags'];
    const table: ToonTable = {
      headers,
      rows,
      fidelity: 'lossy',
      tier: fidelity,
    };

    // 2.4 Caching
    this.storeCached(effectiveKey, table, currentChecksum);

    return table;
  }

  private calculateAstChecksum(ast: DocumentAST, sectionId?: string, fidelity: ToonFidelity = TOON_DEFAULT_FIDELITY): string {
    // BX-08: the fingerprint must see in-place DOM/metadata edits even when
    // the caller forgot the version bump. Besides id:version:checksum it folds
    // in cached text length + head, live DOM text length + head (when the node
    // still points at live XML), and the flag/style inputs — so a row whose
    // rendered content changed can never hit a stale cache entry. Kept O(N)
    // with small constants (lengths + short prefixes, no full serialization).
    // BX-08 follow-up (D1 brutal review): section.heading.text is rendered
    // (Header column + breadcrumb path) but was never folded — a heading-only
    // edit served stale TOON. Fold heading identity/text/live-DOM plus cheap
    // global revision signals (settings, headers/footers, numbering) so those
    // edits invalidate instead of hitting a stale entry. Globals are folded
    // always (even for section-scoped keys): over-invalidation is safe, stale
    // is not. All signals are O(1)/O(small-map) — no full serialization.
    // Track-B: standard/full tiers render run/paragraph/table/chart signals,
    // so the fmt fingerprint is folded whenever fidelity !== compact. Compact
    // keeps the legacy fingerprint byte-identical (no over-invalidation).
    // H1: sectionId may be a nodeId fallback (see resolveSectionsForChecksum)
    // — checksum must cover the same rows generateToon renders.
    let combined = `tier:${fidelity};`;
    const sections = this.resolveSectionsForChecksum(ast, sectionId);
    for (const sec of sections) {
      combined += `${sec.id}:${sec.version}:${sec.checksum};`;
      // Heading signal: rendered as Header column + breadcrumb path + style
      // token (headingL). Fold id/version/checksum/level plus cached text and
      // live DOM text so detached-heading edits cannot slip through.
      const h: HeadingNode | null = sec.heading;
      if (h) {
        let hLiveHead = '';
        let hLiveLen = -1;
        try {
          const live: unknown = h.xml?.textContent;
          if (typeof live === 'string') {
            hLiveLen = live.length;
            hLiveHead = live.slice(0, 48);
          }
        } catch {
          hLiveLen = -1;
          hLiveHead = '';
        }
        combined += `h:${h.id ?? '-'}:${h.version ?? '-'}:${h.checksum ?? '-'}:` +
          `lv${h.level ?? '-'}:t${String(h.text ?? '').length}:${String(h.text ?? '').slice(0, 48)}:` +
          `live${hLiveLen}:${hLiveHead};`;
      } else {
        combined += 'h:-;';
      }
      // Cheap pageSetup signal (orientation/paperSize/margins feed pagination).
      try {
        const ps = sec.pageSetup;
        if (ps && typeof ps === 'object') {
          combined += `ps:${ps.orientation ?? '-'}:${ps.paperSize ?? '-'}:` +
            `${ps.marginTop ?? '-'}:${ps.marginBottom ?? '-'}:` +
            `${ps.marginLeft ?? '-'}:${ps.marginRight ?? '-'};`;
        }
      } catch {
        combined += 'ps:?;';
      }
      for (const n of sec.nodes) {
        let liveHead = '';
        let liveLen = -1;
        try {
          const live: unknown = n.xml?.textContent;
          if (typeof live === 'string') {
            liveLen = live.length;
            liveHead = live.slice(0, 48);
          }
        } catch {
          liveLen = -1;
          liveHead = '';
        }
        const meta = n.metadata;
        combined += `${n.id}:${n.version}:${n.checksum}:` +
          `t${(meta.text ?? '').length}:${String(meta.text ?? '').slice(0, 48)}:` +
          `live${liveLen}:${liveHead}:` +
          `ll${meta.listLevel ?? '-'}:ns${meta.namedStyle ?? '-'}:` +
          `f${meta.hasImage ? 1 : 0}${meta.hasEquation ? 1 : 0}` +
          `${meta.hasComment ? 1 : 0}${meta.hasHyperlink ? 1 : 0}` +
          `${meta.hasBookmark ? 1 : 0}${meta.hasTrackedChanges ? 1 : 0};`;
        if (fidelity !== 'compact') {
          combined += `${fingerprintFmt(n)};`;
        }
      }
    }
    // Global revision signals: cheap counts + short heads, never full XML.
    try {
      const s = ast.settings;
      if (s) combined += `settings:${s.trackChanges ?? '-'}:${s.defaultAuthor ?? '-'}:${ast.nextChangeId ?? '-'};`;
    } catch {
      combined += 'settings:?;';
    }
    try {
      const hdr: Record<string, { type?: string; text?: string }> = ast.headers as Record<string, { type?: string; text?: string }>;
      const ftr: Record<string, { type?: string; text?: string }> = ast.footers as Record<string, { type?: string; text?: string }>;
      combined += `hdr:${Object.keys(hdr).length};`;
      for (const [k, v] of Object.entries(hdr)) {
        combined += `${k}:${v?.type ?? '-'}:t${String(v?.text ?? '').length}:${String(v?.text ?? '').slice(0, 32)};`;
      }
      combined += `ftr:${Object.keys(ftr).length};`;
      for (const [k, v] of Object.entries(ftr)) {
        combined += `${k}:${v?.type ?? '-'}:t${String(v?.text ?? '').length}:${String(v?.text ?? '').slice(0, 32)};`;
      }
    } catch {
      combined += 'hdr:?;ftr:?;';
    }
    try {
      const absIds = Object.keys(ast.numbering?.abstractNum ?? {});
      const conIds = Object.keys(ast.numbering?.concreteNum ?? {});
      // Sorted id lists: numbering maps are tiny (<100), join stays cheap and
      // distinguishes add/remove/renumber (counts alone would miss swaps).
      combined += `num:${absIds.length}:${conIds.length}:${[...absIds].sort().join(',')}:${[...conIds].sort().join(',')};`;
    } catch {
      combined += 'num:?;';
    }
    return combined;
  }

  private buildSectionPaths(sections: Section[]): Map<string, string> {
    const paths = new Map<string, string>();
    const hierarchy: { section: Section; depth: number }[] = [];

    for (const section of sections) {
      if (!section.heading) {
        paths.set(section.id, 'Root');
        continue;
      }

      // Determine hierarchy depth based on heading level
      const level = section.heading.level;

      // Pop sections deeper than current level
      while (hierarchy.length > 0) {
        const top = hierarchy[hierarchy.length - 1];
        const topLevel = top.section.heading ? top.section.heading.level : 0;
        if (topLevel < level) break;
        hierarchy.pop();
      }

      // Build breadcrumb path
      const parts: string[] = [];
      for (const h of hierarchy) {
        const t = h.section.heading ? h.section.heading.text : '';
        if (t) parts.push(t);
      }
      const breadcrumb = parts.length === 0
        ? section.heading.text
        : [...parts, section.heading.text].join(' > ');

      paths.set(section.id, breadcrumb);
      hierarchy.push({ section, depth: level });
    }

    return paths;
  }

  private generateStyleInfo(node: Node, section: Section, ast: DocumentAST): string {
    // Check if this node is the section heading
    if (section.heading && node.id === section.heading.id) {
      return `heading${section.heading.level}:L${section.heading.level - 1}`;
    }

    // Named style (e.g., "Heading 2", "Quote", "Code")
    if (node.metadata.namedStyle) {
      return node.metadata.namedStyle;
    }

    // List item: session metadata when the list op set it, else the live XML
    // (numbered vs bullet resolved through the numbering maps) so reloaded
    // documents label truthfully instead of falling back to body/bullet.
    if (node.metadata.listLevel !== undefined || hasListNumbering(node)) {
      let level = node.metadata.listLevel;
      if (level === undefined) {
        level = 0;
        try {
          const numPr = node.xml.getElementsByTagName('w:numPr')[0] as Element | undefined;
          const v = numPr?.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val');
          if (v !== undefined && v !== null && v !== '') {
            const n = parseInt(String(v), 10);
            if (Number.isFinite(n)) level = Math.max(0, Math.min(8, n));
          }
        } catch { /* level stays 0 */ }
      }
      const listStyle = node.metadata.listStyle
        || (isOrderedListItem(node, ast) ? 'numbered' : 'bullet');
      return `list:${listStyle}:L${level + 1}`; // levels count from 1, like the `level` parameter
    }

    // Chart nodes carry a drawing (hasImage) but are charts first: `chart`/`chart3d`
    // style tokens keep them distinguishable from plain images in TOON.
    if (node.type === 'chart') {
      return 'chart';
    }
    if (node.type === 'chart3d') {
      return 'chart3d';
    }

    // Table: prefer live grid dims from the XML (getAST metadata carries no
    // tableDimensions); fall back to metadata, then to a bare token.
    if (node.type === 'table') {
      try {
        const xml = node.xml;
        if (xml && typeof xml.getElementsByTagName === 'function') {
          const grid = extractTableGrid(xml);
          if (grid.rows > 0 && grid.cols > 0) return `table:${grid.rows}x${grid.cols}`;
        }
      } catch {
        /* fall through to metadata */
      }
      if (node.metadata.tableDimensions) {
        const { rows, cols } = node.metadata.tableDimensions;
        return `table:${rows}x${cols}`;
      }
      return 'table';
    }

    // Image
    if (node.type === 'image' || node.metadata.hasImage) {
      return 'image';
    }

    // Equation
    if (node.type === 'equation' || node.metadata.hasEquation) {
      return 'equation';
    }

    // Default body text
    return 'body:L-1';
  }

  private generatePlaceholderText(node: Node): string {
    // Generate descriptive placeholder for non-text nodes
    if (node.type === 'image' || node.metadata.hasImage) {
      return '[Image]';
    }
    if (node.type === 'chart3d') {
      return '[3D Chart]';
    }
    if (node.type === 'equation' || node.metadata.hasEquation) {
      return '[Equation]';
    }
    if (node.type === 'table') {
      const dims = node.metadata.tableDimensions;
      return dims ? `[Table ${dims.rows}x${dims.cols}]` : '[Table]';
    }
    if (node.metadata.hasField) {
      return '[Field]';
    }

    // Inspect XML for page breaks and charts if it's an empty paragraph
    if (node.type === 'paragraph' && (!node.metadata.text || node.metadata.text.trim() === '')) {
      const xml = node.xml;
      if (xml && typeof xml.getElementsByTagName === 'function') {
        try {
          if (xml.getElementsByTagName('w:drawing').length > 0) {
            if (xml.getElementsByTagName('c:chart').length > 0) {
              return '[3D Chart]';
            }
            return '[Image]';
          }
          const brs = xml.getElementsByTagName('w:br');
          for (let b = 0; b < brs.length; b++) {
            try {
              if ((brs[b] as Element).getAttribute('w:type') === 'page') {
                return '[Page Break]';
              }
            } catch {
              continue;
            }
          }
        } catch {
          return '';
        }
      }
    }
    return '';
  }

  private generateFlags(node: Node): string {
    const flags: string[] = [];
    if (node.type === 'table') flags.push('[TBL]');
    if (node.metadata.hasImage) flags.push('[IMG]');
    if (node.metadata.hasEquation) flags.push('[EQ]');
    if (node.metadata.hasCitation) flags.push('[CITE]');
    if (node.metadata.hasComment) flags.push('[CMT]');
    if (node.metadata.hasTrackedChanges) flags.push('[TRK]');
    if (node.metadata.hasHyperlink) flags.push('[LNK]');
    if (node.metadata.hasBookmark) flags.push('[BKMK]');
    if (node.metadata.listLevel !== undefined) flags.push(`[L${node.metadata.listLevel + 1}]`);
    return flags.join(' ');
  }

  public renderToonText(table: ToonTable, options?: ToonRenderOptions): string {
    const omits = new Set<string>(options?.omitColumns ?? []);
    const compact = options?.compact ?? false;
    const tier: ToonFidelity = table.tier ?? 'compact';
    const wantsDetail = tier === 'full' && (options?.includeDetail ?? true) && !omits.has('detail');

    // Filter headers based on options
    const activeHeaders: string[] = [];
    type ColKey = 'id' | 'version' | 'path' | 'header' | 'style' | 'text' | 'flags' | 'detail';
    const colKeys: ColKey[] = wantsDetail
      ? ['id', 'version', 'path', 'header', 'style', 'text', 'flags', 'detail']
      : ['id', 'version', 'path', 'header', 'style', 'text', 'flags'];
    const fallbackLabels = wantsDetail
      ? ['ID', 'V', 'Path', 'Header', 'Style', 'Text', 'Flags', 'Detail']
      : ['ID', 'V', 'Path', 'Header', 'Style', 'Text', 'Flags'];
    const headerLabels = table.headers.length === colKeys.length
      ? table.headers
      : fallbackLabels;

    colKeys.forEach((key, index) => {
      if (key === 'id' || key === 'version' || key === 'text' || key === 'detail') {
        if (key === 'detail' && (omits.has('detail') || !wantsDetail)) return;
        activeHeaders.push(headerLabels[index] || key.toUpperCase());
      } else if (!omits.has(key)) {
        activeHeaders.push(headerLabels[index] || key.toUpperCase());
      }
    });

    // ST-11 G1b: surface fidelity in the TEXT output (dead-flag fix) — the
    // structured table carries fidelity:'lossy' but agents only read text.
    // First line is machine-readable so prompts/tests can gate on it.
    // Track-B appends the tier (compact|standard|full) after the lossy flag;
    // the '> fidelity: lossy' prefix is preserved for backward compat.
    const fidelity = table.fidelity || TOON_FIDELITY_ALLOWLIST.fidelity || 'lossy';
    let output = `> fidelity: ${fidelity} tier: ${tier}\n`;

    if (options?.includeHeaderRow !== false) {
      const sep = compact ? '|' : ' | ';
      output += `${activeHeaders.join(sep)}\n`;
      if (!compact) {
        output += `${activeHeaders.map(() => '---').join(' | ')}\n`;
      }
    }

    for (const row of table.rows) {
      let textVal = row.text;
      if (options?.maxTextLength && options.maxTextLength > 0 && textVal.length > options.maxTextLength) {
        textVal = textVal.substring(0, options.maxTextLength) + '...';
      }

      const rowParts: string[] = [];
      rowParts.push(row.id);
      rowParts.push(row.version.toString());
      if (!omits.has('path')) rowParts.push(row.path);
      if (!omits.has('header')) rowParts.push(row.header);
      if (!omits.has('style')) rowParts.push(row.style);
      rowParts.push(textVal);
      if (!omits.has('flags')) rowParts.push(row.flags);
      if (wantsDetail) {
        if (omits.has('detail')) {
          // Explicitly omitted — skip.
        } else {
          rowParts.push(row.detail ?? '');
        }
      }

      const sep = compact ? '|' : ' | ';
      const rawWidth = options?.width;
      const width =
        typeof rawWidth === 'number' && Number.isFinite(rawWidth) && rawWidth >= 1
          ? Math.floor(rawWidth)
          : undefined;
      if (width === undefined) {
        output += `${rowParts.join(sep)}\n`;
        continue;
      }
      // Text is the cell after id, version, and whichever of
      // path/header/style are present. Only it may wrap: the prefix (all
      // columns before Text, trailing separator included) and the suffix
      // (flags/detail, leading separator included) stay whole. The suffix
      // rides on the LAST line so folding continuations back recovers the
      // default line byte-identically; a row that fits emits exactly the
      // default line.
      const textPos =
        2 + (omits.has('path') ? 0 : 1) + (omits.has('header') ? 0 : 1) + (omits.has('style') ? 0 : 1);
      const prefix = `${rowParts.slice(0, textPos).join(sep)}${sep}`;
      const tail = rowParts.slice(textPos + 1);
      const suffix = tail.length > 0 ? `${sep}${tail.join(sep)}` : '';
      const prefixLen = displayLen(prefix);
      // The last line carries prefix indent + segment + suffix, so it has
      // the tightest budget; wrapping every segment to it keeps all lines
      // within width. Non-positive budget (tiny width or a prefix that
      // already overflows) falls back to the default line — other columns
      // never wrap, even at the cost of exceeding width.
      const budget = width - prefixLen - displayLen(suffix);
      if (budget <= 0) {
        output += `${rowParts.join(sep)}\n`;
        continue;
      }
      const segments = wrapTextCell(textVal, budget);
      if (segments.length <= 1) {
        output += `${rowParts.join(sep)}\n`;
        continue;
      }
      const indent = ' '.repeat(prefixLen);
      output += `${prefix}${segments[0]}\n`;
      for (let k = 1; k < segments.length - 1; k++) {
        output += `${indent}${segments[k]}\n`;
      }
      output += `${indent}${segments[segments.length - 1]}${suffix}\n`;
    }

    return output;
  }

  // 2.5 TOON Diffs
  public generateDiff(oldTable: ToonTable, newTable: ToonTable, options?: ToonDiffOptions): ToonTable {
    const diffRows: ToonRow[] = [];
    const oldMap = new Map(oldTable.rows.map(r => [r.id, r]));

    for (const newRow of newTable.rows) {
      const oldRow = oldMap.get(newRow.id);
      if (!oldRow) {
        diffRows.push({
          ...newRow,
          flags: `${newRow.flags} [+ADDED]`.trim()
        });
      } else if (oldRow.version !== newRow.version || oldRow.text !== newRow.text || oldRow.flags !== newRow.flags) {
        diffRows.push({
          ...newRow,
          flags: `${newRow.flags} [~MODIFIED]`.trim()
        });
      } else if (options?.showUnchangedStatus) {
        diffRows.push(newRow);
      }
    }

    return {
      headers: ['ID', 'V', 'Path', 'Header', 'Style', 'Text (CHANGED)', 'Flags'],
      rows: diffRows,
      fidelity: 'lossy',
      tier: newTable.tier ?? oldTable.tier ?? 'compact',
    };
  }

  // 2.6 Token Measurement
  public estimateTokens(text: string): number {
    if (!text) return 0;
    // Subword estimation heuristic: ~3.8 characters per token for typical JSON/markdown/XML structured text
    return Math.ceil(text.length / 3.8);
  }

  public estimateTableTokens(table: ToonTable, options?: ToonRenderOptions): number {
    const rendered = this.renderToonText(table, options);
    return this.estimateTokens(rendered);
  }

  /**
   * Generates a TOON representation fitting within a target maxTokens budget.
   * Compression ladder first (whole-document coverage): text 120 → 40 →
   * compact (omitting path/flags columns). When still over budget, slices on
   * clean row boundaries (binary-search largest fitting prefix, fit measured
   * with the same estimateTokens returned in tokenCount) with cursor paging
   * via options.offset. A single row bigger than budget is still returned
   * with truncated set so callers never page forever.
   * Fidelity-aware: defaults to standard; full requires sectionId (throws otherwise).
   */
  public generateTokenBudgetToon(
    ast: DocumentAST,
    maxTokens: number,
    options?: ToonGenerateOptions & ToonRenderOptions & { offset?: number }
  ): { table: ToonTable; rendered: string; tokenCount: number; truncated: boolean; hasMore: boolean; nextOffset: number | null } {
    if (!Number.isFinite(maxTokens) || maxTokens <= 0) {
      throw new Error(
        `generateTokenBudgetToon requires a positive maxTokens (got '${String(maxTokens)}'). ` +
          `Remedy: pass a token budget >= 40, e.g. generateTokenBudgetToon(ast, 800).`
      );
    }
    const fidelity: ToonFidelity = options?.fidelity ?? TOON_DEFAULT_FIDELITY;
    assertToonFidelity(fidelity);
    if (fidelity === 'full' && !(options?.sectionId)) {
      throw new Error(
        `generateTokenBudgetToon with fidelity 'full' requires options.sectionId. ` +
          `Remedy: pass { fidelity: 'full', sectionId: '<sectionId>' } or use fidelity 'standard'.`
      );
    }
    let textLimit = options?.maxTextLength || 120;
    let table = this.generateToon(ast, options?.sectionId, { ...options, fidelity, maxTextLength: textLimit });
    const renderOpts: ToonRenderOptions = { ...options, maxTextLength: textLimit };
    let rendered = this.renderToonText(table, renderOpts);
    let tokenCount = this.estimateTokens(rendered);

    // Iteratively compress if exceeding token budget
    if (tokenCount > maxTokens && textLimit > 40) {
      textLimit = 40;
      renderOpts.maxTextLength = textLimit;
      table = this.generateToon(ast, options?.sectionId, { ...options, fidelity, maxTextLength: textLimit });
      rendered = this.renderToonText(table, renderOpts);
      tokenCount = this.estimateTokens(rendered);
    }

    if (tokenCount > maxTokens) {
      renderOpts.compact = true;
      renderOpts.omitColumns = ['path', 'flags'];
      rendered = this.renderToonText(table, renderOpts);
      tokenCount = this.estimateTokens(rendered);
    }

    const totalRows = table.rows.length;
    const rawOffset = options?.offset ?? 0;
    const start = Number.isFinite(rawOffset)
      ? Math.max(0, Math.min(Math.floor(rawOffset), totalRows))
      : 0;

    // Fast path: whole document from offset 0 fits — byte-identical to the
    // pre-cursor output (same table object, same rendered text).
    if (start === 0 && tokenCount <= maxTokens) {
      return { table, rendered, tokenCount, truncated: false, hasMore: false, nextOffset: null };
    }

    // Row-boundary slicing backstop (mirrors the MCP inspect maxTokens cut):
    // whole rows only — dropping trailing rows never mid-row-truncates.
    // Floor of 1 row: a single row over budget is still returned with
    // truncated set, and nextOffset always advances, so paging terminates.
    let rows = table.rows.slice(start);
    let truncated = false;
    const renderRows = (n: number): string =>
      this.renderToonText(
        { headers: table.headers, rows: rows.slice(0, n), fidelity: 'lossy', tier: table.tier },
        renderOpts
      );
    if (rows.length > 0 && this.estimateTokens(renderRows(rows.length)) > maxTokens) {
      // Binary-search the largest fitting prefix; keep >= 1 row so the
      // caller always sees the window head + truncation flags.
      let lo = 1;
      let hi = rows.length - 1;
      let best = 0;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (this.estimateTokens(renderRows(mid)) <= maxTokens) {
          best = mid;
          lo = mid + 1;
        } else {
          hi = mid - 1;
        }
      }
      if (best < 1) best = 1;
      if (best < rows.length) {
        rows = rows.slice(0, best);
        truncated = true;
      }
    }
    // Single-row over-budget: the search above cannot run when the window
    // holds 1 row (hi = 0), so flag truncation when it still exceeds budget.
    if (rows.length >= 1 && this.estimateTokens(renderRows(rows.length)) > maxTokens) {
      truncated = true;
    }
    const windowed: ToonTable = { headers: table.headers, rows, fidelity: 'lossy', tier: table.tier };
    rendered = this.renderToonText(windowed, renderOpts);
    tokenCount = this.estimateTokens(rendered);
    const hasMore = start + rows.length < totalRows;
    return { table: windowed, rendered, tokenCount, truncated, hasMore, nextOffset: hasMore ? start + rows.length : null };
  }
}
