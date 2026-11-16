/**
 * Whole-table width and row heights (0.5.3): `setTableWidth` and `setRowHeights`.
 *
 * Source of truth for widths is the table grid (`w:gridCol`, always numeric twips). `w:tblW` and `w:tcW` may be
 * `pct`, `dxa`, `auto` or `nil`; they are never converted into numbers, only rewritten from the grid. ASTra never
 * invents a width: when the grid or the container cannot be trusted the operation is refused with a remedy.
 */
import { createWordElement } from '../ooxml/ns';
import { ensureTableGrid, sortChildrenBySchema, TBLPR_CHILD_ORDER, TCPR_CHILD_ORDER, TRPR_CHILD_ORDER } from '../ooxml/schema-order';
import { toError } from '../errors';
import { usableAreaOf, EMU_PER_INCH } from './chart-size';
import type { PageSetupProperties } from '../ast';

const TWIPS_PER_INCH = 1440;
/** 22 in: the largest page and table width Word shows (also the limit for a chart). */
const MAX_TWIPS = 22 * TWIPS_PER_INCH;
/** A column narrower than this (1 pt) cannot hold anything; the operation is refused instead. */
const MIN_COLUMN_TWIPS = 20;
/** Word's default left and right cell margin. */
const DEFAULT_CELL_MARGIN_TWIPS = 108;

function directChildren(parent: Element, tagName: string): Element[] {
  const out: Element[] = [];
  for (let c = parent.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (c as Element).tagName === tagName) out.push(c as Element);
  }
  return out;
}

const fail = (message: string, remedy: string): never => {
  throw toError('E_VALIDATION', message, remedy);
};

function intAttr(el: Element | undefined, name: string): number | null {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined || !/^\d+$/.test(raw.trim())) return null;
  return Number(raw.trim());
}

/** The first number among the keys, looked up flat and inside params. */
type Pick = (key: string) => unknown;

export type TableWidthSpec =
  | { kind: 'dxa'; twips: number }
  | { kind: 'pct'; percent: number };

export interface TableWidthRequest {
  spec: TableWidthSpec;
  columns: 'scale' | 'keep';
}

/** Reads and validates the arguments of `setTableWidth` before anything is touched. */
export function parseTableWidth(pick: Pick): TableWidthRequest {
  const widthIn = pick('widthIn');
  const widthPercent = pick('widthPercent');
  const widthTwips = pick('widthTwips');
  const given = [
    widthIn !== undefined ? 'widthIn' : null,
    widthPercent !== undefined ? 'widthPercent' : null,
    widthTwips !== undefined ? 'widthTwips' : null,
  ].filter((k): k is string => k !== null);
  if (given.length === 0) {
    return fail(
      'setTableWidth needs a width',
      'pass exactly one of widthIn (inches), widthPercent (1 to 100, of the text width) or widthTwips, e.g. { widthIn: 5 }'
    );
  }
  if (given.length > 1) {
    return fail(
      `setTableWidth got more than one width (${given.join(', ')})`,
      'pass exactly one of widthIn, widthPercent or widthTwips'
    );
  }
  const columnsRaw = pick('columns');
  if (columnsRaw !== undefined && columnsRaw !== 'scale' && columnsRaw !== 'keep') {
    return fail(
      `setTableWidth columns must be "scale" or "keep" (got ${JSON.stringify(columnsRaw)})`,
      'use columns: "scale" (resize every column in proportion, the default) or columns: "keep" (change only the table width)'
    );
  }
  const columns = (columnsRaw ?? 'scale') as 'scale' | 'keep';
  const finite = (name: string, v: unknown): number => {
    if (typeof v !== 'number' || !Number.isFinite(v)) {
      return fail(`setTableWidth ${name} must be a finite number (got ${JSON.stringify(v)})`, `pass ${name} as a number`);
    }
    return v;
  };
  if (widthPercent !== undefined) {
    const percent = finite('widthPercent', widthPercent);
    if (percent < 1 || percent > 100) {
      return fail(`setTableWidth widthPercent ${percent} is outside 1 to 100`, 'pass widthPercent from 1 to 100 (percent of the text width)');
    }
    return { spec: { kind: 'pct', percent }, columns };
  }
  let twips: number;
  if (widthIn !== undefined) {
    twips = Math.round(finite('widthIn', widthIn) * TWIPS_PER_INCH);
  } else {
    const t = finite('widthTwips', widthTwips);
    if (!Number.isInteger(t)) return fail(`setTableWidth widthTwips must be a whole number (got ${t})`, 'pass widthTwips as a whole number of twips (1440 per inch)');
    twips = t;
  }
  if (twips < MIN_COLUMN_TWIPS || twips > MAX_TWIPS) {
    return fail(
      `setTableWidth width is outside the supported range (got ${twips} twips, ${Math.round((twips / TWIPS_PER_INCH) * 100) / 100} in)`,
      `use a width from ${MIN_COLUMN_TWIPS} to ${MAX_TWIPS} twips (up to 22 inches); a table wider than the page is clipped by Word`
    );
  }
  return { spec: { kind: 'dxa', twips }, columns };
}

interface RowLayout {
  tr: Element;
  cells: Array<{ tc: Element; start: number; span: number }>;
  before: number;
  after: number;
}

/** Walks every direct row; returns each cell's grid position or throws when the spans do not match `gridLength`. */
function layoutRows(rows: Element[], gridLength: number | null): { rows: RowLayout[]; widest: number } {
  const out: RowLayout[] = [];
  let widest = 0;
  rows.forEach((tr, ri) => {
    const trPr = directChildren(tr, 'w:trPr')[0];
    const before = trPr ? intAttr(directChildren(trPr, 'w:gridBefore')[0], 'w:val') ?? 0 : 0;
    const after = trPr ? intAttr(directChildren(trPr, 'w:gridAfter')[0], 'w:val') ?? 0 : 0;
    let p = before;
    const cells: RowLayout['cells'] = [];
    for (const tc of directChildren(tr, 'w:tc')) {
      const tcPr = directChildren(tc, 'w:tcPr')[0];
      const spanRaw = tcPr ? directChildren(tcPr, 'w:gridSpan')[0] : undefined;
      const span = spanRaw ? intAttr(spanRaw, 'w:val') : 1;
      if (span === null || span < 1) {
        fail(`setTableWidth: row ${ri + 1} has a cell with an invalid gridSpan`, 'repair the table or rebuild it, then retry');
      }
      cells.push({ tc, start: p, span: span as number });
      p += span as number;
    }
    p += after;
    widest = Math.max(widest, p);
    if (gridLength !== null && p !== gridLength) {
      fail(
        `setTableWidth: row ${ri + 1} spans ${p} grid columns but the table grid has ${gridLength}`,
        'the table is structurally inconsistent; fix the merged cells (mergeCells/splitCell) or rebuild the table, then retry'
      );
    }
    out.push({ tr, cells, before, after });
  });
  return { rows: out, widest };
}

/** The numeric grid, or null when it is missing, empty or has a non-numeric or zero entry. */
function readGrid(tbl: Element): number[] | null {
  const grid = directChildren(tbl, 'w:tblGrid')[0];
  if (!grid) return null;
  const cols = directChildren(grid, 'w:gridCol');
  if (cols.length === 0) return null;
  const widths: number[] = [];
  for (const col of cols) {
    const w = intAttr(col, 'w:w');
    if (w === null || w <= 0) return null;
    widths.push(w);
  }
  return widths;
}

/**
 * The grid, rebuilt from one row without merges whose cells all have a numeric dxa width; null when no row can
 * be trusted (never fabricated from `auto`, `nil` or `pct`).
 */
function rebuiltGrid(rows: Element[]): number[] | null {
  for (const tr of rows) {
    const cells = directChildren(tr, 'w:tc');
    if (cells.length === 0) continue;
    const widths: number[] = [];
    let ok = true;
    for (const tc of cells) {
      const tcPr = directChildren(tc, 'w:tcPr')[0];
      const merged = !!tcPr && (directChildren(tcPr, 'w:gridSpan').length > 0 || directChildren(tcPr, 'w:vMerge').length > 0 || directChildren(tcPr, 'w:hMerge').length > 0);
      const tcW = tcPr ? directChildren(tcPr, 'w:tcW')[0] : undefined;
      const w = tcW && tcW.getAttribute('w:type') === 'dxa' ? intAttr(tcW, 'w:w') : null;
      if (merged || w === null || w <= 0) {
        ok = false;
        break;
      }
      widths.push(w);
    }
    if (ok) return widths;
  }
  return null;
}

function isFloatingTable(tbl: Element): boolean {
  const tblPr = directChildren(tbl, 'w:tblPr')[0];
  return !!tblPr && directChildren(tblPr, 'w:tblpPr').length > 0;
}

export interface WidthContainer {
  /** Width in twips of what the table sits in, or null when it cannot be determined. */
  twips: number | null;
  /** Why it is unknown. */
  reason?: string;
  /** Human name, used in warnings ("the text width", "the width of its cell"). */
  label: string;
}

function marginOf(el: Element | undefined, side: 'left' | 'right'): number | null {
  if (!el) return null;
  const m = directChildren(el, `w:${side}`)[0] ?? directChildren(el, side === 'left' ? 'w:start' : 'w:end')[0];
  if (!m) return null;
  const t = m.getAttribute('w:type');
  if (t && t !== 'dxa') return null;
  return intAttr(m, 'w:w');
}

/** The width available to the table: the section's text width, or the width of the cell a nested table sits in. */
export function containerWidth(tbl: Element, pageSetup: PageSetupProperties | null | undefined): WidthContainer {
  let node: Node | null = tbl.parentNode;
  while (node && node.nodeType === 1 && (node as Element).tagName !== 'w:tc' && (node as Element).tagName !== 'w:body') {
    node = node.parentNode;
  }
  const parent = node && node.nodeType === 1 ? (node as Element) : null;
  if (!parent || parent.tagName === 'w:body') {
    if ((pageSetup?.columns ?? 1) > 1) {
      return { twips: null, reason: 'the section has several text columns, so the width of one column is not known', label: 'the text width' };
    }
    const area = usableAreaOf(pageSetup);
    return { twips: Math.round((area.widthEmu / EMU_PER_INCH) * TWIPS_PER_INCH), label: 'the text width' };
  }
  // A nested table: the parent cell's width from the parent grid, less the cell margins.
  const parentRow = parent.parentNode as Element | null;
  const parentTbl = parentRow?.parentNode as Element | null;
  if (!parentRow || parentRow.tagName !== 'w:tr' || !parentTbl || parentTbl.tagName !== 'w:tbl') {
    return { twips: null, reason: 'the table sits in a structure other than a body or a table cell', label: 'its container' };
  }
  const grid = readGrid(parentTbl);
  if (!grid) return { twips: null, reason: "the parent table's grid is missing or not numeric", label: 'the width of its cell' };
  let layout: ReturnType<typeof layoutRows>;
  try {
    layout = layoutRows([parentRow], grid.length);
  } catch {
    return { twips: null, reason: "the parent table's cell spans do not match its grid", label: 'the width of its cell' };
  }
  const cell = layout.rows[0].cells.find((c) => c.tc === parent);
  if (!cell) return { twips: null, reason: 'the parent cell was not found in its row', label: 'the width of its cell' };
  const width = grid.slice(cell.start, cell.start + cell.span).reduce((a, b) => a + b, 0);
  const tcPr = directChildren(parent, 'w:tcPr')[0];
  const parentPr = directChildren(parentTbl, 'w:tblPr')[0];
  const tcMar = tcPr ? directChildren(tcPr, 'w:tcMar')[0] : undefined;
  const tblCellMar = parentPr ? directChildren(parentPr, 'w:tblCellMar')[0] : undefined;
  const left = marginOf(tcMar, 'left') ?? marginOf(tblCellMar, 'left') ?? DEFAULT_CELL_MARGIN_TWIPS;
  const right = marginOf(tcMar, 'right') ?? marginOf(tblCellMar, 'right') ?? DEFAULT_CELL_MARGIN_TWIPS;
  const inner = width - left - right;
  if (inner <= 0) return { twips: null, reason: 'the parent cell is narrower than its margins', label: 'the width of its cell' };
  return { twips: inner, label: 'the width of its cell' };
}

function ensureTblPr(tbl: Element): Element {
  let tblPr = directChildren(tbl, 'w:tblPr')[0];
  if (!tblPr) {
    tblPr = createWordElement(tbl.ownerDocument, 'w:tblPr');
    tbl.insertBefore(tblPr, tbl.firstChild);
  }
  return tblPr;
}

function setTblW(tbl: Element, spec: { w: number; type: 'dxa' | 'pct' }): void {
  const tblPr = ensureTblPr(tbl);
  let tblW = directChildren(tblPr, 'w:tblW')[0];
  if (!tblW) {
    tblW = createWordElement(tbl.ownerDocument, 'w:tblW');
    tblPr.appendChild(tblW);
  }
  tblW.setAttribute('w:w', String(spec.w));
  tblW.setAttribute('w:type', spec.type);
  sortChildrenBySchema(tblPr, TBLPR_CHILD_ORDER);
}

/** Scales `grid` to add up to `total` exactly: each column rounded in proportion, the remainder given to the last. */
export function scaleGrid(grid: number[], total: number): number[] {
  const sum = grid.reduce((a, b) => a + b, 0);
  const out = grid.map((w) => Math.round((w * total) / sum));
  const diff = total - out.reduce((a, b) => a + b, 0);
  out[out.length - 1] += diff;
  return out;
}

export interface TableWidthOutcome {
  warnings: string[];
}

/**
 * Applies `setTableWidth` to one table. `rows` are the table's own rows (nested tables untouched).
 * Everything is checked before the first change, so a refusal leaves the table as it was.
 */
export function applyTableWidth(
  tbl: Element,
  rows: Element[],
  request: TableWidthRequest,
  pageSetup: PageSetupProperties | null | undefined
): TableWidthOutcome {
  const warnings: string[] = [];
  if (isFloatingTable(tbl)) {
    fail(
      'setTableWidth does not support a floating table (it has w:tblpPr)',
      'make the table inline in Word, or set its width there; ASTra does not guess how a floating table should reflow'
    );
  }
  const { spec, columns } = request;
  const container = containerWidth(tbl, pageSetup);

  let tblW: { w: number; type: 'dxa' | 'pct' };
  let targetTwips: number | null;
  if (spec.kind === 'pct') {
    if (container.twips === null) {
      return fail(
        `setTableWidth widthPercent is refused: ${container.reason ?? 'the container width is not known'}`,
        'use widthIn or widthTwips instead'
      );
    }
    targetTwips = Math.round((container.twips * spec.percent) / 100);
    tblW = { w: Math.round(spec.percent * 50), type: 'pct' };
  } else {
    targetTwips = spec.twips;
    tblW = { w: spec.twips, type: 'dxa' };
  }

  if (columns === 'keep') {
    setTblW(tbl, tblW);
  } else {
    // Checks first: nothing changes unless the grid and every row can be trusted.
    let grid = readGrid(tbl);
    if (!grid) {
      grid = rebuiltGrid(rows);
      if (!grid) {
        return fail(
          'setTableWidth: the table grid is missing or not numeric and no row gives plain numeric cell widths to rebuild it from',
          'set the column widths first (setColumnWidths with widths in twips), or use columns: "keep" to change only the table width'
        );
      }
    }
    const { rows: laid } = layoutRows(rows, grid.length);
    if (laid.length === 0) fail('setTableWidth: the table has no rows', 'add a row first');
    const scaled = scaleGrid(grid, targetTwips as number);
    if (scaled.some((w) => w < MIN_COLUMN_TWIPS)) {
      fail(
        `setTableWidth: a width of ${targetTwips} twips leaves a column narrower than ${MIN_COLUMN_TWIPS} twips`,
        'use a larger width, or merge columns first'
      );
    }
    // Apply: grid, every cell width from the grid, table width, fixed layout.
    let tblGrid = directChildren(tbl, 'w:tblGrid')[0];
    if (!tblGrid) {
      ensureTableGrid(tbl, tbl.ownerDocument);
      tblGrid = directChildren(tbl, 'w:tblGrid')[0];
    }
    const gridCols = directChildren(tblGrid, 'w:gridCol');
    if (gridCols.length !== scaled.length) {
      while (tblGrid.firstChild) tblGrid.removeChild(tblGrid.firstChild);
      for (const w of scaled) {
        const col = createWordElement(tbl.ownerDocument, 'w:gridCol');
        col.setAttribute('w:w', String(w));
        tblGrid.appendChild(col);
      }
    } else {
      gridCols.forEach((col, i) => col.setAttribute('w:w', String(scaled[i])));
    }
    for (const row of laid) {
      for (const cell of row.cells) {
        let tcPr = directChildren(cell.tc, 'w:tcPr')[0];
        if (!tcPr) {
          tcPr = createWordElement(tbl.ownerDocument, 'w:tcPr');
          cell.tc.insertBefore(tcPr, cell.tc.firstChild);
        }
        let tcW = directChildren(tcPr, 'w:tcW')[0];
        if (!tcW) {
          tcW = createWordElement(tbl.ownerDocument, 'w:tcW');
          tcPr.appendChild(tcW);
        }
        let width = 0;
        for (let i = cell.start; i < cell.start + cell.span; i++) width += scaled[i];
        tcW.setAttribute('w:w', String(width));
        tcW.setAttribute('w:type', 'dxa');
        sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
      }
    }
    setTblW(tbl, tblW);
    const tblPr = ensureTblPr(tbl);
    let layout = directChildren(tblPr, 'w:tblLayout')[0];
    if (!layout) {
      layout = createWordElement(tbl.ownerDocument, 'w:tblLayout');
      tblPr.appendChild(layout);
    }
    layout.setAttribute('w:type', 'fixed');
    sortChildrenBySchema(tblPr, TBLPR_CHILD_ORDER);
  }

  if (spec.kind === 'dxa' && container.twips !== null && spec.twips > container.twips) {
    warnings.push(
      `layoutWarning: the table is ${Math.round((spec.twips / TWIPS_PER_INCH) * 100) / 100} in wide, ${container.label} is ${Math.round((container.twips / TWIPS_PER_INCH) * 100) / 100} in; Word clips a table at the page edge. Use a smaller width.`
    );
  }
  return { warnings };
}

export interface RowHeightsRequest {
  rule: 'atLeast' | 'exact' | 'auto';
  twips?: number;
}

/** Reads and validates the arguments of `setRowHeights` before anything is touched. */
export function parseRowHeights(pick: Pick): RowHeightsRequest {
  const heightIn = pick('heightIn');
  const heightTwips = pick('heightTwips');
  const ruleRaw = pick('rule');
  if (ruleRaw !== undefined && ruleRaw !== 'atLeast' && ruleRaw !== 'exact' && ruleRaw !== 'auto') {
    return fail(`setRowHeights rule must be "atLeast", "exact" or "auto" (got ${JSON.stringify(ruleRaw)})`, 'use rule "atLeast" (default), "exact" or "auto"');
  }
  const rule = (ruleRaw ?? 'atLeast') as RowHeightsRequest['rule'];
  if (heightIn !== undefined && heightTwips !== undefined) {
    return fail('setRowHeights got both heightIn and heightTwips', 'pass exactly one of heightIn or heightTwips');
  }
  if (rule === 'auto') {
    if (heightIn !== undefined || heightTwips !== undefined) {
      return fail(
        'setRowHeights with rule "auto" takes no height (a height is given)',
        'remove heightIn/heightTwips: "auto" clears the row heights; use "atLeast" or "exact" to set one'
      );
    }
    return { rule };
  }
  if (heightIn === undefined && heightTwips === undefined) {
    return fail(`setRowHeights with rule "${rule}" needs a height`, 'pass exactly one of heightIn (inches) or heightTwips, e.g. { heightIn: 0.3 }');
  }
  let twips: number;
  if (heightIn !== undefined) {
    if (typeof heightIn !== 'number' || !Number.isFinite(heightIn)) return fail(`setRowHeights heightIn must be a finite number (got ${JSON.stringify(heightIn)})`, 'pass heightIn as a number of inches');
    twips = Math.round(heightIn * TWIPS_PER_INCH);
  } else {
    if (typeof heightTwips !== 'number' || !Number.isInteger(heightTwips)) return fail(`setRowHeights heightTwips must be a whole number (got ${JSON.stringify(heightTwips)})`, 'pass heightTwips as a whole number of twips (1440 per inch)');
    twips = heightTwips;
  }
  if (twips < 1 || twips > MAX_TWIPS) {
    return fail(
      `setRowHeights height is outside the supported range (got ${twips} twips)`,
      `use a height from 1 to ${MAX_TWIPS} twips (up to 22 inches)`
    );
  }
  return { rule, twips };
}

/** Sets or clears the height of every row of the table (its own rows only). */
export function applyRowHeights(tbl: Element, rows: Element[], request: RowHeightsRequest): { warnings: string[] } {
  const warnings: string[] = [];
  for (const tr of rows) {
    let trPr = directChildren(tr, 'w:trPr')[0];
    if (request.rule === 'auto') {
      if (!trPr) continue;
      for (const h of directChildren(trPr, 'w:trHeight')) trPr.removeChild(h);
      if (!trPr.firstChild) tr.removeChild(trPr);
      continue;
    }
    if (!trPr) {
      trPr = createWordElement(tbl.ownerDocument, 'w:trPr');
      tr.insertBefore(trPr, tr.firstChild);
    }
    let h = directChildren(trPr, 'w:trHeight')[0];
    if (!h) {
      h = createWordElement(tbl.ownerDocument, 'w:trHeight');
      trPr.appendChild(h);
    }
    h.setAttribute('w:val', String(request.twips));
    h.setAttribute('w:hRule', request.rule);
    sortChildrenBySchema(trPr, TRPR_CHILD_ORDER);
  }
  if (request.rule === 'exact') {
    warnings.push('layoutWarning: rule "exact" fixes every row to this height; Word and LibreOffice clip content that does not fit. Use "atLeast" to let a row grow.');
  }
  return { warnings };
}
