import { DocumentAST, Node, calculateNodeChecksum } from '../ast';
import { BaseOperation, createTextElement, validateVersion } from './base';
import type { StyleGuideConfig } from '../style-guide/types';
import { XMLSerializer } from '@xmldom/xmldom';
import * as crypto from 'crypto';
import { resolveReproducibleMode } from '../zip-hardening';
import { ensureTableGrid, sortChildrenBySchema, RPR_CHILD_ORDER, TCPR_CHILD_ORDER, TRPR_CHILD_ORDER, TBLPR_CHILD_ORDER, PPR_CHILD_ORDER } from '../ooxml/schema-order';
import { createWordElement } from '../ooxml/ns';
import { emitTrackedCellReplace, resolveRevisionStamp } from '../ooxml/track-changes-emitter';
import { toError } from '../errors';

function isTableReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

function tableCaptionHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

function directChildren(parent: Element, tagName: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === tagName) {
      out.push(child as Element);
    }
  }
  return out;
}

// OP-10 helpers: a logical column index is not the physical w:tc index when
// cells span multiple grids (w:gridSpan) or merge horizontally (w:hMerge).
function getCellGridSpan(tc: Element): number {
  const tcPr = directChildren(tc, 'w:tcPr')[0];
  if (!tcPr) return 1;
  const gs = tcPr.getElementsByTagName('w:gridSpan')[0] as Element | undefined;
  if (!gs) return 1;
  const v = parseInt(gs.getAttribute('w:val') || '1', 10);
  return Number.isFinite(v) && v > 0 ? v : 1;
}

function getHMergeVal(tc: Element): string | null {
  const tcPr = directChildren(tc, 'w:tcPr')[0];
  if (!tcPr) return null;
  const hm = tcPr.getElementsByTagName('w:hMerge')[0] as Element | undefined;
  if (!hm) return null;
  // Per ECMA-376 a bare <w:hMerge/> means continue.
  return hm.getAttribute('w:val') || 'continue';
}

export function findLogicalTableCell(tr: Element, colIndex: number): { tc?: Element; isCover?: boolean; logicalWidth: number } {
  const cells = directChildren(tr, 'w:tc');
  let logical = 0;
  let targetTc: Element | undefined;
  let isCover = false;
  for (let ci = 0; ci < cells.length; ci++) {
    const cand = cells[ci];
    const hMerge = getHMergeVal(cand);
    const isContinue = hMerge !== null && hMerge !== 'restart';
    if (isContinue) {
      if (logical === colIndex) {
        isCover = true;
      }
      logical += 1;
      continue;
    }
    const span = getCellGridSpan(cand);
    if (colIndex >= logical && colIndex < logical + span) {
      targetTc = cand;
    }
    logical += span;
  }
  return { tc: targetTc, isCover, logicalWidth: logical };
}

export function resolveTableCell(tr: Element, colIndex: number, opName: string): Element {
  const { tc, isCover, logicalWidth } = findLogicalTableCell(tr, colIndex);
  if (isCover) {
    throw new Error(
      `E_VALIDATION: ${opName} column index ${shown(colIndex)} targets a merged cover cell (hMerge continue) and is not directly editable; target the merge restart instead.`
    );
  }
  if (!tc) {
    throw new Error(
      `E_VALIDATION: ${opName} column index ${shown(colIndex)} out of bounds (row has ${logicalWidth} logical columns). Remedy: pass colIndex in [1, ${logicalWidth}], counting from 1.`
    );
  }
  return tc;
}

/**
 * Contrast guardrail for shaded cells (WCAG 2.x AA, normal text).
 * Shading a cell with a dark fill while its text stays dark is a
 * readability defect, so setCellShading/setShading WARN instead of failing:
 * the fill is still written, and the warning is returned in the op result
 * (typed `warnings` array) so callers and tests can assert it. Library code
 * stays silent: no stdout/stderr logging (CLI remains the logging surface).
 *
 * Threshold: 4.5:1 — WCAG 2.0/2.1 success criterion 1.4.3 (AA) for normal
 * text. Deliberately NOT 3:1 (AA large-text only) and NOT 7:1 (AAA):
 * table body text is normal-size running text, so AA normal-text applies.
 */
export const TABLE_CONTRAST_MIN_RATIO = 4.5;

function srgbChannelToLinear(channel: number): number {
  const s = channel / 255;
  // WCAG 2.x linearization (0.03928 cutoff per the 2.0/2.1 formula).
  return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
}

/**
 * Relative luminance per WCAG 2.x (0 = black, 1 = white).
 * Input: 6 hex digits with optional leading '#'.
 */
export function hexToRelativeLuminance(hex: string): number {
  const h = hex.replace(/^#/, '');
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return 0.2126 * srgbChannelToLinear(r) + 0.7152 * srgbChannelToLinear(g) + 0.0722 * srgbChannelToLinear(b);
}

/** WCAG contrast ratio of two 6-digit hex colors, in [1, 21]. */
export function contrastRatio(firstHex: string, secondHex: string): number {
  const l1 = hexToRelativeLuminance(firstHex);
  const l2 = hexToRelativeLuminance(secondHex);
  const [hi, lo] = l1 >= l2 ? [l1, l2] : [l2, l1];
  return (hi + 0.05) / (lo + 0.05);
}

/**
 * First explicit 6-digit w:color in a cell's OWN paragraphs (direct w:p
 * children only — never descends into a nested table), or null when the
 * cell relies on the default. Callers treat null as black ('000000'):
 * w:color="auto" and absent color both render as black on a light page.
 */
function firstExplicitCellTextColor(tc: Element): string | null {
  for (const p of directChildren(tc, 'w:p')) {
    const runs: Element[] = [...directChildren(p, 'w:r')];
    for (const tag of ['w:hyperlink', 'w:ins', 'w:del']) {
      for (const wrap of directChildren(p, tag)) {
        const wRuns = wrap.getElementsByTagName('w:r');
        for (let i = 0; i < wRuns.length; i++) runs.push(wRuns[i] as Element);
      }
    }
    for (const r of runs) {
      const rPr = directChildren(r, 'w:rPr')[0];
      const colorEl = rPr?.getElementsByTagName('w:color')[0] as Element | undefined;
      const v = colorEl?.getAttribute('w:val');
      if (v && /^[0-9A-Fa-f]{6}$/.test(v)) return v.toUpperCase();
    }
  }
  return null;
}

/**
 * Pushes a low-contrast warning into `warnings` when fill vs text fails
 * TABLE_CONTRAST_MIN_RATIO. Never throws — the shading op still applies;
 * the warning is advisory. `seen` dedupes identical fill/text pairs so
 * whole-table shading warns once, not per cell. Silent by design: the
 * typed `warnings` array is the channel (no stderr logging).
 */
function warnIfLowContrast(
  warnings: string[],
  seen: Set<string>,
  fill: string,
  textColor: string,
  where: string
): void {
  const key = `${fill}:${textColor}`;
  if (seen.has(key)) return;
  const ratio = contrastRatio(textColor, fill);
  if (ratio < TABLE_CONTRAST_MIN_RATIO) {
    seen.add(key);
    const msg =
      `Table contrast warning: fill #${fill} with text #${textColor} at ${where} ` +
      `has contrast ${ratio.toFixed(2)}:1 below WCAG AA (${TABLE_CONTRAST_MIN_RATIO}:1). ` +
      `Remedy: use a lighter fill or a light text color (e.g. FFFFFF).`;
    warnings.push(msg);
  }
}

/**
 * Paints an explicit text color onto every run of a cell's OWN paragraphs
 * (direct w:p children — never into a nested table). Cells without runs
 * are left alone (no empty-run synthesis); the caller still reports the
 * intended color through the contrast guardrail.
 */
function paintCellTextColor(tc: Element, doc: Document, normColor: string): void {
  let ps = directChildren(tc, 'w:p');
  if (ps.length === 0) {
    const p = createWordElement(doc, 'w:p');
    tc.appendChild(p);
    ps = [p];
  }
  for (const p of ps) {
    const runs: Element[] = [...directChildren(p, 'w:r')];
    for (const tag of ['w:hyperlink', 'w:ins', 'w:del']) {
      for (const wrap of directChildren(p, tag)) {
        const wRuns = wrap.getElementsByTagName('w:r');
        for (let i = 0; i < wRuns.length; i++) runs.push(wRuns[i] as Element);
      }
    }
    for (const r of runs) {
      let rPr = directChildren(r, 'w:rPr')[0];
      if (!rPr) {
        rPr = createWordElement(doc, 'w:rPr');
        r.insertBefore(rPr, r.firstChild);
      }
      let colorEl = rPr.getElementsByTagName('w:color')[0];
      if (!colorEl) {
        colorEl = createWordElement(doc, 'w:color');
        rPr.appendChild(colorEl);
      }
      (colorEl as Element).setAttribute('w:val', normColor);
      sortChildrenBySchema(rPr, RPR_CHILD_ORDER);
    }
  }
}

async function refreshTableNode(node: Node, tblEl: Element): Promise<void> {
  // Ensure ECMA-376 compliant tblGrid exists
  if (tblEl.ownerDocument) {
    ensureTableGrid(tblEl, tblEl.ownerDocument);
  }
  const textNodes = tblEl.getElementsByTagName('w:t');
  let text = '';
  for (let i = 0; i < textNodes.length; i++) {
    text += textNodes[i].textContent || '';
  }
  node.metadata.text = text;
  const rows = directChildren(tblEl, 'w:tr');
  node.metadata.tableDimensions = { rows: rows.length, cols: rows.length > 0 ? directChildren(rows[0], 'w:tc').length : 0 };
  node.version += 1;
  node.checksum = await calculateNodeChecksum(new XMLSerializer().serializeToString(tblEl));
}
export type TableParams =
  | { op: "updateCell", rowIndex: number, colIndex: number, text: string }
  | { op: "setCaption", captionText: string, captionPosition?: string }
  | { op: "addRow", afterIndex?: number, cells?: string[] }
  | { op: "deleteRow", index: number }
  | { op: "addColumn", afterIndex: number, width?: number }
  | { op: "deleteColumn", index: number }
  | { op: "mergeCells", fromRow: number, fromCol: number, toRow: number, toCol: number }
  | { op: "splitCell", rowIndex: number, colIndex: number }
  | { op: "setCellShading", rowIndex: number, colIndex: number, fill: string }
  | { op: "setRowHeight", rowIndex: number, height: number, rule?: string }
  | { op: "setCellAlignment", rowIndex?: number, colIndex?: number, align: string }
  | { op: 'setTableAlignment'; align?: string; alignment?: string; [k: string]: unknown }
  | { op: "setCellPadding", rowIndex?: number, colIndex?: number, padding: number | { top?: number; left?: number; bottom?: number; right?: number } }
  | { op: "setCellTextColor", rowIndex?: number, colIndex?: number, color: string }
  | { op: "setKeepTogether" }
  | { op: "setColumnWidths", widths: number[] }
  | { op: "setBorder" }
  | { op: "setShading" };

export interface TableOperation extends BaseOperation {
  action: 'table';
  /** Table node id (e.g. 't_001'). Alias: tableId. */
  target?: string;
  tableId?: string;
  /** Op name may be given here instead of params.op. */
  operation?: string;
  /** Short alias for operation (flat or nested in params). */
  op?: string;
  trackChanges?: boolean;
  params?: Record<string, unknown>;
  /** Flat aliases accepted when params is omitted. */
  rowIndex?: number;
  colIndex?: number;
  /** Alias for colIndex (accepted flat and nested in params; colIndex wins when both present). */
  columnIndex?: number;
  /** Short aliases for rowIndex/colIndex (accepted flat and nested in params). */
  row?: number;
  col?: number;
  text?: string;
  /** Caption text for setCaption (flat alias; also accepted nested in params). */
  captionText?: string;
  /** Caption placement for setCaption: 'above' (default) | 'below'. */
  captionPosition?: string;
  afterIndex?: number;
  cells?: string[];
  index?: number;
  fill?: string;
  headerFill?: string;
  zebra?: boolean;
  height?: number;
  rule?: string;
  /** Horizontal cell alignment for setCellAlignment (flat alias; 'left'|'center'|'right'). */
  align?: string;
  /** Alias for align (flat alias for setCellAlignment). */
  alignment?: string;
  /** Cell padding for setCellPadding: uniform DXA number or per-side object. */
  padding?: number | { top?: number; left?: number; bottom?: number; right?: number };
  /** Attribution for a tracked updateCell: author (else settings.defaultAuthor) and ISO 8601 date (else now). */
  author?: string;
  date?: string;
  /** Text color for setShading header cells / setCellShading: 6 hex digits with optional leading '#' (flat alias; also nested in params). */
  textColor?: string;
  /** Text color for setCellTextColor: 6 hex digits with optional leading '#' (flat alias; also nested in params). */
  color?: string;
  /** Explicit column widths in DXA for setColumnWidths (flat alias). */
  widths?: number[];
  /** Alias for widths (flat alias for setColumnWidths). */
  columnWidths?: number[];
  /** Vertical alignment value for setVerticalAlign (flat alias). */
  vAlign?: string;
  /** Width in DXA for setColumnWidth (flat alias; `w` is the short form). */
  width?: number;
  /** Short alias for width (flat or nested in params). */
  w?: number;
  /** Merge range start/end rows/cols for mergeCells (flat or nested in params). */
  fromRow?: number;
  fromCol?: number;
  toRow?: number;
  toCol?: number;
}

export interface TableOperationResult {
  /** Advisory warnings (e.g. low-contrast fills). Ops still apply; never a failure. */
  warnings: string[];
}

/**
 * Fail-closed gate for content-structural table ops under an explicit
 * per-op trackChanges:true.
 *
 * These ops (row/column insert/delete, merge/split) have no paragraph-level
 * revision form: Word tracks them with structural markers (w:cellIns,
 * w:cellDel, w:cellMerge, trPr change records), which this engine does not
 * emit. Applying them anyway would silently drop the revision — the change
 * would vanish from review — so they throw E_VALIDATION naming the missing
 * marker instead. updateCell is NOT gated: cell text edits already emit
 * paragraph-level w:ins/w:del via emitTrackedCellReplace. Property-only ops
 * (shading, alignment, padding, widths, borders, header repeat, cantSplit,
 * keep-together, cell text color)
 * are NOT gated either: they alter no text runs, so there is no deletion
 * or insertion to revision-mark.
 *
 * Only an explicit per-op flag (flat `trackChanges` or nested in params)
 * triggers the gate. Ambient settings.trackChanges stays on the legacy
 * path (apply directly): the default blank document ships with tracking
 * enabled, and flipping every structural edit to a throw on that signal
 * alone would break the established table suites.
 */
function assertStructuralTrackingAllowed(opName: string): void {
  const markerByOp: Record<string, string> = {
    addRow: 'row insertion requires structural revision markers (w:cellIns / tracked w:trPr) which the engine does not emit',
    deleteRow: 'row deletion requires structural revision markers (w:cellDel) which the engine does not emit',
    addColumn: 'column insertion requires structural revision markers (w:cellIns) which the engine does not emit',
    deleteColumn: 'column deletion requires structural revision markers (w:cellDel) which the engine does not emit',
    mergeCells: 'cell merge requires structural revision markers (w:cellMerge / tracked w:tcPr) which the engine does not emit',
    splitCell: 'cell split requires structural revision markers (tracked w:tcPr change) which the engine does not emit',
  };
  const limitation = markerByOp[opName];
  if (!limitation) return;
  throw Object.assign(
    new Error(
      `E_VALIDATION: ${opName} under trackChanges:true is not supported: ${limitation}; ` +
      `applying it would silently drop the revision. ` +
      `Remedy: re-run with trackChanges:false to apply the structural change outright, ` +
      `or use updateCell for paragraph-level tracked text edits.`
    ),
    {
      code: 'E_VALIDATION' as const,
      retryable: false as const,
      remedy: 'Re-run with trackChanges:false to apply the structural change outright, or use updateCell for paragraph-level tracked text edits.',
    }
  );
}

// Row/column short-alias precedence — uniform across updateCell/splitCell/
// deleteColumn/setCellShading/setCellAlignment/setCellPadding/setVerticalAlign
// (and the P-chain that feeds updateCell/splitCell):
//   row: raw.rowIndex > raw.row > flat.rowIndex > flat.row
//   col: raw.colIndex > raw.col > flat.colIndex > flat.col > raw.columnIndex > flat.columnIndex
// `colIndex` wins over `col` wins over `columnIndex`; nested params win over
// flat within the same name. Non-integer/negative values fail closed at call
// sites via existing E_VALIDATION checks.
function rowIndexAlias(raw: Record<string, unknown>, op: TableOperation): unknown {
  if (raw.rowIndex !== undefined) return raw.rowIndex;
  if (raw.row !== undefined) return raw.row;
  if (op.rowIndex !== undefined) return op.rowIndex;
  if (op.row !== undefined) return op.row;
  return undefined;
}

function colIndexAlias(raw: Record<string, unknown>, op: TableOperation): unknown {
  if (raw.colIndex !== undefined) return raw.colIndex;
  if (raw.col !== undefined) return raw.col;
  if (op.colIndex !== undefined) return op.colIndex;
  if (op.col !== undefined) return op.col;
  if (raw.columnIndex !== undefined) return raw.columnIndex;
  if (op.columnIndex !== undefined) return op.columnIndex;
  return undefined;
}

// R-42: a table caption is a paragraph sibling whose pPr carries the
// Caption style (same marker image setCaption writes and the HTML exporter
// reads at export/html.ts:707-709). Both 'w:val' and bare 'val' spellings
// count, mirroring the image caption check.
function isCaptionParagraphNode(n: Node | undefined): boolean {
  if (!n || n.type !== 'paragraph' || !n.xml) return false;
  const pPr = (n.xml as Element).getElementsByTagName('w:pPr')[0] as Element | undefined;
  const pStyle = pPr?.getElementsByTagName('w:pStyle')[0] as Element | undefined;
  const v = pStyle?.getAttribute('w:val') ?? pStyle?.getAttribute('val');
  return v === 'Caption';
}

// Zebra striping is not a table property: setShading paints alternate data rows cell by cell
// (odd rows one fill, even rows another). A cloned row therefore carries its source row's fill
// and breaks the pattern, so addRow re-applies it, but only when the table really is striped.

/** The one fill shared by every cell of a row, or null when cells differ or any is unshaded. */
function uniformRowFill(tr: Element): string | null {
  const cells = directChildren(tr, 'w:tc');
  if (cells.length === 0) return null;
  let fill: string | null = null;
  for (const tc of cells) {
    const shd = directChildren(directChildren(tc, 'w:tcPr')[0] ?? tc, 'w:shd')[0];
    const value = shd?.getAttribute('w:fill');
    if (!value || value === 'auto') return null;
    if (fill !== null && fill !== value) return null;
    fill = value;
  }
  return fill;
}

/** The two alternating fills of a striped table (odd data rows, even data rows), or null. */
function detectZebra(rows: Element[]): { odd: string; even: string } | null {
  let odd: string | null = null;
  let even: string | null = null;
  const body = rows.slice(1);
  if (body.length < 2) return null;
  for (let i = 0; i < body.length; i++) {
    const fill = uniformRowFill(body[i]);
    if (fill === null) return null;
    if (i % 2 === 0) {
      if (odd !== null && odd !== fill) return null;
      odd = fill;
    } else {
      if (even !== null && even !== fill) return null;
      even = fill;
    }
  }
  return odd !== null && even !== null && odd !== even ? { odd, even } : null;
}

function paintRowFill(tr: Element, fill: string): void {
  for (const tc of directChildren(tr, 'w:tc')) {
    const doc = tc.ownerDocument;
    if (!doc) continue;
    let tcPr = directChildren(tc, 'w:tcPr')[0];
    if (!tcPr) {
      tcPr = createWordElement(doc, 'w:tcPr');
      tc.insertBefore(tcPr, tc.firstChild);
    }
    let shd = directChildren(tcPr, 'w:shd')[0];
    if (!shd) {
      shd = createWordElement(doc, 'w:shd');
      tcPr.appendChild(shd);
    }
    shd.setAttribute('w:val', 'clear');
    shd.setAttribute('w:color', 'auto');
    shd.setAttribute('w:fill', fill);
    sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
  }
}

/** Paint every data row by its position so the rows below an insert or a removal continue the pattern. */
function restripe(tblEl: Element, zebra: { odd: string; even: string }): void {
  directChildren(tblEl, 'w:tr').slice(1).forEach((tr, i) => paintRowFill(tr, i % 2 === 0 ? zebra.odd : zebra.even));
}

// Rows and columns are numbered from 1 in the operation, as Word, Excel and the inspect grid
// (r1c1) number them. The code below indexes arrays, so every such number is shifted down once
// here, flat and inside params; messages add the 1 back before they print a number.
const GRID_NUMBER_FIELDS = [
  'rowIndex', 'row', 'colIndex', 'col', 'columnIndex', 'index', 'afterIndex', 'fromRow', 'fromCol', 'toRow', 'toCol',
] as const;

function shiftGridNumbers(source: Record<string, unknown>): Record<string, unknown> {
  let shifted: Record<string, unknown> | undefined;
  for (const key of GRID_NUMBER_FIELDS) {
    const v = source[key];
    const whole =
      (typeof v === 'number' && Number.isInteger(v)) || (typeof v === 'string' && /^\s*-?\d+\s*$/.test(v));
    if (!whole) continue;
    shifted ??= { ...source };
    shifted[key] = Number(v) - 1;
  }
  return shifted ?? source;
}

/** A row or column number as the caller wrote it: whole numbers were shifted on the way in, anything else was not. */
const shown = (n: number): number => (Number.isInteger(n) ? n + 1 : n);

/** Refuse a row number (already shifted to 0-based) that is not one of the table's rows. */
function assertRowInRange(opName: string, param: string, zeroBased: number, rowCount: number): void {
  if (!Number.isInteger(zeroBased) || zeroBased < 0 || zeroBased >= rowCount) {
    throw new Error(
      `E_VALIDATION: ${opName} row index ${shown(zeroBased)} out of bounds (table has ${rowCount} rows). Remedy: pass ${param} in [1, ${rowCount}], counting from 1.`
    );
  }
}

/** Refuse a column number (already shifted to 0-based) that is not one of the table's columns; `within` finishes "(...)" in the message. */
function assertColumnInRange(opName: string, param: string, zeroBased: number, columnCount: number, within: string): void {
  if (!Number.isInteger(zeroBased) || zeroBased < 0 || zeroBased >= columnCount) {
    throw new Error(
      `E_VALIDATION: ${opName} column index ${shown(zeroBased)} out of bounds (${within}). Remedy: pass ${param} in [1, ${columnCount}], counting from 1.`
    );
  }
}

/** The widest row, counting the columns a merge covers. */
function widestRow(rows: Element[]): number {
  let widest = 0;
  for (const tr of rows) widest = Math.max(widest, findLogicalTableCell(tr, 0).logicalWidth);
  return widest;
}

function shiftTableGridNumbers(op: TableOperation): TableOperation {
  const flat = shiftGridNumbers(op as unknown as Record<string, unknown>);
  const params = typeof op.params === 'object' && op.params !== null ? shiftGridNumbers(op.params as Record<string, unknown>) : op.params;
  if (flat === (op as unknown) && params === op.params) return op;
  return { ...flat, params } as unknown as TableOperation;
}

export async function executeTable(opIn: TableOperation, ast: DocumentAST, guide?: StyleGuideConfig): Promise<TableOperationResult> {
  const op = shiftTableGridNumbers(opIn);
  const targetId = op.target || op.tableId || ((typeof op.params === 'object' && op.params !== null) ? ((op.params as Record<string, unknown>).target as string | undefined || (op.params as Record<string, unknown>).tableId as string | undefined) : undefined) || '';
  const raw = (op.params || {}) as Record<string, unknown>;
  const rawOp = raw.op ?? op.op ?? op.operation ?? op.params?.operation ?? '';
  const resolvedOp = String(rawOp);

  let tableNode: Node | null = null;
  if (targetId) {
    for (const section of ast.sections) {
      const found = section.nodes.find(n =>
        n.id === targetId ||
        ((n.xml as Element | undefined)?.getAttribute?.('data-astra-id') === targetId) ||
        (n.type === 'table' && /^t_\d+$/.test(targetId) && /^t_\d+$/.test(n.id) && targetId.replace(/^t_0*/, 't_') === n.id.replace(/^t_0*/, 't_'))
      );
      if (found) {
        tableNode = found;
        break;
      }
    }
    if (!tableNode) {
      const available = ast.sections.flatMap(s => s.nodes.filter(n => n.type === 'table').map(n => n.id));
      throw new Error(
        `Table target '${targetId}' not found. Available table nodes: ${available.join(', ') || 'none'}. ` +
        `Expected shape: { action:'table', target:'<tableId>', params:{ op:'updateCell', rowIndex:1, colIndex:2, text:'new text' } }`
      );
    }
  } else {
    // If targetId is not specified, try to find the first table node
    for (const section of ast.sections) {
      const found = section.nodes.find(n => n.type === 'table');
      if (found) {
        tableNode = found;
        break;
      }
    }
  }

  if (!tableNode || !tableNode.xml) {
    if (resolvedOp === 'setTableAlignment' && !targetId) {
      throw toError(
        'E_VALIDATION',
        'Table operation "setTableAlignment" requires a target table node id.',
        'Pass target: "<tableNodeId>" in the operation payload.'
      );
    }
    const available = ast.sections.flatMap(s => s.nodes.filter(n => n.type === 'table').map(n => n.id));
    throw new Error(
      `Table target '${targetId || '(missing)'}' not found. Available table nodes: ${available.join(', ') || 'none'}. ` +
      `Expected shape: { action:'table', target:'<tableId>', params:{ op:'updateCell', rowIndex:1, colIndex:2, text:'new text' } }`
    );
  }

  if (tableNode.type !== 'table' || (tableNode.xml as Element)?.localName !== 'tbl') {
    const available = ast.sections.flatMap(s => s.nodes.filter(n => n.type === 'table').map(n => n.id));
    throw toError(
      'E_VALIDATION',
      `Target '${targetId}' is a ${tableNode.type} (<${(tableNode.xml as Element)?.localName || tableNode.type}>), but table operations require a table target.`,
      `Pass a valid table target id (${available.length > 0 ? `e.g. ${available.join(', ')}` : 'inspect document for table nodes'}) from inspect_document.`
    );
  }

  validateVersion(tableNode, op.expectedVersion);

  const tblEl: Element = tableNode.xml;
  // Resolve rows from DIRECT children only: getElementsByTagName would also
  // return rows of nested tables and silently shift every row index.
  const rows = directChildren(tblEl, 'w:tr');

  // Accept params either nested under params:{} or flat on the operation
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : (op[key] !== undefined ? (op[key] as T) : flat);

  const rawCells = pick('cells', op.cells);
  const cellsArray = Array.isArray(rawCells)
    ? (rawCells as unknown[]).map(v => String(v ?? ''))
    : undefined;

  const P = {
    op: resolvedOp,
    rowIndex: Number(rowIndexAlias(raw, op) ?? 0),
    // Alias order (see helpers): colIndex > col > columnIndex, nested params > flat.
    colIndex: Number(colIndexAlias(raw, op) ?? 0),
    text: String(pick('text', op.text) ?? ''),
    afterIndex: Number(pick('afterIndex', op.afterIndex) ?? rows.length - 1),
    cells: cellsArray,
    index: Number(pick('index', op.index) ?? 0),
    fill: String(pick('fill', op.fill) ?? 'F7F9FC').replace(/^#/, ''),
    headerFill: (pick('headerFill', op.headerFill) as string | undefined)?.replace(/^#/, ''),
    zebra: pick('zebra', op.zebra) as boolean | undefined,
    // Explicit header/single-cell text color for setShading/setCellShading
    // (flat or nested in params; wins over the guide per the precedence
    // invariant). Validated below before any mutation.
    textColor: (pick('textColor', op.textColor) as string | undefined),
  };

  // Advisory warnings collected across shading ops (contrast guardrail).
  // Returned to the caller in the typed result; never a failure.
  const warnings: string[] = [];
  const warnedContrastPairs = new Set<string>();

  // Structural tracking gate (explicit per-op flag only; see helper docs).
  const explicitTrack = op.trackChanges ?? raw.trackChanges;
  if (explicitTrack === true) {
    assertStructuralTrackingAllowed(resolvedOp);
  }

  // Explicit text color for the shading ops (validated up-front so a bad
  // value fails closed before any mutation). Applies to the header row for
  // setShading and to the single cell for setCellShading; wins over the
  // guide per the precedence invariant (explicit > guide > built-in).
  let explicitTextColor: string | undefined;
  if ((resolvedOp === 'setShading' || resolvedOp === 'setCellShading') && P.textColor !== undefined) {
    const normText = String(P.textColor).replace(/^#/, '');
    if (!/^[0-9A-Fa-f]{6}$/.test(normText)) {
      throw new Error(
        `E_VALIDATION: ${resolvedOp} requires textColor as a 6-digit hex color with optional leading '#' (got textColor='${String(P.textColor)}'). ` +
          `Remedy: pass textColor as 6 hex digits, e.g. textColor:'FFFFFF'.`
      );
    }
    explicitTextColor = normText.toUpperCase();
  }

  /**
   * Guide default for a freshly shaded fill: the guide header text when the
   * fill is too dark for default dark text (WCAG AA 4.5:1, computed — never
   * string-matched against headerFill, so light fills stay dark-texted even
   * under a dark-header guide). Undefined without a guide, without a guide
   * header text, or when the fill already contrasts with dark text.
   */
  const guideTextForFill = (normFill: string): string | undefined => {
    const headerText = guide?.tables?.headerText?.replace(/^#/, '').toUpperCase();
    if (!headerText || !/^[0-9A-F]{6}$/.test(headerText)) return undefined;
    if (!/^[0-9A-Fa-f]{6}$/.test(normFill)) return undefined;
    return contrastRatio('000000', normFill.toUpperCase()) < TABLE_CONTRAST_MIN_RATIO
      ? headerText
      : undefined;
  };

  switch (P.op) {
    case 'updateCell': {
      const { rowIndex, colIndex, text } = P;
      // Who and when for a tracked cell edit: this call's author/date, else the document default.
      const stamp = resolveRevisionStamp(pick('author', op.author), pick('date', op.date), ast.settings.defaultAuthor);
      const trackChanges = op.trackChanges ?? raw.trackChanges ?? ast.settings.trackChanges;
      const trackOn = trackChanges === true;
      if (!Number.isInteger(rowIndex) || rowIndex < 0) {
        throw new Error(
          `E_VALIDATION: updateCell row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
        );
      }
      if (!Number.isInteger(colIndex) || colIndex < 0) {
        throw new Error(
          `E_VALIDATION: updateCell column index ${shown(colIndex)} out of bounds. Remedy: pass colIndex >= 1 (columns are numbered from 1).`
        );
      }
      const tr = rows[rowIndex];
      if (!tr) {
        throw new Error(
          `E_VALIDATION: updateCell row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
        );
      }
      const tc = resolveTableCell(tr, colIndex, 'updateCell');
      // Only this cell's own paragraphs count — never descend into a nested table.
      const cellPs = directChildren(tc, 'w:p');
      // Collect ALL w:t in this cell's own paragraphs (multi-w:t cells have
      // one w:t per run; writing only the first leaves stale tail text).
      const allT: Element[] = [];
      for (const p of cellPs) {
        const runs = directChildren(p, 'w:r');
        for (const r of runs) {
          const ts = directChildren(r, 'w:t');
          for (const t of ts) allT.push(t);
        }
        // Runs nested in hyperlink/ins/del wrappers inside this paragraph.
        for (const tag of ['w:hyperlink', 'w:ins', 'w:del']) {
          const wraps = directChildren(p, tag);
          for (const wEl of wraps) {
            const wRuns = wEl.getElementsByTagName('w:r');
            for (let i = 0; i < wRuns.length; i++) {
              const wts = directChildren(wRuns[i] as Element, 'w:t');
              for (const t of wts) allT.push(t);
            }
          }
        }
      }
      if (allT.length === 0) {
        let pNode = cellPs[0];
        if (!pNode) {
          pNode = createWordElement(tblEl.ownerDocument, 'w:p');
          tc.appendChild(pNode);
        }
        // Tracking on an empty cell: the inserted text is still a revision
        // (w:ins), mirroring Word. See emitTrackedCellReplace docs for why
        // whole-cell w:cellIns/w:cellDel markers do not apply to text edits.
        const trackEmpty = trackOn;
        if (trackEmpty && text.length > 0) {
          const hasRevs = emitTrackedCellReplace(
            tc,
            '',
            text,
            stamp.author,
            stamp.date,
            { getNextId: () => ast.nextChangeId++ }
          );
          if (hasRevs) {
            tableNode.metadata.hasTrackedChanges = true;
          }
        } else {
          const rNode = createWordElement(tblEl.ownerDocument, 'w:r');
          const tNode = createWordElement(tblEl.ownerDocument, 'w:t');
          tNode.setAttribute('xml:space', 'preserve');
          tNode.textContent = text;
          rNode.appendChild(tNode);
          pNode.appendChild(rNode);
        }
      } else if (
        trackOn &&
        allT.map(t => t.textContent || '').join('') !== text
      ) {
        // Tracked cell edit: paragraph-level w:ins/w:del revisions scoped to
        // this cell's own paragraphs (documented equivalent of cellIns/
        // cellDel, which Word reserves for structural cell insert/delete).
        const oldCellText = allT.map(t => t.textContent || '').join('');
        const hasRevs = emitTrackedCellReplace(
          tc,
          oldCellText,
          text,
          stamp.author,
          stamp.date,
          { getNextId: () => ast.nextChangeId++ }
        );
        if (hasRevs) {
          tableNode.metadata.hasTrackedChanges = true;
        }
      } else {
        // Replace first w:t + clear extras (never concat-loss/duplication).
        allT[0].textContent = text;
        try { allT[0].setAttribute('xml:space', 'preserve'); } catch { /* ignore */ }
        for (let i = 1; i < allT.length; i++) allT[i].textContent = '';
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setCaption': {
      // R-42: table caption with placement ('above' default | 'below').
      // Creates (or adopts + moves) a Caption-styled paragraph sibling on
      // the requested side of the table, in DOM order AND AST order, so
      // every exporter serializes it on that side with no exporter change:
      // document.xml carries the caption paragraph before/after the w:tbl,
      // and HTML/Markdown render the sibling paragraph in node order.
      // The caption text lives ONLY in the sibling (never copied into
      // metadata.caption) so HTML/Markdown render it exactly once.
      // metadata.captionPosition records the placement for the L-3
      // exporter branches (ast.ts is untouched — Record cast, no `any`).
      const captionRaw = raw.captionText ?? op.captionText;
      if (captionRaw === undefined || captionRaw === null || String(captionRaw).trim() === '') {
        throw new Error(
          `E_VALIDATION: setCaption requires captionText (got ${captionRaw === undefined ? 'no captionText' : "'" + String(captionRaw) + "'"}). ` +
            `Remedy: pass captionText as a non-empty string (e.g. { action:'table', target:'<tableId>', operation:'setCaption', captionText:'Table 1: Results' }); run \`astra inspect\` to list table ids.`
        );
      }
      const captionText = String(captionRaw);
      const posRaw = raw.captionPosition ?? op.captionPosition ?? 'above';
      const captionPosition = String(posRaw);
      if (captionPosition !== 'above' && captionPosition !== 'below') {
        throw new Error(
          `E_VALIDATION: setCaption captionPosition must be 'above'|'below' (got '${String(posRaw)}'). ` +
            `Remedy: pass captionPosition:'above' (default) or captionPosition:'below'; run \`astra inspect\` to list table ids.`
        );
      }
      let capSection: { nodes: Node[] } | null = null;
      for (const section of ast.sections) {
        if (section.nodes.some(n => n.id === tableNode.id)) {
          capSection = section;
          break;
        }
      }
      if (!capSection) {
        throw new Error(
          `E_VALIDATION: setCaption target table '${tableNode.id}' is detached from the AST. ` +
            `Remedy: re-read ids via \`astra inspect\` and retry.`
        );
      }
      const parent = tblEl.parentNode as Element | null;
      if (!parent) {
        throw new Error(
          `E_VALIDATION: setCaption target table '${tableNode.id}' is detached from the XML DOM parent. ` +
            `Remedy: re-read ids via \`astra inspect\` and retry.`
        );
      }
      const doc = tblEl.ownerDocument;
      const wantBefore = captionPosition === 'above';
      const tableIdxNow = capSection.nodes.findIndex(n => n.id === tableNode.id);
      const beforeNode = tableIdxNow > 0 ? capSection.nodes[tableIdxNow - 1] : undefined;
      const afterNode = tableIdxNow + 1 < capSection.nodes.length ? capSection.nodes[tableIdxNow + 1] : undefined;
      // Adopt the requested-side caption when present (idempotent re-set),
      // else move the opposite-side one, else mint a fresh paragraph.
      // This op never creates a second caption and never deletes one.
      let capNode: Node | undefined;
      if (wantBefore && isCaptionParagraphNode(beforeNode)) capNode = beforeNode;
      else if (!wantBefore && isCaptionParagraphNode(afterNode)) capNode = afterNode;
      else if (wantBefore && isCaptionParagraphNode(afterNode)) capNode = afterNode;
      else if (!wantBefore && isCaptionParagraphNode(beforeNode)) capNode = beforeNode;
      if (!capNode) {
        const liveIds = new Set<string>();
        for (const s of ast.sections) for (const n of s.nodes) liveIds.add(n.id);
        // Deterministic by default (content-seeded hash, same pattern as
        // callout/section); wall-clock only on explicit/env opt-out.
        let capId: string;
        if (isTableReproducible()) {
          const base = `tcap_${tableNode.id}_${tableCaptionHash8(`${tableNode.id}:${captionText}:${captionPosition}`)}`;
          capId = base;
          let n = 2;
          while (liveIds.has(capId)) capId = `${base}_${n++}`;
        } else {
          capId = `tcap_${tableNode.id}_${Date.now()}`;
          let seq = 0;
          while (liveIds.has(capId)) capId = `tcap_${tableNode.id}_${Date.now()}_${++seq}`;
        }
        const p = createWordElement(doc, 'w:p');
        p.setAttribute('data-astra-id', capId);
        const pPr = createWordElement(doc, 'w:pPr');
        const pStyle = createWordElement(doc, 'w:pStyle');
        pStyle.setAttribute('w:val', 'Caption');
        pPr.appendChild(pStyle);
        p.appendChild(pPr);
        capNode = {
          id: capId,
          type: 'paragraph',
          xml: p,
          metadata: {
            text: captionText,
            namedStyle: 'Caption',
            hasCitation: false,
            hasComment: false,
            hasBookmark: false,
            hasTrackedChanges: false,
            hasEquation: false,
            hasField: false,
            hasImage: false,
            hasHyperlink: false,
          },
          version: 1,
          checksum: '',
        };
      }
      // Rewrite runs (pPr preserved) so re-sets replace, never append.
      const capEl = capNode.xml as Element;
      for (const k of Array.from(capEl.childNodes)) {
        if (k.nodeType === 1 && (k as Element).tagName === 'w:r') capEl.removeChild(k);
      }
      const rNode = createWordElement(doc, 'w:r');
      rNode.appendChild(createTextElement(doc, captionText));
      capEl.appendChild(rNode);
      // Reposition in DOM + AST (remove-then-insert converges moves and
      // fresh nodes through one path).
      if (capEl.parentNode) capEl.parentNode.removeChild(capEl);
      const staleIdx = capSection.nodes.findIndex(n => n.id === capNode!.id);
      if (staleIdx !== -1) capSection.nodes.splice(staleIdx, 1);
      const tIdx = capSection.nodes.findIndex(n => n.id === tableNode.id);
      if (wantBefore) {
        parent.insertBefore(capEl, tblEl);
        capSection.nodes.splice(tIdx, 0, capNode);
      } else {
        parent.insertBefore(capEl, tblEl.nextSibling);
        capSection.nodes.splice(tIdx + 1, 0, capNode);
      }
      capNode.metadata.text = captionText;
      capNode.metadata.namedStyle = 'Caption';
      capNode.version += 1;
      capNode.checksum = await calculateNodeChecksum(new XMLSerializer().serializeToString(capEl));
      (tableNode.metadata as unknown as Record<string, unknown>).captionPosition = captionPosition;
      tableNode.version += 1;
      break;
    }
    case 'addRow': {
      const { afterIndex } = P;
      if (pick('afterIndex', op.afterIndex) !== undefined) assertRowInRange('addRow', 'afterIndex', afterIndex, rows.length);
      const refRow = rows[afterIndex] || rows[rows.length - 1];
      const newRow = refRow.cloneNode(true) as Element;
      // A verbatim clone would duplicate the source row's native w14:paraIds
      // — two paragraphs sharing one id is ambiguous downstream and the
      // save guard refuses it. Stripped ids are re-minted on the next load,
      // so the row keeps a reload-stable identity.
      const clonedParas = newRow.getElementsByTagName('w:p');
      for (let pIdx = 0; pIdx < clonedParas.length; pIdx++) {
        (clonedParas[pIdx] as unknown as Element).removeAttribute('w14:paraId');
      }
      // Clear cell texts in new row
      const tNodes = Array.from(newRow.getElementsByTagName('w:t'));
      tNodes.forEach(t => { t.textContent = ''; });
      if (Array.isArray(P.cells)) {
        const rowCells = directChildren(newRow, 'w:tc');
        for (let cIdx = 0; cIdx < rowCells.length && cIdx < P.cells.length; cIdx++) {
          const tc = rowCells[cIdx];
          const textVal = String(P.cells[cIdx] ?? '');
          const cellPs = directChildren(tc, 'w:p');
          const allT: Element[] = [];
          for (const p of cellPs) {
            const runs = directChildren(p, 'w:r');
            for (const r of runs) {
              const ts = directChildren(r, 'w:t');
              for (const t of ts) allT.push(t);
            }
            for (const tag of ['w:hyperlink', 'w:ins', 'w:del']) {
              const wraps = directChildren(p, tag);
              for (const wEl of wraps) {
                const wRuns = wEl.getElementsByTagName('w:r');
                for (let i = 0; i < wRuns.length; i++) {
                  const wts = directChildren(wRuns[i] as Element, 'w:t');
                  for (const t of wts) allT.push(t);
                }
              }
            }
          }
          if (allT.length === 0) {
            let pNode = cellPs[0];
            if (!pNode) {
              pNode = createWordElement(tblEl.ownerDocument, 'w:p');
              tc.appendChild(pNode);
            }
            const rNode = createWordElement(tblEl.ownerDocument, 'w:r');
            const tNode = createWordElement(tblEl.ownerDocument, 'w:t');
            tNode.setAttribute('xml:space', 'preserve');
            tNode.textContent = textVal;
            rNode.appendChild(tNode);
            pNode.appendChild(rNode);
          } else {
            allT[0].setAttribute('xml:space', 'preserve');
            allT[0].textContent = textVal;
            for (let i = 1; i < allT.length; i++) {
              allT[i].textContent = '';
            }
          }
        }
      }
      // B-6 pagination hygiene: a cloned data row keeps cantSplit but must
      // never inherit tblHeader (only the true header row repeats).
      let newTrPr = directChildren(newRow, 'w:trPr')[0];
      if (!newTrPr) {
        newTrPr = createWordElement(tblEl.ownerDocument, 'w:trPr');
        newRow.insertBefore(newTrPr, newRow.firstChild);
      }
      const staleHeader = newTrPr.getElementsByTagName('w:tblHeader')[0];
      if (staleHeader) newTrPr.removeChild(staleHeader);
      if (!newTrPr.getElementsByTagName('w:cantSplit')[0]) {
        newTrPr.appendChild(createWordElement(tblEl.ownerDocument, 'w:cantSplit'));
        sortChildrenBySchema(newTrPr, TRPR_CHILD_ORDER);
      }
      const zebra = detectZebra(rows);
      if (refRow.nextSibling) {
        tblEl.insertBefore(newRow, refRow.nextSibling);
      } else {
        tblEl.appendChild(newRow);
      }
      if (zebra) restripe(tblEl, zebra);
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'deleteRow': {
      // Precedence: index > rowIndex > row (nested params win over flat within
      // the same name; row aliases per rowIndexAlias). No default: a missing
      // index must never delete the header row.
      const indexSrc = raw.index ?? op.index ?? rowIndexAlias(raw, op);
      if (indexSrc === undefined) {
        throw new Error(
          `E_VALIDATION: deleteRow requires a row index (table has ${rows.length} rows). Remedy: pass index (or rowIndex) in [1, ${rows.length}], counting from 1.`
        );
      }
      const index = typeof indexSrc === 'number' || (typeof indexSrc === 'string' && indexSrc.trim() !== '') ? Number(indexSrc) : NaN;
      if (!Number.isInteger(index) || index < 0 || !rows[index]) {
        throw new Error(
          `E_VALIDATION: deleteRow row index ${shown(index)} out of bounds (table has ${rows.length} rows). Remedy: pass index in [1, ${rows.length}], counting from 1.`
        );
      }
      const targetRow = rows[index];
      // Stripes are painted cell by cell, so removing a body row leaves the rows below it one fill out of step.
      const zebra = index > 0 ? detectZebra(rows) : null;
      tblEl.removeChild(targetRow);
      if (zebra) restripe(tblEl, zebra);
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setShading': {
      const fill = P.fill;
      const isZebra = P.zebra !== false;
      rows.forEach((row, rIdx) => {
        const rowFill = rIdx === 0
          ? (P.headerFill || fill)
          : (isZebra ? (rIdx % 2 === 1 ? fill : 'FFFFFF') : fill);
        if (isZebra && rIdx === 0 && !P.headerFill) return;
        const cells = directChildren(row, 'w:tc');
        cells.forEach((tc, cIdx) => {
          let tcPr = directChildren(tc, 'w:tcPr')[0];
          if (!tcPr) {
            tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
            tc.insertBefore(tcPr, tc.firstChild);
          }
          let shd = tcPr.getElementsByTagName('w:shd')[0];
          if (!shd) {
            shd = createWordElement(tblEl.ownerDocument, 'w:shd');
            tcPr.appendChild(shd);
          }
          shd.setAttribute('w:val', 'clear');
          shd.setAttribute('w:color', 'auto');
          shd.setAttribute('w:fill', rowFill);
          sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
          // Header row taste defaulting (explicit > guide > built-in): an
          // explicit textColor always wins; otherwise the guide header text
          // applies when the fresh fill is too dark for dark text (WCAG-gated,
          // computed per fill). Body rows keep their existing text.
          let effectiveText: string | undefined;
          if (rIdx === 0) {
            effectiveText = explicitTextColor ?? guideTextForFill(rowFill);
            if (effectiveText && tblEl.ownerDocument) {
              paintCellTextColor(tc, tblEl.ownerDocument, effectiveText);
            }
          }
          // Contrast guardrail: advisory only, never fails the op.
          warnIfLowContrast(warnings, warnedContrastPairs, rowFill, effectiveText ?? firstExplicitCellTextColor(tc) ?? '000000', `cell (${rIdx + 1},${cIdx + 1}) of table '${tableNode.id}'`);
        });
      });
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setColumnWidths': {
      const widths = (raw.widths || op.widths || raw.columnWidths || op.columnWidths || op.params?.widths || op.params?.columnWidths) as unknown;
      if (!Array.isArray(widths) || widths.length === 0) {
        throw new Error(
          `E_VALIDATION: setColumnWidths requires an array of widths in DXA (e.g. widths: [2000, 3000, 3000]) (got ${Array.isArray(widths) ? 'empty array' : typeof widths}). ` +
            `Remedy: pass widths (or columnWidths) as a non-empty array of positive integer twips, e.g. widths:[2000, 3000, 3000].`
        );
      }
      widths.forEach((w: unknown, i: number) => {
        if (typeof w !== 'number' || !Number.isInteger(w) || (w as number) <= 0) {
          throw new Error(
            `E_VALIDATION: setColumnWidths widths[${i}] must be a positive integer in DXA (got ${String(w)}). ` +
              `Remedy: pass widths as positive integer twips, e.g. widths:[2000, 3000, 3000].`
          );
        }
      });
      const pinned = widths as number[];
      let tblGrid = directChildren(tblEl, 'w:tblGrid')[0];
      if (!tblGrid) {
        tblGrid = createWordElement(tblEl.ownerDocument, 'w:tblGrid');
        tblEl.insertBefore(tblGrid, rows[0] || null);
      }
      // Honored exactly: rebuild the grid to exactly the requested entries —
      // never averaged, padded, or rescaled.
      while (tblGrid.firstChild) tblGrid.removeChild(tblGrid.firstChild);
      for (const w of pinned) {
        const gridCol = createWordElement(tblEl.ownerDocument, 'w:gridCol');
        gridCol.setAttribute('w:w', String(w));
        tblGrid.appendChild(gridCol);
      }

      for (const row of rows) {
        const cells = directChildren(row, 'w:tc');
        cells.forEach((tc, colIdx) => {
          const w = pinned[colIdx];
          if (w !== undefined) {
            let tcPr = directChildren(tc, 'w:tcPr')[0];
            if (!tcPr) {
              tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
              tc.insertBefore(tcPr, tc.firstChild);
            }
            let tcW = tcPr.getElementsByTagName('w:tcW')[0];
            if (!tcW) {
              tcW = createWordElement(tblEl.ownerDocument, 'w:tcW');
              tcPr.appendChild(tcW);
            }
            tcW.setAttribute('w:w', String(w));
            tcW.setAttribute('w:type', 'dxa');
            sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
          }
        });
      }
      // Pinning: a fixed layout stops Word from re-averaging (autofit) the
      // explicit widths on open. w:tblLayout sorts per TBLPR order.
      let tblPr = directChildren(tblEl, 'w:tblPr')[0];
      if (!tblPr) {
        tblPr = createWordElement(tblEl.ownerDocument, 'w:tblPr');
        tblEl.insertBefore(tblPr, tblEl.firstChild);
      }
      let tblLayout = tblPr.getElementsByTagName('w:tblLayout')[0];
      if (!tblLayout) {
        tblLayout = createWordElement(tblEl.ownerDocument, 'w:tblLayout');
        tblPr.appendChild(tblLayout);
      }
      tblLayout.setAttribute('w:type', 'fixed');
      sortChildrenBySchema(tblPr, TBLPR_CHILD_ORDER);
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setColumnWidth': {
      // Stays physical/grid-based because it mutates w:tblGrid/w:gridCol definitions
      // and row-level tcW where multi-span cells aggregate width, avoiding Cover Policy A
      // rejections during whole-column operations.
      // Precedence per colIndexAlias: colIndex > col > columnIndex.
      const colIdx = Number(colIndexAlias(raw, op) ?? P.colIndex ?? 0);
      const width = Number(raw.width ?? op.width ?? raw.w ?? op.w ?? 2000);
      const tblGrid = directChildren(tblEl, 'w:tblGrid')[0];
      if (colIndexAlias(raw, op) !== undefined) {
        const columnCount = tblGrid ? directChildren(tblGrid, 'w:gridCol').length : widestRow(rows);
        assertColumnInRange('setColumnWidth', 'colIndex', colIdx, columnCount, `table has ${columnCount} columns`);
      }
      if (tblGrid) {
        const gridCols = directChildren(tblGrid, 'w:gridCol');
        if (gridCols[colIdx]) {
          gridCols[colIdx].setAttribute('w:w', String(width));
        }
      }
      for (const row of rows) {
        const cells = directChildren(row, 'w:tc');
        const tc = cells[colIdx];
        if (tc) {
          let tcPr = directChildren(tc, 'w:tcPr')[0];
          if (!tcPr) {
            tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
            tc.insertBefore(tcPr, tc.firstChild);
          }
          let tcW = tcPr.getElementsByTagName('w:tcW')[0];
          if (!tcW) {
            tcW = createWordElement(tblEl.ownerDocument, 'w:tcW');
            tcPr.appendChild(tcW);
          }
          tcW.setAttribute('w:w', String(width));
          tcW.setAttribute('w:type', 'dxa');
          sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'addColumn': {
      const afterIdx = Number(pick<number>('afterIndex', undefined) ?? (rows[0] ? directChildren(rows[0], 'w:tc').length - 1 : 0));
      if (pick<number>('afterIndex', undefined) !== undefined) {
        const columnCount = widestRow(rows);
        assertColumnInRange('addColumn', 'afterIndex', afterIdx, columnCount, `row has ${columnCount} logical columns`);
      }
      const width = Number(pick<number>('width', undefined) ?? 2000);
      const tblGrid = directChildren(tblEl, 'w:tblGrid')[0];
      if (tblGrid) {
        const gridCols = directChildren(tblGrid, 'w:gridCol');
        const newCol = createWordElement(tblEl.ownerDocument, 'w:gridCol');
        newCol.setAttribute('w:w', String(width));
        if (gridCols[afterIdx] && gridCols[afterIdx].nextSibling) {
          tblGrid.insertBefore(newCol, gridCols[afterIdx].nextSibling);
        } else {
          tblGrid.appendChild(newCol);
        }
      }
      for (const row of rows) {
        const cells = directChildren(row, 'w:tc');
        const refCell = cells[afterIdx] || cells[cells.length - 1];
        const newCell = createWordElement(tblEl.ownerDocument, 'w:tc');
        const tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
        const tcW = createWordElement(tblEl.ownerDocument, 'w:tcW');
        tcW.setAttribute('w:w', String(width));
        tcW.setAttribute('w:type', 'dxa');
        tcPr.appendChild(tcW);
        newCell.appendChild(tcPr);
        newCell.appendChild(createWordElement(tblEl.ownerDocument, 'w:p'));

        if (refCell && refCell.nextSibling) {
          row.insertBefore(newCell, refCell.nextSibling);
        } else {
          row.appendChild(newCell);
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'deleteColumn': {
      // Precedence: index > colIndex > col > columnIndex (nested params win
      // over flat within the same name; col order per colIndexAlias).
      const indexSrc = raw.index ?? op.index ?? colIndexAlias(raw, op);
      if (indexSrc === undefined) {
        throw new Error(
          `E_VALIDATION: deleteColumn requires a column index. Remedy: pass index (or colIndex) as an integer in [1, columns], counting from 1.`
        );
      }
      const targetCol = typeof indexSrc === 'number' || (typeof indexSrc === 'string' && indexSrc.trim() !== '') ? Number(indexSrc) : NaN;
      if (!Number.isInteger(targetCol)) {
        throw new Error(
          `E_VALIDATION: deleteColumn column index ${shown(targetCol)} out of bounds. Remedy: pass integer index.`
        );
      }
      {
        const maxLogical = widestRow(rows);
        if (targetCol < 0 || targetCol >= maxLogical) {
          throw new Error(
            `E_VALIDATION: deleteColumn column index ${shown(targetCol)} out of bounds (row has ${maxLogical} logical columns). Remedy: pass index in [1, ${maxLogical}], counting from 1.`
          );
        }
      }
      // OP-10: map logical column -> physical w:tc per row via gridSpan/hMerge.
      // A cell with gridSpan=N covers N logical columns (shrink, don't remove);
      // hMerge continue cells each occupy one logical column (remove one cell).
      let gridRemoved = false;
      const removeOneGridCol = () => {
        if (gridRemoved) return;
        const grid = directChildren(tblEl, 'w:tblGrid')[0];
        if (grid) {
          const gridCols = directChildren(grid, 'w:gridCol');
          if (gridCols[targetCol]) grid.removeChild(gridCols[targetCol]);
          else if (gridCols.length > 0) grid.removeChild(gridCols[gridCols.length - 1]);
        }
        gridRemoved = true;
      };
      for (const row of rows) {
        const cells = directChildren(row, 'w:tc');
        let logical = 0;
        let removed = false;
        for (let ci = 0; ci < cells.length; ci++) {
          const tc = cells[ci];
          const hMerge = getHMergeVal(tc);
          const isContinue = hMerge !== null && hMerge !== 'restart';
          const span = isContinue ? 1 : getCellGridSpan(tc);
          if (targetCol >= logical && targetCol < logical + span) {
            if (span > 1) {
              // Logical column lies inside a gridSpan: shrink the span, keep the cell.
              const tcPr = directChildren(tc, 'w:tcPr')[0];
              const gsEl = tcPr?.getElementsByTagName('w:gridSpan')[0] as Element | undefined;
              if (gsEl && tcPr) {
                if (span - 1 <= 1) tcPr.removeChild(gsEl);
                else gsEl.setAttribute('w:val', String(span - 1));
              }
              removeOneGridCol();
            } else {
              // span == 1: remove the physical cell. If it was a merge restart,
              // promote the following continue cell so the merge stays valid.
              const wasRestart = hMerge === 'restart';
              row.removeChild(tc);
              if (wasRestart) {
                const next = directChildren(row, 'w:tc')[ci];
                if (next && getHMergeVal(next) === 'continue') {
                  const nextTcPr = directChildren(next, 'w:tcPr')[0];
                  const nextHm = nextTcPr?.getElementsByTagName('w:hMerge')[0] as Element | undefined;
                  if (nextHm && nextTcPr) {
                    const after = directChildren(row, 'w:tc')[ci + 1];
                    if (after && getHMergeVal(after) === 'continue') {
                      nextHm.setAttribute('w:val', 'restart');
                    } else {
                      nextTcPr.removeChild(nextHm);
                    }
                  }
                }
              }
              removeOneGridCol();
            }
            removed = true;
            break;
          }
          logical += span;
        }
        if (!removed) {
          throw new Error(
            `E_VALIDATION: deleteColumn column index ${shown(targetCol)} out of bounds (row has ${logical} logical columns). Remedy: pass index in [1, ${logical}], counting from 1.`
          );
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setBorder': {
      const borderType = String(pick<string>('border', undefined) || 'all');
      const color = String(pick<string>('color', undefined) || 'D1D5DB').replace(/^#/, '');
      const rawSz = pick<string | number>('sz', undefined) || (raw.size as string | number | undefined) || 4;
      const szNum = Number(rawSz);
      if (!Number.isInteger(szNum) || szNum < 1) {
        throw toError(
          'E_VALIDATION',
          `setBorder sz must be a whole number of eighths of a point (got ${String(rawSz)}).`,
          'Pass sz:4 for a 0.5 pt line or sz:8 for 1 pt.'
        );
      }
      const sz = String(szNum);
      const val = String(raw.val || 'single');

      let tblPr = directChildren(tblEl, 'w:tblPr')[0];
      if (!tblPr) {
        tblPr = createWordElement(tblEl.ownerDocument, 'w:tblPr');
        tblEl.insertBefore(tblPr, tblEl.firstChild);
      }
      let tblBorders = tblPr.getElementsByTagName('w:tblBorders')[0];
      if (!tblBorders) {
        tblBorders = createWordElement(tblEl.ownerDocument, 'w:tblBorders');
        tblPr.appendChild(tblBorders);
      }

      const borderTags = borderType === 'all'
        ? ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV']
        : [`w:${borderType}`];

      for (const tag of borderTags) {
        let bEl = tblBorders.getElementsByTagName(tag)[0];
        if (!bEl) {
          bEl = createWordElement(tblEl.ownerDocument, tag);
          tblBorders.appendChild(bEl);
        }
        bEl.setAttribute('w:val', val);
        bEl.setAttribute('w:sz', sz);
        bEl.setAttribute('w:space', '0');
        bEl.setAttribute('w:color', color);
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'mergeCells': {
      // B-1 flat parity: accept fromRow/fromCol/toRow/toCol flat on the op
      // (mirrors the P pick() pattern); toRow/toCol default to fromRow/fromCol.
      const fromRow = Number(raw.fromRow ?? op.fromRow ?? 0);
      const fromCol = Number(raw.fromCol ?? op.fromCol ?? 0);
      const toRow = Number(raw.toRow ?? op.toRow ?? fromRow);
      const toCol = Number(raw.toCol ?? op.toCol ?? fromCol);
      for (const [name, v] of [['fromRow', fromRow], ['fromCol', fromCol], ['toRow', toRow], ['toCol', toCol]] as const) {
        if (!Number.isInteger(v)) {
          throw new Error(
            `E_VALIDATION: mergeCells requires integer indexes (got ${name}=${String(raw[name] ?? op[name])}). ` +
              `Remedy: pass integer fromRow/fromCol/toRow/toCol (e.g. { action:'table', target:'<tableId>', operation:'mergeCells', fromRow:1, fromCol:1, toRow:2, toCol:2 }).`
          );
        }
      }

      const columnCount = widestRow(rows);
      assertRowInRange('mergeCells', 'fromRow', fromRow, rows.length);
      assertRowInRange('mergeCells', 'toRow', toRow, rows.length);
      assertColumnInRange('mergeCells', 'fromCol', fromCol, columnCount, `row has ${columnCount} logical columns`);
      assertColumnInRange('mergeCells', 'toCol', toCol, columnCount, `row has ${columnCount} logical columns`);
      if (toRow < fromRow || toCol < fromCol) {
        throw new Error(
          `E_VALIDATION: mergeCells range runs backwards (from ${shown(fromRow)},${shown(fromCol)} to ${shown(toRow)},${shown(toCol)}). Remedy: pass fromRow <= toRow and fromCol <= toCol, counting from 1.`
        );
      }

      for (let r = fromRow; r <= toRow; r++) {
        const row = rows[r];
        if (!row) continue;
        const cells = directChildren(row, 'w:tc');
        for (let c = fromCol; c <= toCol; c++) {
          const tc = cells[c];
          if (!tc) continue;
          let tcPr = directChildren(tc, 'w:tcPr')[0];
          if (!tcPr) {
            tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
            tc.insertBefore(tcPr, tc.firstChild);
          }
          if (toCol > fromCol) {
            let hMerge = tcPr.getElementsByTagName('w:hMerge')[0];
            if (!hMerge) {
              hMerge = createWordElement(tblEl.ownerDocument, 'w:hMerge');
              tcPr.appendChild(hMerge);
            }
            hMerge.setAttribute('w:val', c === fromCol ? 'restart' : 'continue');
          }
          if (toRow > fromRow) {
            let vMerge = tcPr.getElementsByTagName('w:vMerge')[0];
            if (!vMerge) {
              vMerge = createWordElement(tblEl.ownerDocument, 'w:vMerge');
              tcPr.appendChild(vMerge);
            }
            vMerge.setAttribute('w:val', r === fromRow ? 'restart' : 'continue');
          }
          sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'splitCell': {
      // Stays physical because its purpose is to directly strip w:hMerge/w:vMerge
      // from specific <w:tc> DOM nodes (including continuation cells), where Cover Policy A
      // rejection would prevent unmerging. Bounds still mirror resolveTableCell
      // (logical width) but cover cells are allowed here; only OOB rejects.
      const { rowIndex, colIndex } = P;
      if (!Number.isInteger(rowIndex) || rowIndex < 0) {
        throw new Error(
          `E_VALIDATION: splitCell row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
        );
      }
      const tr = rows[rowIndex];
      if (!tr) {
        throw new Error(
          `E_VALIDATION: splitCell row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
        );
      }
      if (!Number.isInteger(colIndex) || colIndex < 0) {
        throw new Error(
          `E_VALIDATION: splitCell column index ${shown(colIndex)} out of bounds. Remedy: pass colIndex >= 1 (columns are numbered from 1).`
        );
      }
      const { logicalWidth, tc: logicalTc } = findLogicalTableCell(tr, colIndex);
      if (colIndex >= logicalWidth) {
        throw new Error(
          `E_VALIDATION: splitCell column index ${shown(colIndex)} out of bounds (row has ${logicalWidth} logical columns). Remedy: pass colIndex in [1, ${logicalWidth}], counting from 1.`
        );
      }
      {
        const tc = directChildren(tr, 'w:tc')[colIndex] ?? logicalTc;
        if (tc) {
          const tcPr = directChildren(tc, 'w:tcPr')[0];
          if (tcPr) {
            const hMerge = tcPr.getElementsByTagName('w:hMerge')[0];
            if (hMerge) tcPr.removeChild(hMerge);
            const vMerge = tcPr.getElementsByTagName('w:vMerge')[0];
            if (vMerge) tcPr.removeChild(vMerge);
          }
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setHeaderRepeat': {
      const targetRowIdx = Number(rowIndexAlias(raw, op) ?? P.rowIndex ?? 0);
      if (rowIndexAlias(raw, op) !== undefined) assertRowInRange('setHeaderRepeat', 'rowIndex', targetRowIdx, rows.length);
      const tr = rows[targetRowIdx];
      if (tr && tblEl.ownerDocument) {
        let trPr = directChildren(tr, 'w:trPr')[0];
        if (!trPr) {
          trPr = createWordElement(tblEl.ownerDocument, 'w:trPr');
          tr.insertBefore(trPr, tr.firstChild);
        }
        let tblHeader = trPr.getElementsByTagName('w:tblHeader')[0];
        if (!tblHeader) {
          tblHeader = createWordElement(tblEl.ownerDocument, 'w:tblHeader');
          trPr.appendChild(tblHeader);
          sortChildrenBySchema(trPr, TRPR_CHILD_ORDER);
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setCantSplit': {
      for (const tr of rows) {
        if (tblEl.ownerDocument) {
          let trPr = directChildren(tr, 'w:trPr')[0];
          if (!trPr) {
            trPr = createWordElement(tblEl.ownerDocument, 'w:trPr');
            tr.insertBefore(trPr, tr.firstChild);
          }
          let cantSplit = trPr.getElementsByTagName('w:cantSplit')[0];
          if (!cantSplit) {
            cantSplit = createWordElement(tblEl.ownerDocument, 'w:cantSplit');
            trPr.appendChild(cantSplit);
            sortChildrenBySchema(trPr, TRPR_CHILD_ORDER);
          }
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setVerticalAlign': {
      const align = String(pick<string>('vAlign', undefined) || pick<string>('alignment', undefined) || 'center').toLowerCase();
      const rowSrc = rowIndexAlias(raw, op);
      const colSrc = colIndexAlias(raw, op);
      const targetRow = rowSrc !== undefined ? Number(rowSrc) : undefined;
      const targetCol = colSrc !== undefined ? Number(colSrc) : undefined;
      if (targetRow !== undefined && !Number.isInteger(targetRow)) {
        throw new Error(
          `E_VALIDATION: setVerticalAlign requires integer indexes (got rowIndex=${String(rowSrc)}). ` +
            `Remedy: pass integer rowIndex/row and colIndex/col.`
        );
      }
      if (targetRow !== undefined && (targetRow < 0 || !rows[targetRow])) {
        throw new Error(
          `E_VALIDATION: setVerticalAlign row index ${shown(targetRow)} out of bounds (table has ${rows.length} rows). ` +
            `Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
        );
      }
      if (targetCol !== undefined && !Number.isInteger(targetCol)) {
        throw new Error(
          `E_VALIDATION: setVerticalAlign requires integer indexes (got colIndex=${String(colSrc)}). ` +
            `Remedy: pass integer rowIndex/row and colIndex/col.`
        );
      }
      if (targetRow === undefined && targetCol !== undefined) {
        let maxLogical = 0;
        for (const tr of rows) {
          const probe = findLogicalTableCell(tr, 0);
          if (probe.logicalWidth > maxLogical) maxLogical = probe.logicalWidth;
        }
        if ((targetCol as number) < 0 || (targetCol as number) >= maxLogical) {
          throw new Error(
            `E_VALIDATION: setVerticalAlign column index ${shown(targetCol)} out of bounds (table has ${maxLogical} logical columns). ` +
              `Remedy: pass colIndex in [1, ${maxLogical}], counting from 1.`
          );
        }
      }

      const applyVAlign = (tc: Element) => {
        let tcPr = directChildren(tc, 'w:tcPr')[0];
        if (!tcPr && tblEl.ownerDocument) {
          tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
          tc.insertBefore(tcPr, tc.firstChild);
        }
        if (tcPr && tblEl.ownerDocument) {
          let vAlignEl = tcPr.getElementsByTagName('w:vAlign')[0];
          if (!vAlignEl) {
            vAlignEl = createWordElement(tblEl.ownerDocument, 'w:vAlign');
            tcPr.appendChild(vAlignEl);
          }
          vAlignEl.setAttribute('w:val', align);
          sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
        }
      };

      for (let r = 0; r < rows.length; r++) {
        if (targetRow !== undefined && r !== targetRow) continue;
        const tr = rows[r];
        if (targetCol !== undefined) {
          if (targetRow !== undefined) {
            const tc = resolveTableCell(tr, targetCol, 'setVerticalAlign');
            applyVAlign(tc);
          } else {
            const { tc, isCover } = findLogicalTableCell(tr, targetCol);
            if (tc && !isCover) {
              applyVAlign(tc);
            }
          }
        } else {
          const cells = directChildren(tr, 'w:tc');
          for (let c = 0; c < cells.length; c++) {
            applyVAlign(cells[c]);
          }
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setCellShading': {
      // E3: single-cell fill (flat + params shapes, mergeCells precedent).
      // Unlike setShading (row/zebra-wide), this touches exactly one cell.
      // row/col are accepted as aliases for rowIndex/colIndex (flat or
      // nested in params). Coordinates are REQUIRED — no (0,0) default.
      // Precedence per helpers: rowIndex > row; colIndex > col > columnIndex.
      const rowSrc = rowIndexAlias(raw, op);
      const colSrc = colIndexAlias(raw, op);
      if (rowSrc === undefined || colSrc === undefined) {
        throw new Error(
          `E_VALIDATION: setCellShading requires explicit rowIndex/row and colIndex/col (got row=${String(rowSrc)} col=${String(colSrc)}). ` +
            `Remedy: pass integer rowIndex/colIndex or row/col (e.g. { action:'table', target:'<tableId>', operation:'setCellShading', rowIndex:1, colIndex:2, fill:'FF0000' }).`
        );
      }
      const rowIndex = Number(rowSrc);
      const colIndex = Number(colSrc);
      for (const [name, v] of [['rowIndex', rowIndex], ['colIndex', colIndex]] as const) {
        if (!Number.isInteger(v)) {
          throw new Error(
            `E_VALIDATION: setCellShading requires integer indexes (got ${name}=${String(name === 'rowIndex' ? rowSrc : colSrc)}). ` +
              `Remedy: pass integer rowIndex/colIndex or row/col (e.g. { action:'table', target:'<tableId>', operation:'setCellShading', rowIndex:1, colIndex:2, fill:'FF0000' }).`
          );
        }
      }
      const tr = rows[rowIndex];
      if (!tr) {
        throw new Error(
          `E_VALIDATION: setCellShading row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). ` +
            `Remedy: pass rowIndex in [1, ${rows.length}], counting from 1 (see inspect_document table dimensions).`
        );
      }
      const tc = resolveTableCell(tr, colIndex, 'setCellShading');
      const fillInput = raw.fill ?? op.fill;
      if (fillInput === undefined) {
        throw new Error(
          `E_VALIDATION: setCellShading requires fill (got no fill). ` +
            `Remedy: pass fill as 6 hex digits without '#', e.g. fill:'FF0000'.`
        );
      }
      const fill = String(fillInput).replace(/^#/, '');
      if (!/^[0-9A-Fa-f]{6}$/.test(fill)) {
        throw new Error(
          `E_VALIDATION: setCellShading requires a 6-digit hex fill with optional leading '#' (got fill='${String(fillInput)}'). ` +
            `Remedy: pass fill as 6 hex digits, e.g. fill:'FF0000' or fill:'#FF0000'.`
        );
      }
      const normFill = fill.toUpperCase();
      let tcPr = directChildren(tc, 'w:tcPr')[0];
      if (!tcPr) {
        tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
        tc.insertBefore(tcPr, tc.firstChild);
      }
      let shd = tcPr.getElementsByTagName('w:shd')[0];
      if (!shd) {
        shd = createWordElement(tblEl.ownerDocument, 'w:shd');
        tcPr.appendChild(shd);
      }
      shd.setAttribute('w:val', 'clear');
      shd.setAttribute('w:color', 'auto');
      shd.setAttribute('w:fill', normFill);
      sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
      // Taste defaulting (explicit > guide > built-in): an explicit
      // textColor always wins; otherwise the guide header text applies when
      // the fresh fill is too dark for dark text (WCAG-gated, computed).
      const effectiveText = explicitTextColor ?? guideTextForFill(normFill);
      if (effectiveText && tblEl.ownerDocument) {
        paintCellTextColor(tc, tblEl.ownerDocument, effectiveText);
      }
      // Contrast guardrail: advisory only, never fails the op.
      warnIfLowContrast(warnings, warnedContrastPairs, normFill, effectiveText ?? firstExplicitCellTextColor(tc) ?? '000000', `cell (${shown(rowIndex)},${shown(colIndex)}) of table '${tableNode.id}'`);
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setRowHeight': {
      // E3: explicit row height in DXA/twips (flat + params shapes).
      // row is accepted as an alias for rowIndex (flat or nested in
      // params). The row coordinate is REQUIRED — no row-0 default.
      // Precedence per rowIndexAlias: rowIndex > row.
      const rowSrc = rowIndexAlias(raw, op);
      if (rowSrc === undefined) {
        throw new Error(
          `E_VALIDATION: setRowHeight requires an explicit rowIndex/row (got rowIndex=undefined). ` +
            `Remedy: pass integer rowIndex or row (e.g. { action:'table', target:'<tableId>', operation:'setRowHeight', rowIndex:1, height:400, rule:'exact' }).`
        );
      }
      const rowIndex = Number(rowSrc);
      if (!Number.isInteger(rowIndex)) {
        throw new Error(
          `E_VALIDATION: setRowHeight requires an integer rowIndex (got rowIndex=${String(rowSrc)}). ` +
            `Remedy: pass integer rowIndex or row (e.g. { action:'table', target:'<tableId>', operation:'setRowHeight', rowIndex:1, height:400, rule:'exact' }).`
        );
      }
      const tr = rows[rowIndex];
      if (!tr) {
        throw new Error(
          `E_VALIDATION: setRowHeight row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). ` +
            `Remedy: pass rowIndex in [1, ${rows.length}], counting from 1 (see inspect_document table dimensions).`
        );
      }
      const height = Number(raw.height ?? op.height);
      if (!Number.isInteger(height) || height <= 0) {
        throw new Error(
          `E_VALIDATION: setRowHeight requires height as a positive integer in DXA/twips (got height=${String(raw.height ?? op.height)}). ` +
            `Remedy: pass height as integer twips, e.g. height:400 with rule:'atLeast' (default), 'exact', or 'auto'.`
        );
      }
      const rule = String(raw.rule ?? op.rule ?? 'atLeast');
      if (rule !== 'atLeast' && rule !== 'exact' && rule !== 'auto') {
        throw new Error(
          `E_VALIDATION: setRowHeight rule must be 'atLeast'|'exact'|'auto' (got rule='${rule}'). ` +
            `Remedy: pass rule:'atLeast' (default), 'exact', or 'auto'.`
        );
      }
      let trPr = directChildren(tr, 'w:trPr')[0];
      if (!trPr) {
        trPr = createWordElement(tblEl.ownerDocument, 'w:trPr');
        tr.insertBefore(trPr, tr.firstChild);
      }
      let trHeight = trPr.getElementsByTagName('w:trHeight')[0];
      if (!trHeight) {
        trHeight = createWordElement(tblEl.ownerDocument, 'w:trHeight');
        trPr.appendChild(trHeight);
      }
      trHeight.setAttribute('w:val', String(height));
      trHeight.setAttribute('w:hRule', rule);
      sortChildrenBySchema(trPr, TRPR_CHILD_ORDER);
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setCellAlignment': {
      // Per-cell / per-column / per-row / whole-table horizontal alignment.
      // Writes w:jc on each targeted cell's OWN paragraphs (direct w:p
      // children — never into a nested table). Logical-column resolution
      // reuses resolveTableCell / findLogicalTableCell, so merged tables
      // behave exactly like setCellShading (cover cells reject loudly when
      // addressed directly, and are skipped for column-wide application).
      const alignRaw = raw.align ?? raw.alignment ?? op.align ?? op.alignment;
      if (alignRaw === undefined || alignRaw === null || String(alignRaw).trim() === '') {
        throw new Error(
          `E_VALIDATION: setCellAlignment requires align/alignment (got ${String(alignRaw)}). ` +
            `Remedy: pass align as 'left', 'center', or 'right' (e.g. { action:'table', target:'<tableId>', operation:'setCellAlignment', rowIndex:1, colIndex:2, align:'center' }).`
        );
      }
      const align = String(alignRaw).toLowerCase();
      if (align !== 'left' && align !== 'center' && align !== 'right') {
        throw new Error(
          `E_VALIDATION: setCellAlignment align must be 'left'|'center'|'right' (got '${String(alignRaw)}'). ` +
            `Remedy: pass align:'left', align:'center', or align:'right'.`
        );
      }
      const rowSrc = rowIndexAlias(raw, op);
      const colSrc = colIndexAlias(raw, op);
      let rowIndex: number | undefined;
      let colIndex: number | undefined;
      if (rowSrc !== undefined) {
        rowIndex = Number(rowSrc);
        if (!Number.isInteger(rowIndex)) {
          throw new Error(
            `E_VALIDATION: setCellAlignment requires integer indexes (got rowIndex=${String(rowSrc)}). ` +
              `Remedy: pass integer rowIndex/row and colIndex/col.`
          );
        }
        if (!rows[rowIndex]) {
          throw new Error(
            `E_VALIDATION: setCellAlignment row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). ` +
              `Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
          );
        }
      }
      if (colSrc !== undefined) {
        colIndex = Number(colSrc);
        if (!Number.isInteger(colIndex)) {
          throw new Error(
            `E_VALIDATION: setCellAlignment requires integer indexes (got colIndex=${String(colSrc)}). ` +
              `Remedy: pass integer rowIndex/row and colIndex/col.`
          );
        }
      }
      const applyAlign = (tc: Element) => {
        let ps = directChildren(tc, 'w:p');
        if (ps.length === 0) {
          const p = createWordElement(tblEl.ownerDocument, 'w:p');
          tc.appendChild(p);
          ps = [p];
        }
        for (const p of ps) {
          let pPr = directChildren(p, 'w:pPr')[0];
          if (!pPr) {
            pPr = createWordElement(tblEl.ownerDocument, 'w:pPr');
            p.insertBefore(pPr, p.firstChild);
          }
          let jc = pPr.getElementsByTagName('w:jc')[0];
          if (!jc) {
            jc = createWordElement(tblEl.ownerDocument, 'w:jc');
            pPr.appendChild(jc);
          }
          jc.setAttribute('w:val', align);
          sortChildrenBySchema(pPr, PPR_CHILD_ORDER);
        }
      };
      if (rowIndex !== undefined && colIndex !== undefined) {
        applyAlign(resolveTableCell(rows[rowIndex], colIndex, 'setCellAlignment'));
      } else if (rowIndex !== undefined) {
        for (const tc of directChildren(rows[rowIndex], 'w:tc')) applyAlign(tc);
      } else if (colIndex !== undefined) {
        let maxLogical = 0;
        for (const tr of rows) {
          const probe = findLogicalTableCell(tr, 0);
          if (probe.logicalWidth > maxLogical) maxLogical = probe.logicalWidth;
        }
        if (colIndex < 0 || colIndex >= maxLogical) {
          throw new Error(
            `E_VALIDATION: setCellAlignment column index ${shown(colIndex)} out of bounds (table has ${maxLogical} logical columns). ` +
              `Remedy: pass colIndex in [1, ${maxLogical}], counting from 1.`
          );
        }
        for (const tr of rows) {
          const { tc, isCover } = findLogicalTableCell(tr, colIndex);
          if (tc && !isCover) applyAlign(tc);
        }
      } else {
        for (const tr of rows) {
          for (const tc of directChildren(tr, 'w:tc')) applyAlign(tc);
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setTableAlignment': {
      if (!targetId) {
        throw toError(
          'E_VALIDATION',
          'Table operation "setTableAlignment" requires a target table node id.',
          'Pass target: "<tableNodeId>" in the operation payload.'
        );
      }
      const rawAlign = (
        (typeof raw === 'object' && raw !== null
          ? (raw as Record<string, unknown>).align ?? (raw as Record<string, unknown>).alignment
          : undefined) ??
        op.align ??
        op.alignment
      );
      const align = typeof rawAlign === 'string' ? rawAlign.toLowerCase() : '';
      if (!align || !['left', 'center', 'right'].includes(align)) {
        throw toError(
          'E_VALIDATION',
          `Invalid table alignment "${rawAlign}". Expected "left", "center", or "right".`,
          'Pass align: "left" | "center" | "right" in the operation payload.'
        );
      }

      const doc = tblEl.ownerDocument;
      let tblPr = directChildren(tblEl, 'w:tblPr')[0];
      if (!tblPr) {
        tblPr = createWordElement(doc, 'w:tblPr');
        tblEl.insertBefore(tblPr, tblEl.firstChild);
      }

      let jc = directChildren(tblPr, 'w:jc')[0];
      if (!jc) {
        jc = createWordElement(doc, 'w:jc');
        tblPr.appendChild(jc);
      }
      jc.setAttribute('w:val', align);
      sortChildrenBySchema(tblPr, TBLPR_CHILD_ORDER);
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setCellPadding': {
      // Per-cell / per-column / per-row / whole-table cell padding (w:tcMar).
      // padding is a uniform DXA number (all four sides) or a per-side
      // object { top?, left?, bottom?, right? } (flat or nested in params);
      // object sides left unspecified keep their existing value. Targeting
      // reuses resolveTableCell / findLogicalTableCell like setCellAlignment.
      const padRaw = raw.padding ?? op.padding;
      const sides: { top?: number; left?: number; bottom?: number; right?: number } = {};
      if (typeof padRaw === 'number') {
        if (!Number.isInteger(padRaw) || padRaw < 0) {
          throw new Error(
            `E_VALIDATION: setCellPadding padding must be a non-negative integer in DXA (got ${String(padRaw)}). ` +
              `Remedy: pass padding as integer twips, e.g. padding:100, or per-side padding:{ top:100, left:150 }.`
          );
        }
        sides.top = sides.left = sides.bottom = sides.right = padRaw;
      } else if (padRaw !== null && typeof padRaw === 'object' && !Array.isArray(padRaw)) {
        for (const side of ['top', 'left', 'bottom', 'right'] as const) {
          const v = (padRaw as Record<string, unknown>)[side];
          if (v === undefined) continue;
          const n = typeof v === 'number' ? v : Number(v);
          if (!Number.isInteger(n) || (n as number) < 0) {
            throw new Error(
              `E_VALIDATION: setCellPadding padding.${side} must be a non-negative integer in DXA (got ${String(v)}). ` +
                `Remedy: pass padding as integer twips, e.g. padding:{ ${side}:100 }.`
            );
          }
          sides[side] = n as number;
        }
        if (Object.keys(sides).length === 0) {
          throw new Error(
            `E_VALIDATION: setCellPadding padding object specifies no sides (got {}). ` +
              `Remedy: pass at least one of top|left|bottom|right, e.g. padding:{ top:100, left:150 }.`
          );
        }
      } else {
        throw new Error(
          `E_VALIDATION: setCellPadding requires padding (got ${padRaw === undefined ? 'no padding' : String(padRaw)}). ` +
            `Remedy: pass padding as a non-negative integer (uniform DXA) or a per-side object, e.g. padding:100 or padding:{ top:100, left:150, bottom:100, right:150 }.`
        );
      }
      const rowSrc = rowIndexAlias(raw, op);
      const colSrc = colIndexAlias(raw, op);
      let rowIndex: number | undefined;
      let colIndex: number | undefined;
      if (rowSrc !== undefined) {
        rowIndex = Number(rowSrc);
        if (!Number.isInteger(rowIndex)) {
          throw new Error(
            `E_VALIDATION: setCellPadding requires integer indexes (got rowIndex=${String(rowSrc)}). ` +
              `Remedy: pass integer rowIndex/row and colIndex/col.`
          );
        }
        if (!rows[rowIndex]) {
          throw new Error(
            `E_VALIDATION: setCellPadding row index ${shown(rowIndex)} out of bounds (table has ${rows.length} rows). ` +
              `Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
          );
        }
      }
      if (colSrc !== undefined) {
        colIndex = Number(colSrc);
        if (!Number.isInteger(colIndex)) {
          throw new Error(
            `E_VALIDATION: setCellPadding requires integer indexes (got colIndex=${String(colSrc)}). ` +
              `Remedy: pass integer rowIndex/row and colIndex/col.`
          );
        }
      }
      const applyPad = (tc: Element) => {
        let tcPr = directChildren(tc, 'w:tcPr')[0];
        if (!tcPr) {
          tcPr = createWordElement(tblEl.ownerDocument, 'w:tcPr');
          tc.insertBefore(tcPr, tc.firstChild);
        }
        let tcMar = tcPr.getElementsByTagName('w:tcMar')[0];
        if (!tcMar) {
          tcMar = createWordElement(tblEl.ownerDocument, 'w:tcMar');
          tcPr.appendChild(tcMar);
        }
        for (const side of ['top', 'left', 'bottom', 'right'] as const) {
          const v = sides[side];
          if (v === undefined) continue;
          let el = directChildren(tcMar as Element, `w:${side}`)[0];
          if (!el) {
            el = createWordElement(tblEl.ownerDocument, `w:${side}`);
            tcMar.appendChild(el);
          }
          el.setAttribute('w:w', String(v));
          el.setAttribute('w:type', 'dxa');
        }
        sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
      };
      if (rowIndex !== undefined && colIndex !== undefined) {
        applyPad(resolveTableCell(rows[rowIndex], colIndex, 'setCellPadding'));
      } else if (rowIndex !== undefined) {
        for (const tc of directChildren(rows[rowIndex], 'w:tc')) applyPad(tc);
      } else if (colIndex !== undefined) {
        let maxLogical = 0;
        for (const tr of rows) {
          const probe = findLogicalTableCell(tr, 0);
          if (probe.logicalWidth > maxLogical) maxLogical = probe.logicalWidth;
        }
        if (colIndex < 0 || colIndex >= maxLogical) {
          throw new Error(
            `E_VALIDATION: setCellPadding column index ${shown(colIndex)} out of bounds (table has ${maxLogical} logical columns). ` +
              `Remedy: pass colIndex in [1, ${maxLogical}], counting from 1.`
          );
        }
        for (const tr of rows) {
          const { tc, isCover } = findLogicalTableCell(tr, colIndex);
          if (tc && !isCover) applyPad(tc);
        }
      } else {
        for (const tr of rows) {
          for (const tc of directChildren(tr, 'w:tc')) applyPad(tc);
        }
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setCellTextColor': {
      // Header-capable text color (row 0 = header row): writes w:color on the
      // runs of each targeted cell's OWN paragraphs (direct w:p children —
      // never into a nested table). Targeting mirrors setCellAlignment:
      // cell (row+col), row, column, or whole table when both omitted.
      const colorInput = raw.color ?? op.color;
      if (colorInput === undefined || colorInput === null || String(colorInput).trim() === '') {
        throw new Error(
          `E_VALIDATION: setCellTextColor requires color (got no color). ` +
            `Remedy: pass color as 6 hex digits without '#', e.g. color:'FFFFFF' for white header text.`
        );
      }
      const color = String(colorInput).replace(/^#/, '');
      if (!/^[0-9A-Fa-f]{6}$/.test(color)) {
        throw new Error(
          `E_VALIDATION: setCellTextColor requires a 6-digit hex color with optional leading '#' (got color='${String(colorInput)}'). ` +
            `Remedy: pass color as 6 hex digits, e.g. color:'FFFFFF' or color:'#FFFFFF'.`
        );
      }
      const normColor = color.toUpperCase();
      const cellFillOf = (tc: Element): string => {
        const tcPr = directChildren(tc, 'w:tcPr')[0];
        const shd = tcPr ? directChildren(tcPr, 'w:shd')[0] : undefined;
        const v = shd?.getAttribute('w:fill');
        return v && /^[0-9A-Fa-f]{6}$/.test(v) ? v.toUpperCase() : 'FFFFFF';
      };
      const applyTextColor = (tc: Element, rIdx: number, cIdx: number) => {
        let ps = directChildren(tc, 'w:p');
        if (ps.length === 0) {
          const p = createWordElement(tblEl.ownerDocument, 'w:p');
          tc.appendChild(p);
          ps = [p];
        }
        for (const p of ps) {
          const runs: Element[] = [...directChildren(p, 'w:r')];
          for (const tag of ['w:hyperlink', 'w:ins', 'w:del']) {
            for (const wrap of directChildren(p, tag)) {
              const wRuns = wrap.getElementsByTagName('w:r');
              for (let i = 0; i < wRuns.length; i++) runs.push(wRuns[i] as Element);
            }
          }
          for (const r of runs) {
            let rPr = directChildren(r, 'w:rPr')[0];
            if (!rPr && tblEl.ownerDocument) {
              rPr = createWordElement(tblEl.ownerDocument, 'w:rPr');
              r.insertBefore(rPr, r.firstChild);
            }
            if (rPr && tblEl.ownerDocument) {
              let colorEl = rPr.getElementsByTagName('w:color')[0];
              if (!colorEl) {
                colorEl = createWordElement(tblEl.ownerDocument, 'w:color');
                rPr.appendChild(colorEl);
              }
              colorEl.setAttribute('w:val', normColor);
              sortChildrenBySchema(rPr, RPR_CHILD_ORDER);
            }
          }
        }
        // Contrast guardrail: advisory only, never fails the op.
        warnIfLowContrast(warnings, warnedContrastPairs, cellFillOf(tc), normColor, `cell (${rIdx + 1},${cIdx + 1}) of table '${tableNode.id}'`);
      };
      const tRowSrc = rowIndexAlias(raw, op);
      const tColSrc = colIndexAlias(raw, op);
      let tRowIndex: number | undefined;
      let tColIndex: number | undefined;
      if (tRowSrc !== undefined) {
        tRowIndex = Number(tRowSrc);
        if (!Number.isInteger(tRowIndex)) {
          throw new Error(
            `E_VALIDATION: setCellTextColor requires integer indexes (got rowIndex=${String(tRowSrc)}). ` +
              `Remedy: pass integer rowIndex/row and colIndex/col.`
          );
        }
        if (!rows[tRowIndex]) {
          throw new Error(
            `E_VALIDATION: setCellTextColor row index ${shown(tRowIndex)} out of bounds (table has ${rows.length} rows). ` +
              `Remedy: pass rowIndex in [1, ${rows.length}], counting from 1.`
          );
        }
      }
      if (tColSrc !== undefined) {
        tColIndex = Number(tColSrc);
        if (!Number.isInteger(tColIndex)) {
          throw new Error(
            `E_VALIDATION: setCellTextColor requires integer indexes (got colIndex=${String(tColSrc)}). ` +
              `Remedy: pass integer rowIndex/row and colIndex/col.`
          );
        }
      }
      if (tRowIndex !== undefined && tColIndex !== undefined) {
        applyTextColor(resolveTableCell(rows[tRowIndex], tColIndex, 'setCellTextColor'), tRowIndex, tColIndex);
      } else if (tRowIndex !== undefined) {
        const cells = directChildren(rows[tRowIndex], 'w:tc');
        cells.forEach((tc, cIdx) => applyTextColor(tc, tRowIndex as number, cIdx));
      } else if (tColIndex !== undefined) {
        let maxLogical = 0;
        for (const tr of rows) {
          const probe = findLogicalTableCell(tr, 0);
          if (probe.logicalWidth > maxLogical) maxLogical = probe.logicalWidth;
        }
        if ((tColIndex as number) < 0 || (tColIndex as number) >= maxLogical) {
          throw new Error(
            `E_VALIDATION: setCellTextColor column index ${shown(tColIndex)} out of bounds (table has ${maxLogical} logical columns). ` +
              `Remedy: pass colIndex in [1, ${maxLogical}], counting from 1.`
          );
        }
        rows.forEach((tr, rIdx) => {
          const { tc, isCover } = findLogicalTableCell(tr, tColIndex as number);
          if (tc && !isCover) applyTextColor(tc, rIdx, tColIndex as number);
        });
      } else {
        rows.forEach((tr, rIdx) => {
          directChildren(tr, 'w:tc').forEach((tc, cIdx) => applyTextColor(tc, rIdx, cIdx));
        });
      }
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    case 'setKeepTogether': {
      // Pagination guard for signature-style tables: every row cantSplit (no
      // row ever straddles a page break) plus keepNext on all cell paragraphs
      // except the last row, so Word lays the whole table out as one block
      // (moves to the next page intact instead of widowing rows). Additive
      // only: existing keepNext markers and the last row are left alone.
      for (const tr of rows) {
        if (tblEl.ownerDocument) {
          let trPr = directChildren(tr, 'w:trPr')[0];
          if (!trPr) {
            trPr = createWordElement(tblEl.ownerDocument, 'w:trPr');
            tr.insertBefore(trPr, tr.firstChild);
          }
          if (!trPr.getElementsByTagName('w:cantSplit')[0]) {
            trPr.appendChild(createWordElement(tblEl.ownerDocument, 'w:cantSplit'));
            sortChildrenBySchema(trPr, TRPR_CHILD_ORDER);
          }
        }
      }
      rows.forEach((tr, rIdx) => {
        if (rIdx === rows.length - 1 || !tblEl.ownerDocument) return;
        for (const tc of directChildren(tr, 'w:tc')) {
          let ps = directChildren(tc, 'w:p');
          if (ps.length === 0) {
            const p = createWordElement(tblEl.ownerDocument, 'w:p');
            tc.appendChild(p);
            ps = [p];
          }
          for (const p of ps) {
            let pPr = directChildren(p, 'w:pPr')[0];
            if (!pPr) {
              pPr = createWordElement(tblEl.ownerDocument, 'w:pPr');
              p.insertBefore(pPr, p.firstChild);
            }
            if (!pPr.getElementsByTagName('w:keepNext')[0]) {
              pPr.appendChild(createWordElement(tblEl.ownerDocument, 'w:keepNext'));
              sortChildrenBySchema(pPr, PPR_CHILD_ORDER);
            }
          }
        }
      });
      await refreshTableNode(tableNode, tblEl);
      break;
    }
    default:
      throw new Error(`table.${P.op} is not supported yet`);
  }
  return { warnings };
}
