import { toError } from './errors';

/**
 * Operations that take a free-form `params` object (the schema cannot list its
 * keys without knowing the sub-operation) used to ignore every key they did not
 * read, so `{params:{bold:true}}` on a table cell came back `ok:true` and did
 * nothing. This module is the other half of that contract: after the schema
 * accepts an operation, every key the caller supplied must be one the chosen
 * operation reads, or the call fails before anything is applied.
 */

/** Keys that select or guard the operation rather than configure it. */
const STRUCTURAL = new Set(['action', 'target', 'tableId', 'operation', 'op', 'params', 'trackChanges', 'expectedVersion']);

const ROW = ['rowIndex', 'row'] as const;
const COL = ['colIndex', 'col', 'columnIndex'] as const;

/**
 * What each table sub-operation reads, flat or nested in params (the two are
 * interchangeable). Kept next to the executor's switch in operations/table.ts;
 * tests/unit/operation-params.test.ts fails when a sub-operation is added to
 * the schema without an entry here.
 */
export const TABLE_SUBOP_KEYS: Readonly<Record<string, readonly string[]>> = {
  updateCell: [...ROW, ...COL, 'text', 'author', 'date'],
  setCaption: ['captionText', 'captionPosition'],
  addRow: ['afterIndex', 'cells'],
  deleteRow: ['index', ...ROW],
  addColumn: ['afterIndex', 'width'],
  deleteColumn: ['index', ...COL],
  mergeCells: ['fromRow', 'fromCol', 'toRow', 'toCol'],
  splitCell: [...ROW, ...COL],
  setBorder: ['border', 'color', 'sz', 'size', 'val'],
  setShading: ['fill', 'headerFill', 'zebra', 'textColor'],
  setColumnWidths: ['widths', 'columnWidths'],
  setColumnWidth: ['width', 'w', ...COL],
  setHeaderRepeat: [...ROW],
  setCantSplit: [],
  setVerticalAlign: ['vAlign', 'alignment', ...ROW, ...COL],
  setCellShading: [...ROW, ...COL, 'fill', 'textColor'],
  setRowHeight: ['height', 'rule', ...ROW],
  setCellAlignment: ['align', 'alignment', ...ROW, ...COL],
  setTableAlignment: ['align', 'alignment'],
  setCellPadding: ['padding', ...ROW, ...COL],
  setCellTextColor: ['color', ...ROW, ...COL],
  setKeepTogether: [],
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function unknownKeyError(label: string, unknown: string[], allowed: readonly string[]): Error {
  const names = unknown.map((k) => `"${k}"`).join(', ');
  const list = allowed.length > 0 ? allowed.join(', ') : '(none besides the target)';
  return toError(
    'E_VALIDATION',
    `unknown parameter ${names} for ${label}; allowed: ${list}.`,
    `Remove ${names} or correct its spelling. ${label} accepts: ${list}.`
  );
}

function tableParamsError(op: Record<string, unknown>, nested: Record<string, unknown> | undefined): Error | null {
  const sub = nested?.['op'] ?? op['op'] ?? op['operation'] ?? nested?.['operation'];
  if (typeof sub !== 'string') return null;
  const allowed = TABLE_SUBOP_KEYS[sub];
  // An unknown sub-operation is reported by the schema or the executor with its own message.
  if (!allowed) return null;
  const supplied = new Set<string>();
  for (const source of [op, nested ?? {}]) {
    for (const key of Object.keys(source)) {
      if (!STRUCTURAL.has(key) && source[key] !== undefined) supplied.add(key);
    }
  }
  const unknown = [...supplied].filter((key) => !allowed.includes(key));
  return unknown.length > 0 ? unknownKeyError(`table ${sub}`, unknown, allowed) : null;
}

/**
 * Returns the error for the first key an operation carries that it does not
 * read, or null. `nestedKeys` lists the keys other free-form operations accept
 * inside params (their flat keys are already policed by the strict schema).
 */
export function unknownParamsError(op: unknown, nestedKeys: Readonly<Record<string, readonly string[]>>): Error | null {
  if (!isRecord(op) || typeof op['action'] !== 'string') return null;
  const nested = isRecord(op['params']) ? op['params'] : undefined;
  if (op['action'] === 'table') return tableParamsError(op, nested);
  const allowed = nestedKeys[op['action']];
  if (!allowed || !nested) return null;
  const unknown = Object.keys(nested).filter((key) => !allowed.includes(key) && nested[key] !== undefined);
  return unknown.length > 0 ? unknownKeyError(String(op['action']), unknown, allowed) : null;
}
