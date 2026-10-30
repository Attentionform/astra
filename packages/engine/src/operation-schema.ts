import { z } from 'zod';
import { toError } from './errors';
import { unknownParamsError } from './operation-params';
import { SUPPORTED_CHART_2D_AND_3D_TYPES } from './operations/chart';

/**
 * Track-B strict operation schema defense.
 * Kills `z.record(z.any())`: every operation is a member of a
 * `z.discriminatedUnion('action', [...])` with per-action required fields.
 * Unknown actions, missing required fields, and unknown flat keys are
 * rejected at the schema gate (with remedy) instead of wasting agent steps
 * on engine errors. No `any` anywhere — unknown payloads use `z.unknown()`.
 *
 * Second gate stays in the engine (loud errors with hints for bad ids,
 * out-of-bounds indexes, unknown table sub-ops). `params` nesting (OP-16)
 * is accepted as a string-keyed record of unknowns where the engine
 * supports it; deep nested-value validation is the engine's job.
 */

const nodeId = z.string().min(1).max(256);
const nodeIdOrList = z.union([nodeId, z.array(nodeId).min(1).max(100)]);
const expectedVersion = z.number().int().nonnegative().optional();
const nestedParams = z.record(z.string(), z.unknown()).optional();
const sectionNumber = z.number().int().positive('sections count from 1 (1 is the first section)').describe('Section number, counting from 1 (1 is the first section).');
const position = z.enum(['before', 'after']);
const operationInsertPosition = z.enum(['before', 'after']).optional();

/**
 * Strict pt-size match: one capture group holding the numeric points value.
 * Requires leading digits (`12pt`, `10.5pt`, `10.5 pt`); surrounding
 * whitespace tolerated (` 12pt `); rejects signs
 * (`-5pt`), leading/trailing/multi dots (`.5pt`, `12.pt`, `12.5.3pt`),
 * exponents (`1e3pt`) and wrong units (`12px`). Case-insensitive.
 */
export const FONT_SIZE_PT_REGEX = /^\s*(\d+(?:\.\d+)?)\s*pt\s*$/i;

/**
 * L-8: Convert font size from pt string (e.g. "12pt" -> 24, "10.5pt" -> 21)
 * or pass numeric half-points through.
 *
 * Unit contract (single source of truth for style + replace): bare numbers
 * are OOXML half-points (`w:sz` units, ECMA-376 file truth) on EVERY path;
 * integers pass through unchanged while floats round to the nearest integer
 * (`10.5` -> `11`) so `w:sz`/`w:szCs` always carry valid integer values;
 * points are NEVER inferred from a bare number, so `12` (number) -> `12`
 * while `"12pt"` (string) -> `24` via Math.round(pt*2). Runtime accepts ONLY numbers + pt-suffixed strings:
 * bare numeric strings (`"24"`, `"+5"`, `"1e3"`, `".5"`) yield undefined,
 * matching the schema gate below which rejects every bare string. Callers
 * (operations/style.ts, operations/replace.ts) pass user values straight
 * through parseFontSize, so style and replace agree exactly.
 * Anything else (negatives, non-finite, pt-suffixed garbage that fails the
 * strict regex) yields undefined at runtime and is rejected at the schema
 * gate. `0` / `"0pt"` -> `0` everywhere (schema-uniform; see watermark note).
 */
export function parseFontSize(sz: number | string | undefined): number | undefined {
  if (sz === undefined || sz === null) return undefined;
  if (typeof sz === 'number') {
    // Round: w:sz takes integer half-points only, so a bare float (10.5)
    // normalizes to the nearest integer (11) instead of emitting invalid XML.
    return Number.isFinite(sz) && sz >= 0 ? Math.round(sz) : undefined;
  }
  if (typeof sz === 'string') {
    const trimmed = sz.trim();
    const match = trimmed.match(FONT_SIZE_PT_REGEX);
    if (match) {
      const pt = parseFloat(match[1]);
      return Number.isFinite(pt) && pt >= 0 ? Math.round(pt * 2) : undefined;
    }
    // No Number() fallback: bare numeric strings ("24", "+5", "1e3", ".5")
    // and pt-suffixed garbage ("+5pt", ".5pt") that failed the strict regex
    // must not coerce — the schema gate rejects them, so runtime matches it.
    return undefined;
  }
  return undefined;
}

/**
 * Paragraph spacing and indents are stored in twips (20 per point, 1440 per inch). A bare number
 * is twips; a string such as "12pt" or "0.5pt" is points. Anything else yields undefined and is
 * rejected at the schema gate, so `spacingAfter: 12` can no longer silently mean 0.6 pt.
 */
export function parseTwips(value: number | string | undefined): number | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'number') return Number.isFinite(value) ? Math.round(value) : undefined;
  const match = typeof value === 'string' ? value.trim().match(FONT_SIZE_PT_REGEX) : null;
  if (!match) return undefined;
  const pt = parseFloat(match[1]);
  return Number.isFinite(pt) && pt >= 0 ? Math.round(pt * 20) : undefined;
}

/**
 * `parseTwips` for an executor. Flat settings are checked at the schema gate, but a value inside `params`
 * is not, and an unreadable one must stop the call instead of being written into the XML as "undefined".
 */
export function requireTwips(name: string, value: unknown): number {
  const twips = parseTwips(value as number | string | undefined);
  if (twips === undefined) {
    throw new Error(
      `E_VALIDATION: ${name} must be a number of twips (20 per point) or a string in points such as "12pt" (got ${JSON.stringify(value)}). ` +
        `Remedy: pass ${name}: 240 or ${name}: "12pt".`
    );
  }
  return twips;
}

const twipsOrPoints = z
  .union([
    z.number().finite(),
    z.string().regex(FONT_SIZE_PT_REGEX, 'must be a number of twips (20 per point) or a string such as "12pt"'),
  ])
  .describe('Twips (20 per point, so 240 is 12 pt), or a string in points such as "12pt".');

const fontSizeOrSzSchema = z.union([
  z.number().finite().nonnegative(),
  z.string().regex(FONT_SIZE_PT_REGEX, 'Font size string must match format like "12pt" or "10.5pt"')
]);

const styleProperties = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  doubleStrikethrough: z.boolean().optional(),
  vertAlign: z.enum(['superscript', 'subscript', 'baseline']).optional(),
  alignment: z.enum(['left', 'center', 'right', 'both', 'justify']).optional(),
  color: z.string().max(32).optional(),
  fontSize: fontSizeOrSzSchema.optional(),
  sz: fontSizeOrSzSchema.optional(),
  fontFamily: z.string().max(128).optional(),
  highlight: z.string().max(32).optional(),
  lineSpacing: z.number().optional(),
  spacingBefore: twipsOrPoints.optional(),
  spacingAfter: twipsOrPoints.optional(),
  indentLeft: twipsOrPoints.optional(),
  indentRight: twipsOrPoints.optional(),
  indentFirstLine: twipsOrPoints.optional(),
  indentHanging: twipsOrPoints.optional(),
  // Word-level span gate: styles only the matched substring (run splitting
  // in operations/style.ts). Flat op.match wins over properties/params
  // nesting; a nested match must never silently degrade to whole-paragraph.
  match: z.string().min(1).max(1000).optional().describe('Word-level span: style only this literal substring (run splitting); an absent match fails with E_NOT_FOUND.'),
  matchAll: z.boolean().optional().describe('Style every occurrence of match (default false: first occurrence only).'),
  matchCase: z.boolean().optional().describe('Case-sensitive match (default true: pass false for case-insensitive).'),
}).catchall(z.unknown());

/** Shared image/chart compat shapes for insert rewriting (M3 backward-compat). */
const insertChartSeries = z.object({
  name: z.string().max(500),
  values: z.array(z.number()).max(1_000),
  color: z.string().max(16).optional(),
}).catchall(z.unknown());

const insertWrapMode = z.enum(['inline', 'square', 'tight', 'through', 'topAndBottom', 'behindText', 'inFrontOfText']);

const tableCellPadding = z.object({
  top: z.number().int().nonnegative().optional(),
  left: z.number().int().nonnegative().optional(),
  bottom: z.number().int().nonnegative().optional(),
  right: z.number().int().nonnegative().optional(),
}).strict().refine(
  (obj) => obj.top !== undefined || obj.left !== undefined || obj.bottom !== undefined || obj.right !== undefined,
  { message: 'cellPadding object must contain at least one valid side (top, bottom, left, right)' }
);

const insertParams = z.object({
  cellPadding: z.union([z.number().int().nonnegative(), tableCellPadding]).optional(),
}).catchall(z.unknown());

/**
 * Who and when a revision or comment is attributed to. Both are optional: the author falls back
 * to the document default (settings.defaultAuthor, set per run or per session) and the date to
 * the current time, or to the pinned stamp when output is reproducible.
 */
const revisionAuthor = z.string().trim().min(1).max(128).optional();
const revisionDate = z
  .string()
  .regex(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$/, 'date must be an ISO 8601 timestamp such as 2026-10-24T09:30:00Z')
  .refine((value) => !Number.isNaN(Date.parse(value)), 'date must be a real calendar date and time')
  .optional();

const insertOp = z.object({
  action: z.literal('insert'),
  // M3: 'image'/'chart' are accepted here for backward-compat and mapped by
  // operation-engine routeOperation to the dedicated image/chart ops
  // (explicit rewrite, never silent). Prefer the dedicated
  // {action:'image',operation:'insert',...} / {action:'chart',...} shapes for
  // new code; insert+type:image/chart keeps old batches working.
  type: z.enum(['paragraph', 'pageBreak', 'sectionBreak', 'table', 'xml', 'divider', 'horizontalRule', 'image', 'chart']),
  // Anchor/position are optional here so header/footer inserts (which the
  // engine appends in call order without ever reading them) need no dummy
  // anchor. Body inserts still require both — enforced by the strict
  // validators below, not by this shape.
  anchor: z.string().min(1).max(256).optional(),
  position: operationInsertPosition,
  text: z.string().max(500_000).optional(),
  xml: z.string().max(500_000).optional(),
  id: z.string().min(1).max(256).optional(),
  headers: z.array(z.string().max(10_000)).max(100).optional(),
  rows: z.array(z.array(z.string().max(10_000)).max(100)).max(1_000).optional(),
  cellPadding: z.union([z.number().int().nonnegative(), tableCellPadding]).optional(),
  color: z.string().max(16).optional(),
  thickness: z.number().optional(),
  lineStyle: z.string().max(32).optional(),
  spacingBefore: twipsOrPoints.optional(),
  spacingAfter: twipsOrPoints.optional(),
  breakType: z.enum(['nextPage', 'nextColumn', 'continuous', 'evenPage', 'oddPage']).optional(),
  // M3 backward-compat args for type:'image'/'chart' (forwarded by the engine
  // rewrite to executeImage/executeChart). Absent for other types.
  fileData: z.string().max(20_000_000).optional(),
  width: z.number().int().positive().max(100_000_000).optional(),
  height: z.number().int().positive().max(100_000_000).optional(),
  widthEmu: z.number().int().positive().max(100_000_000).optional(),
  heightEmu: z.number().int().positive().max(100_000_000).optional(),
  // Chart types: 2D default ('column' when omitted) + opt-in 3D, enumerated
  // from the engine's single source of truth.
  chartType: z.enum(SUPPORTED_CHART_2D_AND_3D_TYPES).optional(),
  holeSize: z.number().int().min(1).max(90).optional(),
  title: z.string().max(1_000).optional(),
  categories: z.array(z.string().max(500)).max(1_000).optional(),
  series: z.array(insertChartSeries).max(100).optional(),
  view3D: z.record(z.string(), z.unknown()).optional(),
  shape: z.string().max(32).optional(),
  wrapMode: insertWrapMode.optional(),
  captionText: z.string().max(5_000).optional(),
  // R-42: table caption placement ('above' default | 'below'). Accepted at
  // the gate so table inserts can request a below-table caption; honored by
  // the table caption path (operations/table.ts setCaption, adjacent to
  // updateCell). Bad values fail closed here with E_VALIDATION.
  captionPosition: z.enum(['above', 'below']).optional(),
  // Table inserts only: keep rows together (cantSplit + tblHeader + keepNext
  // chaining). Explicit wins; otherwise the guide applies; else legacy.
  keepTogether: z.boolean().optional(),
  // Table inserts only: bold the header row text. Explicit wins; otherwise the guide's
  // tables.headerBold applies; otherwise bold.
  headerBold: z.boolean().optional(),
  altText: z.string().max(5_000).optional(),
  // Paragraph inserts may opt into revision markup (runs wrapped in w:ins);
  // the engine honors it for type:'paragraph' and fails closed otherwise.
  trackChanges: z.boolean().optional(),
  author: revisionAuthor,
  date: revisionDate,
  // Destination store for header/footer routing, honored by the insert path
  // (flat or params-nested values, nested wins on conflict; bad values fail
  // closed here with E_VALIDATION).
  target: z.enum(['body', 'header', 'footer']).optional(),
  pageType: z.enum(['default', 'first', 'even']).optional(),
  sectionIndex: sectionNumber.optional(),
  expectedVersion,
  // Optional free-form params bag; the insert path honors routing keys
  // nested here exactly like the flat values above.
  params: insertParams.optional(),
}).strict();

const replaceOp = z.object({
  action: z.literal('replace'),
  target: z.string().min(1).max(256),
  text: z.string().max(500_000),
  preserveFormatting: z.boolean().optional(),
  trackChanges: z.boolean().optional(),
  author: revisionAuthor,
  date: revisionDate,
  fontSize: fontSizeOrSzSchema.optional(),
  sz: fontSizeOrSzSchema.optional(),
  expectedVersion,
  params: nestedParams,
}).strict();

const deleteOp = z.object({
  action: z.literal('delete'),
  target: nodeIdOrList,
  strict: z.boolean().optional(),
  trackChanges: z.boolean().optional(),
  author: revisionAuthor,
  date: revisionDate,
  expectedVersion,
  params: nestedParams,
}).strict();

const styleOp = z.object({
  action: z.literal('style'),
  target: nodeIdOrList,
  properties: styleProperties.optional(),
  params: styleProperties.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  doubleStrikethrough: z.boolean().optional(),
  vertAlign: z.enum(['superscript', 'subscript', 'baseline']).optional(),
  alignment: z.enum(['left', 'center', 'right', 'both', 'justify']).optional(),
  color: z.string().max(32).optional(),
  fontSize: fontSizeOrSzSchema.optional(),
  sz: fontSizeOrSzSchema.optional(),
  fontFamily: z.string().max(128).optional(),
  highlight: z.string().max(32).optional(),
  lineSpacing: z.number().optional(),
  spacingBefore: twipsOrPoints.optional(),
  spacingAfter: twipsOrPoints.optional(),
  indentLeft: twipsOrPoints.optional(),
  indentRight: twipsOrPoints.optional(),
  indentFirstLine: twipsOrPoints.optional(),
  indentHanging: twipsOrPoints.optional(),
  // Word-level span gate (flat; wins over properties/params nesting).
  match: z.string().min(1).max(1000).optional().describe('Word-level span: style only this literal substring (run splitting); an absent match fails with E_NOT_FOUND.'),
  matchAll: z.boolean().optional().describe('Style every occurrence of match (default false: first occurrence only).'),
  matchCase: z.boolean().optional().describe('Case-sensitive match (default true: pass false for case-insensitive).'),
  expectedVersion,
}).strict();

const clearFormatOp = z.object({
  action: z.literal('clearFormat'),
  target: nodeIdOrList,
  expectedVersion,
  params: nestedParams,
}).strict();

const moveOp = z.object({
  action: z.literal('move'),
  target: nodeIdOrList,
  anchor: z.string().min(1).max(256),
  position,
  expectedVersion,
  params: nestedParams,
}).strict();

/**
 * M1: canonical engine-supported table sub-ops (single source mirrored from
 * backend/operations/table.ts switch + operation-engine route).
 * Flat `operation`/`op` are enums (typos rejected at the schema gate with
 * remedy); `params` stays a free record for args (engine validates deep
 * values). Engine case list: updateCell/setCaption/addRow/deleteRow/addColumn/
 * deleteColumn/mergeCells/splitCell/setBorder/setShading/setColumnWidths/
 * setColumnWidth/setHeaderRepeat/setCantSplit/setVerticalAlign/
 * setCellShading/setRowHeight/setCellAlignment/setCellPadding/
 * setCellTextColor/setKeepTogether.
 */
const tableSubOp = z.enum([
  'updateCell',
  'setCaption',
  'addRow',
  'deleteRow',
  'addColumn',
  'deleteColumn',
  'mergeCells',
  'splitCell',
  'setBorder',
  'setShading',
  'setColumnWidths',
  'setColumnWidth',
  'setHeaderRepeat',
  'setCantSplit',
  'setVerticalAlign',
  'setCellShading',
  'setRowHeight',
  'setCellAlignment',
  'setTableAlignment',
  'setCellPadding',
  'setCellTextColor',
  'setKeepTogether',
]);

const tableAlignmentSchema = z.preprocess(
  (val) => (typeof val === 'string' ? val.toLowerCase() : val),
  z.enum(['left', 'center', 'right'])
);

const tableOp = z.object({
  action: z.literal('table'),
  target: z.string().min(1).max(256).optional(),
  tableId: z.string().min(1).max(256).optional(),
  operation: tableSubOp.optional(),
  op: tableSubOp.optional(),
  trackChanges: z.boolean().optional(),
  author: revisionAuthor,
  date: revisionDate,
  params: nestedParams,
  // Rows and columns are numbered from 1, as in the inspect grid (r1c1) and the cell addresses.
  rowIndex: z.number().int().describe('Row number, counting from 1 (row 1 is the first row, usually the header).').optional(),
  colIndex: z.number().int().describe('Column number, counting from 1.').optional(),
  // Short aliases for rowIndex/colIndex (engine resolves row->rowIndex, col->colIndex).
  row: z.number().int().describe('Alias of rowIndex (counting from 1).').optional(),
  col: z.number().int().describe('Alias of colIndex (counting from 1).').optional(),
  columnIndex: z.number().int().describe('Alias of colIndex (counting from 1).').optional(),
  text: z.string().max(100_000).optional(),
  afterIndex: z.number().int().describe('addRow: insert after this row number; addColumn: after this column number (counting from 1; default the last).').optional(),
  index: z.number().int().describe('deleteRow/deleteColumn: the row or column number to remove (counting from 1).').optional(),
  fill: z.string().max(16).optional(),
  headerFill: z.string().max(16).optional(),
  zebra: z.boolean().optional(),
  // Explicit header/single-cell text color for setShading/setCellShading
  // (flat or params-nested; wins over the guide per the precedence invariant).
  textColor: z.string().max(16).optional(),
  widths: z.array(z.number()).max(100).optional(),
  columnWidths: z.array(z.number()).max(100).optional(),
  width: z.number().int().positive().optional(),
  height: z.number().int().positive().optional(),
  rule: z.enum(['atLeast', 'exact', 'auto']).optional(),
  border: z.string().max(32).optional(),
  color: z.string().max(16).optional(),
  sz: fontSizeOrSzSchema.optional(),
  fontSize: fontSizeOrSzSchema.optional(),
  vAlign: z.string().max(16).optional(),
  align: tableAlignmentSchema.optional(),
  alignment: tableAlignmentSchema.optional(),
  padding: z.union([z.number().int().nonnegative(), tableCellPadding]).optional(),
  fromRow: z.number().int().describe('mergeCells: first row of the range (counting from 1).').optional(),
  fromCol: z.number().int().describe('mergeCells: first column of the range (counting from 1).').optional(),
  toRow: z.number().int().describe('mergeCells: last row of the range (counting from 1; default fromRow).').optional(),
  toCol: z.number().int().describe('mergeCells: last column of the range (counting from 1; default fromCol).').optional(),
  // R-42: table caption args (setCaption, adjacent to updateCell).
  // captionPosition defaults to 'above' at the engine honor site.
  captionText: z.string().max(5_000).optional(),
  captionPosition: z.enum(['above', 'below']).optional(),
  expectedVersion,
}).strict();

const sectionOp = z.object({
  action: z.literal('section'),
  operation: z.enum(['setColumns', 'split', 'merge']),
  target: z.string().min(1).max(256).optional(),
  sectionIndex: sectionNumber.optional(),
  columns: z.number().int().min(1).max(12).optional(),
  spacing: z.number().optional(),
  separator: z.boolean().optional(),
  expectedVersion,
  params: nestedParams,
}).strict();

const pageSetupOp = z.object({
  action: z.literal('pageSetup'),
  params: z.object({
    marginTop: z.number().optional(),
    marginBottom: z.number().optional(),
    marginLeft: z.number().optional(),
    marginRight: z.number().optional(),
    headerDistance: z.number().optional(),
    footerDistance: z.number().optional(),
    orientation: z.enum(['portrait', 'landscape']).optional(),
    paperSize: z.enum(['A4', 'Letter', 'Legal', 'A3']).optional(),
    columns: z.number().int().min(1).max(12).optional(),
    sectionIndex: sectionNumber.optional(),
  }).catchall(z.unknown()).optional(),
  marginTop: z.number().optional(),
  marginBottom: z.number().optional(),
  marginLeft: z.number().optional(),
  marginRight: z.number().optional(),
  headerDistance: z.number().optional(),
  footerDistance: z.number().optional(),
  orientation: z.enum(['portrait', 'landscape']).optional(),
  paperSize: z.enum(['A4', 'Letter', 'Legal', 'A3']).optional(),
  columns: z.number().int().min(1).max(12).optional(),
  sectionIndex: sectionNumber.optional(),
  expectedVersion,
}).strict();

const namedStyleOp = z.object({
  action: z.literal('namedStyle'),
  operation: z.enum(['apply']),
  styleName: z.string().min(1).max(128),
  target: nodeIdOrList.optional(),
  expectedVersion,
}).strict();

const fieldOp = z.object({
  action: z.literal('field'),
  operation: z.enum(['insertTOC', 'insertPageNumber', 'insertWordCount', 'updateFields']),
  anchor: z.string().min(1).max(256).optional(),
  position: operationInsertPosition,
  title: z.string().max(1_000).optional(),
  levels: z.number().int().min(1).max(9).optional(),
  // Destination store for header/footer routing, honored by the field path
  // (flat or params-nested; bad values fail closed here with E_VALIDATION).
  target: z.enum(['body', 'header', 'footer']).optional(),
  pageType: z.enum(['default', 'first', 'even']).optional(),
  sectionIndex: sectionNumber.optional(),
  expectedVersion,
  params: nestedParams,
}).strict();

const equationOp = z.object({
  action: z.literal('equation'),
  operation: z.enum(['insert', 'replace']),
  latex: z.string().min(1).max(50_000),
  anchor: z.string().min(1).max(256).optional(),
  position: operationInsertPosition,
  target: z.string().min(1).max(256).optional(),
  id: z.string().min(1).max(256).optional(),
  mode: z.enum(['auto', 'latex', 'plain']).optional(),
  expectedVersion,
  params: nestedParams,
}).strict();

const listOp = z.object({
  action: z.literal('list'),
  operation: z.enum(['create', 'continue', 'changeLevel', 'remove']),
  target: nodeIdOrList,
  listType: z.enum(['bullet', 'numbered']).optional(),
  previousNodeId: z.string().min(1).max(256).optional(),
  level: z.number().int().min(0).max(9).describe('List level, counting from 1 (1 is the top level, up to 9). changeLevel also takes 0 to take the paragraph out of the list.').optional(),
  delta: z.number().int().min(-9).max(9).describe('changeLevel: move this many levels deeper (positive) or shallower (negative); below level 1 leaves the list.').optional(),
  expectedVersion,
  params: nestedParams,
}).strict();

const headerFooterOp = z.object({
  action: z.literal('headerFooter'),
  operation: z.enum(['set', 'remove']),
  targetType: z.enum(['header', 'footer']),
  pageType: z.enum(['default', 'first', 'even']).optional(),
  text: z.string().max(50_000).optional(),
  xmlContent: z.string().max(500_000).optional(),
  sectionIndex: sectionNumber.optional(),
  alignment: z.enum(['left', 'center', 'right']).optional(),
  numberFormat: z.string().max(32).optional(),
  expectedVersion,
}).strict();

/**
 * Chart-specific remedy for a schema-level chartType enum failure. Batch
 * wrappers otherwise replace it with their generic "canonical shapes" line.
 */
export function chartTypeSchemaRemedy(detail: string): string | undefined {
  if (!/chartType/i.test(detail)) return undefined;
  return `use one of ${SUPPORTED_CHART_2D_AND_3D_TYPES.join(' | ')}; omit chartType to keep an existing stacked/percent chart's grouping or a doughnut's hole, or delete and re-insert the chart with the desired type`;
}

export function chartHoleSizeSchemaRemedy(detail: string): string | undefined {
  if (!/holeSize/i.test(detail)) return undefined;
  return 'holeSize must be an integer from 1 to 90 (doughnut charts only); omit it to keep the authored hole';
}

const chartSeries = z.object({
  name: z.string().max(500),
  values: z.array(z.number()).max(1_000),
  color: z.string().max(16).optional(),
  chartType: z.enum(['column', 'line']).optional(),
  axis: z.enum(['primary', 'secondary']).optional(),
}).catchall(z.unknown());

const chartLegendStyle = z.object({
  position: z.enum(['b', 't', 'l', 'r', 'tr', 'none']).optional(),
  overlay: z.boolean().optional(),
}).strict();

const chartDataLabelStyle = z.object({
  show: z.boolean().optional(),
  position: z.enum(['ctr', 'inBase', 'inEnd', 'outEnd', 'bestFit', 'b', 't', 'l', 'r']).optional(),
  numberFormat: z.string().min(1).max(255).optional(),
  showValue: z.boolean().optional(),
  showCategoryName: z.boolean().optional(),
  showSeriesName: z.boolean().optional(),
  showPercent: z.boolean().optional(),
  showLegendKey: z.boolean().optional(),
}).strict();

const chartCategoryAxisStyle = z.object({
  title: z.string().min(1).max(1_000).optional(),
  hidden: z.boolean().optional(),
  tickLabelPosition: z.enum(['nextTo', 'high', 'low', 'none']).optional(),
  majorGridlines: z.boolean().optional(),
}).strict();

const chartValueAxisStyle = z.object({
  title: z.string().min(1).max(1_000).optional(),
  hidden: z.boolean().optional(),
  min: z.number().finite().optional(),
  max: z.number().finite().optional(),
  numberFormat: z.string().min(1).max(255).optional(),
  majorGridlines: z.boolean().optional(),
  tickLabelPosition: z.enum(['nextTo', 'high', 'low', 'none']).optional(),
}).strict();

const chartTrendlineStyle = z.object({
  type: z.enum(['linear', 'exponential', 'logarithmic', 'movingAvg', 'polynomial', 'power', 'none']),
  series: z.array(z.number().int().positive('series numbers count from 1 (1 is the first series)').describe('Series number, counting from 1 in the order the series appear')).min(1).max(100).optional(),
  order: z.number().int().min(2).max(6).optional(),
  period: z.number().int().min(2).max(255).optional(),
  displayEquation: z.boolean().optional(),
  displayRSquared: z.boolean().optional(),
}).strict();

const chartFloatingStyle = z.object({
  horizontalAlign: z.enum(['left', 'center', 'right']).optional(),
  verticalAlign: z.enum(['top', 'middle', 'bottom']).optional(),
  behindText: z.boolean().optional(),
  wrap: z.enum(['none', 'square']).optional(),
}).strict();

const chartStyle = z.object({
  legend: chartLegendStyle.optional(),
  dataLabels: chartDataLabelStyle.optional(),
  axes: z.object({
    category: chartCategoryAxisStyle.optional(),
    value: chartValueAxisStyle.optional(),
  }).strict().optional(),
  trendline: chartTrendlineStyle.optional(),
  floating: chartFloatingStyle.optional(),
}).strict();

const chartOp = z.object({
  action: z.literal('chart'),
  operation: z.enum(['insert', 'update', 'updateView3D']).optional(),
  target: z.string().min(1).max(256).optional(),
  // Chart types: 2D default ('column' when omitted) + opt-in 3D, enumerated
  // from the engine's single source of truth (dispatch validates the same
  // list and still carries the remedy for untyped callers).
  chartType: z.enum(SUPPORTED_CHART_2D_AND_3D_TYPES).optional(),
  holeSize: z.number().int().min(1).max(90).optional(),
  title: z.string().max(1_000).optional(),
  view3D: z.record(z.string(), z.unknown()).optional(),
  shape: z.string().max(32).optional(),
  categories: z.array(z.string().max(500)).max(1_000).optional(),
  series: z.array(chartSeries).max(100).optional(),
  style: chartStyle.optional(),
  anchor: z.string().min(1).max(256).optional(),
  id: z.string().min(1).max(256).optional(),
  widthEmu: z.number().positive().max(100_000_000).optional(),
  heightEmu: z.number().positive().max(100_000_000).optional(),
  expectedVersion,
  params: nestedParams,
}).strict();

const imageOp = z.object({
  action: z.literal('image'),
  operation: z.enum(['insert', 'resize', 'reposition', 'setWrap', 'setCaption', 'delete', 'setAltText']),
  anchor: z.string().min(1).max(256).optional(),
  target: z.string().min(1).max(256).optional(),
  fileData: z.string().max(20_000_000).optional(),
  filePath: z.string().min(1).max(1024).optional(),
  width: z.number().int().positive().max(100_000_000).optional(),
  height: z.number().int().positive().max(100_000_000).optional(),
  wrapMode: z.enum(['inline', 'square', 'tight', 'through', 'topAndBottom', 'behindText', 'inFrontOfText']).optional(),
  captionText: z.string().max(5_000).optional(),
  altText: z.string().max(5_000).optional(),
  expectedVersion,
}).strict();

const calloutOp = z.object({
  action: z.literal('callout'),
  text: z.string().min(1).max(100_000),
  anchor: z.string().min(1).max(256).optional(),
  title: z.string().max(1_000).optional(),
  theme: z.string().max(32).optional(),
  position: operationInsertPosition,
  expectedVersion,
  params: nestedParams,
}).strict();

const footnoteOp = z.object({
  action: z.literal('footnote'),
  operation: z.enum(['insert']).optional(),
  anchor: z.string().min(1).max(256),
  text: z.string().min(1).max(50_000),
  id: z.string().min(1).max(256).optional(),
  expectedVersion,
}).strict();

const endnoteOp = z.object({
  action: z.literal('endnote'),
  operation: z.enum(['insert']).optional(),
  anchor: z.string().min(1).max(256),
  text: z.string().min(1).max(50_000),
  id: z.string().min(1).max(256).optional(),
  expectedVersion,
}).strict();

/**
 * Watermark size gate: numbers are half-points, pt-suffixed strings convert
 * via the same strict regex as body text. `0` / `"0pt"` -> `0` is accepted
 * here for uniformity with fontSizeOrSzSchema (parse returns 0 everywhere);
 * a zero-size watermark is degenerate but harmless and never corrupts — the
 * engine default path covers a missing size. `positive()` was rejected
 * because it split number-0 (fail) from string-"0pt" (pass) for the same
 * value; `nonnegative()` keeps both in agreement.
 */
const watermarkSizeSchema = z.union([z.number().finite().nonnegative().max(500), z.string().regex(FONT_SIZE_PT_REGEX)]);
const watermarkOp = z.object({
  action: z.literal('watermark'),
  text: z.string().max(500).optional(),
  color: z.string().max(16).optional(),
  fontSize: watermarkSizeSchema.optional(),
  sz: watermarkSizeSchema.optional(),
  opacity: z.number().min(0).max(1).optional(),
  rotation: z.number().optional(),
  expectedVersion,
}).strict();

const pageBorderOp = z.object({
  action: z.literal('pageBorder'),
  style: z.string().max(32).optional(),
  sz: fontSizeOrSzSchema.optional(),
  fontSize: fontSizeOrSzSchema.optional(),
  color: z.string().max(16).optional(),
  marginOffset: z.enum(['page', 'text']).optional(),
  sectionIndex: sectionNumber.optional(),
  expectedVersion,
}).strict();

const coverPageParams = z.object({
  topSpacingBefore: z.number().int().nonnegative().optional(),
  dividerSpacingAfter: z.number().int().nonnegative().optional(),
  metaSpacingAfter: z.number().int().nonnegative().optional(),
}).catchall(z.unknown());

const coverPageOp = z.object({
  action: z.literal('coverPage'),
  title: z.string().min(1).max(1_000),
  template: z.enum(['modern_minimalist', 'corporate_navy', 'academic', 'executive']).optional(),
  subtitle: z.string().max(1_000).optional(),
  author: z.string().max(500).optional(),
  organization: z.string().max(500).optional(),
  org: z.string().max(500).optional(),
  date: z.string().max(100).optional(),
  topSpacingBefore: z.number().int().nonnegative().optional(),
  dividerSpacingAfter: z.number().int().nonnegative().optional(),
  metaSpacingAfter: z.number().int().nonnegative().optional(),
  expectedVersion,
  params: coverPageParams.optional(),
}).strict();

const hyperlinkOp = z.object({
  action: z.literal('hyperlink'),
  operation: z.enum(['insert']).optional(),
  target: z.string().min(1).max(256),
  text: z.string().max(5_000).optional(),
  url: z.string().max(2_000).optional(),
  bookmarkName: z.string().max(256).optional(),
  expectedVersion,
}).strict();

const bookmarkOp = z.object({
  action: z.literal('bookmark'),
  operation: z.enum(['insert']).optional(),
  target: z.string().min(1).max(256),
  text: z.string().max(5_000).optional(),
  url: z.string().max(2_000).optional(),
  bookmarkName: z.string().max(256).optional(),
  name: z.string().max(256).optional(),
  expectedVersion,
}).strict();

const commentOp = z.object({
  action: z.literal('comment'),
  operation: z.enum(['insert', 'reply']),
  target: z.string().min(1).max(256),
  text: z.string().min(1).max(50_000),
  author: revisionAuthor,
  date: revisionDate,
  expectedVersion,
}).strict();

const trackedChangeOp = z.object({
  action: z.literal('trackedChange'),
  operation: z.enum(['acceptAll', 'rejectAll', 'insert', 'delete']),
  target: z.string().min(1).optional(),
  text: z.string().optional(),
  author: revisionAuthor,
  date: revisionDate,
  expectedVersion,
}).strict();

const metadataOp = z.object({
  // E2 metadata op: 6 WRITABLE fields (below). READABLE via ast.metadata/
  // inspect is 10 (the 6 plus lastModifiedBy|revision|created|modified —
  // read-only here; revision-via-op is rejected by .strict() unknown-key).
  action: z.literal('metadata'),
  operation: z.enum(['setProperties']),
  title: z.string().max(5_000).optional(),
  creator: z.string().max(5_000).optional(),
  subject: z.string().max(5_000).optional(),
  description: z.string().max(5_000).optional(),
  keywords: z.string().max(5_000).optional(),
  category: z.string().max(5_000).optional(),
}).strict();

export {
  insertOp,
  replaceOp,
  deleteOp,
  styleOp,
  replaceOp as replaceOpSchema,
  styleOp as styleOpSchema,
};

/** Strict per-action operation schema — the `z.record(z.any())` killer. */
export const OperationSchema = z.discriminatedUnion('action', [
  insertOp,
  replaceOp,
  deleteOp,
  styleOp,
  clearFormatOp,
  moveOp,
  tableOp,
  sectionOp,
  pageSetupOp,
  namedStyleOp,
  fieldOp,
  equationOp,
  listOp,
  headerFooterOp,
  chartOp,
  imageOp,
  calloutOp,
  footnoteOp,
  endnoteOp,
  watermarkOp,
  pageBorderOp,
  coverPageOp,
  hyperlinkOp,
  bookmarkOp,
  commentOp,
  trackedChangeOp,
  metadataOp,
]);

export type Operation = z.infer<typeof OperationSchema>;

/**
 * Keys read from `params` that the operation does not declare as flat keys
 * (flat and nested are otherwise interchangeable). Everything else an operation
 * may carry inside params is derived from its own strict schema below, so a new
 * operation that reuses `nestedParams` is covered without being listed here.
 */
const NESTED_ONLY_KEYS: Readonly<Record<string, readonly string[]>> = {
  replace: ['resetStyle'],
  delete: ['trackChanges'],
};

function deriveNestedParamKeys(): Readonly<Record<string, readonly string[]>> {
  const out: Record<string, readonly string[]> = {};
  for (const option of OperationSchema.options) {
    const shape = option.shape as Record<string, unknown>;
    if (shape['params'] !== nestedParams) continue;
    const action = (shape['action'] as z.ZodLiteral<string>).value;
    const keys = Object.keys(shape).filter((key) => key !== 'action' && key !== 'params');
    out[action] = [...keys, ...(NESTED_ONLY_KEYS[action] ?? [])];
  }
  return out;
}

/** Actions that take a free-form `params` object, with the keys each may carry in it. */
export const NESTED_PARAM_KEYS = deriveNestedParamKeys();

/** Batch schema: capped operation list, sequential default true (applied by the engine). */
export const BatchOperationsSchema = z.object({
  // Keys an operation carries but does not read fail here, so every caller of this schema
  // (CLI and MCP alike) refuses them before anything is applied.
  operations: z.array(OperationSchema).min(1).max(500).superRefine((operations, ctx) => {
    operations.forEach((operation, index) => {
      const error = unknownParamsError(operation, NESTED_PARAM_KEYS);
      if (error) ctx.addIssue({ code: 'custom', path: [index], message: error.message });
    });
  }),
  sequential: z.boolean().optional(),
  rollbackOnError: z.boolean().optional(),
  maxSteps: z.number().int().min(1).optional(),
  /**
   * R-8: optional per-batch zip-bomb ceiling overrides (same knobs as env
   * `ASTRA_MAX_RATIO` / `ASTRA_MAX_UNCOMPRESSED_BYTES` and CLI
   * `--max-ratio` / `--max-uncompressed-bytes`; explicit CLI flags win over
   * these file values). Absent = env/default. Bounded positive so a typo
   * cannot silently disable the fail-closed ceilings.
   */
  maxRatio: z.number().finite().positive().max(1_000_000).optional(),
  maxUncompressedBytes: z.number().int().positive().max(Number.MAX_SAFE_INTEGER).optional(),
}).strict();

export type BatchInput = z.infer<typeof BatchOperationsSchema>;

function formatZodIssues(issues: ReadonlyArray<{ path: ReadonlyArray<PropertyKey>; message: string; code: string }>): string {
  return issues.map((i) => `${i.path.map(String).join('.') || '(root)'}: ${i.message} [${i.code}]`).join('; ');
}

/**
 * Body-path anchor rule for insert ops. The engine appends header/footer
 * inserts in call order and never reads anchor/position there, so both stay
 * optional when the effective target is header/footer. Body inserts still
 * need an explicit anchor + position. The effective target mirrors the
 * insert path (params-nested wins on conflict); anything but an exact
 * header/footer string counts as body. Returns the issue detail or null.
 */
function insertBodyAnchorDetail(op: unknown): string | null {
  if (!op || typeof op !== 'object') return null;
  const rec = op as Record<string, unknown>;
  if (rec['action'] !== 'insert') return null;
  const params = rec['params'];
  const nestedTarget =
    params !== null && typeof params === 'object' && !Array.isArray(params)
      ? (params as Record<string, unknown>)['target']
      : undefined;
  const effectiveTarget = nestedTarget !== undefined ? nestedTarget : rec['target'];
  if (effectiveTarget === 'header' || effectiveTarget === 'footer') return null;
  const missing: string[] = [];
  if (typeof rec['anchor'] !== 'string' || (rec['anchor'] as string).length < 1) {
    missing.push('anchor: anchor is required for body inserts [invalid_type]');
  }
  const pos = rec['position'];
  if (pos !== 'before' && pos !== 'after') {
    missing.push('position: position is required for body inserts [invalid_type]');
  }
  return missing.length > 0 ? missing.join('; ') : null;
}

/**
 * Image insert source rule: `fileData` (inline base64) and `filePath` (file
 * inside allowedDir) are mutually exclusive, and an insert needs exactly
 * one. Both shapes stay optional on the strict schema (no refinement
 * wrappers, so `.shape` reflection keeps working); the conflict/absence is
 * enforced here and at `executeImage` entry. Returns the coded error or
 * null. Presence is truthiness (empty string counts as absent, matching the
 * executor).
 */
function imageInsertSourceError(op: unknown): Error | null {
  if (!op || typeof op !== 'object') return null;
  const rec = op as Record<string, unknown>;
  if (rec['action'] !== 'image' || rec['operation'] !== 'insert') return null;
  const hasData = typeof rec['fileData'] === 'string' && rec['fileData'].length > 0;
  const hasPath = typeof rec['filePath'] === 'string' && rec['filePath'].length > 0;
  if (hasData && hasPath) {
    return toError(
      'E_VALIDATION',
      'image insert carries both fileData and filePath (mutually exclusive)',
      'pass exactly one image source: fileData with inline base64 bytes, or filePath with a file inside allowedDir'
    );
  }
  if (!hasData && !hasPath) {
    return toError(
      'E_VALIDATION',
      'image insert needs an image source (neither fileData nor filePath present)',
      'pass fileData with inline base64 image bytes, or filePath with a file inside allowedDir (batch --allowed-dir / execute_batch allowedDir)'
    );
  }
  return null;
}

/**
 * Per-series plot/axis selectors are combo-only. An update without an explicit
 * chartType may target an existing combo, so the engine re-checks against the
 * stored part; inserts and explicit non-combo updates fail here.
 */
function chartSeriesPlotFieldsError(op: unknown): Error | null {
  if (!op || typeof op !== 'object') return null;
  const rec = op as Record<string, unknown>;
  if (rec['action'] !== 'chart') return null;
  const chartType = typeof rec['chartType'] === 'string' ? rec['chartType'] : undefined;
  const isUpdate = rec['operation'] === 'update';
  if (isUpdate && chartType === undefined) return null;
  if (chartType === 'combo') return null;
  if (!Array.isArray(rec['series'])) return null;
  const index = rec['series'].findIndex((entry) => {
    if (!entry || typeof entry !== 'object') return false;
    return (entry as Record<string, unknown>)['chartType'] !== undefined
      || (entry as Record<string, unknown>)['axis'] !== undefined;
  });
  if (index < 0) return null;
  return toError(
    'E_VALIDATION',
    `chart series[${index}] carries chartType/axis outside a combo chart`,
    'omit series[].chartType and series[].axis for single-plot charts, or set chartType:"combo" with at least one column and one line series'
  );
}

/**
 * Strict single-operation validator. Returns ok:true with the parsed op or
 * throws with a remedy (never a silent fallback, never null).
 */
export function validateOperationStrict(op: unknown): { ok: true; operation: Operation } {
  const received =
    op !== null && typeof op === 'object' && 'action' in (op as Record<string, unknown>)
      ? String((op as Record<string, unknown>)['action'])
      : typeof op;
  const parsed = OperationSchema.safeParse(op);
  if (parsed.success) {
    const extra = insertBodyAnchorDetail(op);
    if (extra) {
      throw new Error(
        `E_VALIDATION: invalid operation (action '${received}'). ${extra}. ` +
          `Remedy: pass anchor + position ('before'|'after') for body inserts (re-read node ids via \`astra inspect\`), ` +
          `or set target:'header'/'footer' to append without an anchor.`
      );
    }
    const sourceError = imageInsertSourceError(op);
    if (sourceError) throw sourceError;
    const comboError = chartSeriesPlotFieldsError(op);
    if (comboError) throw comboError;
    const paramsError = unknownParamsError(op, NESTED_PARAM_KEYS);
    if (paramsError) throw paramsError;
    return { ok: true, operation: parsed.data };
  }
  const detail = formatZodIssues(parsed.error.issues);
  throw new Error(
    `E_VALIDATION: invalid operation (action '${received}'). ${detail}. ` +
      `Remedy: use one of insert|replace|delete|table|chart|image|style|list|namedStyle|headerFooter|field|equation|callout|comment|hyperlink|bookmark|footnote|endnote|watermark|pageBorder|pageSetup|section|coverPage|move|clearFormat|trackedChange|metadata ` +
      `with its required fields (e.g. replace needs {action:'replace',target,text}; insert needs {action:'insert',type,anchor,position}). ` +
      `Re-read node ids via \`astra inspect\` and retry.`
  );
}

/**
 * Strict batch validator. Returns ok:true with parsed ops or throws with remedy.
 */
export function validateOperationsStrict(ops: unknown): { ok: true; operations: Operation[] } {
  if (!Array.isArray(ops) || ops.length === 0) {
    throw new Error(
      `E_VALIDATION: "operations" must be a non-empty array (got ${Array.isArray(ops) ? 'empty array' : typeof ops}). ` +
        `Remedy: pass { operations: [{ action: 'replace', target: '<nodeId>', text: '...' }], sequential: true }. ` +
        `Single edit = 1-item batch.`
    );
  }
  const failures: string[] = [];
  const parsedOps: Operation[] = [];
  for (let i = 0; i < ops.length; i++) {
    const parsed = OperationSchema.safeParse(ops[i]);
    if (!parsed.success) {
      failures.push(`operations[${i}]: ${formatZodIssues(parsed.error.issues)}`);
      continue;
    }
    const extra = insertBodyAnchorDetail(ops[i]);
    if (extra) {
      failures.push(`operations[${i}]: ${extra}`);
      continue;
    }
    const laterError =
      imageInsertSourceError(ops[i]) ?? chartSeriesPlotFieldsError(ops[i]) ?? unknownParamsError(ops[i], NESTED_PARAM_KEYS);
    if (laterError) {
      failures.push(`operations[${i}]: ${laterError.message}`);
      continue;
    }
    parsedOps.push(parsed.data);
  }
  if (failures.length > 0) {
    throw new Error(
      `E_VALIDATION: ${failures.length} invalid operation(s). ${failures.slice(0, 3).join(' | ')}${failures.length > 3 ? ` | …+${failures.length - 3} more` : ''}. ` +
        `Remedy: fix the listed operations to their canonical shapes (see validateOperationStrict remedy) and retry the batch.`
    );
  }
  return { ok: true, operations: parsedOps };
}
