import { XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST, Node, NodeMetadata, calculateNodeChecksum } from '../ast';
import { createWordElement } from '../ooxml/ns';
import { toError } from '../errors';

export interface BaseOperation {
  action: string;
  target?: string | string[];
  expectedVersion?: number;
  [key: string]: unknown;
}

/**
 * Runtime expando chain for same-anchor FIFO 'after' inserts. The fields are
 * attached to live AST nodes at runtime (never serialized); the interface
 * documents the duck shape so call sites need no `any`.
 */
export interface InsertChainHost {
  _insertAfterChain?: unknown;
  _lastInsertAfter?: Node | undefined;
}

/**
 * Creates a <w:t> for the given text, setting xml:space="preserve" whenever the
 * text has leading/trailing whitespace so Word does not trim it away.
 */
export function createTextElement(doc: Document, text: string): Element {
  const t = createWordElement(doc, 'w:t');
  if (text.length > 0 && (/^\s/.test(text) || /\s$/.test(text))) {
    t.setAttribute('xml:space', 'preserve');
  }
  t.appendChild(doc.createTextNode(text));
  return t;
}

export function validateVersion(node: Node, expectedVersion?: number) {
  if (expectedVersion !== undefined && node.version !== expectedVersion) {
    throw Object.assign(
      new Error(`E_VALIDATION: Version mismatch on node ${node.id}. Expected ${expectedVersion}, got ${node.version}`),
      {
        code: 'E_VALIDATION' as const,
        retryable: false as const,
        remedy: `Re-read the document via inspect_document and retry with the current version (${node.version}) for node '${node.id}'.`,
      }
    );
  }
}

// ---------------------------------------------------------------------------
// Cell addresses
//
// A table is one node in the document model, so a paragraph inside a cell has
// no id of its own. `<tableId>:r<row>c<col>` names one, with the same 1-based
// coordinates `inspect` prints for the table grid (`r1c1` is the top-left
// cell). A cell with several paragraphs takes `:p<n>` (1-based) as well.
// The operations that edit a paragraph's text or formatting look their target
// up through `findNode`; it hands them a stand-in node whose `xml` is the real
// paragraph, whose `version` is the table's (so `expectedVersion` guards the
// table and each edit bumps it), and whose bookkeeping is folded back into the
// table node by `syncCellStandIns` once the operation is done.
// ---------------------------------------------------------------------------

const CELL_ADDRESS = /^(.+):r(\d+)c(\d+)(?::p(\d+))?$/;

/** Actions whose targets may be a cell address; every other action refuses one. */
export const CELL_ADDRESS_ACTIONS: ReadonlySet<string> = new Set([
  'replace',
  'style',
  'comment',
  'clearFormat',
  'namedStyle',
  'hyperlink',
  'bookmark',
]);

export interface CellAddress {
  tableId: string;
  row: number;
  col: number;
  para?: number;
}

/** Splits `<tableId>:r<row>c<col>[:p<n>]`; null for any other string. */
export function parseCellAddress(id: unknown): CellAddress | null {
  if (typeof id !== 'string') return null;
  const m = CELL_ADDRESS.exec(id);
  if (!m) return null;
  return {
    tableId: m[1],
    row: parseInt(m[2], 10),
    col: parseInt(m[3], 10),
    ...(m[4] !== undefined ? { para: parseInt(m[4], 10) } : {}),
  };
}

function findDirectNode(ast: DocumentAST, id: string): Node | undefined {
  for (const section of ast.sections ?? []) {
    const hit = section.nodes.find((n) => n.id === id);
    if (hit) return hit;
  }
  return undefined;
}

function elementChildren(parent: Element, tag: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i] as Element;
    if (child.nodeType === 1 && child.tagName === tag) out.push(child);
  }
  return out;
}

function cellAddressError(message: string, remedy: string): Error {
  return toError('E_VALIDATION', message, remedy);
}

function paragraphText(p: Element): string {
  const parts = p.getElementsByTagName('w:t');
  let text = '';
  for (let i = 0; i < parts.length; i++) text += parts[i].textContent ?? '';
  return text;
}

function cellMergeState(tc: Element): { span: number; covered: boolean } {
  const tcPr = elementChildren(tc, 'w:tcPr')[0];
  let span = 1;
  let covered = false;
  if (tcPr) {
    const gs = elementChildren(tcPr, 'w:gridSpan')[0];
    const n = gs ? parseInt(gs.getAttribute('w:val') ?? '1', 10) : 1;
    if (Number.isFinite(n) && n > 1) span = n;
    for (const tag of ['w:hMerge', 'w:vMerge']) {
      const el = elementChildren(tcPr, tag)[0];
      if (el) {
        const v = (el.getAttribute('w:val') ?? '').toLowerCase();
        if (v === '' || v === 'continue') covered = true;
      }
    }
  }
  return { span, covered };
}

/** The cell (row, first logical column) at a logical column of a row, with its merge state. */
function cellAtColumn(
  row: Element,
  col: number
): { tc: Element; start: number; state: { span: number; covered: boolean } } | undefined {
  let logical = 0;
  for (const tc of elementChildren(row, 'w:tc')) {
    const state = cellMergeState(tc);
    if (col >= logical + 1 && col <= logical + state.span) return { tc, start: logical + 1, state };
    logical += state.span;
  }
  return undefined;
}

function isContinue(tc: Element, tag: 'w:vMerge' | 'w:hMerge'): boolean {
  const tcPr = elementChildren(tc, 'w:tcPr')[0];
  const el = tcPr ? elementChildren(tcPr, tag)[0] : undefined;
  if (!el) return false;
  const v = (el.getAttribute('w:val') ?? '').toLowerCase();
  return v === '' || v === 'continue';
}

/** Where the merge that covers the cell at (row, col) starts: up through vMerge, then left through hMerge. */
function mergeRestart(rows: Element[], row: number, col: number): { row: number; col: number } | undefined {
  let r = row;
  let c = col;
  for (let guard = 0; guard <= rows.length; guard++) {
    const at = cellAtColumn(rows[r - 1], c);
    if (!at || !isContinue(at.tc, 'w:vMerge') || r === 1) break;
    r -= 1;
  }
  for (let guard = 0; guard <= 64; guard++) {
    const at = cellAtColumn(rows[r - 1], c);
    if (!at || !isContinue(at.tc, 'w:hMerge') || c === 1) break;
    c = Math.max(1, at.start - 1);
  }
  const found = cellAtColumn(rows[r - 1], c);
  return found && !found.state.covered ? { row: r, col: found.start } : undefined;
}

function resolveCellParagraph(table: Node, addr: CellAddress): Element {
  const label = `${addr.tableId}:r${addr.row}c${addr.col}`;
  const rows = elementChildren(table.xml, 'w:tr');
  if (addr.row < 1 || addr.row > rows.length) {
    throw cellAddressError(
      `Cell address ${label}: row ${addr.row} is outside table '${addr.tableId}' (it has ${rows.length} row${rows.length === 1 ? '' : 's'}).`,
      `Cell addresses are 1-based, as in the r1c1 grid inspect prints: use r1 to r${rows.length}, the same numbering the table operation's rowIndex and colIndex use.`
    );
  }
  let logical = 0;
  let cell: Element | undefined;
  let start = 0;
  let covered = false;
  for (const tc of elementChildren(rows[addr.row - 1], 'w:tc')) {
    const state = cellMergeState(tc);
    const first = logical + 1;
    const last = logical + state.span;
    if (addr.col >= first && addr.col <= last) {
      cell = tc;
      start = first;
      covered = state.covered;
      break;
    }
    logical = last;
  }
  if (!cell) {
    throw cellAddressError(
      `Cell address ${label}: column ${addr.col} is outside row ${addr.row} of table '${addr.tableId}' (it has ${logical} column${logical === 1 ? '' : 's'}).`,
      `Cell addresses are 1-based, as in the r1c1 grid inspect prints: use c1 to c${logical}, the same numbering the table operation's rowIndex and colIndex use.`
    );
  }
  if (start !== addr.col) {
    throw cellAddressError(
      `Cell address ${label}: column ${addr.col} lies inside the merged cell that starts at column ${start}.`,
      `Use ${addr.tableId}:r${addr.row}c${start}.`
    );
  }
  if (covered) {
    const restart = mergeRestart(rows, addr.row, addr.col);
    throw cellAddressError(
      `Cell address ${label} is covered by a merged cell, so it has no text of its own.`,
      restart
        ? `Use ${addr.tableId}:r${restart.row}c${restart.col}, where the merge starts (inspect marks covered cells hidden-cover).`
        : 'Address the cell where the merge starts (inspect marks covered cells hidden-cover).'
    );
  }
  const paragraphs = elementChildren(cell, 'w:p');
  if (addr.para !== undefined) {
    const chosen = paragraphs[addr.para - 1];
    if (!chosen) {
      throw cellAddressError(
        `Cell address ${label}:p${addr.para}: the cell has ${paragraphs.length} paragraph${paragraphs.length === 1 ? '' : 's'}.`,
        paragraphs.length > 0 ? `Use :p1 to :p${paragraphs.length}.` : 'The cell has no paragraph to edit; fill it with the table operation updateCell.'
      );
    }
    return chosen;
  }
  if (paragraphs.length === 1) return paragraphs[0];
  if (paragraphs.length === 0) {
    throw cellAddressError(`Cell address ${label} has no paragraph to edit.`, 'Fill the cell with the table operation updateCell first.');
  }
  const listing = paragraphs.map((p, i) => `:p${i + 1} "${paragraphText(p).slice(0, 30)}"`).join(', ');
  throw cellAddressError(
    `Cell address ${label} has ${paragraphs.length} paragraphs (${listing}); which one?`,
    `Add the paragraph number, for example ${label}:p1.`
  );
}

/**
 * Table nodes that had a paragraph edited through a stand-in, per document, until the operation ends.
 * `OperationEngine.routeOperation` drains it; a caller that runs an executor directly must call
 * `syncCellStandIns` itself.
 */
const pendingTables = new WeakMap<DocumentAST, Set<Node>>();

function standInFor(ast: DocumentAST, table: Node, addr: CellAddress, p: Element): Node {
  const flags = (needle: string): boolean => p.getElementsByTagName(needle).length > 0;
  const styleEl = p.getElementsByTagName('w:pStyle')[0];
  const metadata: NodeMetadata = {
    text: paragraphText(p),
    hasCitation: false,
    hasComment: flags('w:commentReference'),
    hasBookmark: flags('w:bookmarkStart'),
    hasTrackedChanges: flags('w:ins') || flags('w:del'),
    hasEquation: flags('m:oMath'),
    hasField: flags('w:instrText') || flags('w:fldSimple'),
    hasImage: flags('w:drawing'),
    hasHyperlink: flags('w:hyperlink'),
    ...(styleEl?.getAttribute('w:val') ? { namedStyle: styleEl.getAttribute('w:val') as string } : {}),
  };
  const standIn = {
    id: `${addr.tableId}:r${addr.row}c${addr.col}${addr.para !== undefined ? `:p${addr.para}` : ''}`,
    type: 'paragraph' as const,
    // Read-only on purpose: an executor that swapped in a new element would edit the stand-in
    // and lose the change from the table, so reassignment throws instead.
    get xml(): Element {
      return p;
    },
    metadata,
    get version(): number {
      return table.version;
    },
    set version(value: number) {
      table.version = value;
    },
    // The table's checksum covers the whole table; syncCellStandIns recomputes it.
    get checksum(): string {
      return table.checksum;
    },
    set checksum(_value: string) {
      /* recomputed from the table XML when the operation ends */
    },
  } as Node;
  let pending = pendingTables.get(ast);
  if (!pending) {
    pending = new Set<Node>();
    pendingTables.set(ast, pending);
  }
  pending.add(table);
  return standIn;
}

/**
 * Finds a node by id, or the paragraph in a table cell by cell address.
 * Throws a validation error (with the fix) when the address names a table but
 * not a usable cell; returns undefined when no such node or table exists.
 */
export function findNode(ast: DocumentAST, id: unknown): Node | undefined {
  if (typeof id !== 'string') return undefined;
  const direct = findDirectNode(ast, id);
  if (direct) return direct;
  const addr = parseCellAddress(id);
  if (!addr) return undefined;
  const table = findDirectNode(ast, addr.tableId);
  if (!table || table.type !== 'table' || (table.xml as Element)?.tagName !== 'w:tbl') return undefined;
  return standInFor(ast, table, addr, resolveCellParagraph(table, addr));
}

/**
 * Cell addresses in a table node whose paragraph text passes `test` (at most
 * `limit`), in reading order. A cell with several paragraphs reports the
 * matching paragraph (`:p<n>`); a merged-away cell is skipped.
 */
export function cellAddressesMatching(table: Node, test: (text: string) => boolean, limit = 5): string[] {
  const out: string[] = [];
  const tbl = table.xml as Element | undefined;
  if (!tbl || tbl.tagName !== 'w:tbl') return out;
  const rows = elementChildren(tbl, 'w:tr');
  for (let r = 0; r < rows.length; r++) {
    let logical = 0;
    for (const tc of elementChildren(rows[r], 'w:tc')) {
      const state = cellMergeState(tc);
      const start = logical + 1;
      logical += state.span;
      if (state.covered) continue;
      const paragraphs = elementChildren(tc, 'w:p');
      for (let i = 0; i < paragraphs.length; i++) {
        let hit = false;
        try {
          hit = test(paragraphText(paragraphs[i]));
        } catch {
          hit = false;
        }
        if (!hit) continue;
        out.push(`${table.id}:r${r + 1}c${start}${paragraphs.length > 1 ? `:p${i + 1}` : ''}`);
        if (out.length >= limit) return out;
      }
    }
  }
  return out;
}

/** Folds edits made through cell stand-ins into their table nodes (text, checksum). */
export async function syncCellStandIns(ast: DocumentAST): Promise<void> {
  const pending = pendingTables.get(ast);
  if (!pending || pending.size === 0) return;
  pendingTables.delete(ast);
  for (const table of pending) {
    const tbl = table.xml as Element;
    const xml = new XMLSerializer().serializeToString(tbl);
    table.metadata.text = paragraphText(tbl);
    // The executors set these flags on the cell's stand-in; derive them for the table the way the parser does.
    table.metadata.hasComment = xml.includes('w:commentReference');
    table.metadata.hasBookmark = xml.includes('w:bookmarkStart');
    table.metadata.hasHyperlink = xml.includes('w:hyperlink');
    table.metadata.hasTrackedChanges = tbl.getElementsByTagName('w:ins').length > 0 || tbl.getElementsByTagName('w:del').length > 0;
    table.checksum = await calculateNodeChecksum(xml);
  }
}

/**
 * Refuses a cell address on an operation that cannot act on one, naming what
 * can. Only addresses whose table exists count, so ordinary ids are untouched.
 */
export function assertNoCellAddress(action: string, fields: Record<string, unknown>, ast: DocumentAST): void {
  if (CELL_ADDRESS_ACTIONS.has(action)) return;
  const params = typeof fields['params'] === 'object' && fields['params'] !== null ? (fields['params'] as Record<string, unknown>) : {};
  const candidates: unknown[] = [fields['target'], fields['anchor'], params['target'], params['anchor']];
  for (const value of candidates.flat()) {
    const addr = parseCellAddress(value);
    if (!addr) continue;
    const table = findDirectNode(ast, addr.tableId);
    if (!table || table.type !== 'table') continue;
    throw cellAddressError(
      `'${action}' cannot act on a cell address (${String(value)}).`,
      action === 'table'
        ? `Use the table id '${addr.tableId}' with rowIndex ${addr.row} and colIndex ${addr.col}.`
        : `Cell addresses work with ${[...CELL_ADDRESS_ACTIONS].join(', ')}. For '${action}', use the table id '${addr.tableId}'.`
    );
  }
}
