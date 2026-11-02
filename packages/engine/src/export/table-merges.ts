/**
 * Merged cells of a `w:tbl`, for the HTML and Markdown exports.
 *
 * OOXML writes a merge in two ways: `w:gridSpan` on one cell (a horizontal merge), and `w:vMerge` /
 * `w:hMerge` markers on several cells (`restart` starts a merge; `continue`, a bare marker or an empty
 * value joins the merge that began before it). The engine's own `mergeCells` writes the markers; Word
 * writes `gridSpan` and `vMerge`. Both are read here into one model: each physical cell is either an
 * origin that owns a rectangle (`colspan` x `rowspan`) or a covered cell that belongs to the rectangle
 * of an origin.
 *
 * A covered cell can still hold text in the XML (the engine's `mergeCells` keeps it, and a renderer shows
 * it inside the merged cell), so each shown cell lists the covered cells that joined it in `absorbed`, and
 * the exports write their text after its own.
 */

export interface MergeCell {
  /** The `w:tc` element. */
  tc: Element;
  /** The `w:gridSpan` of this cell alone (1 when absent). */
  span: number;
  /** Columns the cell spans, merged neighbours included. */
  colspan: number;
  /** Rows the cell spans, merged cells below included. */
  rowspan: number;
  /** True when the cell belongs to the rectangle of an origin cell and is not shown. */
  covered: boolean;
  /** The covered cells (`w:tc`) that joined this cell's rectangle, in document order. */
  absorbed: Element[];
}

export interface TableMerges {
  /** One entry per `w:tr` that has cells, one `MergeCell` per `w:tc`, in document order. */
  rows: MergeCell[][];
  /** True when any cell spans more than one column or row, or is covered. */
  hasMerges: boolean;
  /** Merged ranges: shown cells that span more than one column or row. */
  mergedRanges: number;
}

function directChildren(parent: Element, tag: string): Element[] {
  const out: Element[] = [];
  const kids = parent.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const child = kids[i] as Element;
    if (child.nodeType === 1 && child.tagName === tag) out.push(child);
  }
  return out;
}

/**
 * The elements named `tag` that belong to `root`: its direct children, and those inside wrappers such as
 * `w:sdt` / `w:sdtContent` / `w:ins` (a row or cell can sit in a content control). A nested table is not
 * entered, so its rows and cells stay its own.
 */
function ownElements(root: Element, tag: string): Element[] {
  const out: Element[] = [];
  const walk = (parent: Element): void => {
    const kids = parent.childNodes;
    for (let i = 0; i < kids.length; i++) {
      const child = kids[i] as Element;
      if (child.nodeType !== 1) continue;
      if (child.tagName === tag) out.push(child);
      else if (child.tagName !== 'w:tbl' && child.tagName !== 'w:tc' && child.tagName !== 'w:tcPr' && child.tagName !== 'w:trPr') walk(child);
    }
  };
  walk(root);
  return out;
}

type MarkerState = 'none' | 'restart' | 'continue';

function markerState(tcPr: Element | undefined, tag: string): MarkerState {
  if (!tcPr) return 'none';
  const el = directChildren(tcPr, tag)[0];
  if (!el) return 'none';
  return (el.getAttribute('w:val') ?? '').toLowerCase() === 'restart' ? 'restart' : 'continue';
}

function positiveInt(raw: string | null | undefined, fallback: number): number {
  const n = raw === null || raw === undefined ? NaN : parseInt(raw, 10);
  return Number.isFinite(n) && n >= 1 ? n : fallback;
}

export function readTableMerges(tbl: Element): TableMerges {
  const rows: MergeCell[][] = [];
  // For each logical column, the shown cell that starts a vertical merge and runs down into the current row.
  const owners = new Map<number, MergeCell>();
  // The last row in which a shown cell grew downward, so a wide cell counts a row once.
  const grewInRow = new Map<MergeCell, number>();
  let rowIndex = -1;

  for (const tr of ownElements(tbl, 'w:tr')) {
    const tcs = ownElements(tr, 'w:tc');
    if (tcs.length === 0) continue;
    rowIndex++;
    const trPr = directChildren(tr, 'w:trPr')[0];
    const gridBefore = trPr ? directChildren(trPr, 'w:gridBefore')[0] : undefined;
    let col = 1 + (gridBefore ? positiveInt(gridBefore.getAttribute('w:val'), 0) : 0);

    const rowCells: MergeCell[] = [];
    // Column claims of this row, applied to `owners` after the row is read.
    const claims: Array<{ from: number; to: number; startsVertical: boolean }> = [];
    // The shown cell that a horizontal continuation on the next cell would join. `counts` is false when
    // that cell is already inside a rectangle whose width was counted from its own origin.
    let chain: { head: MergeCell; counts: boolean } | undefined;

    for (const tc of tcs) {
      const tcPr = directChildren(tc, 'w:tcPr')[0];
      const gridSpan = tcPr ? directChildren(tcPr, 'w:gridSpan')[0] : undefined;
      const span = gridSpan ? positiveInt(gridSpan.getAttribute('w:val'), 1) : 1;
      const h = markerState(tcPr, 'w:hMerge');
      const v = markerState(tcPr, 'w:vMerge');
      const cell: MergeCell = { tc, span, colspan: span, rowspan: 1, covered: false, absorbed: [] };
      const from = col;
      const to = col + span - 1;

      const owner = v === 'continue' ? owners.get(col) : undefined;
      if (owner) {
        // Joins the rectangle that comes down from the row above.
        cell.covered = true;
        if (grewInRow.get(owner) !== rowIndex) {
          owner.rowspan += 1;
          grewInRow.set(owner, rowIndex);
        }
        owner.absorbed.push(tc);
        chain = { head: owner, counts: false };
      } else if (h === 'continue' && chain) {
        // Joins the cell on its left.
        cell.covered = true;
        if (chain.counts) chain.head.colspan += span;
        chain.head.absorbed.push(tc);
      } else {
        // A shown cell. A marker that has nothing to join is read as a plain cell.
        claims.push({ from, to, startsVertical: v === 'restart' });
        chain = { head: cell, counts: true };
      }
      rowCells.push(cell);
      col = to + 1;
    }

    // A shown cell that does not start a vertical merge ends any rectangle above in its columns; one that
    // starts a vertical merge takes its columns for the rows below. Covered cells change nothing.
    let claimIndex = 0;
    for (const cell of rowCells) {
      if (cell.covered) continue;
      const claim = claims[claimIndex++];
      for (let c = claim.from; c <= claim.to; c++) {
        if (claim.startsVertical) owners.set(c, cell);
        else owners.delete(c);
      }
    }
    rows.push(rowCells);
  }

  let mergedRanges = 0;
  let hasMerges = false;
  for (const row of rows) {
    for (const cell of row) {
      if (cell.covered) hasMerges = true;
      else if (cell.colspan > 1 || cell.rowspan > 1) {
        hasMerges = true;
        mergedRanges++;
      }
    }
  }
  return { rows, hasMerges, mergedRanges };
}
