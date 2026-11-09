/** WordprocessingML revision elements whose ids must be reserved for new revisions. */
export const REVISION_TAGS = [
  'ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'sectPrChange',
  'tblPrChange', 'tblPrExChange', 'tblGridChange', 'tPrChange', 'trPrChange', 'tcPrChange', 'numberingChange',
  'cellIns', 'cellDel', 'cellMerge',
  'customXmlInsRangeStart', 'customXmlInsRangeEnd', 'customXmlDelRangeStart', 'customXmlDelRangeEnd',
  'customXmlMoveFromRangeStart', 'customXmlMoveFromRangeEnd', 'customXmlMoveToRangeStart', 'customXmlMoveToRangeEnd',
] as const;

/**
 * Move range markers carry an id from the same space as revisions but are not revisions
 * themselves, so they stay out of REVISION_TAGS (that list also drives counting and
 * resolution). New ids must still start above them.
 */
export const MOVE_MARKER_TAGS = ['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd'] as const;

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const REVISION_NAMES: ReadonlySet<string> = new Set<string>(REVISION_TAGS);
const REVISION_AND_MARKER_NAMES: ReadonlySet<string> = new Set<string>([...REVISION_TAGS, ...MOVE_MARKER_TAGS]);

/**
 * Calls `visit` for every WordprocessingML element in `document` whose local name is in `names`, in one walk.
 * One `getElementsByTagNameNS` call per name walks the whole document each time (29 walks, 17 ms on a document of
 * 3,000 paragraphs, against 0.3 ms for one). The walk is iterative, so a deeply nested document cannot overflow
 * the stack. Order is not specified.
 */
export function walkWordElements(document: Document, names: ReadonlySet<string>, visit: (element: Element) => void): void {
  const root = document.documentElement;
  if (!root) return;
  const stack: Element[] = [root];
  while (stack.length > 0) {
    const element = stack.pop() as Element;
    if (element.namespaceURI === W_NS && names.has(element.localName)) visit(element);
    for (let child = element.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === 1) stack.push(child as Element);
    }
  }
}

/** Number of revision elements in `document`, including kinds that cannot be resolved. Move range markers are not counted. */
export function countRevisionElements(document: Document): number {
  let count = 0;
  walkWordElements(document, REVISION_NAMES, () => { count++; });
  return count;
}

/**
 * Highest numeric id on a revision element or move range marker. Non-numeric ids are ignored.
 * A digit string that a JavaScript number cannot hold exactly gives Infinity, never a rounded value.
 */
export function maxRevisionId(document: Document): number {
  let max = 0;
  walkWordElements(document, REVISION_AND_MARKER_NAMES, (el) => {
    const value = el.getAttributeNS(W_NS, 'id');
    if (value !== null && /^\d+$/.test(value)) {
      const n = Number(value);
      max = Number.isSafeInteger(n) ? Math.max(max, n) : Infinity;
    }
  });
  return max;
}

/** First id to hand out after `highest`. At or above the exact limit it is 2^53, which takeRevisionId refuses. */
export function firstFreeRevisionId(highest: number): number {
  return highest >= Number.MAX_SAFE_INTEGER ? Number.MAX_SAFE_INTEGER + 1 : highest + 1;
}

/** Hand out the next revision id. Fails with no side effect when the id cannot be represented exactly. */
export function takeRevisionId(counter: { nextChangeId: number }): number {
  const id = counter.nextChangeId;
  if (!Number.isSafeInteger(id)) {
    throw Object.assign(new Error(
      `E_VALIDATION: Cannot allocate a revision id: the document already uses revision ids at or beyond ${Number.MAX_SAFE_INTEGER}, the largest id that can be numbered exactly. ` +
      'Remedy: accept or reject the revisions with the highest ids in Word, or edit without tracking.'
    ), {code: 'E_VALIDATION'});
  }
  counter.nextChangeId = id + 1;
  return id;
}
