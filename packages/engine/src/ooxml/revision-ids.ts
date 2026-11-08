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

/**
 * Highest numeric id on a revision element or move range marker. Non-numeric ids are ignored.
 * A digit string that a JavaScript number cannot hold exactly gives Infinity, never a rounded value.
 */
export function maxRevisionId(document: Document): number {
  let max = 0;
  for (const tag of [...REVISION_TAGS, ...MOVE_MARKER_TAGS]) {
    for (const el of Array.from(document.getElementsByTagNameNS(W_NS, tag))) {
      const value = el.getAttributeNS(W_NS, 'id');
      if (value !== null && /^\d+$/.test(value)) {
        const n = Number(value);
        if (!Number.isSafeInteger(n)) return Infinity;
        max = Math.max(max, n);
      }
    }
  }
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
