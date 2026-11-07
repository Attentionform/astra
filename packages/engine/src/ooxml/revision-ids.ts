/** WordprocessingML revision elements whose ids must be reserved for new revisions. */
export const REVISION_TAGS = [
  'ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'sectPrChange',
  'tblPrChange', 'tblGridChange', 'trPrChange', 'tcPrChange', 'numberingChange',
  'cellIns', 'cellDel', 'cellMerge',
] as const;

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function maxRevisionId(document: Document): number {
  let max = 0;
  for (const tag of REVISION_TAGS) {
    for (const el of Array.from(document.getElementsByTagNameNS(W_NS, tag))) {
      const value = el.getAttributeNS(W_NS, 'id');
      if (value !== null && /^\d+$/.test(value)) {
        const n = Number(value);
        if (Number.isSafeInteger(n)) max = Math.max(max, n);
      }
    }
  }
  return max;
}
