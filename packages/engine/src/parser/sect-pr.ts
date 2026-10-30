import { createWordElement } from '../ooxml/ns';

// Canonical CT_SectPr child order (subset covering elements we create/mutate).
export const SECT_PR_CHILD_ORDER: string[] = [
  'w:headerReference', 'w:footerReference', 'w:footnotePr', 'w:endnotePr',
  'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc', 'w:pgBorders', 'w:lnNumType',
  'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign', 'w:noEndnote',
  'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid',
  'w:printerSettings'
];

/**
 * Finds or creates a direct child of a <w:sectPr>, inserting it at the
 * schema-correct position so Word does not reject the document.
 */
export function upsertSectPrChild(sectPr: Element, tagName: string): Element {
  for (let i = 0; i < sectPr.childNodes.length; i++) {
    const child = sectPr.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === tagName) {
      return child as Element;
    }
  }

  const doc = sectPr.ownerDocument;
  if (!doc) throw new Error(`Cannot create ${tagName}: sectPr has no owner document`);
  const created = tagName.startsWith('w:') ? createWordElement(doc, tagName) : doc.createElement(tagName);
  const orderOf = (tag: string) => {
    const idx = SECT_PR_CHILD_ORDER.indexOf(tag);
    return idx === -1 ? SECT_PR_CHILD_ORDER.length : idx;
  };
  const targetIdx = orderOf(tagName);
  for (let i = 0; i < sectPr.childNodes.length; i++) {
    const child = sectPr.childNodes[i];
    if (child.nodeType !== 1) continue;
    if (orderOf((child as Element).tagName) > targetIdx) {
      sectPr.insertBefore(created, child);
      return created;
    }
  }
  sectPr.appendChild(created);
  return created;
}

/**
 * IN-07: geometry-aware lines-per-page derived from sectPr params
 * (pgSz/pgMar in twips; 1440 twips = 1in = 96px). Baseline 45 lines at the
 * default usable height (~864px for A4/Letter 1in margins), scaled with the
 * actual usable height so custom page sizes/margins stay renderer-compatible.
 * Exported for unit tests.
 */
export function linesPerPageFromSectPr(sectPrXml?: string): number {
  const BASE_LINES = 45;
  const BASE_USABLE_PX = 864;
  try {
    if (!sectPrXml || typeof sectPrXml !== 'string') return BASE_LINES;
    const numAttr = (tag: string, attr: string): number | null => {
      const tagRe = new RegExp(`<w:${tag}[^>]*?w:${attr}="(\\d+)"`, 'i');
      const m = sectPrXml.match(tagRe);
      if (!m) return null;
      const v = parseInt(m[1], 10);
      return Number.isFinite(v) && v > 0 ? v : null;
    };
    const h = numAttr('pgSz', 'h');
    const top = numAttr('pgMar', 'top');
    const bottom = numAttr('pgMar', 'bottom');
    if (h === null) return BASE_LINES;
    // 1440 twips = 1in = 96px ⇒ px = twips / 15.
    const pageHpx = h / 15;
    const marginPx = ((top ?? 1440) + (bottom ?? 1440)) / 15;
    const usable = pageHpx - marginPx;
    if (!Number.isFinite(usable) || usable <= 200) return BASE_LINES;
    const scaled = Math.round((BASE_LINES * usable) / BASE_USABLE_PX);
    return Math.max(10, Math.min(120, scaled));
  } catch {
    return BASE_LINES;
  }
}
