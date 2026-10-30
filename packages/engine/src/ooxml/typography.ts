/**
 * ISO/IEC 29500 (ECMA-376) Typographic Linter & Formatting Engine
 * 
 * Enforces commercial publication-grade typography rules across OOXML documents:
 * 1. Heading keepNext (prevents orphaned headings at page bottoms)
 * 2. Table row cantSplit (prevents awkward row splitting across page breaks)
 * 3. Table header repeat (tblHeader on first row for multi-page tables)
 * 4. Paragraph widow/orphan control
 * 
 * Guaranteed schema-valid via sortChildrenBySchema and 100% transparent/opt-out.
 */

import { createWordElement } from './ns';
import { sortChildrenBySchema, PPR_CHILD_ORDER, TRPR_CHILD_ORDER } from './schema-order';

export interface TypographicFix {
  rule: 'keepNext' | 'cantSplit' | 'tblHeader' | 'widowControl';
  targetId: string;
  element: string;
  description: string;
}

export interface TypographyLintResult {
  applied: boolean;
  fixes: TypographicFix[];
}

export interface TypographyOptions {
  enabled?: boolean; // default true
  keepNextHeadings?: boolean; // default true
  cantSplitRows?: boolean; // default true
  repeatTableHeader?: boolean; // default true
  widowControl?: boolean; // default true
}

/**
 * Lints and applies commercial typographic rules to a WordprocessingML Document DOM.
 */
export function lintAndFixTypography(doc: Document, options?: TypographyOptions): TypographyLintResult {
  const opts: Required<TypographyOptions> = {
    enabled: options?.enabled ?? true,
    keepNextHeadings: options?.keepNextHeadings ?? true,
    cantSplitRows: options?.cantSplitRows ?? true,
    repeatTableHeader: options?.repeatTableHeader ?? true,
    widowControl: options?.widowControl ?? true,
  };

  if (!opts.enabled) {
    return { applied: false, fixes: [] };
  }

  const fixes: TypographicFix[] = [];

  // Helper to find or create w:pPr
  function getOrCreatePPr(p: Element): Element {
    let pPr = p.getElementsByTagName('w:pPr')[0];
    if (!pPr) {
      pPr = createWordElement(doc, 'w:pPr');
      p.insertBefore(pPr, p.firstChild);
    }
    return pPr;
  }

  // Helper to find or create w:trPr
  function getOrCreateTrPr(tr: Element): Element {
    let trPr = tr.getElementsByTagName('w:trPr')[0];
    if (!trPr) {
      trPr = createWordElement(doc, 'w:trPr');
      tr.insertBefore(trPr, tr.firstChild);
    }
    return trPr;
  }

  // 1. Process Headings & Widow Control on Paragraphs
  const pNodes = doc.getElementsByTagName('w:p');
  for (let i = 0; i < pNodes.length; i++) {
    const p = pNodes[i];
    const nodeId = p.getAttribute('data-astra-id') || p.getAttribute('w14:paraId') || `p_${i + 1}`;

    const pPr = p.getElementsByTagName('w:pPr')[0];
    let isHeading = false;

    if (pPr) {
      const pStyle = pPr.getElementsByTagName('w:pStyle')[0];
      if (pStyle) {
        const val = pStyle.getAttribute('w:val') || '';
        if (/^heading\s*\d/i.test(val) || /^heading\d/i.test(val) || /^title/i.test(val) || /^subtitle/i.test(val)) {
          isHeading = true;
        }
      }
      if (pPr.getElementsByTagName('w:outlineLvl').length > 0) {
        isHeading = true;
      }
    }

    // Rule 1: keepNext on headings
    if (isHeading && opts.keepNextHeadings) {
      const activePPr = getOrCreatePPr(p);
      const keepNextNodes = activePPr.getElementsByTagName('w:keepNext');
      if (keepNextNodes.length === 0) {
        const keepNext = createWordElement(doc, 'w:keepNext');
        activePPr.appendChild(keepNext);
        sortChildrenBySchema(activePPr, PPR_CHILD_ORDER);
        fixes.push({
          rule: 'keepNext',
          targetId: nodeId,
          element: 'w:p',
          description: 'Added w:keepNext to heading paragraph to prevent orphan at page bottom'
        });
      }
    }

    // Rule 4: widowControl on paragraphs
    if (opts.widowControl) {
      const activePPr = getOrCreatePPr(p);
      const widowNodes = activePPr.getElementsByTagName('w:widowControl');
      if (widowNodes.length === 0) {
        const widowControl = createWordElement(doc, 'w:widowControl');
        activePPr.appendChild(widowControl);
        sortChildrenBySchema(activePPr, PPR_CHILD_ORDER);
        fixes.push({
          rule: 'widowControl',
          targetId: nodeId,
          element: 'w:p',
          description: 'Added w:widowControl to paragraph to prevent single orphan lines'
        });
      }
    }
  }

  // 2. Process Tables: Row splitting and header repeating
  const tblNodes = doc.getElementsByTagName('w:tbl');
  for (let t = 0; t < tblNodes.length; t++) {
    const tbl = tblNodes[t];
    const tblId = tbl.getAttribute('data-astra-id') || `tbl_${t + 1}`;

    const trNodes = tbl.getElementsByTagName('w:tr');
    for (let r = 0; r < trNodes.length; r++) {
      const tr = trNodes[r];
      // Only direct children rows or rows belonging to this table
      if (tr.parentNode !== tbl) {
        let parent = tr.parentNode;
        let isNested = false;
        while (parent && parent !== tbl) {
          if ((parent as Element).tagName === 'w:tbl') {
            isNested = true;
            break;
          }
          parent = parent.parentNode;
        }
        if (isNested) continue;
      }

      const rowId = tr.getAttribute('data-astra-id') || `${tblId}_tr_${r + 1}`;

      // Rule 2: cantSplit on table rows
      if (opts.cantSplitRows) {
        const trPr = getOrCreateTrPr(tr);
        const cantSplitNodes = trPr.getElementsByTagName('w:cantSplit');
        if (cantSplitNodes.length === 0) {
          const cantSplit = createWordElement(doc, 'w:cantSplit');
          trPr.appendChild(cantSplit);
          sortChildrenBySchema(trPr, TRPR_CHILD_ORDER);
          fixes.push({
            rule: 'cantSplit',
            targetId: rowId,
            element: 'w:tr',
            description: 'Added w:cantSplit to table row to prevent awkward page break splits'
          });
        }
      }

      // Rule 3: repeatTableHeader on the first row
      if (r === 0 && opts.repeatTableHeader) {
        const trPr = getOrCreateTrPr(tr);
        const tblHeaderNodes = trPr.getElementsByTagName('w:tblHeader');
        if (tblHeaderNodes.length === 0) {
          const tblHeader = createWordElement(doc, 'w:tblHeader');
          trPr.appendChild(tblHeader);
          sortChildrenBySchema(trPr, TRPR_CHILD_ORDER);
          fixes.push({
            rule: 'tblHeader',
            targetId: rowId,
            element: 'w:tr',
            description: 'Added w:tblHeader to table header row to repeat on multi-page tables'
          });
        }
      }
    }
  }

  return {
    applied: fixes.length > 0,
    fixes
  };
}
