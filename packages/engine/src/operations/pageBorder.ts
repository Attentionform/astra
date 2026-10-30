import { DocumentAST } from '../ast';
import { BaseOperation } from './base';
import { DocxParser, upsertSectPrChild } from '../parser';
import { createWordElement } from '../ooxml/ns';
import { sectionPosition } from './section-number';
import { toError } from '../errors';

export interface PageBorderOperation extends BaseOperation {
  action: 'pageBorder';
  style?: string;       // e.g. 'single', 'double', 'dashed', 'triple' (default 'single')
  sz?: number;          // 1/8pt (e.g. 12 = 1.5pt, 24 = 3pt)
  color?: string;       // HEX without '#'
  marginOffset?: 'page' | 'text'; // default 'page'
  sectionIndex?: number;
}

export function executePageBorder(op: PageBorderOperation, ast: DocumentAST, parser?: DocxParser): void {
  // The number names a document section (a w:sectPr); the AST may hold fewer
  // sections than the file, so the AST lookup below falls back to the first.
  const sectionCount = Math.max(parser?.getSectPrElements().length ?? 0, ast.sections.length, 1);
  const sectionIdx = sectionPosition('pageBorder', op.sectionIndex, sectionCount) ?? 0;
  const section = ast.sections[sectionIdx] ?? ast.sections[0];

  if (!section) throw new Error("Section not found for page border");

  const lastNode = section.nodes[section.nodes.length - 1];
  const doc = lastNode?.xml?.ownerDocument || ast.sections[0]?.nodes[0]?.xml?.ownerDocument;

  if (doc) {
    // OP-12: scope strictly to THIS section's sectPr — never sectPr[0].
    let sectPr: Element | null = null;
    try {
      const sectPrs = parser?.getSectPrElements() ?? [];
      if (sectPrs.length > sectionIdx) sectPr = sectPrs[sectionIdx];
    } catch { sectPr = null; }
    if (!sectPr) {
      for (const node of section.nodes) {
        const el = node.xml as Element;
        if (!el || typeof el.getElementsByTagName !== 'function') continue;
        if (el.tagName === 'w:sectPr') { sectPr = el; break; }
        const local = el.getElementsByTagName('w:sectPr')[0] as Element | undefined;
        if (local) { sectPr = local; break; }
      }
    }
    if (!sectPr) {
      const all = doc.getElementsByTagName('w:sectPr');
      if (all.length === 1) sectPr = all[0] as Element;
      else if (all.length > sectionIdx) sectPr = all[sectionIdx] as Element;
    }
    if (!sectPr && parser) {
      throw toError(
        'E_VALIDATION',
        `Cannot set the page border: the document has no section properties (w:sectPr) for section ${sectionIdx + 1}.`,
        'Re-read the document with inspect_document and check the section count; if it looks wrong, save and reopen the file, then retry.',
      );
    }
    if (sectPr) {
      const pgBorders = upsertSectPrChild(sectPr, 'w:pgBorders');
      pgBorders.setAttribute('w:offsetFrom', op.marginOffset || 'page');

      const style = op.style || 'single';
      const sz = String(op.sz || 12);
      const color = (op.color || 'CBD5E1').replace(/^#/, '');

      ['top', 'left', 'bottom', 'right'].forEach(side => {
        let sideEl = pgBorders.getElementsByTagName(`w:${side}`)[0];
        if (!sideEl) {
          sideEl = createWordElement(doc, `w:${side}`);
          pgBorders.appendChild(sideEl);
        }
        sideEl.setAttribute('w:val', style);
        sideEl.setAttribute('w:sz', sz);
        sideEl.setAttribute('w:space', '24');
        sideEl.setAttribute('w:color', color);
      });
    }
  }

  section.version += 1;
}
