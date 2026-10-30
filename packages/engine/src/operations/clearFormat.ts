import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node } from '../ast';
import { BaseOperation, validateVersion, findNode } from './base';
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, appendInSchemaOrder } from '../ooxml/schema-order';

export interface ClearFormatOperation extends BaseOperation {
  action: 'clearFormat';
  target: string | string[];
  params?: Record<string, unknown>;
}

export interface ClearFormatResult {
  cleared: string[];
}

/**
 * Batch 9b: clear-formatting backend op (Eraser prerequisite).
 *
 * Combined Word Ctrl+Space + Ctrl+Q semantics:
 *   - Ctrl+Space clears CHARACTER (run-level) direct formatting: every child
 *     of w:rPr. Canonical list is RPR_CHILD_ORDER (ECMA-376 §17.3.2.28) —
 *     w:b/w:bCs/w:i/w:iCs/w:caps/w:smallCaps/w:color/w:highlight/w:u/
 *     w:strike/w:dstrike/w:sz/w:szCs/w:rFonts/w:vertAlign plus the long tail
 *     the old 13-tag hand list missed (w:shd/w:bdr/w:kern/w:position/
 *     w:spacing/w:w/w:effect/w:fitText/w:rtl/w:cs/w:em/w:lang/
 *     w:eastAsianLayout/w:specVanish/w:oMath/w:rStyle/w:outline/w:shadow/
 *     w:emboss/w:imprint/w:noProof/w:snapToGrid/w:vanish/w:webHidden/etc.).
 *   - Ctrl+Q clears PARAGRAPH direct formatting: w:jc (alignment), w:ind
 *     (indentation), w:pBdr (paragraph borders), w:spacing (before/after/
 *     line). pStyle resets to Normal; numbering/tabs/shading and other pPr
 *     peers are left alone (narrow Ctrl+Q scope pinned by tests).
 * Paragraph-mark defaults (pPr/w:rPr) get the same run-strip so the pilcrow
 * carries no formatting. Text is never touched.
 */
const STRIP_RPR_TAGS: Set<string> = new Set<string>(RPR_CHILD_ORDER);

/** Ctrl+Q paragraph-direct-formatting peers cleared alongside pStyle reset. */
const STRIP_PPR_TAGS: Set<string> = new Set<string>(['w:jc', 'w:ind', 'w:pBdr', 'w:spacing']);

function hasElementChildren(el: Element): boolean {
  for (let i = 0; i < el.childNodes.length; i++) {
    if (el.childNodes[i].nodeType === 1) return true;
  }
  return false;
}

function stripRunProps(rPr: Element): void {
  const victims: Element[] = [];
  for (let i = 0; i < rPr.childNodes.length; i++) {
    const child = rPr.childNodes[i] as Element;
    if (child.nodeType === 1 && STRIP_RPR_TAGS.has((child as Element).tagName)) {
      victims.push(child as Element);
    }
  }
  for (const v of victims) rPr.removeChild(v);
}

function clearParagraphElement(pEl: Element, doc: Document): void {
  // 1. Strip every run in this paragraph.
  const runs = pEl.getElementsByTagName('w:r');
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    // getElementsByTagName on the run returns its own rPr first (runs nest
    // nothing else that owns an rPr except the paragraph-mark pPr/rPr, which
    // is NOT under w:r — so the first hit is the run's own).
    const rPr = run.getElementsByTagName('w:rPr')[0];
    if (rPr && rPr.parentNode === run) {
      stripRunProps(rPr);
      if (!hasElementChildren(rPr)) run.removeChild(rPr);
    }
  }

  // 2. Paragraph properties: reset pStyle to Normal, drop Ctrl+Q peers
  //    (jc/ind/pBdr/spacing), strip paragraph-mark run defaults.
  let pPr = pEl.getElementsByTagName('w:pPr')[0];
  // getElementsByTagName may return a nested pPr only for w:p roots there is
  // exactly one direct child; guard the parent for table-cell edge shapes.
  if (pPr && pPr.parentNode !== pEl) {
    pPr = null as unknown as Element;
    for (let i = 0; i < pEl.childNodes.length; i++) {
      const c = pEl.childNodes[i] as Element;
      if (c.nodeType === 1 && (c as Element).tagName === 'w:pPr') {
        pPr = c as Element;
        break;
      }
    }
  }
  if (!pPr) {
    pPr = createWordElement(doc, 'w:pPr');
    pEl.insertBefore(pPr, pEl.firstChild);
  }
  let pStyle = pPr.getElementsByTagName('w:pStyle')[0];
  if (!pStyle || pStyle.parentNode !== pPr) {
    // A nested pStyle (shouldn't happen) is ignored; ensure a direct one.
    pStyle = null as unknown as Element;
    for (let i = 0; i < pPr.childNodes.length; i++) {
      const c = pPr.childNodes[i] as Element;
      if (c.nodeType === 1 && (c as Element).tagName === 'w:pStyle') {
        pStyle = c as Element;
        break;
      }
    }
    if (!pStyle) {
      pStyle = createWordElement(doc, 'w:pStyle');
      appendInSchemaOrder(pPr, pStyle, PPR_CHILD_ORDER);
    }
  }
  pStyle.setAttribute('w:val', 'Normal');

  // Ctrl+Q: drop paragraph direct-formatting peers (jc/ind/pBdr/spacing).
  // Collect-then-remove so live NodeLists can't skip siblings.
  const pVictims: Element[] = [];
  for (let i = 0; i < pPr.childNodes.length; i++) {
    const child = pPr.childNodes[i] as Element;
    if (child.nodeType === 1 && STRIP_PPR_TAGS.has((child as Element).tagName)) {
      pVictims.push(child as Element);
    }
  }
  for (const v of pVictims) pPr.removeChild(v);

  const markRPr = pPr.getElementsByTagName('w:rPr')[0];
  if (markRPr && markRPr.parentNode === pPr) {
    stripRunProps(markRPr);
    if (!hasElementChildren(markRPr)) pPr.removeChild(markRPr);
  }
}

export function executeClearFormat(op: ClearFormatOperation, ast: DocumentAST): ClearFormatResult {
  const raw = (op.params || {}) as Record<string, unknown>;
  const resolvedTarget = (raw['target'] !== undefined ? (raw['target'] as string | string[]) : op.target) as
    | string
    | string[];
  const targets = Array.isArray(resolvedTarget) ? resolvedTarget : [resolvedTarget];
  if (targets.length === 0 || !targets[0]) throw new Error('ClearFormat operation requires a target');

  const cleared: string[] = [];
  for (const target of targets) {
    const targetNode: Node | undefined = findNode(ast, target);
    if (!targetNode) throw new Error(`Node ${target} not found`);
    validateVersion(targetNode, op.expectedVersion);

    const doc = targetNode.xml.ownerDocument;
    if (!doc) throw new Error(`Node ${target} has no owner document; cannot clear formatting`);

    const root = targetNode.xml as Element;
    if (root.tagName === 'w:p') {
      clearParagraphElement(root, doc);
    } else {
      // Non-paragraph node (e.g. table): clear every enclosed paragraph plus
      // any runs outside paragraphs (defensive), then fall back to a direct
      // run sweep so no shaped target keeps formatting.
      const paras = root.getElementsByTagName('w:p');
      for (let i = 0; i < paras.length; i++) clearParagraphElement(paras[i], doc);
      if (paras.length === 0) {
        const runs = root.getElementsByTagName('w:r');
        for (let i = 0; i < runs.length; i++) {
          const run = runs[i];
          const rPr = run.getElementsByTagName('w:rPr')[0];
          if (rPr && rPr.parentNode === run) {
            stripRunProps(rPr);
            if (!hasElementChildren(rPr)) run.removeChild(rPr);
          }
        }
      }
    }

    targetNode.metadata.namedStyle = 'Normal';
    targetNode.version += 1;
    cleared.push(target);
  }
  return { cleared };
}
