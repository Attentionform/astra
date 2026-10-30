import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node, calculateNodeChecksum } from '../ast';
import { BaseOperation, createTextElement } from './base';
import { XMLSerializer } from '@xmldom/xmldom';
import * as crypto from 'crypto';
import { TOCManipulator } from '../manipulators/toc';
import { DocxParser } from '../parser';
import { bodySectPr, insertIntoSection, lastSection, ownerDocumentOf } from '../parser/sections';
import { resolveReproducibleMode } from '../zip-hardening';
import { appendElementsToHeaderFooter, resolveHeaderFooterTarget } from './headerFooter';
import { toError } from '../errors';

function isFieldReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

function fieldHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

function collectFieldLiveIds(ast: DocumentAST): Set<string> {
  const live = new Set<string>();
  try {
    for (const sec of ast.sections ?? []) {
      for (const n of sec.nodes ?? []) {
        if (typeof n.id === 'string' && n.id) live.add(n.id);
      }
    }
  } catch { /* best-effort */ }
  return live;
}

export interface FieldOperation extends BaseOperation {
  action: 'field';
  operation: 'insertTOC' | 'insertTOF' | 'insertPageNumber' | 'insertWordCount' | 'updateFields' | 'removeField';
  anchor?: string; // Required for inserts
  position?: 'before' | 'after';
  title?: string;
  levels?: number;
  // Destination store: 'body' (default) or a repeating 'header'/'footer'
  // part. A page-number field belongs in the footer so it repeats per page.
  // Accepts flat or params-nested (mirrors the pick pattern below).
  target?: 'body' | 'header' | 'footer';
  pageType?: 'default' | 'first' | 'even';
  sectionIndex?: number;
  params?: Record<string, unknown>;
}

export async function executeField(op: FieldOperation, ast: DocumentAST, parser?: DocxParser): Promise<void> {
  // OP-16: accept args nested under params:{} or flat (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedOperation = String(pick('operation', op.operation) ?? op.operation ?? '');
  const resolvedAnchor = pick('anchor', op.anchor) as string | undefined;
  const resolvedPosition = (pick('position', op.position) ?? op.position ?? 'after') as 'before' | 'after';
  const resolvedTitle = pick('title', op.title) as string | undefined;
  const resolvedLevels = Number(pick('levels', op.levels) ?? op.levels ?? 3);
  // Destination store resolves before any mutation: bad targets fail closed
  // here so a typo can never misplace content.
  const hfTarget = resolveHeaderFooterTarget(
    `field.${resolvedOperation}`,
    pick('target', op.target),
    pick('pageType', op.pageType),
    pick('sectionIndex', op.sectionIndex)
  );
  if (resolvedOperation === 'updateFields') {
    TOCManipulator.enableWordAutoUpdateFields(parser);
    return;
  }

  if (resolvedOperation.startsWith('insert')) {
    // With no anchor the field goes to the end of the document: the last section.
    let targetSection = lastSection(ast.sections);
    let insertIndex = targetSection ? targetSection.nodes.length : 0;
    let anchorNode: Node | undefined = undefined;

    if (resolvedAnchor) {
      for (const section of ast.sections) {
        const idx = section.nodes.findIndex(n => n.id === resolvedAnchor);
        if (idx !== -1) {
          targetSection = section;
          insertIndex = resolvedPosition === 'before' ? idx : idx + 1;
          anchorNode = section.nodes[idx];
          break;
        }
      }
    }

    if (!targetSection) throw new Error("No section found for field insertion");
    const doc = targetSection.nodes[0]?.xml?.ownerDocument || ownerDocumentOf(ast.sections);
    if (!doc) throw new Error("Cannot resolve owner XML document");

    if (resolvedOperation === 'insertTOC') {
      if (hfTarget.target !== 'body') {
        throw new Error(
          `E_VALIDATION: field.insertTOC target '${hfTarget.target}' is not supported (a table of contents is a multi-paragraph body outline). ` +
            `Remedy: omit target (or pass target:'body') to place the TOC in the body; use target:'footer' only for field.insertPageNumber/field.insertWordCount.`
        );
      }
      // Routed through the TOC deterministic path: no wall-clock mint here.
      // generateTOCNode mints a content-seeded id by default and wall-clock
      // only when opted out, so identical input stays byte-identical.
      const fieldId = undefined as unknown as string | undefined;
      // G2-7: prefer the renderer-compatible map when cheap (parser summary
      // nodePages is cached — O(1) hit unless nodes mutated). Falls back to
      // the heuristic estimate inside scanHeadings, which labels entries.
      let rendererMap: Record<string, number> | undefined;
      try {
        const maybeSummary = parser?.getStructuralSummary?.();
        if (maybeSummary?.nodePages && typeof maybeSummary.nodePages === 'object') {
          rendererMap = maybeSummary.nodePages;
        }
      } catch { rendererMap = undefined; }
      const entries = TOCManipulator.scanHeadings(ast, resolvedLevels || 3, rendererMap);
      const tocNodes = await TOCManipulator.generateTOCNode(doc, resolvedTitle || 'Table of Contents', entries, fieldId, resolvedLevels || 3);

      // Insert into live DOM preserving sibling order ([container, ...entries])
      if (anchorNode && anchorNode.xml && anchorNode.xml.parentNode) {
        const parent = anchorNode.xml.parentNode;
        if (resolvedPosition === 'before') {
          for (const n of tocNodes) {
            parent.insertBefore(n.xml, anchorNode.xml);
          }
        } else {
          const ref = anchorNode.xml.nextSibling;
          for (const n of tocNodes) {
            parent.insertBefore(n.xml, ref);
          }
        }
      } else {
        const body = doc.getElementsByTagName('w:body')[0];
        if (body) {
          // The body's own sectPr: a section break paragraph carries one too, and it is not a child of the body.
          const sectPr = bodySectPr(body);
          for (const n of tocNodes) {
            body.insertBefore(n.xml, sectPr);
          }
        }
      }

      if (parser) {
        for (const n of tocNodes) {
          parser.nodeMap?.set(n.id, n.xml);
        }
      }

      insertIntoSection(ast.sections, targetSection, insertIndex, tocNodes).version += 1;

      // Enable updateFields in Word settings so Word recalculates exact page numbers on open
      TOCManipulator.enableWordAutoUpdateFields(parser);
      return;
    }

    if (resolvedOperation === 'insertPageNumber' || resolvedOperation === 'insertWordCount') {
      const instr = resolvedOperation === 'insertPageNumber' ? 'PAGE' : 'NUMWORDS';
      const fieldId = isFieldReproducible()
        ? (() => {
            const base = `field_${fieldHash8(`${resolvedOperation}:${resolvedAnchor ?? ''}:${hfTarget.target}:${resolvedPosition}`)}`;
            const live = collectFieldLiveIds(ast);
            let candidate = base;
            let n = 2;
            while (live.has(candidate)) candidate = `${base}_${n++}`;
            return candidate;
          })()
        : `field_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;

      // Header/footer routing: the field paragraph is appended to the part
      // (repeating per page) instead of the body (rendering once). Uses the
      // house complex-field shape (fldChar begin/instrText/end) so renderers
      // evaluate it like manipulator-built footers.
      if (hfTarget.target === 'header' || hfTarget.target === 'footer') {
        const ownerDoc = targetSection.nodes[0]?.xml?.ownerDocument
          || ownerDocumentOf(ast.sections);
        if (!ownerDoc) throw new Error("Cannot resolve owner XML document");
        const partPara = buildHeaderFooterFieldParagraph(
          ownerDoc,
          hfTarget.target,
          resolvedOperation === 'insertPageNumber' ? 'pageNumber' : 'wordCount'
        );
        // Batch-time preflight: header/footer parts persist verbatim (the
        // save path strips session ids from document.xml only), so a
        // session-only attribute here would pass the batch and die at the
        // validator-gated save. Fail fast with the true error instead.
        assertHeaderFooterParagraphClean(partPara, resolvedOperation, hfTarget.target);
        appendElementsToHeaderFooter(ast, parser, [partPara], {
          targetType: hfTarget.target,
          pageType: hfTarget.pageType,
          sectionIndex: hfTarget.sectionIndex
        });
        return;
      }

      const p = createWordElement(doc, 'w:p');
      p.setAttribute('data-astra-id', fieldId);

      // Valid hierarchy: w:p -> w:fldSimple -> w:r -> w:t
      const fldSimple = createWordElement(doc, 'w:fldSimple');
      fldSimple.setAttribute('w:instr', instr);
      const r = createWordElement(doc, 'w:r');
      const t = createWordElement(doc, 'w:t');
      t.textContent = instr;
      r.appendChild(t);
      fldSimple.appendChild(r);
      p.appendChild(fldSimple);

      const newNode: Node = {
        id: fieldId,
        type: 'paragraph',
        xml: p,
        metadata: {
          text: resolvedOperation === 'insertPageNumber' ? '[Page Number]' : '[Word Count]',
          hasCitation: false,
          hasComment: false,
          hasBookmark: false,
          hasTrackedChanges: false,
          hasEquation: false,
          hasField: true,
          hasImage: false,
          hasHyperlink: false
        },
        version: 1,
        checksum: await calculateNodeChecksum(new XMLSerializer().serializeToString(p))
      };

      insertIntoSection(ast.sections, targetSection, insertIndex, [newNode]).version += 1;

      // Insert into live DOM adjacent to the anchor node (mirrors insertTOC pattern above)
      if (anchorNode && anchorNode.xml && anchorNode.xml.parentNode) {
        if (resolvedPosition === 'before') {
          anchorNode.xml.parentNode.insertBefore(p, anchorNode.xml);
        } else {
          anchorNode.xml.parentNode.insertBefore(p, anchorNode.xml.nextSibling);
        }
      } else {
        const body = doc.getElementsByTagName('w:body')[0];
        if (body) {
          body.insertBefore(p, bodySectPr(body));
        }
      }

      if (parser) {
        parser.nodeMap?.set(fieldId, p);
      }
      return;
    }
  }

  throw new Error(`field.${resolvedOperation} is not supported yet`);
}

/**
 * Batch-time preflight for header/footer part writes: scans the built
 * paragraph subtree for session-only tracking attributes the document
 * validator rejects. Throws E_VALIDATION naming the field op and the part
 * so the batch fails fast instead of the save.
 */
function assertHeaderFooterParagraphClean(
  root: Element,
  operation: string,
  targetType: 'header' | 'footer'
): void {
  const stack: Element[] = [root];
  while (stack.length > 0) {
    const el = stack.pop() as Element;
    if (typeof el.hasAttribute === 'function' && el.hasAttribute('data-astra-id')) {
      throw toError(
        'E_VALIDATION',
        `field.${operation} target '${targetType}' carries a session-only tracking attribute the validator rejects in the ${targetType} part. ` +
        `Remedy: rebuild the ${targetType} content without session ids (re-run the field.${operation} op), then retry save_document.`
      );
    }
    const kids = el.childNodes;
    for (let i = 0; i < kids.length; i++) {
      if (kids[i].nodeType === 1) stack.push(kids[i] as Element);
    }
  }
}

/**
 * Builds a centered (footer) / right-aligned (header) paragraph carrying a
 * page-number ("Page X of Y") or word-count ("Words: N") complex field. The
 * caller appends it to the header/footer part where it repeats per page.
 * Carries no session id: header/footer parts persist verbatim (the save path
 * strips session ids from document.xml only), so any non-OOXML attribute
 * here fails schema validation at save.
 */
function buildHeaderFooterFieldParagraph(
  doc: Document,
  targetType: 'header' | 'footer',
  kind: 'pageNumber' | 'wordCount'
): Element {
  const p = createWordElement(doc, 'w:p');
  const pPr = createWordElement(doc, 'w:pPr');
  const jc = createWordElement(doc, 'w:jc');
  jc.setAttribute('w:val', targetType === 'header' ? 'right' : 'center');
  pPr.appendChild(jc);
  p.appendChild(pPr);

  const addTextRun = (text: string): void => {
    const r = createWordElement(doc, 'w:r');
    r.appendChild(createTextElement(doc, text));
    p.appendChild(r);
  };
  const addField = (instr: string): void => {
    const begin = createWordElement(doc, 'w:r');
    const beginChar = createWordElement(doc, 'w:fldChar');
    beginChar.setAttribute('w:fldCharType', 'begin');
    begin.appendChild(beginChar);
    p.appendChild(begin);
    const instrRun = createWordElement(doc, 'w:r');
    const instrText = createWordElement(doc, 'w:instrText');
    instrText.setAttribute('xml:space', 'preserve');
    instrText.textContent = ` ${instr} `;
    instrRun.appendChild(instrText);
    p.appendChild(instrRun);
    const end = createWordElement(doc, 'w:r');
    const endChar = createWordElement(doc, 'w:fldChar');
    endChar.setAttribute('w:fldCharType', 'end');
    end.appendChild(endChar);
    p.appendChild(end);
  };

  if (kind === 'pageNumber') {
    addTextRun('Page ');
    addField('PAGE');
    addTextRun(' of ');
    addField('NUMPAGES');
  } else {
    addTextRun('Words: ');
    addField('NUMWORDS');
  }
  return p;
}

