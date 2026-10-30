import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Section, Node } from '../ast';
import { BaseOperation } from './base';
import { DocxParser, upsertSectPrChild } from '../parser';
import * as crypto from 'crypto';
import { resolveReproducibleMode } from '../zip-hardening';
import { sectionPosition } from './section-number';
import { toError } from '../errors';

function isSectionReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

function sectionHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

export interface SectionOperation extends BaseOperation {
  action: 'section';
  operation: 'setColumns' | 'split' | 'merge';
  target?: string;
  sectionIndex?: number;
  columns?: number;
  spacing?: number; // DXA gutter spacing between columns (default 720 = 0.5 in)
  separator?: boolean; // Toggles vertical dividing rule between columns (w:sep="1")
  params?: Record<string, unknown>;
}

export function executeSection(op: SectionOperation, ast: DocumentAST, parser?: DocxParser): void {
  // OP-16: accept args nested under params:{} or flat (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedOperation = String(pick('operation', op.operation) ?? op.operation ?? '');
  const resolvedTarget = pick('target', op.target) as string | undefined;
  // Callers number sections from 1; the AST counts from 0.
  const resolvedSectionIndex = sectionPosition(`section.${resolvedOperation}`, pick('sectionIndex', op.sectionIndex), ast.sections.length);
  const resolvedColumns = Number(pick('columns', op.columns) ?? op.columns ?? 2);
  const resolvedSpacing = Number(pick('spacing', op.spacing) ?? op.spacing ?? 720);
  const resolvedSeparator = pick('separator', op.separator) as boolean | undefined;
  if (resolvedOperation === 'setColumns') {
    const colsCount = resolvedColumns || 2;
    const spacing = resolvedSpacing || 720;

    const targetSection = (resolvedSectionIndex !== undefined && ast.sections[resolvedSectionIndex])
      ? ast.sections[resolvedSectionIndex]
      : (resolvedTarget ? ast.sections.find(s => s.id === resolvedTarget) : ast.sections[0]);

    if (!targetSection) throw new Error(`Target section not found`);

    // Modify sectPr XML if available
    const lastNode = targetSection.nodes[targetSection.nodes.length - 1];
    const doc = lastNode?.xml?.ownerDocument || ast.sections[0]?.nodes[0]?.xml?.ownerDocument;
    
    if (targetSection.pageSetup) {
      targetSection.pageSetup.columns = colsCount;
    } else {
      targetSection.pageSetup = { columns: colsCount };
    }

    if (doc) {
      // OP-12: scope strictly to THIS section's sectPr — never sectPr[0].
      // Mirrors the HDRFTR per-section resolution (parser index, then local
      // search, then single-sectPr fallback).
      const sectionIdx = resolvedSectionIndex !== undefined
        ? resolvedSectionIndex
        : (resolvedTarget ? ast.sections.findIndex(s => s.id === resolvedTarget) : 0);
      const safeIdx = sectionIdx >= 0 ? sectionIdx : 0;
      let sectPr: Element | null = null;
      try {
        const sectPrs = parser?.getSectPrElements() ?? [];
        if (sectPrs.length > safeIdx) sectPr = sectPrs[safeIdx];
      } catch { sectPr = null; }
      if (!sectPr) {
        for (const node of targetSection.nodes) {
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
        else if (all.length > safeIdx) sectPr = all[safeIdx] as Element;
      }
      if (!sectPr && parser) {
        throw toError(
          'E_VALIDATION',
          `Cannot set columns: the document has no section properties (w:sectPr) for section ${safeIdx + 1}.`,
          'Re-read the document with inspect_document and check the section count; if it looks wrong, save and reopen the file, then retry.',
        );
      }
      if (sectPr) {
        const cols = upsertSectPrChild(sectPr, 'w:cols');
        cols.setAttribute('w:num', String(colsCount));
        cols.setAttribute('w:space', String(spacing));
        cols.setAttribute('w:equalWidth', '1');
        if (resolvedSeparator !== undefined) {
          cols.setAttribute('w:sep', resolvedSeparator ? '1' : '0');
        }
      }
    }

    targetSection.version += 1;
    return;
  }

  if (resolvedOperation === 'split') {
    if (!resolvedTarget) throw new Error('section.split requires a target node id');
    let foundSecIdx = -1;
    let foundNodeIdx = -1;

    for (let sIdx = 0; sIdx < ast.sections.length; sIdx++) {
      const nIdx = ast.sections[sIdx].nodes.findIndex(n => n.id === resolvedTarget);
      if (nIdx !== -1) {
        foundSecIdx = sIdx;
        foundNodeIdx = nIdx;
        break;
      }
    }

    if (foundSecIdx === -1) throw new Error(`Node ${resolvedTarget} not found for section split`);

    const origSection = ast.sections[foundSecIdx];
    const newNodes = origSection.nodes.splice(foundNodeIdx);
    
    const newSection: Section = isSectionReproducible()
      ? (() => {
          const base = `sec_${sectionHash8(`${resolvedTarget ?? ''}:${foundSecIdx}:${foundNodeIdx}:${newNodes.length}`)}`;
          const live = new Set<string>();
          try {
            for (const s of ast.sections ?? []) {
              if (typeof s.id === 'string' && s.id) live.add(s.id);
            }
          } catch { /* best-effort */ }
          let candidate = base;
          let n = 2;
          while (live.has(candidate)) candidate = `${base}_${n++}`;
          return {
            id: candidate,
            heading: null,
            nodes: newNodes,
            version: 1,
            checksum: `sec_${sectionHash8(`${candidate}:${newNodes.length}`)}`,
            pageSetup: { ...(origSection.pageSetup || {}) }
          };
        })()
      : {
      id: `sec_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      heading: null,
      nodes: newNodes,
      version: 1,
      checksum: `sec_${Date.now()}`,
      pageSetup: { ...(origSection.pageSetup || {}) }
    };

    ast.sections.splice(foundSecIdx + 1, 0, newSection);
    origSection.version += 1;
    // OP-11: mirror the AST split in the live DOM. Intermediate sections are
    // delimited by a w:sectPr inside the preceding paragraph's w:pPr; without
    // this the split is AST-only and lost on save.
    injectSectionBreakOnSplit(origSection, newSection, parser);
    return;
  }

  if (resolvedOperation === 'merge') {
    const secIdx = resolvedSectionIndex !== undefined ? resolvedSectionIndex : ast.sections.findIndex(s => s.id === resolvedTarget);
    if (secIdx === -1 || secIdx >= ast.sections.length - 1) {
      throw new Error(`Invalid section index/target for merge: merge joins a section with the one after it (document has ${ast.sections.length} section${ast.sections.length === 1 ? '' : 's'}); pass sectionIndex in [1, ${Math.max(ast.sections.length - 1, 1)}] or a target section id`);
    }

    const firstSec = ast.sections[secIdx];
    const secondSec = ast.sections[secIdx + 1];
    // OP-11: remove the boundary w:sectPr from the live DOM *before* merging
    // so the combined section inherits the following section's properties
    // (or the body-level sectPr when merging into the final section).
    removeBoundarySectPr(firstSec);
    firstSec.nodes.push(...secondSec.nodes);
    ast.sections.splice(secIdx + 1, 1);
    firstSec.version += 1;
    return;
  }

  throw new Error(`section.${resolvedOperation} is not supported yet`);
}

/**
 * OP-11: inject a w:sectPr for the end of `origSection` into the live DOM.
 * Uses the ownerDocument of the live nodes (same document the parser
 * serializes), inserting a w:pPr > w:sectPr. Clones the document's final
 * (body-level) sectPr so page size/margins/cols survive; creates a fresh
 * sectPr when none exists.
 */
function injectSectionBreakOnSplit(origSection: Section, newSection: Section, parser?: DocxParser): void {
  const prevNode = origSection.nodes.length > 0
    ? origSection.nodes[origSection.nodes.length - 1]
    : undefined;
  const firstNewEl = newSection.nodes.length > 0 ? newSection.nodes[0].xml : undefined;
  const liveDoc: Document | null =
    (prevNode?.xml?.ownerDocument as Document | undefined) ||
    (firstNewEl?.ownerDocument as Document | undefined) ||
    parser?.getDocumentDom() ||
    null;
  if (!liveDoc) return;

  const sourceSectPr = findSourceSectPr(liveDoc, parser);

  const buildSectPr = (doc: Document): Element => {
    let created: Element;
    if (sourceSectPr) {
      created = sourceSectPr.cloneNode(true) as Element;
      // cloneNode keeps the source's ownerDocument; adopt into the live doc
      // when they differ so serialization stays in one DOM.
      if (created.ownerDocument !== doc && typeof doc.importNode === 'function') {
        created = doc.importNode(created, true) as Element;
      }
    } else {
      created = createWordElement(doc, 'w:sectPr');
    }
    return created;
  };

  // Normal case: the last paragraph of the original section carries the break.
  const prevEl = prevNode?.xml as Element | undefined;
  if (prevEl && prevEl.tagName === 'w:p') {
    const doc = (prevEl.ownerDocument as Document) || liveDoc;
    let pPr: Element | null = null;
    for (let i = 0; i < prevEl.childNodes.length; i++) {
      const child = prevEl.childNodes[i];
      if (child.nodeType === 1 && (child as Element).tagName === 'w:pPr') {
        pPr = child as Element;
        break;
      }
    }
    if (!pPr) {
      pPr = createWordElement(doc, 'w:pPr');
      // w:pPr must be the first child of w:p per schema.
      if (prevEl.firstChild) prevEl.insertBefore(pPr, prevEl.firstChild);
      else prevEl.appendChild(pPr);
    }
    // Abort when a break already exists (idempotent retry safety).
    for (let i = 0; i < pPr.childNodes.length; i++) {
      const child = pPr.childNodes[i];
      if (child.nodeType === 1 && (child as Element).tagName === 'w:sectPr') return;
    }
    // w:sectPr is conventionally the last child of w:pPr.
    pPr.appendChild(buildSectPr(doc));
    return;
  }

  // Edge cases: split at index 0 (origSection empty) or preceding node is a
  // table (sectPr cannot live in w:tbl). Insert a synthetic break paragraph
  // before the new section's first node so the DOM gains a real delimiter.
  // No parser nodeMap update is needed: merges/splits regroup AST references
  // in place and never remove these paragraph elements, so stable ids stay
  // valid; the synthetic break is tracked via the AST node below.
  const doc = (firstNewEl?.ownerDocument as Document) || liveDoc;
  const body = doc.getElementsByTagName('w:body')[0];
  if (!body) return;
  const p = createWordElement(doc, 'w:p');
  const breakId = isSectionReproducible()
    ? `p_secbreak_${sectionHash8(`${newSection.id}:break`)}`
    : `p_secbreak_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  p.setAttribute('data-astra-id', breakId);
  const pPr = createWordElement(doc, 'w:pPr');
  p.appendChild(pPr);
  pPr.appendChild(buildSectPr(doc));
  const anchor = firstNewEl?.parentNode === body || firstNewEl?.parentNode
    ? firstNewEl as Element
    : undefined;
  if (anchor && anchor.parentNode) {
    anchor.parentNode.insertBefore(p, anchor);
  } else {
    // Keep the body-level sectPr last: insert before it when present.
    let bodySectPr: Element | null = null;
    for (let i = 0; i < body.childNodes.length; i++) {
      const child = body.childNodes[i];
      if (child.nodeType === 1 && (child as Element).tagName === 'w:sectPr') {
        bodySectPr = child as Element;
        break;
      }
    }
    if (bodySectPr) body.insertBefore(p, bodySectPr);
    else body.appendChild(p);
  }
  const breakNode: Node = {
    id: breakId,
    type: 'paragraph',
    xml: p,
    metadata: {
      text: '',
      hasCitation: false,
      hasComment: false,
      hasBookmark: false,
      hasTrackedChanges: false,
      hasEquation: false,
      hasField: false,
      hasImage: false,
      hasHyperlink: false
    },
    version: 1,
    checksum: isSectionReproducible() ? `secbreak_${sectionHash8(breakId)}` : `secbreak_${Date.now()}`
  };
  origSection.nodes.push(breakNode);
}

/**
 * OP-11: remove the intermediate w:sectPr terminating `firstSec` from the
 * live DOM (the w:sectPr inside its last paragraph's w:pPr). Body-level
 * sectPr elements are never removed here — the merged section keeps the
 * surviving (following/final) section properties.
 */
function removeBoundarySectPr(firstSec: Section): void {
  if (firstSec.nodes.length === 0) return;
  const lastEl = firstSec.nodes[firstSec.nodes.length - 1]?.xml as Element | undefined;
  if (!lastEl || typeof lastEl.getElementsByTagName !== 'function') return;
  const candidates = Array.from(lastEl.getElementsByTagName('w:sectPr'));
  for (const sp of candidates) {
    const parent = sp.parentNode as Element | null;
    if (parent && (parent as Element).tagName === 'w:pPr') {
      parent.removeChild(sp);
      return;
    }
  }
}

/** Last sectPr in document order (the body-level one) used as clone source. */
function findSourceSectPr(doc: Document, parser?: DocxParser): Element | null {
  try {
    const fromParser = parser?.getSectPrElements();
    if (fromParser && fromParser.length > 0) return fromParser[fromParser.length - 1];
  } catch { /* fall through to DOM lookup */ }
  const all = Array.from(doc.getElementsByTagName('w:sectPr'));
  return all.length > 0 ? all[all.length - 1] : null;
}

