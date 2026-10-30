import { createWordElement } from '../ooxml/ns';
import { toError } from '../errors';
import { DocumentAST, Node } from '../ast';
import { BaseOperation } from './base';
import { NumberingManipulator } from '../manipulators/numbering';

export interface ListOperation extends BaseOperation {
  action: 'list';
  operation: 'create' | 'continue' | 'changeLevel' | 'remove';
  target: string | string[]; // Paragraph nodes to convert to/from list items
  listType?: 'bullet' | 'numbered'; // For create
  /** List level, counting from 1 (1 is the top level, up to 9). changeLevel also takes 0: out of the list. */
  level?: number;
  delta?: number; // For changeLevel: relative level shift (cur + delta)
  /** Node id of an existing list item to continue from (for continue; nested params.previousNodeId also honored). */
  previousNodeId?: string;
  params?: Record<string, unknown>;
}

const MAX_LIST_LEVEL = 9;

/**
 * Callers number list levels from 1 (1 is the top level), like headings; on
 * changeLevel, 0 takes the paragraph out of the list. w:ilvl counts from 0, so
 * the number a caller types is checked and converted once, before anything is
 * mutated.
 */
function listLevelNumber(operation: string, raw: unknown): number {
  const n = typeof raw === 'number' ? raw : Number(String(raw));
  const min = operation === 'changeLevel' ? 0 : 1;
  if (!Number.isInteger(n) || n < min || n > MAX_LIST_LEVEL) {
    throw toError(
      'E_VALIDATION',
      `list.${operation} level ${String(raw)} is out of range`,
      operation === 'changeLevel'
        ? `pass level in [1, ${MAX_LIST_LEVEL}], counting from 1 (1 is the top level), or 0 to take the paragraph out of the list`
        : `pass level in [1, ${MAX_LIST_LEVEL}], counting from 1 (1 is the top level)`
    );
  }
  return n;
}

/** Level number of a paragraph: 1 is the top level, 0 when it is not in a list. */
function currentLevelNumber(node: Node): number {
  let inList = node.metadata.listLevel !== undefined;
  if (!inList) {
    try { inList = node.xml.getElementsByTagName('w:numPr').length > 0; } catch { /* not a list item */ }
  }
  return inList ? getCurrentListLevel(node) + 1 : 0;
}

/** Clamp OOXML list level to 0-8 (ECMA-376 w:ilvl range). */
export function clampListLevel(v: unknown): number {
  const n = typeof v === 'number' ? v : parseInt(String(v ?? '0'), 10);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(8, Math.floor(n)));
}

/** Read current list level from metadata first, then w:ilvl in XML. */
export function getCurrentListLevel(node: Node): number {
  try {
    const m = node.metadata.listLevel;
    if (typeof m === 'number' && Number.isFinite(m)) return clampListLevel(m);
    const numPr = node.xml.getElementsByTagName('w:numPr')[0];
    const ilvl = numPr?.getElementsByTagName('w:ilvl')[0];
    const v = ilvl?.getAttribute('w:val');
    if (v !== undefined && v !== null && v !== '') {
      const n = parseInt(String(v), 10);
      if (Number.isFinite(n)) return clampListLevel(n);
    }
  } catch { /* ignore */ }
  return 0;
}

/** Resolve the bullet/numbered style for a concrete numId through the
 * numbering maps (anything but 'bullet' counts as numbered, mirroring the
 * exporters). Unresolvable numbering reads as bullet, never throws. */
export function resolveListStyle(ast: DocumentAST, numId: string): 'bullet' | 'numbered' {
  try {
    const concrete = (ast.numbering?.concreteNum ?? {}) as Record<string, Element>;
    const abstract = (ast.numbering?.abstractNum ?? {}) as Record<string, Element>;
    const numEl = concrete[numId];
    const absId = numEl?.getElementsByTagName?.('w:abstractNumId')?.[0]?.getAttribute?.('w:val');
    const absEl = (absId != null && absId !== '') ? abstract[absId] : undefined;
    if (!absEl) return 'bullet';
    const fmts = Array.from(absEl.getElementsByTagName?.('w:numFmt') || []) as Element[];
    for (const f of fmts) {
      const v = f.getAttribute?.('w:val');
      if (v === 'bullet') return 'bullet';
      if (v) return 'numbered';
    }
    return 'bullet';
  } catch {
    return 'bullet';
  }
}

/** Read the ilvl of a list node (0 when absent/unparseable). */
function readNodeListLevel(xml: Element): number {
  try {
    const numPr = xml.getElementsByTagName('w:numPr')[0] as Element | undefined;
    const v = numPr?.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val');
    if (v !== undefined && v !== null && v !== '') {
      const n = parseInt(String(v), 10);
      if (Number.isFinite(n)) return clampListLevel(n);
    }
  } catch { /* fall through */ }
  return 0;
}

export function executeList(op: ListOperation, ast: DocumentAST): void {
  // OP-16: accept args nested under params:{} or flat (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedOperation = String(pick('operation', op.operation) ?? op.operation ?? '');
  const resolvedTarget = (pick('target', op.target) ?? op.target) as string | string[] | undefined;
  const resolvedListType = (pick('listType', op.listType) ?? op.listType ?? 'bullet') as 'bullet' | 'numbered';
  const rawLevel = pick('level', op.level) as number | undefined;
  const rawDelta = pick('delta', op.delta) as number | undefined;
  const numManipulator = new NumberingManipulator(ast);
  const targets = Array.isArray(resolvedTarget) ? resolvedTarget : (resolvedTarget ? [resolvedTarget] : []);
  
  if (targets.length === 0) throw new Error("List operation requires target(s)");

  if (resolvedOperation !== 'create' && resolvedOperation !== 'continue' &&
      resolvedOperation !== 'changeLevel' && resolvedOperation !== 'remove') {
    throw new Error(
      `Unknown list operation '${resolvedOperation}'. Expected 'create', 'continue', 'changeLevel', or 'remove'. ` +
        `Re-read the document with inspect_document and retry with a valid list operation.`
    );
  }
  const typedLevel = resolvedOperation !== 'remove' && rawLevel !== undefined && rawLevel !== null
    ? listLevelNumber(resolvedOperation, rawLevel)
    : undefined;

  let numId = '1';
  let joinedListStyle: 'bullet' | 'numbered' | undefined;
  let joinedLevel: number | undefined;
  if (resolvedOperation === 'create') {
    numId = numManipulator.getOrCreateList(resolvedListType || 'bullet');
    joinedListStyle = resolvedListType || 'bullet';
  } else if (resolvedOperation === 'continue') {
    // Join the caller's list instead of forking a new sequence: an explicit
    // previousNodeId wins, else the nearest preceding sibling carrying
    // w:numPr. No previous list fails closed (never a silent new list).
    const explicitPrev = (pick('previousNodeId', op.previousNodeId) ?? op.previousNodeId) as string | undefined;
    let prevId = explicitPrev;
    if (!prevId) {
      const firstTarget = targets[0];
      for (const section of ast.sections) {
        const idx = section.nodes.findIndex(n => n.id === firstTarget);
        if (idx !== -1) {
          for (let i = idx - 1; i >= 0; i--) {
            const cand = section.nodes[i];
            try {
              if (cand.xml.getElementsByTagName('w:numPr').length > 0) { prevId = cand.id; break; }
            } catch { /* keep scanning */ }
          }
          break;
        }
      }
    }
    if (!prevId) {
      throw new Error(
        `List 'continue' found no previous list item before '${targets[0]}'. ` +
          `Re-read the document with inspect_document and retry with { operation:'continue', target:'<newParaId>', previousNodeId:'<listItemId>' }.`
      );
    }
    numId = numManipulator.continueList(prevId);
    joinedListStyle = resolveListStyle(ast, numId);
    for (const section of ast.sections) {
      const prev = section.nodes.find(n => n.id === prevId);
      if (prev) { joinedLevel = readNodeListLevel(prev.xml); break; }
    }
  }
  // A list reused from an older file may stop short of the level an item is
  // placed at: give its definition the levels it lacks.
  if (resolvedOperation === 'create' || resolvedOperation === 'continue') numManipulator.ensureLevels(numId);

  for (const target of targets) {
    let targetNode: Node | undefined;
    for (const section of ast.sections) {
      targetNode = section.nodes.find(n => n.id === target);
      if (targetNode) break;
    }

    if (!targetNode) {
      throw new Error(`Node ${target} not found for list operation '${resolvedOperation}'. Re-read the document with inspect_document and retry with valid node ids.`);
    }

    if ((targetNode.xml as Element).tagName !== 'w:p') {
      throw toError(
        'E_VALIDATION',
        `list.${resolvedOperation} applies to paragraphs, but '${target}' is a ${targetNode.type} node.`,
        "Target a paragraph id (see inspect_document). To change a table's contents use action 'table'."
      );
    }

    const doc = targetNode.xml.ownerDocument;
    if (!doc) throw new Error(`Node ${target} has no owner document; cannot apply list operation`);

    let pPr = targetNode.xml.getElementsByTagName('w:pPr')[0];
    if (!pPr) {
      pPr = createWordElement(doc, 'w:pPr');
      targetNode.xml.insertBefore(pPr, targetNode.xml.firstChild);
    }

    if (resolvedOperation === 'remove') {
      const numPr = pPr.getElementsByTagName('w:numPr')[0];
      if (numPr) pPr.removeChild(numPr);
      targetNode.metadata.listLevel = undefined;
      targetNode.metadata.listStyle = undefined;
    } else if (resolvedOperation === 'changeLevel') {
      // Resolve the level number (1 = top level): explicit level wins;
      // otherwise cur + delta, clamped to 0-9; else cur. Level 0 = outdented
      // to plain: remove numPr (mirrors Tab outdent at the top level in the
      // canvas, which dispatches `remove`). A paragraph outside any list is 0.
      let level: number;
      if (typedLevel !== undefined) {
        level = typedLevel;
      } else if (rawDelta !== undefined && Number.isFinite(Number(rawDelta))) {
        level = Math.max(0, Math.min(MAX_LIST_LEVEL, currentLevelNumber(targetNode) + Math.trunc(Number(rawDelta))));
      } else {
        level = currentLevelNumber(targetNode);
      }
      if (level === 0) {
        const numPr = pPr.getElementsByTagName('w:numPr')[0];
        if (numPr) pPr.removeChild(numPr);
        targetNode.metadata.listLevel = undefined;
        targetNode.metadata.listStyle = undefined;
      } else {
        let numPr = pPr.getElementsByTagName('w:numPr')[0];
        if (!numPr) {
          numPr = createWordElement(doc, 'w:numPr');
          pPr.appendChild(numPr);
        }
        let ilvlNode = numPr.getElementsByTagName('w:ilvl')[0];
        if (!ilvlNode) {
          ilvlNode = createWordElement(doc, 'w:ilvl');
        }
        ilvlNode.setAttribute('w:val', (level - 1).toString());
        // Set numId: reuse the paragraph's own list; a paragraph that is not in
        // one yet joins a real list of the requested type (never a dangling id).
        let numIdNode = numPr.getElementsByTagName('w:numId')[0];
        if (!numIdNode) {
          numIdNode = createWordElement(doc, 'w:numId');
          const existing = pPr.getElementsByTagName('w:numId')[0];
          const ownId = existing?.getAttribute('w:val');
          if (!ownId) targetNode.metadata.listStyle = resolvedListType;
          numIdNode.setAttribute('w:val', ownId || numManipulator.getOrCreateList(resolvedListType));
        }
        numManipulator.ensureLevels(numIdNode.getAttribute('w:val') ?? '');
        if (ilvlNode.parentNode === numPr) numPr.removeChild(ilvlNode);
        if (numIdNode.parentNode === numPr) numPr.removeChild(numIdNode);
        numPr.appendChild(ilvlNode);
        numPr.appendChild(numIdNode);
        targetNode.metadata.listLevel = level - 1;
      }
    } else {
      // create | continue (any other operation threw before the loop).
      let numPr = pPr.getElementsByTagName('w:numPr')[0];
      if (!numPr) {
        numPr = createWordElement(doc, 'w:numPr');
        pPr.appendChild(numPr);
      }

      // Per CT_NumPr (ECMA-376), child order is w:ilvl then w:numId.
      let ilvlNode = numPr.getElementsByTagName('w:ilvl')[0];
      if (!ilvlNode) {
        ilvlNode = createWordElement(doc, 'w:ilvl');
      }
      const level = typedLevel !== undefined
        ? typedLevel - 1
        : (resolvedOperation === 'continue' ? (joinedLevel ?? 0) : (targetNode.metadata.listLevel ?? 0));
      ilvlNode.setAttribute('w:val', level.toString());

      // Set numId
      let numIdNode = numPr.getElementsByTagName('w:numId')[0];
      if (!numIdNode) {
        numIdNode = createWordElement(doc, 'w:numId');
      }
      numIdNode.setAttribute('w:val', numId);

      // Enforce schema order: ilvl before numId (fixes inverted append order).
      if (ilvlNode.parentNode === numPr) numPr.removeChild(ilvlNode);
      if (numIdNode.parentNode === numPr) numPr.removeChild(numIdNode);
      numPr.appendChild(ilvlNode);
      numPr.appendChild(numIdNode);
      
      targetNode.metadata.listLevel = level;
      if (joinedListStyle) targetNode.metadata.listStyle = joinedListStyle;
    }
    
    targetNode.version += 1;
  }
}
