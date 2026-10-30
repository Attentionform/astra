import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node } from '../ast';
import { BaseOperation, createTextElement, validateVersion, findNode } from './base';
import { emitTrackedReplace, resolveRevisionStamp } from '../ooxml/track-changes-emitter';
import { RPR_CHILD_ORDER, sortChildrenBySchema } from '../ooxml/schema-order';
import { parseFontSize } from '../operation-schema';

export interface ReplaceOperation extends BaseOperation {
  action: 'replace';
  target: string;
  text: string;
  preserveFormatting?: boolean;
  trackChanges?: boolean;
  author?: string;
  date?: string;
  resetStyle?: boolean;
  fontSize?: number | string;
  sz?: number | string;
  params?: Record<string, unknown>;
}

export interface ReplaceOperationResult {
  /** Advisory warnings (e.g. dropped note references). Ops still apply; never a failure. */
  warnings: string[];
}

export function executeReplace(op: ReplaceOperation, ast: DocumentAST): ReplaceOperationResult {
  // OP-16: accept args nested under params:{} or flat (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedTarget = String(pick('target', op.target) ?? op.target ?? '');
  const resolvedText = String(pick('text', op.text) ?? op.text ?? '');
  // Preserve-or-carry is the default: a replace without flags keeps the
  // target paragraph's run formatting on the rewritten text. An explicit
  // preserveFormatting:false opts out (legacy strip behavior); resetStyle
  // also strips. This matches the tracked path (emitTrackedReplace always
  // carries the first run's properties) and the style op merge rule
  // (explicit properties win, everything else is carried).
  const rawPreserve = pick('preserveFormatting', op.preserveFormatting);
  const resolvedPreserve = rawPreserve === undefined ? true : Boolean(rawPreserve);
  const resolvedReset = pick('resetStyle', op.resetStyle) === true;
  const rawSz = pick('fontSize', op.fontSize) ?? pick('sz', op.sz);
  const resolvedSz = parseFontSize(rawSz as number | string | undefined);
  const targetNode: Node | undefined = findNode(ast, resolvedTarget);

  // Unknown targets must fail loudly: silently retargeting node[0] corrupted
  // unrelated paragraphs whenever callers passed stale ids.

  if (!targetNode) {
    throw new Error(`Node ${resolvedTarget} not found in AST`);
  }

  validateVersion(targetNode, op.expectedVersion);

  if (!targetNode.xml || (targetNode.xml as Element).tagName !== 'w:p') {
    throw Object.assign(
      new Error(`E_VALIDATION: replace operation only supports paragraph nodes ('w:p'), got '${(targetNode.xml as Element)?.tagName || 'unknown'}'. For a table cell, target '<tableId>:r<row>c<col>' (1-based, as inspect prints the grid) or use action: 'table' with operation: 'updateCell'.`),
      { code: 'E_VALIDATION', retryable: false, remedy: "Target a paragraph node id from inspect (e.g. { action: 'replace', target: '<paraId>', text: '...' })." }
    );
  }

  // A picture, chart or equation node is an object, not text: replacing its "text"
  // used to delete the object. Refuse and name the operation that edits it.
  if (targetNode.type === 'image' || targetNode.type === 'chart' || targetNode.type === 'chart3d' || targetNode.type === 'equation') {
    throw Object.assign(
      new Error(`E_VALIDATION: replace changes paragraph text, but '${targetNode.id}' is a ${targetNode.type} node; replacing it would delete it.`),
      {
        code: 'E_VALIDATION',
        retryable: false,
        remedy: "Use the operation for that node: action 'image' (resize, setAltText, setCaption, delete), action 'chart' operation 'update', or action 'equation' operation 'replace'. To remove it, use action 'delete'.",
      }
    );
  }

  // Replace rewrites the runs of a paragraph, but inline math is not a run: it would stay behind, cut off from
  // the sentence around it. No operation edits a substring, so refuse and say what to do instead.
  const mathEl = targetNode.xml as Element;
  if (mathEl.getElementsByTagName('m:oMath').length > 0 || mathEl.getElementsByTagName('m:oMathPara').length > 0) {
    throw Object.assign(
      new Error(`E_VALIDATION: Paragraph '${targetNode.id}' contains math; replace would detach it from its sentence.`),
      {
        code: 'E_VALIDATION',
        retryable: false,
        remedy: "Change an equation with action 'equation' operation 'replace', or delete the paragraph and insert a new one. No operation edits a substring of a paragraph; style with match only formats text.",
      }
    );
  }

  // Runs that hold a picture, chart, shape, embedded object or comment anchor are not text:
  // lift them out before the text is rewritten and put them back after it.
  const keptRuns = detachObjectRuns(targetNode.xml as Element);

  // Tracking applies to every edit including the first one on a fresh
  // (version 1) node: when tracking is on, the original text must survive
  // as w:del/w:delText so reviewers can accept/reject. A version>1 gate
  // here silently overwrote fresh paragraphs, destroying the baseline.
  // Follow the live document setting; an explicit per-op flag wins.
  // When neither carries a value the edit stays silent.
  const explicitTrack = pick('trackChanges', op.trackChanges);
  const docTrack = ast.settings.trackChanges;
  const trackOn = explicitTrack !== undefined ? explicitTrack === true : docTrack === true;

  if (trackOn) {
    const carried = snapshotNoteRefs(targetNode.xml as Element);
    const stamp = resolveRevisionStamp(pick('author', op.author), pick('date', op.date), ast.settings.defaultAuthor);
    const hasRevs = emitTrackedReplace(
      targetNode.xml,
      targetNode.metadata.text,
      resolvedText,
      stamp.author,
      stamp.date,
      { getNextId: () => ast.nextChangeId++ }
    );
    if (hasRevs || targetNode.metadata.text !== resolvedText) {
      targetNode.metadata.hasTrackedChanges = true;
    }
    if (resolvedReset) {
      stripInheritedRunProps(targetNode.xml as Element);
    }
    if (resolvedSz !== undefined && targetNode.xml) {
      applyRunFontSize(targetNode.xml, resolvedSz);
    }
    const warnings = restoreNoteRefs(targetNode.xml as Element, carried);
    warnings.push(...restoreObjectRuns(targetNode.xml as Element, keptRuns));
    // Update AST node metadata
    targetNode.metadata.text = resolvedText;
    targetNode.version += 1;
    return { warnings };
  }
  const out = replaceText(targetNode.xml, resolvedText, resolvedPreserve, resolvedSz, resolvedReset);
  out.warnings.push(...restoreObjectRuns(targetNode.xml as Element, keptRuns));

  // Update AST node metadata
  targetNode.metadata.text = resolvedText;
  targetNode.version += 1;
  return out;
}

function applyRunFontSize(pNode: Element, hps: number): void {
  const doc = pNode.ownerDocument;
  if (!doc) return;
  for (const r of Array.from(pNode.getElementsByTagName('w:r'))) {
    let rPr = r.getElementsByTagName('w:rPr')[0];
    if (!rPr) {
      rPr = createWordElement(doc, 'w:rPr');
      r.insertBefore(rPr, r.firstChild);
    }
    let szEl = rPr.getElementsByTagName('w:sz')[0];
    if (!szEl) {
      szEl = createWordElement(doc, 'w:sz');
      rPr.appendChild(szEl);
    }
    szEl.setAttribute('w:val', String(hps));

    let szCsEl = rPr.getElementsByTagName('w:szCs')[0];
    if (!szCsEl) {
      szCsEl = createWordElement(doc, 'w:szCs');
      rPr.appendChild(szCsEl);
    }
    szCsEl.setAttribute('w:val', String(hps));
    // Explicit size lands wherever it was appended; re-sort to the ECMA-376
    // run-property sequence (w:sz/w:szCs sit before highlight/underline) so
    // the merged properties stay schema-ordered before the save normalizer.
    sortChildrenBySchema(rPr, RPR_CHILD_ORDER);
  }
}

/**
 * Replace text in a paragraph node.
 *
 * Preserve-or-carry merge strategy: the rewritten run clones the whole
 * w:rPr of the first original run, so replacement text keeps the target
 * paragraph's formatting (bold, italic, color, size, fonts, underline,
 * effects). Explicit op properties win over carried ones, mirroring the
 * style op merge rule: fontSize/sz overwrites the carried w:sz/w:szCs pair
 * and leaves every other carried property untouched. resetStyle:true drops
 * all carried properties so the text takes the paragraph style, and an
 * explicit preserveFormatting:false keeps the legacy strip behavior. The
 * tracked path (emitTrackedReplace) already carries unconditionally, so
 * both paths now agree. Carried and merged properties are re-sorted to the
 * ECMA-376 run-property sequence (RPR_CHILD_ORDER); the save normalizer
 * re-enforces the same order, so this is idempotent. Note anchors are
 * carried separately.
 */
export function replaceText(node: Element, newText: string, preserveFormatting: boolean, fontSize?: number | string, resetStyle?: boolean): ReplaceOperationResult {
  if (node.tagName !== 'w:p') return { warnings: [] };

  const doc = node.ownerDocument;
  if (!doc) return { warnings: [] };

  const carried = snapshotNoteRefs(node);
  const runs = Array.from(node.getElementsByTagName('w:r'));
  let firstRunPr: Element | null = null;
  
  if (runs.length > 0) {
    const rPr = runs[0].getElementsByTagName('w:rPr')[0];
    if (rPr) {
      firstRunPr = rPr.cloneNode(true) as Element;
    }
  }

  // Remove all existing runs
  runs.forEach(r => r.parentNode?.removeChild(r));

  // Create new run
  const newRun = createWordElement(doc, 'w:r');
  const inherit = preserveFormatting && resetStyle !== true;
  if (firstRunPr && inherit) {
    newRun.appendChild(firstRunPr);
  }

  const hps = parseFontSize(fontSize);
  if (hps !== undefined) {
    let rPr = newRun.getElementsByTagName('w:rPr')[0];
    if (!rPr) {
      rPr = createWordElement(doc, 'w:rPr');
      newRun.insertBefore(rPr, newRun.firstChild);
    }
    let szEl = rPr.getElementsByTagName('w:sz')[0];
    if (!szEl) {
      szEl = createWordElement(doc, 'w:sz');
      rPr.appendChild(szEl);
    }
    szEl.setAttribute('w:val', String(hps));

    let szCsEl = rPr.getElementsByTagName('w:szCs')[0];
    if (!szCsEl) {
      szCsEl = createWordElement(doc, 'w:szCs');
      rPr.appendChild(szCsEl);
    }
    szCsEl.setAttribute('w:val', String(hps));
  }

  // Carried and merged properties keep ECMA-376 run-property sequence
  // (explicit w:sz/w:szCs may append after later-ordered properties such
  // as highlight or underline); re-sort so the rewritten run stays valid
  // before the save normalizer re-enforces the same order (idempotent).
  const mergedPr = newRun.getElementsByTagName('w:rPr')[0];
  if (mergedPr) sortChildrenBySchema(mergedPr, RPR_CHILD_ORDER);
  
  const textNode = createTextElement(doc, newText);
  newRun.appendChild(textNode);

  // B-3 bookmark-aware replace: when the paragraph carries a bookmark span,
  // appending lands the new run AFTER w:bookmarkEnd (collapsing the bookmark
  // to a 0-width anchor). Insert BEFORE the first bookmarkEnd so the span
  // keeps wrapping the replaced text.
  let bookmarkEnd: Element | null = null;
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i] as Element;
    if (child.nodeType === 1 && (child as Element).tagName === 'w:bookmarkEnd') {
      bookmarkEnd = child as Element;
      break;
    }
  }
  if (!bookmarkEnd) {
    const nested = node.getElementsByTagName('w:bookmarkEnd');
    if (nested.length > 0) bookmarkEnd = nested[0] as unknown as Element;
  }
  if (
    bookmarkEnd &&
    bookmarkEnd.parentNode &&
    (bookmarkEnd.parentNode as Element).tagName !== 'w:r'
  ) {
    bookmarkEnd.parentNode.insertBefore(newRun, bookmarkEnd);
  } else {
    node.appendChild(newRun);
  }
  return { warnings: restoreNoteRefs(node, carried) };
}

const OBJECT_RUN_TAGS = ['w:drawing', 'w:pict', 'w:object', 'w:commentReference'];

function runHoldsObject(run: Element): boolean {
  return OBJECT_RUN_TAGS.some((tag) => run.getElementsByTagName(tag).length > 0);
}

function hasKeptAncestor(run: Element, kept: Element[]): boolean {
  for (let n = run.parentNode; n; n = n.parentNode) {
    if (kept.includes(n as Element)) return true;
  }
  return false;
}

/** Detaches the runs that hold an object (picture, chart, shape, OLE object, comment reference) from the paragraph. */
function detachObjectRuns(p: Element): Element[] {
  const kept: Element[] = [];
  for (const run of Array.from(p.getElementsByTagName('w:r')) as Element[]) {
    if (!runHoldsObject(run) || hasKeptAncestor(run, kept)) continue;
    kept.push(run);
  }
  for (const run of kept) {
    run.parentNode?.removeChild(run);
    // A run can hold text next to the object; the text is being replaced, the object is not.
    for (const child of Array.from(run.childNodes) as Element[]) {
      if (child.nodeType === 1 && child.tagName === 'w:t') run.removeChild(child);
    }
  }
  return kept;
}

/** Puts the detached runs back at the end of the paragraph, in their original order. */
function restoreObjectRuns(p: Element, kept: Element[]): string[] {
  for (const run of kept) p.appendChild(run);
  const objects = kept.filter((run) => ['w:drawing', 'w:pict', 'w:object'].some((tag) => run.getElementsByTagName(tag).length > 0));
  return objects.length === 0
    ? []
    : [`[ASTra Engine] replace kept ${objects.length} picture or object run(s) and moved them to the end of the paragraph.`];
}

type CarriedNote = { tag: string; id: string; rPr: Element | null; ref: Element };

function snapshotNoteRefs(p: Element): CarriedNote[] {
  const out: CarriedNote[] = [];
  for (const tag of ['w:footnoteReference', 'w:endnoteReference']) {
    const els = Array.from(p.getElementsByTagName(tag));
    for (const el of els) {
      const ref = el as unknown as Element;
      let id = '';
      try {
        id = ref.getAttribute('w:id') || ref.getAttribute('id') || '';
      } catch {
        id = '';
      }
      let rPr: Element | null = null;
      try {
        // eslint-disable-next-line no-undef
        let host = ref.parentNode;
        while (host && host !== p) {
          if (host.nodeType === 1 && (host as unknown as Element).tagName === 'w:r') {
            const pr = (host as unknown as Element).getElementsByTagName('w:rPr')[0] as unknown as Element | undefined;
            if (pr) rPr = pr.cloneNode(true) as unknown as Element;
            break;
          }
          host = host.parentNode;
        }
      } catch {
        rPr = null;
      }
      try {
        out.push({ tag, id, rPr, ref: ref.cloneNode(false) as unknown as Element });
      } catch {
        out.push({ tag, id, rPr, ref: ref as unknown as Element });
      }
    }
  }
  return out;
}

function restoreNoteRefs(p: Element, notes: CarriedNote[]): string[] {
  if (!notes || notes.length === 0) return [];
  const label = notes.map(n => `${n.tag} w:id=${n.id || '(missing)'}`).join(', ');
  const doc = p.ownerDocument;
  if (!doc) {
    return [`[ASTra Engine] replace dropped note references (${label}): no owner document to carry onto replacement runs.`];
  }
  try {
    const present = new Set<string>();
    for (const tag of ['w:footnoteReference', 'w:endnoteReference']) {
      const els = p.getElementsByTagName(tag);
      for (let i = 0; i < els.length; i++) {
        let id = '';
        try {
          id = (els[i] as unknown as Element).getAttribute('w:id') || '';
        } catch {
          id = '';
        }
        present.add(`${tag}:${id}`);
      }
    }
    const missing = notes.filter(n => !present.has(`${n.tag}:${n.id}`));
    if (missing.length === 0) return [];
    for (const m of missing) {
      const run = createWordElement(doc, 'w:r');
      if (m.rPr) {
        try {
          run.appendChild(m.rPr.cloneNode(true));
        } catch {
          /* anchor still carried without its style */
        }
      }
      try {
        run.appendChild(m.ref.cloneNode(false));
      } catch {
        continue;
      }
      p.appendChild(run);
    }
    const after = new Set<string>();
    for (const tag of ['w:footnoteReference', 'w:endnoteReference']) {
      const els = p.getElementsByTagName(tag);
      for (let i = 0; i < els.length; i++) {
        let id = '';
        try {
          id = (els[i] as unknown as Element).getAttribute('w:id') || '';
        } catch {
          id = '';
        }
        after.add(`${tag}:${id}`);
      }
    }
    const still = missing.filter(m => !after.has(`${m.tag}:${m.id}`));
    if (still.length > 0) {
      const names = still.map(n => `${n.tag} w:id=${n.id || '(missing)'}`).join(', ');
      return [`[ASTra Engine] replace could not carry note references (${names}) onto replacement runs; ids dropped.`];
    }
    return [];
  } catch {
    return [`[ASTra Engine] replace could not carry note references (${label}) onto replacement runs; ids dropped.`];
  }
}

function stripInheritedRunProps(p: Element): void {
  const runs = Array.from(p.getElementsByTagName('w:r'));
  for (const r of runs) {
    const rEl = r as unknown as Element;
    try {
      if (rEl.getElementsByTagName('w:footnoteReference').length > 0) continue;
      if (rEl.getElementsByTagName('w:endnoteReference').length > 0) continue;
      const direct = Array.from(rEl.childNodes).filter(
        n => n.nodeType === 1 && (n as unknown as Element).tagName === 'w:rPr'
      );
      for (const pr of direct) rEl.removeChild(pr);
    } catch {
      /* keep going; one bad run must not block the rest */
    }
  }
}

