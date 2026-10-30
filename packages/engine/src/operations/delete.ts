import { DocumentAST, Node } from '../ast';
import { BaseOperation, validateVersion } from './base';
import { DocxParser } from '../parser';
import { emitTrackedReplace, resolveRevisionStamp } from '../ooxml/track-changes-emitter';

export interface DeleteOperation extends BaseOperation {
  action: 'delete';
  target: string | string[];
  trackChanges?: boolean;
  /** Explicit author for tracked revisions (flat alias). */
  author?: string;
  /** Explicit date (ISO 8601) for tracked revisions; defaults to the current time. */
  date?: string;
  /** When true, a missing target throws instead of landing in `skipped`. */
  strict?: boolean;
  params?: Record<string, unknown>;
}

/**
 * OX-04: idempotent delete reports what it skipped.
 *
 * Policy (deliberate idempotency kept): a missing target WITHOUT strict:true
 * is a success with an entry in `skipped` — this masks nothing because the
 * caller sees exactly which ids were stale/hallucinated. WITH strict:true
 * the same miss throws (batch then rolls back). Silent by design: the typed
 * `skipped` array (and batch `skippedIds`) is the channel, no stderr logging.
 */
export interface DeleteResult {
  deleted: string[];
  skipped: string[];
}

export function executeDelete(op: DeleteOperation, ast: DocumentAST, parser?: DocxParser): DeleteResult {
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedTarget = (pick('target', op.target) ?? op.target) as string | string[];
  const targets = Array.isArray(resolvedTarget) ? resolvedTarget : [resolvedTarget];
  if (targets.length === 0 || !targets[0]) throw new Error("Delete operation requires a target");

  const deleted: string[] = [];
  const skipped: string[] = [];
  const seenTarget = new Set<string>();

  // Tracked-delete gate: only an explicit per-op trackChanges:true (flat or
  // nested in params) requests revision semantics. Ambient
  // settings.trackChanges stays on the legacy path (remove outright): the
  // default blank document ships with tracking enabled, and flipping delete
  // to mark-instead-of-remove on that signal alone would break the
  // established idempotent-delete contract. Under the explicit flag,
  // paragraph nodes (w:p) are NOT removed — their text is converted to
  // paragraph-level w:del revisions (which ECMA-376 allows), the node is
  // retained, and the id is still reported in `deleted` as the recorded
  // deletion. Non-paragraph nodes (tables and other structural nodes) have
  // no paragraph-level revision form (whole-node removal needs structural
  // markers such as w:cellDel, which the engine does not emit), so they
  // fail closed with E_VALIDATION instead of silently vanishing.
  const explicitTrack = op.trackChanges ?? raw.trackChanges;
  const trackedDelete = explicitTrack === true;
  const stamp = resolveRevisionStamp(op.author ?? raw.author, op.date ?? raw.date, ast.settings.defaultAuthor);
  const author = stamp.author;
  const nowIso = () => stamp.date;

  const findNode = (target: string): { sectionIdx: number; index: number; node: Node } | null => {
    for (let si = 0; si < ast.sections.length; si++) {
      const index = ast.sections[si].nodes.findIndex(n => n.id === target);
      if (index !== -1) return { sectionIdx: si, index, node: ast.sections[si].nodes[index] };
    }
    return null;
  };

  if (trackedDelete) {
    // Fail-closed pre-scan: classify every deduplicated target BEFORE any
    // mutation so a structural target aborts the whole op with nothing
    // applied. Missing targets keep legacy warn/skip/strict behavior.
    for (const target of targets) {
      if (seenTarget.has(target)) continue;
      const found = findNode(target);
      if (!found) continue;
      validateVersion(found.node, op.expectedVersion);
      const tag = (found.node.xml as Element)?.tagName;
      if (tag !== 'w:p') {
        throw Object.assign(
          new Error(
            `E_VALIDATION: delete of ${found.node.type} node '${target}' under trackChanges:true is not supported: ` +
            `whole-node removal requires structural revision markers (w:cellDel and related track-change records) ` +
            `which the engine does not emit; applying it would silently drop the revision. ` +
            `Remedy: re-run with trackChanges:false to delete outright, ` +
            `or delete contained paragraphs individually for paragraph-level (w:del) revisions.`
          ),
          {
            code: 'E_VALIDATION' as const,
            retryable: false as const,
            remedy: "Re-run with trackChanges:false to delete outright, or delete contained paragraphs individually for paragraph-level (w:del) revisions.",
          }
        );
      }
    }
    seenTarget.clear();
  }

  for (const target of targets) {
    // D2 skippedIds dedupe: same ghost twice in one op reports once.
    if (seenTarget.has(target)) continue;
    seenTarget.add(target);
    let found = false;
    for (const section of ast.sections) {
      const index = section.nodes.findIndex(n => n.id === target);
      if (index !== -1) {
        const node = section.nodes[index];
        validateVersion(node, op.expectedVersion);

        if (trackedDelete) {
          // Paragraph-level tracked delete: mark, don't remove.
          const oldText = typeof node.metadata.text === 'string'
            ? node.metadata.text
            : visibleTextOf(node.xml as Element);
          const hasRevs = emitTrackedReplace(
            node.xml as Element,
            oldText,
            '',
            author,
            nowIso(),
            { getNextId: () => ast.nextChangeId++ }
          );
          if (hasRevs || oldText !== '') {
            node.metadata.hasTrackedChanges = true;
          }
          node.metadata.text = '';
          node.version += 1;
          // Node stays addressable: do NOT remove from DOM, AST, or nodeMap.
        } else {
          // Remove from DOM
          node.xml.parentNode?.removeChild(node.xml);

          // Remove from AST
          section.nodes.splice(index, 1);
        }

        found = true;
        deleted.push(target);
        break;
      }
    }
    // Keep parser.nodeMap consistent so stale lookups fail instead of
    // returning detached XML elements. Tracked deletes retain the node, so
    // the map entry stays valid and must NOT be removed.
    if (!trackedDelete) {
      parser?.removeFromNodeMap(target);
    }
    if (!found) {
      skipped.push(target);
      if (op.strict === true) {
        throw new Error(`Node ${target} not found for deletion`);
      }
    }
  }
  return { deleted, skipped };
}

/** Visible text fallback when node metadata carries no text string. */
function visibleTextOf(xml: Element): string {
  if (!xml || typeof xml.getElementsByTagName !== 'function') return '';
  const parts: string[] = [];
  const textNodes = xml.getElementsByTagName('w:t');
  for (let i = 0; i < textNodes.length; i++) {
    parts.push(textNodes[i].textContent || '');
  }
  return parts.join('');
}
