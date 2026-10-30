// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts
// (lines 252-716 + 718-719 + 729-754 + 1778-1924 + 2198-2201 + 2291-2298).
// Zero logic edits (only `export` added to previously-private cross-handler
// helpers: parse/collect/sidecar/path/restore/refresh/COMMIT_OID_RE).
import { createHash } from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import {
  AdmZip,
  DocxParser,
  getRepoDir,
  parseTableOrdinalFromXPath,
  refreshAstInPlace,
} from '@attentionform/astra-engine';
import type { DocHandle } from '../session-store';
import { COMMIT_AUTHOR_DEFAULT, COMMIT_MESSAGE_MAX } from './errors';

// ── Audit-bridge bounds + helpers (all bounded, all best-effort) ─────────────
/** Cap on discarded node ids listed in the revert echo. */
export const DISCARDED_NODE_IDS_MAX = 20;
/** Cap on affected node ids per history entry / audit-trail row. */
export const AUDIT_NODE_IDS_MAX = 100;
/** Cap on affected files listed by preview_revert / audit rows. */
export const AUDIT_FILES_MAX = 100;
/**
 * Value-summary bounds (advisory, additive, bounded):
 * - CHANGED_VALUES_MAX caps the changedValues entries per save/preview
 *   envelope (old→new {nodeId, before, after} in document order).
 * - CHANGED_VALUE_TEXT_MAX caps each before/after string (chars); longer
 *   values are sliced and the envelope truncated flag is set (honest).
 * History entries keep as-is (no value summaries there).
 */
export const CHANGED_VALUES_MAX = 20;
export const CHANGED_VALUE_TEXT_MAX = 200;

/** Advisory old→new value summary for one node (bounded strings). */
export interface ChangedValue {
  nodeId: string;
  before: string;
  after: string;
}

function truncateChangedText(s: string): { text: string; cut: boolean } {
  const str = String(s ?? '');
  if (str.length <= CHANGED_VALUE_TEXT_MAX) return { text: str, cut: false };
  return { text: str.slice(0, CHANGED_VALUE_TEXT_MAX), cut: true };
}

/**
 * Diffs two id->text snapshots into bounded old→new summaries.
 * before/after are id->text maps; primaryOrder is walked first (document
 * order of the side the caller wants listed first), secondaryOrder covers
 * ids absent from the primary side (deletions/restorations). Added nodes
 * use before:'', removed nodes use after:''. Count + per-string caps apply;
 * truncated is true when either cap was hit (honest, never silent).
 */
export function diffChangedValues(
  before: Map<string, string>,
  after: Map<string, string>,
  primaryOrder: string[],
  secondaryOrder: string[]
): { values: ChangedValue[]; truncated: boolean } {
  const values: ChangedValue[] = [];
  let truncated = false;
  const seen = new Set<string>();
  const push = (id: string, b: string, a: string): void => {
    if (values.length >= CHANGED_VALUES_MAX) {
      truncated = true;
      return;
    }
    const tb = truncateChangedText(b);
    const ta = truncateChangedText(a);
    if (tb.cut || ta.cut) truncated = true;
    values.push({ nodeId: id, before: tb.text, after: ta.text });
  };
  try {
    for (const id of primaryOrder) {
      if (seen.has(id)) continue;
      seen.add(id);
      if (!before.has(id) && !after.has(id)) continue;
      const b = before.has(id) ? (before.get(id) ?? '') : '';
      const a = after.has(id) ? (after.get(id) ?? '') : '';
      if (b !== a) push(id, b, a);
    }
    for (const id of secondaryOrder) {
      if (seen.has(id)) continue;
      seen.add(id);
      if (!before.has(id) && !after.has(id)) continue;
      const b = before.has(id) ? (before.get(id) ?? '') : '';
      const a = after.has(id) ? (after.get(id) ?? '') : '';
      if (b !== a) push(id, b, a);
    }
  } catch {
    truncated = true;
  }
  return { values, truncated };
}

/**
 * Across-parse value diff for save/preview envelopes (advisory, additive,
 * bounded). Same-parse ids are stable, but a FRESH parse mints positional
 * ids (t_001…) while in-session inserts keep minted ids (_bk_new_…), so an
 * id-keyed diff across a save/preview boundary reports phantom
 * added+removed pairs for unchanged nodes. Strategy: shared ids diff by id
 * (walked in live document order); remaining orphans zip positionally when
 * both sides hold the same orphan count (silent when the paired texts are
 * equal — the clean-resave case); otherwise orphans degrade honestly to
 * added (before:'') / removed (after:'') entries. nodeId always prefers the
 * live side (parsed-commit/baseline ids are ephemeral). liveIsAfter selects
 * direction: save passes (live, baseline, true) = baseline→live; preview
 * passes (live, target, false) = live→target. Count + per-string caps apply;
 * truncated is true when either cap was hit (honest, never silent).
 * positional is true when either side held orphan ids (no shared id across
 * the boundary — usually a reparse id mismatch, common under duplicate
 * native ids): paired entries were zipped by position, so consumers must
 * treat before/after pairing as positional, not identity.
 */
export function diffChangedValuesAcrossParse(
  live: { order: string[]; texts: Map<string, string> },
  other: { order: string[]; texts: Map<string, string> },
  liveIsAfter: boolean
): { values: ChangedValue[]; truncated: boolean; positional: boolean } {
  const values: ChangedValue[] = [];
  let truncated = false;
  let positional = false;
  const push = (id: string, b: string, a: string): void => {
    if (values.length >= CHANGED_VALUES_MAX) {
      truncated = true;
      return;
    }
    const tb = truncateChangedText(b);
    const ta = truncateChangedText(a);
    if (tb.cut || ta.cut) truncated = true;
    values.push({ nodeId: id, before: tb.text, after: ta.text });
  };
  try {
    const before = liveIsAfter ? other : live;
    const after = liveIsAfter ? live : other;
    // Shared ids: diff by id in live document order.
    for (const id of live.order) {
      if (!other.texts.has(id)) continue;
      const b = before.texts.get(id) ?? '';
      const a = after.texts.get(id) ?? '';
      if (b !== a) push(id, b, a);
    }
    // Orphans (id exists on one side only — usually a reparse id mismatch).
    const liveOrphans = live.order.filter((id) => !other.texts.has(id));
    const otherOrphans = other.order.filter((id) => !live.texts.has(id));
    if (liveOrphans.length > 0 || otherOrphans.length > 0) positional = true;
    if (liveOrphans.length === otherOrphans.length) {
      for (let i = 0; i < liveOrphans.length; i++) {
        const lid = liveOrphans[i];
        const oid = otherOrphans[i];
        const b = before.texts.get(liveIsAfter ? oid : lid) ?? '';
        const a = after.texts.get(liveIsAfter ? lid : oid) ?? '';
        if (b !== a) push(lid, b, a);
      }
    } else {
      for (const id of liveOrphans) {
        const t = live.texts.get(id) ?? '';
        if (liveIsAfter) push(id, '', t);
        else push(id, t, '');
      }
      for (const id of otherOrphans) {
        const t = other.texts.get(id) ?? '';
        if (liveIsAfter) push(id, t, '');
        else push(id, '', t);
      }
    }
  } catch {
    truncated = true;
  }
  return { values, truncated, positional };
}

/**
 * Parses a target commit's document.xml into id->text via a synthetic docx
 * (live package parts + target document.xml, same DocxParser path as live,
 * so para/table ids and w:t concatenation match exactly). Best-effort:
 * returns null when the synthetic package cannot be parsed — callers emit
 * empty changedValues + reason and MUST NOT fail the preview on this probe.
 */
export async function parseTargetIdTextsForPreview(
  liveSnap: Map<string, Buffer>,
  targetXml: string
): Promise<{ order: string[]; texts: Map<string, string> } | null> {
  try {
    const zip = new AdmZip();
    for (const [name, buf] of liveSnap) {
      if (name === 'word/document.xml') continue;
      zip.addFile(name, Buffer.from(buf));
    }
    zip.addFile('word/document.xml', Buffer.from(targetXml, 'utf8'));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    return collectAstIdTexts(ast);
  } catch {
    return null;
  }
}

/**
 * HOLD retention note (honest, no WORM claims): the .astra/ sidecar is
 * deletable local state living beside the .docx — anyone with file access
 * can wipe it. Full server-side retention is OUT of scope: history lives
 * beside the document and dies with it; the audit handlers below only
 * report what is currently on disk.
 */
export const AUDIT_RETENTION_NOTE =
  'HOLD: .astra/ sidecar history is deletable local state (anyone with file access can remove it), ' +
  'NOT WORM/immutable storage. Full server-side retention is OUT of scope: history lives beside the .docx ' +
  'and dies with it; this report covers only what is currently on disk.';

/** One created node per applied op, for the execute_batch result envelope. */
export interface BatchOpCreatedIds {
  index: number;
  action: string;
  createdNodeIds: string[];
}

/** Echo of unsaved edits discarded by a revert (counts + ids only, never text). */
export interface DiscardedUnsavedSummary {
  wasDirty: boolean;
  /** Added + modified + deleted-vs-saved node count (0 when clean). */
  discardedNodeCount: number;
  /** Affected node ids in live-document order (bounded, [] when clean). */
  discardedNodeIds: string[];
  /** True when the id bound above was hit. */
  truncated: boolean;
}

/** Node ids + live texts in document order (single AST walk). */
export function collectAstIdTexts(ast: {
  sections: Array<{ nodes: Array<{ id: string; metadata: { text?: unknown } }> }>;
}): { order: string[]; texts: Map<string, string> } {
  const order: string[] = [];
  const texts = new Map<string, string>();
  try {
    for (const sec of ast.sections ?? []) {
      for (const n of sec.nodes ?? []) {
        if (typeof n.id !== 'string' || !n.id) continue;
        order.push(n.id);
        texts.set(n.id, String(n.metadata?.text ?? ''));
      }
    }
  } catch {
    /* best-effort: return what was collected */
  }
  return { order, texts };
}

/**
 * Best-effort parse of saved .docx bytes into id->text (for the dirty diff).
 * Returns null when the bytes cannot be parsed — callers fall back to a
 * live-only summary and MUST NOT fail the revert on this probe.
 */
export async function parseSavedIdTextsOrdered(
  savedBytes: Buffer | null
): Promise<{ order: string[]; texts: Map<string, string> } | null> {
  if (!savedBytes || savedBytes.length === 0) return null;
  try {
    const parser = new DocxParser(Buffer.from(savedBytes));
    const ast = await parser.getAST();
    return collectAstIdTexts(ast);
  } catch {
    return null;
  }
}

export async function parseSavedIdTexts(savedBytes: Buffer | null): Promise<Map<string, string> | null> {
  const parsed = await parseSavedIdTextsOrdered(savedBytes);
  return parsed ? parsed.texts : null;
}

/**
 * Summarizes the unsaved edits a revert is about to discard: nodes added or
 * text-modified vs the last saved bytes count as discarded, plus nodes
 * deleted-vs-saved (their deletion is the discarded edit; they contribute
 * ids + count). Counts + ids only, never text. Everything bounded; clean
 * sessions yield the zero summary (never null, so callers can always echo it).
 */
export function summarizeDiscardedUnsaved(
  live: { order: string[]; texts: Map<string, string> },
  saved: Map<string, string> | null,
  fallbackOpCount: number
): DiscardedUnsavedSummary {
  const zero: DiscardedUnsavedSummary = {
    wasDirty: false,
    discardedNodeCount: 0,
    discardedNodeIds: [],
    truncated: false,
  };
  if (!saved) {
    // No saved baseline to diff against (unreadable bytes): report the op
    // delta honestly. Node ids stay [] here — without a baseline no id can
    // be attributed, and no text is echoed.
    const ops = Math.max(0, Math.floor(fallbackOpCount));
    if (ops <= 0) return zero;
    return {
      wasDirty: true,
      discardedNodeCount: ops,
      discardedNodeIds: [],
      truncated: false,
    };
  }
  const addedOrModified: string[] = [];
  for (const id of live.order) {
    const liveText = live.texts.get(id) ?? '';
    if (!saved.has(id) || saved.get(id) !== liveText) addedOrModified.push(id);
  }
  const liveSet = new Set(live.order);
  const removed: string[] = [];
  for (const id of saved.keys()) {
    if (!liveSet.has(id)) removed.push(id);
  }
  if (addedOrModified.length === 0 && removed.length === 0) return zero;
  const ids = [...addedOrModified, ...removed];
  return {
    wasDirty: true,
    discardedNodeCount: addedOrModified.length + removed.length,
    discardedNodeIds: ids.slice(0, DISCARDED_NODE_IDS_MAX),
    truncated: ids.length > DISCARDED_NODE_IDS_MAX,
  };
}

/**
 * Diffs two id->text snapshots (pre vs post restore): added, removed, or
 * text-changed ids, returned in post-restore document order, bounded.
 * Used for revert/redo affected-node-ids (honest restore footprint).
 */
export function diffIdTexts(
  before: Map<string, string>,
  afterOrder: string[],
  after: Map<string, string>
): string[] {
  const out: string[] = [];
  try {
    for (const id of afterOrder) {
      if (out.length >= AUDIT_NODE_IDS_MAX) break;
      if (!before.has(id) || before.get(id) !== after.get(id)) out.push(id);
    }
    if (out.length < AUDIT_NODE_IDS_MAX) {
      const afterSet = new Set(afterOrder);
      for (const id of before.keys()) {
        if (out.length >= AUDIT_NODE_IDS_MAX) break;
        if (!afterSet.has(id)) out.push(id);
      }
    }
  } catch {
    /* best-effort */
  }
  return out;
}

/**
 * Creator-op predicate (verified by execution probe 2026-09-22: insert,
 * callout, equation/insert, field/insertTOC|insertPageNumber|insertWordCount,
 * coverPage, section/split add section nodes; replace/delete/style/list/
 * comment/trackedChange/hyperlink/footnote/headerFooter/metadata do not).
 * chart/image inserts also create nodes (media paths need fileData, covered
 * by the same rule). Everything else never claims created ids.
 */
function isCreatorOp(action: string, op: Record<string, unknown>): boolean {
  switch (action) {
    case 'insert':
    case 'callout':
    case 'coverPage':
      return true;
    case 'equation':
      return op['operation'] === 'insert';
    case 'field':
      return op['operation'] === 'insertTOC' || op['operation'] === 'insertPageNumber' || op['operation'] === 'insertWordCount';
    case 'chart':
      return op['operation'] === undefined || op['operation'] === 'insert';
    case 'image':
      return op['operation'] === 'insert';
    case 'section':
      return op['operation'] === 'split';
    default:
      return false;
  }
}

/**
 * Attributes post-batch created ids to per-op results. Client-supplied `id`
 * claims are exact when the id is present in the post-batch id set (the
 * engine uses op.id verbatim as the node id — verified by probe); ids
 * absent from that set are never echoed (footnote honors no id and creates
 * zero AST nodes, so its id stays unclaimed). Remaining ids go to
 * creator-candidate ops: a single unclaimed creator takes all of them
 * (exact for the common single-insert batch and for multi-node creators
 * like insertTOC/coverPage alone in a batch); multiple unclaimed creators
 * split best-effort in batch order with leftovers on the last creator.
 * Non-creators always report [].
 */
export function attributeCreatedIds(
  actions: string[],
  ops: Array<Record<string, unknown>>,
  createdOrdered: string[]
): BatchOpCreatedIds[] {
  const results: BatchOpCreatedIds[] = actions.map((action, index) => ({ index, action, createdNodeIds: [] }));
  const remaining = [...createdOrdered];
  const claim = (index: number, id: string): void => {
    const at = remaining.indexOf(id);
    if (at >= 0) remaining.splice(at, 1);
    results[index].createdNodeIds.push(id);
  };
  // Exact pass: client-supplied ids, claimed ONLY when present in the
  // post-batch id set (footnote honors no id and creates zero AST nodes,
  // so an unbacked id must not echo as a phantom; the op keeps [] via the
  // exact attribution below). Schema allows `id` on insert, equation,
  // footnote, chart; strict schemas reject it elsewhere with E_VALIDATION.
  for (let i = 0; i < ops.length; i++) {
    const rawId = ops[i]?.['id'];
    if (typeof rawId === 'string' && rawId && remaining.includes(rawId)) claim(i, rawId);
  }
  const unclaimedCreators: number[] = [];
  for (let i = 0; i < ops.length; i++) {
    if (results[i].createdNodeIds.length === 0 && isCreatorOp(actions[i], ops[i] ?? {})) unclaimedCreators.push(i);
  }
  if (unclaimedCreators.length === 1) {
    // Exact: the only op that could have created nodes owns the remainder.
    for (const id of remaining) results[unclaimedCreators[0]].createdNodeIds.push(id);
  } else if (unclaimedCreators.length > 1) {
    // Best-effort split: one id per creator in batch order, leftovers on last.
    let k = 0;
    for (const idx of unclaimedCreators) {
      if (k < remaining.length) {
        results[idx].createdNodeIds.push(remaining[k]);
        k += 1;
      }
    }
    const last = unclaimedCreators[unclaimedCreators.length - 1];
    while (k < remaining.length) {
      results[last].createdNodeIds.push(remaining[k]);
      k += 1;
    }
  }
  // Zero creators but created ids exist (should not happen; engine changed?):
  // keep the flat list exact and leave per-op empty rather than misattribute.
  return results;
}

/** sha256 hex over committed document.xml bytes (audit-trail package hash). */
export function hashDocumentXml(xml: string): string {
  return createHash('sha256').update(Buffer.from(xml, 'utf8')).digest('hex');
}

/**
 * R-1: counts pending tracked revisions (w:ins + w:del) in a live parser DOM.
 * Best-effort, never throws. Mirrors OperationEngine.countPendingRevisions
 * for the save path (which does not go through executeBatch).
 */
export function countPendingRevisionsInParser(parser: {
  getDocumentDom?: () => unknown;
  getDom?: () => unknown;
}): number {
  try {
    const dom = (parser?.getDocumentDom?.() ?? parser?.getDom?.()) as
      | { getElementsByTagName?: (tag: string) => { length: number } }
      | undefined;
    if (dom && typeof dom.getElementsByTagName === 'function') {
      return dom.getElementsByTagName('w:ins').length + dom.getElementsByTagName('w:del').length;
    }
  } catch {
    /* best-effort: report 0 rather than fail the envelope */
  }
  return 0;
}

/**
 * T-5 revert commit subject: `Revert [N steps] to <oid> by <author>[: reason]`.
 * N = how far back from the pre-revert head the target sat (1 for the common
 * single-step revert); oid is the 7-char short form; reason is appended only
 * when non-blank. Capped to COMMIT_MESSAGE_MAX like every other subject.
 */
export function formatRevertCommitMessage(
  stepsBack: number,
  targetOid: string,
  author: string,
  reason: string
): string {
  const n = Math.max(0, Math.floor(stepsBack));
  const short = typeof targetOid === 'string' ? targetOid.slice(0, 7) : '';
  const by = author || COMMIT_AUTHOR_DEFAULT;
  const subject = reason ? `Revert [${n} steps] to ${short} by ${by}: ${reason}` : `Revert [${n} steps] to ${short} by ${by}`;
  return subject.length > COMMIT_MESSAGE_MAX ? subject.slice(0, COMMIT_MESSAGE_MAX) : subject;
}

/** 40-hex-char git object id (full oid; history tools never accept prefixes). */
export const COMMIT_OID_RE = /^[0-9a-fA-F]{40}$/;

/**
 * Rebuilds the handle AST in place after a package-level restore
 * (revert/redo): the parser DOMs were rewound underneath, so the live AST
 * object (shared with handle.engine) is refreshed field-by-field, mirroring
 * the engine rollback path so external references stay valid.
 */
export async function refreshHandleAstInPlace(handle: DocHandle): Promise<void> {
  await refreshAstInPlace(handle.ast, handle.parser);
}

/**
 * Phase C-1c: maps a validator failure to an xpath-specific remedy instead
 * of the generic copy-paste. Matches on the validator xPath (falls back to
 * the description text): drawing docPr/@id, table tbl children, and
 * relationships each name their own fix.
 */
export function remedyForValidationError(xPath: unknown, description: unknown): string {
  const xp = typeof xPath === 'string' ? xPath : '';
  const desc = typeof description === 'string' ? description : '';
  const xpLower = xp.toLowerCase();
  const hay = `${xpLower} ${desc.toLowerCase()}`;
  if (xpLower.includes('docpr') || (xpLower.includes('@id') && hay.includes('drawing'))) {
    return (
      'Duplicate or invalid drawing id (wp:docPr): re-inspect the flagged image/drawing node, ' +
      'ensure every drawing carries a unique non-zero docPr id (no copies sharing one id), ' +
      'fix via execute_batch, then retry save_document.'
    );
  }
  if (
    xpLower.includes('w:tbl') ||
    xpLower.includes('/w:tbl') ||
    xpLower.includes('tblpr') ||
    xpLower.includes('tblgrid') ||
    xpLower.includes('/w:tr') ||
    xpLower.includes('/w:tc') ||
    xpLower.includes('gridcol')
  ) {
    return (
      'Table structure invalid (tblPr/tblGrid/tr/tc): re-inspect the flagged table, ' +
      'check tblGrid gridCol count matches every row cell count and children follow schema order, ' +
      'fix via execute_batch, then retry save_document.'
    );
  }
  if (
    hay.includes('relationship') ||
    hay.includes('.rels') ||
    hay.includes('r:id') ||
    /rid\d/.test(hay)
  ) {
    return (
      'Document relationships invalid: re-inspect hyperlink/image/chart targets, ' +
      'ensure every rId referenced from the document exists in document.xml.rels with a valid target, ' +
      'fix via execute_batch, then retry save_document.'
    );
  }
  if (hay.includes('duplicate') && (hay.includes('paraid') || hay.includes('w14:paraid'))) {
    return (
      `Duplicate paragraph id${xp ? ` (${xp})` : ''}: two paragraphs share one native w14:paraId, ` +
      'so a merge/migration cannot tell them apart. ' +
      'Delete the repeat paragraph (its session id carries a _dup suffix — see inspect_document) and reinsert its text for a fresh id, ' +
      'then retry save_document.'
    );
  }
  if (hay.includes('unknown element')) {
    return (
      `Non-OOXML markup${xp ? ` (${xp})` : ''}: the flagged element is not valid WordprocessingML. ` +
      'Remove or repair the element (re-export the file from Word when the package itself is suspect), ' +
      'fix via execute_batch, then retry save_document.'
    );
  }
  if (hay.includes('data-astra-id')) {
    const part =
      xpLower.includes('w:ftr') || xpLower.includes('footer')
        ? 'footer part'
        : xpLower.includes('w:hdr') || xpLower.includes('header')
          ? 'header part'
          : 'saved part';
    return (
      `Session tracking attribute leaked into the ${part}${xp ? ` (${xp})` : ''}: ` +
      'a field or insert op stamped a session-only id the validator rejects. ' +
      'Re-run the field insertPageNumber/insertWordCount op (or headerFooter set) so the part is rebuilt without session ids, ' +
      'then retry save_document.'
    );
  }
  if (
    xpLower.includes('w:ftr') ||
    xpLower.includes('w:hdr') ||
    xpLower.includes('headerreference') ||
    xpLower.includes('footerreference') ||
    xpLower.includes('footer') ||
    xpLower.includes('header')
  ) {
    const part = xpLower.includes('ftr') || xpLower.includes('footer') ? 'footer' : 'header';
    return (
      `Header/footer part invalid${xp ? ` (${xp})` : ''}: the ${part} part failed schema validation. ` +
      'Re-inspect the header/footer content (headerFooter set rebuilds the part), ' +
      'fix via execute_batch, then retry save_document.'
    );
  }
  return 'Fix the flagged operations (re-inspect TOON, correct hex/half-point/twips values) and retry save_document.';
}

/**
 * Table node ids in live document order (session ids such as t_001.. or
 * minted insert ids). Best-effort, never throws. Used to resolve a validator
 * table ordinal to the concrete node id the agent should fix.
 */
export function collectTableNodeIdsInOrder(ast: unknown): string[] {
  try {
    const sections = (ast as { sections?: Array<{ nodes?: Array<Record<string, unknown>> }> } | null | undefined)?.sections;
    if (!Array.isArray(sections)) return [];
    const out: string[] = [];
    for (const sec of sections) {
      const nodes = (sec as { nodes?: Array<Record<string, unknown>> } | null | undefined)?.nodes;
      if (!Array.isArray(nodes)) continue;
      for (const n of nodes) {
        const id = (n as { id?: unknown } | null)?.id;
        if (typeof id !== 'string' || !id) continue;
        if ((n as { type?: unknown }).type === 'table' || /^t_\d+/.test(id)) out.push(id);
      }
    }
    return out;
  } catch {
    return [];
  }
}

/**
 * Builds the location fragment for a save-blocked envelope. Always names a
 * location: the validator xPath verbatim plus the resolved table ordinal and
 * session node id when the failure is table-related; the bare xPath for other
 * parts; an explicit no-xPath fallback (with the live table inventory) when
 * the validator gave no path. Never throws.
 * Ordinals are body-top-level only: a nested-table fault resolves to the
 * outer containing table (safe direction, never a wrong sibling).
 */
export function formatValidationLocation(
  xPath: unknown,
  tableIds: readonly string[] = []
): string {
  const xp = typeof xPath === 'string' && xPath ? xPath : undefined;
  const ordinal = parseTableOrdinalFromXPath(xp);
  if (xp && ordinal !== null) {
    const node = ordinal >= 1 && ordinal <= tableIds.length ? tableIds[ordinal - 1] : undefined;
    return node ? ` at ${xp} (table ${ordinal}, node ${node})` : ` at ${xp} (table ${ordinal})`;
  }
  if (xp && /w:tbl/i.test(xp)) {
    const ids = tableIds.length > 0 ? ` (tables in session: ${tableIds.join(', ')})` : '';
    return ` at ${xp} (table${ids})`;
  }
  if (xp) return ` at ${xp}`;
  const ids = tableIds.length > 0 ? ` ${tableIds.length} table(s) in session: ${tableIds.join(', ')}` : ' no tables in session';
  return ` (validator gave no xPath;${ids}; re-inspect to locate the flagged node)`;
}

/**
 * Shapes a rejected restore candidate (revert/redo) into the same
 * xpath-specific E_VALIDATION throw the validating save uses, so callers get
 * a fix-and-retry hint instead of a bare schema dump.
 */
export function invalidRestoreError(toolName: string, errors: unknown, verb: string, tableIds: readonly string[] = []): Error {
  const list = Array.isArray(errors) ? errors : [];
  const first = list[0] as
    | { description?: unknown; message?: unknown; xPath?: unknown; path?: unknown }
    | undefined;
  const desc =
    (typeof first?.description === 'string' && first.description) ||
    (typeof first?.message === 'string' && first.message) ||
    'OOXML validation failed';
  const xp = typeof first?.xPath === 'string' ? first.xPath : (typeof first?.path === 'string' ? first.path : undefined);
  const at = formatValidationLocation(xp, tableIds);
  const count = list.length > 1 ? ` (${list.length} errors, showing first)` : '';
  return Object.assign(
    new Error(`E_VALIDATION: ${toolName} blocked — ${verb} invalid${at}${count}: ${String(desc).slice(0, 2000)}`),
    {
      code: 'E_VALIDATION',
      remedy: remedyForValidationError(xp, desc),
    }
  );
}

/** Best-effort restore double-fault visibility: file + op context on stderr. */
export function logRestoreFailure(op: string, filePath: string, step: string, err: unknown): void {
  const msg = err instanceof Error ? err.message : String(err);
  console.error(`[${op}] restore failed for "${filePath}" step ${step}: ${msg}`);
}

/**
 * Honest remedy: only claim "left unchanged" when every restore step was
 * verified. When any step failed, point at inspect/history verification.
 */
export function honestRestoreRemedy(
  op: string,
  memoryOk: boolean,
  diskOk: boolean,
  okRemedy: string,
  detail: string
): string {
  if (memoryOk && diskOk) return okRemedy;
  return `${op}: ${detail}, restore attempted, VERIFY via inspect/history before continuing. Run inspect_document and get_document_history to confirm live vs saved state before retrying.`;
}

export function pathDirname(p: string): string {
  const idx = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'));
  return idx > 0 ? p.slice(0, idx) : '.';
}

/** True when the sidecar git repo exists for `docPath` (read-only probe). */
export function sidecarRepoExists(docPath: string): boolean {
  try {
    return fs.existsSync(path.join(getRepoDir(docPath), '.git'));
  } catch {
    return false;
  }
}
