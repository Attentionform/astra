type DiffChunk = { type: 'equal' | 'delete' | 'insert'; text: string };

import { resolveDeterministicZipDate, resolveReproducibleMode } from '../zip-hardening';
import { stripIllegalXmlChars } from './xml-text';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Normalize the (explicit, env) overload: first arg may be legacy env object. */
function splitDateArgs(
  explicitOrEnv?: boolean | NodeJS.ProcessEnv,
  env?: NodeJS.ProcessEnv
): { explicit: boolean | undefined; env: NodeJS.ProcessEnv } {
  if (typeof explicitOrEnv === 'boolean') {
    return { explicit: explicitOrEnv, env: env ?? process.env };
  }
  if (explicitOrEnv && typeof explicitOrEnv === 'object') {
    return { explicit: undefined, env: explicitOrEnv };
  }
  return { explicit: undefined, env: env ?? process.env };
}

/**
 * R-7 deterministic revision stamps (deterministic-by-default).
 *
 * Honors SOURCE_DATE_EPOCH (seconds since unix epoch), the same env var the
 * zip layer uses for entry headers (see zip-hardening.ts
 * resolveDeterministicZipDate). Parsing mirrors that helper (trim, Number,
 * finite, >= 0, valid Date); absent/invalid falls back to the deterministic
 * epoch (1980-01-01Z or SOURCE_DATE_EPOCH when set) when reproducible, or to
 * wall-clock when opted out via explicit false, ASTRA_REPRODUCIBLE=0, or
 * ASTRA_NO_REPRODUCIBLE=1.
 *
 * Default is deterministic: same input twice yields byte-identical w:date
 * stamps without any env. Opt-out restores wall-clock. The explicit flag
 * threads through the save layer, so rezip({ reproducible: false }) fully
 * unpins already-minted stamps (see parser rezip normalization).
 */
export function deterministicDate(explicitOrEnv?: boolean | NodeJS.ProcessEnv, env?: NodeJS.ProcessEnv): Date {
  const { explicit, env: effEnv } = splitDateArgs(explicitOrEnv, env);
  try {
    if (!resolveReproducibleMode(explicit, effEnv)) return new Date();
    return resolveDeterministicZipDate(effEnv);
  } catch {
    try {
      return new Date(Date.UTC(1980, 0, 1, 0, 0, 0));
    } catch {
      return new Date();
    }
  }
}

/** Revision w:date in full ISO form (millis kept, matches legacy delete path). */
export function deterministicRevisionIso(explicitOrEnv?: boolean | NodeJS.ProcessEnv, env?: NodeJS.ProcessEnv): string {
  const { explicit, env: effEnv } = splitDateArgs(explicitOrEnv, env);
  return deterministicDate(explicit, effEnv).toISOString();
}

/** Comment w:date in Word form (millis stripped, matches legacy comments path). */
export function deterministicCommentIso(explicitOrEnv?: boolean | NodeJS.ProcessEnv, env?: NodeJS.ProcessEnv): string {
  const { explicit, env: effEnv } = splitDateArgs(explicitOrEnv, env);
  return deterministicDate(explicit, effEnv).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * Who and when for a revision: the caller's author, else the document default, else "ASTra";
 * the caller's date, else the current time (or the pinned stamp when output is reproducible).
 * Every revision writer goes through here so they cannot disagree about either.
 */
export function resolveRevisionStamp(
  explicitAuthor: unknown,
  explicitDate: unknown,
  settingsAuthor: unknown,
): { author: string; date: string } {
  const date = typeof explicitDate === 'string' && explicitDate.trim() !== '' ? explicitDate : deterministicRevisionIso();
  return { author: resolveDeterministicAuthor(explicitAuthor, settingsAuthor), date };
}

/** Env var naming the document's default author for revisions and comments that do not give one. */
export const AUTHOR_ENV = 'ASTRA_AUTHOR';

/** The author a freshly opened document starts with: ASTRA_AUTHOR when set, else "ASTra". */
export function resolveDefaultAuthor(env: NodeJS.ProcessEnv = process.env): string {
  const raw = env?.[AUTHOR_ENV];
  const author = typeof raw === 'string' ? stripIllegalXmlChars(raw).trim().slice(0, 128) : '';
  return author !== '' ? author : 'ASTra';
}

/**
 * Author stamps are already deterministic (explicit op author or
 * settings.defaultAuthor, fallback 'ASTra'); no wall-clock feeds them. This
 * helper centralizes that fallback so the deterministic path has one
 * documented author route alongside the date helper.
 */
export function resolveDeterministicAuthor(
  explicit: unknown,
  settingsAuthor: unknown,
): string {
  for (const candidate of [explicit, settingsAuthor]) {
    if (typeof candidate !== 'string') continue;
    const author = stripIllegalXmlChars(candidate);
    if (author.trim() !== '') return author;
  }
  return 'ASTra';
}

/**
 * Detach (and return clones of) prior w:del revision runs in a scope so a
 * chained tracked edit accumulates history instead of wiping it. Only the
 * outermost del in any nesting chain is stashed; inner ones travel with the
 * outer clone. Callers re-attach the clones after emitting the new diff.
 */
function stashPriorDeletions(scope: Element): Element[] {
  const stashed: Element[] = [];
  const dels = Array.from(scope.getElementsByTagNameNS(W_NS, 'del')) as Element[];
  for (const d of dels) {
    let ancestor: Node | null = d.parentNode;
    let nested = false;
    while (ancestor && ancestor !== scope) {
      if (ancestor.nodeType === 1 && (ancestor as Element).tagName === 'w:del') {
        nested = true;
        break;
      }
      ancestor = ancestor.parentNode;
    }
    if (!nested) stashed.push(d.cloneNode(true) as Element);
  }
  return stashed;
}

/** Remove every w:r descendant of a scope (runs only; pPr/bookmarks survive). */
function removeDescendantRuns(scope: Element): void {
  const runs = Array.from(scope.getElementsByTagNameNS(W_NS, 'r')) as Element[];
  for (const r of runs) {
    r.parentNode?.removeChild(r);
  }
}

/**
 * Remove wrapper elements left with no element children. Iterates to a
 * fixpoint so wrappers nested inside other emptied wrappers are collected.
 */
function removeEmptyWrappers(root: Element, localNames: string[]): void {
  let removed = true;
  while (removed) {
    removed = false;
    for (const local of localNames) {
      const els = Array.from(root.getElementsByTagNameNS(W_NS, local)) as Element[];
      for (const el of els) {
        if (el === root) continue;
        let hasElementChild = false;
        for (let i = 0; i < el.childNodes.length; i++) {
          if (el.childNodes[i].nodeType === 1) {
            hasElementChild = true;
            break;
          }
        }
        if (!hasElementChild) {
          el.parentNode?.removeChild(el);
          removed = true;
        }
      }
    }
  }
}

/** Remove direct-child w:hyperlink shells left with no run/revision content. */
function removeEmptyHyperlinkShells(p: Element): void {
  for (let i = p.childNodes.length - 1; i >= 0; i--) {
    const child = p.childNodes[i] as Element;
    if (child.nodeType !== 1 || child.tagName !== 'w:hyperlink') continue;
    let hasElementChild = false;
    for (let j = 0; j < child.childNodes.length; j++) {
      if (child.childNodes[j].nodeType === 1) {
        hasElementChild = true;
        break;
      }
    }
    if (!hasElementChild) p.removeChild(child);
  }
}

/**
 * Snapshot hyperlink shells (attributes only) in a scope before clearing.
 * Each entry is a detached shallow clone: all anchor/history/rel attrs
 * travel with it, no children. Callers re-emit one wrapper after the diff
 * so the LINK survives, not just its text.
 */
function snapshotHyperlinkShells(scope: Element): Element[] {
  const out: Element[] = [];
  const found = Array.from(scope.getElementsByTagNameNS(W_NS, 'hyperlink')) as Element[];
  for (const h of found) {
    // Skip nested duplicates (hyperlinks never nest, guard anyway).
    let ancestor: Node | null = h.parentNode;
    let nested = false;
    while (ancestor && ancestor !== scope) {
      if (ancestor.nodeType === 1 && (ancestor as Element).tagName === 'w:hyperlink') {
        nested = true;
        break;
      }
      ancestor = ancestor.parentNode;
    }
    if (!nested) out.push(h.cloneNode(false) as Element);
  }
  return out;
}

/**
 * Re-emit the hyperlink wrapper around freshly emitted revision runs when
 * the original scope carried a link but clearing left none. The wrapper is
 * a fresh clone of the first snapshot (anchor/history/r:id identical), so
 * the href target is unchanged. Stashed deletion history stays outside the
 * wrapper (non-visible w:delText needs no link); only the new diff runs
 * move inside. No-op when there was no link, when new runs are empty
 * (avoids minting a dead shell), or when a non-empty hyperlink already
 * survives.
 *
 * Known limit: per-word link mapping collapses. A mixed paragraph (linked
 * words + plain words) re-emits a single wrapper from the first link
 * around all new runs; the surviving text is complete and valid, but plain
 * words become linked. Word-level link spans would need run-aligned diff.
 */
function restoreHyperlinkWrappers(scope: Element, snapshots: Element[], newNodes: Element[]): void {
  if (snapshots.length === 0 || newNodes.length === 0) return;
  const existing = Array.from(scope.getElementsByTagNameNS(W_NS, 'hyperlink')) as Element[];
  for (const e of existing) {
    for (let i = 0; i < e.childNodes.length; i++) {
      if (e.childNodes[i].nodeType === 1) return;
    }
  }
  // Drop any leftover empty shells so exactly one live wrapper remains.
  for (const e of existing) {
    e.parentNode?.removeChild(e);
  }
  const wrapper = snapshots[0].cloneNode(false) as Element;
  for (const n of newNodes) {
    wrapper.appendChild(n);
  }
  scope.appendChild(wrapper);
}

/**
 * Re-attach stashed deletion history right after leading pPr/bookmark
 * markers so pPr stays first and the new revision runs keep their order.
 * Deletions carry w:delText (non-visible), so the visible text of the
 * paragraph is unaffected by where they sit.
 */
function reattachHistory(targetP: Element, stashed: Element[]): void {
  if (stashed.length === 0) return;
  let anchor: Node | null = null;
  for (let i = 0; i < targetP.childNodes.length; i++) {
    const c = targetP.childNodes[i] as Element;
    // Skip insignificant whitespace text between elements (pretty-printed
    // parts): it must not terminate the leading-marker scan, or history
    // lands before w:pPr and fails ECMA ordering validation.
    if (c.nodeType !== 1) continue;
    if (c.tagName === 'w:pPr' || c.tagName === 'w:bookmarkStart') {
      anchor = c;
    } else {
      break;
    }
  }
  const ref = anchor ? anchor.nextSibling : targetP.firstChild;
  for (const d of stashed) {
    targetP.insertBefore(d, ref);
  }
}

export function wordDiff(oldText: string, newText: string): DiffChunk[] {
  if (oldText === newText) {
    return oldText.length > 0 ? [{ type: 'equal', text: oldText }] : [];
  }

  const oldWords = oldText.split(/(\s+)/);
  const newWords = newText.split(/(\s+)/);

  // Compute LCS grid
  const m = oldWords.length;
  const n = newWords.length;

  if (m * n > 2_500_000) {
    const fallback: DiffChunk[] = [];
    if (oldText.length > 0) {
      fallback.push({ type: 'delete', text: oldText });
    }
    if (newText.length > 0) {
      fallback.push({ type: 'insert', text: newText });
    }
    return fallback;
  }

  const c = Array(m + 1).fill(0).map(() => Array(n + 1).fill(0));

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (oldWords[i - 1] === newWords[j - 1]) {
        c[i][j] = c[i - 1][j - 1] + 1;
      } else {
        c[i][j] = Math.max(c[i][j - 1], c[i - 1][j]);
      }
    }
  }

  // Backtrack to build chunks
  const chunks: DiffChunk[] = [];
  let i = m, j = n;
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && oldWords[i - 1] === newWords[j - 1]) {
      chunks.unshift({ type: 'equal', text: oldWords[i - 1] });
      i--; j--;
    } else if (j > 0 && (i === 0 || c[i][j - 1] >= c[i - 1][j])) {
      chunks.unshift({ type: 'insert', text: newWords[j - 1] });
      j--;
    } else if (i > 0 && (j === 0 || c[i][j - 1] < c[i - 1][j])) {
      chunks.unshift({ type: 'delete', text: oldWords[i - 1] });
      i--;
    }
  }

  // Merge adjacent chunks of the same type
  const merged: DiffChunk[] = [];
  for (const chunk of chunks) {
    if (merged.length > 0 && merged[merged.length - 1].type === chunk.type) {
      merged[merged.length - 1].text += chunk.text;
    } else {
      merged.push(chunk);
    }
  }

  return merged;
}

export function emitTrackedReplace(
  node: Element,
  oldText: string,
  newText: string,
  author: string,
  date: string,
  idFactory: { getNextId: () => number }
): boolean {
  // No-op edits must not touch the DOM: rebuilding would wipe accumulated
  // revision history (and hyperlink shells) while minting nothing.
  if (oldText === newText) return false;
  const doc = node.ownerDocument!;

  // Accumulate: stash prior deletion history (direct and hyperlink-nested)
  // before clearing. Prior insertions hold visible text that the new diff
  // re-emits, so they are intentionally NOT stashed: re-emitting them as
  // insertions would duplicate the visible text. Their surviving text comes
  // back as plain equal runs (insertion provenance of intermediate steps
  // flattens); every deleted word stays addressable in w:delText.
  const priorDels = stashPriorDeletions(node);
  // Snapshot hyperlink shells before clearing so the LINK can be re-emitted
  // (copied anchor/history attrs) around the new revision runs.
  const linkSnapshots = snapshotHyperlinkShells(node);

  // Find first run to preserve formatting
  let firstRunPr: Element | null = null;
  const runs = Array.from(node.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'r'));
  for (const r of runs) {
    const rPrs = r.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rPr');
    if (rPrs.length > 0) {
      firstRunPr = rPrs.item(0)!.cloneNode(true) as Element;
      break;
    }
  }

  // Clear existing runs (including runs nested in prior w:ins/w:del and
  // w:hyperlink wrappers), then drop the wrappers left empty. Emptied
  // hyperlink shells are dropped here and RE-EMITTED below around the new
  // revision runs (copied attrs): a bare <w:hyperlink/> with no runs is a
  // dead shell with an orphaned r:id, but deleting the shell outright kills
  // the LINK. The link text itself is preserved as revision content
  // (w:del/w:ins) inside the re-emitted wrapper at paragraph level.
  removeDescendantRuns(node);
  removeEmptyWrappers(node, ['ins', 'del']);
  removeEmptyHyperlinkShells(node);

  const chunks = wordDiff(oldText, newText);
  let hasRevisions = false;
  const newNodes: Element[] = [];

  for (const chunk of chunks) {
    if (chunk.text.length === 0) continue;
    if (chunk.type === 'equal') {
      const r = makeRun(doc, chunk.text, firstRunPr);
      node.appendChild(r);
      newNodes.push(r);
    } else if (chunk.type === 'delete') {
      const d = makeDelRun(doc, chunk.text, firstRunPr, author, date, idFactory.getNextId());
      node.appendChild(d);
      newNodes.push(d);
      hasRevisions = true;
    } else if (chunk.type === 'insert') {
      const ins = makeInsRun(doc, chunk.text, firstRunPr, author, date, idFactory.getNextId());
      node.appendChild(ins);
      newNodes.push(ins);
      hasRevisions = true;
    }
  }
  reattachHistory(node, priorDels);
  restoreHyperlinkWrappers(node, linkSnapshots, newNodes);
  return hasRevisions;
}

/**
 * Tracked text edit of a table cell (w:tc).
 *
 * Word semantics note: editing the TEXT inside an existing cell is tracked
 * exactly like a body paragraph — w:ins/w:del runs inside the cell's w:p
 * elements. The structural markers w:cellIns/w:cellDel (CT_TrackChange in
 * w:tcPr) are reserved for whole-cell insertion/deletion (add/remove
 * row/column), not for text edits, so emitting them here would
 * misrepresent the change. This helper is therefore the documented
 * equivalent: paragraph-level revisions scoped to the cell's own
 * paragraphs (never descending into a nested w:tbl).
 *
 * The word-level diff of (oldCellText -> newText) is emitted into the
 * cell's first own paragraph; remaining own paragraphs are cleared to
 * plain empty shells so no stale text survives. The cell keeps at least
 * one w:p, as ECMA-376 requires.
 *
 * Chained edits accumulate: prior w:del runs from any own paragraph are
 * re-attached to the first paragraph, so every deleted word stays
 * addressable in w:delText. Hyperlink shells are re-emitted (copied
 * anchor/history attrs) around the new revision runs in the target
 * paragraph so the LINK survives; mixed linked/plain paragraphs collapse
 * to the first link (see restore helper).
 */
export function emitTrackedCellReplace(
  tc: Element,
  oldText: string,
  newText: string,
  author: string,
  date: string,
  idFactory: { getNextId: () => number }
): boolean {
  // No-op edits must not touch the DOM (see emitTrackedReplace).
  if (oldText === newText) return false;
  const doc = tc.ownerDocument!;
  const ownPs: Element[] = [];
  for (let i = 0; i < tc.childNodes.length; i++) {
    const child = tc.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === 'w:p') {
      ownPs.push(child as Element);
    }
  }
  let targetP = ownPs[0];
  if (!targetP) {
    targetP = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:p');
    tc.appendChild(targetP);
    ownPs.push(targetP);
  }

  // First run properties anywhere in the cell's own paragraphs, for
  // formatting continuity of the emitted revision runs.
  let firstRunPr: Element | null = null;
  for (const p of ownPs) {
    const rPrs = p.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rPr');
    if (rPrs.length > 0) {
      firstRunPr = rPrs.item(0)!.cloneNode(true) as Element;
      break;
    }
  }

  // Accumulate deletion history from every own paragraph (never from a
  // nested table: only direct-child w:p subtrees are in scope).
  const priorDels: Element[] = [];
  for (const p of ownPs) {
    for (const d of stashPriorDeletions(p)) priorDels.push(d);
  }
  // Snapshot hyperlink shells from every own paragraph before clearing.
  const linkSnapshots: Element[] = [];
  for (const p of ownPs) {
    for (const s of snapshotHyperlinkShells(p)) linkSnapshots.push(s);
  }

  // Clear revision-bearing content from every own paragraph: runs (even
  // inside prior w:ins/w:del or w:hyperlink wrappers) plus wrappers left
  // empty thereby. pPr and bookmarks survive; emptied hyperlink shells are
  // dropped here and re-emitted below around the new runs in the target
  // paragraph. Prior insertions are NOT
  // stashed: they hold visible text the new diff re-emits, and keeping
  // them would duplicate it.
  for (const p of ownPs) {
    removeDescendantRuns(p);
    removeEmptyWrappers(p, ['ins', 'del']);
    removeEmptyHyperlinkShells(p);
  }

  const chunks = wordDiff(oldText, newText);
  let hasRevisions = false;
  const newNodes: Element[] = [];
  for (const chunk of chunks) {
    if (chunk.text.length === 0) continue;
    if (chunk.type === 'equal') {
      const r = makeRun(doc, chunk.text, firstRunPr);
      targetP.appendChild(r);
      newNodes.push(r);
    } else if (chunk.type === 'delete') {
      const d = makeDelRun(doc, chunk.text, firstRunPr, author, date, idFactory.getNextId());
      targetP.appendChild(d);
      newNodes.push(d);
      hasRevisions = true;
    } else if (chunk.type === 'insert') {
      const ins = makeInsRun(doc, chunk.text, firstRunPr, author, date, idFactory.getNextId());
      targetP.appendChild(ins);
      newNodes.push(ins);
      hasRevisions = true;
    }
  }
  reattachHistory(targetP, priorDels);
  restoreHyperlinkWrappers(targetP, linkSnapshots, newNodes);
  return hasRevisions;
}

function makeRun(doc: Document, text: string, rPr: Element | null): Element {
  const r = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
  if (rPr) r.appendChild(rPr.cloneNode(true));
  const t = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:t');
  if (hasBoundarySpace(text)) {
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  }
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

/**
 * Significant boundary whitespace needs xml:space="preserve" or renderers
 * drop it: a tab (or other non-space whitespace) at a run edge without the
 * marker collapses in print/PDF output, jamming the words it separates.
 * Plain spaces are the common case; the check covers every whitespace kind
 * so adjacent revision runs keep their boundary in visible text.
 */
function hasBoundarySpace(text: string): boolean {
  if (text.length === 0) return false;
  return /\s/.test(text[0]) || /\s/.test(text[text.length - 1]);
}

function makeDelRun(doc: Document, text: string, rPr: Element | null, author: string, date: string, id: number): Element {
  const del = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:del');
  del.setAttribute('w:id', String(id));
  del.setAttribute('w:author', author);
  del.setAttribute('w:date', date);

  const r = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
  if (rPr) r.appendChild(rPr.cloneNode(true));
  const delText = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:delText');
  if (hasBoundarySpace(text)) {
    delText.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  }
  delText.appendChild(doc.createTextNode(text));
  r.appendChild(delText);
  del.appendChild(r);
  return del;
}

function makeInsRun(doc: Document, text: string, rPr: Element | null, author: string, date: string, id: number): Element {
  const ins = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:ins');
  ins.setAttribute('w:id', String(id));
  ins.setAttribute('w:author', author);
  ins.setAttribute('w:date', date);

  ins.appendChild(makeRun(doc, text, rPr));
  return ins;
}
