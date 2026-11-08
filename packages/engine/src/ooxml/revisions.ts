import { REVISION_TAGS } from './revision-ids';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REVISION_NAMES = new Set<string>(REVISION_TAGS);

export interface UnresolvedRevision {
  kind: string;
  count: number;
  reason: string;
  part: string;
  nodeId?: string;
}

export interface RevisionResolution {
  resolved: number;
  resolvedIds: number[];
  byKind: Record<string, number>;
  unresolved: UnresolvedRevision[];
  removedNodeIds?: string[];
  idsMayHaveChanged?: boolean;
  /** Bookmarks removed together with a paragraph whose deletion or move was resolved. */
  removedBookmarks?: string[];
  /** Comment ids whose anchors were removed with such a paragraph; the caller deletes the comments themselves. */
  removedComments?: number[];
  /** Links that still point at a removed bookmark, as Word leaves a stale table-of-contents entry. */
  staleLinks?: { bookmark: string; count: number }[];
}

export interface RevisionSelector { ids?: number[]; author?: string; nodeId?: string }
export interface RevisionEntry {
  id: number | null;
  kind: string;
  author: string;
  date: string;
  part: string;
  nodeId?: string;
  group?: string;
  text: string;
  textLength: number;
}

export function listRevisions(parts: readonly {part: string; document: Document; nodeIdOf?: (element: Element) => string | undefined}[], options: {offset?: number; limit?: number} = {}): {revisions: RevisionEntry[]; total: number; nextOffset?: number} {
  const all: RevisionEntry[] = [];
  for (const {part, document, nodeIdOf} of parts) {
    const moves = findMoveGroups(document);
    for (const element of Array.from(document.getElementsByTagName('*'))) {
      if (element.namespaceURI !== W_NS || !REVISION_NAMES.has(element.localName)) continue;
      const value = element.getAttributeNS(W_NS, 'id');
      const fullText = element.textContent ?? '';
      const entry: RevisionEntry = {
        id: value && /^\d+$/.test(value) ? Number(value) : null,
        kind: element.localName,
        author: element.getAttributeNS(W_NS, 'author') ?? '',
        date: element.getAttributeNS(W_NS, 'date') ?? '',
        part,
        text: fullText.slice(0,60),
        textLength: fullText.length,
      };
      const nodeId = nodeIdOf?.(element);
      if (nodeId) entry.nodeId = nodeId;
      const group = moves.get(element);
      if (group) entry.group = group.name;
      all.push(entry);
    }
  }
  const offset = Math.max(0, options.offset ?? 0);
  const limit = Math.max(0, Math.min(50, options.limit ?? 50));
  const nextOffset = offset + limit < all.length ? offset + limit : undefined;
  return {revisions:all.slice(offset,offset+limit), total:all.length, ...(nextOffset !== undefined ? {nextOffset} : {})};
}

export function readRevision(parts: readonly {part: string; document: Document; nodeIdOf?: (element: Element) => string | undefined}[], id: number): {group: (RevisionEntry & {fullText: string})[]; ambiguous: boolean} {
  const group: (RevisionEntry & {fullText: string})[] = [];
  const moveKeys = new Set<string>();
  for (const part of parts) {
    const moves = findMoveGroups(part.document);
    for (const element of Array.from(part.document.getElementsByTagName('*'))) {
      if (element.namespaceURI !== W_NS || !REVISION_NAMES.has(element.localName) || element.getAttributeNS(W_NS, 'id') !== String(id)) continue;
      const fullText = element.textContent ?? '';
      const move = moves.get(element);
      group.push({
        id, kind:element.localName,
        author:element.getAttributeNS(W_NS, 'author') ?? '',
        date:element.getAttributeNS(W_NS, 'date') ?? '',
        part:part.part,
        ...(part.nodeIdOf?.(element) ? {nodeId:part.nodeIdOf(element)} : {}),
        ...(move ? {group:move.name} : {}),
        text:fullText.slice(0,60), textLength:fullText.length, fullText,
      });
      if (move) moveKeys.add(`${part.part}:${move.name}`);
    }
  }
  return {group, ambiguous:group.length > 1 && (
    group.length !== 2 || moveKeys.size !== 1 ||
    !group.some(entry => entry.kind === 'moveFrom') || !group.some(entry => entry.kind === 'moveTo')
  )};
}

/** A selected side of a move always brings its counterpart into the transaction. */
export function expandMoveSelection(document: Document, selected: Set<Element>): void {
  const groups = findMoveGroups(document);
  for (const element of [...selected]) {
    const group = groups.get(element);
    if (group) for (const member of group.members) selected.add(member);
  }
}

const children = (element: Element): Element[] =>
  Array.from(element.childNodes).filter((node): node is Element => node.nodeType === 1);
const local = (node: Node | null): string => node && node.nodeType === 1 ? (node as Element).localName : '';
/** The revision mark of a paragraph: w:ins, w:del, w:moveFrom or w:moveTo in w:pPr/w:rPr. */
const isParagraphMark = (el: Element): boolean => local(el.parentNode) === 'rPr' && local(el.parentNode!.parentNode) === 'pPr';

function unwrap(element: Element): void {
  const parent = element.parentNode;
  if (!parent) return;
  while (element.firstChild) parent.insertBefore(element.firstChild, element);
  parent.removeChild(element);
}

function restoreDeletedText(element: Element): void {
  for (const [oldName, newName] of [['delText', 't'], ['delInstrText', 'instrText']]) {
    for (const old of Array.from(element.getElementsByTagNameNS(W_NS, oldName))) {
      const replacement = element.ownerDocument!.createElementNS(W_NS, `w:${newName}`);
      for (const attr of Array.from(old.attributes)) replacement.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
      while (old.firstChild) replacement.appendChild(old.firstChild);
      old.parentNode!.replaceChild(replacement, old);
    }
  }
}

function mergeParagraphMark(marker: Element): boolean {
  const paragraph = marker.parentNode?.parentNode?.parentNode as Element | undefined;
  if (!paragraph || local(paragraph) !== 'p') return false;
  let sibling: Node | null = paragraph.nextSibling;
  while (sibling && sibling.nodeType !== 1) sibling = sibling.nextSibling;
  if (!sibling || local(sibling) !== 'p') return false;
  const next = sibling as Element;
  const properties = children(paragraph).find(el => el.localName === 'pPr');
  if (properties && children(properties).some(el => el.localName === 'sectPr')) return false;
  const zeroWidth = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'commentReference', 'proofErr']);
  const plain = (p: Element): boolean => children(p).every(el => {
    if (el.localName === 'pPr' || zeroWidth.has(el.localName)) return true;
    return el.localName === 'r' && children(el).every(child => ['rPr', 't', 'tab', 'br', 'lastRenderedPageBreak', 'noBreakHyphen', 'softHyphen'].includes(child.localName));
  });
  if (!plain(paragraph) || !plain(next)) return false;
  // A paragraph with no text left is a whole-paragraph deletion. The block pass owns it, because Word
  // removes its bookmarks and comments instead of carrying them into the next paragraph.
  const live = children(paragraph).some(el => el.localName !== 'pPr' && !zeroWidth.has(el.localName));
  if (!live && children(paragraph).some(el => el.localName !== 'pPr')) return false;
  const nextPr = children(next).find(el => el.localName === 'pPr');
  const anchor = nextPr ? nextPr.nextSibling : next.firstChild;
  for (const child of children(paragraph)) if (child.localName !== 'pPr') next.insertBefore(child, anchor);
  paragraph.parentNode!.removeChild(paragraph);
  return true;
}

interface MoveGroup { name: string; members: Element[]; marks: Element[]; markers: Element[]; valid: boolean }
function findMoveGroups(doc: Document): Map<Element, MoveGroup> {
  const groups = new Map<string, MoveGroup>();
  const open = new Map<string, { name: string; start: Element }>();
  const associated = new Map<Element, MoveGroup>();
  for (const el of Array.from(doc.getElementsByTagName('*'))) {
    if (el.namespaceURI !== W_NS) continue;
    const tag = el.localName;
    if (tag === 'moveFromRangeStart' || tag === 'moveToRangeStart') {
      const side = tag === 'moveFromRangeStart' ? 'from' : 'to';
      const id = el.getAttributeNS(W_NS, 'id');
      const name = el.getAttributeNS(W_NS, 'name');
      if (!id || !name) continue;
      const group = groups.get(name) ?? { name, members: [], marks: [], markers: [], valid: true };
      groups.set(name, group);
      const key = `${side}:${id}`;
      if (open.has(key)) group.valid = false;
      open.set(key, { name, start: el });
      group.markers.push(el);
    } else if (tag === 'moveFromRangeEnd' || tag === 'moveToRangeEnd') {
      const side = tag === 'moveFromRangeEnd' ? 'from' : 'to';
      const key = `${side}:${el.getAttributeNS(W_NS, 'id')}`;
      const range = open.get(key);
      if (!range) continue;
      groups.get(range.name)!.markers.push(el);
      open.delete(key);
    } else if ((tag === 'moveFrom' || tag === 'moveTo') && isParagraphMark(el)) {
      // The mark of a moved paragraph comes before the range start inside the same paragraph.
      const paragraph = el.parentNode!.parentNode!.parentNode as Element;
      const startTag = tag === 'moveFrom' ? 'moveFromRangeStart' : 'moveToRangeStart';
      const own = children(paragraph).filter(child => child.localName === startTag && child.namespaceURI === W_NS);
      const sidePrefix = tag === 'moveFrom' ? 'from:' : 'to:';
      const open1 = [...open.entries()].filter(([key]) => key.startsWith(sidePrefix));
      const name = own.length === 1 ? own[0].getAttributeNS(W_NS, 'name') : open1.length === 1 ? open1[0][1].name : null;
      if (!name) continue;
      const group = groups.get(name) ?? { name, members: [], marks: [], markers: [], valid: true };
      groups.set(name, group);
      group.members.push(el);
      group.marks.push(el);
      associated.set(el, group);
    } else if (tag === 'moveFrom' || tag === 'moveTo') {
      const side = tag === 'moveFrom' ? 'from:' : 'to:';
      const ranges = [...open.entries()].filter(([key]) => key.startsWith(side));
      if (ranges.length !== 1) continue;
      const group = groups.get(ranges[0][1].name)!;
      group.members.push(el);
      associated.set(el, group);
    }
  }
  for (const range of open.values()) groups.get(range.name)!.valid = false;
  for (const group of groups.values()) {
    const from = group.members.filter(el => el.localName === 'moveFrom').length;
    const to = group.members.filter(el => el.localName === 'moveTo').length;
    if (!from || !to || group.markers.length !== 4) group.valid = false;
    // Avoid resolving one side before nested revision semantics are known.
    if (group.members.some(member => Array.from(member.getElementsByTagName('*')).some(child => child.namespaceURI === W_NS && REVISION_NAMES.has(child.localName)))) group.valid = false;
  }
  return associated;
}

const PROPERTY_KEEP: Record<string, readonly string[]> = {
  rPr: ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'],
  pPr: ['rPr', 'sectPr', 'pPrChange'],
  sectPr: ['headerReference', 'footerReference', 'sectPrChange'],
  tblPr: ['tblPrChange'],
  tblPrEx: ['tblPrExChange'],
  trPr: ['ins', 'del', 'trPrChange'],
  tcPr: ['cellIns', 'cellDel', 'cellMerge', 'tcPrChange'],
};

/** Count revision elements, including unsupported kinds, without changing a document. */
export function countRevisions(documents: readonly Document[]): number {
  let count = 0;
  for (const doc of documents) {
    for (const tag of REVISION_TAGS) count += doc.getElementsByTagNameNS(W_NS, tag).length;
  }
  return count;
}


/** Elements that take no space in a paragraph. */
const ZERO_WIDTH = new Set([
  'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'proofErr',
  'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'permStart', 'permEnd',
]);

/** True when every child of the paragraph is its properties, a zero-width marker, a run that only carries a comment reference, or one of `wrappers`. */
function onlyWrapped(paragraph: Element, wrappers: readonly string[]): boolean {
  return children(paragraph).every(child => {
    if (child.namespaceURI !== W_NS) return false;
    if (child.localName === 'pPr' || ZERO_WIDTH.has(child.localName) || wrappers.includes(child.localName)) return true;
    return child.localName === 'r' && children(child).every(part => part.localName === 'rPr' || part.localName === 'commentReference');
  });
}

type Anchors =
  | { ok: true; bookmarks: string[]; comments: number[]; spill: { element: Element; block: Element; at: 'start' | 'end' }[] }
  | { ok: false; reason: string };

/**
 * The bookmarks and comment threads anchored in blocks that are about to be removed. Following Word, those
 * that lie wholly inside go with them. A comment that starts, ends or is referenced outside makes the caller
 * keep the blocks and report it. A bookmark that ends outside is refused unless `allowSpill` is set, which
 * is the case for a moved paragraph: Word then carries the start to the next paragraph and the end to the
 * previous one.
 */
function inspectAnchors(blocks: readonly Element[], allowSpill: boolean): Anchors {
  const idOf = (el: Element): string => el.getAttributeNS(W_NS, 'id') ?? '';
  const collect = (tag: string): { el: Element; block: Element }[] =>
    blocks.flatMap(block => Array.from(block.getElementsByTagNameNS(W_NS, tag)).map(el => ({ el, block })));
  const bookmarkStarts = collect('bookmarkStart');
  const bookmarkEnds = collect('bookmarkEnd');
  const startIds = new Set(bookmarkStarts.map(x => idOf(x.el)));
  const endIds = new Set(bookmarkEnds.map(x => idOf(x.el)));
  const spill: { element: Element; block: Element; at: 'start' | 'end' }[] = [];
  const whole: string[] = [];
  for (const { el, block } of bookmarkStarts) {
    if (endIds.has(idOf(el))) whole.push(el.getAttributeNS(W_NS, 'name') ?? '');
    else if (allowSpill) spill.push({ element: el, block, at: 'start' });
    else return { ok: false, reason: 'a bookmark starts or ends outside the paragraph that would be removed' };
  }
  for (const { el, block } of bookmarkEnds) {
    if (startIds.has(idOf(el))) continue;
    if (allowSpill) spill.push({ element: el, block, at: 'end' });
    else return { ok: false, reason: 'a bookmark starts or ends outside the paragraph that would be removed' };
  }
  const cs = new Set(collect('commentRangeStart').map(x => idOf(x.el)));
  const ce = new Set(collect('commentRangeEnd').map(x => idOf(x.el)));
  const cr = new Set(collect('commentReference').map(x => idOf(x.el)));
  const comments = new Set([...cs, ...ce, ...cr]);
  for (const id of comments) if (!cs.has(id) || !ce.has(id) || !cr.has(id) || !/^\d+$/.test(id)) return { ok: false, reason: 'a comment starts, ends or is referenced outside the paragraph that would be removed' };
  return { ok: true, bookmarks: whole, comments: [...comments].map(Number), spill };
}

/** A block that may be removed from its parent without leaving the parent without content. */
function removable(block: Element): string | null {
  const parent = block.parentNode as Element | null;
  if (!parent) return 'the paragraph has no parent';
  const properties = children(block).find(el => el.localName === 'pPr');
  if (properties && children(properties).some(el => el.localName === 'sectPr')) return 'the paragraph carries a section break';
  const siblings = children(parent).filter(el => el !== block && ['p', 'tbl', 'sdt'].includes(el.localName));
  if (!siblings.length && ['tc', 'body'].includes(local(parent))) return 'it is the only block of its container';
  return null;
}

/** Remove a paragraph mark element and the property containers it leaves empty. */
function dropMark(mark: Element): void {
  const rPr = mark.parentNode as Element;
  rPr.removeChild(mark);
  if (!children(rPr).length && rPr.parentNode) {
    const pPr = rPr.parentNode as Element;
    pPr.removeChild(rPr);
    if (!children(pPr).length && pPr.parentNode) pPr.parentNode.removeChild(pPr);
  }
}

interface BlockContext {
  mode: 'accept' | 'reject';
  options: { selected?: ReadonlySet<Element>; commentsOk?: (ids: number[]) => boolean };
  /** Revision elements the generic pass must leave alone, with the reason. */
  refused: Map<Element, string>;
  noteResolved: (el: Element) => void;
  bookmarks: Set<string>;
  comments: Set<number>;
}

const revisionElementsIn = (root: Element): Element[] =>
  Array.from(root.getElementsByTagName('*')).filter(el => el.namespaceURI === W_NS && REVISION_NAMES.has(el.localName));

const allSelected = (ctx: BlockContext, elements: readonly Element[]): boolean =>
  !ctx.options.selected || elements.every(el => ctx.options.selected!.has(el));

function refuse(ctx: BlockContext, elements: readonly Element[], reason: string): void {
  for (const el of elements) ctx.refused.set(el, reason);
}

/** The nearest paragraph before or after a block, skipping markers that go with it and blocks that are removed too. */
function neighbourParagraph(block: Element, step: 'previousSibling' | 'nextSibling', gone: ReadonlySet<Element>): Element | null {
  for (let n: Node | null = block[step]; n; n = n[step]) {
    if (n.nodeType !== 1) continue;
    const el = n as Element;
    if (gone.has(el)) continue;
    return local(el) === 'p' ? el : null;
  }
  return null;
}

/** Check anchors and removability for blocks that will be removed. Returns a reason, or null when they may go. */
function checkRemoval(ctx: BlockContext, blocks: readonly Element[], moved?: ReadonlySet<Element>): string | null {
  for (const block of blocks) {
    const why = removable(block);
    if (why) return why;
  }
  const anchors = inspectAnchors(blocks, !!moved);
  if (!anchors.ok) return anchors.reason;
  const gone = new Set<Element>([...blocks, ...(moved ?? [])]);
  for (const item of anchors.spill) {
    if (!neighbourParagraph(item.block, item.at === 'start' ? 'nextSibling' : 'previousSibling', gone))
      return 'a bookmark that ends outside the moved paragraph has no neighbouring paragraph to take it';
  }
  if (anchors.comments.length && ctx.options.commentsOk && !ctx.options.commentsOk(anchors.comments)) return 'a comment thread anchored here is also anchored outside the paragraph that would be removed';
  return null;
}

function commitRemoval(ctx: BlockContext, blocks: readonly Element[], moved?: ReadonlySet<Element>): void {
  const anchors = inspectAnchors(blocks, !!moved);
  const gone = new Set<Element>([...blocks, ...(moved ?? [])]);
  if (anchors.ok) {
    for (const name of anchors.bookmarks) ctx.bookmarks.add(name);
    for (const id of anchors.comments) ctx.comments.add(id);
    for (const item of anchors.spill) {
      const target = neighbourParagraph(item.block, item.at === 'start' ? 'nextSibling' : 'previousSibling', gone);
      if (!target) continue;
      if (item.at === 'start') target.insertBefore(item.element, children(target).find(el => el.localName !== 'pPr') ?? null);
      else target.appendChild(item.element);
    }
  }
  for (const block of blocks) block.parentNode?.removeChild(block);
}

/** Reject of a table of contents (or any block) that Word inserted inside a pair of customXmlIns range markers. */
function resolveInsertedSdt(doc: Document, ctx: BlockContext): void {
  if (ctx.mode !== 'reject') return;
  for (const sdt of Array.from(doc.getElementsByTagNameNS(W_NS, 'sdt'))) {
    if (!sdt.parentNode) continue;
    const siblingElement = (node: Node | null, step: 'previousSibling' | 'nextSibling'): Element | null => {
      for (let n = node?.[step] ?? null; n; n = n[step]) if (n.nodeType === 1) return n as Element;
      return null;
    };
    const before = siblingElement(sdt, 'previousSibling');
    const after = siblingElement(sdt, 'nextSibling');
    const content = children(sdt).find(el => el.localName === 'sdtContent');
    if (!before || !after || !content || before.localName !== 'customXmlInsRangeStart' || after.localName !== 'customXmlInsRangeEnd') continue;
    const inside = children(content);
    const first = inside[0];
    const last = inside[inside.length - 1];
    const id = (el: Element) => el.getAttributeNS(W_NS, 'id');
    if (!first || !last || first.localName !== 'customXmlInsRangeEnd' || last.localName !== 'customXmlInsRangeStart' || id(first) !== id(before) || id(last) !== id(after)) continue;
    const revisions = revisionElementsIn(sdt);
    const all = [before, after, ...revisions];
    if (!allSelected(ctx, all)) continue;
    const reason = revisions.some(el => !['ins', 'customXmlInsRangeStart', 'customXmlInsRangeEnd'].includes(el.localName))
      ? 'the inserted block holds other revisions'
      : checkRemoval(ctx, [sdt]);
    if (reason) { refuse(ctx, all, reason); continue; }
    commitRemoval(ctx, [sdt]);
    before.parentNode?.removeChild(before);
    after.parentNode?.removeChild(after);
    for (const el of all) ctx.noteResolved(el);
  }
}

/** Whole-paragraph moves: the destination stays once, the source goes, every mark and range marker is removed. */
function resolveParagraphMoves(doc: Document, groups: Map<Element, MoveGroup>, ctx: BlockContext): void {
  for (const group of new Set(groups.values())) {
    if (!group.marks.length || !group.valid) continue;
    if (!allSelected(ctx, group.members)) continue;
    const sideOf = (el: Element) => el.localName;
    const marked = new Map<Element, Element>();
    for (const mark of group.marks) marked.set(mark.parentNode!.parentNode!.parentNode as Element, mark);
    const wrappers = group.members.filter(el => !group.marks.includes(el));
    const problem = wrappers.some(w => !marked.has(w.parentNode as Element) || sideOf(marked.get(w.parentNode as Element)!) !== sideOf(w))
      || [...marked].some(([paragraph, mark]) => !onlyWrapped(paragraph, [sideOf(mark)]))
      ? 'only part of the paragraphs is moved (a partial move)' : null;
    if (problem) { refuse(ctx, group.members, problem); continue; }
    const keep = ctx.mode === 'accept' ? 'moveTo' : 'moveFrom';
    const drop = [...marked].filter(([, mark]) => sideOf(mark) !== keep).map(([paragraph]) => paragraph);
    const goneMarkers = new Set(group.markers);
    const reason = checkRemoval(ctx, drop, goneMarkers);
    if (reason) { refuse(ctx, group.members, reason); continue; }
    for (const [paragraph, mark] of marked) {
      if (sideOf(mark) !== keep) continue;
      dropMark(mark);
      for (const wrapper of wrappers.filter(w => w.parentNode === paragraph)) {
        if (keep === 'moveFrom') restoreDeletedText(wrapper);
        unwrap(wrapper);
      }
    }
    commitRemoval(ctx, drop, goneMarkers);
    for (const marker of group.markers) marker.parentNode?.removeChild(marker);
    for (const el of group.members) ctx.noteResolved(el);
  }
}

/** A paragraph whose mark and whole content are one deletion (accept) or one insertion (reject) is removed as a unit. */
function resolveWholeParagraphs(doc: Document, ctx: BlockContext): void {
  const kind = ctx.mode === 'accept' ? 'del' : 'ins';
  for (const mark of revisionElementsIn(doc.documentElement).filter(el => el.localName === kind && isParagraphMark(el))) {
    const paragraph = mark.parentNode!.parentNode!.parentNode as Element;
    if (local(paragraph) !== 'p' || !paragraph.parentNode) continue;
    if (!onlyWrapped(paragraph, [kind])) continue;
    const revisions = revisionElementsIn(paragraph);
    if (!allSelected(ctx, revisions)) continue;
    const reason = revisions.some(el => !['ins', 'del'].includes(el.localName) || (el.localName !== kind && el !== mark && !isInside(el, kind)))
      ? 'the paragraph holds other revisions' : checkRemoval(ctx, [paragraph]);
    if (reason) { refuse(ctx, revisions, reason); continue; }
    commitRemoval(ctx, [paragraph]);
    for (const el of revisions) ctx.noteResolved(el);
  }
}

/** True when `el` sits below a w:<kind> wrapper (an insertion inside a deletion). */
function isInside(el: Element, kind: string): boolean {
  for (let n: Node | null = el.parentNode; n; n = n.parentNode) if (local(n) === kind) return true;
  return false;
}

/** Resolve supported revisions in one part. Unhandled structural revisions remain reported. */
export function resolveRevisions(documents: readonly Document[], mode: 'accept' | 'reject', options: { removeDeletedCells?: boolean; part?: string; selected?: ReadonlySet<Element>; nodeIdOf?: (element: Element) => string | undefined; commentsOk?: (ids: number[]) => boolean } = {}): RevisionResolution {
  const result: RevisionResolution = { resolved: 0, resolvedIds: [], byKind: {}, unresolved: [] };
  const noteResolved = (el: Element): void => {
    result.resolved++;
    result.byKind[el.localName] = (result.byKind[el.localName] ?? 0) + 1;
    const value = el.getAttributeNS(W_NS, 'id');
    if (value && /^\d+$/.test(value)) result.resolvedIds.push(Number(value));
  };
  for (const doc of documents) {
    const pending = new Set<Element>();
    const containsPending = (parent: Element): boolean => {
      for (const item of pending) for (let node: Node | null = item.parentNode; node; node = node.parentNode) if (node === parent) return true;
      if (options.selected) for (const child of Array.from(parent.getElementsByTagName('*'))) {
        if (child.namespaceURI === W_NS && REVISION_NAMES.has(child.localName) && !options.selected.has(child)) return true;
      }
      return false;
    };
    // A marker pass can detach whole subtrees that still hold revision elements; those are already gone.
    const isAttached = (element: Element): boolean => {
      let node: Node | null = element;
      while (node.parentNode) node = node.parentNode;
      return node === doc;
    };
    const isUnselectedRevision = (element: Element): boolean =>
      !!options.selected && element.namespaceURI === W_NS && REVISION_NAMES.has(element.localName) && !options.selected.has(element);
    const matchingRangeStart = (end: Element): Element | null => {
      const wanted = end.localName.replace('RangeEnd', 'RangeStart');
      const id = end.getAttributeNS(W_NS, 'id');
      for (let node: Node | null = end.previousSibling; node; node = node.previousSibling) {
        if (node.nodeType === 1 && (node as Element).namespaceURI === W_NS && (node as Element).localName === wanted && (node as Element).getAttributeNS(W_NS, 'id') === id) return node as Element;
      }
      return null;
    };
    const everyBetween = (start: Node, end: Node, test: (node: Node) => boolean): boolean => {
      for (let node: Node | null = start.nextSibling; node && node !== end; node = node.nextSibling) if (!test(node)) return false;
      return true;
    };
    const moveGroups = findMoveGroups(doc);
    // Block-level pass: whole paragraphs, whole moves and inserted blocks go as a unit or not at all.
    const refused = new Map<Element, string>();
    const blockContext: BlockContext = { mode, options, refused, noteResolved, bookmarks: new Set(), comments: new Set() };
    resolveInsertedSdt(doc, blockContext);
    resolveParagraphMoves(doc, moveGroups, blockContext);
    resolveWholeParagraphs(doc, blockContext);
    for (const name of blockContext.bookmarks) (result.removedBookmarks ??= []).push(name);
    for (const id of blockContext.comments) (result.removedComments ??= []).push(id);
    if (blockContext.bookmarks.size) {
      const stale = new Map<string, number>();
      for (const name of blockContext.bookmarks) {
        let count = 0;
        for (const link of Array.from(doc.getElementsByTagNameNS(W_NS, 'hyperlink'))) if (link.getAttributeNS(W_NS, 'anchor') === name) count++;
        for (const instr of Array.from(doc.getElementsByTagNameNS(W_NS, 'instrText'))) if (new RegExp(`\\b(?:PAGEREF|REF|HYPERLINK\\s+\\\\l)\\s+"?${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`).test(instr.textContent ?? '')) count++;
        if (count) stale.set(name, count);
      }
      if (stale.size) result.staleLinks = [...stale].map(([bookmark, count]) => ({ bookmark, count }));
    }
    const elements = Array.from(doc.getElementsByTagName('*')).filter(el => el.namespaceURI === W_NS && REVISION_NAMES.has(el.localName));
    for (const el of elements.reverse()) {
      if (options.selected && !options.selected.has(el)) continue;
      if (!el.parentNode || !isAttached(el)) continue;
      const kind = el.localName;
      const parent = el.parentNode as Element;
      const parentKind = local(parent);
      const grandparentKind = local(parent.parentNode);
      let insideUnsupportedMove = false;
      for (let ancestor: Node | null = el.parentNode; ancestor; ancestor = ancestor.parentNode) {
        if (ancestor.nodeType !== 1) continue;
        const element = ancestor as Element;
        if (element.namespaceURI === W_NS && (element.localName === 'moveFrom' || element.localName === 'moveTo')) {
          const group = moveGroups.get(element);
          if (!group || !group.valid) insideUnsupportedMove = true;
        }
      }
      let supported = true;
      let reason = 'revision kind or context is not supported';
      if (refused.has(el)) {
        reason = refused.get(el)!;
        supported = false;
      } else if (insideUnsupportedMove) {
        reason = 'the containing move group is incomplete or contains nested revisions';
        supported = false;
      } else if ((kind === 'moveFrom' || kind === 'moveTo') && isParagraphMark(el)) {
        // Handled with its whole group by the block pass. Still attached means that pass left it alone.
        reason = moveGroups.get(el)?.valid ? 'the move is only partly selected' : 'move ranges are incomplete, ambiguous or contain nested revisions';
        supported = false;
      } else if (kind === 'moveFrom' || kind === 'moveTo') {
        const group = moveGroups.get(el);
        if (!group || !group.valid) {
          reason = 'move ranges are incomplete, ambiguous or contain nested revisions';
          supported = false;
        }
        else {
          if ((kind === 'moveTo') === (mode === 'accept')) {
            if (kind === 'moveFrom') restoreDeletedText(el);
            unwrap(el);
          } else if (containsPending(el)) supported = false;
          else parent.removeChild(el);
          if (supported && group.members.every(member => !member.parentNode)) {
            for (const marker of group.markers) marker.parentNode?.removeChild(marker);
          }
        }
      } else if (kind === 'ins' || kind === 'del') {
        if (parentKind === 'rPr' && grandparentKind === 'pPr') {
          const removeMarkOnly = (kind === 'ins') === (mode === 'accept');
          if (removeMarkOnly) parent.removeChild(el);
          else {
            supported = mergeParagraphMark(el);
            if (!supported) reason = 'paragraph mark cannot be merged safely with the next block';
          }
        } else if (parentKind === 'trPr') {
          const keepRow = (kind === 'ins') === (mode === 'accept');
          if (keepRow) parent.removeChild(el);
          else {
            const row = parent.parentNode as Element;
            const table = row?.parentNode as Element;
            if (local(row) !== 'tr' || local(table) !== 'tbl') supported = false;
            else if (containsPending(row)) supported = false;
            else if (children(table).filter(child => child.localName === 'tr').length > 1) table.removeChild(row);
            else if (local(table.parentNode) === 'tc' && children(table.parentNode as Element).filter(child => ['p', 'tbl'].includes(child.localName)).length === 1) supported = false;
            else table.parentNode?.removeChild(table);
          }
        } else if (kind === 'ins' && mode === 'accept') {
          unwrap(el);
        } else if (kind === 'del' && mode === 'reject') {
          restoreDeletedText(el);
          unwrap(el);
        } else {
          if (containsPending(el)) supported = false;
          else parent.removeChild(el);
        }
      } else if ((kind === 'cellIns' || kind === 'cellMerge') && mode === 'accept' && parentKind === 'tcPr') {
        parent.removeChild(el);
      } else if (kind === 'cellDel' && mode === 'reject' && parentKind === 'tcPr') {
        parent.removeChild(el);
      } else if (kind === 'cellDel' && mode === 'accept' && parentKind === 'tcPr' && options.removeDeletedCells) {
        const cell = parent.parentNode as Element;
        if (local(cell) === 'tc' && cell.parentNode && !containsPending(cell)) cell.parentNode.removeChild(cell);
        else supported = false;
      } else if (kind === 'numberingChange' && mode === 'accept') {
        parent.removeChild(el);
      } else if (kind === 'tblGridChange' && parentKind === 'tblGrid') {
        if (mode === 'reject') {
          const old = children(el)[0];
          if (!old || old.localName !== 'tblGrid') supported = false;
          else {
            for (const child of children(parent)) if (child.localName === 'gridCol') parent.removeChild(child);
            for (const child of children(old)) parent.insertBefore(child.cloneNode(true), el);
          }
        }
        if (supported) parent.removeChild(el);
      } else if (kind === `${parentKind}Change` && PROPERTY_KEEP[parentKind]) {
        if (mode === 'accept') parent.removeChild(el);
        else {
          const old = children(el)[0];
          if (!old || old.localName !== parentKind || old.namespaceURI !== W_NS) {
            supported = false;
            reason = 'previous revision properties are missing or have the wrong type';
          } else {
            for (const child of children(parent)) if (!PROPERTY_KEEP[parentKind].includes(child.localName)) parent.removeChild(child);
            const ref = ['pPr', 'trPr', 'tcPr'].includes(parentKind) ? parent.firstChild : el;
            for (const child of children(old)) parent.insertBefore(child.cloneNode(true), ref);
            parent.removeChild(el);
          }
        }
      } else if (/^customXml(Ins|MoveTo)Range(Start|End)$/.test(kind) && mode === 'accept') {
        // The markers only delimit content that stays. Drop them and keep the content.
        parent.removeChild(el);
      } else if (kind === 'customXmlInsRangeEnd' && mode === 'reject') {
        const start = matchingRangeStart(el);
        if (!start) {
          reason = 'the matching start marker is missing or in another container';
          supported = false;
        } else {
          const between: Node[] = [];
          for (let node: Node | null = start.nextSibling; node && node !== el; node = node.nextSibling) between.push(node);
          if (between.some(node => node.nodeType === 1 && (pending.has(node as Element) || containsPending(node as Element) || isUnselectedRevision(node as Element)))) {
            reason = 'the inserted content holds revisions that cannot be resolved';
            supported = false;
          } else {
            for (const node of between) parent.removeChild(node);
            parent.removeChild(start);
            parent.removeChild(el);
          }
        }
      } else if (kind === 'customXmlInsRangeStart' && mode === 'reject') {
        // Resolved together with its end marker, which is visited first.
        reason = 'the matching end marker is missing or in another container';
        supported = false;
      } else if (/^customXml(Del|MoveFrom)RangeEnd$/.test(kind)) {
        const start = matchingRangeStart(el);
        const empty = start && everyBetween(start, el, node => node.nodeType !== 1 && !(node.textContent ?? '').trim());
        if (start && empty) {
          for (let node: Node | null = start.nextSibling; node && node !== el;) { const next: Node | null = node.nextSibling; parent.removeChild(node); node = next; }
          parent.removeChild(start);
          parent.removeChild(el);
        } else {
          reason = 'customXml markers around deleted or moved content are not supported';
          supported = false;
        }
      } else if (/^customXml(Del|MoveFrom)RangeStart$/.test(kind)) {
        reason = 'customXml markers around deleted or moved content are not supported';
        supported = false;
      } else {
        supported = false;
      }
      if (!supported) {
        pending.add(el);
        const nodeId = options.nodeIdOf?.(el);
        result.unresolved.push({ kind, count: 1, reason, part: options.part ?? 'word/document.xml', ...(nodeId ? {nodeId} : {}) });
        continue;
      }
      noteResolved(el);
    }
  }
  return result;
}
