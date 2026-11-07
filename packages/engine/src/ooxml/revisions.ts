import { REVISION_TAGS } from './revision-ids';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REVISION_NAMES = new Set<string>(REVISION_TAGS);

export interface UnresolvedRevision {
  kind: string;
  count: number;
  reason: string;
  part: string;
}

export interface RevisionResolution {
  resolved: number;
  resolvedIds: number[];
  byKind: Record<string, number>;
  unresolved: UnresolvedRevision[];
  removedNodeIds?: string[];
  idsMayHaveChanged?: boolean;
}

const children = (element: Element): Element[] =>
  Array.from(element.childNodes).filter((node): node is Element => node.nodeType === 1);
const local = (node: Node | null): string => node && node.nodeType === 1 ? (node as Element).localName : '';

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
  const plain = (p: Element): boolean => children(p).every(el => {
    if (el.localName === 'pPr') return true;
    return el.localName === 'r' && children(el).every(child => ['rPr', 't', 'tab', 'br'].includes(child.localName));
  });
  if (!plain(paragraph) || !plain(next)) return false;
  const nextPr = children(next).find(el => el.localName === 'pPr');
  const anchor = nextPr ? nextPr.nextSibling : next.firstChild;
  for (const child of children(paragraph)) if (child.localName !== 'pPr') next.insertBefore(child, anchor);
  paragraph.parentNode!.removeChild(paragraph);
  return true;
}

interface MoveGroup { name: string; members: Element[]; markers: Element[]; valid: boolean }
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
      const group = groups.get(name) ?? { name, members: [], markers: [], valid: true };
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

/** Resolve supported revisions in one part. Unhandled structural revisions remain reported. */
export function resolveRevisions(documents: readonly Document[], mode: 'accept' | 'reject', options: { removeDeletedCells?: boolean; part?: string } = {}): RevisionResolution {
  const result: RevisionResolution = { resolved: 0, resolvedIds: [], byKind: {}, unresolved: [] };
  for (const doc of documents) {
    const pending = new Set<Element>();
    const containsPending = (parent: Element): boolean => {
      for (const item of pending) for (let node: Node | null = item.parentNode; node; node = node.parentNode) if (node === parent) return true;
      return false;
    };
    const moveGroups = findMoveGroups(doc);
    const elements = Array.from(doc.getElementsByTagName('*')).filter(el => el.namespaceURI === W_NS && REVISION_NAMES.has(el.localName));
    for (const el of elements.reverse()) {
      if (!el.parentNode) continue;
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
      if (insideUnsupportedMove) {
        reason = 'the containing move group is incomplete or contains nested revisions';
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
      } else {
        supported = false;
      }
      if (!supported) {
        pending.add(el);
        result.unresolved.push({ kind, count: 1, reason, part: options.part ?? 'word/document.xml' });
        continue;
      }
      result.resolved++;
      result.byKind[kind] = (result.byKind[kind] ?? 0) + 1;
      const value = el.getAttributeNS(W_NS, 'id');
      if (value && /^\d+$/.test(value)) result.resolvedIds.push(Number(value));
    }
  }
  return result;
}
