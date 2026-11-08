import type { DocxZipAdapter } from '../parser';
import { XMLSerializer } from '@xmldom/xmldom';
import { DOMParser } from './dom-parser';
import { assertNoXxe } from './xxe-guard';
import { deterministicCommentIso } from './track-changes-emitter';

const REL_TYPE_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const REL_TYPE_COMMENTS_EXTENDED = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const REL_TYPE_PEOPLE = 'http://schemas.microsoft.com/office/2011/relationships/people';

const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const CT_COMMENTS_EXTENDED = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const CT_PEOPLE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml';

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

const COMMENTS_XML_TEMPLATE =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
  ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
  ` xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>`;

const COMMENTS_EXTENDED_XML_TEMPLATE =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"` +
  ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>`;

const PEOPLE_XML_TEMPLATE =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>`;

import { generateValidParaId } from './schema-order';

let commentParaSeq = 0;

/** Test hook: reset the deterministic comment paraId counter (isolates tests). */
export function resetCommentParaSeqForTests(): void {
  commentParaSeq = 0;
}

function isCommentReproducible(): boolean {
  try {
    const noRaw = process.env['ASTRA_NO_REPRODUCIBLE'];
    if (noRaw !== undefined && noRaw !== null && String(noRaw).trim() !== '') {
      const v = String(noRaw).trim().toLowerCase();
      if (v === '1' || v === 'true' || v === 'yes' || v === 'on') return false;
    }
    const repRaw = process.env['ASTRA_REPRODUCIBLE'];
    if (repRaw !== undefined && repRaw !== null && String(repRaw).trim() !== '') {
      const v = String(repRaw).trim().toLowerCase();
      if (v === '0' || v === 'false' || v === 'no' || v === 'off') return false;
    }
    return true;
  } catch {
    return true;
  }
}

function generateParaId(deterministicSeed?: string): string {
  if (isCommentReproducible()) {
    const seed = deterministicSeed ?? `comment:${commentParaSeq++}`;
    return generateValidParaId(seed, 0);
  }
  // Opt-out: wall-clock random (legacy path).
  const val = (Math.floor(Math.random() * 0xFFFFFFFF) & 0x7FFFFFFF) || 1;
  return val.toString(16).toUpperCase().padStart(8, '0');
}

function isoNow(): string {
  return deterministicCommentIso();
}

/** Typed carrier for dropped non-fatal comment-XML warnings (module scope — parseXml has no result channel). */
const commentXmlWarnings: string[] = [];

function parseXml(xml: string, warnings?: string[]): Document {
  assertNoXxe(xml);
  const parser = new DOMParser({
    errorHandler: {
      // Non-fatal XML warnings ride the typed channel (same text as the
      // removed console.warn): per-call `warnings` when supplied, plus the
      // module collector observable via getCommentXmlWarnings().
      warning: (w) => {
        const msg = `[comments] XML warning (non-fatal): ${w}`;
        commentXmlWarnings.push(msg);
        if (warnings) warnings.push(msg);
      },
      error: (e) => { throw new Error(e) },
      fatalError: (f) => { throw new Error(f) }
    }
  });
  return parser.parseFromString(xml, 'text/xml');
}

/** Returns a copy of collected comment-XML warnings since the last clear. */
export function getCommentXmlWarnings(): string[] {
  return [...commentXmlWarnings];
}
/** Test hook: clears the collected comment-XML warnings. */
export function clearCommentXmlWarnings(): void {
  commentXmlWarnings.length = 0;
}
function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

export interface CommentRecord {
  id: number; paraId: string; paraIdParent?: string; author: string; date: string;
  resolved: boolean; text: string;
}

/** The comment element with this id, or null. */
function findCommentElement(commentsDoc: Document, id: number): Element | null {
  return Array.from(commentsDoc.getElementsByTagNameNS(W, 'comment')).find(el => el.getAttributeNS(W, 'id') === String(id)) ?? null;
}

/** The thread key of a comment: the w14:paraId of its LAST paragraph (MS-DOCX), or '' when it has none. */
export function commentThreadKey(commentEl: Element): string {
  const paragraphs = Array.from(commentEl.getElementsByTagNameNS(W, 'p'));
  return paragraphs[paragraphs.length - 1]?.getAttributeNS(W14, 'paraId') ?? '';
}

/**
 * The thread key of a comment, giving its last paragraph a fresh id when it has none (comments from
 * Word 2007 and other generators). Changes the DOM only; the caller writes comments.xml.
 */
function ensureThreadKey(commentsDoc: Document, commentEl: Element, seed: string): string {
  const existing = commentThreadKey(commentEl);
  if (existing) return existing;
  const paragraphs = Array.from(commentEl.getElementsByTagNameNS(W, 'p'));
  const last = paragraphs[paragraphs.length - 1];
  if (!last) throw new Error('E_VALIDATION: The comment has no paragraph to carry a thread id.');
  const used = new Set(Array.from(commentsDoc.getElementsByTagNameNS(W, 'p')).map(p => p.getAttributeNS(W14, 'paraId')).filter(Boolean));
  let paraId = generateParaId(seed);
  for (let n = 1; used.has(paraId); n++) paraId = generateParaId(`${seed}:${n}`);
  ensureCommentPartNamespaceAliases(commentsDoc);
  last.setAttributeNS(W14, 'w14:paraId', paraId);
  return paraId;
}

/** Read comments and thread metadata without modifying the package. */
export function listComments(zip: DocxZipAdapter): CommentRecord[] {
  const xml = zip.readTextOrNull('word/comments.xml');
  if (!xml) return [];
  const comments = parseXml(xml);
  const extXml = zip.readTextOrNull('word/commentsExtended.xml');
  const extended = extXml ? parseXml(extXml) : null;
  const byPara = new Map<string, Element>();
  if (extended) for (const el of Array.from(extended.getElementsByTagNameNS(W15, 'commentEx'))) {
    const id = el.getAttributeNS(W15, 'paraId');
    if (id) byPara.set(id, el);
  }
  return Array.from(comments.getElementsByTagNameNS(W, 'comment')).map(el => {
    const rawId = el.getAttributeNS(W, 'id');
    const paraId = commentThreadKey(el);
    const ex = byPara.get(paraId);
    const parent = ex?.getAttributeNS(W15, 'paraIdParent');
    return {
      id:rawId !== null && /^\d+$/.test(rawId) ? Number(rawId) : -1, paraId,
      ...(parent ? {paraIdParent:parent} : {}),
      author:el.getAttributeNS(W, 'author') ?? '', date:el.getAttributeNS(W, 'date') ?? '',
      resolved:['1','true','on'].includes(ex?.getAttributeNS(W15, 'done') ?? ''),
      text:Array.from(el.getElementsByTagNameNS(W, 't')).map(t => t.textContent ?? '').join(''),
    };
  });
}

export function setCommentResolved(zip: DocxZipAdapter, id: number, done: boolean): void {
  const commentsXml = zip.readTextOrNull('word/comments.xml');
  const commentsDoc = commentsXml ? parseXml(commentsXml) : null;
  const commentEl = commentsDoc ? findCommentElement(commentsDoc, id) : null;
  if (!commentsDoc || !commentEl) throw new Error(`E_VALIDATION: Comment ${id} was not found.`);
  // Optional parts that a generator left out are created; nothing is written until the key exists.
  const key = ensureThreadKey(commentsDoc, commentEl, `comment:${id}:resolve`);
  bootstrapCommentParts(zip);
  zip.writeText('word/comments.xml', serializeXml(commentsDoc));
  ensureCommentExEntry(zip, key);
  const ext = parseXml(zip.readText('word/commentsExtended.xml'));
  const entry = Array.from(ext.getElementsByTagNameNS(W15, 'commentEx')).find(el => el.getAttributeNS(W15, 'paraId') === key);
  if (!entry) throw new Error('E_VALIDATION: Comment metadata could not be created.');
  entry.setAttributeNS(W15, 'w15:done', done ? '1' : '0');
  zip.writeText('word/commentsExtended.xml', serializeXml(ext));
}

/**
 * Remove a whole comment thread. Every reference is checked before anything is changed. Optional parts
 * that a generator did not write (commentsExtended, commentsIds, commentsExtensible, a paragraph id) are
 * skipped. References that contradict each other make the call fail with no change.
 */
export function deleteCommentThread(zip: DocxZipAdapter, document: Document, id: number): void {
  const records = listComments(zip);
  const root = records.find(c => c.id === id);
  if (!root) throw new Error(`E_VALIDATION: Comment ${id} was not found.`);
  const thread = [root];
  for (let i = 0; i < thread.length; i++) {
    if (!thread[i].paraId) continue;
    for (const child of records.filter(c => c.paraIdParent === thread[i].paraId)) {
      if (thread.includes(child)) throw new Error('E_VALIDATION: Comment thread contains a cycle.');
      thread.push(child);
    }
  }
  const commentsXml = zip.readTextOrNull('word/comments.xml');
  if (!commentsXml) throw new Error('E_VALIDATION: Comment parts are incomplete.');
  const extendedXml = zip.readTextOrNull('word/commentsExtended.xml');
  const comments = parseXml(commentsXml);
  const extended = extendedXml ? parseXml(extendedXml) : null;
  const commentElements = Array.from(comments.getElementsByTagNameNS(W, 'comment'));
  const extendedElements = extended ? Array.from(extended.getElementsByTagNameNS(W15, 'commentEx')) : [];
  const idsXml = zip.readTextOrNull('word/commentsIds.xml');
  const idsDoc = idsXml ? parseXml(idsXml) : null;
  const extensibleXml = zip.readTextOrNull('word/commentsExtensible.xml');
  const extensible = extensibleXml ? parseXml(extensibleXml) : null;
  const markers: Element[] = [];
  const durableIds = new Set<string>();
  for (const comment of thread) {
    if (comment.id < 0 || commentElements.filter(el => el.getAttributeNS(W, 'id') === String(comment.id)).length !== 1)
      throw new Error('E_VALIDATION: Comment records are inconsistent.');
    if (comment.paraId) {
      if (records.filter(item => item.paraId === comment.paraId).length !== 1 ||
          extendedElements.filter(el => el.getAttributeNS(W15, 'paraId') === comment.paraId).length > 1)
        throw new Error('E_VALIDATION: Comment paragraph metadata is ambiguous.');
    }
    if (comment.paraIdParent && extendedElements.filter(el => el.getAttributeNS(W15, 'paraId') === comment.paraId).length !== 1)
      throw new Error('E_VALIDATION: Reply metadata is inconsistent.');
    const found: Record<string, Element[]> = {};
    for (const kind of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
      found[kind] = Array.from(document.getElementsByTagNameNS(W, kind)).filter(el => (el.getAttributeNS(W, 'id') || el.getAttribute('w:id')) === String(comment.id));
      if (found[kind].length > 1) throw new Error(`E_VALIDATION: Comment ${comment.id} has inconsistent ${kind} markup.`);
    }
    // A range always has both ends; a comment without a range keeps only its reference.
    if (found['commentRangeStart'].length !== found['commentRangeEnd'].length)
      throw new Error(`E_VALIDATION: Comment ${comment.id} has inconsistent commentRangeStart/commentRangeEnd markup.`);
    markers.push(...found['commentRangeStart'], ...found['commentRangeEnd'], ...found['commentReference']);
    if (idsDoc && comment.paraId) {
      const matches = Array.from(idsDoc.getElementsByTagName('*')).filter(el => Array.from(el.attributes).some(attr => attr.localName === 'paraId' && attr.value === comment.paraId));
      if (matches.length > 1) throw new Error('E_VALIDATION: Comment ID mapping is ambiguous.');
      for (const match of matches) for (const attr of Array.from(match.attributes)) if (attr.localName === 'durableId') durableIds.add(attr.value);
    }
  }
  const paraIds = new Set(thread.map(c => c.paraId).filter(Boolean));
  const commentIds = new Set(thread.map(c => String(c.id)));
  for (const el of commentElements) if (commentIds.has(el.getAttributeNS(W, 'id') ?? '')) el.parentNode?.removeChild(el);
  for (const el of extendedElements) if (paraIds.has(el.getAttributeNS(W15, 'paraId') ?? '')) el.parentNode?.removeChild(el);
  if (idsDoc) for (const el of Array.from(idsDoc.getElementsByTagName('*'))) {
    if (Array.from(el.attributes).some(attr => attr.localName === 'paraId' && paraIds.has(attr.value))) el.parentNode?.removeChild(el);
  }
  if (extensible && durableIds.size) for (const el of Array.from(extensible.getElementsByTagName('*'))) {
    if (Array.from(el.attributes).some(attr => attr.localName === 'durableId' && durableIds.has(attr.value))) el.parentNode?.removeChild(el);
  }
  for (const marker of markers) {
    const parent = marker.parentNode;
    parent?.removeChild(marker);
    if (marker.localName === 'commentReference' && parent?.nodeType === 1 && (parent as Element).localName === 'r' &&
        Array.from(parent.childNodes).every(n => n.nodeType !== 1 || (n as Element).localName === 'rPr')) parent.parentNode?.removeChild(parent);
  }
  zip.writeText('word/comments.xml', serializeXml(comments));
  if (extended) zip.writeText('word/commentsExtended.xml', serializeXml(extended));
  if (idsDoc) zip.writeText('word/commentsIds.xml', serializeXml(idsDoc));
  if (extensible && durableIds.size) zip.writeText('word/commentsExtensible.xml', serializeXml(extensible));
}

export function bootstrapCommentParts(zip: DocxZipAdapter): void {
  const created: string[] = [];

  if (!zip.hasFile('word/comments.xml')) {
    zip.writeText('word/comments.xml', COMMENTS_XML_TEMPLATE);
    created.push('word/comments.xml');
  }
  if (!zip.hasFile('word/commentsExtended.xml')) {
    zip.writeText('word/commentsExtended.xml', COMMENTS_EXTENDED_XML_TEMPLATE);
    created.push('word/commentsExtended.xml');
  }
  if (!zip.hasFile('word/people.xml')) {
    zip.writeText('word/people.xml', PEOPLE_XML_TEMPLATE);
    created.push('word/people.xml');
  }

  if (created.length === 0) return;
  ensureContentTypes(zip, created);
  ensureRelationships(zip, created);
}

function ensureContentTypes(zip: DocxZipAdapter, newParts: string[]): void {
  const ctPath = '[Content_Types].xml';
  const ctXml = zip.readTextOrNull(ctPath);
  if (!ctXml) return;
  const ctDoc = parseXml(ctXml);
  const typesEl = ctDoc.documentElement;
  const ctNs = 'http://schemas.openxmlformats.org/package/2006/content-types';

  const partToCt: Record<string, string> = {
    'word/comments.xml': CT_COMMENTS,
    'word/commentsExtended.xml': CT_COMMENTS_EXTENDED,
    'word/people.xml': CT_PEOPLE,
  };

  const overrides = Array.from(typesEl.getElementsByTagNameNS(ctNs, 'Override'));
  const existingPartNames = new Set(overrides.map((o) => o.getAttribute('PartName')));

  for (const part of newParts) {
    const partName = `/${part}`;
    const contentType = partToCt[part];
    if (!contentType || existingPartNames.has(partName)) continue;

    const override = ctDoc.createElementNS(ctNs, 'Override');
    override.setAttribute('PartName', partName);
    override.setAttribute('ContentType', contentType);
    typesEl.appendChild(override);
  }
  zip.writeText(ctPath, serializeXml(ctDoc));
}

function ensureRelationships(zip: DocxZipAdapter, newParts: string[]): void {
  const relsPath = 'word/_rels/document.xml.rels';
  const relsXml = zip.readTextOrNull(relsPath);
  if (!relsXml) return;
  
  const relsDoc = parseXml(relsXml);
  const relsEl = relsDoc.documentElement;
  const relNs = 'http://schemas.openxmlformats.org/package/2006/relationships';

  const partToRelType: Record<string, string> = {
    'word/comments.xml': REL_TYPE_COMMENTS,
    'word/commentsExtended.xml': REL_TYPE_COMMENTS_EXTENDED,
    'word/people.xml': REL_TYPE_PEOPLE,
  };

  const existingRels = Array.from(relsEl.getElementsByTagNameNS(relNs, 'Relationship'));
  const existingTypes = new Set(existingRels.map((r) => r.getAttribute('Type')));

  // BX-05: seed from canonical rIdN ids only (timestamp/rId_link_* bypass ids
  // must not drag the counter into the trillions), then existence-loop every
  // allocation so ids minted outside this counter can never collide.
  const usedIds = new Set(existingRels.map((r) => r.getAttribute('Id')));
  let maxId = 0;
  for (const rel of existingRels) {
    const id = rel.getAttribute('Id') ?? '';
    const match = /^rId(\d+)$/.exec(id);
    if (match) maxId = Math.max(maxId, parseInt(match[1]!, 10));
  }
  const nextFreeRId = (): string => {
    do {
      maxId++;
    } while (usedIds.has(`rId${maxId}`));
    const fresh = `rId${maxId}`;
    usedIds.add(fresh);
    return fresh;
  };

  for (const part of newParts) {
    const relType = partToRelType[part];
    if (!relType || existingTypes.has(relType)) continue;

    const freshId = nextFreeRId();
    const rel = relsDoc.createElementNS(relNs, 'Relationship');
    rel.setAttribute('Id', freshId);
    rel.setAttribute('Type', relType);
    rel.setAttribute('Target', part.replace('word/', ''));
    relsEl.appendChild(rel);
  }
  zip.writeText(relsPath, serializeXml(relsDoc));
}

export function allocateNextCommentId(commentsDoc: Document): number {
  const commentEls = commentsDoc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'comment');
  let maxId = -1;
  for (let i = 0; i < commentEls.length; i++) {
    const el = commentEls.item(i) as Element;
    const idStr = el.getAttribute('w:id');
    if (idStr) {
      const id = parseInt(idStr, 10);
      if (id > maxId) maxId = id;
    }
  }
  return maxId + 1;
}

export function addComment(
  zip: DocxZipAdapter,
  paragraphEl: Element,
  author: string,
  text: string,
  isoDate?: string,
): void {
  bootstrapCommentParts(zip);
  
  const commentsXml = zip.readText('word/comments.xml');
  const commentsDoc = parseXml(commentsXml);
  const commentId = allocateNextCommentId(commentsDoc);

  insertCommentMarkers(paragraphEl, commentId);

  const paraId = generateParaId(`comment:${commentId}:${author}:${text}`);
  addCommentElement(commentsDoc, {
    id: commentId,
    author,
    initials: (Array.from(author)[0] ?? '').toUpperCase(),
    text,
    paraId,
    date: isoDate,
  });
  zip.writeText('word/comments.xml', serializeXml(commentsDoc));

  ensureAuthorInPeople(zip, author);
}

export function addCommentReply(
  zip: DocxZipAdapter,
  parentCommentId: number,
  author: string,
  text: string,
  isoDate?: string,
): number {
  const commentsXml = zip.readText('word/comments.xml');
  const commentsDoc = parseXml(commentsXml);
  
  const parentEl = findCommentElement(commentsDoc, parentCommentId);
  if (!parentEl) throw new Error(`Parent comment ID ${parentCommentId} not found in comments.xml`);
  // Replies point at the last paragraph of the parent. A parent from Word 2007 has no paragraph id yet.
  const parentParaId = ensureThreadKey(commentsDoc, parentEl, `comment:${parentCommentId}:thread`);
  bootstrapCommentParts(zip);

  const commentId = allocateNextCommentId(commentsDoc);
  const replyParaId = generateParaId(`comment:${commentId}:${author}:${text}:reply:${parentCommentId}`);
  
  addCommentElement(commentsDoc, {
    id: commentId,
    author,
    initials: (Array.from(author)[0] ?? '').toUpperCase(),
    text,
    paraId: replyParaId,
    date: isoDate,
  });
  zip.writeText('word/comments.xml', serializeXml(commentsDoc));

  linkReplyInCommentsExtended(zip, replyParaId, parentParaId);
  ensureCommentExEntry(zip, parentParaId);
  ensureAuthorInPeople(zip, author);

  return commentId;
}

function linkReplyInCommentsExtended(zip: DocxZipAdapter, replyParaId: string, parentParaId: string): void {
  const extXml = zip.readText('word/commentsExtended.xml');
  const extDoc = parseXml(extXml);
  ensureCommentPartNamespaceAliases(extDoc);
  
  const exEl = extDoc.createElementNS('http://schemas.microsoft.com/office/word/2012/wordml', 'w15:commentEx');
  exEl.setAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'w15:paraId', replyParaId);
  exEl.setAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'w15:paraIdParent', parentParaId);
  exEl.setAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'w15:done', '0');
  extDoc.documentElement.appendChild(exEl);

  zip.writeText('word/commentsExtended.xml', serializeXml(extDoc));
}

function ensureCommentExEntry(zip: DocxZipAdapter, paraId: string): void {
  const extXml = zip.readText('word/commentsExtended.xml');
  const extDoc = parseXml(extXml);
  ensureCommentPartNamespaceAliases(extDoc);
  
  const existing = extDoc.documentElement.getElementsByTagNameNS('http://schemas.microsoft.com/office/word/2012/wordml', 'commentEx');
  for (let i = 0; i < existing.length; i++) {
    const el = existing.item(i) as Element;
    if ((el.getAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'paraId') ?? el.getAttribute('w15:paraId')) === paraId) return;
  }

  const exEl = extDoc.createElementNS('http://schemas.microsoft.com/office/word/2012/wordml', 'w15:commentEx');
  exEl.setAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'w15:paraId', paraId);
  exEl.setAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'w15:done', '0');
  extDoc.documentElement.appendChild(exEl);

  zip.writeText('word/commentsExtended.xml', serializeXml(extDoc));
}

function insertCommentMarkers(paragraphEl: Element, commentId: number): void {
  const doc = paragraphEl.ownerDocument;
  if (!doc) return;
  const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  
  const rangeStart = doc.createElementNS(wNs, 'w:commentRangeStart');
  rangeStart.setAttribute('w:id', String(commentId));

  const rangeEnd = doc.createElementNS(wNs, 'w:commentRangeEnd');
  rangeEnd.setAttribute('w:id', String(commentId));

  const refRun = doc.createElementNS(wNs, 'w:r');
  const commentRef = doc.createElementNS(wNs, 'w:commentReference');
  commentRef.setAttribute('w:id', String(commentId));
  refRun.appendChild(commentRef);

  // w:pPr must remain the first child per CT_P; insert range start after it when present
  const pPrList = paragraphEl.getElementsByTagName('w:pPr');
  let directPPr: Element | null = null;
  for (let i = 0; i < pPrList.length; i++) {
    const el = pPrList.item(i) as Element | null;
    if (el && el.parentNode === paragraphEl) {
      directPPr = el;
      break;
    }
  }
  if (!directPPr) {
    const nsList = paragraphEl.getElementsByTagNameNS(wNs, 'pPr');
    for (let i = 0; i < nsList.length; i++) {
      const el = nsList.item(i) as Element | null;
      if (el && el.parentNode === paragraphEl) {
        directPPr = el;
        break;
      }
    }
  }
  if (directPPr) {
    paragraphEl.insertBefore(rangeStart, directPPr.nextSibling);
  } else {
    paragraphEl.insertBefore(rangeStart, paragraphEl.firstChild);
  }
  paragraphEl.appendChild(rangeEnd);
  paragraphEl.appendChild(refRun);
}

function ensureCommentPartNamespaceAliases(commentsDoc: Document): void {
  const root = commentsDoc.documentElement;
  const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const w15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
  if (root.lookupNamespaceURI('w14') !== w14) {
    root.setAttributeNS(XMLNS_NS, 'xmlns:w14', w14);
  }
  if (root.lookupNamespaceURI('w15') !== w15) {
    root.setAttributeNS(XMLNS_NS, 'xmlns:w15', w15);
  }
}

function addCommentElement(
  commentsDoc: Document,
  params: { id: number; author: string; initials: string; text: string; paraId: string; date?: string },
): void {
  ensureCommentPartNamespaceAliases(commentsDoc);
  const root = commentsDoc.documentElement;
  const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const w14Ns = 'http://schemas.microsoft.com/office/word/2010/wordml';

  const commentEl = commentsDoc.createElementNS(wNs, 'w:comment');
  commentEl.setAttribute('w:id', String(params.id));
  commentEl.setAttribute('w:author', params.author);
  commentEl.setAttribute('w:date', params.date ?? isoNow());
  commentEl.setAttribute('w:initials', params.initials);

  const p = commentsDoc.createElementNS(wNs, 'w:p');
  p.setAttributeNS(w14Ns, 'w14:paraId', params.paraId);

  const refRun = commentsDoc.createElementNS(wNs, 'w:r');
  const annotRef = commentsDoc.createElementNS(wNs, 'w:annotationRef');
  refRun.appendChild(annotRef);
  p.appendChild(refRun);

  const textRun = commentsDoc.createElementNS(wNs, 'w:r');
  const t = commentsDoc.createElementNS(wNs, 'w:t');
  t.appendChild(commentsDoc.createTextNode(params.text));
  textRun.appendChild(t);
  p.appendChild(textRun);

  commentEl.appendChild(p);
  root.appendChild(commentEl);
}

function ensureAuthorInPeople(zip: DocxZipAdapter, author: string): void {
  const peopleXml = zip.readText('word/people.xml');
  const peopleDoc = parseXml(peopleXml);
  ensureCommentPartNamespaceAliases(peopleDoc);
  const root = peopleDoc.documentElement;
  const w15Ns = 'http://schemas.microsoft.com/office/word/2012/wordml';

  const persons = root.getElementsByTagNameNS(w15Ns, 'person');
  for (let i = 0; i < persons.length; i++) {
    const el = persons.item(i) as Element;
    const name = el.getAttribute('w15:author');
    if (name === author) return;
  }

  const personEl = peopleDoc.createElementNS(w15Ns, 'w15:person');
  personEl.setAttributeNS(w15Ns, 'w15:author', author);

  const presenceInfo = peopleDoc.createElementNS(w15Ns, 'w15:presenceInfo');
  presenceInfo.setAttributeNS(w15Ns, 'w15:providerId', 'None');
  presenceInfo.setAttributeNS(w15Ns, 'w15:userId', author);
  personEl.appendChild(presenceInfo);

  root.appendChild(personEl);
  zip.writeText('word/people.xml', serializeXml(peopleDoc));
}

/** Anchors a reply next to its parent: same range, reference run right after the parent's. Returns false when the parent has no markers. */
export function anchorCommentReply(docDom: Document, parentId: number, replyId: number): boolean {
  const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const find = (tag: string): Element | null => {
    const list = docDom.getElementsByTagNameNS(wNs, tag);
    for (let i = 0; i < list.length; i++) {
      const el = list.item(i) as Element;
      if (el.getAttribute('w:id') === String(parentId)) return el;
    }
    return null;
  };
  const start = find('commentRangeStart');
  const end = find('commentRangeEnd');
  const ref = find('commentReference');
  const refRun = ref?.parentNode as Element | null | undefined;
  if (!start || !end || !refRun || !start.parentNode || !end.parentNode || !refRun.parentNode) return false;
  const mk = (tag: string): Element => {
    const el = docDom.createElementNS(wNs, `w:${tag}`);
    el.setAttribute('w:id', String(replyId));
    return el;
  };
  start.parentNode.insertBefore(mk('commentRangeStart'), start.nextSibling);
  end.parentNode.insertBefore(mk('commentRangeEnd'), end.nextSibling);
  const run = docDom.createElementNS(wNs, 'w:r');
  run.appendChild(mk('commentReference'));
  refRun.parentNode.insertBefore(run, refRun.nextSibling);
  return true;
}

/**
 * Reads w:commentReference ids inside the paragraph, identifies which comments
 * are top-level (not replies in commentsExtended.xml), and returns the last
 * top-level comment id, or null if none.
 */
export function findLastTopLevelCommentId(zip: DocxZipAdapter, paragraphEl: Element): number | null {
  const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const w14Ns = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const w15Ns = 'http://schemas.microsoft.com/office/word/2012/wordml';

  const refList = paragraphEl.getElementsByTagNameNS(wNs, 'commentReference');
  const refIds: number[] = [];
  for (let i = 0; i < refList.length; i++) {
    const el = refList.item(i) as Element;
    const id = el.getAttribute('w:id');
    if (id !== null && id !== '') {
      const num = parseInt(id, 10);
      if (!Number.isNaN(num)) {
        refIds.push(num);
      }
    }
  }
  if (refIds.length === 0) return null;

  const commentsXml = zip.readTextOrNull('word/comments.xml');
  if (!commentsXml) return null;
  const commentsDoc = parseXml(commentsXml);

  const commentIdToParaId = new Map<number, string>();
  const commentEls = commentsDoc.getElementsByTagNameNS(wNs, 'comment');
  for (let i = 0; i < commentEls.length; i++) {
    const el = commentEls.item(i) as Element;
    const idAttr = el.getAttribute('w:id');
    if (idAttr === null || idAttr === '') continue;
    const id = parseInt(idAttr, 10);
    const p = el.getElementsByTagNameNS(wNs, 'p').item(0) as Element | null;
    if (p) {
      const paraId = p.getAttributeNS(w14Ns, 'paraId') ?? p.getAttribute('w14:paraId');
      if (paraId) commentIdToParaId.set(id, paraId);
    }
  }

  const replyParaIds = new Set<string>();
  const extXml = zip.readTextOrNull('word/commentsExtended.xml');
  if (extXml) {
    const extDoc = parseXml(extXml);
    const exEls = extDoc.getElementsByTagNameNS(w15Ns, 'commentEx');
    for (let i = 0; i < exEls.length; i++) {
      const exEl = exEls.item(i) as Element;
      const parentParaId = exEl.getAttributeNS(w15Ns, 'paraIdParent') ?? exEl.getAttribute('w15:paraIdParent');
      if (parentParaId) {
        const paraId = exEl.getAttributeNS(w15Ns, 'paraId') ?? exEl.getAttribute('w15:paraId');
        if (paraId) replyParaIds.add(paraId);
      }
    }
  }

  for (let i = refIds.length - 1; i >= 0; i--) {
    const cId = refIds[i]!;
    const paraId = commentIdToParaId.get(cId);
    if (paraId && !replyParaIds.has(paraId)) {
      return cId;
    }
  }

  return null;
}
