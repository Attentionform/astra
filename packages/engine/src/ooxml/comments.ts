import { DocxZipAdapter } from '../parser';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
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
  
  let parentParaId: string | null = null;
  const commentEls = commentsDoc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'comment');
  for (let i = 0; i < commentEls.length; i++) {
    const el = commentEls.item(i) as Element;
    if (el.getAttribute('w:id') === String(parentCommentId)) {
      const p = el.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'p').item(0) as Element;
      if (p) {
        parentParaId = p.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'paraId') ?? p.getAttribute('w14:paraId');
      }
      break;
    }
  }

  if (!parentParaId) {
    throw new Error(`Parent comment ID ${parentCommentId} not found in comments.xml`);
  }

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
