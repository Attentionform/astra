import { createWordElement } from '../ooxml/ns';
import { toError } from '../errors';
import { DocumentAST, Node } from '../ast';
import { DocxParser } from '../parser';
import { BaseOperation, createTextElement, findNode } from './base';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { assertNoXxe } from '../ooxml/xxe-guard';
import {
  allocateNextBookmarkId,
  allocateSafeHyperlinkRId,
} from '../manipulators/relationships';

export const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

export interface HyperlinkOperation extends BaseOperation {
  action: 'hyperlink' | 'bookmark';
  operation?: 'insert';
  target: string; // anchor node ID
  text?: string;
  url?: string;       // For external hyperlinks
  bookmarkName?: string; // For internal bookmark anchors
}

export function executeHyperlink(op: HyperlinkOperation, ast: DocumentAST, parser?: DocxParser): void {
  const targetNode: Node | undefined = findNode(ast, op.target);

  if (!targetNode) throw new Error(`Target node '${op.target}' not found`);

  if (op.action === 'hyperlink' && (targetNode.xml as Element).tagName !== 'w:p') {
    throw toError(
      'E_VALIDATION',
      `hyperlink applies to paragraph text, but '${op.target}' is a ${targetNode.type} node.`,
      "Target a paragraph id (see inspect_document)."
    );
  }

  const doc = targetNode.xml.ownerDocument;
  if (!doc) throw new Error("Document not found on target node");

  if (op.action === 'bookmark') {
    // OX-12 owning-batch fix: the old path hardcoded w:id="0" on EVERY
    // bookmark, so any two bookmarks shared an id (Word requires unique ids
    // per bookmarkStart/bookmarkEnd; the start/end of ONE pair share their
    // id by spec). Allocate via the monotonic allocator (scan max + process
    // high-water reservation, so even two pre-insert allocations diverge).
    // The allocator's floor is 1 on a bookmark-free document; 0 was the
    // bypass's hardcoded value, so reserve past it here to prove the bypass
    // is gone (no w:id 0 is ever minted by this path).
    let bmId = allocateNextBookmarkId(doc);
    if (bmId <= 0) bmId = allocateNextBookmarkId(doc);
    // Default name derives from the unique id (Date.now() names collided
    // within the same millisecond); explicit bookmarkName is kept as given.
    const bName = op.bookmarkName || `bm_${bmId}`;

    const bmStart = createWordElement(doc, 'w:bookmarkStart');
    bmStart.setAttribute('w:id', String(bmId));
    bmStart.setAttribute('w:name', bName);

    const bmEnd = createWordElement(doc, 'w:bookmarkEnd');
    bmEnd.setAttribute('w:id', String(bmId));

    // DEFECT-01 (ECMA-376 §17.3.1.22): w:pPr must be the first child of w:p.
    // Inserting bookmarkStart as firstChild before an existing pPr violates
    // the schema (validator: unexpected pPr). Insert after pPr when present.
    const pPrList = targetNode.xml.getElementsByTagName('w:pPr');
    const pPrEl = pPrList.length > 0 ? (pPrList[0] as unknown as Element) : undefined;
    if (pPrEl?.nextSibling) {
      targetNode.xml.insertBefore(bmStart, pPrEl.nextSibling);
    } else if (pPrEl) {
      targetNode.xml.appendChild(bmStart);
    } else {
      targetNode.xml.insertBefore(bmStart, targetNode.xml.firstChild);
    }
    targetNode.xml.appendChild(bmEnd);
    targetNode.version += 1;
    return;
  }

  // B-4 internal hyperlink: bookmarkName means "jump within this document".
  // Emits <w:hyperlink w:anchor="..." w:history="1"> with the standard styled
  // run and NO External relationship (the old path always minted one, so an
  // internal link validated as an external URL and never jumped).
  // Precedence note: when BOTH url and bookmarkName are present, bookmarkName
  // (internal) wins and url is ignored — never mints an External rel for a
  // bookmark jump. Pass exactly one of the two to avoid ambiguity.
  const bookmarkName = op.bookmarkName as string | undefined;
  if (bookmarkName) {
    const text = op.text || bookmarkName;
    const hyperlink = createWordElement(doc, 'w:hyperlink');
    hyperlink.setAttribute('w:anchor', bookmarkName);
    hyperlink.setAttribute('w:history', '1');

    const run = createWordElement(doc, 'w:r');
    const rPr = createWordElement(doc, 'w:rPr');
    const color = createWordElement(doc, 'w:color');
    color.setAttribute('w:val', '0066CC');
    const underline = createWordElement(doc, 'w:u');
    underline.setAttribute('w:val', 'single');

    rPr.appendChild(color);
    rPr.appendChild(underline);
    run.appendChild(rPr);
    run.appendChild(createTextElement(doc, text));

    hyperlink.appendChild(run);
    targetNode.xml.appendChild(hyperlink);
    targetNode.version += 1;
    return;
  }

  // External Hyperlink — OX-12 owning-batch fix: the old path minted
  // `rId_link_${Date.now()}` straight to the zip (two same-ms hyperlinks
  // shared one Relationship Id, and ast.relationships never learned about
  // either, so later allocator passes could reuse them). Now: canonical rIdN
  // via the atomic allocator (AST-aware + process-wide reservation), written
  // to ast.relationships BEFORE the zip so every later pass (explicit
  const rawUrl = op.url || 'https://example.com';
  const text = op.text || rawUrl;

  // P0-11: Validate external URL schemes. tel:/sms: are legit call/message
  // links; scheme-less URLs (relative paths, #fragment anchors) pass through
  // as relative targets. javascript:/data:/vbscript:/file: still throw.
  const ALLOWED_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'ftp:', 'tel:', 'sms:']);
  let scheme = '';
  let hasScheme = false;
  const schemeMatch = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(rawUrl.trim());
  if (schemeMatch) {
    scheme = schemeMatch[1].toLowerCase() + ':';
    hasScheme = true;
  } else {
    try {
      const parsed = new URL(rawUrl);
      scheme = parsed.protocol.toLowerCase();
      hasScheme = true;
    } catch {
      // No parseable scheme: relative URL (./page, /abs/path, #anchor, ?q=1).
      scheme = '';
      hasScheme = false;
    }
  }

  if (hasScheme && !ALLOWED_SCHEMES.has(scheme)) {
    throw new Error(
      `E_VALIDATION: Invalid hyperlink URL scheme '${scheme}'. Remedy: Hyperlink URL must use http, https, mailto, or ftp scheme (tel, sms, and relative URLs also accepted)`
    );
  }

  const url = rawUrl;
  const relId = allocateSafeHyperlinkRId(ast);
  ast.relationships[relId] = {
    type: HYPERLINK_REL_TYPE,
    target: url,
    targetMode: 'External',
  };

  // 1. Register relationship in word/_rels/document.xml.rels
  if (parser) {
    const zipAdapter = parser.getZipAdapter();
    const relsPath = 'word/_rels/document.xml.rels';
    if (zipAdapter.hasFile(relsPath)) {
      const relsRaw = zipAdapter.readText(relsPath);
      assertNoXxe(relsRaw, relsPath);
      const relsDoc = new DOMParser().parseFromString(relsRaw, 'text/xml');
      const relsEl = relsDoc.getElementsByTagName('Relationships')[0];
      if (relsEl) {
        const rel = relsDoc.createElement('Relationship');
        rel.setAttribute('Id', relId);
        rel.setAttribute('Type', HYPERLINK_REL_TYPE);
        rel.setAttribute('Target', url);
        rel.setAttribute('TargetMode', 'External');
        relsEl.appendChild(rel);
        zipAdapter.writeText(relsPath, new XMLSerializer().serializeToString(relsDoc));
      }
    }
  }

  // 2. Create <w:hyperlink r:id="..."> element in document.xml
  const hyperlink = createWordElement(doc, 'w:hyperlink');
  hyperlink.setAttribute('r:id', relId);
  hyperlink.setAttribute('w:history', '1');

  const run = createWordElement(doc, 'w:r');
  const rPr = createWordElement(doc, 'w:rPr');
  const color = createWordElement(doc, 'w:color');
  color.setAttribute('w:val', '0066CC');
  const underline = createWordElement(doc, 'w:u');
  underline.setAttribute('w:val', 'single');

  rPr.appendChild(color);
  rPr.appendChild(underline);
  run.appendChild(rPr);
  run.appendChild(createTextElement(doc, text));

  hyperlink.appendChild(run);
  targetNode.xml.appendChild(hyperlink);
  targetNode.version += 1;
}
