import { DocumentAST, Node, Section, calculateNodeChecksum } from '../ast';
import { BaseOperation, InsertChainHost, createTextElement, validateVersion } from './base';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import * as crypto from 'crypto';
import { DocxParser, isValidParaId, W14_NS, upsertSectPrChild } from '../parser';
import { insertIntoSection } from '../parser/sections';
import { executeTable } from './table';
import type { StyleGuideConfig } from '../style-guide/types';
import { createWordElement } from '../ooxml/ns';
import { allocateNextBookmarkId } from '../manipulators/relationships';
import { requireTwips } from '../operation-schema';
import { resolveRevisionStamp } from '../ooxml/track-changes-emitter';
import { generateValidParaId, sortChildrenBySchema, PPR_CHILD_ORDER, RPR_CHILD_ORDER, TBLPR_CHILD_ORDER } from '../ooxml/schema-order';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { resolveReproducibleMode } from '../zip-hardening';
import { appendElementsToHeaderFooter, resolveHeaderFooterTarget } from './headerFooter';
import { toError } from '../errors';

export interface InsertOperation extends BaseOperation {
  action: 'insert';
  type: 'paragraph' | 'pageBreak' | 'sectionBreak' | 'table' | 'xml' | 'divider' | 'horizontalRule';
  anchor: string;
  position: 'before' | 'after';
  /** Attribution for a tracked insert (trackChanges:true). */
  author?: string;
  date?: string;
  text?: string;
  xml?: string;
  id?: string;
  headers?: string[];
  rows?: string[][];
  cellPadding?: number | { top?: number; bottom?: number; left?: number; right?: number };
  // R-42: table caption args, honored by the table caption path
  // (operations/table.ts setCaption) after the table is placed.
  captionText?: string;
  captionPosition?: string;
  /**
   * Table inserts only: keep every row on one page (cantSplit) with a
   * repeating header (tblHeader) plus keepNext chaining. Explicit wins;
   * otherwise the guide `tables.keepTogether` applies; otherwise the legacy
   * markers (cantSplit + tblHeader, no keepNext) are kept byte-identical.
   * Accepts flat or params-nested values (nested wins).
   */
  keepTogether?: boolean;
  /**
   * Table inserts only: bold the header row text. Explicit wins (flat or params-nested,
   * nested wins); otherwise the guide `tables.headerBold`; otherwise bold.
   */
  headerBold?: boolean;
  trackChanges?: boolean;
  color?: string;
  thickness?: number; // eighths of a point: 8 = 1pt, 12 = 1.5pt, 24 = 3pt
  lineStyle?: 'single' | 'double' | 'dashed' | 'dotted' | 'thick' | 'wave' | string;
  spacingBefore?: number | string; // twips, or points as "12pt"
  spacingAfter?: number | string;  // twips, or points as "12pt"
  breakType?: 'nextPage' | 'nextColumn' | 'continuous' | 'evenPage' | 'oddPage' | string;
  // Destination store: 'body' (default) or a repeating 'header'/'footer'
  // part. Paragraph-style content only; other types fail closed for
  // header/footer targets. Anchor/position are ignored there (appends in
  // call order). Accepts flat or params-nested values.
  target?: 'body' | 'header' | 'footer';
  pageType?: 'default' | 'first' | 'even';
  sectionIndex?: number;
  // Legacy flat aliases honored alongside `type` (mirrors table.ts pick).
  insertType?: string;
  paragraphType?: string;
  nodeType?: string;
  params?: Record<string, unknown>;
}

// Module-level counter so generated ids stay unique even within the same millisecond.
let generatedIdSeq = 0;

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

function isInsertReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

function insertHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

function mintInsertId(seed: string, live: Set<string>, prefix: string): string {
  if (!isInsertReproducible()) return `${prefix}_${Date.now()}_${++generatedIdSeq}`;
  const base = `${prefix}_${insertHash8(seed)}`;
  let candidate = base;
  let n = 2;
  while (live.has(candidate)) candidate = `${base}_${n++}`;
  return candidate;
}

function deriveInsertSeed(op: InsertOperation): string {
  if (op.id) return op.id;
  const parts = [
    op.anchor || '',
    op.position || '',
    op.type || 'paragraph',
    op.text || '',
    op.xml || '',
    op.breakType || '',
    op.color || '',
    op.thickness !== undefined ? String(op.thickness) : '',
    op.lineStyle || '',
    (op.headers || []).join(','),
    (op.rows || []).map((r) => r.join(',')).join(';'),
  ];
  return parts.join('|');
}

// Mints a valid w14:paraId on the insert root and its w:p descendants that
// lack one, so inserted paragraphs carry a reload-surviving identity from
// creation and later snapshot restores round-trip byte-identically.
// A caller-supplied valid id is preserved ONLY when it does not collide with
// the live-document used set (or an earlier root of the same fragment); on
// collision a fresh id is minted so no two paragraphs ever share an identity.
// Minted ids are loop-guarded against every valid id already present in the
// live document.
function stampMissingParaIds(root: Element, doc: Document, seed?: string, usedSet?: Set<string>): void {
  const used = usedSet || new Set<string>();
  if (!usedSet) {
    const allParas = doc.getElementsByTagName('w:p');
    for (let i = 0; i < allParas.length; i++) {
      const v = (allParas[i] as unknown as Element).getAttribute?.('w14:paraId');
      if (v && isValidParaId(v)) used.add(v.toUpperCase());
    }
  }
  const targets: Element[] = [];
  if ((root as Element).tagName === 'w:p') targets.push(root);
  const descendants = (root as Element).getElementsByTagName?.('w:p');
  if (descendants) {
    for (let i = 0; i < descendants.length; i++) {
      const el = descendants[i] as unknown as Element;
      if (el !== root) targets.push(el);
    }
  }
  let minted = false;
  for (let tIdx = 0; tIdx < targets.length; tIdx++) {
    const p = targets[tIdx];
    const v = p.getAttribute('w14:paraId');
    if (v && isValidParaId(v)) {
      const upper = v.toUpperCase();
      if (!used.has(upper)) {
        used.add(upper);
        continue;
      }
      // Collision with a live-document id or an earlier fragment root:
      // fall through and mint a fresh identity instead of duplicating.
    }
    let id: string;
    let attempt = 0;
    do {
      if (seed) {
        id = generateValidParaId(`${seed}:${tIdx}`, attempt);
      } else if (isInsertReproducible()) {
        id = generateValidParaId(`insert-fallback:${tIdx}`, attempt);
      } else {
        id = generateValidParaId();
      }
      attempt++;
    } while (used.has(id));
    used.add(id);
    p.setAttributeNS(W14_NS, 'w14:paraId', id);
    minted = true;
  }
  if (minted) {
    const docRoot = doc.documentElement;
    if (docRoot && !docRoot.hasAttribute('xmlns:w14')) {
      docRoot.setAttributeNS(XMLNS_NS, 'xmlns:w14', W14_NS);
    }
  }
}

export class InsertValidationError extends Error {
  readonly code = 'E_VALIDATION' as const;
  readonly remedy: string;
  readonly retryable = false as const;

  constructor(message: string, remedy: string) {
    super(`E_VALIDATION: ${message}. Remedy: ${remedy}`);
    this.name = 'InsertValidationError';
    this.remedy = remedy;
    Object.setPrototypeOf(this, InsertValidationError.prototype);
  }
}

const ALLOWED_OOXML_PREFIXES = new Set([
  'w', 'm', 'wp', 'a', 'pic', 'w14', 'w15', 'w16', 'w16cex', 'w16cid', 'w16se',
  'r', 'mc', 'v', 'o', 'w10', 'c', 'cx', 'dgm', 'wp14', 'xml'
]);

const ALLOWED_OOXML_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://schemas.microsoft.com/office/word/2006/wordml',
  'http://schemas.microsoft.com/office/word/2010/wordml',
  'http://schemas.microsoft.com/office/word/2012/wordml',
  'http://schemas.microsoft.com/office/word/2014/wordml',
  'http://schemas.microsoft.com/office/word/2016/wordml',
  'http://schemas.microsoft.com/office/word/2018/wordml',
  'http://schemas.microsoft.com/office/word/2020/wordml',
  'http://schemas.microsoft.com/office/word/2023/wordml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://schemas.microsoft.com/office/2006/relationships',
  'http://schemas.openxmlformats.org/officeDocument/2006/math',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'http://schemas.openxmlformats.org/drawingml/2006/chart',
  'http://schemas.microsoft.com/office/drawing/2014/chartex',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'urn:schemas-microsoft-com:vml',
  'urn:schemas-microsoft-com:office:office',
  'urn:schemas-microsoft-com:office:word',
  'http://www.w3.org/2000/xmlns/',
  'http://www.w3.org/XML/1998/namespace'
]);

const DANGEROUS_TAGS = new Set([
  'script', 'iframe', 'embed', 'object', 'applet', 'form', 'input',
  'meta', 'link', 'style', 'html', 'body', 'head', 'base', 'textarea', 'button',
  'svg', 'img', 'video', 'audio', 'canvas'
]);

// Bare (unprefixed) attributes the engine itself emits on drawing anchors.
// wp:docPr/pic:cNvPr carry schema-level `id`/`name` with no prefix; every
// other formatting attribute the engine emits is namespace-prefixed
// (w:val, w:sz, w:color, ...) or is data-astra-id / xml:* / xmlns:*.
const ALLOWED_BARE_ATTRIBUTES = new Set(['id', 'name']);

function validateOxmlAttribute(el: Element, attr: Attr): void {
  const tagName = el.tagName || el.nodeName;
  const name: string = attr.name || '';
  const lower = name.toLowerCase();
  if (lower === 'data-astra-id') return;
  const colon = name.indexOf(':');
  const prefix = colon !== -1 ? name.slice(0, colon).toLowerCase() : '';
  const local = (colon !== -1 ? name.slice(colon + 1) : name).toLowerCase();
  if (local === 'style' || local.indexOf('on') === 0) {
    throw new InsertValidationError(
      `Disallowed attribute '${name}' on tag '<${tagName}>' in XML insert`,
      `Remove event-handler-style ('${name}') attributes; only OOXML formatting attributes (e.g. w:val, w:color, w:sz) and data-astra-id are permitted.`
    );
  }
  if (name === 'xmlns' || prefix === 'xmlns' || prefix === 'xml') return;
  if (colon !== -1 && ALLOWED_OOXML_PREFIXES.has(prefix)) return;
  const ans = (attr as Attr).namespaceURI;
  if (ans && ALLOWED_OOXML_NAMESPACES.has(ans)) return;
  if (colon === -1 && ALLOWED_BARE_ATTRIBUTES.has(lower)) return;
  throw new InsertValidationError(
    `Non-allowlisted attribute '${name}' on tag '<${tagName}>' in XML insert`,
    'Only OOXML namespace attributes (e.g. w:val, w:color, xml:space), data-astra-id, and drawing id/name are permitted.'
  );
}

function validateOxmlElement(el: Element): void {
  const tagName = el.tagName || el.nodeName;
  const lowerTag = tagName.toLowerCase();
  const localName = el.localName || (tagName.includes(':') ? tagName.split(':')[1] : tagName);
  const lowerLocal = localName.toLowerCase();

  if (DANGEROUS_TAGS.has(lowerTag) || DANGEROUS_TAGS.has(lowerLocal)) {
    throw new InsertValidationError(
      `Disallowed or dangerous tag '<${tagName}>' in XML insert`,
      'Only valid OOXML WordprocessingML elements (e.g. w:p, w:tbl, w:r) are permitted.'
    );
  }

  const prefix = el.prefix;
  const ns = el.namespaceURI;

  if (ns && !ALLOWED_OOXML_NAMESPACES.has(ns)) {
    throw new InsertValidationError(
      `Non-OOXML namespace '${ns}' on tag '<${tagName}>' in XML insert`,
      'Elements must belong to valid OOXML namespaces (e.g. w:*, m:*, wp:*, a:*, pic:*, w14:*, w15:*).'
    );
  }

  if (!ns && (!prefix || !ALLOWED_OOXML_PREFIXES.has(prefix))) {
    throw new InsertValidationError(
      `Non-OOXML tag '<${tagName}>' detected in XML insert`,
      'Elements must belong to valid OOXML namespaces (e.g. w:*, m:*, wp:*, a:*, pic:*, w14:*, w15:*).'
    );
  }

  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    validateOxmlAttribute(el, attrs[i] as unknown as Attr);
  }
}

function validateAllOxmlElements(root: Element): void {
  validateOxmlElement(root);
  const descendants = root.getElementsByTagName('*');
  for (let i = 0; i < descendants.length; i++) {
    validateOxmlElement(descendants[i] as Element);
  }
}

// A table caption is a paragraph sibling whose pPr carries the Caption
// style (same marker operations/table.ts setCaption writes). Mirrored here
// so after-table inserts can land past a trailing caption sibling.
function isCaptionParagraph(n: Node | undefined): boolean {
  if (!n || n.type !== 'paragraph' || !n.xml) return false;
  const xml = n.xml as Element;
  if (typeof xml.getElementsByTagName !== 'function') return false;
  const pPr = xml.getElementsByTagName('w:pPr')[0] as Element | undefined;
  const pStyle = pPr?.getElementsByTagName('w:pStyle')[0] as Element | undefined;
  const v = pStyle?.getAttribute('w:val') ?? pStyle?.getAttribute('val');
  return v === 'Caption';
}

/**
 * After a section break: the break's sectPr becomes a copy of the section it
 * splits (page setup, columns, header/footer references, start settings), and
 * the requested break type moves to the following section, because w:type
 * says how its own section starts (ECMA-376 17.6.22). The following section
 * drops its header/footer references and title page, so it inherits them.
 */
function finishSectionBreak(breakP: Element): void {
  const newSectPr = breakP.getElementsByTagName('w:sectPr')[0];
  const governing = findGoverningSectPr(breakP);
  if (!newSectPr || !governing) return; // no later section: keep the minimal sectPr (not expected in a valid document)
  const breakVal = newSectPr.getElementsByTagName('w:type')[0]?.getAttribute('w:val') || 'nextPage';
  while (newSectPr.firstChild) newSectPr.removeChild(newSectPr.firstChild);
  for (const child of Array.from(governing.childNodes)) {
    if (child.nodeType === 1) newSectPr.appendChild(child.cloneNode(true));
  }
  upsertSectPrChild(governing, 'w:type').setAttribute('w:val', breakVal);
  for (const child of Array.from(governing.childNodes)) {
    const tag = child.nodeType === 1 ? (child as Element).tagName : '';
    if (tag === 'w:titlePg' || tag === 'w:headerReference' || tag === 'w:footerReference') {
      governing.removeChild(child);
    }
  }
  for (const child of Array.from(governing.childNodes)) {
    if (child.nodeType === 1 && (child as Element).tagName === 'w:pgNumType') {
      (child as Element).removeAttribute('w:start');
    }
  }
}

/** First w:sectPr after `from` in document order: paragraph-level in a later block, else body-level. */
function findGoverningSectPr(from: Element): Element | null {
  let sib: ChildNode | null = from.nextSibling;
  while (sib) {
    if (sib.nodeType === 1) {
      const el = sib as Element;
      if (el.tagName === 'w:sectPr') return el;
      if (el.tagName === 'w:p') {
        const hit = el.getElementsByTagName('w:sectPr')[0];
        if (hit) return hit;
      }
    }
    sib = sib.nextSibling;
  }
  return null;
}

interface ResolvedCellPadding {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

function validateTableCellPadding(val: unknown): ResolvedCellPadding {
  if (typeof val === 'number') {
    if (!Number.isInteger(val) || val < 0) {
      throw toError(
        'E_VALIDATION',
        `Invalid table cellPadding: expected non-negative integer dxa, got ${val}.`,
        'pass cellPadding as a non-negative integer (e.g. 120) or an object specifying sides { top, bottom, left, right }'
      );
    }
    return { top: val, bottom: val, left: val, right: val };
  }
  if (typeof val === 'object' && val !== null && !Array.isArray(val)) {
    const obj = val as Record<string, unknown>;
    const allowedSides = new Set(['top', 'bottom', 'left', 'right']);
    const keys = Object.keys(obj);
    for (const key of keys) {
      if (!allowedSides.has(key)) {
        throw toError(
          'E_VALIDATION',
          `Invalid table cellPadding: unknown side '${key}'.`,
          'use only valid sides: top, bottom, left, right with non-negative integer values'
        );
      }
    }
    const sides = ['top', 'bottom', 'left', 'right'] as const;
    let foundSide = false;
    const resolved: ResolvedCellPadding = {};
    for (const side of sides) {
      const sideVal = obj[side];
      if (sideVal !== undefined) {
        if (typeof sideVal !== 'number' || !Number.isInteger(sideVal) || sideVal < 0) {
          throw toError(
            'E_VALIDATION',
            `Invalid table cellPadding.${side}: expected non-negative integer dxa, got ${typeof sideVal === 'number' ? sideVal : JSON.stringify(sideVal)}.`,
            `pass ${side} as a non-negative integer (e.g. 120)`
          );
        }
        resolved[side] = sideVal;
        foundSide = true;
      }
    }
    if (!foundSide) {
      throw toError(
        'E_VALIDATION',
        'Invalid table cellPadding: object must contain at least one valid side (top, bottom, left, right).',
        'specify at least one side with a non-negative integer value, e.g. { top: 120, bottom: 120, left: 160, right: 160 }'
      );
    }
    return resolved;
  }
  throw toError(
    'E_VALIDATION',
    `Invalid table cellPadding: expected non-negative integer or object, got ${typeof val === 'string' ? `"${val}"` : JSON.stringify(val)}.`,
    'pass cellPadding as a non-negative integer (e.g. 120) or an object specifying sides { top, bottom, left, right }'
  );
}

export async function executeInsert(op: InsertOperation, ast: DocumentAST, parser?: DocxParser, guide?: StyleGuideConfig): Promise<void> {  let anchorSection: Section | null = null;
  let anchorIndex = -1;
  let anchorNode: Node | undefined;

  // Destination store resolves first: bad targets fail closed before any
  // lookup or mutation. Accepts flat or params-nested values so the strict
  // schema gate (which allowlists params content) can carry them.
  const hfParams = (op.params ?? {}) as Record<string, unknown>;
  // Nested params win on conflict (params is the deliberate channel);
  // mirrors the field pick pattern.
  const hfPick = (key: string): unknown =>
    hfParams[key] !== undefined ? hfParams[key] : op[key];
  // insertType is op-derived only, so hoisting it here (ahead of anchor
  // resolution) changes nothing for the body path.
  const insertType = ((op.type as string) === 'insert' || !op.type)
    ? (op.insertType || op.paragraphType || op.nodeType || (op.headers || op.rows ? 'table' : op.xml ? 'xml' : 'paragraph'))
    : (op.type || 'paragraph');
  const hfTarget = resolveHeaderFooterTarget('insert', hfPick('target'), hfPick('pageType'), hfPick('sectionIndex'));
  const toHeaderFooter = hfTarget.target === 'header' || hfTarget.target === 'footer';
  // Only paragraph-rooted content (paragraph, divider, horizontalRule all
  // build w:p roots) can append to a header/footer part. Tables, raw XML,
  // and breaks fail closed with a working alternative instead of landing
  // half-supported markup in a repeating part.
  if (toHeaderFooter) {
    if (insertType !== 'paragraph' && insertType !== 'divider' && insertType !== 'horizontalRule') {
      throw new Error(
        `E_VALIDATION: insert type '${String(insertType)}' cannot target '${hfTarget.target}' (only paragraph-style content: paragraph, divider, horizontalRule). ` +
          `Remedy: omit target (or pass target:'body') for '${String(insertType)}', or use { action:'headerFooter', operation:'set', targetType:'${hfTarget.target}', xmlContent:'...' } for full-part content.`
      );
    }
    if (!parser) {
      throw new Error(
        `E_VALIDATION: insert to ${hfTarget.target} needs a parser-attached session (parts live in the document package). ` +
          `Remedy: run against an opened document (open_document first) instead of a headless AST, or use target:'body'.`
      );
    }
  }

  if (!toHeaderFooter) {
    for (const section of ast.sections) {
      anchorIndex = section.nodes.findIndex(n => n.id === op.anchor);
      if (anchorIndex !== -1) {
        anchorSection = section;
        anchorNode = section.nodes[anchorIndex];
        break;
      }
    }

    // Fallback: only a missing/'default' anchor may resolve to the last node;
    // an explicit but unknown/stale anchor must fail loudly so callers self-correct.
    // The end of the document is the last block of the last section that has one.
    if (!anchorNode && (!op.anchor || op.anchor === 'default')) {
      const lastWithBlocks = [...ast.sections].reverse().find((section) => section.nodes.length > 0);
      if (lastWithBlocks) {
        anchorSection = lastWithBlocks;
        anchorIndex = anchorSection.nodes.length - 1;
        anchorNode = anchorSection.nodes[anchorIndex];
        op.position = 'after';
      }
    }
  }

  if (!anchorNode && !toHeaderFooter) {
    throw new Error(
      `Anchor node '${op.anchor}' not found in document. ` +
      `Re-read the current structure with inspect_document and retry with a valid node id.`
    );
  }

  // Body-path guarantee for the branches below: the throw above leaves
  // anchorNode defined whenever !toHeaderFooter, so this alias is the
  // defined node on every body path (header/footer paths never read it).
  const bodyAnchorNode = anchorNode as Node;
  // Same guarantee for the section: anchorSection is assigned everywhere
  // anchorNode is, so it is defined on every body path below.
  const bodyAnchorSection = anchorSection as Section;

  if (anchorNode) {
    validateVersion(anchorNode, op.expectedVersion);
  }

  // Header/footer inserts ignore anchor/position (append in call order);
  // their owner document is the live package DOM.
  const doc = anchorNode ? anchorNode.xml.ownerDocument : parser!.getDocumentDom();
  if (!doc) throw new Error("Document not found on anchor node");

  const astNodeIds = new Set<string>();
  for (const section of ast.sections) {
    for (const n of section.nodes) {
      astNodeIds.add(n.id);
    }
  }

  // Create new XML
  const newId = op.id || mintInsertId(deriveInsertSeed(op), astNodeIds, '_bk_new'); // Generate unique ID or use provided
  // insertType was hoisted above anchor resolution (op-derived only).

  // Explicit per-op revision flag only: the document-level tracking setting
  // does not alter inserts (a fresh paragraph has no baseline to diff, so
  // silent creation stays the default unless the caller opts in here).
  // Non-boolean values fail closed before any mutation; true on anything
  // but a paragraph insert fails closed (no other type emits w:ins).
  const rawTrack = op.trackChanges;
  if (rawTrack !== undefined && rawTrack !== true && rawTrack !== false) {
    throw new Error(
      `E_VALIDATION: insert trackChanges must be a boolean (got '${String(rawTrack)}'). ` +
        `Remedy: pass trackChanges:true to wrap the inserted runs in w:ins, or omit it for a silent insert.`
    );
  }
  const trackOn = rawTrack === true;
  if (trackOn && insertType !== 'paragraph') {
    throw new Error(
      `E_VALIDATION: insert trackChanges:true is supported only for type:'paragraph' (got '${String(insertType)}'). ` +
        `Remedy: insert a paragraph with trackChanges:true, or omit trackChanges for other types.`
    );
  }

  interface CreatedElementInfo {
    xml: Element;
    id: string;
    type: 'paragraph' | 'table';
  }
  const createdElements: CreatedElementInfo[] = [];

  // True when the paragraph path wrapped its runs in w:ins (drives the
  // metadata flag below so it never claims markup that is not in the XML).
  let trackedWrapped = false;

  // Code-bookmark id minted by the paragraph path for fenced code (null for
  // every other insert type); drives the hasBookmark metadata flag below.
  let codeBookmarkId: string | null = null;

  if (insertType === 'paragraph') {
    const newNodeXml = createWordElement(doc, 'w:p');
    newNodeXml.setAttribute('data-astra-id', newId);
    // Code blocks (fenced code via compose, or explicit params.codeBlock):
    // text runs carry the template Code character style (Consolas/indigo/
    // 10pt) with direct formatting mirroring it, and the paragraph gets
    // block shading, a box border, and an indent. A fenced language rides in
    // a bookmark (_AstraCode_<lang>) so exporters rebuild the fence with its
    // tag instead of silently dropping it.
    const codeBag = (op.params ?? {}) as Record<string, unknown>;
    const codeLangRaw = codeBag.codeLanguage;
    const isCodeBlock = codeBag.codeBlock === true || typeof codeLangRaw === 'string';
    const codeLang = typeof codeLangRaw === 'string'
      ? codeLangRaw.replace(/[^A-Za-z0-9_]/g, '').slice(0, 24)
      : '';
    if (isCodeBlock) {
      const pPr = createWordElement(doc, 'w:pPr');
      const shd = createWordElement(doc, 'w:shd');
      shd.setAttribute('w:val', 'clear');
      shd.setAttribute('w:color', 'auto');
      // Guide skin (explicit op colors do not exist on this path — the guide
      // fills the omitted values, else the legacy grays stay byte-identical).
      shd.setAttribute('w:fill', (guide?.codeBlocks?.background ?? 'F2F2F2').replace(/^#/, '').toUpperCase());
      pPr.appendChild(shd);
      const pBdr = createWordElement(doc, 'w:pBdr');
      for (const side of ['top', 'left', 'bottom', 'right']) {
        const edge = createWordElement(doc, `w:${side}`);
        edge.setAttribute('w:val', 'single');
        edge.setAttribute('w:sz', '6');
        edge.setAttribute('w:space', '4');
        edge.setAttribute('w:color', (guide?.codeBlocks?.borderColor ?? 'D1D5DB').replace(/^#/, '').toUpperCase());
        pBdr.appendChild(edge);
      }
      pPr.appendChild(pBdr);
      const spacing = createWordElement(doc, 'w:spacing');
      spacing.setAttribute('w:before', '120');
      spacing.setAttribute('w:after', '120');
      pPr.appendChild(spacing);
      const ind = createWordElement(doc, 'w:ind');
      ind.setAttribute('w:left', '180');
      pPr.appendChild(ind);
      sortChildrenBySchema(pPr, PPR_CHILD_ORDER);
      newNodeXml.appendChild(pPr);
    }
    if (isCodeBlock && codeLang) {
      // Canonical allocator (scan + process-wide high-water reservation):
      // every code bookmark id is unique even for pre-insert batching.
      codeBookmarkId = String(allocateNextBookmarkId(doc as unknown as Document));
      const bmStart = createWordElement(doc, 'w:bookmarkStart');
      bmStart.setAttribute('w:id', codeBookmarkId);
      bmStart.setAttribute('w:name', `_AstraCode_${codeLang}`);
      newNodeXml.appendChild(bmStart);
    }
    if (op.text) {
      // Embedded newlines become break-separated runs (one w:r per line,
      // w:br between lines) so multi-line text keeps its line breaks in
      // Word and in the exporters instead of collapsing into one w:t.
      const lines = op.text.split(/\r\n|\r|\n/);
      const contentRuns: Element[] = [];
      for (let li = 0; li < lines.length; li++) {
        if (li > 0) {
          const brRun = createWordElement(doc, 'w:r');
          brRun.appendChild(createWordElement(doc, 'w:br'));
          contentRuns.push(brRun);
        }
        if (lines[li].length > 0) {
          const run = createWordElement(doc, 'w:r');
          if (isCodeBlock) {
            const rPr = createWordElement(doc, 'w:rPr');
            const rStyle = createWordElement(doc, 'w:rStyle');
            rStyle.setAttribute('w:val', 'Code');
            rPr.appendChild(rStyle);
            const rFonts = createWordElement(doc, 'w:rFonts');
            rFonts.setAttribute('w:ascii', 'Consolas');
            rFonts.setAttribute('w:hAnsi', 'Consolas');
            rPr.appendChild(rFonts);
            const color = createWordElement(doc, 'w:color');
            color.setAttribute('w:val', '4338CA');
            rPr.appendChild(color);
            const sz = createWordElement(doc, 'w:sz');
            sz.setAttribute('w:val', '20');
            rPr.appendChild(sz);
            const szCs = createWordElement(doc, 'w:szCs');
            szCs.setAttribute('w:val', '20');
            rPr.appendChild(szCs);
            sortChildrenBySchema(rPr, RPR_CHILD_ORDER);
            run.appendChild(rPr);
          }
          run.appendChild(createTextElement(doc, lines[li]));
          contentRuns.push(run);
        }
      }
      if (trackOn) {
        const ins = createWordElement(doc, 'w:ins');
        const stamp = resolveRevisionStamp(op.author, op.date, ast.settings.defaultAuthor);
        ins.setAttribute('w:author', stamp.author);
        ins.setAttribute('w:date', stamp.date);
        ins.setAttribute('w:id', String(ast.nextChangeId++));
        for (const r of contentRuns) ins.appendChild(r);
        if (contentRuns.length > 0) {
          newNodeXml.appendChild(ins);
          trackedWrapped = true;
        }
      } else {
        for (const r of contentRuns) newNodeXml.appendChild(r);
      }
    }
    if (codeBookmarkId !== null) {
      const bmEnd = createWordElement(doc, 'w:bookmarkEnd');
      bmEnd.setAttribute('w:id', codeBookmarkId);
      newNodeXml.appendChild(bmEnd);
    }
    createdElements.push({ xml: newNodeXml, id: newId, type: 'paragraph' });
  } else if (insertType === 'pageBreak') {
    const newNodeXml = createWordElement(doc, 'w:p');
    newNodeXml.setAttribute('data-astra-id', newId);
    const run = createWordElement(doc, 'w:r');
    const br = createWordElement(doc, 'w:br');
    br.setAttribute('w:type', 'page');
    run.appendChild(br);
    newNodeXml.appendChild(run);
    createdElements.push({ xml: newNodeXml, id: newId, type: 'paragraph' });
  } else if (insertType === 'sectionBreak') {
    // B-2: a section break is a paragraph whose pPr carries the w:sectPr
    // delimiter (ECMA-376 §17.6.18: w:type inside sectPr selects the break).
    // Minimal sectPr (type only); finishSectionBreak completes the sectPr after placement.
    // Shape note: the strict insertOp schema carries top-level breakType only
    // (no params key — params would be rejected at the schema gate); the
    // params.breakType fallback below is honored for direct-engine calls only.
    const VALID_BREAK_TYPES = new Set(['nextPage', 'nextColumn', 'continuous', 'evenPage', 'oddPage']);
    const rawBreak = op.breakType ?? op.params?.breakType;
    const breakVal = String(rawBreak ?? 'nextPage');
    if (!VALID_BREAK_TYPES.has(breakVal)) {
      throw new Error(
        `E_VALIDATION: invalid breakType '${breakVal}'. ` +
          `Allowed: nextPage|nextColumn|continuous|evenPage|oddPage (ST_SectionMark). ` +
          `Remedy: pass { action:'insert', type:'sectionBreak', anchor, position, breakType:'nextPage' } with one of the allowed values.`
      );
    }
    const newNodeXml = createWordElement(doc, 'w:p');
    newNodeXml.setAttribute('data-astra-id', newId);
    const pPr = createWordElement(doc, 'w:pPr');
    const sectPr = createWordElement(doc, 'w:sectPr');
    const typeEl = createWordElement(doc, 'w:type');
    typeEl.setAttribute('w:val', breakVal);
    sectPr.appendChild(typeEl);
    pPr.appendChild(sectPr);
    newNodeXml.appendChild(pPr);
    createdElements.push({ xml: newNodeXml, id: newId, type: 'paragraph' });
  } else if (insertType === 'divider' || insertType === 'horizontalRule') {
    const newNodeXml = createWordElement(doc, 'w:p');
    newNodeXml.setAttribute('data-astra-id', newId);
    const pPr = createWordElement(doc, 'w:pPr');
    const pBdr = createWordElement(doc, 'w:pBdr');
    const bottom = createWordElement(doc, 'w:bottom');

    const lineStyle = op.lineStyle || 'single';
    const color = (op.color || 'CBD5E1').replace(/^#/, '');
    const thickness = op.thickness !== undefined ? String(op.thickness) : '12';

    bottom.setAttribute('w:val', lineStyle);
    bottom.setAttribute('w:sz', thickness);
    bottom.setAttribute('w:space', '1');
    bottom.setAttribute('w:color', color);
    pBdr.appendChild(bottom);
    pPr.appendChild(pBdr);

    if (op.spacingBefore !== undefined || op.spacingAfter !== undefined) {
      const sp = createWordElement(doc, 'w:spacing');
      if (op.spacingBefore !== undefined) sp.setAttribute('w:before', String(requireTwips('spacingBefore', op.spacingBefore)));
      if (op.spacingAfter !== undefined) sp.setAttribute('w:after', String(requireTwips('spacingAfter', op.spacingAfter)));
      pPr.appendChild(sp);
    }
    newNodeXml.appendChild(pPr);
    createdElements.push({ xml: newNodeXml, id: newId, type: 'paragraph' });
  } else if (insertType === 'table') {
    // R-42: fail closed on bad caption args before mutating the document,
    // so a rejected caption never leaves a half-captioned table behind.
    // Creation itself is delegated to setCaption below (single honor site).
    const insertCapPosRaw = op.captionPosition;
    if (insertCapPosRaw !== undefined && String(insertCapPosRaw) !== 'above' && String(insertCapPosRaw) !== 'below') {
      throw new Error(
        `E_VALIDATION: insert table captionPosition must be 'above'|'below' (got '${String(insertCapPosRaw)}'). ` +
          `Remedy: pass captionPosition:'above' (default) or captionPosition:'below' with captionText; run \`astra inspect\` to list anchor ids.`
      );
    }
    const insertCapTextRaw = op.captionText;
    if (insertCapTextRaw !== undefined && String(insertCapTextRaw).trim() === '') {
      throw new Error(
        `E_VALIDATION: insert table captionText must be a non-empty string (got '${String(insertCapTextRaw)}'). ` +
          `Remedy: pass captionText as a non-empty string (e.g. captionText:'Table 1: Results'); run \`astra inspect\` to list anchor ids.`
      );
    }
    const newNodeXml = createWordElement(doc, 'w:tbl');
    newNodeXml.setAttribute('data-astra-id', newId);
    const tblPr = createWordElement(doc, 'w:tblPr');
    const tblBorders = createWordElement(doc, 'w:tblBorders');
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(border => {
      const b = createWordElement(doc, `w:${border}`);
      b.setAttribute('w:val', 'single');
      b.setAttribute('w:sz', '4');
      b.setAttribute('w:space', '0');
      b.setAttribute('w:color', 'auto');
      tblBorders.appendChild(b);
    });
    const tblStyle = createWordElement(doc, 'w:tblStyle');
    tblStyle.setAttribute('w:val', 'TableGrid');
    tblPr.appendChild(tblStyle);
    tblPr.appendChild(tblBorders);

    const flatPaddingRaw = op.cellPadding;
    const nestedPaddingRaw = (op.params as Record<string, unknown> | undefined)?.cellPadding;

    if (flatPaddingRaw !== undefined) {
      validateTableCellPadding(flatPaddingRaw);
    }
    if (nestedPaddingRaw !== undefined) {
      validateTableCellPadding(nestedPaddingRaw);
    }

    const cellPaddingRaw = nestedPaddingRaw !== undefined ? nestedPaddingRaw : flatPaddingRaw;
    if (cellPaddingRaw !== undefined) {
      const pad = validateTableCellPadding(cellPaddingRaw);
      const tblCellMar = createWordElement(doc, 'w:tblCellMar');
      if (pad.top !== undefined) {
        const topEl = createWordElement(doc, 'w:top');
        topEl.setAttribute('w:w', String(pad.top));
        topEl.setAttribute('w:type', 'dxa');
        tblCellMar.appendChild(topEl);
      }
      if (pad.left !== undefined) {
        const leftEl = createWordElement(doc, 'w:left');
        leftEl.setAttribute('w:w', String(pad.left));
        leftEl.setAttribute('w:type', 'dxa');
        tblCellMar.appendChild(leftEl);
      }
      if (pad.bottom !== undefined) {
        const bottomEl = createWordElement(doc, 'w:bottom');
        bottomEl.setAttribute('w:w', String(pad.bottom));
        bottomEl.setAttribute('w:type', 'dxa');
        tblCellMar.appendChild(bottomEl);
      }
      if (pad.right !== undefined) {
        const rightEl = createWordElement(doc, 'w:right');
        rightEl.setAttribute('w:w', String(pad.right));
        rightEl.setAttribute('w:type', 'dxa');
        tblCellMar.appendChild(rightEl);
      }

      if (tblCellMar.childNodes.length > 0) {
        tblPr.appendChild(tblCellMar);
        sortChildrenBySchema(tblPr, TBLPR_CHILD_ORDER);
      }
    }

    newNodeXml.appendChild(tblPr);

    const defaultHeaders = op.headers || ['Column 1', 'Column 2'];
    const defaultRows = op.rows || [['Data 1', 'Data 2']];

    // ECMA-376 requires <w:tblGrid> between <w:tblPr> and <w:tr>
    const colCount = Math.max(defaultHeaders.length, defaultRows.length > 0 ? defaultRows[0].length : 1);
    const tblGrid = createWordElement(doc, 'w:tblGrid');
    const colWidth = Math.round(9360 / colCount);
    for (let c = 0; c < colCount; c++) {
      const gridCol = createWordElement(doc, 'w:gridCol');
      gridCol.setAttribute('w:w', String(colWidth));
      tblGrid.appendChild(gridCol);
    }
    newNodeXml.appendChild(tblGrid);

    // Header Row (B-6: repeat on every page + never split across pages)
    const headerTr = createWordElement(doc, 'w:tr');
    const headerTrPr = createWordElement(doc, 'w:trPr');
    headerTrPr.appendChild(createWordElement(doc, 'w:cantSplit'));
    headerTrPr.appendChild(createWordElement(doc, 'w:tblHeader'));
    headerTr.appendChild(headerTrPr);
    const nestedHeaderBold = (op.params as Record<string, unknown> | undefined)?.headerBold;
    const explicitHeaderBold = nestedHeaderBold !== undefined ? nestedHeaderBold : op.headerBold;
    if (explicitHeaderBold !== undefined && typeof explicitHeaderBold !== 'boolean') {
      throw new Error(
        `E_VALIDATION: insert table headerBold must be a boolean (got '${String(explicitHeaderBold)}'). ` +
          `Remedy: pass headerBold:true or headerBold:false, or omit it to follow the style guide.`
      );
    }
    const headerBold = typeof explicitHeaderBold === 'boolean' ? explicitHeaderBold : guide?.tables?.headerBold !== false;
    defaultHeaders.forEach(h => {
      const tc = createWordElement(doc, 'w:tc');
      const p = createWordElement(doc, 'w:p');
      const r = createWordElement(doc, 'w:r');
      if (headerBold) {
        const rPr = createWordElement(doc, 'w:rPr');
        rPr.appendChild(createWordElement(doc, 'w:b'));
        r.appendChild(rPr);
      }
      r.appendChild(createTextElement(doc, String(h)));
      p.appendChild(r);
      tc.appendChild(p);
      headerTr.appendChild(tc);
    });
    newNodeXml.appendChild(headerTr);

    // Data Rows (B-6: never split a row across pages)
    defaultRows.forEach(rowCells => {
      const tr = createWordElement(doc, 'w:tr');
      const trPr = createWordElement(doc, 'w:trPr');
      trPr.appendChild(createWordElement(doc, 'w:cantSplit'));
      tr.appendChild(trPr);
      rowCells.forEach(cellText => {
        const tc = createWordElement(doc, 'w:tc');
        const p = createWordElement(doc, 'w:p');
        const r = createWordElement(doc, 'w:r');
        r.appendChild(createTextElement(doc, String(cellText)));
        p.appendChild(r);
        tc.appendChild(p);
        tr.appendChild(tc);
      });
      newNodeXml.appendChild(tr);
    });
    // Keep-together chaining (explicit > guide > legacy): cantSplit +
    // tblHeader above are unconditional legacy markers (kept as-is); an
    // effective keepTogether additionally chains every row but the last with
    // keepNext so the table lays out as one block. Absent both, no keepNext
    // is written and the bytes match the legacy path exactly.
    {
      const nestedKeep = (op.params as Record<string, unknown> | undefined)?.keepTogether;
      const explicitKeep = nestedKeep !== undefined ? nestedKeep : op.keepTogether;
      if (explicitKeep !== undefined && typeof explicitKeep !== 'boolean') {
        throw new Error(
          `E_VALIDATION: insert table keepTogether must be a boolean (got '${String(explicitKeep)}'). ` +
            `Remedy: pass keepTogether:true to chain rows with keepNext, or omit it for legacy markers.`
        );
      }
      const effectiveKeep = typeof explicitKeep === 'boolean'
        ? explicitKeep
        : guide?.tables?.keepTogether === true;
      if (effectiveKeep) {
        const builtRows = Array.from(newNodeXml.childNodes).filter(
          (k) => (k as Element).nodeType === 1 && (k as Element).tagName === 'w:tr',
        ) as Element[];
        builtRows.forEach((tr, rIdx) => {
          if (rIdx === builtRows.length - 1) return;
          for (const tc of Array.from(tr.childNodes).filter(
            (k) => (k as Element).nodeType === 1 && (k as Element).tagName === 'w:tc',
          ) as Element[]) {
            let ps = Array.from(tc.childNodes).filter(
              (k) => (k as Element).nodeType === 1 && (k as Element).tagName === 'w:p',
            ) as Element[];
            if (ps.length === 0) {
              const p = createWordElement(doc, 'w:p');
              tc.appendChild(p);
              ps = [p];
            }
            for (const p of ps) {
              let pPr = Array.from(p.childNodes).find(
                (k) => (k as Element).nodeType === 1 && (k as Element).tagName === 'w:pPr',
              ) as Element | undefined;
              if (!pPr) {
                pPr = createWordElement(doc, 'w:pPr');
                p.insertBefore(pPr, p.firstChild);
              }
              if (!pPr.getElementsByTagName('w:keepNext')[0]) {
                pPr.appendChild(createWordElement(doc, 'w:keepNext'));
                sortChildrenBySchema(pPr, PPR_CHILD_ORDER);
              }
            }
          }
        });
      }
    }
    createdElements.push({ xml: newNodeXml, id: newId, type: 'table' });
  } else if (insertType === 'xml') {
    if (!op.xml || !op.xml.trim()) {
      throw new InsertValidationError(
        "Insert operation of type 'xml' requires a non-empty 'xml' string property",
        "Provide a valid OOXML XML fragment string in the 'xml' property."
      );
    }
    const xmlContent = op.xml.trim();
    // Single detector for the whole codebase: route through the central
    // XXE assert. assertNoXxe throws a plain Error, so re-wrap into the
    // coded fail-closed contract this operation guarantees.
    try {
      assertNoXxe(xmlContent, 'insert xml fragment');
    } catch {
      throw new InsertValidationError(
        'Security Error: XXE / DTD entity declarations are strictly prohibited',
        'Remove <!DOCTYPE and <!ENTITY declarations from the XML content.'
      );
    }
    // Strict well-formedness pre-check: the lenient parser silently repairs
    // a missing final '>' (e.g. '...</w:p' becomes '...</w:p>') with no
    // warning. Truncated input never ends exactly at a tag boundary close.
    if (!xmlContent.endsWith('>')) {
      throw new InsertValidationError(
        'Malformed or truncated XML (input does not end with a closing ">")',
        'Ensure the XML fragment is well-formed with matching opening and closing tags and valid OOXML structure.'
      );
    }
    const wrappedXml = `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
      `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ` +
      `xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ` +
      `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
      `xmlns:v="urn:schemas-microsoft-com:vml" ` +
      `xmlns:o="urn:schemas-microsoft-com:office:office" ` +
      `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">${xmlContent}</root>`;

    let parseErrorMsg = '';
    let parsedDoc: Document | null = null;
    try {
      parsedDoc = new DOMParser({
        errorHandler: {
          warning: (w: unknown) => {
            if (!parseErrorMsg) parseErrorMsg = String(w);
          },
          error: (e: unknown) => {
            if (!parseErrorMsg) parseErrorMsg = String(e);
          },
          fatalError: (f: unknown) => {
            if (!parseErrorMsg) parseErrorMsg = String(f);
          }
        }
      }).parseFromString(wrappedXml, 'text/xml');
    } catch (e: unknown) {
      if (!parseErrorMsg) parseErrorMsg = e instanceof Error ? e.message : String(e);
    }

    if (parseErrorMsg) {
      throw new InsertValidationError(
        `Malformed or truncated XML (${parseErrorMsg.trim().replace(/\r?\n|\r/g, ' ')})`,
        'Ensure the XML fragment is well-formed with matching opening and closing tags and valid OOXML structure.'
      );
    }

    if (!parsedDoc || !parsedDoc.documentElement) {
      throw new InsertValidationError(
        'Failed to parse XML fragment (document is empty or invalid)',
        'Ensure the XML fragment is well-formed with matching opening and closing tags and valid OOXML structure.'
      );
    }

    const parseErrors = parsedDoc.getElementsByTagName('parsererror');
    if (parseErrors.length > 0) {
      const detail = parseErrors[0].textContent?.trim().replace(/\r?\n|\r/g, ' ') || 'XML parser error';
      throw new InsertValidationError(
        `Malformed XML in insert operation (${detail})`,
        'Ensure the XML fragment is well-formed with matching opening and closing tags and valid OOXML structure.'
      );
    }

    // Parser-normalization net: the lenient parser folds '<!doctype' (any
    // case) into a real DocumentType node without emitting any error, so a
    // declaration smuggled past the raw guard would otherwise be dropped
    // silently here while its sibling payload is imported. Any surviving
    // doctype/entity node fails closed.
    const hasDeclarationNode = (() => {
      // Duck-typed DOM walk: parsedDoc children are xmldom nodes, which the
      // imported AST Node type does not describe (no nodeType/childNodes).
      interface DomWalkNode {
        nodeType?: number;
        childNodes?: ArrayLike<DomWalkNode> | null;
      }
      const stack: DomWalkNode[] = [parsedDoc as unknown as DomWalkNode];
      while (stack.length > 0) {
        const cur = stack.pop();
        if (!cur) continue;
        if (cur.nodeType === 10 || cur.nodeType === 5) return true;
        const kids = cur.childNodes;
        if (kids) {
          for (let i = 0; i < kids.length; i++) stack.push(kids[i]);
        }
      }
      return false;
    })();
    if ((parsedDoc as Document).doctype || hasDeclarationNode) {
      throw new InsertValidationError(
        'Security Error: XXE / DTD entity declarations are strictly prohibited',
        'Remove <!DOCTYPE and <!ENTITY declarations from the XML content.'
      );
    }

    const rootElements: Element[] = [];
    const childNodes = parsedDoc.documentElement.childNodes;
    for (let i = 0; i < childNodes.length; i++) {
      const child = childNodes[i];
      if (child.nodeType === 1) { // Node.ELEMENT_NODE
        rootElements.push(child as Element);
      } else if (child.nodeType === 3 && child.nodeValue && child.nodeValue.trim().length > 0) {
        throw new InsertValidationError(
          `Malformed XML fragment: non-whitespace text found outside elements ('${child.nodeValue.trim().slice(0, 50)}')`,
          'Wrap all text in valid OOXML elements such as <w:p><w:r><w:t>...</w:t></w:r></w:p>.'
        );
      }
    }

    if (rootElements.length === 0) {
      throw new InsertValidationError(
        `No valid XML element found in provided 'xml' string. Provided: ${xmlContent.substring(0, 100)}`,
        'Provide at least one valid OOXML element (e.g. <w:p>...</w:p>).'
      );
    }

    // P0-13: Validate all root elements and their descendants belong to valid OOXML namespaces
    for (const el of rootElements) {
      validateAllOxmlElements(el);
    }

    for (let k = 0; k < rootElements.length; k++) {
      const rootEl = rootElements[k];
      const imported = doc.importNode(rootEl, true) as Element;
      let elemId: string;
      if (k === 0 && op.id) {
        elemId = op.id;
      } else {
        const existingId = imported.getAttribute('data-astra-id');
        if (existingId && !astNodeIds.has(existingId)) {
          elemId = existingId;
        } else if (op.id) {
          elemId = `${op.id}_${k}`;
        } else {
          elemId = mintInsertId(`${deriveInsertSeed(op)}:xml:${k}`, astNodeIds, '_bk_new');
        }
      }
      imported.setAttribute('data-astra-id', elemId);
      astNodeIds.add(elemId);
      const isTable = imported.tagName === 'w:tbl' || imported.localName === 'tbl';
      createdElements.push({
        xml: imported,
        id: elemId,
        type: isTable ? 'table' : 'paragraph'
      });
    }
  } else {
    throw new Error(
      `Unsupported insert type '${insertType}'. ` +
      `Supported: paragraph, pageBreak, sectionBreak, table, xml, divider, horizontalRule. ` +
      `Re-read the current structure with inspect_document and retry with a supported type.`
    );
  }

  // Every insert root (paragraph-style w:p roots, table-cell w:p
  // descendants, caller-supplied xml fragments) leaves here with valid
  // w14:paraIds already stamped; caller-supplied valid ids are preserved.
  const usedParaIds = new Set<string>();
  const allParas = doc.getElementsByTagName('w:p');
  for (let i = 0; i < allParas.length; i++) {
    const v = (allParas[i] as unknown as Element).getAttribute?.('w14:paraId');
    if (v && isValidParaId(v)) usedParaIds.add(v.toUpperCase());
  }

  const baseSeed = deriveInsertSeed(op);
  for (let k = 0; k < createdElements.length; k++) {
    const item = createdElements[k];
    const seed = k === 0 ? baseSeed : `${baseSeed}:${k}`;
    stampMissingParaIds(item.xml, doc, seed, usedParaIds);
  }

  // Create AST Nodes
  const createdNodes: Node[] = [];
  for (let k = 0; k < createdElements.length; k++) {
    const item = createdElements[k];
    const xmlStr = new XMLSerializer().serializeToString(item.xml);
    const text = (insertType === 'paragraph' && op.text)
      ? op.text
      : (item.xml.textContent || '');
    const newNode: Node = {
      id: item.id,
      type: item.type,
      xml: item.xml,
      metadata: {
        text,
        headers: op.headers || [],
        rows: op.rows || [],
        hasCitation: false,
        hasComment: false,
        hasBookmark: codeBookmarkId !== null,
        hasTrackedChanges: trackedWrapped,
        hasEquation: false,
        hasField: false,
        hasImage: false,
        hasHyperlink: false
      },
      version: 1,
      checksum: await calculateNodeChecksum(xmlStr)
    };
    createdNodes.push(newNode);
  }

  // The section the new blocks end up in: the anchor's, unless they go after
  // that section's break paragraph, which puts them in the next one.
  let placedSection: Section | undefined = anchorSection ?? undefined;

  // Insert into DOM and AST. Header/footer targets append to the part
  // (call order = part order) and never touch body sections or the anchor
  // chain; the body path below is unchanged.
  if (toHeaderFooter) {
    appendElementsToHeaderFooter(
      ast,
      parser,
      createdElements.map((item) => item.xml),
      { targetType: hfTarget.target as 'header' | 'footer', pageType: hfTarget.pageType, sectionIndex: hfTarget.sectionIndex }
    );
  } else if (op.position === 'before') {
    for (const node of createdNodes) {
      bodyAnchorNode.xml.parentNode?.insertBefore(node.xml, bodyAnchorNode.xml);
    }
    bodyAnchorSection.nodes.splice(anchorIndex, 0, ...createdNodes);
  } else {
    // Same-anchor 'after' inserts preserve FIFO order via a per-anchor
    // chain. The chain is re-validated against the live DOM/AST at every
    // use: a deleted tail (or middle link) drops out, so the next insert
    // lands after the latest still-alive link instead of a detached ghost
    // (which would misplace it directly after the anchor). A legacy single
    // '_lastInsertAfter' pointer is adopted into the chain when present.
    const anchorChain = bodyAnchorNode as unknown as InsertChainHost;
    const rawChain: Node[] = Array.isArray(anchorChain._insertAfterChain)
      ? anchorChain._insertAfterChain as Node[]
      : [];
    const legacyTail = anchorChain._lastInsertAfter;
    if (rawChain.length === 0 && legacyTail) {
      rawChain.push(legacyTail);
    }
    const liveChain = rawChain.filter((n) => {
      if (!n || !n.xml || !(n.xml as Element).parentNode) return false;
      return ast.sections.some((s: Section) => s.nodes.some((x: Node) => x.id === n.id));
    });
    anchorChain._insertAfterChain = liveChain;
    if (legacyTail && !liveChain.some((n) => n.id === legacyTail.id)) {
      delete anchorChain._lastInsertAfter;
    }
    const effectiveAnchor: Node = liveChain.length > 0 ? liveChain[liveChain.length - 1] : bodyAnchorNode;

    // A table followed by its below-caption sibling must not be split: an
    // 'after' insert anchored on (or chained behind) the table lands past
    // the caption paragraph, in both DOM and AST order.
    let captionSkipped = false;
    let prevXml: Element = effectiveAnchor.xml;
    if ((effectiveAnchor.xml as Element)?.tagName === 'w:tbl') {
      const hostSection: Section | undefined = ast.sections.find((s: Section) => s.nodes.some((n: Node) => n.id === effectiveAnchor.id));
      const hostIdx: number = hostSection ? hostSection.nodes.findIndex((n: Node) => n.id === effectiveAnchor.id) : -1;
      const nextNode: Node | undefined = (hostIdx !== -1 && hostSection) ? hostSection.nodes[hostIdx + 1] : undefined;
      if (nextNode && isCaptionParagraph(nextNode)) {
        // DOM adjacency guard: skip only when the caption is the immediate
        // element sibling of the table in the live document.
        let sib = (effectiveAnchor.xml as Element).nextSibling;
        while (sib && sib.nodeType !== 1) sib = sib.nextSibling;
        if (sib === (nextNode.xml as Element)) {
          prevXml = nextNode.xml as Element;
          captionSkipped = true;
        }
      }
    }
    for (const node of createdNodes) {
      prevXml.parentNode?.insertBefore(node.xml, prevXml.nextSibling);
      prevXml = node.xml;
    }

    let effSection: Section = bodyAnchorSection;
    let effIdx = bodyAnchorSection.nodes.findIndex((n: Node) => n.id === effectiveAnchor.id);
    if (effIdx === -1) {
      for (const s of ast.sections) {
        const i = s.nodes.findIndex((n: Node) => n.id === effectiveAnchor.id);
        if (i !== -1) {
          effSection = s;
          effIdx = i;
          break;
        }
      }
    }
    const insertIndex = (effIdx !== -1 ? effIdx + 1 : anchorIndex + 1) + (captionSkipped ? 1 : 0);
    placedSection = insertIntoSection(ast.sections, effSection, insertIndex, createdNodes);

    liveChain.push(...createdNodes);
    anchorChain._insertAfterChain = liveChain;
    anchorChain._lastInsertAfter = createdNodes[createdNodes.length - 1];
  }

  // B-2: the break paragraph delimits sections — nodes after it move to a new
  // section (mirrors section.split; the break itself terminates the first).
  if (insertType === 'sectionBreak') {
    const breakNode = createdNodes[0];
    if (!toHeaderFooter) finishSectionBreak(breakNode.xml as Element);
    const breakSection = placedSection ?? bodyAnchorSection;
    const breakIdx = breakSection.nodes.findIndex((n: Node) => n.id === breakNode.id);
    const secIdx = ast.sections.indexOf(breakSection);
    const newNodes = breakSection.nodes.splice(breakIdx + 1);
    const secSeed = `${deriveInsertSeed(op)}:sectionBreak:${breakNode.id}:${newNodes.length}`;
    const liveSecIds = new Set<string>();
    try {
      for (const s of ast.sections ?? []) {
        if (typeof s.id === 'string' && s.id) liveSecIds.add(s.id);
      }
    } catch { /* best-effort */ }
    let secId: string;
    let secChecksum: string;
    if (isInsertReproducible()) {
      const base = `sec_${insertHash8(secSeed)}`;
      secId = base;
      let n = 2;
      while (liveSecIds.has(secId)) secId = `${base}_${n++}`;
      secChecksum = `sec_${insertHash8(`${secId}:${newNodes.length}`)}`;
    } else {
      secId = `sec_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
      secChecksum = `sec_${Date.now()}`;
    }
    const newSection: Section = {
      id: secId,
      heading: null,
      nodes: newNodes,
      version: 1,
      checksum: secChecksum,
      pageSetup: { ...(breakSection.pageSetup || {}) }
    };
    ast.sections.splice(secIdx + 1, 0, newSection);
    breakSection.version += 1;
  }

  // Header/footer content lives in the part, not in section nodes, so its
  // ids stay out of the body node map (id lookups resolve body nodes).
  if (parser && !toHeaderFooter) {
    for (const node of createdNodes) {
      parser.nodeMap?.set(node.id, node.xml);
    }
  }

  // R-42: honor insert-time table captions through the single setCaption
  // path (operations/table.ts, adjacent to updateCell) — no duplicate
  // caption logic here. Runs after DOM+AST placement so setCaption sees
  // the table in final order and writes the sibling paragraph once.
  if (insertType === 'table' && op.captionText !== undefined && String(op.captionText).trim() !== '') {
    const tableNodeId = createdNodes.length > 0 ? createdNodes[0].id : newId;
    await executeTable({
      action: 'table',
      target: tableNodeId,
      params: {
        op: 'setCaption',
        captionText: String(op.captionText),
        captionPosition: op.captionPosition ?? 'above',
      },
    }, ast);
  }
}
