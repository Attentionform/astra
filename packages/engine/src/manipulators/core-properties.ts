import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST, DocumentMetadata } from '../ast';
import { assertNoXxe } from '../ooxml/xxe-guard';

/**
 * E2 (F-2): read/update/synthesize the OPC core-properties part
 * (`docProps/core.xml`), the round-trip companion to the F-1 ingest in
 * `DocxParser.getAST()`.
 *
 * Design notes:
 * - Structural interface instead of importing `DocxParser` (which owns the
 *   real adapter): `DocxZipAdapter` is structurally identical, so the
 *   parser's `getZipAdapter()` satisfies it with zero import cycle.
 * - Updates are DOM-surgical: only touched elements are upserted, unknown
 *   siblings (lastPrinted, contentStatus, language, …) are preserved
 *   verbatim. Fresh synthesis emits Word-typical child order so the
 *   validator gate stays green.
 * - All text goes through DOM text nodes (no string-concat escaping bugs).
 * - Never throws on absent parts; XXE/DTD input throws (matches parser).
 */

export const CORE_PROPERTIES_PART = 'docProps/core.xml';
export const CORE_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml';
export const CORE_PROPERTIES_REL_TYPE =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

const CP_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC_NS = 'http://purl.org/dc/elements/1.1/';
const DCTERMS_NS = 'http://purl.org/dc/terms/';
const XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance';

/** Writable core fields (subset of DocumentMetadata that lives in core.xml). */
export interface CorePropertiesPatch {
  title?: string;
  creator?: string;
  subject?: string;
  description?: string;
  keywords?: string;
  category?: string;
  lastModifiedBy?: string;
  revision?: number;
  created?: Date | string;
  modified?: Date | string;
}

export interface CorePropsZipAdapter {
  hasFile(path: string): boolean;
  readText(path: string): string;
  readTextOrNull(path: string): string | null;
  writeText(path: string, content: string): void;
  writeBuffer(path: string, content: Buffer): void;
}

const CORE_FIELD_TAGS: Record<keyof Required<Omit<CorePropertiesPatch, 'created' | 'modified'>>, string> = {
  title: 'dc:title',
  subject: 'dc:subject',
  creator: 'dc:creator',
  keywords: 'cp:keywords',
  description: 'dc:description',
  lastModifiedBy: 'cp:lastModifiedBy',
  revision: 'cp:revision',
  category: 'cp:category',
};

/** Word-typical child order (ECMA-376 core-properties sequence).
 * lastPrinted sits between revision/created; category/contentStatus are the
 * tail; truly-unknown elements append after contentStatus (orderOf fallback).
 */
const CORE_CHILD_ORDER: string[] = [
  'dc:title',
  'dc:subject',
  'dc:creator',
  'cp:keywords',
  'dc:description',
  'cp:lastModifiedBy',
  'cp:revision',
  'cp:lastPrinted',
  'dcterms:created',
  'dcterms:modified',
  'cp:category',
  'cp:contentStatus',
];

function orderOf(tag: string): number {
  const idx = CORE_CHILD_ORDER.indexOf(tag);
  return idx === -1 ? CORE_CHILD_ORDER.length : idx;
}

/**
 * Canonicalizes an element to its `prefix:local` core tag so prefix-less
 * parts (default-NS `<title>`, or a renamed prefix) still rank correctly in
 * insertOrdered/findDirectChild. Returns null for truly-unknown elements
 * (they append at the tail via orderOf fallback).
 */
function canonicalCoreTag(el: Element): string | null {
  const rawTag: string = (el as Element).tagName || '';
  if (CORE_CHILD_ORDER.includes(rawTag)) return rawTag;
  const local: string =
    (el as unknown as { localName?: string }).localName ||
    (rawTag.includes(':') ? rawTag.split(':').pop() || rawTag : rawTag);
  const ns: string = (el as unknown as { namespaceURI?: string }).namespaceURI || '';
  if (ns === DC_NS) {
    if (local === 'title') return 'dc:title';
    if (local === 'subject') return 'dc:subject';
    if (local === 'creator') return 'dc:creator';
    if (local === 'description') return 'dc:description';
  }
  if (ns === CP_NS) {
    if (local === 'keywords') return 'cp:keywords';
    if (local === 'lastModifiedBy') return 'cp:lastModifiedBy';
    if (local === 'revision') return 'cp:revision';
    if (local === 'lastPrinted') return 'cp:lastPrinted';
    if (local === 'category') return 'cp:category';
    if (local === 'contentStatus') return 'cp:contentStatus';
  }
  if (ns === DCTERMS_NS) {
    if (local === 'created') return 'dcterms:created';
    if (local === 'modified') return 'dcterms:modified';
  }
  // Prefix-less without a usable NS (e.g. bare `<title>`): map by local name
  // alone — inside coreProperties these locals are unambiguous.
  if (!rawTag.includes(':')) {
    if (local === 'title') return 'dc:title';
    if (local === 'subject') return 'dc:subject';
    if (local === 'creator') return 'dc:creator';
    if (local === 'description') return 'dc:description';
    if (local === 'keywords') return 'cp:keywords';
    if (local === 'lastModifiedBy') return 'cp:lastModifiedBy';
    if (local === 'revision') return 'cp:revision';
    if (local === 'lastPrinted') return 'cp:lastPrinted';
    if (local === 'created') return 'dcterms:created';
    if (local === 'modified') return 'dcterms:modified';
    if (local === 'category') return 'cp:category';
    if (local === 'contentStatus') return 'cp:contentStatus';
  }
  return null;
}

function orderOfElement(el: Element): number {
  return orderOf(canonicalCoreTag(el) ?? (el as Element).tagName);
}

/**
 * Strict core-XML parse (fail-closed): any xmldom warning/error, parsererror
 * node, missing root, or non-coreProperties root throws E_VALIDATION.
 * Lenient parseLenient stays for Content_Types/rels (different owners).
 */
function parseCoreStrict(xml: string): Document {
  const problems: string[] = [];
  const errorHandler = {
    warning: (msg: unknown) => {
      problems.push(`warning: ${String(msg).split('\n')[0]}`);
    },
    error: (msg: unknown) => {
      problems.push(`error: ${String(msg).split('\n')[0]}`);
    },
    fatalError: (msg: unknown) => {
      problems.push(`fatal: ${String(msg).split('\n')[0]}`);
    },
  };
  let doc: Document | null = null;
  try {
    doc = new DOMParser({ locator: {}, errorHandler } as unknown as Record<string, unknown>).parseFromString(
      xml,
      'text/xml'
    ) as unknown as Document;
  } catch (err: unknown) {
    throw new Error(
      `E_VALIDATION: malformed docProps/core.xml (${err instanceof Error ? err.message : String(err)}). ` +
        `Remedy: repair the core-properties part (well-formed XML with a coreProperties root) and retry.`
    );
  }
  if (problems.length > 0 || !doc || !doc.documentElement) {
    throw new Error(
      `E_VALIDATION: malformed docProps/core.xml (${problems[0] ?? 'unparseable XML'}). ` +
        `Remedy: repair the core-properties part (well-formed XML with a coreProperties root) and retry.`
    );
  }
  try {
    if (doc.getElementsByTagName('parsererror').length > 0) {
      throw new Error(
        `E_VALIDATION: malformed docProps/core.xml (parsererror). ` +
          `Remedy: repair the core-properties part (well-formed XML with a coreProperties root) and retry.`
      );
    }
  } catch (err: unknown) {
    if (err instanceof Error && err.message.startsWith('E_VALIDATION')) throw err;
    /* ignore probe failure — root check below still applies */
  }
  const root = doc.documentElement as unknown as Element;
  const rootLocal: string =
    (root as unknown as { localName?: string }).localName ||
    (root.tagName.includes(':') ? root.tagName.split(':').pop() || root.tagName : root.tagName);
  if (rootLocal !== 'coreProperties') {
    throw new Error(
      `E_VALIDATION: malformed docProps/core.xml (root <${root.tagName}> is not coreProperties). ` +
        `Remedy: repair the core-properties part (well-formed XML with a coreProperties root) and retry.`
    );
  }
  return doc;
}

function parseLenient(xml: string): Document | null {
  try {
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    if (!doc || !doc.documentElement) return null;
    if (doc.getElementsByTagName('parsererror').length > 0) return null;
    return doc;
  } catch {
    return null;
  }
}

function directChildren(root: Element): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < root.childNodes.length; i++) {
    const c = root.childNodes[i] as Element;
    if ((c as unknown as { nodeType?: number }).nodeType === 1) out.push(c);
  }
  return out;
}

function findDirectChild(root: Element, tag: string): Element | null {
  for (const c of directChildren(root)) {
    if (c.tagName === tag) return c;
    if ((canonicalCoreTag(c) ?? c.tagName) === tag) return c;
  }
  return null;
}

function textOf(doc: Document, tag: string): string | undefined {
  try {
    const els = doc.getElementsByTagName(tag);
    if (els.length > 0) {
      const t = (els[0] as Element).textContent;
      if (t != null) {
        const trimmed = t.trim();
        if (trimmed !== '') return trimmed;
      }
    }
  } catch {
    /* fall through to NS fallback */
  }
  // Prefix-less / renamed-prefix fallback: same localName in the canonical
  // NS-URI (getElementsByTagNameNS), then bare-local lookup for parts that
  // carry no usable namespace.
  try {
    const nsFallback: Record<string, [string, string]> = {
      'dc:title': [DC_NS, 'title'],
      'dc:subject': [DC_NS, 'subject'],
      'dc:creator': [DC_NS, 'creator'],
      'dc:description': [DC_NS, 'description'],
      'cp:keywords': [CP_NS, 'keywords'],
      'cp:lastModifiedBy': [CP_NS, 'lastModifiedBy'],
      'cp:revision': [CP_NS, 'revision'],
      'cp:lastPrinted': [CP_NS, 'lastPrinted'],
      'cp:category': [CP_NS, 'category'],
      'cp:contentStatus': [CP_NS, 'contentStatus'],
      'dcterms:created': [DCTERMS_NS, 'created'],
      'dcterms:modified': [DCTERMS_NS, 'modified'],
    };
    const fb = nsFallback[tag];
    if (fb) {
      const [nsUri, local] = fb;
      try {
        const nsEls = (doc as unknown as {
          getElementsByTagNameNS(ns: string, local: string): ArrayLike<Element>;
        }).getElementsByTagNameNS(nsUri, local);
        if (nsEls && nsEls.length > 0) {
          const t = (nsEls[0] as Element).textContent;
          if (t != null) {
            const trimmed = t.trim();
            if (trimmed !== '') return trimmed;
          }
        }
      } catch {
        /* fall through to bare-local lookup */
      }
      const bare = doc.getElementsByTagName(local);
      for (let i = 0; i < bare.length; i++) {
        const el = bare[i] as unknown as Element;
        // Bare `<title>` (no colon) only; a prefixed miss already returned.
        if ((el as Element).tagName.includes(':')) continue;
        const t = (el as Element).textContent;
        if (t != null) {
          const trimmed = t.trim();
          if (trimmed !== '') return trimmed;
        }
      }
    }
  } catch {
    return undefined;
  }
  return undefined;
}

function toDate(value: string | undefined): Date | undefined {
  if (value === undefined) return undefined;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

function normalizeDate(value: Date | string | undefined): Date | undefined {
  if (value === undefined) return undefined;
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? undefined : value;
  return toDate(String(value));
}

/**
 * Parses a docProps/core.xml payload into DocumentMetadata fields.
 * Missing fields stay undefined. Fail-closed: XXE/DTD input throws
 * (security, matches the parser) and malformed XML throws E_VALIDATION
 * (never a silent {} — a corrupt core part must not masquerade as absent).
 */
export function parseCorePropertiesXml(xml: string): Partial<DocumentMetadata> {
  if (!xml || typeof xml !== 'string' || !xml.includes('<')) return {};
  assertNoXxe(xml, 'docProps/core.xml');
  const doc = parseCoreStrict(xml);
  const out: Partial<DocumentMetadata> = {};
  const title = textOf(doc, 'dc:title');
  if (title !== undefined) out.title = title;
  const subject = textOf(doc, 'dc:subject');
  if (subject !== undefined) out.subject = subject;
  const creator = textOf(doc, 'dc:creator');
  if (creator !== undefined) out.creator = creator;
  const keywords = textOf(doc, 'cp:keywords');
  if (keywords !== undefined) out.keywords = keywords;
  const description = textOf(doc, 'dc:description');
  if (description !== undefined) out.description = description;
  const lastModifiedBy = textOf(doc, 'cp:lastModifiedBy');
  if (lastModifiedBy !== undefined) out.lastModifiedBy = lastModifiedBy;
  const revisionRaw = textOf(doc, 'cp:revision');
  if (revisionRaw !== undefined) {
    if (!/^-?\d+$/.test(revisionRaw.trim())) {
      throw new Error(
        `E_VALIDATION: metadata revision must be an integer (got '${revisionRaw}'). ` +
          `Remedy: repair the core-properties part with an integer cp:revision or omit it.`
      );
    }
    const n = parseInt(revisionRaw.trim(), 10);
    if (Number.isFinite(n)) out.revision = n;
  }
  const category = textOf(doc, 'cp:category');
  if (category !== undefined) out.category = category;
  const created = toDate(textOf(doc, 'dcterms:created'));
  if (created !== undefined) out.created = created;
  const modified = toDate(textOf(doc, 'dcterms:modified'));
  if (modified !== undefined) out.modified = modified;
  return out;
}

function ensureRootNamespaces(root: Element): void {
  const needs: Array<[string, string]> = [
    ['xmlns:cp', CP_NS],
    ['xmlns:dc', DC_NS],
    ['xmlns:dcterms', DCTERMS_NS],
    ['xmlns:xsi', XSI_NS],
  ];
  for (const [attr, uri] of needs) {
    try {
      if (!root.getAttribute(attr)) root.setAttribute(attr, uri);
    } catch {
      /* best-effort */
    }
  }
}

function setElementText(doc: Document, el: Element, value: string): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(doc.createTextNode(value));
}

/** Inserts `el` before the first direct child ordered after `tag`.
 * Unknown siblings rank by the ECMA sequence (lastPrinted between
 * revision/created, category/contentStatus tail); truly-unknown tags append.
 */
function insertOrdered(root: Element, el: Element, tag: string): void {
  const target = orderOf(tag);
  for (const child of directChildren(root)) {
    if (orderOfElement(child) > target) {
      root.insertBefore(el, child);
      return;
    }
  }
  root.appendChild(el);
}

function upsertTextField(doc: Document, root: Element, tag: string, value: string): void {
  const existing = findDirectChild(root, tag);
  if (existing) {
    setElementText(doc, existing, value);
    return;
  }
  const created = doc.createElement(tag);
  created.appendChild(doc.createTextNode(value));
  insertOrdered(root, created, tag);
}

function upsertDateField(doc: Document, root: Element, tag: string, value: Date): void {
  const existing = findDirectChild(root, tag);
  const iso = value.toISOString();
  if (existing) {
    setElementText(doc, existing, iso);
    try {
      existing.setAttribute('xsi:type', 'dcterms:W3CDTF');
    } catch {
      /* best-effort */
    }
    return;
  }
  const created = doc.createElement(tag);
  try {
    created.setAttribute('xsi:type', 'dcterms:W3CDTF');
  } catch {
    /* best-effort */
  }
  created.appendChild(doc.createTextNode(iso));
  insertOrdered(root, created, tag);
}

function applyPatchToDom(doc: Document, root: Element, patch: CorePropertiesPatch): string[] {
  const touched: string[] = [];
  for (const [field, tag] of Object.entries(CORE_FIELD_TAGS) as Array<[keyof typeof CORE_FIELD_TAGS, string]>) {
    const value = patch[field];
    if (value === undefined) continue;
    if (field === 'revision') {
      const n = patch.revision;
      if (typeof n !== 'number' || !Number.isInteger(n)) {
        throw new Error(
          `E_VALIDATION: metadata revision must be an integer (got '${String(n)}'). ` +
            `Remedy: pass an integer revision or omit the field to preserve the current value.`
        );
      }
      upsertTextField(doc, root, tag, String(n));
    } else {
      if (typeof value !== 'string') {
        throw new Error(
          `E_VALIDATION: metadata field '${field}' must be a string (got ${typeof value}). ` +
            `Remedy: pass a string or omit the field to preserve the current value.`
        );
      }
      upsertTextField(doc, root, tag, value);
    }
    touched.push(field);
  }
  const created = normalizeDate(patch.created);
  if (created !== undefined) {
    upsertDateField(doc, root, 'dcterms:created', created);
    touched.push('created');
  }
  const modified = normalizeDate(patch.modified);
  if (modified !== undefined) {
    upsertDateField(doc, root, 'dcterms:modified', modified);
    touched.push('modified');
  }
  return touched;
}

function freshCorePropertiesDoc(): { doc: Document; root: Element } {
  const skeleton =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cp:coreProperties xmlns:cp="${CP_NS}" xmlns:dc="${DC_NS}" ` +
    `xmlns:dcterms="${DCTERMS_NS}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" ` +
    `xmlns:xsi="${XSI_NS}"></cp:coreProperties>`;
  const doc = new DOMParser().parseFromString(skeleton, 'text/xml');
  return { doc, root: doc.documentElement as unknown as Element };
}

/**
 * Serializes `meta` to a fresh docProps/core.xml payload (Word-typical
 * order). Only defined fields are emitted.
 */
export function serializeCorePropertiesXml(meta: Partial<DocumentMetadata>): string {
  const { doc, root } = freshCorePropertiesDoc();
  applyPatchToDom(doc, root, meta as CorePropertiesPatch);
  return new XMLSerializer().serializeToString(doc);
}

/**
 * Merges `patch` into `existingXml` (surgical: untouched + unknown elements
 * preserved), or synthesizes a fresh part from `fallback` when there is no
 * existing part. Returns the resulting XML plus the touched field names.
 * Fail-closed: XXE/DTD input throws loudly (security); malformed existing
 * XML throws E_VALIDATION (never a silent rebuild — corruption must surface).
 */
export function mergeCorePropertiesXml(
  existingXml: string | null,
  patch: CorePropertiesPatch,
  fallback?: Partial<DocumentMetadata>
): { xml: string; touched: string[] } {
  const serializer = new XMLSerializer();
  if (existingXml) {
    assertNoXxe(existingXml, 'docProps/core.xml');
    // Strict: malformed existing part throws E_VALIDATION (loud, fail-closed).
    const doc = parseCoreStrict(existingXml);
    const root = doc.documentElement as unknown as Element;
    ensureRootNamespaces(root);
    const touched = applyPatchToDom(doc, root, patch);
    return { xml: serializer.serializeToString(doc), touched };
  }
  const { doc, root } = freshCorePropertiesDoc();
  const merged: CorePropertiesPatch = { ...(fallback as CorePropertiesPatch | undefined), ...patch };
  const touched = applyPatchToDom(doc, root, merged);
  return { xml: serializer.serializeToString(doc), touched };
}

function ensureContentTypeOverride(adapter: CorePropsZipAdapter): void {
  const path = '[Content_Types].xml';
  let raw: string | null = null;
  try {
    raw = adapter.readTextOrNull(path);
  } catch {
    return;
  }
  if (!raw) return;
  const doc = parseLenient(raw);
  if (!doc || !doc.documentElement) return;
  const root = doc.documentElement as unknown as Element;
  const overrides = root.getElementsByTagName('Override');
  for (let i = 0; i < overrides.length; i++) {
    const el = overrides[i] as unknown as Element;
    if (el.getAttribute('PartName') === '/docProps/core.xml') return;
  }
  const created = doc.createElement('Override');
  created.setAttribute('PartName', '/docProps/core.xml');
  created.setAttribute('ContentType', CORE_PROPERTIES_CONTENT_TYPE);
  root.appendChild(created);
  adapter.writeText(path, new XMLSerializer().serializeToString(doc));
}

function ensurePackageRel(adapter: CorePropsZipAdapter): void {
  const path = '_rels/.rels';
  let raw: string | null = null;
  try {
    raw = adapter.readTextOrNull(path);
  } catch {
    return;
  }
  if (!raw) return;
  const doc = parseLenient(raw);
  if (!doc || !doc.documentElement) return;
  const root = doc.documentElement as unknown as Element;
  const rels = root.getElementsByTagName('Relationship');
  let maxId = 0;
  for (let i = 0; i < rels.length; i++) {
    const el = rels[i] as unknown as Element;
    if ((el.getAttribute('Type') || '') === CORE_PROPERTIES_REL_TYPE) return;
    const m = /^rId(\d+)$/.exec(el.getAttribute('Id') || '');
    if (m) {
      const n = parseInt(m[1], 10);
      if (Number.isFinite(n) && n > maxId) maxId = n;
    }
  }
  let candidate = `rId${maxId + 1}`;
  const taken = new Set<string>();
  for (let i = 0; i < rels.length; i++) {
    taken.add((rels[i] as unknown as Element).getAttribute('Id') || '');
  }
  while (taken.has(candidate)) {
    maxId += 1;
    candidate = `rId${maxId + 1}`;
  }
  const created = doc.createElement('Relationship');
  created.setAttribute('Id', candidate);
  created.setAttribute('Type', CORE_PROPERTIES_REL_TYPE);
  created.setAttribute('Target', CORE_PROPERTIES_PART);
  root.appendChild(created);
  adapter.writeText(path, new XMLSerializer().serializeToString(doc));
}

/**
 * Writes `xml` to docProps/core.xml and ensures the package stays
 * well-formed (Content_Types Override + package rel entry). The part write
 * throws on adapter failure; the two ensure-steps are best-effort so a
 * hand-built minimal package never blocks a save.
 */
export function writeCorePropertiesPart(adapter: CorePropsZipAdapter, xml: string): void {
  adapter.writeText(CORE_PROPERTIES_PART, xml);
  try {
    ensureContentTypeOverride(adapter);
  } catch {
    /* best-effort — the part itself is written */
  }
  try {
    ensurePackageRel(adapter);
  } catch {
    /* best-effort — the part itself is written */
  }
}

function patchToMetadata(patch: CorePropertiesPatch): Partial<DocumentMetadata> {
  const out: Partial<DocumentMetadata> = {};
  if (patch.title !== undefined) out.title = patch.title;
  if (patch.creator !== undefined) out.creator = patch.creator;
  if (patch.subject !== undefined) out.subject = patch.subject;
  if (patch.description !== undefined) out.description = patch.description;
  if (patch.keywords !== undefined) out.keywords = patch.keywords;
  if (patch.category !== undefined) out.category = patch.category;
  if (patch.lastModifiedBy !== undefined) out.lastModifiedBy = patch.lastModifiedBy;
  if (patch.revision !== undefined) out.revision = patch.revision;
  const created = normalizeDate(patch.created);
  if (created !== undefined) out.created = created;
  const modified = normalizeDate(patch.modified);
  if (modified !== undefined) out.modified = modified;
  return out;
}

export class CorePropertiesManipulator {
  private ast: DocumentAST;
  private adapter?: CorePropsZipAdapter;

  constructor(ast: DocumentAST, adapter?: CorePropsZipAdapter) {
    this.ast = ast;
    this.adapter = adapter;
  }

  /** Reads the live part (absent file/fields → {}, never throws on absent). */
  public read(): Partial<DocumentMetadata> {
    if (!this.adapter) return {};
    let raw: string | null = null;
    try {
      raw = this.adapter.readTextOrNull(CORE_PROPERTIES_PART);
    } catch {
      return {};
    }
    if (!raw) return {};
    return parseCorePropertiesXml(raw);
  }

  /**
   * Applies a partial `patch`: only provided fields are touched (AST +
   * part); unprovided fields are preserved. Without an adapter only the
   * in-memory AST is updated. Returns the touched field names.
   */
  public update(patch: CorePropertiesPatch): { updated: string[]; created: boolean } {
    // Fail-closed pre-check: reject non-integer revision BEFORE touching the
    // AST so a throwing update leaves in-memory metadata unpolluted.
    if (patch.revision !== undefined && (typeof patch.revision !== 'number' || !Number.isInteger(patch.revision))) {
      throw new Error(
        `E_VALIDATION: metadata revision must be an integer (got '${String(patch.revision)}'). ` +
          `Remedy: pass an integer revision or omit the field to preserve the current value.`
      );
    }
    const metaPatch = patchToMetadata(patch);
    const updated = Object.keys(metaPatch);
    Object.assign(this.ast.metadata, metaPatch);
    if (!this.adapter) return { updated, created: false };
    const existed = this.adapter.hasFile(CORE_PROPERTIES_PART);
    let existing: string | null = null;
    if (existed) {
      try {
        existing = this.adapter.readTextOrNull(CORE_PROPERTIES_PART);
      } catch {
        existing = null;
      }
    }
    const { xml } = mergeCorePropertiesXml(existing, patch, this.ast.metadata);
    writeCorePropertiesPart(this.adapter, xml);
    return { updated, created: !existed };
  }
}
