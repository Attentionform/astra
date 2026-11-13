/**
 * Allowlist OOXML normalizer.
 *
 * Position: immediately before the validator-gated save. Runs after DOM parse,
 * so syntax errors (unclosed tags) fail closed first and never reach the
 * passes below. Three passes, all structure-only — text, numbers, and styling
 * values are never altered:
 *
 * Pass 1 — tag renames. Strict lookup table; the ONLY verified alias is
 *   w:tcBdr -> w:tcBorders. tblBdr / pBdr stay OUT (same guess class, but no
 *   specimen + test each, so never added blind). Attributes and children are
 *   transferred to the canonical node; unknown tags are untouched and left
 *   for the strict gate to fail closed.
 * Pass 2 — default infill. Whitelisted element+attribute pairs only, each
 *   with a test: w:shd without w:val -> w:val="clear".
 * Pass 3 — sequence sort. Containers w:pPr / w:rPr / w:tblPr / w:tcPr are
 *   stable-sorted by index in the canonical order arrays from schema-order.ts
 *   (TCPR_CHILD_ORDER et al). Only the elements the order array lists are
 *   moved, and only among the places those elements already occupy. An element
 *   the array does not list (w:ins or w:del on a paragraph mark, mc:AlternateContent,
 *   a w14 effect) keeps its place and is never dropped. This is the same rule
 *   as sortChildrenBySchema, so the two sorts agree on every container.
 *
 * Invariants: idempotent (second run is a fixpoint with an empty log);
 * every rewrite is logged per item with its xPath, enclosing node id when
 * known, and before→after detail (so migration audits can name each
 * healing); clean packages yield an empty list and byte-identical output
 * (parts with no transforms keep their exact input string — no
 * parse/serialize round-trip).
 */

import { XMLSerializer } from '@xmldom/xmldom';
import { DOMParser } from './dom-parser';
import {
  PPR_CHILD_ORDER,
  RPR_CHILD_ORDER,
  TBLPR_CHILD_ORDER,
  TCPR_CHILD_ORDER,
} from './schema-order';
import { createWordElement } from './ns';
import { assertNoXxe } from './xxe-guard';
import { elementXPath } from '../parser/para-ids';

/** Verified alias table. Grows ONLY by incident (real failing file + test). */
const TAG_ALIASES: Readonly<Record<string, string>> = {
  'w:tcBdr': 'w:tcBorders',
};

/** Whitelisted default infills: element tag + attribute + spec default value. */
const DEFAULT_INFILLS: ReadonlyArray<{ tag: string; attr: string; value: string }> = [
  { tag: 'w:shd', attr: 'w:val', value: 'clear' },
];

/** Containers sorted by schema sequence, in fixed log order. */
const SORT_CONTAINERS: ReadonlyArray<{ tag: string; order: readonly string[] }> = [
  { tag: 'w:pPr', order: PPR_CHILD_ORDER },
  { tag: 'w:rPr', order: RPR_CHILD_ORDER },
  { tag: 'w:tblPr', order: TBLPR_CHILD_ORDER },
  { tag: 'w:tcPr', order: TCPR_CHILD_ORDER },
  // Border containers: w:tcBorders follows CT_TcBorders order (top, left,
  // bottom, right, insideH, insideV, plus diagonal tail) and w:pBdr follows
  // CT_PBdr order (top, left, bottom, right, between, bar).
  { tag: 'w:tcBorders', order: ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV', 'w:tl2br', 'w:tr2bl'] },
  { tag: 'w:pBdr', order: ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:between', 'w:bar'] },
];

/**
 * Location of one rewrite: the changed node's xPath plus the enclosing
 * addressable node id (session `data-astra-id` when the live DOM carries
 * one, else the native `w14:paraId`) when known. Never throws.
 */
function describeLocation(el: Element): { nodeId: string | null; xPath: string } {
  let xPath: string;
  try {
    xPath = elementXPath(el);
  } catch {
    xPath = '//w:p';
  }
  let nodeId: string | null = null;
  try {
    let cur: Element | null = el;
    while (cur && cur.nodeType === 1) {
      const tag = (cur as Element).tagName;
      if (tag === 'w:p' || tag === 'w:tbl') {
        const sid = (cur as Element).getAttribute?.('data-astra-id');
        if (sid) {
          nodeId = sid;
          break;
        }
        const pid = (cur as Element).getAttribute?.('w14:paraId');
        if (pid) {
          nodeId = pid;
          break;
        }
      }
      const parent = cur.parentNode as Element | null;
      if (!parent || parent.nodeType !== 1) break;
      cur = parent;
    }
  } catch {
    nodeId = null;
  }
  return { nodeId, xPath };
}

/** ` at <xPath>` plus ` (node <id>)` when the enclosing node is addressable. */
function locationSuffix(el: Element): string {
  const { nodeId, xPath } = describeLocation(el);
  return nodeId ? ` at ${xPath} (node ${nodeId})` : ` at ${xPath}`;
}

function snapshotElements(doc: Document, tag: string): Element[] {
  const list = doc.getElementsByTagName(tag);
  const out: Element[] = [];
  for (let i = 0; i < list.length; i++) out.push(list[i] as unknown as Element);
  return out;
}

/** Pass 1: strict alias renames (attributes + children transferred). */
function passRenames(doc: Document, entries: string[]): void {
  for (const [alias, canonical] of Object.entries(TAG_ALIASES)) {
    for (const old of snapshotElements(doc, alias)) {
      const fresh = createWordElement(doc, canonical);
      if (old.attributes) {
        for (let i = 0; i < old.attributes.length; i++) {
          const attr = old.attributes[i];
          fresh.setAttribute(attr.name, attr.value);
        }
      }
      while (old.firstChild) fresh.appendChild(old.firstChild);
      const parent = old.parentNode;
      if (!parent) continue;
      parent.insertBefore(fresh, old);
      parent.removeChild(old);
      entries.push(`Renamed ${alias} \u2192 ${canonical}${locationSuffix(fresh)}`);
    }
  }
}

/** Pass 2: whitelisted default infills (guard skips already-correct nodes). */
function passInfills(doc: Document, entries: string[]): void {
  for (const { tag, attr, value } of DEFAULT_INFILLS) {
    for (const el of snapshotElements(doc, tag)) {
      if (!el.hasAttribute(attr)) {
        el.setAttribute(attr, value);
        entries.push(`Infilled ${attr}="${value}" on ${tag}${locationSuffix(el)}`);
      }
    }
  }
}

/**
 * Pass 3: stable sequence sort of one container's element children.
 * Elements the order array lists are sorted among the places they occupy;
 * unlisted elements and non-element nodes (whitespace) stay where they are.
 * Returns the before/after child-tag sequences plus whether the container moved.
 */
function sortContainerChildren(
  container: Element,
  order: readonly string[]
): { reordered: boolean; before: string[]; after: string[] } {
  const nodes: Node[] = [];
  for (let i = 0; i < container.childNodes.length; i++) nodes.push(container.childNodes[i]);
  const before: string[] = [];
  const slots: number[] = [];
  const listed: Element[] = [];
  for (let i = 0; i < nodes.length; i++) {
    if (nodes[i].nodeType !== 1) continue;
    const el = nodes[i] as Element;
    before.push(el.tagName);
    if (order.indexOf(el.tagName) !== -1) {
      slots.push(i);
      listed.push(el);
    }
  }
  if (listed.length <= 1) return { reordered: false, before, after: before };
  const sorted = listed
    .map((el, pos) => ({ el, pos }))
    .sort((a, b) => order.indexOf(a.el.tagName) - order.indexOf(b.el.tagName) || a.pos - b.pos);
  let changed = false;
  for (let i = 0; i < listed.length; i++) {
    if (sorted[i].el !== listed[i]) {
      changed = true;
      break;
    }
  }
  if (!changed) return { reordered: false, before, after: before };
  const next = nodes.slice();
  for (let i = 0; i < slots.length; i++) next[slots[i]] = sorted[i].el;
  while (container.firstChild) container.removeChild(container.firstChild);
  for (const node of next) container.appendChild(node);
  const after = next.filter((node) => node.nodeType === 1).map((node) => (node as Element).tagName);
  return { reordered: true, before, after };
}

/** Pass 3 driver over the four whitelisted containers. */
function passSorts(doc: Document, entries: string[], reorderedContainers?: string[]): void {
  for (const { tag, order } of SORT_CONTAINERS) {
    for (const container of snapshotElements(doc, tag)) {
      const { reordered, before, after } = sortContainerChildren(container, order);
      if (reordered) {
        reorderedContainers?.push(elementXPath(container));
        entries.push(`Reordered ${tag} into schema order${locationSuffix(container)}: [${before.join(',')}] \u2192 [${after.join(',')}]`);
      }
    }
  }
}

/**
 * Applies the three passes to a live DOM in place. Returns per-run log
 * entries (empty when the DOM was already canonical). Never touches text,
 * numbers, or styling values — tags, one whitelisted attribute, and child
 * order only. Each entry names its xPath, enclosing node id when known,
 * and before→after detail (pass order, document order within a pass).
 */
export function normalizeDocumentDom(doc: Document): string[] {
  const entries: string[] = [];
  passRenames(doc, entries);
  passInfills(doc, entries);
  passSorts(doc, entries);
  return entries;
}

/** Dry-run over a live DOM: counts would-be transforms without mutating. */
export function previewDocumentDom(doc: Document): string[] {
  const clone = doc.cloneNode(true) as Document;
  return normalizeDocumentDom(clone);
}

function validationFailure(partName: string, detail: string): Error {
  return Object.assign(
    new Error(
      `E_VALIDATION: normalizer refused ${partName}: ${detail}. ` +
        `Remedy: repair the XML well-formedness first, then retry.`
    ),
    { code: 'E_VALIDATION', remedy: 'repair the XML well-formedness first, then retry', retryable: false as const }
  );
}

/** Strict part parse: XXE sweep, then syntax errors fail closed first. */
export function parsePartStrict(partName: string, xml: string): Document | null {
  if (!xml || !xml.includes('<')) return null;
  assertNoXxe(xml, partName);
  let failed = false;
  let detail = '';
  const parser = new DOMParser({
    errorHandler: {
      warning: () => {},
      error: (msg: string) => {
        failed = true;
        detail += msg + '; ';
      },
      fatalError: (msg: string) => {
        failed = true;
        detail += msg + '; ';
      },
    },
  });
  const doc = parser.parseFromString(xml, 'text/xml');
  const parserErrors = doc.getElementsByTagName('parsererror');
  if (failed || (parserErrors && parserErrors.length > 0)) {
    const hint = detail.trim().slice(0, 300) || String(parserErrors[0]?.textContent ?? 'invalid XML').slice(0, 300);
    throw validationFailure(partName, `invalid XML (${hint})`);
  }
  return doc;
}

const partSerializer = new XMLSerializer();

/**
 * Normalizes one XML part string. Returns the (possibly unchanged) XML plus
 * the part's log entries. Clean parts return their EXACT input string, so
 * clean packages stay byte-identical.
 */
export function normalizePartXml(partName: string, xml: string): { xml: string; normalized: string[] } {
  const doc = parsePartStrict(partName, xml);
  if (!doc) return { xml, normalized: [] };
  const entries: string[] = [];
  passRenames(doc, entries);
  passInfills(doc, entries);
  passSorts(doc, entries);
  if (entries.length === 0) return { xml, normalized: [] };
  return { xml: partSerializer.serializeToString(doc), normalized: entries };
}

/**
 * Reorder-only repair of one XML part: runs pass 3 (sequence sort) and nothing else. No alias renames and no
 * default infills, so element names, attributes and text are never touched. Returns the (possibly unchanged)
 * XML, the log entries, and the xPath of every container whose children were reordered. A part with nothing to
 * reorder returns its exact input string. PDF export uses this to repair a copy, then proves the repair with
 * `verifyOrderOnlyRepair`.
 */
export function reorderPartXml(
  partName: string,
  xml: string
): { xml: string; reordered: string[]; containers: string[] } {
  const doc = parsePartStrict(partName, xml);
  if (!doc) return { xml, reordered: [], containers: [] };
  const entries: string[] = [];
  const containers: string[] = [];
  passSorts(doc, entries, containers);
  if (entries.length === 0) return { xml, reordered: [], containers: [] };
  return { xml: partSerializer.serializeToString(doc), reordered: entries, containers };
}

const WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The containers an order-only repair may reorder: the tags the normalizer sorts, with each container's child
 * order. The verifier reads this table, never the list a repair reports, so a wrong report cannot widen what is
 * accepted. A container is matched by namespace URI (WordprocessingML main) and local name; its listed children
 * are matched by the qualified name as written, the same rule `sortContainerChildren` uses.
 */
export const ORDERED_CONTAINERS: ReadonlyMap<string, ReadonlySet<string>> = new Map(
  SORT_CONTAINERS.map((c) => [c.tag.slice(c.tag.indexOf(':') + 1), new Set(c.order)])
);

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

function isWhitespaceText(node: Node): boolean {
  return node.nodeType === 3 && /^[ \t\r\n]*$/.test(node.nodeValue ?? '');
}

function nodeKey(node: Node): string | null {
  if (node.nodeType === 3 || node.nodeType === 4) return `#text:${JSON.stringify(node.nodeValue ?? '')}`;
  if (node.nodeType === 7) return `#pi:${(node as ProcessingInstruction).target}:${JSON.stringify((node as ProcessingInstruction).data)}`;
  if (node.nodeType === 8) return `#comment:${JSON.stringify(node.nodeValue ?? '')}`;
  return null;
}

/** Element identity and attributes. Throws when a prefixed name has no namespace URI (equivalence cannot be shown). */
function elementHead(el: Element): string {
  const prefix = el.prefix ?? '';
  if (prefix && !el.namespaceURI) throw new Error(`no namespace for ${el.tagName}`);
  const attrs: string[] = [];
  for (let i = 0; el.attributes && i < el.attributes.length; i++) {
    const at = el.attributes[i];
    const isDecl = at.name === 'xmlns' || at.name.startsWith('xmlns:');
    if (!isDecl && at.name.includes(':') && !at.namespaceURI) throw new Error(`no namespace for @${at.name}`);
    attrs.push(`${isDecl ? XMLNS_NS : at.namespaceURI ?? ''}|${at.name}=${JSON.stringify(at.value)}`);
  }
  attrs.sort();
  return `<${el.namespaceURI ?? ''}|${prefix}|${el.localName ?? el.tagName} ${attrs.join(' ')}>`;
}

/**
 * Canonical string of an element, built bottom-up. In an ordered container the listed children are sorted among
 * themselves while every other child (unlisted element, whitespace, text, comment, processing instruction) is kept
 * in its slot, which is exactly what the repair moves. A container that holds non-whitespace text, a comment or a
 * processing instruction is not treated as reorderable: its children stay in sequence, so a reorder there is
 * refused as ambiguous.
 */
function canonical(el: Element): string {
  const head = elementHead(el);
  const listed = el.namespaceURI === WML_NS ? ORDERED_CONTAINERS.get(el.localName ?? '') : undefined;
  const parts: Array<string | null> = [];
  const movable: string[] = [];
  let ambiguous = false;
  for (let node = el.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1) {
      const child = node as Element;
      const key = canonical(child);
      if (listed && listed.has(child.tagName)) {
        parts.push(null);
        movable.push(key);
      } else parts.push(key);
    } else {
      const key = nodeKey(node);
      if (key === null) throw new Error(`unsupported node type ${node.nodeType}`);
      if (!isWhitespaceText(node)) ambiguous = true;
      parts.push(key);
    }
  }
  if (!listed || ambiguous) {
    // Not reorderable: every child, listed or not, stays in sequence.
    const inSeq: string[] = [];
    let m = 0;
    for (const p of parts) inSeq.push(p === null ? movable[m++] : p);
    return `${head}${inSeq.join('\u0001')}</>`;
  }
  movable.sort();
  let m = 0;
  const slots = parts.map((p) => (p === null ? `\u0002${movable[m++]}` : p));
  return `${head}${slots.join('\u0001')}</>`;
}

/** Everything outside the root element, with whitespace-only runs removed (the serializer drops a trailing newline). */
function outsideRoot(xml: string): string | null {
  let i = 0;
  const out: string[] = [];
  const skipTo = (end: string): boolean => {
    const j = xml.indexOf(end, i);
    if (j < 0) return false;
    out.push(xml.slice(i, j + end.length));
    i = j + end.length;
    return true;
  };
  const text = xml;
  if (text.charCodeAt(0) === 0xfeff) {
    out.push('\uFEFF');
    i = 1;
  }
  for (;;) {
    while (i < text.length && /[ \t\r\n]/.test(text[i])) i++;
    if (text.startsWith('<?', i)) { if (!skipTo('?>')) return null; }
    else if (text.startsWith('<!--', i)) { if (!skipTo('-->')) return null; }
    else break;
  }
  const rootStart = i;
  let end = text.length;
  const tail: string[] = [];
  for (;;) {
    let e = end;
    while (e > rootStart && /[ \t\r\n]/.test(text[e - 1])) e--;
    if (text.slice(0, e).endsWith('-->')) {
      const k = text.lastIndexOf('<!--', e);
      if (k <= rootStart) break;
      tail.unshift(text.slice(k, e)); end = k;
    } else if (text.slice(0, e).endsWith('?>')) {
      const k = text.lastIndexOf('<?', e);
      if (k <= rootStart) break;
      tail.unshift(text.slice(k, e)); end = k;
    } else break;
  }
  return `${out.join('')}|${tail.join('')}`;
}

/**
 * Independent proof that `repairedXml` differs from `originalXml` only in the order of listed children inside
 * the ordered containers. It does not use the sort code and does not read the list the repair reported. Both
 * parts are parsed strictly and reduced to canonical strings (see `canonical`): names compare by namespace URI,
 * prefix and local name, attributes by namespace URI, name and value, with namespace declarations compared as
 * attributes; text, comments and processing instructions take part; the XML declaration, comments and processing
 * instructions outside the root must be equal. Any difference, any parse failure and any case where equivalence
 * cannot be shown returns false.
 */
export function verifyOrderOnlyRepair(partName: string, originalXml: string, repairedXml: string): boolean {
  try {
    const original = parsePartStrict(partName, originalXml);
    const repaired = parsePartStrict(partName, repairedXml);
    if (!original || !repaired || !original.documentElement || !repaired.documentElement) return false;
    const outA = outsideRoot(originalXml);
    const outB = outsideRoot(repairedXml);
    if (outA === null || outB === null || outA !== outB) return false;
    return canonical(original.documentElement) === canonical(repaired.documentElement);
  } catch {
    return false;
  }
}

export interface NormalizePackageResult {
  parts: Record<string, string>;
  normalized: string[];
}

/**
 * Normalizes a package (part name -> XML string). Iterates parts in sorted
 * name order; counts aggregate across parts. Syntax errors in ANY part fail
 * closed (E_VALIDATION naming the part) before any output is produced.
 * Idempotent: normalizing the result yields itself with an empty list.
 */
export function normalizePackage(parts: Record<string, string>): NormalizePackageResult {
  const names = Object.keys(parts).sort();
  const parsed = new Map<string, Document | null>();
  for (const name of names) parsed.set(name, parsePartStrict(name, parts[name]));
  const entries: string[] = [];
  const rewritten = new Map<string, string>();
  for (const name of names) {
    const doc = parsed.get(name);
    if (!doc) continue;
    const before = entries.length;
    passRenames(doc, entries);
    passInfills(doc, entries);
    passSorts(doc, entries);
    if (entries.length > before) rewritten.set(name, partSerializer.serializeToString(doc));
  }
  const out: Record<string, string> = {};
  for (const name of names) out[name] = rewritten.get(name) ?? parts[name];
  return { parts: out, normalized: entries };
}

/**
 * Dry-run preview: enumerates the would-be normalizations for a package
 * without writing anything. Same fail-closed syntax gate as normalizePackage.
 */
export function previewNormalizations(parts: Record<string, string>): string[] {
  const names = Object.keys(parts).sort();
  const entries: string[] = [];
  for (const name of names) {
    const doc = parsePartStrict(name, parts[name]);
    if (!doc) continue;
    passRenames(doc, entries);
    passInfills(doc, entries);
    passSorts(doc, entries);
  }
  return entries;
}
