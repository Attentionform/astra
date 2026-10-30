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
 *   (TCPR_CHILD_ORDER et al). Unknown tags sort last (rank 999, stable, so
 *   their relative order is preserved and none is ever dropped).
 *
 * Invariants: idempotent (second run is a fixpoint with an empty log);
 * every rewrite is logged per item with its xPath, enclosing node id when
 * known, and before→after detail (so migration audits can name each
 * healing); clean packages yield an empty list and byte-identical output
 * (parts with no transforms keep their exact input string — no
 * parse/serialize round-trip).
 */

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
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

/** Unknown-tag sort rank: after every known schema child, stable. */
const UNKNOWN_TAG_RANK = 999;

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
 * Unknown tags rank last (stable — relative order preserved, never dropped).
 * Non-element nodes (whitespace) stay pinned at their indices. Returns the
 * before/after child-tag sequences plus whether the container moved.
 */
function sortContainerChildren(
  container: Element,
  order: readonly string[]
): { reordered: boolean; before: string[]; after: string[] } {
  const nodes: Node[] = [];
  for (let i = 0; i < container.childNodes.length; i++) nodes.push(container.childNodes[i]);
  const slots: number[] = [];
  const elements: Element[] = [];
  for (let i = 0; i < nodes.length; i++) {
    if (nodes[i].nodeType === 1) {
      slots.push(i);
      elements.push(nodes[i] as Element);
    }
  }
  const before = elements.map((el) => el.tagName);
  if (elements.length <= 1) return { reordered: false, before, after: before };
  const rank = (tag: string): number => {
    const idx = order.indexOf(tag);
    return idx === -1 ? UNKNOWN_TAG_RANK : idx;
  };
  const sorted = elements
    .map((el, pos) => ({ el, pos }))
    .sort((a, b) => rank(a.el.tagName) - rank(b.el.tagName) || a.pos - b.pos);
  let changed = false;
  for (let i = 0; i < elements.length; i++) {
    if (sorted[i].el !== elements[i]) {
      changed = true;
      break;
    }
  }
  if (!changed) return { reordered: false, before, after: before };
  const next = nodes.slice();
  for (let i = 0; i < slots.length; i++) next[slots[i]] = sorted[i].el;
  while (container.firstChild) container.removeChild(container.firstChild);
  for (const node of next) container.appendChild(node);
  return { reordered: true, before, after: sorted.map((s) => s.el.tagName) };
}

/** Pass 3 driver over the four whitelisted containers. */
function passSorts(doc: Document, entries: string[]): void {
  for (const { tag, order } of SORT_CONTAINERS) {
    for (const container of snapshotElements(doc, tag)) {
      const { reordered, before, after } = sortContainerChildren(container, order);
      if (reordered) {
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
function parsePartStrict(partName: string, xml: string): Document | null {
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
