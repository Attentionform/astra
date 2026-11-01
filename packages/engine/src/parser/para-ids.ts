import { resolveDeterministicZipDate, resolveReproducibleMode } from '../zip-hardening';

/**
 * A-4: validates an ST_LongHexNumber paraId candidate — 8 hex chars, nonzero,
 * high bit clear (< 0x80000000 per ECMA-376 WordML 2010 usage).
 */
export function isValidParaId(value: string | null | undefined): boolean {
  if (!value || !/^[0-9A-Fa-f]{8}$/.test(value)) return false;
  const num = parseInt(value, 16);
  return num > 0 && num < 0x80000000;
}

export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

/**
 * Ancestor-chain xPath for any element (1-based same-tag ordinals, e.g.
 * `/w:document[1]/w:body[1]/w:p[2]`). Never throws (falls back to `//w:p`).
 * Shared by the duplicate-paraId scan, the structural tag sweep, and the
 * normalizer detail log so every surface names one addressable location.
 */
export function elementXPath(node: Element): string {
  try {
    const parts: string[] = [];
    let el: Element | null = node;
    while (el && el.nodeType === 1) {
      const tag = el.tagName || 'w:p';
      let ordinal = 1;
      let sib = el.previousSibling;
      while (sib) {
        if (sib.nodeType === 1 && (sib as Element).tagName === tag) ordinal++;
        sib = sib.previousSibling;
      }
      parts.unshift(`${tag}[${ordinal}]`);
      const parent = el.parentNode as Element | null;
      if (!parent || parent.nodeType !== 1) break;
      el = parent;
    }
    return parts.length > 0 ? `/${parts.join('/')}` : '//w:p';
  } catch {
    return '//w:p';
  }
}

/** Paragraph specialization of elementXPath (same output, documents intent). */
export function paragraphXPath(node: Element): string {
  return elementXPath(node);
}

/** One natively duplicated w14:paraId: the id, repeat count, repeat xPaths. */
export interface DuplicateParaId {
  paraId: string;
  count: number;
  xPaths: string[];
}

/**
 * True when the two paragraphs sit in different branches of one
 * mc:AlternateContent (their closest shared ancestor is that element). Only one
 * branch is ever read, and Word writes the same paragraph ids in the mc:Choice
 * copy and in the mc:Fallback copy of a text box, so such a pair is one
 * paragraph written twice and not a collision.
 */
function areAlternativeBranches(a: Element, b: Element): boolean {
  const ancestorsOfA = new Set<Node>();
  for (let node = a.parentNode; node; node = node.parentNode) ancestorsOfA.add(node);
  for (let node = b.parentNode; node; node = node.parentNode) {
    if (ancestorsOfA.has(node)) return node.nodeType === 1 && (node as Element).tagName === 'mc:AlternateContent';
  }
  return false;
}

/** Largest group of equal ids whose pairs are all checked; a bigger group is reported as it was. */
const MAX_MIRRORED_GROUP = 32;

function allAlternativeBranches(group: Element[]): boolean {
  if (group.length > MAX_MIRRORED_GROUP) return false;
  for (let i = 0; i < group.length; i++) {
    for (let j = i + 1; j < group.length; j++) {
      if (!areAlternativeBranches(group[i], group[j])) return false;
    }
  }
  return true;
}

/**
 * Live-DOM scan for duplicated w14:paraId values. Groups by exact string
 * (mirroring the session-id map), skips missing/invalid ids. Covers both
 * load-time duplicates and edit-introduced ones with no load-time state to
 * go stale. Header/footer parts are out of scope (never session-mapped).
 * An id shared only by the alternative branches of an mc:AlternateContent is
 * not a duplicate; every other repeat still is.
 */
export function findDuplicateParaIds(documentDom: Document): DuplicateParaId[] {
  const out: DuplicateParaId[] = [];
  try {
    const all = documentDom.getElementsByTagName('w:p');
    const seen = new Map<string, Element[]>();
    for (let i = 0; i < all.length; i++) {
      const el = all[i] as unknown as Element;
      const v = el.getAttribute?.('w14:paraId');
      if (!v || !isValidParaId(v)) continue;
      const group = seen.get(v);
      if (group) group.push(el);
      else seen.set(v, [el]);
    }
    for (const [paraId, els] of seen) {
      if (els.length < 2) continue;
      if (allAlternativeBranches(els)) continue;
      out.push({ paraId, count: els.length, xPaths: els.slice(1).map(paragraphXPath) });
    }
  } catch {
    return out;
  }
  return out;
}

/**
 * The paragraphs a renumber changes: in every group of equal ids the first
 * paragraph (document order) keeps its id, and so does a later one that is the
 * other branch of an mc:AlternateContent with a paragraph that kept it. Every
 * other paragraph of the group is returned with its id and its xPath. A document
 * for which this returns nothing is one the duplicate guard accepts.
 */
export function findParagraphsToRenumber(documentDom: Document): { element: Element; paraId: string; xPath: string }[] {
  const out: { element: Element; paraId: string; xPath: string }[] = [];
  try {
    const all = documentDom.getElementsByTagName('w:p');
    const seen = new Map<string, Element[]>();
    for (let i = 0; i < all.length; i++) {
      const el = all[i] as unknown as Element;
      const v = el.getAttribute?.('w14:paraId');
      if (!v || !isValidParaId(v)) continue;
      const group = seen.get(v);
      if (group) group.push(el);
      else seen.set(v, [el]);
    }
    for (const [paraId, els] of seen) {
      if (els.length < 2 || allAlternativeBranches(els)) continue;
      const kept: Element[] = [els[0]];
      for (const el of els.slice(1)) {
        if (els.length <= MAX_MIRRORED_GROUP && kept.some((k) => areAlternativeBranches(k, el))) kept.push(el);
        else out.push({ element: el, paraId, xPath: paragraphXPath(el) });
      }
    }
  } catch {
    return out;
  }
  return out;
}

/**
 * Clock/seed injection for snapshot + node ids (test hook + deterministic default).
 * Injected clock returns ms since epoch; injected random returns [0,1).
 * When no injection is set, reproducible mode pins to the deterministic epoch
 * and opt-out falls back to wall-clock. Exported for tests.
 */
let injectedParserClock: (() => number) | null = null;
let injectedParserRandom: (() => number) | null = null;
let parserDeterministicSeq = 0;

export function setParserClockForTests(fn: (() => number) | null): void {
  injectedParserClock = fn;
}

export function setParserRandomForTests(fn: (() => number) | null): void {
  injectedParserRandom = fn;
}

export function resetParserClockForTests(): void {
  injectedParserClock = null;
  injectedParserRandom = null;
  parserDeterministicSeq = 0;
}

export function parserNowMs(): number {
  if (injectedParserClock) return injectedParserClock();
  if (resolveReproducibleMode(undefined)) return resolveDeterministicZipDate().getTime();
  return Date.now();
}

export function parserRandomUnit(): number {
  if (injectedParserRandom) return injectedParserRandom();
  if (resolveReproducibleMode(undefined)) {
    parserDeterministicSeq += 1;
    const hash = ((parserDeterministicSeq * 2654435761) >>> 0) / 4294967296;
    return hash === 0 ? 0.5 : hash;
  }
  return Math.random();
}
