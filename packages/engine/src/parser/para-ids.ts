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
 * Live-DOM scan for duplicated w14:paraId values. Groups by exact string
 * (mirroring the session-id map), skips missing/invalid ids. Covers both
 * load-time duplicates and edit-introduced ones with no load-time state to
 * go stale. Header/footer parts are out of scope (never session-mapped).
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
      out.push({ paraId, count: els.length, xPaths: els.slice(1).map(paragraphXPath) });
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
