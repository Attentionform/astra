import { createHash } from 'crypto';

// Shared by the template generator (which records the expected hashes) and the compose shortcut (which checks them).
// Keep this module free of imports from the rest of the engine so the generator can bundle it alone.

/** Attributes that change without the content changing: namespace declarations, revision ids, paragraph ids, ASTra ids. */
const VOLATILE_ATTRIBUTE = /^(xmlns(:|$)|w:rsid|w14:paraId$|w14:textId$|data-astra-id$)/;

const escapeText = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (text: string): string => escapeText(text).replace(/"/g, '&quot;');

/**
 * Canonical XML of a node: every element, attribute and text node, attributes sorted, the volatile attributes above left out,
 * and formatting white space between elements dropped. Anything the function does not recognise is still included, so
 * unknown structure changes the result.
 */
export function canonicalXml(node: Node): string {
  if (node.nodeType === 3) {
    const text = node.nodeValue ?? '';
    const parent = (node.parentNode as Element | null)?.localName;
    if (/^\s*$/.test(text) && parent !== 't' && parent !== 'instrText') return '';
    return escapeText(text);
  }
  if (node.nodeType !== 1) return '';
  const el = node as Element;
  const attrs: Array<[string, string]> = [];
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (!VOLATILE_ATTRIBUTE.test(a.name)) attrs.push([a.name, a.value]);
  }
  attrs.sort((x, y) => (x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0));
  const open = `<${el.tagName}${attrs.map(([n, v]) => ` ${n}="${escapeAttr(v)}"`).join('')}`;
  const kids: string[] = [];
  for (let c = el.firstChild; c; c = c.nextSibling) kids.push(canonicalXml(c));
  const inner = kids.join('');
  return inner === '' ? `${open}/>` : `${open}>${inner}</${el.tagName}>`;
}

/** SHA-256 (hex) of the canonical XML of a run of sibling nodes. */
export function hashNodes(nodes: Node[]): string {
  return createHash('sha256').update(nodes.map(canonicalXml).join('')).digest('hex');
}

export const REGION_START = 'astra_body_start';
export const REGION_END = 'astra_body_end';
export const SLOT_PREFIX = 'astra_slot_';

const isNamedBookmarkStart = (n: Node, name: string): boolean => n.nodeType === 1 && (n as Element).tagName === 'w:bookmarkStart' && (n as Element).getAttribute('w:name') === name;

/** The body-level siblings strictly between the two boundary bookmarks, or null when the markers are not exactly one ordered pair. */
export function findRegion(body: Element): { start: Element; end: Element; nodes: Element[] } | null {
  const kids: Element[] = [];
  for (let c = body.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) kids.push(c as Element);
  const starts = kids.filter((k) => isNamedBookmarkStart(k, REGION_START));
  const ends = kids.filter((k) => isNamedBookmarkStart(k, REGION_END));
  const all = Array.from(body.getElementsByTagName('w:bookmarkStart'));
  if (starts.length !== 1 || ends.length !== 1) return null;
  if (all.filter((b) => b.getAttribute('w:name') === REGION_START || b.getAttribute('w:name') === REGION_END).length !== 2) return null;
  const si = kids.indexOf(starts[0]);
  const ei = kids.indexOf(ends[0]);
  if (si < 0 || ei < si) return null;
  const startEnd = kids[si + 1];
  const endEnd = kids[ei + 1];
  // Each boundary is an empty bookmarkStart/bookmarkEnd pair: the bookmarkEnd follows its start directly.
  const isEndFor = (el: Element | undefined, id: string | null): boolean => !!el && el.tagName === 'w:bookmarkEnd' && el.getAttribute('w:id') === id;
  if (!isEndFor(startEnd, starts[0].getAttribute('w:id')) || !isEndFor(endEnd, ends[0].getAttribute('w:id'))) return null;
  return { start: starts[0], end: ends[0], nodes: kids.slice(si + 2, ei) };
}

/** A slot: the nodes between `astra_slot_<name>` bookmarkStart and its bookmarkEnd, inside one paragraph. */
export function findSlot(body: Element, name: string): { paragraph: Element; start: Element; end: Element; nodes: Node[] } | null {
  const starts = Array.from(body.getElementsByTagName('w:bookmarkStart')).filter((b) => b.getAttribute('w:name') === `${SLOT_PREFIX}${name}`);
  if (starts.length !== 1) return null;
  const start = starts[0];
  const id = start.getAttribute('w:id');
  const nodes: Node[] = [];
  let c: Node | null = start.nextSibling;
  while (c) {
    if (c.nodeType === 1 && (c as Element).tagName === 'w:bookmarkEnd' && (c as Element).getAttribute('w:id') === id) {
      const paragraph = start.parentNode as Element;
      if (paragraph.tagName !== 'w:p') return null;
      return { paragraph, start, end: c as Element, nodes };
    }
    nodes.push(c);
    c = c.nextSibling;
  }
  return null;
}
