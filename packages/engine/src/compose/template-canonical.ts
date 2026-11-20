import { createHash } from 'crypto';

// Shared by the template generator (which records the expected hashes) and the compose shortcut (which checks them).
// Keep this module free of imports from the rest of the engine so the generator can bundle it alone.

/** Attributes that change without the content changing: namespace declarations, revision ids, paragraph ids, ASTra ids. */
// xml:space is included because Word drops it from text that has no leading or trailing space.
const VOLATILE_ATTRIBUTE = /^(xmlns(:|$)|w:rsid|w14:paraId$|w14:textId$|data-astra-id$|xml:space$)/;

const NO_IDS: ReadonlySet<string> = new Set();

const escapeText = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (text: string): string => escapeText(text).replace(/"/g, '&quot;');

/** A run that holds only formatting and text: its canonical formatting and its text. Anything else is not a plain text run. */
function plainTextRun(node: Node, skip: ReadonlySet<string>): { rPr: string; text: string } | null {
  if (node.nodeType !== 1 || (node as Element).tagName !== 'w:r') return null;
  const run = node as Element;
  let rPr = '';
  let text = '';
  for (let c = run.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 3) {
      if (/^\s*$/.test(c.nodeValue ?? '')) continue;
      return null;
    }
    if (c.nodeType !== 1) continue;
    const child = c as Element;
    if (child.tagName === 'w:rPr') rPr = canonicalXml(child, skip);
    else if (child.tagName === 'w:t') text += child.textContent ?? '';
    else if (child.tagName === 'w:lastRenderedPageBreak') continue;
    else return null;
  }
  for (let i = 0; i < run.attributes.length; i++) if (!VOLATILE_ATTRIBUTE.test(run.attributes[i].name)) return null;
  return { rPr, text };
}

/**
 * Canonical XML of a node: every element, attribute and text node, attributes sorted, the volatile attributes above left out,
 * and formatting white space between elements dropped. Anything the function does not recognise is still included, so
 * unknown structure changes the result.
 */
export function canonicalXml(node: Node, skipBookmarkIds: ReadonlySet<string> = NO_IDS): string {
  if (node.nodeType === 3) {
    const text = node.nodeValue ?? '';
    const parent = (node.parentNode as Element | null)?.localName;
    if (/^\s*$/.test(text) && parent !== 't' && parent !== 'instrText') return '';
    return escapeText(text);
  }
  if (node.nodeType !== 1) return '';
  const el = node as Element;
  // Word writes this marker of where a page break fell when it last laid the page out; it is not content.
  if (el.tagName === 'w:lastRenderedPageBreak') return '';
  // Word records which table-style conditions apply to a row, cell or paragraph; the table style already implies it.
  if (el.tagName === 'w:cnfStyle') return '';
  // The boundary bookmarks are not content: Word moves a body-level bookmark into the neighbouring paragraph when it saves.
  if ((el.tagName === 'w:bookmarkStart' || el.tagName === 'w:bookmarkEnd') && skipBookmarkIds.has(el.getAttribute('w:id') ?? '')) return '';
  const attrs: Array<[string, string]> = [];
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (VOLATILE_ATTRIBUTE.test(a.name)) continue;
    // Word recomputes the legacy hexadecimal w:val of w:tblLook from the flag attributes next to it.
    if (el.tagName === 'w:tblLook' && a.name === 'w:val') continue;
    attrs.push([a.name, a.value]);
  }
  attrs.sort((x, y) => (x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0));
  const open = `<${el.tagName}${attrs.map(([n, v]) => ` ${n}="${escapeAttr(v)}"`).join('')}`;
  const kids: string[] = [];
  // Word splits one run into several runs of identical formatting (for example at a spelling-check boundary), so in a paragraph
  // consecutive text runs with the same formatting are one run.
  let pending: { rPr: string; text: string } | null = null;
  const flush = (): void => {
    if (pending) {
      const body = `${pending.rPr}${pending.text === '' ? '' : `<w:t>${escapeText(pending.text)}</w:t>`}`;
      kids.push(body === '' ? '<w:r/>' : `<w:r>${body}</w:r>`);
    }
    pending = null;
  };
  for (let c = el.firstChild; c; c = c.nextSibling) {
    const run = el.tagName === 'w:p' ? plainTextRun(c, skipBookmarkIds) : null;
    if (run) {
      if (pending && pending.rPr === run.rPr) pending.text += run.text;
      else {
        flush();
        pending = { ...run };
      }
      continue;
    }
    flush();
    kids.push(canonicalXml(c, skipBookmarkIds));
  }
  flush();
  const inner = kids.join('');
  return inner === '' ? `${open}/>` : `${open}>${inner}</${el.tagName}>`;
}

/** SHA-256 (hex) of the canonical XML of a run of sibling nodes. */
export function canonicalNodes(nodes: Node[]): string[] {
  const skip = new Set<string>();
  for (const n of nodes) {
    if (n.nodeType !== 1) continue;
    for (const b of Array.from((n as Element).getElementsByTagName('w:bookmarkStart'))) {
      // The boundary bookmarks, and the hidden bookmarks Word adds itself (a table of contents adds _Toc..., a cross-reference _Ref...).
      const name = b.getAttribute('w:name') ?? '';
      if (name === REGION_START || name === REGION_END || /^_(Toc|Ref|GoBack)/.test(name)) skip.add(b.getAttribute('w:id') ?? '');
    }
  }
  return nodes.map((n) => canonicalXml(n, skip));
}

export function hashNodes(nodes: Node[]): string {
  return createHash('sha256').update(canonicalNodes(nodes).join('')).digest('hex');
}

export const REGION_START = 'astra_body_start';
export const REGION_END = 'astra_body_end';
export const SLOT_PREFIX = 'astra_slot_';

/**
 * The placeholder body: the body-level blocks from the one that holds (or follows) `astra_body_start` to the one that holds
 * (or precedes) `astra_body_end`. The templates hold each marker as an empty bookmarkStart/bookmarkEnd pair inside a paragraph
 * (the start at the beginning of the first paragraph, the end at the end of the last one), which is where Word keeps it when it
 * saves the file; a pair between two body-level blocks is read too. Null unless there is exactly one ordered pair.
 */
export function findRegion(body: Element): { start: Element; end: Element; nodes: Element[] } | null {
  const marked = (name: string): Element[] => Array.from(body.getElementsByTagName('w:bookmarkStart')).filter((b) => b.getAttribute('w:name') === name);
  const starts = marked(REGION_START);
  const ends = marked(REGION_END);
  if (starts.length !== 1 || ends.length !== 1) return null;
  const start = starts[0];
  const end = ends[0];
  // Each marker is a bookmarkStart with its bookmarkEnd in the same parent.
  const endOf = (s: Element): Element | null => {
    for (let c = s.nextSibling; c; c = c.nextSibling) {
      if (c.nodeType === 1 && (c as Element).tagName === 'w:bookmarkEnd' && (c as Element).getAttribute('w:id') === s.getAttribute('w:id')) return c as Element;
    }
    return null;
  };
  const startEnd = endOf(start);
  if (!startEnd || !endOf(end)) return null;
  const kids: Element[] = [];
  for (let c = body.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) kids.push(c as Element);
  const blockOf = (el: Element): number => {
    let n: Node = el;
    while (n.parentNode && n.parentNode !== body) n = n.parentNode;
    return n.parentNode === body ? kids.indexOf(n as Element) : -1;
  };
  const si = start.parentNode === body ? kids.indexOf(startEnd) + 1 : blockOf(start);
  const ei = end.parentNode === body ? kids.indexOf(end) - 1 : blockOf(end);
  if (si < 0 || ei < si || ei >= kids.length) return null;
  return { start, end, nodes: kids.slice(si, ei + 1) };
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
