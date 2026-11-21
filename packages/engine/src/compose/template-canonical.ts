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

const HIDDEN_BOOKMARK = /^_(Toc|GoBack)/;

/** The paragraph style of the paragraph that holds a node, or an empty string. */
function paragraphStyleOf(node: Node): string {
  let p: Node | null = node;
  while (p && !(p.nodeType === 1 && (p as Element).tagName === 'w:p')) p = p.parentNode;
  if (!p) return '';
  for (let c = p.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (c as Element).tagName === 'w:pPr') {
      const style = Array.from((c as Element).childNodes).find((k) => k.nodeType === 1 && (k as Element).tagName === 'w:pStyle') as Element | undefined;
      return style?.getAttribute('w:val') ?? '';
    }
  }
  return '';
}

/**
 * What mentions a hidden bookmark by name, as text to search: the entries of a table of contents (paragraphs in a TOC style)
 * and everything else. Every attribute value and text node of the document is read, so a hyperlink anchor, a field
 * instruction (complex or simple) and anything else that carries the name count; nothing is looked up by element type.
 * `external` is text from other parts of the package (headers, footers, notes), which count as everything else.
 */
function bookmarkReferences(root: Element, external: string): { toc: string; other: string } {
  const toc: string[] = [];
  const other: string[] = [];
  const add = (node: Node, value: string): void => {
    if (!/_(Toc|GoBack)/.test(value)) return;
    (/^TOC/i.test(paragraphStyleOf(node)) ? toc : other).push(value);
  };
  const walk = (el: Element): void => {
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes[i];
      if (el.tagName === 'w:bookmarkStart' && a.name === 'w:name') continue;
      add(el, a.value);
    }
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) walk(c as Element);
      else if (c.nodeType === 3) add(c, c.nodeValue ?? '');
    }
  };
  walk(root);
  other.push(external);
  return { toc: toc.join('\n'), other: other.join('\n') };
}

const names = (text: string, name: string): boolean => new RegExp(`${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![A-Za-z0-9_])`).test(text);

/**
 * Whether a hidden bookmark is one that Word added by itself and that may be left out of the hash. Word adds `_Toc...` to a
 * heading for an entry of a table of contents: the pair closes in the same paragraph, and only table of contents entries point
 * at it. Word adds `_GoBack` as an empty pair where the cursor was. Any other bookmark was added by a person or a tool, so it
 * counts as content, as does one that a link or field outside a table of contents points at, because deleting it would leave
 * that link broken. Such a bookmark makes the region look edited, and nothing is deleted. `_Ref...` bookmarks are never
 * ignored: a cross-reference to the body is an edit.
 */
function isWordHiddenBookmark(start: Element, refs: { toc: string; other: string }): boolean {
  const name = start.getAttribute('w:name') ?? '';
  if (!HIDDEN_BOOKMARK.test(name)) return false;
  const id = start.getAttribute('w:id');
  let end: Node | null = start.nextSibling;
  let onlyBookmarksBetween = true;
  while (end && !(end.nodeType === 1 && (end as Element).tagName === 'w:bookmarkEnd' && (end as Element).getAttribute('w:id') === id)) {
    if (end.nodeType === 1 || (end.nodeType === 3 && !/^\s*$/.test(end.nodeValue ?? ''))) onlyBookmarksBetween = false;
    end = end.nextSibling;
  }
  // The pair must close in the same paragraph; a bookmark that crosses paragraphs is not one that Word wrote.
  if (!end || (start.parentNode as Element | null)?.tagName !== 'w:p') return false;
  if (names(refs.other, name)) return false;
  if (name.startsWith('_GoBack')) return onlyBookmarksBetween;
  return names(refs.toc, name);
}

/** The canonical XML of each node of a run of siblings. The boundary bookmarks and the hidden bookmarks that Word adds itself are left out. */
export function canonicalNodes(nodes: Node[], externalText = ''): string[] {
  const skip = new Set<string>();
  const root = nodes.find((n) => n.ownerDocument)?.ownerDocument?.documentElement;
  const refs = root ? bookmarkReferences(root, externalText) : { toc: '', other: externalText };
  for (const n of nodes) {
    if (n.nodeType !== 1) continue;
    for (const b of Array.from((n as Element).getElementsByTagName('w:bookmarkStart'))) {
      const name = b.getAttribute('w:name') ?? '';
      if (name === REGION_START || name === REGION_END || isWordHiddenBookmark(b, refs)) skip.add(b.getAttribute('w:id') ?? '');
    }
  }
  return nodes.map((n) => canonicalXml(n, skip));
}

/** `externalText` is the text of the other parts of the package (headers, footers, notes), which may point at a bookmark in these nodes. */
export function hashNodes(nodes: Node[], externalText = ''): string {
  return createHash('sha256').update(canonicalNodes(nodes, externalText).join('')).digest('hex');
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
