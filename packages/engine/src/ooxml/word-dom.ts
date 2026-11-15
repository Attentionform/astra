/**
 * Namespace-aware access to WordprocessingML elements and attributes. A package may bind the main namespace to
 * any prefix (`<x:tbl x:val="1">`); code that compares qualified names such as `w:tbl` only sees the usual form.
 * These helpers take the usual `w:name` spelling and match by namespace URI and local name, so both read the same.
 */
import { WORD_NS } from './ns';

const localOf = (qualified: string): string => qualified.replace(/^[^:]+:/, '');

/** True when `node` is the element `qualified` (`w:tc`) in the WordprocessingML namespace, whatever its prefix. */
export function isWordElement(node: Node | null | undefined, qualified: string): boolean {
  if (!node || node.nodeType !== 1) return false;
  const el = node as Element;
  if (el.namespaceURI) return el.namespaceURI === WORD_NS && el.localName === localOf(qualified);
  return el.tagName === qualified;
}

/** Direct children that are the element `qualified`, matched by namespace and local name. */
export function wordChildren(parent: Element, qualified: string): Element[] {
  const out: Element[] = [];
  for (let c = parent.firstChild; c; c = c.nextSibling) {
    if (isWordElement(c, qualified)) out.push(c as Element);
  }
  return out;
}

/** Key used by the schema-order tables: `w:` plus the local name for a WordprocessingML element, else the tag name. */
export function schemaKey(el: Element): string {
  return el.namespaceURI === WORD_NS && el.localName ? `w:${el.localName}` : el.tagName;
}

function wordAttribute(el: Element, qualified: string): Attr | null {
  const local = localOf(qualified);
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (a.namespaceURI ? a.namespaceURI === WORD_NS && a.localName === local : a.name === qualified) return a;
  }
  return null;
}

/** Value of the WordprocessingML attribute `qualified` (`w:val`), whatever its prefix; null when absent. */
export function getWordAttr(el: Element | undefined | null, qualified: string): string | null {
  if (!el) return null;
  return wordAttribute(el, qualified)?.value ?? null;
}

/** The prefix bound to the main namespace in scope at `el` (null when it is the default namespace). */
export function wordPrefixAt(el: Element): string | null {
  if (el.namespaceURI === WORD_NS) return el.prefix ?? null;
  return el.lookupPrefix?.(WORD_NS) ?? 'w';
}

/** Sets the WordprocessingML attribute `qualified`; an existing attribute keeps its own prefix. */
export function setWordAttr(el: Element, qualified: string, value: string): void {
  const existing = wordAttribute(el, qualified);
  if (existing) {
    existing.value = value;
    return;
  }
  const prefix = wordPrefixAt(el) ?? 'w';
  el.setAttributeNS(WORD_NS, `${prefix}:${localOf(qualified)}`, value);
}

/** A new WordprocessingML element written with the same prefix as `like`. */
export function createWordLike(like: Element, qualified: string): Element {
  const prefix = wordPrefixAt(like);
  const local = localOf(qualified);
  return like.ownerDocument.createElementNS(WORD_NS, prefix ? `${prefix}:${local}` : local);
}
