import { createWordElement } from '../ooxml/ns';

/**
 * Pure DOM-tree helpers for DocxParser (verbatim move from parser.ts).
 * None touch parser instance state: every input arrives as a parameter.
 */
export function nodeHasTrackedChanges(node: Element): boolean {
  const insNodes = node.getElementsByTagName('w:ins');
  const delNodes = node.getElementsByTagName('w:del');
  return insNodes.length > 0 || delNodes.length > 0;
}

export function findDirectChild(parent: Element, tagName: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === tagName) {
      return child as Element;
    }
  }
  return null;
}

export function getOrCreateChild(parent: Element, tagName: string, doc: Document): Element {
  const existing = findDirectChild(parent, tagName);
  if (existing) return existing;
  const created = tagName.startsWith('w:') ? createWordElement(doc, tagName) : doc.createElement(tagName);
  parent.appendChild(created);
  return created;
}

export function unwrapElement(element: Element): void {
  const parent = element.parentNode;
  if (!parent) return;
  while (element.firstChild) {
    parent.insertBefore(element.firstChild, element);
  }
  parent.removeChild(element);
}
