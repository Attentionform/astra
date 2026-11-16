/**
 * Relationship elements by namespace and local name. A relationships part may bind the package
 * relationships namespace to a prefix (`<p:Relationship>`); matching the literal tag name would miss every entry.
 */
export function relationshipElements(doc: Document | Element): Element[] {
  return Array.from(doc.getElementsByTagNameNS('*', 'Relationship'));
}

/** A new `Relationship` element written the way the part's root is (same namespace, same prefix). */
export function createRelationshipElement(doc: Document): Element {
  const root = doc.documentElement;
  const ns = root?.namespaceURI;
  if (!ns) return doc.createElement('Relationship');
  const prefix = root.prefix;
  return doc.createElementNS(ns, prefix ? `${prefix}:Relationship` : 'Relationship');
}

/** The `Relationships` root of a relationships part, by namespace and local name. */
export function relationshipsRoot(doc: Document): Element | undefined {
  return doc.getElementsByTagNameNS('*', 'Relationships')[0] as Element | undefined;
}
