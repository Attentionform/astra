/**
 * Relationship/content-type mapping helpers for DocxParser save sync
 * (verbatim move from parser.ts). Pure: no parser instance state.
 */
export function partNameForRelationshipTarget(target: string): string | null {
  if (!target || target.includes('..')) return null;
  // External hyperlink-style targets are not package parts.
  if (/^https?:\/\//i.test(target) || target.startsWith('#')) return null;
  if (target.startsWith('/')) return target;
  if (target.startsWith('word/')) return `/${target}`;
  // document.xml.rels targets resolve relative to word/
  return `/word/${target}`;
}

export function contentTypeForRelationship(relType: string, target: string): string | null {
  const lowerType = relType.toLowerCase();
  const lowerTarget = target.toLowerCase();
  if (lowerType.endsWith('/officeDocument') || lowerTarget.endsWith('/document.xml')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
  }
  if (lowerType.endsWith('/styles') || lowerTarget.endsWith('/styles.xml')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
  }
  if (lowerType.endsWith('/settings') || lowerTarget.endsWith('/settings.xml')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
  }
  if (lowerType.endsWith('/websettings') || lowerTarget.endsWith('/websettings.xml')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml';
  }
  if (lowerType.endsWith('/glossarydocument') || lowerTarget.includes('glossary/')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml';
  }
  if (lowerType.endsWith('/fonttable') || lowerTarget.includes('fonttable')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml';
  }
  if (lowerType.endsWith('/chart') || lowerTarget.includes('charts/')) {
    return 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
  }
  if (lowerType.endsWith('/header') || lowerTarget.includes('header')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
  }
  if (lowerType.endsWith('/footer') || lowerTarget.includes('footer')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
  }
  if (lowerType.endsWith('/footnotes') || lowerTarget.includes('footnotes')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml';
  }
  if (lowerType.endsWith('/endnotes') || lowerTarget.includes('endnotes')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml';
  }
  // commentsExtended first: its target also contains the substring "comments".
  if (lowerType.endsWith('/commentsextended') || lowerTarget.includes('commentsextended')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
  }
  if (lowerType.endsWith('/comments') || lowerTarget.includes('comments')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
  }
  if (lowerType.endsWith('/people') || lowerTarget.includes('people')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml';
  }
  if (lowerType.endsWith('/numbering') || lowerTarget.includes('numbering')) {
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
  }
  if (lowerType.endsWith('/theme') || lowerTarget.includes('theme')) {
    return 'application/vnd.openxmlformats-officedocument.theme+xml';
  }
  if (lowerType.endsWith('/embeddings') || lowerType.includes('oleobject') || lowerTarget.includes('embeddings/')) {
    return 'application/vnd.openxmlformats-officedocument.oleObject';
  }
  // No generic application/xml fallback: stamping an explicit Override with
  // the generic type onto word/* parts (styles.xml, settings.xml, …)
  // actively breaks strict validation — the validator expects each part's
  // specific content type. Unmapped custom parts stay covered by the
  // Default Extension="xml" entry, so skipping the Override is lossless.
  return null;
}
