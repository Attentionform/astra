import { toError } from '../errors';
import { WORD_NS } from './ns';

/**
 * Fallback guard (DN3, section 8.2). The engine reads the main document by the usual `w:` prefix. The package
 * ingress rewrites other prefixes before this runs, so for a supported file the guard passes. If an element in the
 * WordprocessingML namespace still has another prefix (or the default namespace), the document is refused with the
 * reason instead of loading as an empty document and having every edit miss it.
 */
export function assertWordPrefix(doc: Document): void {
  const all = doc.getElementsByTagNameNS(WORD_NS, '*');
  for (let i = 0; i < all.length; i++) {
    const prefix = (all[i] as Element).prefix ?? '';
    if (prefix !== 'w') {
      throw toError(
        'E_VALIDATION',
        `Unsupported package: word/document.xml writes WordprocessingML with ${prefix === '' ? 'the default namespace (no prefix)' : `the prefix "${prefix}"`} instead of "w" (first seen on <${(all[i] as Element).tagName}>). ASTra reads the document by the "w" prefix and would see an empty document.`,
        'open the file in Word and save it again (Word writes the "w" prefix), then retry'
      );
    }
  }
}
