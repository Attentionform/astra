import type { DocumentAST } from '../ast';

/**
 * Numbering sync guard for DocxParser (verbatim move from parser.ts).
 * Pure: the AST arrives as a parameter, no parser instance state.
 */
export function hasLiveNumberingForSync(ast: DocumentAST): boolean {
  try {
    const abstract = (ast.numbering?.abstractNum ?? {}) as Record<string, Element>;
    const concrete = (ast.numbering?.concreteNum ?? {}) as Record<string, Element>;
    for (const numEl of Object.values(concrete)) {
      try {
        const absRef = (numEl as Element).getElementsByTagName?.('w:abstractNumId')?.[0] as Element | undefined;
        const absId = absRef?.getAttribute?.('w:val');
        if (absId != null && absId !== '' && Object.prototype.hasOwnProperty.call(abstract, absId)) {
          return true;
        }
      } catch {
        continue;
      }
    }
    return false;
  } catch {
    return false;
  }
}
