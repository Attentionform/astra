import type { DocumentAST, Node as AstNode } from '../ast';

/**
 * Pure AST lookup for DocxParser (verbatim move from parser.ts).
 * No parser instance state: the AST arrives as a parameter.
 */
export function getNodeFromAST(ast: DocumentAST, id: string): AstNode | undefined {
  for (const sec of ast.sections) {
    const found = sec.nodes.find(n => n.id === id);
    if (found) return found;
  }
  return undefined;
}
