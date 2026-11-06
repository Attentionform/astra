import { type DocumentAST, type DocxParser } from '@attentionform/astra-engine';

/**
 * Give a serialization fork the live numbering, including its original part
 * metadata. The engine uses the source hashes to retain untouched bytes and
 * the root attributes/extra children when definitions really have changed.
 * Clone every numbering element and source record so fork serialization cannot
 * change the session. Pass only numbering/relationships on purpose: the bare
 * rezip path reads only those fields and must not receive other live AST data.
 */
export function adoptForkNumbering(fork: DocxParser, ast: DocumentAST): void {
  const source = ast.numbering?.source;
  const cloneDefinitions = (definitions: Record<string, Element>) => Object.fromEntries(
    Object.entries(definitions).map(([id, element]) => [id, element.cloneNode(true) as Element]),
  );
  fork.adoptAstForNumbering({
    numbering: {
      abstractNum: cloneDefinitions(ast.numbering?.abstractNum ?? {}),
      concreteNum: cloneDefinitions(ast.numbering?.concreteNum ?? {}),
      // Mirrors the NumberingSource clone in engine ast.ts (cloneAST); keep both
      // in step when NumberingSource gains fields.
      ...(source ? {
        source: {
          ...source,
          rootAttributes: source.rootAttributes.map(attribute => ({ ...attribute })),
          leading: [...source.leading],
          trailing: [...source.trailing],
        },
      } : {}),
    },
    relationships: Object.fromEntries(
      Object.entries(ast.relationships ?? {}).map(([id, relationship]) => [id, { ...relationship }]),
    ),
  } as unknown as Parameters<DocxParser['adoptAstForNumbering']>[0]);
}
