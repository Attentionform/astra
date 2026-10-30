import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  createBlankDocxBuffer,
} from '@attentionform/astra-engine';
import { contentTypeForRelationship } from '../../packages/engine/src/parser/rel-utils';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function runBatch(build: (id: string) => unknown[]) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const id = ast.sections[0].nodes[0].id;
  const result = await engine.executeBatch({ operations: build(id) as never });
  const bytes = parser.rezip();
  const rels = new AdmZip(bytes).readAsText('word/_rels/document.xml.rels');
  const ids = rels.match(/Id="[^"]+"/g) ?? [];
  return { result, bytes, duplicateRelationshipIds: ids.length - new Set(ids).size };
}

describe('comment insert followed by other relationship writers in one batch', () => {
  const comment = (id: string) => ({ action: 'comment', operation: 'insert', target: id, text: 'c' });
  const cases: Array<[string, (id: string) => unknown[]]> = [
    ['footnote', (id) => [comment(id), { action: 'footnote', operation: 'insert', anchor: id, text: 'n' }]],
    ['endnote', (id) => [comment(id), { action: 'endnote', operation: 'insert', anchor: id, text: 'e' }]],
    ['image', (id) => [comment(id), { action: 'image', operation: 'insert', anchor: id, fileData: PNG, altText: 'x' }]],
  ];
  it.each(cases)('comment then %s gives unique relationship ids and a valid package', async (_name, build) => {
    const { result, bytes, duplicateRelationshipIds } = await runBatch(build);
    expect(result.success).toBe(true);
    expect(duplicateRelationshipIds).toBe(0);
    expect((await validateDocxBuffer(bytes)).ok).toBe(true);
  });
});

describe('content type for the comment parts', () => {
  it('maps commentsExtended to its own content type, not the comments one', () => {
    expect(
      contentTypeForRelationship('http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml')
    ).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml');
    expect(
      contentTypeForRelationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml')
    ).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  });
});
