import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { refreshAstInPlace } from '../../packages/engine/src/ast-refresh';

function fixture(headerRevision = false): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const xml = zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, '<w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>$1');
  zip.updateFile('word/document.xml', Buffer.from(xml));
  if (headerRevision) zip.addFile('word/header1.xml', Buffer.from('<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:ins w:id="40" w:author="Alice"><w:r><w:t>Header</w:t></w:r></w:ins></w:p></w:hdr>'));
  return zip.toBuffer();
}

describe('revision ID allocation', () => {
  it('does not reuse an ID between tracked insertion and tracked replacement', async () => {
    const parser = new DocxParser(fixture());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const [first, second] = ast.sections[0].nodes;
    const result = await engine.executeBatch({ sequential: true, operations: [
      { action: 'trackedChange', operation: 'insert', target: first.id, text: ' inserted' },
      { action: 'replace', target: second.id, text: 'Changed', trackChanges: true },
    ] } as never);
    expect(result.success).toBe(true);
    const dom = parser.getDocumentDom();
    const ids = ['w:ins', 'w:del'].flatMap(tag => Array.from(dom.getElementsByTagName(tag)).map(el => el.getAttribute('w:id')));
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('seeds the counter from revisions in headers', async () => {
    const parser = new DocxParser(fixture(true));
    const ast = await parser.getAST();
    expect(ast.nextChangeId).toBeGreaterThan(40);
  });
  it('keeps the high-water mark after a refresh', async () => {
    const parser = new DocxParser(fixture());
    const ast = await parser.getAST();
    ast.nextChangeId = 100;
    await refreshAstInPlace(ast, parser);
    expect(ast.nextChangeId).toBe(100);
  });

  it('restores the counter when a batch rolls back', async () => {
    const parser = new DocxParser(fixture());
    const ast = await parser.getAST();
    const initial = ast.nextChangeId;
    const engine = new OperationEngine(ast, parser);
    const result = await engine.executeBatch({ sequential: true, operations: [
      { action: 'trackedChange', operation: 'insert', target: ast.sections[0].nodes[0].id, text: ' temporary' },
      { action: 'trackedChange', operation: 'insert', target: 'missing-node', text: ' fails' },
    ] } as never);
    expect(result.success).toBe(false);
    expect(ast.nextChangeId).toBe(initial);
    expect(parser.getDocumentDom().getElementsByTagName('w:ins').length).toBe(0);
  });

});
