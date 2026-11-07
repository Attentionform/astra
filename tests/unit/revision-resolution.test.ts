import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { sanitizeDocxBuffer } from '../../packages/engine/src/sanitizer';

function documentWithRevisions(overrideBody?: string): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const body = overrideBody ?? '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="4" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>Example</w:t></w:r></w:p>';
  const xml = zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
  zip.updateFile('word/document.xml', Buffer.from(xml));
  return zip.toBuffer();
}

describe('revision resolution', () => {
  it.each(['acceptAll', 'rejectAll'] as const)('%s resolves property changes and refreshes the agent view', async mode => {
    const parser = new DocxParser(documentWithRevisions());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const result = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: mode }] } as never);
    expect(result.success).toBe(true);
    const dom = parser.getDocumentDom();
    expect(dom.getElementsByTagName('w:rPrChange').length).toBe(0);
    const props = dom.getElementsByTagName('w:rPr')[0];
    expect(props.getElementsByTagName(mode === 'acceptAll' ? 'w:b' : 'w:i').length).toBe(1);
    expect(props.getElementsByTagName(mode === 'acceptAll' ? 'w:i' : 'w:b').length).toBe(0);
    expect(result.revisionsPending).toBe(0);
  });

  it('updates search text when a tracked replacement is rejected', async () => {
    const parser = new DocxParser(documentWithRevisions());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const id = ast.sections[0].nodes[0].id;
    await engine.executeBatch({ operations: [{ action: 'replace', target: id, text: 'Replacement', trackChanges: true }] } as never);
    await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'rejectAll' }] } as never);
    expect(ast.sections[0].nodes[0].metadata.text).toBe('Example');
  });
  it('reports revisions that cannot be resolved rather than claiming zero remain', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const xml = zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, '<w:body><w:p><w:moveFrom w:id="8" w:author="Alice"><w:r><w:t>Unpaired</w:t></w:r></w:moveFrom></w:p>$1');
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    const nodeId = ast.sections[0].nodes[0].id;
    const engine = new OperationEngine(ast, parser);
    const result = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
    expect(result.success).toBe(true);
    expect(result.revisionsPending).toBe(1);
    expect(result.revisionResolution?.unresolved).toEqual(expect.arrayContaining([expect.objectContaining({kind: 'moveFrom', nodeId, part:'word/document.xml'})]));
    expect(parser.getDocumentDom().getElementsByTagName('w:moveFrom').length).toBe(1);
  });

  it('accepts a deleted paragraph mark by merging the two paragraphs', async () => {
    const parser = new DocxParser(documentWithRevisions('<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Alice"/></w:rPr></w:pPr><w:r><w:t>First </w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>'));
    const ast = await parser.getAST();
    const [removed, surviving] = ast.sections[0].nodes.map(node => node.id);
    const result = await new OperationEngine(ast, parser).executeBatch({ operations: [{action:'trackedChange', operation:'acceptAll'}] } as never);
    expect(result.success).toBe(true);
    expect(result.revisionsPending).toBe(0);
    expect(result.revisionResolution?.removedNodeIds).toContain(removed);
    expect(ast.sections.flatMap(section => section.nodes)[0].id).toBe(surviving);
    const paragraphs = parser.getDocumentDom().getElementsByTagName('w:p');
    expect(paragraphs.length).toBe(1);
    expect(paragraphs[0].textContent).toBe('First Second');
    expect(ast.sections.flatMap(section => section.nodes)).toHaveLength(1);
    expect(ast.sections.flatMap(section => section.nodes)[0].metadata.text).toBe('First Second');
  });

  it('accepts a complete named move as a single group', async () => {
    const body = '<w:p><w:moveFromRangeStart w:id="1" w:name="moveA"/><w:moveFrom w:id="2" w:author="Alice"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/><w:r><w:t> middle </w:t></w:r><w:moveToRangeStart w:id="3" w:name="moveA"/><w:moveTo w:id="4" w:author="Alice"><w:r><w:t>new</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({ operations: [{action:'trackedChange', operation:'acceptAll'}] } as never);
    expect(result.success).toBe(true);
    expect(result.revisionsPending).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:p')[0].textContent).toBe(' middle new');
    expect(parser.getDocumentDom().getElementsByTagName('w:moveFromRangeStart').length).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:moveToRangeEnd').length).toBe(0);
  });

  it('expands a selected move side to both sides', async () => {
    const body = '<w:p><w:moveFromRangeStart w:id="1" w:name="moveA"/><w:moveFrom w:id="2" w:author="Alice"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/><w:moveToRangeStart w:id="3" w:name="moveA"/><w:moveTo w:id="4" w:author="Alice"><w:r><w:t>new</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'accept',ids:[2]}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionsPending).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:p')[0].textContent).toBe('new');
    expect(parser.getDocumentDom().getElementsByTagName('w:moveFromRangeStart').length).toBe(0);
  });

  it('refuses duplicate IDs across body and header without changing either part', async () => {
    const zip = new AdmZip(documentWithRevisions('<w:p><w:ins w:id="7" w:author="Alice"><w:r><w:t>Body</w:t></w:r></w:ins></w:p>'));
    zip.addFile('word/header1.xml', Buffer.from('<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:ins w:id="7" w:author="Alice"><w:r><w:t>Header</w:t></w:r></w:ins></w:p></w:hdr>'));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    const body = parser.getDocumentDom().toString();
    const header = parser.getZipAdapter().readText('word/header1.xml');
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'accept',ids:[7]}]} as never);
    expect(result.success).toBe(false);
    expect(result.errors.some(e => e.error.includes('ambiguous'))).toBe(true);
    expect(parser.getDocumentDom().toString()).toBe(body);
    expect(parser.getZipAdapter().readText('word/header1.xml')).toBe(header);
  });

  it('accepts a deleted row while preserving the rest of the table', async () => {
    const body = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:trPr><w:del w:id="10" w:author="Alice"/></w:trPr><w:tc><w:p><w:r><w:t>Remove</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({ operations: [{action:'trackedChange', operation:'acceptAll'}] } as never);
    expect(result.success).toBe(true);
    expect(result.revisionsPending).toBe(0);
    const rows = parser.getDocumentDom().getElementsByTagName('w:tr');
    expect(rows.length).toBe(1);
    expect(rows[0].textContent).toBe('Keep');
  });

  it('restores prior table grid columns on rejection', async () => {
    const body = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:tblGridChange w:id="10" w:author="Alice"><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid></w:tblGridChange></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({ operations: [{action:'trackedChange', operation:'rejectAll'}] } as never);
    expect(result.success).toBe(true);
    expect(result.revisionsPending).toBe(0);
    const grid = parser.getDocumentDom().getElementsByTagName('w:tblGrid')[0];
    expect(grid.getElementsByTagName('w:gridCol')[0].getAttribute('w:w')).toBe('4000');
  });

  it('counts, accepts, saves and reopens a tracked insertion in a header', async () => {
    const zip = new AdmZip(documentWithRevisions('<w:p><w:r><w:t>Body</w:t></w:r></w:p>'));
    const header = '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:ins w:id="77" w:author="Alice"><w:r><w:t>Header text</w:t></w:r></w:ins></w:p></w:hdr>';
    zip.addFile('word/header1.xml', Buffer.from(header));
    const types = zip.readAsText('[Content_Types].xml').replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
    zip.updateFile('[Content_Types].xml', Buffer.from(types));
    const rels = zip.readAsText('word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rIdRevisionHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>');
    zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels));
    const document = zip.readAsText('word/document.xml').replace('<w:sectPr>', '<w:sectPr><w:headerReference w:type="default" r:id="rIdRevisionHeader"/>');
    zip.updateFile('word/document.xml', Buffer.from(document));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    expect(engine.countPendingRevisions()).toBe(1);
    const result = await engine.executeBatch({ operations: [{action:'trackedChange', operation:'acceptAll'}] } as never);
    expect(result.revisionsPending).toBe(0);
    const saved = new AdmZip(parser.rezip(ast));
    expect(saved.readAsText('word/header1.xml')).not.toContain('<w:ins');
    expect(saved.readAsText('word/header1.xml')).toContain('Header text');
  });

  it('accepts inserted-cell markers and retains the cell content', async () => {
    const body = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:cellIns w:id="10" w:author="Alice"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({ operations: [{action:'trackedChange', operation:'acceptAll'}] } as never);
    expect(result.revisionsPending).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:cellIns').length).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:tc')[0].textContent).toBe('Cell');
  });

  it('preserves an incomplete nested move and its containing insertion', async () => {
    const body = '<w:p><w:ins w:id="1" w:author="Alice"><w:moveFrom w:id="2" w:author="Bob"><w:r><w:t>Nested</w:t></w:r></w:moveFrom></w:ins></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'rejectAll'}]} as never);
    expect(result.revisionsPending).toBe(2);
    expect(result.revisionResolution?.unresolved).toHaveLength(2);
    expect(parser.getDocumentDom().getElementsByTagName('w:moveFrom').length).toBe(1);
  });

  it('direct parser callers use the same property revision resolver', async () => {
    const parser = new DocxParser(documentWithRevisions());
    const result = parser.applyTrackedChanges('reject');
    expect(result?.resolved).toBe(1);
    expect(parser.getDocumentDom().getElementsByTagName('w:rPrChange').length).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:i').length).toBe(1);
  });

  it('reports an unsupported table property exception revision', async () => {
    const body = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tblPrEx><w:tPrChange w:id="7" w:author="Alice"><w:tblPrEx/></w:tPrChange></w:tblPrEx><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'rejectAll'}]} as never);
    expect(result.revisionsPending).toBe(1);
    expect(result.revisionResolution?.unresolved).toEqual(expect.arrayContaining([expect.objectContaining({kind:'tPrChange'})]));
  });

  it('restores tracked table-property exceptions', async () => {
    const body = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tblPrEx><w:tblW w:w="3000" w:type="dxa"/><w:tblPrExChange w:id="7" w:author="Alice"><w:tblPrEx><w:tblW w:w="2000" w:type="dxa"/></w:tblPrEx></w:tblPrExChange></w:tblPrEx><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'rejectAll'}]} as never);
    expect(result.revisionsPending).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:tblPrExChange').length).toBe(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:tblPrEx')[0].getElementsByTagName('w:tblW')[0].getAttribute('w:w')).toBe('2000');
  });

  it('reports custom XML revision ranges instead of silently hiding them', async () => {
    const body = '<w:p><w:customXmlInsRangeStart w:id="19" w:author="Alice"/><w:r><w:t>Value</w:t></w:r><w:customXmlInsRangeEnd w:id="19"/></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'acceptAll'}]} as never);
    expect(result.revisionsPending).toBe(2);
    expect(result.revisionResolution?.unresolved).toHaveLength(2);
    expect(parser.getDocumentDom().getElementsByTagName('w:customXmlInsRangeStart').length).toBe(1);
  });

  it('leaves every part of a move group intact when nested revisions are unsupported', async () => {
    const body = '<w:p><w:moveFromRangeStart w:id="1" w:name="moveA"/><w:moveFrom w:id="2" w:author="Alice"><w:ins w:id="5" w:author="Bob"><w:r><w:t>old</w:t></w:r></w:ins></w:moveFrom><w:moveFromRangeEnd w:id="1"/><w:moveToRangeStart w:id="3" w:name="moveA"/><w:moveTo w:id="4" w:author="Alice"><w:r><w:t>new</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'acceptAll'}]} as never);
    expect(result.revisionsPending).toBe(3);
    expect(parser.getDocumentDom().getElementsByTagName('w:ins').length).toBe(1);
    expect(parser.getDocumentDom().getElementsByTagName('w:moveFrom').length).toBe(1);
    expect(parser.getDocumentDom().getElementsByTagName('w:moveTo').length).toBe(1);
  });

  it.each([
    ['footer', 'footer1.xml', 'ftr', 'footer', 'footer'],
    ['footnote', 'footnotes.xml', 'footnotes', 'footnotes', 'footnotes'],
    ['endnote', 'endnotes.xml', 'endnotes', 'endnotes', 'endnotes'],
  ])('resolves revisions in %s parts after save and reopen', async (label, name, root, type, relType) => {
    const reference = label === 'footer'
      ? '<w:headerReference w:type="default" r:id="rIdUnused"/>'
      : `<w:r><w:${label}Reference w:id="2"/></w:r>`;
    const body = label === 'footer' ? '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' : `<w:p>${reference}</w:p>`;
    const zip = new AdmZip(documentWithRevisions(body));
    const revision = '<w:ins w:id="88" w:author="Alice"><w:r><w:t>Part text</w:t></w:r></w:ins>';
    const contents = label === 'footer' ? `<w:p>${revision}</w:p>` : `<w:${label} w:id="2"><w:p>${revision}</w:p></w:${label}>`;
    zip.addFile(`word/${name}`, Buffer.from(`<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${contents}</w:${root}>`));
    zip.updateFile('[Content_Types].xml', Buffer.from(zip.readAsText('[Content_Types].xml').replace('</Types>', `<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/></Types>`)));
    zip.updateFile('word/_rels/document.xml.rels', Buffer.from(zip.readAsText('word/_rels/document.xml.rels').replace('</Relationships>', `<Relationship Id="rIdRevisionPart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${relType}" Target="${name}"/></Relationships>`)));
    if (label === 'footer') zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace('<w:sectPr>', '<w:sectPr><w:footerReference w:type="default" r:id="rIdRevisionPart"/>')));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    expect(engine.countPendingRevisions()).toBe(1);
    const result = await engine.executeBatch({operations:[{action:'trackedChange',operation:'acceptAll'}]} as never);
    expect(result.revisionsPending).toBe(0);
    const saved = parser.rezip(ast);
    expect(new AdmZip(saved).readAsText(`word/${name}`)).toContain('Part text');
    expect(new AdmZip(saved).readAsText(`word/${name}`)).not.toContain('<w:ins');
    const reopened = new DocxParser(saved);
    const reopenedAst = await reopened.getAST();
    expect(new OperationEngine(reopenedAst, reopened).countPendingRevisions()).toBe(0);
  });

  it('does not erase current properties when previous revision properties are missing', async () => {
    const body = '<w:p><w:pPr><w:jc w:val="right"/><w:pPrChange w:id="22" w:author="Alice"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'trackedChange',operation:'rejectAll'}]} as never);
    expect(result.revisionsPending).toBe(1);
    expect(parser.getDocumentDom().getElementsByTagName('w:jc')[0].getAttribute('w:val')).toBe('right');
    expect(parser.getDocumentDom().getElementsByTagName('w:pPrChange').length).toBe(1);
  });

  it('merges paragraph marks with bookmarks (such as _Toc) on accept', async () => {
    const body = '<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Alice"/></w:rPr></w:pPr><w:bookmarkStart w:id="10" w:name="_TocHeading"/><w:r><w:t>Heading</w:t></w:r><w:bookmarkEnd w:id="10"/></w:p><w:p><w:r><w:t>Body text</w:t></w:r></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
    expect(result.success).toBe(true);
    expect(result.revisionsPending).toBe(0);
    const paragraphs = parser.getDocumentDom().getElementsByTagName('w:p');
    expect(paragraphs.length).toBe(1);
    expect(paragraphs[0].textContent).toBe('HeadingBody text');
    expect(paragraphs[0].getElementsByTagName('w:bookmarkStart').length).toBe(1);
  });

  it('merges revisionResolution across multiple operations in a batch', async () => {
    const body = '<w:p><w:ins w:id="10" w:author="Alice"><w:r><w:t>one</w:t></w:r></w:ins><w:ins w:id="20" w:author="Bob"><w:r><w:t>two</w:t></w:r></w:ins></w:p>';
    const parser = new DocxParser(documentWithRevisions(body));
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({
      operations: [
        { action: 'trackedChange', operation: 'accept', ids: [10] },
        { action: 'trackedChange', operation: 'reject', ids: [20] },
      ]
    } as never);
    expect(result.success).toBe(true);
    expect(result.revisionResolution).toBeDefined();
    expect(result.revisionResolution?.resolved).toBe(2);
    expect(result.revisionResolution?.resolvedIds).toEqual(expect.arrayContaining([10, 20]));
  });

  it('sanitizer acceptAllRevisions scrubs author and date from unhandled revision markup', async () => {
    const body = '<w:p><w:customXmlInsRangeStart w:id="5" w:author="SecretAuthor" w:date="2026-01-01T00:00:00Z"/><w:r><w:t>Text</w:t></w:r><w:customXmlInsRangeEnd w:id="5"/></w:p>';
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const xml = zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const { buffer } = sanitizeDocxBuffer(zip.toBuffer(), { acceptAllRevisions: true });
    const cleanZip = new AdmZip(buffer);
    const cleanDoc = cleanZip.readAsText('word/document.xml');
    expect(cleanDoc).not.toContain('SecretAuthor');
    expect(cleanDoc).not.toContain('2026-01-01T00:00:00Z');
  });

});
