import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';
import { DOMParser, XMLSerializer } from '../../packages/engine/src/ooxml/dom-parser';
import { protectLineSeparators } from '../../packages/engine/src/ooxml/dom-parser';

/**
 * U+2028 (line separator) and U+0085 (next line) are ordinary characters in XML 1.0, which is what
 * Word writes. The parser library treats them as line ends and turns them into a line feed, so a
 * save changed the text of the document. They are written as character references before parsing.
 */
const LS = String.fromCharCode(0x2028);
const NEL = String.fromCharCode(0x85);
const TEXT = `before${LS}middle${NEL}after`;

async function blankWith(documentBodyXml: string): Promise<Buffer> {
  const blank = DocxParser.createNewDocument('blank');
  await blank.getAST();
  const zip = new AdmZip(blank.rezip());
  zip.updateFile(
    'word/document.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${documentBodyXml}<w:sectPr/></w:body></w:document>`
    )
  );
  return zip.toBuffer();
}

describe('protectLineSeparators', () => {
  it('leaves a string without the characters alone', () => {
    const xml = '<a>plain\r\ntext</a>';
    expect(protectLineSeparators(xml)).toBe(xml);
  });

  it('writes both characters as references in text and attribute values', () => {
    expect(protectLineSeparators(`<a b="x${LS}y">p${NEL}q</a>`)).toBe('<a b="x&#x2028;y">p&#x85;q</a>');
  });

  it('does not touch comments, CDATA sections and processing instructions, where a reference would be literal text', () => {
    const xml = `<a><!-- c${LS}c --><![CDATA[d${NEL}d]]><?pi e${LS}e?>t${LS}t</a>`;
    expect(protectLineSeparators(xml)).toBe(`<a><!-- c${LS}c --><![CDATA[d${NEL}d]]><?pi e${LS}e?>t&#x2028;t</a>`);
  });

  it('the parser wrapper keeps the characters in the text and writes them back unchanged', () => {
    const doc = new DOMParser().parseFromString(`<a>x${LS}y${NEL}z</a>`, 'text/xml');
    expect(doc.documentElement.textContent).toBe(`x${LS}y${NEL}z`);
    expect(new XMLSerializer().serializeToString(doc)).toContain(`x${LS}y${NEL}z`);
  });

  it('a line feed and a carriage return in the source behave as before', () => {
    const doc = new DOMParser().parseFromString('<a>x\r\ny\nz</a>', 'text/xml');
    expect(doc.documentElement.textContent).toBe('x\ny\nz');
  });
});

describe('a save keeps U+2028 and U+0085 in the text', () => {
  it('the body text of a document is the same after a save with no edit', async () => {
    const input = await blankWith(`<w:p><w:r><w:t>${TEXT}</w:t></w:r></w:p>`);
    const parser = new DocxParser(input);
    await parser.getAST();
    const out = parser.rezip();
    const body = new AdmZip(out).readAsText('word/document.xml');
    expect(body).toContain(TEXT);
    expect(body).not.toContain('before\nmiddle');
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('the text of a paragraph that is not the one edited is the same after an edit', async () => {
    const input = await blankWith(`<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>${TEXT}</w:t></w:r></w:p>`);
    const parser = new DocxParser(input);
    const ast = await parser.getAST();
    const first = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'paragraph')!;
    const result = await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'replace', target: first.id, text: 'changed' }],
    });
    expect(result.success).toBe(true);
    const body = new AdmZip(parser.rezip()).readAsText('word/document.xml');
    expect(body).toContain('changed');
    expect(body).toContain(TEXT);
  });

  it('the text of a header is the same after a save with no edit', async () => {
    const withHeader = DocxParser.createNewDocument('blank');
    const ast0 = await withHeader.getAST();
    await new OperationEngine(ast0, withHeader).executeBatch({
      operations: [{ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'HEADERTEXT' }],
    });
    const zip = new AdmZip(withHeader.rezip());
    zip.updateFile('word/header1.xml', Buffer.from(zip.readAsText('word/header1.xml').replace('HEADERTEXT', `H${TEXT}`)));
    const parser = new DocxParser(zip.toBuffer());
    await parser.getAST();
    const header = new AdmZip(parser.rezip()).readAsText('word/header1.xml');
    expect(header).toContain(`H${TEXT}`);
  });

  it('the text of an existing footnote is the same after another footnote is added', async () => {
    const start = DocxParser.createNewDocument('blank');
    const ast0 = await start.getAST();
    const anchor = ast0.sections[0].nodes[0].id;
    await new OperationEngine(ast0, start).executeBatch({
      operations: [{ action: 'footnote', anchor, text: 'FOOTNOTEONE' }],
    });
    const zip = new AdmZip(start.rezip());
    zip.updateFile('word/footnotes.xml', Buffer.from(zip.readAsText('word/footnotes.xml').replace('FOOTNOTEONE', `N${TEXT}`)));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    const result = await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'footnote', anchor: ast.sections[0].nodes[0].id, text: 'second' }],
    });
    expect(result.success).toBe(true);
    const footnotes = new AdmZip(parser.rezip()).readAsText('word/footnotes.xml');
    expect(footnotes).toContain(`N${TEXT}`);
    expect(footnotes).toContain('second');
  });
});
