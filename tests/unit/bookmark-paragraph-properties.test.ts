import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml"';

/** A paragraph whose only w:pPr sits inside a text box, so it has no paragraph properties of its own. */
const TEXT_BOX_RUN =
  '<w:r><w:pict><v:shape id="box1" type="#_x0000_t202" style="width:100pt;height:30pt"><v:textbox><w:txbxContent>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>in the box</w:t></w:r></w:p>' +
  '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>';

async function docWithParagraph(paragraphXml: string): Promise<Buffer> {
  const blank = DocxParser.createNewDocument('blank');
  await blank.getAST();
  const zip = new AdmZip(blank.rezip());
  zip.updateFile(
    'word/document.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
        `${paragraphXml}<w:sectPr/></w:body></w:document>`
    )
  );
  return zip.toBuffer();
}

async function addBookmark(input: Buffer, extra: Record<string, unknown> = {}): Promise<{ out: Buffer; childNames: string[] }> {
  const parser = new DocxParser(input);
  const ast = await parser.getAST();
  const para = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'paragraph')!;
  const result = await new OperationEngine(ast, parser).executeBatch({
    operations: [{ action: 'bookmark', target: para.id, bookmarkName: 'astra_bm', ...extra } as never],
  });
  expect(result.errors ?? []).toEqual([]);
  expect(result.success).toBe(true);
  const out = parser.rezip();
  const doc = new DOMParser().parseFromString(new AdmZip(out).readAsText('word/document.xml'), 'text/xml');
  const body = doc.getElementsByTagName('w:body')[0];
  const outer = Array.from(body.childNodes).find((n) => n.nodeType === 1 && (n as Element).tagName === 'w:p') as Element;
  const childNames = Array.from(outer.childNodes)
    .filter((n) => n.nodeType === 1)
    .map((n) => (n as Element).tagName);
  return { out, childNames };
}

describe('bookmark goes by the paragraph\'s own properties, not by a w:pPr deeper inside it', () => {
  it('a paragraph with a text box and no w:pPr of its own gets the bookmark at its start and end', async () => {
    const input = await docWithParagraph(`<w:p>${TEXT_BOX_RUN}</w:p>`);
    expect((await validateDocxBuffer(input)).ok).toBe(true);
    const { out, childNames } = await addBookmark(input);
    expect(childNames).toEqual(['w:bookmarkStart', 'w:r', 'w:bookmarkEnd']);
    const res = await validateDocxBuffer(out);
    expect(res.errors).toEqual([]);
    expect(res.ok).toBe(true);
  });

  it('the name alias works on the same paragraph', async () => {
    const input = await docWithParagraph(`<w:p>${TEXT_BOX_RUN}</w:p>`);
    const { out } = await addBookmark(input, { bookmarkName: undefined, name: 'alias_bm' });
    expect(new AdmZip(out).readAsText('word/document.xml')).toContain('w:name="alias_bm"');
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('a paragraph with its own w:pPr still gets the bookmark right after it', async () => {
    const input = await docWithParagraph(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>plain</w:t></w:r></w:p>`);
    const { out, childNames } = await addBookmark(input);
    expect(childNames).toEqual(['w:pPr', 'w:bookmarkStart', 'w:r', 'w:bookmarkEnd']);
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('a paragraph with its own w:pPr and a text box puts the bookmark after its own w:pPr', async () => {
    const input = await docWithParagraph(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr>${TEXT_BOX_RUN}</w:p>`);
    const { out, childNames } = await addBookmark(input);
    expect(childNames).toEqual(['w:pPr', 'w:bookmarkStart', 'w:r', 'w:bookmarkEnd']);
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('a paragraph with only a w:pPr and nothing after it still works', async () => {
    const input = await docWithParagraph(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p>`);
    const { out, childNames } = await addBookmark(input);
    expect(childNames).toEqual(['w:pPr', 'w:bookmarkStart', 'w:bookmarkEnd']);
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });
});
