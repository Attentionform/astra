/**
 * What the OOXML validator (Open XML SDK) reports, pinned so a change in its
 * behaviour is noticed:
 *
 *  - `w:tblHeader` takes no value, `on` or `off`. `true`, `1`, `false` and `0` are
 *    reported, although they are legal for most other on/off elements. ASTra
 *    writes the bare element, which passes.
 *  - Children out of schema order inside property containers (`w:pPr`, `w:tblPr`)
 *    are reported on the raw bytes. ASTra puts children in schema order whenever
 *    it saves, and `DocxParser.validateOOXML` (so `astra validate` and
 *    `validate_document`) judges that repaired form; the repairs it would make
 *    are listed by `getPendingNormalizations`.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { validateBuffer } from '@attentionform/ooxml-validator';
import { DocxParser, composeDocument } from '@attentionform/astra-engine';

async function composedTable(): Promise<Buffer> {
  const parser = DocxParser.createNewDocument('blank');
  const result = await composeDocument(parser, '| A | B |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |\n\nBody text.\n');
  expect(result.success).toBe(true);
  return parser.rezip();
}

function editDocumentXml(buf: Buffer, edit: (xml: string) => string): Buffer {
  const zip = new AdmZip(buf);
  const before = zip.readAsText('word/document.xml');
  const after = edit(before);
  expect(after).not.toBe(before);
  zip.updateFile('word/document.xml', Buffer.from(after, 'utf8'));
  return zip.toBuffer();
}

const BODY_PARAGRAPH = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*Body text\.(?:(?!<\/w:p>)[\s\S])*<\/w:p>/;
const paragraphWith = (pPr: string): string => `<w:p>${pPr}<w:r><w:t>Body text.</w:t></w:r></w:p>`;
const BORDER = '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>';
const SPACING = '<w:spacing w:before="0" w:after="120"/>';
const INDENT = '<w:ind w:left="720"/>';
const SHADING = '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>';

/** Child element names of the first <w:pPr> that holds `marker`, nested property blocks collapsed. */
function pPrChildren(xml: string, marker: string): string[] {
  const block = [...xml.matchAll(/<w:pPr>[\s\S]*?<\/w:pPr>/g)].map((m) => m[0]).find((b) => b.includes(marker));
  expect(block).toBeDefined();
  return [...(block as string).replace(/<w:pBdr>[\s\S]*?<\/w:pBdr>/g, '<w:pBdr/>').slice('<w:pPr>'.length).matchAll(/<w:(\w+)[ />]/g)].map((m) => m[1]);
}

const ORDER_CASES: Array<{ name: string; element: string; edit: (xml: string) => string }> = [
  { name: 'pBdr after spacing', element: 'pBdr', edit: (xml) => xml.replace(BODY_PARAGRAPH, paragraphWith(`<w:pPr>${SPACING}${BORDER}</w:pPr>`)) },
  { name: 'shd after ind', element: 'shd', edit: (xml) => xml.replace(BODY_PARAGRAPH, paragraphWith(`<w:pPr>${INDENT}${SHADING}</w:pPr>`)) },
  { name: 'tblCellMar after tblLook', element: 'tblCellMar', edit: (xml) => xml.replace('</w:tblBorders><w:tblCellMar>', '</w:tblBorders><w:tblLook w:val="04A0"/><w:tblCellMar>') },
];

describe('tblHeader value (B2)', () => {
  it('the bare element ASTra writes validates', async () => {
    const buf = await composedTable();
    expect(new AdmZip(buf).readAsText('word/document.xml')).toContain('<w:tblHeader/>');
    expect((await validateBuffer(buf)).ok).toBe(true);
  });

  it('accepts on and off, and reports true, 1, false and 0 with the offending element', async () => {
    const buf = await composedTable();
    const verdict = async (val: string) => validateBuffer(editDocumentXml(buf, (xml) => xml.replace('<w:tblHeader/>', `<w:tblHeader w:val="${val}"/>`)));
    for (const ok of ['on', 'off']) expect((await verdict(ok)).ok).toBe(true);
    for (const bad of ['true', '1', 'false', '0']) {
      const result = await verdict(bad);
      expect(result.ok).toBe(false);
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0].description).toContain(`invalid value '${bad}'`);
      expect(result.errors[0].xPath).toContain('w:tblHeader');
    }
  });
});

describe('child order inside property containers (B3)', () => {
  it('the validator reports each out-of-order child on the raw bytes, and passes the right order', async () => {
    const buf = await composedTable();
    for (const { name, element, edit } of ORDER_CASES) {
      const result = await validateBuffer(editDocumentXml(buf, edit));
      expect(result.ok, name).toBe(false);
      expect(result.errors, name).toHaveLength(1);
      expect(result.errors[0].id, name).toBe('Sch_UnexpectedElementContentExpectingComplex');
      expect(result.errors[0].description, name).toContain(`'http://schemas.openxmlformats.org/wordprocessingml/2006/main:${element}'`);
    }
    const inOrder = editDocumentXml(buf, (xml) => xml.replace(BODY_PARAGRAPH, paragraphWith(`<w:pPr>${BORDER}${SHADING}${SPACING}${INDENT}</w:pPr>`)));
    expect((await validateBuffer(inOrder)).ok).toBe(true);
  });

  it('ASTra judges the repaired form and lists the repair, so its verdict is ok', async () => {
    const buf = await composedTable();
    for (const { name, edit } of ORDER_CASES) {
      const misordered = editDocumentXml(buf, edit);
      expect((await validateBuffer(misordered)).ok, `${name}: raw bytes`).toBe(false);
      const parser = new DocxParser(misordered);
      await parser.getAST();
      expect(parser.getPendingNormalizations().join('\n'), name).toMatch(/Reordered w:(pPr|tblPr) into schema order/);
      expect((await parser.validateOOXML()).ok, `${name}: repaired form`).toBe(true);
    }
  });

  it('saving through ASTra puts the children back in schema order', async () => {
    const buf = await composedTable();
    const misordered = editDocumentXml(buf, (xml) => xml
      .replace(BODY_PARAGRAPH, paragraphWith(`<w:pPr>${INDENT}${SHADING}${SPACING}${BORDER}</w:pPr>`))
      .replace('</w:tblBorders><w:tblCellMar>', '</w:tblBorders><w:tblLook w:val="04A0"/><w:tblCellMar>'));
    const parser = new DocxParser(misordered);
    await parser.getAST();
    const saved = new AdmZip(parser.rezip()).readAsText('word/document.xml');
    expect(pPrChildren(saved, 'pBdr')).toEqual(['pBdr', 'shd', 'spacing', 'ind']);
    const tblPr = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(saved)![0]
      .replace(/<w:(tblBorders|tblCellMar)>[\s\S]*?<\/w:\1>/g, '<w:$1/>');
    expect([...tblPr.matchAll(/<w:(\w+)[ />]/g)].map((m) => m[1]).slice(1)).toEqual(['tblStyle', 'tblBorders', 'tblCellMar', 'tblLook']);
  });
});
