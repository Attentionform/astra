import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';
import { RPR_CHILD_ORDER, sortChildrenBySchema } from '../../packages/engine/src/ooxml/schema-order';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"';

/** Word writes these after every w: child of w:rPr, in this order (MS-DOCX 2.6.1). */
const W14_IN_WORD_ORDER = [
  'w14:glow',
  'w14:shadow',
  'w14:reflection',
  'w14:textOutline',
  'w14:textFill',
  'w14:scene3d',
  'w14:props3d',
  'w14:ligatures',
  'w14:numForm',
  'w14:numSpacing',
  'w14:stylisticSets',
  'w14:cntxtAlts',
];

async function docWithRunProps(rPrInner: string): Promise<Buffer> {
  const blank = DocxParser.createNewDocument('blank');
  await blank.getAST();
  const zip = new AdmZip(blank.rezip());
  zip.updateFile(
    'word/document.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
        `<w:p><w:r><w:rPr>${rPrInner}</w:rPr><w:t>text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return zip.toBuffer();
}

async function boldFirstParagraph(input: Buffer): Promise<{ out: Buffer; rPrChildren: string[] }> {
  const parser = new DocxParser(input);
  const ast = await parser.getAST();
  const para = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'paragraph')!;
  const result = await new OperationEngine(ast, parser).executeBatch({
    operations: [{ action: 'style', target: para.id, bold: true }],
  });
  expect(result.success).toBe(true);
  const out = parser.rezip();
  const xml = new AdmZip(out).readAsText('word/document.xml');
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const rPr = doc.getElementsByTagName('w:r')[0].getElementsByTagName('w:rPr')[0];
  const rPrChildren = Array.from(rPr.childNodes)
    .filter((n) => n.nodeType === 1)
    .map((n) => (n as Element).tagName);
  return { out, rPrChildren };
}

describe('run properties keep the w14 elements after every w: element', () => {
  it('lists the w14 run properties after w:oMath and before w:rPrChange, in the order Word writes them', () => {
    const w14 = RPR_CHILD_ORDER.filter((t) => t.startsWith('w14:'));
    expect(w14).toEqual(W14_IN_WORD_ORDER);
    const firstW14 = RPR_CHILD_ORDER.indexOf('w14:glow');
    expect(firstW14).toBeGreaterThan(RPR_CHILD_ORDER.indexOf('w:oMath'));
    expect(RPR_CHILD_ORDER.indexOf('w14:cntxtAlts')).toBeLessThan(RPR_CHILD_ORDER.indexOf('w:rPrChange'));
    expect(RPR_CHILD_ORDER[RPR_CHILD_ORDER.length - 1]).toBe('w:rPrChange');
  });

  it('sortChildrenBySchema puts w:lang before w14:ligatures even when the w14 element sits in the middle', () => {
    const doc = new DOMParser().parseFromString(
      `<w:rPr ${NS}><w:rFonts/><w14:ligatures w14:val="standard"/><w:lang w:val="en-US"/><w:b/></w:rPr>`,
      'text/xml'
    );
    const rPr = doc.documentElement as unknown as Element;
    sortChildrenBySchema(rPr, RPR_CHILD_ORDER);
    const names = Array.from(rPr.childNodes).map((n) => (n as Element).tagName);
    expect(names).toEqual(['w:rFonts', 'w:b', 'w:lang', 'w14:ligatures']);
    expect(new XMLSerializer().serializeToString(doc)).toContain('w14:val="standard"');
  });

  it('bold on a run with w:lang followed by w14:ligatures stays valid (the form of 9 corpus documents)', async () => {
    const input = await docWithRunProps(
      '<w:rFonts w:ascii="Calibri"/><w:color w:val="000000"/><w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/>' +
        '<w:lang w:val="en-US"/><w14:ligatures w14:val="standardContextual"/>'
    );
    expect((await validateDocxBuffer(input)).ok).toBe(true);
    const { out, rPrChildren } = await boldFirstParagraph(input);
    expect(rPrChildren).toEqual(['w:rFonts', 'w:b', 'w:color', 'w:kern', 'w:sz', 'w:szCs', 'w:lang', 'w14:ligatures']);
    const res = await validateDocxBuffer(out);
    expect(res.errors).toEqual([]);
    expect(res.ok).toBe(true);
  });

  it('bold on a run with w14:textOutline after w:sz keeps w:szCs before the w14 element', async () => {
    const input = await docWithRunProps(
      '<w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/>' +
        '<w14:textOutline w14:w="9525"><w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'
    );
    expect((await validateDocxBuffer(input)).ok).toBe(true);
    const { out, rPrChildren } = await boldFirstParagraph(input);
    expect(rPrChildren).toEqual(['w:rFonts', 'w:b', 'w:sz', 'w:szCs', 'w14:textOutline']);
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('clear formatting removes the w14 text effects together with the w: run properties, as Word does', async () => {
    const input = await docWithRunProps('<w:b/><w:lang w:val="en-US"/><w14:ligatures w14:val="standard"/>');
    const parser = new DocxParser(input);
    const ast = await parser.getAST();
    const para = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'paragraph')!;
    const result = await new OperationEngine(ast, parser).executeBatch({ operations: [{ action: 'clearFormat', target: para.id }] });
    expect(result.success).toBe(true);
    const xml = new AdmZip(parser.rezip()).readAsText('word/document.xml');
    expect(xml).not.toContain('<w:b/>');
    expect(xml).not.toContain('w:lang');
    expect(xml).not.toContain('w14:ligatures');
  });

  it('a full set of w14 properties in Word order is left as it is', async () => {
    const w14xml =
      '<w14:glow w14:rad="63500"><w14:srgbClr w14:val="FF0000"/></w14:glow>' +
      '<w14:shadow w14:blurRad="0" w14:dist="0" w14:dir="0" w14:sx="0" w14:sy="0" w14:kx="0" w14:ky="0" w14:algn="b"><w14:srgbClr w14:val="000000"/></w14:shadow>' +
      '<w14:textOutline w14:w="9525"><w14:noFill/></w14:textOutline><w14:textFill><w14:noFill/></w14:textFill>' +
      '<w14:ligatures w14:val="standard"/><w14:numForm w14:val="lining"/><w14:numSpacing w14:val="proportional"/>' +
      '<w14:stylisticSets><w14:ss w14:id="1"/></w14:stylisticSets><w14:cntxtAlts/>';
    const input = await docWithRunProps(`<w:lang w:val="en-US"/>${w14xml}`);
    expect((await validateDocxBuffer(input)).ok).toBe(true);
    const { out, rPrChildren } = await boldFirstParagraph(input);
    expect(rPrChildren).toEqual(['w:b', 'w:lang', ...W14_IN_WORD_ORDER.filter((t) => t !== 'w14:reflection' && t !== 'w14:scene3d' && t !== 'w14:props3d')]);
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });
});
