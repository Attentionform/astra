import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import {
  createBlankDocxBuffer,
  DocxParser,
  ensureTableGrid,
  OperationEngine,
  PPR_CHILD_ORDER,
  RPR_CHILD_ORDER,
  sortChildrenBySchema,
} from '@attentionform/astra-engine';

describe('OOXML Schema Ordering & Canonicalization Suite', () => {
  it('should sort w:rPr child elements into strict ECMA-376 schema order', () => {
    const xml = `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:color w:val="FF0000"/>
      <w:sz w:val="32"/>
      <w:rFonts w:ascii="Calibri"/>
      <w:b/>
      <w:highlight w:val="yellow"/>
    </w:rPr>`;

    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const rPr = doc.documentElement;
    sortChildrenBySchema(rPr, RPR_CHILD_ORDER);

    const childNames = Array.from(rPr.childNodes)
      .filter((n: any) => n.nodeType === 1)
      .map((n: any) => n.tagName);

    // ECMA-376 sequence: rFonts -> b -> color -> sz -> highlight
    expect(childNames).toEqual(['w:rFonts', 'w:b', 'w:color', 'w:sz', 'w:highlight']);
  });

  it('should keep w:suppressLineNumbers ahead of spacing siblings in w:pPr', () => {
    const xml = `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:spacing w:after="160"/>
      <w:suppressLineNumbers/>
      <w:jc w:val="both"/>
    </w:pPr>`;

    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const pPr = doc.documentElement;
    sortChildrenBySchema(pPr, PPR_CHILD_ORDER);

    const childNames = Array.from(pPr.childNodes)
      .filter((n: any) => n.nodeType === 1)
      .map((n: any) => n.tagName);

    // ECMA-376 CT_PPr sequence: suppressLineNumbers -> spacing -> jc
    expect(childNames).toEqual(['w:suppressLineNumbers', 'w:spacing', 'w:jc']);
  });

  it('should automatically inject <w:tblGrid> into tables missing grid definition', () => {
    const xml = `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Col 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Col 2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Col 3</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;

    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const tbl = doc.documentElement;
    ensureTableGrid(tbl, doc);

    const tblGrid = tbl.getElementsByTagName('w:tblGrid')[0];
    expect(tblGrid).toBeDefined();

    const cols = tblGrid.getElementsByTagName('w:gridCol');
    expect(cols.length).toBe(3);
    expect(cols[0].getAttribute('w:w')).toBe('3120'); // 9360 / 3
  });

  it('should canonicalize entire document DOM on rezip()', async () => {
    const blankBuf = createBlankDocxBuffer('novel');
    const parser = new DocxParser(blankBuf);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    // Style first paragraph out of order: color, then bold, then font family
    const targetNode = ast.sections[0].nodes[0];
    await engine.executeBatch({
      operations: [
        {
          action: 'style',
          target: targetNode.id,
          properties: {
            color: '003366',
            fontSize: 28,
            bold: true,
            fontFamily: 'Georgia'
          }
        }
      ]
    });

    const validation = await parser.validateOOXML();
    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
  });
});
