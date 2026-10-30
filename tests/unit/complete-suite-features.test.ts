import { describe, it, expect } from 'vitest';
import {
  convertAstToHtml,
  convertAstToMarkdown,
  CoverPageOperation,
  DocxParser,
  FootnoteOperation,
  HyperlinkOperation,
  OperationEngine,
  PageBorderOperation,
  WatermarkOperation,
} from '@attentionform/astra-engine';

import { XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';

function createMinimalDocx(): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
  );
  zip.addFile(
    'word/document.xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p data-astra-id="p_001" w14:paraId="11111111">
      <w:r><w:t>Main research methodology section</w:t></w:r>
    </w:p>
    <w:p data-astra-id="p_002" w14:paraId="22222222">
      <w:r><w:t>Second paragraph with data</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Complete Suite Pillars Suite', () => {
  it('inserts footnote and synchronizes word/footnotes.xml OPC part', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const fnOp: FootnoteOperation = {
      action: 'footnote',
      operation: 'insert',
      anchor: '11111111',
      text: 'Johnson et al., Science 2026, Vol 384, pp. 112-120.'
    };

    const result = await engine.executeBatch({ operations: [fnOp] });
    expect(result.success).toBe(true);

    const pNode = ast.sections[0].nodes.find(n => n.id === '11111111');
    const xmlStr = new XMLSerializer().serializeToString(pNode!.xml);
    expect(xmlStr).toContain('w:footnoteReference');
    expect(xmlStr).toContain('w:vertAlign w:val="superscript"');

    // Check footnotes.xml in zipAdapter
    const zipAdapter = parser.getZipAdapter();
    expect(zipAdapter.hasFile('word/footnotes.xml')).toBe(true);
    const fnXml = zipAdapter.readText('word/footnotes.xml');
    expect(fnXml).toContain('Johnson et al., Science 2026');
  });

  it('sets document watermark in header layer', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const wmOp: WatermarkOperation = {
      action: 'watermark',
      text: 'CONFIDENTIAL DRAFT',
      color: 'CBD5E1',
      fontSize: 72
    };

    const result = await engine.executeBatch({ operations: [wmOp] });
    expect(result.success).toBe(true);

    const zipAdapter = parser.getZipAdapter();
    expect(zipAdapter.hasFile('word/header1.xml')).toBe(true);
    const hdrXml = zipAdapter.readText('word/header1.xml');
    expect(hdrXml).toContain('CONFIDENTIAL DRAFT');
    expect(hdrXml).toContain('PowerPlusWaterMarkObject');
  });

  it('applies page border in sectPr', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const borderOp: PageBorderOperation = {
      action: 'pageBorder',
      style: 'double',
      sz: 18,
      color: '4F46E5'
    };

    const result = await engine.executeBatch({ operations: [borderOp] });
    expect(result.success).toBe(true);

    const doc = ast.sections[0].nodes[0].xml.ownerDocument;
    const docXml = new XMLSerializer().serializeToString(doc!);
    expect(docXml).toContain('w:pgBorders');
    expect(docXml).toContain('w:val="double"');
    expect(docXml).toContain('w:sz="18"');
  });

  it('inserts executive cover page and enables titlePg', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const coverOp: CoverPageOperation = {
      action: 'coverPage',
      template: 'modern_minimalist',
      title: '2026 AI Roadmap',
      subtitle: 'Confidential Executive Brief',
      author: 'ASTra Core Team'
    };

    const result = await engine.executeBatch({ operations: [coverOp] });
    expect(result.success).toBe(true);

    const doc = ast.sections[0].nodes[0].xml.ownerDocument;
    const docXml = new XMLSerializer().serializeToString(doc!);
    expect(docXml).toContain('w:titlePg');
    expect(docXml).toContain('2026 AI Roadmap');
    expect(docXml).toContain('Confidential Executive Brief');
  });

  it('inserts external hyperlink and bookmark anchors', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const linkOp: HyperlinkOperation = {
      action: 'hyperlink',
      target: '11111111',
      text: 'ASTra Portal',
      url: 'https://astra.dev'
    };

    const bmOp: HyperlinkOperation = {
      action: 'bookmark',
      target: '22222222',
      bookmarkName: 'MethodologySection'
    };

    const result = await engine.executeBatch({ operations: [linkOp, bmOp] });
    expect(result.success).toBe(true);

    const p1 = ast.sections[0].nodes.find(n => n.id === '11111111');
    const p1Xml = new XMLSerializer().serializeToString(p1!.xml);
    expect(p1Xml).toContain('w:hyperlink');
    expect(p1Xml).toContain('ASTra Portal');

    const p2 = ast.sections[0].nodes.find(n => n.id === '22222222');
    const p2Xml = new XMLSerializer().serializeToString(p2!.xml);
    expect(p2Xml).toContain('w:bookmarkStart');
    expect(p2Xml).toContain('MethodologySection');
  });

  it('exports AST to GFM Markdown and standalone HTML', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();

    const md = convertAstToMarkdown(ast);
    expect(md).toContain('Main research methodology section');
    expect(md).toContain('Second paragraph with data');

    const html = convertAstToHtml(ast, 'Research Paper');
    expect(html).toContain('<!DOCTYPE html>');
    expect(html).toContain('Main research methodology section');
    expect(html).toContain('<title>Research Paper</title>');
  });
});
