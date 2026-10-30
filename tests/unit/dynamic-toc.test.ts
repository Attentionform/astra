import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  OperationEngine,
  TOCManipulator,
} from '@attentionform/astra-engine';

import AdmZip from 'adm-zip';

function createDocumentWithHeadings(): Buffer {
  const zip = new AdmZip();
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Executive Summary</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Introductory content paragraph.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Methodology &amp; Scope</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Detailed methodology.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Clinical Recommendations</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
  return zip.toBuffer();
}

describe('Dynamic Table of Contents (TOC) Engine Test Suite', () => {
  it('should scan headings and assign bookmarks', async () => {
    const parser = new DocxParser(createDocumentWithHeadings());
    const ast = await parser.getAST();

    const entries = TOCManipulator.scanHeadings(ast, 3);
    expect(entries.length).toBe(3);
    expect(entries[0].headingText).toBe('Executive Summary');
    expect(entries[0].level).toBe(1);
    expect(entries[1].headingText).toBe('Methodology & Scope');
    expect(entries[1].level).toBe(2);
    expect(entries[2].headingText).toBe('Clinical Recommendations');
    expect(entries[2].level).toBe(1);
  });

  it('should insert dynamic TOC with dot leader tabs and hyperlinks via OperationEngine', async () => {
    const parser = new DocxParser(createDocumentWithHeadings());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const res = await engine.executeBatch({
      operations: [
        {
          action: 'field',
          operation: 'insertTOC',
          title: 'Contents Overview',
          levels: 3,
          position: 'before',
          anchor: ast.sections[0].nodes[0].id
        }
      ]
    });

    expect(res.success).toBe(true);

    const xml = parser.readDocumentXml();
    expect(xml).toContain('Contents Overview');
    expect(xml).toContain('w:val="right" w:leader="dot" w:pos="9360"');
    expect(xml).toContain('w:hyperlink');
    expect(xml).toContain('Executive Summary');
  });

  it('should enable w:updateFields in word/settings.xml', async () => {
    const parser = new DocxParser(createDocumentWithHeadings());
    TOCManipulator.enableWordAutoUpdateFields(parser);

    const settingsXml = parser.getZipAdapter().readTextOrNull('word/settings.xml');
    expect(settingsXml).toBeTruthy();
    expect(settingsXml).toContain('<w:updateFields w:val="true"/>');
  });
});
