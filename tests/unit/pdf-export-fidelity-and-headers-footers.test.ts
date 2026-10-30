import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  convertAstToHtml,
  DocumentAST,
  DocxParser,
} from '@attentionform/astra-engine';

function createSampleDocx(options?: {
  headerText?: string;
  footerText?: string;
  paperSize?: 'A4' | 'Letter';
  orientation?: 'portrait' | 'landscape';
  addMedia?: boolean;
}): Buffer {
  const zip = new AdmZip();
  const isA4 = options?.paperSize === 'A4';
  const isLandscape = options?.orientation === 'landscape';

  const w = isLandscape ? (isA4 ? '16838' : '15840') : (isA4 ? '11906' : '12240');
  const h = isLandscape ? (isA4 ? '11906' : '12240') : (isA4 ? '16838' : '15840');
  const orientAttr = isLandscape ? ' w:orient="landscape"' : '';

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:color w:val="1E40AF"/>
          <w:sz w:val="36"/>
          <w:rFonts w:ascii="Helvetica"/>
        </w:rPr>
        <w:t>Styled Centered Heading</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="right"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:i/>
          <w:u w:val="single"/>
        </w:rPr>
        <w:t>Italic Underlined Right</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="both"/>
      </w:pPr>
      <w:r>
        <w:t>Justified text body paragraph.</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdH1"/>
      <w:footerReference w:type="default" r:id="rIdF1"/>
      <w:pgSz w:w="${w}" w:h="${h}"${orientAttr}/>
      <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));

  const headerText = options?.headerText ?? 'Confidential Document Header';
  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>${headerText}</w:t></w:r></w:p>
</w:hdr>`;
  zip.addFile('word/header1.xml', Buffer.from(headerXml, 'utf8'));

  const footerText = options?.footerText ?? 'Internal Use Only Footer';
  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>${footerText}</w:t></w:r></w:p>
</w:ftr>`;
  zip.addFile('word/footer1.xml', Buffer.from(footerXml, 'utf8'));

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;
  zip.addFile('word/_rels/document.xml.rels', Buffer.from(relsXml, 'utf8'));

  if (options?.addMedia) {
    const fakePng = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00]);
    zip.addFile('word/media/diagram.png', fakePng);
  }

  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));

  return zip.toBuffer();
}

describe('PDF Export Fidelity & Headers/Footers Pipeline', () => {
  it('DocxParser: extracts pageSetup, headers, footers, and mediaFiles into AST', async () => {
    const docxBuf = createSampleDocx({
      paperSize: 'A4',
      orientation: 'landscape',
      headerText: 'A4 Landscape Header',
      footerText: 'A4 Landscape Footer',
      addMedia: true
    });

    const parser = new DocxParser(docxBuf);
    const ast = await parser.getAST();

    // 1. PageSetup verification
    expect(ast.sections[0].pageSetup).toBeDefined();
    expect(ast.sections[0].pageSetup?.paperSize).toBe('A4');
    expect(ast.sections[0].pageSetup?.orientation).toBe('landscape');
    expect(ast.sections[0].pageSetup?.marginTop).toBe(1440);

    // 2. Headers & Footers in AST
    expect(ast.headers['rIdH1'] || ast.headers['header1.xml']).toBeDefined();
    expect(ast.footers['rIdF1'] || ast.footers['footer1.xml']).toBeDefined();
    expect(ast.metadata.headerText).toBe('A4 Landscape Header');
    expect(ast.metadata.footerText).toBe('A4 Landscape Footer');

    // 3. Media files in AST
    expect(ast.mediaFiles['diagram.png']).toBeDefined();
    expect(ast.mediaFiles['word/media/diagram.png']).toBeDefined();
  });

  it('DocxParser: getStructuralSummary populates mediaFiles lazily with base64 data URLs', () => {
    const docxBuf = createSampleDocx({ addMedia: true });
    const parser = new DocxParser(docxBuf);

    const summary1 = parser.getStructuralSummary();
    expect(summary1.mediaFiles).toBeDefined();
    expect(summary1.mediaFiles!['diagram.png']).toBeDefined();
    expect(summary1.mediaFiles!['diagram.png'].mimeType).toBe('image/png');
    expect(summary1.mediaFiles!['diagram.png'].base64).toMatch(/^data:image\/png;base64,/);

    // Test invalidateMediaCache
    parser.invalidateMediaCache();
    const summary2 = parser.getStructuralSummary();
    expect(summary2.mediaFiles!['diagram.png']).toBeDefined();
  });

  it('convertAstToHtml: preserves paragraph alignment (text-align)', async () => {
    const docxBuf = createSampleDocx();
    const parser = new DocxParser(docxBuf);
    const ast = await parser.getAST();

    const html = convertAstToHtml(ast, 'Alignment Test');
    expect(html).toContain('style="text-align: center;"');
    expect(html).toContain('style="text-align: right;"');
    expect(html).toContain('style="text-align: justify;"');
  });

  it('convertAstToHtml: preserves run formatting with span styles', async () => {
    const docxBuf = createSampleDocx();
    const parser = new DocxParser(docxBuf);
    const ast = await parser.getAST();

    const html = convertAstToHtml(ast, 'Run Styles Test');
    // Bold, color, size (36/2 = 18pt), font-family
    expect(html).toContain('font-weight: bold;');
    expect(html).toContain('color: #1E40AF;');
    expect(html).toContain('font-size: 18pt;');
    expect(html).toContain("font-family: 'Helvetica', sans-serif;");

    // Italic, underline
    expect(html).toContain('font-style: italic;');
    expect(html).toContain('text-decoration: underline;');
  });

  it('convertAstToHtml: heading styles inherit color and border-bottom is removed from h1', () => {
    const ast: DocumentAST = {
      sections: [{
        id: 's1',
        heading: null,
        nodes: [{
          id: 'p1',
          type: 'paragraph',
          xml: null as any,
          metadata: { text: 'Title Text', namedStyle: 'Heading 1', hasCitation: false, hasComment: false, hasBookmark: false, hasTrackedChanges: false, hasEquation: false, hasField: false, hasImage: false, hasHyperlink: false },
          version: 1,
          checksum: 'chk1'
        }],
        pageSetup: null,
        version: 1,
        checksum: 'chk_sec'
      }],
      metadata: {},
      styles: {},
      numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {},
      headers: {},
      footers: {},
      mediaFiles: {},
      settings: { trackChanges: false, defaultAuthor: 'ASTra' },
      nextChangeId: 1
    };

    const html = convertAstToHtml(ast, 'CSS Test', { printMode: true });
    // h1 styling
    expect(html).toContain('h1 { font-size: 24px; color: inherit;');
    expect(html).not.toContain('h1 { font-size: 26px; color: #0F172A; border-bottom: 1px solid #E2E8F0;');
    // No unsupported @top-center or @bottom-center margin boxes in @page
    expect(html).not.toContain('@top-center');
    expect(html).not.toContain('@bottom-center');
    expect(html).toContain('@page {');
  });

  it('DocxParser: parses Letter paper size correctly', async () => {
    const docxBuf = createSampleDocx({ paperSize: 'Letter', orientation: 'portrait' });
    const parser = new DocxParser(docxBuf);
    const ast = await parser.getAST();

    expect(ast.sections[0].pageSetup?.paperSize).toBe('Letter');
    expect(ast.sections[0].pageSetup?.orientation).toBe('portrait');
  });
});
