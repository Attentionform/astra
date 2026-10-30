import { describe, it, expect, beforeEach } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocumentAST,
  HeaderFooterManipulator,
  DocxParser,
  createBlankDocxBuffer,
  assertValidDocx,
} from '@attentionform/astra-engine';

describe('HeaderFooterManipulator Unit Suite', () => {
  let ast: DocumentAST;
  let domParser: DOMParser;

  beforeEach(() => {
    domParser = new DOMParser();
    const mockDoc = domParser.parseFromString(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p data-astra-id="p_001">
            <w:r><w:t>Body text</w:t></w:r>
          </w:p>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
          </w:sectPr>
        </w:body>
      </w:document>`,
      'text/xml'
    );

    const bodyP = mockDoc.getElementsByTagName('w:p')[0];

    ast = {
      metadata: {},
      sections: [{
        id: 'sec_root',
        heading: null,
        nodes: [{
          id: 'p_001',
          type: 'paragraph',
          xml: bodyP,
          metadata: {
            text: 'Body text',
            hasCitation: false,
            hasComment: false,
            hasBookmark: false,
            hasTrackedChanges: false,
            hasEquation: false,
            hasField: false,
            hasImage: false,
            hasHyperlink: false
          },
          version: 1,
          checksum: 'abc'
        }],
        pageSetup: null,
        version: 1,
        checksum: '123'
      }],
      styles: {},
      numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {},
      headers: {},
      footers: {},
      mediaFiles: {},
      settings: { trackChanges: true, defaultAuthor: 'ASTra' },
      nextChangeId: 1
    };
  });

  it('should set default header and update AST relationships and section references', () => {
    const manipulator = new HeaderFooterManipulator(ast);
    const result = manipulator.setHeaderFooter({
      targetType: 'header',
      pageType: 'default',
      text: 'ASTra Header Confidential'
    });

    expect(result.rId).toBeTruthy();
    expect(ast.relationships[result.rId]).toBeDefined();
    expect(ast.relationships[result.rId].type).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/header');
    expect(ast.headers[result.rId]).toBeDefined();

    const node = ast.sections[0].nodes[0];
    const sectPr = (node.xml.ownerDocument || node.xml).getElementsByTagName('w:sectPr')[0];
    expect(sectPr).toBeDefined();

    const headerRef = sectPr.getElementsByTagName('w:headerReference')[0];
    expect(headerRef).toBeDefined();
    expect(headerRef.getAttribute('r:id')).toBe(result.rId);
    expect(headerRef.getAttribute('w:type')).toBe('default');
  });

  it('should set first page header and include w:titlePg in w:sectPr', () => {
    const manipulator = new HeaderFooterManipulator(ast);
    const result = manipulator.setHeaderFooter({
      targetType: 'header',
      pageType: 'first',
      text: 'Cover Page Header'
    });

    expect(result.rId).toBeTruthy();
    const node = ast.sections[0].nodes[0];
    const sectPr = (node.xml.ownerDocument || node.xml).getElementsByTagName('w:sectPr')[0];

    const firstHeaderRef = sectPr.getElementsByTagName('w:headerReference')[0];
    expect(firstHeaderRef.getAttribute('w:type')).toBe('first');
    expect(sectPr.getElementsByTagName('w:titlePg').length).toBe(1);
  });

  it('should remove header reference and clean AST maps', () => {
    const manipulator = new HeaderFooterManipulator(ast);
    const result = manipulator.setHeaderFooter({
      targetType: 'footer',
      pageType: 'default',
      text: 'Page 1 of 10'
    });

    expect(ast.footers[result.rId]).toBeDefined();

    manipulator.removeHeaderFooter({
      targetType: 'footer',
      pageType: 'default'
    });

    expect(ast.footers[result.rId]).toBeUndefined();

    const node = ast.sections[0].nodes[0];
    const sectPr = (node.xml.ownerDocument || node.xml).getElementsByTagName('w:sectPr')[0];
    expect(sectPr.getElementsByTagName('w:footerReference').length).toBe(0);
  });

  describe('header and footer text and page field token handling', () => {
    it('omits field instructions when text contains page in general prose', async () => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const manipulator = new HeaderFooterManipulator(ast, parser);
      manipulator.setHeaderFooter({
        targetType: 'footer',
        text: 'Draft page layout - confidential',
      });
      const zip = new AdmZip(parser.rezip());
      const xml = zip.readAsText('word/footer1.xml');
      expect(xml).not.toContain('instrText');
      expect(xml).toContain('Draft page layout - confidential');
    });

    it('expands PAGE and NUMPAGES explicit tokens in footer', async () => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const manipulator = new HeaderFooterManipulator(ast, parser);
      manipulator.setHeaderFooter({
        targetType: 'footer',
        text: 'Page {PAGE} of {NUMPAGES}',
      });
      const zip = new AdmZip(parser.rezip());
      const xml = zip.readAsText('word/footer1.xml');
      const instrMatches = xml.match(/<w:instrText[^>]*>(.*?)<\/w:instrText>/g) || [];
      expect(instrMatches.length).toBe(2);
      expect(xml).toContain('PAGE');
      expect(xml).toContain('NUMPAGES');
      expect(xml).toContain('<w:t xml:space="preserve">Page </w:t>');
      expect(xml).toContain('<w:t xml:space="preserve"> of </w:t>');
      expect(xml).not.toContain('{');
      expect(xml).not.toContain('}');
      await assertValidDocx(parser.rezip());
    });

    it('replaces page label with Page and page fields while dropping digits', async () => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const manipulator = new HeaderFooterManipulator(ast, parser);
      manipulator.setHeaderFooter({
        targetType: 'footer',
        text: 'Page 1 of 10',
      });
      const zip = new AdmZip(parser.rezip());
      const xml = zip.readAsText('word/footer1.xml');
      const instrMatches = xml.match(/<w:instrText[^>]*>(.*?)<\/w:instrText>/g) || [];
      expect(instrMatches.length).toBe(2);
      expect(xml).toContain('<w:t xml:space="preserve">Page </w:t>');
      expect(xml).not.toContain('1 of 10');
    });

    it('handles explicit tokens separated by delimiter run', async () => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const manipulator = new HeaderFooterManipulator(ast, parser);
      manipulator.setHeaderFooter({
        targetType: 'footer',
        text: '{PAGE} / {NUMPAGES}',
      });
      const zip = new AdmZip(parser.rezip());
      const xml = zip.readAsText('word/footer1.xml');
      const instrMatches = xml.match(/<w:instrText[^>]*>(.*?)<\/w:instrText>/g) || [];
      expect(instrMatches.length).toBe(2);
      expect(xml).toContain('<w:t xml:space="preserve"> / </w:t>');
    });

    it('supports PAGE token in header inside w:hdr', async () => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const manipulator = new HeaderFooterManipulator(ast, parser);
      manipulator.setHeaderFooter({
        targetType: 'header',
        text: 'Report - {PAGE}',
      });
      const zip = new AdmZip(parser.rezip());
      const xml = zip.readAsText('word/header1.xml');
      expect(xml).toContain('<w:hdr');
      expect(xml).toContain('PAGE');
      const instrMatches = xml.match(/<w:instrText[^>]*>(.*?)<\/w:instrText>/g) || [];
      expect(instrMatches.length).toBe(1);
      await assertValidDocx(parser.rezip());
    });

    it('applies number format switch to explicit page tokens', async () => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const manipulator = new HeaderFooterManipulator(ast, parser);
      manipulator.setHeaderFooter({
        targetType: 'footer',
        text: 'Page {PAGE}',
        numberFormat: 'lowerRoman',
      });
      const zip = new AdmZip(parser.rezip());
      const xml = zip.readAsText('word/footer1.xml');
      expect(xml).toContain('PAGE \\* roman');
    });
  });
});
