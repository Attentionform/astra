import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import { getOOXMLSnippet, validateNode } from '@attentionform/astra-engine';

const parser = new DOMParser();

describe('Raw XML Operations Test Suite', () => {

  describe('Header / Footer Injection', () => {
    it('should correctly format and parse header part XML template', () => {
      const headerXml = getOOXMLSnippet('headersFooters', 'headerPart');
      expect(headerXml).not.toBeNull();
      const doc = parser.parseFromString(headerXml!, 'text/xml');
      expect(doc.getElementsByTagName('w:hdr').length).toBe(1);

      const titleRun = doc.getElementsByTagName('w:t')[0];
      expect(titleRun.textContent).toBe('ASTra Executive Confidential Report');

      const border = doc.getElementsByTagName('w:bottom')[0];
      expect(border.getAttribute('w:color')).toBe('003366');
    });

    it('should correctly inject header & footer references into w:sectPr', () => {
      const sectPrXml = getOOXMLSnippet('headersFooters', 'sectPrReferences');
      expect(sectPrXml).not.toBeNull();
      const doc = parser.parseFromString(sectPrXml!, 'text/xml');

      const headerRef = doc.getElementsByTagName('w:headerReference')[0];
      expect(headerRef).toBeDefined();
      expect(headerRef.getAttribute('w:type')).toBe('default');
      expect(headerRef.getAttribute('r:id')).toBe('rIdH1');

      const footerRef = doc.getElementsByTagName('w:footerReference')[0];
      expect(footerRef).toBeDefined();
      expect(footerRef.getAttribute('w:type')).toBe('default');
      expect(footerRef.getAttribute('r:id')).toBe('rIdF1');
    });
  });

  describe('Dynamic Page Numbers (PAGE & NUMPAGES)', () => {
    it('should parse and structure simple dynamic page number fields', () => {
      const pageSnippet = getOOXMLSnippet('pageNumbers', 'pageSimple');
      const numpagesSnippet = getOOXMLSnippet('pageNumbers', 'numpagesSimple');

      const docPage = parser.parseFromString(pageSnippet!, 'text/xml');
      const fldPage = docPage.getElementsByTagName('w:fldSimple')[0];
      expect(fldPage.getAttribute('w:instr')).toContain('PAGE');

      const docNum = parser.parseFromString(numpagesSnippet!, 'text/xml');
      const fldNum = docNum.getElementsByTagName('w:fldSimple')[0];
      expect(fldNum.getAttribute('w:instr')).toContain('NUMPAGES');
    });

    it('should combine Page X of Y inside footer paragraph structure', () => {
      const footerPartXml = getOOXMLSnippet('headersFooters', 'footerPart');
      const doc = parser.parseFromString(footerPartXml!, 'text/xml');

      const fldSimples = doc.getElementsByTagName('w:fldSimple');
      expect(fldSimples.length).toBe(2);
      expect(fldSimples[0].getAttribute('w:instr')).toBe('PAGE');
      expect(fldSimples[1].getAttribute('w:instr')).toBe('NUMPAGES');

      const textNodes = Array.from(doc.getElementsByTagName('w:t')).map(t => t.textContent);
      expect(textNodes).toContain('Page ');
      expect(textNodes).toContain(' of ');
    });
  });

  describe('Executive Table Zebra Styling', () => {
    it('should parse executive table XML with w:tblPr, w:tblHeader, w:cantSplit, and zebra w:shd', () => {
      const tableXml = getOOXMLSnippet('tables', 'executiveTable');
      expect(tableXml).not.toBeNull();
      const doc = parser.parseFromString(tableXml!, 'text/xml');

      // Table properties
      const tblPr = doc.getElementsByTagName('w:tblPr')[0];
      expect(tblPr).toBeDefined();
      const tblJustification = tblPr.getElementsByTagName('w:jc')[0];
      expect(tblJustification.getAttribute('w:val')).toBe('center');

      // Row properties
      const rows = doc.getElementsByTagName('w:tr');
      expect(rows.length).toBe(2);

      const headerTrPr = rows[0].getElementsByTagName('w:trPr')[0];
      expect(headerTrPr.getElementsByTagName('w:tblHeader').length).toBe(1);
      expect(headerTrPr.getElementsByTagName('w:cantSplit').length).toBe(1);

      // Shading: Header row has #003366, Data row has zebra fill #F7F9FC
      const headerShading = rows[0].getElementsByTagName('w:shd')[0];
      expect(headerShading.getAttribute('w:fill')).toBe('003366');

      const dataShading = rows[1].getElementsByTagName('w:shd')[0];
      expect(dataShading.getAttribute('w:fill')).toBe('F7F9FC');
    });

    it('should validate table DOM manipulation without breaking schema', () => {
      const originalTableXml = `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
        <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>`;

      const editedTableXml = `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:tblPr>
          <w:tblW w:w="5000" w:type="pct"/>
          <w:tblJustification w:val="center"/>
        </w:tblPr>
        <w:tr>
          <w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>
          <w:tc>
            <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="003366"/></w:tcPr>
            <w:p><w:r><w:t>Header Updated</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>`;

      const isValid = validateNode(originalTableXml, editedTableXml, { allowStyleChanges: true, allowRunCountChanges: true });
      expect(isValid).toBe(true);
    });
  });

  describe('3D Chart DrawingML Injection', () => {
    it('should parse 3D chartSpace XML with c:view3D and perspective controls', () => {
      const chartXml = getOOXMLSnippet('charts', 'barChart3DSpace');
      expect(chartXml).not.toBeNull();
      const doc = parser.parseFromString(chartXml!, 'text/xml');

      const chartSpace = doc.getElementsByTagName('c:chartSpace')[0];
      expect(chartSpace).toBeDefined();

      const view3D = doc.getElementsByTagName('c:view3D')[0];
      expect(view3D).toBeDefined();
      expect(doc.getElementsByTagName('c:rotX')[0].getAttribute('val')).toBe('30');
      expect(doc.getElementsByTagName('c:rotY')[0].getAttribute('val')).toBe('20');
      expect(doc.getElementsByTagName('c:perspective')[0].getAttribute('val')).toBe('30');

      const bar3D = doc.getElementsByTagName('c:bar3DChart')[0];
      expect(bar3D).toBeDefined();
    });

    it('should embed chart into paragraph via w:drawing element', () => {
      const embedXml = getOOXMLSnippet('charts', 'drawingEmbedding');
      expect(embedXml).not.toBeNull();
      const doc = parser.parseFromString(embedXml!, 'text/xml');

      const drawing = doc.getElementsByTagName('w:drawing')[0];
      expect(drawing).toBeDefined();

      const graphicData = doc.getElementsByTagName('a:graphicData')[0];
      expect(graphicData.getAttribute('uri')).toBe('http://schemas.openxmlformats.org/drawingml/2006/chart');

      const chartRef = doc.getElementsByTagName('c:chart')[0];
      expect(chartRef.getAttribute('r:id')).toBe('rIdC1');
    });
  });

  describe('Typography & Callout Box Updates', () => {
    it('should parse styled heading 1 paragraph with proper font, size (32 = 16pt), and color', () => {
      const headingXml = getOOXMLSnippet('paragraphsRuns', 'heading1');
      const doc = parser.parseFromString(headingXml!, 'text/xml');

      const pStyle = doc.getElementsByTagName('w:pStyle')[0];
      expect(pStyle.getAttribute('w:val')).toBe('Heading1');

      const font = doc.getElementsByTagName('w:rFonts')[0];
      expect(font.getAttribute('w:ascii')).toBe('Helvetica');

      const size = doc.getElementsByTagName('w:sz')[0];
      expect(size.getAttribute('w:val')).toBe('32'); // 16pt

      const color = doc.getElementsByTagName('w:color')[0];
      expect(color.getAttribute('w:val')).toBe('003366');
    });

    it('should parse executive callout box with thick left border and background shading', () => {
      const calloutXml = getOOXMLSnippet('paragraphsRuns', 'calloutBox');
      const doc = parser.parseFromString(calloutXml!, 'text/xml');

      const leftBorder = doc.getElementsByTagName('w:left')[0];
      expect(leftBorder.getAttribute('w:val')).toBe('single');
      expect(leftBorder.getAttribute('w:sz')).toBe('36');
      expect(leftBorder.getAttribute('w:color')).toBe('0066CC');

      const shd = doc.getElementsByTagName('w:shd')[0];
      expect(shd.getAttribute('w:fill')).toBe('F0F4F8');
    });
  });

});
