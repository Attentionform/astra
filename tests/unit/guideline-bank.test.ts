import { describe, it, expect } from 'vitest';
import {
  buildSystemPromptContext,
  getOOXMLGuidelinesContext,
  getOOXMLSnippet,
  OOXML_NAMESPACES,
  OOXML_SNIPPETS,
  SCHEMA_RULES,
} from '@attentionform/astra-engine';

describe('OOXML Schema Knowledge Bank (guideline-bank)', () => {
  it('should define authoritative OOXML namespaces', () => {
    expect(OOXML_NAMESPACES.w).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    expect(OOXML_NAMESPACES.r).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    expect(OOXML_NAMESPACES.a).toBe('http://schemas.openxmlformats.org/drawingml/2006/main');
    expect(OOXML_NAMESPACES.c).toBe('http://schemas.openxmlformats.org/drawingml/2006/chart');
    expect(OOXML_NAMESPACES.wp).toBe('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing');
    expect(OOXML_NAMESPACES.rel).toBe('http://schemas.openxmlformats.org/package/2006/relationships');
  });

  it('should contain all required snippet categories', () => {
    const categories = [
      'headersFooters',
      'pageNumbers',
      'tables',
      'charts',
      'paragraphsRuns',
      'sections',
      'metadataTracked'
    ];
    for (const cat of categories) {
      expect(OOXML_SNIPPETS[cat]).toBeDefined();
      expect(Object.keys(OOXML_SNIPPETS[cat]).length).toBeGreaterThan(0);
    }
  });

  it('should provide header and footer snippets with proper XML structure', () => {
    const hdrPart = getOOXMLSnippet('headersFooters', 'headerPart');
    expect(hdrPart).toContain('<w:hdr');
    expect(hdrPart).toContain('ASTra Executive Confidential Report');

    const ftrPart = getOOXMLSnippet('headersFooters', 'footerPart');
    expect(ftrPart).toContain('<w:ftr');
    expect(ftrPart).toContain('<w:fldSimple w:instr="PAGE"/>');
    expect(ftrPart).toContain('<w:fldSimple w:instr="NUMPAGES"/>');
  });

  it('should provide dynamic page number snippets', () => {
    const pageSimple = getOOXMLSnippet('pageNumbers', 'pageSimple');
    expect(pageSimple).toContain('w:fldSimple w:instr="PAGE \\* MERGEFORMAT"');

    const pageXofY = getOOXMLSnippet('pageNumbers', 'pageXofYParagraph');
    expect(pageXofY).toContain('w:fldSimple w:instr="PAGE"');
    expect(pageXofY).toContain('w:fldSimple w:instr="NUMPAGES"');
  });

  it('should provide executive zebra table snippet', () => {
    const tableXml = getOOXMLSnippet('tables', 'executiveTable');
    expect(tableXml).toContain('<w:tblHeader/>');
    expect(tableXml).toContain('<w:cantSplit/>');
    expect(tableXml).toContain('w:fill="003366"');
    expect(tableXml).toContain('w:fill="F7F9FC"');
  });

  it('should provide 3D chart DrawingML snippet', () => {
    const chart3d = getOOXMLSnippet('charts', 'barChart3DSpace');
    expect(chart3d).toContain('<c:chartSpace');
    expect(chart3d).toContain('<c:view3D>');
    expect(chart3d).toContain('<c:rotX val="30"/>');
    expect(chart3d).toContain('<c:bar3DChart>');
  });

  it('should retrieve snippets using getOOXMLSnippet helper', () => {
    const valid = getOOXMLSnippet('paragraphsRuns', 'heading1');
    expect(valid).toContain('Heading1');
    expect(valid).toContain('Executive Summary');

    const invalidCat = getOOXMLSnippet('nonExistentCat', 'heading1');
    expect(invalidCat).toBeNull();

    const invalidName = getOOXMLSnippet('paragraphsRuns', 'nonExistentSnippet');
    expect(invalidName).toBeNull();
  });

  it('should validate schema rules list', () => {
    expect(SCHEMA_RULES.length).toBeGreaterThanOrEqual(5);
    const elementOrderRule = SCHEMA_RULES.find(r => r.id === 'rule_element_ordering');
    expect(elementOrderRule).toBeDefined();
    expect(elementOrderRule?.exampleGood).toContain('<w:pPr>');
  });

  it('should format RAG context using getOOXMLGuidelinesContext', () => {
    const contextGeneral = getOOXMLGuidelinesContext();
    expect(contextGeneral).toContain('ASTRA AUTHORITATIVE OOXML SCHEMA GUIDELINES');
    expect(contextGeneral).toContain(OOXML_NAMESPACES.w);
    expect(contextGeneral).toContain('Strict Element Child Ordering');

    const contextTopic = getOOXMLGuidelinesContext('tables');
    expect(contextTopic).toContain('RELEVANT TEMPLATES FOR CATEGORY: TABLES');
    expect(contextTopic).toContain('Executive Styled Table with Zebra Striping');

    const systemPromptContext = buildSystemPromptContext();
    expect(systemPromptContext).toBe(contextGeneral);
  });
});
