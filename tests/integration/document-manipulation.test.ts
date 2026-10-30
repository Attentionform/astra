import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as path from 'path';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import os from 'os';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocumentManager,
  DocxParser,
  getOOXMLSnippet,
  OperationEngine,
  validateNode,
} from '@attentionform/astra-engine';

const domParser = new DOMParser();

describe('End-to-End Document Manipulation Features Integration Suite', () => {
  let tmpDir: string;
  let testDocPath: string;

  beforeEach(async () => {
    tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-e2e-doc-'));
    testDocPath = path.join(tmpDir, 'sample_doc.docx');

    // Create a valid baseline DOCX package using AdmZip
    const zip = new AdmZip();
    zip.addFile('word/document.xml', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="11111111">
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Document Header Title</w:t></w:r>
    </w:p>
    <w:p w14:paraId="22222222">
      <w:r><w:t>Initial body content paragraph for integration test.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`, 'utf8'));

    zip.addFile('word/_rels/document.xml.rels', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`, 'utf8'));

    zip.addFile('[Content_Types].xml', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`, 'utf8'));

    zip.addFile('_rels/.rels', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`, 'utf8'));

    zip.writeZip(testDocPath);
  });

  afterEach(async () => {
    await fsp.rm(tmpDir, { recursive: true, force: true });
  });

  // --------------------------------------------------------------------------
  // 1. Headers and Footers Integration
  // --------------------------------------------------------------------------
  describe('1. Headers & Footers Integration', () => {
    it('should inject header/footer XML parts, link relationships, and update w:sectPr references', async () => {
      const buffer = fs.readFileSync(testDocPath);
      const parser = new DocxParser(buffer);
      const zipAdapter = parser.getZipAdapter();

      // Get templates from Guideline Bank
      const headerXml = getOOXMLSnippet('headersFooters', 'headerPart');
      const footerXml = getOOXMLSnippet('headersFooters', 'footerPart');
      expect(headerXml).not.toBeNull();
      expect(footerXml).not.toBeNull();

      // Write header and footer parts
      zipAdapter.writeText('word/header1.xml', headerXml!);
      zipAdapter.writeText('word/footer1.xml', footerXml!);

      // Add relationships
      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;
      zipAdapter.writeText('word/_rels/document.xml.rels', relsXml);

      // Add Content_Types overrides
      const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`;
      zipAdapter.writeText('[Content_Types].xml', contentTypes);

      // Execute headerFooter operation on AST
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      await engine.executeBatch({
        operations: [
          { action: 'headerFooter', operation: 'set', targetType: 'header', pageType: 'default', text: 'ASTra Executive Confidential Report' },
          { action: 'headerFooter', operation: 'set', targetType: 'footer', pageType: 'default', text: 'Page 1 of 1' }
        ] as any,
        sequential: true
      });

      // Save AST / Rezip document
      const outBuffer = parser.rezip();
      expect(outBuffer.length).toBeGreaterThan(0);

      // Verify re-packed DOCX
      const reloadedParser = new DocxParser(outBuffer);
      const reloadedAdapter = reloadedParser.getZipAdapter();

      expect(reloadedAdapter.hasFile('word/header1.xml')).toBe(true);
      expect(reloadedAdapter.hasFile('word/footer1.xml')).toBe(true);

      const savedHeaderStr = reloadedAdapter.readText('word/header1.xml');
      const savedFooterStr = reloadedAdapter.readText('word/footer1.xml');

      expect(savedHeaderStr).toContain('ASTra Executive Confidential Report');
      expect(savedFooterStr).toContain('PAGE');
      expect(savedFooterStr).toContain('NUMPAGES');

      const savedDocXml = reloadedParser.readDocumentXml();
      expect(savedDocXml).toContain('w:headerReference');
      expect(savedDocXml).toContain('w:footerReference');
    });
  });

  // --------------------------------------------------------------------------
  // 2. 3D Charts DrawingML Integration
  // --------------------------------------------------------------------------
  describe('2. 3D Charts DrawingML Integration', () => {
    it('should inject 3D bar chart space XML, add DrawingML reference, and verify perspective controls', async () => {
      const buffer = fs.readFileSync(testDocPath);
      const parser = new DocxParser(buffer);
      const zipAdapter = parser.getZipAdapter();

      // Retrieve 3D ChartML snippet from Guideline Bank
      const chart3DXml = getOOXMLSnippet('charts', 'barChart3DSpace');
      const drawingEmbedXml = getOOXMLSnippet('charts', 'drawingEmbedding');

      expect(chart3DXml).not.toBeNull();
      expect(drawingEmbedXml).not.toBeNull();

      // Write chart part to zip
      zipAdapter.writeText('word/charts/chart1.xml', chart3DXml!);

      // Add relationship for chart
      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>`;
      zipAdapter.writeText('word/_rels/document.xml.rels', relsXml);

      // Insert Drawing element into document DOM
      const drawingId = parser.insertRootNode(drawingEmbedXml!, '22222222', 'after');
      expect(drawingId).toBeTruthy();

      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      // Execute chart operation via engine
      await engine.execute({
        action: 'chart',
        operation: 'insert',
        chartType: 'bar3D',
        title: 'Quarterly Revenue Performance'
      } as any);

      // Re-serialize AST & Rezip
      const outBuffer = parser.rezip();
      const reloadedParser = new DocxParser(outBuffer);
      const reloadedAdapter = reloadedParser.getZipAdapter();

      // Verify chart XML part
      expect(reloadedAdapter.hasFile('word/charts/chart1.xml')).toBe(true);
      const chartContent = reloadedAdapter.readText('word/charts/chart1.xml');

      const chartDom = domParser.parseFromString(chartContent, 'text/xml');
      expect(chartDom.getElementsByTagName('c:view3D').length).toBe(1);
      expect(chartDom.getElementsByTagName('c:rotX')[0].getAttribute('val')).toBe('30');
      expect(chartDom.getElementsByTagName('c:rotY')[0].getAttribute('val')).toBe('20');
      expect(chartDom.getElementsByTagName('c:perspective')[0].getAttribute('val')).toBe('30');

      // Verify drawing node in main document
      const docXml = reloadedParser.readDocumentXml();
      expect(docXml).toContain('w:drawing');
      expect(docXml).toContain('http://schemas.openxmlformats.org/drawingml/2006/chart');
    });
  });

  // --------------------------------------------------------------------------
  // 3. Zebra Tables Integration
  // --------------------------------------------------------------------------
  describe('3. Zebra Tables Integration', () => {
    it('should insert executive zebra table, manipulate rows/cells, and enforce zebra shading', async () => {
      const buffer = fs.readFileSync(testDocPath);
      const parser = new DocxParser(buffer);

      const tableXml = getOOXMLSnippet('tables', 'executiveTable');
      expect(tableXml).not.toBeNull();

      // Insert executive table
      const tableNodeId = parser.insertRootNode(tableXml!, '22222222', 'after');
      expect(tableNodeId).toBeTruthy();

      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      // Execute table operations (update cell, add row, set zebra shading)
      const batchResult = await engine.executeBatch({
        operations: [
          { action: 'table', target: tableNodeId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'Q1 2026 Updated' } },
          { action: 'table', target: tableNodeId, params: { op: 'addRow', afterIndex: 2 } },
          { action: 'table', target: tableNodeId, params: { op: 'updateCell', rowIndex: 3, colIndex: 1, text: 'Q2 2026' } },
          { action: 'table', target: tableNodeId, params: { op: 'setShading', fill: 'F7F9FC', zebra: true } }
        ] as any,
        sequential: true
      });

      expect(batchResult.success).toBe(true);

      // Validate modified node XML against schema rules
      const readXml = parser.readNode(tableNodeId);
      const isValid = validateNode(tableXml!, readXml, { allowStyleChanges: true, allowRunCountChanges: true });
      expect(isValid).toBe(true);

      // Save and verify rezipped DOCX XML output
      const outBuffer = parser.rezip();
      const reloadedParser = new DocxParser(outBuffer);
      const docXml = reloadedParser.readDocumentXml();

      const docDom = domParser.parseFromString(docXml, 'text/xml');
      const tableDom = docDom.getElementsByTagName('w:tbl')[0];
      expect(tableDom).toBeDefined();

      const rows = Array.from(tableDom.getElementsByTagName('w:tr'));
      expect(rows.length).toBeGreaterThanOrEqual(3);

      // Header row check
      const headerTrPr = rows[0].getElementsByTagName('w:trPr')[0];
      expect(headerTrPr.getElementsByTagName('w:tblHeader').length).toBe(1);
      expect(headerTrPr.getElementsByTagName('w:cantSplit').length).toBe(1);

      // Zebra shading check: row 1 shading = #F7F9FC
      const dataShading = rows[1].getElementsByTagName('w:shd')[0];
      expect(dataShading.getAttribute('w:fill')).toBe('F7F9FC');
    });
  });

  // --------------------------------------------------------------------------
  // 4. Raw XML Operations Integration
  // --------------------------------------------------------------------------
  describe('4. Raw XML Operations Integration', () => {
    it('should perform raw XML insert, read, update, replace, delete and enforce AST integrity', async () => {
      const buffer = fs.readFileSync(testDocPath);
      const parser = new DocxParser(buffer);
      let ast = await parser.getAST();

      // 1. Insert raw XML paragraph
      const rawInsertXml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Raw Injected Node</w:t></w:r></w:p>`;
      const insertedId = parser.insertRootNode(rawInsertXml, '11111111', 'after');
      expect(insertedId).toBeTruthy();
      ast = await parser.getAST();

      // 2. Read node
      const readXml = parser.readNode(insertedId);
      expect(readXml).toContain('Raw Injected Node');

      // 3. Update node in AST with version and checksum increment
      const updatedXml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Raw Injected Node Updated</w:t></w:r></w:p>`;
      await parser.updateNodeInAST(ast, insertedId, updatedXml);

      const astNode = parser.getNodeFromAST(ast, insertedId);
      expect(astNode).toBeDefined();
      expect(astNode?.version).toBe(2);
      expect(astNode?.metadata.text).toBe('Raw Injected Node Updated');

      // 4. Replace root node
      const replacementXml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Replaced Root Content</w:t></w:r></w:p>`;
      parser.replaceRootNode(insertedId, replacementXml);
      expect(parser.readNode(insertedId)).toContain('Replaced Root Content');

      // 5. Delete root node
      parser.deleteRootNode(insertedId);
      expect(() => parser.readNode(insertedId)).toThrow();

      // 6. Security Check: Reject XXE / DTD expansion attempts
      const xxeXml = `<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>&xxe;</w:t></w:r></w:p>`;
      expect(() => parser.insertRootNode(xxeXml)).toThrow(/XXE \/ DTD entity declarations are strictly prohibited/);
    });
  });

  // --------------------------------------------------------------------------
  // 5. Git Snapshot Undo/Redo Integration
  // --------------------------------------------------------------------------
  describe('5. Git Snapshot Undo/Redo Integration', () => {
    it('should track commits in isomorphic-git and restore exact XML state via undoLastAction', async () => {
      const docManager = new DocumentManager();
      
      // 1. Open document and initialize git repository
      const openResult = await docManager.openDocument(testDocPath);
      expect(openResult.astraDocPath).toBe(testDocPath);
      expect(openResult.summary.nodes.length).toBeGreaterThan(0);

      // 2. Perform Edit 1: Update node p_001
      const edit1Xml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Edit 1 Title</w:t></w:r></w:p>`;
      await docManager.updateNode('11111111', edit1Xml, 'Edit 1: Modify title');

      // 3. Perform Edit 2: Update node p_002
      const edit2Xml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:r><w:t>Edit 2 Body Text</w:t></w:r></w:p>`;
      await docManager.updateNode('22222222', edit2Xml, 'Edit 2: Modify body');

      // Check version history stack
      let history = await docManager.getVersionHistory();
      expect(history.length).toBe(3); // Initial + Edit 1 + Edit 2
      expect(history[0].message).toBe('Edit 2: Modify body\n');
      expect(history[1].message).toBe('Edit 1: Modify title\n');
      expect(history[2].message).toBe('Initial state\n');

      // Verify node content after Edit 2
      expect(docManager.getParser().readNode('22222222')).toContain('Edit 2 Body Text');

      // 4. Undo Edit 2
      await docManager.undoLastAction();

      history = await docManager.getVersionHistory();
      expect(history.length).toBe(2);
      expect(history[0].message).toBe('Edit 1: Modify title\n');
      expect(docManager.getParser().readNode('11111111')).toContain('Edit 1 Title');
      expect(docManager.getParser().readNode('22222222')).toContain('Initial body content paragraph');

      // 5. Undo Edit 1
      await docManager.undoLastAction();

      history = await docManager.getVersionHistory();
      expect(history.length).toBe(1);
      expect(history[0].message).toBe('Initial state\n');
      expect(docManager.getParser().readNode('11111111')).toContain('Document Header Title');

      // 6. Attempting to undo past initial state should return atInitialState: true cleanly
      const initialRes = await docManager.undoLastAction();
      expect(initialRes.atInitialState).toBe(true);
      expect(initialRes.success).toBe(false);
    });
  });
});
