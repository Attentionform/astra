import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  createBlankDocxBuffer,
  DocumentManager,
  DocxParser,
} from '@attentionform/astra-engine';

import fs from 'fs';
import fsp from 'fs/promises';
import os from 'os';
import path from 'path';
import AdmZip from 'adm-zip';

describe('New Empty Document & Template Creation Suite', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-create-doc-'));
  });

  afterEach(() => {
    if (tempDir && fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  });

  describe('Blank Document Generation', () => {
    it('should create a valid minimal OOXML document for blank template by default', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const summary = parser.getStructuralSummary();

      expect(summary).toBeDefined();
      expect(summary.nodes.length).toBeGreaterThanOrEqual(2);
      expect(summary.nodes[0].text).toContain('Blank Document');
      expect(summary.nodes[1].text).toContain('Start typing your content here');
      
      const validation = parser.validateZipIntegrity();
      expect(validation.isValid).toBe(true);
      expect(validation.errors).toHaveLength(0);
    });

    it('should fallback to blank template when unknown template name is provided', async () => {
      const parser = DocxParser.createNewDocument('unknown-template-xyz');
      const summary = parser.getStructuralSummary();

      expect(summary.nodes[0].text).toContain('Blank Document');
    });

    it('should produce a zip package with all required OOXML structures', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const buf = parser.rezip();
      const zip = new AdmZip(buf);

      expect(zip.getEntry('[Content_Types].xml')).not.toBeNull();
      expect(zip.getEntry('_rels/.rels')).not.toBeNull();
      expect(zip.getEntry('word/_rels/document.xml.rels')).not.toBeNull();
      expect(zip.getEntry('word/document.xml')).not.toBeNull();
      expect(zip.getEntry('word/styles.xml')).not.toBeNull();

      const docXml = zip.getEntry('word/document.xml')?.getData().toString('utf8');
      expect(docXml).toContain('<w:document');
      expect(docXml).toContain('<w:body>');
      expect(docXml).toMatch(/<w:pgSz\s+[^>]*w:w="12240"\s+[^>]*w:h="15840"/);
      expect(docXml).toContain('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    });

    it('should generate valid standalone buffer using createBlankDocxBuffer utility function', async () => {
      const blankBuf = createBlankDocxBuffer('blank');
      expect(blankBuf).toBeDefined();
      expect(blankBuf.length).toBeGreaterThan(0);

      const zip = new AdmZip(blankBuf);
      expect(zip.getEntry('word/document.xml')).not.toBeNull();
      expect(zip.getEntry('word/styles.xml')).not.toBeNull();

      const parser = new DocxParser(blankBuf);
      const summary = parser.getStructuralSummary();
      expect(summary.nodes[0].text).toContain('New Document');
    });

    it('should generate template parsers for novel, financial, and api-spec', async () => {
      const novelParser = DocxParser.createNewDocument('novel');
      const novelSummary = novelParser.getStructuralSummary();
      expect(novelSummary.nodes[0].preview).toContain('Manuscript Title');

      const reportParser = DocxParser.createNewDocument('financial');
      const reportSummary = reportParser.getStructuralSummary();
      expect(reportSummary.nodes[0].preview).toContain('Financial Report');

      const apiParser = DocxParser.createNewDocument('api-spec');
      const apiSummary = apiParser.getStructuralSummary();
      expect(apiSummary.nodes[0].preview).toContain('API Specification');
    });
  });

  describe('AST Node Structure Verification', () => {
    it('should build a complete DocumentAST representation with root section', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      expect(ast).toBeDefined();
      expect(ast.sections).toHaveLength(1);
      expect(ast.sections[0].id).toBe('sec_root');
      expect(ast.sections[0].nodes.length).toBeGreaterThanOrEqual(2);
    });

    it('should assign stable node IDs and data-astra-id attributes to document nodes', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const summary = parser.getStructuralSummary();

      expect(summary.nodes[0].id).toBeDefined();
      expect(summary.nodes[1].id).toBeDefined();
      // A-4: ids are stable minted w14:paraIds (valid ST_LongHexNumber), not positional p_NNN
      for (const n of [summary.nodes[0], summary.nodes[1]]) {
        expect(/^[0-9A-Fa-f]{8}$/.test(n.id)).toBe(true);
      }
      expect(summary.nodes[0].id).not.toBe(summary.nodes[1].id);

      const nodeXml = parser.readNode(summary.nodes[0].id);
      expect(nodeXml).toContain(`data-astra-id="${summary.nodes[0].id}"`);
      expect(nodeXml).toContain('Blank Document');
    });

    it('should correctly build AST node metadata and checksums', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      const firstNode = ast.sections[0].nodes[0];

      expect(firstNode.id).toBeDefined();
      expect(firstNode.type).toBe('paragraph');
      expect(firstNode.version).toBe(1);
      expect(firstNode.checksum).toBeDefined();
      expect(firstNode.checksum.length).toBeGreaterThan(0);
      expect(firstNode.metadata.text).toContain('Blank Document');
    });

    it('should correctly classify heading nodes vs regular body text nodes', async () => {
      const parser = DocxParser.createNewDocument('novel');
      const summary = parser.getStructuralSummary();

      const headingNode = summary.nodes[2];
      const bodyNode = summary.nodes[3];

      expect(headingNode?.isHeading).toBe(true);
      expect(bodyNode?.isHeading).toBe(false);
    });
  });

  describe('State Reset & DocumentManager Life-Cycle', () => {
    it('should initialize DocumentManager state, reset undo/redo stacks, and create Git repo', async () => {
      const docManager = new DocumentManager();
      const docPath = path.join(tempDir, 'NewBlankDoc.docx');

      const result = await docManager.createNewDocument('blank', docPath);

      expect(result.astraDocPath).toBe(docPath);
      expect(fs.existsSync(docPath)).toBe(true);

      const repoDir = path.join(tempDir, '.astra');
      expect(fs.existsSync(repoDir)).toBe(true);

      expect(docManager.getActiveDocPath()).toBe(docPath);
      expect(docManager.canUndo()).toBe(false);
      expect(docManager.canRedo()).toBe(false);

      const snapshotHistory = docManager.getSnapshotHistory();
      expect(snapshotHistory.length).toBe(1);
      expect(snapshotHistory[0].label).toBe('Initial state');
    });

    it('should completely reset active state when creating a new document after edits on previous doc', async () => {
      const docManager = new DocumentManager();
      const docPath1 = path.join(tempDir, 'Doc1.docx');
      const docPath2 = path.join(tempDir, 'Doc2.docx');

      // Create doc 1 and make an edit
      await docManager.createNewDocument('blank', docPath1);
      const firstId = docManager.getParser().getStructuralSummary().nodes[0].id;
      await docManager.updateNode(firstId, '<w:p><w:r><w:t>Modified Doc 1</w:t></w:r></w:p>', 'Edited doc 1');
      
      expect(docManager.canUndo()).toBe(true);
      expect(docManager.getSnapshotHistory().length).toBe(2);

      // Now create doc 2 - state must reset completely
      await docManager.createNewDocument('novel', docPath2);

      expect(docManager.getActiveDocPath()).toBe(docPath2);
      expect(docManager.canUndo()).toBe(false);
      expect(docManager.canRedo()).toBe(false);
      expect(docManager.getSnapshotHistory().length).toBe(1);
      
      const newSummary = docManager.getParser().getStructuralSummary();
      expect(newSummary.nodes[0].text).toBe('Manuscript Title');
    });

    it('should commit new document changes cleanly when overwriting/recreating in an existing Git repo', async () => {
      const docManager = new DocumentManager();
      const docPath = path.join(tempDir, 'RecreatedDoc.docx');

      // First creation initializes Git repo
      await docManager.createNewDocument('blank', docPath);
      
      // Second creation in same directory should commit change rather than erroring on init
      const result2 = await docManager.createNewDocument('api-spec', docPath);
      
      expect(result2.summary.nodes[0].text).toContain('API Specification');
      expect(docManager.canUndo()).toBe(false);
      expect(docManager.canRedo()).toBe(false);
    });

    it('should gracefully fall back and create document when target path fails write', async () => {
      const docManager = new DocumentManager();
      // S1b sandbox: fail-closed validation throws on outside paths before
      // any write, so exercise the write-failure fallback with a
      // sandbox-legit (tmpdir) target whose write is forced to ENOENT.
      // Basename is unique per run: the fallback lands in the SHARED
      // <cwd>/documents dir, and a fixed name collides with other suites'
      // fallback tests running in parallel workers (same git sidecar,
      // torn .git/index). See tests/helpers note pattern in
      // qa-document-fallback.test.ts.
      const invalidPath = path.join(tempDir, 'write_fail', `BlockedDoc_${Date.now()}_${process.pid}.docx`);

      const originalWriteFile = fsp.writeFile;
      const writeFileSpy = vi.spyOn(fsp, 'writeFile').mockImplementation(async (targetFile: any, data: any, opts: any) => {
        if (targetFile.toString().includes('write_fail')) {
          const err: any = new Error(`ENOENT: no such file or directory, open '${targetFile}'`);
          err.code = 'ENOENT';
          throw err;
        }
        return originalWriteFile(targetFile, data, opts);
      });

      let result!: Awaited<ReturnType<DocumentManager['createNewDocument']>>;
      try {
        result = await docManager.createNewDocument('blank', invalidPath);
      } finally {
        writeFileSpy.mockRestore();
      }
      expect(result).toBeDefined();
      expect(result.astraDocPath).toBeDefined();
      expect(result.summary.nodes.length).toBeGreaterThanOrEqual(2);
      expect(fs.existsSync(result.astraDocPath)).toBe(true);
      // Clean up fallback file if in documents or temp, plus its git sidecar
      // (.astra/<basename>) so shared dirs stay hermetic across runs.
      if (fs.existsSync(result.astraDocPath)) {
        try { fs.unlinkSync(result.astraDocPath); } catch { /* ignore cleanup failures */ }
      }
      try {
        fs.rmSync(path.join(path.dirname(result.astraDocPath), '.astra', path.basename(result.astraDocPath)), { recursive: true, force: true });
      } catch { /* ignore cleanup failures */ }
    });
  });
});
