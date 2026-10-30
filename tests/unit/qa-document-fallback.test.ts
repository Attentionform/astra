import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import fsp from 'fs/promises';
import path from 'path';
import os from 'os';
import AdmZip from 'adm-zip';

import {
  DocumentManager,
  DocxParser,
  OperationEngine,
} from '@attentionform/astra-engine';

describe('QA Verification Suite: Document Fallbacks & Engine Lifecycle', () => {
  // Suite budget: each create below runs OOXML validation plus a fresh git
  // repo init/commit; 5s quiet, 31s seen with the full suite competing for
  // CPU. 120s keeps slow runs green without hiding a real hang.
  const qaTempDir = path.join(os.tmpdir(), `astra_qa_test_${Date.now()}`);
  // Unique per run so fallback artifacts landing in SHARED dirs
  // (<cwd>/documents, os.tmpdir()/ASTra) never collide across runs, and their
  // git sidecars never get reused by a later run.
  const runTag = `${Date.now()}_${process.pid}`;
  // Fallback outputs tracked for post-test removal (file + git sidecar), so
  // runs stay hermetic and documents/ does not accumulate leftovers.
  const fallbackArtifacts: string[] = [];
  function trackFallback(docPath: string, targetPath: string): void {
    if (docPath !== targetPath) fallbackArtifacts.push(docPath);
  }

  beforeEach(() => {
    if (!fs.existsSync(qaTempDir)) {
      fs.mkdirSync(qaTempDir, { recursive: true });
    }
  });

  afterEach(() => {
    vi.restoreAllMocks();
    while (fallbackArtifacts.length > 0) {
      const p = fallbackArtifacts.pop() as string;
      try { fs.unlinkSync(p); } catch { /* ignore cleanup failures */ }
      try { fs.rmSync(path.join(path.dirname(p), '.astra', path.basename(p)), { recursive: true, force: true }); } catch { /* ignore cleanup failures */ }
    }
    if (fs.existsSync(qaTempDir)) {
      try {
        fs.rmSync(qaTempDir, { recursive: true, force: true });
      } catch { /* ignore cleanup failures */ }
    }
  });

  describe('DocumentManager Multi-Tier Creation & Fallbacks', () => {
    it('succeeds normally when target path is writable', async () => {
      const docManager = new DocumentManager();
      const validPath = path.join(qaTempDir, 'NormalDoc.docx');

      const result = await docManager.createNewDocument('blank', validPath);

      expect(result.astraDocPath).toBe(validPath);
      expect(fs.existsSync(validPath)).toBe(true);
      expect(docManager.getActiveDocPath()).toBe(validPath);
      expect(docManager.getIsUntitled()).toBe(false);

      // Verify OOXML validity
      const zip = new AdmZip(validPath);
      expect(zip.getEntry('word/document.xml')).not.toBeNull();
      expect(zip.getEntry('[Content_Types].xml')).not.toBeNull();
    });

    it('falls back to workspace documents/ when primary write fails with ENOENT (e.g. CFA blocked)', async () => {
      const docManager = new DocumentManager();
      // S1b sandbox: the primary must be sandbox-legit (fail-closed sandbox
      // throws on outside paths before any write) but unwritable — simulate
      // CFA/ENOENT via a write spy, preserving the test's fallback intent.
      const blockedPath = path.join(qaTempDir, 'cfa_blocked', `BlockedDoc_${runTag}.docx`);

      const originalWriteFile = fsp.writeFile;
      const writeFileSpy = vi.spyOn(fsp, 'writeFile').mockImplementation(async (targetFile: any, data: any, opts: any) => {
        if (targetFile.toString().includes('cfa_blocked')) {
          const err: any = new Error(`ENOENT: no such file or directory, open '${targetFile}'`);
          err.code = 'ENOENT';
          throw err;
        }
        return originalWriteFile(targetFile, data, opts);
      });

      let result!: Awaited<ReturnType<DocumentManager['createNewDocument']>>;
      try {
        result = await docManager.createNewDocument('blank', blockedPath);
      } finally {
        writeFileSpy.mockRestore();
      }
      trackFallback(result.astraDocPath, blockedPath);

      expect(result).toBeDefined();
      expect(result.astraDocPath).toBeDefined();
      expect(result.astraDocPath).not.toBe(blockedPath);
      expect(fs.existsSync(result.astraDocPath)).toBe(true);

      // Verify that isUntitled is set to true because finalPath !== targetPath
      expect(docManager.getIsUntitled()).toBe(true);

      // Verify valid DOCX was written
      const zip = new AdmZip(result.astraDocPath);
      expect(zip.getEntry('word/document.xml')).not.toBeNull();

      // Clean up fallback file
      try { fs.unlinkSync(result.astraDocPath); } catch { /* ignore cleanup failures */ }
    });

    it('falls back to OS temp directory when both primary path and workspace documents/ fail write', async () => {
      const docManager = new DocumentManager();
      const primaryPath = path.join(qaTempDir, 'primary_fail', `Doc_${runTag}.docx`);

      const originalWriteFile = fsp.writeFile;
      const writeFileSpy = vi.spyOn(fsp, 'writeFile').mockImplementation(async (targetFile: any, data: any, opts: any) => {
        const strPath = targetFile.toString();
        // Fail primary path
        if (strPath.includes('primary_fail')) {
          const err: any = new Error(`ENOENT: no such file or directory, open '${strPath}'`);
          err.code = 'ENOENT';
          throw err;
        }
        // Fail workspace documents directory
        const workspaceDocs = path.join(process.cwd(), 'documents');
        if (strPath.startsWith(workspaceDocs)) {
          const err: any = new Error(`EACCES: permission denied, open '${strPath}'`);
          err.code = 'EACCES';
          throw err;
        }
        // Allow OS temp directory
        return originalWriteFile(targetFile, data, opts);
      });

      const result = await docManager.createNewDocument('blank', primaryPath);
      trackFallback(result.astraDocPath, primaryPath);

      expect(result).toBeDefined();
      expect(result.astraDocPath).toBeDefined();
      // Should have fallen back to os.tmpdir()/ASTra/
      const expectedTempPrefix = path.join(os.tmpdir(), 'ASTra');
      expect(result.astraDocPath.startsWith(expectedTempPrefix)).toBe(true);
      expect(fs.existsSync(result.astraDocPath)).toBe(true);
      expect(docManager.getIsUntitled()).toBe(true);

      // Verify valid DOCX was written in temp
      const zip = new AdmZip(result.astraDocPath);
      expect(zip.getEntry('word/document.xml')).not.toBeNull();

      writeFileSpy.mockRestore();
      try { fs.unlinkSync(result.astraDocPath); } catch { /* ignore cleanup failures */ }
    });

    it('identifies silent Drafts fallback as untitled so save triggers Save-As', async () => {
      const docManager = new DocumentManager();
      const draftsDocPath = path.join(qaTempDir, 'Drafts', `ASTra_Document_${Date.now()}.docx`);

      await docManager.createNewDocument('blank', draftsDocPath);

      expect(docManager.getActiveDocPath()).toBe(draftsDocPath);
      expect(docManager.getIsUntitled()).toBe(true);
    });

    it('initializes clean git repo, resets undo/redo, and produces valid OOXML across templates', async () => {
      const docManager = new DocumentManager();
      const templates = ['blank', 'novel', 'financial', 'api-spec'];

      for (const tpl of templates) {
        const docPath = path.join(qaTempDir, `Template_${tpl}.docx`);
        const result = await docManager.createNewDocument(tpl, docPath);

        expect(result.summary.nodes.length).toBeGreaterThanOrEqual(1);
        expect(docManager.canUndo()).toBe(false);
        expect(docManager.canRedo()).toBe(false);
        expect(docManager.getIsDirty()).toBe(false);

        const parser = new DocxParser(fs.readFileSync(docPath));
        const validation = parser.validateZipIntegrity();
        expect(validation.isValid).toBe(true);
      }
    });
  });

  describe('Empty Document (0 Nodes) AST & Engine Lifecycle', () => {
    it('supports deleting all nodes down to 0 nodes without corrupting AST or parser', async () => {
      const docManager = new DocumentManager();
      const docPath = path.join(qaTempDir, 'DeleteAllNodes.docx');
      await docManager.createNewDocument('blank', docPath);

      const summary = docManager.getParser().getStructuralSummary();
      expect(summary.nodes.length).toBeGreaterThanOrEqual(2);

      // Delete all nodes one by one
      const parser = docManager.getParser();
      const ast = await parser.getAST();
      const opEngine = new OperationEngine(ast, parser);
      for (const node of summary.nodes) {
        await opEngine.execute({ action: 'delete', target: node.id });
      }

      const emptySummary = docManager.getParser().getStructuralSummary();
      expect(emptySummary.nodes.length).toBe(0);

      // DocxParser rezip should still succeed and maintain valid zip package
      const buf = docManager.getParser().rezip();
      expect(buf).toBeDefined();
      expect(buf.length).toBeGreaterThan(0);

      const zip = new AdmZip(buf);
      expect(zip.getEntry('word/document.xml')).not.toBeNull();
    });
  });
}, 120000);
