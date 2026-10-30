import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { DocxParser, OperationEngine, type CommentOperation, type ReplaceOperation } from '@attentionform/astra-engine';

import * as path from 'path';
import * as os from 'os';
import * as fs from 'fs';
import AdmZip from 'adm-zip';

describe('Phase 10: Legal Features', () => {
  let parser: DocxParser;
  let engine: OperationEngine;
  let tmpDir: string;
  let tempFilePath: string;

  beforeEach(async () => {
    // Create a fresh test file under the system temp dir (no repo writes).
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-legal-'));
    tempFilePath = path.join(tmpDir, 'legal_test.docx');
    const zip = new AdmZip();
    zip.addFile('word/document.xml', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="11111111"><w:r><w:t>Hello world, this is a test.</w:t></w:r></w:p>
    <w:p w14:paraId="22222222"><w:r><w:t>Second paragraph for tracking.</w:t></w:r></w:p>
  </w:body>
</w:document>`, 'utf8'));
    zip.writeZip(tempFilePath);

    parser = new DocxParser(fs.readFileSync(tempFilePath));
    const ast = await parser.getAST();
    // Simulate these are edits to existing content (version > 1)
    ast.sections[0].nodes.forEach(n => n.version = 2);
    engine = new OperationEngine(ast, parser);
  });

  afterEach(async () => {
    try {
      if (tmpDir && fs.existsSync(tmpDir)) {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    } catch {
      /* best-effort cleanup */
    }
  });

  it('should emit unique w:id for tracked changes across multiple nodes in one batch', async () => {
    const result = await engine.executeBatch({
      operations: [
        { action: 'replace', target: '11111111', text: 'Hello ASTra, this is a test.', expectedVersion: 2, trackChanges: true },
        { action: 'replace', target: '22222222', text: 'Second paragraph is tracked.', expectedVersion: 2, trackChanges: true }
      ] as ReplaceOperation[],
      sequential: true
    });

    expect(result.success).toBe(true);
    
    // Save to trigger AST serialization back to zip
    const outBuffer = parser.serializeAST();
    fs.writeFileSync(tempFilePath, outBuffer);

    const outZip = new AdmZip(tempFilePath);
    const docXml = outZip.getEntry('word/document.xml')?.getData().toString('utf8');
    
    // Extract all w:id from w:ins and w:del
    const insIds = Array.from(docXml!.matchAll(/<w:ins[^>]*w:id="(\d+)"/g)).map(m => m[1]);
    const delIds = Array.from(docXml!.matchAll(/<w:del[^>]*w:id="(\d+)"/g)).map(m => m[1]);
    
    const allIds = [...insIds, ...delIds];
    const uniqueIds = new Set(allIds);
    
    expect(allIds.length).toBeGreaterThan(0);
    expect(allIds.length).toBe(uniqueIds.size); // No duplicates
  });

  it('should add comment and bootstrap side parts', async () => {
    const result = await engine.executeBatch({
      operations: [
        { action: 'comment', operation: 'insert', target: '11111111', text: 'Please review.', expectedVersion: 2 }
      ] as CommentOperation[],
      sequential: true
    });

    expect(result.success).toBe(true);

    const outBuffer = parser.serializeAST();
    fs.writeFileSync(tempFilePath, outBuffer);

    const outZip = new AdmZip(tempFilePath);
    expect(outZip.getEntry('word/comments.xml')).toBeDefined();
    expect(outZip.getEntry('word/commentsExtended.xml')).toBeDefined();
    expect(outZip.getEntry('word/people.xml')).toBeDefined();

    const commentsXml = outZip.getEntry('word/comments.xml')?.getData().toString('utf8');
    expect(commentsXml).toContain('Please review.');
  });
});
