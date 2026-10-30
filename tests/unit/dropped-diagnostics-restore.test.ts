/**
 * D1/D2/D3 dropped-diagnostic restore proofs.
 * Each dropped console signal now rides a typed channel with identical text;
 * library calls stay silent (console gate = 0 in every test).
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import fs from 'fs';
import fsp from 'fs/promises';
import os from 'os';
import path from 'path';
import {
  DocxParser,
  DocumentManager,
  OperationEngine,
  createBlankDocxBuffer,
  getCommentXmlWarnings,
  clearCommentXmlWarnings,
  bootstrapCommentParts,
  addComment,
} from '@attentionform/astra-engine';

afterEach(() => {
  vi.restoreAllMocks();
});

function silence() {
  const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  const error = vi.spyOn(console, 'error').mockImplementation(() => undefined);
  const info = vi.spyOn(console, 'info').mockImplementation(() => undefined);
  return {
    calls() {
      return log.mock.calls.length + warn.mock.calls.length + error.mock.calls.length + info.mock.calls.length;
    },
    restore() {
      log.mockRestore();
      warn.mockRestore();
      error.mockRestore();
      info.mockRestore();
    },
  };
}

describe('D1 createNewDocument write-failure reason rides warnings[]', () => {
  it('success carries an empty warnings array', async () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-d1-ok-'));
    try {
      const mgr = new DocumentManager();
      const target = path.join(tmp, 'Ok.docx');
      const gate = silence();
      let result!: Awaited<ReturnType<DocumentManager['createNewDocument']>>;
      try {
        result = await mgr.createNewDocument('blank', target);
      } finally {
        gate.restore();
      }
      expect(result.astraDocPath).toBe(target);
      expect(result.warnings).toEqual([]);
      expect(fs.existsSync(target)).toBe(true);
      expect(gate.calls()).toBe(0);
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  });

  it('primary-write failure reason preserved verbatim in warnings[]', async () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-d1-fail-'));
    try {
      const mgr = new DocumentManager();
      // Basename is unique per run: the fallback lands in the SHARED
      // <cwd>/documents dir, and a fixed name collides with other suites'
      // fallback tests in parallel workers (same git sidecar, torn
      // .git/index). The write_fail dir substring stays fixed for the
      // spy and the warnings assertions below.
      const invalidPath = path.join(tmp, 'write_fail', `BlockedDoc_${Date.now()}_${process.pid}.docx`);
      const originalWriteFile = fsp.writeFile;
      const writeFileSpy = vi.spyOn(fsp, 'writeFile').mockImplementation(async (targetFile: any, data: any, opts: any) => {
        if (targetFile.toString().includes('write_fail')) {
          const err: any = new Error(`ENOENT: no such file or directory, open '${targetFile}'`);
          err.code = 'ENOENT';
          throw err;
        }
        return originalWriteFile(targetFile, data, opts);
      });
      const gate = silence();
      let result!: Awaited<ReturnType<DocumentManager['createNewDocument']>>;
      try {
        result = await mgr.createNewDocument('blank', invalidPath);
      } finally {
        gate.restore();
        writeFileSpy.mockRestore();
      }
      expect(result.astraDocPath).toBeDefined();
      expect(result.astraDocPath).not.toBe(invalidPath);
      expect(Array.isArray(result.warnings)).toBe(true);
      expect(result.warnings.length).toBeGreaterThanOrEqual(1);
      // Same text as the dropped console.warn, WHY preserved alongside the path.
      expect(result.warnings[0]).toMatch(/\[DocumentManager\] Primary write failed for .*write_fail.*ENOENT/);
      expect(result.warnings[0]).toMatch(/Trying workspace documents fallback/);
      expect(fs.existsSync(result.astraDocPath)).toBe(true);
      expect(gate.calls()).toBe(0);
      if (fs.existsSync(result.astraDocPath)) {
        try { fs.unlinkSync(result.astraDocPath); } catch { /* best-effort */ }
      }
      // Plus its git sidecar (.astra/<basename>) so shared dirs stay
      // hermetic across runs.
      try {
        fs.rmSync(path.join(path.dirname(result.astraDocPath), '.astra', path.basename(result.astraDocPath)), { recursive: true, force: true });
      } catch { /* best-effort */ }
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  });
});

describe('D2 non-blocking XML warnings ride typed channels', () => {
  it('validateZipIntegrity carries warnings[] with the dropped warn text', () => {
    const parser = DocxParser.createNewDocument('blank');
    const clean = parser.validateZipIntegrity();
    expect(clean.isValid).toBe(true);
    expect(clean.errors).toEqual([]);
    expect(clean.warnings).toEqual([]);

    // '<root attr></root>' (valueless attribute) trips xmldom's `warning`
    // callback only — no error/fatalError, no parsererror — so integrity
    // stays valid while the typed warnings channel observes the signal.
    parser.getZipAdapter().writeText('word/warning-probe.xml', '<root attr></root>');
    const gate = silence();
    let out!: ReturnType<DocxParser['validateZipIntegrity']>;
    try {
      out = parser.validateZipIntegrity();
    } finally {
      gate.restore();
    }
    expect(out.isValid).toBe(true);
    expect(out.errors).toEqual([]);
    expect(out.warnings.length).toBeGreaterThan(0);
    expect(out.warnings[0]).toMatch(/\[OOXML Integrity Guard\] XML warning \(non-blocking\):/);
    expect(out.warnings[0]).toMatch(/missed value/);
    expect(gate.calls()).toBe(0);
  });

  it('comment XML warnings ride getCommentXmlWarnings() with the dropped warn text', () => {
    const parser = DocxParser.createNewDocument('blank');
    const adapter = parser.getZipAdapter();
    bootstrapCommentParts(adapter);
    const valid = adapter.readText('word/comments.xml');
    // Valueless attribute on the root: warning-grade only (no error throw),
    // still parseable so the comment flow completes and the collector
    // observes the live signal.
    const poisoned = valid.replace(/(<w:comments)/, '$1 dummyattr');
    expect(poisoned).not.toBe(valid);
    adapter.writeText('word/comments.xml', poisoned);

    const doc = parser.getDocumentDom();
    const pEl = doc.getElementsByTagName('w:p').item(0) as unknown as Element;
    expect(pEl).toBeTruthy();

    clearCommentXmlWarnings();
    expect(getCommentXmlWarnings()).toEqual([]);
    const gate = silence();
    try {
      addComment(adapter, pEl, 'Proof Author', 'proof text');
    } finally {
      gate.restore();
    }
    const collected = getCommentXmlWarnings();
    expect(collected.length).toBeGreaterThan(0);
    expect(collected[0]).toMatch(/\[comments\] XML warning \(non-fatal\):/);
    expect(collected[0]).toMatch(/missed value/);
    expect(gate.calls()).toBe(0);
    clearCommentXmlWarnings();
    expect(getCommentXmlWarnings()).toEqual([]);
  });
});

describe('D3 critical/rollback detail rides typed errors[]', () => {
  it('parser-attached critical error preserves stack + rollback sentence', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const gate = silence();
    let res!: Awaited<ReturnType<OperationEngine['executeBatch']>>;
    try {
      res = await engine.executeBatch({
        operations: [
          { action: 'delete', target: 'ghost_before_fail' } as any,
          { action: 'delete', target: 'ghost_boom', strict: true } as any,
        ],
        sequential: true,
      });
    } finally {
      gate.restore();
    }
    expect(res.success).toBe(false);
    expect(res.rollbackTriggered).toBe(true);
    const critical = res.errors.find((e) => e.operationIndex === -1);
    expect(critical).toBeDefined();
    // Full diagnostic text: batch message + stack frames (was message-only)
    // + the dropped rollback console.warn sentence (same text preserved).
    expect(critical!.error).toMatch(/Batch operation 1 failed/);
    expect(critical!.error).toMatch(/Rollback triggered\. AST restored to previous snapshot\./);
    expect(critical!.error).toMatch(/\bat \b/);
    expect(gate.calls()).toBe(0);
  });

  it('no-parser critical error additionally preserves the legacy rollback detail', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast);
    const gate = silence();
    let res!: Awaited<ReturnType<OperationEngine['executeBatch']>>;
    try {
      res = await engine.executeBatch({
        operations: [
          { action: 'delete', target: 'ghost_legacy', strict: true } as any,
        ],
        sequential: true,
      });
    } finally {
      gate.restore();
    }
    expect(res.success).toBe(false);
    expect(res.rollbackTriggered).toBe(true);
    const critical = res.errors.find((e) => e.operationIndex === -1);
    expect(critical).toBeDefined();
    expect(critical!.error).toMatch(/Rollback triggered\. AST restored to previous snapshot\./);
    // Same text as the dropped no-parser console.warn.
    expect(critical!.error).toMatch(/no parser attached — using legacy AST-only rollback/);
    expect(gate.calls()).toBe(0);
  });
});
