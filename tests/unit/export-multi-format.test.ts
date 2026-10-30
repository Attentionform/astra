import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import {
  DocxParser,
  convertAstToHtml,
  convertAstToMarkdown,
  convertDocxToPdf,
  hasPdfBackend,
  setPdfBackendBinaryForTesting,
  resetPdfBackendCache,
} from '../../packages/engine/src';
import { run as runCli, EXIT_OK, EXIT_FAIL, EXIT_USAGE } from '../../packages/cli/src/cli';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleOpenDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { handleExportDocument } from '../../packages/mcp-server/src/tools/export';
import { withLoSlot } from '../helpers/lo-slot';

// LO-backed `it`s below take the suite-wide LO slot and carry a 300s budget:
// one quiet conversion costs 13-18s here, 20-109s with several running at
// once, and the slot queue can hold a test behind other files' conversions.
// The slot removes the contention; the budget covers queue + own run.

const GALLERY_DOC = path.resolve(__dirname, '../../test-gallery/T1_coffee_onepager.docx');

describe('Sprint 6 Track A: Multi-Format Headless Exports', () => {
  let tempDir: string;

  beforeAll(async () => {
    tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-export-test-'));
  });

  afterAll(async () => {
    resetPdfBackendCache();
    try {
      await fsp.rm(tempDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  });

  describe('Engine Exports', () => {
    it('detects LibreOffice PDF backend on host system', () => {
      const available = hasPdfBackend();
      expect(available).toBe(true);
    });

    it('exports DocumentAST to valid HTML with DOCTYPE and styling', async () => {
      const buf = await fsp.readFile(GALLERY_DOC);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const html = convertAstToHtml(ast, 'Kaffa Coffee');

      expect(html).toContain('<!DOCTYPE html>');
      expect(html).toContain('Kaffa Coffee');
      expect(html).toContain('Why Kaffa');
      expect(html).toContain('<table');
      expect(html).toContain('Discovery');
    });

    it('exports DocumentAST to valid Markdown with headings and tables', async () => {
      const buf = await fsp.readFile(GALLERY_DOC);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const md = convertAstToMarkdown(ast);

      expect(md).toContain('Why Kaffa');
      expect(md).toContain('Discovery');
      expect(md).toContain('Roastery Reserve');
    });

    it('exports docx buffer to valid PDF using headless LibreOffice', async () => {
      const buf = await fsp.readFile(GALLERY_DOC);
      const outPdf = path.join(tempDir, 'coffee.pdf');
      const resPath = await withLoSlot(() => convertDocxToPdf(buf, outPdf));

      expect(resPath).toBe(outPdf);
      expect(fs.existsSync(outPdf)).toBe(true);
      const stat = await fsp.stat(outPdf);
      expect(stat.size).toBeGreaterThan(1000);

      // Verify PDF magic header %PDF-
      const headerBuf = Buffer.alloc(5);
      const fd = await fsp.open(outPdf, 'r');
      await fd.read(headerBuf, 0, 5, 0);
      await fd.close();
      expect(headerBuf.toString('ascii')).toBe('%PDF-');
    }, 300000);

    it('throws DependencyError with E_DEPENDENCY when PDF backend is missing', async () => {
      setPdfBackendBinaryForTesting(null);
      try {
        const buf = await fsp.readFile(GALLERY_DOC);
        const outPdf = path.join(tempDir, 'fail.pdf');
        await expect(convertDocxToPdf(buf, outPdf)).rejects.toThrow(/\[E_DEPENDENCY\]/);
      } finally {
        resetPdfBackendCache();
      }
    });
  });

  describe('CLI Export Subcommand (astra export)', () => {
    it('displays help message with --help and exits 0', async () => {
      const code = await runCli(['export', '--help']);
      expect(code).toBe(EXIT_OK);
    });

    it('exports docx to HTML via CLI', async () => {
      const outHtml = path.join(tempDir, 'cli_coffee.html');
      const code = await runCli(['export', GALLERY_DOC, '--format', 'html', '-o', outHtml]);
      expect(code).toBe(EXIT_OK);
      expect(fs.existsSync(outHtml)).toBe(true);
      const content = await fsp.readFile(outHtml, 'utf-8');
      expect(content).toContain('Why Kaffa');
    });

    it('exports docx to Markdown via CLI', async () => {
      const outMd = path.join(tempDir, 'cli_coffee.md');
      const code = await runCli(['export', GALLERY_DOC, '--format', 'markdown', '-o', outMd]);
      expect(code).toBe(EXIT_OK);
      expect(fs.existsSync(outMd)).toBe(true);
      const content = await fsp.readFile(outMd, 'utf-8');
      expect(content).toContain('Why Kaffa');
    });

    it('exports docx to PDF via CLI', async () => {
      const outPdf = path.join(tempDir, 'cli_coffee.pdf');
      const code = await withLoSlot(() => runCli(['export', GALLERY_DOC, '--format', 'pdf', '-o', outPdf]));
      expect(code).toBe(EXIT_OK);
      expect(fs.existsSync(outPdf)).toBe(true);
      const stat = await fsp.stat(outPdf);
      expect(stat.size).toBeGreaterThan(1000);
    }, 300000);

    it('fails closed on non-existent document with exit code 1', async () => {
      const code = await runCli(['export', path.join(tempDir, 'non_existent.docx')]);
      expect(code).toBe(EXIT_FAIL);
    });

    it('fails closed on unsupported format with exit code 2 (usage)', async () => {
      const code = await runCli(['export', GALLERY_DOC, '--format', 'invalid_format']);
      expect(code).toBe(EXIT_USAGE);
    });
  });

  describe('MCP Server export_document Tool', () => {
    let mcpDoc: string;

    beforeAll(async () => {
      mcpDoc = path.join(tempDir, 'coffee_mcp.docx');
      await fsp.copyFile(GALLERY_DOC, mcpDoc);
    });

    it('exports open document session to HTML, Markdown, and PDF', async () => {
      const store = createMcpSessionStore();
      const openRes = await handleOpenDocument(store, { path: mcpDoc });
      expect(openRes.ok).toBe(true);
      const sessionId = (openRes as { sessionId: string }).sessionId;

      // 1. Export HTML
      const htmlRes = await handleExportDocument(store, { sessionId, format: 'html' });
      expect(htmlRes.ok).toBe(true);
      expect(htmlRes.content).toContain('Why Kaffa');

      // 2. Export Markdown
      const mdRes = await handleExportDocument(store, { sessionId, format: 'markdown' });
      expect(mdRes.ok).toBe(true);
      expect(mdRes.content).toContain('Why Kaffa');

      // 3. Export PDF
      const outPdf = path.join(tempDir, 'mcp_export.pdf');
      const pdfRes = await withLoSlot(() => handleExportDocument(store, { sessionId, format: 'pdf', outPath: outPdf }));
      expect(pdfRes.ok).toBe(true);
      expect(pdfRes.path).toBe(outPdf);
      expect(fs.existsSync(outPdf)).toBe(true);
      expect(pdfRes.bytes).toBeGreaterThan(1000);
    }, 300000);

    it('enforces sandbox jail and rejects outPath traversing outside allowed directory', async () => {
      const store = createMcpSessionStore();
      const openRes = await handleOpenDocument(store, { path: mcpDoc });
      expect(openRes.ok).toBe(true);
      const sessionId = (openRes as { sessionId: string }).sessionId;

      const escapedRes = await handleExportDocument(store, {
        sessionId,
        format: 'html',
        outPath: '/etc/passwd.html',
      });
      expect(escapedRes.ok).toBe(false);
      expect(escapedRes.code).toBe('E_SANDBOX');
    });

    it('fails closed with E_SESSION for non-existent session ID', async () => {
      const store = createMcpSessionStore();
      const res = await handleExportDocument(store, { sessionId: 'invalid-sess', format: 'html' });
      expect(res.ok).toBe(false);
      expect(res.code).toBe('E_SESSION');
    });

    it('supports format alias "md" normalized to markdown', async () => {
      const store = createMcpSessionStore();
      const openRes = await handleOpenDocument(store, { path: mcpDoc });
      expect(openRes.ok).toBe(true);
      const sessionId = (openRes as { sessionId: string }).sessionId;

      const res = await handleExportDocument(store, { sessionId, format: 'md' });
      expect(res.ok).toBe(true);
      expect(res.format).toBe('markdown');
      expect(res.content).toContain('Why Kaffa');
    });

    it('returns E_SANDBOX and never smears to E_IO if default PDF path violates sandbox', async () => {
      const store = createMcpSessionStore();
      const openRes = await handleOpenDocument(store, { path: mcpDoc });
      expect(openRes.ok).toBe(true);
      const sessionId = (openRes as { sessionId: string }).sessionId;
      const session = store.get(sessionId);
      const originalPath = session.filePath;
      try {
        session.filePath = '/etc/secret.docx';
        const res = await handleExportDocument(store, { sessionId, format: 'pdf' });
        expect(res.ok).toBe(false);
        expect(res.code).toBe('E_SANDBOX');
      } finally {
        session.filePath = originalPath;
      }
    });
  });
});
