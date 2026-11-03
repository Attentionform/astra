import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { handleOpenDocument } from '../../src/lib/mcp-tool-adapter';
import {
  handleExportDocument,
  validateAssetTarget,
} from '../../src/tools/export';

const IMAGE_DOC = path.resolve(__dirname, '../../../../test-gallery/T5_image_doc.docx');

describe('MCP export_document: Markdown image asset export mode', () => {
  let tempDir: string;
  let store: ReturnType<typeof createMcpSessionStore>;
  let sessionId: string;

  beforeAll(async () => {
    tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-mcp-asset-export-'));
    store = createMcpSessionStore();

    // Copy IMAGE_DOC into tempDir to operate within allowed MCP root
    const testDoc = path.join(tempDir, 'sample_image.docx');
    await fsp.copyFile(IMAGE_DOC, testDoc);

    const openRes = await handleOpenDocument(store, { path: testDoc });
    expect(openRes.ok).toBe(true);
    sessionId = (openRes as { sessionId: string }).sessionId;
  });

  afterAll(async () => {
    try {
      await fsp.rm(tempDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  });

  describe('validateAssetTarget sandbox security', () => {
    it('rejects empty or non-string paths with E_SANDBOX', () => {
      expect(() => validateAssetTarget('')).toThrow(/E_SANDBOX/);
      expect(() => validateAssetTarget('   ')).toThrow(/E_SANDBOX/);
      expect(() => validateAssetTarget(null as unknown as string)).toThrow(/E_SANDBOX/);
    });

    it('rejects null bytes and line breaks with E_SANDBOX', () => {
      expect(() => validateAssetTarget('/tmp/test\0file.png')).toThrow(/null byte/);
      expect(() => validateAssetTarget('/tmp/test%00file.png')).toThrow(/null byte/);
      expect(() => validateAssetTarget('/tmp/test\nfile.png')).toThrow(/line break/);
      expect(() => validateAssetTarget('/tmp/test\rfile.png')).toThrow(/line break/);
    });

    it('rejects path traversal (..) with E_SANDBOX', () => {
      expect(() => validateAssetTarget(`${tempDir}/../outside.png`)).toThrow(/\.\./);
    });

    it('rejects paths outside allowed roots with E_SANDBOX', () => {
      expect(() => validateAssetTarget('/etc/shadow.png')).toThrow(/outside allowed directories/);
    });

    it('accepts valid path inside allowed directory', () => {
      const valid = path.join(tempDir, 'assets', 'image-0001.png');
      const resolved = validateAssetTarget(valid);
      expect(resolved).toBe(path.resolve(valid));
    });
  });

  describe('Format restrictions on imageMode and assetPrefix', () => {
    it('refuses imageMode on HTML export with E_VALIDATION', async () => {
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'html',
        imageMode: 'file',
      });
      expect(res.ok).toBe(false);
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('imageMode applies to markdown format only');
    });

    it('refuses assetPrefix on HTML export with E_VALIDATION', async () => {
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'html',
        assetPrefix: 'my_assets',
      });
      expect(res.ok).toBe(false);
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('assetPrefix applies to markdown format only');
    });

    it('refuses imageMode on PDF export with E_VALIDATION', async () => {
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'pdf',
        imageMode: 'file',
      });
      expect(res.ok).toBe(false);
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('imageMode applies to markdown format only');
    });
  });

  describe('In-memory Markdown export (no outPath)', () => {
    it('defaults to imageMode: embed for in-memory exports', async () => {
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'markdown',
      });
      expect(res.ok).toBe(true);
      expect(res.format).toBe('markdown');
      expect(res.imageMode).toBe('embed');
      expect(res.assetCount).toBe(0);
      expect(res.assets).toEqual([]);
      expect(res.content).toContain('data:image/');
      expect(res.content).toContain(';base64,');
    });

    it('supports explicit imageMode: file for in-memory exports with relative links', async () => {
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'markdown',
        imageMode: 'file',
      });
      expect(res.ok).toBe(true);
      expect(res.format).toBe('markdown');
      expect(res.imageMode).toBe('file');
      expect(res.assetCount).toBeGreaterThan(0);
      expect(res.assets).toEqual([]);
      expect(res.content).not.toContain('data:image/');
      expect(res.content).toMatch(/!\[.*\]\(.*_assets\/image-0001\./);
    });
  });

  describe('Disk Markdown export with outPath', () => {
    it('defaults to imageMode: file and extracts images into sibling <doc>_assets/ directory', async () => {
      const outMd = path.join(tempDir, 'research_report.md');
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'markdown',
        outPath: outMd,
      });

      expect(res.ok).toBe(true);
      expect(res.format).toBe('markdown');
      expect(res.path).toBe(outMd);
      expect(res.imageMode).toBe('file');
      expect(res.assetPrefix).toBe('research_report_assets');
      expect(res.assetCount).toBe(1);
      expect(res.assets).toHaveLength(1);

      // Verify the markdown file contains relative link to sibling folder
      const mdContent = await fsp.readFile(outMd, 'utf-8');
      expect(mdContent).toContain('research_report_assets/image-0001.');
      expect(mdContent).not.toContain('data:image/');

      // Verify the extracted image file exists on disk and has non-zero size
      const expectedAssetPath = path.join(tempDir, 'research_report_assets', 'image-0001.png');
      expect(res.assets![0]).toBe(expectedAssetPath);
      expect(fs.existsSync(expectedAssetPath)).toBe(true);
      const stat = await fsp.stat(expectedAssetPath);
      expect(stat.size).toBeGreaterThan(0);
    });

    it('honors custom assetPrefix parameter', async () => {
      const outMd = path.join(tempDir, 'custom_prefixed.md');
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'markdown',
        outPath: outMd,
        assetPrefix: 'extracted_figures',
      });

      expect(res.ok).toBe(true);
      expect(res.imageMode).toBe('file');
      expect(res.assetPrefix).toBe('extracted_figures');
      expect(res.assetCount).toBe(1);

      const mdContent = await fsp.readFile(outMd, 'utf-8');
      expect(mdContent).toContain('extracted_figures/image-0001.');

      const expectedAssetPath = path.join(tempDir, 'extracted_figures', 'image-0001.png');
      expect(res.assets![0]).toBe(expectedAssetPath);
      expect(fs.existsSync(expectedAssetPath)).toBe(true);
    });

    it('supports explicit imageMode: embed with outPath for single-file bundle', async () => {
      const outMd = path.join(tempDir, 'bundled.md');
      const res = await handleExportDocument(store, {
        sessionId,
        format: 'markdown',
        outPath: outMd,
        imageMode: 'embed',
      });

      expect(res.ok).toBe(true);
      expect(res.imageMode).toBe('embed');
      expect(res.assetCount).toBe(0);
      expect(res.assets).toEqual([]);

      const mdContent = await fsp.readFile(outMd, 'utf-8');
      expect(mdContent).toContain('data:image/');
      expect(mdContent).toContain(';base64,');

      // Ensure no asset directory was created
      expect(fs.existsSync(path.join(tempDir, 'bundled_assets'))).toBe(false);
    });

    it('handles documents with no images cleanly in file mode', async () => {
      // Create a clean blank document with no images
      const blankDoc = path.join(tempDir, 'blank_doc.docx');
      await fsp.writeFile(blankDoc, createBlankDocxBuffer('blank'));
      const blankOpen = await handleOpenDocument(store, { path: blankDoc });
      expect(blankOpen.ok).toBe(true);
      const blankSessionId = (blankOpen as { sessionId: string }).sessionId;

      const outMd = path.join(tempDir, 'blank_export.md');
      const res = await handleExportDocument(store, {
        sessionId: blankSessionId,
        format: 'markdown',
        outPath: outMd,
      });

      expect(res.ok).toBe(true);
      expect(res.imageMode).toBe('file');
      expect(res.assetCount).toBe(0);
      expect(res.assets).toEqual([]);
      expect(fs.existsSync(outMd)).toBe(true);
      expect(fs.existsSync(path.join(tempDir, 'blank_export_assets'))).toBe(false);
    });
  });
});
