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

    it('refuses file mode when no output path is supplied', async () => {
      const res = await handleExportDocument(store, { sessionId, format: 'markdown', imageMode: 'file' });
      expect(res.ok).toBe(false);
      expect(res.code).toBe('E_VALIDATION');
      expect(res.remedy).toContain('outPath');
    });
  });

  describe('Disk Markdown export with outPath', () => {
    it('cleans unpublished assets when final rename fails', async () => {
      const outPath = path.join(tempDir, 'directory.md');
      await fsp.mkdir(outPath);
      const res = await handleExportDocument(store, { sessionId, format: 'markdown', outPath, assetPrefix: 'rename_failure_assets' });
      expect(res.ok).toBe(false);
      expect((await fsp.stat(outPath)).isDirectory()).toBe(true);
      expect(await fsp.readdir(path.join(tempDir, 'rename_failure_assets'))).toEqual([]);
    });

    it('keeps image assets separate when exports share a prefix', async () => {
      const firstPath = path.join(tempDir, 'first.md');
      const secondPath = path.join(tempDir, 'second.md');
      const args = { sessionId, format: 'markdown', assetPrefix: 'shared_figures' };
      const first = await handleExportDocument(store, { ...args, outPath: firstPath });
      const second = await handleExportDocument(store, { ...args, outPath: secondPath });
      expect(first.ok).toBe(true);
      expect(second.ok).toBe(true);
      expect(first.assets![0]).not.toBe(second.assets![0]);
      expect(fs.existsSync(first.assets![0])).toBe(true);
      expect(fs.existsSync(second.assets![0])).toBe(true);
      const firstMd = await fsp.readFile(firstPath, 'utf8');
      expect(firstMd).toContain(path.relative(tempDir, first.assets![0]).split(path.sep).join('/'));
    });

    it('leaves existing Markdown unchanged when an asset directory is refused', async () => {
      const outMd = path.join(tempDir, 'existing.md');
      await fsp.writeFile(outMd, 'previous document');
      const outside = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-outside-'));
      const link = path.join(tempDir, 'unsafe_assets');
      try {
        await fsp.symlink(outside, link, 'dir');
        const res = await handleExportDocument(store, { sessionId, format: 'markdown', outPath: outMd, assetPrefix: 'unsafe_assets' });
        expect(res.ok).toBe(false);
        expect(await fsp.readFile(outMd, 'utf8')).toBe('previous document');
      } finally {
        await fsp.unlink(link);
        await fsp.rm(outside, { recursive: true, force: true });
      }
    });

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
      expect(mdContent).toMatch(/research_report_assets\/export-[^/]+\/image-0001\./);
      expect(mdContent).not.toContain('data:image/');

      // Verify the extracted image file exists on disk and has non-zero size
      const expectedAssetPath = res.assets![0];
      expect(expectedAssetPath).toMatch(/research_report_assets\/export-[^/]+\/image-0001\.png$/);
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
      expect(mdContent).toMatch(/extracted_figures\/export-[^/]+\/image-0001\./);

      const expectedAssetPath = res.assets![0];
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
