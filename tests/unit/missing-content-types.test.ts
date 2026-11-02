import { describe, it, expect, vi, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleOpenDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';

/**
 * A package without [Content_Types].xml is not an Open Packaging Convention package: no part has a type.
 * It is refused at open, in the same way as a package without word/document.xml, instead of being edited
 * and saved still broken.
 */
function withoutPart(name: string): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  expect(zip.getEntry(name)).not.toBeNull();
  zip.deleteFile(name);
  return zip.toBuffer();
}

describe('a package without [Content_Types].xml', () => {
  it('is refused at open with a message that names the part', () => {
    expect(() => new DocxParser(withoutPart('[Content_Types].xml'))).toThrow(
      'Invalid DOCX: missing [Content_Types].xml',
    );
  });

  it('carries the same error code as a missing word/document.xml', () => {
    let content: unknown;
    let main: unknown;
    try {
      new DocxParser(withoutPart('[Content_Types].xml'));
    } catch (e) {
      content = e;
    }
    try {
      new DocxParser(withoutPart('word/document.xml'));
    } catch (e) {
      main = e;
    }
    expect((content as { code?: string }).code).toBe('E_VALIDATION');
    expect((content as { code?: string }).code).toBe((main as { code?: string }).code);
    expect(String((main as Error).message)).toContain('missing word/document.xml');
  });

  it('still reports INVALID from the validator, which does not open the package for editing', async () => {
    const result = await validateDocxBuffer(withoutPart('[Content_Types].xml'));
    expect(result.ok).toBe(false);
  });

  it('a package that has the part opens as before', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    expect(ast.sections.length).toBeGreaterThan(0);
  });

  describe('through the command line', () => {
    const dirs: string[] = [];
    afterEach(() => {
      vi.restoreAllMocks();
      for (const d of dirs.splice(0)) fs.rmSync(d, { recursive: true, force: true });
    });

    function setup(): { dir: string; file: string } {
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-noct-'));
      dirs.push(dir);
      const file = path.join(dir, 'broken.docx');
      fs.writeFileSync(file, withoutPart('[Content_Types].xml'));
      return { dir, file };
    }

    function capture(): { lines: () => string } {
      const seen: string[] = [];
      vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
      vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
      return { lines: () => seen.join('\n') };
    }

    it('inspect refuses and names the part', async () => {
      const { file } = setup();
      const c = capture();
      expect(await run(['inspect', file])).toBe(1);
      expect(c.lines()).toContain('missing [Content_Types].xml');
    });

    it('batch refuses and writes nothing', async () => {
      const { dir, file } = setup();
      const ops = path.join(dir, 'ops.json');
      fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'x' }] }));
      const out = path.join(dir, 'out.docx');
      const c = capture();
      expect(await run(['batch', file, ops, '--out', out])).toBe(1);
      expect(c.lines()).toContain('missing [Content_Types].xml');
      expect(fs.existsSync(out)).toBe(false);
      expect(fs.readFileSync(file).equals(withoutPart('[Content_Types].xml'))).toBe(true);
    });

    it('the MCP open_document refuses and creates no session', async () => {
      const { file } = setup();
      const store = createMcpSessionStore();
      const res = (await handleOpenDocument(store, { path: file })) as { ok: boolean; error?: string; code?: string };
      expect(res.ok).toBe(false);
      expect(JSON.stringify(res)).toContain('missing [Content_Types].xml');
    });

    it('validate still reports the file as INVALID', async () => {
      const { file } = setup();
      const c = capture();
      expect(await run(['validate', file])).toBe(1);
      expect(c.lines()).toContain('INVALID');
    });
  });
});
