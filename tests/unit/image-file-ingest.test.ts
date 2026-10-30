import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  executeImage,
  validateOperationStrict,
  validateOperationsStrict,
  IMAGE_FILE_MAX_BYTES,
  type DocumentAST,
} from '@attentionform/astra-engine';
import { DOMParser } from '@xmldom/xmldom';
import { useReproducibleOutput } from '../helpers/reproducible';

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const TINY_PNG = Buffer.from(TINY_PNG_BASE64, 'base64');

function createMockAST(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
    'text/xml'
  );
  const bodyP = doc.getElementsByTagName('w:p')[0];
  return {
    metadata: {},
    sections: [{
      id: 'sec_root',
      heading: null,
      nodes: [{
        id: 'p_001',
        type: 'paragraph',
        xml: bodyP,
        metadata: {
          text: 'Intro paragraph',
          hasCitation: false,
          hasComment: false,
          hasBookmark: false,
          hasTrackedChanges: false,
          hasEquation: false,
          hasField: false,
          hasImage: false,
          hasHyperlink: false,
        },
        version: 1,
        checksum: 'chk_1',
      }],
      pageSetup: {},
      version: 1,
      checksum: 'sec_chk_1',
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1,
  };
}

function assetDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-imgfile-'));
  fs.writeFileSync(path.join(dir, 'tiny.png'), TINY_PNG);
  return dir;
}

function codeOf(err: unknown): string | undefined {
  return (err as { code?: unknown } | null)?.code as string | undefined;
}

async function blankEngine() {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}

function packageOf(parser: DocxParser) {
  const zip = new AdmZip(parser.rezip());
  const doc = zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
  const rels = zip.getEntry('word/_rels/document.xml.rels')?.getData().toString('utf8') ?? '';
  const mediaNames = zip.getEntries().map((e) => e.entryName).filter((n) => n.startsWith('word/media/')).sort();
  const mediaBytes = mediaNames.map((n) => zip.getEntry(n)?.getData());
  return { doc, rels, mediaNames, mediaBytes };
}

useReproducibleOutput();

describe('image filePath schema + strict validators', () => {
  it('accepts a filePath insert (fileData stays optional)', () => {
    const res = validateOperationStrict({
      action: 'image', operation: 'insert', anchor: 'p_001', filePath: 'tiny.png',
    });
    expect(res.ok).toBe(true);
  });

  it('rejects both fileData and filePath with E_VALIDATION naming the conflict', () => {
    const op = {
      action: 'image', operation: 'insert', anchor: 'p_001',
      fileData: `data:image/png;base64,${TINY_PNG_BASE64}`, filePath: 'tiny.png',
    };
    expect(() => validateOperationStrict(op)).toThrow(/both fileData and filePath/);
    try {
      validateOperationStrict(op);
      expect.unreachable();
    } catch (err: unknown) {
      expect(codeOf(err)).toBe('E_VALIDATION');
      expect(String((err as Error).message)).toContain('Remedy:');
    }
    expect(() => validateOperationsStrict([op])).toThrow(/both fileData and filePath/);
    try {
      validateOperationsStrict([op]);
      expect.unreachable();
    } catch (err: unknown) {
      expect(String((err as Error).message)).toContain('E_VALIDATION');
    }
  });

  it('rejects an insert with neither source with E_VALIDATION', () => {
    const op = { action: 'image', operation: 'insert', anchor: 'p_001' };
    expect(() => validateOperationStrict(op)).toThrow(/neither fileData nor filePath/);
    expect(() => validateOperationsStrict([op])).toThrow(/neither fileData nor filePath/);
  });

  it('leaves the other six image operations untouched (filePath ignored)', () => {
    const res = validateOperationStrict({ action: 'image', operation: 'resize', target: 'img_1' });
    expect(res.ok).toBe(true);
  });

  it('keeps the strict shape (unknown keys still rejected)', () => {
    expect(() =>
      validateOperationStrict({ action: 'image', operation: 'insert', anchor: 'p_001', filePath: 'a.png', bogus: 1 })
    ).toThrow(/E_VALIDATION/);
  });
});

describe('executeImage filePath entry', () => {
  it('inserts from inside the jail with wrap/size/alt honored and no warning', () => {
    const dir = assetDir();
    const ast = createMockAST();
    const out = executeImage(
      {
        action: 'image', operation: 'insert', anchor: 'p_001', filePath: 'tiny.png',
        width: 914400, height: 457200, wrapMode: 'square', altText: 'Tiny dot',
      },
      ast, undefined, dir
    );
    expect(out).toBeUndefined();
    expect(ast.sections[0].nodes.length).toBe(2);
    const img = ast.sections[0].nodes[1];
    expect(img.type).toBe('image');
    expect(img.metadata.hasImage).toBe(true);
    const mediaKeys = Object.keys(ast.mediaFiles ?? {});
    expect(mediaKeys.length).toBe(1);
    expect((ast.mediaFiles as Record<string, Buffer>)[mediaKeys[0]].equals(TINY_PNG)).toBe(true);
  });

  it('refuses outside-jail refs with E_SANDBOX + remedy', () => {
    const dir = assetDir();
    const ast = createMockAST();
    try {
      executeImage(
        { action: 'image', operation: 'insert', anchor: 'p_001', filePath: '../escape.png' },
        ast, undefined, dir
      );
      expect.unreachable();
    } catch (err: unknown) {
      expect(codeOf(err)).toBe('E_SANDBOX');
      expect(String((err as Error).message)).toContain('Remedy:');
    }
    expect(ast.sections[0].nodes.length).toBe(1);
  });

  it('refuses missing files with E_IO', () => {
    const dir = assetDir();
    const ast = createMockAST();
    try {
      executeImage(
        { action: 'image', operation: 'insert', anchor: 'p_001', filePath: 'nope.png' },
        ast, undefined, dir
      );
      expect.unreachable();
    } catch (err: unknown) {
      expect(codeOf(err)).toBe('E_IO');
    }
  });

  it('refuses oversize files with E_LIMIT', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-imgbig-'));
    fs.writeFileSync(path.join(dir, 'big.png'), Buffer.alloc(IMAGE_FILE_MAX_BYTES + 1));
    const ast = createMockAST();
    try {
      executeImage(
        { action: 'image', operation: 'insert', anchor: 'p_001', filePath: 'big.png' },
        ast, undefined, dir
      );
      expect.unreachable();
    } catch (err: unknown) {
      expect(codeOf(err)).toBe('E_LIMIT');
    }
  });

  it('fails closed with E_SANDBOX when no allowedDir is configured', () => {
    const ast = createMockAST();
    try {
      executeImage(
        { action: 'image', operation: 'insert', anchor: 'p_001', filePath: 'tiny.png' },
        ast
      );
      expect.unreachable();
    } catch (err: unknown) {
      expect(codeOf(err)).toBe('E_SANDBOX');
      expect(String((err as Error).message)).toContain('allowedDir');
    }
  });

  it('rejects both sources at entry with E_VALIDATION', () => {
    const dir = assetDir();
    const ast = createMockAST();
    try {
      executeImage(
        {
          action: 'image', operation: 'insert', anchor: 'p_001',
          fileData: `data:image/png;base64,${TINY_PNG_BASE64}`, filePath: 'tiny.png',
        },
        ast, undefined, dir
      );
      expect.unreachable();
    } catch (err: unknown) {
      expect(codeOf(err)).toBe('E_VALIDATION');
    }
  });

  it('keeps the fileData path byte-identical (same bytes both ways)', () => {
    const dir = assetDir();
    const viaData = createMockAST();
    executeImage(
      { action: 'image', operation: 'insert', anchor: 'p_001', fileData: `data:image/png;base64,${TINY_PNG_BASE64}`, width: 914400, height: 457200, altText: 'Tiny dot' },
      viaData
    );
    const viaPath = createMockAST();
    executeImage(
      { action: 'image', operation: 'insert', anchor: 'p_001', filePath: 'tiny.png', width: 914400, height: 457200, altText: 'Tiny dot' },
      viaPath, undefined, dir
    );
    expect(viaPath.sections[0].nodes[1].id).toBe(viaData.sections[0].nodes[1].id);
    const dataKeys = Object.keys(viaData.mediaFiles ?? {});
    const pathKeys = Object.keys(viaPath.mediaFiles ?? {});
    expect(pathKeys).toEqual(dataKeys);
    expect(
      (viaPath.mediaFiles as Record<string, Buffer>)[pathKeys[0]].equals(
        (viaData.mediaFiles as Record<string, Buffer>)[dataKeys[0]]
      )
    ).toBe(true);
  });
});

describe('engine batch allowedDir plumbing', () => {
  it('inserts by path through per-batch allowedDir with rendered bytes intact', async () => {
    const dir = assetDir();
    const { parser, ast, engine } = await blankEngine();
    const anchor = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [{
        action: 'image', operation: 'insert', anchor, filePath: 'tiny.png',
        width: 914400, height: 457200, altText: 'Tiny dot',
      }],
      sequential: true,
      allowedDir: dir,
    });
    expect(res.success).toBe(true);
    expect(res.appliedCount).toBe(1);
    const pkg = packageOf(parser);
    expect(pkg.doc).toContain('Tiny dot');
    expect(pkg.mediaNames.length).toBe(1);
    expect(pkg.mediaBytes[0]?.equals(TINY_PNG)).toBe(true);
  });

  it('fileData-vs-filePath equivalence: identical document XML and media', async () => {
    const dir = assetDir();
    const first = await blankEngine();
    const anchorA = first.ast.sections[0].nodes[0].id;
    await first.engine.executeBatch({
      operations: [{
        action: 'image', operation: 'insert', anchor: anchorA,
        fileData: `data:image/png;base64,${TINY_PNG_BASE64}`, width: 914400, height: 457200, altText: 'Tiny dot',
      }],
    });
    const second = await blankEngine();
    const anchorB = second.ast.sections[0].nodes[0].id;
    await second.engine.executeBatch({
      operations: [{
        action: 'image', operation: 'insert', anchor: anchorB, filePath: 'tiny.png',
        width: 914400, height: 457200, altText: 'Tiny dot',
      }],
      allowedDir: dir,
    });
    const a = packageOf(first.parser);
    const b = packageOf(second.parser);
    expect(b.doc).toBe(a.doc);
    expect(b.mediaNames).toEqual(a.mediaNames);
    expect(b.mediaBytes[0]?.equals(a.mediaBytes[0] as Buffer)).toBe(true);
  });

  it('honors constructor OperationEngineOptions allowedDir', async () => {
    const dir = assetDir();
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser, { allowedDir: dir });
    const anchor = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [{ action: 'image', operation: 'insert', anchor, filePath: 'tiny.png' }],
    });
    expect(res.success).toBe(true);
  });

  it('fails closed without any jail (E_SANDBOX)', async () => {
    const { ast, engine } = await blankEngine();
    const anchor = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [{ action: 'image', operation: 'insert', anchor, filePath: 'tiny.png' }],
    });
    expect(res.success).toBe(false);
    expect(res.errors[0].error).toContain('E_SANDBOX');
  });

  it('surfaces outside/missing/oversize codes through the batch', async () => {
    const dir = assetDir();
    const big = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-imgbig-'));
    fs.writeFileSync(path.join(big, 'big.png'), Buffer.alloc(IMAGE_FILE_MAX_BYTES + 1));
    for (const [allowedDir, filePath, code] of [
      [dir, '../escape.png', 'E_SANDBOX'],
      [dir, 'nope.png', 'E_IO'],
      [big, 'big.png', 'E_LIMIT'],
    ] as Array<[string, string, string]>) {
      const { ast, engine } = await blankEngine();
      const anchor = ast.sections[0].nodes[0].id;
      const res = await engine.executeBatch({
        operations: [{ action: 'image', operation: 'insert', anchor, filePath }],
        allowedDir,
      });
      expect(res.success).toBe(false);
      expect(res.errors[0].error).toContain(code);
    }
  });

  it('batches without filePath ops never touch the filesystem (no jail needed)', async () => {
    const { ast, engine } = await blankEngine();
    const anchor = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'plain batch' }],
    });
    expect(res.success).toBe(true);
  });
});
