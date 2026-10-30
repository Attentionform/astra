import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, OperationEngine, executeImage, assertValidDocx } from '@attentionform/astra-engine';
import type { ImageOperation } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const GALLERY = path.resolve(__dirname, '../../test-gallery');

async function load(fixture: string) {
  const parser = new DocxParser(fs.readFileSync(path.join(GALLERY, fixture)));
  const ast = await parser.getAST();
  const nodes = ast.sections.flatMap((s) => s.nodes);
  return { parser, ast, nodes };
}

/** Runs fn, which must throw, and returns the error's code and message. */
function thrown(fn: () => unknown): { code?: string; message: string } {
  try {
    fn();
  } catch (e) {
    const err = e as { code?: string; message?: string };
    return { code: err.code, message: String(err.message) };
  }
  throw new Error('expected the call to throw');
}

describe('image target guards and version validation', () => {
  const opTemplates: Array<{ name: string; op: Partial<ImageOperation> }> = [
    { name: 'delete', op: { operation: 'delete' } },
    { name: 'resize', op: { operation: 'resize', width: 914400, height: 914400 } },
    { name: 'setAltText', op: { operation: 'setAltText', altText: 'x' } },
    { name: 'setCaption', op: { operation: 'setCaption', captionText: 'x' } },
    { name: 'setWrap', op: { operation: 'setWrap', wrapMode: 'square' } },
    { name: 'reposition', op: { operation: 'reposition', wrapMode: 'square' } },
  ];

  it('refuses non-image paragraph targets with E_VALIDATION and leaves document unchanged', async () => {
    for (const t of opTemplates) {
      const { parser, ast, nodes } = await load('T2_edit_target.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      expect(para).toBeDefined();

      const before = parser.rezip();
      const op: ImageOperation = {
        action: 'image',
        target: para.id,
        ...t.op,
      } as ImageOperation;

      const res = thrown(() => executeImage(op, ast, parser));
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('requires an image node');
      expect(Buffer.compare(parser.rezip(), before)).toBe(0);
    }
  });

  it('refuses non-image table targets with E_VALIDATION and leaves document unchanged', async () => {
    for (const t of opTemplates) {
      const { parser, ast, nodes } = await load('T2_edit_target.docx');
      const tbl = nodes.find((n) => n.type === 'table')!;
      expect(tbl).toBeDefined();

      const before = parser.rezip();
      const op: ImageOperation = {
        action: 'image',
        target: tbl.id,
        ...t.op,
      } as ImageOperation;

      const res = thrown(() => executeImage(op, ast, parser));
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('requires an image node');
      expect(Buffer.compare(parser.rezip(), before)).toBe(0);
    }
  });

  it('refuses non-existent target ids with E_NOT_FOUND and leaves document unchanged', async () => {
    for (const t of opTemplates) {
      const { parser, ast } = await load('T2_edit_target.docx');
      const before = parser.rezip();
      const op: ImageOperation = {
        action: 'image',
        target: 'NO_SUCH_ID',
        ...t.op,
      } as ImageOperation;

      const res = thrown(() => executeImage(op, ast, parser));
      expect(res.code).toBe('E_NOT_FOUND');
      expect(res.message).toContain("Target image node 'NO_SUCH_ID' not found");
      expect(Buffer.compare(parser.rezip(), before)).toBe(0);
    }
  });

  it('succeeds for image operations on genuine image node in T5', async () => {
    for (const t of opTemplates) {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const imgNode = nodes.find((n) => n.type === 'image')!;
      expect(imgNode).toBeDefined();

      const op: ImageOperation = {
        action: 'image',
        target: imgNode.id,
        ...t.op,
      } as ImageOperation;

      expect(() => executeImage(op, ast, parser)).not.toThrow();
    }
  });

  it('refuses operations through engine batch and rolls back', async () => {
    const { parser, ast, nodes } = await load('T2_edit_target.docx');
    const tbl = nodes.find((n) => n.type === 'table')!;
    expect(tbl).toBeDefined();

    const engine = new OperationEngine(ast, parser);
    const deleteOnTable: ImageOperation = {
      action: 'image',
      operation: 'delete',
      target: tbl.id,
    };

    const res = await engine.executeBatch({
      operations: [deleteOnTable],
      rollbackOnError: true,
    });

    expect(res.success).toBe(false);
    expect(res.errors[0]?.error).toContain('[E_VALIDATION]');
    const refreshedNodes = ast.sections.flatMap((s) => s.nodes);
    expect(refreshedNodes.some((n) => n.id === tbl.id)).toBe(true);
  });

  it('refuses stale expectedVersion with E_VALIDATION and byte-identical document', async () => {
    for (const t of opTemplates) {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const imgNode = nodes.find((n) => n.type === 'image')!;
      expect(imgNode).toBeDefined();

      const before = parser.rezip();
      const op: ImageOperation = {
        action: 'image',
        target: imgNode.id,
        expectedVersion: 9999,
        ...t.op,
      } as ImageOperation;

      const res = thrown(() => executeImage(op, ast, parser));
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('Version mismatch');
      expect(Buffer.compare(parser.rezip(), before)).toBe(0);
    }
  });

  it('accepts matching expectedVersion', async () => {
    for (const t of opTemplates) {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const imgNode = nodes.find((n) => n.type === 'image')!;
      expect(imgNode).toBeDefined();

      const op: ImageOperation = {
        action: 'image',
        target: imgNode.id,
        expectedVersion: 1,
        ...t.op,
      } as ImageOperation;

      expect(() => executeImage(op, ast, parser)).not.toThrow();
    }
  });

  it('observes version bump on mutation so second batch with same expectedVersion fails', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    expect(imgNode).toBeDefined();

    const engine = new OperationEngine(ast, parser);
    const resizeOp: ImageOperation = {
      action: 'image',
      operation: 'resize',
      target: imgNode.id,
      width: 914400,
      height: 914400,
      expectedVersion: 1,
    };

    const firstRun = await engine.executeBatch({
      operations: [resizeOp],
    });
    expect(firstRun.success).toBe(true);
    expect(imgNode.version).toBe(2);

    const secondRun = await engine.executeBatch({
      operations: [resizeOp],
    });
    expect(secondRun.success).toBe(false);
    expect(secondRun.errors[0]?.error).toContain('Version mismatch');
  });
});

describe('wrap schema ordering', () => {
  it('replaces existing wrap preserving schema order', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    executeImage({ action: 'image', operation: 'setWrap', target: imgNode.id, wrapMode: 'square' }, ast, parser);
    executeImage({ action: 'image', operation: 'setWrap', target: imgNode.id, wrapMode: 'tight' }, ast, parser);
    const anchor = (imgNode.xml as Element).getElementsByTagName('wp:anchor')[0];
    const localNames = Array.from(anchor.childNodes)
      .filter((k) => k.nodeType === 1)
      .map((k) => (k as Element).localName || (k as Element).nodeName.split(':').pop()!);
    expect(localNames.slice(0, 6)).toEqual(['simplePos', 'positionH', 'positionV', 'extent', 'wrapTight', 'docPr']);
  });

  it('inserts wrap in schema order before docPr when anchor has no wrap element', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    executeImage({ action: 'image', operation: 'setWrap', target: imgNode.id, wrapMode: 'square' }, ast, parser);
    const anchor = (imgNode.xml as Element).getElementsByTagName('wp:anchor')[0];
    const wrapSquare = anchor.getElementsByTagName('wp:wrapSquare')[0];
    expect(wrapSquare).toBeDefined();
    anchor.removeChild(wrapSquare);

    executeImage({ action: 'image', operation: 'setWrap', target: imgNode.id, wrapMode: 'tight' }, ast, parser);
    const wrapTight = anchor.getElementsByTagName('wp:wrapTight')[0];
    expect(wrapTight).toBeDefined();
    let nextEl = wrapTight.nextSibling;
    while (nextEl && nextEl.nodeType !== 1) {
      nextEl = nextEl.nextSibling;
    }
    const nextLocal = (nextEl as Element)?.localName || (nextEl as Element)?.nodeName?.split(':').pop();
    expect(nextLocal).toBe('docPr');

    await expect(assertValidDocx(parser.rezip())).resolves.not.toThrow();
  });

  it('sets behindDoc attribute for behindText mode', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    executeImage({ action: 'image', operation: 'setWrap', target: imgNode.id, wrapMode: 'behindText' }, ast, parser);
    const anchor = (imgNode.xml as Element).getElementsByTagName('wp:anchor')[0];
    expect(anchor.getAttribute('behindDoc')).toBe('1');
  });
});

describe('width-only and height-only resize', () => {
  it('infers missing height from aspect ratio when only width is provided', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    executeImage({ action: 'image', operation: 'resize', target: imgNode.id, width: 1828800, height: 914400 }, ast, parser);
    executeImage({ action: 'image', operation: 'resize', target: imgNode.id, width: 3657600 }, ast, parser);
    const ext = (imgNode.xml as Element).getElementsByTagName('wp:extent')[0];
    expect(ext.getAttribute('cx')).toBe('3657600');
    expect(ext.getAttribute('cy')).toBe('1828800');
  });

  it('infers missing width from aspect ratio when only height is provided', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    executeImage({ action: 'image', operation: 'resize', target: imgNode.id, width: 1828800, height: 914400 }, ast, parser);
    executeImage({ action: 'image', operation: 'resize', target: imgNode.id, height: 457200 }, ast, parser);
    const ext = (imgNode.xml as Element).getElementsByTagName('wp:extent')[0];
    expect(ext.getAttribute('cx')).toBe('914400');
    expect(ext.getAttribute('cy')).toBe('457200');
  });

  it('sets both dimensions directly when both are provided', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    executeImage({ action: 'image', operation: 'resize', target: imgNode.id, width: 1828800, height: 914400 }, ast, parser);
    const ext = (imgNode.xml as Element).getElementsByTagName('wp:extent')[0];
    expect(ext.getAttribute('cx')).toBe('1828800');
    expect(ext.getAttribute('cy')).toBe('914400');
  });

  it('refuses resize with neither width nor height and leaves document unchanged', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    const before = parser.rezip();
    const res = thrown(() => executeImage({ action: 'image', operation: 'resize', target: imgNode.id } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('requires width and/or height');
    expect(Buffer.compare(parser.rezip(), before)).toBe(0);
  });

  it('refuses dimension inference when existing image extent is unreadable or zero', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    const ext = (imgNode.xml as Element).getElementsByTagName('wp:extent')[0];
    ext.setAttribute('cx', '0');
    const res = thrown(() => executeImage({ action: 'image', operation: 'resize', target: imgNode.id, width: 3657600 }, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Cannot infer aspect ratio');
  });
});

describe('typed errors and error codes', () => {
  it('refuses delete with missing target as E_VALIDATION', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({ action: 'image', operation: 'delete' } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Delete image requires a target node id');
  });

  it('refuses resize with missing target as E_VALIDATION', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({ action: 'image', operation: 'resize', width: 100 } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Resize image requires a target node id');
  });

  it('refuses setAltText with missing target as E_VALIDATION', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({ action: 'image', operation: 'setAltText', altText: 'x' } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('SetAltText requires a target node id');
  });

  it('refuses setCaption with missing target as E_VALIDATION', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({ action: 'image', operation: 'setCaption', captionText: 'x' } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('SetCaption requires a target node id');
  });

  it('refuses setCaption with missing captionText as E_VALIDATION', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    const res = thrown(() => executeImage({ action: 'image', operation: 'setCaption', target: imgNode.id } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('SetCaption requires captionText');
  });

  it('refuses setWrap with missing target as E_VALIDATION', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({ action: 'image', operation: 'setWrap', wrapMode: 'square' } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Reposition/SetWrap requires a target node id');
  });

  it('refuses unsupported operation as E_VALIDATION', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({ action: 'image', operation: 'unsupported' as unknown as ImageOperation['operation'] } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('image.unsupported is not supported yet');
  });

  it('refuses insert with missing anchor as E_VALIDATION', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({ action: 'image', operation: 'insert', fileData: 'dummy' } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Insert image requires an anchor');
  });

  it('refuses insert when anchor node is not found in document as E_NOT_FOUND', async () => {
    const { parser, ast } = await load('T5_image_doc.docx');
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'insert',
      anchor: 'NO_SUCH_ANCHOR',
      fileData: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_NOT_FOUND');
    expect(res.message).toContain("Anchor node 'NO_SUCH_ANCHOR' not found in document");
  });

  it('refuses insert with invalid data URL format as E_VALIDATION', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const para = nodes.find((n) => n.type === 'paragraph')!;
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'insert',
      anchor: para.id,
      fileData: 'data:not-valid',
    } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Invalid image data URL format');
  });

  it('refuses insert with invalid base64 in data URL as E_VALIDATION', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const para = nodes.find((n) => n.type === 'paragraph')!;
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'insert',
      anchor: para.id,
      fileData: 'data:image/png;base64,???invalid',
    } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Invalid base64 in image data URL');
  });

  it('refuses insert with magic bytes mismatch in data URL as E_VALIDATION', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const para = nodes.find((n) => n.type === 'paragraph')!;
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'insert',
      anchor: para.id,
      fileData: 'data:image/png;base64,AAAA',
    } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Invalid image payload or magic bytes mismatch');
  });

  it('refuses insert with invalid base64 string as E_VALIDATION', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const para = nodes.find((n) => n.type === 'paragraph')!;
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'insert',
      anchor: para.id,
      fileData: '???notbase64',
    } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Invalid image data string');
  });

  it('refuses insert with magic bytes mismatch in base64 string as E_VALIDATION', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const para = nodes.find((n) => n.type === 'paragraph')!;
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'insert',
      anchor: para.id,
      fileData: Buffer.from('not an image at all').toString('base64'),
    } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('Invalid image payload or magic bytes mismatch');
  });

  it('refuses setCaption when target node is detached from DOM parent as E_STATE', async () => {
    const { parser, ast, nodes } = await load('T5_image_doc.docx');
    const imgNode = nodes.find((n) => n.type === 'image')!;
    (imgNode.xml as Element).parentNode?.removeChild(imgNode.xml as Element);
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'setCaption',
      target: imgNode.id,
      captionText: 'caption for detached',
    } as ImageOperation, ast, parser));
    expect(res.code).toBe('E_STATE');
    expect(res.message).toContain('detached from XML DOM parent');
  });

  it('contains zero bare new Error expressions in image.ts source', () => {
    const src = fs.readFileSync(path.resolve(__dirname, '../../packages/engine/src/operations/image.ts'), 'utf8');
    expect(src.match(/new Error\(/g)).toBeNull();
  });

  describe('image insert size calculation and sub-pixel guard', () => {
    function createPng200x100(): Buffer {
      const png = Buffer.alloc(33);
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52]).copy(png, 0);
      png.writeUInt32BE(200, 16);
      png.writeUInt32BE(100, 20);
      return png;
    }

    it('infers intrinsic size when neither width nor height is provided', async () => {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      const png = createPng200x100();
      executeImage({
        action: 'image',
        operation: 'insert',
        anchor: para.id,
        fileData: `data:image/png;base64,${png.toString('base64')}`,
      }, ast, parser);
      const p2 = new DocxParser(parser.rezip());
      const ext = p2.getDocumentDom().getElementsByTagName('wp:extent')[0];
      expect(ext.getAttribute('cx')).toBe('1905000');
      expect(ext.getAttribute('cy')).toBe('952500');
    });

    it('infers height preserving intrinsic aspect ratio when only width is provided', async () => {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      const png = createPng200x100();
      executeImage({
        action: 'image',
        operation: 'insert',
        anchor: para.id,
        fileData: `data:image/png;base64,${png.toString('base64')}`,
        width: 1828800,
      }, ast, parser);
      const p2 = new DocxParser(parser.rezip());
      const ext = p2.getDocumentDom().getElementsByTagName('wp:extent')[0];
      expect(ext.getAttribute('cx')).toBe('1828800');
      expect(ext.getAttribute('cy')).toBe('914400');
    });

    it('infers width preserving intrinsic aspect ratio when only height is provided', async () => {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      const png = createPng200x100();
      executeImage({
        action: 'image',
        operation: 'insert',
        anchor: para.id,
        fileData: `data:image/png;base64,${png.toString('base64')}`,
        height: 457200,
      }, ast, parser);
      const p2 = new DocxParser(parser.rezip());
      const ext = p2.getDocumentDom().getElementsByTagName('wp:extent')[0];
      expect(ext.getAttribute('cx')).toBe('914400');
      expect(ext.getAttribute('cy')).toBe('457200');
    });

    it('uses both width and height exactly when both are provided', async () => {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      const png = createPng200x100();
      executeImage({
        action: 'image',
        operation: 'insert',
        anchor: para.id,
        fileData: `data:image/png;base64,${png.toString('base64')}`,
        width: 1000000,
        height: 500000,
      }, ast, parser);
      const p2 = new DocxParser(parser.rezip());
      const ext = p2.getDocumentDom().getElementsByTagName('wp:extent')[0];
      expect(ext.getAttribute('cx')).toBe('1000000');
      expect(ext.getAttribute('cy')).toBe('500000');
    });

    it('refuses width smaller than one pixel with E_VALIDATION and leaves document unchanged', async () => {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      const png = createPng200x100();
      const beforeBytes = parser.rezip();
      const beforeMediaCount = Object.keys(ast.mediaFiles ?? {}).length;
      const res = thrown(() => executeImage({
        action: 'image',
        operation: 'insert',
        anchor: para.id,
        fileData: `data:image/png;base64,${png.toString('base64')}`,
        width: 100,
      }, ast, parser));
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('smaller than one pixel (9525 EMU)');
      const afterBytes = parser.rezip();
      expect(Buffer.compare(beforeBytes, afterBytes)).toBe(0);
      expect(Object.keys(ast.mediaFiles ?? {}).length).toBe(beforeMediaCount);
    });

    it('refuses height smaller than one pixel with E_VALIDATION and leaves document unchanged', async () => {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      const png = createPng200x100();
      const beforeBytes = parser.rezip();
      const beforeMediaCount = Object.keys(ast.mediaFiles ?? {}).length;
      const res = thrown(() => executeImage({
        action: 'image',
        operation: 'insert',
        anchor: para.id,
        fileData: `data:image/png;base64,${png.toString('base64')}`,
        height: 100,
      }, ast, parser));
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain('smaller than one pixel (9525 EMU)');
      const afterBytes = parser.rezip();
      expect(Buffer.compare(beforeBytes, afterBytes)).toBe(0);
      expect(Object.keys(ast.mediaFiles ?? {}).length).toBe(beforeMediaCount);
    });

    it('keeps default dimensions for SVG without width height or viewBox', async () => {
      const { parser, ast, nodes } = await load('T5_image_doc.docx');
      const para = nodes.find((n) => n.type === 'paragraph')!;
      const svg = '<svg xmlns="http://www.w3.org/2000/svg"><circle cx="50" cy="50" r="40"/></svg>';
      executeImage({
        action: 'image',
        operation: 'insert',
        anchor: para.id,
        fileData: `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`,
      }, ast, parser);
      const p2 = new DocxParser(parser.rezip());
      const ext = p2.getDocumentDom().getElementsByTagName('wp:extent')[0];
      expect(ext.getAttribute('cx')).toBe('5486400');
      expect(ext.getAttribute('cy')).toBe('3200400');
    });
  });
});

describe('paragraph holding text and picture', () => {
  async function loadParagraphWithPicture(sameRun = false) {
    const buf = fs.readFileSync(path.join(GALLERY, 'T5_image_doc.docx'));
    const p1 = new DocxParser(buf);
    const ast1 = await p1.getAST();
    const imgNode = ast1.sections.flatMap((s) => s.nodes).find((n) => n.type === 'image')!;
    const xml = imgNode.xml as Element;
    const doc = xml.ownerDocument;
    const r = doc.createElement('w:r');
    const t = doc.createElement('w:t');
    t.textContent = 'Text beside the picture. ';
    if (!sameRun) {
      r.appendChild(t);
      const pPr = xml.getElementsByTagName('w:pPr')[0];
      if (pPr && pPr.nextSibling) {
        xml.insertBefore(r, pPr.nextSibling);
      } else {
        xml.insertBefore(r, xml.firstChild);
      }
    } else {
      const drawing = xml.getElementsByTagName('w:drawing')[0];
      const drawingRun = drawing.parentNode as Element;
      drawingRun.insertBefore(t, drawingRun.firstChild);
    }
    const modifiedBuf = p1.rezip();
    const parser = new DocxParser(modifiedBuf);
    const ast = await parser.getAST();
    const node = ast.sections.flatMap((s) => s.nodes).find((n) => n.id === imgNode.id)!;
    return { parser, ast, node };
  }

  it('resize with both sizes writes the new wp:extent', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture();
    expect(node.type).toBe('paragraph');
    expect(node.metadata.hasImage).toBe(true);
    executeImage({ action: 'image', operation: 'resize', target: node.id, width: 1828800, height: 914400 }, ast, parser);
    const ext = (node.xml as Element).getElementsByTagName('wp:extent')[0];
    expect(ext.getAttribute('cx')).toBe('1828800');
    expect(ext.getAttribute('cy')).toBe('914400');
  });

  it('resize with only width keeps the ratio', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture();
    expect(node.type).toBe('paragraph');
    expect(node.metadata.hasImage).toBe(true);
    executeImage({ action: 'image', operation: 'resize', target: node.id, width: 1828800, height: 914400 }, ast, parser);
    executeImage({ action: 'image', operation: 'resize', target: node.id, width: 3657600 }, ast, parser);
    const ext = (node.xml as Element).getElementsByTagName('wp:extent')[0];
    expect(ext.getAttribute('cx')).toBe('3657600');
    expect(ext.getAttribute('cy')).toBe('1828800');
  });

  it('setAltText writes descr', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture();
    expect(node.type).toBe('paragraph');
    expect(node.metadata.hasImage).toBe(true);
    executeImage({ action: 'image', operation: 'setAltText', target: node.id, altText: 'Sample description' }, ast, parser);
    const docPr = (node.xml as Element).getElementsByTagName('wp:docPr')[0];
    expect(docPr.getAttribute('descr')).toBe('Sample description');
  });

  it('setWrap square writes wp:wrapSquare', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture();
    expect(node.type).toBe('paragraph');
    expect(node.metadata.hasImage).toBe(true);
    executeImage({ action: 'image', operation: 'setWrap', target: node.id, wrapMode: 'square' }, ast, parser);
    const wrapSquare = (node.xml as Element).getElementsByTagName('wp:wrapSquare')[0];
    expect(wrapSquare).toBeDefined();
    expect(wrapSquare.getAttribute('wrapText')).toBe('bothSides');
  });

  it('delete keeps the text, leaves no w:drawing, and sets hasImage to false', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture();
    expect(node.type).toBe('paragraph');
    expect(node.metadata.hasImage).toBe(true);
    executeImage({ action: 'image', operation: 'delete', target: node.id }, ast, parser);
    expect((node.xml as Element).getElementsByTagName('w:drawing').length).toBe(0);
    expect(node.metadata.hasImage).toBe(false);
    expect((node.xml as Element).textContent).toContain('Text beside the picture.');
    const remainingNode = ast.sections.flatMap((s) => s.nodes).find((n) => n.id === node.id);
    expect(remainingNode).toBeDefined();
    expect(remainingNode?.metadata.hasImage).toBe(false);
  });

  it('delete keeps the text when the picture and the text share one run', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture(true);
    expect(node.type).toBe('paragraph');
    expect(node.metadata.hasImage).toBe(true);
    executeImage({ action: 'image', operation: 'delete', target: node.id }, ast, parser);
    expect((node.xml as Element).getElementsByTagName('w:drawing').length).toBe(0);
    expect(node.metadata.hasImage).toBe(false);
    expect((node.xml as Element).textContent).toContain('Text beside the picture.');
    const remainingNode = ast.sections.flatMap((s) => s.nodes).find((n) => n.id === node.id);
    expect(remainingNode).toBeDefined();
    expect(remainingNode?.metadata.hasImage).toBe(false);
  });

  it('delete removes an mc:AlternateContent wrapper together with the picture', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture();
    const xml = node.xml as Element;
    const doc = xml.ownerDocument;
    const drawing = xml.getElementsByTagName('w:drawing')[0];
    const drawingRun = drawing.parentNode as Element;
    const alt = doc.createElement('mc:AlternateContent');
    const choice = doc.createElement('mc:Choice');
    choice.setAttribute('Requires', 'w14');
    drawingRun.replaceChild(alt, drawing);
    choice.appendChild(drawing);
    alt.appendChild(choice);

    executeImage({ action: 'image', operation: 'delete', target: node.id }, ast, parser);
    expect(xml.getElementsByTagName('w:drawing').length).toBe(0);
    expect(xml.getElementsByTagName('mc:AlternateContent').length).toBe(0);
    expect(node.metadata.hasImage).toBe(false);
    expect(xml.textContent).toContain('Text beside the picture.');
  });

  it('plain text paragraph still gives E_VALIDATION containing requires an image node', async () => {
    const { parser, ast, nodes } = await load('T2_edit_target.docx');
    const plainPara = nodes.find((n) => n.type === 'paragraph')!;
    expect(plainPara).toBeDefined();
    expect(plainPara.metadata.hasImage).toBeFalsy();
    const res = thrown(() => executeImage({
      action: 'image',
      operation: 'resize',
      target: plainPara.id,
      width: 914400,
      height: 914400,
    }, ast, parser));
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('requires an image node');
  });
});
