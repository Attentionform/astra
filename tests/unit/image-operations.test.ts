import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocumentAST,
  executeImage,
  ImageOperation,
} from '@attentionform/astra-engine';

function createMockAST(): DocumentAST {
  const doc = new DOMParser().parseFromString('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>', 'text/xml');
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
          hasHyperlink: false
        },
        version: 1,
        checksum: 'chk_1'
      }],
      pageSetup: {},
      version: 1,
      checksum: 'sec_chk_1'
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1
  };
}

describe('Image Operations', () => {
  it('should insert an inline image and allow resizing', () => {
    const ast = createMockAST();
    const dummyBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const insertOp: ImageOperation = {
      action: 'image',
      operation: 'insert',
      anchor: 'p_001',
      fileData: `data:image/png;base64,${dummyBase64}`,
      width: 914400,
      height: 457200
    };

    executeImage(insertOp, ast);

    expect(ast.sections[0].nodes.length).toBe(2);
    const imgNode = ast.sections[0].nodes[1];
    expect(imgNode.type).toBe('image');
    expect(imgNode.metadata.hasImage).toBe(true);

    // Resize image
    const resizeOp: ImageOperation = {
      action: 'image',
      operation: 'resize',
      target: imgNode.id,
      width: 1828800,
      height: 914400
    };

    executeImage(resizeOp, ast);
    const extent = imgNode.xml.getElementsByTagName('wp:extent')[0];
    expect(extent.getAttribute('cx')).toBe('1828800');
    expect(extent.getAttribute('cy')).toBe('914400');
  });

  it('should support wrapMode conversion, altText, and delete', () => {
    const ast = createMockAST();
    const dummyBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const insertOp: ImageOperation = {
      action: 'image',
      operation: 'insert',
      anchor: 'p_001',
      fileData: `data:image/png;base64,${dummyBase64}`
    };

    executeImage(insertOp, ast);
    const imgNode = ast.sections[0].nodes[1];

    // Set Alt text
    executeImage({
      action: 'image',
      operation: 'setAltText',
      target: imgNode.id,
      altText: 'Chart diagram showing quarterly revenue'
    }, ast);

    const docPr = imgNode.xml.getElementsByTagName('wp:docPr')[0];
    expect(docPr.getAttribute('descr')).toBe('Chart diagram showing quarterly revenue');

    // Set wrap to square
    executeImage({
      action: 'image',
      operation: 'setWrap',
      target: imgNode.id,
      wrapMode: 'square'
    }, ast);

    expect(imgNode.xml.getElementsByTagName('wp:anchor').length).toBe(1);
    expect(imgNode.xml.getElementsByTagName('wp:wrapSquare').length).toBe(1);

    // Delete image
    executeImage({
      action: 'image',
      operation: 'delete',
      target: imgNode.id
    }, ast);

    expect(ast.sections[0].nodes.length).toBe(1);
  });
});
