import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import {
  OperationSchema,
  validateOperationStrict,
  validateOperationsStrict,
  executeFootnote,
  executeEndnote,
  OperationEngine,
  convertAstToMarkdown,
  convertAstToHtml,
  executeHyperlink,
  executeImage,
  executeComment,
  DocumentAST,
  DocxParser,
  validateDocxBuffer,
  WORDPROCESSING_DRAWING_NS,
} from '@attentionform/astra-engine';
import { DOMParser } from '@xmldom/xmldom';

function createMockImageAST(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:id="p_001"/><w:p w:id="img_p_002"/></w:body></w:document>',
    'text/xml'
  );
  const p1 = doc.getElementsByTagName('w:p')[0];
  const p2 = doc.getElementsByTagName('w:p')[1];

  return {
    metadata: {},
    sections: [
      {
        id: 'sec_1',
        heading: null,
        nodes: [
          {
            id: 'p_001',
            type: 'paragraph',
            xml: p1,
            metadata: {
              text: 'Introductory text',
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
            checksum: 'chk_p1',
          },
          {
            id: 'img_node_001',
            type: 'image',
            xml: p2,
            metadata: {
              text: '',
              hasImage: true,
              hasCitation: false,
              hasComment: false,
              hasBookmark: false,
              hasTrackedChanges: false,
              hasEquation: false,
              hasField: false,
              hasHyperlink: false,
            },
            version: 1,
            checksum: 'chk_img1',
          },
        ],
        pageSetup: {},
        version: 1,
        checksum: 'sec_chk_1',
      },
    ],
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

describe('Operation Schema & F4 Alignment Negative Tests', () => {
  describe('namedStyleOp schema alignment', () => {
    it('accepts supported "apply" operation', () => {
      const valid = {
        action: 'namedStyle',
        operation: 'apply',
        styleName: 'Heading1',
        target: 'p_001',
      };
      const res = OperationSchema.safeParse(valid);
      expect(res.success).toBe(true);
      expect(() => validateOperationStrict(valid)).not.toThrow();
    });

    it('rejects unsupported "create", "modify", "delete" operations fail-closed', () => {
      for (const op of ['create', 'modify', 'delete', 'rename', 'purge']) {
        const invalid = {
          action: 'namedStyle',
          operation: op,
          styleName: 'CustomStyle',
          target: 'p_001',
        };
        const res = OperationSchema.safeParse(invalid);
        expect(res.success).toBe(false);
        expect(() => validateOperationStrict(invalid as any)).toThrow(/E_VALIDATION/);
      }
    });
  });

  describe('fieldOp schema alignment', () => {
    it('accepts supported operations (insertTOC, insertPageNumber, insertWordCount, updateFields)', () => {
      const ops = ['insertTOC', 'insertPageNumber', 'insertWordCount', 'updateFields'] as const;
      for (const op of ops) {
        const valid = {
          action: 'field',
          operation: op,
        };
        const res = OperationSchema.safeParse(valid);
        expect(res.success).toBe(true);
        expect(() => validateOperationStrict(valid)).not.toThrow();
      }
    });

    it('rejects unsupported field operations (insertTOF, removeField, etc.) fail-closed', () => {
      for (const op of ['insertTOF', 'removeField', 'editField', 'toggleField']) {
        const invalid = {
          action: 'field',
          operation: op,
        };
        const res = OperationSchema.safeParse(invalid);
        expect(res.success).toBe(false);
        expect(() => validateOperationStrict(invalid as any)).toThrow(/E_VALIDATION/);
      }
    });
  });

  describe('imageOp setCaption schema and execution', () => {
    it('accepts setCaption operation schema with captionText', () => {
      const valid = {
        action: 'image',
        operation: 'setCaption',
        target: 'img_node_001',
        captionText: 'Figure 1. Operational flow diagram.',
      };
      const res = OperationSchema.safeParse(valid);
      expect(res.success).toBe(true);
      expect(() => validateOperationStrict(valid)).not.toThrow();
    });

    it('throws when target node is not found for setCaption', () => {
      const ast = createMockImageAST();
      expect(() =>
        executeImage(
          {
            action: 'image',
            operation: 'setCaption',
            target: 'non_existent_node',
            captionText: 'Figure 1',
          },
          ast
        )
      ).toThrow("Target node 'non_existent_node' not found for setCaption");
    });

    it('throws when captionText is missing for setCaption', () => {
      const ast = createMockImageAST();
      expect(() =>
        executeImage(
          {
            action: 'image',
            operation: 'setCaption',
            target: 'img_node_001',
          } as any,
          ast
        )
      ).toThrow('SetCaption requires captionText');
    });

    it('creates a caption paragraph node with namedStyle="Caption" and splices it after target', () => {
      const ast = createMockImageAST();
      const targetNode = ast.sections[0].nodes[1];
      const initialVersion = targetNode.version;
      const initialSecVersion = ast.sections[0].version;

      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: 'img_node_001',
          captionText: 'Figure 1. Architectural schema overview.',
        },
        ast
      );

      // Spliced into section immediately following image node
      expect(ast.sections[0].nodes.length).toBe(3);
      const captionNode = ast.sections[0].nodes[2];
      expect(captionNode.type).toBe('paragraph');
      expect(captionNode.metadata.namedStyle).toBe('Caption');
      expect(captionNode.metadata.text).toBe('Figure 1. Architectural schema overview.');
      expect(captionNode.checksum).toBeTruthy();
      expect(captionNode.checksum.length).toBe(64); // SHA-256 hex
      expect(targetNode.version).toBe(initialVersion + 1);
      expect(ast.sections[0].version).toBe(initialSecVersion + 1);

      // Verify DOM structure
      const pStyle = captionNode.xml.getElementsByTagName('w:pStyle')[0];
      expect(pStyle).toBeDefined();
      expect(pStyle.getAttribute('w:val')).toBe('Caption');
      const t = captionNode.xml.getElementsByTagName('w:t')[0];
      expect(t.textContent).toBe('Figure 1. Architectural schema overview.');
    });

    it('updates existing caption node when setCaption is called again on the same image', () => {
      const ast = createMockImageAST();
      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: 'img_node_001',
          captionText: 'Initial Caption',
        },
        ast
      );

      expect(ast.sections[0].nodes.length).toBe(3);
      const captionNode = ast.sections[0].nodes[2];
      const prevChecksum = captionNode.checksum;

      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: 'img_node_001',
          captionText: 'Updated Caption Text',
        },
        ast
      );

      // Does not create a duplicate paragraph node; updates existing one
      expect(ast.sections[0].nodes.length).toBe(3);
      expect(captionNode.metadata.text).toBe('Updated Caption Text');
      expect(captionNode.checksum).not.toBe(prevChecksum);
      const t = captionNode.xml.getElementsByTagName('w:t')[0];
      expect(t.textContent).toBe('Updated Caption Text');
    });

    it('Defect 1.1: respects custom id or generates collision-free id with sequence', () => {
      const ast = createMockImageAST();
      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: 'img_node_001',
          captionText: 'Custom ID Caption',
          id: 'custom_cap_42',
        } as any,
        ast
      );

      const capNode = ast.sections[0].nodes[2];
      expect(capNode.id).toBe('custom_cap_42');
      expect(capNode.xml.getAttribute('data-astra-id')).toBe('custom_cap_42');
    });

    it('Defect 1.2: registers new caption node in parser.nodeMap when parser is present', () => {
      const ast = createMockImageAST();
      const mockNodeMap = new Map<string, any>();
      const mockParser = { nodeMap: mockNodeMap } as any;

      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: 'img_node_001',
          captionText: 'Parser Map Caption',
          id: 'mapped_caption_1',
        } as any,
        ast,
        mockParser
      );

      expect(mockNodeMap.has('mapped_caption_1')).toBe(true);
      expect(mockNodeMap.get('mapped_caption_1')).toBe(ast.sections[0].nodes[2].xml);
    });

    it('Defect 1.3: consolidates multiple runs and removes trailing runs when updating existing caption', () => {
      const ast = createMockImageAST();
      // First create caption
      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: 'img_node_001',
          captionText: 'Initial Part 1',
        },
        ast
      );

      const capNode = ast.sections[0].nodes[2];
      const doc = capNode.xml.ownerDocument;
      // Artificially append additional runs simulating Word/split runs
      const r2 = doc.createElement('w:r');
      const t2 = doc.createElement('w:t');
      t2.textContent = ' and Part 2';
      r2.appendChild(t2);
      capNode.xml.appendChild(r2);

      const r3 = doc.createElement('w:r');
      const t3 = doc.createElement('w:t');
      t3.textContent = ' and Part 3';
      r3.appendChild(t3);
      capNode.xml.appendChild(r3);

      expect(capNode.xml.getElementsByTagName('w:r').length).toBe(3);

      // Now update caption
      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: 'img_node_001',
          captionText: 'Consolidated Single Caption',
        },
        ast
      );

      const runsAfter = capNode.xml.getElementsByTagName('w:r');
      expect(runsAfter.length).toBe(1);
      expect(runsAfter[0].getElementsByTagName('w:t')[0].textContent).toBe('Consolidated Single Caption');
    });

    it('Defect 1.4: throws fail-fast error when target node is detached from XML DOM parent', () => {
      const ast = createMockImageAST();
      const targetNode = ast.sections[0].nodes[1];
      // Detach XML element from parent
      targetNode.xml.parentNode?.removeChild(targetNode.xml);
      expect(targetNode.xml.parentNode).toBeNull();

      expect(() =>
        executeImage(
          {
            action: 'image',
            operation: 'setCaption',
            target: 'img_node_001',
            captionText: 'Detached Caption',
          },
          ast
        )
      ).toThrow("Cannot insert caption: target node 'img_node_001' is detached from XML DOM parent.");
    });

    it('applies setCaption to real gallery document T5_image_doc.docx and passes strict OOXML validation', async () => {
      const docPath = path.resolve(__dirname, '../../test-gallery/T5_image_doc.docx');
      const buf = fs.readFileSync(docPath);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();

      // Find the image node in T5
      const imageNode = ast.sections[0].nodes.find((n) => n.metadata.hasImage);
      expect(imageNode).toBeDefined();
      if (!imageNode) {
        throw new Error('Failed to locate image node in T5_image_doc.docx');
      }

      executeImage(
        {
          action: 'image',
          operation: 'setCaption',
          target: imageNode.id,
          captionText: 'Figure 1: Verified canonical OOXML image caption.',
        },
        ast,
        parser
      );

      const mutatedBuf = parser.rezip();
      const valResult = await validateDocxBuffer(mutatedBuf);
      expect(valResult.ok).toBe(true);
      expect(valResult.errors).toEqual([]);
    });

    it('applies setWrap on T5_image_doc.docx with canonical WORDPROCESSING_DRAWING_NS and passes OOXML validation', async () => {
      const docPath = path.resolve(__dirname, '../../test-gallery/T5_image_doc.docx');
      const buf = fs.readFileSync(docPath);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();

      const imageNode = ast.sections[0].nodes.find((n) => n.metadata.hasImage);
      expect(imageNode).toBeDefined();
      if (!imageNode) {
        throw new Error('Failed to locate image node in T5_image_doc.docx');
      }

      executeImage(
        {
          action: 'image',
          operation: 'setWrap',
          target: imageNode.id,
          wrapMode: 'square',
        },
        ast,
        parser
      );

      // Verify resulting wp:anchor and wp:wrapSquare elements have WORDPROCESSING_DRAWING_NS
      const anchorEl = imageNode.xml.getElementsByTagName('wp:anchor')[0];
      expect(anchorEl).toBeDefined();
      expect(anchorEl.namespaceURI).toBe(WORDPROCESSING_DRAWING_NS);

      const wrapSquareEl = anchorEl.getElementsByTagName('wp:wrapSquare')[0];
      expect(wrapSquareEl).toBeDefined();
      expect(wrapSquareEl.namespaceURI).toBe(WORDPROCESSING_DRAWING_NS);

      const mutatedBuf = parser.rezip();
      const valResult = await validateDocxBuffer(mutatedBuf);
      expect(valResult.ok).toBe(true);
      expect(valResult.errors).toEqual([]);
    });
  });

  describe('pre-prod regression: bookmark pPr order + footnote id scan', () => {
    it('bookmark-after-pPr: bookmarkStart lands after w:pPr (ECMA-376 §17.3.1.22) and blank+bookmark validates ok:true', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      const targetId = ast.sections[0].nodes[0].id;
      executeHyperlink(
        { action: 'bookmark', target: targetId, bookmarkName: 'bm_regression' } as never,
        ast,
        parser
      );
      const targetNode = ast.sections[0].nodes[0];
      const kids = Array.from(targetNode.xml.childNodes).filter(
        (n) => n.nodeType === 1
      ) as Element[];
      const pPrIdx = kids.findIndex((el) => el.tagName === 'w:pPr');
      const bmIdx = kids.findIndex((el) => el.tagName === 'w:bookmarkStart');
      expect(pPrIdx).toBeGreaterThanOrEqual(0);
      expect(bmIdx).toBeGreaterThan(pPrIdx);
      expect(kids[0].tagName).toBe('w:pPr');
      const buf = parser.rezip();
      const res = await validateDocxBuffer(buf);
      expect(res.ok).toBe(true);
      expect(res.errors).toEqual([]);
    });

    it('footnote-no-dup: footnote on T8 (existing footnotes.xml ids) allocates max+1 and validates ok:true', async () => {
      const docPath = path.resolve(__dirname, '../../test-gallery/T8_footnotes_endnotes.docx');
      const buf = fs.readFileSync(docPath);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const anchor = ast.sections[0].nodes[0].id;
      executeFootnote(
        { action: 'footnote', operation: 'insert', anchor, text: 'regression footnote' } as never,
        ast,
        parser
      );
      const out = parser.rezip();
      const res = await validateDocxBuffer(out);
      expect(res.ok).toBe(true);
      expect(res.errors).toEqual([]);
      // No duplicate w:id in footnotes.xml
      const fnXml: string = (parser.getZipAdapter() as unknown as { readText: (p: string) => string }).readText(
        'word/footnotes.xml'
      );
      const ids = [...fnXml.matchAll(/w:id="(-?\d+)"/g)].map((m) => m[1]);
      const seen = new Set<string>();
      for (const id of ids) {
        // Separator ids (-1, 0) are singletons by construction; real notes must be unique
        if (id === '-1' || id === '0') continue;
        expect(seen.has(id)).toBe(false);
        seen.add(id);
      }
    });
  });
});

describe('Phase A engine P0 regressions', () => {
  const PNG_1PX =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const T5 = path.resolve(__dirname, '../../test-gallery/T5_image_doc.docx');
  const T4 = path.resolve(__dirname, '../../test-gallery/T4_budget.docx');
  const T2 = path.resolve(__dirname, '../../test-gallery/T2_edit_target.docx');

  it('A-1: setWrap validates ok:true for all 7 wrap modes (tight/through carry wrapText + polygon)', async () => {
    const modes = ['inline', 'square', 'tight', 'through', 'topAndBottom', 'behindText', 'inFrontOfText'] as const;
    for (const wrapMode of modes) {
      const parser = new DocxParser(fs.readFileSync(T5));
      const ast = await parser.getAST();
      const imageNode = ast.sections[0].nodes.find((n) => n.metadata.hasImage);
      expect(imageNode).toBeDefined();
      executeImage({ action: 'image', operation: 'setWrap', target: imageNode!.id, wrapMode }, ast, parser);
      if (wrapMode === 'tight' || wrapMode === 'through') {
        const drawing = imageNode!.xml.getElementsByTagName('w:drawing')[0] || imageNode!.xml;
        const anchor = drawing.getElementsByTagName('wp:anchor')[0];
        expect(anchor).toBeDefined();
        const tag = wrapMode === 'tight' ? 'wp:wrapTight' : 'wp:wrapThrough';
        const wrap = anchor.getElementsByTagName(tag)[0];
        expect(wrap).toBeDefined();
        expect(wrap.getAttribute('wrapText')).toBe('bothSides');
        expect(wrap.getElementsByTagName('wp:wrapPolygon').length).toBe(1);
      }
      const res = await validateDocxBuffer(parser.rezip());
      expect(res.ok).toBe(true);
      expect(res.errors).toEqual([]);
    }
  }, 90000);

  it('A-2: 5-image insert mints unique drawing ids and saves ok:true; insert(tight)+caption valid', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    for (let i = 0; i < 5; i++) {
      executeImage({ action: 'image', operation: 'insert', anchor, fileData: PNG_1PX } as any, ast, parser);
    }
    const dom = parser.getDocumentDom();
    const docPrIds = Array.from(dom.getElementsByTagName('wp:docPr')).map((e) => (e as Element).getAttribute('id'));
    const cNvIds = Array.from(dom.getElementsByTagName('pic:cNvPr')).map((e) => (e as Element).getAttribute('id'));
    expect(docPrIds.length).toBe(5);
    expect(new Set(docPrIds).size).toBe(docPrIds.length);
    expect(new Set(cNvIds).size).toBe(cNvIds.length);
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);

    // Quarantined insert hunk: wrapMode + captionText honored on insert, tight stays valid
    const parser2 = new DocxParser(fs.readFileSync(T5));
    const ast2 = await parser2.getAST();
    executeImage(
      {
        action: 'image',
        operation: 'insert',
        anchor: ast2.sections[0].nodes[0].id,
        fileData: PNG_1PX,
        wrapMode: 'tight',
        captionText: 'Figure A-2. Tight-wrapped insert.',
      } as any,
      ast2,
      parser2
    );
    const caption = ast2.sections.flatMap((s) => s.nodes).find((n) => n.metadata.namedStyle === 'Caption');
    expect(caption?.metadata.text).toBe('Figure A-2. Tight-wrapped insert.');
    const res2 = await validateDocxBuffer(parser2.rezip());
    expect(res2.ok).toBe(true);
    expect(res2.errors).toEqual([]);
  }, 90000);

  it('A-3: comment-on-table routes to first-cell paragraph, never w:r-under-w:tbl', async () => {
    const parser = new DocxParser(fs.readFileSync(T4));
    const ast = await parser.getAST();
    const tableNode = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'table');
    expect(tableNode).toBeDefined();
    executeComment(
      { action: 'comment', operation: 'insert', target: tableNode!.id, text: 'A-3 table comment' } as any,
      ast,
      parser
    );
    const tbls = parser.getDocumentDom().getElementsByTagName('w:tbl');
    expect(tbls.length).toBeGreaterThan(0);
    for (let i = 0; i < tbls.length; i++) {
      const kids = Array.from((tbls[i] as unknown as Element).childNodes);
      for (const k of kids) {
        const el = k as Element;
        if ((el as any).nodeType !== 1) continue;
        expect(['w:r', 'w:commentRangeStart', 'w:commentRangeEnd']).not.toContain(el.tagName);
      }
    }
    const firstCellP = ((tbls[0] as unknown as Element).getElementsByTagName('w:tr')[0] as unknown as Element)
      .getElementsByTagName('w:tc')[0]
      .getElementsByTagName('w:p')[0] as unknown as Element;
    expect(firstCellP.getElementsByTagName('w:commentRangeStart').length).toBe(1);
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  }, 60000);

  it('A-4: paragraph ids survive 3 save/reopen cycles and every save validates', async () => {
    let parser = new DocxParser(fs.readFileSync(T2));
    let ast = await parser.getAST();
    const paraIds = () =>
      ast.sections.flatMap((s) => s.nodes.filter((n) => n.type === 'paragraph').map((n) => n.id));
    const ids = paraIds();
    expect(ids.length).toBeGreaterThan(0);
    // Every paragraph carries a valid w14:paraId after load (minted or native)
    for (const id of ids) {
      expect(/^[0-9A-Fa-f]{8}$/.test(id)).toBe(true);
    }
    for (let cycle = 1; cycle <= 3; cycle++) {
      const out = parser.rezip();
      const res = await validateDocxBuffer(out);
      expect(res.ok).toBe(true);
      expect(res.errors).toEqual([]);
      parser = new DocxParser(out);
      ast = await parser.getAST();
      expect(paraIds()).toEqual(ids);
    }
  }, 120000);
});

describe('endnote insert (footnote mirror)', () => {
  const T8 = path.resolve(__dirname, '../../test-gallery/T8_footnotes_endnotes.docx');

  it('schema gate: accepts canonical endnote insert, rejects missing text', () => {
    const ok = validateOperationStrict({
      action: 'endnote',
      operation: 'insert',
      anchor: 'p_001',
      text: 'Archive note.',
    });
    expect(ok.ok).toBe(true);
    expect((ok.operation as { action: string }).action).toBe('endnote');
    const batch = validateOperationsStrict([
      { action: 'endnote', operation: 'insert', anchor: 'p_001', text: 'Batch endnote.' },
    ]);
    expect(batch.ok).toBe(true);
    expect(batch.operations).toHaveLength(1);
    expect(() => validateOperationStrict({ action: 'endnote', anchor: 'p_001' })).toThrow(/E_VALIDATION/);
  });

  it('engine route: OperationEngine executes action endnote', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;
    await engine.execute({ action: 'endnote', operation: 'insert', anchor, text: 'Routed endnote.' } as never);
    expect(parser.readDocumentXml()).toContain('w:endnoteReference');
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('blank insert: creates word/endnotes.xml (separator intact, EndnoteText) and validates', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor, text: 'Blank endnote alpha' } as never,
      ast,
      parser
    );
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string; hasFile: (p: string) => boolean };
    expect(zip.hasFile('word/endnotes.xml')).toBe(true);
    const enXml = zip.readText('word/endnotes.xml');
    // Separator pair intact (mirrors T8 ground truth)
    expect(enXml).toContain('w:type="separator" w:id="-1"');
    expect(enXml).toContain('w:type="continuationSeparator" w:id="0"');
    // Endnote style + ref markers (not footnote names)
    expect(enXml).toContain('EndnoteText');
    expect(enXml).toContain('w:endnoteRef');
    expect(enXml).toContain('Blank endnote alpha');
    expect(enXml).not.toContain('w:footnote');
    expect(parser.readDocumentXml()).toContain('w:endnoteReference');
    // Well-formed + no duplicate real ids
    new DOMParser().parseFromString(enXml, 'text/xml');
    const ids = [...enXml.matchAll(/w:id="(-?\d+)"/g)].map((m) => m[1]).filter((id) => id !== '-1' && id !== '0');
    expect(new Set(ids).size).toBe(ids.length);
    const out = parser.rezip();
    const res = await validateDocxBuffer(out);
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('T8 insert: allocates max+1 (id 2, no dup) and validates', async () => {
    const parser = new DocxParser(fs.readFileSync(T8));
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor, text: 'T8 endnote three' } as never,
      ast,
      parser
    );
    const enXml: string = (parser.getZipAdapter() as unknown as { readText: (p: string) => string }).readText(
      'word/endnotes.xml'
    );
    const ids = [...enXml.matchAll(/w:id="(-?\d+)"/g)].map((m) => m[1]);
    expect(ids).toContain('2');
    const seen = new Set<string>();
    for (const id of ids) {
      if (id === '-1' || id === '0') continue;
      expect(seen.has(id)).toBe(false);
      seen.add(id);
    }
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('export: footnote and endnote bodies render distinctly in MD + HTML (en: namespacing)', async () => {
    const parser = new DocxParser(fs.readFileSync(T8));
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    executeFootnote(
      { action: 'footnote', operation: 'insert', anchor, text: 'FN distinct body ZZZ' } as never,
      ast,
      parser
    );
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor, text: 'EN distinct body QQQ' } as never,
      ast,
      parser
    );
    const md = convertAstToMarkdown(ast, parser);
    expect(md).toContain('FN distinct body ZZZ');
    expect(md).toContain('EN distinct body QQQ');
    // Footnote def is [^N]; endnote def is namespaced [^en-N] — never collide
    expect(md).toMatch(/\[\^[0-9]+\]: FN distinct body ZZZ/);
    expect(md).toMatch(/\[\^en-[0-9]+\]: EN distinct body QQQ/);
    const html = convertAstToHtml(ast, 'Notes', parser);
    expect(html).toContain('FN distinct body ZZZ');
    expect(html).toContain('EN distinct body QQQ');
    expect(html).toMatch(/id="fn-[0-9]+"/);
    expect(html).toMatch(/id="en-[0-9]+"/);
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('settings part: endnote insert on blank creates settings with both note properties', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor, text: 'Settings endnote' } as never,
      ast,
      parser
    );
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string; hasFile: (p: string) => boolean };
    expect(zip.hasFile('word/settings.xml')).toBe(true);
    const settingsXml = zip.readText('word/settings.xml');
    expect(settingsXml).toContain('w:footnotePr');
    expect(settingsXml).toContain('w:endnotePr');
    expect(settingsXml).toContain('w:val="docEnd"');
    // Part is registered in content types and document rels.
    expect(zip.readText('[Content_Types].xml')).toContain('/word/settings.xml');
    expect(zip.readText('word/_rels/document.xml.rels')).toContain('settings.xml');
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('settings part: footnote insert ensures both properties and repeat inserts stay single', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    executeFootnote(
      { action: 'footnote', operation: 'insert', anchor, text: 'Settings footnote' } as never,
      ast,
      parser
    );
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor, text: 'Settings endnote two' } as never,
      ast,
      parser
    );
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string; hasFile: (p: string) => boolean };
    const settingsXml = zip.readText('word/settings.xml');
    expect(settingsXml).toContain('w:footnotePr');
    expect(settingsXml).toContain('w:endnotePr');
    expect(settingsXml).toContain('w:val="docEnd"');
    // Repeat inserts never duplicate the part registration.
    const overrides = (zip.readText('[Content_Types].xml').match(/\/word\/settings\.xml/g) ?? []).length;
    expect(overrides).toBe(1);
    const rels = (zip.readText('word/_rels/document.xml.rels').match(/settings\.xml/g) ?? []).length;
    expect(rels).toBe(1);
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });
});

describe('note settings patch: ordering, degenerate parts, staged writes', () => {
  function settingsOrder(settingsXml: string): string[] {
    return [...settingsXml.matchAll(/<w:(footnotePr|endnotePr|compat|rsids|updateFields|zoom|defaultTabStop)[\s/>]/g)].map(
      (m) => m[1]
    );
  }

  it('inserts footnotePr before a pre-existing endnotePr and before compat', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string; writeText: (p: string, s: string) => void };
    zip.writeText(
      'word/settings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:zoom w:percent="100"/>` +
        `<w:endnotePr><w:pos w:val="docEnd"/></w:endnotePr>` +
        `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>` +
        `<w:rsids><w:rsid w:val="00A1B2C3"/></w:rsids>` +
        `</w:settings>`
    );
    executeFootnote(
      { action: 'footnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'ordered footnote' } as never,
      ast,
      parser
    );
    const order = settingsOrder(zip.readText('word/settings.xml'));
    expect(order).toEqual(['zoom', 'footnotePr', 'endnotePr', 'compat', 'rsids']);
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('rejects an empty settings part with a remedy instead of a TypeError', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as {
      readText: (p: string) => string;
      writeText: (p: string, s: string) => void;
      hasFile: (p: string) => boolean;
    };
    zip.writeText('word/settings.xml', '');
    expect(() =>
      executeFootnote(
        { action: 'footnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'never lands' } as never,
        ast,
        parser
      )
    ).toThrow(/Remedy/);
    expect(() =>
      executeFootnote(
        { action: 'footnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'never lands' } as never,
        ast,
        parser
      )
    ).not.toThrow(/TypeError/);
  });

  it('rejects an unprefixed settings root with a remedy instead of duplicating elements', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as { writeText: (p: string, s: string) => void };
    zip.writeText('word/settings.xml', `<settings><p>t</p></settings>`);
    expect(() =>
      executeEndnote(
        { action: 'endnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'never lands' } as never,
        ast,
        parser
      )
    ).toThrow(/Remedy/);
  });

  it('failed staging leaves content types, rels and relationships untouched', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as {
      readText: (p: string) => string;
      writeText: (p: string, s: string) => void;
      hasFile: (p: string) => boolean;
    };
    zip.writeText('word/settings.xml', '');
    const anchorNode = ast.sections[0].nodes[0];
    const anchorKidsBefore = Array.from(anchorNode.xml.childNodes).length;
    const versionBefore = anchorNode.version;
    const hadFootnotesBefore = zip.hasFile('word/footnotes.xml');
    const ctBefore = zip.hasFile('[Content_Types].xml') ? zip.readText('[Content_Types].xml') : null;
    const relsBefore = zip.hasFile('word/_rels/document.xml.rels')
      ? zip.readText('word/_rels/document.xml.rels')
      : null;
    const relKeysBefore = Object.keys(ast.relationships).sort();
    expect(() =>
      executeFootnote(
        { action: 'footnote', operation: 'insert', anchor: anchorNode.id, text: 'never lands' } as never,
        ast,
        parser
      )
    ).toThrow();
    if (ctBefore !== null) expect(zip.readText('[Content_Types].xml')).toBe(ctBefore);
    if (relsBefore !== null) expect(zip.readText('word/_rels/document.xml.rels')).toBe(relsBefore);
    expect(Object.keys(ast.relationships).sort()).toEqual(relKeysBefore);
    // Anchor paragraph clean: no reference run leaked into the live DOM, no
    // note part created, version unbumped.
    expect(Array.from(anchorNode.xml.childNodes).length).toBe(anchorKidsBefore);
    expect(anchorNode.xml.getElementsByTagName('w:footnoteReference').length).toBe(0);
    expect(anchorNode.version).toBe(versionBefore);
    expect(zip.hasFile('word/footnotes.xml')).toBe(hadFootnotesBefore);
    expect(parser.readDocumentXml()).not.toContain('w:footnoteReference');
  });

  it('failed staging leaves the endnote anchor paragraph clean', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as {
      writeText: (p: string, s: string) => void;
      hasFile: (p: string) => boolean;
    };
    zip.writeText('word/settings.xml', '');
    const anchorNode = ast.sections[0].nodes[0];
    const anchorKidsBefore = Array.from(anchorNode.xml.childNodes).length;
    const versionBefore = anchorNode.version;
    expect(() =>
      executeEndnote(
        { action: 'endnote', operation: 'insert', anchor: anchorNode.id, text: 'never lands' } as never,
        ast,
        parser
      )
    ).toThrow(/Remedy/);
    expect(Array.from(anchorNode.xml.childNodes).length).toBe(anchorKidsBefore);
    expect(anchorNode.xml.getElementsByTagName('w:endnoteReference').length).toBe(0);
    expect(anchorNode.version).toBe(versionBefore);
    expect(zip.hasFile('word/endnotes.xml')).toBe(false);
    expect(parser.readDocumentXml()).not.toContain('w:endnoteReference');
  });

  it('duplicate direct footnotePr siblings normalize to a single child', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string; writeText: (p: string, s: string) => void };
    zip.writeText(
      'word/settings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr>` +
        `<w:footnotePr><w:numFmt w:val="upperRoman"/></w:footnotePr>` +
        `<w:endnotePr><w:pos w:val="docEnd"/></w:endnotePr>` +
        `</w:settings>`
    );
    executeFootnote(
      { action: 'footnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'dedupe footnote' } as never,
      ast,
      parser
    );
    const settingsXml = zip.readText('word/settings.xml');
    expect((settingsXml.match(/<w:footnotePr[\s/>]/g) ?? []).length).toBe(1);
    expect((settingsXml.match(/<w:endnotePr[\s/>]/g) ?? []).length).toBe(1);
    // First sibling wins; ordering kept valid.
    expect(settingsXml).toContain('w:val="decimal"');
    expect(settingsXml.indexOf('w:footnotePr')).toBeLessThan(settingsXml.indexOf('w:endnotePr'));
    expect(parser.readDocumentXml()).toContain('w:footnoteReference');
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('duplicate direct endnotePr siblings normalize to a single child', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string; writeText: (p: string, s: string) => void };
    zip.writeText(
      'word/settings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr>` +
        `<w:endnotePr><w:numFmt w:val="lowerLetter"/></w:endnotePr>` +
        `</w:settings>`
    );
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'dedupe endnote' } as never,
      ast,
      parser
    );
    const settingsXml = zip.readText('word/settings.xml');
    expect((settingsXml.match(/<w:footnotePr[\s/>]/g) ?? []).length).toBe(1);
    expect((settingsXml.match(/<w:endnotePr[\s/>]/g) ?? []).length).toBe(1);
    expect(settingsXml).toContain('w:val="decimal"');
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });
});

describe('note numbering restarts once per document', () => {
  function propertyBody(settingsXml: string, tag: string): string {
    const m = settingsXml.match(new RegExp(`<w:${tag}>([\\s\\S]*?)</w:${tag}>`));
    expect(m).toBeTruthy();
    return m![1];
  }

  it('footnote and endnote settings declare continuous restart', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    executeFootnote(
      { action: 'footnote', operation: 'insert', anchor, text: 'restart footnote' } as never,
      ast,
      parser
    );
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor, text: 'restart endnote' } as never,
      ast,
      parser
    );
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string };
    const settingsXml = zip.readText('word/settings.xml');
    for (const tag of ['footnotePr', 'endnotePr']) {
      const body = propertyBody(settingsXml, tag);
      expect(body).toContain('<w:numRestart w:val="continuous"/>');
      expect((body.match(/w:numRestart/g) ?? []).length).toBe(1);
    }
    // Restart sorts last in (pos, numFmt, numStart, numRestart).
    const endBody = propertyBody(settingsXml, 'endnotePr');
    expect(endBody.indexOf('w:pos')).toBeLessThan(endBody.indexOf('w:numRestart'));
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('an explicit per-file restart choice is preserved, document markers stay plain numbers', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const zip = parser.getZipAdapter() as unknown as { readText: (p: string) => string; writeText: (p: string, s: string) => void };
    zip.writeText(
      'word/settings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:endnotePr><w:pos w:val="docEnd"/><w:numRestart w:val="eachSect"/></w:endnotePr>` +
        `</w:settings>`
    );
    executeFootnote(
      { action: 'footnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'kept footnote' } as never,
      ast,
      parser
    );
    executeEndnote(
      { action: 'endnote', operation: 'insert', anchor: ast.sections[0].nodes[0].id, text: 'kept endnote' } as never,
      ast,
      parser
    );
    const settingsXml = zip.readText('word/settings.xml');
    expect(propertyBody(settingsXml, 'endnotePr')).toContain('<w:numRestart w:val="eachSect"/>');
    expect(propertyBody(settingsXml, 'endnotePr')).not.toContain('w:val="continuous"');
    expect(propertyBody(settingsXml, 'footnotePr')).toContain('<w:numRestart w:val="continuous"/>');
    const docXml = parser.readDocumentXml();
    expect(docXml).toMatch(/<w:footnoteReference w:id="1"\/>/);
    expect(docXml).toMatch(/<w:endnoteReference w:id="1"\/>/);
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });
});

describe('strict gate accepts engine-honored routing fields', () => {
  it('accepts a footer paragraph insert with the full triple', () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'Footer note',
      target: 'footer',
      pageType: 'default',
      sectionIndex: 1,
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
    const batch = validateOperationsStrict([op]);
    expect(batch.ok).toBe(true);
    expect(batch.operations).toHaveLength(1);
  });

  it('accepts a footer insert with only the store set', () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'Footer note',
      target: 'footer',
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
  });

  it('accepts a body insert with explicit store values', () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'Body note',
      target: 'body',
      pageType: 'first',
      sectionIndex: 1,
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
  });

  it('accepts a field page-number insert routed to the footer', () => {
    const op = {
      action: 'field',
      operation: 'insertPageNumber',
      anchor: 'p_001',
      position: 'after',
      target: 'footer',
      pageType: 'default',
      sectionIndex: 1,
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
    const batch = validateOperationsStrict([op]);
    expect(batch.ok).toBe(true);
  });

  it('accepts a chart insert carrying a target selector', () => {
    const op = {
      action: 'chart',
      operation: 'insert',
      anchor: 'p_001',
      target: 'c_001',
      chartType: 'column',
      title: 'Revenue',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'ARR', values: [10, 25] }],
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
  });

  it('accepts a headerFooter set with its own triple', () => {
    const op = {
      action: 'headerFooter',
      operation: 'set',
      targetType: 'footer',
      pageType: 'default',
      sectionIndex: 1,
      text: 'Page footer',
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
  });

  it('rejects bad store values for insert and field', () => {
    const badTargets = [
      {
        action: 'insert',
        type: 'paragraph',
        anchor: 'p_001',
        position: 'after',
        text: 'hi',
        target: 'sidebar',
      },
      {
        action: 'insert',
        type: 'paragraph',
        anchor: 'p_001',
        position: 'after',
        text: 'hi',
        target: 'footer',
        pageType: 'middle',
      },
      {
        action: 'insert',
        type: 'paragraph',
        anchor: 'p_001',
        position: 'after',
        text: 'hi',
        target: 'footer',
        sectionIndex: -1,
      },
      {
        action: 'field',
        operation: 'insertPageNumber',
        target: 'sidebar',
      },
      {
        action: 'field',
        operation: 'insertPageNumber',
        target: 'footer',
        pageType: 'middle',
      },
      {
        action: 'field',
        operation: 'insertPageNumber',
        target: 'footer',
        sectionIndex: -2,
      },
      {
        action: 'field',
        operation: 'insertPageNumber',
        target: 'footer',
        sectionIndex: 0, // sections count from 1
      },
    ];
    for (const op of badTargets) {
      expect(OperationSchema.safeParse(op).success).toBe(false);
      expect(() => validateOperationStrict(op)).toThrow(/E_VALIDATION/);
    }
  });

  it('still rejects unknown keys on the touched shapes', () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'hi',
      target: 'footer',
      bogusKey: 'nope',
    };
    expect(OperationSchema.safeParse(op).success).toBe(false);
    expect(() => validateOperationStrict(op)).toThrow(/E_VALIDATION/);
    const fieldOp = {
      action: 'field',
      operation: 'insertPageNumber',
      target: 'footer',
      bogusKey: 1,
    };
    expect(OperationSchema.safeParse(fieldOp).success).toBe(false);
    expect(() => validateOperationStrict(fieldOp)).toThrow(/E_VALIDATION/);
  });

  it('accepts params-nested routing for insert and routes to the footer part', async () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'Footer via params',
      params: { target: 'footer', pageType: 'default', sectionIndex: 1 },
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
    const batch = validateOperationsStrict([op]);
    expect(batch.ok).toBe(true);
    expect(batch.operations).toHaveLength(1);

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const bodyCount = ast.sections[0].nodes.length;
    const anchor = ast.sections[0].nodes[0].id;
    const result = await engine.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'paragraph',
          anchor,
          position: 'after',
          text: 'Footer via params',
          params: { target: 'footer', pageType: 'default', sectionIndex: 1 },
        } as any,
      ],
    });
    expect(result.success).toBe(true);
    expect(ast.sections[0].nodes.length).toBe(bodyCount);
    expect(parser.readDocumentXml()).not.toContain('Footer via params');
    const footerKeys = Object.keys(ast.footers);
    expect(footerKeys.length).toBeGreaterThan(0);
    const footerText = footerKeys
      .map((k) => ((ast.footers[k].xml as unknown as Element).textContent ?? ''))
      .join('');
    expect(footerText).toContain('Footer via params');

    const withUnknown = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'hi',
      params: { target: 'footer' },
      bogusKey: 'nope',
    };
    expect(OperationSchema.safeParse(withUnknown).success).toBe(false);
    expect(() => validateOperationStrict(withUnknown)).toThrow(/E_VALIDATION/);
  });

  it('flat plus params conflict resolves to the nested value', async () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'Conflict nested wins',
      target: 'body',
      params: { target: 'footer', pageType: 'default', sectionIndex: 1 },
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const bodyCount = ast.sections[0].nodes.length;
    const anchor = ast.sections[0].nodes[0].id;
    const result = await engine.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'paragraph',
          anchor,
          position: 'after',
          text: 'Conflict nested wins',
          target: 'body',
          params: { target: 'footer', pageType: 'default', sectionIndex: 1 },
        } as any,
      ],
    });
    expect(result.success).toBe(true);
    expect(ast.sections[0].nodes.length).toBe(bodyCount);
    expect(parser.readDocumentXml()).not.toContain('Conflict nested wins');
    const footerKeys = Object.keys(ast.footers);
    expect(footerKeys.length).toBeGreaterThan(0);
    const footerText = footerKeys
      .map((k) => ((ast.footers[k].xml as unknown as Element).textContent ?? ''))
      .join('');
    expect(footerText).toContain('Conflict nested wins');
  });

  it('bad params-nested routing values fail closed at the engine', async () => {
    const badNested = [
      { target: 'sidebar' },
      { target: 'footer', pageType: 'middle' },
      { target: 'footer', sectionIndex: -1 },
      { target: 'footer', sectionIndex: 0 },
    ];
    for (const params of badNested) {
      const op = {
        action: 'insert',
        type: 'paragraph',
        anchor: 'p_001',
        position: 'after',
        text: 'never lands',
        params,
      };
      expect(OperationSchema.safeParse(op).success).toBe(true);
      expect(() => validateOperationStrict(op)).not.toThrow();

      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);
      const bodyCount = ast.sections[0].nodes.length;
      const anchor = ast.sections[0].nodes[0].id;
      await expect(
        engine.execute({
          action: 'insert',
          type: 'paragraph',
          anchor,
          position: 'after',
          text: 'never lands',
          params,
        } as any)
      ).rejects.toThrow(/E_VALIDATION/);
      expect(ast.sections[0].nodes.length).toBe(bodyCount);
      expect(parser.readDocumentXml()).not.toContain('never lands');
    }
  });
});

describe('insert without anchor for header/footer targets', () => {
  it('parses a header insert with no anchor or position', () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      text: 'Header note',
      target: 'header',
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
    const batch = validateOperationsStrict([op]);
    expect(batch.ok).toBe(true);
    expect(batch.operations).toHaveLength(1);
  });

  it('parses a footer insert with no anchor or position', () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      text: 'Footer note',
      target: 'footer',
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
    const batch = validateOperationsStrict([op]);
    expect(batch.ok).toBe(true);
    expect(batch.operations).toHaveLength(1);
  });

  it('routes an anchorless footer insert to the footer part, body untouched', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const bodyCount = ast.sections[0].nodes.length;
    const op = {
      action: 'insert',
      type: 'paragraph',
      text: 'Anchorless footer line',
      target: 'footer',
    };
    expect(() => validateOperationStrict(op)).not.toThrow();
    const result = await engine.executeBatch({ operations: [op as any] });
    expect(result.success).toBe(true);
    expect(ast.sections[0].nodes.length).toBe(bodyCount);
    expect(parser.readDocumentXml()).not.toContain('Anchorless footer line');
    const footerKeys = Object.keys(ast.footers);
    expect(footerKeys.length).toBeGreaterThan(0);
    const footerText = footerKeys
      .map((k) => ((ast.footers[k].xml as unknown as Element).textContent ?? ''))
      .join('');
    expect(footerText).toContain('Anchorless footer line');
  });

  it('routes an anchorless header insert to the header part, body untouched', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const bodyCount = ast.sections[0].nodes.length;
    const op = {
      action: 'insert',
      type: 'paragraph',
      text: 'Anchorless header line',
      target: 'header',
    };
    expect(() => validateOperationStrict(op)).not.toThrow();
    const result = await engine.executeBatch({ operations: [op as any] });
    expect(result.success).toBe(true);
    expect(ast.sections[0].nodes.length).toBe(bodyCount);
    expect(parser.readDocumentXml()).not.toContain('Anchorless header line');
    const headerKeys = Object.keys(ast.headers);
    expect(headerKeys.length).toBeGreaterThan(0);
    const headerText = headerKeys
      .map((k) => ((ast.headers[k].xml as unknown as Element).textContent ?? ''))
      .join('');
    expect(headerText).toContain('Anchorless header line');
  });

  it('parses params-nested footer routing without anchor and routes it', async () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      text: 'Nested footer line',
      params: { target: 'footer' },
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
    const batch = validateOperationsStrict([op]);
    expect(batch.ok).toBe(true);

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const bodyCount = ast.sections[0].nodes.length;
    const result = await engine.executeBatch({ operations: [op as any] });
    expect(result.success).toBe(true);
    expect(ast.sections[0].nodes.length).toBe(bodyCount);
    expect(parser.readDocumentXml()).not.toContain('Nested footer line');
    const footerText = Object.keys(ast.footers)
      .map((k) => ((ast.footers[k].xml as unknown as Element).textContent ?? ''))
      .join('');
    expect(footerText).toContain('Nested footer line');
  });

  it('still rejects body inserts missing anchor or position', () => {
    const noAnchor = {
      action: 'insert',
      type: 'paragraph',
      position: 'after',
      text: 'Body orphan',
    };
    expect(() => validateOperationStrict(noAnchor)).toThrow(/E_VALIDATION/);
    expect(() => validateOperationsStrict([noAnchor])).toThrow(/E_VALIDATION/);

    const noPosition = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      text: 'Body orphan',
    };
    expect(() => validateOperationStrict(noPosition)).toThrow(/E_VALIDATION/);
    expect(() => validateOperationsStrict([noPosition])).toThrow(/E_VALIDATION/);

    const explicitBody = {
      action: 'insert',
      type: 'paragraph',
      text: 'Body orphan',
      target: 'body',
    };
    expect(() => validateOperationStrict(explicitBody)).toThrow(/E_VALIDATION/);
    expect(() => validateOperationsStrict([explicitBody])).toThrow(/E_VALIDATION/);

    const nestedBody = {
      action: 'insert',
      type: 'paragraph',
      text: 'Body orphan',
      params: { target: 'body' },
    };
    expect(() => validateOperationStrict(nestedBody)).toThrow(/E_VALIDATION/);
    expect(() => validateOperationsStrict([nestedBody])).toThrow(/E_VALIDATION/);

    const emptyAnchor = {
      action: 'insert',
      type: 'paragraph',
      anchor: '',
      position: 'after',
      text: 'Body orphan',
    };
    expect(OperationSchema.safeParse(emptyAnchor).success).toBe(false);
    expect(() => validateOperationStrict(emptyAnchor)).toThrow(/E_VALIDATION/);
  });

  it('keeps header/footer inserts with anchor working', () => {
    const op = {
      action: 'insert',
      type: 'paragraph',
      anchor: 'p_001',
      position: 'after',
      text: 'Footer note',
      target: 'footer',
    };
    expect(OperationSchema.safeParse(op).success).toBe(true);
    expect(() => validateOperationStrict(op)).not.toThrow();
    const batch = validateOperationsStrict([op]);
    expect(batch.ok).toBe(true);
    expect(batch.operations).toHaveLength(1);
  });

  it('requires integer dimensions for image resize and legacy insert', () => {
    expect(() =>
      validateOperationStrict({
        action: 'image',
        operation: 'resize',
        target: 'x',
        width: 500000.75,
        height: 914400,
      })
    ).toThrow(/width/);

    expect(() =>
      validateOperationStrict({
        action: 'image',
        operation: 'resize',
        target: 'x',
        width: 914400,
        height: 914400,
      })
    ).not.toThrow();

    expect(() =>
      validateOperationStrict({
        action: 'insert',
        type: 'image',
        anchor: 'x',
        fileData: 'data:image/png;base64,AAAA',
        widthEmu: 1.5,
      })
    ).toThrow(/E_VALIDATION/);
  });

  it('honors legacy widthEmu and heightEmu on insert type:image', async () => {
    const png1x1 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

    const parser1 = DocxParser.createNewDocument('blank');
    const ast1 = await parser1.getAST();
    const engine1 = new OperationEngine(ast1, parser1);
    const anchor1 = ast1.sections[0].nodes[0].id;
    const res1 = await engine1.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'image',
          anchor: anchor1,
          position: 'after',
          fileData: png1x1,
          widthEmu: 1828800,
          heightEmu: 914400,
        } as any,
      ],
      sequential: true,
    });
    expect(res1.success).toBe(true);
    expect(parser1.readDocumentXml()).toContain('cx="1828800" cy="914400"');

    const parser2 = DocxParser.createNewDocument('blank');
    const ast2 = await parser2.getAST();
    const engine2 = new OperationEngine(ast2, parser2);
    const anchor2 = ast2.sections[0].nodes[0].id;
    const res2 = await engine2.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'image',
          anchor: anchor2,
          position: 'after',
          fileData: png1x1,
          width: 1828800,
          widthEmu: 1828800,
          height: 914400,
        } as any,
      ],
      sequential: true,
    });
    expect(res2.success).toBe(true);

    const parser3 = DocxParser.createNewDocument('blank');
    const ast3 = await parser3.getAST();
    const engine3 = new OperationEngine(ast3, parser3);
    const anchor3 = ast3.sections[0].nodes[0].id;
    const res3 = await engine3.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'image',
          anchor: anchor3,
          position: 'after',
          fileData: png1x1,
          width: 1000000,
          widthEmu: 1828800,
          height: 914400,
        } as any,
      ],
      sequential: true,
    });
    expect(res3.success).toBe(false);
    expect(res3.errors[0]?.error).toContain('E_VALIDATION');

    const parser4 = DocxParser.createNewDocument('blank');
    const ast4 = await parser4.getAST();
    const engine4 = new OperationEngine(ast4, parser4);
    const anchor4 = ast4.sections[0].nodes[0].id;
    const res4 = await engine4.executeBatch({
      operations: [
        {
          action: 'image',
          operation: 'insert',
          anchor: anchor4,
          fileData: png1x1,
          width: 1828800,
          height: 914400,
        } as any,
      ],
      sequential: true,
    });
    expect(res4.success).toBe(true);
    expect(parser4.readDocumentXml()).toContain('cx="1828800" cy="914400"');
  });
});
