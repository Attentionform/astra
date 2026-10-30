/**
 * Phase B regression suite — one test per UNIFIED_FIX_PLAN §4 item (B-1..B-7).
 * Each test asserts on produced MARKUP / reload TYPES, never just appliedCount.
 */
import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  Chart3DManipulator,
  DocxParser,
  OperationEngine,
  createBlankDocxBuffer,
  executeEquation,
  executeHyperlink,
  executeImage,
  executeInsert,
  executeReplace,
  executeTable,
  replaceText,
  validateOperationStrict,
  OperationSchema,
} from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const serializer = new XMLSerializer();

function mockTableAST(): any {
  const xmlStr = `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
  const tblEl = new DOMParser().parseFromString(xmlStr, 'text/xml').documentElement;
  return {
    metadata: {},
    sections: [{
      id: 'sec_root', heading: null,
      nodes: [{
        id: 't_001', type: 'table', xml: tblEl,
        metadata: { text: 'H1 H2 C1 C2', tableDimensions: { rows: 2, cols: 2 } },
        version: 1, checksum: 'chk',
      }],
      pageSetup: {}, version: 1, checksum: 'sec',
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'phase-b' },
    nextChangeId: 1,
  };
}

describe('Phase B fixes (B-1..B-7)', () => {
  it('B-1: flat mergeCells resolves like params-nested (writes gridSpan/vMerge)', async () => {
    const ast = mockTableAST();
    await executeTable({
      action: 'table', target: 't_001', operation: 'mergeCells',
      fromRow: 1, fromCol: 1, toRow: 2, toCol: 2,
    } as any, ast);
    const xml = serializer.serializeToString(ast.sections[0].nodes[0].xml);
    expect(xml).toContain('w:hMerge');
    expect(xml).toContain('w:val="restart"');
    expect(xml).toContain('w:val="continue"');
    expect(xml).toContain('w:vMerge');
  });

  it('B-1 gate: batch that empties a table fails loud + rolls back (never success=true)', async () => {
    const ast = mockTableAST();
    const engine = new OperationEngine(ast);
    const result = await engine.executeBatch({
      operations: [
        { action: 'table', target: 't_001', params: { op: 'deleteRow', index: 1 } },
        { action: 'table', target: 't_001', params: { op: 'deleteRow', index: 1 } },
      ] as any,
      sequential: true,
    });
    expect(result.success).toBe(false);
    expect(result.rollbackTriggered).toBe(true);
    expect(result.errors.some(e => String(e.error).includes('structural gate'))).toBe(true);
    const rows = ast.sections[0].nodes[0].xml.getElementsByTagName('w:tr').length;
    expect(rows).toBe(2);
  });

  it('B-2: sectionBreak insert emits sectPr break + splits sections (breakType in schema)', async () => {
    const parsed = validateOperationStrict({
      action: 'insert', type: 'sectionBreak', anchor: 'p_1', position: 'after', breakType: 'continuous',
    });
    expect(parsed.ok).toBe(true);

    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const before = ast.sections.length;
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType: 'continuous' } as any, ast, parser);
    expect(ast.sections.length).toBe(before + 1);
    const breakXml = serializer.serializeToString(ast.sections[0].nodes[1].xml);
    expect(breakXml).toContain('w:sectPr');
    // The break type describes how the following section starts.
    const sps = parser.getSectPrElements();
    expect(serializer.serializeToString(sps[sps.length - 1])).toContain('w:val="continuous"');

    // Default break type is nextPage.
    const parser2 = new DocxParser(createBlankDocxBuffer('blank'));
    const ast2: any = await parser2.getAST();
    const anchor2 = ast2.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor: anchor2, position: 'after' } as any, ast2, parser2);
    // The break type describes how the following section starts.
    const sps2 = parser2.getSectPrElements();
    expect(serializer.serializeToString(sps2[sps2.length - 1])).toContain('w:val="nextPage"');
  });

  it('B-2 gate: bogus breakType fails loudly (schema + engine); all 5 enum values accepted', async () => {
    // Schema gate rejects non-enum values.
    for (const bogus of ['foobar-bogus', '', 'PAGE']) {
      const op = { action: 'insert', type: 'sectionBreak', anchor: 'p_1', position: 'after', breakType: bogus };
      expect(OperationSchema.safeParse(op).success).toBe(false);
      expect(() => validateOperationStrict(op)).toThrow(/E_VALIDATION/);
    }

    // Engine fails loudly on bogus values via both top-level and params shapes.
    for (const badOp of [
      { action: 'insert', type: 'sectionBreak', anchor: '', position: 'after', breakType: 'foobar-bogus' },
      { action: 'insert', type: 'sectionBreak', anchor: '', position: 'after', params: { breakType: 'foobar-bogus' } },
    ] as any[]) {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast: any = await parser.getAST();
      badOp.anchor = ast.sections[0].nodes[0].id;
      await expect(executeInsert(badOp, ast, parser)).rejects.toThrow(/E_VALIDATION.*breakType/);
    }

    // All 5 ST_SectionMark values pass the schema gate and land in w:val.
    for (const breakType of ['nextPage', 'nextColumn', 'continuous', 'evenPage', 'oddPage']) {
      expect(validateOperationStrict({
        action: 'insert', type: 'sectionBreak', anchor: 'p_1', position: 'after', breakType,
      }).ok).toBe(true);
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast: any = await parser.getAST();
      const anchor = ast.sections[0].nodes[0].id;
      await executeInsert({ action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType } as any, ast, parser);
      // The break type describes how the following section starts.
      const sps = parser.getSectPrElements();
      expect(serializer.serializeToString(sps[sps.length - 1])).toContain(`w:val="${breakType}"`);
    }
  });

  it('B-3: bookmark-aware replace keeps the span wrapping the new text (offsets)', () => {
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}"><w:bookmarkStart w:id="7" w:name="bm_phaseb"/><w:r><w:t>old</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>`,
      'text/xml',
    ).documentElement;
    const ast: any = {
      metadata: {},
      sections: [{
        id: 'sec_root', heading: null,
        nodes: [{
          id: 'p_bm', type: 'paragraph', xml: p,
          metadata: { text: 'old', hasCitation: false, hasComment: false, hasBookmark: true, hasTrackedChanges: false, hasEquation: false, hasField: false, hasImage: false, hasHyperlink: false },
          version: 1, checksum: 'c',
        }],
        pageSetup: {}, version: 1, checksum: 's',
      }],
      styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {}, headers: {}, footers: {}, mediaFiles: {},
      settings: { trackChanges: false, defaultAuthor: 'phase-b' }, nextChangeId: 1,
    };
    executeReplace({ action: 'replace', target: 'p_bm', text: 'brand new text' } as any, ast);
    const xml = serializer.serializeToString(ast.sections[0].nodes[0].xml);
    const iStart = xml.indexOf('bookmarkStart');
    const iText = xml.indexOf('brand new text');
    const iEnd = xml.indexOf('bookmarkEnd');
    expect(iStart).toBeGreaterThanOrEqual(0);
    expect(iText).toBeGreaterThanOrEqual(0);
    expect(iEnd).toBeGreaterThanOrEqual(0);
    expect(iStart < iText && iText < iEnd).toBe(true);
  });

  it('B-3b: bookmark-aware replace with nested bookmark inside w:hyperlink preserves hyperlink and keeps run inside span', () => {
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}"><w:bookmarkStart w:id="12" w:name="bm_nested"/><w:hyperlink w:anchor="bm_target"><w:r><w:t>old linked text</w:t></w:r><w:bookmarkEnd w:id="12"/></w:hyperlink></w:p>`,
      'text/xml',
    ).documentElement;
    const ast: any = {
      metadata: {},
      sections: [{
        id: 'sec_root', heading: null,
        nodes: [{
          id: 'p_bm_nested', type: 'paragraph', xml: p,
          metadata: { text: 'old linked text', hasCitation: false, hasComment: false, hasBookmark: true, hasTrackedChanges: false, hasEquation: false, hasField: false, hasImage: false, hasHyperlink: true },
          version: 1, checksum: 'c',
        }],
        pageSetup: {}, version: 1, checksum: 's',
      }],
      styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {}, headers: {}, footers: {}, mediaFiles: {},
      settings: { trackChanges: false, defaultAuthor: 'phase-b' }, nextChangeId: 1,
    };
    executeReplace({ action: 'replace', target: 'p_bm_nested', text: 'new updated text' } as any, ast);
    const xml = serializer.serializeToString(ast.sections[0].nodes[0].xml);
    const iStart = xml.indexOf('bookmarkStart');
    const iText = xml.indexOf('new updated text');
    const iEnd = xml.indexOf('bookmarkEnd');
    expect(iStart).toBeGreaterThanOrEqual(0);
    expect(iText).toBeGreaterThanOrEqual(0);
    expect(iEnd).toBeGreaterThanOrEqual(0);
    expect(iStart < iText && iText < iEnd).toBe(true);

    const hyperlinks = ast.sections[0].nodes[0].xml.getElementsByTagName('w:hyperlink');
    expect(hyperlinks.length).toBeGreaterThanOrEqual(1);
    expect(hyperlinks[0].getElementsByTagName('w:r').length).toBeGreaterThanOrEqual(1);
    expect(hyperlinks[0].textContent).toContain('new updated text');
  });

  it('B-3c: malformed document with bookmarkEnd inside w:r falls back safely to appendChild without creating nested w:r', () => {
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}"><w:bookmarkStart w:id="13" w:name="bm_malformed"/><w:r><w:t>malformed run</w:t><w:bookmarkEnd w:id="13"/></w:r></w:p>`,
      'text/xml',
    ).documentElement;

    replaceText(p, 'safe text', false);

    const runs = Array.from(p.getElementsByTagName('w:r'));
    expect(runs.length).toBeGreaterThanOrEqual(1);
    for (const r of runs) {
      expect((r as Element).getElementsByTagName('w:r').length).toBe(0);
    }
  });

  it('B-4: internal hyperlink emits w:anchor with styled run and NO External rel', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    return parser.getAST().then((ast: any) => {
      const anchor = ast.sections[0].nodes[0].id;
      const extBefore = Object.values(ast.relationships).filter((r: any) => r.targetMode === 'External').length;
      executeHyperlink({ action: 'hyperlink', target: anchor, text: 'jump inside', bookmarkName: 'bm_phaseb' } as any, ast, parser);
      const node = ast.sections[0].nodes.find((n: any) => n.id === anchor);
      const links = node.xml.getElementsByTagName('w:hyperlink');
      expect(links.length).toBeGreaterThan(0);
      const link = links[links.length - 1] as Element;
      expect(link.getAttribute('w:anchor')).toBe('bm_phaseb');
      expect(link.getAttribute('w:history')).toBe('1');
      expect(link.getAttribute('r:id')).toBeFalsy();
      const extAfter = Object.values(ast.relationships).filter((r: any) => r.targetMode === 'External').length;
      expect(extAfter).toBe(extBefore);
      // Styled run preserved.
      expect(serializer.serializeToString(link)).toContain('w:val="0066CC"');
    });
  });

  it('B-5: chart XML carries c:numCache; line3D has spPr stroke; pie3D has fills + varyColors', () => {
    const ast: any = {
      metadata: {}, sections: [], styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {}, headers: {}, footers: {}, mediaFiles: {},
      settings: { trackChanges: true, defaultAuthor: 'phase-b' }, nextChangeId: 1,
    };
    const m = new Chart3DManipulator(ast);
    const opts = { categories: ['Q1', 'Q2'], series: [{ name: 'S', values: [10, 20], color: '0066CC' }] } as any;

    const bar = m.generate3DChartXml({ ...opts, chartType: 'bar3D' });
    expect(bar).toContain('c:numCache');
    expect(bar).toContain('<c:v>10</c:v>');

    const line = m.generate3DChartXml({ ...opts, chartType: 'line3D' });
    expect(line).toContain('c:numCache');
    expect(line).toContain('<a:ln');

    const pie = m.generate3DChartXml({ ...opts, chartType: 'pie3D' });
    expect(pie).toContain('c:numCache');
    expect(pie).toContain('<c:varyColors val="1"/>');
    expect(pie).toContain('<c:dPt>');
    expect(pie).toContain('c:spPr');

    // All three stay well-formed XML.
    for (const xml of [bar, line, pie]) {
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      expect(doc.getElementsByTagName('parsererror').length).toBe(0);
    }
  });

  it('B-6: table create/addRow carry cantSplit on all rows + tblHeader on the header row', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'table', anchor, position: 'after',
      headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']],
    } as any, ast, parser);
    const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
    expect(tblNode).toBeDefined();
    const rows = tblNode.xml.getElementsByTagName('w:tr');
    expect(rows.length).toBe(3);
    for (let i = 0; i < rows.length; i++) {
      const trPr = (rows[i] as Element).getElementsByTagName('w:trPr')[0];
      expect(trPr, `row ${i} has trPr`).toBeDefined();
      expect(trPr.getElementsByTagName('w:cantSplit').length, `row ${i} cantSplit`).toBe(1);
    }
    expect((rows[0] as Element).getElementsByTagName('w:tblHeader').length).toBe(1);

    // addRow clones the header row here — the clone must keep cantSplit but
    // must NOT duplicate tblHeader.
    await executeTable({ action: 'table', target: tblNode.id, params: { op: 'addRow', afterIndex: 1 } } as any, ast);
    const rowsAfter = tblNode.xml.getElementsByTagName('w:tr');
    expect(rowsAfter.length).toBe(4);
    for (let i = 0; i < rowsAfter.length; i++) {
      expect((rowsAfter[i] as Element).getElementsByTagName('w:cantSplit').length, `row ${i} cantSplit after addRow`).toBe(1);
    }
    expect(tblNode.xml.getElementsByTagName('w:tblHeader').length).toBe(1);
  });

  it('B-7: reload types match pre-save (image stays image, equation stays equation)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast1: any = await parser.getAST();
    const anchor = ast1.sections[0].nodes[0].id;
    const eq = await executeEquation({ action: 'equation', operation: 'insert', anchor, position: 'after', latex: 'x^2', id: 'eq_phaseb' } as any, ast1, parser);
    const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    executeImage({ action: 'image', operation: 'insert', anchor, fileData: png1x1, id: 'img_phaseb' } as any, ast1, parser);

    const preEq = ast1.sections[0].nodes.find((n: any) => n.id === (eq as any).id);
    const preImg = ast1.sections.flatMap((s: any) => s.nodes).find((n: any) => n.id === 'img_phaseb');
    expect(preEq.type).toBe('equation');
    expect(preImg.type).toBe('image');

    const buf = parser.rezip(ast1);
    const reload = new DocxParser(buf);
    const ast2: any = await reload.getAST();
    const all = ast2.sections.flatMap((s: any) => s.nodes);
    const reEq = all.find((n: any) => serializer.serializeToString(n.xml).includes('m:oMath'));
    const reImg = all.find((n: any) => {
      const x = serializer.serializeToString(n.xml);
      return x.includes('w:drawing') && x.includes('a:blip');
    });
    expect(reEq).toBeDefined();
    expect(reImg).toBeDefined();
    expect(reEq.type).toBe(preEq.type);
    expect(reImg.type).toBe(preImg.type);
  });

  it('B-7 guard: shape-only drawing stays paragraph; text+image stays paragraph on reload', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    // Shape-only drawing: w:drawing present but NO a:blip (e.g. a floating
    // shape/textbox), no text. Must not classify as image.
    const shapeXml =
      `<w:p xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      `<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      `<wps:wsp><wps:cNvSpPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    // Text + real picture: has w:t text AND drawing+blip. The empty-text
    // guard keeps it a paragraph (deliberate: mixed content is not a pure image).
    const textImgXml =
      `<w:p xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:r><w:t>Captioned figure</w:t></w:r>` +
      `<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    parser.insertRootNode(shapeXml, undefined, 'append', 'shape_only_phaseb');
    parser.insertRootNode(textImgXml, undefined, 'append', 'textimg_phaseb');

    const buf = parser.rezip();
    const reload = new DocxParser(buf);
    const ast2: any = await reload.getAST();
    const all = ast2.sections.flatMap((s: any) => s.nodes);
    // NOTE: reload re-keys paragraphs by w14:paraId (not the insert-time
    // custom id), so locate by content markers instead of id.
    const reShape = all.find((n: any) => serializer.serializeToString(n.xml).includes('wordprocessingShape'));
    const reTextImg = all.find((n: any) => serializer.serializeToString(n.xml).includes('Captioned figure'));
    expect(reShape).toBeDefined();
    expect(reTextImg).toBeDefined();
    expect(serializer.serializeToString(reShape.xml)).toContain('w:drawing');
    expect(serializer.serializeToString(reShape.xml)).not.toContain('a:blip');
    expect(serializer.serializeToString(reTextImg.xml)).toContain('a:blip');
    expect(reShape.type).toBe('paragraph');
    expect(reTextImg.type).toBe('paragraph');
  });
});
