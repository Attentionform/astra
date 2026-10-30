import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';
import {
  Chart3DManipulator,
  DocxParser,
  DocumentAST,
  executeImage,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

// P1 (Word text-overlay fix): chart/image drawing paragraphs must carry a
// centered, auto-spaced w:pPr so Word expands the line box for the inline
// graphic even when the Normal style uses exact line spacing.

function makeMinimalAst(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body><w:p><w:r><w:t>anchor</w:t></w:r></w:p></w:body>
    </w:document>`,
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
          text: 'anchor', hasCitation: false, hasComment: false, hasBookmark: false,
          hasTrackedChanges: false, hasEquation: false, hasField: false,
          hasImage: false, hasHyperlink: false,
        },
        version: 1,
        checksum: 'chk_1',
      }],
      pageSetup: null,
      version: 1,
      checksum: 'sec_1',
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

// Give the live Normal style an exact-spaced pPr (the Word overlay trigger).
function makeNormalStyleExact(parser: DocxParser): void {
  const stylesDom = parser.getStylesDom();
  if (!stylesDom) throw new Error('styles DOM unavailable');
  const styles = Array.from(stylesDom.getElementsByTagName('w:style'));
  const normal = styles.find((s) => (s as Element).getAttribute('w:styleId') === 'Normal') as Element | undefined;
  if (!normal) throw new Error('Normal style not found');
  const pPr = stylesDom.createElement('w:pPr');
  const spacing = stylesDom.createElement('w:spacing');
  spacing.setAttribute('w:line', '276');
  spacing.setAttribute('w:lineRule', 'exact');
  pPr.appendChild(spacing);
  normal.appendChild(pPr);
}

// Serialize every w:p ancestor of a w:drawing in document.xml.
function getDrawingParagraphs(docxBuf: Buffer): string[] {
  const zip = new AdmZip(docxBuf);
  const entry = zip.getEntry('word/document.xml');
  if (!entry) throw new Error('word/document.xml missing from rezipped docx');
  const dom = new DOMParser().parseFromString(entry.getData().toString('utf8'), 'text/xml');
  const drawings = dom.getElementsByTagName('w:drawing');
  const out: string[] = [];
  const serializer = new XMLSerializer();
  for (let i = 0; i < drawings.length; i++) {
    let el: any = drawings[i];
    while (el && el.nodeName !== 'w:p') el = el.parentNode;
    if (!el) throw new Error('w:drawing without ancestor w:p');
    out.push(serializer.serializeToString(el));
  }
  return out;
}

describe('P1 chart/image drawing paragraph layout (Word overlay fix)', () => {
  it('chart drawing paragraph carries centered auto w:pPr', () => {
    const m = new Chart3DManipulator(makeMinimalAst());
    const xml = m.generateDrawingEmbedXml('rIdC1', 1, 'Chart 1', 5486400, 3200400);
    expect(xml).toContain('<w:pPr>');
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('w:lineRule="auto"');
    expect(xml).toContain('<w:spacing w:before="140" w:after="140" w:line="240" w:lineRule="auto"/>');
  });

  it('chart drawing defaults unchanged (widthEmu 5486400, heightEmu passthrough)', () => {
    const m = new Chart3DManipulator(makeMinimalAst());
    const xml = m.generateDrawingEmbedXml('rIdC1');
    expect(xml).toContain('<wp:extent cx="5486400" cy="3200400"/>');
  });

  it('chart in exact-spaced Normal doc validates and drawing paragraph carries auto rule', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const m = new Chart3DManipulator(ast, parser);
    m.insert3DChart({
      chartType: 'bar3D',
      title: 'P1 Layout',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'S', values: [1, 2], color: '0066CC' }],
    });
    makeNormalStyleExact(parser);
    const buf = parser.rezip();

    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);

    const zip = new AdmZip(buf);
    const stylesXml = zip.getEntry('word/styles.xml')?.getData().toString('utf8') ?? '';
    expect(stylesXml).toContain('w:lineRule="exact"');

    const paras = getDrawingParagraphs(buf);
    expect(paras.length).toBeGreaterThan(0);
    for (const p of paras) {
      expect(p).toContain('<w:pPr>');
      expect(p).toContain('w:lineRule="auto"');
      expect(p).toContain('<w:jc w:val="center"/>');
    }
  });

  it('image drawing paragraph likewise carries centered auto w:pPr and validates', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const anchorId = ast.sections[0].nodes[0].id;
    const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    executeImage(
      {
        action: 'image',
        operation: 'insert',
        anchor: anchorId,
        fileData: `data:image/png;base64,${png1x1}`,
        width: 2743200,
        height: 1828800,
      },
      ast,
      parser
    );
    makeNormalStyleExact(parser);
    const buf = parser.rezip();

    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);

    const paras = getDrawingParagraphs(buf);
    expect(paras.length).toBeGreaterThan(0);
    for (const p of paras) {
      expect(p).toContain('<w:pPr>');
      expect(p).toContain('w:lineRule="auto"');
      expect(p).toContain('<w:jc w:val="center"/>');
    }
  });
});
