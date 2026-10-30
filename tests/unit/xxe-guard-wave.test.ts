import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  executeEquation,
  executeChart,
  executeChartUpdate,
  HeaderFooterManipulator,
  NumberingManipulator,
  parseCorePropertiesXml,
  mergeCorePropertiesXml,
  extractStylePalette,
  sanitizeDocumentXml,
  sanitizeDocxBuffer,
  assertNoForbiddenMarkup,
  addComment,
  sanitizeXml,
  validateTocContainerXml,
  canonicalizePartXml,
  convertAstToMarkdown,
  executeFootnote,
  executeHyperlink,
  executeImage,
  Chart2DManipulator,
  Chart3DManipulator,
} from '@attentionform/astra-engine';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  assertNoXxe,
  containsForbiddenMarkup,
} from '../../packages/engine/src/ooxml/xxe-guard';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Every declaration spelling the central guard must catch.
const EVIL_DECLARATIONS = [
  '<!DOCTYPE foo>',
  '<!doctype foo>',
  '<!DoCtYpE foo>',
  '<!dOcTyPe foo SYSTEM "http://evil.example/x">',
  '<!ENTITY xxe "pwn">',
  '<!entity xxe "pwn">',
  '<!EnTiTy xxe SYSTEM "file:///etc/passwd">',
  '<!\nDOCTYPE foo>',
  '<!\tENTITY xxe "pwn">',
  '<! ENTITY xxe "pwn">',
  '<!\r\nDOCTYPE foo>',
  '<!  DOCTYPE foo>',
  '<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>',
  '<!\x00DOCTYPE foo>',
];

const INERT_ENCODED_FORMS = [
  '&lt;!DOCTYPE foo>',
  '&lt;!ENTITY xxe "pwn">',
  '&#60;!DOCTYPE foo>',
  '&#x3c;!ENTITY xxe "pwn">',
  '%3C!DOCTYPE foo>',
  'profit &lt; loss and a &amp; b',
  'plain math: x &lt; y, y &gt; z',
];

const LEGIT_XML = [
  '<w:document xmlns:w="' + W_NS + '"><w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body></w:document>',
  '<!-- a plain comment carries no declaration keyword -->',
  '<w:t xmlns:w="' + W_NS + '">AT&amp;T &lt;3 math &amp; prose</w:t>',
  '<![CDATA[plain cdata without keywords]]>',
  '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="' + W_NS + '"/>',
];

function evilDocXml(decl: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    decl +
    '<w:document xmlns:w="' + W_NS + '"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>'
  );
}

function plantInLiveZip(parser: DocxParser, partPath: string, content: string): void {
  const zip = (parser as unknown as { zip: { getEntry(n: string): unknown; addFile(n: string, b: Buffer): void; updateFile(n: string, b: Buffer): void } }).zip;
  const buf = Buffer.from(content, 'utf8');
  if (zip.getEntry(partPath)) zip.updateFile(partPath, buf);
  else zip.addFile(partPath, buf);
}

describe('central XXE/DTD guard', () => {
  it('flags every case variant of DOCTYPE and ENTITY', () => {
    for (const decl of EVIL_DECLARATIONS) {
      expect(containsForbiddenMarkup(decl), decl).toBe(true);
    }
  });

  it('flags tab/newline-split evasions between markup open and keyword', () => {
    expect(containsForbiddenMarkup('<!\nDOCTYPE r>')).toBe(true);
    expect(containsForbiddenMarkup('<!\tENTITY x "y">')).toBe(true);
    expect(containsForbiddenMarkup('<!\r\nDoCtYpE r>')).toBe(true);
    expect(containsForbiddenMarkup('<! ENTITY x "y">')).toBe(true);
  });

  it('fails closed with E_VALIDATION plus a remedy on every variant', () => {
    for (const decl of EVIL_DECLARATIONS) {
      expect(() => assertNoXxe(decl, 'probe'), decl).toThrowError(/E_VALIDATION/);
      expect(() => assertNoXxe(decl, 'probe'), decl).toThrowError(/Remedy/);
      expect(() => assertNoForbiddenMarkup(decl, 'probe'), decl).toThrowError(/E_VALIDATION/);
    }
  });

  it('leaves inert encoded forms alone (never decodes before scanning)', () => {
    for (const inert of INERT_ENCODED_FORMS) {
      expect(containsForbiddenMarkup(inert), inert).toBe(false);
      expect(() => assertNoXxe(inert, 'probe')).not.toThrow();
    }
  });

  it('passes legit XML: comments, escaped math text, plain parts', () => {
    for (const legit of LEGIT_XML) {
      expect(containsForbiddenMarkup(legit), legit).toBe(false);
      expect(() => assertNoXxe(legit, 'probe')).not.toThrow();
    }
  });

  it('ignores empty and non-string input instead of crashing', () => {
    expect(containsForbiddenMarkup('')).toBe(false);
    expect(containsForbiddenMarkup(undefined)).toBe(false);
    expect(containsForbiddenMarkup(null)).toBe(false);
    expect(() => assertNoXxe('', 'probe')).not.toThrow();
  });
});

describe('sanitizer ingress', () => {
  it('sanitizeDocumentXml rejects case/split variants with E_VALIDATION', () => {
    for (const decl of ['<!doctype x>', '<!DoCtYpE x>', '<!\nENTITY x "y">', '<! ENTITY x "y">']) {
      expect(() => sanitizeDocumentXml(evilDocXml(decl)), decl).toThrowError(/E_VALIDATION/);
    }
  });

  it('sanitizeDocumentXml still accepts legit docs with entity-like math text', () => {
    const legit =
      '<?xml version="1.0"?><w:document xmlns:w="' + W_NS + '">' +
      '<w:body><w:p><w:r><w:t>AT&amp;T profit &lt; forecast</w:t></w:r></w:p></w:body></w:document>';
    expect(() => sanitizeDocumentXml(legit)).not.toThrow();
  });

  it('sanitizeDocxBuffer fails closed on an evil part smuggled into the zip', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const buf = parser.rezip();
    const zip = new AdmZip(buf);
    zip.updateFile('word/document.xml', Buffer.from(evilDocXml('<!doctype evil>'), 'utf8'));
    expect(() => sanitizeDocxBuffer(zip.toBuffer())).toThrowError(/E_VALIDATION/);
  });
});

describe('parser ingress', () => {
  it('constructor rejects evil document.xml in every spelling', () => {
    for (const decl of ['<!DOCTYPE x>', '<!doctype x>', '<!DoCtYpE x>', '<!\nDOCTYPE x>']) {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const zip = new AdmZip(parser.rezip());
      zip.updateFile('word/document.xml', Buffer.from(evilDocXml(decl), 'utf8'));
      expect(() => new DocxParser(zip.toBuffer()), decl).toThrowError(/E_VALIDATION/);
    }
  });

  it('reloadDocumentXml, updateNode, insertRootNode and writeText reject before parsing', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    expect(() => parser.reloadDocumentXml(evilDocXml('<!DoCtYpE x>'))).toThrowError(/E_VALIDATION/);
    expect(() => parser.reloadDocumentXml(evilDocXml('<!\nENTITY x "y">'))).toThrowError(/E_VALIDATION/);
    const ast = await parser.getAST();
    const nodeId = (ast.sections[0].nodes[0] as { id: string }).id;
    expect(() => parser.updateNode(nodeId, '<!ENTITY x "y"><w:p/>')).toThrowError(/E_VALIDATION/);
    expect(() => parser.insertRootNode('<!doctype x><w:p/>')).toThrowError(/E_VALIDATION/);
    expect(() => parser.getZipAdapter().writeText('word/document.xml', '<!ENTITY x "y"><w:p/>')).toThrowError(
      /E_VALIDATION/
    );
  });

  it('getAST hydration stays loud on evil numbering and core parts', async () => {
    const evilNumbering = '<!doctype n><w:numbering xmlns:w="' + W_NS + '"/>';
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    plantInLiveZip(parser, 'word/numbering.xml', evilNumbering);
    await expect(parser.getAST()).rejects.toThrow(/E_VALIDATION/);

    const parser2 = new DocxParser(createBlankDocxBuffer('blank'));
    plantInLiveZip(parser2, 'docProps/core.xml', '<!DoCtYpE c><cp:coreProperties/>');
    await expect(parser2.getAST()).rejects.toThrow(/E_VALIDATION/);
  });

  it('syncNumberingToZip path throws instead of silently dropping the payload', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const numbering = new NumberingManipulator(ast);
    expect(() => numbering.parseNumbering('<!entity x "y"><w:numbering/>')).toThrowError(/E_VALIDATION/);
    expect(() => numbering.parseNumbering('<w:numbering xmlns:w="' + W_NS + '"/>')).not.toThrow();
  });
});

describe('operation ingress points', () => {
  it('xml insert rejects case/split variants with E_VALIDATION and keeps typed errors', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = (ast.sections[0].nodes[0] as { id: string }).id;
    for (const decl of ['<!DOCTYPE x>', '<!entity x "y">', '<!\nDOCTYPE x>', '<! ENTITY x "y">']) {
      await expect(
        executeInsert(
          { action: 'insert', type: 'xml', anchor, position: 'after', xml: decl + '<w:p/>' } as never,
          ast as never
        )
      ).rejects.toMatchObject({ code: 'E_VALIDATION' });
    }
    await expect(
      executeInsert(
        {
          action: 'insert',
          type: 'paragraph',
          anchor,
          position: 'after',
          text: 'AT&T math a < b stays legit',
          id: 'legit-para',
        } as never,
        ast as never
      )
    ).resolves.toBeUndefined();
  });

  it('equation latex rejects declaration smuggling, keeps plain math working', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = (ast.sections[0].nodes[0] as { id: string }).id;
    for (const latex of ['x <!DOCTYPE foo>', 'y <!entity x "p">', 'z <!\nENTITY x "p">']) {
      await expect(
        executeEquation({ action: 'equation', operation: 'insert', anchor, latex } as never, ast as never)
      ).rejects.toThrow(/E_VALIDATION/);
    }
    const legit = await executeEquation(
      { action: 'equation', operation: 'insert', anchor, latex: 'profit & loss, a < b' } as never,
      ast as never,
      parser as never
    );
    expect((legit as { id: string }).id).toBeTruthy();
  });
});

describe('part-manipulator ingress points', () => {
  it('core-properties parse and merge reject every spelling', () => {
    for (const decl of ['<!DOCTYPE c>', '<!doctype c>', '<!ENTITY x "y">', '<!\nDOCTYPE c>']) {
      expect(() => parseCorePropertiesXml(decl + '<cp:coreProperties/>'), decl).toThrowError(/E_VALIDATION/);
      expect(() => mergeCorePropertiesXml(decl + '<cp:coreProperties/>', {}), decl).toThrowError(/E_VALIDATION/);
    }
    expect(() =>
      parseCorePropertiesXml(
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>'
      )
    ).not.toThrow();
  });

  it('comments hydration rejects an evil comments part before building markers', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const paraEl = (ast.sections[0].nodes[0] as { xml: Element }).xml;
    plantInLiveZip(parser, 'word/comments.xml', '<!DoCtYpE c><w:comments xmlns:w="' + W_NS + '"/>');
    expect(() => addComment(parser.getZipAdapter(), paraEl, 'Eve', 'hi')).toThrowError(/E_VALIDATION/);
  });

  it('header/footer setter rejects evil xmlContent before touching the package', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const hf = new HeaderFooterManipulator(ast as never, parser);
    for (const decl of ['<!DOCTYPE h>', '<!doctype h>', '<!\nENTITY x "y">']) {
      expect(
        () => hf.setHeaderFooter({ targetType: 'header', xmlContent: decl + '<w:hdr/>' }),
        decl
      ).toThrowError(/E_VALIDATION/);
    }
    const ok = hf.setHeaderFooter({ targetType: 'header', text: 'Legit header & prose' });
    expect(ok.rId).toBeTruthy();
  });

  it('styles hydration rejects evil payloads and keeps clean palettes', () => {
    expect(() =>
      extractStylePalette('<!doctype s><w:styles xmlns:w="' + W_NS + '"/>')
    ).toThrowError(/E_VALIDATION/);
    expect(() =>
      extractStylePalette('<!\nENTITY x "y"><w:styles xmlns:w="' + W_NS + '"/>')
    ).toThrowError(/E_VALIDATION/);
    const palette = extractStylePalette('<w:styles xmlns:w="' + W_NS + '"/>');
    expect(palette.primaryFont).toBe('Calibri');
  });

  it('chart update rejects a poisoned chart part instead of parsing it', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart(
      {
        action: 'chart',
        chartType: 'column',
        title: 'Seed',
        categories: ['A'],
        series: [{ name: 'S', values: [1] }],
        id: 'chart_guard_node',
      } as never,
      ast as never,
      parser
    );
    plantInLiveZip(parser, 'word/charts/chart1.xml', '<!DoCtYpE c><c:chartSpace/>');
    expect(() =>
      executeChartUpdate(
        { action: 'chart', operation: 'update', target: 'chart_guard_node', title: 'X' } as never,
        ast as never,
        parser
      )
    ).toThrowError(/E_VALIDATION/);
  });
});

describe('sanitizeXml central routing (validator divergence kill)', () => {
  const SPLIT_VARIANTS = [
    '<!DOCTYPE x>',
    '<!doctype x>',
    '<!DoCtYpE x>',
    '<!\nDOCTYPE x>',
    '<!\r\nDOCTYPE x>',
    '<! ENTITY x "y">',
    '<!  DOCTYPE x>',
    '<!\tENTITY x "y">',
    '<!\x00DOCTYPE x>',
  ];
  it('sanitizeXml throws E_VALIDATION on every split/space/NUL/lowercase variant', () => {
    for (const decl of SPLIT_VARIANTS) {
      expect(() => sanitizeXml(decl + '<w:p/>'), decl).toThrowError(/E_VALIDATION/);
      expect(() => sanitizeXml('<?xml version="1.0"?>' + decl + '<w:p/>'), decl).toThrowError(/E_VALIDATION/);
    }
  });
  it('sanitizeXml keeps legit XML and script gate intact', () => {
    expect(sanitizeXml('<w:p>hello</w:p>')).toBe('<w:p>hello</w:p>');
    expect(() => sanitizeXml('<w:p><script>alert(1)</script></w:p>')).toThrowError(/script/i);
    expect(() => sanitizeXml('<!-- a plain comment --><w:p>hi</w:p>')).not.toThrow();
  });
  it('toc container path rejects the same variants before parsing', () => {
    for (const decl of ['<!doctype x>', '<!DoCtYpE x>', '<!\nENTITY x "y">', '<! ENTITY x "y">', '<!\x00DOCTYPE x>']) {
      expect(() => validateTocContainerXml(decl + '<w:p/>'), decl).toThrowError(/E_VALIDATION/);
    }
    const legit =
      '<w:fldSimple w:instr="TOC \\o &quot;1-3&quot; \\h \\z \\u">' +
      '<w:r><w:t>cached</w:t></w:r></w:fldSimple>';
    expect(() => validateTocContainerXml(legit)).not.toThrow();
  });
});

describe('rels/CT/footnotes ingress gates (evil-rels fixtures)', () => {
  const REL_DECLS = ['<!DOCTYPE r>', '<!doctype r>', '<!\nDOCTYPE r>', '<! ENTITY x "y">', '<!\x00DOCTYPE r>'];
  const evilRels = (decl: string): string =>
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    decl +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';
  const evilCt = (decl: string): string =>
    '<?xml version="1.0" encoding="UTF-8"?>' +
    decl +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="xml" ContentType="application/xml"/></Types>';
  const evilFootnotes = (decl: string): string =>
    '<?xml version="1.0" encoding="UTF-8"?>' +
    decl +
    '<w:footnotes xmlns:w="' + W_NS + '">' +
    '<w:footnote w:id="1"><w:p><w:r><w:t>evil</w:t></w:r></w:p></w:footnote></w:footnotes>';

  it('parser summary + AST reject evil document rels in every spelling', async () => {
    for (const decl of REL_DECLS) {
      const p1 = new DocxParser(createBlankDocxBuffer('blank'));
      plantInLiveZip(p1, 'word/_rels/document.xml.rels', evilRels(decl));
      expect(() => p1.getStructuralSummary(), decl).toThrowError(/E_VALIDATION/);
      const p2 = new DocxParser(createBlankDocxBuffer('blank'));
      plantInLiveZip(p2, 'word/_rels/document.xml.rels', evilRels(decl));
      await expect(p2.getAST(), decl).rejects.toThrow(/E_VALIDATION/);
    }
  });

  it('footnote op rejects evil footnotes part before parsing', async () => {
    for (const decl of ['<!DOCTYPE f>', '<!doctype f>', '<!\nENTITY x "y">', '<! ENTITY x "y">']) {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const anchor = (ast.sections[0].nodes[0] as { id: string }).id;
      plantInLiveZip(parser, 'word/footnotes.xml', evilFootnotes(decl));
      expect(
        () => executeFootnote({ action: 'footnote', anchor, text: 'note' } as never, ast as never, parser as never),
        decl
      ).toThrowError(/E_VALIDATION/);
    }
  });

  it('footnote op rejects evil CT + evil rels when registering a fresh part', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = (ast.sections[0].nodes[0] as { id: string }).id;
    plantInLiveZip(parser, '[Content_Types].xml', evilCt('<!doctype c>'));
    expect(
      () => executeFootnote({ action: 'footnote', anchor, text: 'note' } as never, ast as never, parser as never)
    ).toThrowError(/E_VALIDATION/);
    const parser2 = new DocxParser(createBlankDocxBuffer('blank'));
    const ast2 = await parser2.getAST();
    const anchor2 = (ast2.sections[0].nodes[0] as { id: string }).id;
    plantInLiveZip(parser2, 'word/_rels/document.xml.rels', evilRels('<!DoCtYpE r>'));
    expect(
      () => executeFootnote({ action: 'footnote', anchor: anchor2, text: 'note' } as never, ast2 as never, parser2 as never)
    ).toThrowError(/E_VALIDATION/);
  });

  it('hyperlink op rejects evil document rels before parsing', async () => {
    for (const decl of ['<!DOCTYPE r>', '<!doctype r>', '<!\nDOCTYPE r>']) {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const target = (ast.sections[0].nodes[0] as { id: string }).id;
      plantInLiveZip(parser, 'word/_rels/document.xml.rels', evilRels(decl));
      expect(
        () =>
          executeHyperlink(
            { action: 'hyperlink', target, url: 'https://example.com', text: 'x' } as never,
            ast as never,
            parser as never
          ),
        decl
      ).toThrowError(/E_VALIDATION/);
    }
  });

  it('image op rejects evil rels + evil content-types before parsing', async () => {
    const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = (ast.sections[0].nodes[0] as { id: string }).id;
    plantInLiveZip(parser, 'word/_rels/document.xml.rels', evilRels('<!doctype r>'));
    expect(() =>
      executeImage(
        { action: 'image', operation: 'insert', anchor, fileData: `data:image/png;base64,${png1x1}` } as never,
        ast as never,
        parser as never
      )
    ).toThrowError(/E_VALIDATION/);
    const parser2 = new DocxParser(createBlankDocxBuffer('blank'));
    const ast2 = await parser2.getAST();
    const anchor2 = (ast2.sections[0].nodes[0] as { id: string }).id;
    plantInLiveZip(parser2, '[Content_Types].xml', evilCt('<!ENTITY x "y">'));
    expect(() =>
      executeImage(
        { action: 'image', operation: 'insert', anchor: anchor2, fileData: `data:image/png;base64,${png1x1}` } as never,
        ast2 as never,
        parser2 as never
      )
    ).toThrowError(/E_VALIDATION/);
  });

  it('chart rels readers reject evil doc/header/chart rels parts', async () => {
    const seed = async (): Promise<{ parser: DocxParser; ast: any; nodeId: string; rId: string }> => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = (await parser.getAST()) as any;
      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Seed',
          categories: ['A'],
          series: [{ name: 'S', values: [1] }],
          id: 'chart_rels_seed',
        } as never,
        ast as never,
        parser as never
      );
      const node = ast.sections[0].nodes.find((n: any) => n.id === 'chart_rels_seed') as any;
      const chartEl = node.xml.getElementsByTagName('c:chart')[0] as Element;
      const rIdAttr =
        chartEl.getAttribute('r:id') ||
        chartEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      if (typeof rIdAttr !== 'string' || rIdAttr.length === 0) {
        throw new Error('seed chart is missing its relationship id');
      }
      return { parser: parser as DocxParser, ast, nodeId: 'chart_rels_seed', rId: rIdAttr };
    };
    {
      const { parser, ast, nodeId, rId } = await seed();
      delete (ast.relationships as Record<string, unknown>)[rId];
      plantInLiveZip(parser, 'word/_rels/document.xml.rels', evilRels('<!doctype r>'));
      expect(() =>
        executeChartUpdate(
          { action: 'chart', operation: 'update', target: nodeId, title: 'X' } as never,
          ast as never,
          parser as never
        )
      ).toThrowError(/E_VALIDATION/);
    }
    {
      const { parser, ast, nodeId, rId } = await seed();
      delete (ast.relationships as Record<string, unknown>)[rId];
      const cleanRels =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>';
      plantInLiveZip(parser, 'word/_rels/document.xml.rels', cleanRels);
      plantInLiveZip(parser, 'word/_rels/header1.xml.rels', evilRels('<!ENTITY x "y">'));
      expect(() =>
        executeChartUpdate(
          { action: 'chart', operation: 'update', target: nodeId, title: 'X' } as never,
          ast as never,
          parser as never
        )
      ).toThrowError(/E_VALIDATION/);
    }
    {
      const { parser, ast, nodeId } = await seed();
      plantInLiveZip(parser, 'word/charts/_rels/chart1.xml.rels', evilRels('<!DoCtYpE r>'));
      expect(() =>
        executeChartUpdate(
          { action: 'chart', operation: 'update', target: nodeId, title: 'X', categories: ['A'], series: [{ name: 'S', values: [2] }] } as never,
          ast as never,
          parser as never
        )
      ).toThrowError(/E_VALIDATION/);
    }
    {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      plantInLiveZip(parser, 'word/_rels/document.xml.rels', evilRels('<!doctype r>'));
      const m2d = new Chart2DManipulator(ast as never, parser as never);
      expect(() =>
        m2d.insert2DChart({ title: 'T', categories: ['A'], series: [{ name: 'S', values: [1] }] })
      ).toThrowError(/E_VALIDATION/);
    }
  });

  it('header/footer + chart manipulators reject evil rels/CT on ensure paths', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    plantInLiveZip(parser, 'word/_rels/document.xml.rels', evilRels('<!DoCtYpE r>'));
    const hf = new HeaderFooterManipulator(ast as never, parser as never);
    expect(() => hf.setHeaderFooter({ targetType: 'header', text: 'hi' })).toThrowError(/E_VALIDATION/);
    const parser2 = new DocxParser(createBlankDocxBuffer('blank'));
    const ast2 = await parser2.getAST();
    plantInLiveZip(parser2, '[Content_Types].xml', evilCt('<!doctype c>'));
    const m3d = new Chart3DManipulator(ast2 as never, parser2 as never);
    expect(() =>
      m3d.insert3DChart({ title: 'T', categories: ['A'], series: [{ name: 'S', values: [1] }] })
    ).toThrowError(/E_VALIDATION/);
  });

  it('markdown footnotes reader + schema-order parse throw before parsing', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    plantInLiveZip(parser, 'word/footnotes.xml', evilFootnotes('<!DoCtYpE f>'));
    expect(() => convertAstToMarkdown(ast as never, parser as never)).toThrowError(/E_VALIDATION/);
    for (const decl of ['<!DOCTYPE x>', '<!doctype x>', '<!\nENTITY x "y">']) {
      expect(() => canonicalizePartXml(decl + '<w:p/>'), decl).toThrowError(/E_VALIDATION/);
    }
    expect(canonicalizePartXml('<w:p xmlns:w="' + W_NS + '"><w:r><w:t>ok</w:t></w:r></w:p>')).toContain('<w:p');
  });

  it('sanitizer ingress rejects evil rels smuggled into the package', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const buf = parser.rezip();
    const zip = new AdmZip(buf);
    zip.updateFile('word/_rels/document.xml.rels', Buffer.from(evilRels('<!doctype r>'), 'utf8'));
    expect(() => sanitizeDocxBuffer(zip.toBuffer())).toThrowError(/E_VALIDATION/);
    const zip2 = new AdmZip(buf);
    zip2.updateFile('[Content_Types].xml', Buffer.from(evilCt('<!ENTITY x "y">'), 'utf8'));
    expect(() => sanitizeDocxBuffer(zip2.toBuffer())).toThrowError(/E_VALIDATION/);
  });
});

describe('comment-split inertness (no gate change, literal passthrough)', () => {
  const SPLITS = ['<!<!--c-->DOCTYPE x>', '<!DO<!--c-->CTYPE x>', '<! --DOCTYPE x>'];
  it('guard leaves comment-split shapes alone', () => {
    for (const s of SPLITS) {
      expect(containsForbiddenMarkup(s), s).toBe(false);
      expect(() => assertNoXxe(s, 'probe')).not.toThrow();
    }
  });
  it('xmldom renders comment-split shapes inert (escaped text or dropped comment, never a declaration)', () => {
    for (const s of SPLITS) {
      const xml = '<?xml version="1.0"?><root>' + s + '<a>x</a></root>';
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      expect(doc.getElementsByTagName('parsererror').length).toBe(0);
      const out = new XMLSerializer().serializeToString(doc);
      expect(out).not.toMatch(/<!DOCTYPE/i);
      expect(out).not.toMatch(/<!ENTITY/i);
      expect(out.includes('&lt;!') || out.includes('<!--')).toBe(true);
    }
  });
  it('legit comment-heavy documents keep passing (no false positives)', () => {
    const heavy =
      '<?xml version="1.0"?><w:document xmlns:w="' + W_NS + '">' +
      '<!-- ASTra TOC: baked page numbers are ESTIMATES; Word updates fields on open (updateFields=true). -->' +
      '<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p><!-- c1 --><!-- c2 --></w:body></w:document>';
    expect(containsForbiddenMarkup(heavy)).toBe(false);
    expect(() => assertNoXxe(heavy, 'probe')).not.toThrow();
    expect(() => sanitizeXml(heavy)).not.toThrow();
    expect(() => canonicalizePartXml(heavy)).not.toThrow();
  });
});

describe('bare declaration escape-to-text (no gate change)', () => {
  const BARES = ['<!ELEMENT a ANY>', '<!ATTLIST a b CDATA "x">', '<!NOTATION n SYSTEM "x">'];
  it('guard leaves bare declarations alone', () => {
    for (const s of BARES) {
      expect(containsForbiddenMarkup(s), s).toBe(false);
      expect(() => assertNoXxe(s, 'probe')).not.toThrow();
    }
  });
  it('xmldom escapes bare declarations to text with no declaration node', () => {
    for (const s of BARES) {
      const xml = '<?xml version="1.0"?><root>' + s + '<a>x</a></root>';
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      expect(doc.getElementsByTagName('parsererror').length).toBe(0);
      expect(doc.documentElement.tagName).toBe('root');
      const out = new XMLSerializer().serializeToString(doc);
      expect(out).toContain('&lt;!' + s.slice(2, 6));
      expect(out).toContain('<a>x</a>');
    }
  });
});

describe('multi-bang declaration runs (adjacent bangs form a real doctype)', () => {
  const ADJACENT_BANGS = ['<!!DOCTYPE x>', '<!!!DOCTYPE x>', '<!!!!DOCTYPE x>'];
  it('guard flags double/triple/quad-bang runs plus whitespace mixes', () => {
    for (const decl of [...ADJACENT_BANGS, '<!!ENTITY x "y">', '<!!!ENTITY x "y">', '<!!\nDOCTYPE x>', '<!! DOCTYPE x>']) {
      expect(containsForbiddenMarkup(decl), decl).toBe(true);
      expect(() => assertNoXxe(decl, 'probe'), decl).toThrowError(/E_VALIDATION/);
    }
  });
  it('xmldom accepts adjacent multi-bang opens as doctype nodes (why the !+ run is load-bearing)', () => {
    for (const decl of ADJACENT_BANGS) {
      const doc = new DOMParser().parseFromString('<?xml version="1.0"?><root>' + decl + '<a>x</a></root>', 'text/xml');
      expect(doc.getElementsByTagName('parsererror').length).toBe(0);
      expect(doc.doctype, decl).not.toBeNull();
    }
  });
  it('whitespace-mixed multi-bang escapes to inert text yet stays blocked fail-closed', () => {
    for (const decl of ['<!! DOCTYPE x>', '<!!\nDOCTYPE x>']) {
      const doc = new DOMParser().parseFromString('<?xml version="1.0"?><root>' + decl + '<a>x</a></root>', 'text/xml');
      expect(doc.doctype, decl).toBeNull();
      expect(containsForbiddenMarkup(decl), decl).toBe(true);
      expect(() => assertNoXxe(decl, 'probe'), decl).toThrowError(/E_VALIDATION/);
    }
  });
  it('space-split bangs stay unflagged and inert (no keyword adjacency, same line as comment-splits)', () => {
    for (const decl of ['<! !DOCTYPE x>', '<! \n !DOCTYPE x>']) {
      const doc = new DOMParser().parseFromString('<?xml version="1.0"?><root>' + decl + '<a>x</a></root>', 'text/xml');
      expect(doc.doctype, decl).toBeNull();
      expect(containsForbiddenMarkup(decl), decl).toBe(false);
      expect(() => assertNoXxe(decl, 'probe')).not.toThrow();
    }
  });
  it('double-bang comments and cdata keep passing (no false positives)', () => {
    for (const legit of ['<!-- !! just bangs -->', '<!-- a !! b -->', '<![CDATA[x]]>']) {
      expect(containsForbiddenMarkup(legit), legit).toBe(false);
      expect(() => assertNoXxe(legit, 'probe')).not.toThrow();
    }
  });
});

describe('summary and rezip auxiliary-part gates', () => {
  const evilTheme =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<!DOCTYPE foo>' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:evil>x</a:evil></a:theme>';
  const evilHeader = (decl: string): string =>
    '<?xml version="1.0" encoding="UTF-8"?>' + decl + '<w:hdr xmlns:w="' + W_NS + '"/>';
  const evilFooter = (decl: string): string =>
    '<?xml version="1.0" encoding="UTF-8"?>' + decl + '<w:ftr xmlns:w="' + W_NS + '"/>';
  const evilChart = (decl: string): string =>
    '<?xml version="1.0" encoding="UTF-8"?>' + decl + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>';

  it('summary rejects evil header/footer/chart parts like rels', () => {
    for (const decl of ['<!DOCTYPE h>', '<!doctype h>', '<!!DOCTYPE h>', '<!\nENTITY x "y">']) {
      const p1 = new DocxParser(createBlankDocxBuffer('blank'));
      plantInLiveZip(p1, 'word/header1.xml', evilHeader(decl));
      expect(() => p1.getStructuralSummary(), 'header:' + decl).toThrowError(/E_VALIDATION/);
      const p2 = new DocxParser(createBlankDocxBuffer('blank'));
      plantInLiveZip(p2, 'word/footer1.xml', evilFooter(decl));
      expect(() => p2.getStructuralSummary(), 'footer:' + decl).toThrowError(/E_VALIDATION/);
      const p3 = new DocxParser(createBlankDocxBuffer('blank'));
      plantInLiveZip(p3, 'word/charts/chart1.xml', evilChart(decl));
      expect(() => p3.getStructuralSummary(), 'chart:' + decl).toThrowError(/E_VALIDATION/);
    }
  });

  it('summary rejects an evil theme part instead of plumbing it through', () => {
    for (const decl of ['<!DOCTYPE t>', '<!doctype t>', '<!!DOCTYPE t>']) {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      plantInLiveZip(parser, 'word/theme/theme1.xml', evilTheme.replace('<!DOCTYPE foo>', decl));
      expect(() => parser.getStructuralSummary(), decl).toThrowError(/E_VALIDATION/);
    }
  });

  it('evil theme does not survive rezip (export throws fail-closed)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    plantInLiveZip(parser, 'word/theme/theme1.xml', evilTheme);
    await parser.getAST();
    expect(() => parser.rezip()).toThrowError(/E_VALIDATION/);
  });

  it('writeBuffer rejects evil payloads exactly like writeText', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    for (const decl of ['<!DOCTYPE t>', '<!doctype t>', '<!!DOCTYPE t>', '<!ENTITY x "y">']) {
      expect(
        () => parser.getZipAdapter().writeBuffer('word/theme/theme1.xml', Buffer.from(decl + '<a:theme/>', 'utf8')),
        decl
      ).toThrowError(/E_VALIDATION/);
    }
    expect(() =>
      parser.getZipAdapter().writeBuffer('word/media/blob.bin', Buffer.from([0x89, 0x50, 0x4e, 0x47]))
    ).not.toThrow();
    expect(parser.getZipAdapter().hasFile('word/media/blob.bin')).toBe(true);
  });
});

describe('utf16 storage boundary (no detector in this slice)', () => {
  it('guard scans utf8-decoded text only: utf16 bytes evade the string pattern (pinned limitation)', () => {
    const evil = '<?xml version="1.0"?><!DOCTYPE foo><root><a>x</a></root>';
    const mojibake = Buffer.from(evil, 'utf16le').toString('utf8');
    expect(containsForbiddenMarkup(mojibake)).toBe(false);
  });
  it('a utf16-encoded document.xml still chokes loudly at parse instead of loading silently', () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.updateFile(
      'word/document.xml',
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-16"?><!DOCTYPE foo>' +
          '<w:document xmlns:w="' + W_NS + '"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>',
        'utf16le'
      )
    );
    expect(() => new DocxParser(zip.toBuffer())).toThrow();
  });
});
