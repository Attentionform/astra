import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  Chart3DManipulator,
  DocxParser,
  DocumentAST,
  validateDocxBuffer,
} from '@attentionform/astra-engine';
import { DOMParser } from '@xmldom/xmldom';

// P3 (3D multi-series spacing mitigation, opt-in 3D path only):
// bar3D/column3D charts must carry c:gapWidth + c:gapDepth so renderers
// don't collapse cluster spacing and jam category labels.
// NOTE: c:overlap is NOT a legal child of c:bar3DChart per ECMA-376
// CT_Bar3DChart (barDir, grouping, varyColors?, ser*, dLbls?, gapWidth?,
// gapDepth?, shape?, axId*) — @attentionform/ooxml-validator rejects it with
// Sch_InvalidElementContentExpectingComplex, so it is intentionally absent.

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

const THREE_SERIES = [
  { name: 'ARR', values: [10, 25, 50, 100], color: '10B981' },
  { name: 'Headcount', values: [15, 30, 60, 120], color: '6366F1' },
  { name: 'Churn', values: [5, 8, 12, 9], color: 'F59E0B' },
];
const CATS = ['Q1', 'Q2', 'Q3', 'Q4'];

describe('P3 3D multi-series spacing (bar3D/column3D gap hints)', () => {
  it.each(['bar3D', 'column3D'] as const)('%s multi-series XML carries gapWidth + gapDepth, no overlap', (chartType) => {
    const m = new Chart3DManipulator(makeMinimalAst());
    const xml = m.generate3DChartXml({ chartType, title: 'P3', categories: CATS, series: THREE_SERIES });
    expect(xml).toContain('<c:bar3DChart>');
    expect(xml).toContain('<c:gapWidth val="150"/>');
    expect(xml).toContain('<c:gapDepth val="150"/>');
    expect(xml).not.toContain('<c:overlap');
  });

  it('3D multi-series bar doc validates ok:true and chart part carries spacing', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const m = new Chart3DManipulator(ast, parser);
    m.insert3DChart({ chartType: 'bar3D', title: 'P3', categories: CATS, series: THREE_SERIES });
    const buf = parser.rezip();

    const v = await validateDocxBuffer(buf);
    expect(v.ok).toBe(true);
    expect(v.errors).toEqual([]);

    const zip = new AdmZip(buf);
    const chartXml = zip.getEntry('word/charts/chart1.xml')?.getData().toString('utf8') ?? '';
    expect(chartXml).toContain('<c:gapWidth val="150"/>');
    expect(chartXml).toContain('<c:gapDepth val="150"/>');
  });

  it('single-series 3D output still validates (behavior unchanged apart from spacing)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const m = new Chart3DManipulator(ast, parser);
    m.insert3DChart({
      chartType: 'column3D',
      title: 'P3 single',
      categories: CATS,
      series: [{ name: 'Solo', values: [1, 2, 3, 4], color: '0066CC' }],
    });
    const buf = parser.rezip();
    const v = await validateDocxBuffer(buf);
    expect(v.ok).toBe(true);
    expect(v.errors).toEqual([]);
  });

  it('non-bar 3D types untouched (no gap hints in line3D/pie3D)', () => {
    const m = new Chart3DManipulator(makeMinimalAst());
    const line = m.generate3DChartXml({ chartType: 'line3D', title: 'L', categories: CATS, series: THREE_SERIES.slice(0, 1) });
    expect(line).not.toContain('<c:gapWidth');
    expect(line).not.toContain('<c:gapDepth');
    const pie = m.generate3DChartXml({ chartType: 'pie3D', title: 'P', categories: CATS, series: THREE_SERIES.slice(0, 1) });
    expect(pie).not.toContain('<c:gapWidth');
    expect(pie).not.toContain('<c:gapDepth');
  });
});
