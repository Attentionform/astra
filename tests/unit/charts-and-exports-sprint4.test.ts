import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import { convertAstToHtml, convertAstToMarkdown } from '@attentionform/astra-engine';
import {
  DocxParser,
  executeChart,
  executeChartUpdate,
  executeInsert,
} from '@attentionform/astra-engine';

describe('Sprint 4: Vector SVG Chart Export, PDF Headers/Footers & Cell Colors', () => {
  it('html.ts: renders vector SVG chart when chart node is present in AST', () => {
    const ast: any = {
      sections: [
        {
          id: 'sec_1',
          nodes: [
            {
              id: 'chart_1',
              type: 'chart',
              xml: null,
              metadata: {
                title: 'Revenue Projection 2026',
                categories: ['Q1', 'Q2', 'Q3', 'Q4'],
                series: [
                  { name: 'Revenue', values: [100, 150, 200, 280], color: '2563EB' }
                ]
              },
              version: 1,
              checksum: 'mock-1'
            }
          ],
          pageSetup: {
            orientation: 'portrait',
            paperSize: 'Letter'
          },
          version: 1
        }
      ],
      metadata: { hasTrackedChanges: false },
      settings: { trackChanges: false, defaultAuthor: 'ASTra' },
      nextChangeId: 1
    };

    const html = convertAstToHtml(ast, 'Chart Test', { printMode: true });
    expect(html).toContain('Revenue Projection 2026');
    expect(html).toContain('<svg');
    expect(html).toContain('Q1');
    expect(html).toContain('Q4');
    expect(html).toContain('@page');
  });

  it('html.ts: extracts custom cell background colors and high contrast text', () => {
    const xmlStr = `
      <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:tr>
          <w:tc>
            <w:tcPr><w:shd w:fill="1E3A5F"/></w:tcPr>
            <w:p><w:r><w:t>Dark Header</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:tcPr><w:shd w:fill="EFF6FF"/></w:tcPr>
            <w:p><w:r><w:t>Light Cell</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `;
    const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
    const tblEl = doc.getElementsByTagName('w:tbl')[0];

    const ast: any = {
      sections: [
        {
          id: 'sec_1',
          nodes: [
            {
              id: 't_1',
              type: 'table',
              xml: tblEl,
              metadata: {},
              version: 1,
              checksum: 'mock-tbl'
            }
          ],
          pageSetup: {
            orientation: 'portrait',
            paperSize: 'Letter'
          },
          version: 1
        }
      ],
      metadata: { hasTrackedChanges: false },
      settings: { trackChanges: false, defaultAuthor: 'ASTra' },
      nextChangeId: 1
    };

    const html = convertAstToHtml(ast);
    expect(html).toContain('background-color: #1E3A5F');
    expect(html).toContain('color: #FFFFFF');
    expect(html).toContain('background-color: #EFF6FF');
  });

  it('markdown.ts: renders structured data table for chart nodes', () => {
    const ast: any = {
      sections: [
        {
          id: 'sec_1',
          nodes: [
            {
              id: 'chart_1',
              type: 'chart',
              xml: null,
              metadata: {
                title: 'Market Share',
                categories: ['North America', 'Europe', 'Asia'],
                series: [
                  { name: 'Share %', values: [45, 30, 25] }
                ]
              },
              version: 1,
              checksum: 'mock-1'
            }
          ],
          pageSetup: {
            orientation: 'portrait',
            paperSize: 'Letter'
          },
          version: 1
        }
      ],
      metadata: { hasTrackedChanges: false },
      settings: { trackChanges: false, defaultAuthor: 'ASTra' },
      nextChangeId: 1
    };

    const md = convertAstToMarkdown(ast);
    expect(md).toContain('### 📊 Market Share');
    expect(md).toContain('| Category | Share % |');
    expect(md).toContain('| North America | 45 |');
    expect(md).toContain('| Europe | 30 |');
    expect(md).toContain('| Asia | 25 |');
  });
});

describe('R-48: chart insert joins the shared same-anchor after-chain', () => {
  function astIdSequence(ast: any): string[] {
    return ast.sections.flatMap((s: any) => s.nodes.map((n: any) => n.id));
  }

  function domSiblingIds(fromXml: Element): string[] {
    const ids: string[] = [];
    let curr: Element | null = fromXml;
    while (curr) {
      if (curr.nodeType === 1 && (curr as Element).tagName !== 'w:sectPr') {
        const id = (curr as Element).getAttribute?.('data-astra-id');
        if (id) ids.push(id);
      }
      curr = curr.nextSibling as Element | null;
    }
    return ids;
  }

  function expectConsecutive(ids: string[], window: string[]): void {
    const start = ids.indexOf(window[0]);
    expect(start).toBeGreaterThanOrEqual(0);
    expect(ids.slice(start, start + window.length)).toEqual(window);
  }

  function chartOp(anchor: string, id: string, title: string): any {
    return {
      action: 'chart',
      chartType: 'column',
      title,
      categories: ['Q1', 'Q2'],
      series: [{ name: 'S', values: [1, 2], color: '0066CC' }],
      anchor,
      id,
    };
  }

  it('cross-type FIFO: table insert then chart insert land in call order (AST+DOM)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const anchor: string = ast.sections[0].nodes[0].id;
    const anchorXml: Element = ast.sections[0].nodes[0].xml;

    await executeInsert(
      { action: 'insert', type: 'table', anchor, position: 'after', id: 'tbl-1' },
      ast,
      parser
    );
    executeChart(chartOp(anchor, 'chart-1', 'After table'), ast, parser);

    expectConsecutive(astIdSequence(ast), [anchor, 'tbl-1', 'chart-1']);
    expectConsecutive(domSiblingIds(anchorXml), [anchor, 'tbl-1', 'chart-1']);
  });

  it('cross-type FIFO: chart insert then table insert land in call order (AST+DOM)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const anchor: string = ast.sections[0].nodes[0].id;
    const anchorXml: Element = ast.sections[0].nodes[0].xml;

    executeChart(chartOp(anchor, 'chart-1', 'Before table'), ast, parser);
    await executeInsert(
      { action: 'insert', type: 'table', anchor, position: 'after', id: 'tbl-1' },
      ast,
      parser
    );

    expectConsecutive(astIdSequence(ast), [anchor, 'chart-1', 'tbl-1']);
    expectConsecutive(domSiblingIds(anchorXml), [anchor, 'chart-1', 'tbl-1']);
  });

  it('three same-anchor chart inserts land FIFO (AST+DOM)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const anchor: string = ast.sections[0].nodes[0].id;
    const anchorXml: Element = ast.sections[0].nodes[0].xml;

    executeChart(chartOp(anchor, 'c-1', 'First'), ast, parser);
    executeChart(chartOp(anchor, 'c-2', 'Second'), ast, parser);
    executeChart(chartOp(anchor, 'c-3', 'Third'), ast, parser);

    expectConsecutive(astIdSequence(ast), [anchor, 'c-1', 'c-2', 'c-3']);
    expectConsecutive(domSiblingIds(anchorXml), [anchor, 'c-1', 'c-2', 'c-3']);
  });

  it('single anchored chart insert still lands directly after the anchor', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const before = astIdSequence(ast).length;
    const anchor: string = ast.sections[0].nodes[0].id;
    const anchorXml: Element = ast.sections[0].nodes[0].xml;

    executeChart(chartOp(anchor, 'solo-chart', 'Solo'), ast, parser);

    const seq = astIdSequence(ast);
    expect(seq.length).toBe(before + 1);
    expect(seq[seq.indexOf(anchor) + 1]).toBe('solo-chart');
    expectConsecutive(domSiblingIds(anchorXml), [anchor, 'solo-chart']);
  });

  it('chart update path untouched: in-place, node count stable, no second chart part', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();

    executeChart(chartOp(ast.sections[0].nodes[0].id, 'chart-upd', 'Original'), ast, parser);
    const nodesAfterInsert = astIdSequence(ast).length;
    const nodeBefore = ast.sections.flatMap((s: any) => s.nodes).find((n: any) => n.id === 'chart-upd');
    expect(nodeBefore.version).toBe(1);

    executeChartUpdate(
      {
        action: 'chart',
        operation: 'update',
        target: 'chart-upd',
        title: 'Updated',
        categories: ['Q1', 'Q2'],
        series: [{ name: 'S', values: [9, 8], color: '0066CC' }],
      },
      ast,
      parser
    );

    expect(astIdSequence(ast).length).toBe(nodesAfterInsert);
    const nodeAfter = ast.sections.flatMap((s: any) => s.nodes).find((n: any) => n.id === 'chart-upd');
    expect(nodeAfter.version).toBe(2);
    expect(nodeAfter.metadata.title).toBe('Updated');
    const zipAdapter = parser.getZipAdapter();
    expect(zipAdapter.hasFile('word/charts/chart1.xml')).toBe(true);
    expect(zipAdapter.hasFile('word/charts/chart2.xml')).toBe(false);
  });
});

describe('R-14: chart update stale body advisory', () => {
  const OLD_TITLE = 'H1 Revenue January\u2013June';
  const OLD_CATS = ['January', 'February', 'March', 'April', 'May', 'June'];
  const OLD_VALS = [1200, 2400, 3100, 4200, 5300, 6000];
  const CAPTION = 'Figure 1: H1 Revenue January\u2013June sales trend (see chart).';
  const CTA = 'Act now \u2014 6,000 early-bird bonus ends Friday.';
  const UNRELATED = 'Staffing plan for Q3 covers onboarding and mentoring.';

  async function setupDocWithBody() {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const anchor: string = ast.sections[0].nodes[0].id;
    executeChart(
      {
        action: 'chart',
        chartType: 'column',
        title: OLD_TITLE,
        categories: OLD_CATS,
        series: [{ name: 'Revenue', values: OLD_VALS, color: '0066CC' }],
        anchor,
        id: 'chart-r14',
      },
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: CAPTION, id: 'cap-r14' },
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: CTA, id: 'cta-r14' },
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: UNRELATED, id: 'unrelated-r14' },
      ast,
      parser
    );
    return { parser, ast };
  }

  function nodeText(ast: any, id: string): string {
    return ast.sections.flatMap((s: any) => s.nodes).find((n: any) => n.id === id)?.metadata?.text ?? '';
  }

  it('Jul update returns stale refs to January-June caption + 6,000 CTA, skips unrelated, stays in-place', async () => {
    const { parser, ast } = await setupDocWithBody();
    const countBefore = ast.sections.flatMap((s: any) => s.nodes).length;
    const capBefore = nodeText(ast, 'cap-r14');
    const ctaBefore = nodeText(ast, 'cta-r14');
    const unrelatedBefore = nodeText(ast, 'unrelated-r14');
    expect(capBefore).toContain('January');
    expect(ctaBefore).toContain('6,000');

    const res: any = executeChartUpdate(
      {
        action: 'chart',
        operation: 'update',
        target: 'chart-r14',
        title: 'July Revenue Update',
        categories: ['Jul'],
        series: [{ name: 'Revenue', values: [7500], color: '0066CC' }],
      },
      ast,
      parser
    );

    expect(res).toBeDefined();
    expect(Array.isArray(res.possiblyStaleBodyRefs)).toBe(true);
    const ids = res.possiblyStaleBodyRefs.map((r: any) => r.nodeId);
    expect(ids).toContain('cap-r14');
    expect(ids).toContain('cta-r14');
    expect(ids).not.toContain('unrelated-r14');
    expect(ids).not.toContain('chart-r14');
    const byId: Record<string, string> = {};
    for (const r of res.possiblyStaleBodyRefs) byId[r.nodeId] = r.preview;
    expect(byId['cap-r14']).toContain('January');
    expect(byId['cta-r14']).toContain('6,000');
    for (const r of res.possiblyStaleBodyRefs) {
      expect(typeof r.nodeId).toBe('string');
      expect(typeof r.preview).toBe('string');
      expect(r.preview.length).toBeGreaterThan(0);
    }
    expect(nodeText(ast, 'cap-r14')).toBe(capBefore);
    expect(nodeText(ast, 'cta-r14')).toBe(ctaBefore);
    expect(nodeText(ast, 'unrelated-r14')).toBe(unrelatedBefore);
    expect(ast.sections.flatMap((s: any) => s.nodes).length).toBe(countBefore);
    const chartNode = ast.sections.flatMap((s: any) => s.nodes).find((n: any) => n.id === 'chart-r14');
    expect(chartNode.version).toBe(2);
    expect(chartNode.metadata.title).toBe('July Revenue Update');
    const zipAdapter = parser.getZipAdapter();
    expect(zipAdapter.hasFile('word/charts/chart1.xml')).toBe(true);
    expect(zipAdapter.hasFile('word/charts/chart2.xml')).toBe(false);
  });

  it('update via executeChart dispatch also returns stale refs and leaves body untouched', async () => {
    const { parser, ast } = await setupDocWithBody();
    const res: any = executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: 'chart-r14',
        title: 'July Revenue Update',
        categories: ['Jul'],
        series: [{ name: 'Revenue', values: [7500], color: '0066CC' }],
      },
      ast,
      parser
    );
    expect(res).toBeDefined();
    expect(Array.isArray(res.possiblyStaleBodyRefs)).toBe(true);
    const ids = res.possiblyStaleBodyRefs.map((r: any) => r.nodeId);
    expect(ids).toContain('cap-r14');
    expect(ids).toContain('cta-r14');
    expect(ids).not.toContain('unrelated-r14');
    expect(nodeText(ast, 'cap-r14')).toBe(CAPTION);
    expect(nodeText(ast, 'cta-r14')).toBe(CTA);
    const zipAdapter = parser.getZipAdapter();
    expect(zipAdapter.hasFile('word/charts/chart1.xml')).toBe(true);
    expect(zipAdapter.hasFile('word/charts/chart2.xml')).toBe(false);
  });

  it('no echoes in body returns empty refs with no false positives', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast: any = await parser.getAST();
    const anchor: string = ast.sections[0].nodes[0].id;
    executeChart(
      {
        action: 'chart',
        chartType: 'column',
        title: 'Alpha Summary',
        categories: ['Alpha', 'Beta'],
        series: [{ name: 'S', values: [150, 250], color: '0066CC' }],
        anchor,
        id: 'chart-clean',
      },
      ast,
      parser
    );
    await executeInsert(
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: UNRELATED, id: 'plain-1' },
      ast,
      parser
    );
    const res: any = executeChartUpdate(
      {
        action: 'chart',
        operation: 'update',
        target: 'chart-clean',
        title: 'Gamma Summary',
        categories: ['Gamma'],
        series: [{ name: 'S', values: [350], color: '0066CC' }],
      },
      ast,
      parser
    );
    expect(res.possiblyStaleBodyRefs).toEqual([]);
    expect(nodeText(ast, 'plain-1')).toBe(UNRELATED);
    const zipAdapter = parser.getZipAdapter();
    expect(zipAdapter.hasFile('word/charts/chart2.xml')).toBe(false);
  });
});
