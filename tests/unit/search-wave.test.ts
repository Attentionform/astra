import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  executeChart,
  executeEquation,
  executeFootnote,
  executeReplace,
} from '@attentionform/astra-engine';

async function freshBlank() {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  return { parser, ast };
}

describe('search projection covers chart, equation, note, and revision text', () => {
  it('post-chart-update search hits updated categories and series values', async () => {
    const { parser, ast } = await freshBlank();
    executeChart({
      action: 'chart',
      chartType: 'column',
      id: 'chart_search_probe',
      title: 'Subscriber Growth',
      categories: ['Jan', 'Feb'],
      series: [{ name: 'Active Subscribers', values: [1250, 2300], color: '0066CC' }],
    } as never, ast, parser);

    // Pre-update terms are indexed from the insert metadata.
    expect(parser.searchNodesInAST(ast, 'Subscriber').map((n) => n.id)).toContain('chart_search_probe');

    executeChart({
      action: 'chart',
      operation: 'update',
      target: 'chart_search_probe',
      categories: ['Jun', 'Jul'],
      series: [{ name: 'Active Subscribers', values: [6050, 6900], color: '10B981' }],
    } as never, ast, parser);

    const julHits = parser.searchNodesInAST(ast, 'Jul');
    expect(julHits.map((n) => n.id)).toContain('chart_search_probe');
    const valueHits = parser.searchNodesInAST(ast, '6900');
    expect(valueHits.map((n) => n.id)).toContain('chart_search_probe');
    // Stale pre-update terms no longer match the chart node.
    expect(parser.searchNodesInAST(ast, 'Jan').map((n) => n.id)).not.toContain('chart_search_probe');
  });

  it('reload-path chart search hits title, categories, and series from the chart part', async () => {
    const { parser, ast } = await freshBlank();
    executeChart({
      action: 'chart',
      chartType: 'column',
      id: 'chart_reload_probe',
      title: 'Jul Revenue Deep Dive',
      categories: ['Jul', 'Aug'],
      series: [{ name: 'RevTracker', values: [6900, 7100], color: '0066CC' }],
    } as never, ast, parser);

    const reloaded = new DocxParser(parser.rezip());
    const reloadedAst = await reloaded.getAST();
    const chartNode = reloadedAst.sections[0].nodes.find(
      (n) => n.type === 'chart' || n.type === 'chart3d',
    );
    expect(chartNode).toBeDefined();
    // Reloaded chart paragraphs carry no w:t, so the base text is empty.
    expect(chartNode!.metadata.text).toBe('');
    expect(reloaded.searchNodesInAST(reloadedAst, 'Deep Dive').map((n) => n.id)).toContain(chartNode!.id);
    expect(reloaded.searchNodesInAST(reloadedAst, 'Jul').map((n) => n.id)).toContain(chartNode!.id);
    expect(reloaded.searchNodesInAST(reloadedAst, '6900').map((n) => n.id)).toContain(chartNode!.id);
    expect(reloaded.searchNodesInAST(reloadedAst, 'RevTracker').map((n) => n.id)).toContain(chartNode!.id);
  });

  it('equation term hits in-memory and after reload', async () => {
    const { parser, ast } = await freshBlank();
    const anchor = ast.sections[0].nodes[0].id;
    const out = (await executeEquation(
      { action: 'equation', operation: 'insert', anchor, latex: 'NavierStokesMomentumTerm' } as never,
      ast,
      parser,
    )) as { id: string };
    expect(parser.searchNodesInAST(ast, 'NavierStokes').map((n) => n.id)).toContain(out.id);

    const reloaded = new DocxParser(parser.rezip());
    const reloadedAst = await reloaded.getAST();
    const hits = reloaded.searchNodesInAST(reloadedAst, 'NavierStokes');
    expect(hits.length).toBeGreaterThan(0);
    expect(hits[0].type).toBe('equation');
  });

  it('footnote body term hits the anchor paragraph', async () => {
    const { parser, ast } = await freshBlank();
    const anchor = ast.sections[0].nodes[0].id;
    executeFootnote(
      { action: 'footnote', operation: 'insert', anchor, text: 'FootnoteBodyZebraQuasar' } as never,
      ast,
      parser,
    );
    const hits = parser.searchNodesInAST(ast, 'ZebraQuasar');
    expect(hits.map((n) => n.id)).toContain(anchor);

    // Same anchor resolves after a save/reload round-trip.
    const reloaded = new DocxParser(parser.rezip());
    const reloadedAst = await reloaded.getAST();
    expect(reloaded.searchNodesInAST(reloadedAst, 'ZebraQuasar').map((n) => n.id)).toContain(anchor);
  });

  it('endnote body term hits the referencing paragraph', async () => {
    const { parser, ast } = await freshBlank();
    const anchor = ast.sections[0].nodes[0].id;
    const before = parser.readNode(anchor);
    expect(before).toContain('</w:p>');
    const withRef = before.replace(
      '</w:p>',
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="7"/></w:r></w:p>',
    );
    parser.updateNode(anchor, withRef);
    parser.getZipAdapter().writeText(
      'word/endnotes.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
        '<w:endnote w:id="7"><w:p><w:r><w:t>EndnoteBodyCometHarbor</w:t></w:r></w:p></w:endnote>' +
        '</w:endnotes>',
    );
    const live = await parser.getAST();
    expect(parser.searchNodesInAST(live, 'CometHarbor').map((n) => n.id)).toContain(anchor);
  });

  it('tracked-deletion text hits the host node flagged as a revision', async () => {
    const { parser, ast } = await freshBlank();
    const anchor = ast.sections[0].nodes[0].id;
    executeReplace(
      { action: 'replace', target: anchor, text: 'UniqueDelMarkerKiloWatt', trackChanges: false } as never,
      ast,
    );
    parser.deleteTrackedChange(anchor, 'ASTra', new Date().toISOString());
    const live = await parser.getAST();
    const host = live.sections[0].nodes.find((n) => n.id === anchor);
    expect(host).toBeDefined();
    // The visible w:t is gone; only w:delText carries the deleted wording.
    expect(host!.metadata.text).toBe('');
    const hits = parser.searchNodesInAST(live, 'UniqueDelMarker');
    expect(hits.map((n) => n.id)).toContain(anchor);
    expect(hits.find((n) => n.id === anchor)!.metadata.hasTrackedChanges).toBe(true);
  });

  it('legacy behavior intact: case-insensitive substring, no-match empty, bad input throws', async () => {
    const { parser, ast } = await freshBlank();
    const lower = parser.searchNodesInAST(ast, 'blank document').map((n) => n.id);
    const upper = parser.searchNodesInAST(ast, 'BLANK DOCUMENT').map((n) => n.id);
    expect(lower).toEqual(upper);
    expect(lower.length).toBeGreaterThan(0);
    expect(parser.searchNodesInAST(ast, 'zzz-no-such-term-qwerty')).toEqual([]);
    expect(() => (parser.searchNodesInAST as any)(ast, undefined)).toThrow(TypeError);
    expect(() => (parser.searchNodesInAST as any)(ast, 123)).toThrow(TypeError);
  });
});
