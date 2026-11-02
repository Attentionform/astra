import { describe, it, expect } from 'vitest';
import {
  buildEmbeddedNodePages,
  buildPageMap,
  lintVisualDefects,
  type DocumentAST,
} from '@attentionform/astra-engine';
import { buildCliVisualTelemetry } from '../../packages/cli/src/cli/visual-cmd';

/**
 * A node that fills its page used to open the next page at once, so a document whose last node ended a page
 * got one page too many (and a blank-page warning for it). The total is the highest page that holds a node.
 */
const LINES = 45; // lines per page with no sectPr

/** A table that takes `LINES` lines by itself: rows + 2 header/spacing lines. */
const fullPageTable = (id: string) => ({ id, type: 'table', metadata: { tableDimensions: { rows: LINES - 2 } } });
const paragraph = (id: string, text = 'a short paragraph') => ({ id, type: 'paragraph', text });

function astOf(nodes: Array<{ id: string; type: string; metadata?: Record<string, unknown> }>): DocumentAST {
  return { sections: [{ id: 'sec_root', nodes }] } as unknown as DocumentAST;
}

describe('buildEmbeddedNodePages total', () => {
  it('a last node that fills the page does not open another page', () => {
    const r = buildEmbeddedNodePages([fullPageTable('t1')]);
    expect(r.totalPages).toBe(1);
    expect(r.nodePages['t1']).toBe(1);
  });

  it('two full pages are two pages, not three', () => {
    const r = buildEmbeddedNodePages([fullPageTable('t1'), fullPageTable('t2')]);
    expect(r.totalPages).toBe(2);
    expect(r.nodePages).toEqual({ t1: 1, t2: 2 });
  });

  it('a node after a full page still lands on the next page and counts it', () => {
    const r = buildEmbeddedNodePages([fullPageTable('t1'), paragraph('p1')]);
    expect(r.totalPages).toBe(2);
    expect(r.nodePages['p1']).toBe(2);
  });

  it('an explicit page break puts the break and the content after it on the next page', () => {
    const r = buildEmbeddedNodePages([paragraph('p1'), { id: 'b1', type: 'pageBreak' }, paragraph('p2')]);
    expect(r.totalPages).toBe(2);
    expect(r.nodePages).toEqual({ p1: 1, b1: 2, p2: 2 });
  });

  it('a page break as the last node still counts its own page', () => {
    const r = buildEmbeddedNodePages([paragraph('p1'), { id: 'b1', type: 'pageBreak' }]);
    expect(r.totalPages).toBe(2);
  });

  it('no nodes is one page', () => {
    expect(buildEmbeddedNodePages([]).totalPages).toBe(1);
  });

  it('an ordinary short document is unchanged', () => {
    const r = buildEmbeddedNodePages([paragraph('p1'), paragraph('p2')]);
    expect(r.totalPages).toBe(1);
  });
});

describe('buildPageMap estimate total', () => {
  it('follows the same rule', () => {
    expect(buildPageMap(astOf([fullPageTable('t1')])).totalPages).toBe(1);
    expect(buildPageMap(astOf([fullPageTable('t1'), fullPageTable('t2')])).totalPages).toBe(2);
    const withTail = buildPageMap(astOf([fullPageTable('t1'), { id: 'p1', type: 'paragraph', metadata: { text: 'tail' } }]));
    expect(withTail.totalPages).toBe(2);
    expect(withTail.source).toBe('estimate');
  });
});

describe('lintVisualDefects blank pages', () => {
  const node = (id: string) => ({ id, type: 'paragraph', metadata: { text: 'text' } });
  const map = (totalPages: number, nodePages: Record<string, number>) => ({
    totalPages,
    nodePages,
    source: 'estimate' as const,
    isEstimate: true,
  });

  it('does not report a trailing page after the last node', () => {
    const warnings = lintVisualDefects(astOf([node('a'), node('b')]), map(3, { a: 1, b: 2 }));
    expect(warnings.filter((w) => w.startsWith('Blank page'))).toEqual([]);
  });

  it('reports a blank page between two populated pages', () => {
    const warnings = lintVisualDefects(astOf([node('a'), node('b')]), map(3, { a: 1, b: 3 }));
    expect(warnings.filter((w) => w.startsWith('Blank page'))).toEqual(['Blank page: page 2 of 3 has no content nodes.']);
  });

  it('still reports a document with no node at all as blank', () => {
    const warnings = lintVisualDefects(astOf([]), map(1, {}));
    expect(warnings.filter((w) => w.startsWith('Blank page'))).toEqual(['Blank page: page 1 of 1 has no content nodes.']);
  });
});

describe('buildCliVisualTelemetry page-map source', () => {
  const ast = astOf([
    { id: 'a', type: 'paragraph', metadata: { text: 'first' } },
    { id: 'b', type: 'paragraph', metadata: { text: 'second' } },
  ]);
  const summary = { nodePages: { a: 1, b: 2 }, totalPages: 2, pageMapSource: 'estimate' };

  it('labels the embedded map as an estimate', () => {
    const t = buildCliVisualTelemetry(ast, summary, 1);
    expect(t.mapSource).toBe('estimate');
    expect(t.wireframe).toContain('[PAGE 1 OF 2 (estimate)]');
  });

  it('keeps a renderer map labelled as the renderer', () => {
    const t = buildCliVisualTelemetry(ast, { ...summary, pageMapSource: 'renderer' }, 1);
    expect(t.mapSource).toBe('renderer');
    expect(t.wireframe).toContain('[PAGE 1 OF 2]');
  });

  it('treats a summary without a source label as a renderer map, as before', () => {
    const t = buildCliVisualTelemetry(ast, { nodePages: summary.nodePages, totalPages: 2 }, 1);
    expect(t.mapSource).toBe('renderer');
    expect(t.wireframe).not.toContain('(estimate)');
  });

  it('does not label the renderer count as an estimate when the render disagrees with the map', () => {
    const t = buildCliVisualTelemetry(ast, summary, 1, 3);
    expect(t.wireframe).toContain('[PAGE 1 OF 3]');
    expect(t.wireframe).not.toContain('(estimate)');
  });

  it('a render of 3 pages against an estimate of 2 prints the mismatch note', () => {
    const t = buildCliVisualTelemetry(ast, summary, 1, 3);
    expect(t.totalPages).toBe(3);
    expect(t.wordLayoutTotalPages).toBe(2);
    expect(t.warnings.some((w) => w.startsWith('Pagination mismatch: renderer produced 3 pages; word-layout page map totals 2'))).toBe(true);
  });

  it('a render that agrees with the estimate prints no mismatch note', () => {
    const t = buildCliVisualTelemetry(ast, summary, 1, 2);
    expect(t.warnings.some((w) => w.startsWith('Pagination mismatch'))).toBe(false);
  });
});
