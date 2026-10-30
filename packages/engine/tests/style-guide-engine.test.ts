/**
 * Style guide engine pins: omitted-param defaulting (tables, charts,
 * callouts, code blocks), the advisory linter, and named/file resolution.
 *
 * Precedence under test everywhere: explicit op param > guide > built-in.
 * Guide-absent runs pin the legacy bytes (no guide = no behavior change).
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import fs from 'fs';
import path from 'path';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { DocumentManager } from '../src/document-manager';
import { OperationEngine } from '../src/operation-engine';
import { validateOperationsStrict } from '../src/operation-schema';
import { composeDocument } from '../src/compose/markdown';
import type { StyleGuideConfig } from '../src/style-guide/types';
import { lintDocumentStyle } from '../src/style-guide/lint';
import { resolveStyleGuide } from '../src/style-guide/resolve';
import type { Node } from '../src/ast';

const GUIDE_FIXTURE: StyleGuideConfig = {
  name: 'Engine Test',
  version: 1,
  palette: {
    primary: '1E3A8A',
    chartPalette: ['1E3A8A', '0284C7', '10B981', 'F59E0B'],
  },
  tables: {
    headerFill: '1E3A8A',
    headerText: 'FFFFFF',
    keepTogether: true,
  },
  typography: {},
  callouts: { defaultTheme: 'slate' },
  codeBlocks: { background: 'F1F5F9', borderColor: 'CBD5E1' },
  layout: {},
};

function freshGuide(): StyleGuideConfig {
  return JSON.parse(JSON.stringify(GUIDE_FIXTURE)) as StyleGuideConfig;
}

async function freshDoc() {
  const parser = new DocxParser(createBlankDocxBuffer('Guide Test'));
  const ast = await parser.getAST();
  return { parser, ast };
}

function lastNodeId(ast: { sections: Array<{ nodes: Node[] }> }): string {
  const nodes = ast.sections[0].nodes;
  return nodes[nodes.length - 1].id;
}

function xmlOf(node: Node): string {
  return new XMLSerializer().serializeToString(node.xml);
}

function tableNodeOf(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const found = ast.sections[0].nodes.find((n) => n.type === 'table');
  if (!found) throw new Error('test setup: no table node found');
  return found;
}

async function insertTable(engine: OperationEngine, anchor: string) {
  await engine.execute({
    action: 'insert',
    type: 'table',
    position: 'after',
    anchor,
    headers: ['Metric', 'Value'],
    rows: [
      ['Alpha', '1'],
      ['Beta', '2'],
    ],
  } as never);
}

describe('table taste defaulting', () => {
  it('shades the header with guide text when textColor is omitted', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const result = await engine.execute({
      action: 'table',
      target: table.id,
      operation: 'setShading',
      headerFill: '1E3A8A',
      fill: 'F7F9FC',
    } as never);
    expect(result.warnings).toEqual([]);
    expect(xmlOf(table)).toContain('w:val="FFFFFF"');
  });

  it('leaves header text alone without a guide (legacy bytes)', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    await engine.execute({
      action: 'table',
      target: table.id,
      operation: 'setShading',
      headerFill: '1E3A8A',
      fill: 'F7F9FC',
    } as never);
    expect(xmlOf(table)).not.toContain('w:val="FFFFFF"');
  });

  it('lets an explicit textColor beat the guide', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    await engine.execute({
      action: 'table',
      target: table.id,
      operation: 'setShading',
      headerFill: '1E3A8A',
      fill: 'F7F9FC',
      textColor: 'FF0000',
    } as never);
    const xml = xmlOf(table);
    expect(xml).toContain('w:val="FF0000"');
    expect(xml).not.toContain('w:val="FFFFFF"');
  });

  it('computes the contrast ratio instead of matching the header fill', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    // Light fill: dark text already contrasts, so no guide text is applied.
    await engine.execute({
      action: 'table',
      target: table.id,
      operation: 'setCellShading',
      rowIndex: 1,
      colIndex: 1,
      fill: 'F8FAFC',
    } as never);
    expect(xmlOf(table)).not.toContain('w:val="FFFFFF"');
    // Dark fill: the same guide applies its header text.
    await engine.execute({
      action: 'table',
      target: table.id,
      operation: 'setCellShading',
      rowIndex: 1,
      colIndex: 2,
      fill: '1E3A8A',
    } as never);
    expect(xmlOf(table)).toContain('w:val="FFFFFF"');
  });

  it('rejects a malformed explicit textColor before mutating', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    await expect(
      engine.execute({
        action: 'table',
        target: table.id,
        operation: 'setShading',
        headerFill: '1E3A8A',
        textColor: 'not-a-color',
      } as never),
    ).rejects.toThrow(/textColor/);
  });
});

describe('table keep-together on insert', () => {
  it('chains rows with keepNext when the guide asks for it', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await insertTable(engine, lastNodeId(ast));
    const xml = xmlOf(tableNodeOf(ast));
    expect(xml).toContain('w:tblHeader');
    expect(xml).toContain('w:cantSplit');
    expect(xml).toContain('w:keepNext');
  });

  it('keeps legacy markers without keepNext when no guide is active', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await insertTable(engine, lastNodeId(ast));
    const xml = xmlOf(tableNodeOf(ast));
    expect(xml).toContain('w:tblHeader');
    expect(xml).toContain('w:cantSplit');
    expect(xml).not.toContain('w:keepNext');
  });

  it('honors an explicit keepTogether without any guide', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await engine.execute({
      action: 'insert',
      type: 'table',
      position: 'after',
      anchor: lastNodeId(ast),
      headers: ['A'],
      rows: [['1']],
      keepTogether: true,
    } as never);
    expect(xmlOf(tableNodeOf(ast))).toContain('w:keepNext');
  });
});

describe('chart palette defaulting', () => {
  const series3 = [
    { name: 'S1', values: [1, 2] },
    { name: 'S2', values: [3, 4] },
    { name: 'S3', values: [5, 6] },
  ];

  it('fills uncolored series from the guide palette in order', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Palette',
      categories: ['A', 'B'],
      series: JSON.parse(JSON.stringify(series3)),
      anchor: lastNodeId(ast),
    } as never);
    const chartXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(chartXml).toContain('val="1E3A8A"');
    expect(chartXml).toContain('val="0284C7"');
    expect(chartXml).toContain('val="10B981"');
    expect(chartXml.indexOf('val="1E3A8A"') < chartXml.indexOf('val="0284C7"')).toBe(true);
    expect(chartXml.indexOf('val="0284C7"') < chartXml.indexOf('val="10B981"')).toBe(true);
  });

  it('never shifts palette indices for explicitly colored series', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Explicit middle',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [1, 2] },
        { name: 'S2', values: [3, 4], color: 'FF0000' },
        { name: 'S3', values: [5, 6] },
      ],
      anchor: lastNodeId(ast),
    } as never);
    const chartXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(chartXml).toContain('val="1E3A8A"');
    expect(chartXml).toContain('val="FF0000"');
    // Third series still takes palette[2], not palette[1].
    expect(chartXml).toContain('val="10B981"');
    expect(chartXml).not.toContain('val="0284C7"');
  });

  it('passes series through untouched without a guide', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Legacy',
      categories: ['A', 'B'],
      series: JSON.parse(JSON.stringify(series3)),
      anchor: lastNodeId(ast),
    } as never);
    const chartXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(chartXml).toContain('val="0066CC"');
    expect(chartXml).not.toContain('val="1E3A8A"');
  });

  it('fills uncolored series on update without moving stored colors', async () => {
    const { parser, ast } = await freshDoc();
    const plain = new OperationEngine(ast, parser);
    await plain.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Updatable',
      categories: ['A', 'B'],
      series: [{ name: 'Old', values: [1, 2], color: 'ABCDEF' }],
      anchor: lastNodeId(ast),
    } as never);
    const chartId = ast.sections[0].nodes.find((n) => n.type === 'chart')?.id;
    if (!chartId) throw new Error('test setup: chart node missing');
    const guided = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await guided.execute({
      action: 'chart',
      operation: 'update',
      target: chartId,
      series: [
        { name: 'Kept', values: [7, 8] },
        { name: 'Fresh', values: [9, 10] },
      ],
    } as never);
    const chartXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    // Stored ABCDEF counts as explicit for series 0; series 1 takes palette[1].
    expect(chartXml).toContain('val="ABCDEF"');
    expect(chartXml).toContain('val="0284C7"');
  });

  it('keeps an explicit #RGB color as valid expanded 6-hex', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Shorthand',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [1, 2], color: '#09C' },
        { name: 'S2', values: [3, 4] },
      ],
      anchor: lastNodeId(ast),
    } as never);
    const chartXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    // Shorthand survives (not palette-overwritten) and is written expanded.
    expect(chartXml).toContain('val="0099CC"');
    expect(chartXml).not.toContain('val="09C"');
    expect(chartXml).toContain('val="0284C7"');
  });

  it('gives 7 series on a 6-palette distinct colors with no duplicate warning', async () => {
    const guide = freshGuide();
    guide.palette.chartPalette = ['1E3A8A', '0284C7', '10B981', 'F59E0B', 'EF4444', '8B5CF6'];
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: guide });
    const series7 = Array.from({ length: 7 }, (_, i) => ({ name: `S${i}`, values: [i + 1, i + 2] }));
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Overflow',
      categories: ['A', 'B'],
      series: series7,
      anchor: lastNodeId(ast),
    } as never);
    const chartXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const fills = Array.from(chartXml.matchAll(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/g)).map((m) => m[1].toUpperCase());
    expect(fills).toHaveLength(7);
    expect(new Set(fills).size).toBe(7);
    const dups = lintDocumentStyle(ast, guide).filter((w) => w.rule === 'chart-duplicate-colors');
    expect(dups).toEqual([]);
  });

  it('keeps 12 series on a 6-palette all distinct and valid', async () => {
    const guide = freshGuide();
    guide.palette.chartPalette = ['1E3A8A', '0284C7', '10B981', 'F59E0B', 'EF4444', '8B5CF6'];
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: guide });
    const series12 = Array.from({ length: 12 }, (_, i) => ({ name: `S${i}`, values: [i + 1, i + 2] }));
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Deep overflow',
      categories: ['A', 'B'],
      series: series12,
      anchor: lastNodeId(ast),
    } as never);
    const chartXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const fills = Array.from(chartXml.matchAll(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/g)).map((m) => m[1].toUpperCase());
    expect(fills).toHaveLength(12);
    for (const f of fills) expect(f).toMatch(/^[0-9A-F]{6}$/);
    expect(new Set(fills).size).toBe(12);
    const dups = lintDocumentStyle(ast, guide).filter((w) => w.rule === 'chart-duplicate-colors');
    expect(dups).toEqual([]);
  });
});

describe('callout theme defaulting', () => {
  it('uses the guide default when theme is omitted', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    const result = await engine.executeBatch({
      operations: [
        { action: 'callout', anchor: lastNodeId(ast), text: 'Guide note' } as never,
      ],
    });
    expect(result.success).toBe(true);
    const node = ast.sections[0].nodes.find(
      (n) => (n.metadata as { callout?: { title?: string } }).callout !== undefined,
    );
    expect(node).toBeDefined();
    // Slate border, not legacy indigo.
    expect(xmlOf(node as Node)).toContain('w:color="475569"');
    expect(xmlOf(node as Node)).not.toContain('w:color="4F46E5"');
  });

  it('keeps legacy indigo without a guide', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await engine.executeBatch({
      operations: [
        { action: 'callout', anchor: lastNodeId(ast), text: 'Legacy note' } as never,
      ],
    });
    const node = ast.sections[0].nodes.find(
      (n) => (n.metadata as { callout?: { title?: string } }).callout !== undefined,
    );
    expect(xmlOf(node as Node)).toContain('w:color="4F46E5"');
  });

  it('falls back loudly on an unknown theme instead of silently', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    const result = await engine.executeBatch({
      operations: [
        { action: 'callout', anchor: lastNodeId(ast), text: 'Odd', theme: 'mauve' } as never,
      ],
    });
    expect(result.success).toBe(true);
    expect(result.warnings.some((w) => w.includes('mauve') && w.includes('slate'))).toBe(true);
    const node = ast.sections[0].nodes.find(
      (n) => (n.metadata as { callout?: { title?: string } }).callout !== undefined,
    );
    expect(xmlOf(node as Node)).toContain('w:color="475569"');
  });
});

describe('code block skin', () => {
  it('paints guide background and borders through the insert path', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await engine.execute({
      action: 'insert',
      type: 'paragraph',
      position: 'after',
      anchor: lastNodeId(ast),
      text: 'const x = 1;',
      params: { codeBlock: true, codeLanguage: 'ts' },
    } as never);
    const node = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
    const xml = xmlOf(node);
    expect(xml).toContain('w:fill="F1F5F9"');
    expect(xml).toContain('w:color="CBD5E1"');
  });

  it('keeps legacy grays without a guide', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await engine.execute({
      action: 'insert',
      type: 'paragraph',
      position: 'after',
      anchor: lastNodeId(ast),
      text: 'const x = 1;',
      params: { codeBlock: true },
    } as never);
    const node = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
    const xml = xmlOf(node);
    expect(xml).toContain('w:fill="F2F2F2"');
    expect(xml).toContain('w:color="D1D5DB"');
  });

  it('skins fenced code through the compose path', async () => {
    const { parser } = await freshDoc();
    const result = await composeDocument(parser, '```ts\nconst x = 1;\n```\n', {
      styleGuide: freshGuide(),
    });
    expect(result.success).toBe(true);
    const ast = await parser.getAST();
    const dump = ast.sections[0].nodes.map(xmlOf).join('');
    expect(dump).toContain('w:fill="F1F5F9"');
  });
});

describe('advisory linter', () => {
  it('warns on dark-on-dark headers without blocking the batch', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser, { styleGuide: freshGuide() });
    await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const result = await engine.execute({
      action: 'table',
      target: table.id,
      operation: 'setShading',
      headerFill: '1E3A8A',
      fill: 'F7F9FC',
      textColor: '000000',
    } as never);
    // Single-op path applies (no throw) and reports its advisories.
    expect(result.warnings.length).toBeGreaterThan(0);
    const findings = lintDocumentStyle(ast, freshGuide());
    const contrast = findings.filter((w) => w.rule === 'table-header-contrast');
    expect(contrast).toHaveLength(1);
    expect(contrast[0].nodeId).toBe(table.id);
    expect(contrast[0].message).toContain('1E3A8A');
    expect(contrast[0].message).toContain('FFFFFF');
  });

  it('flags duplicate chart colors per chart node', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Twins',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [1, 2], color: '0066CC' },
        { name: 'S2', values: [3, 4], color: '0066CC' },
      ],
      anchor: lastNodeId(ast),
    } as never);
    const chartId = ast.sections[0].nodes.find((n) => n.type === 'chart')?.id;
    const findings = lintDocumentStyle(ast, freshGuide());
    const dups = findings.filter((w) => w.rule === 'chart-duplicate-colors');
    expect(dups).toHaveLength(1);
    expect(dups[0].nodeId).toBe(chartId);
    expect(dups[0].message).toContain('0066CC');
  });

  it('flags tables whose rows can split across pages', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await insertTable(engine, lastNodeId(ast));
    // Fresh inserts carry cantSplit everywhere, so the baseline is clean.
    expect(
      lintDocumentStyle(ast, freshGuide()).filter((w) => w.rule === 'table-orphan'),
    ).toEqual([]);
    const table = tableNodeOf(ast);
    const trPrs = Array.from((table.xml as Element).getElementsByTagName('w:trPr'));
    for (const trPr of trPrs) {
      for (const cs of Array.from((trPr as Element).getElementsByTagName('w:cantSplit'))) {
        (trPr as Element).removeChild(cs);
      }
    }
    const findings = lintDocumentStyle(ast, freshGuide());
    const orphans = findings.filter((w) => w.rule === 'table-orphan');
    expect(orphans).toHaveLength(1);
    expect(orphans[0].nodeId).toBe(table.id);
    expect(orphans[0].message).toContain('setKeepTogether');
  });

  it('stays silent on splitable rows when the guide disables keepTogether', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const trPrs = Array.from((table.xml as Element).getElementsByTagName('w:trPr'));
    for (const trPr of trPrs) {
      for (const cs of Array.from((trPr as Element).getElementsByTagName('w:cantSplit'))) {
        (trPr as Element).removeChild(cs);
      }
    }
    const flagged = lintDocumentStyle(ast, freshGuide()).filter((w) => w.rule === 'table-orphan');
    expect(flagged).toHaveLength(1);
    const relaxed = freshGuide();
    relaxed.tables.keepTogether = false;
    expect(lintDocumentStyle(ast, relaxed).filter((w) => w.rule === 'table-orphan')).toEqual([]);
  });

  it('stays silent without a guide and never throws on odd input', async () => {
    const { ast } = await freshDoc();
    expect(lintDocumentStyle(ast, undefined)).toEqual([]);
    expect(lintDocumentStyle({ sections: [] } as never, freshGuide())).toEqual([]);
    expect(
      lintDocumentStyle({ sections: [{ nodes: [{ id: 'x' }] }] } as never, freshGuide()),
    ).toEqual([]);
  });
});

describe('guide resolution', () => {
  let scratchDir: string;

  beforeEach(() => {
    const root = '/tmp/opencode';
    if (!fs.existsSync(root)) fs.mkdirSync(root, { recursive: true });
    scratchDir = fs.mkdtempSync(path.join(root, 'astra-guide-'));
  });

  afterEach(() => {
    if (scratchDir && fs.existsSync(scratchDir)) {
      fs.rmSync(scratchDir, { recursive: true, force: true });
    }
  });

  it('reads a file guide with parser diagnostics attached', () => {
    const file = path.join(scratchDir, 'brand.md');
    fs.writeFileSync(
      file,
      '# Style Guide: Probe\n\n## Palette\n- Primary: #123456\n- Chart Palette: #111111, #222222, #333333, #444444\n',
    );
    const resolved = resolveStyleGuide('./brand.md', { allowedDir: scratchDir });
    expect(resolved.config.palette.primary).toBe('123456');
    expect(resolved.warnings.length).toBeGreaterThan(0);
    expect(resolved.filePath.endsWith('brand.md')).toBe(true);
  });

  it('fails closed on unknown names, missing files, remote urls, and jail escapes', () => {
    expect(() => resolveStyleGuide('definitely-not-a-guide')).toThrow(/unknown style guide/);
    try {
      resolveStyleGuide('definitely-not-a-guide');
      expect.unreachable();
    } catch (err) {
      expect((err as { code?: string }).code).toBe('E_VALIDATION');
    }
    expect(() => resolveStyleGuide('./missing.md', { allowedDir: scratchDir })).toThrow(
      /not found/,
    );
    expect(() => resolveStyleGuide('https://example.com/g.md')).toThrow(/remote URL/);
    expect(() => resolveStyleGuide('../escape.md', { allowedDir: scratchDir })).toThrow(
      /outside the allowed directory/,
    );
    try {
      resolveStyleGuide('../escape.md', { allowedDir: scratchDir });
      expect.unreachable();
    } catch (err) {
      expect((err as { code?: string }).code).toBe('E_SANDBOX');
    }
  });
});

describe('schema gate agreement', () => {
  it('accepts the new explicit params the engine honors', () => {
    expect(() =>
      validateOperationsStrict([
        {
          action: 'table',
          target: 't_001',
          operation: 'setShading',
          headerFill: '1E3A8A',
          textColor: 'FFFFFF',
        },
      ] as never),
    ).not.toThrow();
    expect(() =>
      validateOperationsStrict([
        {
          action: 'insert',
          type: 'table',
          anchor: 'p_001',
          position: 'after',
          headers: ['A'],
          rows: [['1']],
          keepTogether: true,
        },
      ] as never),
    ).not.toThrow();
  });
});

describe('save and validate carry style warnings', () => {
  let scratchDir: string;

  beforeEach(() => {
    const root = '/tmp/opencode';
    if (!fs.existsSync(root)) fs.mkdirSync(root, { recursive: true });
    scratchDir = fs.mkdtempSync(path.join(root, 'astra-manager-'));
  });

  afterEach(() => {
    if (scratchDir && fs.existsSync(scratchDir)) {
      fs.rmSync(scratchDir, { recursive: true, force: true });
    }
  });

  it('warns without blocking when a guide is set, stays quiet otherwise', async () => {
    // Build the offending bytes through a headless engine first.
    const builder = new DocxParser(createBlankDocxBuffer('Manager Probe'));
    let ast = await builder.getAST();
    const engine = new OperationEngine(ast, builder, { styleGuide: freshGuide() });
    const anchor = lastNodeId(ast);
    await engine.execute({
      action: 'insert',
      type: 'table',
      position: 'after',
      anchor,
      headers: ['H'],
      rows: [['body']],
    } as never);
    ast = await builder.getAST();
    const tableId = ast.sections[0].nodes.find((n) => n.type === 'table')?.id as string;
    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setShading',
      headerFill: '1E3A8A',
      fill: 'F7F9FC',
      textColor: '000000',
    } as never);

    const file = path.join(scratchDir, 'probe.docx');
    fs.writeFileSync(file, builder.rezip());

    const manager = new DocumentManager();
    await manager.openDocument(file);
    manager.setStyleGuide(freshGuide());
    const saved = await manager.saveDocument();
    expect(saved.styleWarnings.some((w) => w.rule === 'table-header-contrast')).toBe(true);
    const validated = await manager.validateCurrentDocument();
    expect(validated.styleWarnings.some((w) => w.rule === 'table-header-contrast')).toBe(true);

    manager.setStyleGuide(undefined);
    const quiet = await manager.saveDocument();
    expect(quiet.styleWarnings).toEqual([]);
  });
});
