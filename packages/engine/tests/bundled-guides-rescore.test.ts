/**
 * Bundled guide re-score: the three guides under
 * `packages/engine/templates/guides/` parse zero-warning, named resolution
 * succeeds for each, and rebuilt validation-job-shaped documents (status
 * report, NDA, API reference, proposal) carry zero style warnings on save
 * and validate.
 *
 * Render check (headless LibreOffice 25.2.7.2, 2026-10-15, PDF/PNG eyeballed
 * by the implementer — all four PASS):
 *   - status report (report guide): PASS — navy/white header, three distinct
 *     chart series, both callouts shaded; no contrast or overflow defects.
 *   - NDA (legal guide): PASS — grayscale, white-on-midnight signature
 *     header, signature rows intact on one page.
 *   - API reference (technical guide): PASS — code fences shaded (no naked
 *     code), both tables readable, footnote marker plus footnote present.
 *     Tables keep the compose default gray headers, still warning-free.
 *   - proposal (legal guide): PASS — grayscale options table, recommendation
 *     callout shaded. Note: the blank-document seed subtitle
 *     ("Proposal template ...") still prints above the composed title;
 *     pre-existing seed behavior, not a guide defect.
 *
 * Debug aid only: set ASTRA_KEEP_RESCORE=1 to keep the emitted .docx files
 * (paths print to stdout) instead of removing the scratch dir.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { fileURLToPath } from 'url';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { DocumentManager } from '../src/document-manager';
import { OperationEngine } from '../src/operation-engine';
import { composeDocument } from '../src/compose/markdown';
import { parseStyleGuide } from '../src/style-guide/parser';
import {
  resolveStyleGuide,
  listBundledGuides,
  bundledGuideCandidates,
} from '../src/style-guide/resolve';
import {
  lintDocumentStyle,
  chartXmlGetterForAst,
} from '../src/style-guide/lint';
import type { StyleGuideConfig } from '../src/style-guide/types';
import type { Node } from '../src/ast';

const GUIDE_NAMES = ['report', 'technical', 'legal'] as const;

function guideDir(): string {
  const fromCwd = path.resolve(process.cwd(), 'packages/engine/templates/guides');
  if (fs.existsSync(fromCwd)) return fromCwd;
  const here = path.dirname(fileURLToPath(import.meta.url));
  const fromTest = path.resolve(here, '..', 'templates', 'guides');
  if (fs.existsSync(fromTest)) return fromTest;
  throw new Error(`bundled guide dir missing (tried ${fromCwd} and ${fromTest})`);
}

function readGuide(name: string): string {
  return fs.readFileSync(path.join(guideDir(), `${name}.md`), 'utf8');
}

function lastNodeId(ast: { sections: Array<{ nodes: Node[] }> }): string {
  const nodes = ast.sections[0].nodes;
  return nodes[nodes.length - 1].id;
}

function tableIds(ast: { sections: Array<{ nodes: Node[] }> }): string[] {
  return ast.sections[0].nodes.filter((n) => n.type === 'table').map((n) => n.id);
}

async function saveWithGuide(
  parser: DocxParser,
  guide: StyleGuideConfig,
  scratchDir: string,
  fileName: string,
): Promise<{ saveWarnings: unknown[]; validateWarnings: unknown[]; file: string }> {
  const file = path.join(scratchDir, fileName);
  fs.writeFileSync(file, parser.rezip());
  const manager = new DocumentManager();
  await manager.openDocument(file);
  manager.setStyleGuide(guide);
  const saved = await manager.saveDocument();
  const validated = await manager.validateCurrentDocument();
  return { saveWarnings: saved.styleWarnings, validateWarnings: validated.styleWarnings, file };
}

async function addChartAndCallouts(
  parser: DocxParser,
  guide: StyleGuideConfig,
  chartTitle: string,
  calloutTexts: string[],
): Promise<void> {
  let ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser, { styleGuide: guide });
  await engine.execute({
    action: 'chart',
    operation: 'insert',
    chartType: 'column',
    title: chartTitle,
    categories: ['Batch 1', 'Batch 2', 'Batch 3'],
    series: [
      { name: 'Planned', values: [30, 34, 32] },
      { name: 'Done', values: [28, 35, 33] },
      { name: 'Stretch', values: [8, 6, 9] },
    ],
    anchor: lastNodeId(ast),
  } as never);
  ast = await parser.getAST();
  const guided = new OperationEngine(ast, parser, { styleGuide: guide });
  const result = await guided.executeBatch({
    operations: calloutTexts.map((text) => ({
      action: 'callout',
      anchor: lastNodeId(ast),
      text,
    })) as never,
  });
  expect(result.success).toBe(true);
}

async function shadeFirstTableWithGuide(
  parser: DocxParser,
  guide: StyleGuideConfig,
): Promise<void> {
  const ast = await parser.getAST();
  const ids = tableIds(ast);
  expect(ids.length).toBeGreaterThan(0);
  const engine = new OperationEngine(ast, parser, { styleGuide: guide });
  await engine.execute({
    action: 'table',
    target: ids[0],
    operation: 'setShading',
    headerFill: guide.tables.headerFill,
    fill: 'F7F9FC',
  } as never);
}

describe('bundled guides parse zero-warning', () => {
  it('every bundled guide parses with no diagnostics', () => {
    for (const name of GUIDE_NAMES) {
      const { config, warnings } = parseStyleGuide(readGuide(name));
      expect(warnings).toEqual([]);
      expect(config.version).toBe(1);
    }
  });

  it('report guide carries the executive palette and contrast-safe tables', () => {
    const { config } = parseStyleGuide(readGuide('report'));
    expect(config.name).toBe('Corporate');
    expect(config.palette.primary).toBe('1E3A8A');
    expect(config.palette.secondary).toBe('0284C7');
    expect(config.palette.accent).toBe('D97706');
    expect(config.palette.chartPalette.length).toBeGreaterThanOrEqual(4);
    expect(config.tables.headerFill).toBe('1E3A8A');
    expect(config.tables.headerText).toBe('FFFFFF');
    expect(config.tables.keepTogether).toBe(true);
    expect(config.typography.fontFamily).toBe('Calibri');
    expect(config.callouts.defaultTheme).toBe('slate');
  });

  it('technical guide carries the engineering palette and shaded code blocks', () => {
    const { config } = parseStyleGuide(readGuide('technical'));
    expect(config.name).toBe('Technical');
    expect(config.palette.primary).toBe('334155');
    expect(config.palette.secondary).toBe('4F46E5');
    expect(config.palette.accent).toBe('10B981');
    expect(config.palette.chartPalette.length).toBeGreaterThanOrEqual(4);
    expect(config.tables.headerFill).toBe('334155');
    expect(config.tables.headerText).toBe('FFFFFF');
    expect(config.codeBlocks.background).toBe('F8FAFC');
  });

  it('legal guide is grayscale with unsplittable rows', () => {
    const { config } = parseStyleGuide(readGuide('legal'));
    expect(config.name).toBe('Legal');
    expect(config.palette.primary).toBe('0F172A');
    expect(config.palette.secondary).toBe('475569');
    expect(config.tables.headerFill).toBe('0F172A');
    expect(config.tables.headerText).toBe('FFFFFF');
    expect(config.tables.keepTogether).toBe(true);
    expect(config.palette.chartPalette.length).toBeGreaterThanOrEqual(4);
  });
});

describe('named-guide resolution succeeds for bundled guides', () => {
  it('lists all three bundled names', () => {
    expect(listBundledGuides()).toEqual(expect.arrayContaining([...GUIDE_NAMES]));
  });

  it('resolves each name to its file with zero parse warnings', () => {
    for (const name of GUIDE_NAMES) {
      const resolved = resolveStyleGuide(name);
      expect(resolved.warnings).toEqual([]);
      expect(resolved.filePath.endsWith(`${name}.md`)).toBe(true);
      const direct = parseStyleGuide(readGuide(name));
      expect(resolved.config).toEqual(direct.config);
    }
  });

  it('still fails closed on unknown names', () => {
    try {
      resolveStyleGuide('definitely-not-a-guide');
      expect.unreachable();
    } catch (err) {
      expect((err as { code?: string }).code).toBe('E_VALIDATION');
      expect(String((err as Error).message)).toContain('report');
    }
  });
});

describe('re-score rebuilds carry zero style warnings', () => {
  let scratchDir: string;

  beforeEach(() => {
    const root = path.join(os.tmpdir(), 'opencode');
    if (!fs.existsSync(root)) fs.mkdirSync(root, { recursive: true });
    scratchDir = fs.mkdtempSync(path.join(root, 'astra-rescore-'));
  });

  afterEach(() => {
    if (process.env.ASTRA_KEEP_RESCORE) {
      process.stdout.write(`kept rescore docs in ${scratchDir}\n`);
      return;
    }
    if (scratchDir && fs.existsSync(scratchDir)) {
      fs.rmSync(scratchDir, { recursive: true, force: true });
    }
  });

  it('status report under the report guide saves warning-free', async () => {
    const guide = resolveStyleGuide('report').config;
    const parser = new DocxParser(createBlankDocxBuffer('Status Report'));
    const composed = await composeDocument(
      parser,
      [
        '# Q3 Status Report',
        '',
        'Delivery summary for the third quarter across all active streams.',
        '',
        '## Milestones',
        '',
        '| Milestone | Owner | Status | Due |',
        '| --- | --- | --- | --- |',
        '| Checkout rewrite | Priya | Done | Jul 12 |',
        '| Search rollout | Tomas | Done | Aug 02 |',
        '| Billing migration | Aisha | On track | Sep 20 |',
        '| Mobile parity | Leo | At risk | Oct 05 |',
        '',
        '## Velocity',
        '',
        'Throughput held steady while scope grew in the second half.',
        '',
        '## Risks',
        '',
        '- Billing migration depends on one external vendor.',
        '- Mobile parity needs two more device-lab slots.',
        '',
      ].join('\n'),
      { styleGuide: guide },
    );
    expect(composed.success).toBe(true);
    await shadeFirstTableWithGuide(parser, guide);
    await addChartAndCallouts(parser, guide, 'Delivery velocity', [
      'Risk: billing migration depends on one external vendor.',
      'Win: search rollout landed two weeks early.',
    ]);
    const { saveWarnings, validateWarnings } = await saveWithGuide(
      parser,
      guide,
      scratchDir,
      'status-report.docx',
    );
    expect(saveWarnings).toEqual([]);
    expect(validateWarnings).toEqual([]);
  });

  it('NDA under the legal guide saves warning-free', async () => {
    const guide = resolveStyleGuide('legal').config;
    const parser = new DocxParser(createBlankDocxBuffer('NDA'));
    const composed = await composeDocument(
      parser,
      [
        '# Mutual Non-Disclosure Agreement',
        '',
        'This agreement is entered into by **Acme Corp** and **Globex Inc**.',
        '',
        '## 1. Confidential Information',
        '',
        'Each party agrees to hold all **Confidential Information** in strict confidence.',
        '',
        '## 2. Term',
        '',
        'This agreement remains in effect for three years from the date below.',
        '',
        '## Signatures',
        '',
        '| Party | Signatory | Date |',
        '| --- | --- | --- |',
        '| Acme Corp | _________________ | __________ |',
        '| Globex Inc | _________________ | __________ |',
        '',
        'Document control id: NDA-2026-0142.',
        '',
      ].join('\n'),
      { styleGuide: guide },
    );
    expect(composed.success).toBe(true);
    await shadeFirstTableWithGuide(parser, guide);
    const { saveWarnings, validateWarnings } = await saveWithGuide(
      parser,
      guide,
      scratchDir,
      'nda.docx',
    );
    expect(saveWarnings).toEqual([]);
    expect(validateWarnings).toEqual([]);
  });

  it('API reference under the technical guide saves warning-free', async () => {
    const guide = resolveStyleGuide('technical').config;
    const parser = new DocxParser(createBlankDocxBuffer('API Reference'));
    const composed = await composeDocument(
      parser,
      [
        '# Widget API Reference',
        '',
        'Version 2 of the widget service API.[^1]',
        '',
        '## Authentication',
        '',
        'Pass a bearer token on every request.',
        '',
        '```http',
        'GET /v2/widgets HTTP/1.1',
        'Authorization: Bearer <token>',
        '```',
        '',
        '## List widgets',
        '',
        '```http',
        'GET /v2/widgets?limit=50 HTTP/1.1',
        'Authorization: Bearer <token>',
        '```',
        '',
        '| Parameter | Type | Required |',
        '| --- | --- | --- |',
        '| limit | integer | no |',
        '| cursor | string | no |',
        '',
        '## Create a widget',
        '',
        '```json',
        '{"name": "gadget", "size": 3}',
        '```',
        '',
        '| Field | Type | Notes |',
        '| --- | --- | --- |',
        '| name | string | unique per account |',
        '| size | integer | between 1 and 9 |',
        '',
        '[^1]: Rate limits are published separately.',
        '',
      ].join('\n'),
      { styleGuide: guide },
    );
    expect(composed.success).toBe(true);
    const { saveWarnings, validateWarnings } = await saveWithGuide(
      parser,
      guide,
      scratchDir,
      'api-reference.docx',
    );
    expect(saveWarnings).toEqual([]);
    expect(validateWarnings).toEqual([]);
  });

  it('proposal under the legal guide saves warning-free', async () => {
    const guide = resolveStyleGuide('legal').config;
    const parser = new DocxParser(createBlankDocxBuffer('Proposal'));
    const composed = await composeDocument(
      parser,
      [
        '# Proposal: Regional Expansion',
        '',
        '## Executive Summary',
        '',
        'We propose opening two regional hubs to cut delivery latency.',
        '',
        '## Options',
        '',
        '| Option | Cost | Risk | Effort |',
        '| --- | --- | --- | --- |',
        '| North hub only | 420k | medium | 6 weeks |',
        '| South hub only | 380k | medium | 5 weeks |',
        '| Both hubs | 690k | low | 9 weeks |',
        '',
        '## Appendix: Assumptions',
        '',
        '- Lease terms stay within last quarter bands.',
        '- Hiring plan funds six field roles.',
        '',
      ].join('\n'),
      { styleGuide: guide },
    );
    expect(composed.success).toBe(true);
    await shadeFirstTableWithGuide(parser, guide);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser, { styleGuide: guide });
    const result = await engine.executeBatch({
      operations: [
        {
          action: 'callout',
          anchor: lastNodeId(ast),
          text: 'Recommendation: approve both hubs; the combined option has the lowest risk.',
        } as never,
      ],
    });
    expect(result.success).toBe(true);
    const { saveWarnings, validateWarnings } = await saveWithGuide(
      parser,
      guide,
      scratchDir,
      'proposal.docx',
    );
    expect(saveWarnings).toEqual([]);
    expect(validateWarnings).toEqual([]);
  });
});

describe('shipped dist copy stays loadable', () => {
  it('dist guides match source bytes', () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const srcDir = path.resolve(here, '..', 'templates', 'guides');
    const distDir = path.resolve(here, '..', 'dist', 'templates', 'guides');
    for (const name of GUIDE_NAMES) {
      const src = fs.readFileSync(path.join(srcDir, `${name}.md`), 'utf8');
      const shipped = fs.readFileSync(path.join(distDir, `${name}.md`), 'utf8');
      expect(shipped).toBe(src);
    }
  });

  it('candidates include a shipped dist path that exists', () => {
    for (const name of GUIDE_NAMES) {
      const cands = bundledGuideCandidates(name);
      const hit = cands.find((c) => c.includes(path.join('dist', 'templates', 'guides')));
      expect(hit).toBeDefined();
      expect(fs.existsSync(hit as string)).toBe(true);
    }
  });

  it('unknown names fail with the available set named', () => {
    try {
      resolveStyleGuide(' modern ');
      expect.unreachable();
    } catch (err) {
      expect((err as { code?: string }).code).toBe('E_VALIDATION');
      const msg = String((err as Error).message);
      expect(msg).toContain('report');
      expect(msg).not.toContain('no bundled guides installed');
    }
    try {
      resolveStyleGuide('definitely-not-a-guide');
      expect.unreachable();
    } catch (err) {
      expect((err as { code?: string }).code).toBe('E_VALIDATION');
    }
  });
});

describe('stored chart colors drive the duplicate rule', () => {
  async function twinDocXml(colors: [string, string]): Promise<{ ast: { sections: Array<{ nodes: Node[] }> } & Record<string, unknown>; getXml: (id: string) => string | undefined }> {
    const guide = resolveStyleGuide('report').config;
    const builder = new DocxParser(createBlankDocxBuffer('Twin Probe'));
    const live = await builder.getAST();
    const engine = new OperationEngine(live, builder, { styleGuide: guide });
    const anchor = live.sections[0].nodes[live.sections[0].nodes.length - 1].id;
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Twins',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [1, 2], color: colors[0] },
        { name: 'S2', values: [3, 4], color: colors[1] },
      ],
      anchor,
    } as never);
    const buf = builder.rezip();
    const fresh = new DocxParser(buf);
    const ast = (await fresh.getAST()) as unknown as { sections: Array<{ nodes: Node[] }> } & Record<string, unknown>;
    const zip = fresh.getZipAdapter();
    const getXml = chartXmlGetterForAst(
      ast as never,
      (part) => zip.readTextOrNull(part),
    );
    return { ast, getXml };
  }

  it('flags twin stored colors after reload', async () => {
    const guide = resolveStyleGuide('report').config;
    const { ast, getXml } = await twinDocXml(['FF0000', 'FF0000']);
    const findings = lintDocumentStyle(ast as never, guide, getXml);
    const dups = findings.filter((w) => w.rule === 'chart-duplicate-colors');
    expect(dups).toHaveLength(1);
    expect(dups[0].message).toContain('FF0000');
  });

  it('stays quiet for distinct stored colors after reload', async () => {
    const guide = resolveStyleGuide('report').config;
    const { ast, getXml } = await twinDocXml(['FF0000', '00FF00']);
    const findings = lintDocumentStyle(ast as never, guide, getXml);
    expect(findings.filter((w) => w.rule === 'chart-duplicate-colors')).toEqual([]);
  });

  it('live twin detail still fires without the stored reader', async () => {
    const guide = resolveStyleGuide('report').config;
    const builder = new DocxParser(createBlankDocxBuffer('Live Probe'));
    const ast = await builder.getAST();
    const engine = new OperationEngine(ast, builder, { styleGuide: guide });
    const anchor = ast.sections[0].nodes[ast.sections[0].nodes.length - 1].id;
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      title: 'Twins',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [1, 2], color: 'FF0000' },
        { name: 'S2', values: [3, 4], color: 'FF0000' },
      ],
      anchor,
    } as never);
    const findings = lintDocumentStyle(ast, guide);
    expect(findings.filter((w) => w.rule === 'chart-duplicate-colors')).toHaveLength(1);
    expect(lintDocumentStyle(ast, undefined)).toEqual([]);
  });
});
