/**
 * Per-tag callout theme pins: guide `Note/Tip/Important/Warning/Caution
 * Theme` lines parse into `callouts.tagThemes`, compose resolves
 * `tagThemes[tag] ?? defaultTheme ?? semantic`, and the summary carries an
 * additive `calloutTagThemes` copy. Guide-absent and tagless-guide runs pin
 * the legacy bytes (no behavior change).
 */
import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { parseStyleGuide } from '../src/style-guide/parser';
import { DEFAULT_STYLE_GUIDE } from '../src/style-guide/defaults';
import { summarizeStyleGuide } from '../src/style-guide/summary';
import { compileMarkdown } from '../src/compose/compile';
import { composeDocument } from '../src/compose/markdown';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import type { StyleGuideConfig } from '../src/style-guide/types';
import type { Node } from '../src/ast';

function guideDoc(calloutLines: string[]): string {
  return [
    '# Style Guide: Tag Themes',
    '',
    '## Palette',
    '- Primary: #1E3A8A',
    '- Chart Palette: #1E3A8A, #0284C7, #10B981, #F59E0B',
    '',
    '## Tables',
    '- Keep Together: true',
    '',
    '## Typography',
    '- Font Family: Calibri',
    '',
    '## Callouts',
    ...calloutLines,
    '',
    '## Code Blocks',
    '- Background: #F1F5F9',
    '',
    '## Layout',
    '- Margin: standard',
    '',
  ].join('\n');
}

function opThemes(markdown: string, styleGuide?: StyleGuideConfig): string[] {
  const ops = compileMarkdown(
    markdown,
    styleGuide !== undefined ? { styleGuide } : undefined,
  );
  return ops
    .filter((op) => (op as { action?: string }).action === 'callout')
    .map((op) => String((op as Record<string, unknown>)['theme']));
}

const ALL_ALERTS = [
  '> [!NOTE]',
  '> A note.',
  '',
  '> [!TIP]',
  '> A tip.',
  '',
  '> [!IMPORTANT]',
  '> Important.',
  '',
  '> [!WARNING]',
  '> A warning.',
  '',
  '> [!CAUTION]',
  '> Caution.',
  '',
].join('\n');

describe('per-tag theme parsing', () => {
  it('parses all five tag lines with case-insensitive keys and values', () => {
    const { config, warnings } = parseStyleGuide(
      guideDoc([
        '- Default Theme: slate',
        '- Note Theme: indigo',
        '- Tip Theme: EMERALD',
        '- Important Theme: Slate',
        '- Warning Theme: amber',
        '- Caution Theme: rose',
      ]),
    );
    expect(warnings).toEqual([]);
    expect(config.callouts.tagThemes).toEqual({
      note: 'indigo',
      tip: 'emerald',
      important: 'slate',
      warning: 'amber',
      caution: 'rose',
    });
  });

  it('warns and ignores an unknown tag key', () => {
    const { config, warnings } = parseStyleGuide(
      guideDoc(['- Default Theme: slate', '- Alert Theme: amber']),
    );
    expect(warnings.some((w) => /unrecognized key 'Alert Theme' in 'callouts' ignored/.test(w))).toBe(true);
    expect('tagThemes' in config.callouts).toBe(false);
  });

  it('warns on a bad tag value and leaves that tag to the default theme', () => {
    const { config, warnings } = parseStyleGuide(
      guideDoc([
        '- Default Theme: slate',
        '- Warning Theme: puce',
        '- Note Theme: indigo',
      ]),
    );
    expect(
      warnings.some((w) =>
        /unknown callout theme 'puce' for Warning Theme; default theme applied/.test(w),
      ),
    ).toBe(true);
    expect(config.callouts.tagThemes).toEqual({ note: 'indigo' });
    expect(config.callouts.tagThemes?.warning).toBeUndefined();
  });

  it('keeps the last line on repeats', () => {
    const { config, warnings } = parseStyleGuide(
      guideDoc([
        '- Default Theme: slate',
        '- Warning Theme: rose',
        '- Warning Theme: amber',
      ]),
    );
    expect(config.callouts.tagThemes).toEqual({ warning: 'amber' });
    expect(warnings.some((w) => /duplicate key.*Warning Theme/.test(w))).toBe(true);
  });

  it('carries no tagThemes without tag lines', () => {
    expect('tagThemes' in DEFAULT_STYLE_GUIDE.callouts).toBe(false);
    const sparse = parseStyleGuide(
      guideDoc(['- Default Theme: slate']),
    );
    expect('tagThemes' in sparse.config.callouts).toBe(false);
    expect(sparse.warnings).toEqual([]);
  });
});

describe('per-tag theme resolution', () => {
  it('uses the semantic map without a guide', () => {
    expect(opThemes(ALL_ALERTS)).toEqual([
      'indigo',
      'emerald',
      'slate',
      'amber',
      'rose',
    ]);
  });

  it('uses the guide default for every tag when no tag lines exist', () => {
    const { config } = parseStyleGuide(guideDoc(['- Default Theme: slate']));
    expect(opThemes(ALL_ALERTS, config)).toEqual([
      'slate',
      'slate',
      'slate',
      'slate',
      'slate',
    ]);
  });

  it('prefers the configured tag theme over the guide default', () => {
    const { config } = parseStyleGuide(
      guideDoc(['- Default Theme: slate', '- Warning Theme: amber']),
    );
    const ops = compileMarkdown('> [!WARNING]\n> Watch out.\n\n> [!NOTE]\n> Take note.\n', {
      styleGuide: config,
    });
    const themes = ops
      .filter((op) => (op as { action?: string }).action === 'callout')
      .map((op) => String((op as Record<string, unknown>)['theme']));
    expect(themes).toEqual(['amber', 'slate']);
  });

  it('resolves every tag independently under a fully configured guide', () => {
    const { config } = parseStyleGuide(
      guideDoc([
        '- Default Theme: slate',
        '- Note Theme: rose',
        '- Tip Theme: indigo',
        '- Important Theme: emerald',
        '- Warning Theme: amber',
        '- Caution Theme: slate',
      ]),
    );
    expect(opThemes(ALL_ALERTS, config)).toEqual([
      'rose',
      'indigo',
      'emerald',
      'amber',
      'slate',
    ]);
  });

  it('falls back to the guide default for a tag whose value was rejected', () => {
    const { config } = parseStyleGuide(
      guideDoc(['- Default Theme: slate', '- Warning Theme: puce']),
    );
    expect(opThemes('> [!WARNING]\n> Watch out.\n', config)).toEqual(['slate']);
  });

  it('lets an explicit batch callout theme beat guide tag themes', async () => {
    const { config } = parseStyleGuide(
      guideDoc(['- Default Theme: slate', '- Warning Theme: amber']),
    );
    const parser = new DocxParser(createBlankDocxBuffer('Explicit Wins'));
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[ast.sections[0].nodes.length - 1].id;
    const engine = new OperationEngine(ast, parser, { styleGuide: config });
    const result = await engine.executeBatch({
      operations: [
        { action: 'callout', anchor, text: 'Explicit rose', theme: 'rose' } as never,
      ],
    });
    expect(result.success).toBe(true);
    const node = ast.sections[0].nodes.find(
      (n) => (n.metadata as { callout?: { title?: string } }).callout !== undefined,
    ) as Node;
    const xml = new XMLSerializer().serializeToString(node.xml);
    expect(xml).toContain('w:color="E11D48"');
    expect(xml).not.toContain('w:color="D97706"');
  });
});

describe('per-tag themes end to end', () => {
  it('renders warning amber and note slate under a custom guide', async () => {
    const { config, warnings } = parseStyleGuide(
      guideDoc(['- Default Theme: slate', '- Warning Theme: amber']),
    );
    expect(warnings).toEqual([]);
    const parser = new DocxParser(createBlankDocxBuffer('Tag Theme Compose'));
    const result = await composeDocument(
      parser,
      '> [!WARNING]\n> Watch out here.\n\n> [!NOTE]\n> Take note here.\n',
      { styleGuide: config },
    );
    expect(result.success).toBe(true);
    const ast = await parser.getAST();
    const callouts = ast.sections[0].nodes.filter(
      (n) => (n.metadata as { callout?: { title?: string } }).callout !== undefined,
    );
    expect(callouts).toHaveLength(2);
    const xml = callouts.map((n) => new XMLSerializer().serializeToString(n.xml)).join('');
    expect(xml).toContain('w:color="D97706"');
    expect(xml).toContain('w:color="475569"');
    expect(xml).not.toContain('w:color="4F46E5"');
  });
});

describe('per-tag theme summary', () => {
  it('keeps the singular theme and copies tag themes when present', () => {
    const { config } = parseStyleGuide(
      guideDoc(['- Default Theme: slate', '- Warning Theme: amber']),
    );
    const summary = summarizeStyleGuide(config);
    expect(summary.calloutTheme).toBe('slate');
    expect(summary.calloutTagThemes).toEqual({ warning: 'amber' });
    summary.calloutTagThemes!.warning = 'rose';
    expect(config.callouts.tagThemes).toEqual({ warning: 'amber' });
  });

  it('omits tag themes when unconfigured', () => {
    const { config } = parseStyleGuide(guideDoc(['- Default Theme: slate']));
    const summary = summarizeStyleGuide(config);
    expect(summary.calloutTheme).toBe('slate');
    expect('calloutTagThemes' in summary).toBe(false);
  });
});
