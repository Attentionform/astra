/**
 * Style guide parser pins: fail-safe defaulting, hex normalization, and
 * the diagnostics channel. The parser never throws on user input.
 */
import { describe, it, expect } from 'vitest';
import { parseStyleGuide, bundledGuidePath, isFileGuideRef } from '../src/style-guide/parser';
import { DEFAULT_STYLE_GUIDE } from '../src/style-guide/defaults';

const CORPORATE_NAVY = `# Style Guide: Corporate Navy

Executive reporting theme emphasizing authority, clarity, and high contrast.

## Palette
- Primary: #1E3A8A
- Secondary: #0284C7
- Accent: #D97706
- Neutral Dark: #1E293B
- Neutral Light: #F8FAFC
- Chart Palette: #1E3A8A, #0284C7, #10B981, #F59E0B, #6366F1, #EC4899

## Typography
- Font Family: Calibri
- Code Font: Consolas
- Heading 1: #1E3A8A, bold
- Heading 2: #0284C7, bold

## Tables
- Header Fill: #1E3A8A
- Header Text: #FFFFFF
- Alternating Fill: #F8FAFC
- Border Color: #CBD5E1
- Keep Together: true

## Callouts
- Default Theme: slate
- Preferred Themes: slate, emerald, amber

## Code Blocks
- Background: #F1F5F9
- Border Color: #CBD5E1

## Layout
- Margin: standard
- Prevent Orphan Headings: true
- Widow Control: true
`;

describe('style guide parser', () => {
  it('parses the canonical corporate example with zero warnings', () => {
    const { config, warnings } = parseStyleGuide(CORPORATE_NAVY);
    expect(warnings).toEqual([]);
    expect(config.version).toBe(1);
    expect(config.name).toBe('Corporate Navy');
    expect(config.description).toContain('authority');
    expect(config.palette.primary).toBe('1E3A8A');
    expect(config.palette.chartPalette).toEqual([
      '1E3A8A', '0284C7', '10B981', 'F59E0B', '6366F1', 'EC4899',
    ]);
    expect(config.tables.headerText).toBe('FFFFFF');
    expect(config.tables.keepTogether).toBe(true);
    expect(config.typography.fontFamily).toBe('Calibri');
    expect(config.callouts.defaultTheme).toBe('slate');
    expect(config.codeBlocks.background).toBe('F1F5F9');
    expect(config.layout.margin).toBe('standard');
  });

  it('defaults every missing section with one warning per section', () => {
    const { config, warnings } = parseStyleGuide(
      '# Style Guide: Sparse\n\n## Palette\n- Primary: #112233\n',
    );
    expect(config.palette.primary).toBe('112233');
    for (const s of ['tables', 'typography', 'callouts', 'codeBlocks', 'layout']) {
      expect(warnings.some((w) => w.includes(`section '${s}' missing`))).toBe(true);
    }
    expect(config.tables).toEqual(DEFAULT_STYLE_GUIDE.tables);
    expect(config.layout).toEqual(DEFAULT_STYLE_GUIDE.layout);
    // palette section present -> no palette warning
    expect(warnings.some((w) => w.includes(`section 'palette'`))).toBe(false);
  });

  it('falls back per-key on malformed hex and names the line', () => {
    const md = [
      '# Style Guide: Broken',
      '',
      '## Palette',
      '- Primary: notacolor',
      '- Accent: #ZZZZZZ',
      '',
      '## Tables',
      '- Header Fill: #12345',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.primary).toBe(DEFAULT_STYLE_GUIDE.palette.primary);
    expect(config.palette.accent).toBe(DEFAULT_STYLE_GUIDE.palette.accent);
    expect(config.tables.headerFill).toBe(DEFAULT_STYLE_GUIDE.tables.headerFill);
    expect(warnings.some((w) => /line 4.*palette\.primary/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 5.*palette\.accent/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 8.*tables\.headerFill/.test(w))).toBe(true);
  });

  it('drops malformed chart entries, pads short palettes, keeps valid order', () => {
    const md = [
      '# Style Guide: Charts',
      '',
      '## Palette',
      '- Chart Palette: #1E3A8A, bogus, #10B981',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.chartPalette.slice(0, 2)).toEqual(['1E3A8A', '10B981']);
    expect(config.palette.chartPalette.length).toBeGreaterThanOrEqual(4);
    expect(warnings.some((w) => /line 4.*bogus/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 4.*padded/.test(w))).toBe(true);
  });

  it('ignores html comments and empty lines without warnings', () => {
    const md = [
      '<!-- house theme, do not edit lightly -->',
      '# Style Guide: Commented',
      '',
      '',
      '## Palette <!-- primary section -->',
      '<!-- - Primary: #000000 -->',
      '- Primary: #112233',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.primary).toBe('112233');
    expect(warnings).toEqual([]);
  });

  it('warns on unknown version but still parses best-effort', () => {
    const md = CORPORATE_NAVY.replace(
      'Executive reporting',
      '- Version: 99\n\nExecutive reporting',
    );
    const { config, warnings } = parseStyleGuide(md);
    expect(config.version).toBe(1);
    expect(config.palette.primary).toBe('1E3A8A');
    expect(config.tables.keepTogether).toBe(true);
    expect(warnings.some((w) => /unknown guide version '99'/.test(w))).toBe(true);
  });

  it('accepts explicit version 1 silently', () => {
    const md = `# Style Guide: V1\n\n- Version: 1\n\n${CORPORATE_NAVY.split('\n').slice(4).join('\n')}`;
    const { warnings } = parseStyleGuide(md);
    expect(warnings).toEqual([]);
  });

  it('normalizes lowercase and hashless hex silently; expands shorthand with a warning', () => {
    const md = [
      '# Style Guide: Case',
      '',
      '## Palette',
      '- Primary: #1e3a8a',
      '- Secondary: 0284c7',
      '- Accent: #D97',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.primary).toBe('1E3A8A');
    expect(config.palette.secondary).toBe('0284C7');
    expect(config.palette.accent).toBe('DD9977');
    expect(warnings.some((w) => /line 6.*expanded/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 4/.test(w))).toBe(false);
  });

  it('falls back on unknown theme, margin, and boolean with line warnings', () => {
    const md = [
      '# Style Guide: Unknowns',
      '',
      '## Palette',
      '- Primary: #112233',
      '',
      '## Tables',
      '- Keep Together: maybe',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: ultraviolet',
      '- Preferred Themes: slate, ultraviolet',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: enormous',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.callouts.defaultTheme).toBe(DEFAULT_STYLE_GUIDE.callouts.defaultTheme);
    expect(config.callouts.preferredThemes).toEqual(['slate']);
    expect(config.tables.keepTogether).toBe(DEFAULT_STYLE_GUIDE.tables.keepTogether);
    expect(config.layout.margin).toBe('standard');
    expect(warnings.some((w) => /line 13.*ultraviolet/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 14.*ultraviolet/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 7.*maybe/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 20.*enormous/.test(w))).toBe(true);
  });

  it('warns on unknown sections and unrecognized keys without throwing', () => {
    const md = [
      '# Style Guide: Odd',
      '',
      '## Palette',
      '- Primary: #112233',
      '- Favorite Snack: pretzels',
      '',
      '## Charts',
      '- Primary: #000000',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.primary).toBe('112233');
    expect(warnings.some((w) => /line 5.*Favorite Snack/.test(w))).toBe(true);
    expect(warnings.some((w) => /line 7.*unknown section/.test(w))).toBe(true);
  });

  it('never throws on empty or garbage input and returns full defaults', () => {
    for (const input of ['', '\n\n', '}}}{{[[', '- just\n- bullets\n- no sections']) {
      const { config, warnings } = parseStyleGuide(input);
      expect(config.version).toBe(1);
      expect(config.palette.chartPalette.length).toBeGreaterThanOrEqual(4);
      expect(warnings.length).toBeGreaterThan(0);
    }
    const empty = parseStyleGuide('');
    expect(empty.warnings.filter((w) => w.startsWith('section'))).toHaveLength(6);
  });

  it('section headers match case-insensitively and code-blocks spacing variants', () => {
    const md = [
      '# Style Guide: Variants',
      '',
      '## PALETTE',
      '- Primary: #112233',
      '',
      '## code blocks',
      '- Background: #AABBCC',
      '',
      '## tables',
      '- Keep Together: false',
      '',
      '## TYPOGRAPHY',
      '- Font Family: Arial',
      '',
      '## CallOuts',
      '- Default Theme: rose',
      '',
      '## LAYOUT',
      '- Margin: narrow',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(warnings).toEqual([]);
    expect(config.codeBlocks.background).toBe('AABBCC');
    expect(config.tables.keepTogether).toBe(false);
    expect(config.callouts.defaultTheme).toBe('rose');
    expect(config.layout.margin).toBe('narrow');
  });

  it('named guides resolve to the bundled path convention', () => {
    expect(bundledGuidePath('corporate')).toBe('packages/engine/templates/guides/corporate.md');
    expect(bundledGuidePath(' Report ')).toBe('packages/engine/templates/guides/report.md');
    expect(bundledGuidePath('legal.md')).toBe('packages/engine/templates/guides/legal.md');
  });

  it('distinguishes file refs from named guides', () => {
    expect(isFileGuideRef('./my-brand.md')).toBe(true);
    expect(isFileGuideRef('/workspace/style-guide.md')).toBe(true);
    expect(isFileGuideRef('corporate')).toBe(false);
    expect(isFileGuideRef('report')).toBe(false);
  });

  it('warns when chart palette shorthand expands', () => {
    const md = [
      '# Style Guide: Chart Short',
      '',
      '## Palette',
      '- Chart Palette: #D97, #1E3A8A, #0284C7, #10B981',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.chartPalette[0]).toBe('DD9977');
    expect(warnings.some((w) => /line 4.*shorthand.*DD9977/.test(w))).toBe(true);
  });

  it('parses heading 1/2 styles warn-free with color and bold', () => {
    const { config, warnings } = parseStyleGuide(CORPORATE_NAVY);
    expect(warnings).toEqual([]);
    expect(config.typography.heading1).toEqual({ color: '1E3A8A', bold: true });
    expect(config.typography.heading2).toEqual({ color: '0284C7', bold: true });
  });

  it('warns on duplicate keys naming the line and kept value', () => {
    const md = [
      '# Style Guide: Dup',
      '',
      '## Palette',
      '- Primary: #111111',
      '- Primary: #222222',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.primary).toBe('222222');
    expect(warnings.some((w) => /line 5.*duplicate key.*Primary.*222222/.test(w))).toBe(true);
  });

  it('pads short chart palettes without duplicating user colors', () => {
    const md = [
      '# Style Guide: Pad',
      '',
      '## Palette',
      '- Chart Palette: #1E3A8A, bogus',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.palette.chartPalette.length).toBeGreaterThanOrEqual(4);
    expect(new Set(config.palette.chartPalette).size).toBe(config.palette.chartPalette.length);
    expect(config.palette.chartPalette[0]).toBe('1E3A8A');
    expect(warnings.some((w) => /line 4.*padded/.test(w))).toBe(true);
  });

  it('warns on a later guide title and keeps the first', () => {
    const md = [
      '# Style Guide: First',
      '# Second Title',
      '',
      '## Palette',
      '- Primary: #112233',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(config.name).toBe('First');
    expect(warnings.some((w) => /line 2.*duplicate guide title.*Second Title.*kept.*First/.test(w))).toBe(true);
  });

  it('warns on a deep heading outside any section', () => {
    const md = [
      '# Style Guide: First',
      '',
      '### Stray Heading',
      '',
      '## Palette',
      '- Primary: #112233',
      '',
      '## Tables',
      '- Keep Together: true',
      '',
      '## Typography',
      '- Font Family: Calibri',
      '',
      '## Callouts',
      '- Default Theme: slate',
      '',
      '## Code Blocks',
      '- Background: #F1F5F9',
      '',
      '## Layout',
      '- Margin: standard',
      '',
    ].join('\n');
    const { config, warnings } = parseStyleGuide(md);
    expect(warnings.some((w) => /line 3.*heading outside any section.*Stray Heading/.test(w))).toBe(true);
    expect(config.description ?? '').not.toContain('Stray Heading');
  });

  it('keeps the built-in default guide frozen and parse copies independent', () => {
    expect(Object.isFrozen(DEFAULT_STYLE_GUIDE)).toBe(true);
    expect(Object.isFrozen(DEFAULT_STYLE_GUIDE.palette.chartPalette)).toBe(true);
    const { config } = parseStyleGuide(CORPORATE_NAVY);
    config.palette.primary = '000000';
    config.palette.chartPalette.push('000000');
    expect(DEFAULT_STYLE_GUIDE.palette.primary).toBe('1E3A8A');
    expect(DEFAULT_STYLE_GUIDE.palette.chartPalette).not.toContain('000000');
  });
});
