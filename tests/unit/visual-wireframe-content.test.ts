/**
 * The wireframe shows what is on the page: table columns and the first rows,
 * list markers, and a one-cell shaded table as a callout. It reads the table
 * grid from the XML, as `inspect` does, because the model carries no table size.
 */
import { describe, it, expect } from 'vitest';
import { DocxParser, composeDocument, buildVisualTelemetry } from '@attentionform/astra-engine';

async function wireframeOf(markdown: string): Promise<{ wireframe: string; warnings: string[]; tokens: number }> {
  const parser = DocxParser.createNewDocument('blank');
  const result = await composeDocument(parser, markdown);
  expect(result.success).toBe(true);
  const telemetry = buildVisualTelemetry(await parser.getAST(), { page: 1 });
  return { wireframe: telemetry.wireframe, warnings: telemetry.warnings, tokens: telemetry.tokens };
}

const body = (frame: string): string[] => frame.split('\n').map((l) => l.replace(/^\| ?/, '').replace(/ *\|$/, ''));

describe('wireframe content', () => {
  it('prints the column count, the header row and the first rows of a table', async () => {
    const { wireframe } = await wireframeOf(
      '# T\n\n| Region | Revenue | Growth |\n| --- | --- | --- |\n| North | 120 | 4% |\n| South | 98 | 2% |\n| East | 143 | 7% |\n| West | 87 | -1% |\n',
    );
    const lines = body(wireframe);
    expect(lines).toContain('[TABLE: 3 cols]');
    expect(lines).toContain('  Region | Revenue | Growth');
    expect(lines).toContain('  North | 120 | 4%');
    expect(lines).toContain('  South | 98 | 2%');
    expect(lines).toContain('  (+2 more rows)');
    expect(wireframe).not.toContain('[TABLE: ? cols]');
  });

  it('marks bullets and numbers each numbered list from 1', async () => {
    const { wireframe } = await wireframeOf('- alpha\n- beta\n\n1. one\n2. two\n3. three\n');
    const lines = body(wireframe);
    expect(lines).toContain('- alpha');
    expect(lines).toContain('- beta');
    expect(lines).toContain('1. one');
    expect(lines).toContain('3. three');
  });

  it('labels a one-cell shaded table as a callout, not a table', async () => {
    const { wireframe } = await wireframeOf('Intro.\n\n> [!TIP]\n> Remember the deadline.\n');
    const lines = body(wireframe);
    expect(lines).toContain('[CALLOUT] Remember the deadline.');
    expect(wireframe).not.toContain('[TABLE');
  });

  it('stays inside the token budget for a dense page', async () => {
    const rows = Array.from({ length: 30 }, (_, i) => `| r${i} | ${i} | ${i * 2} |`).join('\n');
    const { tokens } = await wireframeOf(`# T\n\n- a\n- b\n\n| A | B | C |\n| --- | --- | --- |\n${rows}\n\n- c\n- d\n`);
    expect(tokens).toBeLessThanOrEqual(500);
  });

  it('warns about a table too wide for the page, now that columns are known', async () => {
    const header = '| ' + Array.from({ length: 8 }, (_, i) => `H${i}`).join(' | ') + ' |';
    const rule = '| ' + Array.from({ length: 8 }, () => '---').join(' | ') + ' |';
    const row = '| ' + Array.from({ length: 8 }, (_, i) => `${i}`).join(' | ') + ' |';
    const { warnings } = await wireframeOf(`${header}\n${rule}\n${row}\n`);
    expect(warnings.join('\n')).toMatch(/8 col|wide|margin/i);
  });

  it('does not read "paragraph" as a chart mention, but still flags "Figure 1"', async () => {
    const plain = await wireframeOf('A closing paragraph with nothing else.\n');
    expect(plain.warnings.join('\n')).not.toMatch(/chart/i);
    const figure = await wireframeOf('See Figure 1 for the trend.\n');
    expect(figure.warnings.join('\n')).toMatch(/chart/i);
  });
});
