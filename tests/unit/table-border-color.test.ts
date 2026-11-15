/**
 * 0.5.3: table borders are not `auto`.
 *
 * Word draws a border with w:color="auto" black; LibreOffice draws it light grey. A table inserted by ASTra now
 * gets the style guide's `tables.borderColor` when it is a usable hex colour, otherwise a light grey, and an
 * explicit `setBorder` after creation still wins.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, parseStyleGuide, validateDocxBuffer } from '@attentionform/astra-engine';
import { composeDocument } from '../../packages/engine/src/compose/markdown';
import { DEFAULT_TABLE_BORDER_COLOR, resolveTableBorderColor } from '../../packages/engine/src/operations/insert';

async function session(guide?: unknown) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser, guide ? { styleGuide: guide as any } : undefined);
  const xml = () => new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
  const run = (operations: unknown[]) => engine.executeBatch({ operations } as any);
  const anchor = ast.sections[0].nodes[0].id as string;
  return { run, xml, anchor, parser, ast };
}

const insertTable = (anchor: string) => ({
  action: 'insert', type: 'table', id: 'tbl_a', anchor, position: 'after', headers: ['Item', 'Qty'], rows: [['a', '1'], ['b', '2']],
});

function tableBorders(xml: string): Array<{ edge: string; color: string; sz: string }> {
  const block = xml.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/)![0];
  return [...block.matchAll(/<w:(\w+) [^>]*w:sz="(\d+)"[^>]*w:color="([^"]+)"/g)].map((m) => ({ edge: m[1], sz: m[2], color: m[3] }));
}

describe('table border colour', () => {
  it('uses a light grey, not auto, without a style guide', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const borders = tableBorders(s.xml());
    expect(borders.map((b) => b.edge)).toEqual(['top', 'left', 'bottom', 'right', 'insideH', 'insideV']);
    for (const b of borders) {
      expect(b.color).toBe(DEFAULT_TABLE_BORDER_COLOR);
      expect(b.sz).toBe('4');
    }
    expect(s.xml()).not.toMatch(/<w:tblBorders>[\s\S]*w:color="auto"[\s\S]*<\/w:tblBorders>/);
  });

  it('follows the guide tables.borderColor', async () => {
    const guide = parseStyleGuide('# Style Guide: t\n\n## Tables\n- Border Color: #1E3A8A\n').config;
    const s = await session(guide);
    await s.run([insertTable(s.anchor)]);
    for (const b of tableBorders(s.xml())) expect(b.color).toBe('1E3A8A');
  });

  it('upper-cases a lower-case guide colour', async () => {
    const guide = parseStyleGuide('# Style Guide: t\n\n## Tables\n- Border Color: #cbd5e1\n').config;
    const s = await session(guide);
    await s.run([insertTable(s.anchor)]);
    for (const b of tableBorders(s.xml())) expect(b.color).toBe('CBD5E1');
  });

  it('falls back to the default for values that are not 6-digit hex', () => {
    for (const bad of ['red', '#12', '12345G', '', 'auto', undefined, 42]) {
      expect(resolveTableBorderColor({ tables: { borderColor: bad } } as never)).toBe(DEFAULT_TABLE_BORDER_COLOR);
    }
    expect(resolveTableBorderColor(undefined)).toBe(DEFAULT_TABLE_BORDER_COLOR);
    expect(resolveTableBorderColor({} as never)).toBe(DEFAULT_TABLE_BORDER_COLOR);
  });

  it('an invalid colour in the guide file warns and the table still gets a grey border', async () => {
    const parsed = parseStyleGuide('# Style Guide: t\n\n## Tables\n- Border Color: not-a-colour\n');
    expect(parsed.warnings.length).toBeGreaterThan(0);
    const s = await session(parsed.config);
    await s.run([insertTable(s.anchor)]);
    for (const b of tableBorders(s.xml())) expect(b.color).toMatch(/^[0-9A-F]{6}$/);
  });

  it('an explicit setBorder after creation wins', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    await s.run([{ action: 'table', target: 'tbl_a', operation: 'setBorder', border: 'bottom', color: 'FF0000', sz: 8 }]);
    const borders = tableBorders(s.xml());
    expect(borders.find((b) => b.edge === 'bottom')).toMatchObject({ color: 'FF0000', sz: '8' });
    expect(borders.find((b) => b.edge === 'top')?.color).toBe(DEFAULT_TABLE_BORDER_COLOR);
  });

  it('a Markdown table with the guide colour carries it, and the file stays valid', async () => {
    const guide = parseStyleGuide('# Style Guide: t\n\n## Tables\n- Border Color: #CBD5E1\n').config;
    const parser = DocxParser.createNewDocument('blank');
    await composeDocument(parser, '# T\n\n| A | B |\n|---|---|\n| 1 | 2 |\n', { styleGuide: guide });
    const xml = new AdmZip(parser.rezip()).readAsText('word/document.xml');
    for (const b of tableBorders(xml)) expect(b.color).toBe('CBD5E1');
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.errors).toEqual([]);
  });

  it('a Markdown table without a guide uses the built-in guide colour, never auto', async () => {
    const parser = DocxParser.createNewDocument('blank');
    await composeDocument(parser, '# T\n\n| A | B |\n|---|---|\n| 1 | 2 |\n');
    const xml = new AdmZip(parser.rezip()).readAsText('word/document.xml');
    for (const b of tableBorders(xml)) expect(b.color).not.toBe('auto');
  });
});
