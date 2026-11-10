import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, createBlankDocxBuffer, executeInsert, executeTable, contrastRatio } from '@attentionform/astra-engine';

async function freshTable() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  await executeInsert(
    { action: 'insert', type: 'table', anchor: ast.sections[0].nodes[0].id, position: 'after', headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']] } as any,
    ast,
    parser,
  );
  const tableId = ast.sections[0].nodes.find((n: any) => n.type === 'table').id as string;
  const xml = () => new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
  return { parser, ast, tableId, xml };
}
const headerRowColors = (xml: string): string[] => {
  const firstRow = xml.slice(xml.indexOf('<w:tr'), xml.indexOf('</w:tr>'));
  return [...firstRow.matchAll(/<w:color w:val="([0-9A-Fa-f]{6})"/g)].map((m) => m[1].toUpperCase());
};

describe('a dark fill gets readable text without the caller asking', () => {
  it('setShading headerFill 0F172A paints the header text white and says so', async () => {
    const { ast, tableId, xml } = await freshTable();
    const res: any = await executeTable({ action: 'table', target: tableId, operation: 'setShading', headerFill: '0F172A', fill: 'F7F9FC' } as any, ast);
    expect(headerRowColors(xml())).toEqual(['FFFFFF', 'FFFFFF']);
    expect(res.warnings.some((w: string) => /^Table note: text set to #FFFFFF .*0F172A/.test(w))).toBe(true);
    expect(res.warnings.some((w: string) => /contrast warning/.test(w))).toBe(false);
  });

  it('setCellShading on one dark cell paints that cell white', async () => {
    const { ast, tableId, xml } = await freshTable();
    const res: any = await executeTable({ action: 'table', target: tableId, operation: 'setCellShading', row: 2, col: 1, fill: '000080' } as any, ast);
    expect(xml()).toContain('w:fill="000080"');
    expect(xml()).toContain('<w:color w:val="FFFFFF"/>');
    expect(res.warnings.length).toBe(1);
    expect(res.warnings[0]).toMatch(/^Table note: text set to #FFFFFF/);
  });

  it('a light fill leaves the text alone and adds nothing', async () => {
    const { ast, tableId, xml } = await freshTable();
    const res: any = await executeTable({ action: 'table', target: tableId, operation: 'setShading', headerFill: 'D9D9D9', fill: 'F7F9FC' } as any, ast);
    expect(headerRowColors(xml())).toEqual([]);
    expect(res.warnings).toEqual([]);
  });

  it('an explicit textColor always wins, and the contrast warning stays for it', async () => {
    const { ast, tableId, xml } = await freshTable();
    const res: any = await executeTable({ action: 'table', target: tableId, operation: 'setShading', headerFill: '0F172A', fill: 'F7F9FC', textColor: '000000' } as any, ast);
    expect(headerRowColors(xml())).toEqual(['000000', '000000']);
    expect(res.warnings.some((w: string) => /contrast warning/.test(w))).toBe(true);
    expect(res.warnings.some((w: string) => /Table note/.test(w))).toBe(false);
  });

  it('a mid-tone fill gets whichever of white and black reads better', async () => {
    for (const fill of ['777777', '3B82F6', 'F59E0B', '10B981']) {
      const { ast, tableId, xml } = await freshTable();
      await executeTable({ action: 'table', target: tableId, operation: 'setCellShading', row: 1, col: 1, fill } as any, ast);
      const chosen = headerRowColors(xml())[0] ?? '000000';
      const other = chosen === 'FFFFFF' ? '000000' : 'FFFFFF';
      expect(contrastRatio(chosen, fill), `${fill} -> ${chosen}`).toBeGreaterThanOrEqual(contrastRatio(other, fill));
    }
  });

  it('a cell whose own text already reads well on the fill is not repainted', async () => {
    const { ast, tableId, xml } = await freshTable();
    await executeTable({ action: 'table', target: tableId, operation: 'setCellTextColor', row: 1, col: 1, color: 'FFFF00' } as any, ast);
    const res: any = await executeTable({ action: 'table', target: tableId, operation: 'setCellShading', row: 1, col: 1, fill: '000080' } as any, ast);
    expect(headerRowColors(xml())[0]).toBe('FFFF00');
    expect(res.warnings).toEqual([]);
  });
});
