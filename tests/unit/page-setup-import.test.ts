/**
 * Width arithmetic (percentage table widths, chart width warnings) must use the section's real page: its text
 * columns, and its actual page size, for an imported file as well as for one set through pageSetup.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { usableAreaOf } from '../../packages/engine/src/operations/chart-size';

const sectPrOf = (pgSz: string, extra = ''): string =>
  `<w:sectPr>${pgSz}<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${extra}</w:sectPr>`;

async function importWith(sectPr: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const doc = zip.readAsText('word/document.xml').replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/, sectPr);
  zip.updateFile('word/document.xml', Buffer.from(doc));
  const parser = new DocxParser(zip.toBuffer());
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const run = (operations: unknown[]) => engine.executeBatch({ operations } as any);
  const anchor = ast.sections[0].nodes[0].id as string;
  await run([{ action: 'insert', type: 'table', id: 't', anchor, position: 'after', headers: ['A', 'B'], rows: [['1', '2']] }]);
  const gridSum = (): number => {
    const xml = new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
    return [...xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)![0].matchAll(/<w:gridCol w:w="(\d+)"/g)].reduce((n, m) => n + Number(m[1]), 0);
  };
  return { ast, run, gridSum };
}
const full = { action: 'table', target: 't', operation: 'setTableWidth', widthPercent: 100 };

describe('percentage table width follows the real page of an imported section', () => {
  it('Letter portrait, one column: 6.5 in (control)', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="12240" w:h="15840"/>'));
    expect((await s.run([full])).success).toBe(true);
    expect(s.gridSum()).toBe(9360);
  });
  it('several text columns: widthPercent is refused, as for a section set through pageSetup', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="12240" w:h="15840"/>', '<w:cols w:num="2" w:space="720"/>'));
    expect(s.ast.sections[0].pageSetup.columns).toBe(2);
    const res = await s.run([full]);
    expect(res.success).toBe(false);
    expect(JSON.stringify(res)).toMatch(/several text columns/);
  });
  it('columns given as a list of w:col are counted', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="12240" w:h="15840"/>', '<w:cols w:equalWidth="0"><w:col w:w="3000" w:space="720"/><w:col w:w="5000"/></w:cols>'));
    expect(s.ast.sections[0].pageSetup.columns).toBe(2);
  });
  it('Legal landscape: 12 in of text width, not the Letter width', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="20160" w:h="12240" w:orient="landscape"/>'));
    expect(s.ast.sections[0].pageSetup.paperSize).toBe('Legal');
    expect((await s.run([full])).success).toBe(true);
    expect(s.gridSum()).toBe(20160 - 2880);
  });
  it('A3 portrait', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="16838" w:h="23811"/>'));
    expect(s.ast.sections[0].pageSetup.paperSize).toBe('A3');
    expect((await s.run([full])).success).toBe(true);
    expect(s.gridSum()).toBe(16838 - 2880);
  });
  it('a custom page size is used as written', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="10000" w:h="15000"/>'));
    expect((await s.run([full])).success).toBe(true);
    expect(s.gridSum()).toBe(10000 - 2880);
  });
  it('pageSetup to Legal landscape after import: the new page is used, not the stale one', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="12240" w:h="15840"/>'));
    expect((await s.run([{ action: 'pageSetup', paperSize: 'Legal', orientation: 'landscape' }])).success).toBe(true);
    expect((await s.run([full])).success).toBe(true);
    expect(s.gridSum()).toBe(20160 - 2880);
  });
  it('pageSetup to one column after an imported two-column section lifts the refusal', async () => {
    const s = await importWith(sectPrOf('<w:pgSz w:w="12240" w:h="15840"/>', '<w:cols w:num="2" w:space="720"/>'));
    expect((await s.run([{ action: 'pageSetup', columns: 1 }])).success).toBe(true);
    expect((await s.run([full])).success).toBe(true);
    expect(s.gridSum()).toBe(9360);
  });
});

describe('usableAreaOf', () => {
  it('knows Legal and A3 by name when no size is given', () => {
    expect(usableAreaOf({ paperSize: 'Legal', marginLeft: 1440, marginRight: 1440 }).widthEmu).toBe(Math.round(((12240 - 2880) / 1440) * 914400));
    expect(usableAreaOf({ paperSize: 'A3', orientation: 'landscape', marginLeft: 1440, marginRight: 1440 }).widthEmu).toBe(Math.round(((23811 - 2880) / 1440) * 914400));
  });
  it('an actual size wins over the paper name and is not turned again', () => {
    const area = usableAreaOf({ paperSize: 'Letter', orientation: 'landscape', pageWidthTwips: 20160, pageHeightTwips: 12240 });
    expect(area.widthEmu).toBe(Math.round(((20160 - 2880) / 1440) * 914400));
    expect(area.heightEmu).toBe(Math.round(((12240 - 2880) / 1440) * 914400));
  });
});
