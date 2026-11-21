import { describe, expect, it } from 'vitest';
import zlib from 'node:zlib';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, composeDocument } from '../src/internal';
import { contentBoxOf, textWidthTwipsOf, usableAreaOf } from '../src/operations/chart-size';
import { fitImageEmus } from '../src/compose/assets';

const EMU_IN = 914400;

function png(width: number, height: number): Buffer {
  const chunk = (type: string, data: Buffer): Buffer => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(body) >>> 0);
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(Buffer.alloc(0))), chunk('IEND', Buffer.alloc(0))]);
}

async function docOn(paper: string, margins?: unknown): Promise<{ parser: DocxParser; engine: OperationEngine }> {
  const parser = DocxParser.createNewDocument('blank');
  const engine = new OperationEngine(await parser.getAST(), parser);
  const res = await engine.executeBatch({ operations: [{ action: 'pageSetup', paperSize: paper, ...(margins ? { margins } : {}) }] as never });
  expect(res.success, JSON.stringify(res.errors)).toBe(true);
  return { parser, engine };
}
const docXml = (p: DocxParser): string => new AdmZip(p.rezip()).readAsText('word/document.xml');

describe('the text area that content has to fit', () => {
  it('the gutter comes off the width on the left and off the height on the top', () => {
    const page = { pageWidthTwips: 12240, pageHeightTwips: 15840, marginTop: 1440, marginBottom: 1440, marginLeft: 1440, marginRight: 1440, gutter: 720 };
    expect(usableAreaOf(page)).toEqual({ widthEmu: 6 * EMU_IN, heightEmu: 9 * EMU_IN });
    expect(usableAreaOf({ ...page, gutterAtTop: true })).toEqual({ widthEmu: 6.5 * EMU_IN, heightEmu: 8.5 * EMU_IN });
  });
  it('columns share the width, less a 0.5 in gap', () => {
    const page = { pageWidthTwips: 12240, pageHeightTwips: 15840, marginTop: 1440, marginBottom: 1440, marginLeft: 1440, marginRight: 1440, columns: 2 };
    expect(contentBoxOf(page).widthEmu).toBe(Math.floor((6.5 * EMU_IN - 0.5 * EMU_IN) / 2));
    expect(textWidthTwipsOf(page)).toBe(4320);
  });
  it('a section without a page size gets the Letter box and 9360 twips', () => {
    expect(contentBoxOf(null)).toEqual({ widthEmu: 6.5 * EMU_IN, heightEmu: 9 * EMU_IN });
    expect(textWidthTwipsOf(undefined)).toBe(9360);
  });
  it('the image cap is the old 6 in by 9 in on Letter and follows a smaller page', () => {
    const big = { widthPx: 4000, heightPx: 4000 };
    expect(fitImageEmus(big)).toEqual({ width: 6 * EMU_IN, height: 6 * EMU_IN });
    expect(fitImageEmus(big, contentBoxOf({ pageWidthTwips: 12240, pageHeightTwips: 15840, marginTop: 1440, marginBottom: 1440, marginLeft: 1440, marginRight: 1440 }))).toEqual({ width: 6 * EMU_IN, height: 6 * EMU_IN });
    const a5 = contentBoxOf({ pageWidthTwips: 8391, pageHeightTwips: 11906, marginTop: 1134, marginBottom: 1134, marginLeft: 1134, marginRight: 1134 });
    const fitted = fitImageEmus(big, a5)!;
    expect(fitted.width).toBeLessThan(a5.widthEmu);
    expect(fitted.width).toBeLessThanOrEqual(a5.widthEmu - 100000);
    expect(fitted.width).toBe(fitted.height);
  });
});

describe('what the engine writes follows the page', () => {
  it('a composed image on A5 fits the A5 text width, and on Letter keeps the old size', async () => {
    const data = `data:image/png;base64,${png(3000, 2000).toString('base64')}`;
    const widthOf = async (parser: DocxParser): Promise<number> => Number(/<wp:extent cx="(\d+)"/.exec(docXml(parser))?.[1]);
    const a5 = await docOn('A5', '20mm');
    await composeDocument(a5.parser, `![chart](${data})\n`);
    const a5Width = await widthOf(a5.parser);
    const a5Text = (8391 - 2 * 1134) * 635;
    expect(a5Width).toBeGreaterThan(0);
    expect(a5Width).toBeLessThanOrEqual(a5Text);
    const letter = DocxParser.createNewDocument('blank');
    await composeDocument(letter, `![chart](${data})\n`);
    expect(await widthOf(letter)).toBe(6 * EMU_IN);
  });
  it('an inserted table grid fills the text width of its page', async () => {
    const { parser, engine } = await docOn('A5', '20mm');
    const ast = await parser.getAST();
    const res = await engine.executeBatch({ operations: [{ action: 'insert', type: 'table', headers: ['A', 'B', 'C'], rows: [['1', '2', '3']], anchor: ast.sections[0].nodes[0].id, position: 'after' }] as never });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const widths = [...docXml(parser).matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => Number(m[1]));
    expect(widths).toHaveLength(3);
    const text = 8391 - 2 * 1134;
    expect(Math.abs(widths.reduce((a, b) => a + b, 0) - text)).toBeLessThanOrEqual(3);
  });
  it('a table of contents puts its tab stop at the text width, with a gutter taken off', async () => {
    const { parser, engine } = await docOn('A5', '20mm');
    const ast = await parser.getAST();
    await composeDocument(parser, '# One\n\ntext\n\n## Two\n\nmore\n');
    const first = (await parser.getAST()).sections[0].nodes[0].id;
    const res = await engine.executeBatch({ operations: [{ action: 'field', operation: 'insertTOC', levels: 3, position: 'before', anchor: first }] as never });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(docXml(parser)).toContain(`w:leader="dot" w:pos="${8391 - 2 * 1134}"`);
    void ast;
  });
});
