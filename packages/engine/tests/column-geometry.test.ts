/**
 * Content placed in a section with columns is sized to the real column, read from the file: the space between equal columns, the listed
 * widths of columns of different widths. Where the width of the column that content lands in cannot be known, the content is sized to the
 * narrowest column and the caller gets a note.
 */
import { describe, expect, it } from 'vitest';
import zlib from 'node:zlib';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, composeDocument, validateDocxBuffer } from '../src/internal';
import { columnLayoutOf, contentBoxOf, textWidthTwipsOf } from '../src/operations/chart-size';

const png = (w: number, h: number): Buffer => {
  const chunk = (type: string, data: Buffer): Buffer => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(body) >>> 0);
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(Buffer.alloc(0))), chunk('IEND', Buffer.alloc(0))]);
};
const IMAGE = `data:image/png;base64,${png(3000, 2000).toString('base64')}`;
const IN = 914400;
const TW = 635; // EMU per twip

/** A blank Letter document whose only section has the given `w:cols`. */
function withCols(cols: string): DocxParser {
  const zip = new AdmZip(Buffer.from(DocxParser.createNewDocument('blank').rezip()));
  const xml = zip.readAsText('word/document.xml');
  const sect = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(xml)![0];
  const next = /<w:cols\b[^>]*\/>|<w:cols\b[\s\S]*?<\/w:cols>/.test(sect)
    ? sect.replace(/<w:cols\b[^>]*\/>|<w:cols\b[\s\S]*?<\/w:cols>/, cols)
    : sect.replace(/(<w:pgMar\b[^>]*\/>)/, `$1${cols}`);
  zip.updateFile('word/document.xml', Buffer.from(xml.replace(sect, next)));
  return new DocxParser(zip.toBuffer());
}
const run = async (parser: DocxParser, operations: object[]) => {
  const engine = new OperationEngine(await parser.getAST(), parser);
  return engine.executeBatch({ operations: operations as never });
};
const doc = (parser: DocxParser): string => new AdmZip(parser.rezip()).readAsText('word/document.xml');
const extents = (parser: DocxParser): number[] => [...doc(parser).matchAll(/<wp:extent cx="(\d+)"/g)].map((m) => Number(m[1]));
const chart = (extra: object = {}) => ({ action: 'chart', operation: 'insert', chartType: 'column', title: 'T', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }], anchor: 'end', position: 'after', ...extra });

const EQUAL_2IN_GAP = '<w:cols w:num="2" w:space="2880"/>';
const UNEQUAL = '<w:cols w:num="2" w:equalWidth="0"><w:col w:w="2000" w:space="1360"/><w:col w:w="6000"/></w:cols>';

describe('the column that content has to fit', () => {
  it('reads the gap of equal columns: 2 in gap on Letter leaves 2.25 in a column', async () => {
    const parser = withCols(EQUAL_2IN_GAP);
    const page = (await parser.getAST()).sections[0].pageSetup!;
    expect(page.columns).toBe(2);
    expect(page.columnSpace).toBe(2880);
    expect(contentBoxOf(page).widthEmu).toBe(Math.floor((9360 - 2880) / 2 * TW));
    expect(textWidthTwipsOf(page)).toBe(3240);
  });
  it('uses the 0.5 in Word default when the file gives no gap', async () => {
    const page = (await withCols('<w:cols w:num="2"/>').getAST()).sections[0].pageSetup!;
    expect(page.columnSpace).toBeUndefined();
    expect(textWidthTwipsOf(page)).toBe((9360 - 720) / 2);
  });
  it('reads listed widths of unequal columns and takes the narrowest', async () => {
    const page = (await withCols(UNEQUAL).getAST()).sections[0].pageSetup!;
    expect(page.columnWidths).toEqual([2000, 6000]);
    expect(columnLayoutOf(page)).toMatchObject({ unequal: true, unknown: false, widthEmu: 2000 * TW });
  });
  it('listed columns that all have one width are equal, not flagged', async () => {
    const page = (await withCols('<w:cols w:num="2" w:equalWidth="0"><w:col w:w="4320" w:space="720"/><w:col w:w="4320"/></w:cols>').getAST()).sections[0].pageSetup!;
    expect(columnLayoutOf(page)).toMatchObject({ unequal: false, unknown: false, widthEmu: 4320 * TW });
  });
  it('equalWidth="1" ignores stale listed columns', async () => {
    const page = (await withCols('<w:cols w:num="3" w:equalWidth="1" w:space="360"><w:col w:w="100"/><w:col w:w="200"/></w:cols>').getAST()).sections[0].pageSetup!;
    expect(page.columns).toBe(3);
    expect(page.columnWidths).toBeUndefined();
    expect(textWidthTwipsOf(page)).toBe(Math.round((9360 - 720) / 3));
  });
  it('listed columns with a missing width are flagged as not known', async () => {
    const page = (await withCols('<w:cols w:num="2" w:equalWidth="0"><w:col w:w="2000" w:space="720"/><w:col/></w:cols>').getAST()).sections[0].pageSetup!;
    expect(page.columnsUnknown).toBe(true);
    expect(columnLayoutOf(page)).toMatchObject({ unknown: true });
  });
});

describe('a chart', () => {
  it('without a size fits one of two equal columns (3 in on Letter), with no warning', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const res = await run(parser, [{ action: 'pageSetup', columns: 2 }, chart()]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(res.warnings ?? []).toEqual([]);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(3 * IN);
  });
  it('without a size follows a custom gap', async () => {
    const parser = withCols(EQUAL_2IN_GAP);
    const res = await run(parser, [chart()]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(3240 * TW + 1);
  });
  it('without a size in unequal columns takes the narrowest and says so', async () => {
    const parser = withCols(UNEQUAL);
    const res = await run(parser, [chart()]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(2000 * TW + 1);
    expect((res.warnings ?? []).join(' ')).toMatch(/narrowest column/);
  });
  it('with a width wider than the column is kept and warns against the column width', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const res = await run(parser, [{ action: 'pageSetup', columns: 2 }, chart({ widthIn: 6 })]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect((res.warnings ?? []).join(' ')).toMatch(/column width is 3 in/);
    expect(Math.max(...extents(parser))).toBe(6 * IN);
  });
});

describe('an image', () => {
  it('composed into two equal columns with a 2 in gap is 2.25 in wide at most', async () => {
    const parser = withCols(EQUAL_2IN_GAP);
    const res = await composeDocument(parser, `![figure](${IMAGE})\n`);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(2.25 * IN);
  });
  it('inserted by an operation without a size in unequal columns takes the narrowest and says so', async () => {
    const parser = withCols(UNEQUAL);
    const res = await run(parser, [{ action: 'image', operation: 'insert', fileData: IMAGE, altText: 'x', anchor: 'end', position: 'after' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(2000 * TW);
    expect((res.warnings ?? []).join(' ')).toMatch(/narrowest column/);
  });
  it('composed into unequal columns takes the narrowest and says so', async () => {
    const parser = withCols(UNEQUAL);
    const res = await composeDocument(parser, `![figure](${IMAGE})\n`);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(2000 * TW);
    expect(res.warnings.join(' ')).toMatch(/narrowest column/);
  });
  it('composed into equal columns says nothing about the columns', async () => {
    const parser = withCols(EQUAL_2IN_GAP);
    const res = await composeDocument(parser, `![figure](${IMAGE})\n`);
    expect(res.warnings.join(' ')).not.toMatch(/column/);
  });
});

describe('a table', () => {
  it('inserted into unequal columns has a grid within the narrowest column', async () => {
    const parser = withCols(UNEQUAL);
    const res = await run(parser, [{ action: 'insert', type: 'table', headers: ['A', 'B'], rows: [['1', '2']], anchor: 'end', position: 'after' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const grid = [...doc(parser).matchAll(/<w:gridCol w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0);
    expect(grid).toBeLessThanOrEqual(2000 + 2);
  });
  it('composed into two equal columns with a 2 in gap has a grid within one column', async () => {
    const parser = withCols(EQUAL_2IN_GAP);
    const res = await composeDocument(parser, '| A | B |\n|---|---|\n| 1 | 2 |\n');
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const grid = [...doc(parser).matchAll(/<w:gridCol w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0);
    expect(grid).toBeLessThanOrEqual(3240 + 2);
  });
});

describe('changing the column count', () => {
  it('replaces a list of unequal columns with equal ones and drops the stale geometry', async () => {
    const parser = withCols(UNEQUAL);
    const res = await run(parser, [{ action: 'pageSetup', columns: 3 }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const sect = [...doc(parser).matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]).pop()!;
    expect(sect).toMatch(/<w:cols\b[^>]*w:num="3"/);
    expect(sect).not.toMatch(/<w:col\b/);
    const page = (await parser.getAST()).sections[0].pageSetup!;
    expect(page.columns).toBe(3);
    expect(page.columnWidths).toBeUndefined();
    const valid = await validateDocxBuffer(parser.rezip());
    expect(valid.ok, JSON.stringify(valid.errors)).toBe(true);
  });
});

describe('section setColumns', () => {
  it('writes the gap it was given, drops a list of unequal columns, and content follows the new gap', async () => {
    const parser = withCols(UNEQUAL);
    const res = await run(parser, [{ action: 'section', operation: 'setColumns', sectionIndex: 1, columns: 2, spacing: 1440 }, chart()]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const sect = [...doc(parser).matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]).pop()!;
    expect(sect).not.toMatch(/<w:col\b/);
    const page = (await parser.getAST()).sections[0].pageSetup!;
    expect(page.columnSpace).toBe(1440);
    expect(page.columnWidths).toBeUndefined();
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(((9360 - 1440) / 2) * TW + 1);
    expect((res.warnings ?? []).join(' ')).not.toMatch(/narrowest/);
  });
});

/** A blank document whose only section has no `w:pgSz`: the page falls back to Letter, and the margins and columns in the file still apply. */
function withoutPageSize(margins: string, cols: string): DocxParser {
  const zip = new AdmZip(Buffer.from(DocxParser.createNewDocument('blank').rezip()));
  const xml = zip.readAsText('word/document.xml');
  const sect = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(xml)![0];
  const next = sect.replace(/<w:pgSz\b[^>]*\/>/, '').replace(/<w:pgMar\b[^>]*\/>/, '').replace(/<w:cols\b[^>]*\/>|<w:cols\b[\s\S]*?<\/w:cols>/, '').replace(/(<w:sectPr\b[^>]*>)/, `$1${margins}${cols}`);
  zip.updateFile('word/document.xml', Buffer.from(xml.replace(sect, next)));
  return new DocxParser(zip.toBuffer());
}
const MARGINS_2IN = '<w:pgMar w:top="1440" w:right="2880" w:bottom="1440" w:left="2880" w:header="720" w:footer="720" w:gutter="0"/>';

describe('a section without a page size', () => {
  it('has no w:pgSz in the test file, and reads margins and columns', async () => {
    const parser = withoutPageSize(MARGINS_2IN, '<w:cols w:num="2" w:space="720"/>');
    expect(doc(parser)).not.toMatch(/<w:pgSz\b/);
    const page = (await parser.getAST()).sections[0].pageSetup!;
    expect(page.pageWidthTwips).toBeUndefined();
    expect(page.marginLeft).toBe(2880);
    expect(page.columns).toBe(2);
  });
  it('text width is Letter less the file margins: 2 in each side leaves 4.5 in', async () => {
    const page = (await withoutPageSize(MARGINS_2IN, '').getAST()).sections[0].pageSetup!;
    expect(textWidthTwipsOf(page)).toBe(12240 - 2 * 2880);
  });
  it('text width is one column of what the margins leave', async () => {
    const page = (await withoutPageSize(MARGINS_2IN, '<w:cols w:num="2" w:space="720"/>').getAST()).sections[0].pageSetup!;
    expect(textWidthTwipsOf(page)).toBe((12240 - 2 * 2880 - 720) / 2);
  });
  it('columns alone (no margins, no size) still narrow the content', async () => {
    const page = (await withoutPageSize('', '<w:cols w:num="2" w:space="1440"/>').getAST()).sections[0].pageSetup!;
    expect(textWidthTwipsOf(page)).toBe((12240 - 2 * 1440 - 1440) / 2);
  });
  it('a chart, an image and a table fit the columns that the file gives', async () => {
    const column = (12240 - 2 * 2880 - 720) / 2; // 2520 twips
    const parser = withoutPageSize(MARGINS_2IN, '<w:cols w:num="2" w:space="720"/>');
    const res = await run(parser, [chart(), { action: 'image', operation: 'insert', fileData: IMAGE, altText: 'x', anchor: 'end', position: 'after' }, { action: 'insert', type: 'table', headers: ['A', 'B'], rows: [['1', '2']], anchor: 'end', position: 'after' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual(column * TW + 1);
    const grid = [...doc(parser).matchAll(/<w:gridCol w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0);
    expect(grid).toBeLessThanOrEqual(column + 2);
  });
  it('composed content fits the margins that the file gives', async () => {
    const parser = withoutPageSize(MARGINS_2IN, '');
    const res = await composeDocument(parser, `![figure](${IMAGE})\n\n| A | B |\n|---|---|\n| 1 | 2 |\n`);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(Math.max(...extents(parser))).toBeLessThanOrEqual((12240 - 2 * 2880) * TW + 1);
    const grid = [...doc(parser).matchAll(/<w:gridCol w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0);
    expect(grid).toBeLessThanOrEqual(12240 - 2 * 2880 + 2);
  });
  it('a section with nothing at all keeps the Letter text area', () => {
    expect(textWidthTwipsOf(undefined)).toBe(9360);
    expect(textWidthTwipsOf({})).toBe(9360);
    expect(contentBoxOf({}).widthEmu).toBe(6.5 * IN);
  });
  it('a named paper without a stored size uses the paper table (A5, not Letter)', () => {
    expect(textWidthTwipsOf({ paperSize: 'A5', marginLeft: 1134, marginRight: 1134 })).toBe(8391 - 2 * 1134);
  });
});

