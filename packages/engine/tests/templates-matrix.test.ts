/**
 * Every designed template on ten pages (named papers, custom sizes, landscape): create, compose Markdown with a table, an image and a callout,
 * then insert a native chart without a size, then check the file is valid, the section size is what was asked for, and no table or tab stop reaches past the text area.
 */
import { describe, expect, it } from 'vitest';
import zlib from 'node:zlib';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, TEMPLATE_NAMES, composeDocument, validateDocxBuffer } from '../src/internal';

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
const MD = `# Matrix check\n\nIntro paragraph.\n\n| Name | Qty | Note |\n|---|---|---|\n| One | 1 | first |\n| Two | 2 | second |\n\n> [!NOTE]\n> A callout.\n\n![figure](${IMAGE})\n`;

const DESIGNED = TEMPLATE_NAMES.filter((n) => n !== 'blank');
const PAGES: Array<{ label: string; page: { size: string | { width: string; height: string }; orientation?: 'landscape' }; w: number; h: number }> = [
  { label: 'Letter', page: { size: 'Letter' }, w: 12240, h: 15840 },
  { label: 'A4', page: { size: 'A4' }, w: 11906, h: 16838 },
  { label: 'Legal', page: { size: 'Legal' }, w: 12240, h: 20160 },
  { label: 'A3', page: { size: 'A3' }, w: 16838, h: 23811 },
  { label: 'A5', page: { size: 'A5' }, w: 8391, h: 11906 },
  { label: 'B5', page: { size: 'B5' }, w: 9979, h: 14173 },
  { label: '6x9in', page: { size: '6x9in' }, w: 8640, h: 12960 },
  { label: '150x200mm', page: { size: '150x200mm' }, w: 8504, h: 11339 },
  { label: 'Letter landscape', page: { size: 'Letter', orientation: 'landscape' }, w: 15840, h: 12240 },
  { label: 'A4 landscape', page: { size: 'A4', orientation: 'landscape' }, w: 16838, h: 11906 },
];

describe.each(PAGES)('templates on $label', ({ page, w, h }) => {
  for (const name of DESIGNED) {
    it(`${name}: valid, sized as asked, nothing wider than the text`, async () => {
      const parser = DocxParser.createNewDocument(name, { page });
      const res = await composeDocument(parser, MD, { template: name } as never);
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      const ast = await parser.getAST();
      const last = ast.sections[ast.sections.length - 1].nodes;
      const chart = await new OperationEngine(ast, parser).executeBatch({
        operations: [{ action: 'chart', operation: 'insert', chartType: 'column', title: 'Matrix', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }], anchor: last[last.length - 1].id, position: 'after' }] as never,
      });
      expect(chart.success, JSON.stringify(chart.errors)).toBe(true);
      const out = parser.rezip();
      const valid = await validateDocxBuffer(out);
      expect(valid.ok, JSON.stringify(valid.errors)).toBe(true);
      const zip = new AdmZip(out);
      const doc = zip.readAsText('word/document.xml');
      // The body section: the last sectPr.
      const sect = [...doc.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]).pop() as string;
      const num = (re: RegExp): number => Number(re.exec(sect)?.[1]);
      expect(num(/<w:pgSz[^>]*w:w="(\d+)"/)).toBeGreaterThanOrEqual(Math.abs(w - 1));
      expect(Math.abs(num(/<w:pgSz[^>]*w:w="(\d+)"/) - w)).toBeLessThanOrEqual(1);
      expect(Math.abs(num(/<w:pgSz[^>]*w:h="(\d+)"/) - h)).toBeLessThanOrEqual(1);
      const text = num(/<w:pgSz[^>]*w:w="(\d+)"/) - num(/w:left="(\d+)"/) - num(/w:right="(\d+)"/) - num(/w:gutter="(\d+)"/);
      // A chart inserted without a size fits the text width (an image of the same kind is capped the same way).
      // With columns in the body section, one column is the bound.
      const colsTag = /<w:cols\b[^>]*>/.exec(sect)?.[0] ?? '';
      const colCount = Number(/w:num="(\d+)"/.exec(colsTag)?.[1] ?? 1);
      const colSpace = Number(/w:space="(\d+)"/.exec(colsTag)?.[1] ?? 720);
      const room = colCount > 1 ? (text - colSpace * (colCount - 1)) / colCount : text;
      for (const m of doc.matchAll(/<wp:extent cx="(\d+)"/g)) expect(Number(m[1]), `an image or chart ${m[1]} EMU wide in a ${room} twip column`).toBeLessThanOrEqual((room + 3) * 635);
      expect(doc).toContain('c:chart');
      // Every table's grid fits the text width of the widest section that holds it (columns only narrow it, so the page text width is the bound).
      for (const tbl of doc.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) {
        if (/<w:tbl>[\s\S]*<w:tbl>/.test(tbl[0].slice(7))) continue; // nested: the outer grid is checked
        const grid = [...tbl[0].matchAll(/<w:gridCol w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0);
        expect(grid, `a table grid of ${grid} twips on a ${text} twip text width`).toBeLessThanOrEqual(text + 3);
      }
      // Every tab stop in the document and in the styles stays inside the text area (right tabs may sit at the edge).
      for (const part of ['word/document.xml', 'word/styles.xml', 'word/header1.xml', 'word/footer1.xml', 'word/header2.xml', 'word/footer2.xml']) {
        const xml = zip.readAsText(part);
        for (const m of xml.matchAll(/<w:tab w:val="(?:right|center|decimal)"[^>]*w:pos="(\d+)"/g)) {
          expect(Number(m[1]), `${part} tab stop ${m[1]} on a ${text} twip text width`).toBeLessThanOrEqual(text + 3);
        }
      }
    }, 60_000);
  }
});
