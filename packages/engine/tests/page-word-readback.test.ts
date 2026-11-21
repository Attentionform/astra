/**
 * What Word 16 (16.0.14334) read back from files that ASTra made for these requests, in points, turned to twips. The requests are the
 * ones the Word check used; the numbers are what Word reported, so a change in what ASTra writes shows up here.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine } from '../src/internal';

const sectOf = (parser: DocxParser): string => [...parser.readDocumentXml().matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]).pop() as string;
const read = (sect: string, re: RegExp): number => Number(re.exec(sect)?.[1]);
const make = async (tpl: string, page: object, ops: object[] = []): Promise<DocxParser> => {
  const parser = DocxParser.createNewDocument(tpl as never, { page } as never);
  if (ops.length > 0) {
    const res = await new OperationEngine(await parser.getAST(), parser).executeBatch({ operations: ops as never });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
  }
  return parser;
};
const near = (twips: number, points: number): void => expect(Math.abs(twips - points * 20)).toBeLessThanOrEqual(1);

describe('files as Word reads them', () => {
  it('A5 report with five margin values: Word reads the sizes and margins asked for', async () => {
    const s = sectOf(await make('report', { size: 'A5', margins: { top: '22mm', bottom: '18mm', left: '17mm', right: '19mm', header: '9mm', footer: '8mm' } }));
    near(read(s, /w:w="(\d+)"/), 419.55); near(read(s, /w:h="(\d+)"/), 595.3);
    near(read(s, /w:top="(\d+)"/), 62.35); near(read(s, /w:bottom="(\d+)"/), 51); near(read(s, /w:left="(\d+)"/), 48.2); near(read(s, /w:right="(\d+)"/), 53.85);
    near(read(s, /w:header="(\d+)"/), 25.5); near(read(s, /w:footer="(\d+)"/), 22.7);
  });
  it('custom 150 x 200 mm memo', async () => {
    const s = sectOf(await make('memo', { size: '150x200mm', margins: { top: '15mm', bottom: '16mm', left: '14mm', right: '13mm' } }));
    near(read(s, /w:w="(\d+)"/), 425.2); near(read(s, /w:h="(\d+)"/), 566.95);
    near(read(s, /w:top="(\d+)"/), 42.5); near(read(s, /w:bottom="(\d+)"/), 45.35); near(read(s, /w:left="(\d+)"/), 39.7); near(read(s, /w:right="(\d+)"/), 36.85);
  });
  it('gutter with mirrored margins: 12 mm gutter on the left edge, mirror on', async () => {
    const parser = await make('report', { size: 'A5' }, [{ action: 'pageSetup', margins: { gutter: '12mm' }, mirrorMargins: true }]);
    near(read(sectOf(parser), /w:gutter="(\d+)"/), 34);
    expect(new AdmZip(parser.rezip()).readAsText('word/settings.xml')).toMatch(/<w:mirrorMargins\b/);
    expect(sectOf(parser)).not.toMatch(/gutterAtTop/);
  });
  it('page with the minimum text width: 0.75 in, which is 54 pt in Word', async () => {
    const s = sectOf(await make('blank', { size: '5x7in', margins: { left: '2.125in', right: '2.125in', top: '1in', bottom: '1in' } }));
    expect(read(s, /w:w="(\d+)"/) - read(s, /w:left="(\d+)"/) - read(s, /w:right="(\d+)"/)).toBe(1080);
  });
});
