import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '../src/internal';
import { OperationSchema } from '../src/operation-schema';
import { operationIssueRemedy } from '../src/operation-hints';

const sectPr = (parser: DocxParser): string => new AdmZip(parser.rezip()).readAsText('word/document.xml').match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)![0];
const pgMar = (parser: DocxParser): Record<string, string> => Object.fromEntries([...(/<w:pgMar([^>]*)\/>/.exec(sectPr(parser))?.[1] ?? '').matchAll(/w:(\w+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));

async function withPage(pgSz: string, pgMarXml: string): Promise<DocxParser> {
  const parser = DocxParser.createNewDocument('blank');
  const zip = new AdmZip(parser.rezip());
  const xml = zip.readAsText('word/document.xml').replace(/<w:pgSz[^>]*\/>/, pgSz).replace(/<w:pgMar[^>]*\/>/, pgMarXml);
  zip.updateFile('word/document.xml', Buffer.from(xml));
  return new DocxParser(zip.toBuffer());
}
async function apply(parser: DocxParser, op: object) {
  const engine = new OperationEngine(await parser.getAST(), parser);
  return engine.executeBatch({ operations: [op] as never });
}

describe('the names that an earlier version took are refused, never ignored', () => {
  const cases: Array<[string, object, RegExp]> = [
    ['marginTop at the top level', { action: 'pageSetup', marginTop: 720 }, /marginTop.*margins\.top/],
    ['marginBottom', { action: 'pageSetup', marginBottom: 720 }, /margins\.bottom/],
    ['marginLeft', { action: 'pageSetup', marginLeft: 720 }, /margins\.left/],
    ['marginRight', { action: 'pageSetup', marginRight: 720 }, /margins\.right/],
    ['headerDistance', { action: 'pageSetup', headerDistance: 720 }, /margins\.header/],
    ['footerDistance', { action: 'pageSetup', footerDistance: 720 }, /margins\.footer/],
    ['the header alias', { action: 'pageSetup', header: 720 }, /margins\.header/],
    ['the footer alias', { action: 'pageSetup', footer: 720 }, /margins\.footer/],
    ['marginTop in params', { action: 'pageSetup', params: { marginTop: 720 } }, /marginTop.*margins\.top/],
    ['headerDistance in params', { action: 'pageSetup', params: { headerDistance: 720 } }, /margins\.header/],
    ['pageSize', { action: 'pageSetup', pageSize: 'a4' }, /paperSize/],
    ['an unknown name in params', { action: 'pageSetup', params: { colour: 'red' } }, /not a page setup field/],
  ];
  for (const [label, op, pattern] of cases) {
    it(`${label}: the library refuses it with the replacement and changes nothing`, async () => {
      const parser = DocxParser.createNewDocument('blank');
      const before = sectPr(parser);
      const res = await apply(parser, op);
      expect(res.success).toBe(false);
      expect(JSON.stringify(res.errors)).toMatch(pattern);
      expect(sectPr(parser)).toBe(before);
    });
  }
  it('no accepted request is a silent no-op: every field that is given changes the file', async () => {
    const fields: object[] = [{ paperSize: 'A5' }, { orientation: 'landscape' }, { margins: { top: 1000 } }, { margins: { header: 300 } }, { margins: { gutter: 200 } }, { columns: 2 }, { mirrorMargins: true }, { gutterPosition: 'top' }];
    for (const f of fields) {
      const parser = DocxParser.createNewDocument('blank');
      const before = sectPr(parser) + new AdmZip(parser.rezip()).readAsText('word/settings.xml');
      const res = await apply(parser, { action: 'pageSetup', ...f });
      expect(res.success, JSON.stringify(f) + JSON.stringify(res.errors)).toBe(true);
      expect(sectPr(parser) + new AdmZip(parser.rezip()).readAsText('word/settings.xml'), JSON.stringify(f)).not.toBe(before);
    }
  });
  it('the schema gate (the MCP and CLI path) refuses them too, and the hint names the replacement', () => {
    for (const op of [{ action: 'pageSetup', marginTop: 720 }, { action: 'pageSetup', params: { marginLeft: 720 } }, { action: 'pageSetup', pageSize: 'a4' }]) {
      const parsed = OperationSchema.safeParse(op);
      expect(parsed.success).toBe(false);
      if (parsed.success) continue;
      const hint = operationIssueRemedy([op], parsed.error.issues.map((i) => ({ code: i.code, path: ['operations', 0, ...i.path], keys: (i as { keys?: string[] }).keys })));
      expect(hint, JSON.stringify(op)).toMatch(/margins|paperSize/);
    }
    expect(OperationSchema.safeParse({ action: 'pageSetup', margins: { top: '20mm' }, paperSize: 'A5' }).success).toBe(true);
    expect(OperationSchema.safeParse({ action: 'pageSetup', params: { paperSize: { width: '150mm', height: '200mm' }, margins: '1in' } }).success).toBe(true);
  });
});

describe('an existing page is kept as it is unless the request changes it', () => {
  it('a negative top margin that Word accepts survives orientation, paper and other margin changes, with its value', async () => {
    for (const op of [{ orientation: 'landscape' }, { paperSize: 'Legal' }, { margins: { left: '1.2in' } }, { columns: 2 }]) {
      const parser = await withPage('<w:pgSz w:w="12240" w:h="15840"/>', '<w:pgMar w:top="-720" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>');
      const res = await apply(parser, { action: 'pageSetup', ...op });
      expect(res.success, JSON.stringify(op) + JSON.stringify(res.errors)).toBe(true);
      expect(pgMar(parser).top, JSON.stringify(op)).toBe('-720');
      expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
    }
  });
  it('a new negative margin is refused', async () => {
    const parser = await withPage('<w:pgSz w:w="12240" w:h="15840"/>', '<w:pgMar w:top="-720" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>');
    const res = await apply(parser, { action: 'pageSetup', margins: { top: '-5mm' } });
    expect(res.success).toBe(false);
    expect(pgMar(parser).top).toBe('-720');
  });
  it('the margins that the request does not set keep their parsed values', async () => {
    const parser = await withPage('<w:pgSz w:w="11906" w:h="16838" w:code="9"/>', '<w:pgMar w:top="1234" w:right="1111" w:bottom="999" w:left="2222" w:header="333" w:footer="444" w:gutter="55"/>');
    const res = await apply(parser, { action: 'pageSetup', margins: { left: '1in' } });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(pgMar(parser)).toEqual({ top: '1234', right: '1111', bottom: '999', left: '1440', header: '333', footer: '444', gutter: '55' });
  });
  it('a page above 22 in that is already in the file is not refused for an unrelated edit, and a size request is', async () => {
    const parser = await withPage('<w:pgSz w:w="43200" w:h="15840"/>', '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>');
    expect((await apply(parser, { action: 'pageSetup', margins: { top: '2in' } })).success).toBe(true);
    expect((await apply(parser, { action: 'pageSetup', paperSize: '30x10in' })).success).toBe(false);
  });
});

describe('a page that is already narrower than the minimum is still a document that works', () => {
  const narrowWidth = (text: number): string => {
    const side = Math.floor((12240 - text) / 2);
    return `<w:pgMar w:top="1440" w:right="${side}" w:bottom="1440" w:left="${side}" w:header="720" w:footer="720" w:gutter="0"/>`;
  };
  for (const text of [720, 288, 20]) {
    it(`text width ${text} twips: it validates, takes an edit, and orientation that widens it passes`, async () => {
      const parser = await withPage('<w:pgSz w:w="12240" w:h="15840"/>', narrowWidth(text));
      expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
      const edit = await apply(parser, { action: 'insert', type: 'paragraph', text: 'Added line.', position: 'after', anchor: (await parser.getAST()).sections[0].nodes[0]?.id });
      expect(edit.success, JSON.stringify(edit.errors)).toBe(true);
      expect((await apply(parser, { action: 'pageSetup', orientation: 'landscape' })).success).toBe(true);
    });
    it(`text width ${text} twips: a request that makes it narrower is refused, and one that widens it passes`, async () => {
      const parser = await withPage('<w:pgSz w:w="12240" w:h="15840"/>', narrowWidth(text));
      const side = Math.floor((12240 - text) / 2);
      expect((await apply(parser, { action: 'pageSetup', margins: { left: side + 100 } })).success).toBe(false);
      expect((await apply(parser, { action: 'pageSetup', margins: { left: side - 100 } })).success).toBe(true);
    });
  }
  it('a text height that is already too small: orientation that shrinks it further is refused, and the file is unchanged', async () => {
    const side = Math.floor((15840 - 72) / 2);
    const parser = await withPage('<w:pgSz w:w="12240" w:h="15840"/>', `<w:pgMar w:top="${side}" w:right="1440" w:bottom="${side}" w:left="1440" w:header="100" w:footer="100" w:gutter="0"/>`);
    const before = sectPr(parser);
    const res = await apply(parser, { action: 'pageSetup', orientation: 'landscape' });
    expect(res.success).toBe(false);
    expect(JSON.stringify(res.errors)).toMatch(/text height/);
    expect(sectPr(parser)).toBe(before);
  });
  it('a narrow width: orientation that does not shrink it goes through, and one that shrinks the height of a roomy page is judged by that height', async () => {
    const parser = await withPage('<w:pgSz w:w="12240" w:h="15840"/>', '<w:pgMar w:top="7200" w:right="1440" w:bottom="7200" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>');
    // Height after margins is 1440 on portrait (fine), and 12240 - 14400 on landscape (negative): refused, because it drops under 0.5 in and is smaller than before.
    const res = await apply(parser, { action: 'pageSetup', orientation: 'landscape' });
    expect(res.success).toBe(false);
    expect(JSON.stringify(res.errors)).toMatch(/text height/);
  });
});
