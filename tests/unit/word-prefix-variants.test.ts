/**
 * 0.5.3: WordprocessingML bound to another prefix. Table layout reads and writes by namespace and local name, so an
 * equivalent prefix gives the same result; a main document that does not use "w" is refused at open (the engine
 * reads the body by "w" and would otherwise see an empty document).
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { applyRowHeights, applyTableWidth, parseRowHeights, parseTableWidth } from '../../packages/engine/src/operations/table-layout';

const WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function tableWith(prefix: string): { doc: Document; tbl: Element; rows: Element[] } {
  const p = prefix ? `${prefix}:` : '';
  const xml = `<root xmlns:${prefix || 'zz'}="${WNS}"><${p}tbl>`
    + `<${p}tblPr><${p}tblW ${p}w="2000" ${p}type="dxa"/></${p}tblPr>`
    + `<${p}tblGrid><${p}gridCol ${p}w="2000"/></${p}tblGrid>`
    + `<${p}tr><${p}trPr><${p}trHeight ${p}val="400" ${p}hRule="exact"/></${p}trPr>`
    + `<${p}tc><${p}tcPr><${p}tcW ${p}w="2000" ${p}type="dxa"/></${p}tcPr><${p}p/></${p}tc></${p}tr>`
    + `</${p}tbl></root>`;
  const doc = new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;
  const tbl = doc.getElementsByTagNameNS(WNS, 'tbl')[0];
  const rows = Array.from(doc.getElementsByTagNameNS(WNS, 'tr'));
  return { doc, tbl, rows };
}
const out = (doc: Document): string => new XMLSerializer().serializeToString(doc as any);
const count = (xml: string, re: RegExp): number => (xml.match(re) ?? []).length;
const pick = (values: Record<string, unknown>) => (key: string) => values[key];

describe.each(['w', 'x'])('table layout with the WordprocessingML prefix "%s"', (prefix) => {
  it('setTableWidth scale resizes the grid, the cell and the table width', () => {
    const { doc, tbl, rows } = tableWith(prefix);
    applyTableWidth(tbl, rows, parseTableWidth(pick({ widthTwips: 1500 })), null);
    const xml = out(doc);
    expect(xml).toMatch(new RegExp(`gridCol ${prefix}:w="1500"`));
    expect(xml).toMatch(new RegExp(`tcW ${prefix}:w="1500"`));
    expect(xml).toMatch(new RegExp(`tblW ${prefix}:w="1500" ${prefix}:type="dxa"`));
    expect(xml).toMatch(new RegExp(`tblLayout ${prefix}:type="fixed"`));
    expect(count(xml, /tblLayout/g)).toBe(1);
  });

  it('setRowHeights auto clears the existing height and the empty row properties', () => {
    const { doc, tbl, rows } = tableWith(prefix);
    applyRowHeights(tbl, rows, parseRowHeights(pick({ rule: 'auto' })));
    const xml = out(doc);
    expect(xml).not.toContain('trHeight');
    expect(xml).not.toContain('trPr');
  });

  it('setRowHeights exact updates the one row-properties element', () => {
    const { doc, tbl, rows } = tableWith(prefix);
    applyRowHeights(tbl, rows, parseRowHeights(pick({ rule: 'exact', heightTwips: 500 })));
    const xml = out(doc);
    expect(count(xml, /<\w+:trPr>/g)).toBe(1);
    expect(count(xml, /trHeight/g)).toBe(1);
    expect(xml).toMatch(new RegExp(`trHeight ${prefix}:val="500" ${prefix}:hRule="exact"`));
  });

  it('a row height added to a row without row properties uses the table prefix', () => {
    const { doc, tbl, rows } = tableWith(prefix);
    const trPr = Array.from(doc.getElementsByTagNameNS(WNS, 'trPr'))[0];
    trPr.parentNode!.removeChild(trPr);
    applyRowHeights(tbl, rows, parseRowHeights(pick({ rule: 'atLeast', heightTwips: 300 })));
    expect(out(doc)).toMatch(new RegExp(`<${prefix}:trPr><${prefix}:trHeight ${prefix}:val="300" ${prefix}:hRule="atLeast"/></${prefix}:trPr>`));
  });
});

describe('a main document that does not write the "w" prefix is rewritten to "w" at open (DN3 ingress)', () => {
  const blankDocument = (): { zip: AdmZip; xml: string } => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    return { zip, xml: zip.readAsText('word/document.xml') };
  };
  const reopen = (zip: AdmZip, xml: string): DocxParser => {
    zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
    return new DocxParser(zip.toBuffer());
  };

  it('opens the usual form', async () => {
    const { zip, xml } = blankDocument();
    const ast: any = await reopen(zip, xml).getAST();
    expect(ast.sections.flatMap((s: any) => s.nodes).length).toBeGreaterThan(0);
  });

  it('rewrites a renamed prefix and reads the body', async () => {
    const { zip, xml } = blankDocument();
    const renamed = xml.replace(/<(\/?)w:/g, '<$1x:').replace(/ w:/g, ' x:').replace(/xmlns:w=/, 'xmlns:x=');
    expect(renamed).toContain('<x:body>');
    const parser = reopen(zip, renamed);
    expect(parser.canonicalisedParts.map((p) => p.part)).toContain('word/document.xml');
    const ast: any = await parser.getAST();
    expect(ast.sections.flatMap((s: any) => s.nodes).length).toBeGreaterThan(0);
  });

  it('rewrites the default namespace', async () => {
    const { zip, xml } = blankDocument();
    const dflt = xml.replace(/<(\/?)w:/g, '<$1').replace(/ w:/g, ' ').replace(/xmlns:w=/, 'xmlns=');
    const parser = reopen(zip, dflt);
    const ast: any = await parser.getAST();
    expect(ast.sections.flatMap((s: any) => s.nodes).length).toBeGreaterThan(0);
  });

  it('rewrites a single table that uses another prefix inside a "w" document', async () => {
    const { zip, xml } = blankDocument();
    const table = `<x:tbl xmlns:x="${WNS}"><x:tblPr/><x:tblGrid><x:gridCol x:w="2000"/></x:tblGrid><x:tr><x:tc><x:p/></x:tc></x:tr></x:tbl>`;
    const mixed = xml.replace('<w:body>', `<w:body>${table}`);
    expect(mixed).toContain('<x:tbl ');
    const parser = reopen(zip, mixed);
    expect(parser.readDocumentXml()).not.toContain('<x:tbl');
    expect(parser.readDocumentXml()).toContain('<w:tbl');
  });

  it('the engine still works end to end on the usual form', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id as string;
    const res = await (engine as any).executeBatch({ operations: [
      { action: 'insert', type: 'table', id: 't', anchor, position: 'after', headers: ['A'], rows: [['1']] },
      { action: 'table', target: 't', operation: 'setRowHeights', rule: 'exact', heightTwips: 500 },
      { action: 'table', target: 't', operation: 'setTableWidth', widthTwips: 3000 },
    ] });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
  });
});
