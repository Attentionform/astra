import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, convertAstToHtml, convertAstToMarkdown } from '@attentionform/astra-engine';
import { DOMParser } from '../../packages/engine/src/ooxml/dom-parser';
import { readTableMerges } from '../../packages/engine/src/export/table-merges';

/**
 * Merged cells in the HTML and Markdown exports. A table with gridSpan, vMerge or hMerge used to export one
 * cell per physical w:tc with no colspan or rowspan, so the columns slid out of line. HTML now writes colspan
 * and rowspan and leaves covered cells out; Markdown writes the text of a merged range once and says so.
 */
interface Cell {
  t?: string;
  /** w:gridSpan */
  gs?: number;
  /** w:vMerge: '' is a bare marker (continue) */
  v?: 'restart' | 'continue' | '';
  /** w:hMerge */
  h?: 'restart' | 'continue' | '';
  /** raw XML put in the cell after its text, e.g. a nested table */
  extra?: string;
  /** put the cell in a content control, as a row-level w:sdt does */
  sdt?: boolean;
}

function tcXml(c: Cell): string {
  const marks: string[] = [];
  if (c.gs !== undefined) marks.push(`<w:gridSpan w:val="${c.gs}"/>`);
  if (c.h !== undefined) marks.push(c.h === '' ? '<w:hMerge/>' : `<w:hMerge w:val="${c.h}"/>`);
  if (c.v !== undefined) marks.push(c.v === '' ? '<w:vMerge/>' : `<w:vMerge w:val="${c.v}"/>`);
  const pr = marks.length > 0 ? `<w:tcPr>${marks.join('')}</w:tcPr>` : '';
  const text = c.t !== undefined && c.t !== '' ? `<w:r><w:t>${c.t}</w:t></w:r>` : '';
  const tc = `<w:tc>${pr}<w:p>${text}</w:p>${c.extra ?? ''}</w:tc>`;
  return c.sdt ? `<w:sdt><w:sdtContent>${tc}</w:sdtContent></w:sdt>` : tc;
}

function tblXml(rows: Cell[][]): string {
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${rows
    .map((row) => `<w:tr>${row.map(tcXml).join('')}</w:tr>`)
    .join('')}</w:tbl>`;
}

function docxWithTable(rows: Cell[][]): Buffer {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const xml = zip.readAsText('word/document.xml');
  const start = xml.indexOf('<w:body>') + '<w:body>'.length;
  const end = xml.lastIndexOf('<w:sectPr');
  zip.updateFile('word/document.xml', Buffer.from(xml.slice(0, start) + tblXml(rows) + '<w:p/>' + xml.slice(end), 'utf8'));
  return zip.toBuffer();
}

async function exportBoth(rows: Cell[][]): Promise<{ html: string; md: string }> {
  const parser = new DocxParser(docxWithTable(rows));
  const ast = await parser.getAST();
  return { html: convertAstToHtml(ast, 'T', parser), md: convertAstToMarkdown(ast, parser) };
}

const table = (html: string): string => html.slice(html.indexOf('<table'), html.indexOf('</table>') + '</table>'.length);
const bodyRows = (html: string): string[] => (table(html).match(/<tr>.*?<\/tr>/g) ?? []).map((r) => r.replace(/ style="[^"]*"/g, ''));

describe('HTML export of merged cells', () => {
  it('writes colspan for a gridSpan cell (Word style)', async () => {
    const { html } = await exportBoth([
      [{ t: 'A', gs: 2 }, { t: 'B' }],
      [{ t: 'c' }, { t: 'd' }, { t: 'e' }],
    ]);
    expect(bodyRows(html)).toEqual(['<tr><th colspan="2">A</th><th>B</th></tr>', '<tr><td>c</td><td>d</td><td>e</td></tr>']);
  });

  it('writes rowspan for vMerge restart and leaves the covered cell out (Word style)', async () => {
    const { html } = await exportBoth([
      [{ t: 'H1' }, { t: 'H2' }],
      [{ t: 'X', v: 'restart' }, { t: 'y' }],
      [{ v: '' }, { t: 'z' }],
    ]);
    expect(bodyRows(html)).toEqual([
      '<tr><th>H1</th><th>H2</th></tr>',
      '<tr><td rowspan="2">X</td><td>y</td></tr>',
      '<tr><td>z</td></tr>',
    ]);
  });

  it('counts a vertical merge over three rows', async () => {
    const { html } = await exportBoth([
      [{ t: 'H' }, { t: 'I' }],
      [{ t: 'X', v: 'restart' }, { t: 'a' }],
      [{ v: 'continue' }, { t: 'b' }],
      [{ v: 'continue' }, { t: 'c' }],
    ]);
    expect(bodyRows(html)[1]).toBe('<tr><td rowspan="3">X</td><td>a</td></tr>');
    expect(bodyRows(html)[3]).toBe('<tr><td>c</td></tr>');
  });

  it('reads the legacy hMerge markers the engine writes', async () => {
    const { html } = await exportBoth([
      [{ t: 'h1', h: 'restart' }, { t: 'h2', h: 'continue' }, { t: 'h3' }],
      [{ t: 'a' }, { t: 'b' }, { t: 'c' }],
    ]);
    expect(bodyRows(html)[0]).toBe('<tr><th colspan="2">h1 h2</th><th>h3</th></tr>');
  });

  it('reads a 2x2 block written with both markers on every cell, as mergeCells writes it', async () => {
    const { html } = await exportBoth([
      [{ t: 'a' }, { t: 'b' }, { t: 'c' }, { t: 'd' }],
      [{ t: 'e' }, { t: 'f' }, { t: 'g', h: 'restart', v: 'restart' }, { t: 'h', h: 'continue', v: 'restart' }],
      [{ t: 'i' }, { t: 'j' }, { t: 'k', h: 'restart', v: 'continue' }, { t: 'l', h: 'continue', v: 'continue' }],
    ]);
    const rows = bodyRows(html);
    // A renderer shows the text of the cell merged in from the same row ("h") and hides the cells below ("k", "l").
    expect(rows[1]).toBe('<tr><td>e</td><td>f</td><td colspan="2" rowspan="2">g h</td></tr>');
    expect(html).not.toContain('>k<');
    expect(html).not.toContain('l</td>');
    expect(rows[2]).toBe('<tr><td>i</td><td>j</td></tr>');
  });

  it('puts the rows that a header cell spans into the head', async () => {
    const { html } = await exportBoth([
      [{ t: 'Name', v: 'restart' }, { t: 'Score', gs: 2 }],
      [{ v: 'continue' }, { t: 'A' }, { t: 'B' }],
      [{ t: 'Ann' }, { t: '1' }, { t: '2' }],
    ]);
    const t = table(html);
    expect(t.indexOf('</thead>')).toBeGreaterThan(t.indexOf('<th rowspan="2"'));
    expect(t.slice(0, t.indexOf('</thead>'))).toMatch(/<th[^>]*>A<\/th>/);
    expect(t.slice(t.indexOf('<tbody>'))).toContain('<td>Ann</td>');
    expect(t).toContain('<th colspan="2"');
  });

  it('shows an orphan vMerge continue as a plain cell instead of dropping it', async () => {
    const { html } = await exportBoth([
      [{ t: 'a', v: 'continue' }, { t: 'b' }],
      [{ t: 'c' }, { t: 'd' }],
    ]);
    expect(bodyRows(html)[0]).toBe('<tr><th>a</th><th>b</th></tr>');
  });

  it('ends a vertical merge at a plain cell: a later continue has nothing to join', async () => {
    const { html } = await exportBoth([
      [{ t: 'H' }, { t: 'I' }],
      [{ t: 'X', v: 'restart' }, { t: 'a' }],
      [{ t: 'plain' }, { t: 'b' }],
      [{ t: 'late', v: 'continue' }, { t: 'c' }],
    ]);
    const rows = bodyRows(html);
    expect(rows[1]).toBe('<tr><td>X</td><td>a</td></tr>');
    expect(rows[3]).toBe('<tr><td>late</td><td>c</td></tr>');
  });

  it('keeps the shading and text colour of a merged cell', async () => {
    const zip = new AdmZip(docxWithTable([[{ t: 'A', gs: 2 }], [{ t: 'b' }, { t: 'c' }]]));
    const xml = zip.readAsText('word/document.xml').replace('<w:tcPr><w:gridSpan w:val="2"/>', '<w:tcPr><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="FF0000"/>');
    zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
    const parser = new DocxParser(zip.toBuffer());
    const html = convertAstToHtml(await parser.getAST(), 'T', parser);
    expect(html).toMatch(/<th colspan="2" style="background-color: #FF0000; color: #FFFFFF;">A<\/th>/);
  });

  it('does not take rows of a nested table as rows of the outer table', async () => {
    const nested = tblXml([[{ t: 'n1' }, { t: 'n2' }], [{ t: 'n3' }, { t: 'n4' }]]);
    const { html } = await exportBoth([
      [{ t: 'H', gs: 2 }],
      [{ t: 'outer', extra: nested }, { t: 'x' }],
    ]);
    // the outer table has two rows; the nested rows are not added to it
    expect(bodyRows(html).length).toBe(2);
    expect(bodyRows(html)[0]).toBe('<tr><th colspan="2">H</th></tr>');
  });

  it('reads cells that sit in a content control instead of dropping them', async () => {
    const { html } = await exportBoth([
      [{ t: 'A', gs: 2, sdt: true }, { t: 'B', sdt: true }],
      [{ t: 'c', sdt: true }, { t: 'd' }, { t: 'e' }],
    ]);
    expect(bodyRows(html)).toEqual(['<tr><th colspan="2">A</th><th>B</th></tr>', '<tr><td>c</td><td>d</td><td>e</td></tr>']);
  });

  it('leaves a table without merges exactly as before', async () => {
    const { html, md } = await exportBoth([
      [{ t: 'a' }, { t: 'b' }],
      [{ t: 'c' }, { t: 'd' }],
    ]);
    expect(html).not.toContain('colspan');
    expect(html).not.toContain('rowspan');
    expect(table(html)).toBe(
      '<table class="doc-table">\n<thead>\n<tr><th style="background-color: #F1F5F9; color: #0F172A;">a</th><th style="background-color: #F1F5F9; color: #0F172A;">b</th></tr>\n</thead>\n<tbody>\n<tr><td>c</td><td>d</td></tr>\n</tbody>\n</table>',
    );
    expect(md).not.toContain('Note: this table');
  });
});

describe('Markdown export of merged cells', () => {
  it('writes the merged text once and keeps the columns aligned, with one note', async () => {
    const { md } = await exportBoth([
      [{ t: 'A', gs: 2 }, { t: 'B' }],
      [{ t: 'c' }, { t: 'd' }, { t: 'e' }],
    ]);
    expect(md).toContain('| A |  | B |\n| --- | --- | --- |\n| c | d | e |');
    expect(md.match(/Note: this table has 1 merged range\./g)?.length).toBe(1);
    expect(md).toContain('Markdown cannot show merged cells');
  });

  it('puts the text of a vertical merge in its first cell and leaves the cells below empty', async () => {
    const { md } = await exportBoth([
      [{ t: 'H1' }, { t: 'H2' }],
      [{ t: 'X', v: 'restart' }, { t: 'y' }],
      [{ v: '' }, { t: 'z' }],
    ]);
    expect(md).toContain('| H1 | H2 |\n| --- | --- |\n| X | y |\n|  | z |');
  });

  it('counts ranges in the note and joins the text of a same-row covered cell once', async () => {
    const { md } = await exportBoth([
      [{ t: 'h1', h: 'restart' }, { t: 'h2', h: 'continue' }, { t: 'h3' }, { t: 'h4' }],
      [{ t: 'a', v: 'restart' }, { t: 'b' }, { t: 'c' }, { t: 'd' }],
      [{ t: 'a2', v: 'continue' }, { t: 'e' }, { t: 'f' }, { t: 'g' }],
    ]);
    expect(md).toContain('| h1 h2 |  | h3 | h4 |');
    // The text of a cell merged in from below is hidden by a renderer, so it is not written.
    expect(md).toContain('| a | b | c | d |');
    expect(md).toContain('|  | e | f | g |');
    expect(md).toContain('this table has 2 merged ranges.');
    // The text of a cell merged in from the same row is written once, in the cell it joined.
    expect(md.split('h2').length - 1).toBe(1);
    expect(md).not.toContain('a2');
  });

  it('escapes a pipe in merged text', async () => {
    const { md } = await exportBoth([
      [{ t: 'a|b', gs: 2 }],
      [{ t: 'c' }, { t: 'd' }],
    ]);
    expect(md).toContain('| a\\|b |  |');
  });
});

describe('readTableMerges', () => {
  const read = (rows: Cell[][]) => {
    const doc = new DOMParser().parseFromString(
      `<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${tblXml(rows)}</w:root>`,
      'text/xml',
    );
    return readTableMerges(doc.getElementsByTagName('w:tbl')[0] as unknown as Element);
  };

  it('reports no merges for a plain table', () => {
    const r = read([[{ t: 'a' }, { t: 'b' }]]);
    expect(r.hasMerges).toBe(false);
    expect(r.mergedRanges).toBe(0);
  });

  it('reads a gridSpan of 3 as colspan 3 and a wide vertical merge once per row', () => {
    const r = read([
      [{ t: 'A', gs: 3, v: 'restart' }],
      [{ gs: 3, v: 'continue' }],
      [{ t: 'x' }, { t: 'y' }, { t: 'z' }],
    ]);
    expect(r.rows[0][0].colspan).toBe(3);
    expect(r.rows[0][0].rowspan).toBe(2);
    expect(r.rows[1][0].covered).toBe(true);
    expect(r.mergedRanges).toBe(1);
  });

  it('ignores a gridSpan that is not a positive integer', () => {
    const r = read([[{ t: 'a', gs: 0 }, { t: 'b' }]]);
    expect(r.rows[0][0].colspan).toBe(1);
    expect(r.hasMerges).toBe(false);
  });
});
