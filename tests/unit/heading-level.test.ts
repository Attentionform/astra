import { describe, it, expect } from 'vitest';
import { AdmZip, DocxParser, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { headingLevelOfStyle } from '../../packages/engine/src/parser';

/**
 * Which paragraphs are headings is unchanged (a top-level body paragraph whose style id contains `heading`,
 * `title` or `subtitle`, case-insensitive). The level of numbered styles and of Subtitle is now the real one:
 * Heading1 to Heading9 give 1 to 9, Title 1, Subtitle 2; any other matching style keeps level 1.
 */
function docWith(styles: string[], extra = ''): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const xml = zip.readAsText('word/document.xml');
  const paras = styles
    .map((s, i) => `<w:p><w:pPr><w:pStyle w:val="${s}"/></w:pPr><w:r><w:t>H${i}-${s}</w:t></w:r></w:p><w:p><w:r><w:t>Body ${i}.</w:t></w:r></w:p>`)
    .join('');
  const at = xml.lastIndexOf('<w:sectPr');
  zip.updateFile('word/document.xml', Buffer.from(`${xml.slice(0, at)}${paras}${extra}${xml.slice(at)}`, 'utf8'));
  return zip.toBuffer();
}

async function headings(buf: Buffer): Promise<Array<{ text: string; level: number }>> {
  // The parser fills `section.heading` never: sections are the file's page sections. Headings are read from the
  // structural summary, which is where the level is computed.
  const parser = new DocxParser(buf);
  await parser.getAST();
  return parser
    .getStructuralSummary()
    .nodes.filter((n) => n.isHeading && /^H\d+-/.test(n.text))
    .map((n) => ({ text: n.text, level: n.headingLevel as number }));
}

describe('headingLevelOfStyle', () => {
  it('numbered styles give their number, spaces ignored, 1 to 9 only', () => {
    for (let n = 1; n <= 9; n++) {
      expect(headingLevelOfStyle(`heading${n}`)).toBe(n);
      expect(headingLevelOfStyle(`heading ${n}`)).toBe(n);
    }
    expect(headingLevelOfStyle('heading10')).toBe(1);
    expect(headingLevelOfStyle('heading0')).toBe(1);
  });

  it('Title is 1, Subtitle is 2, other matching styles keep level 1', () => {
    expect(headingLevelOfStyle('title')).toBe(1);
    expect(headingLevelOfStyle('subtitle')).toBe(2);
    for (const other of ['tocheading', 'heading1char', 'booktitle', 'heading']) expect(headingLevelOfStyle(other), other).toBe(1);
  });
});

describe('headings in a parsed document', () => {
  it('Heading1 to Heading9, Title and Subtitle carry their levels', async () => {
    const styles = ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6', 'Heading7', 'Heading8', 'Heading9', 'Title', 'Subtitle'];
    const found = await headings(docWith(styles));
    expect(found.map((h) => h.level)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 1, 2]);
  });

  it('which paragraphs are headings is unchanged: custom styles and h1 are not, TOCHeading is, at level 1', async () => {
    const found = await headings(docWith(['Kapitel', 'h1', 'TOCHeading', 'Heading1Char', 'Heading10']));
    expect(found.map((h) => h.text)).toEqual(['H2-TOCHeading', 'H3-Heading1Char', 'H4-Heading10']);
    expect(found.every((h) => h.level === 1)).toBe(true);
  });

  it('a paragraph that holds a textbox does not take the style of a paragraph inside the textbox', async () => {
    const box = '<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>H9-in a box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>';
    const parser = new DocxParser(docWith(['Heading1'], box));
    await parser.getAST();
    const hosts = parser.getStructuralSummary().nodes.filter((n) => n.text.includes('in a box'));
    expect(hosts.every((n) => !n.isHeading)).toBe(true);
  });

  it('a table that holds a heading-styled cell paragraph is not a heading', async () => {
    const parser = new DocxParser(docWith(['Heading1'], '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>In a cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>'));
    await parser.getAST();
    const tables = parser.getStructuralSummary().nodes.filter((n) => n.type === 'table');
    expect(tables).toHaveLength(1);
    expect(tables[0].isHeading).toBe(false);
  });

  it('a heading inside a table cell is not a heading', async () => {
    const cell = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>In a cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const found = await headings(docWith(['Heading1'], cell));
    expect(found.map((h) => h.text)).toEqual(['H0-Heading1']);
  });

  it('a Heading4 after a Heading2 keeps level 4, not 1', async () => {
    const buf = docWith(['Heading1', 'Heading2', 'Heading4', 'Heading1']);
    const found = await headings(buf);
    expect(found.map((h) => h.level)).toEqual([1, 2, 4, 1]);
  });
});

describe('namedStyle of a node is the style of its own paragraph', () => {
  it('a table and a textbox host paragraph report no style of a paragraph nested inside them', async () => {
    const box = '<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>H9-in a box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>';
    const cell = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>In a cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const ast = await new DocxParser(docWith(['Heading1'], cell + box)).getAST();
    const nodes = ast.sections.flatMap((s) => s.nodes);
    const table = nodes.find((n) => n.type === 'table')!;
    const host = nodes.find((n) => n.metadata.text?.includes('in a box'))!;
    expect(table.metadata.namedStyle).toBeUndefined();
    expect(host.metadata.namedStyle).toBeUndefined();
    expect(nodes.find((n) => n.metadata.text === 'H0-Heading1')!.metadata.namedStyle).toBe('Heading1');
  });
});
