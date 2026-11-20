/**
 * Body helpers and package assembly shared by every template. A template file builds its body with these
 * helpers and hands it to `buildPackage`, which writes a complete, deterministic .docx.
 */
import AdmZip from '../../src/adm-zip-safe';
import { clampZipEntryTimes } from '../../src/zip-hardening';
import { W_NS, XML_DECL, esc, fontTableXml, numberingXml, settingsXml, stylesXml, themeXml, type Theme } from './design';

export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Fixed stamp so the same inputs always give the same bytes. */
export const BUILD_DATE = new Date(Date.UTC(2026, 0, 1, 0, 0, 0));
const BUILD_ISO = '2026-01-01T00:00:00Z';

// ---- runs and paragraphs ------------------------------------------------------------------------

export interface RunOpts {
  style?: string;
  bold?: boolean;
  italic?: boolean;
  /** Extra run properties, already in schema order. */
  raw?: string;
}

export function run(text: string, o: RunOpts = {}): string {
  const rPr = (o.style ? `<w:rStyle w:val="${o.style}"/>` : '') + (o.bold ? '<w:b/>' : '') + (o.italic ? '<w:i/>' : '') + (o.raw ?? '');
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

/** A bracketed placeholder in the Placeholder character style. */
export function ph(label: string): string {
  return run(`[${label}]`, { style: 'Placeholder' });
}

/** A complex field: instruction plus the text Word shows until it updates. */
export function field(instr: string, cached: string, o: { dirty?: boolean; rPr?: string } = {}): string {
  const rPr = o.rPr ? `<w:rPr>${o.rPr}</w:rPr>` : '';
  return (
    `<w:r>${rPr}<w:fldChar w:fldCharType="begin"${o.dirty ? ' w:dirty="true"' : ''}/></w:r>` +
    `<w:r>${rPr}<w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${rPr}<w:t xml:space="preserve">${esc(cached)}</w:t></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
  );
}

export interface ParaOpts {
  style?: string;
  /** Extra paragraph properties after pStyle, already in schema order. */
  raw?: string;
}

export function p(content: string | string[], o: ParaOpts = {}): string {
  const body = Array.isArray(content) ? content.join('') : content;
  const pPr = (o.style ? `<w:pStyle w:val="${o.style}"/>` : '') + (o.raw ?? '');
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${body}</w:p>`;
}

/** Plain text paragraph. */
export function para(text: string, style?: string): string {
  return p(text === '' ? '' : run(text), style ? { style } : {});
}

/** A paragraph whose whole content is one placeholder. */
export function phPara(label: string, style?: string): string {
  return p(ph(label), style ? { style } : {});
}

export function pageBreak(): string {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

/** Empty bookmark pair at body level. `id` must be unique in the document. */
export function marker(id: number, name: string): string {
  return `<w:bookmarkStart w:id="${id}" w:name="${name}"/><w:bookmarkEnd w:id="${id}"/>`;
}

// ---- tables -------------------------------------------------------------------------------------

export interface Cell {
  /** Cell content as runs, or paragraphs when `blocks` is set. */
  content: string;
  /** Use the content as-is (already paragraphs). */
  blocks?: boolean;
  align?: 'left' | 'right' | 'center';
  span?: number;
  shade?: string;
  /** Raw tcPr children after tcW, in schema order. */
  raw?: string;
}

export function cell(content: string, o: Omit<Cell, 'content'> = {}): Cell {
  return { content, ...o };
}

export interface TableOpts {
  style: string;
  /** Column widths in twips; their sum is the table width. */
  widths: number[];
  rows: Cell[][];
  /** First row is a header: repeats on each page and is read as a header by assistive tools. */
  header?: boolean;
  /** tblLook flags. */
  look?: { firstRow?: boolean; banded?: boolean };
  /** Keep rows from splitting across pages. */
  cantSplit?: boolean;
  /** Paragraph style for cell text. */
  textStyle?: string;
}

export function table(o: TableOpts): string {
  const total = o.widths.reduce((a, b) => a + b, 0);
  const look = o.look ?? {};
  const grid = o.widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('');
  const textStyle = o.textStyle ?? 'TableText';
  const rows = o.rows
    .map((row, ri) => {
      const isHeader = o.header && ri === 0;
      const trPr = (o.cantSplit ? '<w:cantSplit/>' : '') + (isHeader ? '<w:tblHeader/>' : '');
      let col = 0;
      const cells = row
        .map((c) => {
          const span = c.span ?? 1;
          const w = o.widths.slice(col, col + span).reduce((a, b) => a + b, 0);
          col += span;
          const tcPr =
            `<w:tcW w:w="${w}" w:type="dxa"/>` + (span > 1 ? `<w:gridSpan w:val="${span}"/>` : '') + (c.raw ?? '') + (c.shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${c.shade}"/>` : '');
          const jc = c.align && c.align !== 'left' ? `<w:jc w:val="${c.align}"/>` : '';
          const inner = c.blocks ? c.content : `<w:p><w:pPr><w:pStyle w:val="${textStyle}"/>${jc}</w:pPr>${c.content}</w:p>`;
          return `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`;
        })
        .join('');
      return `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
    })
    .join('');
  const lookAttr = `w:val="0000" w:firstRow="${look.firstRow === false ? 0 : o.header ? 1 : 0}" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="${look.banded ? 0 : 1}" w:noVBand="1"`;
  return (
    `<w:tbl><w:tblPr><w:tblStyle w:val="${o.style}"/><w:tblW w:w="${total}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook ${lookAttr}/></w:tblPr>` +
    `<w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`
  );
}

// ---- page, header and footer --------------------------------------------------------------------

export interface PageSetup {
  /** Twips. */
  width: number;
  height: number;
  top: number;
  bottom: number;
  left: number;
  right: number;
  header: number;
  footer: number;
}

export const LETTER: PageSetup = { width: 12240, height: 15840, top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720 };
export const A4: PageSetup = { width: 11906, height: 16838, top: 1440, bottom: 1440, left: 1417, right: 1417, header: 709, footer: 709 };

/** An empty paragraph that ends a section (next section starts on a new page). Holds the ending section's page setup. */
export function sectionBreak(page: PageSetup, o: { continuous?: boolean; cols?: number } = {}): string {
  return `<w:p><w:pPr><w:pStyle w:val="NoSpacing"/>${sectPr(page, o)}</w:pPr></w:p>`;
}

export interface HeaderFooterSpec {
  /** Header XML body (paragraphs). */
  header?: string;
  footer?: string;
  /** Different first page: the first page header and footer are the empty ones. */
  titlePg?: boolean;
}

export function sectPr(page: PageSetup, hf: { defaultHeader?: string; defaultFooter?: string; firstHeader?: string; firstFooter?: string; titlePg?: boolean; restartNumbering?: boolean; cols?: number; continuous?: boolean }): string {
  const refs =
    (hf.defaultHeader ? `<w:headerReference w:type="default" r:id="${hf.defaultHeader}"/>` : '') +
    (hf.firstHeader ? `<w:headerReference w:type="first" r:id="${hf.firstHeader}"/>` : '') +
    (hf.defaultFooter ? `<w:footerReference w:type="default" r:id="${hf.defaultFooter}"/>` : '') +
    (hf.firstFooter ? `<w:footerReference w:type="first" r:id="${hf.firstFooter}"/>` : '');
  return (
    `<w:sectPr>${refs}${hf.continuous ? '<w:type w:val="continuous"/>' : ''}<w:pgSz w:w="${page.width}" w:h="${page.height}"/>` +
    `<w:pgMar w:top="${page.top}" w:right="${page.right}" w:bottom="${page.bottom}" w:left="${page.left}" w:header="${page.header}" w:footer="${page.footer}" w:gutter="0"/>` +
    (hf.restartNumbering ? '<w:pgNumType w:start="1"/>' : '') +
    (hf.cols && hf.cols > 1 ? `<w:cols w:num="${hf.cols}" w:space="480"/>` : '') +
    (hf.titlePg ? '<w:titlePg/>' : '') +
    '</w:sectPr>'
  );
}

const hdrFtr = (tag: 'hdr' | 'ftr', body: string): string => `${XML_DECL}<w:${tag} xmlns:w="${W_NS}" xmlns:r="${R_NS}">${body}</w:${tag}>`;

/** "Page X of Y" paragraph for a footer, with the right tab for the page number. */
export function pageXofY(left = '', total: 'NUMPAGES' | 'SECTIONPAGES' = 'NUMPAGES'): string {
  return p([left, run('\t'), run('Page '), field('PAGE', '1'), run(' of '), field(total, '1')], { style: 'Footer' });
}

/** Running title paragraph for a header: document title field at the left, optional text at the right. */
export function runningTitle(right = ''): string {
  return p([field('TITLE', ''), run('\t'), right ? run(right) : ''], { style: 'Header' });
}

// ---- package assembly ---------------------------------------------------------------------------

export interface BuildInput {
  theme: Theme;
  /** Template name, written to the AstraTemplate custom property with the design version. */
  template: string;
  designVersion: number;
  /** Everything inside w:body except the final sectPr. */
  body: string;
  page: PageSetup;
  /** Header and footer parts by role. */
  parts?: { defaultHeader?: string; defaultFooter?: string; firstHeader?: string; firstFooter?: string; titlePg?: boolean; restartNumbering?: boolean; cols?: number; continuous?: boolean };
  /** dc:title. Empty for a starter. */
  title?: string;
  subject?: string;
  keywords?: string;
}

export function buildPackage(inp: BuildInput): Buffer {
  const parts = inp.parts ?? {};
  const rels: string[] = [
    `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`,
    `<Relationship Id="rId2" Type="${REL}/settings" Target="settings.xml"/>`,
    `<Relationship Id="rId3" Type="${REL}/fontTable" Target="fontTable.xml"/>`,
    `<Relationship Id="rId4" Type="${REL}/theme" Target="theme/theme1.xml"/>`,
    `<Relationship Id="rId5" Type="${REL}/numbering" Target="numbering.xml"/>`,
  ];
  const files: Record<string, string> = {};
  const overrides: string[] = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>',
  ];
  let n = 6;
  const addHf = (kind: 'header' | 'footer', role: string, xmlBody: string | undefined, key: 'defaultHeader' | 'defaultFooter' | 'firstHeader' | 'firstFooter'): string | undefined => {
    if (xmlBody === undefined) return undefined;
    const id = `rId${n++}`;
    const file = `${kind}${role === 'first' ? '2' : '1'}.xml`;
    files[`word/${file}`] = hdrFtr(kind === 'header' ? 'hdr' : 'ftr', xmlBody);
    rels.push(`<Relationship Id="${id}" Type="${REL}/${kind}" Target="${file}"/>`);
    overrides.push(`<Override PartName="/word/${file}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`);
    void key;
    return id;
  };
  const refs = {
    defaultHeader: addHf('header', 'default', parts.defaultHeader, 'defaultHeader'),
    firstHeader: addHf('header', 'first', parts.firstHeader, 'firstHeader'),
    defaultFooter: addHf('footer', 'default', parts.defaultFooter, 'defaultFooter'),
    firstFooter: addHf('footer', 'first', parts.firstFooter, 'firstFooter'),
    titlePg: parts.titlePg,
    restartNumbering: parts.restartNumbering,
    cols: parts.cols,
    continuous: parts.continuous,
  };

  const document =
    `${XML_DECL}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${inp.body}${sectPr(inp.page, refs)}</w:body></w:document>`;

  const core =
    `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
    'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${esc(inp.title ?? '')}</dc:title><dc:subject>${esc(inp.subject ?? '')}</dc:subject><dc:creator></dc:creator>` +
    `<cp:keywords>${esc(inp.keywords ?? '')}</cp:keywords><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>1</cp:revision>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${BUILD_ISO}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${BUILD_ISO}</dcterms:modified></cp:coreProperties>`;
  const app =
    `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    '<Application>ASTra</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged></Properties>';
  const custom =
    `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="AstraTemplate"><vt:lpwstr>${esc(`${inp.template}@${inp.designVersion}`)}</vt:lpwstr></property></Properties>`;

  const contentTypes =
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    `${overrides.join('')}</Types>`;
  const rootRels =
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>` +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` +
    `<Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/></Relationships>`;
  const documentRels = `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`;

  const zip = new AdmZip();
  const add = (name: string, text: string) => zip.addFile(name, Buffer.from(text, 'utf8'));
  add('[Content_Types].xml', contentTypes);
  add('_rels/.rels', rootRels);
  add('word/document.xml', document);
  add('word/_rels/document.xml.rels', documentRels);
  add('word/styles.xml', stylesXml(inp.theme));
  add('word/numbering.xml', numberingXml(inp.theme));
  add('word/settings.xml', settingsXml());
  add('word/fontTable.xml', fontTableXml(inp.theme));
  add('word/theme/theme1.xml', themeXml(inp.theme));
  add('docProps/core.xml', core);
  add('docProps/app.xml', app);
  add('docProps/custom.xml', custom);
  for (const [name, text] of Object.entries(files)) add(name, text);
  clampZipEntryTimes(zip, BUILD_DATE);
  return zip.toBuffer();
}
