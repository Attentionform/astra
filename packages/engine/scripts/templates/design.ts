/**
 * The shared design system of the ASTra document templates: palette, type scale, paragraph, character
 * and table styles, list numbering, header and footer, theme. Every template is built from one `Theme`
 * so a change here reaches all of them. Nothing in this file is specific to one template.
 *
 * Element order inside w:pPr, w:rPr, w:style and w:tblPr follows the Open XML schema; the validator
 * and Word both reject other orders.
 */

export interface Theme {
  name: string;
  fonts: { body: string; heading: string; mono: string };
  colors: {
    /** Text. Near-black, never pure black. */
    ink: string;
    /** Secondary text: captions, header, footer, placeholders. */
    muted: string;
    /** The one accent of the family: headings, rules, table header, links. */
    accent: string;
    /** A very light tint of the accent: callout and table header fills. */
    accentSoft: string;
    /** Hairlines and table rules. */
    rule: string;
    /** Code and quote background. */
    wash: string;
    /** Placeholder text colour (must keep 4.5:1 against white). */
    placeholder: string;
  };
  /** `manuscript`: standard manuscript format (12 pt serif, double spacing, first-line indent, centred chapter headings). */
  variant?: 'manuscript';
  /** Half-points. */
  sizes: { body: number; title: number; subtitle: number; h1: number; h2: number; h3: number; h4: number; small: number; code: number };
}

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function esc(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const fonts = (t: Theme, which: 'body' | 'heading' | 'mono'): string => {
  const f = t.fonts[which];
  return `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`;
};

interface ParaStyle {
  id: string;
  name: string;
  basedOn?: string;
  next?: string;
  uiPriority?: number;
  qFormat?: boolean;
  /** Pre-built children, already in schema order. */
  pPr?: string;
  rPr?: string;
}

function paraStyle(s: ParaStyle): string {
  return (
    `<w:style w:type="paragraph" w:styleId="${s.id}">` +
    `<w:name w:val="${esc(s.name)}"/>` +
    (s.basedOn ? `<w:basedOn w:val="${s.basedOn}"/>` : '') +
    (s.next ? `<w:next w:val="${s.next}"/>` : '') +
    (s.uiPriority !== undefined ? `<w:uiPriority w:val="${s.uiPriority}"/>` : '') +
    (s.qFormat ? '<w:qFormat/>' : '') +
    (s.pPr ? `<w:pPr>${s.pPr}</w:pPr>` : '') +
    (s.rPr ? `<w:rPr>${s.rPr}</w:rPr>` : '') +
    '</w:style>'
  );
}

function charStyle(id: string, name: string, rPr: string, extra = ''): string {
  return `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${esc(name)}"/>${extra}<w:rPr>${rPr}</w:rPr></w:style>`;
}

/** Twips of one point. */
const pt = (n: number): number => Math.round(n * 20);

/** numId of the bullet list and of the decimal list defined by `numberingXml()`. */
export const NUM_BULLET = 1;
export const NUM_DECIMAL = 2;
/** numId of the multilevel heading numbering (1, 1.1, 1.1.1) used by headings that carry it directly. */
export const NUM_HEADINGS = 3;

export function stylesXml(t: Theme): string {
  const c = t.colors;
  const s = t.sizes;
  const keep = '<w:keepNext/><w:keepLines/>';
  const rule = (color: string, sz = 4, space = 4) => `<w:bottom w:val="single" w:sz="${sz}" w:space="${space}" w:color="${color}"/>`;

  const ms = t.variant === 'manuscript';
  const paragraphs: string[] = [
    paraStyle({ id: 'Normal', name: 'Normal', qFormat: true }),
    paraStyle({
      id: 'BodyText',
      name: 'Body Text',
      basedOn: 'Normal',
      uiPriority: 1,
      qFormat: true,
      pPr: ms ? '<w:spacing w:after="0" w:line="480" w:lineRule="auto"/><w:ind w:firstLine="720"/>' : '<w:spacing w:after="160" w:line="288" w:lineRule="auto"/>',
    }),
    paraStyle({
      id: 'Title',
      name: 'Title',
      basedOn: 'Normal',
      next: 'Subtitle',
      uiPriority: 10,
      qFormat: true,
      pPr: ms
        ? `<w:keepNext/><w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>`
        : `<w:keepNext/><w:spacing w:before="0" w:after="120" w:line="216" w:lineRule="auto"/><w:contextualSpacing/>`,
      rPr: ms
        ? `${fonts(t, 'heading')}<w:b/><w:caps/><w:color w:val="${c.ink}"/><w:sz w:val="28"/><w:szCs w:val="28"/>`
        : `${fonts(t, 'heading')}<w:b/><w:color w:val="${c.ink}"/><w:spacing w:val="-6"/><w:kern w:val="28"/><w:sz w:val="${s.title}"/><w:szCs w:val="${s.title}"/>`,
    }),
    paraStyle({
      id: 'Subtitle',
      name: 'Subtitle',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 11,
      qFormat: true,
      pPr: ms ? '<w:spacing w:before="0" w:after="240" w:line="264" w:lineRule="auto"/><w:jc w:val="center"/>' : '<w:spacing w:before="0" w:after="240" w:line="264" w:lineRule="auto"/>',
      rPr: ms ? `${fonts(t, 'heading')}<w:sz w:val="${s.body}"/><w:szCs w:val="${s.body}"/>` : `${fonts(t, 'heading')}<w:color w:val="${c.muted}"/><w:sz w:val="${s.subtitle}"/><w:szCs w:val="${s.subtitle}"/>`,
    }),
    paraStyle({
      id: 'Heading1',
      name: 'heading 1',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 9,
      qFormat: true,
      pPr: ms
        ? `${keep}<w:pageBreakBefore/><w:spacing w:before="2880" w:after="480" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/>`
        : `${keep}<w:spacing w:before="${pt(24)}" w:after="${pt(8)}" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="0"/>`,
      rPr: ms
        ? `${fonts(t, 'heading')}<w:b/><w:color w:val="${c.ink}"/><w:sz w:val="${s.body}"/><w:szCs w:val="${s.body}"/>`
        : `${fonts(t, 'heading')}<w:b/><w:color w:val="${c.accent}"/><w:sz w:val="${s.h1}"/><w:szCs w:val="${s.h1}"/>`,
    }),
    paraStyle({
      id: 'Heading2',
      name: 'heading 2',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 9,
      qFormat: true,
      pPr: `${keep}<w:spacing w:before="${pt(16)}" w:after="${pt(6)}" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="1"/>`,
      rPr: `${fonts(t, 'heading')}<w:b/><w:color w:val="${c.ink}"/><w:sz w:val="${s.h2}"/><w:szCs w:val="${s.h2}"/>`,
    }),
    paraStyle({
      id: 'Heading3',
      name: 'heading 3',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 9,
      qFormat: true,
      pPr: `${keep}<w:spacing w:before="${pt(12)}" w:after="${pt(4)}" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="2"/>`,
      rPr: `${fonts(t, 'heading')}<w:b/><w:color w:val="${c.ink}"/><w:sz w:val="${s.h3}"/><w:szCs w:val="${s.h3}"/>`,
    }),
    paraStyle({
      id: 'Heading4',
      name: 'heading 4',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 9,
      qFormat: true,
      pPr: `${keep}<w:spacing w:before="${pt(10)}" w:after="${pt(3)}" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="3"/>`,
      rPr: `${fonts(t, 'heading')}<w:b/><w:i/><w:color w:val="${c.muted}"/><w:sz w:val="${s.h4}"/><w:szCs w:val="${s.h4}"/>`,
    }),
    paraStyle({
      id: 'ListBullet',
      name: 'List Bullet',
      basedOn: 'BodyText',
      uiPriority: 36,
      qFormat: true,
      pPr: `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${NUM_BULLET}"/></w:numPr><w:spacing w:after="60"/><w:contextualSpacing/>`,
    }),
    paraStyle({
      id: 'ListBullet2',
      name: 'List Bullet 2',
      basedOn: 'BodyText',
      uiPriority: 36,
      qFormat: true,
      pPr: `<w:numPr><w:ilvl w:val="1"/><w:numId w:val="${NUM_BULLET}"/></w:numPr><w:spacing w:after="60"/><w:contextualSpacing/>`,
    }),
    paraStyle({
      id: 'ListNumber',
      name: 'List Number',
      basedOn: 'BodyText',
      uiPriority: 36,
      qFormat: true,
      pPr: `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${NUM_DECIMAL}"/></w:numPr><w:spacing w:after="60"/><w:contextualSpacing/>`,
    }),
    paraStyle({
      id: 'ListNumber2',
      name: 'List Number 2',
      basedOn: 'BodyText',
      uiPriority: 36,
      qFormat: true,
      pPr: `<w:numPr><w:ilvl w:val="1"/><w:numId w:val="${NUM_DECIMAL}"/></w:numPr><w:spacing w:after="60"/><w:contextualSpacing/>`,
    }),
    paraStyle({
      id: 'Quote',
      name: 'Quote',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 29,
      qFormat: true,
      pPr: `<w:pBdr><w:left w:val="single" w:sz="24" w:space="10" w:color="${c.accent}"/></w:pBdr><w:spacing w:before="160" w:after="160" w:line="288" w:lineRule="auto"/><w:ind w:left="360" w:right="360"/>`,
      rPr: `<w:i/><w:color w:val="${c.muted}"/>`,
    }),
    paraStyle({
      id: 'Callout',
      name: 'Callout',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 30,
      qFormat: true,
      pPr:
        `<w:keepLines/><w:pBdr><w:top w:val="single" w:sz="4" w:space="6" w:color="${c.accentSoft}"/>` +
        `<w:left w:val="single" w:sz="24" w:space="8" w:color="${c.accent}"/>` +
        `<w:bottom w:val="single" w:sz="4" w:space="6" w:color="${c.accentSoft}"/>` +
        `<w:right w:val="single" w:sz="4" w:space="8" w:color="${c.accentSoft}"/></w:pBdr>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${c.accentSoft}"/><w:spacing w:before="200" w:after="200" w:line="276" w:lineRule="auto"/><w:ind w:left="240" w:right="240"/>`,
    }),
    paraStyle({
      id: 'Caption',
      name: 'caption',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 35,
      qFormat: true,
      pPr: `<w:keepNext/><w:spacing w:before="120" w:after="80" w:line="240" w:lineRule="auto"/>`,
      rPr: `<w:b/><w:color w:val="${c.muted}"/><w:sz w:val="${s.small}"/><w:szCs w:val="${s.small}"/>`,
    }),
    paraStyle({
      id: 'Code',
      name: 'Code',
      basedOn: 'Normal',
      uiPriority: 40,
      qFormat: true,
      pPr: `<w:keepLines/><w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="${c.rule}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${c.wash}"/><w:spacing w:before="0" w:after="0" w:line="252" w:lineRule="auto"/><w:ind w:left="240"/>`,
      rPr: `${fonts(t, 'mono')}<w:sz w:val="${s.code}"/><w:szCs w:val="${s.code}"/>`,
    }),
    paraStyle({
      id: 'NoSpacing',
      name: 'No Spacing',
      basedOn: 'Normal',
      uiPriority: 1,
      qFormat: true,
      pPr: '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
    }),
    paraStyle({
      id: 'Header',
      name: 'header',
      basedOn: 'Normal',
      uiPriority: 99,
      pPr: `<w:pBdr>${rule(c.rule)}</w:pBdr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>`,
      rPr: `<w:color w:val="${c.muted}"/><w:sz w:val="${s.small}"/><w:szCs w:val="${s.small}"/>`,
    }),
    paraStyle({
      id: 'Footer',
      name: 'footer',
      basedOn: 'Normal',
      uiPriority: 99,
      pPr: `<w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>`,
      rPr: `<w:color w:val="${c.muted}"/><w:sz w:val="${s.small}"/><w:szCs w:val="${s.small}"/>`,
    }),
    paraStyle({
      id: 'TOCHeading',
      name: 'TOC Heading',
      basedOn: 'Heading1',
      next: 'BodyText',
      uiPriority: 39,
      qFormat: true,
      pPr: '<w:outlineLvl w:val="9"/>',
    }),
    paraStyle({
      id: 'TOC1',
      name: 'toc 1',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 39,
      pPr: `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs><w:spacing w:before="80" w:after="40"/>`,
      rPr: '<w:b/>',
    }),
    paraStyle({
      id: 'TOC2',
      name: 'toc 2',
      basedOn: 'Normal',
      next: 'BodyText',
      uiPriority: 39,
      pPr: `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs><w:spacing w:after="40"/><w:ind w:left="240"/>`,
    }),
    paraStyle({
      id: 'FootnoteText',
      name: 'footnote text',
      basedOn: 'Normal',
      uiPriority: 99,
      pPr: '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
      rPr: `<w:sz w:val="${s.small}"/><w:szCs w:val="${s.small}"/>`,
    }),
    paraStyle({
      id: 'EndnoteText',
      name: 'endnote text',
      basedOn: 'Normal',
      uiPriority: 99,
      pPr: '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
      rPr: `<w:sz w:val="${s.small}"/><w:szCs w:val="${s.small}"/>`,
    }),
    paraStyle({
      id: 'ListParagraph',
      name: 'List Paragraph',
      basedOn: 'Normal',
      uiPriority: 34,
      qFormat: true,
      pPr: '<w:ind w:left="720"/><w:contextualSpacing/>',
    }),
    // Cover and form furniture.
    paraStyle({
      id: 'CoverLabel',
      name: 'Cover Label',
      basedOn: 'Normal',
      uiPriority: 41,
      qFormat: true,
      pPr: '<w:keepNext/><w:spacing w:before="0" w:after="20" w:line="240" w:lineRule="auto"/>',
      rPr: `<w:b/><w:caps/><w:color w:val="${c.muted}"/><w:spacing w:val="20"/><w:sz w:val="${s.small}"/><w:szCs w:val="${s.small}"/>`,
    }),
    paraStyle({
      id: 'CoverValue',
      name: 'Cover Value',
      basedOn: 'Normal',
      uiPriority: 41,
      qFormat: true,
      pPr: '<w:spacing w:before="0" w:after="200" w:line="240" w:lineRule="auto"/>',
      rPr: `<w:sz w:val="${s.body + 2}"/><w:szCs w:val="${s.body + 2}"/>`,
    }),
    paraStyle({
      id: 'Reference',
      name: 'Reference',
      basedOn: 'BodyText',
      uiPriority: 44,
      qFormat: true,
      pPr: '<w:spacing w:after="80"/><w:ind w:left="480" w:hanging="480"/>',
    }),
    paraStyle({
      id: 'TableText',
      name: 'Table Text',
      basedOn: 'Normal',
      uiPriority: 42,
      qFormat: true,
      pPr: '<w:spacing w:before="40" w:after="40" w:line="252" w:lineRule="auto"/>',
      rPr: `<w:sz w:val="${s.body - 1}"/><w:szCs w:val="${s.body - 1}"/>`,
    }),
  ];

  const characters: string[] = [
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>',
    charStyle('Strong', 'Strong', '<w:b/><w:bCs/>', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="22"/><w:qFormat/>'),
    charStyle('Emphasis', 'Emphasis', '<w:i/><w:iCs/>', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="20"/><w:qFormat/>'),
    charStyle('Hyperlink', 'Hyperlink', `<w:color w:val="${c.accent}"/><w:u w:val="single"/>`, '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
    charStyle('CodeChar', 'Code Char', `${fonts(t, 'mono')}<w:sz w:val="${s.code}"/><w:szCs w:val="${s.code}"/>`, '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="40"/>'),
    charStyle('Placeholder', 'Placeholder', `<w:i/><w:color w:val="${c.placeholder}"/>`, '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="43"/><w:qFormat/>'),
    charStyle('EndnoteReference', 'endnote reference', '<w:vertAlign w:val="superscript"/>', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
    charStyle('FootnoteReference', 'footnote reference', '<w:vertAlign w:val="superscript"/>', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
  ];

  const hairline = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${c.rule}"/>`;
  const tables: string[] = [
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>' +
      '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>',
    // The style the engine names on every table it writes: single spacing in cells, thin grid lines.
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/>' +
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
      '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>',
    // Clean: hairline horizontal rules, tinted header row.
    '<w:style w:type="table" w:styleId="AstraTable"><w:name w:val="Astra Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="59"/><w:qFormat/>' +
      `<w:pPr><w:spacing w:before="40" w:after="40" w:line="252" w:lineRule="auto"/></w:pPr>` +
      `<w:tblPr><w:tblBorders>${hairline('top')}${hairline('bottom')}${hairline('insideH')}</w:tblBorders>` +
      '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="${c.ink}"/></w:rPr><w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="12" w:space="0" w:color="${c.accent}"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="${c.accentSoft}"/></w:tcPr></w:tblStylePr>` +
      '</w:style>',
    // Banded: the clean table plus alternate row tint.
    '<w:style w:type="table" w:styleId="AstraTableBanded"><w:name w:val="Astra Table Banded"/><w:basedOn w:val="AstraTable"/><w:uiPriority w:val="60"/><w:qFormat/>' +
      `<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>` +
      `<w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${c.wash}"/></w:tcPr></w:tblStylePr>` +
      '</w:style>',
    // Minimal: no vertical or inner rules, only a rule under the header. For invoices and form blocks.
    '<w:style w:type="table" w:styleId="AstraTableMinimal"><w:name w:val="Astra Table Minimal"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="61"/><w:qFormat/>' +
      '<w:tblPr><w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="${c.muted}"/></w:rPr><w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="8" w:space="0" w:color="${c.ink}"/></w:tcBorders></w:tcPr></w:tblStylePr>` +
      '</w:style>',
    // Borderless layout grid for memo and letter header blocks (no tblHeader: not a data table).
    '<w:style w:type="table" w:styleId="AstraTableGrid"><w:name w:val="Astra Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="62"/>' +
      '<w:tblPr><w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>',
  ];

  return (
    XML_DECL +
    `<w:styles xmlns:w="${W_NS}">` +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    `${fonts(t, 'body')}<w:color w:val="${c.ink}"/><w:sz w:val="${s.body}"/><w:szCs w:val="${s.body}"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>` +
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl/><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    paragraphs.slice(0, 1).join('') +
    characters.slice(0, 1).join('') +
    tables.slice(0, 1).join('') +
    paragraphs.slice(1).join('') +
    characters.slice(1).join('') +
    tables.slice(1).join('') +
    '</w:styles>'
  );
}

export function numberingXml(t: Theme): string {
  const bullets = ['•', '–', '•'];
  const level = (i: number, fmt: string, text: string, left: number) =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr><w:tabs><w:tab w:val="num" w:pos="${left}"/></w:tabs><w:ind w:left="${left}" w:hanging="300"/></w:pPr>` +
    (fmt === 'bullet' ? `<w:rPr><w:rFonts w:ascii="${t.fonts.body}" w:hAnsi="${t.fonts.body}" w:hint="default"/><w:color w:val="${t.colors.accent}"/></w:rPr>` : `<w:rPr><w:color w:val="${t.colors.accent}"/></w:rPr>`) +
    '</w:lvl>';
  const bulletLevels = [0, 1, 2].map((i) => level(i, 'bullet', bullets[i], 480 + i * 360)).join('');
  const decimalLevels = [
    level(0, 'decimal', '%1.', 480),
    level(1, 'lowerLetter', '%2.', 840),
    level(2, 'lowerRoman', '%3.', 1200),
  ].join('');
  return (
    XML_DECL +
    `<w:numbering xmlns:w="${W_NS}">` +
    `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLevels}</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${decimalLevels}</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/>` +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="540" w:hanging="540"/></w:pPr></w:lvl>` +
    `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="660" w:hanging="660"/></w:pPr></w:lvl>` +
    `<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="780" w:hanging="780"/></w:pPr></w:lvl>` +
    `</w:abstractNum>` +
    `<w:num w:numId="${NUM_BULLET}"><w:abstractNumId w:val="0"/></w:num>` +
    `<w:num w:numId="${NUM_DECIMAL}"><w:abstractNumId w:val="1"/></w:num>` +
    `<w:num w:numId="${NUM_HEADINGS}"><w:abstractNumId w:val="2"/></w:num>` +
    '</w:numbering>'
  );
}

export function settingsXml(): string {
  return (
    XML_DECL +
    `<w:settings xmlns:w="${W_NS}">` +
    '<w:zoom w:percent="100"/>' +
    '<w:defaultTabStop w:val="720"/>' +
    '<w:characterSpacingControl w:val="doNotCompress"/>' +
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
    '<w:themeFontLang w:val="en-US"/>' +
    '<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" ' +
    'w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>' +
    '<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>' +
    '</w:settings>'
  );
}

export function fontTableXml(t: Theme): string {
  const f = (name: string, family: string, pitch: string) =>
    `<w:font w:name="${name}"><w:charset w:val="00"/><w:family w:val="${family}"/><w:pitch w:val="${pitch}"/></w:font>`;
  const seen = new Set<string>();
  const out: string[] = [];
  for (const [name, family, pitch] of [
    [t.fonts.body, 'auto', 'variable'],
    [t.fonts.heading, 'auto', 'variable'],
    [t.fonts.mono, 'modern', 'fixed'],
  ] as const) {
    if (seen.has(name)) continue;
    seen.add(name);
    out.push(f(name, family, pitch));
  }
  return XML_DECL + `<w:fonts xmlns:w="${W_NS}">${out.join('')}</w:fonts>`;
}

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export function themeXml(t: Theme): string {
  const c = t.colors;
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const line = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  const fill = `<a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst>`;
  return (
    XML_DECL +
    `<a:theme xmlns:a="${A_NS}" name="${esc(t.name)}"><a:themeElements>` +
    `<a:clrScheme name="${esc(t.name)}">` +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    `<a:dk2><a:srgbClr val="${c.ink}"/></a:dk2><a:lt2><a:srgbClr val="${c.wash}"/></a:lt2>` +
    `<a:accent1><a:srgbClr val="${c.accent}"/></a:accent1><a:accent2><a:srgbClr val="${c.muted}"/></a:accent2>` +
    `<a:accent3><a:srgbClr val="${c.rule}"/></a:accent3><a:accent4><a:srgbClr val="${c.accentSoft}"/></a:accent4>` +
    `<a:accent5><a:srgbClr val="${c.placeholder}"/></a:accent5><a:accent6><a:srgbClr val="${c.ink}"/></a:accent6>` +
    `<a:hlink><a:srgbClr val="${c.accent}"/></a:hlink><a:folHlink><a:srgbClr val="${c.muted}"/></a:folHlink></a:clrScheme>` +
    `<a:fontScheme name="${esc(t.name)}"><a:majorFont><a:latin typeface="${t.fonts.heading}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="${t.fonts.body}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="${esc(t.name)}">${fill}<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    `<a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme>` +
    '</a:themeElements></a:theme>'
  );
}

/** Relative luminance and WCAG contrast ratio of two RRGGBB colours. */
export function contrast(a: string, b: string): number {
  const lum = (hex: string): number => {
    const ch = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((v) => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)));
    return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
  };
  const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
}

/** Text and background pairs that must reach WCAG AA (4.5:1). */
export function contrastChecks(t: Theme): Array<{ pair: string; ratio: number; ok: boolean }> {
  const c = t.colors;
  const pairs: Array<[string, string, string]> = [
    ['ink on white', c.ink, 'FFFFFF'],
    ['muted on white', c.muted, 'FFFFFF'],
    ['accent on white (Heading 1, links)', c.accent, 'FFFFFF'],
    ['placeholder on white', c.placeholder, 'FFFFFF'],
    ['ink on accentSoft (callout, table header)', c.ink, c.accentSoft],
    ['ink on wash (code, banded rows)', c.ink, c.wash],
    ['muted on wash', c.muted, c.wash],
  ];
  return pairs.map(([pair, fg, bg]) => {
    const ratio = Math.round(contrast(fg, bg) * 100) / 100;
    return { pair, ratio, ok: ratio >= 4.5 };
  });
}
