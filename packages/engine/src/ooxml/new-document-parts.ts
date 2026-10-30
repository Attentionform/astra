/**
 * The package parts a brand-new document carries besides document.xml and
 * styles.xml: core and extended properties, settings, a font table and a
 * theme. Without them Word opens the file in compatibility mode, shows no
 * title or author, and has no theme fonts or colours to resolve against.
 */

import { stripIllegalXmlChars } from './xml-text';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** `[Content_Types].xml` entries for the parts below. */
export const NEW_DOCUMENT_CONTENT_TYPE_OVERRIDES = `
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`;

/** `_rels/.rels` entries (rId1 is the main document). */
export const NEW_DOCUMENT_ROOT_RELS = `
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="${REL_BASE}/extended-properties" Target="docProps/app.xml"/>`;

/** `word/_rels/document.xml.rels` entries (rId1 is the styles part). */
export const NEW_DOCUMENT_DOCUMENT_RELS = `
  <Relationship Id="rId2" Type="${REL_BASE}/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="${REL_BASE}/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId4" Type="${REL_BASE}/theme" Target="theme/theme1.xml"/>`;

function escapeXml(value: string): string {
  return stripIllegalXmlChars(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** W3CDTF without milliseconds, the form Word writes. */
function w3cdtf(at: Date): string {
  return at.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function coreProperties(title: string, author: string, at: Date): string {
  const stamp = w3cdtf(at);
  return (
    XML_DECL +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${escapeXml(title)}</dc:title>` +
    `<dc:creator>${escapeXml(author)}</dc:creator>` +
    `<cp:lastModifiedBy>${escapeXml(author)}</cp:lastModifiedBy>` +
    '<cp:revision>1</cp:revision>' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>` +
    '</cp:coreProperties>'
  );
}

const APP_PROPERTIES =
  XML_DECL +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<Application>ASTra</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
  '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>' +
  '</Properties>';

// compatibilityMode 15 is what stops Word from labelling the file "Compatibility Mode".
const SETTINGS =
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
  '</w:settings>';

const FONT_TABLE =
  XML_DECL +
  `<w:fonts xmlns:w="${W_NS}">` +
  '<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>' +
  '<w:font w:name="Consolas"><w:panose1 w:val="020B0609020204030204"/><w:charset w:val="00"/><w:family w:val="modern"/><w:pitch w:val="fixed"/></w:font>' +
  '</w:fonts>';

const SOLID = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
const line = (width: number): string =>
  `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr">${SOLID}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
const EFFECT = '<a:effectStyle><a:effectLst/></a:effectStyle>';
const FONT_FACES = '<a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>';

// Colours follow the built-in Default Corporate palette; fonts keep the
// Calibri the styles already use, so adding the theme changes no layout.
const THEME =
  XML_DECL +
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ASTra">' +
  '<a:themeElements>' +
  '<a:clrScheme name="ASTra">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="1E293B"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="1E3A8A"/></a:accent1><a:accent2><a:srgbClr val="0284C7"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="10B981"/></a:accent3><a:accent4><a:srgbClr val="F59E0B"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="6366F1"/></a:accent5><a:accent6><a:srgbClr val="EC4899"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme>' +
  `<a:fontScheme name="ASTra"><a:majorFont>${FONT_FACES}</a:majorFont><a:minorFont>${FONT_FACES}</a:minorFont></a:fontScheme>` +
  '<a:fmtScheme name="ASTra">' +
  `<a:fillStyleLst>${SOLID}${SOLID}${SOLID}</a:fillStyleLst>` +
  `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
  `<a:effectStyleLst>${EFFECT}${EFFECT}${EFFECT}</a:effectStyleLst>` +
  `<a:bgFillStyleLst>${SOLID}${SOLID}${SOLID}</a:bgFillStyleLst>` +
  '</a:fmtScheme>' +
  '</a:themeElements>' +
  '</a:theme>';

/**
 * Part path to XML text for a new document. `at` stamps both created and
 * modified (the caller passes the pinned stamp in reproducible mode).
 */
export function newDocumentParts(opts: { title: string; author: string; at: Date }): Record<string, string> {
  return {
    'docProps/core.xml': coreProperties(opts.title, opts.author, opts.at),
    'docProps/app.xml': APP_PROPERTIES,
    'word/settings.xml': SETTINGS,
    'word/fontTable.xml': FONT_TABLE,
    'word/theme/theme1.xml': THEME,
  };
}

/** Document title for a template name: "report" gives "Report", "api-spec" gives "Api Spec". */
export function newDocumentTitle(templateType: string): string {
  const name = (templateType || 'blank').toLowerCase().trim() || 'blank';
  if (name === 'blank') return 'Blank Document';
  return name
    .split(/[^a-z0-9]+/)
    .filter(Boolean)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}
