import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { AdmZip, createBlankDocxBuffer, REPAIR_RULES, repairDocxBuffer, type RepairRule, type ValidationResult } from '@attentionform/astra-engine';
import { parsePartStrict } from '../../packages/engine/src/ooxml/normalize';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const rule = (id: string): RepairRule => REPAIR_RULES.find((r) => r.id === id) as RepairRule;

/** Runs one rule on one part. Returns the new XML and the log. */
function run(id: string, part: string, xml: string, hasBidi = false): { xml: string; log: string[] } {
  const doc = parsePartStrict(part, xml);
  if (!doc) throw new Error('no document');
  const log = rule(id).apply(doc, part, { hasBidi });
  return { xml: log.length > 0 ? new XMLSerializer().serializeToString(doc) : xml, log };
}

/** Applies the rule twice: the second run must change nothing. */
function runTwice(id: string, part: string, xml: string, hasBidi = false) {
  const first = run(id, part, xml, hasBidi);
  const second = run(id, part, first.xml, hasBidi);
  expect(second.log, `${id} is idempotent`).toEqual([]);
  return first;
}

const doc = (body: string) => `<w:document ${W}><w:body>${body}</w:body></w:document>`;

describe('repair rules', () => {
  it('settings-order puts w:zoom before w:compat', () => {
    const r = runTwice('settings-order', 'word/settings.xml', `<w:settings ${W}><w:compat/><w:zoom w:percent="100"/></w:settings>`);
    expect(r.xml.indexOf('<w:zoom')).toBeLessThan(r.xml.indexOf('<w:compat'));
    expect(r.log).toHaveLength(1);
  });

  it('style-order puts w:name before w:basedOn and w:rPr after w:pPr', () => {
    const r = runTwice('style-order', 'word/styles.xml', `<w:styles ${W}><w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="N"/><w:rPr><w:sz w:val="20"/></w:rPr><w:name w:val="A"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style></w:styles>`);
    expect(r.xml.indexOf('<w:name')).toBeLessThan(r.xml.indexOf('<w:basedOn'));
    expect(r.xml.indexOf('<w:pPr')).toBeLessThan(r.xml.indexOf('<w:rPr'));
  });

  it('style-order sorts a table property container inside a style', () => {
    const r = runTwice('style-order', 'word/styles.xml', `<w:styles ${W}><w:style w:type="table" w:styleId="T"><w:name w:val="T"/><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblStyleRowBandSize w:val="1"/></w:tblPr></w:style></w:styles>`);
    expect(r.xml.indexOf('tblStyleRowBandSize')).toBeLessThan(r.xml.indexOf('<w:tblW'));
  });

  it('sectpr-order puts w:pgSz before w:pgMar and keeps header and footer references where they are', () => {
    const sect = '<w:sectPr><w:footerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><w:headerReference w:type="default" r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>';
    const r = runTwice('sectpr-order', 'word/document.xml', doc(`${sect}<w:pgMar w:top="1"/><w:pgSz w:w="1" w:h="1"/></w:sectPr>`));
    expect(r.xml.indexOf('<w:pgSz')).toBeLessThan(r.xml.indexOf('<w:pgMar'));
    expect(r.xml.indexOf('footerReference')).toBeLessThan(r.xml.indexOf('headerReference'));
  });

  it('table-children-order puts w:tblPr before w:tblGrid', () => {
    const r = runTwice('table-children-order', 'word/document.xml', doc('<w:tbl><w:tblGrid><w:gridCol w:w="1"/></w:tblGrid><w:tblPr/><w:tr/></w:tbl>'));
    expect(r.xml.indexOf('<w:tblPr')).toBeLessThan(r.xml.indexOf('<w:tblGrid'));
  });

  it('run-properties-first moves w:rPr to the start of the run and keeps the text', () => {
    const r = runTwice('run-properties-first', 'word/document.xml', doc('<w:p><w:r><w:t>hi</w:t><w:rPr><w:b/></w:rPr></w:r></w:p>'));
    expect(r.xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t>hi</w:t></w:r>');
  });

  it('integer-attribute rounds page size and cell margins only', () => {
    const r = runTwice('integer-attribute', 'word/document.xml', doc('<w:sectPr><w:pgSz w:w="11905.5" w:h="16837.795"/></w:sectPr><w:tbl><w:tblPr><w:tblCellMar><w:left w:w="70.0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:tbl><w:p><w:r><w:t w:foo="1.5">x</w:t></w:r></w:p>'));
    expect(r.xml).toContain('w:w="11906"');
    expect(r.xml).toContain('w:h="16838"');
    expect(r.xml).toContain('w:w="70"');
    expect(r.xml).toContain('w:foo="1.5"');
  });

  it('on-off-attribute writes 1 and 0 as on and off for the listed elements only', () => {
    const r = runTwice('on-off-attribute', 'word/styles.xml', `<w:styles ${W}><w:style w:type="paragraph" w:styleId="A"><w:qFormat w:val="1"/><w:semiHidden w:val="0"/><w:rPr><w:b w:val="0"/></w:rPr></w:style></w:styles>`);
    expect(r.xml).toContain('<w:qFormat w:val="on"/>');
    expect(r.xml).toContain('<w:semiHidden w:val="off"/>');
    expect(r.xml).toContain('<w:b w:val="0"/>');
  });

  it('grid-column-type removes the undeclared attribute', () => {
    const r = runTwice('grid-column-type', 'word/document.xml', doc('<w:tbl><w:tblGrid><w:gridCol w:w="5" w:type="dxa"/></w:tblGrid></w:tbl>'));
    expect(r.xml).toContain('<w:gridCol w:w="5"/>');
  });

  it('font-family-case writes Auto as auto', () => {
    const r = runTwice('font-family-case', 'word/fontTable.xml', `<w:fonts ${W}><w:font w:name="X"><w:family w:val="Auto"/></w:font></w:fonts>`);
    expect(r.xml).toContain('w:val="auto"');
  });

  it('alignment-start-end maps start and end when no part is right to left, and does nothing otherwise', () => {
    const xml = doc('<w:p><w:pPr><w:jc w:val="start"/></w:pPr></w:p><w:p><w:pPr><w:jc w:val="end"/></w:pPr></w:p>');
    const r = runTwice('alignment-start-end', 'word/document.xml', xml);
    expect(r.xml).toContain('w:val="left"');
    expect(r.xml).toContain('w:val="right"');
    expect(run('alignment-start-end', 'word/document.xml', xml, true).log).toEqual([]);
  });

  it('smarttag-unwrap keeps the runs and drops the tag and its properties', () => {
    const r = runTwice('smarttag-unwrap', 'word/document.xml', doc('<w:p><w:smartTag w:uri="u" w:element="e"><w:smartTagPr><w:attr w:name="a" w:val="b"/></w:smartTagPr><w:r><w:t>kept</w:t></w:r></w:smartTag></w:p>'));
    expect(r.xml).toContain('<w:p><w:r><w:t>kept</w:t></w:r></w:p>');
    expect(r.xml).not.toContain('smartTag');
  });

  it('shapetype-duplicate removes an identical repeat and keeps a different one with the same id', () => {
    const v = 'xmlns:v="urn:schemas-microsoft-com:vml"';
    const same = doc(`<w:p><w:pict ${v}><v:shapetype id="t1" coordsize="1,1"/><v:shapetype id="t1" coordsize="1,1"/><v:shapetype id="t1" coordsize="2,2"/></w:pict></w:p>`);
    const r = runTwice('shapetype-duplicate', 'word/document.xml', same);
    expect((r.xml.match(/<v:shapetype/g) ?? []).length).toBe(2);
  });
});

describe('rules added after the first measurement', () => {
  it('settings-legacy-smarttags removes w:smartTagType and w:doNotEmbedSmartTags and keeps the rest', () => {
    const r = runTwice('settings-legacy-smarttags', 'word/settings.xml', `<w:settings ${W}><w:zoom w:percent="100"/><w:smartTagType w:namespaceuri="urn:x" w:name="y"/><w:doNotEmbedSmartTags/><w:decimalSymbol w:val="."/></w:settings>`);
    expect(r.xml).not.toContain('smartTag');
    expect(r.xml).not.toContain('doNotEmbedSmartTags');
    expect(r.xml).toContain('<w:zoom w:percent="100"/>');
    expect(r.xml).toContain('<w:decimalSymbol w:val="."/>');
  });

  it('style-table-defaults removes default table width, layout and bidiVisual from a style, not other values', () => {
    const r = runTwice('style-table-defaults', 'word/styles.xml', `<w:styles ${W}><w:style w:type="table" w:styleId="A"><w:name w:val="A"/><w:tblPr><w:bidiVisual w:val="off"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single"/></w:tblBorders><w:tblLayout w:type="autofit"/></w:tblPr></w:style><w:style w:type="table" w:styleId="B"><w:name w:val="B"/><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/></w:tblPr></w:style></w:styles>`);
    expect(r.xml).not.toContain('w:bidiVisual');
    expect(r.xml).toContain('<w:tblW w:w="5000" w:type="pct"/>');
    expect(r.xml).toContain('<w:tblLayout w:type="fixed"/>');
    expect(r.xml).toContain('<w:tblBorders>');
    expect((r.xml.match(/<w:tblW /g) ?? []).length).toBe(1);
  });

  it('style-table-defaults leaves a table in the document alone', () => {
    const xml = doc('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="autofit"/></w:tblPr></w:tbl>');
    expect(run('style-table-defaults', 'word/styles.xml', xml).log).toEqual([]);
  });

  it('paragraph-bare-run-content puts a bare w:br, w:tab and w:drawing into runs and keeps the order and the text', () => {
    const r = runTwice('paragraph-bare-run-content', 'word/document.xml', doc('<w:p><w:r><w:t>a</w:t></w:r><w:br/><w:tab/><w:r><w:t>b</w:t></w:r><w:drawing/></w:p>'));
    expect(r.xml).toContain('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:br/></w:r><w:r><w:tab/></w:r><w:r><w:t>b</w:t></w:r><w:r><w:drawing/></w:r></w:p>');
  });

  it('paragraph-bare-run-content leaves a w:br inside a run alone', () => {
    expect(run('paragraph-bare-run-content', 'word/document.xml', doc('<w:p><w:r><w:br/></w:r></w:p>')).log).toEqual([]);
  });

  const W10 = 'xmlns:w10="urn:schemas-microsoft-com:office:word"';
  const withW10 = (inner: string) => `<w:document ${W} ${W10}><w:body><w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p></w:body></w:document>`;

  it('vml-wrap-inline removes a w10:wrap of type inline and keeps the other wrap types', () => {
    const r = runTwice('vml-wrap-inline', 'word/document.xml', withW10('<w10:wrap type="inline"/><w10:wrap type="square"/>'));
    expect(r.log).toHaveLength(1);
    expect(r.xml).not.toContain('type="inline"');
    expect(r.xml).toContain('<w10:wrap type="square"/>');
  });

  it('vml-wrap-inline leaves a w10:wrap with a valid type alone', () => {
    expect(run('vml-wrap-inline', 'word/document.xml', withW10('<w10:wrap type="topAndBottom"/>')).log).toEqual([]);
  });

  it('color-default-value gives a w:color without w:val the value auto and keeps the other attributes', () => {
    const r = runTwice('color-default-value', 'word/document.xml', doc('<w:p><w:r><w:rPr><w:color w:themeColor="accent1"/></w:rPr><w:t>x</w:t></w:r></w:p>'));
    expect(r.log).toHaveLength(1);
    expect(r.xml).toContain('w:val="auto"');
    expect(r.xml).toContain('w:themeColor="accent1"');
  });

  it('color-default-value leaves a w:color that has a w:val alone', () => {
    expect(run('color-default-value', 'word/document.xml', doc('<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>x</w:t></w:r></w:p>')).log).toEqual([]);
  });
});

describe('cell-valign-duplicate, shading-empty-fill and the order of the Word extension elements in settings', () => {
  const cell = (props: string) => doc(`<w:tbl><w:tr><w:tc><w:tcPr>${props}</w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);

  it('cell-valign-duplicate keeps the last w:vAlign and removes the others', () => {
    const r = runTwice('cell-valign-duplicate', 'word/document.xml', cell('<w:vAlign w:val="top"/><w:tcMar><w:left w:w="1" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/>'));
    expect(r.xml.match(/<w:vAlign/g)).toHaveLength(1);
    expect(r.xml).toContain('<w:vAlign w:val="center"/>');
    expect(r.xml).not.toContain('w:val="top"');
    expect(r.xml).toContain('<w:tcMar>');
    expect(r.log).toEqual(['Removed 1 duplicate w:vAlign from /w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:tcPr[1], kept w:val="center"']);
  });

  it('cell-valign-duplicate handles three, in a style as well, and leaves one alone', () => {
    const three = runTwice('cell-valign-duplicate', 'word/document.xml', cell('<w:vAlign w:val="top"/><w:vAlign w:val="bottom"/><w:vAlign w:val="center"/>'));
    expect(three.xml.match(/<w:vAlign/g)).toHaveLength(1);
    expect(three.xml).toContain('w:val="center"');
    const style = runTwice('cell-valign-duplicate', 'word/styles.xml', `<w:styles ${W}><w:style w:type="table" w:styleId="T"><w:tblStylePr w:type="firstRow"><w:tcPr><w:vAlign w:val="top"/><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr></w:style></w:styles>`);
    expect(style.xml.match(/<w:vAlign/g)).toHaveLength(1);
    expect(run('cell-valign-duplicate', 'word/document.xml', cell('<w:vAlign w:val="center"/>')).log).toEqual([]);
  });

  it('cell-valign-duplicate keeps the text of the cell', () => {
    const r = run('cell-valign-duplicate', 'word/document.xml', cell('<w:vAlign w:val="top"/><w:vAlign w:val="center"/>'));
    expect(r.xml).toContain('<w:t>cell</w:t>');
  });

  it('shading-empty-fill removes an empty w:fill and keeps the other attributes', () => {
    const r = runTwice('shading-empty-fill', 'word/document.xml', cell('<w:shd w:val="clear" w:color="auto" w:fill=""/>'));
    expect(r.xml).toContain('<w:shd w:val="clear" w:color="auto"/>');
    expect(r.xml).not.toContain('w:fill');
    expect(r.log).toHaveLength(1);
  });

  it('shading-empty-fill also reads a fill of white space, and works in styles and paragraph properties', () => {
    expect(run('shading-empty-fill', 'word/document.xml', doc('<w:p><w:pPr><w:shd w:val="clear" w:fill="  "/></w:pPr></w:p>')).xml).not.toContain('w:fill');
    const style = run('shading-empty-fill', 'word/styles.xml', `<w:styles ${W}><w:style w:type="paragraph" w:styleId="S"><w:pPr><w:shd w:val="clear" w:fill=""/></w:pPr></w:style></w:styles>`);
    expect(style.xml).not.toContain('w:fill');
  });

  it('shading-empty-fill leaves a real fill, auto, and a shd without a fill alone', () => {
    for (const shd of ['<w:shd w:val="clear" w:fill="FF0000"/>', '<w:shd w:val="clear" w:fill="auto"/>', '<w:shd w:val="clear"/>']) {
      expect(run('shading-empty-fill', 'word/document.xml', cell(shd)).log).toEqual([]);
    }
  });

  it('settings-order keeps w:listSeparator before the w14 and w15 elements, and puts a late w:updateFields where it belongs', () => {
    // The order of a document from the corpus: Word's elements follow w:listSeparator, and w:updateFields is last.
    const settings = `<w:settings ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:zoom w:percent="100"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat/><w:themeFontLang w:val="en-US"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/><w14:docId w14:val="1"/><w15:chartTrackingRefBased/><w15:docId w15:val="{A}"/><w:updateFields w:val="true"/></w:settings>`;
    const r = runTwice('settings-order', 'word/settings.xml', settings);
    const tags = Array.from(r.xml.matchAll(/<((?:w|w14|w15):[A-Za-z]+)[ />]/g), (m) => m[1]).filter((t) => t !== 'w:settings');
    expect(tags).toEqual(['w:zoom', 'w:characterSpacingControl', 'w:updateFields', 'w:compat', 'w:themeFontLang', 'w:decimalSymbol', 'w:listSeparator', 'w14:docId', 'w15:chartTrackingRefBased', 'w15:docId']);
  });

  it('settings-order puts the w14 and w15 elements after w:listSeparator when they come before it', () => {
    const settings = `<w:settings ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:docId w14:val="1"/><w14:defaultImageDpi w14:val="220"/><w14:discardImageEditingData w14:val="1"/><w:listSeparator w:val=","/><w:decimalSymbol w:val="."/></w:settings>`;
    const r = runTwice('settings-order', 'word/settings.xml', settings);
    const tags = Array.from(r.xml.matchAll(/<((?:w|w14):[A-Za-z]+)[ />]/g), (m) => m[1]).filter((t) => t !== 'w:settings');
    expect(tags).toEqual(['w:decimalSymbol', 'w:listSeparator', 'w14:docId', 'w14:discardImageEditingData', 'w14:defaultImageDpi']);
  });

  it('paragraph-properties-first moves w:pPr to the start of the paragraph and keeps text', () => {
    const r = runTwice('paragraph-properties-first', 'word/document.xml', doc('<w:p><w:r><w:t>hello</w:t></w:r><w:pPr><w:jc w:val="center"/></w:pPr></w:p>'));
    expect(r.xml).toContain('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>hello</w:t></w:r></w:p>');
    expect(r.log).toHaveLength(1);
  });

  it('paragraph-properties-first keeps the last of two w:pPr, as Word does, and the text', () => {
    const body = '<w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="25"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r>';
    const r = runTwice('paragraph-properties-first', 'word/document.xml', doc(`<w:p>${body}</w:p>`));
    expect(r.xml).toContain('<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="25"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>');
    expect(r.xml).not.toContain('BodyText');
    expect(r.log).toHaveLength(1);
    expect(r.log[0]).toContain('Word uses the last one');
  });

  it('paragraph-properties-first moves the last of two w:pPr to the start when runs come first', () => {
    const r = runTwice('paragraph-properties-first', 'word/document.xml', doc('<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>two</w:t></w:r><w:pPr><w:jc w:val="right"/></w:pPr></w:p>'));
    expect(r.xml).toContain('<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>');
    expect(r.log).toHaveLength(2);
  });

  it('paragraph-properties-first leaves a paragraph with three w:pPr elements alone', () => {
    const xml = doc('<w:p><w:r><w:t>three</w:t></w:r><w:pPr><w:jc w:val="left"/></w:pPr><w:pPr><w:jc w:val="right"/></w:pPr><w:pPr><w:jc w:val="center"/></w:pPr></w:p>');
    expect(run('paragraph-properties-first', 'word/document.xml', xml).log).toEqual([]);
  });

  it('cell-margins-order puts w:tcMar children in schema order', () => {
    const r = runTwice('cell-margins-order', 'word/document.xml', cell('<w:tcMar><w:bottom w:w="120" w:type="dxa"/><w:top w:w="100" w:type="dxa"/><w:right w:w="80" w:type="dxa"/><w:left w:w="60" w:type="dxa"/></w:tcMar>'));
    expect(r.xml.indexOf('<w:top')).toBeLessThan(r.xml.indexOf('<w:left'));
    expect(r.xml.indexOf('<w:left')).toBeLessThan(r.xml.indexOf('<w:bottom'));
    expect(r.xml.indexOf('<w:bottom')).toBeLessThan(r.xml.indexOf('<w:right'));
    expect(r.log).toHaveLength(1);
  });

  it('cell-margins-order handles start and end attributes and leaves ordered margins alone', () => {
    const r = runTwice('cell-margins-order', 'word/document.xml', cell('<w:tcMar><w:end w:w="80" w:type="dxa"/><w:start w:w="60" w:type="dxa"/><w:top w:w="100" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/></w:tcMar>'));
    expect(r.xml.indexOf('<w:top')).toBeLessThan(r.xml.indexOf('<w:start'));
    expect(r.xml.indexOf('<w:start')).toBeLessThan(r.xml.indexOf('<w:bottom'));
    expect(r.xml.indexOf('<w:bottom')).toBeLessThan(r.xml.indexOf('<w:end'));
    expect(run('cell-margins-order', 'word/document.xml', cell('<w:tcMar><w:top w:w="10" w:type="dxa"/><w:left w:w="10" w:type="dxa"/></w:tcMar>')).log).toEqual([]);
  });

  it('all three repairs in one package keep the visible text and leave the package valid', async () => {
    const input = packageWith({
      'word/document.xml': doc(
        '<w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill=""/><w:vAlign w:val="top"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>',
      ),
    });
    const out = await repairDocxBuffer(input);
    const part = partOf(out.buffer, 'word/document.xml');
    expect(out.rolledBack).toEqual([]);
    expect(part).toContain('cell text');
    expect(part.match(/<w:vAlign/g)).toHaveLength(1);
    expect(part).not.toContain('w:fill=""');
    expect(out.remaining).toBe(0);
  });
});

describe('every rule leaves a valid part byte for byte', () => {
  const valid = doc('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ok</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1" w:right="1" w:bottom="1" w:left="1"/></w:sectPr>');
  for (const r of REPAIR_RULES) {
    it(r.id, () => {
      const part = ['word/settings.xml', 'word/styles.xml', 'word/fontTable.xml', 'word/document.xml'].find((n) => r.appliesTo(n)) as string;
      expect(run(r.id, part, valid).log).toEqual([]);
    });
  }
});

// ── runner ──────────────────────────────────────────────────────────────

const okResult = { file: 'x', ok: true, errors: [] } as unknown as ValidationResult;
const noValidator = async () => okResult;

function packageWith(parts: Record<string, string>): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  for (const [name, xml] of Object.entries(parts)) {
    if (zip.getEntry(name)) zip.updateFile(name, Buffer.from(xml, 'utf8'));
    else zip.addFile(name, Buffer.from(xml, 'utf8'));
  }
  return zip.toBuffer();
}
const partOf = (buf: Buffer, name: string) => new AdmZip(buf).getEntry(name)?.getData().toString('utf8') ?? '';

describe('repairDocxBuffer', () => {
  const settings = `<w:settings ${W}><w:compat/><w:zoom w:percent="100"/></w:settings>`;
  const body = doc('<w:p><w:smartTag w:uri="u" w:element="e"><w:r><w:t>same text</w:t></w:r></w:smartTag></w:p>');

  it('repairs a package and reports what it changed, with the same visible text', async () => {
    const input = packageWith({ 'word/settings.xml': settings, 'word/document.xml': body });
    const out = await repairDocxBuffer(input, { validate: noValidator });
    expect(out.normalized.length).toBeGreaterThanOrEqual(2);
    expect(out.rolledBack).toEqual([]);
    expect(partOf(out.buffer, 'word/document.xml')).toContain('same text');
    expect(partOf(out.buffer, 'word/document.xml')).not.toContain('smartTag');
    expect(partOf(out.buffer, 'word/settings.xml').indexOf('<w:zoom')).toBeLessThan(partOf(out.buffer, 'word/settings.xml').indexOf('<w:compat'));
  });

  it('a second run changes nothing and returns the same bytes', async () => {
    const first = await repairDocxBuffer(packageWith({ 'word/settings.xml': settings }), { validate: noValidator });
    const second = await repairDocxBuffer(first.buffer, { validate: noValidator });
    expect(second.normalized).toEqual([]);
    expect(second.buffer.equals(first.buffer)).toBe(true);
  });

  it('returns the input untouched when no rule applies', async () => {
    const input = createBlankDocxBuffer('blank');
    const out = await repairDocxBuffer(input, { validate: noValidator });
    expect(out.normalized).toEqual([]);
    expect(out.buffer.equals(input)).toBe(true);
  });

  it('drops only the rule that adds a validator error, and keeps the others', async () => {
    const input = packageWith({ 'word/settings.xml': settings, 'word/document.xml': body });
    // A validator test double that reports a new error once the smart tag is gone.
    const validate = async (buf: Buffer): Promise<ValidationResult> => {
      const gone = !partOf(buf, 'word/document.xml').includes('smartTag');
      return gone
        ? ({ file: 'x', ok: false, errors: [{ part: '/word/document.xml', id: 'Sch_Test', xPath: '/w:document[1]/w:body[1]', description: 'new', errorType: 'Schema' }] } as unknown as ValidationResult)
        : okResult;
    };
    const out = await repairDocxBuffer(input, { validate });
    expect(out.rolledBack.map((r) => r.rule)).toEqual(['smarttag-unwrap']);
    expect(out.rolledBack[0].reason).toContain('adds 1 validator error');
    expect(partOf(out.buffer, 'word/document.xml')).toContain('smartTag');
    expect(partOf(out.buffer, 'word/settings.xml').indexOf('<w:zoom')).toBeLessThan(partOf(out.buffer, 'word/settings.xml').indexOf('<w:compat'));
  });

  it('drops a rule that would change the visible text', async () => {
    const changing: RepairRule = {
      id: 'text-changer',
      description: 'test double',
      appliesTo: (p) => p === 'word/document.xml',
      apply(d) {
        const t = d.getElementsByTagName('w:t')[0];
        (t.firstChild as Text).data = 'different';
        return ['changed text'];
      },
    };
    const input = packageWith({ 'word/document.xml': body });
    const out = await repairDocxBuffer(input, { rules: [changing], validate: noValidator });
    expect(out.normalized).toEqual([]);
    expect(out.rolledBack).toEqual([{ rule: 'text-changer', reason: 'visible text would change' }]);
    expect(out.buffer.equals(input)).toBe(true);
  });

  it('reports the number of validator errors that remain', async () => {
    const still = async () => ({ file: 'x', ok: false, errors: [{ part: '/word/styles.xml', id: 'Sch_Other', xPath: '/w:styles[1]', description: 'left', errorType: 'Schema' }] }) as unknown as ValidationResult;
    const out = await repairDocxBuffer(packageWith({ 'word/settings.xml': settings }), { validate: still });
    expect(out.remaining).toBe(1);
    expect(out.rolledBack).toEqual([]);
  });
});
