/**
 * Fidelity corpus fixture generator.
 *
 * Builds every hand-made .docx in fixtures/ (byte-deterministic input XML),
 * copies gallery baselines, and derives one hostile twin per family in
 * hostile/ (same shape, exactly one corruption). Run:
 *   npx tsx tests/harness/fidelity/generate-fixtures.ts
 * Paths resolve relative to this file, so cwd does not matter.
 */
import AdmZip from 'adm-zip';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '../../..');
const FIX = path.join(HERE, 'fixtures');
const HOSTILE = path.join(HERE, 'hostile');
const GALLERY = path.join(ROOT, 'test-gallery');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PNG1PX = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);
const DNS =
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const D = '2026-01-02T10:00:00Z';
// Pinned zip entry mtime: AdmZip defaults to Date.now(), which makes regen
// bytes drift run-to-run and breaks byte-determinism. Every writer pins to
// this fixed date before toBuffer() so regen twice is byte-identical.
const FIXED_DATE = new Date(D);

function pinZipTimes(zip: AdmZip): void {
  for (const entry of zip.getEntries()) {
    entry.header.time = FIXED_DATE;
  }
}

function pkg(documentXml: string, extra: Record<string, string | Buffer> = {}, rels = ''): Buffer {
  const hasNumbering = 'word/numbering.xml' in extra;
  const hasFootnotes = 'word/footnotes.xml' in extra;
  const hasEndnotes = 'word/endnotes.xml' in extra;
  const hasComments = 'word/comments.xml' in extra;
  const hasHeader = Object.keys(extra).some((k) => k.startsWith('word/header'));
  const hasFooter = Object.keys(extra).some((k) => k.startsWith('word/footer'));
  const hasStyles = 'word/styles.xml' in extra;
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${hasNumbering ? '\n  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ''}${hasFootnotes ? '\n  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : ''}${hasEndnotes ? '\n  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>' : ''}${hasComments ? '\n  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : ''}${hasHeader ? '\n  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}${hasFooter ? '\n  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}${hasStyles ? '\n  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' : ''}
</Types>`),
  );
  zip.addFile(
    '_rels/.rels',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`),
  );
  zip.addFile('word/document.xml', Buffer.from(documentXml));
  for (const [k, v] of Object.entries(extra)) zip.addFile(k, typeof v === 'string' ? Buffer.from(v) : v);
  pinZipTimes(zip);
  return zip.toBuffer();
}

const docWrap = (inner: string, nsExtra = ''): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W}"${nsExtra}><w:body>${inner}<w:sectPr/></w:body></w:document>`;
const para = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const picGraphic = (cnvId: string): string =>
  `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${cnvId}" name="p1"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic>`;

const IMG_EXTRA = { 'word/media/image1.png': PNG1PX };
const IMG_REL =
  '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>';
const NUM_REL =
  '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>';
const FN_REL =
  '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>';
const EN_REL =
  '<Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>';
const CM_REL =
  '<Relationship Id="rIdCm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>';
const HD_REL =
  '<Relationship Id="rIdHd" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>';
const FT_REL =
  '<Relationship Id="rIdFt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>';
const ST_REL =
  '<Relationship Id="rIdSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
const HL_REL =
  '<Relationship Id="rIdHl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/report" TargetMode="External"/>';

const MNS = ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

const LEAD_A =
  '<w:p><w:r><w:t xml:space="preserve">A first plain sentence stands before the pathology under test in this fixture.</w:t></w:r></w:p>';
const LEAD_B =
  '<w:p><w:r><w:t xml:space="preserve">A second plain sentence gives the mutation runner its required pair for this fixture.</w:t></w:r></w:p>';

const footnotesXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>Footnote one carries a citation gloss for the corpus.</w:t></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Footnote two holds a second gloss with more detail.</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const endnotesXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:endnotes xmlns:w="${W}"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="1"><w:p><w:r><w:t>Endnote one collects a closing remark for the chapter.</w:t></w:r></w:p></w:endnote></w:endnotes>`;
const commentsXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="reviewer" w:date="${D}" w:initials="RV"><w:p><w:r><w:t>First review note asks to verify this clause wording.</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="editor" w:date="${D}" w:initials="ED"><w:p><w:r><w:t>Second note confirms the figure caption matches the art.</w:t></w:r></w:p></w:comment></w:comments>`;
const headerXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Running head for the probe document section.</w:t></w:r></w:p></w:hdr>`;
const footerXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p></w:ftr>`;
const stylesXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:latentStyles w:defLockedState="0" w:count="2"><w:lsdException w:name="Normal" w:qFormat="1"/><w:lsdException w:name="heading 1" w:qFormat="1"/></w:latentStyles><w:style w:type="paragraph" w:styleId="Quoteblock"><w:name w:val="Quote block"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="120"/><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style></w:styles>`;

const numberingRestart =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W}">` +
  `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:lvl>` +
  `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>` +
  `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
  `<w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride></w:num></w:numbering>`;

const numberingDuo =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W}">` +
  `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
  `</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>` +
  `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
  `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

function numParas(items: Array<{ ilvl: string; numId: string; text: string }>): string {
  return items
    .map(
      (it) =>
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="${it.ilvl}"/><w:numId w:val="${it.numId}"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${it.text}</w:t></w:r></w:p>`,
    )
    .join('');
}

const builders: Record<string, () => Buffer> = {
  'merged-cells-basic': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the merged cell tangle.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
          `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide cell</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>side cell</w:t></w:r></w:p></w:tc></w:tr>` +
          `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>merge top</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>middle</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>edge</w:t></w:r></w:p></w:tc></w:tr>` +
          `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>lower middle</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>lower edge</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'merged-cells-header-repeat': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the repeating header table.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>` +
          `<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Header one</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Header two</w:t></w:r></w:p></w:tc></w:tr>` +
          `<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'nested-table-single': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the nested table.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>` +
          `<w:tr><w:tc><w:p><w:r><w:t>outer cell opens</w:t></w:r></w:p>` +
          `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
          `<w:tr><w:tc><w:p><w:r><w:t>inner A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>inner B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
          `<w:p><w:r><w:t>outer cell closes</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'nested-table-deep': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the deep nesting.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level zero</w:t></w:r></w:p>` +
          `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level one</w:t></w:r></w:p>` +
          `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="7000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level two leaf</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
          `</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'revision-multiscript': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the multiscript revision.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:p><w:ins w:id="10" w:author="layla" w:date="${D}"><w:r><w:t xml:space="preserve">Inserted Latin plus </w:t></w:r><w:r><w:t xml:space="preserve">Arabic مرحبا plus </w:t></w:r><w:r><w:t xml:space="preserve">CJK 你好 end.</w:t></w:r></w:ins>` +
          `<w:del w:id="11" w:author="omar" w:date="${D}"><w:r><w:delText xml:space="preserve">Deleted Latin plus </w:delText></w:r><w:r><w:delText xml:space="preserve">Arabic وداعا plus </w:delText></w:r><w:r><w:delText xml:space="preserve">CJK 再见 end.</w:delText></w:r></w:del>` +
          `<w:r><w:t>plain tail.</w:t></w:r></w:p>`,
      ),
    ),
  'revision-nested-change': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the nested change.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:p><w:del w:id="30" w:author="ed" w:date="${D}"><w:r><w:delText xml:space="preserve">outer deletion holding </w:delText></w:r><w:ins w:id="31" w:author="jo" w:date="${D}"><w:r><w:t>an inner insertion</w:t></w:r></w:ins><w:r><w:delText xml:space="preserve"> and more deletion.</w:delText></w:r></w:del></w:p>`,
      ),
    ),
  'revision-soup': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the revision soup.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:p><w:ins w:id="20" w:author="a1" w:date="${D}"><w:r><w:rPr><w:b/></w:rPr><w:t>bold insert</w:t></w:r></w:ins>` +
          `<w:r><w:t xml:space="preserve"> middle </w:t></w:r>` +
          `<w:del w:id="21" w:author="a2" w:date="${D}"><w:r><w:rPr><w:i/></w:rPr><w:delText>italic delete</w:delText></w:r></w:del>` +
          `<w:ins w:id="22" w:author="a3" w:date="${D}"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>underline insert</w:t></w:r></w:ins></w:p>` +
          `<w:p><w:del w:id="24" w:author="a4" w:date="${D}"><w:r><w:delText>fourth author deletion</w:delText></w:r><w:r><w:delText xml:space="preserve"> across two runs.</w:delText></w:r></w:del>` +
          `<w:ins w:id="25" w:author="a5" w:date="${D}"><w:r><w:t>fifth author insertion.</w:t></w:r></w:ins></w:p>`,
      ),
    ),
  'drawing-negative-anchor': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the anchored drawing.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">` +
          `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>-914400</wp:posOffset></wp:positionH>` +
          `<wp:positionV relativeFrom="paragraph"><wp:posOffset>-457200</wp:posOffset></wp:positionV>` +
          `<wp:extent cx="1828800" cy="1828800"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>${picGraphic('1')}</wp:anchor></w:drawing></w:r></w:p>`,
        DNS,
      ),
      IMG_EXTRA,
      IMG_REL,
    ),
  'drawing-inline-wrap': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the inline drawing.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="Inline 1"/><wp:cNvGraphicFramePr/>${picGraphic('2')}</wp:inline></w:drawing></w:r></w:p>` +
          `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">` +
          `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
          `<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>` +
          `<wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="3" name="Wrapped 1"/><wp:cNvGraphicFramePr/>${picGraphic('3')}</wp:anchor></w:drawing></w:r></w:p>`,
        DNS,
      ),
      IMG_EXTRA,
      IMG_REL,
    ),
  'numbering-restart': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the restarted list.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          numParas([
            { ilvl: '0', numId: '1', text: 'Item one of list A.' },
            { ilvl: '0', numId: '1', text: 'Item two of list A.' },
            { ilvl: '1', numId: '1', text: 'Nested letter item.' },
            { ilvl: '0', numId: '2', text: 'Restarted list at seven.' },
          ]),
      ),
      { 'word/numbering.xml': numberingRestart },
      NUM_REL,
    ),
  'numbering-abstract-duo': () =>
    pkg(
      docWrap(
        para('A first plain sentence stands before the dual list.') +
          para('A second plain sentence gives the mutation runner its pair.') +
          numParas([
            { ilvl: '0', numId: '1', text: 'Decimal item one.' },
            { ilvl: '0', numId: '1', text: 'Decimal item two.' },
            { ilvl: '0', numId: '2', text: 'Bullet item one.' },
          ]),
      ),
      { 'word/numbering.xml': numberingDuo },
      NUM_REL,
    ),
  // --- anchored-drawing ---
  'drawing-wrap-tight': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="914400" y="0"/><wp:lineTo x="914400" y="914400"/><wp:lineTo x="0" y="914400"/></wp:wrapPolygon></wp:wrapTight><wp:docPr id="9" name="Tight 1" descr="tight wrapped figure"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="9" name="p9"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`,
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
      ),
      IMG_EXTRA,
      IMG_REL,
    ),
  // --- bookmarks-fields ---
  'bookmark-pair': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:bookmarkStart w:id="7" w:name="ClauseNine"/><w:r><w:t>clause copy enclosed by a named bookmark range.</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p><w:p><w:bookmarkStart w:id="8" w:name="_GoBack"/><w:r><w:t>cursor restore marker copy for the probe.</w:t></w:r><w:bookmarkEnd w:id="8"/></w:p>`,
      ),
    ),
  'field-simple-page': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr="PAGE \\* MERGEFORMAT"><w:r><w:t>3</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr="NUMPAGES \\* MERGEFORMAT"><w:r><w:t>41</w:t></w:r></w:fldSimple></w:p><w:p><w:r><w:t>plain tail copy after the field paragraph.</w:t></w:r></w:p>`,
      ),
    ),
  'field-complex-date': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p><w:p><w:r><w:instrText xml:space="preserve"> DATE \\@ "d MMMM yyyy" </w:instrText></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:t>2 January 2026</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
      ),
    ),
  'field-pageref': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:bookmarkStart w:id="21" w:name="FigureTwo"/><w:r><w:t>Figure 2 caption anchor copy.</w:t></w:r><w:bookmarkEnd w:id="21"/></w:p><w:p><w:r><w:t xml:space="preserve">See page </w:t></w:r><w:fldSimple w:instr="PAGEREF FigureTwo \\h"><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>`,
      ),
    ),
  // --- breaks-symbols ---
  'break-page-column': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>copy before a hard page break marker.</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>copy before a column break marker.</w:t></w:r><w:r><w:br w:type="column"/></w:r></w:p>`,
      ),
    ),
  'break-wrap-hyphen': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>line one of wrapped copy</w:t></w:r><w:r><w:br w:type="textWrapping" w:clear="all"/></w:r><w:r><w:t>line two after the wrap clear.</w:t></w:r></w:p><w:p><w:r><w:t>extraordinarilylongtoken</w:t></w:r><w:r><w:noBreakHyphen/></w:r><w:r><w:t>plus soft</w:t></w:r><w:r><w:softHyphen/></w:r><w:r><w:t>hyphen tail.</w:t></w:r></w:p>`,
      ),
    ),
  'sym-tab-chars': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:sym w:font="Wingdings" w:char="F0A8"/></w:r><w:r><w:t xml:space="preserve"> flagged item copy </w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>trailing copy after tab.</w:t></w:r></w:p>`,
      ),
    ),
  // --- content-controls ---
  'sdt-block-rich': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:sdt><w:sdtPr><w:alias w:val="Client name"/><w:tag w:val="clientName"/><w:lock w:val="contentLocked"/><w:placeholder><w:docPart w:val="DefaultPlaceholder_1082065158"/></w:placeholder><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Acme Industrial Holdings Ltd.</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
      ),
    ),
  'sdt-inline-date': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>Signed on </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Signing date"/><w:tag w:val="signDate"/><w:date w:fullDate="${D}"><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-GB"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>2 January 2026</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> in London.</w:t></w:r></w:p>`,
      ),
    ),
  'sdt-nested-repeat': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:sdt><w:sdtPr><w:alias w:val="Clause block"/><w:tag w:val="clause"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>outer clause copy with a nested control inside.</w:t></w:r></w:p><w:sdt><w:sdtPr><w:alias w:val="Amount"/><w:tag w:val="amount"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>twelve thousand pounds</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>`,
      ),
    ),
  'sdt-combobox': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:sdt><w:sdtPr><w:alias w:val="Department"/><w:tag w:val="dept"/><w:comboBox w:lastValue="LEG"><w:listItem w:displayText="Finance" w:value="FIN"/><w:listItem w:displayText="Legal" w:value="LEG"/></w:comboBox></w:sdtPr><w:sdtContent><w:r><w:t>Legal</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
      ),
    ),
  // --- document-parts ---
  'styles-custom-quote': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:pStyle w:val="Quoteblock"/></w:pPr><w:r><w:t>quoted contract language styled through the custom style.</w:t></w:r></w:p>`,
      ),
      { 'word/styles.xml': stylesXml },
      ST_REL,
    ),
  'footnote-pair': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>body copy carrying the first footnote marker</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> and the second marker</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t>.</w:t></w:r></w:p>`,
      ),
      { 'word/footnotes.xml': footnotesXml },
      FN_REL,
    ),
  'endnote-single': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>chapter copy carrying an endnote marker at its close</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>`,
      ),
      { 'word/endnotes.xml': endnotesXml },
      EN_REL,
    ),
  'comment-range-pair': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>clause copy under the first reviewer note.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>caption copy under the second editor note.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`,
      ),
      { 'word/comments.xml': commentsXml },
      CM_REL,
    ),
  'header-footer-basic': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>body copy under a running head and page footer.</w:t></w:r></w:p>`,
      ),
      { 'word/header1.xml': headerXml, 'word/footer1.xml': footerXml },
      HD_REL + FT_REL,
    ),
  // --- hyperlinks ---
  'hyperlink-external': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:hyperlink r:id="rIdHl" w:history="1"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>annual report landing page</w:t></w:r></w:hyperlink></w:p>`,
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      ),
      {},
      HL_REL,
    ),
  'hyperlink-bookmark': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:bookmarkStart w:id="31" w:name="AnnexB"/><w:r><w:t>Annex B anchor copy for internal linking.</w:t></w:r><w:bookmarkEnd w:id="31"/></w:p><w:p><w:hyperlink w:anchor="AnnexB" w:history="1"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>jump to Annex B</w:t></w:r></w:hyperlink></w:p>`,
      ),
    ),
  // --- math ---
  'math-inline-fraction': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>y</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p>`,
        MNS,
      ),
    ),
  'math-sup-nary': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t xml:space="preserve"> + </m:t></m:r><m:nary><m:naryPr><m:chr m:val="&#x222B;"/><m:limLoc m:val="subSup"/><m:subHide m:val="off"/><m:supHide m:val="off"/></m:naryPr><m:sub><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup><m:e><m:r><m:t>f</m:t></m:r></m:e></m:nary></m:oMath></m:oMathPara></w:p>`,
        MNS,
      ),
    ),
  'math-delimiters': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:d><m:r><m:t xml:space="preserve"> = </m:t></m:r><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>c</m:t></m:r></m:e></m:rad></m:oMath></m:oMathPara></w:p>`,
        MNS,
      ),
    ),
  // --- merged-cells ---
  'merged-cells-tall': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>tall merge</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>r1c2</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:gridSpan w:val="1"/></w:tcPr><w:p><w:r><w:t>r1c3</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide middle</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>r3c2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>r3c3</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  // --- nested-tables ---
  'nested-merge-combo': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide outer</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>inner top</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>inner side</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>inner lower</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>`,
      ),
    ),
  // --- numbering ---
  'numbering-roman-article': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Article One opening provision copy.</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Article Two following provision copy.</w:t></w:r></w:p>`,
      ),
      { 'word/numbering.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="Article %1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Section %1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1800" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>` },
      NUM_REL,
    ),
  'numbering-override-duo': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>resumed list item restarting at twelve.</w:t></w:r></w:p>`,
      ),
      { 'word/numbering.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="12"/></w:lvlOverride></w:num></w:numbering>` },
      NUM_REL,
    ),
  // --- paragraph-flow ---
  'para-keep-flow': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:keepNext/><w:keepLines/><w:widowControl/></w:pPr><w:r><w:t>heading copy pinned to the paragraph that follows it.</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/><w:widowControl w:val="0"/></w:pPr><w:r><w:t>body copy forced onto a fresh page by the break.</w:t></w:r></w:p>`,
      ),
    ),
  'para-context-spacing': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:suppressLineNumbers/><w:spacing w:before="200" w:after="200"/><w:contextualSpacing/></w:pPr><w:r><w:t>list-adjacent copy with contextual spacing enabled.</w:t></w:r></w:p>`,
      ),
    ),
  'para-mirror-indent': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:ind w:left="720" w:right="720"/><w:mirrorIndents/><w:suppressOverlap/></w:pPr><w:r><w:t>mirrored margin copy for bound book layout.</w:t></w:r></w:p>`,
      ),
    ),
  // --- paragraph-layout ---
  'para-spacing-indent': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="exact"/><w:ind w:left="1440" w:right="720" w:firstLine="360"/></w:pPr><w:r><w:t>spaced and indented body copy for the probe.</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:after="120"/><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>hanging indent follow-on copy for the probe.</w:t></w:r></w:p>`,
      ),
    ),
  'para-justify-levels': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:jc w:val="center"/><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>centered heading-styled copy for the probe.</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="both"/><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>justified body copy exercising distributed line breaking behavior here.</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:t>distributed copy exercising east asian justification behavior.</w:t></w:r></w:p>`,
      ),
    ),
  'para-borders-shading': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="12" w:space="4" w:color="1F4E79"/><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1F4E79"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/></w:pPr><w:r><w:t>callout copy with borders and shading applied.</w:t></w:r></w:p>`,
      ),
    ),
  'para-tabs-stops': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="center" w:pos="2880"/><w:tab w:val="right" w:pos="4320"/><w:tab w:val="decimal" w:pos="5760"/></w:tabs></w:pPr><w:r><w:t>Name</w:t></w:r><w:r><w:tab/><w:t>Value</w:t></w:r><w:r><w:tab/><w:t>1.25</w:t></w:r></w:p>`,
      ),
    ),
  // --- run-scripts ---
  'run-super-sub': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>Formula H</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:t>O plus x</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:spacing w:val="40"/><w:kern w:val="8"/><w:position w:val="8"/></w:rPr><w:t>raised tracked copy.</w:t></w:r></w:p>`,
      ),
    ),
  'run-lang-bidi': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:bCs/><w:szCs w:val="24"/><w:rtl/><w:lang w:val="ar-SA" w:bidi="ar-SA"/></w:rPr><w:t>مرحبا بالعالم في فقرة تجريبية</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="SimSun"/><w:lang w:val="fr-FR" w:eastAsia="zh-CN"/></w:rPr><w:t>bonjour monde with language tagging applied.</w:t></w:r></w:p>`,
      ),
    ),
  'run-noproof-err': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:t>ZZZQ code token skipped by spellcheck here.</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>teh flagged token with proof markup around.</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>`,
      ),
    ),
  // --- run-typography ---
  'run-color-fonts': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/><w:color w:val="C00000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>colored display copy in a distinct face.</w:t></w:r><w:r><w:rPr><w:color w:val="404040" w:themeColor="text1" w:themeShade="80"/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:rPr><w:t>theme-tinted copy with background shading.</w:t></w:r></w:p>`,
      ),
    ),
  'run-emphasis-stack': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:smallCaps/><w:strike/><w:u w:val="double"/></w:rPr><w:t>heavy emphasis stack for the probe run.</w:t></w:r><w:r><w:rPr><w:caps/><w:dstrike/><w:u w:val="wave" w:color="FF0000"/></w:rPr><w:t>second emphasis stack with wave underline.</w:t></w:r></w:p>`,
      ),
    ),
  'run-highlight-shadow': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:rPr><w:shadow w:val="1"/><w:emboss/><w:highlight w:val="yellow"/></w:rPr><w:t>highlighted embossed marker copy for review.</w:t></w:r><w:r><w:rPr><w:outline/><w:imprint/><w:highlight w:val="cyan"/></w:rPr><w:t>second marked copy with imprint and outline.</w:t></w:r></w:p>`,
      ),
    ),
  'run-vanish-hidden': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:r><w:t>visible anchor copy before hidden runs.</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden annotation copy not for print.</w:t></w:r><w:r><w:rPr><w:webHidden/></w:rPr><w:t>web-only hidden copy for the probe.</w:t></w:r></w:p>`,
      ),
    ),
  // --- section-layout ---
  'sect-landscape-margins': () =>
    pkg(
      docWrap(
        `<w:p><w:r><w:t>portrait section copy before the break marker.</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/><w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr><w:r><w:t>landscape two-column copy after the break.</w:t></w:r></w:p>` + `<w:p><w:r><w:t>second plain sentence after landscape for mutation pair.</w:t></w:r></w:p>`,
      ),
    ),
  'sect-pagenum-docgrid': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:sectPr><w:pgNumType w:start="7" w:chapStyle="1"/><w:cols w:num="1" w:space="720"/><w:rtlGutter/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:pPr><w:r><w:t>section copy with page numbering restart and grid.</w:t></w:r></w:p>`,
      ),
    ),
  'sect-titlepg-gutter': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="240"/><w:titlePg/></w:sectPr></w:pPr><w:r><w:t>title-page section copy with wide margins set.</w:t></w:r></w:p>`,
      ),
    ),
  // --- table-borders ---
  'tbl-borders-shading': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="808080"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="808080"/></w:tblBorders><w:shd w:val="clear" w:color="auto" w:fill="E8E8E8"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="6" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFE699"/></w:tcPr><w:p><w:r><w:t>shaded cell</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>plain cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'tbl-cell-margins': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellSpacing w:w="20" w:type="dxa"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcMar><w:top w:w="200" w:type="dxa"/><w:left w:w="200" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>margin cell</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>other cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'tbl-look-banded': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>H3</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  // --- table-layout ---
  'tbl-fixed-layout': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>fixed A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:t>fixed B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'tbl-float-overlap': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblpPr w:horzAnchor="margin" w:vertAnchor="text" w:tblpX="200" w:tblpY="400" w:tblpXSpec="center" w:tblpYSpec="top"/><w:tblOverlap w:val="never"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>floating table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  'tbl-row-height': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="atLeast"/><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>tall header row</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:trPr><w:trHeight w:val="300" w:hRule="exact"/></w:trPr><w:tc><w:p><w:r><w:t>exact row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
  // --- tracked-revisions ---
  'revision-move-pair': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:p><w:moveFrom w:id="60" w:author="mv" w:date="${D}"><w:r><w:delText xml:space="preserve">moved clause leaving </w:delText></w:r></w:moveFrom><w:r><w:t xml:space="preserve">stable middle </w:t></w:r><w:moveTo w:id="61" w:author="mv" w:date="${D}"><w:r><w:t>arriving clause copy.</w:t></w:r></w:moveTo></w:p>`,
      ),
    ),
  'revision-cell-split': () =>
    pkg(
      docWrap(
        LEAD_A + LEAD_B + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>steady cell</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:cellIns w:id="70" w:author="ed" w:date="${D}"/></w:tcPr><w:p><w:r><w:t>inserted cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ),
};

const galleryCopies: Record<string, string> = {
  'gallery-coffee-onepager': 'T1_coffee_onepager.docx',
  'gallery-edit-target': 'T2_edit_target.docx',
  'gallery-image-doc': 'T5_image_doc.docx',
  'gallery-footnotes-endnotes': 'T8_footnotes_endnotes.docx',
  'gallery-proposal': 'T3_proposal.docx',
  'gallery-budget': 'T4_budget.docx',
  'gallery-honesty': 'T6_honesty.docx',
  'gallery-visual-probe': 'T7_visual_probe.docx',
  'gallery-postmigration': 'T9_postmigration.docx',
};

interface TwinSpec {
  family: string;
  base: string;
  corruption: string;
  expectedRefusal: 'open' | 'ingest' | 'rezip' | 'validate';
  corrupt: (documentXml: string, parts: Record<string, Buffer>) => { documentXml?: string; parts?: Record<string, Buffer> };
}

function swapEntry(buf: Buffer, entry: string, replacement: Buffer): Buffer {
  const zip = new AdmZip(buf);
  zip.updateFile(entry, replacement);
  pinZipTimes(zip);
  return zip.toBuffer();
}

const twins: TwinSpec[] = [
  {
    family: 'merged-cells',
    base: 'merged-cells-basic',
    corruption: 'w:gridSpan loses its required w:val attribute (schema violation, same element shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:gridSpan w:val="2"/>', '<w:gridSpan/>') }),
  },
  {
    family: 'nested-tables',
    base: 'nested-table-single',
    corruption: 'one w:tc close tag is swapped for w:tr (mismatched markup, same element shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({
      documentXml: documentXml.replace(
        '<w:tc><w:p><w:r><w:t>inner B</w:t></w:r></w:p></w:tc></w:tr>',
        '<w:tc><w:p><w:r><w:t>inner B</w:t></w:r></w:p></w:tr>',
      ),
    }),
  },
  {
    family: 'tracked-revisions',
    base: 'revision-multiscript',
    corruption: 'w:ins loses its required w:id attribute (schema violation, same range shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:ins w:id="10"', '<w:ins') }),
  },
  {
    family: 'anchored-drawing',
    base: 'drawing-negative-anchor',
    corruption: 'required wp:extent child removed from wp:anchor (schema violation, same anchor shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<wp:extent cx="1828800" cy="1828800"/>', '') }),
  },
  {
    family: 'numbering',
    base: 'numbering-restart',
    corruption: 'word/numbering.xml truncated mid-element (same package shape, corrupt part).',
    expectedRefusal: 'rezip',
    corrupt: (_documentXml, parts) => ({
      parts: { 'word/numbering.xml': parts['word/numbering.xml'].subarray(0, 200) },
    }),
  },
  {
    family: 'gallery-baseline',
    base: 'gallery-coffee-onepager',
    corruption: 'first w:t close tag truncated by one byte (same package shape, corrupt part).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => {
      const at = documentXml.indexOf('</w:t>');
      return { documentXml: at < 0 ? documentXml : documentXml.slice(0, at) + '</w:t' + documentXml.slice(at + 6) };
    },
  },
  {
    family: 'table-borders',
    base: 'tbl-borders-shading',
    corruption: 'w:top inside w:tblBorders loses its required w:val attribute (schema violation, same border shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:top w:val="single" w:sz="8"', '<w:top w:sz="8"') }),
  },
  {
    family: 'table-layout',
    base: 'tbl-fixed-layout',
    corruption: 'w:jc loses its required w:val attribute (schema violation, same layout shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:jc w:val="center"/>', '<w:jc/>') }),
  },
  {
    family: 'paragraph-layout',
    base: 'para-justify-levels',
    corruption: 'w:jc loses its required w:val attribute (schema violation, same paragraph shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:jc w:val="center"/>', '<w:jc/>') }),
  },
  {
    family: 'paragraph-flow',
    base: 'para-keep-flow',
    corruption: 'first w:t close tag truncated by one byte (same package shape, corrupt part).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => {
      const at = documentXml.indexOf('</w:t>');
      return { documentXml: at < 0 ? documentXml : documentXml.slice(0, at) + '</w:t' + documentXml.slice(at + 6) };
    },
  },
  {
    family: 'run-typography',
    base: 'run-color-fonts',
    corruption: 'w:sz loses its required w:val attribute (schema violation, same run shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:sz w:val="28"/>', '<w:sz/>') }),
  },
  {
    family: 'run-scripts',
    base: 'run-super-sub',
    corruption: 'w:vertAlign loses its required w:val attribute (schema violation, same run shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:vertAlign w:val="subscript"/>', '<w:vertAlign/>') }),
  },
  {
    family: 'breaks-symbols',
    base: 'break-page-column',
    corruption: 'first w:t close tag truncated by one byte (same package shape, corrupt part).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => {
      const at = documentXml.indexOf('</w:t>');
      return { documentXml: at < 0 ? documentXml : documentXml.slice(0, at) + '</w:t' + documentXml.slice(at + 6) };
    },
  },
  {
    family: 'bookmarks-fields',
    base: 'bookmark-pair',
    corruption: 'w:bookmarkStart loses its required w:name attribute (schema violation, same range shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:bookmarkStart w:id="7" w:name="ClauseNine"/>', '<w:bookmarkStart w:id="7"/>') }),
  },
  {
    family: 'content-controls',
    base: 'sdt-block-rich',
    corruption: 'closing w:sdtContent tag removed (mismatched markup, same control shape).',
    expectedRefusal: 'open',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('</w:sdtContent>', '') }),
  },
  {
    family: 'math',
    base: 'math-delimiters',
    corruption: 'm:begChr loses its required m:val attribute (schema violation, same formula shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<m:begChr m:val="("/>', '<m:begChr/>') }),
  },
  {
    family: 'section-layout',
    base: 'sect-landscape-margins',
    corruption: 'w:type loses its required w:val attribute (schema violation, same section shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:type w:val="continuous"/>', '<w:type/>') }),
  },
  {
    family: 'document-parts',
    base: 'footnote-pair',
    corruption: 'w:footnoteReference loses its required w:id attribute (schema violation, same note shape).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('<w:footnoteReference w:id="1"/>', '<w:footnoteReference/>') }),
  },
  {
    family: 'hyperlinks',
    base: 'hyperlink-external',
    corruption: 'w:hyperlink points at a dangling r:id with no matching relationship (same link shape, broken reference).',
    expectedRefusal: 'validate',
    corrupt: (documentXml) => ({ documentXml: documentXml.replace('r:id="rIdHl"', 'r:id="rIdHlX"') }),
  },
];

function main(): void {
  fs.mkdirSync(FIX, { recursive: true });
  fs.mkdirSync(HOSTILE, { recursive: true });
  for (const [name, build] of Object.entries(builders)) {
    fs.writeFileSync(path.join(FIX, `${name}.docx`), build());
    console.log(`fixture ${name}.docx`);
  }
  for (const [name, file] of Object.entries(galleryCopies)) {
    fs.copyFileSync(path.join(GALLERY, file), path.join(FIX, `${name}.docx`));
    console.log(`gallery ${name}.docx <- ${file}`);
  }
  for (const twin of twins) {
    const baseBuf = fs.readFileSync(path.join(FIX, `${twin.base}.docx`));
    const zip = new AdmZip(baseBuf);
    const parts: Record<string, Buffer> = {};
    for (const entry of zip.getEntries()) {
      if (!entry.entryName.endsWith('/')) parts[entry.entryName] = entry.getData();
    }
    const edits = twin.corrupt(parts['word/document.xml'].toString('utf8'), parts);
    let out: Buffer = baseBuf;
    if (edits.documentXml !== undefined) out = swapEntry(out, 'word/document.xml', Buffer.from(edits.documentXml));
    if (edits.parts) {
      for (const [entry, data] of Object.entries(edits.parts)) out = swapEntry(out, entry, data);
    }
    const twinName = `${twin.family}.twin`;
    fs.writeFileSync(path.join(HOSTILE, `${twinName}.docx`), out);
    fs.writeFileSync(
      path.join(HOSTILE, `${twinName}.json`),
      JSON.stringify(
        { family: twin.family, twin: twinName, base: twin.base, corruption: twin.corruption, expectedRefusal: twin.expectedRefusal },
        null,
        2,
      ) + '\n',
    );
    console.log(`twin ${twinName}.docx (from ${twin.base})`);
  }
}

main();
