/**
 * Validator allowance proofs: ECMA-376 schema validation expectations for
 * paragraph line numbers, combo-box attributes, and n-ary limit flags.
 * Every assertion runs the production path (DocxParser.validateOOXML,
 * normalize + sweep + native gate), never a mock.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser } from '@attentionform/astra-engine';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const FIXED_DATE = new Date('2026-01-02T10:00:00Z');

function packageBuffer(documentXml: string): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n' +
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n' +
        '  <Default Extension="xml" ContentType="application/xml"/>\n' +
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n' +
        '</Types>',
    ),
  );
  zip.addFile(
    '_rels/.rels',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n' +
        '</Relationships>',
    ),
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    ),
  );
  zip.addFile('word/document.xml', Buffer.from(documentXml, 'utf8'));
  for (const entry of zip.getEntries()) entry.header.time = FIXED_DATE;
  return zip.toBuffer();
}

const wrap = (inner: string, nsExtra = ''): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:document xmlns:w="${WORD_NS}"${nsExtra}><w:body>${inner}<w:sectPr/></w:body></w:document>`;

const MNS = ` xmlns:m="${MATH_NS}"`;

function comboBoxBuffer(inner: string): Buffer {
  return packageBuffer(
    wrap(
      `<w:p><w:sdt><w:sdtPr><w:alias w:val="Department"/><w:tag w:val="dept"/>` +
        `<w:comboBox>${inner}</w:comboBox></w:sdtPr>` +
        `<w:sdtContent><w:r><w:t>Legal</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
    ),
  );
}

function naryBuffer(prInner: string): Buffer {
  return packageBuffer(
    wrap(
      `<w:p><m:oMathPara><m:oMath><m:nary><m:naryPr>${prInner}</m:naryPr>` +
        `<m:sub><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup>` +
        `<m:e><m:r><m:t>f</m:t></m:r></m:e></m:nary></m:oMath></m:oMathPara></w:p>`,
      MNS,
    ),
  );
}

describe('paragraph line-number suppression beside siblings validates', () => {
  it('accepts suppressLineNumbers with spacing and justification in ECMA order', async () => {
    const buf = packageBuffer(
      wrap(
        `<w:p><w:pPr><w:suppressLineNumbers/>` +
          `<w:spacing w:after="160" w:line="259" w:lineRule="auto"/><w:jc w:val="both"/>` +
          `</w:pPr><w:r><w:t>flow copy.</w:t></w:r></w:p>`,
      ),
    );
    const validation = await new DocxParser(buf).validateOOXML();
    expect(validation.ok).toBe(true);
  });

  it('heals scrambled sibling order instead of rejecting', async () => {
    const buf = packageBuffer(
      wrap(
        `<w:p><w:pPr><w:jc w:val="both"/>` +
          `<w:spacing w:after="160"/><w:suppressLineNumbers/>` +
          `</w:pPr><w:r><w:t>flow copy.</w:t></w:r></w:p>`,
      ),
    );
    const validation = await new DocxParser(buf).validateOOXML();
    expect(validation.ok).toBe(true);
  });
});

describe('combo-box selection cache validates', () => {
  const items =
    `<w:listItem w:displayText="Finance" w:value="FIN"/>` +
    `<w:listItem w:displayText="Legal" w:value="LEG"/>`;

  it('accepts attribute form of lastValue', async () => {
    const buf = packageBuffer(
      wrap(
        `<w:p><w:sdt><w:sdtPr><w:alias w:val="Department"/><w:tag w:val="dept"/>` +
          `<w:comboBox w:lastValue="LEG">${items}</w:comboBox></w:sdtPr>` +
          `<w:sdtContent><w:r><w:t>Legal</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
      ),
    );
    const validation = await new DocxParser(buf).validateOOXML();
    expect(validation.ok).toBe(true);
  });

  it('rejects child lastValue element under the schema validator', async () => {
    const validation = await new DocxParser(
      comboBoxBuffer(`${items}<w:lastValue w:val="LEG"/>`),
    ).validateOOXML();
    expect(validation.ok).toBe(false);
  });

  it('still rejects a leading lastValue ahead of the list items', async () => {
    const validation = await new DocxParser(
      comboBoxBuffer(`<w:lastValue w:val="LEG"/>${items}`),
    ).validateOOXML();
    expect(validation.ok).toBe(false);
  });

  it('still rejects a duplicated lastValue', async () => {
    const validation = await new DocxParser(
      comboBoxBuffer(`${items}<w:lastValue w:val="LEG"/><w:lastValue w:val="FIN"/>`),
    ).validateOOXML();
    expect(validation.ok).toBe(false);
  });

  it('still rejects lastValue outside the combo box', async () => {
    const buf = packageBuffer(
      wrap(
        `<w:p><w:sdt><w:sdtPr><w:alias w:val="x"/><w:lastValue w:val="LEG"/>` +
          `</w:sdtPr><w:sdtContent><w:r><w:t>Legal</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
      ),
    );
    const validation = await new DocxParser(buf).validateOOXML();
    expect(validation.ok).toBe(false);
  });
});

describe('n-ary limit display flags validate', () => {
  it('accepts subHide and supHide after chr and limLoc in ECMA order', async () => {
    const validation = await new DocxParser(
      naryBuffer(
        `<m:chr m:val="&#x222B;"/><m:limLoc m:val="subSup"/>` +
          `<m:subHide m:val="off"/><m:supHide m:val="off"/>`,
      ),
    ).validateOOXML();
    expect(validation.ok).toBe(true);
  });

  it('accepts grow between limLoc and the hide flags', async () => {
    const validation = await new DocxParser(
      naryBuffer(
        `<m:chr m:val="&#x222B;"/><m:limLoc m:val="subSup"/>` +
          `<m:grow m:val="off"/><m:subHide m:val="off"/><m:supHide m:val="off"/>`,
      ),
    ).validateOOXML();
    expect(validation.ok).toBe(true);
  });

  it('still rejects a duplicated subHide', async () => {
    const validation = await new DocxParser(
      naryBuffer(
        `<m:chr m:val="&#x222B;"/><m:limLoc m:val="subSup"/>` +
          `<m:subHide m:val="off"/><m:subHide m:val="on"/>`,
      ),
    ).validateOOXML();
    expect(validation.ok).toBe(false);
  });

  it('still rejects flags ahead of the operator character', async () => {
    const validation = await new DocxParser(
      naryBuffer(`<m:subHide m:val="off"/><m:chr m:val="&#x222B;"/>`),
    ).validateOOXML();
    expect(validation.ok).toBe(false);
  });

  it('still rejects flags interleaved before the limit location', async () => {
    const validation = await new DocxParser(
      naryBuffer(
        `<m:chr m:val="&#x222B;"/><m:subHide m:val="off"/><m:limLoc m:val="subSup"/>`,
      ),
    ).validateOOXML();
    expect(validation.ok).toBe(false);
  });
});
