import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocumentAST,
  executeImage,
  executeHyperlink,
  isRogueZipEntry,
  redactBodyPiiInWordText,
  sanitizeCorePropertiesXml,
  sanitizeDocxBuffer,
  sanitizeDocumentXml,
  sanitizePackageEntries,
} from '@attentionform/astra-engine';

const W_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function createMockAST(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    `<w:document xmlns:w="${W_MAIN}"><w:body><w:p/></w:body></w:document>`,
    'text/xml'
  );
  const bodyP = doc.getElementsByTagName('w:p')[0];
  return {
    metadata: {},
    sections: [
      {
        id: 'sec_root',
        heading: null,
        nodes: [
          {
            id: 'p_001',
            type: 'paragraph',
            xml: bodyP,
            metadata: {
              text: 'Intro paragraph',
              hasCitation: false,
              hasComment: false,
              hasBookmark: false,
              hasTrackedChanges: false,
              hasEquation: false,
              hasField: false,
              hasImage: false,
              hasHyperlink: false,
            },
            version: 1,
            checksum: 'chk_1',
          },
        ],
        pageSetup: {},
        version: 1,
        checksum: 'sec_chk_1',
      },
    ],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1,
  };
}

function packageBuffer(entries: Record<string, string>): Buffer {
  const zip = new AdmZip();
  for (const [name, text] of Object.entries(entries)) {
    zip.addFile(name, Buffer.from(text, 'utf8'));
  }
  if (!entries['[Content_Types].xml']) {
    zip.addFile(
      '[Content_Types].xml',
      Buffer.from('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8')
    );
  }
  return zip.toBuffer();
}

const TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`;

describe('Input gates follow-ups: body-PII closures', () => {
  it('redacts PII planted in word/comments.xml and counts it as bodyPII', () => {
    const comments =
      `<w:comments xmlns:w="${W_MAIN}">` +
      `<w:comment w:id="0" w:author="Reviewer" w:initials="R" w:date="2026-02-01T00:00:00Z">` +
      `<w:p><w:r><w:t>ping jane.doe@example.com or 212-555-0147 about this</w:t></w:r></w:p>` +
      `</w:comment></w:comments>`;
    const files = new Map<string, Buffer>([['word/comments.xml', Buffer.from(comments, 'utf8')]]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    const out = clean.get('word/comments.xml')!.toString('utf8');
    expect(out).not.toContain('jane.doe@example.com');
    expect(out).not.toContain('212-555-0147');
    expect(out).toContain('[redacted]');
    expect(report.bodyPII).toBe(2);
    expect(report.propsPII).toBe(0);
    expect(report.purgedPII).toBe(report.bodyPII + report.propsPII);
  });

  it('redacts PII planted in word/people.xml element text and w15:userId attributes', () => {
    const people =
      `<w15:people xmlns:w="${W_MAIN}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">` +
      `<w15:person w15:author="Jane Doe">` +
      `<w15:presenceInfo w15:providerId="AD" w15:userId="jane.doe@example.com"/>` +
      `<w15:contact>backup 415-555-0199</w15:contact>` +
      `</w15:person></w15:people>`;
    const files = new Map<string, Buffer>([['word/people.xml', Buffer.from(people, 'utf8')]]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    const out = clean.get('word/people.xml')!.toString('utf8');
    expect(out).not.toContain('jane.doe@example.com');
    expect(out).not.toContain('415-555-0199');
    expect(report.bodyPII).toBe(2);
    expect(report.propsPII).toBe(0);
  });

  it('redacts PII planted in word/customXml/item1.xml element text', () => {
    const item =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<my:record xmlns:my="https://example.com/schema">` +
      `<my:contact>jane.doe@example.com 555-123-4567</my:contact>` +
      `</my:record>`;
    const files = new Map<string, Buffer>([['word/customXml/item1.xml', Buffer.from(item, 'utf8')]]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    const out = clean.get('word/customXml/item1.xml')!.toString('utf8');
    expect(out).not.toContain('jane.doe@example.com');
    expect(out).not.toContain('555-123-4567');
    expect(report.bodyPII).toBe(2);
    expect(report.propsPII).toBe(0);
  });

  it('redacts PII planted in word/charts/chart1.xml a:t and c:v element text', () => {
    const chart =
      `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>owner jane.doe@example.com</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:ser><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>4111111111111111</c:v></c:pt></c:numLit></c:val></c:ser>` +
      `</c:chart></c:chartSpace>`;
    const files = new Map<string, Buffer>([['word/charts/chart1.xml', Buffer.from(chart, 'utf8')]]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    const out = clean.get('word/charts/chart1.xml')!.toString('utf8');
    expect(out).not.toContain('jane.doe@example.com');
    expect(out).not.toContain('4111111111111111');
    expect(report.bodyPII).toBe(2);
    expect(report.propsPII).toBe(0);
  });

  it('redacts PII planted in style-family name attributes, keeping style structure', () => {
    const styles =
      `<w:styles xmlns:w="${W_MAIN}">` +
      `<w:style w:type="paragraph" w:styleId="ContactNote">` +
      `<w:name w:val="Note jane.doe@example.com"/>` +
      `<w:rPr><w:sz w:val="22"/></w:rPr>` +
      `</w:style></w:styles>`;
    const numbering =
      `<w:numbering xmlns:w="${W_MAIN}">` +
      `<w:abstractNum w:abstractNumId="0">` +
      `<w:lvl w:ilvl="0"><w:lvlText w:val="Call 555-123-4567 %1."/></w:lvl>` +
      `</w:abstractNum></w:numbering>`;
    const files = new Map<string, Buffer>([
      ['word/styles.xml', Buffer.from(styles, 'utf8')],
      ['word/numbering.xml', Buffer.from(numbering, 'utf8')],
    ]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    const outStyles = clean.get('word/styles.xml')!.toString('utf8');
    const outNumbering = clean.get('word/numbering.xml')!.toString('utf8');
    expect(outStyles).not.toContain('jane.doe@example.com');
    expect(outStyles).toContain('w:styleId="ContactNote"');
    expect(outNumbering).not.toContain('555-123-4567');
    expect(report.bodyPII).toBe(2);
    expect(report.propsPII).toBe(0);
  });

  it('closes split-run evasion: PII straddling adjacent w:t is redacted with honest counts', () => {
    const docXml =
      `<w:document xmlns:w="${W_MAIN}"><w:body>` +
      `<w:p><w:r><w:t>Contact jane.doe@exa</w:t></w:r><w:r><w:t>mple.com today</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>SSN 123-</w:t></w:r><w:r><w:t>45-</w:t></w:r><w:r><w:t>6789 filed</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const r = redactBodyPiiInWordText(docXml, true);
    expect(r.count).toBe(2);
    expect(r.xml).not.toContain('jane.doe@exa');
    expect(r.xml).not.toContain('mple.com');
    expect(r.xml).not.toContain('123-');
    expect(r.xml).not.toContain('6789');
    expect(r.xml).toContain('[redacted]');
    // Rendered text carries no PII remnant.
    const rendered = [...r.xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join('');
    expect(rendered).not.toMatch(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    expect(rendered).not.toMatch(/\d{3}-\d{2}-\d{4}/);
  });

  it('split-run scan counts without mutating when redact is false', () => {
    const docXml =
      `<w:document xmlns:w="${W_MAIN}"><w:body>` +
      `<w:p><w:r><w:t>mail jane@exa</w:t></w:r><w:r><w:t>mple.com</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const r = redactBodyPiiInWordText(docXml, false);
    expect(r.count).toBe(1);
    expect(r.xml).toBe(docXml);
  });

  it('single-run PII keeps legacy counts (no double counting across passes)', () => {
    const docXml =
      `<w:document xmlns:w="${W_MAIN}"><w:body>` +
      `<w:p><w:r><w:t>reach partner.jane@biglaw.example or 212-555-0147</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    expect(redactBodyPiiInWordText(docXml, true).count).toBe(2);
    const r = sanitizeDocumentXml(docXml, {}, 'word/document.xml');
    expect(r.bodyPII).toBe(2);
  });

  it('core.xml under zero-retention empties creator/modified/revision instead of restamping', () => {
    const core =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
      `xmlns:dcterms="http://purl.org/dc/terms/">` +
      `<dc:creator>Alice Doe</dc:creator>` +
      `<cp:lastModifiedBy>Bob Smith</cp:lastModifiedBy>` +
      `<dcterms:created>2026-01-01T00:00:00Z</dcterms:created>` +
      `<dcterms:modified>2026-02-01T00:00:00Z</dcterms:modified>` +
      `<cp:revision>7</cp:revision>` +
      `<dc:title>Deal notes</dc:title>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, { zeroRetention: true });
    expect(r.purgedAuthors).toBe(6); // 2 author tags + created/modified/revision/title
    expect(r.xml).toContain('<dc:creator></dc:creator>');
    expect(r.xml).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');
    expect(r.xml).toContain('<dcterms:modified></dcterms:modified>');
    expect(r.xml).toContain('<cp:revision></cp:revision>');
    expect(r.xml).not.toContain('Alice Doe');
    expect(r.xml).not.toContain('Bob Smith');
    expect(r.xml).not.toContain('2026-02-01');
    expect(r.xml).not.toContain('ASTra');
  });

  it('core.xml without zero-retention empties authors and keeps titles', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:creator>Alice Doe</dc:creator>` +
      `<cp:lastModifiedBy>Bob Smith</cp:lastModifiedBy>` +
      `<cp:revision>7</cp:revision>` +
      `<dc:title>Public Report</dc:title>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, { zeroRetention: false });
    expect(r.xml).toContain('<dc:creator></dc:creator>');
    expect(r.xml).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');
    expect(r.xml).not.toContain('Alice Doe');
    expect(r.xml).not.toContain('Bob Smith');
    expect(r.xml).toContain('<cp:revision>7</cp:revision>');
    expect(r.xml).toContain('<dc:title>Public Report</dc:title>');
  });

  it('end-to-end report keeps an honest props-vs-body split with no original values echoed', () => {
    const buf = packageBuffer({
      'word/document.xml':
        `<w:document xmlns:w="${W_MAIN}"><w:body>` +
        `<w:p><w:r><w:t>mail jane@exa</w:t></w:r><w:r><w:t>mple.com split</w:t></w:r></w:p>` +
        `</w:body></w:document>`,
      'word/comments.xml':
        `<w:comments xmlns:w="${W_MAIN}">` +
        `<w:comment w:id="0" w:author="R"><w:p><w:r><w:t>ssn 123-45-6789</w:t></w:r></w:p></w:comment>` +
        `</w:comments>`,
      'docProps/core.xml':
        `<?xml version="1.0" encoding="UTF-8"?>` +
        `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
        `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">` +
        `<dc:creator>leak-creator@example.com</dc:creator>` +
        `<dcterms:modified>2026-02-01T00:00:00Z</dcterms:modified>` +
        `<cp:revision>9</cp:revision>` +
        `</cp:coreProperties>`,
      '[Content_Types].xml': TYPES_XML,
    });
    const { buffer, report } = sanitizeDocxBuffer(buf, { zeroRetention: true });
    expect(report.ok).toBe(true);
    expect(report.bodyPII).toBe(2);
    expect(report.propsPII).toBe(1);
    expect(report.purgedPII).toBe(report.bodyPII + report.propsPII);
    expect(JSON.stringify(report)).not.toContain('jane@example.com');
    expect(JSON.stringify(report)).not.toContain('leak-creator@example.com');
    expect(JSON.stringify(report)).not.toContain('123-45-6789');
    const out = new AdmZip(buffer);
    expect(out.readAsText('word/document.xml')).not.toContain('mple.com');
    expect(out.readAsText('word/comments.xml')).not.toContain('123-45-6789');
    const coreOut = out.readAsText('docProps/core.xml');
    expect(coreOut).not.toContain('leak-creator@example.com');
    expect(coreOut).toContain('<cp:revision></cp:revision>');
  });
});

describe('Input gates follow-ups: over-blocking regressions', () => {
  const TIFF_LE = Buffer.concat([
    Buffer.from([0x49, 0x49, 0x2a, 0x00, 0x08, 0x00, 0x00, 0x00]),
    Buffer.alloc(20),
  ]).toString('base64');
  const TIFF_BE = Buffer.concat([
    Buffer.from([0x4d, 0x4d, 0x00, 0x2a, 0x00, 0x00, 0x00, 0x08]),
    Buffer.alloc(20),
  ]).toString('base64');
  const SVG = Buffer.from(
    '<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>'
  ).toString('base64');
  const SVG_PROLOG = Buffer.from(
    '<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg"><rect width="1" height="1"/></svg>'
  ).toString('base64');
  const SVG_XXE = Buffer.from(
    '<?xml version="1.0"?><!DOCTYPE svg [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>' +
      '<svg xmlns="http://www.w3.org/2000/svg"><rect width="1" height="1"/></svg>'
  ).toString('base64');
  const PNG_1PX =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

  it('accepts TIFF magic (little- and big-endian) via data URL and raw base64', () => {
    for (const data of [
      `data:image/tiff;base64,${TIFF_LE}`,
      `data:image/tiff;base64,${TIFF_BE}`,
      TIFF_LE,
      TIFF_BE,
    ]) {
      const ast = createMockAST();
      expect(() =>
        executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: data }, ast)
      ).not.toThrow();
      const keys = Object.keys(ast.mediaFiles || {});
      expect(keys.length).toBe(1);
      expect(keys[0]).toMatch(/\.tiff$/);
    }
  });

  it('accepts SVG content but keeps SVG XXE fail-closed', () => {
    for (const data of [`data:image/svg+xml;base64,${SVG}`, SVG_PROLOG]) {
      const ast = createMockAST();
      expect(() =>
        executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: data }, ast)
      ).not.toThrow();
      expect(Object.keys(ast.mediaFiles || {})[0]).toMatch(/\.svg$/);
    }
    const evil = createMockAST();
    expect(() =>
      executeImage(
        { action: 'image', operation: 'insert', anchor: 'p_001', fileData: `data:image/svg+xml;base64,${SVG_XXE}` },
        evil
      )
    ).toThrowError(/E_VALIDATION.*XXE/);
  });

  it('accepts case-insensitive DATA: URI prefixes while attacks stay blocked', () => {
    for (const data of [
      `DATA:image/png;base64,${PNG_1PX}`,
      `Data:image/png;base64,${PNG_1PX}`,
      `data:IMAGE/PNG;BASE64,${PNG_1PX}`,
    ]) {
      const ast = createMockAST();
      expect(() =>
        executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: data }, ast)
      ).not.toThrow();
    }
    const pathAst = createMockAST();
    expect(() =>
      executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: '/etc/passwd' }, pathAst)
    ).toThrowError(/E_VALIDATION.*Provide image as data: URL or base64 string/);
    const badAst = createMockAST();
    const bad = Buffer.from('NOT_AN_IMAGE_HEADER').toString('base64');
    expect(() =>
      executeImage(
        { action: 'image', operation: 'insert', anchor: 'p_001', fileData: `DATA:image/png;base64,${bad}` },
        badAst
      )
    ).toThrowError(/E_VALIDATION.*magic bytes mismatch/);
  });

  it('treats legit dotted filenames as benign while traversal stays rogue', () => {
    expect(isRogueZipEntry('word/media/my..photo.png')).toBe(false);
    expect(isRogueZipEntry('word/media/report...final.docx')).toBe(false);
    expect(isRogueZipEntry('word/document.xml')).toBe(false);
    expect(isRogueZipEntry('../evil.txt')).toBe(true);
    expect(isRogueZipEntry('word/../../evil.txt')).toBe(true);
    expect(isRogueZipEntry('..\\evil.txt')).toBe(true);
    expect(isRogueZipEntry('/etc/passwd')).toBe(true);
    expect(isRogueZipEntry('C:\\windows\\evil.exe')).toBe(true);

    const files = new Map<string, Buffer>([
      ['word/media/my..photo.png', Buffer.from([0x89, 0x50, 0x4e, 0x47])],
      ['word/document.xml', Buffer.from(`<w:document xmlns:w="${W_MAIN}"><w:body/></w:document>`, 'utf8')],
    ]);
    const { files: clean } = sanitizePackageEntries(files, {});
    expect(clean.has('word/media/my..photo.png')).toBe(true);
  });

  it('accepts tel:/sms: and relative anchor URLs while script schemes stay blocked', () => {
    for (const url of [
      'tel:+1-555-123-4567',
      'sms:+15551234567',
      '#myAnchor',
      'page.html',
      './docs/page.html',
      '/abs/path/page',
      '?q=1',
    ]) {
      const ast = createMockAST();
      expect(() => executeHyperlink({ action: 'hyperlink', target: 'p_001', url }, ast)).not.toThrow();
      const targets = Object.values(ast.relationships || {}).map((r) => (r as { target: string }).target);
      expect(targets).toContain(url);
    }
    for (const url of [
      'javascript:alert(1)',
      'JAVASCRIPT:alert(document.cookie)',
      'data:text/html,<script>alert(1)</script>',
      'vbscript:msgbox(1)',
      'file:///etc/passwd',
    ]) {
      const ast = createMockAST();
      expect(() => executeHyperlink({ action: 'hyperlink', target: 'p_001', url }, ast)).toThrowError(
        /E_VALIDATION.*Hyperlink URL must use http, https, mailto, or ftp scheme/
      );
    }
  });
});
