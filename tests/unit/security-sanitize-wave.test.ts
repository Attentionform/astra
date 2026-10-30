import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocumentAST,
  executeImage,
  executeHyperlink,
  DocxParser,
  sanitizeDocxBuffer,
  sanitizeCorePropertiesXml,
  parseCorePropertiesXml,
  assertNoForbiddenMarkup,
  isRogueZipEntry,
  redactBodyPiiInWordText,
} from '@attentionform/astra-engine';

const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

// Minimal valid JPEG (SOI + APP0 + SOF0 + SOS + EOI)
const MINIMAL_JPEG_BASE64 = Buffer.from([
  0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x01, 0x00, 0x48,
  0x00, 0x48, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x08, 0x06, 0x06, 0x07, 0x06, 0x05, 0x08,
  0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0a, 0x0c, 0x14, 0x0d, 0x0c, 0x0b, 0x0b, 0x0c, 0x19, 0x12,
  0x13, 0x0f, 0x14, 0x1d, 0x1a, 0x1f, 0x1e, 0x1d, 0x1a, 0x1c, 0x1c, 0x20, 0x24, 0x2e, 0x27, 0x20,
  0x22, 0x2c, 0x23, 0x1c, 0x1c, 0x28, 0x37, 0x29, 0x2c, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1f, 0x27,
  0x39, 0x3d, 0x38, 0x32, 0x3c, 0x2e, 0x33, 0x34, 0x32, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01,
  0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xff, 0xc4, 0x00, 0x1f, 0x00, 0x00, 0x01, 0x05, 0x01, 0x01,
  0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04,
  0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f,
  0x00, 0x7f, 0x00, 0xff, 0xd9,
]).toString('base64');

// Minimal valid GIF (GIF89a 1x1)
const MINIMAL_GIF_BASE64 = Buffer.from(
  'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
  'base64'
).toString('base64');

// Minimal valid WebP (RIFF....WEBPVP8 )
const MINIMAL_WEBP_BASE64 = Buffer.from([
  0x52, 0x49, 0x46, 0x46, 0x1a, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50, 0x56, 0x50, 0x38, 0x20,
  0x0e, 0x00, 0x00, 0x00, 0x30, 0x01, 0x00, 0x9d, 0x01, 0x2a, 0x01, 0x00, 0x01, 0x00, 0x02, 0x00,
  0x34, 0x25, 0xa4, 0x00,
]).toString('base64');

// Minimal valid BMP (BM header 14 bytes)
const MINIMAL_BMP_BASE64 = Buffer.from([
  0x42, 0x4d, 0x3a, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x36, 0x00, 0x00, 0x00,
]).toString('base64');

function createMockAST(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
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

function buildCleanDocxZip(): AdmZip {
  const zip = new AdmZip();
  const docXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t>Clean content</w:t></w:r></w:p></w:body></w:document>`;
  const coreXml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
    `<dc:title>Test Title</dc:title>` +
    `<dc:creator>Test Author</dc:creator>` +
    `</cp:coreProperties>`;
  const typesXml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `</Types>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile('docProps/core.xml', Buffer.from(coreXml, 'utf8'));
  zip.addFile('[Content_Types].xml', Buffer.from(typesXml, 'utf8'));
  return zip;
}

describe('Wave 1A Security & Sanitization Gates', () => {
  describe('P0-1: Image Path Traversal Prevention & Magic Bytes', () => {
    it('rejects filesystem paths (/etc/passwd) with E_VALIDATION and remedy', () => {
      const ast = createMockAST();
      expect(() =>
        executeImage(
          {
            action: 'image',
            operation: 'insert',
            anchor: 'p_001',
            fileData: '/etc/passwd',
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Provide image as data: URL or base64 string/);
    });

    it('rejects relative file paths (.env) with E_VALIDATION and remedy', () => {
      const ast = createMockAST();
      expect(() =>
        executeImage(
          {
            action: 'image',
            operation: 'insert',
            anchor: 'p_001',
            fileData: '../../.env',
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Provide image as data: URL or base64 string/);
    });

    it('rejects base64 of non-image payload (invalid magic bytes)', () => {
      const ast = createMockAST();
      const fakeBase64 = Buffer.from('this is not an image').toString('base64');
      expect(() =>
        executeImage(
          {
            action: 'image',
            operation: 'insert',
            anchor: 'p_001',
            fileData: fakeBase64,
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Provide image as data: URL or base64 string/);
    });

    it('rejects data URL containing invalid magic bytes', () => {
      const ast = createMockAST();
      const fakeBase64 = Buffer.from('NOT_AN_IMAGE_HEADER').toString('base64');
      expect(() =>
        executeImage(
          {
            action: 'image',
            operation: 'insert',
            anchor: 'p_001',
            fileData: `data:image/png;base64,${fakeBase64}`,
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Provide image as data: URL or base64 string/);
    });

    it('accepts data URLs and raw base64 with valid magic bytes (PNG, JPEG, GIF, WebP, BMP)', () => {
      const formats = [
        { name: 'PNG (data URL)', data: `data:image/png;base64,${PNG_1PX}` },
        { name: 'PNG (raw base64)', data: PNG_1PX },
        { name: 'JPEG (data URL)', data: `data:image/jpeg;base64,${MINIMAL_JPEG_BASE64}` },
        { name: 'JPEG (raw base64)', data: MINIMAL_JPEG_BASE64 },
        { name: 'GIF (data URL)', data: `data:image/gif;base64,${MINIMAL_GIF_BASE64}` },
        { name: 'WebP (data URL)', data: `data:image/webp;base64,${MINIMAL_WEBP_BASE64}` },
        { name: 'BMP (data URL)', data: `data:image/bmp;base64,${MINIMAL_BMP_BASE64}` },
      ];

      for (const fmt of formats) {
        const ast = createMockAST();
        expect(() =>
          executeImage(
            {
              action: 'image',
              operation: 'insert',
              anchor: 'p_001',
              fileData: fmt.data,
            },
            ast
          )
        ).not.toThrow();
        expect(ast.sections[0].nodes.length).toBe(2);
        expect(ast.sections[0].nodes[1].type).toBe('image');
      }
    });

    it('prefers magic bytes over declared MIME in data URL when they diverge', () => {
      const ast = createMockAST();
      executeImage(
        {
          action: 'image',
          operation: 'insert',
          anchor: 'p_001',
          fileData: `data:image/jpeg;base64,${PNG_1PX}`, // declared jpeg, actual png
        },
        ast
      );
      const mediaKeys = Object.keys(ast.mediaFiles || {});
      expect(mediaKeys.length).toBe(1);
      expect(mediaKeys[0]).toMatch(/\.png$/);
    });
  });

  describe('P0-11: Hyperlink Scheme Allowlist', () => {
    it('rejects javascript:alert(1) with E_VALIDATION and remedy', () => {
      const ast = createMockAST();
      expect(() =>
        executeHyperlink(
          {
            action: 'hyperlink',
            target: 'p_001',
            url: 'javascript:alert(1)',
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Hyperlink URL must use http, https, mailto, or ftp scheme/);
    });

    it('rejects JAVASCRIPT: (case-insensitive) with E_VALIDATION', () => {
      const ast = createMockAST();
      expect(() =>
        executeHyperlink(
          {
            action: 'hyperlink',
            target: 'p_001',
            url: 'JAVASCRIPT:alert(document.cookie)',
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Hyperlink URL must use http, https, mailto, or ftp scheme/);
    });

    it('rejects data: URLs with E_VALIDATION', () => {
      const ast = createMockAST();
      expect(() =>
        executeHyperlink(
          {
            action: 'hyperlink',
            target: 'p_001',
            url: 'data:text/html,<script>alert(1)</script>',
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Hyperlink URL must use http, https, mailto, or ftp scheme/);
    });

    it('rejects vbscript: with E_VALIDATION', () => {
      const ast = createMockAST();
      expect(() =>
        executeHyperlink(
          {
            action: 'hyperlink',
            target: 'p_001',
            url: 'vbscript:msgbox(1)',
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Hyperlink URL must use http, https, mailto, or ftp scheme/);
    });

    it('rejects file: protocol with E_VALIDATION', () => {
      const ast = createMockAST();
      expect(() =>
        executeHyperlink(
          {
            action: 'hyperlink',
            target: 'p_001',
            url: 'file:///etc/passwd',
          },
          ast
        )
      ).toThrowError(/E_VALIDATION.*Hyperlink URL must use http, https, mailto, or ftp scheme/);
    });

    it('accepts allowed protocols: http, https, mailto, ftp', () => {
      const allowed = [
        'http://example.com',
        'https://example.com/page',
        'mailto:user@example.com',
        'ftp://files.example.com/doc.pdf',
      ];
      for (const url of allowed) {
        const ast = createMockAST();
        expect(() =>
          executeHyperlink(
            {
              action: 'hyperlink',
              target: 'p_001',
              url,
            },
            ast
          )
        ).not.toThrow();
      }
    });

    it('allows internal bookmark hyperlinks without URL scheme check', () => {
      const ast = createMockAST();
      expect(() =>
        executeHyperlink(
          {
            action: 'hyperlink',
            target: 'p_001',
            bookmarkName: 'InternalSection',
          },
          ast
        )
      ).not.toThrow();
    });
  });

  describe('P0-6 & L-20: Case-Insensitive XXE / DTD Detector', () => {
    it('parser fails closed on lowercase <!doctype and <!entity', () => {
      const zip = buildCleanDocxZip();
      const evilDocXml = `<?xml version="1.0"?><!doctype doc [<!entity xxe "pwn">]><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>`;
      zip.updateFile('word/document.xml', Buffer.from(evilDocXml, 'utf8'));

      expect(() => new DocxParser(zip.toBuffer())).toThrowError(
        /Security Error: XXE \/ DTD entity declarations are strictly prohibited/
      );
    });

    it('parser reloadDocumentXml rejects case-insensitive <!DocType', () => {
      const zip = buildCleanDocxZip();
      const parser = new DocxParser(zip.toBuffer());
      expect(() =>
        parser.reloadDocumentXml('<!DocType root SYSTEM "http://evil.com"><w:document/>')
      ).toThrowError(/Security Error: XXE \/ DTD entity declarations are strictly prohibited/);
    });

    it('assertNoForbiddenMarkup detects <!doctype, <!entity, and uppercase variations', () => {
      expect(() => assertNoForbiddenMarkup('<!doctype test>')).toThrowError(/E_VALIDATION/);
      expect(() => assertNoForbiddenMarkup('<!DOCTYPE test>')).toThrowError(/E_VALIDATION/);
      expect(() => assertNoForbiddenMarkup('<!entity test "foo">')).toThrowError(/E_VALIDATION/);
      expect(() => assertNoForbiddenMarkup('<!ENTITY test "foo">')).toThrowError(/E_VALIDATION/);
      expect(() => assertNoForbiddenMarkup('<w:document><w:body/></w:document>')).not.toThrow();
    });

    it('sanitizeDocxBuffer fails closed on ingress XXE in any XML part', () => {
      const zip = buildCleanDocxZip();
      zip.addFile(
        'word/comments.xml',
        Buffer.from('<?xml version="1.0"?><!doctype root SYSTEM "http://attacker.com"><w:comments/>', 'utf8')
      );
      expect(() => sanitizeDocxBuffer(zip.toBuffer())).toThrowError(/E_VALIDATION.*XXE/);
    });

    it('parseCorePropertiesXml fails closed on case-insensitive <!doctype and <!entity', () => {
      expect(() => parseCorePropertiesXml('<?xml version="1.0"?><!doctype doc><cp:coreProperties/>')).toThrowError(
        /Security Error: XXE \/ DTD entity declarations are strictly prohibited/
      );
      expect(() => parseCorePropertiesXml('<?xml version="1.0"?><!DocType doc><cp:coreProperties/>')).toThrowError(
        /Security Error: XXE \/ DTD entity declarations are strictly prohibited/
      );
      expect(() => parseCorePropertiesXml('<?xml version="1.0"?><!entity test "foo"><cp:coreProperties/>')).toThrowError(
        /Security Error: XXE \/ DTD entity declarations are strictly prohibited/
      );
      expect(() => parseCorePropertiesXml('<?xml version="1.0"?><!Entity test "foo"><cp:coreProperties/>')).toThrowError(
        /Security Error: XXE \/ DTD entity declarations are strictly prohibited/
      );
    });
  });

  describe('P0-6: Rogue Zip Entry Defense', () => {
    it('isRogueZipEntry detects traversal, absolute paths, drive letters, and control characters', () => {
      expect(isRogueZipEntry('../evil.txt')).toBe(true);
      expect(isRogueZipEntry('word/../../evil.txt')).toBe(true);
      expect(isRogueZipEntry('/etc/passwd')).toBe(true);
      expect(isRogueZipEntry('  /etc/passwd')).toBe(true);
      expect(isRogueZipEntry('\\Windows\\System32')).toBe(true);
      expect(isRogueZipEntry('  \\Windows\\System32')).toBe(true);
      expect(isRogueZipEntry('C:\\test.txt')).toBe(true);
      expect(isRogueZipEntry('  C:\\test.txt')).toBe(true);
      expect(isRogueZipEntry('d:/evil.txt')).toBe(true);
      expect(isRogueZipEntry('  d:/evil.txt')).toBe(true);
      expect(isRogueZipEntry('word/\x00evil.xml')).toBe(true);

      expect(isRogueZipEntry('word/document.xml')).toBe(false);
      expect(isRogueZipEntry('docProps/core.xml')).toBe(false);
      expect(isRogueZipEntry('[Content_Types].xml')).toBe(false);
    });

    it('sanitizeDocxBuffer strips rogue entries while preserving valid parts', () => {
      const zip = buildCleanDocxZip();
      zip.addFile('/etc/passwd', Buffer.from('root:x:0:0:root:/root:/bin/bash'));
      zip.addFile('../../secret.env', Buffer.from('API_KEY=secret'));
      zip.addFile('C:\\windows\\system32\\cmd.exe', Buffer.from('binary-content'));

      const { buffer, report } = sanitizeDocxBuffer(zip.toBuffer());
      expect(report.ok).toBe(true);

      const outZip = new AdmZip(buffer);
      expect(outZip.getEntry('/etc/passwd')).toBeNull();
      expect(outZip.getEntry('../../secret.env')).toBeNull();
      expect(outZip.getEntry('C:\\windows\\system32\\cmd.exe')).toBeNull();
      expect(outZip.getEntry('word/document.xml')).not.toBeNull();
    });
  });

  describe('P0-7: Body PII Scan & Core Properties Purge', () => {
    it('purges dc:title, dc:subject, cp:keywords, dc:description, cp:category under zeroRetention: true', () => {
      const coreXml =
        `<?xml version="1.0" encoding="UTF-8"?>` +
        `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
        `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
        `<dc:title>Merger Agreement CONFIDENTIAL</dc:title>` +
        `<dc:subject>Project Pegasus</dc:subject>` +
        `<cp:keywords>acquisition, confidential, M&amp;A</cp:keywords>` +
        `<dc:description>Private deal notes</dc:description>` +
        `<cp:category>Top Secret</cp:category>` +
        `<dc:creator>Alice Doe</dc:creator>` +
        `</cp:coreProperties>`;

      const r = sanitizeCorePropertiesXml(coreXml, { zeroRetention: true });
      expect(r.xml).not.toContain('Merger Agreement CONFIDENTIAL');
      expect(r.xml).not.toContain('Project Pegasus');
      expect(r.xml).not.toContain('acquisition');
      expect(r.xml).not.toContain('Private deal notes');
      expect(r.xml).not.toContain('Top Secret');
    });

    it('preserves dc:title when zeroRetention is false and no PII pattern matches', () => {
      const coreXml =
        `<?xml version="1.0" encoding="UTF-8"?>` +
        `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
        `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
        `<dc:title>Public Annual Report</dc:title>` +
        `</cp:coreProperties>`;

      const r = sanitizeCorePropertiesXml(coreXml, { zeroRetention: false });
      expect(r.xml).toContain('<dc:title>Public Annual Report</dc:title>');
    });

    it('scans and redacts body text PII (email, phone, SSN, credit card) inside <w:t>', () => {
      const docXml =
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:body>` +
        `<w:p><w:r><w:t>Contact Jane at partner.jane@biglaw.example or 212-555-0147</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>Confidential SSN: 123-45-6789 and Card: 4111-1111-1111-1111</w:t></w:r></w:p>` +
        `</w:body></w:document>`;

      const r = redactBodyPiiInWordText(docXml, true);
      expect(r.count).toBe(4);
      expect(r.xml).not.toContain('partner.jane@biglaw.example');
      expect(r.xml).not.toContain('212-555-0147');
      expect(r.xml).not.toContain('123-45-6789');
      expect(r.xml).not.toContain('4111-1111-1111-1111');
      expect(r.xml).toContain('[redacted]');
    });

    it('sanitizeDocxBuffer records bodyPII, propsPII, and purgedPII = bodyPII + propsPII', () => {
      const zip = new AdmZip();
      const docXml =
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:body><w:p><w:r><w:t>Employee SSN: 123-45-6789</w:t></w:r></w:p></w:body></w:document>`;
      const coreXml =
        `<?xml version="1.0" encoding="UTF-8"?>` +
        `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
        `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
        `<dc:creator>author@company.org</dc:creator>` +
        `</cp:coreProperties>`;
      zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
      zip.addFile('docProps/core.xml', Buffer.from(coreXml, 'utf8'));
      zip.addFile(
        '[Content_Types].xml',
        Buffer.from('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8')
      );

      const { buffer, report } = sanitizeDocxBuffer(zip.toBuffer(), { zeroRetention: true });
      expect(report.ok).toBe(true);
      expect(report.bodyPII).toBe(1);
      expect(report.propsPII).toBe(1);
      expect(report.purgedPII).toBe(2);

      const outZip = new AdmZip(buffer);
      const cleanDoc = outZip.readAsText('word/document.xml');
      expect(cleanDoc).not.toContain('123-45-6789');
      expect(cleanDoc).toContain('[redacted]');
    });
  });
});
