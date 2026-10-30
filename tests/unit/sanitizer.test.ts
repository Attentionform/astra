import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  isBlockedExecutableEntry,
  isZeroRetentionEnabled,
  redactAttributePii,
  redactBodyPiiInWordText,
  redactDrawingAltPii,
  redactTextNodePii,
  sanitizeAppPropertiesXml,
  sanitizeCorePropertiesXml,
  sanitizeCustomPropertiesXml,
  sanitizeDocumentXml,
  sanitizeDocxBuffer,
  sanitizeEmbeddedWorkbookBytes,
  sanitizePackageEntries,
  stripJpegExif,
  stripPngTextChunks,
  validateSanitizeOptionsStrict,
} from '@attentionform/astra-engine';

function buildJpegWithExif(): Buffer {
  // SOI + APP1(Exif) + APP13 + SOS + minimal scan + EOI.
  const soi = Buffer.from([0xff, 0xd8]);
  const exifPayload = Buffer.concat([Buffer.from('Exif\0\0', 'binary'), Buffer.alloc(32, 0x41)]);
  const app1 = Buffer.concat([
    Buffer.from([0xff, 0xe1]),
    Buffer.alloc(2),
    exifPayload,
  ]);
  app1.writeUInt16BE(exifPayload.length + 2, 2);
  const psPayload = Buffer.from('Photoshop 3.0\0author-pii', 'binary');
  const app13 = Buffer.concat([Buffer.from([0xff, 0xed]), Buffer.alloc(2), psPayload]);
  app13.writeUInt16BE(psPayload.length + 2, 2);
  const sos = Buffer.from([0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f, 0x00]);
  const eoi = Buffer.from([0xff, 0xd9]);
  return Buffer.concat([soi, app1, app13, sos, Buffer.from([0x00]), eoi]);
}

function buildPngWithAuthor(): Buffer {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const chunk = (type: string, data: Buffer): Buffer => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length, 0);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    // CRC is not validated by the sanitizer — zeros suffice for the unit.
    return Buffer.concat([len, body, Buffer.alloc(4)]);
  };
  const ihdr = chunk('IHDR', Buffer.alloc(13));
  const text = chunk('tEXt', Buffer.from('Author\0Jane Doe <jane@example.com>', 'latin1'));
  const idat = chunk('IDAT', Buffer.from([0x00]));
  const iend = chunk('IEND', Buffer.alloc(0));
  return Buffer.concat([sig, ihdr, text, idat, iend]);
}

function buildDirtyDocx(): Buffer {
  const zip = new AdmZip();
  const docXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>` +
    `<w:p w:rsidR="00A1B2C3" w:rsidP="00D4E5F6"><w:r><w:t>Hello</w:t></w:r>` +
    `<w:ins w:id="1" w:author="Jane Doe" w:initials="JD" w:date="2026-01-01T00:00:00Z"><w:r><w:t> tracked</w:t></w:r></w:ins></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>` +
    `</w:body></w:document>`;
  const coreXml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
    `<dc:creator>jane.doe@example.com</dc:creator>` +
    `<cp:lastModifiedBy>Jane Doe +1 (555) 123-4567</cp:lastModifiedBy>` +
    `</cp:coreProperties>`;
  const appXml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
    `<Company>Acme Corp jane@example.com</Company><Manager>Jane Doe</Manager>` +
    `</Properties>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile('docProps/core.xml', Buffer.from(coreXml, 'utf8'));
  zip.addFile('docProps/app.xml', Buffer.from(appXml, 'utf8'));
  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
  zip.addFile('word/media/image1.jpg', buildJpegWithExif());
  zip.addFile('word/media/image2.png', buildPngWithAuthor());
  return zip.toBuffer();
}

describe('Sprint-3 sanitizer: RSID/author/EXIF/PII purge + zero-retention', () => {
  it('purges RSIDs and authors from document.xml with counts', () => {
    const xml =
      `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidR="00A1" w:rsidP="00B2">` +
      `<w:ins w:author="Evil Author" w:initials="EA"><w:r><w:t>x</w:t></w:r></w:ins></w:p>`;
    const r = sanitizeDocumentXml(xml, { replacementAuthor: 'ASTra' });
    expect(r.removedRSIDs).toBe(2);
    expect(r.purgedAuthors).toBe(2);
    expect(r.xml).not.toContain('rsid');
    expect(r.xml).not.toContain('Evil Author');
    expect(r.xml).toContain('w:author="ASTra"');
  });

  it('purges core/app authors and redacts PII emails + phones', () => {
    // Note: PII inside pure author fields (creator/lastModifiedBy) is counted
    // then overwritten by the author replacement — so the surviving
    // '[redacted]' marker is asserted on dc:description (a non-author field).
    const core = `<cp:coreProperties xmlns:cp="x" xmlns:dc="y"><dc:creator>jane@example.com</dc:creator><cp:lastModifiedBy>Jane +1 (555) 123-4567</cp:lastModifiedBy><dc:description>Contact jane@example.com</dc:description></cp:coreProperties>`;
    const rc = sanitizeCorePropertiesXml(core, {});
    expect(rc.purgedAuthors).toBe(2);
    expect(rc.purgedPII).toBeGreaterThanOrEqual(3);
    expect(rc.xml).not.toContain('jane@example.com');
    expect(rc.xml).not.toContain('(555)');
    expect(rc.xml).toContain('[redacted]');

    const app = `<Properties xmlns="z"><Company>Acme jane@example.com</Company></Properties>`;
    const ra = sanitizeAppPropertiesXml(app, {});
    expect(ra.purgedAuthors).toBe(1);
    expect(ra.purgedPII).toBe(1);
    expect(ra.xml).not.toContain('jane@example.com');
  });

  it('strips JPEG APP1/APP13 EXIF and PNG author text chunks', () => {
    const jpeg = buildJpegWithExif();
    const rj = stripJpegExif(jpeg);
    expect(rj.strippedBytes).toBeGreaterThan(0);
    expect(rj.buf.length).toBeLessThan(jpeg.length);
    expect(rj.buf.subarray(0, 2).equals(Buffer.from([0xff, 0xd8]))).toBe(true);

    const png = buildPngWithAuthor();
    const rp = stripPngTextChunks(png);
    expect(rp.strippedBytes).toBeGreaterThan(0);
    expect(rp.buf.length).toBeLessThan(png.length);
    expect(rp.buf.toString('latin1')).not.toContain('Jane Doe');
  });

  it('sanitizeDocxBuffer end-to-end purges all carriers and stays loadable', () => {
    const dirty = buildDirtyDocx();
    const { buffer, report } = sanitizeDocxBuffer(dirty, { zeroRetention: true });
    expect(report.ok).toBe(true);
    expect(report.zeroRetention).toBe(true);
    expect(report.removedRSIDs).toBeGreaterThanOrEqual(2);
    expect(report.purgedAuthors).toBeGreaterThanOrEqual(3);
    expect(report.strippedExifBytes).toBeGreaterThan(0);
    expect(report.purgedPII).toBeGreaterThanOrEqual(2);
    expect(report.filesTouched.length).toBeGreaterThan(0);
    // Zero-retention: counts only — original PII never echoed in the report.
    expect(JSON.stringify(report)).not.toContain('jane.doe@example.com');
    expect(JSON.stringify(report)).not.toContain('Jane Doe');

    const out = new AdmZip(buffer);
    const docXml = out.readAsText('word/document.xml');
    expect(docXml).not.toContain('rsid');
    expect(docXml).not.toContain('Jane Doe');
    const coreXml = out.readAsText('docProps/core.xml');
    expect(coreXml).not.toContain('jane.doe@example.com');

    // Sanitized package still parses + passes integrity.
    const parser = new DocxParser(buffer);
    expect(parser.validateZipIntegrity().isValid).toBe(true);
    expect(parser.readDocumentXml()).toContain('Hello');
  });

  it('sanitizePackageEntries never mutates its input map', () => {
    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from('<w:p xmlns:w="x" w:rsidR="00A1"><w:r><w:t>t</w:t></w:r></w:p>', 'utf8')],
    ]);
    const before = files.get('word/document.xml')!.toString('utf8');
    const { files: clean } = sanitizePackageEntries(files, {});
    expect(files.get('word/document.xml')!.toString('utf8')).toBe(before);
    expect(clean.get('word/document.xml')!.toString('utf8')).not.toContain('rsid');
  });

  it('strict zod rejects unknown sanitizer keys', () => {
    expect(() => validateSanitizeOptionsStrict({ purgeRSIDs: true, evil: 1 })).toThrow(
      /E_VALIDATION/
    );
    expect(validateSanitizeOptionsStrict({}).ok).toBe(true);
  });

  it('zero-retention flag resolves from env when unset', () => {
    const saved = process.env.ASTRA_ZERO_RETENTION;
    try {
      process.env.ASTRA_ZERO_RETENTION = '1';
      expect(isZeroRetentionEnabled({})).toBe(true);
      expect(isZeroRetentionEnabled({ zeroRetention: false })).toBe(false);
      process.env.ASTRA_ZERO_RETENTION = '0';
      expect(isZeroRetentionEnabled({})).toBe(false);
    } finally {
      if (saved === undefined) delete process.env.ASTRA_ZERO_RETENTION;
      else process.env.ASTRA_ZERO_RETENTION = saved;
    }
  });

  it('purges word/settings.xml rsid list elements alongside rsid attributes', () => {
    const settings =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:rsids><w:rsid w:val="00A1B2C3"/><w:rsid w:val="00D4E5F6"/></w:rsids>` +
      `<w:p w:rsidR="00A1B2C3"><w:r><w:t>s</w:t></w:r></w:p>` +
      `</w:settings>`;
    const r = sanitizeDocumentXml(settings, {});
    expect(r.removedRSIDs).toBe(3);
    expect(r.xml).not.toContain('00A1B2C3');
    expect(r.xml).not.toContain('00D4E5F6');
    expect(r.xml).not.toContain('w:rsidR');
    expect(r.xml).not.toContain('<w:rsid ');
  });

  it('redacts bare US phones (555-123-4567) in core props', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:description>Call 555-123-4567 or 555.123.4567</dc:description>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, {});
    expect(r.purgedPII).toBe(2);
    expect(r.xml).not.toContain('555-123-4567');
    expect(r.xml).not.toContain('555.123.4567');
  });

  it('redacts docProps/custom.xml PII without touching structure', () => {
    const custom =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ` +
      `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Contact">` +
      `<vt:lpwstr>jane@example.com 555-123-4567</vt:lpwstr></property>` +
      `</Properties>`;
    const r = sanitizeCustomPropertiesXml(custom, {});
    expect(r.purgedPII).toBe(2);
    expect(r.purgedAuthors).toBe(0);
    expect(r.xml).not.toContain('jane@example.com');
    expect(r.xml).not.toContain('555-123-4567');
    expect(r.xml).toContain('name="Contact"');
  });

  it('sanitizePackageEntries covers custom.xml + settings.xml + commentsExtended.xml', () => {
    const files = new Map<string, Buffer>([
      [
        'word/settings.xml',
        Buffer.from(
          `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            `<w:rsids><w:rsid w:val="00A1B2C3"/></w:rsids></w:settings>`,
          'utf8'
        ),
      ],
      [
        'word/commentsExtended.xml',
        Buffer.from(
          `<w:commentsExtended xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            `<w:comment w:author="Jane Doe" w:initials="JD" w:rsidR="00A1"/>` +
            `</w:commentsExtended>`,
          'utf8'
        ),
      ],
      [
        'docProps/custom.xml',
        Buffer.from(
          `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">` +
            `<property pid="2" name="Email"><vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">jane@example.com</vt:lpwstr></property>` +
            `</Properties>`,
          'utf8'
        ),
      ],
    ]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    expect(report.removedRSIDs).toBeGreaterThanOrEqual(2);
    expect(report.purgedAuthors).toBeGreaterThanOrEqual(1);
    expect(report.purgedPII).toBeGreaterThanOrEqual(1);
    expect(clean.get('word/settings.xml')!.toString('utf8')).not.toContain('00A1B2C3');
    expect(clean.get('word/commentsExtended.xml')!.toString('utf8')).not.toContain('Jane Doe');
    expect(clean.get('docProps/custom.xml')!.toString('utf8')).not.toContain('jane@example.com');
    expect(report.filesTouched).toContain('word/settings.xml');
    expect(report.filesTouched).toContain('word/commentsExtended.xml');
    expect(report.filesTouched).toContain('docProps/custom.xml');
  });

  it('strips PNG Copyright/Artist/Disclaimer/Source/Warning/Creation Time chunks', () => {
    const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const chunk = (type: string, data: Buffer): Buffer => {
      const len = Buffer.alloc(4);
      len.writeUInt32BE(data.length, 0);
      return Buffer.concat([len, Buffer.from(type, 'ascii'), data, Buffer.alloc(4)]);
    };
    const png = Buffer.concat([
      sig,
      chunk('IHDR', Buffer.alloc(13)),
      chunk('tEXt', Buffer.from('Copyright\0Acme 2026', 'latin1')),
      chunk('tEXt', Buffer.from('Artist\0Jane Doe', 'latin1')),
      chunk('tEXt', Buffer.from('Creation Time\x002026-01-01', 'latin1')),
      chunk('tEXt', Buffer.from('Source\0camera', 'latin1')),
      chunk('IDAT', Buffer.from([0x00])),
      chunk('IEND', Buffer.alloc(0)),
    ]);
    const r = stripPngTextChunks(png);
    expect(r.strippedBytes).toBeGreaterThan(0);
    const text = r.buf.toString('latin1');
    expect(text).not.toContain('Jane Doe');
    expect(text).not.toContain('Copyright');
  });

  it('drops Exif/XMP APP1 but preserves non-Exif APP1 segments', () => {
    const soi = Buffer.from([0xff, 0xd8]);
    const sos = Buffer.from([0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f, 0x00]);
    const eoi = Buffer.from([0xff, 0xd9]);
    const app1Of = (payload: Buffer): Buffer => {
      const seg = Buffer.concat([Buffer.from([0xff, 0xe1]), Buffer.alloc(2), payload]);
      seg.writeUInt16BE(payload.length + 2, 2);
      return seg;
    };
    // Non-Exif APP1 must survive (proves the isExifApp1-only gate, no drop-all).
    const benign = app1Of(Buffer.from('http://example.com/not-exif-payload', 'binary'));
    const benignJpeg = Buffer.concat([soi, benign, sos, Buffer.from([0x00]), eoi]);
    const rb = stripJpegExif(benignJpeg);
    expect(rb.strippedBytes).toBe(0);
    expect(rb.buf.length).toBe(benignJpeg.length);
    // Genuine Exif APP1 must still be stripped.
    const exif = app1Of(
      Buffer.concat([Buffer.from('Exif\0\0', 'binary'), Buffer.alloc(16, 0x41)])
    );
    const exifJpeg = Buffer.concat([soi, exif, sos, Buffer.from([0x00]), eoi]);
    const re = stripJpegExif(exifJpeg);
    expect(re.strippedBytes).toBeGreaterThan(0);
    expect(re.buf.length).toBeLessThan(exifJpeg.length);
  });

  it('redacts SSN (123-45-6789) in core/app/custom props', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:description>SSN 123-45-6789 on file</dc:description>` +
      `</cp:coreProperties>`;
    const rc = sanitizeCorePropertiesXml(core, {});
    expect(rc.purgedPII).toBe(1);
    expect(rc.xml).not.toContain('123-45-6789');
    expect(rc.xml).toContain('[redacted]');

    const app = `<Properties xmlns="z"><Company>SSN 123-45-6789</Company></Properties>`;
    const ra = sanitizeAppPropertiesXml(app, {});
    expect(ra.purgedPII).toBe(1);
    expect(ra.xml).not.toContain('123-45-6789');

    const custom =
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">` +
      `<property pid="2" name="SSN"><vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">123-45-6789</vt:lpwstr></property>` +
      `</Properties>`;
    const rx = sanitizeCustomPropertiesXml(custom, {});
    expect(rx.purgedPII).toBe(1);
    expect(rx.xml).not.toContain('123-45-6789');
  });

  it('redacts credit cards (contiguous + spaced/dashed 13-19 digit) in core/app/custom', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:description>Card 4111111111111111 and 4111 1111 1111 1111 and 4111-1111-1111-1111</dc:description>` +
      `</cp:coreProperties>`;
    const rc = sanitizeCorePropertiesXml(core, {});
    expect(rc.purgedPII).toBe(3);
    expect(rc.xml).not.toContain('4111111111111111');
    expect(rc.xml).not.toContain('4111 1111 1111 1111');
    expect(rc.xml).not.toContain('4111-1111-1111-1111');

    const app = `<Properties xmlns="z"><Company>Pay 5500005555555559</Company></Properties>`;
    const ra = sanitizeAppPropertiesXml(app, {});
    expect(ra.purgedPII).toBe(1);
    expect(ra.xml).not.toContain('5500005555555559');

    const custom =
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">` +
      `<property pid="2" name="Card"><vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">378282246310005</vt:lpwstr></property>` +
      `</Properties>`;
    const rx = sanitizeCustomPropertiesXml(custom, {});
    expect(rx.purgedPII).toBe(1);
    expect(rx.xml).not.toContain('378282246310005');
  });

  it('strips PNG eXIf binary chunk unconditionally', () => {
    const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const chunk = (type: string, data: Buffer): Buffer => {
      const len = Buffer.alloc(4);
      len.writeUInt32BE(data.length, 0);
      return Buffer.concat([len, Buffer.from(type, 'ascii'), data, Buffer.alloc(4)]);
    };
    const exifPayload = Buffer.concat([Buffer.from('Exif\0\0', 'binary'), Buffer.from([0x49, 0x49, 0x2a, 0x00])]);
    const png = Buffer.concat([
      sig,
      chunk('IHDR', Buffer.alloc(13)),
      chunk('eXIf', exifPayload),
      chunk('IDAT', Buffer.from([0x00])),
      chunk('IEND', Buffer.alloc(0)),
    ]);
    const r = stripPngTextChunks(png);
    expect(r.strippedBytes).toBeGreaterThan(0);
    expect(r.buf.length).toBeLessThan(png.length);
    expect(r.buf.toString('latin1')).not.toContain('Exif');
  });

  it('strips PNG Adobe XMP keywords case-insensitively (XML:com.adobe.xmp / XMP)', () => {
    const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const chunk = (type: string, data: Buffer): Buffer => {
      const len = Buffer.alloc(4);
      len.writeUInt32BE(data.length, 0);
      return Buffer.concat([len, Buffer.from(type, 'ascii'), data, Buffer.alloc(4)]);
    };
    const png = Buffer.concat([
      sig,
      chunk('IHDR', Buffer.alloc(13)),
      // Photoshop XMP carrier — mixed case proves case-insensitive lookup.
      chunk('iTXt', Buffer.from('XML:com.adobe.xmp\0xmp-payload-gps', 'latin1')),
      chunk('tEXt', Buffer.from('XMP\0xmp-tracking-id', 'latin1')),
      chunk('IDAT', Buffer.from([0x00])),
      chunk('IEND', Buffer.alloc(0)),
    ]);
    const r = stripPngTextChunks(png);
    expect(r.strippedBytes).toBeGreaterThan(0);
    const text = r.buf.toString('latin1');
    expect(text).not.toContain('xmp-payload-gps');
    expect(text).not.toContain('xmp-tracking-id');
  });

  it('purges <w:rsidRoot> and strips empty <w:rsids> wrapper cleanly', () => {
    const settings =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:rsids><w:rsidRoot w:val="00A1B2C3"/><w:rsid w:val="00D4E5F6"/></w:rsids>` +
      `</w:settings>`;
    const r = sanitizeDocumentXml(settings, {});
    expect(r.removedRSIDs).toBe(2);
    expect(r.xml).not.toContain('rsidRoot');
    expect(r.xml).not.toContain('00A1B2C3');
    expect(r.xml).not.toContain('00D4E5F6');
    expect(r.xml).not.toContain('<w:rsid ');
    // Empty wrapper must be gone entirely — not left as <w:rsids></w:rsids>.
    expect(r.xml).not.toContain('<w:rsids');
    expect(r.xml).not.toContain('</w:rsids>');
  });

  it('purges w15:author (people.xml) alongside w:author', () => {
    const xml =
      `<w:people xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">` +
      `<w15:person w15:author="Jane Doe" w:author="Legacy Author"><w15:presenceInfo/></w15:person>` +
      `</w:people>`;
    const r = sanitizeDocumentXml(xml, { replacementAuthor: 'ASTra' });
    expect(r.purgedAuthors).toBe(2);
    expect(r.xml).not.toContain('Jane Doe');
    expect(r.xml).not.toContain('Legacy Author');
  });

  it('clean package reports all-zero purges WITH scan evidence (trustworthy zeros)', () => {
    const files = new Map<string, Buffer>([
      [
        'word/document.xml',
        Buffer.from(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            `<w:body><w:p><w:r><w:t>Hello clean world</w:t></w:r></w:p>` +
            `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>` +
            `</w:body></w:document>`,
          'utf8'
        ),
      ],
      [
        'docProps/core.xml',
        Buffer.from(
          `<?xml version="1.0" encoding="UTF-8"?>` +
            `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
            `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
            `<dc:title>Clean</dc:title>` +
            `</cp:coreProperties>`,
          'utf8'
        ),
      ],
      [
        'docProps/app.xml',
        Buffer.from(
          `<?xml version="1.0" encoding="UTF-8"?>` +
            `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
            `<Application>ASTra</Application>` +
            `</Properties>`,
          'utf8'
        ),
      ],
      [
        '[Content_Types].xml',
        Buffer.from(
          `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
          'utf8'
        ),
      ],
    ]);
    const { report } = sanitizePackageEntries(files, { zeroRetention: false });
    // All-zero purges — but the scan provably ran:
    expect(report.removedRSIDs).toBe(0);
    expect(report.purgedAuthors).toBe(0);
    expect(report.purgedPII).toBe(0);
    expect(report.acceptedRevisions).toBe(0);
    expect(report.strippedComments).toBe(0);
    expect(report.strippedExifBytes).toBe(0);
    expect(report.scannedCount).toBeGreaterThan(0);
    expect(report.scannedCount).toBe(report.scannedParts.length);
    expect(report.scannedParts).toContain('word/document.xml');
    expect(report.scannedParts).toContain('docProps/core.xml');
    expect(report.scannedParts).toContain('docProps/app.xml');
    expect([...report.scannedParts].sort()).toEqual(report.scannedParts);
  });

  it('dirty package moves purge counts AND lists every scanned part', () => {
    const dirty = buildDirtyDocx();
    const { report } = sanitizeDocxBuffer(dirty, { zeroRetention: true });
    expect(report.scannedCount).toBeGreaterThan(0);
    expect(report.scannedCount).toBe(report.scannedParts.length);
    expect(report.scannedParts).toContain('word/document.xml');
    expect(report.scannedParts).toContain('docProps/core.xml');
    expect(report.scannedParts).toContain('docProps/app.xml');
    expect(report.scannedParts).toContain('word/media/image1.jpg');
    expect(report.scannedParts).toContain('word/media/image2.png');
    // Counts moved off zero — the dirty counterpart to the clean test above.
    expect(report.removedRSIDs).toBeGreaterThan(0);
    expect(report.purgedAuthors).toBeGreaterThan(0);
    expect(report.purgedPII).toBeGreaterThan(0);
    expect(report.strippedExifBytes).toBeGreaterThan(0);
    // Backward compatible: old shape intact, new fields survive JSON round-trip.
    const json = JSON.parse(JSON.stringify(report)) as typeof report;
    expect(json.removedRSIDs).toBe(report.removedRSIDs);
    expect(json.scannedCount).toBe(report.scannedCount);
    expect(json.scannedParts).toEqual(report.scannedParts);
  });

  it('preserves clause ids while still redacting real phones and SSNs', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:description>DFARS 252.204-7012 FAR 252.204-7012 CFR 252.204-7012 NIST 252.204-7012 Call 555-123-4567 SSN 123-45-6789</dc:description>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, {});
    expect(r.xml).toContain('DFARS 252.204-7012');
    expect(r.xml).toContain('FAR 252.204-7012');
    expect(r.xml).toContain('CFR 252.204-7012');
    expect(r.xml).toContain('NIST 252.204-7012');
    expect(r.xml).not.toContain('555-123-4567');
    expect(r.xml).not.toContain('123-45-6789');
    expect(r.purgedPII).toBe(2);

    const doc =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t>DFARS 252.204-7012 and call 555-123-4567</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const rb = sanitizeDocumentXml(doc, {}, 'word/document.xml');
    expect(rb.xml).toContain('252.204-7012');
    expect(rb.xml).not.toContain('555-123-4567');
    expect(rb.bodyPII).toBe(1);
  });

  it('preserves clause ids across colon and whitespace separator variants', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:description>DFARS:252.204-7012 FAR: 252.204-7012 CFR\t252.204-7012 NIST\n252.204-7012 DFARS  252.204-7012 DFARS.252.204-7012 DFARS. 252.204-7012 DFARS:\t252.204-7012 dfars:252.204-7012 Call 555-123-4567</dc:description>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, {});
    expect(r.xml).toContain('DFARS:252.204-7012');
    expect(r.xml).toContain('FAR: 252.204-7012');
    expect(r.xml).toContain('CFR\t252.204-7012');
    expect(r.xml).toContain('NIST\n252.204-7012');
    expect(r.xml).toContain('DFARS  252.204-7012');
    expect(r.xml).toContain('DFARS.252.204-7012');
    expect(r.xml).toContain('DFARS. 252.204-7012');
    expect(r.xml).toContain('DFARS:\t252.204-7012');
    expect(r.xml).toContain('dfars:252.204-7012');
    expect(r.xml).not.toContain('555-123-4567');
    expect(r.purgedPII).toBe(1);

    const doc =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t>DFARS:252.204-7012 and FAR  252.204-7012 and call 555-123-4567</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const rb = sanitizeDocumentXml(doc, {}, 'word/document.xml');
    expect(rb.xml).toContain('DFARS:252.204-7012');
    expect(rb.xml).toContain('FAR  252.204-7012');
    expect(rb.xml).not.toContain('555-123-4567');
    expect(rb.bodyPII).toBe(1);
  });

  it('redacts paren-shaped numbers even after an acronym (phone-shaped, never a clause id)', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:description>DFARS (252) 204-7012 and +1 (555) 123-4567</dc:description>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, {});
    // Accepted residual, pinned: the leading-+ and paren branches stay
    // unguarded because clause ids never render as `(252) ...` or `+...`;
    // redacting here avoids leaking real numbers mentioned near an acronym.
    expect(r.xml).not.toContain('(252) 204-7012');
    expect(r.xml).not.toContain('+1 (555) 123-4567');
    expect(r.purgedPII).toBe(2);
  });
});

describe('phone guard: clause citations survive, agency numbers redact', () => {
  it('keeps clause and section-sign citations while redacting nearby real numbers', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:description>DFARS clause 252.204-7012 and DFARS \u00A7 252.204-7012 and FAR clause 52.204-21 ` +
      `but NIST: 301-975-2000 and DFARS: 703-697-7241 and 555-123-4567</dc:description>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, {});
    expect(r.xml).toContain('DFARS clause 252.204-7012');
    expect(r.xml).toContain('DFARS \u00A7 252.204-7012');
    expect(r.xml).toContain('FAR clause 52.204-21');
    expect(r.xml).not.toContain('301-975-2000');
    expect(r.xml).not.toContain('703-697-7241');
    expect(r.xml).not.toContain('555-123-4567');
    expect(r.purgedPII).toBe(3);
  });

  it('keeps clause lists and multi-separator gaps in body prose', () => {
    const doc =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t>DFARS clauses 252.204-7012, 252.204-7013 and DFARS:  252.204-7012 ` +
      `but call NIST: 301-975-2000</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const r = sanitizeDocumentXml(doc, {}, 'word/document.xml');
    expect(r.xml).toContain('252.204-7012');
    expect(r.xml).toContain('252.204-7013');
    expect(r.xml).not.toContain('301-975-2000');
    expect(r.bodyPII).toBe(1);
  });
});

describe('location counts match actual redactions; stylesWithEffects rsids purged', () => {
  it('count-only per-paragraph collectors agree with the redacting pass (3 hits, 3 locations)', () => {
    const xml =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      `<w:p><w:r><w:t>Contact jane@example.com</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Call 555-123-4567</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>SSN 123-45-6789</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const actual = sanitizeDocumentXml(xml, {}, 'word/document.xml');
    expect(actual.bodyPII).toBe(3);
    // Same shape as the sanitize location display: one entry per hit.
    const locations: string[] = [];
    const paraRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
    let idx = 0;
    let m: RegExpExecArray | null;
    const re = new RegExp(paraRe.source, 'g');
    while ((m = re.exec(xml)) !== null) {
      const count =
        redactBodyPiiInWordText(m[0], false).count + redactTextNodePii(m[0], false).count;
      for (let k = 0; k < count; k++) locations.push(`word/document.xml#p${idx}`);
      idx++;
    }
    expect(locations).toEqual([
      'word/document.xml#p0',
      'word/document.xml#p1',
      'word/document.xml#p2',
    ]);
    expect(locations.length).toBe(actual.bodyPII);
  });

  it('count-only and redacting counts agree on single-run and split-run prose', () => {
    const single = `<w:p><w:r><w:t>Contact jane@example.com</w:t></w:r></w:p>`;
    expect(redactBodyPiiInWordText(single, false).count).toBe(1);
    expect(redactBodyPiiInWordText(single, true).count).toBe(1);
    expect(redactTextNodePii(single, false).count).toBe(0);
    const split =
      `<w:document xmlns:w="x"><w:body>` +
      `<w:p><w:r><w:t>mail jane@exa</w:t></w:r><w:r><w:t>mple.com</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    expect(redactBodyPiiInWordText(split, false).count).toBe(1);
    expect(redactBodyPiiInWordText(split, true).count).toBe(1);
    const outside =
      `<w:document xmlns:w="x"><w:r><w:t>jane@example.com</w:t></w:r></w:document>`;
    expect(redactBodyPiiInWordText(outside, false).count).toBe(1);
    expect(redactBodyPiiInWordText(outside, true).count).toBe(1);
  });

  it('text-node pass skips w:t prose but still catches revision text', () => {
    const del = `<w:del><w:r><w:delText>jane@example.com</w:delText></w:r></w:del>`;
    expect(redactTextNodePii(del, false).count).toBe(1);
    expect(redactBodyPiiInWordText(del, false).count).toBe(0);
    const redacted = redactTextNodePii(del, true);
    expect(redacted.count).toBe(1);
    expect(redacted.xml).not.toContain('jane@example.com');
    expect(redacted.xml).toContain('[redacted]');
  });

  it('purges rsids from word/stylesWithEffects.xml with zero survivors package-wide', () => {
    const styled =
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:style w:styleId="Normal" w:rsidR="00A1B2C3" w:rsidP="00D4E5F6">` +
      `<w:name w:val="Normal"/></w:style></w:styles>`;
    const files = new Map<string, Buffer>([
      ['word/stylesWithEffects.xml', Buffer.from(styled, 'utf8')],
      ['word/styles.xml', Buffer.from(styled, 'utf8')],
    ]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    expect(report.removedRSIDs).toBe(4);
    expect(report.scannedParts).toContain('word/stylesWithEffects.xml');
    for (const [, data] of clean) {
      expect(data.toString('utf8')).not.toMatch(/w:rsid[A-Za-z]*="/);
    }
    expect(clean.get('word/stylesWithEffects.xml')!.toString('utf8')).not.toContain('00A1B2C3');
    expect(report.filesTouched).toContain('word/stylesWithEffects.xml');
  });

  it('strips hostile SVG active content loudly while clean SVG stays byte-identical (F2 pin)', () => {
    // Mirrors /tmp evidence f2-live.ts: sanitize path strips (never refuses),
    // reports the part loudly, leaves clean bytes verbatim.
    const hostile =
      '<?xml-stylesheet href="e.xsl"?>' +
      '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" onload="e()">' +
      '<script>alert(1)</script>' +
      '<foreignObject><body>x</body></foreignObject>' +
      '<rect width="10" height="10"/>' +
      '</svg>';
    const clean =
      '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>';
    const pngBytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from('<w:document xmlns:w="x"><w:body/></w:document>', 'utf8')],
      ['word/media/evil.svg', Buffer.from(hostile, 'utf8')],
      ['word/media/ok.svg', Buffer.from(clean, 'utf8')],
      ['word/media/pic.png', Buffer.from(pngBytes)],
    ]);
    const { files: out, report } = sanitizePackageEntries(files, {});
    const evilOut = out.get('word/media/evil.svg')!.toString('utf8');
    expect(evilOut).not.toMatch(/<script/i);
    expect(evilOut).not.toMatch(/onload/i);
    expect(evilOut).not.toMatch(/foreignObject/i);
    expect(evilOut).toContain('<rect');
    expect(report.strippedSvgActive).toContain('word/media/evil.svg');
    expect(out.get('word/media/ok.svg')!.equals(Buffer.from(clean, 'utf8'))).toBe(true);
    expect(out.get('word/media/pic.png')!.equals(pngBytes)).toBe(true);
    expect(report.filesTouched).toContain('word/media/evil.svg');
    expect(report.scannedParts).toContain('word/media/evil.svg');
    expect(report.scannedParts).toContain('word/media/ok.svg');
  });
});

describe('drawing alternative text, chart/customXml attributes, labeled bare digits, extended deny list', () => {
  it('redacts email and formatted phone in drawing alternative text', () => {
    const xml =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<w:body><w:p><w:r><w:drawing><wp:inline>` +
      `<wp:docPr id="1" name="Photo jane@example.com" descr="Call 555-123-4567"/>` +
      `<pic:pic><pic:cNvPr id="2" name="Figure 2026" descr="Contact jane@example.com"/>` +
      `</pic:pic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`;
    const r = sanitizeDocumentXml(xml, {}, 'word/document.xml');
    expect(r.bodyPII).toBeGreaterThanOrEqual(3);
    expect(r.xml).not.toContain('jane@example.com');
    expect(r.xml).not.toContain('555-123-4567');
    // Structural ids survive; benign figure label without PII survives.
    expect(r.xml).toContain('id="1"');
    expect(r.xml).toContain('Figure 2026');
  });

  it('redacts labeled bare digits in drawing alt but leaves unlabeled years and counts intact', () => {
    const xml =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<w:body><w:p><w:r><w:drawing><wp:inline>` +
      `<wp:docPr id="1" name="x" descr="Phone: 5551234567"/>` +
      `</wp:inline></w:drawing></w:r></w:p></w:body></w:document>`;
    const r = sanitizeDocumentXml(xml, {}, 'word/document.xml');
    expect(r.bodyPII).toBe(1);
    expect(r.xml).not.toContain('5551234567');
    expect(r.xml).toContain('Phone: [redacted]');

    const cleanAlt =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<w:body><w:p><w:r><w:drawing><wp:inline>` +
      `<wp:docPr id="7" name="Figure 2026" descr="Q3 count 12345, year 2026"/>` +
      `</wp:inline></w:drawing></w:r></w:p></w:body></w:document>`;
    const rc = sanitizeDocumentXml(cleanAlt, {}, 'word/document.xml');
    expect(rc.bodyPII).toBe(0);
    expect(rc.xml).toBe(cleanAlt);
  });

  it('labeled helper covers tel/mobile/fax variants and ignores short fragments', () => {
    for (const label of ['tel: 5551234567', 'Phone 5551234567', 'mobile: 5551234567', 'FAX 5551234567']) {
      const r = redactDrawingAltPii(
        `<wp:docPr xmlns:wp="x" id="1" name="x" descr="${label}"/>`,
        true
      );
      expect(r.count).toBe(1);
      expect(r.xml).not.toContain('5551234567');
    }
    // Short fragments and years without the length gate stay intact.
    for (const quiet of ['Phone: 2026', 'tel: 123', 'fax: 12345']) {
      const r = redactDrawingAltPii(
        `<wp:docPr xmlns:wp="x" id="1" name="x" descr="${quiet}"/>`,
        true
      );
      expect(r.count).toBe(0);
    }
    // Bare digits without a label never redact, even at length.
    const bare = redactDrawingAltPii(
      `<wp:docPr xmlns:wp="x" id="1" name="x" descr="code 5551234567"/>`,
      true
    );
    expect(bare.count).toBe(0);
    expect(bare.xml).toContain('5551234567');
  });

  it('M2: body w:t prose redacts labeled bare digits, keeps unlabeled years/counts', () => {
    const xml =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t>Phone: 5551234567 in year 2026 with count 12345</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const r = sanitizeDocumentXml(xml, {}, 'word/document.xml');
    // Labeled bare-digit phone redacts (label kept); year + short count stay.
    expect(r.bodyPII).toBe(1);
    expect(r.xml).toContain('Phone: [redacted]');
    expect(r.xml).toContain('2026');
    expect(r.xml).toContain('12345');
    // Truly unlabeled bare runs stay intact (global matching rejected).
    const xml2 =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t>value 5551234567 in 2026</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const r2 = sanitizeDocumentXml(xml2, {}, 'word/document.xml');
    expect(r2.bodyPII).toBe(0);
    expect(r2.xml).toBe(xml2);
  });

  it('redacts attribute PII in chart and customXml parts while structural attrs survive', () => {
    const chart =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<c:title><c:tx><c:rich><a:p><a:t>Sales jane@example.com</a:t></a:p></c:rich></c:tx></c:title>` +
      `<c:txPr info="Contact 555-123-4567" note="Phone: 5551234567" count="2026"/>` +
      `</c:chartSpace>`;
    const rc = sanitizeDocumentXml(chart, {}, 'word/charts/chart1.xml');
    expect(rc.bodyPII).toBeGreaterThanOrEqual(3);
    expect(rc.xml).not.toContain('jane@example.com');
    expect(rc.xml).not.toContain('555-123-4567');
    expect(rc.xml).not.toContain('5551234567');
    expect(rc.xml).toContain('2026');

    const custom =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<root xmlns:ex="http://example.com" ex:owner="jane@example.com" ex:ref="Phone: 5551234567" ex:year="2026">` +
      `<item>plain</item></root>`;
    const rx = sanitizeDocumentXml(custom, {}, 'word/customXml/item1.xml');
    expect(rx.bodyPII).toBeGreaterThanOrEqual(2);
    expect(rx.xml).not.toContain('jane@example.com');
    expect(rx.xml).not.toContain('5551234567');
    expect(rx.xml).toContain('2026');
  });

  it('redacts field-code text carriers in document parts', () => {
    const xml =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:instrText xml:space="preserve"> HYPERLINK "mailto:jane@example.com" </w:instrText></w:r></w:p>` +
      `</w:body></w:document>`;
    const r = sanitizeDocumentXml(xml, {}, 'word/document.xml');
    expect(r.bodyPII).toBeGreaterThanOrEqual(1);
    expect(r.xml).not.toContain('jane@example.com');
  });

  it('clean document parts stay byte-identical (no drawing or chart churn)', () => {
    const doc =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<w:body><w:p><w:r><w:t>Hello clean world</w:t></w:r>` +
      `<w:drawing><wp:inline><wp:docPr id="1" name="Figure 1" descr="Quarterly chart"/></wp:inline></w:drawing>` +
      `</w:p></w:body></w:document>`;
    const r = sanitizeDocumentXml(doc, {}, 'word/document.xml');
    expect(r.bodyPII).toBe(0);
    expect(r.xml).toBe(doc);

    const chart =
      `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:lang val="en-US"/></c:chartSpace>`;
    const rc = sanitizeDocumentXml(chart, {}, 'word/charts/chart1.xml');
    expect(rc.bodyPII).toBe(0);
    expect(rc.xml).toBe(chart);
  });

  it('extended deny list drops each new carrier with silent drop+count semantics', () => {
    for (const ext of ['lnk', 'jar', 'hta', 'cpl', 'ws', 'wsf', 'psm1', 'com', 'pif', 'reg', 'url']) {
      expect(isBlockedExecutableEntry(`evil/payload.${ext}`)).toBe(true);
      expect(isBlockedExecutableEntry(`evil/PAYLOAD.${ext.toUpperCase()}`)).toBe(true);
    }
    expect(isBlockedExecutableEntry('word/media/image.exe.png')).toBe(false);
    expect(isBlockedExecutableEntry('word/document.xml')).toBe(false);

    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from('<w:document xmlns:w="x"><w:body/></w:document>', 'utf8')],
      ['evil/a.lnk', Buffer.from('x', 'utf8')],
      ['evil/b.jar', Buffer.from('x', 'utf8')],
      ['evil/c.hta', Buffer.from('x', 'utf8')],
      ['evil/d.cpl', Buffer.from('x', 'utf8')],
      ['evil/e.ws', Buffer.from('x', 'utf8')],
      ['evil/f.wsf', Buffer.from('x', 'utf8')],
      ['evil/g.psm1', Buffer.from('x', 'utf8')],
      ['evil/h.com', Buffer.from('x', 'utf8')],
      ['evil/i.pif', Buffer.from('x', 'utf8')],
      ['evil/j.reg', Buffer.from('x', 'utf8')],
      ['evil/k.url', Buffer.from('x', 'utf8')],
    ]);
    const { files: out, report } = sanitizePackageEntries(files, {});
    expect(out.has('word/document.xml')).toBe(true);
    for (const name of files.keys()) {
      if (name === 'word/document.xml') continue;
      expect(out.has(name)).toBe(false);
      expect(report.droppedRogue).toContain(name);
    }
    expect(report.droppedRogue.length).toBe(11);
    expect(report.filesTouched.length).toBeGreaterThanOrEqual(11);
  });

  it('attribute helper redacts labeled digits and leaves structural values alone', () => {
    const r = redactAttributePii(`<x a="Phone: 5551234567" b="2026" c="plain"/>`, true);
    expect(r.count).toBe(1);
    expect(r.xml).toContain('Phone: [redacted]');
    expect(r.xml).toContain('2026');
    const c = redactAttributePii(`<x a="2026" b="count 12345"/>`, true);
    expect(c.count).toBe(0);
  });
});

describe('embedded parts: hostile text redacted, clean bytes intact', () => {
  function buildWorkbookZip(sheetInner: string, sharedInner?: string): Buffer {
    const zip = new AdmZip();
    zip.addFile(
      '[Content_Types].xml',
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-8"?>' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        'utf8'
      )
    );
    zip.addFile(
      'xl/worksheets/sheet1.xml',
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${sheetInner}</t></is></c></row></sheetData>` +
          `</worksheet>`,
        'utf8'
      )
    );
    if (sharedInner !== undefined) {
      zip.addFile(
        'xl/sharedStrings.xml',
        Buffer.from(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
            `<si><t>${sharedInner}</t></si></sst>`,
          'utf8'
        )
      );
    }
    return zip.toBuffer();
  }

  it('redacts element text and attribute PII in embedded XML parts', () => {
    const hostile =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<root xmlns:ex="http://example.com" ex:owner="jane@example.com">` +
      `<item>Contact jane@example.com or 555-123-4567</item></root>`;
    const r = sanitizeDocumentXml(hostile, {}, 'word/embeddings/oleObject1.xml');
    expect(r.bodyPII).toBeGreaterThanOrEqual(3);
    expect(r.xml).not.toContain('jane@example.com');
    expect(r.xml).not.toContain('555-123-4567');
    expect(r.xml).toContain('[redacted]');
  });

  it('clean embedded XML stays byte-identical but still scanned', () => {
    const clean =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<root xmlns:ex="http://example.com"><item>Quarterly totals 2026</item></root>`;
    const r = sanitizeDocumentXml(clean, {}, 'word/embeddings/oleObject1.xml');
    expect(r.bodyPII).toBe(0);
    expect(r.xml).toBe(clean);
    const files = new Map<string, Buffer>([
      ['word/embeddings/oleObject1.xml', Buffer.from(clean, 'utf8')],
    ]);
    const { report } = sanitizePackageEntries(files, {});
    expect(report.scannedParts).toContain('word/embeddings/oleObject1.xml');
    expect(report.purgedPII).toBe(0);
  });

  it('package pass redacts hostile embedded XML with scan evidence', () => {
    const hostile =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<ole><t>Reach jane@example.com</t></ole>`;
    const oleBytes = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0x01]);
    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from('<w:document xmlns:w="x"><w:body/></w:document>', 'utf8')],
      ['word/embeddings/oleObject1.xml', Buffer.from(hostile, 'utf8')],
      // M4: embeddings bins refuse by default — the legit-opaque control
      // lives outside embeddings (R-10 rule keeps it legit).
      ['word/oleObject1.bin', Buffer.from(oleBytes)],
    ]);
    const { files: out, report } = sanitizePackageEntries(files, {});
    expect(out.get('word/embeddings/oleObject1.xml')!.toString('utf8')).not.toContain(
      'jane@example.com'
    );
    expect(report.purgedPII).toBeGreaterThanOrEqual(1);
    expect(report.scannedParts).toContain('word/embeddings/oleObject1.xml');
    expect(report.filesTouched).toContain('word/embeddings/oleObject1.xml');
    expect(out.get('word/oleObject1.bin')!.equals(oleBytes)).toBe(true);
  });

  it('redacts hostile workbook sheet text and keeps clean workbooks byte-identical', () => {
    const hostileWb = buildWorkbookZip('Contact jane@example.com 555-123-4567', 'Note jane@example.com');
    const r = sanitizeEmbeddedWorkbookBytes(hostileWb, 'word/embeddings/workbook.xlsx');
    expect(r.purgedPII).toBeGreaterThanOrEqual(3);
    const inner = new AdmZip(r.buf);
    const sheet = inner.getEntry('xl/worksheets/sheet1.xml')!.getData().toString('utf8');
    expect(sheet).not.toContain('jane@example.com');
    expect(sheet).not.toContain('555-123-4567');
    expect(sheet).toContain('[redacted]');
    const shared = inner.getEntry('xl/sharedStrings.xml')!.getData().toString('utf8');
    expect(shared).not.toContain('jane@example.com');

    const cleanWb = buildWorkbookZip('Quarterly totals 2026', 'Plain label');
    const rc = sanitizeEmbeddedWorkbookBytes(cleanWb, 'word/embeddings/workbook.xlsx');
    expect(rc.purgedPII).toBe(0);
    expect(rc.buf.equals(cleanWb)).toBe(true);
  });

  it('package pass scans workbook entries and preserves legit workbooks', () => {
    const cleanWb = buildWorkbookZip('Quarterly totals 2026');
    const oleBytes = Buffer.from([0xd0, 0xcf, 0x11, 0xe0]);
    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from('<w:document xmlns:w="x"><w:body/></w:document>', 'utf8')],
      ['word/embeddings/Microsoft_Excel_Worksheet1.xlsx', Buffer.from(cleanWb)],
      // M4: embeddings bins refuse by default — legit-opaque control outside.
      ['word/oleObject1.bin', Buffer.from(oleBytes)],
    ]);
    const { files: out, report } = sanitizePackageEntries(files, {});
    expect(report.scannedParts).toContain('word/embeddings/Microsoft_Excel_Worksheet1.xlsx');
    expect(out.get('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.equals(cleanWb)).toBe(true);
    expect(out.get('word/oleObject1.bin')!.equals(oleBytes)).toBe(true);
  });

  it('hostile workbook inside a package is neutralized end to end', () => {
    const hostileWb = buildWorkbookZip('SSN 123-45-6789 on file');
    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from('<w:document xmlns:w="x"><w:body/></w:document>', 'utf8')],
      ['word/embeddings/Microsoft_Excel_Worksheet1.xlsx', Buffer.from(hostileWb)],
    ]);
    const { files: out, report } = sanitizePackageEntries(files, {});
    expect(report.purgedPII).toBeGreaterThanOrEqual(1);
    const inner = new AdmZip(out.get('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!);
    expect(inner.getEntry('xl/worksheets/sheet1.xml')!.getData().toString('utf8')).not.toContain(
      '123-45-6789'
    );
  });

  it('forbidden markup inside a workbook fails closed', () => {
    const zip = new AdmZip();
    zip.addFile(
      'xl/worksheets/sheet1.xml',
      Buffer.from(
        '<?xml version="1.0"?><!DOCTYPE foo><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        'utf8'
      )
    );
    expect(() =>
      sanitizeEmbeddedWorkbookBytes(zip.toBuffer(), 'word/embeddings/workbook.xlsx')
    ).toThrow(/E_VALIDATION/);
  });
});
