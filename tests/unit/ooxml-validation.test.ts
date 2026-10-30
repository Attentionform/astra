import { describe, it, expect } from 'vitest';
import { DocxParser } from '@attentionform/astra-engine';
import AdmZip from 'adm-zip';

describe('OOXML Automated Integrity Guard Suite', () => {
  it('should detect valid XML parts and pass validateZipIntegrity', async () => {
    const zip = new AdmZip();
    zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
    zip.addFile('word/document.xml', Buffer.from('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Valid XML</w:t></w:r></w:p></w:body></w:document>'));

    const parser = new DocxParser(zip.toBuffer());
    const validation = parser.validateZipIntegrity();

    expect(validation.isValid).toBe(true);
    expect(validation.errors).toHaveLength(0);
    expect(() => parser.rezip()).not.toThrow();
  });

  it('should block rezip and throw error when unescaped ampersand or corrupt XML is present', async () => {
    const zip = new AdmZip();
    zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
    // Corrupt XML part with unclosed tag and undefined entity
    zip.addFile('word/footer1.xml', Buffer.from('<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Confidential &badentity;</w:t></w:r><unclosed_tag></w:ftr>'));
    zip.addFile('word/document.xml', Buffer.from('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>'));

    const parser = new DocxParser(zip.toBuffer());
    const validation = parser.validateZipIntegrity();

    expect(validation.isValid).toBe(false);
    expect(validation.errors.length).toBeGreaterThan(0);
    expect(validation.errors[0]).toContain('word/footer1.xml');

    expect(() => parser.rezip()).toThrow('[OOXML Integrity Guard] Export blocked due to XML corruption');
  });
});
