import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { decodeXmlPart } from '../../packages/engine/src/ooxml/part-text';

const SOURCES = '<?xml version="1.0" encoding="UTF-16" standalone="no"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="\\APA.XSL"/>';
const le = (s: string) => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(s, 'utf16le')]);
const be = (s: string) => { const b = Buffer.from(s, 'utf16le'); b.swap16(); return Buffer.concat([Buffer.from([0xfe, 0xff]), b]); };
const pkg = (part: Buffer) => { const z = new AdmZip(createBlankDocxBuffer('blank')); z.addFile('customXml/item2.xml', part); return z.toBuffer(); };
const bytesOf = (buf: Buffer) => new AdmZip(buf).getEntry('customXml/item2.xml')!.getData();

describe('UTF-16 XML parts', () => {
  it('decodes by byte order mark and by the first characters of a declaration', () => {
    expect(decodeXmlPart(le(SOURCES))).toBe(SOURCES);
    expect(decodeXmlPart(be(SOURCES))).toBe(SOURCES);
    expect(decodeXmlPart(Buffer.from(SOURCES, 'utf16le'))).toBe(SOURCES);
    const beNoBom = Buffer.from(SOURCES, 'utf16le'); beNoBom.swap16();
    expect(decodeXmlPart(beNoBom)).toBe(SOURCES);
    expect(decodeXmlPart(Buffer.from('<a>é</a>', 'utf8'))).toBe('<a>é</a>');
  });
  it('a UTF-16LE and a UTF-16BE part are saved and keep their bytes', () => {
    for (const part of [le(SOURCES), be(SOURCES)]) {
      const parser = new DocxParser(pkg(part));
      const out = parser.rezip();
      expect(bytesOf(out).equals(part)).toBe(true);
    }
  });
  it('a broken UTF-16 part is still refused, like the same part in UTF-8', () => {
    const text = '<?xml version="1.0"?><a>&foo;</a>';
    expect(() => new DocxParser(pkg(Buffer.from(text, 'utf8'))).rezip()).toThrow(/Well-Formedness|corruption/);
    expect(() => new DocxParser(pkg(le(text))).rezip()).toThrow(/Well-Formedness|corruption/);
    expect(() => new DocxParser(pkg(le(SOURCES).subarray(0, 41))).rezip()).toThrow(/Well-Formedness|corruption/);
  });
  it('a document type declaration with an entity in a UTF-16 part is refused as XXE', () => {
    const evil = le('<?xml version="1.0"?><!DOCTYPE d [<!ENTITY e SYSTEM "file:///etc/passwd">]><d>&e;</d>');
    expect(() => new DocxParser(pkg(evil))).toThrow(/DOCTYPE|XXE|ENTITY|prohibited/i);
  });
});
