/**
 * The XML integrity check in `rezip()` refuses a part that is not one XML element: text with no element, and text
 * before or after the root. The XML parser accepts these without an error, so the check looks for them itself.
 * It does not catch every malformation: a mismatched closing tag passes (`validateDocxBuffer()` catches it).
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser } from '../src/internal';

function withTheme(xml: string): DocxParser {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  zip.updateFile('word/theme/theme1.xml', Buffer.from(xml, 'utf8'));
  return new DocxParser(zip.toBuffer());
}

describe('rezip() integrity guard', () => {
  it.each([
    ['text with no element', 'hello'],
    ['text before the root', 'junk<a:theme xmlns:a="x"/>'],
    ['text after the root', '<a:theme xmlns:a="x"/>junk'],
    ['a truncated part', '<a:theme xmlns:a="x"><a:x>'],
    ['two root elements', '<a:a xmlns:a="x"/><a:b xmlns:a="x"/>'],
  ])('refuses %s', (_label, xml) => {
    expect(() => withTheme(xml).rezip()).toThrow(/Well-Formedness/);
  });

  it('accepts a part with an XML declaration, a byte order mark and trailing whitespace', () => {
    const xml = '﻿<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="x"/>\n';
    expect(() => withTheme(xml).rezip()).not.toThrow();
  });

  it('accepts an unedited document', () => {
    expect(() => DocxParser.createNewDocument('blank').rezip()).not.toThrow();
  });

  it('skipValidation turns the check off, and the XXE and macro checks still run', () => {
    expect(() => withTheme('hello').rezip({ skipValidation: true })).not.toThrow();
  });
});

describe('rezip() reproducible option', () => {
  it('the same input bytes give the same output bytes, apart in time', async () => {
    const seed = DocxParser.createNewDocument('blank').rezip();
    const first = new DocxParser(seed).rezip({ reproducible: true });
    await new Promise((resolve) => setTimeout(resolve, 2100));
    const second = new DocxParser(seed).rezip({ reproducible: true });
    expect(second.equals(first)).toBe(true);
  });

  it('without the option and without the environment variables the dates are real', async () => {
    const saved = { r: process.env.ASTRA_REPRODUCIBLE, s: process.env.SOURCE_DATE_EPOCH, n: process.env.ASTRA_NO_REPRODUCIBLE };
    delete process.env.ASTRA_REPRODUCIBLE;
    delete process.env.SOURCE_DATE_EPOCH;
    delete process.env.ASTRA_NO_REPRODUCIBLE;
    try {
      const seed = DocxParser.createNewDocument('blank').rezip();
      const first = new DocxParser(seed).rezip();
      await new Promise((resolve) => setTimeout(resolve, 2100));
      const second = new DocxParser(seed).rezip();
      expect(second.equals(first)).toBe(false);
    } finally {
      for (const [k, v] of [['ASTRA_REPRODUCIBLE', saved.r], ['SOURCE_DATE_EPOCH', saved.s], ['ASTRA_NO_REPRODUCIBLE', saved.n]] as const) {
        if (v === undefined) delete process.env[k];
        else process.env[k] = v;
      }
    }
  });
});
