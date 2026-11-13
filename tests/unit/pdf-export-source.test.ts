import { describe, it, expect } from 'vitest';
import { createHash } from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import { AdmZip, createBlankDocxBuffer, validateDocxBuffer } from '../../packages/engine/src';
import { prepareDocxForPdfExport, PDF_SOURCE_REMEDY } from '../../packages/engine/src/export/pdf-source';

/**
 * PDF export renders either the original bytes (a valid file) or a verified temporary copy in which only the
 * order of children in schema-ordered containers differs. Everything else is refused.
 * Needs the native validator, so it runs in CI and locally with OOXML_VALIDATOR_CLI.
 */
const FIXTURE = path.resolve(__dirname, '../fixtures/word-order.docx');
const sha = (buf: Buffer): string => createHash('sha256').update(buf).digest('hex');

function wordOrder(): Buffer {
  return fs.readFileSync(FIXTURE);
}

function edit(buf: Buffer, change: (xml: string) => string, part = 'word/document.xml'): Buffer {
  const zip = new AdmZip(buf);
  zip.updateFile(part, Buffer.from(change(zip.readAsText(part)), 'utf8'));
  return zip.toBuffer();
}

function parts(buf: Buffer): Map<string, string> {
  const zip = new AdmZip(buf);
  return new Map(zip.getEntries().filter((e) => !e.isDirectory).map((e) => [e.entryName, sha(e.getData())]));
}

function changedParts(a: Buffer, b: Buffer): string[] {
  const left = parts(a);
  const right = parts(b);
  const names = new Set([...left.keys(), ...right.keys()]);
  return [...names].filter((n) => left.get(n) !== right.get(n)).sort();
}

describe('prepareDocxForPdfExport', () => {
  it('repairs a Word-order file on a copy: valid result, only document.xml changed, source untouched', async () => {
    const original = wordOrder();
    const before = sha(original);
    expect((await validateDocxBuffer(original)).ok).toBe(false);

    const prepared = await prepareDocxForPdfExport(original);
    expect(prepared.repaired).toBe(true);
    expect(prepared.repairedParts).toEqual(['word/document.xml']);
    expect(prepared.reorderedContainers).toBe(1);
    expect((await validateDocxBuffer(prepared.buffer)).ok).toBe(true);
    expect(changedParts(original, prepared.buffer)).toEqual(['word/document.xml']);
    expect(sha(original)).toBe(before);
    expect(prepared.buffer).not.toBe(original);
  });

  it('a valid document goes to the renderer as the same bytes, and the repair code is not entered', async () => {
    const valid = createBlankDocxBuffer('blank');
    const prepared = await prepareDocxForPdfExport(valid);
    expect(prepared.repaired).toBe(false);
    expect(prepared.buffer).toBe(valid);
    expect(prepared.repairedParts).toEqual([]);
  });

  it('a mixed file (an order error and a real error) is refused with what remains, and nothing is produced', async () => {
    const mixed = edit(wordOrder(), (xml) => xml.replace('<w:t>Plain paragraph.</w:t>', '<w:t>Plain paragraph.</w:t><w:bogus/>'));
    await expect(prepareDocxForPdfExport(mixed)).rejects.toMatchObject({
      code: 'E_VALIDATION',
      remedy: PDF_SOURCE_REMEDY,
    });
    await expect(prepareDocxForPdfExport(mixed)).rejects.toThrow(/bogus/);
  });

  it('a refusal never says to re-export the file from Word', async () => {
    const mixed = edit(wordOrder(), (xml) => xml.replace('<w:t>Plain paragraph.</w:t>', '<w:t>Plain paragraph.</w:t><w:bogus/>'));
    const error = (await prepareDocxForPdfExport(mixed).catch((e: unknown) => e)) as { remedy: string };
    expect(error.remedy).not.toMatch(/re-export the file from Word/i);
    expect(error.remedy).toContain('astra validate');
    expect(error.remedy).toContain('astra normalize');
  });

  it('paragraphs without w14:paraId get none, and existing ids are unchanged', async () => {
    const countIds = (buf: Buffer): string[] =>
      [...new AdmZip(buf).readAsText('word/document.xml').matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1]);

    const withoutIds = await prepareDocxForPdfExport(wordOrder());
    expect(countIds(wordOrder())).toEqual([]);
    expect(countIds(withoutIds.buffer)).toEqual([]);

    const withIds = edit(wordOrder(), (xml) =>
      xml.replace('<w:p><w:r><w:rPr>', '<w:p w14:paraId="1A2B3C4D" w14:textId="77777777"><w:r><w:rPr>').replace('<w:p><w:r><w:t>', '<w:p w14:paraId="5E6F7A8B"><w:r><w:t>')
    );
    expect(countIds(withIds)).toEqual(['1A2B3C4D', '5E6F7A8B']);
    const prepared = await prepareDocxForPdfExport(withIds);
    expect(prepared.repaired).toBe(true);
    expect(countIds(prepared.buffer)).toEqual(['1A2B3C4D', '5E6F7A8B']);
  });

  it('every part other than the repaired one is content-identical, including extra parts', async () => {
    const zip = new AdmZip(wordOrder());
    const media = Buffer.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    zip.addFile('customXml/item1.xml', Buffer.from('<?xml version="1.0"?><root><b>2</b><a>1</a></root>', 'utf8'));
    zip.addFile('docProps/thumbnail.bin', media);
    const original = zip.toBuffer();
    if (!(await validateDocxBuffer(original)).ok === false) throw new Error('fixture must still be invalid');
    const prepared = await prepareDocxForPdfExport(original);
    expect(prepared.repaired).toBe(true);
    expect(changedParts(original, prepared.buffer)).toEqual(['word/document.xml']);
    expect(new AdmZip(prepared.buffer).getEntry('docProps/thumbnail.bin')?.getData().equals(media)).toBe(true);
  });

  it('a file that needs a rename or an infill is refused, not repaired', async () => {
    // w:shd without the required w:val is the whitelisted infill of the normalizer, which export does not apply.
    const needsInfill = edit(createBlankDocxBuffer('blank'), (xml) =>
      xml.replace('</w:body>', '<w:p><w:r><w:rPr><w:shd w:fill="FF0000" w:color="auto"/></w:rPr><w:t>x</w:t></w:r></w:p></w:body>')
    );
    if ((await validateDocxBuffer(needsInfill)).ok) throw new Error('fixture must be invalid');
    await expect(prepareDocxForPdfExport(needsInfill)).rejects.toMatchObject({ code: 'E_VALIDATION', remedy: PDF_SOURCE_REMEDY });
  });

  it('cannot complete: a part that is not well-formed XML refuses the export', async () => {
    const zip = new AdmZip(wordOrder());
    zip.addFile('customXml/broken.xml', Buffer.from('<a><b>', 'utf8'));
    const refused = await prepareDocxForPdfExport(zip.toBuffer()).catch((e: unknown) => e);
    expect(refused).toMatchObject({ code: 'E_VALIDATION' });
  });

  // ---- package safety and encoding (design M5 v3 to v5) ----

  const customPart = (buf: Buffer, name: string, data: Buffer): Buffer => {
    const zip = new AdmZip(buf);
    zip.addFile(name, data);
    return zip.toBuffer();
  };
  const entryBytes = (buf: Buffer, name: string): Buffer | undefined => new AdmZip(buf).getEntry(name)?.getData();
  const utf16 = (xml: string, bom = true): Buffer => Buffer.concat([bom ? Buffer.from([0xff, 0xfe]) : Buffer.alloc(0), Buffer.from(xml, 'utf16le')]);
  const BAD_ORDER_XML = '<?xml version="1.0"?><w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sz w:val="2"/><w:color w:val="000000"/></w:rPr>';

  it('an unrelated valid UTF-16 part is copied byte for byte and the export succeeds when the validator accepts it', async () => {
    const part = utf16('<?xml version="1.0" encoding="UTF-16"?><root><a>1</a></root>');
    const original = customPart(wordOrder(), 'customXml/item1.xml', part);
    const verdict = await validateDocxBuffer(original);
    // The validator's own verdict on the original: it can only be invalid because of the order error in document.xml.
    expect(verdict.ok).toBe(false);
    const prepared = await prepareDocxForPdfExport(original).catch((e: unknown) => e);
    if ((prepared as { buffer?: Buffer }).buffer) {
      const out = prepared as { buffer: Buffer; repairedParts: string[] };
      expect(out.repairedParts).toEqual(['word/document.xml']);
      expect(entryBytes(out.buffer, 'customXml/item1.xml')!.equals(part)).toBe(true);
    } else {
      // The validator refuses a package with that part: the export is refused with its errors, never repaired around.
      expect(prepared).toMatchObject({ code: 'E_VALIDATION' });
    }
  });

  it('an order error inside a UTF-16 part is not repaired: the export is refused', async () => {
    const original = customPart(wordOrder(), 'word/theme/x.xml', utf16(BAD_ORDER_XML));
    const result = await prepareDocxForPdfExport(original).catch((e: unknown) => e);
    // document.xml is repaired, the UTF-16 part is left as it is; the result is either success (the validator does
    // not check that part) with the part untouched, or a refusal. It is never a rewritten UTF-16 part.
    if ((result as { buffer?: Buffer }).buffer) {
      expect(entryBytes((result as { buffer: Buffer }).buffer, 'word/theme/x.xml')!.equals(utf16(BAD_ORDER_XML))).toBe(true);
    } else {
      expect(result).toMatchObject({ code: 'E_VALIDATION' });
    }
  });

  it('a part declaring windows-1252 is not examined and keeps its bytes', async () => {
    const part = Buffer.from('<?xml version="1.0" encoding="windows-1252"?><r>caf\xe9</r>', 'latin1');
    const original = customPart(wordOrder(), 'customXml/item1.xml', part);
    const result = await prepareDocxForPdfExport(original).catch((e: unknown) => e);
    if ((result as { buffer?: Buffer }).buffer) {
      expect(entryBytes((result as { buffer: Buffer }).buffer, 'customXml/item1.xml')!.equals(part)).toBe(true);
    } else {
      expect(result).toMatchObject({ code: 'E_VALIDATION' });
    }
  });

  it('a UTF-8 part with a BOM and non-ASCII text round-trips through the repair exactly', async () => {
    const bomDoc = edit(wordOrder(), (xml) => '\uFEFF' + xml.replace('Plain paragraph.', 'Café 日本語 — ok'));
    const prepared = await prepareDocxForPdfExport(bomDoc);
    const out = entryBytes(prepared.buffer, 'word/document.xml')!;
    expect(out.subarray(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf]))).toBe(true);
    expect(out.toString('utf8')).toContain('Café 日本語 — ok');
    expect(out.toString('utf8')).not.toContain('\uFFFD');
    expect(out.toString('utf8')).toContain('standalone="yes"');
  });

  it('invalid UTF-8 in the part that holds the order error is refused: nothing is rewritten with replacement characters', async () => {
    const zip = new AdmZip(wordOrder());
    const xml = Buffer.from(zip.readAsText('word/document.xml'), 'utf8');
    const marker = Buffer.from('Plain paragraph.', 'utf8');
    const at = xml.indexOf(marker);
    const broken = Buffer.concat([xml.subarray(0, at), Buffer.from([0xc3, 0x28]), xml.subarray(at)]);
    zip.updateFile('word/document.xml', broken);
    const result = await prepareDocxForPdfExport(zip.toBuffer()).catch((e: unknown) => e);
    expect(result).toMatchObject({ code: 'E_VALIDATION' });
  });

  it('invalid UTF-8 in an unrelated part: recorded verdict, and the part is never rewritten', async () => {
    const part = Buffer.concat([Buffer.from('<?xml version="1.0"?><a>'), Buffer.from([0xc3, 0x28]), Buffer.from('</a>')]);
    const original = customPart(wordOrder(), 'customXml/item1.xml', part);
    // Recorded on the native validator (Linux build): the only error of the original is the order error, and
    // the validator accepts the package with the invalid bytes in the unrelated part once that is repaired.
    expect((await validateDocxBuffer(original)).ok).toBe(false);
    const result = await prepareDocxForPdfExport(original).catch((e: unknown) => e);
    if ((result as { buffer?: Buffer }).buffer) {
      const out = (result as { buffer: Buffer }).buffer;
      expect((await validateDocxBuffer(out)).ok).toBe(true);
      expect(entryBytes(out, 'customXml/item1.xml')!.equals(part)).toBe(true);
      expect(entryBytes(out, 'customXml/item1.xml')!.includes(Buffer.from('\uFFFD', 'utf8'))).toBe(false);
    } else {
      expect(result).toMatchObject({ code: 'E_VALIDATION' });
    }
  });

  it('a truncated multi-byte sequence and an overlong encoding are not established as UTF-8', async () => {
    const { establishedUtf8Text } = await import('../../packages/engine/src/export/pdf-source');
    expect(establishedUtf8Text(Buffer.from([0x3c, 0x61, 0x2f, 0x3e, 0xe2, 0x82]))).toBeNull();
    expect(establishedUtf8Text(Buffer.from([0x3c, 0x61, 0x3e, 0xc0, 0xaf, 0x3c, 0x2f, 0x61, 0x3e]))).toBeNull();
    expect(establishedUtf8Text(Buffer.from('<a>é</a>', 'utf8'))).toBe('<a>é</a>');
    expect(establishedUtf8Text(utf16('<a/>'))).toBeNull();
    expect(establishedUtf8Text(Buffer.from('<?xml version="1.0" encoding="iso-8859-1"?><a/>', 'latin1'))).toBeNull();
    expect(establishedUtf8Text(Buffer.from('<?xml version="1.0" encoding="utf-8"?><a/>', 'utf8'))).not.toBeNull();
  });

  it('a package with two entries of the same name is refused, not repaired around', async () => {
    const zip = new AdmZip(wordOrder());
    zip.addFile('customXml/item1.xml', Buffer.from('<a/>', 'utf8'));
    zip.addFile('customXml/item2.xml', Buffer.from('<b/>', 'utf8'));
    const raw = Buffer.from(zip.toBuffer());
    let at = raw.indexOf('item2');
    while (at >= 0) {
      raw.write('item1', at, 'latin1');
      at = raw.indexOf('item2', at + 5);
    }
    // The zip library itself refuses two entries of one name; the export must refuse too, whatever the validator says.
    expect(() => new AdmZip(raw).getEntries()).toThrow(/Duplicate entry name/);
    await expect(prepareDocxForPdfExport(raw)).rejects.toMatchObject({ code: 'E_VALIDATION', remedy: PDF_SOURCE_REMEDY });
    await expect(prepareDocxForPdfExport(raw)).rejects.toThrow(/Format error in package|Duplicate entry name|cannot be read as a zip/);
  });

  it('an entry whose name climbs out of the package is refused', async () => {
    const zip = new AdmZip(wordOrder());
    zip.addFile('customXml/item1.xml', Buffer.from('<a/>', 'utf8'));
    const raw = Buffer.from(zip.toBuffer());
    const name = Buffer.from('customXml/item1.xml');
    const evil = Buffer.from('customXml/../x1.xml'.padEnd(name.length, 'x').slice(0, name.length));
    expect(evil.length).toBe(name.length);
    let at = raw.indexOf(name);
    while (at >= 0) {
      evil.copy(raw, at);
      at = raw.indexOf(name, at + name.length);
    }
    expect(new AdmZip(raw).getEntries().some((e) => e.entryName.includes('..'))).toBe(true);
    await expect(prepareDocxForPdfExport(raw)).rejects.toMatchObject({ code: 'E_VALIDATION', remedy: PDF_SOURCE_REMEDY });
  });

  it('a part with a DOCTYPE or an entity is refused through the XXE guard', async () => {
    const xxe = customPart(wordOrder(), 'customXml/item1.xml', Buffer.from('<?xml version="1.0"?><!DOCTYPE a [<!ENTITY e SYSTEM "file:///etc/passwd">]><a>&e;</a>', 'utf8'));
    await expect(prepareDocxForPdfExport(xxe)).rejects.toMatchObject({ code: 'E_VALIDATION' });
  });

  it('the package ceilings apply: a ratio above the ceiling is refused with E_LIMIT, for an invalid input and a valid one', async () => {
    // Not very compressible, so it passes the default ratio of 100 and the low ceiling below catches it.
    let seed = 12345;
    const noise = Array.from({ length: 6000 }, () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) >> 8) & 0xff);
    const big = Buffer.from('<?xml version="1.0"?><r>' + Buffer.from(noise).toString('base64') + '</r>', 'utf8');
    const invalid = customPart(wordOrder(), 'customXml/big.xml', big);
    const valid = customPart(createBlankDocxBuffer('blank'), 'customXml/big.xml', big);
    const previous = process.env.ASTRA_MAX_RATIO;
    process.env.ASTRA_MAX_RATIO = '1.05';
    try {
      await expect(prepareDocxForPdfExport(invalid)).rejects.toThrow(/E_LIMIT/);
      await expect(prepareDocxForPdfExport(valid)).rejects.toThrow(/E_LIMIT/);
    } finally {
      if (previous === undefined) delete process.env.ASTRA_MAX_RATIO;
      else process.env.ASTRA_MAX_RATIO = previous;
    }
    // Under the default ceilings both behave as before.
    expect((await prepareDocxForPdfExport(valid)).repaired).toBe(false);
    expect((await prepareDocxForPdfExport(invalid)).repaired).toBe(true);
  });
});
