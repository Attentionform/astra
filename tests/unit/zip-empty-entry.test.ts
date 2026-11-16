/**
 * An empty part written by a deflating producer has size 0 and a 2-byte compressed payload. System.IO.Packaging writes
 * `_xmlsignatures/origin.sigs` of a signed package that way (checked: size 0, compressed 2, method 8). The zip ceiling
 * used to refuse any size-0 entry with a payload, which refused those signed files at open.
 */
import { describe, expect, it } from 'vitest';
import { deflateRawSync } from 'node:zlib';
import AdmZip from 'adm-zip';
import { DocxParser, createBlankDocxBuffer } from '@attentionform/astra-engine';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
const crc32 = (buf: Buffer): number => {
  let c = 0xffffffff;
  for (const byte of buf) c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};

/** A minimal zip writer that deflates every entry, so an empty entry gets the 2-byte empty deflate stream (03 00). */
function zipDeflatingEverything(entries: Array<[string, Buffer]>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const [name, data] of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const packed = deflateRawSync(data);
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0, 6); local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0, 10); local.writeUInt16LE(0x21, 12); local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(nameBuf.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0, 8);
    central.writeUInt16LE(8, 10); central.writeUInt16LE(0, 12); central.writeUInt16LE(0x21, 14); central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(packed.length, 20); central.writeUInt32LE(data.length, 24); central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, packed);
    centrals.push(central, nameBuf);
    offset += local.length + nameBuf.length + packed.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralBuf.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, centralBuf, end]);
}

/** A signed package whose `_xmlsignatures/origin.sigs` is stored the way a deflating producer writes an empty part. */
function signedWithDeflatedEmptyPart(): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const entries: Array<[string, Buffer]> = zip.getEntries().filter((e) => !e.isDirectory).map((e) => [
    e.entryName,
    e.entryName === '[Content_Types].xml'
      ? Buffer.from(e.getData().toString('utf8').replace('<Default Extension="xml"', '<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/><Default Extension="xml"'), 'utf8')
      : e.getData(),
  ]);
  entries.push(['_xmlsignatures/origin.sigs', Buffer.alloc(0)]);
  entries.push(['_xmlsignatures/sig1.xml', Buffer.from('<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>', 'utf8')]);
  return zipDeflatingEverything(entries);
}

describe('empty parts with a deflate stream', () => {
  it('the fixture really has size 0 and a non-empty compressed payload', () => {
    const entry = new AdmZip(signedWithDeflatedEmptyPart()).getEntry('_xmlsignatures/origin.sigs')!;
    expect(entry.header.size).toBe(0);
    expect(entry.header.compressedSize).toBeGreaterThan(0);
    expect(entry.header.compressedSize).toBeLessThanOrEqual(8);
  });

  it('a signed package with such a part opens and is reported as signed', () => {
    const parser = new DocxParser(signedWithDeflatedEmptyPart());
    expect(parser.getPackageReport().signed).toBe(true);
    expect(parser.getStructuralSummary().nodes.length).toBeGreaterThan(0);
  });

  it('a size-0 entry with a real payload is still refused (the lying-size pattern)', () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const buf = zip.toBuffer();
    // Patch the document part to declare size 0 with a 64-byte payload in its central directory record.
    const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
    let off = buf.readUInt32LE(eocd + 16);
    const patched = Buffer.from(buf);
    for (let i = 0; i < buf.readUInt16LE(eocd + 10); i++) {
      const nameLen = patched.readUInt16LE(off + 28);
      const name = patched.toString('utf8', off + 46, off + 46 + nameLen);
      if (name === 'word/document.xml') {
        patched.writeUInt32LE(64, off + 20); // compressed size
        patched.writeUInt32LE(0, off + 24); // uncompressed size
      }
      off += 46 + nameLen + patched.readUInt16LE(off + 30) + patched.readUInt16LE(off + 32);
    }
    expect(() => new DocxParser(patched)).toThrow(/declared size 0 with non-zero compressed payload \(64 bytes\)/);
  });
});
