import { describe, it, expect, beforeAll, afterAll, vi } from 'vitest';
import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import {
  DocxParser,
  createBlankDocxBuffer,
  AdmZip,
  clampZipEntryTimes,
  ZIP_FIXED_EPOCH_MS,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

// Seeded hostile-input suite for the document parser.
//
// Fixed seed drives every mutation below (mulberry32). No wall-clock,
// no Math.random, no network: rebuilding the corpus twice yields identical
// bytes (pinned by hash assertion inside the run).
const SEED = 0x51eed;

/** Deterministic PRNG (mulberry32). */
function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function docXmlWithBody(bodyInner: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyInner}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>'
  );
}

function para(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

// Minimal STORE-method zip builder: keeps entry names byte-exact (no
// normalization), so traversal/absolute/drive shapes reach the parser raw.
const CRC_TABLE: Int32Array = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(b: Buffer): number {
  let c = -1;
  for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

interface RawEntry {
  name: string;
  data: Buffer;
}

function buildRawZip(entries: RawEntry[]): Buffer {
  const enc = (s: string): Buffer => Buffer.from(s, 'utf8');
  const chunks: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const name = enc(e.name);
    const crc = crc32(e.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(e.data.length, 18);
    local.writeUInt32LE(e.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    chunks.push(local, name, e.data);
    const cen = Buffer.alloc(46);
    cen.writeUInt32LE(0x02014b50, 0);
    cen.writeUInt16LE(20, 4);
    cen.writeUInt16LE(20, 6);
    cen.writeUInt16LE(0x0800, 8);
    cen.writeUInt16LE(0, 10);
    cen.writeUInt16LE(0, 12);
    cen.writeUInt16LE(0, 14);
    cen.writeUInt32LE(crc, 16);
    cen.writeUInt32LE(e.data.length, 20);
    cen.writeUInt32LE(e.data.length, 24);
    cen.writeUInt16LE(name.length, 28);
    cen.writeUInt16LE(0, 30);
    cen.writeUInt16LE(0, 32);
    cen.writeUInt16LE(0, 34);
    cen.writeUInt16LE(0, 36);
    cen.writeUInt32LE(0, 38);
    cen.writeUInt32LE(offset, 42);
    central.push(cen, name);
    offset += local.length + name.length + e.data.length;
  }
  const cenBuf = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(cenBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...chunks, cenBuf, end]);
}

// Rewrites central-directory size headers without inflating (ceiling probe).
function patchCentralDirectory(
  buffer: Buffer,
  patchFn: (entry: { name: string; compressedSize: number; uncompressedSize: number }) => {
    compressedSize?: number;
    uncompressedSize?: number;
  } | void,
): Buffer {
  const copy = Buffer.from(buffer);
  let offset = 0;
  const cdSig = Buffer.from([0x50, 0x4b, 0x01, 0x02]);
  for (;;) {
    offset = copy.indexOf(cdSig, offset);
    if (offset === -1) break;
    const compSize = copy.readUInt32LE(offset + 20);
    const uncompSize = copy.readUInt32LE(offset + 24);
    const nameLen = copy.readUInt16LE(offset + 28);
    const name = copy.toString('utf8', offset + 46, offset + 46 + nameLen);
    const res = patchFn({ name, compressedSize: compSize, uncompressedSize: uncompSize });
    if (res) {
      if (typeof res.compressedSize === 'number') copy.writeUInt32LE(res.compressedSize, offset + 20);
      if (typeof res.uncompressedSize === 'number') copy.writeUInt32LE(res.uncompressedSize, offset + 24);
    }
    offset += 4;
  }
  return copy;
}

interface FuzzCase {
  name: string;
  kind: string;
  buffer: Buffer;
}

function baseEntries(): RawEntry[] {
  const zip = new AdmZip(createBlankDocxBuffer('blank')) as unknown as {
    getEntries(): { entryName: string; getData(): Buffer }[];
  };
  return zip.getEntries().map((e) => ({ name: e.entryName, data: Buffer.from(e.getData()) }));
}

function mutateViaZip(label: string, fn: (zip: IZipEdit) => void): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank')) as unknown as IZipEdit;
  fn(zip);
  void label;
  // Clamp entry stamps: the writer stamps added/updated entries with the
  // wall clock, which would make corpus bytes differ run to run.
  clampZipEntryTimes(
    zip as unknown as Parameters<typeof clampZipEntryTimes>[0],
    new Date(ZIP_FIXED_EPOCH_MS),
  );
  return (zip as unknown as { toBuffer(): Buffer }).toBuffer();
}

interface IZipEdit {
  addFile(name: string, data: Buffer): void;
  updateFile(name: string, data: Buffer): void;
  deleteFile(name: string): void;
}

const DECL_VARIANTS = [
  '<!DOCTYPE foo>',
  '<!doctype foo>',
  '<!DoCtYpE foo>',
  '<!dOcTyPe foo SYSTEM "http://evil.example/x">',
  '<!ENTITY xxe "pwn">',
  '<!entity xxe SYSTEM "file:///etc/passwd">',
  '<!\nDOCTYPE foo>',
  '<!\tENTITY xxe "pwn">',
  '<! ENTITY xxe "pwn">',
  '<!  DOCTYPE foo>',
  '<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>',
  '<!!DOCTYPE foo>',
];

function buildCorpus(): FuzzCase[] {
  const cases: FuzzCase[] = [];
  const base = createBlankDocxBuffer('blank');

  // 1. Raw byte shapes (not parseable as packages).
  {
    const rnd = mulberry32(SEED ^ 0x01);
    const push = (name: string, buffer: Buffer): void => {
      cases.push({ name: `raw/${name}`, kind: 'raw', buffer });
    };
    push('empty', Buffer.alloc(0));
    push('one-byte', Buffer.from([0x50]));
    push('nul-64', Buffer.alloc(64, 0));
    push('text-garbage', Buffer.from('this is not a zip package at all', 'utf8'));
    push('utf16-text', Buffer.from('not a zip', 'utf16le'));
    push('local-header-only', Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0x01, 0x02, 0x03]));
    push('end-only', Buffer.from([0x50, 0x4b, 0x05, 0x06]));
    for (const cut of [16, 64, 512, base.length - 12]) {
      push(`truncated@${cut}`, base.subarray(0, cut));
    }
    const png = Buffer.from(base);
    png[0] = 0x89;
    png[1] = 0x50;
    push('magic-swapped', png);
    push('trailing-junk', Buffer.concat([base, Buffer.from('TRAILINGJUNK')]));;
    const randLen = 32 + Math.floor(rnd() * 96);
    const rand = Buffer.alloc(randLen);
    for (let i = 0; i < randLen; i++) rand[i] = Math.floor(rnd() * 256);
    push('random-seeded', rand);
  }

  // 2. Declaration smuggling across package parts.
  {
    const rnd = mulberry32(SEED ^ 0x02);
    const evilDoc = (decl: string): string =>
      '<?xml version="1.0" encoding="UTF-8"?>' +
      decl +
      `<w:document xmlns:w="${W_NS}"><w:body>${para('hi')}</w:body></w:document>`;
    const targets: { part: string; wrap: (decl: string) => string; mode: 'update' | 'add' }[] = [
      { part: 'word/document.xml', wrap: evilDoc, mode: 'update' },
      { part: 'word/styles.xml', wrap: (d) => `<?xml version="1.0"?>${d}<w:styles xmlns:w="${W_NS}"/>`, mode: 'update' },
      { part: 'word/numbering.xml', wrap: (d) => `<?xml version="1.0"?>${d}<w:numbering xmlns:w="${W_NS}"/>`, mode: 'add' },
      {
        part: 'word/_rels/document.xml.rels',
        wrap: (d) =>
          `<?xml version="1.0"?>${d}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
        mode: 'update',
      },
      { part: 'word/header1.xml', wrap: (d) => `<?xml version="1.0"?>${d}<w:hdr xmlns:w="${W_NS}"/>`, mode: 'add' },
      { part: 'word/footer1.xml', wrap: (d) => `<?xml version="1.0"?>${d}<w:ftr xmlns:w="${W_NS}"/>`, mode: 'add' },
      {
        part: 'word/charts/chart1.xml',
        wrap: (d) => `<?xml version="1.0"?>${d}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`,
        mode: 'add',
      },
      {
        part: 'word/theme/theme1.xml',
        wrap: (d) => `<?xml version="1.0"?>${d}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`,
        mode: 'add',
      },
      { part: 'word/footnotes.xml', wrap: (d) => `<?xml version="1.0"?>${d}<w:footnotes xmlns:w="${W_NS}"/>`, mode: 'add' },
      {
        part: 'docProps/core.xml',
        wrap: (d) => `<?xml version="1.0"?>${d}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`,
        mode: 'add',
      },
    ];
    for (let i = 0; i < 22; i++) {
      const t = targets[Math.floor(rnd() * targets.length)];
      const decl = DECL_VARIANTS[Math.floor(rnd() * DECL_VARIANTS.length)];
      const payload = t.wrap(decl);
      cases.push({
        name: `decl/${t.part}#${i}`,
        kind: 'decl',
        buffer: mutateViaZip(t.part, (z) => {
          const data = Buffer.from(payload, 'utf8');
          if (t.mode === 'update' && t.part !== 'word/document.xml' && t.part !== 'word/styles.xml') {
            z.addFile(t.part, data);
          } else if (t.mode === 'add') {
            z.addFile(t.part, data);
          } else {
            z.updateFile(t.part, data);
          }
        }),
      });
    }
    // Adjacent markup that must stay inert at the XML layer (valid outcomes).
    cases.push({
      name: 'decl/xinclude-stays-valid',
      kind: 'decl',
      buffer: mutateViaZip('xinclude', (z) =>
        z.updateFile(
          'word/document.xml',
          Buffer.from(
            docXmlWithBody(
              `${para('hi')}<w:p><w:r><w:t>ref</w:t></w:r></w:p>`.replace(
                '</w:body>',
                '<xi:include xmlns:xi="http://www.w3.org/2001/XInclude" href="other.xml"/></w:body>',
              ),
            ),
            'utf8',
          ),
        ),
      ),
    });
    cases.push({
      name: 'decl/comment-heavy-stays-valid',
      kind: 'decl',
      buffer: mutateViaZip('comments', (z) =>
        z.updateFile(
          'word/document.xml',
          Buffer.from(
            `<?xml version="1.0"?><!-- note one --><w:document xmlns:w="${W_NS}"><w:body><!-- c2 -->${para('hi')}<!-- c3 --></w:body></w:document>`,
            'utf8',
          ),
        ),
      ),
    });
  }

  // 3. Traversal / hostile entry names (raw builder keeps them verbatim).
  {
    const rnd = mulberry32(SEED ^ 0x03);
    const entries = baseEntries();
    const hostileNames = [
      '../../evil.txt',
      '..\\evil.txt',
      '/abs/evil.txt',
      'C:\\win\\evil.txt',
      'C:/win/evil.txt',
      'word/../../evil2.txt',
      'a/b/../../../evil3.txt',
      '..%2f..%2fevil.txt',
      'word/media/payload.exe',
      'payload.vbs',
      'word/media/run.bat',
      'evil\x00.txt',
      'word/ctl\x01name.txt',
      'word/' + 'n'.repeat(180) + '.txt',
      'word/héllo–文档.txt',
      '.../dots.txt',
    ];
    for (let i = 0; i < hostileNames.length; i++) {
      const nm = hostileNames[i];
      const filler = Buffer.from(`hostile-${i}`, 'utf8');
      cases.push({
        name: `traversal/raw#${i}`,
        kind: 'traversal',
        buffer: buildRawZip([...entries, { name: nm, data: filler }]),
      });
    }
    // Same shapes through the tolerant writer (names may normalize there).
    const viaWriter = ['../../w-evil.txt', 'word/media/w-payload.exe', '..\\w-evil.txt'];
    for (let i = 0; i < viaWriter.length; i++) {
      const nm = viaWriter[i];
      cases.push({
        name: `traversal/writer#${i}`,
        kind: 'traversal',
        buffer: mutateViaZip(nm, (z) => z.addFile(nm, Buffer.from(`w-${i}`))),
      });
    }
    void rnd;
  }

  // 4. Empty / huge / missing parts.
  {
    const rnd = mulberry32(SEED ^ 0x04);
    const entries = baseEntries();
    const emptyDoc = entries.map((e) =>
      e.name === 'word/document.xml' ? { name: e.name, data: Buffer.alloc(0) } : e,
    );
    cases.push({ name: 'size/empty-document-xml', kind: 'size', buffer: buildRawZip(emptyDoc) });
    const emptyStyles = entries.map((e) =>
      e.name === 'word/styles.xml' ? { name: e.name, data: Buffer.alloc(0) } : e,
    );
    cases.push({ name: 'size/empty-styles-xml', kind: 'size', buffer: buildRawZip(emptyStyles) });
    cases.push({
      name: 'size/whitespace-document-xml',
      kind: 'size',
      buffer: mutateViaZip('ws', (z) => z.updateFile('word/document.xml', Buffer.from('  \n  ', 'utf8'))),
    });
    cases.push({
      name: 'size/text-document-xml',
      kind: 'size',
      buffer: mutateViaZip('text', (z) => z.updateFile('word/document.xml', Buffer.from('just some text', 'utf8'))),
    });
    cases.push({ name: 'size/zero-entry-archive', kind: 'size', buffer: buildRawZip([]) });
    cases.push({
      name: 'size/missing-content-types',
      kind: 'size',
      buffer: mutateViaZip('noct', (z) => z.deleteFile('[Content_Types].xml')),
    });
    cases.push({
      name: 'size/missing-styles',
      kind: 'size',
      buffer: mutateViaZip('nostyles', (z) => z.deleteFile('word/styles.xml')),
    });
    cases.push({
      name: 'size/missing-rels',
      kind: 'size',
      buffer: mutateViaZip('norels', (z) => z.deleteFile('word/_rels/document.xml.rels')),
    });
    // Huge but incompressible paragraph (ratio ~1:1, must load).
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,;:';
    let huge = '';
    {
      const parts: string[] = [];
      for (let i = 0; i < 200000; i++) parts.push(alphabet[Math.floor(rnd() * alphabet.length)]);
      huge = parts.join('');
    }
    cases.push({
      name: 'size/huge-paragraph-200k',
      kind: 'size',
      buffer: mutateViaZip('huge', (z) =>
        z.updateFile('word/document.xml', Buffer.from(docXmlWithBody(para(huge)), 'utf8')),
      ),
    });
    // Huge stored (uncompressed) sidecar part via raw builder.
    {
      const bytes = Buffer.alloc(1024 * 1024);
      for (let i = 0; i < bytes.length; i++) bytes[i] = Math.floor(rnd() * 256);
      cases.push({
        name: 'size/huge-stored-1mib',
        kind: 'size',
        buffer: buildRawZip([...entries, { name: 'word/media/blob.bin', data: bytes }]),
      });
    }
    // Many small parts.
    cases.push({
      name: 'size/many-parts-150',
      kind: 'size',
      buffer: mutateViaZip('many', (z) => {
        for (let i = 0; i < 150; i++) z.addFile(`word/custom/p${i}.txt`, Buffer.from(`p${i}`));
      }),
    });
    // Duplicate entry name.
    cases.push({
      name: 'size/duplicate-name',
      kind: 'size',
      buffer: mutateViaZip('dup', (z) => {
        z.addFile('word/dup.txt', Buffer.from('one'));
        z.addFile('word/dup.txt', Buffer.from('two'));
      }),
    });
    // Garbage sidecar + garbage rels + garbage content-types (tolerated lanes).
    cases.push({
      name: 'size/garbage-rels',
      kind: 'size',
      buffer: mutateViaZip('grels', (z) =>
        z.updateFile('word/_rels/document.xml.rels', Buffer.from('not xml {{{', 'utf8')),
      ),
    });
    cases.push({
      name: 'size/garbage-content-types',
      kind: 'size',
      buffer: mutateViaZip('gct', (z) => z.updateFile('[Content_Types].xml', Buffer.from('garbage{{{', 'utf8'))),
    });
    cases.push({
      name: 'size/empty-sidecar',
      kind: 'size',
      buffer: buildRawZip([...entries, { name: 'word/custom/empty.txt', data: Buffer.alloc(0) }]),
    });
    cases.push({
      name: 'size/bom-document-xml',
      kind: 'size',
      buffer: mutateViaZip('bom', (z) =>
        z.updateFile(
          'word/document.xml',
          Buffer.concat([
            Buffer.from([0xef, 0xbb, 0xbf]),
            Buffer.from(docXmlWithBody(para('bom-hi')), 'utf8'),
          ]),
        ),
      ),
    });
  }

  // 5. Truncated / lying central-directory headers (ceiling lane).
  {
    const mk = (name: string, fn: Parameters<typeof patchCentralDirectory>[1]): FuzzCase => ({
      name: `ceiling/${name}`,
      kind: 'ceiling',
      buffer: patchCentralDirectory(base, fn),
    });
    cases.push(
      mk('total-over-200mib', (e) =>
        e.name === 'word/document.xml' ? { uncompressedSize: 210 * 1024 * 1024, compressedSize: 5 * 1024 * 1024 } : undefined,
      ),
      mk('entry-ratio-250to1', (e) =>
        e.name === 'word/styles.xml' ? { uncompressedSize: 5000, compressedSize: 20 } : undefined,
      ),
      mk('zero-size-nonzero-payload', (e) =>
        e.name === 'word/document.xml' ? { uncompressedSize: 0, compressedSize: 1024 } : undefined,
      ),
      mk('overall-ratio', () => ({ uncompressedSize: 50000, compressedSize: 600 })),
      mk('max-u32-sizes', (e) =>
        e.name === 'word/document.xml' ? { uncompressedSize: 0xffffffff, compressedSize: 0xffffffff } : undefined,
      ),
      mk('tiny-shrunk-styles', (e) =>
        e.name === 'word/styles.xml' ? { uncompressedSize: 8, compressedSize: 8 } : undefined,
      ),
    );
  }

  // 6. Odd content-types shapes.
  {
    const ct = (inner: string): string =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${inner}</Types>`;
    const full = ct(
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    );
    void full;
    const variants: [string, string][] = [
      ['empty', ''],
      ['whitespace', '   '],
      ['garbage', 'not xml at all {{{'],
      ['no-document-default', ct('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')],
      ['no-overrides', ct('<Default Extension="xml" ContentType="application/xml"/>')],
      ['doctype', '<?xml version="1.0"?><!DOCTYPE c><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
      ['entity', '<?xml version="1.0"?><!ENTITY x "y"><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
      ['unclosed', '<Types><Default Extension="xml"'],
    ];
    for (const [label, payload] of variants) {
      cases.push({
        name: `contenttype/${label}`,
        kind: 'contenttype',
        buffer: mutateViaZip(label, (z) => z.updateFile('[Content_Types].xml', Buffer.from(payload, 'utf8'))),
      });
    }
  }

  // 7. Deep nesting shapes.
  {
    const rnd = mulberry32(SEED ^ 0x07);
    for (const depth of [50, 150, 400]) {
      let inner = para('deep');
      for (let i = 0; i < depth; i++) inner = `<w:p>${inner}</w:p>`;
      cases.push({
        name: `nesting/xml-depth-${depth}`,
        kind: 'nesting',
        buffer: mutateViaZip(`depth${depth}`, (z) =>
          z.updateFile('word/document.xml', Buffer.from(docXmlWithBody(inner), 'utf8')),
        ),
      });
    }
    let deepPath = 'word';
    for (let i = 0; i < 60; i++) deepPath += `/d${i}`;
    cases.push({
      name: 'nesting/dir-depth-60',
      kind: 'nesting',
      buffer: mutateViaZip('deepdir', (z) => z.addFile(`${deepPath}/leaf.txt`, Buffer.from('leaf'))),
    });
    cases.push({
      name: 'nesting/tables-20',
      kind: 'nesting',
      buffer: mutateViaZip('tables', (z) => {
        let rows = '';
        for (let i = 0; i < 20; i++) {
          rows += `<w:tr><w:tc><w:p><w:r><w:t>c${i}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>d${i}</w:t></w:r></w:p></w:tc></w:tr>`;
        }
        z.updateFile('word/document.xml', Buffer.from(docXmlWithBody(`<w:tbl>${rows}</w:tbl>`), 'utf8'));
      }),
    });
    void rnd;
  }

  // 8. Valid controls (must parse to shaped output + round-trip).
  {
    for (const template of ['blank', 'novel', 'financial', 'report'] as const) {
      cases.push({
        name: `control/${template}`,
        kind: 'control',
        buffer: createBlankDocxBuffer(template),
      });
    }
  }

  return cases;
}

function sha256Hex(b: Buffer): string {
  return crypto.createHash('sha256').update(b).digest('hex');
}

// Coded-error contract for hostile inputs (all fuzz shapes carry E_* codes):
// - zip-format: raw archive layer rejects (empty/garbage/truncated bytes,
//   corrupt local headers) throw E_VALIDATION at the open boundary.
// - missing-part: package without the main part throws E_VALIDATION.
// - xml-parse: malformed part XML throws E_VALIDATION (no toolkit internals
//   leak uncoded).
// - crc / truncated-data: archive CRC mismatch or inflate short-read throws
//   E_VALIDATION naming the part (size-shaped lies still hit the E_LIMIT
//   bomb ceiling first).
// - integrity-gate: the save-gate OOXML block throws E_VALIDATION.
// Any uncoded throw below is a suite violation (no known-raw lane).

const CODED_RE = /E_VALIDATION|E_LIMIT|E_SECURITY/;

interface CaseOutcome {
  name: string;
  kind: string;
  verdict: 'valid' | 'coded' | 'violation';
  detail: string;
}

function classifyThrow(err: unknown): { verdict: 'coded' | 'violation'; detail: string } {
  const msg = err instanceof Error ? err.message : String(err);
  const short = msg.slice(0, 160).replace(/\s+/g, ' ');
  if (CODED_RE.test(msg)) {
    const code = (['E_VALIDATION', 'E_LIMIT', 'E_SECURITY'] as const).find((c) => msg.includes(c)) ?? 'E_*';
    return { verdict: 'coded', detail: `${code}: ${short}` };
  }
  return { verdict: 'violation', detail: `UNCODED: ${short}` };
}

async function runCase(c: FuzzCase): Promise<CaseOutcome> {
  let parser: DocxParser;
  try {
    parser = new DocxParser(c.buffer);
  } catch (err) {
    const r = classifyThrow(err);
    return { name: c.name, kind: c.kind, verdict: r.verdict, detail: `ctor:${r.detail}` };
  }
  try {
    const summary = parser.getStructuralSummary() as unknown as {
      nodes?: { id?: unknown }[];
    };
    if (!summary || !Array.isArray(summary.nodes)) {
      return { name: c.name, kind: c.kind, verdict: 'violation', detail: 'INVALID SUMMARY SHAPE' };
    }
    for (const n of summary.nodes) {
      if (!n || typeof n.id !== 'string') {
        return { name: c.name, kind: c.kind, verdict: 'violation', detail: 'NODE WITHOUT STRING ID' };
      }
    }
  } catch (err) {
    const r = classifyThrow(err);
    return { name: c.name, kind: c.kind, verdict: r.verdict, detail: `summary:${r.detail}` };
  }
  try {
    const ast = (await parser.getAST()) as unknown as { sections?: unknown[] };
    if (!ast || !Array.isArray(ast.sections)) {
      return { name: c.name, kind: c.kind, verdict: 'violation', detail: 'INVALID AST SHAPE' };
    }
  } catch (err) {
    const r = classifyThrow(err);
    return { name: c.name, kind: c.kind, verdict: r.verdict, detail: `ast:${r.detail}` };
  }
  try {
    const out = parser.rezip();
    if (!Buffer.isBuffer(out) || out.length === 0) {
      return { name: c.name, kind: c.kind, verdict: 'violation', detail: 'EMPTY REZIP' };
    }
  } catch (err) {
    const r = classifyThrow(err);
    return { name: c.name, kind: c.kind, verdict: r.verdict, detail: `rezip:${r.detail}` };
  }
  return { name: c.name, kind: c.kind, verdict: 'valid', detail: 'parsed + rezip ok' };
}

useReproducibleOutput();

describe('seeded hostile-input suite for the document parser', () => {
  const outsideWrites: string[] = [];
  const tmpRoot = os.tmpdir();
  let spies: { mockRestore(): void }[] = [];

  beforeAll(() => {
    const guard = (target: unknown): void => {
      const p = String(target ?? '');
      if (p !== '' && !(p === tmpRoot || p.startsWith(tmpRoot + '/'))) {
        outsideWrites.push(p);
        throw new Error(`[E_SECURITY] test sandbox: write outside tmp blocked: ${p}`);
      }
    };
    const s1 = vi.spyOn(fs, 'writeFileSync').mockImplementation(((p: unknown) => {
      guard(p);
      return undefined;
    }) as typeof fs.writeFileSync);
    const s2 = vi.spyOn(fs, 'appendFileSync').mockImplementation(((p: unknown) => {
      guard(p);
      return undefined;
    }) as typeof fs.appendFileSync);
    const s3 = vi
      .spyOn(fs, 'createWriteStream')
      .mockImplementation(((p: unknown) => {
        guard(p);
        throw new Error('[E_SECURITY] test sandbox: streams disabled in suite');
      }) as unknown as typeof fs.createWriteStream);
    spies = [s1, s2, s3];
  });

  afterAll(() => {
    for (const s of spies) s.mockRestore();
  });

  it('disk extraction entry points stay disabled', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const zip = parser as unknown as Record<string, unknown>;
    void zip;
    const inner = (parser as unknown as { zip: Record<string, (...a: unknown[]) => unknown> }).zip;
    for (const method of ['extractAllTo', 'extractAllToAsync', 'extractEntryTo'] as const) {
      expect(() => (inner[method] as (...a: unknown[]) => unknown)(os.tmpdir())).toThrowError(/E_SECURITY/);
    }
  });

  it('corpus is deterministic for the fixed seed', () => {
    const a = buildCorpus();
    const b = buildCorpus();
    expect(a.length).toBe(b.length);
    const ha = sha256Hex(Buffer.concat(a.map((c) => c.buffer)));
    const hb = sha256Hex(Buffer.concat(b.map((c) => c.buffer)));
    expect(ha).toBe(hb);
    expect(a.map((c) => c.name)).toEqual(b.map((c) => c.name));
  });

  it('every hostile input is contained: coded error or shaped output, no outside writes', async () => {
    const corpus = buildCorpus();
    const started = performance.now();
    const outcomes: CaseOutcome[] = [];
    for (const c of corpus) {
      outcomes.push(await runCase(c));
    }
    const elapsedMs = performance.now() - started;

    const count = (v: CaseOutcome['verdict']): number => outcomes.filter((o) => o.verdict === v).length;
    const byCode = (code: string): number => outcomes.filter((o) => o.detail.includes(code)).length;
    const corpusHash = sha256Hex(Buffer.concat(corpus.map((c) => c.buffer)));
    const kinds = [...new Set(corpus.map((c) => c.kind))].sort();

    // eslint-disable-next-line no-console
    console.log(
      [
        `seed=0x${SEED.toString(16)} cases=${corpus.length} kinds=${kinds.join(',')}`,
        `corpus-sha256=${corpusHash}`,
        `valid=${count('valid')} coded=${count('coded')} violations=${count('violation')}`,
        `E_VALIDATION=${byCode('E_VALIDATION')} E_LIMIT=${byCode('E_LIMIT')} E_SECURITY=${byCode('E_SECURITY')}`,
        `runtime-ms=${Math.round(elapsedMs)}`,
      ].join('\n'),
    );

    expect(corpus.length).toBe(96);
    expect(elapsedMs).toBeLessThan(60000);
    expect(outsideWrites).toEqual([]);
    const violations = outcomes.filter((o) => o.verdict === 'violation');
    expect(violations.map((v) => `${v.name} :: ${v.detail}`)).toEqual([]);
    // Controls always parse.
    for (const o of outcomes.filter((o) => o.kind === 'control')) {
      expect(o.verdict).toBe('valid');
    }
  }, 55000);
});
