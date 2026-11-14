/**
 * Small .docx builders for the reference-neutraliser and renderer-selection tests.
 * `buildDocx` makes a minimal package from raw part text; `rawZip` writes a stored (uncompressed) zip with
 * any entry names, including duplicates, which AdmZip cannot create.
 */
import * as zlib from 'zlib';
import { AdmZip } from '@attentionform/astra-engine';

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function documentXml(body: string, opts: { ns?: string; prefix?: string; defaultNs?: boolean } = {}): string {
  const ns = opts.ns ?? W_NS;
  const prefix = opts.prefix ?? 'w';
  if (opts.defaultNs) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><document xmlns="${ns}" xmlns:r="${R_NS}"><body>${body}<sectPr/></body></document>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${prefix}:document xmlns:${prefix}="${ns}" xmlns:r="${R_NS}"><${prefix}:body>${body}<${prefix}:sectPr/></${prefix}:body></${prefix}:document>`;
}

export function relsXml(rels: Array<{ id: string; type: string; target: string; mode?: string }>): string {
  const items = rels
    .map((r) => `<Relationship Id="${r.id}" Type="${r.type.includes('/') ? r.type : `${REL_TYPE}/${r.type}`}" Target="${r.target}"${r.mode ? ` TargetMode="${r.mode}"` : ''}/>`)
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}">${items}</Relationships>`;
}

/** A complex field written the way Word does it, with the instruction split over the given runs. */
export function complexField(instrParts: string[], result = 'x'): string {
  const runs = instrParts.map((p) => `<w:r><w:instrText xml:space="preserve">${p}</w:instrText></w:r>`).join('');
  return `<w:r><w:fldChar w:fldCharType="begin"/></w:r>${runs}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

export interface BuildOptions {
  body?: string;
  document?: string;
  rels?: Array<{ id: string; type: string; target: string; mode?: string }>;
  extra?: Record<string, string | Buffer>;
}

export function buildDocx(o: BuildOptions = {}): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )
  );
  zip.addFile('_rels/.rels', Buffer.from(relsXml([{ id: 'rId1', type: 'officeDocument', target: 'word/document.xml' }])));
  zip.addFile('word/document.xml', Buffer.from(o.document ?? documentXml(o.body ?? '<w:p><w:r><w:t>hello</w:t></w:r></w:p>')));
  if (o.rels) zip.addFile('word/_rels/document.xml.rels', Buffer.from(relsXml(o.rels)));
  for (const [name, data] of Object.entries(o.extra ?? {})) zip.addFile(name, Buffer.isBuffer(data) ? data : Buffer.from(data));
  return zip.toBuffer();
}

export function readPart(docx: Buffer, name: string): string {
  const e = new AdmZip(docx).getEntry(name);
  if (!e) throw new Error(`no part ${name}`);
  return e.getData().toString('utf8');
}

export function listParts(docx: Buffer): string[] {
  return new AdmZip(docx).getEntries().map((e) => e.entryName);
}

/** Stored zip with arbitrary (possibly duplicate) entry names. */
export function rawZip(entries: Array<{ name: string; data: Buffer }>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const name = Buffer.from(e.name, 'utf8');
    const crc = zlib.crc32(e.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(crc >>> 0, 14);
    local.writeUInt32LE(e.data.length, 18);
    local.writeUInt32LE(e.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    locals.push(local, name, e.data);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(crc >>> 0, 16);
    central.writeUInt32LE(e.data.length, 20);
    central.writeUInt32LE(e.data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, name);
    offset += local.length + name.length + e.data.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, centralBuf, end]);
}
