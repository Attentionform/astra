/**
 * 0.5.3: the macro strip reads relationships by local name, either quote style and decoded attribute values, so a
 * relationships part that binds the package namespace to a prefix, or spells a target with a character reference,
 * cannot hide a macro or control relationship from it.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, stripMacroPartsFromBuffer } from '@attentionform/astra-engine';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const prefixed = (xml: string): string => xml
  .replace(/<Relationships\b/, '<p:Relationships').replace(/<\/Relationships>/, '</p:Relationships>')
  .replace(/xmlns=/, 'xmlns:p=').replace(/<Relationship\b/g, '<p:Relationship');

const ACTIVEX_BIN_REL = (target: string, quote: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PKG}"><Relationship Id=${quote}rId1${quote} Type=${quote}http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary${quote} Target=${quote}${target}${quote}/></Relationships>`;

function build(opts: { prefixRels: boolean; target: string; quote: string }): Buffer {
  const z = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const doc = z.readAsText('word/document.xml');
  z.updateFile('word/document.xml', Buffer.from(doc.replace('<w:body>',
    '<w:body><w:p><w:r><w:pict><w:control r:id="rId90" w:name="C" w:shapeid="_x0000_s1"/></w:pict></w:r><w:r><w:t>kept</w:t></w:r></w:p>'), 'utf8'));
  const rels = z.readAsText('word/_rels/document.xml.rels')
    .replace('</Relationships>', `<Relationship Id="rId90" Type="${REL}/control" Target="activeX/activeX1.xml"/></Relationships>`);
  z.updateFile('word/_rels/document.xml.rels', Buffer.from(opts.prefixRels ? prefixed(rels) : rels, 'utf8'));
  z.addFile('word/activeX/activeX1.xml', Buffer.from(
    `<?xml version="1.0" encoding="UTF-8"?><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="${REL}" ax:classid="{868FBE40-B964-453C-BEA2-77F32F2C294F}" ax:persistence="persistStorage" r:id="rId1"/>`, 'utf8'));
  z.addFile('word/activeX/activeX1.bin', Buffer.from('binary'));
  const binRels = ACTIVEX_BIN_REL(opts.target, opts.quote);
  z.addFile('word/activeX/_rels/activeX1.xml.rels', Buffer.from(opts.prefixRels ? prefixed(binRels) : binRels, 'utf8'));
  const ct = z.readAsText('[Content_Types].xml');
  z.updateFile('[Content_Types].xml', Buffer.from(ct.replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="xml"'), 'utf8'));
  return z.toBuffer();
}

const cases: Array<[string, { prefixRels: boolean; target: string; quote: string }]> = [
  ['plain', { prefixRels: false, target: 'activeX1.bin', quote: '"' }],
  ['prefixed relationships part', { prefixRels: true, target: 'activeX1.bin', quote: '"' }],
  ['single-quoted attributes', { prefixRels: false, target: 'activeX1.bin', quote: "'" }],
  ['character reference in the target', { prefixRels: false, target: 'activeX&#49;.bin', quote: '"' }],
  ['percent-escape in the target', { prefixRels: false, target: 'activeX%31.bin', quote: '"' }],
  ['prefixed, single-quoted and escaped together', { prefixRels: true, target: 'activeX&#49;.bin', quote: "'" }],
];

describe('macro strip with relationship spellings the plain reader would miss', () => {
  for (const [label, opts] of cases) {
    it(`${label}: the control, its binary and every reference to them are removed`, () => {
      const { buffer, removed } = stripMacroPartsFromBuffer(build(opts));
      const names = new AdmZip(buffer).getEntries().map((e) => e.entryName);
      expect(removed).toContain('word/activeX/activeX1.bin');
      expect(names.filter((n) => /activeX/i.test(n))).toEqual([]);
      const zip = new AdmZip(buffer);
      expect(zip.readAsText('word/_rels/document.xml.rels')).not.toContain('rId90');
      expect(zip.readAsText('word/_rels/document.xml.rels')).not.toMatch(/activeX/i);
      expect(zip.readAsText('word/document.xml')).not.toContain('w:control');
      expect(zip.readAsText('word/document.xml')).toContain('kept');
      expect(zip.readAsText('[Content_Types].xml')).not.toMatch(/activeX/i);
    });
  }
});
