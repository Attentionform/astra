/**
 * DN3: macro strip on a package whose WordprocessingML prefix is not "w". The strip removes the control parts and the
 * references to them in the document; a reference written with another prefix must not survive as a dangling r:id.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, sanitizeDocxBuffer, stripMacroPartsFromBuffer } from '@attentionform/astra-engine';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const renameW = (xml: string): string =>
  xml.replace(/(<\/?)w:/g, '$1x:').replace(/(\s)w:([A-Za-z])/g, '$1x:$2').replace(/xmlns:w=/g, 'xmlns:x=');

function build(prefixed: boolean): Buffer {
  const z = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const doc = z.readAsText('word/document.xml').replace('<w:body>',
    '<w:body><w:p><w:r><w:pict><w:control r:id="rId90" w:name="C" w:shapeid="_x0000_s1"/></w:pict></w:r><w:r><w:t>kept text</w:t></w:r></w:p>');
  z.updateFile('word/document.xml', Buffer.from(prefixed ? renameW(doc) : doc, 'utf8'));
  z.updateFile('word/_rels/document.xml.rels', Buffer.from(z.readAsText('word/_rels/document.xml.rels')
    .replace('</Relationships>', `<Relationship Id="rId90" Type="${REL}/control" Target="activeX/activeX1.xml"/></Relationships>`), 'utf8'));
  z.addFile('word/activeX/activeX1.xml', Buffer.from(`<?xml version="1.0" encoding="UTF-8"?><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="${REL}" ax:classid="{868FBE40-B964-453C-BEA2-77F32F2C294F}" ax:persistence="persistStorage" r:id="rId1"/>`, 'utf8'));
  z.addFile('word/activeX/activeX1.bin', Buffer.from('binary'));
  z.addFile('word/activeX/_rels/activeX1.xml.rels', Buffer.from(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/></Relationships>`, 'utf8'));
  z.updateFile('[Content_Types].xml', Buffer.from(z.readAsText('[Content_Types].xml').replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="xml"'), 'utf8'));
  return z.toBuffer();
}
const view = (buf: Buffer) => {
  const z = new AdmZip(buf);
  const doc = z.readAsText('word/document.xml');
  return {
    parts: z.getEntries().map((e) => e.entryName).filter((n) => n.includes('activeX')),
    danglingId: /r:id="rId90"/.test(doc),
    controlElement: /<\w+:control[\s/>]/.test(doc),
    text: doc.includes('kept text'),
    rels: z.readAsText('word/_rels/document.xml.rels').includes('rId90'),
  };
};

describe('macro strip, prefixed document', () => {
  const usual = view(stripMacroPartsFromBuffer(build(false)).buffer);
  it('the usual form is stripped completely (baseline)', () => {
    expect(usual).toEqual({ parts: [], danglingId: false, controlElement: false, text: true, rels: false });
  });
  it('stripMacroPartsFromBuffer: a prefixed document gives the same result', () => {
    expect(view(stripMacroPartsFromBuffer(build(true)).buffer)).toEqual(usual);
  });
  it('sanitizeDocxBuffer with stripMacros: a prefixed document gives the same result', () => {
    expect(view(sanitizeDocxBuffer(build(true), { stripMacros: true }).buffer)).toEqual(view(sanitizeDocxBuffer(build(false), { stripMacros: true }).buffer));
  });
  it('DocxParser with stripMacros opens the prefixed package and drops the control', () => {
    const parser = new DocxParser(build(true), { stripMacros: true });
    expect(view(parser.rezip())).toEqual(usual);
  });
});
