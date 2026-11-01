import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  sanitizeDocxBuffer,
  stripMacroPartsFromBuffer,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const O_NS = 'urn:schemas-microsoft-com:office:office';

interface Fixture {
  body: string;
  rels?: string[];
  parts?: Record<string, string>;
}

/** Blank document plus the given body XML, relationships and extra parts. */
function fixture(f: Fixture): Buffer {
  const z = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const doc = z.getEntry('word/document.xml')!.getData().toString('utf8');
  z.updateFile('word/document.xml', Buffer.from(doc.replace('<w:body>', `<w:body>${f.body}`), 'utf8'));
  if (f.rels) {
    const rels = z.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
    z.updateFile(
      'word/_rels/document.xml.rels',
      Buffer.from(rels.replace('</Relationships>', `${f.rels.join('')}</Relationships>`), 'utf8')
    );
  }
  for (const [name, text] of Object.entries(f.parts ?? {})) z.addFile(name, Buffer.from(text, 'utf8'));
  const ct = z.getEntry('[Content_Types].xml')!.getData().toString('utf8');
  z.updateFile(
    '[Content_Types].xml',
    Buffer.from(ct.replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="xml"'), 'utf8')
  );
  return z.toBuffer();
}

const text = (buf: Buffer, name: string): string => new AdmZip(buf).getEntry(name)!.getData().toString('utf8');
const names = (buf: Buffer): string[] => new AdmZip(buf).getEntries().map((e) => e.entryName);

const ACTIVEX_PARTS = {
  'word/activeX/activeX1.xml':
    `<?xml version="1.0" encoding="UTF-8"?><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="${REL}" ax:classid="{868FBE40-B964-453C-BEA2-77F32F2C294F}" ax:persistence="persistStorage" r:id="rId1"/>`,
  'word/activeX/activeX1.bin': 'binary',
  'word/activeX/_rels/activeX1.xml.rels':
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/></Relationships>',
};
const CONTROL_REL = `<Relationship Id="rId90" Type="${REL}/control" Target="activeX/activeX1.xml"/>`;

describe('strip-macros leaves no reference to a removed part', () => {
  it('removes a w:control and keeps the picture next to it; the validator does not crash', async () => {
    const input = fixture({
      body:
        '<w:p><w:r><w:t>before</w:t></w:r><w:r><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t201" style="width:1pt;height:1pt"/>' +
        '<w:control r:id="rId90" w:name="Control 1" w:shapeid="_x0000_s1026"/></w:pict></w:r><w:r><w:t>after</w:t></w:r></w:p>',
      rels: [CONTROL_REL],
      parts: ACTIVEX_PARTS,
    });
    const { buffer, removed, edited } = stripMacroPartsFromBuffer(input);
    expect(removed).toContain('word/activeX/activeX1.bin');
    expect(edited).toContain('word/document.xml');
    const doc = text(buffer, 'word/document.xml');
    expect(doc).not.toContain('w:control');
    expect(doc).toContain('v:shape');
    expect(doc).toContain('before');
    expect(doc).toContain('after');
    expect(text(buffer, 'word/_rels/document.xml.rels')).not.toContain('rId90');
    const res = await validateDocxBuffer(buffer);
    expect(res.errors.filter((e) => e.errorType === 'Exception')).toEqual([]);
  });

  it('drops a w:pict that only held the control', () => {
    const input = fixture({
      body: '<w:p><w:r><w:pict><w:control r:id="rId90" w:name="C" w:shapeid="_x0000_s1"/></w:pict></w:r><w:r><w:t>kept</w:t></w:r></w:p>',
      rels: [CONTROL_REL],
      parts: ACTIVEX_PARTS,
    });
    const doc = text(stripMacroPartsFromBuffer(input).buffer, 'word/document.xml');
    expect(doc).not.toContain('w:pict');
    expect(doc).toContain('kept');
  });

  it('removes every control when there are many', () => {
    const one = '<w:p><w:r><w:pict><w:control r:id="rId90" w:name="C" w:shapeid="_x0000_s1"/></w:pict></w:r></w:p>';
    const input = fixture({ body: one.repeat(40), rels: [CONTROL_REL], parts: ACTIVEX_PARTS });
    const doc = text(stripMacroPartsFromBuffer(input).buffer, 'word/document.xml');
    expect(doc).not.toContain('w:control');
  });

  it('removes c:externalData from a chart whose data link is an oleObject relationship', async () => {
    const chart =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="' +
      REL +
      '"><c:chart><c:plotArea><c:layout/></c:plotArea><c:plotVisOnly val="1"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>';
    const chartRels =
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' +
      REL +
      '/oleObject" Target="file:///C:\\data\\book.xlsx" TargetMode="External"/></Relationships>';
    const input = fixture({
      body: '<w:p><w:r><w:t>chart below</w:t></w:r></w:p>',
      parts: { 'word/charts/chart1.xml': chart, 'word/charts/_rels/chart1.xml.rels': chartRels },
    });
    const z = new AdmZip(input);
    z.updateFile(
      'word/_rels/document.xml.rels',
      Buffer.from(
        text(input, 'word/_rels/document.xml.rels').replace(
          '</Relationships>',
          `<Relationship Id="rId91" Type="${REL}/chart" Target="charts/chart1.xml"/></Relationships>`
        ),
        'utf8'
      )
    );
    const { buffer } = stripMacroPartsFromBuffer(z.toBuffer());
    const out = text(buffer, 'word/charts/chart1.xml');
    expect(out).not.toContain('externalData');
    expect(out).toContain('c:plotArea');
    expect(text(buffer, 'word/charts/_rels/chart1.xml.rels')).not.toContain('rId1');
  });

  it('removes only the o:OLEObject inside a w:object and keeps the preview shape and the field codes of siblings', async () => {
    const input = fixture({
      body:
        `<w:p><w:r><w:t>x</w:t></w:r><w:r><w:object w:dxaOrig="100" w:dyaOrig="100"><v:shape id="_x0000_s1027" type="#_x0000_t75" style="width:1pt;height:1pt"/>` +
        `<o:OLEObject xmlns:o="${O_NS}" Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_s1027" DrawAspect="Content" ObjectID="_1" r:id="rId92"><o:FieldCodes>\\s</o:FieldCodes></o:OLEObject></w:object></w:r></w:p>`,
      rels: [`<Relationship Id="rId92" Type="${REL}/oleObject" Target="embeddings/oleObject1.bin"/>`],
      parts: { 'word/embeddings/oleObject1.bin': 'ole' },
    });
    const { buffer, removed } = stripMacroPartsFromBuffer(input);
    expect(removed).toContain('word/embeddings/oleObject1.bin');
    const doc = text(buffer, 'word/document.xml');
    expect(doc).not.toContain('OLEObject');
    expect(doc).toContain('<w:object');
    expect(doc).toContain('v:shape');
    expect(doc).toContain('<w:t>x</w:t>');
    const res = await validateDocxBuffer(buffer);
    expect(res.errors.filter((e) => e.errorType === 'Exception')).toEqual([]);
  });

  it('removes an o:OLEObject that sits directly in a w:pict (Word 97 style) and drops a container left empty', () => {
    const input = fixture({
      body:
        `<w:p><w:r><w:pict><v:shape id="_x0000_s1030" type="#_x0000_t75" style="width:1pt;height:1pt"/>` +
        `<o:OLEObject xmlns:o="${O_NS}" Type="Embed" ProgID="Word.Picture.8" ShapeID="_x0000_s1030" DrawAspect="Content" ObjectID="_3" r:id="rId96"/></w:pict></w:r>` +
        `<w:r><w:object><o:OLEObject xmlns:o="${O_NS}" Type="Embed" ProgID="Word.Picture.8" ShapeID="_x0000_s1031" DrawAspect="Content" ObjectID="_4" r:id="rId96"/></w:object></w:r></w:p>`,
      rels: [`<Relationship Id="rId96" Type="${REL}/oleObject" Target="embeddings/oleObject4.bin"/>`],
      parts: { 'word/embeddings/oleObject4.bin': 'ole' },
    });
    const doc = text(stripMacroPartsFromBuffer(input).buffer, 'word/document.xml');
    expect(doc).not.toContain('OLEObject');
    expect(doc).toContain('<w:pict>');
    expect(doc).not.toContain('<w:object');
  });

  it('refuses a carrier it does not know, naming the part and the element', () => {
    const input = fixture({
      body:
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="x"><a:blip r:embed="rId93"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
      rels: [`<Relationship Id="rId93" Type="${REL}/oleObject" Target="embeddings/oleObject2.bin"/>`],
      parts: { 'word/embeddings/oleObject2.bin': 'ole' },
    });
    expect(() => stripMacroPartsFromBuffer(input)).toThrow(/E_VALIDATION.*word\/document\.xml.*a:blip r:embed="rId93"/);
  });

  it('refuses a carrier that holds a text run, so no visible text is deleted', () => {
    const input = fixture({
      body:
        `<w:p><w:r><w:object><v:shape id="_x0000_s1028" type="#_x0000_t75"/>` +
        `<o:OLEObject xmlns:o="${O_NS}" Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_s1028" DrawAspect="Content" ObjectID="_2" r:id="rId94"><w:r><w:t>words</w:t></w:r></o:OLEObject></w:object></w:r></w:p>`,
      rels: [`<Relationship Id="rId94" Type="${REL}/oleObject" Target="embeddings/oleObject3.bin"/>`],
      parts: { 'word/embeddings/oleObject3.bin': 'ole' },
    });
    expect(() => stripMacroPartsFromBuffer(input)).toThrow(/E_VALIDATION.*word\/document\.xml.*holds text/);
  });

  it('does not touch a part that never referenced the removed relationship', () => {
    const input = fixture({
      body: '<w:p><w:r><w:t>plain</w:t></w:r></w:p>',
      rels: [`<Relationship Id="rId95" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`],
      parts: { 'word/vbaProject.bin': 'vba' },
    });
    const before = text(input, 'word/document.xml');
    const { buffer, edited } = stripMacroPartsFromBuffer(input);
    expect(text(buffer, 'word/document.xml')).toBe(before);
    expect(edited).not.toContain('word/document.xml');
  });

  it('works the same through open with stripMacros and through sanitize', () => {
    const input = fixture({
      body: '<w:p><w:r><w:pict><w:control r:id="rId90" w:name="C" w:shapeid="_x0000_s1"/></w:pict></w:r></w:p>',
      rels: [CONTROL_REL],
      parts: ACTIVEX_PARTS,
    });
    const parser = new DocxParser(input, { stripMacros: true });
    const out = parser.rezip();
    expect(text(out, 'word/document.xml')).not.toContain('w:control');
    expect(names(out).some((n) => n.startsWith('word/activeX/'))).toBe(false);
    const { buffer } = sanitizeDocxBuffer(input, { stripMacros: true });
    expect(text(buffer, 'word/document.xml')).not.toContain('w:control');
  });
});

describe('strip-macros on a macro-enabled document leaves a plain document package', () => {
  const VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
  const VBA_DATA_REL = 'http://schemas.microsoft.com/office/2006/relationships/wordVbaData';

  function docm(): Buffer {
    const z = new AdmZip(
      fixture({
        body: '<w:p><w:r><w:t>text</w:t></w:r></w:p>',
        rels: [`<Relationship Id="rId91" Type="${VBA_REL}" Target="vbaProject.bin"/>`],
        parts: {
          'word/vbaProject.bin': 'binary',
          'word/_rels/vbaProject.bin.rels':
            `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${VBA_DATA_REL}" Target="vbaData.xml"/></Relationships>`,
          'word/vbaData.xml': '<?xml version="1.0" encoding="UTF-8"?><wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>',
        },
      })
    );
    let ct = z.getEntry('[Content_Types].xml')!.getData().toString('utf8');
    ct = ct
      .replace('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', 'application/vnd.ms-word.document.macroEnabled.main+xml')
      .replace('<Default Extension="bin" ContentType="application/octet-stream"/>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>')
      .replace('</Types>', '<Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/></Types>');
    z.updateFile('[Content_Types].xml', Buffer.from(ct, 'utf8'));
    return z.toBuffer();
  }

  it('removes the VBA data part and the rels part of the project, and the main part is no longer macro-enabled', () => {
    const out = stripMacroPartsFromBuffer(docm()).buffer;
    const list = names(out);
    expect(list.filter((n) => /vba/i.test(n))).toEqual([]);
    const ct = text(out, '[Content_Types].xml');
    expect(ct).not.toMatch(/macroEnabled|vbaProject|vbaData/);
    expect(ct).toContain('wordprocessingml.document.main+xml');
    expect(text(out, 'word/_rels/document.xml.rels')).not.toContain('vbaProject');
  });
});
