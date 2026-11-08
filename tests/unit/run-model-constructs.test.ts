import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { XMLSerializer } from '@xmldom/xmldom';

async function fromBody(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const node = ast.sections[0].nodes[0];
  return { parser, ast, engine, id: node.id, text: node.metadata.text as string };
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const SECOND = '<w:p><w:r><w:t>second</w:t></w:r></w:p>';

const VML_BOX = '<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="s1" style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>';

// The form Word writes: mc:AlternateContent is a child of the run, a drawing text box in
// mc:Choice and a VML copy in mc:Fallback.
const AC_BOX = '<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">'
  + '<mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
  + '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="2" name="tb"/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>'
  + '<wps:txbx><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
  + '<mc:Fallback><w:pict><v:shape id="s1" style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>';

const count = (xml: string, needle: string) => xml.split(needle).length - 1;
const bodyXml = (parser: DocxParser) => new XMLSerializer().serializeToString(parser.getDocumentDom().getElementsByTagName('w:body')[0]);

describe('text boxes in a paragraph (FX-06)', () => {
  for (const [name, box, copies] of [['VML text box', VML_BOX, 1], ['mc:AlternateContent text box', AC_BOX, 2]] as const) {
    describe(name, () => {
      it('has a baseline that validates', async () => {
        const doc = await fromBody(`<w:p>${box}${run('alpha beta gamma')}</w:p>${SECOND}`);
        expect((await validateDocxBuffer(doc.parser.rezip(doc.ast))).ok).toBe(true);
        expect(doc.text).toBe('boxed'.repeat(copies) + 'alpha beta gamma');
      });

      for (const tracked of [false, true]) {
        it(`${tracked ? 'tracked' : 'untracked'} replace outside the box keeps the box and its text`, async () => {
          const doc = await fromBody(`<w:p>${box}${run('alpha beta gamma')}</w:p>${SECOND}`);
          const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'boxed'.repeat(copies) + 'alpha BETA gamma', trackChanges: tracked }] } as never);
          expect(result.success, JSON.stringify(result.errors)).toBe(true);
          const xml = bodyXml(doc.parser);
          expect(count(xml, 'boxed')).toBe(copies);
          expect(xml).toContain('BETA');
          expect((await validateDocxBuffer(doc.parser.rezip(doc.ast))).ok).toBe(true);
        });
      }

      it('refuses a replace that would remove the box text and changes nothing', async () => {
        const doc = await fromBody(`<w:p>${box}${run('alpha beta gamma')}</w:p>${SECOND}`);
        const before = bodyXml(doc.parser);
        const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'alpha BETA gamma' }] } as never);
        expect(result.success).toBe(false);
        expect(JSON.stringify(result.errors)).toMatch(/text box|drawing/i);
        expect(bodyXml(doc.parser)).toBe(before);
      });

      it('tracked whole-paragraph delete keeps the box object and can be rejected', async () => {
        const doc = await fromBody(`<w:p>${box}${run('alpha beta gamma')}</w:p>${SECOND}`);
        const result = await doc.engine.executeBatch({ operations: [{ action: 'delete', target: doc.id, trackChanges: true }] } as never);
        expect(result.success, JSON.stringify(result.errors)).toBe(true);
        expect((await validateDocxBuffer(doc.parser.rezip(doc.ast))).ok).toBe(true);
        const rejected = await doc.engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'rejectAll' }] } as never);
        expect(rejected.success, JSON.stringify(rejected.errors)).toBe(true);
        const xml = bodyXml(doc.parser);
        expect(count(xml, 'boxed')).toBe(copies);
        expect(xml).toContain('alpha beta gamma');
        expect(xml).not.toContain('<w:del');
        expect((await validateDocxBuffer(doc.parser.rezip(doc.ast))).ok).toBe(true);
      });
    });
  }
});

describe('a deletion that contains an insertion', () => {
  it('can be edited around: the model and the stored text agree', async () => {
    const body = `<w:p>${run('keep this ')}<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">gone </w:delText></w:r><w:ins w:id="2" w:author="B" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inner</w:t></w:r></w:ins></w:del></w:p>${SECOND}`;
    const doc = await fromBody(body);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: doc.text.replace('keep this', 'keep that') }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const xml = bodyXml(doc.parser);
    expect(xml).toMatch(/<w:t[^>]*>that<\/w:t>/);
    expect(xml).toContain('inner');
    expect((await validateDocxBuffer(doc.parser.rezip(doc.ast))).ok).toBe(true);
  });
});

describe('deliberate refusals with a precise reason (rulings R-1 and R-2)', () => {
  it('R-1: a paragraph that holds the start of a field spanning paragraphs', async () => {
    const body = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run('Entry one') + '</w:p>'
      + '<w:p>' + run('Entry two') + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const doc = await fromBody(body);
    const before = bodyXml(doc.parser);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'Entry uno' }] } as never);
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/spans paragraphs.*table of contents/i);
    expect(bodyXml(doc.parser)).toBe(before);
  });

  it('R-1: a paragraph that holds the end of a field spanning paragraphs', async () => {
    const body = '<w:p>' + run('Entry two') + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + SECOND;
    const doc = await fromBody(body);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'Entry deux' }] } as never);
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/spans paragraphs.*table of contents/i);
  });

  it('R-2: a paragraph with an inline partial move', async () => {
    const stamp = 'w:author="A" w:date="2026-01-01T00:00:00Z"';
    const body = `<w:p><w:moveFrom w:id="1" ${stamp}><w:r><w:delText xml:space="preserve">moved clause </w:delText></w:r></w:moveFrom>${run('stable middle ')}<w:moveTo w:id="2" ${stamp}><w:r><w:t>arriving clause.</w:t></w:r></w:moveTo></w:p>${SECOND}`;
    const doc = await fromBody(body);
    const before = bodyXml(doc.parser);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'stable middle arriving clause!' }] } as never);
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/partial move.*(accept|reject)/i);
    expect(bodyXml(doc.parser)).toBe(before);
  });
});

describe('refusal allowlist over the repository fixtures (FX-06)', () => {
  it('every refused paragraph is refused for a deliberate reason', async () => {
    const { readdirSync, readFileSync, statSync } = await import('node:fs');
    const { join } = await import('node:path');
    const { buildRunModel } = await import('../../packages/engine/src/ooxml/run-model');
    const files: string[] = [];
    const walk = (dir: string): void => {
      for (const entry of readdirSync(dir)) {
        const full = join(dir, entry);
        if (statSync(full).isDirectory()) walk(full);
        else if (entry.endsWith('.docx')) files.push(full);
      }
    };
    walk(join(__dirname, '..', 'harness', 'fidelity'));
    expect(files.length).toBeGreaterThan(50);
    const allowed = [
      /spans paragraphs, such as a table of contents/, // R-1
      /a partial move/, // R-2
      /the paragraph text and text map differ/, // two hostile fixtures nest w:p inside w:t on purpose
    ];
    const unexpected: string[] = [];
    let refused = 0;
    for (const file of files) {
      let ast;
      try { ast = await new DocxParser(readFileSync(file)).getAST(); } catch { continue; }
      for (const section of ast.sections) for (const node of section.nodes) {
        if (node.type !== 'paragraph' || !node.xml) continue;
        try { buildRunModel(node.xml as Element, node.metadata.text as string); } catch (error) {
          refused++;
          const message = String((error as Error).message);
          if (!allowed.some(re => re.test(message))) unexpected.push(`${file}: ${message}`);
        }
      }
    }
    expect(unexpected).toEqual([]);
    expect(refused).toBeLessThan(15);
  });
});
