import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';

const PNG_1x1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/** Default entries of [Content_Types].xml as [extension, content type]. */
function defaultsOf(docx: Buffer): Array<[string, string]> {
  const xml = new AdmZip(docx).readAsText('[Content_Types].xml');
  return [...xml.matchAll(/<Default\s+[^>]*?Extension="([^"]*)"[^>]*?ContentType="([^"]*)"/g)].map((m) => [m[1], m[2]]);
}

/** A blank document whose content types declare the image extension in capitals, as Word on Windows can write it. */
async function blankWithUpperCaseDefault(ext: string, contentType: string): Promise<Buffer> {
  const blank = DocxParser.createNewDocument('blank');
  await blank.getAST();
  const zip = new AdmZip(blank.rezip());
  const types = zip.readAsText('[Content_Types].xml');
  // Drop any lower-case declaration of the same extension, then declare it in capitals.
  const stripped = types.replace(new RegExp(`<Default\\s+[^>]*Extension="${ext.toLowerCase()}"[^>]*/>`, 'g'), '');
  zip.updateFile(
    '[Content_Types].xml',
    Buffer.from(stripped.replace('</Types>', `<Default Extension="${ext}" ContentType="${contentType}"/></Types>`))
  );
  return zip.toBuffer();
}

describe('image content types treat the extension without regard to case', () => {
  it('an image insert does not declare png a second time when the package declares PNG', async () => {
    const input = await blankWithUpperCaseDefault('PNG', 'image/png');
    expect(defaultsOf(input).filter(([e]) => e.toLowerCase() === 'png')).toEqual([['PNG', 'image/png']]);
    expect((await validateDocxBuffer(input)).ok).toBe(true);

    const parser = new DocxParser(input);
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    const result = await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'image', operation: 'insert', anchor, fileData: PNG_1x1, width: 914400, height: 914400 }],
    });
    expect(result.success).toBe(true);
    const out = parser.rezip();
    expect(defaultsOf(out).filter(([e]) => e.toLowerCase() === 'png')).toEqual([['PNG', 'image/png']]);
    const res = await validateDocxBuffer(out);
    expect(res.errors).toEqual([]);
    expect(res.ok).toBe(true);
  });

  it('an image insert still declares png when the package declares no png extension', async () => {
    const blank = DocxParser.createNewDocument('blank');
    await blank.getAST();
    const zip = new AdmZip(blank.rezip());
    const types = zip.readAsText('[Content_Types].xml').replace(/<Default\s+[^>]*Extension="png"[^>]*\/>/gi, '');
    zip.updateFile('[Content_Types].xml', Buffer.from(types));
    const input = zip.toBuffer();
    expect(defaultsOf(input).some(([e]) => e.toLowerCase() === 'png')).toBe(false);

    const parser = new DocxParser(input);
    const ast = await parser.getAST();
    await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'image', operation: 'insert', anchor: ast.sections[0].nodes[0].id, fileData: PNG_1x1, width: 914400, height: 914400 }],
    });
    const out = parser.rezip();
    expect(defaultsOf(out).filter(([e]) => e.toLowerCase() === 'png')).toEqual([['png', 'image/png']]);
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('a save of a document with an existing .PNG media part does not add a second png declaration', async () => {
    const input0 = await blankWithUpperCaseDefault('PNG', 'image/png');
    const zip = new AdmZip(input0);
    zip.addFile('word/media/image1.PNG', Buffer.from(PNG_1x1, 'base64'));
    const rels = zip.readAsText('word/_rels/document.xml.rels');
    zip.updateFile(
      'word/_rels/document.xml.rels',
      Buffer.from(
        rels.replace(
          '</Relationships>',
          '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.PNG"/></Relationships>'
        )
      )
    );
    const input = zip.toBuffer();
    expect((await validateDocxBuffer(input)).ok).toBe(true);

    const parser = new DocxParser(input);
    const ast = await parser.getAST();
    await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: ast.sections[0].nodes[0].id, position: 'after', text: 'after the image part' }],
    });
    // Saving with the AST is the path that declares the content type of every media part it holds.
    const out = parser.rezip(ast);
    expect(defaultsOf(out).filter(([e]) => e.toLowerCase() === 'png')).toEqual([['PNG', 'image/png']]);
    const res = await validateDocxBuffer(out);
    expect(res.errors).toEqual([]);
    expect(res.ok).toBe(true);
  });
});
