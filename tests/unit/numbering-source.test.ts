import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, NumberingManipulator } from '@attentionform/astra-engine';

/**
 * A numbering part keeps its root element (namespaces and mc:Ignorable), its numPicBullet and
 * numIdMacAtCleanup children, and is left as it was when no list changed.
 */
const RICH = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" mc:Ignorable="w14 w15 w16cid">
<w:numPicBullet w:numPicBulletId="0"><w:pict/></w:numPicBullet>
<w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0"><w:nsid w:val="1A2B3C4D"/><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>
<w:num w:numId="1" w16cid:durableId="123456"><w:abstractNumId w:val="0"/></w:num>
<w:numIdMacAtCleanup w:val="1"/>
</w:numbering>`;

function withNumbering(xml: string): Buffer {
  const z = new AdmZip(createBlankDocxBuffer('blank'));
  z.addFile('word/numbering.xml', Buffer.from(xml, 'utf8'));
  return z.toBuffer();
}
const part = (b: Buffer) => new AdmZip(b).getEntry('word/numbering.xml')?.getData().toString('utf8');

describe('numbering part source', () => {
  it('a save without edits keeps the numbering part byte for byte', async () => {
    const p = new DocxParser(withNumbering(RICH));
    await p.getAST();
    expect(part(p.rezip())).toBe(RICH);
  });
  it('a save with the AST and then a bare save keep the numbering part byte for byte', async () => {
    const p = new DocxParser(withNumbering(RICH));
    const ast = await p.getAST();
    expect(part(p.rezip(ast))).toBe(RICH);
    expect(part(p.rezip())).toBe(RICH);
  });
  it('a part rewritten after a list was created keeps the root, numPicBullet first and numIdMacAtCleanup last', async () => {
    const p = new DocxParser(withNumbering(RICH));
    const ast = await p.getAST();
    new NumberingManipulator(ast).getOrCreateList('numbered');
    const out = part(p.rezip())!;
    expect(out).toContain('mc:Ignorable="w14 w15 w16cid"');
    expect(out).toContain('xmlns:w16cid=');
    expect(out.indexOf('<w:numPicBullet')).toBeGreaterThan(-1);
    expect(out.indexOf('<w:numPicBullet')).toBeLessThan(out.indexOf('<w:abstractNum'));
    expect(out.lastIndexOf('<w:numIdMacAtCleanup')).toBeGreaterThan(out.lastIndexOf('<w:num '));
    expect(out).not.toBe(RICH);
  });
  it('an orphan made after load is still pruned', async () => {
    const p = new DocxParser(withNumbering(RICH));
    const ast = await p.getAST();
    delete ast.numbering.concreteNum['1'];
    expect(part(p.rezip(ast))).toBeUndefined();
  });
  it('a rolled back batch leaves the numbering part byte for byte', async () => {
    const p = new DocxParser(withNumbering(RICH));
    const ast = await p.getAST();
    const engine = new OperationEngine(ast, p);
    const P1 = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: P1, position: 'before', text: 'x' },
        { action: 'move', target: P1, anchor: 'nope_missing', position: 'after' },
      ],
      sequential: true,
    });
    expect(res.rollbackTriggered).toBe(true);
    expect(part(p.rezip())).toBe(RICH);
  });
  it('a successful text batch leaves the numbering part byte for byte', async () => {
    const p = new DocxParser(withNumbering(RICH));
    const ast = await p.getAST();
    const engine = new OperationEngine(ast, p);
    const P1 = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: P1, position: 'before', text: 'hello' }],
    });
    expect(res.success).toBe(true);
    expect(part(p.rezip())).toBe(RICH);
  });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function pkg(numbering: string): Buffer {
  const z = new AdmZip(createBlankDocxBuffer('blank'));
  z.addFile('word/numbering.xml', Buffer.from(numbering, 'utf8'));
  let rels = z.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
  if (!rels.includes('relationships/numbering')) {
    rels = rels.replace('</Relationships>', '<Relationship Id="rIdNum1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
    z.updateFile('word/_rels/document.xml.rels', Buffer.from(rels, 'utf8'));
  }
  let ct = z.getEntry('[Content_Types].xml')!.getData().toString('utf8');
  if (!ct.includes('/word/numbering.xml')) {
    ct = ct.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
    z.updateFile('[Content_Types].xml', Buffer.from(ct, 'utf8'));
  }
  return z.toBuffer();
}
const get = (b: Buffer, n: string) => new AdmZip(b).getEntry(n)?.getData().toString('utf8');
it('an empty numbering part stays with its relationship and content type', async () => {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"/>`;
  const p = new DocxParser(pkg(xml)); await p.getAST();
  const out = p.rezip();
  expect(get(out, 'word/numbering.xml')).toBe(xml);
  expect(get(out, 'word/_rels/document.xml.rels')).toContain('relationships/numbering');
  expect(get(out, '[Content_Types].xml')).toContain('/word/numbering.xml');
});
it('an orphan part on load stays until a list operation touches it', async () => {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="11111111"/><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum></w:numbering>`;
  const p = new DocxParser(pkg(xml)); await p.getAST();
  expect(get(p.rezip(), 'word/numbering.xml')).toBe(xml);
});
it('a plain root is kept when untouched and starts the rewritten part when a list is created', async () => {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="11111111"/><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
  const p = new DocxParser(pkg(xml)); const ast = await p.getAST();
  expect(get(p.rezip(), 'word/numbering.xml')).toBe(xml);
  new NumberingManipulator(ast).getOrCreateList('bullet');
  const out = get(p.rezip(), 'word/numbering.xml')!;
  expect(out.startsWith(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}">`)).toBe(true);
  expect(out).not.toBe(xml);
});
