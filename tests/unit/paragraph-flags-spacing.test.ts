/**
 * 0.5.3: keepWithNext / keepLines / pageBreakBefore on the style op, the guide's preventOrphanHeadings,
 * and the space after a callout or a table that compose creates.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, parseStyleGuide, validateDocxBuffer, validateOperationStrict } from '@attentionform/astra-engine';
import { composeDocument } from '../../packages/engine/src/compose/markdown';

async function session() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const xml = () => new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
  const run = (operations: unknown[]) => engine.executeBatch({ operations } as any);
  const anchor = ast.sections[0].nodes[0].id as string;
  return { run, xml, anchor, parser, ast };
}
const para = (anchor: string, id: string, text = 'Body text') => ({ action: 'insert', type: 'paragraph', anchor, position: 'after', id, text });
const pPrOf = (xml: string, text: string) => {
  const p = xml.match(new RegExp(`<w:p[ >](?:(?!</w:p>)[\\s\\S])*?${text}[\\s\\S]*?</w:p>`));
  return p ? (p[0].match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '') : null;
};

describe('paragraph flags on the style op', () => {
  it('true writes the element, false writes an explicit off, null removes it', async () => {
    const s = await session();
    await s.run([para(s.anchor, 'p1', 'FlagPara')]);
    await s.run([{ action: 'style', target: 'p1', keepWithNext: true, keepLines: true, pageBreakBefore: true }]);
    let pPr = pPrOf(s.xml(), 'FlagPara')!;
    expect(pPr).toContain('<w:keepNext/>');
    expect(pPr).toContain('<w:keepLines/>');
    expect(pPr).toContain('<w:pageBreakBefore/>');
    await s.run([{ action: 'style', target: 'p1', keepWithNext: false, pageBreakBefore: false }]);
    pPr = pPrOf(s.xml(), 'FlagPara')!;
    expect(pPr).toContain('<w:keepNext w:val="0"/>');
    expect(pPr).toContain('<w:pageBreakBefore w:val="0"/>');
    expect(pPr).toContain('<w:keepLines/>'); // omitted key: unchanged
    await s.run([{ action: 'style', target: 'p1', keepWithNext: null, keepLines: null, pageBreakBefore: null }]);
    pPr = pPrOf(s.xml(), 'FlagPara') ?? '';
    expect(pPr).not.toMatch(/keepNext|keepLines|pageBreakBefore/);
    const docx = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(docx.errors).toEqual([]);
  });

  it('keeps the schema order next to other paragraph properties', async () => {
    const s = await session();
    await s.run([para(s.anchor, 'p1', 'OrderPara'), { action: 'style', target: 'p1', alignment: 'center', spacingAfter: 120, keepWithNext: true, pageBreakBefore: true }]);
    const pPr = pPrOf(s.xml(), 'OrderPara')!;
    expect(pPr.indexOf('<w:keepNext/>')).toBeLessThan(pPr.indexOf('<w:pageBreakBefore/>'));
    expect(pPr.indexOf('<w:pageBreakBefore/>')).toBeLessThan(pPr.indexOf('<w:spacing'));
    expect(pPr.indexOf('<w:spacing')).toBeLessThan(pPr.indexOf('<w:jc'));
    const docx = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(docx.errors).toEqual([]);
  });

  it('null on a paragraph without the property does not leave an empty pPr', async () => {
    const s = await session();
    await s.run([para(s.anchor, 'p1', 'PlainPara')]);
    const res = await s.run([{ action: 'style', target: 'p1', keepWithNext: null }]);
    expect(res.success).toBe(true);
    expect(pPrOf(s.xml(), 'PlainPara') ?? '').not.toContain('<w:pPr/>');
  });

  it('works inside params and properties, and for several targets', async () => {
    const s = await session();
    await s.run([para(s.anchor, 'p1', 'One'), para('p1', 'p2', 'Two')]);
    await s.run([{ action: 'style', target: ['p1', 'p2'], params: { keepLines: true }, properties: { pageBreakBefore: true } }]);
    for (const t of ['One', 'Two']) {
      const pPr = pPrOf(s.xml(), t)!;
      expect(pPr).toContain('<w:keepLines/>');
      expect(pPr).toContain('<w:pageBreakBefore/>');
    }
  });

  it('refuses a non-boolean before changing anything', async () => {
    const s = await session();
    await s.run([para(s.anchor, 'p1', 'StrictPara')]);
    const before = s.xml();
    const res: any = await s.run([{ action: 'style', target: 'p1', bold: true, keepWithNext: 'yes' }]);
    expect(res.success).toBe(false);
    expect(JSON.stringify(res)).toMatch(/keepWithNext must be true, false or null/);
    expect(s.xml()).toBe(before);
  });

  it('a style op that only has a null flag still counts as a change; the schema accepts true, false and null', () => {
    for (const v of [true, false, null]) {
      expect(() => validateOperationStrict({ action: 'style', target: 'p1', keepWithNext: v })).not.toThrow();
      expect(() => validateOperationStrict({ action: 'style', target: 'p1', params: { pageBreakBefore: v } })).not.toThrow();
    }
    expect(() => validateOperationStrict({ action: 'style', target: 'p1', keepLines: 'no' })).toThrow();
  });

  it('a flag works on a paragraph with a span match, and a heading style keepNext can be switched off', async () => {
    const s = await session();
    await s.run([para(s.anchor, 'p1', 'Alpha beta'), { action: 'style', target: 'p1', match: 'beta', bold: true, keepWithNext: true }]);
    const pPr = pPrOf(s.xml(), 'Alpha')!;
    expect(pPr).toContain('<w:keepNext/>');
  });
});

describe('compose: preventOrphanHeadings and spacing', () => {
  async function compose(md: string, guideText?: string) {
    const parser = DocxParser.createNewDocument('blank');
    const styleGuide = guideText ? parseStyleGuide(guideText).config : undefined;
    await composeDocument(parser, md, { styleGuide, allowedDir: process.cwd() });
    return new AdmZip(parser.rezip()).readAsText('word/document.xml');
  }
  const md = '## Heading\n\nBody under heading.\n';

  it('without a guide key the heading keeps with the next paragraph (the heading style sets it)', async () => {
    expect(pPrOf(await compose(md), 'Heading')).toContain('<w:keepNext/>');
  });

  it('Prevent Orphan Headings: false writes an explicit keepNext off on every heading', async () => {
    const xml = await compose(`# Title\n\n${md}`, '# Style Guide: t\n\n## Layout\n- Prevent Orphan Headings: false\n');
    expect(pPrOf(xml, 'Heading')).toContain('<w:keepNext w:val="0"/>');
    expect(pPrOf(xml, 'Title')).toContain('<w:keepNext w:val="0"/>');
    expect(pPrOf(xml, 'Body under')).not.toContain('keepNext');
  });

  it('Prevent Orphan Headings: true keeps the heading with the next paragraph', async () => {
    const pPr = pPrOf(await compose(md, '# Style Guide: t\n\n## Layout\n- Prevent Orphan Headings: true\n'), 'Heading')!;
    expect(pPr).toContain('<w:keepNext/>');
    expect(pPr).not.toContain('w:val="0"');
  });

  const callout = '> [!NOTE]\n> A note.\n\nAfter the note.\n';
  const spacingBeforeOf = (xml: string, text: string) => Number((pPrOf(xml, text)?.match(/w:before="(\d+)"/) ?? [])[1] ?? 0);

  it('the paragraph after a callout gets 8 pt of space before by default; after a table 12 pt as before', async () => {
    expect(spacingBeforeOf(await compose(callout), 'After the note')).toBe(160);
    expect(spacingBeforeOf(await compose('| a | b |\n| --- | --- |\n| 1 | 2 |\n\nAfter the table.\n'), 'After the table')).toBe(240);
  });

  it('a list and a heading after a callout get the space too, and the guide sets it', async () => {
    const xml = await compose('> [!NOTE]\n> A note.\n\n- First item\n- Second item\n', '# Style Guide: t\n\n## Callouts\n- Space After: 14 pt\n');
    expect(spacingBeforeOf(xml, 'First item')).toBe(280);
    // Word drops the space before a list item whose paragraph style matches the one before it unless contextual spacing is off.
    expect(pPrOf(xml, 'First item')).toContain('<w:contextualSpacing w:val="0"/>');
    expect(pPrOf(xml, 'Second item')).toContain('<w:contextualSpacing/>');
    expect(spacingBeforeOf(xml, 'Second item')).toBe(0);
    const t = await compose('| a | b |\n| --- | --- |\n| 1 | 2 |\n\nAfter.\n', '# Style Guide: t\n\n## Tables\n- Space After: 0 pt\n');
    expect(spacingBeforeOf(t, 'After.')).toBe(0);
  });

  it('only blocks this call creates change: a block before the callout and a block two after it do not', async () => {
    const xml = await compose('Before the note.\n\n> [!NOTE]\n> A note.\n\nFirst after.\n\nSecond after.\n');
    expect(spacingBeforeOf(xml, 'Before the note')).toBe(0);
    expect(spacingBeforeOf(xml, 'First after')).toBe(160);
    expect(spacingBeforeOf(xml, 'Second after')).toBe(0);
  });

  it('guide parsing: pt only, 0 to 72, warnings name the key', () => {
    const ok = parseStyleGuide('# Style Guide: t\n\n## Callouts\n- Space After: 10\n\n## Tables\n- Space After: 6 pt\n');
    expect(ok.config.callouts.spaceAfterPt).toBe(10);
    expect(ok.config.tables.spaceAfterPt).toBe(6);
    const bad = parseStyleGuide('# Style Guide: t\n\n## Callouts\n- Space After: 1 in\n- Space After: 100\n\n## Tables\n- Space After: wide\n');
    expect(bad.config.callouts.spaceAfterPt).toBeUndefined();
    expect(bad.config.tables.spaceAfterPt).toBeUndefined();
    expect(bad.warnings.join('\n')).toContain("callouts.spaceAfter unit 'in' is not supported (use pt)");
    expect(bad.warnings.join('\n')).toContain('callouts.spaceAfter \'100\' is outside 0 to 72 pt');
    expect(bad.warnings.join('\n')).toContain('tables.spaceAfter');
  });
});
