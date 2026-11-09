import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

/**
 * A simple field (`w:fldSimple`) cannot sit inside `w:del`, and a `w:del` inside it deletes the cached result only
 * and leaves the field behind. A tracked edit that covers a whole simple field now rewrites that field as a complex
 * field (begin, instruction, separate, result, end) and tracks each part, the way Word does. The edit is refused, as
 * before, when the field result holds anything but plain runs or the edit cuts the field.
 */
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const SECOND = '<w:p><w:r><w:t>second</w:t></w:r></w:p>';
const STAMP = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';
const FIELD = '<w:fldSimple w:instr=" DATE "><w:r><w:rPr><w:b/></w:rPr><w:t>2026</w:t></w:r></w:fldSimple>';
const BODY = `<w:p>${r('before ')}${FIELD}${r(' after')}</w:p>${SECOND}`;

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const node = ast.sections[0].nodes[0];
  return { parser, ast, engine, id: node.id, text: node.metadata.text as string };
}
type Doc = Awaited<ReturnType<typeof open>>;

const xmlOf = (parser: DocxParser) => new AdmZip(parser.rezip({ skipValidation: true })).readAsText('word/document.xml');
const paragraph = (xml: string, index = 0) => xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g)![index];
const textOf = (xml: string, index = 0) => (paragraph(xml, index).match(/<w:t(?:\s[^>]*)?>[^<]*/g) ?? []).map(t => t.replace(/<[^>]*>/g, '')).join('');
const count = (xml: string, needle: string) => xml.split(needle).length - 1;
const errorIds = async (parser: DocxParser) => ((await validateDocxBuffer(parser.rezip({ skipValidation: true }))).errors ?? []).map(e => e.id ?? '');
const revisionIds = (xml: string) => Array.from(xml.matchAll(/<w:(?:ins|del) [^>]*w:id="(\d+)"/g), m => Number(m[1]));

type Operation = Record<string, unknown>;
const EDITS: Record<string, (doc: Doc) => Operation> = {
  'tracked delete of the field result': doc => ({ action: 'trackedChange', operation: 'delete', target: doc.id, text: '2026' }),
  'tracked delete of the whole paragraph': doc => ({ action: 'delete', target: doc.id, trackChanges: true }),
  'tracked replace that covers the field': doc => ({ action: 'replace', target: doc.id, text: 'before X after', trackChanges: true }),
};

async function run(doc: Doc, operation: Operation) {
  return doc.engine.executeBatch({ operations: [operation] } as never);
}
/** Saves, opens the saved file as a fresh session, resolves every revision there and returns what it holds. */
async function afterResolve(parser: DocxParser, operation: 'acceptAll' | 'rejectAll') {
  const saved = new DocxParser(parser.rezip({ skipValidation: true }));
  const ast = await saved.getAST();
  const engine = new OperationEngine(ast, saved);
  const result = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation }] } as never);
  expect(result.success, JSON.stringify(result.errors)).toBe(true);
  return { xml: xmlOf(saved), errors: await errorIds(saved) };
}

describe('the baseline is valid', () => {
  it('a paragraph with a simple field validates and counts the cached result as text', async () => {
    const doc = await open(BODY);
    expect(doc.text).toBe('before 2026 after');
    expect(await errorIds(doc.parser)).toEqual([]);
  });
});

for (const [label, build] of Object.entries(EDITS)) {
  describe(label, () => {
    it('succeeds, writes a complex field and no simple field, and the file validates', async () => {
      const doc = await open(BODY);
      const result = await run(doc, build(doc));
      expect(result.success, JSON.stringify(result.errors)).toBe(true);
      const xml = xmlOf(doc.parser);
      expect(count(xml, '<w:fldSimple')).toBe(0);
      expect(count(xml, 'w:fldCharType="begin"')).toBe(1);
      expect(count(xml, 'w:fldCharType="separate"')).toBe(1);
      expect(count(xml, 'w:fldCharType="end"')).toBe(1);
      expect(count(xml, '<w:delInstrText')).toBe(1);
      expect(await errorIds(doc.parser)).toEqual([]);
      expect((await validateDocxBuffer(doc.parser.rezip(doc.ast))).ok).toBe(true);
    });

    it('tracks every part of the field: five deletions (begin, instruction, separate, result, end) with unique ids', async () => {
      const doc = await open(BODY);
      const before = doc.ast.nextChangeId;
      const result = await run(doc, build(doc));
      expect(result.success, JSON.stringify(result.errors)).toBe(true);
      const xml = xmlOf(doc.parser);
      const field = paragraph(xml);
      // A self-closing w:del is the deleted paragraph mark; it holds nothing and is not a part of the field.
      const deletions = field.match(/<w:del (?:[^>]*[^/>])?>[\s\S]*?<\/w:del>/g) ?? [];
      // The whole-paragraph delete also wraps the text around the field; the other edits wrap the field only.
      const parts = deletions.filter(d => /<w:fldChar |<w:delInstrText|<w:delText>2026</.test(d));
      expect(parts.length).toBe(5);
      const ids = revisionIds(xml);
      expect(new Set(ids).size).toBe(ids.length);
      expect(doc.ast.nextChangeId - before).toBe(ids.length);
      expect(count(xml, '<w:delText')).toBeGreaterThanOrEqual(1);
      expect(xml).not.toMatch(/<w:t[ >][^<]*2026/);
    });

    it('reject all restores the field as a complex field with the same instruction, result and formatting', async () => {
      const doc = await open(BODY);
      expect((await run(doc, build(doc))).success).toBe(true);
      const { xml, errors } = await afterResolve(doc.parser, 'rejectAll');
      expect(errors).toEqual([]);
      expect(textOf(xml)).toBe(doc.text);
      expect(count(xml, '<w:del ')).toBe(0);
      expect(count(xml, '<w:ins ')).toBe(0);
      expect(count(xml, '<w:delText')).toBe(0);
      expect(count(xml, '<w:delInstrText')).toBe(0);
      expect(xml).toMatch(/<w:instrText[^>]*> DATE <\/w:instrText>/);
      expect(count(xml, 'w:fldCharType="begin"')).toBe(1);
      expect(count(xml, 'w:fldCharType="end"')).toBe(1);
      expect(paragraph(xml)).toMatch(/<w:rPr><w:b\/><\/w:rPr><w:t[^>]*>2026<\/w:t>/);
    });

    it('accept all removes the whole field and leaves a valid file', async () => {
      const doc = await open(BODY);
      expect((await run(doc, build(doc))).success).toBe(true);
      const { xml, errors } = await afterResolve(doc.parser, 'acceptAll');
      expect(errors).toEqual([]);
      expect(xml).not.toMatch(/fldChar|instrText|fldSimple|2026/);
      // Accepting the delete of the whole paragraph removes it, so the first paragraph left is the second one.
      expect(textOf(xml)).toBe(label.includes('whole paragraph') ? 'second' : label.includes('replace') ? 'before X after' : 'before  after');
    });
  });
}

describe('the converted field keeps what the simple field had', () => {
  it('carries w:fldLock and w:dirty onto the begin marker', async () => {
    const doc = await open(`<w:p>${r('x ')}<w:fldSimple w:instr=" PAGE " w:fldLock="1" w:dirty="true"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>${SECOND}`);
    expect((await run(doc, { action: 'trackedChange', operation: 'delete', target: doc.id, text: '3' })).success).toBe(true);
    expect(xmlOf(doc.parser)).toMatch(/<w:fldChar [^>]*w:fldCharType="begin"[^>]*>/);
    const begin = xmlOf(doc.parser).match(/<w:fldChar [^>]*w:fldCharType="begin"[^>]*>/)![0];
    expect(begin).toContain('w:fldLock="1"');
    expect(begin).toContain('w:dirty="true"');
  });

  it('copies the result formatting to the markers, without a tracked formatting change, and keeps that change on the result', async () => {
    const field = '<w:fldSimple w:instr=" DATE "><w:r><w:rPr><w:b/><w:rPrChange w:id="900" w:author="Old" w:date="2025-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>2026</w:t></w:r></w:fldSimple>';
    const doc = await open(`<w:p>${r('before ')}${field}${r(' after')}</w:p>${SECOND}`);
    expect((await run(doc, { action: 'trackedChange', operation: 'delete', target: doc.id, text: '2026' })).success).toBe(true);
    const xml = xmlOf(doc.parser);
    expect(count(xml, '<w:rPrChange')).toBe(1);
    const markers = (paragraph(xml).match(/<w:r>(?:(?!<\/w:r>)[\s\S])*?(?:fldChar|instrText|delInstrText)[\s\S]*?<\/w:r>/g) ?? []);
    expect(markers.length).toBe(4);
    for (const marker of markers) expect(marker).toContain('<w:rPr><w:b/></w:rPr>');
    expect(await errorIds(doc.parser)).toEqual([]);
  });

  it('keeps several result runs, in order, with their own formatting', async () => {
    const field = '<w:fldSimple w:instr=" AUTHOR "><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Ann </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Lee</w:t></w:r></w:fldSimple>';
    const doc = await open(`<w:p>${r('by ')}${field}${r(' today')}</w:p>${SECOND}`);
    expect(doc.text).toBe('by Ann Lee today');
    expect((await run(doc, { action: 'replace', target: doc.id, text: 'by nobody today', trackChanges: true })).success).toBe(true);
    const { xml, errors } = await afterResolve(doc.parser, 'rejectAll');
    expect(errors).toEqual([]);
    expect(textOf(xml)).toBe('by Ann Lee today');
    expect(paragraph(xml)).toMatch(/<w:rPr><w:i\/><\/w:rPr><w:t[^>]*>Ann <\/w:t>[\s\S]*<w:rPr><w:b\/><\/w:rPr><w:t[^>]*>Lee<\/w:t>/);
    expect(count(xml, ' AUTHOR ')).toBe(1);
  });

  it('an instruction with quotes, backslashes and angle brackets survives a reject unchanged', async () => {
    const instruction = ' IF "a" = "b" "x &lt; y" "\\* MERGEFORMAT" ';
    const doc = await open(`<w:p>${r('s ')}<w:fldSimple w:instr="${instruction.replace(/"/g, '&quot;')}"><w:r><w:t>y</w:t></w:r></w:fldSimple></w:p>${SECOND}`);
    expect((await run(doc, { action: 'trackedChange', operation: 'delete', target: doc.id, text: 'y' })).success).toBe(true);
    const { xml, errors } = await afterResolve(doc.parser, 'rejectAll');
    expect(errors).toEqual([]);
    expect(xml).toContain('IF "a" = "b" "x &lt; y" "\\* MERGEFORMAT"');
  });

  it('a field with no result becomes begin, instruction, separate and end, and rejects back', async () => {
    const doc = await open(`<w:p>${r('page ')}<w:fldSimple w:instr=" PAGE "/>${r(' end')}</w:p>${SECOND}`);
    expect(doc.text).toBe('page  end');
    const result = await run(doc, { action: 'delete', target: doc.id, trackChanges: true });
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const xml = xmlOf(doc.parser);
    expect(count(xml, '<w:fldSimple')).toBe(0);
    expect(count(xml, 'w:fldCharType="separate"')).toBe(1);
    expect(await errorIds(doc.parser)).toEqual([]);
    const rejected = await afterResolve(doc.parser, 'rejectAll');
    expect(rejected.errors).toEqual([]);
    expect(textOf(rejected.xml)).toBe('page  end');
    expect(rejected.xml).toMatch(/<w:instrText[^>]*> PAGE <\/w:instrText>/);
  });

  it('a field inside a hyperlink stays inside it', async () => {
    const link = `<w:hyperlink w:anchor="top">${r('see ')}<w:fldSimple w:instr=" REF top "><w:r><w:t>Top</w:t></w:r></w:fldSimple></w:hyperlink>`;
    const doc = await open(`<w:p>${link}${r(' now')}</w:p>${SECOND}`);
    expect((await run(doc, { action: 'trackedChange', operation: 'delete', target: doc.id, text: 'Top' })).success).toBe(true);
    const xml = xmlOf(doc.parser);
    const anchored = xml.match(/<w:hyperlink [\s\S]*?<\/w:hyperlink>/)![0];
    expect(anchored).toContain('w:fldCharType="begin"');
    expect(anchored).toContain('w:fldCharType="end"');
    expect(count(xml, '<w:fldSimple')).toBe(0);
    expect(await errorIds(doc.parser)).toEqual([]);
  });

  it('two simple fields in one paragraph are both converted when one edit covers both', async () => {
    const a = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
    const b = '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple>';
    const doc = await open(`<w:p>${r('p. ')}${a}${r(' of ')}${b}</w:p>${SECOND}`);
    expect(doc.text).toBe('p. 1 of 9');
    expect((await run(doc, { action: 'delete', target: doc.id, trackChanges: true })).success).toBe(true);
    const xml = xmlOf(doc.parser);
    expect(count(xml, '<w:fldSimple')).toBe(0);
    expect(count(xml, 'w:fldCharType="begin"')).toBe(2);
    const ids = revisionIds(xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect(await errorIds(doc.parser)).toEqual([]);
    const rejected = await afterResolve(doc.parser, 'rejectAll');
    expect(rejected.errors).toEqual([]);
    expect(textOf(rejected.xml)).toBe('p. 1 of 9');
  });
});

describe('what stays as it was', () => {
  it('a tracked edit outside the field leaves the simple field byte for byte as it was', async () => {
    const doc = await open(BODY);
    const result = await run(doc, { action: 'replace', target: doc.id, text: 'BEFORE 2026 after', trackChanges: true });
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(xmlOf(doc.parser)).toContain(FIELD);
    expect(await errorIds(doc.parser)).toEqual([]);
  });

  it('an untracked edit outside the field leaves it as it was', async () => {
    const doc = await open(BODY);
    expect((await run(doc, { action: 'replace', target: doc.id, text: 'BEFORE 2026 after' })).success).toBe(true);
    expect(xmlOf(doc.parser)).toContain(FIELD);
  });

  it('an untracked edit that removes the field removes it without writing a complex field', async () => {
    const doc = await open(BODY);
    const result = await run(doc, { action: 'replace', target: doc.id, text: 'before X after' });
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const xml = xmlOf(doc.parser);
    expect(xml).not.toMatch(/fldSimple|fldChar|instrText/);
    expect(textOf(xml)).toBe('before X after');
    expect(await errorIds(doc.parser)).toEqual([]);
  });

  it('an edit that cuts the field is still refused, with the paragraph and the id counter unchanged', async () => {
    const field = '<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ann Lee</w:t></w:r></w:fldSimple>';
    const doc = await open(`<w:p>${r('by ')}${field}${r(' today')}</w:p>${SECOND}`);
    const before = xmlOf(doc.parser);
    const ids = doc.ast.nextChangeId;
    const result = await run(doc, { action: 'replace', target: doc.id, text: 'by Ann Kim today', trackChanges: true });
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/cuts through field/);
    expect(JSON.stringify(result.errors)).toMatch(/change all of the text of the field in the same edit, or edit the text outside the field/);
    expect(JSON.stringify(result.errors)).not.toMatch(/replace the paragraph explicitly/);
    expect(xmlOf(doc.parser)).toBe(before);
    expect(doc.ast.nextChangeId).toBe(ids);
  });

  it('a result that holds a tracked change is refused with the remedy, and nothing changes', async () => {
    const field = `<w:fldSimple w:instr=" DATE "><w:ins w:id="5" ${STAMP}><w:r><w:t>2026</w:t></w:r></w:ins></w:fldSimple>`;
    const doc = await open(`<w:p>${r('before ')}${field}${r(' after')}</w:p>${SECOND}`);
    const before = xmlOf(doc.parser);
    const ids = doc.ast.nextChangeId;
    const result = await run(doc, { action: 'replace', target: doc.id, text: 'before X after', trackChanges: true });
    expect(result.success).toBe(false);
    const message = JSON.stringify(result.errors);
    expect(message).toMatch(/simple field/);
    expect(message).toMatch(/edit the text outside the field, or make the change in Word/);
    expect(xmlOf(doc.parser)).toBe(before);
    expect(doc.ast.nextChangeId).toBe(ids);
  });

  it('a result that holds a link is refused, and nothing changes', async () => {
    const field = `<w:fldSimple w:instr=" DATE "><w:hyperlink w:anchor="a">${r('2026')}</w:hyperlink></w:fldSimple>`;
    const doc = await open(`<w:p>${r('before ')}${field}${r(' after')}</w:p>${SECOND}`);
    const before = xmlOf(doc.parser);
    const result = await run(doc, { action: 'delete', target: doc.id, trackChanges: true });
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/simple field/);
    expect(xmlOf(doc.parser)).toBe(before);
  });

  it('a later refusal in the same edit leaves the simple field unconverted (the conversion comes after every check)', async () => {
    const sdt = `<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${r('INSIDE-SDT')}</w:sdtContent></w:sdt>`;
    const doc = await open(`<w:p>${r('before ')}${FIELD}${sdt}${r(' after')}</w:p>${SECOND}`);
    expect(doc.text).toBe('before 2026INSIDE-SDT after');
    const before = xmlOf(doc.parser);
    const ids = doc.ast.nextChangeId;
    const result = await run(doc, { action: 'replace', target: doc.id, text: 'before INSIDE after', trackChanges: true });
    expect(result.success).toBe(false);
    expect(xmlOf(doc.parser)).toBe(before);
    expect(xmlOf(doc.parser)).toContain(FIELD);
    expect(doc.ast.nextChangeId).toBe(ids);
  });

  it('an edit that changes nothing does not convert the field', async () => {
    const doc = await open(BODY);
    const before = xmlOf(doc.parser);
    await run(doc, { action: 'replace', target: doc.id, text: doc.text, trackChanges: true });
    expect(xmlOf(doc.parser)).toBe(before);
  });

  it('a batch whose second operation is refused rolls the first back, field included', async () => {
    const doc = await open(`${BODY.replace(SECOND, '')}${SECOND}`);
    const before = xmlOf(doc.parser);
    const result = await doc.engine.executeBatch({
      operations: [
        { action: 'delete', target: doc.id, trackChanges: true },
        { action: 'replace', target: 'no-such-node', text: 'x' },
      ],
    } as never);
    expect(result.success).toBe(false);
    expect(xmlOf(doc.parser)).toBe(before);
    expect(xmlOf(doc.parser)).toContain(FIELD);
  });
});

describe('a field of several words', () => {
  const date = '<w:fldSimple w:instr=" DATE "><w:r><w:t>November 9, 2026</w:t></w:r></w:fldSimple>';
  const body = `<w:p>${r('Printed on ')}${date}${r(' by the office.')}</w:p>${SECOND}`;

  it('a replace that changes every word of the field converts it', async () => {
    const doc = await open(body);
    const result = await run(doc, { action: 'replace', target: doc.id, text: 'Printed on soon by the office.', trackChanges: true });
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(xmlOf(doc.parser)).not.toContain('<w:fldSimple');
    expect(await errorIds(doc.parser)).toEqual([]);
    const rejected = await afterResolve(doc.parser, 'rejectAll');
    expect(textOf(rejected.xml)).toBe('Printed on November 9, 2026 by the office.');
  });

  it('a replace that keeps a word or a space of the field splits the change, is refused with a remedy that works, and following it succeeds', async () => {
    const doc = await open(body);
    const before = xmlOf(doc.parser);
    const refused = await run(doc, { action: 'replace', target: doc.id, text: 'Printed on January 1 by the office.', trackChanges: true });
    expect(refused.success).toBe(false);
    expect(JSON.stringify(refused.errors)).toMatch(/change all of the text of the field in the same edit/);
    expect(xmlOf(doc.parser)).toBe(before);
    // The remedy: put the whole new text where the whole field text was, as one change.
    const done = await run(doc, { action: 'replace', target: doc.id, text: 'Printed on January-1 by the office.', trackChanges: true });
    expect(done.success, JSON.stringify(done.errors)).toBe(true);
    expect(await errorIds(doc.parser)).toEqual([]);
  });

  it('a tracked delete of the whole field text converts it', async () => {
    const doc = await open(body);
    const result = await run(doc, { action: 'trackedChange', operation: 'delete', target: doc.id, text: 'November 9, 2026' });
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(xmlOf(doc.parser)).not.toContain('<w:fldSimple');
    expect(await errorIds(doc.parser)).toEqual([]);
  });
});

describe('the converted field round-trips through save and reopen', () => {
  it('the saved file opens, shows the tracked deletion, and a selective reject by id restores the field', async () => {
    const doc = await open(BODY);
    const result = await run(doc, { action: 'replace', target: doc.id, text: 'before X after', trackChanges: true });
    expect(result.success).toBe(true);
    const saved = new DocxParser(doc.parser.rezip(doc.ast));
    const ast = await saved.getAST();
    expect(ast.sections[0].nodes[0].metadata.text).toContain('before');
    const engine = new OperationEngine(ast, saved);
    const rejected = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'reject', ids: result.revisionIds }] } as never);
    expect(rejected.success, JSON.stringify(rejected.errors)).toBe(true);
    const xml = xmlOf(saved);
    expect(textOf(xml)).toBe('before 2026 after');
    expect(xml).toMatch(/<w:instrText[^>]*> DATE <\/w:instrText>/);
    expect(await errorIds(saved)).toEqual([]);
  });
});
