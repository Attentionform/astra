import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

/**
 * Paragraph content that the run model used to refuse although `main` edited it correctly:
 *  - range permission markers (w:permStart, w:permEnd), which are zero-width like bookmarks;
 *  - w:bdo, w:dir and w:customXml, which wrap ordinary runs like w:hyperlink does.
 * And ruby text (furigana), which `main` damaged: the model now counts it as one block that no edit may cut.
 */
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const RUBY = '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="12"/><w:hpsRaise w:val="20"/><w:hpsBaseText w:val="24"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:t>RT</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>BASE</w:t></w:r></w:rubyBase></w:ruby></w:r>';
const SECOND = '<w:p><w:r><w:t>second</w:t></w:r></w:p>';

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const node = ast.sections[0].nodes[0];
  return { parser, ast, engine, id: node.id, text: node.metadata.text as string };
}
const xmlOf = (parser: DocxParser) => new AdmZip(parser.rezip({ skipValidation: true })).readAsText('word/document.xml');
const plainText = (xml: string) => (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).map(p => (p.match(/<w:t(?:\s[^>]*)?>[^<]*/g) ?? []).map(t => t.replace(/<[^>]*>/g, '')).join(''));
const errorIds = async (parser: DocxParser) => ((await validateDocxBuffer(parser.rezip({ skipValidation: true }))).errors ?? []).map((e: { id: string }) => e.id);
async function afterResolve(parser: DocxParser, operation: 'acceptAll' | 'rejectAll') {
  const saved = new DocxParser(parser.rezip({ skipValidation: true }));
  const ast = await saved.getAST();
  const engine = new OperationEngine(ast, saved);
  const result = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation }] } as never);
  expect(result.success, JSON.stringify(result.errors)).toBe(true);
  return { xml: xmlOf(saved), errors: await errorIds(saved) };
}

const WRAPPERS: Array<[string, string, string, string]> = [
  // name, markup around the word "mid", element that must survive, word
  ['w:bdo', `<w:bdo w:val="rtl">${r('mid')}</w:bdo>`, '<w:bdo ', 'mid'],
  ['w:dir', `<w:dir w:val="rtl">${r('mid')}</w:dir>`, '<w:dir ', 'mid'],
  ['w:customXml', `<w:customXml w:element="tag">${r('mid')}</w:customXml>`, '<w:customXml ', 'mid'],
  ['w:permStart and w:permEnd', `<w:permStart w:id="1" w:edGrp="everyone"/>${r('mid')}<w:permEnd w:id="1"/>`, '<w:permStart ', 'mid'],
];

describe('paragraphs with wrappers and markers that hold ordinary runs can be edited', () => {
  for (const [label, markup, element, word] of WRAPPERS) {
    describe(label, () => {
      const body = `<w:p>${r('before ')}${markup}${r(' after')}</w:p>${SECOND}`;

      it('has a baseline that validates and a text that counts the wrapped run', async () => {
        const doc = await open(body);
        expect(doc.text).toBe(`before ${word} after`);
        expect(await errorIds(doc.parser)).toEqual([]);
      });

      for (const tracked of [false, true]) {
        it(`${tracked ? 'tracked' : 'untracked'} replace outside the wrapper keeps it`, async () => {
          const doc = await open(body);
          const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: `BEFORE ${word} after`, trackChanges: tracked }] } as never);
          expect(result.success, JSON.stringify(result.errors)).toBe(true);
          const xml = xmlOf(doc.parser);
          expect(xml).toContain(element);
          expect(xml).toContain('BEFORE');
          expect(await errorIds(doc.parser)).toEqual([]);
          if (tracked) {
            expect((await afterResolve(doc.parser, 'acceptAll')).xml).toContain(element);
            const rejected = await afterResolve(doc.parser, 'rejectAll');
            expect(plainText(rejected.xml)[0]).toBe(`before ${word} after`);
            expect(rejected.xml).toContain(element);
            expect(rejected.errors).toEqual([]);
          }
        });

        it(`${tracked ? 'tracked' : 'untracked'} replace of the wrapped text keeps the wrapper and validates`, async () => {
          const doc = await open(body);
          const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'before CHANGED after', trackChanges: tracked }] } as never);
          expect(result.success, JSON.stringify(result.errors)).toBe(true);
          const xml = xmlOf(doc.parser);
          expect(xml).toContain('CHANGED');
          expect(await errorIds(doc.parser)).toEqual([]);
          if (tracked) {
            const accepted = await afterResolve(doc.parser, 'acceptAll');
            expect(plainText(accepted.xml)[0]).toBe('before CHANGED after');
            const rejected = await afterResolve(doc.parser, 'rejectAll');
            expect(plainText(rejected.xml)[0]).toBe(`before ${word} after`);
            expect(rejected.errors).toEqual([]);
          } else expect(plainText(xml)[0]).toBe('before CHANGED after');
        });
      }

      it('tracked delete of the paragraph: Accept All leaves no text, Reject All restores it', async () => {
        const doc = await open(body);
        const result = await doc.engine.executeBatch({ operations: [{ action: 'delete', target: doc.id, trackChanges: true }] } as never);
        expect(result.success, JSON.stringify(result.errors)).toBe(true);
        expect(await errorIds(doc.parser)).toEqual([]);
        const rejected = await afterResolve(doc.parser, 'rejectAll');
        expect(plainText(rejected.xml)[0]).toBe(`before ${word} after`);
        expect(rejected.xml).toContain(element);
        expect(rejected.errors).toEqual([]);
        const accepted = await afterResolve(doc.parser, 'acceptAll');
        expect(plainText(accepted.xml)).toEqual(['second']);
      });

      it('untracked delete of the paragraph removes it', async () => {
        const doc = await open(body);
        const result = await doc.engine.executeBatch({ operations: [{ action: 'delete', target: doc.id }] } as never);
        expect(result.success, JSON.stringify(result.errors)).toBe(true);
        expect(plainText(xmlOf(doc.parser))).toEqual(['second']);
      });
    });
  }

  it('a run split by an edit stays inside its wrapper', async () => {
    const doc = await open(`<w:p>${r('before ')}<w:bdo w:val="rtl">${r('one two three')}</w:bdo>${r(' after')}</w:p>${SECOND}`);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'before one TWO three after', trackChanges: true }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const xml = xmlOf(doc.parser);
    const inside = xml.match(/<w:bdo\s[^>]*>([\s\S]*?)<\/w:bdo>/)?.[1] ?? '';
    expect(inside).toContain('TWO');
    expect(inside).toContain('one');
    expect(inside).toContain('three');
    expect(await errorIds(doc.parser)).toEqual([]);
  });

  it('a paragraph that mixes all of them is edited around them', async () => {
    const doc = await open(`<w:p>${r('a ')}<w:permStart w:id="1" w:edGrp="everyone"/><w:bdo w:val="rtl">${r('b')}</w:bdo><w:permEnd w:id="1"/>${r(' c ')}<w:customXml w:element="t"><w:dir w:val="ltr">${r('d')}</w:dir></w:customXml>${r(' e')}</w:p>${SECOND}`);
    expect(doc.text).toBe('a b c d e');
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'A b C d E', trackChanges: true }] } as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const xml = xmlOf(doc.parser);
    for (const element of ['<w:permStart ', '<w:permEnd ', '<w:bdo ', '<w:customXml ', '<w:dir ']) expect(xml).toContain(element);
    expect(await errorIds(doc.parser)).toEqual([]);
    expect(plainText((await afterResolve(doc.parser, 'rejectAll')).xml)[0]).toBe('a b c d e');
  });
});

describe('ruby text (furigana)', () => {
  const body = `<w:p>${r('before ')}${RUBY}${r(' after')}</w:p>${SECOND}`;

  it('is counted in the stored text and in the model, so the paragraph can be edited', async () => {
    const doc = await open(body);
    expect(doc.text).toBe('before RTBASE after');
  });

  for (const tracked of [false, true]) {
    it(`${tracked ? 'tracked' : 'untracked'} replace outside the ruby keeps the ruby and its text`, async () => {
      const doc = await open(body);
      const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'BEFORE RTBASE after', trackChanges: tracked }] } as never);
      expect(result.success, JSON.stringify(result.errors)).toBe(true);
      const xml = xmlOf(doc.parser);
      expect(xml).toContain('<w:ruby>');
      expect(xml).toContain('<w:rubyBase>');
      expect(xml.match(/<w:t(?:\s[^>]*)?>RT</g)?.length).toBe(1);
      expect(xml.match(/<w:t(?:\s[^>]*)?>BASE</g)?.length).toBe(1);
      expect(xml).toContain('BEFORE');
      expect(await errorIds(doc.parser)).toEqual([]);
    });
  }

  it('refuses an edit that cuts or removes the ruby, names ruby text, and changes nothing', async () => {
    for (const operation of [
      { action: 'replace', text: 'before X after' },
      { action: 'replace', text: 'before X after', trackChanges: true },
      { action: 'delete', trackChanges: true },
    ]) {
      const doc = await open(body);
      const before = xmlOf(doc.parser);
      const result = await doc.engine.executeBatch({ operations: [{ target: doc.id, ...operation }] } as never);
      expect(result.success).toBe(false);
      const message = JSON.stringify(result.errors);
      expect(message).toMatch(/ruby text/);
      expect(message).not.toMatch(/field marker/);
      expect(message).toMatch(/trackChanges false|outside the ruby/);
      expect(xmlOf(doc.parser)).toBe(before);
    }
  });

  it('untracked delete, style and an insert next to it still work', async () => {
    for (const operation of [
      { action: 'delete' },
      { action: 'style', bold: true },
      { action: 'insert', type: 'paragraph', anchor: '<id>', position: 'after', text: 'NEW' },
    ]) {
      const doc = await open(body);
      const op = JSON.parse(JSON.stringify(operation).replace('<id>', doc.id));
      const result = await doc.engine.executeBatch({ operations: [{ ...(op.action === 'insert' ? {} : { target: doc.id }), ...op }] } as never);
      expect(result.success, `${op.action}: ${JSON.stringify(result.errors)}`).toBe(true);
    }
  });

  it('Accept All and Reject All keep the ruby', async () => {
    const doc = await open(`<w:p>${r('before ')}${RUBY}<w:ins w:id="5" w:author="R" w:date="2026-01-01T00:00:00Z">${r(' NEW')}</w:ins></w:p>${SECOND}`);
    for (const operation of ['acceptAll', 'rejectAll'] as const) {
      const resolved = await afterResolve(doc.parser, operation);
      expect(resolved.xml).toContain('<w:ruby>');
      expect(resolved.errors).toEqual([]);
    }
  });

  it('an unknown run child is named in the message and is not called a field marker', async () => {
    const doc = await open(`<w:p>${r('before ')}<w:r><w:t>x</w:t><w:unknownThing/></w:r>${r(' after')}</w:p>${SECOND}`);
    const result = await doc.engine.executeBatch({ operations: [{ action: 'replace', target: doc.id, text: 'changed' }] } as never);
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/does not model \(w:unknownThing\)/);
    expect(JSON.stringify(result.errors)).not.toMatch(/field marker/);
  });
});

describe('a tracked-deleted or tracked-inserted paragraph whose runs sit inside a container is removed whole', () => {
  const A = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';
  const CONTAINERS: Array<[string, (inner: string) => string]> = [
    ['w:hyperlink', inner => `<w:hyperlink w:anchor="a">${inner}</w:hyperlink>`],
    ['w:smartTag', inner => `<w:smartTag w:uri="urn:x" w:element="e">${inner}</w:smartTag>`],
    ['w:bdo', inner => `<w:bdo w:val="rtl">${inner}</w:bdo>`],
    ['w:dir', inner => `<w:dir w:val="rtl">${inner}</w:dir>`],
    ['w:customXml', inner => `<w:customXml w:element="t">${inner}</w:customXml>`],
  ];
  for (const [label, wrap] of CONTAINERS) {
    it(`${label}: Accept All removes a deleted paragraph and Reject All keeps it`, async () => {
      const deleted = `<w:del w:id="1" ${A}><w:r><w:delText>gone</w:delText></w:r></w:del>`;
      const body = `<w:p><w:pPr><w:rPr><w:del w:id="2" ${A}/></w:rPr></w:pPr>${wrap(deleted)}</w:p>${SECOND}`;
      const doc = await open(body);
      const accepted = await afterResolve(doc.parser, 'acceptAll');
      expect(plainText(accepted.xml)).toEqual(['second']);
      if (label !== 'w:smartTag') expect(accepted.errors).toEqual([]);
      const rejected = await afterResolve(doc.parser, 'rejectAll');
      expect(plainText(rejected.xml)).toEqual(['gone', 'second']);
      // The validator rejects w:smartTag itself in every file (also without revisions), so it cannot judge this case.
      if (label !== 'w:smartTag') expect(rejected.errors).toEqual([]);
    });

    it(`${label}: Reject All removes an inserted paragraph and Accept All keeps it`, async () => {
      const inserted = `<w:ins w:id="1" ${A}>${r('new')}</w:ins>`;
      const body = `<w:p><w:pPr><w:rPr><w:ins w:id="2" ${A}/></w:rPr></w:pPr>${wrap(inserted)}</w:p>${SECOND}`;
      const doc = await open(body);
      const rejected = await afterResolve(doc.parser, 'rejectAll');
      expect(plainText(rejected.xml)).toEqual(['second']);
      const accepted = await afterResolve(doc.parser, 'acceptAll');
      expect(plainText(accepted.xml)).toEqual(['new', 'second']);
    });

    it(`${label}: a paragraph that keeps live text in the container is not removed`, async () => {
      const body = `<w:p><w:pPr><w:rPr><w:del w:id="2" ${A}/></w:rPr></w:pPr>${wrap(r('live'))}</w:p>${SECOND}`;
      const doc = await open(body);
      const accepted = await afterResolve(doc.parser, 'acceptAll');
      expect(plainText(accepted.xml).join('|')).toContain('live');
    });
  }
});
