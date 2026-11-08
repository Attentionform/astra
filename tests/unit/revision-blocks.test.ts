import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS = `xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"`;
const stamp = (author: string) => `w:author="${author}" w:date="2026-01-01T00:00:00Z"`;

interface Spec { body: string; comments?: string; extended?: string }
function build(spec: Spec): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${spec.body}$1`)));
  let rels = zip.readAsText('word/_rels/document.xml.rels');
  let types = zip.readAsText('[Content_Types].xml');
  const add = (part: string, xml: string, rel: string, ct: string) => {
    zip.addFile(`word/${part}`, Buffer.from(xml));
    rels = rels.replace('</Relationships>', `<Relationship Id="rId${part.replace(/\W/g, '')}" Type="${rel}" Target="${part}"/></Relationships>`);
    types = types.replace('</Types>', `<Override PartName="/word/${part}" ContentType="${ct}"/></Types>`);
  };
  if (spec.comments) add('comments.xml', `<w:comments ${NS}>${spec.comments}</w:comments>`, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  if (spec.extended) add('commentsExtended.xml', `<w15:commentsEx ${NS}>${spec.extended}</w15:commentsEx>`, 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml');
  zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels));
  zip.updateFile('[Content_Types].xml', Buffer.from(types));
  return zip.toBuffer();
}

async function open(spec: Spec) {
  const parser = new DocxParser(build(spec));
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const run = (engine: OperationEngine, operations: unknown[]) => engine.executeBatch({ operations } as never);
const texts = (parser: DocxParser) => Array.from(parser.getDocumentDom().getElementsByTagNameNS(W, 'p'))
  .map(p => Array.from(p.getElementsByTagNameNS(W, 't')).map(t => t.textContent ?? '').join(''));
const count = (parser: DocxParser, tag: string) => parser.getDocumentDom().getElementsByTagNameNS(W, tag).length;
const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

// Whole-paragraph moves as desktop Word writes them (see tests/fixtures/word-revisions): the paragraph mark
// carries the move, the range start sits in the paragraph, the range end sits between paragraphs.
const movedTo = (text: string, id: number, name: string, author = 'B') =>
  `<w:p><w:pPr><w:rPr><w:moveTo w:id="${id}" ${stamp(author)}/></w:rPr></w:pPr><w:moveToRangeStart w:id="${id + 1}" w:name="${name}" ${stamp(author)}/><w:moveTo w:id="${id + 2}" ${stamp(author)}><w:r><w:t>${text}</w:t></w:r></w:moveTo></w:p><w:moveToRangeEnd w:id="${id + 1}"/>`;
const movedFrom = (text: string, id: number, name: string, author = 'A') =>
  `<w:p><w:pPr><w:rPr><w:moveFrom w:id="${id}" ${stamp(author)}/></w:rPr></w:pPr><w:moveFromRangeStart w:id="${id + 1}" w:name="${name}" ${stamp(author)}/><w:moveFrom w:id="${id + 2}" ${stamp(author)}><w:r><w:t>${text}</w:t></w:r></w:moveFrom></w:p><w:moveFromRangeEnd w:id="${id + 1}"/>`;
const MOVE = movedTo('MOVED', 0, 'm1') + para('middle') + movedFrom('MOVED', 3, 'm1');

describe('whole-paragraph moves (FX-05)', () => {
  it('accept keeps the paragraph once, at the destination, with no marks', async () => {
    const { parser, engine } = await open({ body: MOVE });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(texts(parser)).toEqual(['MOVED', 'middle']);
    expect(result.revisionsPending).toBe(0);
    expect(result.revisionResolution?.unresolved).toEqual([]);
    expect(count(parser, 'moveFromRangeStart') + count(parser, 'moveToRangeEnd') + count(parser, 'moveFrom') + count(parser, 'moveTo')).toBe(0);
  });

  it('reject keeps the paragraph once, at the source, with no marks', async () => {
    const { parser, ast, engine } = await open({ body: MOVE });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'rejectAll' }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(texts(parser)).toEqual(['middle', 'MOVED']);
    expect(result.revisionsPending).toBe(0);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('selecting only the source side by id resolves both sides', async () => {
    const { parser, engine } = await open({ body: MOVE });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'accept', ids: [5] }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(texts(parser)).toEqual(['MOVED', 'middle']);
    expect(result.revisionsPending).toBe(0);
  });

  it('selecting only the destination mark by id resolves both sides', async () => {
    const { parser, engine } = await open({ body: MOVE });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'reject', ids: [0] }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(texts(parser)).toEqual(['middle', 'MOVED']);
  });

  it('an author selector that names only one side still moves both sides together', async () => {
    const { parser, engine } = await open({ body: MOVE });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'accept', author: 'A' }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(texts(parser)).toEqual(['MOVED', 'middle']);
    expect(result.revisionsPending).toBe(0);
  });

  it('a revision nested inside the move leaves both sides untouched and reports them', async () => {
    const nested = `<w:p><w:pPr><w:rPr><w:moveTo w:id="0" ${stamp('B')}/></w:rPr></w:pPr><w:moveToRangeStart w:id="1" w:name="m1" ${stamp('B')}/><w:moveTo w:id="2" ${stamp('B')}><w:ins w:id="9" ${stamp('C')}><w:r><w:t>MOVED</w:t></w:r></w:ins></w:moveTo></w:p><w:moveToRangeEnd w:id="1"/>`;
    const { parser, engine } = await open({ body: nested + para('middle') + movedFrom('MOVED', 3, 'm1') });
    const before = texts(parser);
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(result.revisionResolution!.unresolved.length).toBeGreaterThan(0);
    expect(texts(parser)).toEqual(before);
    expect(count(parser, 'moveFrom')).toBe(2);
    expect(count(parser, 'moveTo')).toBe(2);
  });

  it('a paragraph that is only partly moved is left untouched and reported', async () => {
    const partial = `<w:p><w:pPr><w:rPr><w:moveFrom w:id="3" ${stamp('A')}/></w:rPr></w:pPr><w:moveFromRangeStart w:id="4" w:name="m1" ${stamp('A')}/><w:moveFrom w:id="5" ${stamp('A')}><w:r><w:t>MOVED</w:t></w:r></w:moveFrom><w:r><w:t>loose text</w:t></w:r></w:p><w:moveFromRangeEnd w:id="4"/>`;
    const { parser, engine } = await open({ body: movedTo('MOVED', 0, 'm1') + para('middle') + partial });
    const before = texts(parser);
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(JSON.stringify(result.revisionResolution!.unresolved)).toMatch(/partial move/);
    expect(texts(parser)).toEqual(before);
  });
});

describe('deleted paragraphs take their bookmarks and comments with them, as in Word (FX-08, decision D-1)', () => {
  const del = (text: string, id: number) => `<w:del w:id="${id}" ${stamp('A')}><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
  const deletedPara = (inner: string, id: number) => `<w:p><w:pPr><w:rPr><w:del w:id="${id}" ${stamp('A')}/></w:rPr></w:pPr>${inner}</w:p>`;

  it('removes a bookmark that lies wholly inside, and reports it and the links that now dangle', async () => {
    const body = para('before')
      + deletedPara(`<w:bookmarkStart w:id="1" w:name="_Toc1"/>${del('Gone heading', 2)}<w:bookmarkEnd w:id="1"/>`, 3)
      + `<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:t>see heading</w:t></w:r></w:hyperlink></w:p>` + para('after');
    const { parser, engine } = await open({ body });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(texts(parser)).toEqual(['before', 'see heading', 'after']);
    expect(count(parser, 'bookmarkStart')).toBe(0);
    expect(result.revisionResolution?.removedBookmarks).toEqual(['_Toc1']);
    expect(result.revisionResolution?.staleLinks).toEqual([{ bookmark: '_Toc1', count: 1 }]);
    expect(result.revisionsPending).toBe(0);
  });

  it('keeps the paragraph and reports it when a bookmark ends outside', async () => {
    const body = para('before') + deletedPara(`<w:bookmarkStart w:id="1" w:name="span"/>${del('Gone', 2)}`, 3) + para('after') + `<w:p><w:bookmarkEnd w:id="1"/><w:r><w:t>end here</w:t></w:r></w:p>`;
    const { parser, engine } = await open({ body });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(JSON.stringify(result.revisionResolution!.unresolved)).toMatch(/bookmark starts or ends outside/);
    expect(count(parser, 'bookmarkStart')).toBe(1);
    expect(count(parser, 'delText')).toBe(1);
    expect(result.revisionsPending).toBeGreaterThan(0);
  });

  const COMMENTED = (id: number, extra = '') => `<w:commentRangeStart w:id="${id}"/>${del('Gone', 10 + id)}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>${extra}`;
  const comment = (id: number, paraId: string, text: string) => `<w:comment w:id="${id}" w:author="Ann" w:date="2026-01-01T00:00:00Z" w:initials="A"><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;

  it('deletes a comment thread whose anchors all lie inside, and keeps other comments and the comment parts', async () => {
    const body = para('before') + deletedPara(COMMENTED(0), 3) + `<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>kept anchor</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`;
    const { parser, ast, engine } = await open({
      body,
      comments: comment(0, '10000001', 'goes') + comment(1, '10000002', 'stays'),
      extended: '<w15:commentEx w15:paraId="10000001" w15:done="0"/><w15:commentEx w15:paraId="10000002" w15:done="0"/>',
    });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.listPackageComments().map(c => c.text)).toEqual(['stays']);
    expect(result.revisionResolution?.removedComments).toEqual([0]);
    const zip = parser.getZipAdapter();
    expect(zip.hasFile('word/comments.xml')).toBe(true);
    expect(zip.readText('word/commentsExtended.xml')).not.toContain('10000001');
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('removes the comment parts, relationships and content types when the last comment goes', async () => {
    const { parser, ast, engine } = await open({ body: para('before') + deletedPara(COMMENTED(0), 3) + para('after'), comments: comment(0, '10000001', 'only'), extended: '<w15:commentEx w15:paraId="10000001" w15:done="0"/>' });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const zip = parser.getZipAdapter();
    expect(zip.hasFile('word/comments.xml')).toBe(false);
    expect(zip.hasFile('word/commentsExtended.xml')).toBe(false);
    const saved = new AdmZip(parser.rezip(ast));
    expect(saved.readAsText('word/_rels/document.xml.rels')).not.toContain('comments');
    expect(saved.readAsText('[Content_Types].xml')).not.toContain('comments');
    expect((await validateDocxBuffer(saved.toBuffer())).ok).toBe(true);
  });

  it('keeps the paragraph when a reply in the thread is anchored outside', async () => {
    const body = para('before') + deletedPara(COMMENTED(0), 3)
      + `<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>reply anchor</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`;
    const { parser, engine } = await open({
      body,
      comments: comment(0, '10000001', 'parent') + comment(1, '10000002', 'reply'),
      extended: '<w15:commentEx w15:paraId="10000001" w15:done="0"/><w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001" w15:done="0"/>',
    });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(JSON.stringify(result.revisionResolution!.unresolved)).toMatch(/comment thread/);
    expect(parser.listPackageComments()).toHaveLength(2);
    expect(count(parser, 'commentRangeStart')).toBe(2);
  });

  it('keeps the paragraph when a comment reference has no range', async () => {
    const body = para('before') + deletedPara(`${del('Gone', 4)}<w:r><w:commentReference w:id="0"/></w:r>`, 3) + para('after');
    const { parser, engine } = await open({ body, comments: comment(0, '10000001', 'ref only') });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(JSON.stringify(result.revisionResolution!.unresolved)).toMatch(/comment starts, ends or is referenced outside/);
    expect(parser.listPackageComments()).toHaveLength(1);
  });

  it('a deleted paragraph mark with live text still joins the next paragraph', async () => {
    const body = `<w:p><w:pPr><w:rPr><w:del w:id="1" ${stamp('A')}/></w:rPr></w:pPr><w:r><w:t>first </w:t></w:r></w:p>` + para('second');
    const { parser, engine } = await open({ body });
    const result = await run(engine, [{ action: 'trackedChange', operation: 'acceptAll' }]);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(texts(parser)).toEqual(['first second']);
  });

  it('reject of a tracked-inserted paragraph removes it, accept keeps it', async () => {
    const inserted = `<w:p><w:pPr><w:rPr><w:ins w:id="1" ${stamp('A')}/></w:rPr></w:pPr><w:ins w:id="2" ${stamp('A')}><w:r><w:t>new paragraph</w:t></w:r></w:ins></w:p>`;
    for (const [operation, expected] of [['rejectAll', ['before', 'after']], ['acceptAll', ['before', 'new paragraph', 'after']]] as const) {
      const { parser, engine } = await open({ body: para('before') + inserted + para('after') });
      const result = await run(engine, [{ action: 'trackedChange', operation }]);
      expect(result.success, JSON.stringify(result.errors)).toBe(true);
      expect(texts(parser)).toEqual(expected);
      expect(result.revisionsPending).toBe(0);
    }
  });
});
