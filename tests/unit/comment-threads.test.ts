import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const NS = `xmlns:w="${W}" xmlns:w14="${W14}" xmlns:w15="${W15}"`;
const DATE = '2026-01-01T00:00:00Z';

const cp = (text: string, paraId?: string) => `<w:p${paraId ? ` w14:paraId="${paraId}"` : ''}><w:r><w:t>${text}</w:t></w:r></w:p>`;
const comment = (id: number, paragraphs: string, author = 'Ann') => `<w:comment w:id="${id}" w:author="${author}" w:date="${DATE}" w:initials="A">${paragraphs}</w:comment>`;
const markers = (id: number, text: string) => `<w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t>${text}</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>`;

interface Spec { body: string; comments: string; extended?: string; ids?: string }

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
  add('comments.xml', `<w:comments ${NS}>${spec.comments}</w:comments>`, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  if (spec.extended !== undefined) add('commentsExtended.xml', `<w15:commentsEx ${NS}>${spec.extended}</w15:commentsEx>`, 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml');
  if (spec.ids !== undefined) add('commentsIds.xml', `<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">${spec.ids}</w16cid:commentsIds>`, 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml');
  zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels));
  zip.updateFile('[Content_Types].xml', Buffer.from(types));
  return zip.toBuffer();
}

async function open(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const op = (engine: OperationEngine, operation: string, target: string, text?: string) =>
  engine.executeBatch({ operations: [{ action: 'comment', operation, target, ...(text ? { text } : {}) }] } as never);
const part = (parser: DocxParser, name: string) => parser.getZipAdapter().readTextOrNull(`word/${name}`);
const docXml = (parser: DocxParser) => new XMLSerializer().serializeToString(parser.getDocumentDom());
const valid = async (parser: DocxParser, ast: Parameters<DocxParser['rezip']>[0]) => (await validateDocxBuffer(parser.rezip(ast))).ok;

const MULTI = {
  body: markers(0, 'anchored text'),
  comments: comment(0, cp('first', '11111111') + cp('second', '22222222') + cp('third', '33333333')),
  extended: '<w15:commentEx w15:paraId="33333333" w15:done="0"/>',
};

describe('comment threads key on the last paragraph of the parent (FX-04)', () => {
  it('a reply to a three-paragraph comment points at its last paragraph', async () => {
    const { parser, engine } = await open(build(MULTI));
    const result = await op(engine, 'reply', '0', 'a reply');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const list = parser.listPackageComments();
    const reply = list.find(c => c.id !== 0)!;
    expect(reply.paraIdParent).toBe('33333333');
    expect(part(parser, 'commentsExtended.xml')).not.toContain('paraIdParent="11111111"');
  });

  it('resolve and reopen write the entry for the last paragraph', async () => {
    const { parser, engine } = await open(build(MULTI));
    expect((await op(engine, 'resolve', '0')).success).toBe(true);
    expect(parser.listPackageComments().find(c => c.id === 0)!.resolved).toBe(true);
    expect(part(parser, 'commentsExtended.xml')).toMatch(/paraId="33333333"[^>]*done="1"|done="1"[^>]*paraId="33333333"/);
    expect((await op(engine, 'unresolve', '0')).success).toBe(true);
    expect(parser.listPackageComments().find(c => c.id === 0)!.resolved).toBe(false);
  });

  it('deleting the parent removes the reply and every trace of both', async () => {
    const { parser, ast, engine } = await open(build(MULTI));
    await op(engine, 'reply', '0', 'a reply');
    expect(parser.listPackageComments()).toHaveLength(2);
    const result = await op(engine, 'delete', '0');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.listPackageComments()).toHaveLength(0);
    expect(part(parser, 'commentsExtended.xml')).not.toContain('commentEx');
    expect(docXml(parser)).not.toMatch(/comment(RangeStart|RangeEnd|Reference)/);
    expect(await valid(parser, ast)).toBe(true);
  });

  it('a two-paragraph reply made elsewhere is removed with its parent', async () => {
    const { parser, engine } = await open(build({
      body: markers(0, 'anchored') + markers(1, 'reply anchor'),
      comments: comment(0, cp('parent', 'AAAA0001')) + comment(1, cp('r1', 'BBBB0001') + cp('r2', 'BBBB0002'), 'Bob'),
      extended: '<w15:commentEx w15:paraId="AAAA0001" w15:done="0"/><w15:commentEx w15:paraId="BBBB0002" w15:paraIdParent="AAAA0001" w15:done="0"/>',
    }));
    const list = parser.listPackageComments();
    expect(list.find(c => c.id === 1)!.paraIdParent).toBe('AAAA0001');
    const result = await op(engine, 'delete', '0');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.listPackageComments()).toHaveLength(0);
  });
});

describe('comments from generators that write no extended parts (FX-11)', () => {
  const FLAT = { body: markers(0, 'anchored'), comments: comment(0, cp('flat comment')) };

  it('lists the comment', async () => {
    const { parser } = await open(build(FLAT));
    expect(parser.listPackageComments().map(c => c.text)).toEqual(['flat comment']);
  });

  it('delete works with no extended part and no paragraph id', async () => {
    const { parser, ast, engine } = await open(build(FLAT));
    const result = await op(engine, 'delete', '0');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.listPackageComments()).toHaveLength(0);
    expect(docXml(parser)).not.toMatch(/comment(RangeStart|RangeEnd|Reference)/);
    expect(await valid(parser, ast)).toBe(true);
  });

  it('resolve creates the paragraph id, the extended part, its relationship and content type', async () => {
    const { parser, ast, engine } = await open(build(FLAT));
    const result = await op(engine, 'resolve', '0');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.listPackageComments()[0].resolved).toBe(true);
    const zip = parser.getZipAdapter();
    expect(zip.hasFile('word/commentsExtended.xml')).toBe(true);
    expect(zip.readText('word/_rels/document.xml.rels')).toContain('commentsExtended');
    expect(zip.readText('[Content_Types].xml')).toContain('/word/commentsExtended.xml');
    expect(part(parser, 'comments.xml')).toMatch(/w14:paraId="[0-9A-F]{8}"/);
    expect(await valid(parser, ast)).toBe(true);
  });

  it('delete works for a reference-only comment (no range markers)', async () => {
    const { parser, engine } = await open(build({ body: '<w:p><w:r><w:t>x</w:t></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p>', comments: comment(0, cp('ref only')) }));
    const result = await op(engine, 'delete', '0');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(docXml(parser)).not.toContain('commentReference');
  });

  it('delete works when the range spans two paragraphs', async () => {
    const body = '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>';
    const { parser, ast, engine } = await open(build({ body, comments: comment(0, cp('spans')) }));
    const result = await op(engine, 'delete', '0');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(docXml(parser)).not.toMatch(/comment(RangeStart|RangeEnd|Reference)/);
    expect(await valid(parser, ast)).toBe(true);
  });

  it('delete refuses and changes nothing when the anchors are inconsistent', async () => {
    const body = markers(0, 'one') + '<w:p><w:r><w:commentReference w:id="0"/></w:r></w:p>';
    const { parser, engine } = await open(build({ body, comments: comment(0, cp('dup')) }));
    const before = docXml(parser);
    const result = await op(engine, 'delete', '0');
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/inconsistent/);
    expect(docXml(parser)).toBe(before);
    expect(parser.listPackageComments()).toHaveLength(1);
  });
});

describe('operations accept the comment:<n> ids that search returns (FX-12)', () => {
  it.each(['resolve', 'unresolve', 'delete'])('%s accepts comment:0 and plain 0', async operation => {
    for (const target of ['comment:0', '0']) {
      const { parser, engine } = await open(build(MULTI));
      const result = await op(engine, operation, target);
      expect(result.success, `${operation} ${target}: ${JSON.stringify(result.errors)}`).toBe(true);
      if (operation === 'delete') expect(parser.listPackageComments()).toHaveLength(0);
    }
  });

  it('reply accepts comment:0', async () => {
    const { parser, engine } = await open(build(MULTI));
    const result = await op(engine, 'reply', 'comment:0', 'via search id');
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.listPackageComments()).toHaveLength(2);
  });

  it('a malformed id is still refused with a clear message', async () => {
    const { engine } = await open(build(MULTI));
    const result = await op(engine, 'delete', 'comment:abc');
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(/comment id/i);
  });
});
