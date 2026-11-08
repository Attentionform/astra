import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DIR = join(__dirname, '..', 'fixtures', 'word-revisions');
const fixture = (name: string): Buffer => readFileSync(join(DIR, name));

/** What a reader sees: paragraph texts, bookmarks, comment markers, comment parts, revision markup left. */
async function summarize(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  await parser.getAST();
  const dom = parser.getDocumentDom();
  const names = (tag: string) => Array.from(dom.getElementsByTagNameNS(W, tag));
  const text = (p: Element) => Array.from(p.getElementsByTagNameNS(W, 't')).map(t => t.textContent ?? '').join('');
  const zip = new AdmZip(buffer);
  return {
    paragraphs: names('p').map(text),
    bookmarks: names('bookmarkStart').map(b => b.getAttributeNS(W, 'name')).sort(),
    comments: { starts: names('commentRangeStart').length, ends: names('commentRangeEnd').length, refs: names('commentReference').length },
    commentParts: ['comments', 'commentsExtended', 'commentsIds', 'commentsExtensible'].filter(p => zip.getEntry(`word/${p}.xml`)),
    revisions: parser.countPackageRevisions(),
    moveMarkers: ['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd'].reduce((n, t) => n + names(t).length, 0),
    customXml: names('customXmlInsRangeStart').length + names('customXmlInsRangeEnd').length,
  };
}

async function resolveWith(name: string, operation: 'acceptAll' | 'rejectAll') {
  const parser = new DocxParser(fixture(name));
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const result = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation }] } as never);
  const buffer = parser.rezip(ast);
  return { result, buffer, summary: await summarize(buffer) };
}

const CASES: [string, string][] = [
  ['move-last-paragraph', 'whole paragraph moved from the end to the top'],
  ['move-middle-paragraph', 'middle paragraph moved to the top'],
  ['move-heading-toc', 'heading moved, with a table of contents inserted while tracking'],
  ['delete-paragraph-comment-thread', 'deleted paragraph holding a comment with two replies'],
  ['delete-heading-bookmark-comment', 'deleted heading with a _Toc bookmark and a deleted paragraph with a comment'],
];

describe('resolving revisions in files made by desktop Word gives the result Word gives', () => {
  for (const [base, description] of CASES) {
    for (const [operation, suffix] of [['acceptAll', 'accepted'], ['rejectAll', 'rejected']] as const) {
      it(`${description}: ${operation}`, async () => {
        const expected = await summarize(fixture(`${base}.${suffix}.docx`));
        const { result, buffer, summary } = await resolveWith(`${base}.docx`, operation);
        expect(result.success, JSON.stringify(result.errors)).toBe(true);
        expect(result.revisionResolution?.unresolved ?? [], 'nothing left unresolved').toEqual([]);
        expect(result.revisionsPending).toBe(0);
        expect(summary.revisions).toBe(0);
        expect(summary.moveMarkers).toBe(0);
        expect(summary.customXml).toBe(0);
        expect(summary.paragraphs).toEqual(expected.paragraphs);
        expect(summary.bookmarks).toEqual(expected.bookmarks);
        expect(summary.comments).toEqual(expected.comments);
        expect(summary.commentParts).toEqual(expected.commentParts);
        expect((await validateDocxBuffer(buffer)).ok).toBe(true);
      });
    }
  }
});
