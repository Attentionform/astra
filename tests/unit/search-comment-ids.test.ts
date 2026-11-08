import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleExecuteBatch, handleSearchContent } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function docWith(paragraphs: number, withComment: boolean): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const body = Array.from({length: paragraphs}, (_, i) => `<w:p><w:r><w:t>needle ${i}</w:t></w:r></w:p>`).join('')
    + (withComment ? '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>anchored</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>' : '');
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  if (withComment) {
    zip.addFile('word/comments.xml', Buffer.from(`<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>needle comment</w:t></w:r></w:p></w:comment></w:comments>`));
    zip.updateFile('word/_rels/document.xml.rels', Buffer.from(zip.readAsText('word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')));
    zip.updateFile('[Content_Types].xml', Buffer.from(zip.readAsText('[Content_Types].xml').replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')));
  }
  return zip.toBuffer();
}

describe('search_content (FX-12)', () => {
  it('flags truncation when more than 50 nodes match, literal and regex', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(docWith(60, false), 'many.docx');
    for (const extra of [{}, {regex: true}]) {
      const result = await handleSearchContent(store, {sessionId: handle.sessionId, query: 'needle', ...extra});
      if (!result.ok) throw new Error(JSON.stringify(result));
      expect(result.matches).toHaveLength(50);
      expect((result as {truncated?: boolean}).truncated).toBe(true);
    }
  });

  it('does not flag truncation when everything fits', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(docWith(3, false), 'few.docx');
    const result = await handleSearchContent(store, {sessionId: handle.sessionId, query: 'needle'});
    if (!result.ok) throw new Error(JSON.stringify(result));
    expect(result.matches).toHaveLength(3);
    expect((result as {truncated?: boolean}).truncated).toBeFalsy();
  });

  it('a comment id from search can be resolved and deleted with no conversion', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(docWith(0, true), 'comment.docx');
    const found = await handleSearchContent(store, {sessionId: handle.sessionId, query: 'needle comment', scope: 'comments'});
    if (!found.ok) throw new Error(JSON.stringify(found));
    expect(found.matches).toEqual([{id: 'comment:0', preview: 'needle comment'}]);
    const id = found.matches[0].id;
    const resolved = await handleExecuteBatch(store, {sessionId: handle.sessionId, operations: [{action: 'comment', operation: 'resolve', target: id}]});
    expect(resolved.ok, JSON.stringify(resolved)).toBe(true);
    const deleted = await handleExecuteBatch(store, {sessionId: handle.sessionId, operations: [{action: 'comment', operation: 'delete', target: id}]});
    expect(deleted.ok, JSON.stringify(deleted)).toBe(true);
    const again = await handleSearchContent(store, {sessionId: handle.sessionId, query: 'needle comment', scope: 'comments'});
    if (!again.ok) throw new Error(JSON.stringify(again));
    expect(again.matches).toEqual([]);
  });
});
