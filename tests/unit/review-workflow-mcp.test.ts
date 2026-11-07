import { describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import AdmZip from 'adm-zip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleExecuteBatch, handleInspectDocument, handleOpenDocument, handleSearchContent } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';

describe('MCP review workflow', () => {
  it('pages revisions even when the document has fewer body nodes than the review offset', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const revisions = Array.from({length:55}, (_, i) => `<w:ins w:id="${i + 1}" w:author="A"><w:r><w:t>word${i}</w:t></w:r></w:ins>`).join('');
    const xml = zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:p>${revisions}</w:p>$1`);
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(zip.toBuffer(), 'many-revisions.docx');
    const first = await handleInspectDocument(store, {sessionId:handle.sessionId, include:['revisions']});
    expect(first).toMatchObject({ok:true, revisions:{total:55, nextOffset:50}});
    const second = await handleInspectDocument(store, {sessionId:handle.sessionId, include:['revisions'], offset:50});
    expect(second).toMatchObject({ok:true, revisions:{total:55}});
    if (!second.ok) throw new Error('review page failed');
    expect((second.revisions as {revisions:unknown[]}).revisions).toHaveLength(5);
  });
  it('exposes review operations over the built stdio server', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'astra-review-wire-'));
    const file = join(dir, 'document.docx');
    writeFileSync(file, createBlankDocxBuffer('blank'));
    const env = Object.fromEntries(Object.entries(process.env).filter((entry): entry is [string,string] => entry[1] !== undefined));
    const transport = new StdioClientTransport({command:process.execPath, args:[join(process.cwd(),'packages/mcp-server/dist/index.js'), '--allowed-dir', dir], env});
    const client = new Client({name:'review-test', version:'1.0'}, {capabilities:{}});
    try {
      await client.connect(transport);
      const catalog = await client.listTools();
      expect(JSON.stringify(catalog.tools).length).toBeLessThan(90_000);
      expect(JSON.stringify(catalog.tools.find(tool => tool.name === 'execute_batch')?.inputSchema)).toContain('$ref');
      const call = async (name:string, args:Record<string, unknown>) => {
        const result = await client.callTool({name, arguments:args});
        const content = result.content as Array<{type:string; text?:string}>;
        return JSON.parse(content.find(item => item.type === 'text')!.text!);
      };
      const open = await call('open_document', {path:file});
      expect(open.ok).toBe(true);
      expect(open.summary.summaryOnly).toBe(true);
      const inspected = await call('inspect_document', {sessionId:open.sessionId});
      const node = inspected.toonText.match(/[A-F0-9]{8}/)?.[0];
      expect(node).toBeTruthy();
      const edit = await call('execute_batch', {sessionId:open.sessionId,operations:[{action:'replace',target:node,text:'Wire edit',trackChanges:true}]});
      expect(edit.ok, JSON.stringify(edit)).toBe(true);
      expect(edit.revisionIds.length).toBeGreaterThan(0);
      const revisions = await call('inspect_document', {sessionId:open.sessionId, include:['revisions']});
      expect(revisions.revisions.total).toBeGreaterThan(0);
      const reject = await call('execute_batch', {sessionId:open.sessionId,operations:[{action:'trackedChange',operation:'reject',ids:edit.revisionIds}]});
      expect(reject.ok, JSON.stringify(reject)).toBe(true);
    } finally {
      await client.close();
      rmSync(dir, {recursive:true, force:true});
    }
  }, 30000);
  it('does not log engine stack frames unless debug logging is enabled', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'error.docx');
    const previous = process.env.ASTRA_DEBUG;
    delete process.env.ASTRA_DEBUG;
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      const out = await handleExecuteBatch(store, {sessionId:handle.sessionId, operations:[{action:'replace',target:'missing',text:'x'}]});
      expect(out.ok).toBe(false);
      expect(log.mock.calls.length).toBeGreaterThan(0);
      expect(log.mock.calls.map(args => args.join(' ')).join('\n')).not.toMatch(/\n\s+at /);
    } finally {
      log.mockRestore();
      if (previous !== undefined) process.env.ASTRA_DEBUG = previous;
    }
  });
  it('opens with a compact summary and retains the explicit full summary', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'astra-review-'));
    try {
      const file = join(dir, 'document.docx');
      writeFileSync(file, createBlankDocxBuffer('blank'));
      const store = createMcpSessionStore();
      const compact = await handleOpenDocument(store, {path:file});
      expect(compact).toMatchObject({ok:true, summary:{summaryOnly:true, hasTrackedChanges:false}});
      expect(JSON.stringify(compact).length).toBeLessThan(10000);
      const full = await handleOpenDocument(store, {path:file, summaryOnly:false});
      expect(full.ok).toBe(true);
      expect(JSON.stringify(full)).toContain('stylesXml');
    } finally { rmSync(dir, {recursive:true, force:true}); }
  });
  it('writes, inspects, searches and resolves a comment', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'review.docx');
    const target = handle.ast.sections[0].nodes[0].id;
    const created = await handleExecuteBatch(store, {sessionId:handle.sessionId, operations:[{action:'comment',operation:'insert',target,text:'Unique review note'}]});
    expect(created).toMatchObject({ok:true});
    const inspect = await handleInspectDocument(store, {sessionId:handle.sessionId, include:['comments']});
    expect(inspect).toMatchObject({ok:true, comments:{comments:[{id:0, text:'Unique review note', resolved:false}]}});
    const body = await handleSearchContent(store, {sessionId:handle.sessionId, query:'Unique review note'});
    expect(body).toMatchObject({ok:true, matches:[]});
    const comments = await handleSearchContent(store, {sessionId:handle.sessionId, query:'Unique review note', scope:'comments'});
    expect(comments).toMatchObject({ok:true, matches:[{id:'comment:0'}]});
    const resolved = await handleExecuteBatch(store, {sessionId:handle.sessionId, operations:[{action:'comment',operation:'resolve',target:'0'}]});
    expect(resolved).toMatchObject({ok:true});
    expect(await handleInspectDocument(store, {sessionId:handle.sessionId, include:['comments'], commentId:0}))
      .toMatchObject({ok:true, comments:{thread:[{id:0, resolved:true}]}});
  });

  it('lists revision IDs and resolves a selected edit through MCP', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'tracked.docx');
    const target = handle.ast.sections[0].nodes[0].id;
    const edit = await handleExecuteBatch(store, {sessionId:handle.sessionId, operations:[{action:'replace',target,text:'Tracked text',trackChanges:true}]});
    expect(edit).toMatchObject({ok:true});
    if (!edit.ok) throw new Error('edit failed');
    const ids = (edit as unknown as {revisionIds:number[]}).revisionIds;
    expect(ids.length).toBeGreaterThan(0);
    expect(await handleInspectDocument(store, {sessionId:handle.sessionId,include:['revisions']})).toMatchObject({ok:true, revisions:{total:ids.length}});
    const rejected = await handleExecuteBatch(store, {sessionId:handle.sessionId,operations:[{action:'trackedChange',operation:'reject',ids}]});
    expect(rejected).toMatchObject({ok:true, revisionsPending:0});
  });
});
