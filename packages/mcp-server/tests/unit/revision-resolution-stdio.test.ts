import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';

const server = path.resolve(__dirname, '../../dist/index.js');

describe('revision resolution over MCP stdio', () => {
  let client: Client;
  let dir: string;
  beforeAll(async () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-revision-stdio-'));
    fs.writeFileSync(path.join(dir, 'document.docx'), createBlankDocxBuffer('blank'));
    client = new Client({name:'revision-resolution-test',version:'1.0.0'}, {capabilities:{}});
    await client.connect(new StdioClientTransport({command:process.execPath,args:[server,'--allowed-dir',dir],env:{...process.env}}));
  });
  afterAll(async () => {
    await client?.close();
    if (dir) fs.rmSync(dir, {recursive:true,force:true});
  });
  it('returns restored search text after rejecting a tracked edit', async () => {
    async function call(name: string, args: Record<string, unknown>) {
      const result = await client.callTool({name, arguments:args});
      expect(result.isError).toBeFalsy();
      return JSON.parse((result.content as Array<{text:string}>)[0].text);
    }
    const opened = await call('open_document', {path:path.join(dir,'document.docx'),summaryOnly:true});
    const sessionId = opened.sessionId;
    const inspected = await call('inspect_document', {sessionId,fidelity:'standard'});
    const id = inspected.toonText.match(/\b([A-F0-9]{8})\b/)?.[1];
    expect(id).toBeTruthy();
    await call('execute_batch', {sessionId,operations:[{action:'replace',target:id,text:'Original words'}]});
    await call('execute_batch', {sessionId,operations:[{action:'replace',target:id,text:'Changed words',trackChanges:true}]});
    const resolved = await call('execute_batch', {sessionId,operations:[{action:'trackedChange',operation:'rejectAll'}]});
    expect(resolved.ok).toBe(true);
    const changed = await call('search_content', {sessionId,query:'Changed'});
    const original = await call('search_content', {sessionId,query:'Original'});
    expect(changed.matches).toHaveLength(0);
    expect(original.matches).toHaveLength(1);
    await call('close_document', {sessionId});
  });
});
