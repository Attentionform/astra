// Thin MCP client around the PUBLISHED @attentionform/astra-mcp-server (stdio).
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const here = path.dirname(fileURLToPath(import.meta.url));
export const SERVER_BIN = path.resolve(here, '../node_modules/.bin/astra-mcp-server');

export async function connect(allowedDir, extraArgs = []) {
  const transport = new StdioClientTransport({
    command: SERVER_BIN,
    args: ['--allowed-dir', allowedDir, ...extraArgs],
    stderr: 'pipe',
  });
  const client = new Client({ name: 'astra-benchmarks', version: '0.0.0' });
  await client.connect(transport);
  async function call(name, args = {}) {
    const t0 = process.hrtime.bigint();
    const res = await client.callTool({ name, arguments: args });
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    const inputRejected = !!res.isError && /Input validation error/.test(JSON.stringify(res.content || ''));
    if (inputRejected && ['close_document', 'open_document', 'save_document', 'inspect_document'].includes(name)) throw new Error(`harness bug: ${name} rejected by input validation: ${JSON.stringify(res.content).slice(0, 300)}`);
    const text = (res.content || []).filter((c) => c.type === 'text').map((c) => c.text).join('\n');
    let json;
    try { json = JSON.parse(text); } catch { json = undefined; }
    return { ms, isError: !!res.isError, inputRejected, text, json, raw: res };
  }
  return { client, call, pid: transport.pid, close: () => client.close() };
}
