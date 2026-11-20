/**
 * Agents self-correct from the Remedy text, so the MCP server's errors must keep naming MCP tools,
 * even though the engine now also offers library wording (`libraryRemedy`) next to them.
 */
import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';

describe('MCP error remedies keep naming MCP tools', () => {
  it('a batch that targets a missing node says inspect_document', async () => {
    const file = path.join(os.tmpdir(), `astra-remedy-${process.pid}-${Date.now()}.docx`);
    fs.writeFileSync(file, DocxParser.createNewDocument('blank').rezip());
    try {
      const store = createMcpSessionStore();
      const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
      if (!opened.ok) throw new Error('open failed');
      const res = (await LEAN6_HANDLERS.execute_batch(store, {
        sessionId: opened.sessionId,
        operations: [{ action: 'replace', target: 'p_missing', text: 'x' }],
      } as never)) as { ok: boolean; remedy?: string; message?: string };
      expect(res.ok).toBe(false);
      expect(`${res.remedy ?? ''} ${res.message ?? ''}`).toContain('inspect_document');
    } finally {
      fs.rmSync(file, { force: true });
    }
  });

  it("a call without a session says open_document", async () => {
    const store = createMcpSessionStore();
    const res = (await LEAN6_HANDLERS.inspect_document(store, { sessionId: 'sess_nope' } as never)) as { ok: boolean; remedy?: string };
    expect(res.ok).toBe(false);
    expect(res.remedy ?? '').toContain('open_document');
  });
});
