import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { existsSync, mkdtempSync, rmSync, symlinkSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleSaveDocument } from '../../packages/mcp-server/src/lib/mcp/handlers/save-close';
import { handleExportDocument } from '../../packages/mcp-server/src/tools/export';

describe('default MCP write roots', () => {
  let dir: string;
  const previous = process.env.ASTRA_EXTRA_DOC_ROOTS;
  beforeEach(() => { delete process.env.ASTRA_EXTRA_DOC_ROOTS; dir = mkdtempSync(join(tmpdir(), 'astra-roots-')); });
  afterEach(() => {
    if (previous === undefined) delete process.env.ASTRA_EXTRA_DOC_ROOTS;
    else process.env.ASTRA_EXTRA_DOC_ROOTS = previous;
    rmSync(dir, {recursive:true, force:true});
  });

  it('refuses outside, traversal and symlinked save and export targets without writing', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'sandbox.docx');
    symlinkSync(dir, join(dir, 'link'), 'dir');
    const targets = [
      `/etc/astra-review-refusal-${process.pid}.docx`,
      `${dir}/../astra-review-traversal-${process.pid}.docx`,
      join(dir, 'link', 'linked.docx'),
      `${dir}/../../etc/astra-review-normalized-${process.pid}.docx`,
    ];
    for (const target of targets) {
      const save = await handleSaveDocument(store, {sessionId:handle.sessionId, outPath:target});
      expect(save, target).toMatchObject({ok:false, code:'E_SANDBOX'});
      const exported = await handleExportDocument(store, {sessionId:handle.sessionId, format:'html', outPath:target.replace(/\.docx$/, '.html')});
      expect(exported, `${target}: ${JSON.stringify(exported)}`).toMatchObject({ok:false, code:'E_SANDBOX'});
      expect(existsSync(resolve(target))).toBe(false);
      expect(existsSync(resolve(target.replace(/\.docx$/, '.html')))).toBe(false);
    }
    expect(existsSync(join(dir, 'linked.docx'))).toBe(false);
    expect(existsSync(join(dir, 'linked.html'))).toBe(false);
  });
});
