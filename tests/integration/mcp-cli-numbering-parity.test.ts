import { afterEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip } from '@attentionform/astra-engine';
import { run as runCli } from '../../packages/cli/src/cli';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleValidateDocument } from '../../packages/mcp-server/src/lib/mcp/handlers/parity';

afterEach(() => vi.restoreAllMocks());

describe('CLI and MCP numbering validation parity', () => {
  it.each([false, true])('agrees on Word-style numbering (unrelated invalid markup: %s)', async (invalid) => {
    const zip = new AdmZip(fs.readFileSync(path.resolve(__dirname, '../fixtures/numbering-w15.docx')));
    if (invalid) {
      zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(
        '</w:body>', '<w:p><w:r><w:bogus/></w:r></w:p></w:body>',
      )));
    }
    const bytes = zip.toBuffer();
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-mcp-numbering-'));
    const store = createMcpSessionStore();
    let sessionId: string | undefined;
    try {
      const file = path.join(dir, 'input.docx');
      fs.writeFileSync(file, bytes);
      const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
      const code = await runCli(['validate', file, '--json']);
      const cli = JSON.parse(log.mock.calls.map(([text]) => String(text)).join(''));
      expect(code).toBe(invalid ? 1 : 0);
      sessionId = (await store.openBuffer(bytes, 'numbering-w15.docx')).sessionId;
      const result = await handleValidateDocument(store, { sessionId });
      expect(result.ok).toBe(true);
      if (!result.ok) throw new Error(JSON.stringify(result));
      expect(result.validation.valid).toBe(!invalid);
      expect(result.validation.errorCount).toBe(cli.errorCount ?? 0);
      expect(result.validation.errors.some(error => error.description.includes('restartNumberingAfterBreak'))).toBe(false);
    } finally {
      if (sessionId) await store.closeDocument(sessionId);
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});
