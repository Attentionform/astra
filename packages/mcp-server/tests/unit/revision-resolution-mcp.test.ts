import { describe, expect, it } from 'vitest';
import { AdmZip, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { handleSanitizeDocument } from '../../src/lib/mcp/handlers/parity';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { handleExecuteBatch, handleSearchContent } from '../../src/lib/mcp/handlers/search-batch';

describe('revision resolution MCP workflow', () => {
  it('reports the restored text after rejecting a tracked replacement', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'example.docx');
    const id = handle.ast.sections[0].nodes[0].id;
    await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'replace', target: id, text: 'Original text' }] });
    await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'replace', target: id, text: 'Inserted text', trackChanges: true }] });
    const before = await handleSearchContent(store, { sessionId: handle.sessionId, query: 'Inserted' });
    expect(before.ok && before.matches).toHaveLength(1);
    const result = await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'trackedChange', operation: 'rejectAll' }] });
    expect(result.ok).toBe(true);
    const after = await handleSearchContent(store, { sessionId: handle.sessionId, query: 'Inserted' });
    const restored = await handleSearchContent(store, { sessionId: handle.sessionId, query: 'Original' });
    expect(after.ok && after.matches).toHaveLength(0);
    expect(restored.ok && restored.matches).toHaveLength(1);
    await store.closeDocument(handle.sessionId);
  });
  it('refuses to sanitize when an unpaired move cannot be resolved, and says why', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const xml = zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, '<w:body><w:p><w:moveFrom w:id="8" w:author="Alice"><w:r><w:t>Keep</w:t></w:r></w:moveFrom></w:p>$1');
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(zip.toBuffer(), 'move.docx');
    const result = await handleSanitizeDocument(store, {sessionId: handle.sessionId, acceptAllRevisions: true});
    expect(result.ok).toBe(false);
    expect(JSON.stringify(result)).toMatch(/revision markup remains/);
    expect(JSON.stringify(result)).toContain('moveFrom');
    await store.closeDocument(handle.sessionId);
  });

});
