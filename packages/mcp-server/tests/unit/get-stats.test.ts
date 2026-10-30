import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, computeDocStats, DocxParser, estimateTokenCount } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { handleInspectDocument } from '../../src/lib/mcp/handlers/open-inspect';
import { handleComposeDocument } from '../../src/tools/compose-document';
import { handleGetStats } from '../../src/lib/mcp/handlers/parity';

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/** Two lead paragraphs + a two-level nested table (shape of the fidelity
 *  nested-table-deep fixture): 9 + 10 + 2 + 2 + 3 = 26 words, 3 w:tbl. */
function nestedTableBuffer(): Buffer {
  const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    p('A first plain sentence stands before the deep nesting.') +
    p('A second plain sentence gives the mutation runner its pair.') +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level zero</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level one</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="7000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level two leaf</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
    `</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`;
  const zip = new AdmZip();
  zip.addFile('word/document.xml', Buffer.from(documentXml));
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
  );
  return zip.toBuffer();
}

async function openBlankSession(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const store = createMcpSessionStore();
  const buf = DocxParser.createNewDocument('blank').rezip();
  const handle = await store.openBuffer(buf, 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

const RICH_MARKDOWN = [
  '# Census Proof',
  '',
  'Opening paragraph with several words for the counter.',
  '',
  '## Sections',
  '',
  '- first bullet',
  '- second bullet',
  '',
  '| Name | Value |',
  '| --- | --- |',
  '| alpha | 1 |',
  '| beta | 2 |',
  '',
].join('\n');

describe('get_stats handler', () => {
  it('counts words paragraphs tables and images with the engine counters', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-stats-'));
    try {
      fs.writeFileSync(path.join(dir, 'tiny.png'), Buffer.from(TINY_PNG_BASE64, 'base64'));
      const { store, sessionId } = await openBlankSession();
      const composed = await handleComposeDocument(store, {
        sessionId,
        markdown: `${RICH_MARKDOWN}\n![Tiny](tiny.png)\n`,
        allowedDir: dir,
      });
      expect(composed.ok).toBe(true);
      const res = await handleGetStats(store, { sessionId });
      expect(res.ok).toBe(true);
      if (!res.ok) throw new Error(`stats failed: ${JSON.stringify(res)}`);
      // Single-source proof: every field equals the engine counter on the same AST.
      const expected = computeDocStats(store.get(sessionId).ast);
      expect(res.words).toBe(expected.words);
      expect(res.paragraphs).toBe(expected.paragraphs);
      expect(res.tables).toBe(expected.tables);
      expect(res.images).toBe(expected.images);
      expect(res.pages).toBe(expected.pages);
      expect(res.words).toBeGreaterThan(10);
      expect(res.paragraphs).toBeGreaterThan(3);
      expect(res.tables).toBe(1);
      expect(res.images).toBe(1);
      expect(res.revision).toBe(2);
      expect(res.dirty).toBe(true);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('costs a fraction of an inspect baseline on the same document', async () => {
    const { store, sessionId } = await openBlankSession();
    const lines: string[] = ['# Census Proof', ''];
    for (let i = 1; i <= 300; i++) lines.push(`Paragraph number ${i} with enough words to be real.`);
    const composed = await handleComposeDocument(store, { sessionId, markdown: lines.join('\n\n') + '\n' });
    expect(composed.ok).toBe(true);
    const stats = await handleGetStats(store, { sessionId });
    expect(stats.ok).toBe(true);
    if (!stats.ok) throw new Error('stats failed');
    const inspected = await handleInspectDocument(store, { sessionId });
    expect(inspected.ok).toBe(true);
    if (!inspected.ok) throw new Error('inspect failed');
    const statsTokens = estimateTokenCount(JSON.stringify(stats));
    const inspectTokens = estimateTokenCount(inspected.toonText);
    expect(inspectTokens).toBeGreaterThan(1000);
    expect(statsTokens).toBeLessThan(50);
    expect(statsTokens).toBeLessThan(inspectTokens / 10);
  });

  it('counts a nested table once per w:t and tallies nested w:tbl', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(nestedTableBuffer(), 'nested.docx');
    const res = await handleGetStats(store, { sessionId: handle.sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`stats failed: ${JSON.stringify(res)}`);
    // Oracle pins (raw document.xml census): no per-ancestor repetition,
    // nested tables included in the tally.
    expect(res.words).toBe(26);
    expect(res.paragraphs).toBe(2);
    expect(res.tables).toBe(3);
    // Single-source proof stays true on nested input.
    const expected = computeDocStats(store.get(handle.sessionId).ast);
    expect(res.words).toBe(expected.words);
    expect(res.tables).toBe(expected.tables);
  });

  it('never bumps the revision (pure read)', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleGetStats(store, { sessionId });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error('stats failed');
    expect(res.revision).toBe(1);
    expect(res.dirty).toBe(false);
    expect(store.get(sessionId).revision).toBe(1);
  });

  it('refuses an unknown session with code and remedy', async () => {
    const store = createMcpSessionStore();
    const res = await handleGetStats(store, { sessionId: 'sess_nope' });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('unknown session unexpectedly counted');
    // Cross-surface parity (session tools preserve the store's E_SESSION verbatim).
    expect(res.code).toBe('E_SESSION');
    expect(res.message).toBe('Unknown session "sess_nope" (expired or never opened)');
    expect(res.remedy).toBe('Call open_document first, then retry with the returned sessionId.');
  });

  it('rejects wire violations with code and remedy', async () => {
    const { store } = await openBlankSession();
    const res = await handleGetStats(store, {});
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('missing sessionId unexpectedly counted');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.remedy.length).toBeGreaterThan(0);
  });
});
