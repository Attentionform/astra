import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createMcpSessionStore } from '../src/lib/session-store';
import { LEAN6_HANDLERS } from '../src/lib/mcp-tool-adapter';
import { DocxParser } from '@attentionform/astra-engine';

const tmp: string[] = [];
afterEach(() => {
  while (tmp.length > 0) {
    try { fs.unlinkSync(tmp.pop() as string); } catch { /* best effort */ }
  }
});

async function openNew(template: string): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const file = path.join(os.tmpdir(), `astra-page-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}.docx`);
  fs.writeFileSync(file, DocxParser.createNewDocument(template).rezip());
  tmp.push(file);
  const store = createMcpSessionStore();
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  if (!opened.ok) throw new Error('open failed');
  return { store, sessionId: opened.sessionId };
}

describe('inspect_document reports the page as the file has it', () => {
  it('a Letter document is called Letter, with its size and margins in mm and in', async () => {
    const { store, sessionId } = await openNew('blank');
    const r = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.page).toMatchObject({ paper: 'Letter', orientation: 'portrait', columns: 1 });
    expect(r.page?.width).toEqual({ twips: 12240, mm: 215.9, in: 8.5 });
    expect(r.page?.margins.top).toEqual({ twips: 1440, mm: 25.4, in: 1 });
    expect(r.page?.summary).toMatch(/^Letter portrait, 215.9 x 279.4 mm \(8.5 x 11 in\); margins top 25.4 mm/);
    expect(r.pageBySection).toBeUndefined();
  });

  it('a custom size after pageSetup is called Custom, and a changed paper reads back by name', async () => {
    const { store, sessionId } = await openNew('blank');
    const set = await LEAN6_HANDLERS.execute_batch(store, { sessionId, operations: [{ action: 'pageSetup', paperSize: { width: '150mm', height: '200mm' }, margins: { left: '18mm', gutter: '5mm' } }] });
    expect(set.ok, JSON.stringify(set)).toBe(true);
    const custom = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    if (!custom.ok) throw new Error('inspect failed');
    expect(custom.page?.paper).toBe('Custom');
    expect(custom.page?.width.mm).toBe(150);
    expect(custom.page?.margins.left?.mm).toBe(18);
    expect(custom.page?.margins.gutter?.mm).toBe(5);
    const a5 = await LEAN6_HANDLERS.execute_batch(store, { sessionId, operations: [{ action: 'pageSetup', paperSize: 'A5' }] });
    expect(a5.ok).toBe(true);
    const named = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    if (!named.ok) throw new Error('inspect failed');
    expect(named.page?.paper).toBe('A5');
  });

  it('a document whose sections differ lists each section', async () => {
    const { store, sessionId } = await openNew('report');
    const set = await LEAN6_HANDLERS.execute_batch(store, { sessionId, operations: [{ action: 'pageSetup', sectionIndex: 2, paperSize: 'A5' }] });
    expect(set.ok, JSON.stringify(set)).toBe(true);
    const r = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    if (!r.ok) throw new Error('inspect failed');
    expect(r.pageBySection?.length).toBeGreaterThan(1);
    expect(r.pageBySection?.[1].summary).toMatch(/^A5 /);
  });
});
