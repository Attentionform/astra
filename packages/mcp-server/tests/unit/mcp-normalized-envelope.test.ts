import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleExportDocument } from '../../src/tools/export';
import { handleVisualInspect } from '../../src/tools/visual';

const tmpFiles: string[] = [];

// Live binary probes run only when explicitly enabled; default runs stay
// mocked (no process spawns) and fast. Enable with ASTRA_RUN_SLOW=1.
const RUN_LIVE = process.env.ASTRA_RUN_SLOW === '1';

function writeTmpDocx(prefix: string): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = path.join(os.tmpdir(), `astra-normalized-${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}.docx`);
  fs.writeFileSync(file, buf);
  tmpFiles.push(file);
  return file;
}

/** Minimal legacy file: body paragraph with `w:shd` missing required `w:val`. */
function writeLegacyTmpDocx(prefix: string): string {
  const blank = DocxParser.createNewDocument('blank').rezip();
  const zip = new AdmZip(blank);
  const xml = zip.getEntry('word/document.xml')?.getData().toString('utf8');
  if (!xml) throw new Error('blank template has no document.xml');
  const sect = xml.indexOf('<w:sectPr');
  const anchor = sect >= 0 ? sect : xml.lastIndexOf('</w:body>');
  if (anchor < 0) throw new Error('blank template has no body anchor');
  const legacy =
    xml.slice(0, anchor) +
    '<w:p><w:pPr><w:shd w:fill="FF0000"/></w:pPr><w:r><w:t>legacy</w:t></w:r></w:p>' +
    xml.slice(anchor);
  zip.updateFile('word/document.xml', Buffer.from(legacy, 'utf8'));
  const file = path.join(os.tmpdir(), `astra-normalized-${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}.docx`);
  fs.writeFileSync(file, zip.toBuffer());
  tmpFiles.push(file);
  return file;
}

afterEach(() => {
  while (tmpFiles.length > 0) {
    const f = tmpFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
});

describe('save/batch envelopes carry normalized[]', () => {
  it('execute_batch success envelope carries normalized (array, empty on clean files)', async () => {
    const store = createMcpSessionStore();
    const src = writeTmpDocx('batch');
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: opened.sessionId,
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Envelope proof' }],
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(Array.isArray(res.normalized)).toBe(true);
  });

  it('save_document success envelope carries normalized (array, empty on clean files)', async () => {
    const store = createMcpSessionStore();
    const src = writeTmpDocx('save');
    const out = path.join(os.tmpdir(), `astra-normalized-save-out-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(out);
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId, outPath: out });
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    expect(Array.isArray(saved.normalized)).toBe(true);
    expect(fs.existsSync(out)).toBe(true);
    await LEAN6_HANDLERS.close_document(store, { sessionId: opened.sessionId });
  });

  it('save_document on a legacy file reports the core rewrites', async () => {    const store = createMcpSessionStore();
    const src = writeLegacyTmpDocx('save-legacy');
    const out = path.join(os.tmpdir(), `astra-normalized-save-legacy-out-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(out);
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId, outPath: out });
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    expect(saved.normalized.length).toBe(1);
    expect(saved.normalized[0]).toContain('w:shd');
    expect(fs.existsSync(out)).toBe(true);
    await LEAN6_HANDLERS.close_document(store, { sessionId: opened.sessionId });
  });

  it('execute_batch on a legacy file reports the bookkeeping heal, save then reports nothing new', async () => {
    const store = createMcpSessionStore();
    const src = writeLegacyTmpDocx('batch-legacy');
    const out = path.join(os.tmpdir(), `astra-normalized-batch-legacy-out-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(out);
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const batched = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: opened.sessionId,
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Batch legacy proof' }],
    });
    expect(batched.ok).toBe(true);
    if (!batched.ok) return;
    expect(batched.normalized.length).toBe(1);
    expect(batched.normalized[0]).toContain('w:shd');
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId, outPath: out });
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    // The heal already landed (and was reported) during the batch — the save
    // healed nothing new, so its envelope is honestly empty.
    expect(saved.normalized).toEqual([]);
    expect(fs.existsSync(out)).toBe(true);
    await LEAN6_HANDLERS.close_document(store, { sessionId: opened.sessionId });
  });
});

describe('render paths leave the live session untouched', () => {
  it.runIf(RUN_LIVE)('visual render plus pdf export heal nothing live; the later save reports the rewrite itself', async () => {
    const store = createMcpSessionStore();
    const src = writeLegacyTmpDocx('render-legacy');
    const pdfOut = path.join(os.tmpdir(), `astra-render-legacy-pdf-${process.pid}-${Date.now()}.pdf`);
    tmpFiles.push(pdfOut);
    const saveOut = path.join(os.tmpdir(), `astra-render-legacy-save-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(saveOut);
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const sessionId = opened.sessionId;
    const beforeXml = store.get(sessionId).parser.readDocumentXml();
    const beforeRogues = store.get(sessionId).parser.getDroppedRogueEntries();

    const seen = await handleVisualInspect(store, { sessionId, withVisual: true });
    if (!seen.ok) expect(seen.code).toBe('E_DEPENDENCY');
    expect(store.get(sessionId).parser.readDocumentXml()).toBe(beforeXml);
    expect(store.get(sessionId).parser.lastNormalizations).toEqual([]);
    expect(store.get(sessionId).parser.getDroppedRogueEntries()).toEqual(beforeRogues);
    expect(store.get(sessionId).revision).toBe(1);

    const exported = await handleExportDocument(store, { sessionId, format: 'pdf', outPath: pdfOut });
    expect(exported.ok).toBe(true);
    if (!exported.ok) return;
    expect(fs.existsSync(pdfOut)).toBe(true);
    expect(store.get(sessionId).parser.readDocumentXml()).toBe(beforeXml);
    expect(store.get(sessionId).parser.lastNormalizations).toEqual([]);
    expect(store.get(sessionId).parser.getDroppedRogueEntries()).toEqual(beforeRogues);

    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, outPath: saveOut });
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    // Neither render pass healed the live package, so the save is the first
    // writer and honestly reports the repair it applied.
    expect(saved.normalized.length).toBe(1);
    expect(saved.normalized[0]).toContain('w:shd');
    expect(fs.existsSync(saveOut)).toBe(true);
    await LEAN6_HANDLERS.close_document(store, { sessionId });
  }, 180000);

  it.runIf(RUN_LIVE)('clean document: render plus export then save stays quiet', async () => {
    const store = createMcpSessionStore();
    const src = writeTmpDocx('render-clean');
    const pdfOut = path.join(os.tmpdir(), `astra-render-clean-pdf-${process.pid}-${Date.now()}.pdf`);
    tmpFiles.push(pdfOut);
    const saveOut = path.join(os.tmpdir(), `astra-render-clean-save-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(saveOut);
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const sessionId = opened.sessionId;
    const beforeXml = store.get(sessionId).parser.readDocumentXml();

    const seen = await handleVisualInspect(store, { sessionId, withVisual: true });
    if (!seen.ok) expect(seen.code).toBe('E_DEPENDENCY');
    const exported = await handleExportDocument(store, { sessionId, format: 'pdf', outPath: pdfOut });
    expect(exported.ok).toBe(true);
    if (!exported.ok) return;
    expect(store.get(sessionId).parser.readDocumentXml()).toBe(beforeXml);
    expect(store.get(sessionId).parser.lastNormalizations).toEqual([]);

    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, outPath: saveOut });
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    // Nothing needed repair at any point, so the save envelope is empty.
    expect(saved.normalized).toEqual([]);
    expect(fs.existsSync(saveOut)).toBe(true);
    await LEAN6_HANDLERS.close_document(store, { sessionId });
  }, 180000);
});
