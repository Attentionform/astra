import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';

const tmpFiles: string[] = [];

/** A bad value for w:type in a style: a schema error that nothing in the engine repairs. */
const BAD_STYLE = (id: string) => `<w:style w:type="bogus" w:styleId="${id}"><w:name w:val="${id}"/></w:style>`;

function writeDocx(prefix: string, badStyles: number): string {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  if (badStyles > 0) {
    const styles = zip.getEntry('word/styles.xml')?.getData().toString('utf8') ?? '';
    const bad = Array.from({ length: badStyles }, (_, i) => BAD_STYLE(`Bad${i + 1}`)).join('');
    zip.updateFile('word/styles.xml', Buffer.from(styles.replace('</w:styles>', `${bad}</w:styles>`), 'utf8'));
  }
  const file = path.join(os.tmpdir(), `astra-baseline-${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}.docx`);
  fs.writeFileSync(file, zip.toBuffer());
  tmpFiles.push(file);
  return file;
}

afterEach(() => {
  while (tmpFiles.length > 0) {
    try {
      fs.unlinkSync(tmpFiles.pop() as string);
    } catch {
      /* best-effort */
    }
  }
});

async function open(file: string, extra: Record<string, unknown> = {}) {
  const store = createMcpSessionStore();
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file, ...extra });
  if (!opened.ok) throw new Error(`open failed: ${JSON.stringify(opened)}`);
  return { store, opened, sessionId: opened.sessionId };
}

const insertText = (sessionId: string, text: string) =>
  ({
    sessionId,
    operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text }],
  }) as never;

describe('open_document validation', () => {
  it('reports a valid file as valid', async () => {
    const { opened } = await open(writeDocx('valid', 0));
    expect(opened.validation).toEqual({ status: 'valid' });
  });

  it('reports the errors a file already has, and that edits adding none can be saved', async () => {
    const { opened } = await open(writeDocx('invalid', 1));
    const v = opened.validation as { status: string; preExistingErrors: number; baselineComparable: boolean; degraded: boolean; note: string };
    expect(v.status).toBe('invalid');
    expect(v.preExistingErrors).toBeGreaterThanOrEqual(1);
    expect(v.baselineComparable).toBe(true);
    expect(v.degraded).toBe(true);
    expect(v.note).toContain('Edits that add no new error can be saved');
  });

  it('skips the check with validationTimeoutMs 0', async () => {
    const { opened } = await open(writeDocx('skipped', 1), { validationTimeoutMs: 0 });
    expect(opened.validation).toEqual({ status: 'skipped' });
  });
});

describe('save_document against the errors the file already had', () => {
  it('saves a harmless edit and says how many errors were already there', async () => {
    const { store, opened, sessionId } = await open(writeDocx('tolerated', 1));
    const before = (opened.validation as { preExistingErrors: number }).preExistingErrors;
    expect((await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'hello'))).ok).toBe(true);
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    expect(saved.validation).toMatchObject({ ok: false, tolerated: true, preExistingErrors: before, introducedErrors: 0 });
    expect(saved.commitMessage).toContain(`saved with ${before} existing validator errors, 0 introduced`);
  });

  it('a second save compares with the file just written', async () => {
    const { store, sessionId } = await open(writeDocx('second', 1));
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'one'));
    expect((await LEAN6_HANDLERS.save_document(store, { sessionId })).ok).toBe(true);
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'two'));
    const again = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(again.ok).toBe(true);
    if (again.ok) expect(again.validation).toMatchObject({ tolerated: true, introducedErrors: 0 });
  });

  it('refuses an edit that adds an error and names the numbers', async () => {
    const { store, opened, sessionId } = await open(writeDocx('introduced', 1));
    const before = (opened.validation as { preExistingErrors: number }).preExistingErrors;
    await store.withSession(sessionId, async (handle) => {
      // A part the engine does not rebuild on save, so the error reaches the validator as written.
      const zip = handle.parser.getZipAdapter();
      const settings = zip.readText('word/settings.xml');
      zip.writeText('word/settings.xml', settings.replace('</w:settings>', '<w:zoom w:percent="bogus"/></w:settings>'));
    });
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(false);
    if (saved.ok) return;
    expect(JSON.stringify(saved)).toMatch(new RegExp(`this edit adds \\d+ validator error\\(s\\) \\(the file already had ${before}\\)`));
  });

  it('refuses a tolerated save when strict is set', async () => {
    const { store, sessionId } = await open(writeDocx('strict', 1));
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'hello'));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, strict: true });
    expect(saved.ok).toBe(false);
    if (!saved.ok) expect(JSON.stringify(saved)).toContain('save blocked: package invalid');
  });

  it('works out the baseline at the first save when the open check was skipped', async () => {
    const { store, sessionId } = await open(writeDocx('lazy', 1), { validationTimeoutMs: 0 });
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'hello'));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(true);
    if (saved.ok) expect(saved.validation).toMatchObject({ tolerated: true, introducedErrors: 0 });
  });

  it('refuses the save when the file on disk changed after the open check was skipped', async () => {
    const file = writeDocx('changed', 1);
    const { store, sessionId } = await open(file, { validationTimeoutMs: 0 });
    fs.appendFileSync(file, Buffer.alloc(1));
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'hello'));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(false);
  });

  it('a clean save reports ok', async () => {
    const { store, sessionId } = await open(writeDocx('clean', 0));
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'hello'));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(true);
    if (saved.ok) expect(saved.validation).toEqual({ ok: true, valid: true, errorCount: 0 });
  });
});

describe('save_document result shapes read like validate_document', () => {
  it('a tolerated save carries valid:false and the error count next to ok:false', async () => {
    const { store, opened, sessionId } = await open(writeDocx('shape-tolerated', 2));
    const before = (opened.validation as { preExistingErrors: number }).preExistingErrors;
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'hello'));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    expect(saved.validation).toMatchObject({ ok: false, valid: false, tolerated: true });
    expect((saved.validation as { errorCount: number }).errorCount).toBe(before);
  });

  it('a refused save lists up to 10 errors, not only the first, and points to validate_document', async () => {
    const { store, sessionId } = await open(writeDocx('shape-refused', 12));
    await LEAN6_HANDLERS.execute_batch(store, insertText(sessionId, 'hello'));
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, strict: true });
    expect(saved.ok).toBe(false);
    if (saved.ok) return;
    const errors = (saved as unknown as { errors?: Array<{ description: string }> }).errors;
    expect(errors).toBeDefined();
    expect(errors!.length).toBe(10);
    expect(errors![0].description.length).toBeGreaterThan(0);
    expect((saved as unknown as { errorCount: number }).errorCount).toBeGreaterThanOrEqual(12);
    expect(saved.remedy).toContain('validate_document');
  });

  it('an edit that adds errors is refused with the errors it added', async () => {
    const { store, sessionId } = await open(writeDocx('shape-introduced', 1));
    await store.withSession(sessionId, async (handle) => {
      const zip = handle.parser.getZipAdapter();
      const settings = zip.readText('word/settings.xml');
      zip.writeText('word/settings.xml', settings.replace('</w:settings>', '<w:zoom w:percent="bogus"/></w:settings>'));
    });
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(false);
    if (saved.ok) return;
    const errors = (saved as unknown as { errors?: unknown[] }).errors;
    expect(Array.isArray(errors) && errors.length > 0).toBe(true);
  });
});
