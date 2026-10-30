import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '@attentionform/astra-engine';
import {
  CLIENT_SESSION_ID_ENV,
  createMcpSessionStore,
  hasClientSessionIdHatch,
  isClientSessionIdHatchEnabledByEnv,
} from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNewDocument } from '../../src/tools/new-document';

function blankBuf(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

const UUID_SHAPE = /^sess_[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

const tmpFiles: string[] = [];

function writeTmpDocx(prefix: string): string {
  const file = path.join(os.tmpdir(), `astra-sess-${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}.docx`);
  fs.writeFileSync(file, blankBuf());
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
  delete process.env[CLIENT_SESSION_ID_ENV];
});

describe('server-minted session ids', () => {
  it('omitted id mints sess_<uuid>; mints are unique', async () => {
    const store = createMcpSessionStore();
    const a = await store.openBuffer(blankBuf(), 'a.docx');
    const b = await store.openBuffer(blankBuf(), 'b.docx');
    expect(a.sessionId).toMatch(UUID_SHAPE);
    expect(b.sessionId).toMatch(UUID_SHAPE);
    expect(a.sessionId).not.toBe(b.sessionId);
    expect(store.size()).toBe(2);
  });

  it('open_document handler omits -> uuid shape', async () => {
    const store = createMcpSessionStore();
    const res = await LEAN6_HANDLERS.open_document(store, { path: writeTmpDocx('omit') });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.sessionId).toMatch(UUID_SHAPE);
  });

  it('squatted id rejected E_VALIDATION with remedy; no session created', async () => {
    const store = createMcpSessionStore();
    const res = await LEAN6_HANDLERS.open_document(store, {
      path: writeTmpDocx('squat'),
      sessionId: 'sess_custom',
    });
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toMatch(/server-minted/);
    expect(res.remedy).toMatch(/Omit sessionId/);
    expect(store.size()).toBe(0);
    expect(store.has('sess_custom')).toBe(false);
  });

  it('store-level squat throws E_VALIDATION (openDocument + openBuffer)', async () => {
    const store = createMcpSessionStore();
    await expect(store.openBuffer(blankBuf(), 'x.docx', 'sess_x')).rejects.toMatchObject({
      code: 'E_VALIDATION',
    });
    await expect(store.openDocument(writeTmpDocx('y'), 'sess_y')).rejects.toMatchObject({
      code: 'E_VALIDATION',
    });
    expect(store.size()).toBe(0);
  });

  it('hatch: deterministic id honored via store and handler', async () => {
    const store = createMcpSessionStore({ allowClientSessionIds: true });
    const h = await store.openBuffer(blankBuf(), 'det.docx', 'sess_det');
    expect(h.sessionId).toBe('sess_det');
    const res = await LEAN6_HANDLERS.open_document(store, {
      path: writeTmpDocx('det2'),
      sessionId: 'sess_det2',
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.sessionId).toBe('sess_det2');
  });

  it('hatch collision: second open E_CONFLICT, live session intact', async () => {
    const store = createMcpSessionStore({ allowClientSessionIds: true });
    const first = await store.openBuffer(blankBuf(), 'live.docx', 'sess_live');
    const revBefore = first.revision;
    await expect(store.openBuffer(blankBuf(), 'clash.docx', 'sess_live')).rejects.toMatchObject({
      code: 'E_CONFLICT',
    });
    const res = await LEAN6_HANDLERS.open_document(store, {
      path: writeTmpDocx('clash2'),
      sessionId: 'sess_live',
    });
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.code).toBe('E_CONFLICT');
    const live = store.get('sess_live');
    expect(live).toBe(first);
    expect(live.filePath).toBe('live.docx');
    expect(live.revision).toBe(revBefore);
  });

  it('hatch collision while locked: E_CONFLICT busy, never replaced', async () => {
    const store = createMcpSessionStore({ allowClientSessionIds: true });
    const first = await store.openBuffer(blankBuf(), 'busy.docx', 'sess_busy');
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const inFlight = store.withSession('sess_busy', async () => {
      await gate;
      return 'done';
    });
    await expect(store.openBuffer(blankBuf(), 'busy2.docx', 'sess_busy')).rejects.toMatchObject({
      code: 'E_CONFLICT',
    });
    release();
    await expect(inFlight).resolves.toBe('done');
    expect(store.get('sess_busy')).toBe(first);
  });

  it('new_document squat rejected E_VALIDATION before writing', async () => {
    const store = createMcpSessionStore();
    const target = path.join(os.tmpdir(), `astra-sess-new-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(target);
    const res = await handleNewDocument(store, { outPath: target, sessionId: 'sess_custom' });
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.code).toBe('E_VALIDATION');
    expect(res.remedy).toMatch(/Omit sessionId/);
    expect(fs.existsSync(target)).toBe(false);
  });

  it('new_document hatch: deterministic id honored', async () => {
    const store = createMcpSessionStore({ allowClientSessionIds: true });
    const target = path.join(os.tmpdir(), `astra-sess-newh-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(target);
    const res = await handleNewDocument(store, { outPath: target, sessionId: 'sess_picked' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.sessionId).toBe('sess_picked');
  });

  it('env enables the hatch; unknown values stay off', () => {
    expect(isClientSessionIdHatchEnabledByEnv({} as NodeJS.ProcessEnv)).toBe(false);
    expect(isClientSessionIdHatchEnabledByEnv({ [CLIENT_SESSION_ID_ENV]: '1' })).toBe(true);
    expect(isClientSessionIdHatchEnabledByEnv({ [CLIENT_SESSION_ID_ENV]: 'TRUE' })).toBe(true);
    expect(isClientSessionIdHatchEnabledByEnv({ [CLIENT_SESSION_ID_ENV]: 'yes' })).toBe(false);
    process.env[CLIENT_SESSION_ID_ENV] = '1';
    expect(createMcpSessionStore().getAllowClientSessionIds()).toBe(true);
    delete process.env[CLIENT_SESSION_ID_ENV];
    expect(createMcpSessionStore().getAllowClientSessionIds()).toBe(false);
  });

  it('argv parsing honors the flag and stops at --', () => {
    expect(hasClientSessionIdHatch(['node', 'x', '--allow-client-session-ids'])).toBe(true);
    expect(hasClientSessionIdHatch(['node', 'x'])).toBe(false);
    expect(hasClientSessionIdHatch(['node', 'x', '--', '--allow-client-session-ids'])).toBe(false);
  });
});
