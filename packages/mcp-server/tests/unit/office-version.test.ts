/**
 * The Office schema version of the server: set once at startup, reaching the validator on every path.
 * Each test spies on the backend and checks the `officeVersion` argument it actually received, not
 * only the field a result reports.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { nativeBackend } from '@attentionform/astra-engine/ooxml/validator-backend';
import { createAstraMcpServer, main } from '../../src/index';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { configureOfficeVersion } from '../../src/lib/mcp/validation-target';

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-office-version-'));
let counter = 0;
const file = (ext = '.docx') => path.join(dir, `f${process.pid}-${counter++}${ext}`);

function blankDocx(): string {
  const f = file();
  fs.writeFileSync(f, DocxParser.createNewDocument('blank').rezip());
  return f;
}

/** A document with one schema error that nothing in the engine repairs. */
function invalidDocx(): string {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const styles = zip.getEntry('word/styles.xml')?.getData().toString('utf8') ?? '';
  zip.updateFile('word/styles.xml', Buffer.from(styles.replace('</w:styles>', '<w:style w:type="bogus" w:styleId="Bad1"><w:name w:val="Bad1"/></w:style></w:styles>'), 'utf8'));
  const f = file();
  fs.writeFileSync(f, zip.toBuffer());
  return f;
}

let received: Array<string | undefined> = [];

async function connect() {
  const server = createAstraMcpServer(createMcpSessionStore());
  const [a, b] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'office-version-test', version: '1' });
  await Promise.all([server.connect(a), client.connect(b)]);
  const call = async (name: string, args: Record<string, unknown>) => {
    const res = await client.callTool({ name, arguments: args });
    const text = (res.content as Array<{ text?: string }>)[0]?.text ?? '{}';
    return JSON.parse(text) as Record<string, any>;
  };
  return { call, close: () => client.close() };
}

beforeEach(() => {
  received = [];
  const realBuffer = nativeBackend.validateBuffer.bind(nativeBackend);
  const realFile = nativeBackend.validateFile.bind(nativeBackend);
  vi.spyOn(nativeBackend, 'validateBuffer').mockImplementation((buf, opts) => {
    received.push(opts?.officeVersion);
    return realBuffer(buf, opts);
  });
  vi.spyOn(nativeBackend, 'validateFile').mockImplementation((p, opts) => {
    received.push(opts?.officeVersion);
    return realFile(p, opts);
  });
});

afterEach(() => {
  vi.restoreAllMocks();
  configureOfficeVersion(undefined);
  delete process.env.ASTRA_OFFICE_VERSION;
});

describe('every validation the server runs uses the configured Office version', () => {
  it.each(['Microsoft365', 'Office2010'] as const)('%s reaches the validator on every path', async (version) => {
    configureOfficeVersion(version);
    const { call, close } = await connect();
    try {
      // open: the check made at open
      const opened = await call('open_document', { path: blankDocx() });
      expect(opened.ok).toBe(true);
      expect(opened.validation).toMatchObject({ status: 'valid', requestedOfficeVersion: version, office: version });
      const sessionId = opened.sessionId as string;

      // edit, then save
      expect((await call('execute_batch', { sessionId, operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'one' }] })).ok).toBe(true);
      const validated = await call('validate_document', { sessionId });
      expect(validated.validation ?? validated).toMatchObject({ requestedOfficeVersion: version, office: version });
      const saved = await call('save_document', { sessionId });
      expect(saved.ok).toBe(true);
      expect(saved.validation).toMatchObject({ ok: true, requestedOfficeVersion: version, office: version });

      // compose, normalize, revert and redo
      expect((await call('compose_document', { sessionId, markdown: '# Title\n\nBody.' })).ok).toBe(true);
      expect((await call('save_document', { sessionId })).ok).toBe(true);
      expect((await call('normalize_document', { sessionId })).ok).toBe(true);
      expect((await call('revert_document', { sessionId, discardUnsaved: true })).ok).toBeDefined();
      expect((await call('redo_document', { sessionId })).ok).toBeDefined();

      // a new document
      const created = await call('new_document', { outPath: file(), template: 'blank' });
      expect(created.ok).toBe(true);

    } finally {
      await close();
    }
    expect(received.length).toBeGreaterThan(5);
    // The argument the validator received, not a reported field: every call carried the configured version.
    expect(new Set(received)).toEqual(new Set([version]));
  });

  it('the PDF export source check uses the configured version', async () => {
    configureOfficeVersion('Office2021');
    const { call, close } = await connect();
    try {
      const opened = await call('open_document', { path: blankDocx(), validationTimeoutMs: 0 });
      // The source check runs before any renderer starts; whether a renderer exists here does not matter.
      await call('export_document', { sessionId: opened.sessionId, format: 'pdf', outPath: file('.pdf'), renderer: 'libreoffice' });
    } finally {
      await close();
    }
    expect(received.length).toBeGreaterThan(0);
    expect(new Set(received)).toEqual(new Set(['Office2021']));
  }, 120_000);

  it('a file that already has errors is judged against the same version at open and at save', async () => {
    configureOfficeVersion('Microsoft365');
    const { call, close } = await connect();
    try {
      const opened = await call('open_document', { path: invalidDocx() });
      expect(opened.validation).toMatchObject({ status: 'invalid', requestedOfficeVersion: 'Microsoft365', office: 'Microsoft365' });
      const sessionId = opened.sessionId as string;
      await call('execute_batch', { sessionId, operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'x' }] });
      const saved = await call('save_document', { sessionId });
      expect(saved.validation).toMatchObject({ tolerated: true, requestedOfficeVersion: 'Microsoft365', office: 'Microsoft365' });
    } finally {
      await close();
    }
    expect(new Set(received)).toEqual(new Set(['Microsoft365']));
  });

  it('the default is Office2019 and the environment variable applies when no flag is set', async () => {
    const first = await connect();
    expect((await first.call('open_document', { path: blankDocx() })).validation).toMatchObject({ requestedOfficeVersion: 'Office2019', office: 'Office2019' });
    await first.close();

    process.env.ASTRA_OFFICE_VERSION = 'Office2016';
    const second = await connect();
    expect((await second.call('open_document', { path: blankDocx() })).validation).toMatchObject({ requestedOfficeVersion: 'Office2016', office: 'Office2016' });
    await second.close();
    expect(new Set(received)).toEqual(new Set(['Office2019', 'Office2016']));
  });

  it('a skipped check reports the requested version and no version that ran', async () => {
    configureOfficeVersion('Office2013');
    const { call, close } = await connect();
    try {
      const opened = await call('open_document', { path: blankDocx(), validationTimeoutMs: 0 });
      expect(opened.validation).toEqual({ status: 'skipped', requestedOfficeVersion: 'Office2013' });
      expect(opened.validation).not.toHaveProperty('office');
    } finally {
      await close();
    }
  });

  it('a check that times out reports the requested version and no version that ran', async () => {
    configureOfficeVersion('Office2013');
    vi.mocked(nativeBackend.validateBuffer).mockImplementation(() => new Promise(() => undefined));
    const { call, close } = await connect();
    try {
      const opened = await call('open_document', { path: blankDocx(), validationTimeoutMs: 50 });
      expect(opened.validation).toMatchObject({ status: 'timed-out', requestedOfficeVersion: 'Office2013' });
      expect(opened.validation).not.toHaveProperty('office');
    } finally {
      await close();
    }
  });

  it('a validator that cannot start reports the requested version and no version that ran', async () => {
    configureOfficeVersion('Office2013');
    vi.mocked(nativeBackend.validateBuffer).mockRejectedValue(new Error('no validator binary'));
    const { call, close } = await connect();
    try {
      const opened = await call('open_document', { path: blankDocx() });
      expect(opened.validation).toMatchObject({ status: 'skipped', requestedOfficeVersion: 'Office2013' });
      expect(opened.validation.note).toContain('the validator is not available');
      expect(opened.validation).not.toHaveProperty('office');
    } finally {
      await close();
    }
  });
});

describe('--office-version flag', () => {
  const run = async (args: string[]) => {
    const err = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
    const code = await main(['node', 'astra-mcp-server', ...args]);
    const text = err.mock.calls.map((c) => String(c[0])).join('');
    err.mockRestore();
    return { code, text };
  };

  it('rejects an unknown version and lists the seven valid ones', async () => {
    const { code, text } = await run(['--office-version', 'Office1999']);
    expect(code).toBe(2);
    expect(text).toContain('Office2007, Office2010, Office2013, Office2016, Office2019, Office2021, Microsoft365');
  });

  it('rejects a missing value', async () => {
    expect((await run(['--office-version'])).code).toBe(2);
  });

  it('rejects an invalid ASTRA_OFFICE_VERSION at startup', async () => {
    process.env.ASTRA_OFFICE_VERSION = 'nope';
    const { code, text } = await run(['--dump-schemas', path.join(dir, 'schemas')]);
    expect(code).toBe(2);
    expect(text).toContain('Invalid ASTRA_OFFICE_VERSION');
  });

  it('the usage text names the flag', async () => {
    const { text } = await run(['--bogus']);
    expect(text).toContain('--office-version <version>');
  });

  it('the flag beats the environment variable', async () => {
    process.env.ASTRA_OFFICE_VERSION = 'Office2007';
    const { code } = await run(['--office-version=Microsoft365', '--dump-schemas', path.join(dir, 'schemas2')]);
    expect(code).toBe(0);
    const { serverOfficeVersion } = await import('../../src/lib/mcp/validation-target');
    expect(serverOfficeVersion()).toBe('Microsoft365');
  });
});
