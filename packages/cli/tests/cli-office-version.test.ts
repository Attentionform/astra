/**
 * The Office schema version of the CLI: `--office-version` or ASTRA_OFFICE_VERSION, resolved once per
 * run and passed to the validator on every path. Each test spies on the backend and checks the
 * `officeVersion` argument it actually received.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { nativeBackend } from '@attentionform/astra-engine/ooxml/validator-backend';
import { createMcpSessionStore } from '../../mcp-server/src/lib/session-store';
import { LEAN6_HANDLERS } from '../../mcp-server/src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-office-'));
let n = 0;
const tmp = (ext = '.docx') => path.join(dir, `f${n++}${ext}`);

function blank(invalid = false): string {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  if (invalid) {
    const styles = zip.getEntry('word/styles.xml')?.getData().toString('utf8') ?? '';
    zip.updateFile('word/styles.xml', Buffer.from(styles.replace('</w:styles>', '<w:style w:type="bogus" w:styleId="Bad1"><w:name w:val="Bad1"/></w:style></w:styles>'), 'utf8'));
  }
  const f = tmp();
  fs.writeFileSync(f, zip.toBuffer());
  return f;
}

function ops(): string {
  const f = tmp('.json');
  fs.writeFileSync(f, JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'hello' }] }));
  return f;
}

let received: Array<string | undefined> = [];
const stdout = (): string => vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
const stderr = (): string => vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
const lastJson = () => JSON.parse(stdout().split('\n').filter((l) => l.startsWith('{')).pop() as string);

beforeEach(() => {
  received = [];
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
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
  delete process.env.ASTRA_OFFICE_VERSION;
});

const only = (v: string) => expect(new Set(received)).toEqual(new Set([v]));

describe('the target is chosen once per run', () => {
  it('defaults to Office2019', async () => {
    expect(await run(['--json', 'validate', blank()])).toBe(0);
    expect(lastJson()).toMatchObject({ ok: true, requestedOfficeVersion: 'Office2019', office: 'Office2019' });
    only('Office2019');
  });
  it('the flag selects the version, before or after the command, with = or a space', async () => {
    expect(await run(['--json', '--office-version', 'Office2013', 'validate', blank()])).toBe(0);
    expect(lastJson()).toMatchObject({ requestedOfficeVersion: 'Office2013', office: 'Office2013' });
    expect(await run(['--json', 'validate', blank(), '--office-version=Microsoft365'])).toBe(0);
    expect(lastJson()).toMatchObject({ requestedOfficeVersion: 'Microsoft365', office: 'Microsoft365' });
    expect(new Set(received)).toEqual(new Set(['Office2013', 'Microsoft365']));
  });
  it('the environment variable applies, and the flag beats it', async () => {
    process.env.ASTRA_OFFICE_VERSION = 'Office2016';
    expect(await run(['--json', 'validate', blank()])).toBe(0);
    expect(lastJson().requestedOfficeVersion).toBe('Office2016');
    expect(await run(['--json', '--office-version', 'Office2021', 'validate', blank()])).toBe(0);
    expect(lastJson().requestedOfficeVersion).toBe('Office2021');
    expect(new Set(received)).toEqual(new Set(['Office2016', 'Office2021']));
  });
  it('an invalid flag value is a usage error (exit 2) that lists the seven versions', async () => {
    expect(await run(['--office-version', 'Office1999', 'validate', blank()])).toBe(2);
    expect(stderr()).toContain('Office2007, Office2010, Office2013, Office2016, Office2019, Office2021, Microsoft365');
    expect(received).toEqual([]);
  });
  it('a flag without a value and an invalid environment variable are usage errors', async () => {
    expect(await run(['validate', blank(), '--office-version'])).toBe(2);
    process.env.ASTRA_OFFICE_VERSION = 'nope';
    expect(await run(['validate', blank()])).toBe(2);
    expect(stderr()).toContain('Invalid ASTRA_OFFICE_VERSION');
  });
  it('the usage text names the flag', async () => {
    await run([]);
    expect(stderr()).toContain('--office-version <version>');
  });
  it('a document valid only for Office 2010 and later fails under Office2007 and passes under Office2010', async () => {
    const f = tmp();
    fs.copyFileSync(path.join(__dirname, '../../ooxml-validator/tests/fixtures/office2010-only.docx'), f);
    expect(await run(['validate', f, '--office-version', 'Office2007'])).toBe(1);
    expect(await run(['validate', f, '--office-version', 'Office2010'])).toBe(0);
  });
});

describe('every validation path of the CLI receives the chosen version', () => {
  it('batch (clean file) and the baseline check of batch (file with errors)', async () => {
    expect(await run(['batch', blank(), ops(), '--office-version', 'Office2010'])).toBe(0);
    expect(await run(['batch', blank(true), ops(), '--office-version', 'Office2010'])).toBe(0);
    expect(stderr()).toContain('NOTE: saved with');
    expect(received.length).toBeGreaterThanOrEqual(3); // gate, then gate + baseline for the file with errors
    only('Office2010');
  });
  it('batch --dry-run', async () => {
    expect(await run(['batch', blank(), ops(), '--dry-run', '--office-version', 'Office2021'])).toBe(0);
    only('Office2021');
  });
  it('normalize', async () => {
    expect(await run(['normalize', blank(), '--office-version', 'Office2016'])).toBe(0);
    expect(received.length).toBeGreaterThan(0);
    only('Office2016');
  });
  it('create', async () => {
    expect(await run(['create', tmp(), '--template', 'blank', '--office-version', 'Office2013'])).toBe(0);
    expect(received.length).toBeGreaterThan(0);
    only('Office2013');
  });
  it('import', async () => {
    const md = tmp('.md');
    fs.writeFileSync(md, '# Title\n\nBody.');
    expect(await run(['import', md, '-o', tmp(), '--office-version', 'Office2013'])).toBe(0);
    expect(received.length).toBeGreaterThan(0);
    only('Office2013');
  });
  it('sanitize', async () => {
    expect(await run(['sanitize', blank(), tmp(), '--office-version', 'Microsoft365'])).toBe(0);
    expect(received.length).toBeGreaterThan(0);
    only('Microsoft365');
  });
  it('export to PDF checks its source with the chosen version', async () => {
    await run(['export', blank(), '--format', 'pdf', '-o', tmp('.pdf'), '--renderer', 'libreoffice', '--office-version', 'Office2007']);
    expect(received.length).toBeGreaterThan(0);
    only('Office2007');
  }, 120_000);
  it('revert and redo check the restored package with the chosen version', async () => {
    const file = blank();
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error('open failed');
    for (const text of ['one', 'two']) {
      await LEAN6_HANDLERS.execute_batch(store, {
        sessionId: opened.sessionId,
        operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text }],
      } as never);
      expect((await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId, message: text })).ok).toBe(true);
    }
    received = [];
    expect(await run(['revert', file, '--steps', '1', '--office-version', 'Office2013'])).toBe(0);
    expect(await run(['redo', file, '--office-version', 'Office2013'])).toBe(0);
    expect(received.length).toBeGreaterThanOrEqual(2);
    only('Office2013');
  });
});
