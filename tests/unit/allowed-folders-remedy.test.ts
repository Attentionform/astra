import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { describeAllowedFolders, validateSaveTarget } from '../../packages/mcp-server/src/lib/session-store';
import { NEW_DOCUMENT_DESCRIPTION } from '../../packages/mcp-server/src/tools/new-document';

let work = '';
const saved = process.env.ASTRA_EXTRA_DOC_ROOTS;

beforeEach(() => {
  work = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'allowed-remedy-')));
});
afterEach(() => {
  if (saved === undefined) delete process.env.ASTRA_EXTRA_DOC_ROOTS;
  else process.env.ASTRA_EXTRA_DOC_ROOTS = saved;
  fs.rmSync(work, { recursive: true, force: true });
});

describe('error remedies name the folders that are really allowed', () => {
  it('with a named folder, the text names that folder and not the system temp folder', () => {
    process.env.ASTRA_EXTRA_DOC_ROOTS = work;
    const text = describeAllowedFolders();
    expect(text).toContain(work);
    expect(text).not.toMatch(/system temp/);
  });

  it('with no named folder, the text names the defaults and how to allow another folder', () => {
    delete process.env.ASTRA_EXTRA_DOC_ROOTS;
    const text = describeAllowedFolders();
    expect(text).toMatch(/\.\/documents/);
    expect(text).toMatch(/system temp folder/);
    expect(text).toMatch(/--allowed-dir/);
  });

  it('a hidden-folder refusal gives the true remedy', () => {
    process.env.ASTRA_EXTRA_DOC_ROOTS = work;
    let remedy = '';
    try {
      validateSaveTarget(path.join(work, '.hidden', 'a.docx'));
    } catch (err: unknown) {
      remedy = (err as { remedy?: string }).remedy ?? '';
    }
    expect(remedy).toContain(work);
    expect(remedy).not.toMatch(/documents\/ or the system temp/);
  });

  it('the new_document description states the rule', () => {
    expect(NEW_DOCUMENT_DESCRIPTION).toMatch(/only when none are named/);
    expect(NEW_DOCUMENT_DESCRIPTION).not.toMatch(/--allowed-dir \/ documents \/ system temp/);
  });
});
