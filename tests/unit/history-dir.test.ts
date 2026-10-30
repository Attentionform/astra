/**
 * ASTRA_HISTORY_DIR keeps the history sidecar and its tombstone registry out of
 * the document's folder. Unset, history stays in `<docDir>/.astra/` as before.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  commitChange,
  getHistory,
  getRepoDir,
  getSidecarHistoryWarning,
  initAstraRepo,
  undo,
} from '@attentionform/astra-engine';

let root: string;
let docsDir: string;
let historyDir: string;
const saved = process.env['ASTRA_HISTORY_DIR'];

beforeEach(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-histdir-'));
  docsDir = path.join(root, 'docs');
  historyDir = path.join(root, 'history');
  fs.mkdirSync(docsDir, { recursive: true });
});

afterEach(() => {
  if (saved === undefined) delete process.env['ASTRA_HISTORY_DIR'];
  else process.env['ASTRA_HISTORY_DIR'] = saved;
  fs.rmSync(root, { recursive: true, force: true });
});

describe('ASTRA_HISTORY_DIR', () => {
  it('keeps the default <docDir>/.astra/ when unset', async () => {
    delete process.env['ASTRA_HISTORY_DIR'];
    const doc = path.join(docsDir, 'a.docx');
    expect(getRepoDir(doc)).toBe(path.join(docsDir, '.astra', 'a.docx'));
    await initAstraRepo(doc, '<w:p>one</w:p>');
    expect(fs.existsSync(path.join(docsDir, '.astra', 'a.docx', '.git'))).toBe(true);
    expect(fs.existsSync(path.join(docsDir, '.astra-tombstones.json'))).toBe(true);
  });

  it('moves the sidecar and the tombstone registry under the history dir, leaving the document folder clean', async () => {
    process.env['ASTRA_HISTORY_DIR'] = historyDir;
    const doc = path.join(docsDir, 'a.docx');
    const repo = getRepoDir(doc);
    expect(path.dirname(path.dirname(repo))).toBe(historyDir);
    expect(path.basename(repo)).toBe('a.docx');

    await initAstraRepo(doc, '<w:p>one</w:p>');
    await commitChange(doc, '<w:p>two</w:p>', 'Second');
    expect((await getHistory(doc)).map((c) => c.message)).toEqual(['Second\n', 'Initial state\n']);
    expect(fs.existsSync(path.join(repo, '.git'))).toBe(true);
    expect(fs.existsSync(path.join(path.dirname(repo), '.astra-tombstones.json'))).toBe(true);

    // Nothing was written next to the document.
    expect(fs.readdirSync(docsDir)).toEqual([]);

    expect(await undo(doc)).toBe('<w:p>one</w:p>');
    expect((await getHistory(doc)).length).toBe(1);
  });

  it('gives two folders with the same file name separate histories', async () => {
    process.env['ASTRA_HISTORY_DIR'] = historyDir;
    const other = path.join(root, 'other');
    fs.mkdirSync(other);
    const a = path.join(docsDir, 'same.docx');
    const b = path.join(other, 'same.docx');
    expect(getRepoDir(a)).not.toBe(getRepoDir(b));
    await initAstraRepo(a, '<w:p>a</w:p>');
    await initAstraRepo(b, '<w:p>b</w:p>');
    await commitChange(a, '<w:p>a2</w:p>', 'Only in a');
    expect((await getHistory(a)).length).toBe(2);
    expect((await getHistory(b)).length).toBe(1);
  });

  it('keeps history for a document that lives inside the history dir itself', async () => {
    process.env['ASTRA_HISTORY_DIR'] = root;
    const doc = path.join(docsDir, 'a.docx');
    fs.writeFileSync(doc, 'not a real package');
    const repo = getRepoDir(doc);
    expect(repo).not.toBe(doc);
    await initAstraRepo(doc, '<w:p>one</w:p>');
    await commitChange(doc, '<w:p>two</w:p>', 'Second');
    expect((await getHistory(doc)).map((c) => c.message)).toEqual(['Second\n', 'Initial state\n']);
    expect(fs.statSync(doc).isFile()).toBe(true);
    expect(fs.readFileSync(doc, 'utf8')).toBe('not a real package');
    expect(fs.existsSync(path.join(repo, '.git'))).toBe(true);
  });

  it('gives a folder reached through a symlink the same history as the folder itself', () => {
    process.env['ASTRA_HISTORY_DIR'] = historyDir;
    const link = path.join(root, 'docs-link');
    fs.symlinkSync(docsDir, link, 'dir');
    expect(getRepoDir(path.join(link, 'a.docx'))).toBe(getRepoDir(path.join(docsDir, 'a.docx')));
  });

  it('still warns when a recorded sidecar was deleted', async () => {
    process.env['ASTRA_HISTORY_DIR'] = historyDir;
    const doc = path.join(docsDir, 'a.docx');
    await initAstraRepo(doc, '<w:p>one</w:p>');
    await commitChange(doc, '<w:p>two</w:p>', 'Second');
    fs.rmSync(getRepoDir(doc), { recursive: true, force: true });
    await initAstraRepo(doc, '<w:p>fresh</w:p>');
    expect(getSidecarHistoryWarning(doc)).toMatch(/possible history removal/);
  });

  it('treats an empty value as unset', () => {
    process.env['ASTRA_HISTORY_DIR'] = '   ';
    const doc = path.join(docsDir, 'a.docx');
    expect(getRepoDir(doc)).toBe(path.join(docsDir, '.astra', 'a.docx'));
  });
});
