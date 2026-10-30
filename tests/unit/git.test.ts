import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import path from 'path';
import fs from 'fs';
import fsp from 'fs/promises';
import git from 'isomorphic-git';
import {
  checkoutCommit,
  commitChange,
  getHistory,
  getRepoDir,
  initAstraRepo,
  undo,
} from '@attentionform/astra-engine';
import os from 'os';

describe('Git Undo Stack via isomorphic-git', () => {
  let docPath: string;
  let repoDir: string;

  beforeEach(async () => {
    // Setup a temp file path for the test
    const tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-test-'));
    docPath = path.join(tmpDir, 'test_draft.astra.docx');
    repoDir = getRepoDir(docPath);
  });

  afterEach(async () => {
    // Cleanup
    await fsp.rm(path.dirname(repoDir), { recursive: true, force: true });
  });

  it('should initialize a git repo and make the first commit', async () => {
    const initialXml = '<w:document><w:body><w:p>Initial 1</w:p></w:body></w:document>';
    
    await initAstraRepo(docPath, initialXml);
    
    const history = await getHistory(docPath);
    expect(history.length).toBe(1);
    expect(history[0].message).toBe('Initial state\n');
  });

  it('should commit changes and track history', async () => {
    const initialXml = '<w:p>Initial</w:p>';
    await initAstraRepo(docPath, initialXml);
    
    const modifiedXml = '<w:p>Modified</w:p>';
    await commitChange(docPath, modifiedXml, 'Updated paragraph');
    
    const history = await getHistory(docPath);
    expect(history.length).toBe(2);
    expect(history[0].message).toBe('Updated paragraph\n');
    expect(history[1].message).toBe('Initial state\n');
  });

  it('should undo a commit and restore previous file content', async () => {
    const v1 = '<w:p>Version 1</w:p>';
    const v2 = '<w:p>Version 2</w:p>';
    const v3 = '<w:p>Version 3</w:p>';

    await initAstraRepo(docPath, v1);
    await commitChange(docPath, v2, 'v2');
    await commitChange(docPath, v3, 'v3');
    
    let history = await getHistory(docPath);
    expect(history.length).toBe(3);

    // First undo
    const restoredV2 = await undo(docPath);
    expect(restoredV2).toBe(v2);
    history = await getHistory(docPath);
    expect(history.length).toBe(2);

    // Second undo
    const restoredV1 = await undo(docPath);
    expect(restoredV1).toBe(v1);
    history = await getHistory(docPath);
    expect(history.length).toBe(1);

    // Trying to undo past initial state should throw
    await expect(undo(docPath)).rejects.toThrow('Cannot undo: already at the initial state.');
  });
});

describe('checkout-by-oid read-only restore via isomorphic-git', () => {
  let docPath: string;
  let repoDir: string;

  beforeEach(async () => {
    const tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-test-'));
    docPath = path.join(tmpDir, 'test_draft.astra.docx');
    repoDir = getRepoDir(docPath);
  });

  afterEach(async () => {
    await fsp.rm(path.dirname(repoDir), { recursive: true, force: true });
  });

  it('checkout-by-oid returns middle xml and leaves HEAD/history untouched', async () => {
    const v1 = '<w:p>Version 1</w:p>';
    const v2 = '<w:p>Version 2</w:p>';
    const v3 = '<w:p>Version 3</w:p>';

    await initAstraRepo(docPath, v1);
    await commitChange(docPath, v2, 'v2');
    await commitChange(docPath, v3, 'v3');

    const before = await getHistory(docPath);
    expect(before.length).toBe(3);
    const headBefore = before[0].oid;
    // git.log is newest-first: [v3, v2, v1]
    const middleOid = before[1].oid;

    const result = await checkoutCommit(docPath, middleOid);
    expect(result.xml).toBe(v2);
    expect(result.extra.size).toBe(0);

    const after = await getHistory(docPath);
    expect(after.length).toBe(3);
    expect(after[0].oid).toBe(headBefore);
    expect(after.map(c => c.oid)).toEqual(before.map(c => c.oid));

    // Working tree still holds the head version (read-only: no checkout).
    const liveXml = await fsp.readFile(path.join(repoDir, 'document.xml'), 'utf8');
    expect(liveXml).toBe(v3);
  });

  it('checkout-by-oid with unknown oid rejects as not found in history', async () => {
    await initAstraRepo(docPath, '<w:p>v1</w:p>');
    await commitChange(docPath, '<w:p>v2</w:p>', 'v2');
    await expect(
      checkoutCommit(docPath, 'deadbeefdeadbeefdeadbeefdeadbeefdeadbeef')
    ).rejects.toThrow(/not found in history/);
    const history = await getHistory(docPath);
    expect(history.length).toBe(2);
  });

  it('checkout-by-oid restores tracked extras from the target commit', async () => {
    const v1 = '<w:p>v1</w:p>';
    const v2 = '<w:p>v2</w:p>';
    const styles = '<w:styles><w:style>custom</w:style></w:styles>';
    await initAstraRepo(docPath, v1);
    await commitChange(docPath, v2, 'v2 with styles', {
      'word/styles.xml': styles,
    });

    const history = await getHistory(docPath);
    expect(history.length).toBe(2);
    const headOid = history[0].oid;

    const result = await checkoutCommit(docPath, headOid);
    expect(result.xml).toBe(v2);
    expect(result.extra.get('word/styles.xml')?.toString('utf8')).toBe(styles);
  });
});

describe('git hardening extras', () => {
  let docPath: string;
  let repoDir: string;

  beforeEach(async () => {
    const tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-test-'));
    docPath = path.join(tmpDir, 'test_draft.astra.docx');
    repoDir = getRepoDir(docPath);
  });

  afterEach(async () => {
    await fsp.rm(path.dirname(repoDir), { recursive: true, force: true });
  });

  it('rich-extras round-trip preserves charts/embeddings/theme/header-rels bytes', async () => {
    const v1 = '<w:p>v1</w:p>';
    const v2 = '<w:p>v2</w:p>';
    const chartBytes = Buffer.from([0x00, 0x3c, 0x63, 0xff, 0x00, 0x41, 0x42, 0xff]);
    const embeddingBytes = Buffer.from([0x00, 0xff, 0x89, 0x50, 0x4e, 0x47, 0x00, 0xff, 0x01]);
    const themeBytes = Buffer.from([0x3c, 0x00, 0x74, 0xff, 0x68, 0x00, 0x65, 0xff]);
    const headerRelBytes = Buffer.from([0x3c, 0x52, 0x00, 0x65, 0xff, 0x6c, 0x00, 0xff]);
    const footerRelBytes = Buffer.from([0x3c, 0x52, 0x00, 0x66, 0xff, 0x6c, 0x00, 0xff]);
    const docRelBytes = Buffer.from([0x3c, 0x52, 0x00, 0x64, 0xff, 0x6c, 0x00, 0xff]);
    const extras: Record<string, Buffer> = {
      'word/charts/chart1.xml': chartBytes,
      'word/embeddings/oleObject1.bin': embeddingBytes,
      'word/theme/theme1.xml': themeBytes,
      'word/_rels/header1.xml.rels': headerRelBytes,
      'word/_rels/footer1.xml.rels': footerRelBytes,
      'word/_rels/document.xml.rels': docRelBytes,
    };

    await initAstraRepo(docPath, v1);
    await commitChange(docPath, v2, 'v2 with rich extras', extras);

    const history = await getHistory(docPath);
    expect(history.length).toBe(2);
    const headOid = history[0].oid;
    const baseOid = history[1].oid;

    const head = await checkoutCommit(docPath, headOid);
    expect(head.xml).toBe(v2);
    for (const [rel, want] of Object.entries(extras)) {
      const got = head.extra.get(rel);
      expect(got, rel).toBeDefined();
      expect(got!.equals(want), rel).toBe(true);
    }

    const base = await checkoutCommit(docPath, baseOid);
    expect(base.xml).toBe(v1);
    expect(base.extra.size).toBe(0);
  });

  it('missing-document.xml commit rejection', async () => {
    await initAstraRepo(docPath, '<w:p>v1</w:p>', {
      'word/styles.xml': '<w:styles/>',
    });
    await fsp.unlink(path.join(repoDir, 'document.xml'));
    await git.remove({ fs, dir: repoDir, filepath: 'document.xml' });
    await git.commit({
      fs,
      dir: repoDir,
      author: { name: 'ASTra', email: 'astra@attentionform.com' },
      message: 'drop document',
    });

    const history = await getHistory(docPath);
    expect(history.length).toBe(2);
    await expect(checkoutCommit(docPath, history[0].oid)).rejects.toThrow(
      /document\.xml missing/
    );
  });

  it('uninitialized-path domain error mentions doc path and suggests open/save', async () => {
    const historyErr = await getHistory(docPath).catch((e: unknown) => e);
    expect(historyErr).toBeInstanceOf(Error);
    const historyMsg = String((historyErr as Error).message);
    expect(historyMsg).toContain(docPath);
    expect(historyMsg).toMatch(/no history exists yet/);
    expect(historyMsg).toMatch(/open or save/);

    const checkoutErr = await checkoutCommit(
      docPath,
      'deadbeefdeadbeefdeadbeefdeadbeefdeadbeef'
    ).catch((e: unknown) => e);
    expect(checkoutErr).toBeInstanceOf(Error);
    const checkoutMsg = String((checkoutErr as Error).message);
    expect(checkoutMsg).toContain(docPath);
    expect(checkoutMsg).toMatch(/no history exists yet/);
    expect(checkoutMsg).not.toMatch(/not found in history/);
  });

  it('concurrent checkouts leave working tree untouched', async () => {
    const v1 = '<w:p>Version 1</w:p>';
    const v2 = '<w:p>Version 2</w:p>';
    const v3 = '<w:p>Version 3</w:p>';

    await initAstraRepo(docPath, v1);
    await commitChange(docPath, v2, 'v2');
    await commitChange(docPath, v3, 'v3');

    const before = await getHistory(docPath);
    expect(before.length).toBe(3);
    const oidsNewestFirst = before.map(c => c.oid);

    const [rNewest, rMiddle, rOldest] = await Promise.all(
      oidsNewestFirst.map(oid => checkoutCommit(docPath, oid))
    );
    expect(rNewest.xml).toBe(v3);
    expect(rMiddle.xml).toBe(v2);
    expect(rOldest.xml).toBe(v1);

    const liveXml = await fsp.readFile(path.join(repoDir, 'document.xml'), 'utf8');
    expect(liveXml).toBe(v3);

    const after = await getHistory(docPath);
    expect(after.map(c => c.oid)).toEqual(oidsNewestFirst);
  });
});
