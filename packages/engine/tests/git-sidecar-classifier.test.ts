/**
 * Sidecar `.git` classifier pins for the engine write path
 * (`ensureAstraGitRepo` via `initAstraRepo` / `commitChange`), mirroring the
 * CLI history classifier (`sidecarGitState` in
 * packages/cli/src/cli/history-cmd.ts): a real directory — directly or
 * through a symlink to a directory — is a usable repo; a file (e.g. a
 * leftover `gitdir:` redirect), a symlink to a file, or a broken symlink is
 * a stray entry that refuses fail-closed with E_IO and writes nothing
 * wherever the redirect points. Isomorphic-git follows `.git` redirect
 * files, so an unclassified file previously let init/commit write objects
 * and refs into the redirect target outside the sidecar.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import fsp from 'fs/promises';
import path from 'path';
import os from 'os';
import git from 'isomorphic-git';
import { commitChange, getHistory, getRepoDir, initAstraRepo, undoPackage } from '../src/git';

const XML = '<w:document><w:body><w:p>probe</w:p></w:body></w:document>';

let root: string;

function listTree(dir: string): string[] {
  const out: string[] = [];
  const walk = (d: string, base: string): void => {
    let entries: fs.Dirent[] = [];
    try {
      entries = fs.readdirSync(d, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      const rel = base ? `${base}/${e.name}` : e.name;
      out.push(rel);
      if (e.isDirectory()) walk(path.join(d, e.name), rel);
    }
  };
  walk(dir, '');
  return out.sort();
}

async function captureError(p: Promise<unknown>): Promise<{ code?: unknown; message: string; remedy: string }> {
  try {
    await p;
    throw new Error('expected the call to refuse');
  } catch (e) {
    return {
      code: (e as { code?: unknown }).code,
      message: String((e as Error).message),
      remedy: String((e as { remedy?: unknown }).remedy ?? ''),
    };
  }
}

describe('sidecar .git classifier (engine write path)', () => {
  beforeEach(async () => {
    root = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-git-classifier-'));
  });

  afterEach(async () => {
    await fsp.rm(root, { recursive: true, force: true });
  });

  it('stray .git file: commitChange refuses with E_IO + remedy and writes nothing into the target', async () => {
    const evil = path.join(root, 'evil-target');
    await fsp.mkdir(evil, { recursive: true });
    await git.init({ fs, dir: evil });
    const docPath = path.join(root, 'doc.docx');
    const sidecarDir = getRepoDir(docPath);
    await fsp.mkdir(sidecarDir, { recursive: true });
    await fsp.writeFile(path.join(sidecarDir, '.git'), `gitdir: ${path.join(evil, '.git')}\n`);

    const before = listTree(evil);
    const err = await captureError(commitChange(docPath, XML, 'probe commit'));

    expect(err.code).toBe('E_IO');
    expect(err.message).toContain('stray');
    expect(err.message).toContain('redirect');
    expect(err.remedy).toContain('Remove the stray');
    expect(listTree(evil)).toEqual(before);
    expect(listTree(sidecarDir)).toEqual(['.git']);
    expect(fs.readFileSync(path.join(sidecarDir, '.git'), 'utf8')).toBe(`gitdir: ${path.join(evil, '.git')}\n`);
  });

  it('stray .git file: initAstraRepo refuses and never creates the missing redirect target', async () => {
    const docPath = path.join(root, 'doc2.docx');
    const sidecarDir = getRepoDir(docPath);
    await fsp.mkdir(sidecarDir, { recursive: true });
    const missing = path.join(root, 'never-created');
    await fsp.writeFile(path.join(sidecarDir, '.git'), `gitdir: ${missing}\n`);

    const err = await captureError(initAstraRepo(docPath, XML));

    expect(err.code).toBe('E_IO');
    expect(err.remedy).toContain('.astra');
    expect(fs.existsSync(missing)).toBe(false);
    expect(listTree(sidecarDir)).toEqual(['.git']);
  });

  it('broken .git symlink and symlink-to-file are stray, not absent', async () => {
    const docPath = path.join(root, 'doc3.docx');
    const sidecarDir = getRepoDir(docPath);
    await fsp.mkdir(sidecarDir, { recursive: true });
    const fileTarget = path.join(root, 'plain.txt');
    await fsp.writeFile(fileTarget, 'not a gitdir');

    await fsp.symlink(path.join(sidecarDir, 'missing-target'), path.join(sidecarDir, '.git'));
    const broken = await captureError(initAstraRepo(docPath, XML));
    expect(broken.code).toBe('E_IO');
    expect(broken.message).toContain('stray');

    await fsp.unlink(path.join(sidecarDir, '.git'));
    await fsp.symlink(fileTarget, path.join(sidecarDir, '.git'));
    const toFile = await captureError(initAstraRepo(docPath, XML));
    expect(toFile.code).toBe('E_IO');
    expect(listTree(sidecarDir)).toEqual(['.git']);
  });

  it('stray .git file: undoPackage refuses before any ref/checkout write', async () => {
    const evil = path.join(root, 'undo-target');
    await fsp.mkdir(evil, { recursive: true });
    await git.init({ fs, dir: evil });
    const docPath = path.join(root, 'doc6.docx');
    const sidecarDir = getRepoDir(docPath);
    await fsp.mkdir(sidecarDir, { recursive: true });
    await fsp.writeFile(path.join(sidecarDir, '.git'), `gitdir: ${path.join(evil, '.git')}\n`);

    const before = listTree(evil);
    const err = await captureError(undoPackage(docPath));
    expect(err.code).toBe('E_IO');
    expect(listTree(evil)).toEqual(before);
    expect(listTree(sidecarDir)).toEqual(['.git']);
  });

  it('legit .git directory: absent auto-inits, then commits append history', async () => {
    const docPath = path.join(root, 'doc4.docx');
    await initAstraRepo(docPath, XML);
    await commitChange(docPath, XML.replace('probe', 'probe two'), 'second');

    const history = await getHistory(docPath);
    expect(history.length).toBe(2);
    expect(fs.lstatSync(path.join(getRepoDir(docPath), '.git')).isDirectory()).toBe(true);
  });

  it('.git symlink to a real directory stays usable (CLI classifier parity)', async () => {
    const host = path.join(root, 'git-host');
    await fsp.mkdir(host);
    await git.init({ fs, dir: host });
    const docPath = path.join(root, 'doc5.docx');
    const sidecarDir = getRepoDir(docPath);
    await fsp.mkdir(sidecarDir, { recursive: true });
    await fsp.symlink(path.join(host, '.git'), path.join(sidecarDir, '.git'));

    await initAstraRepo(docPath, XML);

    const history = await getHistory(docPath);
    expect(history.length).toBe(1);
    expect(fs.lstatSync(path.join(sidecarDir, '.git')).isSymbolicLink()).toBe(true);
    expect(listTree(path.join(host, '.git')).length).toBeGreaterThan(0);
  });
});
