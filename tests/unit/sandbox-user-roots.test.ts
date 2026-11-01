import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { assertWriteTargetNoSymlink, writeFileNoFollow } from '@attentionform/astra-engine';
import { getMcpAllowedRoots, validateSaveTarget } from '../../packages/mcp-server/src/lib/session-store';

let work = '';
const saved = process.env.ASTRA_EXTRA_DOC_ROOTS;

beforeEach(() => {
  work = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'user-roots-')));
});
afterEach(() => {
  if (saved === undefined) delete process.env.ASTRA_EXTRA_DOC_ROOTS;
  else process.env.ASTRA_EXTRA_DOC_ROOTS = saved;
  fs.rmSync(work, { recursive: true, force: true });
});

const mkdir = (...p: string[]) => { const d = path.join(work, ...p); fs.mkdirSync(d, { recursive: true }); return d; };
const bytes = Buffer.from('x');

describe('a folder the user allowed is the whole jail', () => {
  it('a root reached through a link can be written', async () => {
    const real = mkdir('real');
    const link = path.join(work, 'link');
    fs.symlinkSync(real, link, 'dir');
    process.env.ASTRA_EXTRA_DOC_ROOTS = link;
    const target = validateSaveTarget(path.join(link, 'a.docx'));
    await writeFileNoFollow(target, bytes);
    expect(fs.readFileSync(path.join(real, 'a.docx')).toString()).toBe('x');
  });

  it('a link above the root is the user\'s choice and is not examined', async () => {
    const real = mkdir('real', 'sub');
    const link = path.join(work, 'up');
    fs.symlinkSync(path.join(work, 'real'), link, 'dir');
    const root = path.join(link, 'sub');
    process.env.ASTRA_EXTRA_DOC_ROOTS = root;
    await writeFileNoFollow(validateSaveTarget(path.join(root, 'a.docx')), bytes);
    expect(fs.existsSync(path.join(real, 'a.docx'))).toBe(true);
  });

  it('a link inside the root that points outside is refused', () => {
    const root = mkdir('root');
    const outside = mkdir('outside');
    fs.symlinkSync(outside, path.join(root, 'escape'), 'dir');
    process.env.ASTRA_EXTRA_DOC_ROOTS = root;
    expect(() => validateSaveTarget(path.join(root, 'escape', 'a.docx'))).toThrow(/outside the allowed/);
  });

  it('a link inside the root that points inside is still refused for writing', () => {
    const root = mkdir('root', 'real');
    fs.symlinkSync(root, path.join(path.dirname(root), 'inner'), 'dir');
    process.env.ASTRA_EXTRA_DOC_ROOTS = path.dirname(root);
    expect(() => assertWriteTargetNoSymlink(path.join(path.dirname(root), 'inner', 'a.docx'))).toThrow(/symlink/);
  });

  it('a final component that is a link is refused', async () => {
    const root = mkdir('root');
    fs.writeFileSync(path.join(root, 'other.docx'), 'o');
    fs.symlinkSync(path.join(root, 'other.docx'), path.join(root, 'final.docx'));
    process.env.ASTRA_EXTRA_DOC_ROOTS = root;
    await expect(writeFileNoFollow(path.join(root, 'final.docx'), bytes)).rejects.toThrow(/symlink/);
  });

  it('.. traversal is refused', () => {
    const root = mkdir('root');
    process.env.ASTRA_EXTRA_DOC_ROOTS = root;
    expect(() => validateSaveTarget(path.join(root, '..', 'x.docx'))).toThrow(/traversal|outside/);
  });

  it('with a named folder the system temp folder is not allowed; without one it is', () => {
    const root = mkdir('root');
    const elsewhere = path.join(os.tmpdir(), `elsewhere-${process.pid}.docx`);
    process.env.ASTRA_EXTRA_DOC_ROOTS = root;
    expect(getMcpAllowedRoots()).toEqual([root]);
    expect(() => validateSaveTarget(elsewhere)).toThrow(/outside the allowed/);
    delete process.env.ASTRA_EXTRA_DOC_ROOTS;
    expect(getMcpAllowedRoots()).toContain(path.resolve(os.tmpdir()));
    expect(() => validateSaveTarget(elsewhere)).not.toThrow();
  });

  it('a path outside every root still sees every link up to the filesystem root', () => {
    const real = mkdir('real');
    const link = path.join(work, 'link');
    fs.symlinkSync(real, link, 'dir');
    process.env.ASTRA_EXTRA_DOC_ROOTS = mkdir('root');
    expect(() => assertWriteTargetNoSymlink(path.join(link, 'a.docx'))).toThrow(/symlink/);
  });
});
