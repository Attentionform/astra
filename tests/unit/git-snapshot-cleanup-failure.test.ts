import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import path from 'path';
import fsp from 'fs/promises';
import os from 'os';
import fs from 'fs';
import git from 'isomorphic-git';
import { SNAPSHOT_FORMAT_MARKER, checkoutCommit, commitChange, getHistory, getRepoDir, initAstraRepo } from '@attentionform/astra-engine';

// A commit may only claim to hold a complete snapshot when the stale parts were really removed.
// Here the removal of one stale part fails: the commit must still be written (the save must not
// break), but it must not carry the completeness marker.
const hooks = vi.hoisted(() => ({ failRemoveFor: null as string | null }));

vi.mock('isomorphic-git', async (importOriginal) => {
  const actual = await importOriginal<typeof import('isomorphic-git')>();
  const real = actual.default as unknown as Record<string, unknown>;
  const wrapped = new Proxy(real, {
    get(target, key) {
      if (key === 'remove') {
        return async (opts: { filepath: string }) => {
          if (hooks.failRemoveFor !== null && opts.filepath === hooks.failRemoveFor) throw new Error('injected remove failure');
          return (target.remove as (o: unknown) => Promise<unknown>)(opts);
        };
      }
      return target[key as string];
    },
  });
  return { ...actual, default: wrapped };
});

describe('completeness marker when cleanup of stale parts fails', () => {
  let docPath: string;
  let repoDir: string;

  beforeEach(async () => {
    const dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-snapfail-'));
    docPath = path.join(dir, 'doc.docx');
    repoDir = getRepoDir(docPath);
  });

  afterEach(async () => {
    hooks.failRemoveFor = null;
    await fsp.rm(path.dirname(repoDir), { recursive: true, force: true });
    await fsp.rm(path.dirname(docPath), { recursive: true, force: true });
  });

  it('leaves the commit unmarked when a stale part cannot be removed, and marks it again once cleanup works', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/people.xml': Buffer.from('stale'), 'word/styles.xml': Buffer.from('s') });
    hooks.failRemoveFor = 'word/people.xml';
    await commitChange(docPath, '<b/>', 'people.xml dropped from the package', { 'word/styles.xml': Buffer.from('s') });
    const oid = (await getHistory(docPath))[0].oid;
    const tree = await git.listFiles({ fs, dir: repoDir, ref: oid });
    expect(tree).not.toContain(SNAPSHOT_FORMAT_MARKER);
    expect((await checkoutCommit(docPath, oid)).complete).toBe(false);

    hooks.failRemoveFor = null;
    await commitChange(docPath, '<c/>', 'cleanup works again', { 'word/styles.xml': Buffer.from('s') });
    const next = await checkoutCommit(docPath, (await getHistory(docPath))[0].oid);
    expect(next.complete).toBe(true);
    expect([...next.extra.keys()]).toEqual(['word/styles.xml']);
  });
});
