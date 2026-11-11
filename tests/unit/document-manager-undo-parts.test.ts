import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import fsp from 'fs/promises';
import os from 'os';
import path from 'path';
import git from 'isomorphic-git';
import {
  AdmZip,
  DocumentManager,
  SNAPSHOT_FORMAT_MARKER,
  createBlankDocxBuffer,
  getRepoDir,
  isLegacyTrackedPackageFile,
} from '@attentionform/astra-engine';

/**
 * The desktop app's undo (DocumentManager.undoLastAction) rewinds the history branch one commit and
 * rebuilds the package. It must remove parts the earlier commit did not have, and leave alone the parts
 * an older commit never recorded.
 */
const PEOPLE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
  '<w15:person w15:author="Jane Secret Reviewer"><w15:presenceInfo w15:providerId="None" w15:userId="Jane Secret Reviewer"/></w15:person></w15:people>';

describe('DocumentManager undo and package parts', () => {
  let dir: string;
  let file: string;

  beforeEach(async () => {
    dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-dm-undo-'));
    file = path.join(dir, 'doc.docx');
    await fsp.writeFile(file, createBlankDocxBuffer('blank'));
  });

  afterEach(async () => {
    await fsp.rm(dir, { recursive: true, force: true });
  });

  const names = (): string[] => new AdmZip(fs.readFileSync(file)).getEntries().map((e) => e.entryName);

  async function addPeopleAndCommit(dm: DocumentManager): Promise<void> {
    const parser = dm.getParser();
    const snapshot = parser.capturePackageSnapshot();
    snapshot.set('word/people.xml', Buffer.from(PEOPLE, 'utf8'));
    parser.restorePackageSnapshot(snapshot);
    await dm.commitTransaction('add people.xml');
  }

  it('undo removes a part the earlier commit did not have, author name included', async () => {
    const dm = new DocumentManager();
    await dm.openDocument(file);
    const before = names();
    await addPeopleAndCommit(dm);
    expect(names()).toContain('word/people.xml');

    const res = await dm.undoLastAction();
    expect(res.success).toBe(true);
    expect(names()).toEqual(before);
    expect(new AdmZip(fs.readFileSync(file)).getEntries().some((e) => e.getData().toString('latin1').includes('Jane Secret'))).toBe(false);
  });

  it('undo to a commit recorded before complete snapshots leaves the parts it never recorded', async () => {
    const dm = new DocumentManager();
    await dm.openDocument(file);
    // Turn the first commit into a 0.4.1-style one: no marker, only the old tracked set.
    const repoDir = getRepoDir(file);
    for (const rel of await git.listFiles({ fs, dir: repoDir })) {
      if (rel === 'document.xml') continue;
      if (rel === SNAPSHOT_FORMAT_MARKER || !isLegacyTrackedPackageFile(rel)) {
        await git.remove({ fs, dir: repoDir, filepath: rel });
        await fsp.rm(path.join(repoDir, rel), { force: true });
      }
    }
    await git.commit({ fs, dir: repoDir, author: { name: 'x', email: 'x@example.com' }, message: 'legacy' });
    await addPeopleAndCommit(dm);
    const unrecorded = names().filter((n) => n !== 'word/document.xml' && !isLegacyTrackedPackageFile(n));
    expect(unrecorded).toContain('word/people.xml');

    const res = await dm.undoLastAction();
    expect(res.success).toBe(true);
    for (const name of unrecorded) expect(names(), name).toContain(name);
  });
});
