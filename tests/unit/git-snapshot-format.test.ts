import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import path from 'path';
import fs from 'fs';
import fsp from 'fs/promises';
import os from 'os';
import zlib from 'zlib';
import git from 'isomorphic-git';
import {
  AdmZip,
  DocxParser,
  SNAPSHOT_FORMAT_MARKER,
  SNAPSHOT_FORMAT_VERSION,
  checkoutCommit,
  commitChange,
  createBlankDocxBuffer,
  getHistory,
  getRepoDir,
  initAstraRepo,
  isLegacyTrackedPackageFile,
  isTrackedPackageFile,
  entryNameFor,
  historyPathFor,
  isCompleteSnapshotMarker,
  needsHistoryEscape,
  trackedPredicateFor,
  undoPackage,
} from '@attentionform/astra-engine';

describe('which package parts history records', () => {
  it('records every part except word/document.xml, including the comment metadata parts', () => {
    for (const name of [
      'word/people.xml',
      'word/commentsExtended.xml',
      'word/commentsIds.xml',
      'word/commentsExtensible.xml',
      'word/comments.xml',
      'word/styles.xml',
      'word/media/image1.png',
      'docProps/core.xml',
      'docProps/app.xml',
      'docProps/thumbnail.jpeg',
      'customXml/item1.xml',
      'word/fontTable.xml',
      'word/webSettings.xml',
      '_rels/.rels',
      '[Content_Types].xml',
      'word/_rels/document.xml.rels',
      'word/my part ünï.xml',
    ]) {
      expect(isTrackedPackageFile(name), name).toBe(true);
    }
    expect(isTrackedPackageFile('word/document.xml')).toBe(false);
  });

  it('never records entries without content or names that can never be stored', () => {
    for (const name of ['', 'word/', '/etc/passwd', 'C:/x.xml', '../evil.xml', 'word/../../evil.xml', 'word\\..\\evil.xml']) {
      expect(isTrackedPackageFile(name), name).toBe(false);
    }
  });

  it('records names git cannot hold as they are, by escaping them', () => {
    for (const name of [
      '.git/config',
      'word/.git/config',
      '.gitignore',
      'customXml/.gitignore',
      '.gitattributes',
      '.gitmodules',
      '.GIT/config',
      SNAPSHOT_FORMAT_MARKER,
      'document.xml',
      '__astra_escaped__/looks-like-ours',
    ]) {
      expect(needsHistoryEscape(name), name).toBe(true);
      expect(isTrackedPackageFile(name), name).toBe(true);
      const repoPath = historyPathFor(name);
      expect(repoPath.startsWith('__astra_escaped__/')).toBe(true);
      expect(repoPath).not.toMatch(/\.git/i);
      expect(entryNameFor(repoPath)).toBe(name);
    }
    expect(historyPathFor('word/people.xml')).toBe('word/people.xml');
    expect(entryNameFor('word/people.xml')).toBe('word/people.xml');
  });

  it('only a marker with exactly the current version counts as complete', () => {
    expect(isCompleteSnapshotMarker(`${SNAPSHOT_FORMAT_VERSION}\n`)).toBe(true);
    expect(isCompleteSnapshotMarker(Buffer.from(SNAPSHOT_FORMAT_VERSION))).toBe(true);
    for (const bad of ['', '1\n', '3\n', '99', 'garbage', '2 extra', undefined]) {
      expect(isCompleteSnapshotMarker(bad as never), String(bad)).toBe(false);
    }
  });

  it('keeps the old whitelist available for commits recorded before the marker', () => {
    expect(isLegacyTrackedPackageFile('word/comments.xml')).toBe(true);
    expect(isLegacyTrackedPackageFile('word/media/a.png')).toBe(true);
    expect(isLegacyTrackedPackageFile('word/people.xml')).toBe(false);
    expect(isLegacyTrackedPackageFile('docProps/core.xml')).toBe(false);
    expect(trackedPredicateFor({ complete: true })).toBe(isTrackedPackageFile);
    expect(trackedPredicateFor({ complete: false })).toBe(isLegacyTrackedPackageFile);
  });
});

describe('history commits carry the snapshot marker and store parts safely', () => {
  let docPath: string;
  let repoDir: string;

  beforeEach(async () => {
    const tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-snapfmt-'));
    docPath = path.join(tmpDir, 'doc.docx');
    repoDir = getRepoDir(docPath);
  });

  afterEach(async () => {
    await fsp.rm(path.dirname(repoDir), { recursive: true, force: true });
    await fsp.rm(path.dirname(docPath), { recursive: true, force: true });
  });

  const head = async (): Promise<string> => (await getHistory(docPath))[0].oid;
  const treeAt = async (oid: string): Promise<string[]> => git.listFiles({ fs, dir: repoDir, ref: oid });

  it('writes the marker in the first commit and in every later commit', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/people.xml': Buffer.from('p1') });
    const first = await head();
    expect(await treeAt(first)).toContain(SNAPSHOT_FORMAT_MARKER);
    await commitChange(docPath, '<b/>', 'second', { 'word/people.xml': Buffer.from('p2') });
    const second = await head();
    expect(await treeAt(second)).toContain(SNAPSHOT_FORMAT_MARKER);
    const marker = await git.readBlob({ fs, dir: repoDir, oid: second, filepath: SNAPSHOT_FORMAT_MARKER });
    expect(Buffer.from(marker.blob).toString('utf8')).toBe(`${SNAPSHOT_FORMAT_VERSION}\n`);
  });

  it('checkoutCommit reports complete for marked commits and the exact bytes of every part', async () => {
    const binary = Buffer.from([0, 1, 2, 255, 254, 10, 13, 0]);
    await initAstraRepo(docPath, '<a/>', {
      'word/people.xml': Buffer.from('<people/>'),
      'docProps/thumbnail.jpeg': binary,
      'word/my part ünï.xml': Buffer.from('<x/>'),
    });
    const target = await checkoutCommit(docPath, await head());
    expect(target.complete).toBe(true);
    expect(target.extra.get('word/people.xml')!.toString()).toBe('<people/>');
    expect(target.extra.get('docProps/thumbnail.jpeg')!.equals(binary)).toBe(true);
    expect(target.extra.get('word/my part ünï.xml')!.toString()).toBe('<x/>');
    expect(target.extra.has(SNAPSHOT_FORMAT_MARKER)).toBe(false);
  });

  it('a part removed from the package is removed from the next commit', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/people.xml': Buffer.from('p'), 'word/styles.xml': Buffer.from('s') });
    await commitChange(docPath, '<b/>', 'drop people', { 'word/styles.xml': Buffer.from('s') });
    const target = await checkoutCommit(docPath, await head());
    expect([...target.extra.keys()]).toEqual(['word/styles.xml']);
  });

  it("stores git's own names escaped, round-trips them exactly, and never lets a document change the history repo", async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/styles.xml': Buffer.from('s') });
    const gitConfigBefore = await fsp.readFile(path.join(repoDir, '.git', 'config'), 'utf8');
    const hostile: Record<string, Buffer> = {
      'word/styles.xml': Buffer.from('s2'),
      '.git/config': Buffer.from('[core]\n\tbare = true\n'),
      '.gitignore': Buffer.from('*\n'),
      'customXml/.gitignore': Buffer.from('*\n'),
      '.gitattributes': Buffer.from('* filter=x\n'),
      [SNAPSHOT_FORMAT_MARKER]: Buffer.from('999\n'),
      'document.xml': Buffer.from('<not-the-document/>'),
      '__astra_escaped__/looks-like-ours': Buffer.from('collision'),
      'word/people.xml': Buffer.from('<people/>'),
    };
    await commitChange(docPath, '<b/>', 'hostile names', hostile);
    const oid = await head();
    expect(await fsp.readFile(path.join(repoDir, '.git', 'config'), 'utf8')).toBe(gitConfigBefore);
    const tree = await treeAt(oid);
    for (const bad of ['.gitignore', 'customXml/.gitignore', '.gitattributes', 'document.xml.bak']) expect(tree).not.toContain(bad);
    const marker = await git.readBlob({ fs, dir: repoDir, oid, filepath: SNAPSHOT_FORMAT_MARKER });
    expect(Buffer.from(marker.blob).toString('utf8')).toBe(`${SNAPSHOT_FORMAT_VERSION}\n`);
    const doc = await git.readBlob({ fs, dir: repoDir, oid, filepath: 'document.xml' });
    expect(Buffer.from(doc.blob).toString('utf8')).toBe('<b/>');
    const target = await checkoutCommit(docPath, oid);
    expect(target.complete).toBe(true);
    expect([...target.extra.keys()].sort()).toEqual(Object.keys(hostile).sort());
    for (const [name, buf] of Object.entries(hostile)) expect(target.extra.get(name)!.equals(buf), name).toBe(true);
    // The other parts still index normally: a .gitignore part must not hide them.
    expect(tree).toContain('word/people.xml');
    expect(tree).toContain('word/styles.xml');
  });

  it('a part too long to escape fails the commit clearly instead of being skipped', async () => {
    await initAstraRepo(docPath, '<a/>');
    const longName = `.git${'x'.repeat(300)}.xml`;
    await expect(commitChange(docPath, '<b/>', 'long', { [longName]: Buffer.from('x') })).rejects.toMatchObject({
      code: 'E_VALIDATION',
    });
  });

  it('a commit made without a package snapshot is not marked, and clears the marker of the commit before it', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/people.xml': Buffer.from('p') });
    expect(await treeAt(await head())).toContain(SNAPSHOT_FORMAT_MARKER);
    await commitChange(docPath, '<b/>', 'document.xml only');
    const oid = await head();
    expect(await treeAt(oid)).not.toContain(SNAPSHOT_FORMAT_MARKER);
    expect((await checkoutCommit(docPath, oid)).complete).toBe(false);
    // A later commit with a full snapshot is marked again.
    await commitChange(docPath, '<c/>', 'full snapshot', { 'word/people.xml': Buffer.from('p3') });
    expect((await checkoutCommit(docPath, await head())).complete).toBe(true);
  });

  it('undoPackage reads a malformed marker on the commit it returns to as incomplete', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/styles.xml': Buffer.from('s') });
    await fsp.writeFile(path.join(repoDir, SNAPSHOT_FORMAT_MARKER), '3\n');
    await git.add({ fs, dir: repoDir, filepath: SNAPSHOT_FORMAT_MARKER });
    await git.commit({ fs, dir: repoDir, author: { name: 'x', email: 'x@example.com' }, message: 'future marker' });
    await commitChange(docPath, '<b/>', 'next', { 'word/styles.xml': Buffer.from('s2') });
    const undone = await undoPackage(docPath);
    expect(undone.complete).toBe(false);
  });

  it('a malformed or future marker reads as incomplete', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/styles.xml': Buffer.from('s') });
    for (const content of ['3\n', 'garbage', '']) {
      await fsp.writeFile(path.join(repoDir, SNAPSHOT_FORMAT_MARKER), content);
      await git.add({ fs, dir: repoDir, filepath: SNAPSHOT_FORMAT_MARKER });
      await git.commit({ fs, dir: repoDir, author: { name: 'x', email: 'x@example.com' }, message: `marker ${content}` });
      expect((await checkoutCommit(docPath, await head())).complete, JSON.stringify(content)).toBe(false);
    }
  });

  it('still refuses a path that climbs out of the history repo', async () => {
    await initAstraRepo(docPath, '<a/>');
    await expect(commitChange(docPath, '<b/>', 'traversal', { '../outside.xml': Buffer.from('x') })).rejects.toThrow(
      /Security Violation/
    );
  });

  it('undoPackage reports complete for a marked commit', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/people.xml': Buffer.from('p') });
    await commitChange(docPath, '<b/>', 'second', { 'word/people.xml': Buffer.from('p2') });
    const undone = await undoPackage(docPath);
    expect(undone.complete).toBe(true);
    expect(undone.extra.get('word/people.xml')!.toString()).toBe('p');
  });

  it('an old commit without the marker reports incomplete', async () => {
    await initAstraRepo(docPath, '<a/>', { 'word/styles.xml': Buffer.from('s') });
    await git.remove({ fs, dir: repoDir, filepath: SNAPSHOT_FORMAT_MARKER });
    await fsp.rm(path.join(repoDir, SNAPSHOT_FORMAT_MARKER), { force: true });
    await git.commit({ fs, dir: repoDir, author: { name: 'x', email: 'x@example.com' }, message: 'legacy' });
    const target = await checkoutCommit(docPath, await head());
    expect(target.complete).toBe(false);
    expect(target.extra.get('word/styles.xml')!.toString()).toBe('s');
  });
});

/** Minimal zip writer so entries can carry a deflate level other than the one adm-zip uses. */
function writeZip(entries: Array<{ name: string; data: Buffer; level: number }>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const name = Buffer.from(e.name, 'utf8');
    const packed = zlib.deflateRawSync(e.data, { level: e.level });
    const crc = zlib.crc32(e.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0x21, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18);
    local.writeUInt32LE(e.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    locals.push(local, name, packed);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0x21, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(packed.length, 20);
    central.writeUInt32LE(e.data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, name);
    offset += 30 + name.length + packed.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, centralBuf, end]);
}

describe('restoring a package snapshot keeps untouched entries as they were stored', () => {
  function sourceDocx(): { buffer: Buffer } {
    const blank = new AdmZip(createBlankDocxBuffer('blank'));
    const entries = blank
      .getEntries()
      .filter((e) => !e.isDirectory)
      .map((e) => ({ name: e.entryName, data: e.getData(), level: 1 }));
    const filler = Buffer.from('<root>' + 'repeated text '.repeat(400) + '</root>', 'utf8');
    entries.push({ name: 'customXml/item1.xml', data: filler, level: 1 });
    return { buffer: writeZip(entries) };
  }

  it('does not recompress an entry whose bytes did not change, and keeps entry order', () => {
    const { buffer } = sourceDocx();
    const control = new AdmZip(new DocxParser(buffer).rezip()).getEntries().map((e) => e.entryName);
    const original = new AdmZip(buffer);
    const storedSize = original.getEntry('customXml/item1.xml')!.header.compressedSize;
    expect(zlib.deflateRawSync(original.getEntry('customXml/item1.xml')!.getData()).length).not.toBe(storedSize);

    const parser = new DocxParser(buffer);
    parser.restorePackageSnapshot(parser.capturePackageSnapshot());
    const out = new AdmZip(parser.rezip());
    expect(out.getEntry('customXml/item1.xml')!.header.compressedSize).toBe(storedSize);
    expect(out.getEntries().map((e) => e.entryName)).toEqual(control);
  });

  it('still writes an entry whose bytes changed and removes an entry the snapshot dropped', () => {
    const { buffer } = sourceDocx();
    const parser = new DocxParser(buffer);
    const snapshot = parser.capturePackageSnapshot();
    snapshot.set('customXml/item1.xml', Buffer.from('<root>changed</root>'));
    snapshot.delete('docProps/core.xml');
    snapshot.set('word/people.xml', Buffer.from('<people/>'));
    parser.restorePackageSnapshot(snapshot);
    const out = new AdmZip(parser.rezip());
    expect(out.getEntry('customXml/item1.xml')!.getData().toString()).toBe('<root>changed</root>');
    expect(out.getEntry('docProps/core.xml')).toBeNull();
    expect(out.getEntry('word/people.xml')!.getData().toString()).toBe('<people/>');
  });
});
