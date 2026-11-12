import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import fsp from 'fs/promises';
import os from 'os';
import path from 'path';
import { DocumentManager, DocxParser, checkoutCommit, createBlankDocxBuffer, getHistory } from '@attentionform/astra-engine';

/**
 * A package snapshot that cannot be collected must never reach the history as an empty set of parts: the history
 * reads a supplied set as the complete package, drops the parts missing from it and marks the commit complete.
 * Every DocumentManager commit collects the parts before it writes, so a failure leaves file, history and session
 * as they were.
 */
const PEOPLE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
  '<w15:person w15:author="Reviewer"><w15:presenceInfo w15:providerId="None" w15:userId="Reviewer"/></w15:person></w15:people>';

let dir: string;
let file: string;
beforeEach(async () => {
  dir = await fsp.realpath(await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-dm-snapshot-fail-')));
  vi.stubEnv('ASTRA_EXTRA_DOC_ROOTS', dir);
  file = path.join(dir, 'doc.docx');
  await fsp.writeFile(file, createBlankDocxBuffer('blank'));
});
afterEach(async () => {
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
  await fsp.rm(dir, { recursive: true, force: true });
});

async function head(): Promise<{ commits: number; parts: string[]; complete: boolean }> {
  const entries = await getHistory(file);
  const state = await checkoutCommit(file, entries[0].oid);
  return { commits: entries.length, parts: [...state.extra.keys()].sort(), complete: state.complete };
}
function addPeople(dm: DocumentManager): void {
  const parser = dm.getParser();
  const snapshot = parser.capturePackageSnapshot();
  snapshot.set('word/people.xml', Buffer.from(PEOPLE, 'utf8'));
  parser.restorePackageSnapshot(snapshot);
}
function breakSnapshots() {
  return vi.spyOn(DocxParser.prototype, 'capturePackageSnapshot').mockImplementation(() => { throw new Error('injected snapshot collection failure'); });
}

describe('DocumentManager history when package parts cannot be collected', () => {
  it('has real parts to lose', async () => {
    await new DocumentManager().openDocument(file);
    const state = await head();
    expect(state.parts.length).toBeGreaterThan(3);
    expect(state.complete).toBe(true);
  });

  it('Save As over a document that already has history writes nothing and records nothing', async () => {
    const dm = new DocumentManager();
    await dm.openDocument(file);
    const before = await head();
    const bytes = await fsp.readFile(file);
    const snapshot = breakSnapshots();
    await expect(dm.saveAsDocument(file, { applyTypography: false })).rejects.toMatchObject({ code: 'E_STATE' });
    snapshot.mockRestore();
    expect(await fsp.readFile(file)).toEqual(bytes);
    expect(await head()).toEqual(before);
  });

  it('commitTransaction fails before the write, keeps the history, and puts the in-memory edit back', async () => {
    const dm = new DocumentManager();
    await dm.openDocument(file);
    const before = await head();
    const bytes = await fsp.readFile(file);
    addPeople(dm);
    const snapshot = breakSnapshots();
    await expect(dm.commitTransaction('add people part')).rejects.toMatchObject({ code: 'E_STATE' });
    snapshot.mockRestore();
    expect(await fsp.readFile(file)).toEqual(bytes);
    expect(await head()).toEqual(before);
    expect([...dm.getParser().capturePackageSnapshot().keys()]).not.toContain('word/people.xml');
  });

  it('applyTrackedChanges fails before the write and keeps the history', async () => {
    const dm = new DocumentManager();
    await dm.openDocument(file);
    const before = await head();
    const bytes = await fsp.readFile(file);
    const snapshot = breakSnapshots();
    await expect(dm.applyTrackedChanges('accept')).rejects.toMatchObject({ code: 'E_STATE' });
    snapshot.mockRestore();
    expect(await fsp.readFile(file)).toEqual(bytes);
    expect(await head()).toEqual(before);
  });

  it('redo fails before the write, keeps the redo step available, and keeps the history', async () => {
    const dm = new DocumentManager();
    await dm.openDocument(file);
    addPeople(dm);
    await dm.commitTransaction('add people part');
    expect((await dm.undoLastAction()).success).toBe(true);
    const before = await head();
    const bytes = await fsp.readFile(file);
    const snapshot = breakSnapshots();
    await expect(dm.redoLastAction()).rejects.toMatchObject({ code: 'E_STATE' });
    snapshot.mockRestore();
    expect(await fsp.readFile(file)).toEqual(bytes);
    expect(await head()).toEqual(before);
    expect((await dm.redoLastAction()).success).toBe(true);
    expect((await head()).commits).toBe(before.commits + 1);
  });
});
