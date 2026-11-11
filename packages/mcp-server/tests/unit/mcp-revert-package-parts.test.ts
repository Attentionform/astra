import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import git from 'isomorphic-git';
import { createBlankDocxBuffer, getRepoDir, isLegacyTrackedPackageFile } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS, GIT_HISTORY_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleSanitizeDocument } from '../../src/lib/mcp/handlers/parity';
import { useReproducibleOutput } from '../../../../tests/helpers/reproducible';

useReproducibleOutput();

/**
 * End-user view of revert: what is inside the .docx zip after the agent adds comments, saves,
 * reverts and redoes. Every assertion reads decompressed zip entries (raw DOCX bytes can hide text).
 */

const MARKER = '.astra-snapshot-format';
const tmpDirs: string[] = [];
type Store = ReturnType<typeof createMcpSessionStore>;

afterEach(async () => {
  while (tmpDirs.length > 0) {
    await fs.promises.rm(tmpDirs.pop() as string, { recursive: true, force: true }).catch(() => undefined);
  }
});

async function makeSession(store: Store): Promise<{ sessionId: string; file: string }> {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-revert-parts-'));
  tmpDirs.push(dir);
  const file = path.join(dir, 'doc.docx');
  await fs.promises.writeFile(file, createBlankDocxBuffer('blank'));
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  if (!opened.ok) throw new Error('open failed in test setup');
  return { sessionId: opened.sessionId, file };
}

async function firstNodeId(store: Store, sessionId: string): Promise<string> {
  const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId, offset: 0, limit: 1 });
  if (!first.ok) throw new Error('inspect failed in test setup');
  const id = first.toonText
    .split('\n')
    .map((l) => l.split(/\s+/)[0])
    .filter((x) => !!x && !['>', 'ID'].includes(x) && !/^---/.test(x))[0];
  if (!id) throw new Error('no node id');
  return id;
}

async function save(store: Store, sessionId: string, message: string): Promise<string> {
  const res = await LEAN6_HANDLERS.save_document(store, { sessionId, message });
  if (!res.ok) throw new Error(`save failed: ${JSON.stringify(res)}`);
  return (res as unknown as { commitOid: string }).commitOid;
}

async function batch(store: Store, sessionId: string, operations: unknown[]): Promise<void> {
  const res = await LEAN6_HANDLERS.execute_batch(store, { sessionId, operations } as never);
  if (!res.ok) throw new Error(`batch failed: ${JSON.stringify(res)}`);
}

async function comment(store: Store, sessionId: string, author: string, text: string): Promise<void> {
  const target = await firstNodeId(store, sessionId);
  await batch(store, sessionId, [{ action: 'comment', operation: 'insert', target, text, author }]);
}

async function reply(store: Store, sessionId: string, author: string, text: string): Promise<void> {
  const target = await firstNodeId(store, sessionId);
  await batch(store, sessionId, [{ action: 'comment', operation: 'reply', target, text, author }]);
}

/** Decompressed contents of every zip entry, by name. */
function parts(file: string): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  for (const e of new AdmZip(file).getEntries()) if (!e.isDirectory) out.set(e.entryName, e.getData());
  return out;
}

function anywhere(file: string, needle: string): string[] {
  return [...parts(file)].filter(([, buf]) => buf.toString('latin1').includes(needle)).map(([n]) => n);
}

function expectSamePackage(actual: Map<string, Buffer>, expected: Map<string, Buffer>): void {
  expect([...actual.keys()].sort()).toEqual([...expected.keys()].sort());
  for (const [name, buf] of expected) expect(actual.get(name)?.equals(buf), `part ${name} differs`).toBe(true);
}

async function revertTo(store: Store, sessionId: string, commitOid: string) {
  const res = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, commitOid });
  if (!res.ok) throw new Error(`revert failed: ${JSON.stringify(res)}`);
  return res as typeof res & { snapshotComplete: boolean; restoreWarning?: string };
}

describe('revert restores every package part (0.4.2)', () => {
  it('reverting past the first comment leaves no comment parts and no author name', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    const c0 = await save(store, sessionId, 'before comments');
    const before = parts(file);
    await comment(store, sessionId, 'Jane Secret Reviewer', 'review me');
    await save(store, sessionId, 'with comment');
    expect(anywhere(file, 'Jane Secret Reviewer')).toContain('word/people.xml');

    await revertTo(store, sessionId, c0);

    const after = parts(file);
    expect(after.has('word/people.xml')).toBe(false);
    expect(after.has('word/commentsExtended.xml')).toBe(false);
    expect(after.has('word/comments.xml')).toBe(false);
    expect(anywhere(file, 'Jane Secret Reviewer')).toEqual([]);
    expectSamePackage(after, before);
  });

  it('reverting to a commit with fewer authors drops the later author and the dangling reply entry', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    await comment(store, sessionId, 'Alice', 'first');
    const c1 = await save(store, sessionId, 'alice comment');
    const atC1 = parts(file);
    await reply(store, sessionId, 'Bob Confidential', 'reply text');
    await save(store, sessionId, 'bob reply');
    expect(anywhere(file, 'Bob Confidential').length).toBeGreaterThan(0);

    await revertTo(store, sessionId, c1);

    expect(anywhere(file, 'Bob Confidential')).toEqual([]);
    expect(parts(file).get('word/commentsExtended.xml')?.toString()).not.toContain('paraIdParent');
    expectSamePackage(parts(file), atC1);
  });

  it('redo brings the later comment parts back', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    const c0 = await save(store, sessionId, 'before comments');
    await comment(store, sessionId, 'Jane Secret Reviewer', 'review me');
    await save(store, sessionId, 'with comment');
    const withComment = parts(file);

    await revertTo(store, sessionId, c0);
    expect(anywhere(file, 'Jane Secret Reviewer')).toEqual([]);
    const redo = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId });
    expect(redo.ok).toBe(true);

    expectSamePackage(parts(file), withComment);
  });

  it('a revert that touches no comments keeps every part byte-equal and the same entry order', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    const c0 = await save(store, sessionId, 'start');
    const originalOrder = new AdmZip(file).getEntries().map((e) => e.entryName);
    const target = await firstNodeId(store, sessionId);
    await batch(store, sessionId, [{ action: 'replace', target, text: 'changed text' }]);
    await save(store, sessionId, 'edit');
    await revertTo(store, sessionId, c0);
    expect(new AdmZip(file).getEntries().map((e) => e.entryName)).toEqual(originalOrder);
  });

  it('preview_revert lists the comment parts and reports a complete snapshot', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store);
    const c0 = await save(store, sessionId, 'before comments');
    await comment(store, sessionId, 'Jane Secret Reviewer', 'review me');
    await save(store, sessionId, 'with comment');

    const preview = (await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, commitOid: c0 })) as {
      ok: boolean;
      affectedFiles: string[];
      snapshotComplete: boolean;
      restoreWarning?: string;
    };
    expect(preview.ok).toBe(true);
    expect(preview.affectedFiles).toEqual(
      expect.arrayContaining(['word/comments.xml', 'word/commentsExtended.xml', 'word/people.xml'])
    );
    expect(preview.snapshotComplete).toBe(true);
    expect(preview.restoreWarning).toBeUndefined();
  });

  it('the snapshot marker lives in the history tree, never in the .docx', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    await save(store, sessionId, 'start');
    expect([...parts(file).keys()]).not.toContain(MARKER);
    const repoDir = getRepoDir(file);
    const head = await git.resolveRef({ fs, dir: repoDir, ref: 'HEAD' });
    const tree = await git.listFiles({ fs, dir: repoDir, ref: head });
    expect(tree).toContain(MARKER);
  });
});

describe('revert against a history store written before complete snapshots', () => {
  /** Turn the head commit into a 0.4.1-style commit: no marker, only the old tracked set. */
  async function rewriteHeadAsLegacy(file: string): Promise<string> {
    const repoDir = getRepoDir(file);
    for (const rel of await git.listFiles({ fs, dir: repoDir })) {
      if (rel === 'document.xml') continue;
      if (rel === MARKER || !isLegacyTrackedPackageFile(rel)) {
        await git.remove({ fs, dir: repoDir, filepath: rel });
        await fs.promises.rm(path.join(repoDir, rel), { force: true });
      }
    }
    return git.commit({
      fs,
      dir: repoDir,
      author: { name: 'ASTra Agent', email: 'astra@attentionform.com' },
      message: 'legacy snapshot',
    });
  }

  it('warns on preview and revert, never deletes parts the old commit never recorded', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    await save(store, sessionId, 'start');
    const legacyOid = await rewriteHeadAsLegacy(file);
    const target = await firstNodeId(store, sessionId);
    await batch(store, sessionId, [{ action: 'replace', target, text: 'later text' }]);
    await save(store, sessionId, 'later edit');
    const unrecorded = [...parts(file).keys()].filter(
      (n) => n !== 'word/document.xml' && !isLegacyTrackedPackageFile(n) && !n.endsWith('/')
    );
    expect(unrecorded.length).toBeGreaterThan(0);

    const preview = (await GIT_HISTORY_HANDLERS.preview_revert(store, { sessionId, commitOid: legacyOid })) as {
      ok: boolean;
      snapshotComplete: boolean;
      restoreWarning?: string;
    };
    expect(preview.ok).toBe(true);
    expect(preview.snapshotComplete).toBe(false);
    expect(preview.restoreWarning).toMatch(/older ASTra version/i);
    expect(preview.restoreWarning).toMatch(/not exact/i);

    const res = await revertTo(store, sessionId, legacyOid);
    expect(res.snapshotComplete).toBe(false);
    expect(res.restoreWarning).toMatch(/older ASTra version/i);
    for (const name of unrecorded) expect(parts(file).has(name), `${name} was deleted`).toBe(true);
    // The restored file still opens cleanly.
    const reopened = await LEAN6_HANDLERS.open_document(store, { path: file });
    expect(reopened.ok).toBe(true);
  });

  it('a new commit in an old store is marked, so a revert to it is complete and does not warn', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    await save(store, sessionId, 'start');
    await rewriteHeadAsLegacy(file);
    const target = await firstNodeId(store, sessionId);
    await batch(store, sessionId, [{ action: 'replace', target, text: 'new era' }]);
    const marked = await save(store, sessionId, 'marked commit');
    await batch(store, sessionId, [{ action: 'replace', target, text: 'newer era' }]);
    await save(store, sessionId, 'another');

    const res = await revertTo(store, sessionId, marked);
    expect(res.snapshotComplete).toBe(true);
    expect(res.restoreWarning).toBeUndefined();
  });
});

describe('sanitize_document removes comment author names from people.xml', () => {
  it('the default run removes the name from a document with a live comment', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    await comment(store, sessionId, 'Jane Secret Reviewer', 'review me');
    await save(store, sessionId, 'with comment');
    expect(anywhere(file, 'Jane Secret Reviewer')).toEqual(expect.arrayContaining(['word/comments.xml', 'word/people.xml']));
    const res = await handleSanitizeDocument(store, { sessionId });
    expect(res.ok).toBe(true);
    await save(store, sessionId, 'sanitized');
    expect(anywhere(file, 'Jane Secret Reviewer')).toEqual([]);
  });

  it('cleans a file that a 0.4.1 revert left with an orphan people.xml (no comments.xml)', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    await save(store, sessionId, 'start');
    await LEAN6_HANDLERS.close_document(store, { sessionId });
    // The 0.4.1 leftover: people.xml and commentsExtended.xml with no relationship and no comments.xml.
    const zip = new AdmZip(file);
    zip.addFile(
      'word/people.xml',
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
          '<w15:person w15:author="Jane Secret Reviewer"><w15:presenceInfo w15:providerId="None" w15:userId="Jane Secret Reviewer"/></w15:person></w15:people>'
      )
    );
    zip.addFile(
      'word/commentsExtended.xml',
      Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>')
    );
    fs.writeFileSync(file, zip.toBuffer());
    expect(anywhere(file, 'Jane Secret Reviewer')).toEqual(['word/people.xml']);

    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error('open failed');
    const res = await handleSanitizeDocument(store, { sessionId: opened.sessionId });
    expect(res.ok).toBe(true);
    await save(store, opened.sessionId, 'sanitized');
    expect(anywhere(file, 'Jane Secret Reviewer')).toEqual([]);
  });
});
