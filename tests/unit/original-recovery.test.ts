import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {
  captureOriginal, commitChange, exportOriginal, getOriginalInfo, getRepoDir,
  initAstraRepo, DocumentManager, createBlankDocxBuffer, originalSha256, readOriginal, replaceFileAtomic, writeFileExclusive, checkoutCommit, getHistory,
} from '@attentionform/astra-engine';

let dir: string;
let doc: string;
beforeEach(async () => { dir = await fsp.realpath(await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-original-'))); doc = path.join(dir, 'test.docx'); });
afterEach(async () => { vi.restoreAllMocks(); vi.unstubAllEnvs(); await fsp.rm(dir, { recursive: true, force: true }); });
const data = Buffer.from('PK original bytes: sensitive reviewer');
describe('exact original capture and export', () => {
  it('records exact input, retains it across saves and keeps it out of package parts', async () => {
    await initAstraRepo(doc, '<a/>', {}, data);
    const info = await getOriginalInfo(doc);
    expect(info).toMatchObject({ available: true, sha256: originalSha256(data), bytes: data.length });
    for (let i = 0; i < 5; i++) await commitChange(doc, `<a>${i}</a>`, 'edit', {});
    expect(await readOriginal(doc)).toEqual(data);
    const target = await checkoutCommit(doc, (await getHistory(doc))[0].oid);
    expect([...target.extra.keys()].some((p) => p.includes('astra-original'))).toBe(false);
    const out = path.join(dir, 'out.docx');
    await exportOriginal(doc, out);
    expect(await fsp.readFile(out)).toEqual(data);
    await expect(exportOriginal(doc, out)).rejects.toMatchObject({ code: 'E_CONFLICT' });
    await exportOriginal(doc, out, { overwrite: true });
    expect(originalSha256(await fsp.readFile(out))).toBe(originalSha256(data));
  });
  it('Save As captures the preexisting destination while preserving the source original', async () => {
    const sourceBytes = createBlankDocxBuffer('blank');
    await fsp.writeFile(doc, sourceBytes);
    const destination = path.join(dir, 'existing.docx');
    const destinationBytes = Buffer.concat([createBlankDocxBuffer('blank'), Buffer.from('original trailer')]);
    await fsp.writeFile(destination, destinationBytes);
    const dm = new DocumentManager();
    await dm.openDocument(doc);
    await dm.saveAsDocument(destination, { applyTypography: false });
    expect(await readOriginal(doc)).toEqual(sourceBytes);
    expect(await readOriginal(destination)).toEqual(destinationBytes);
    await dm.saveAsDocument(destination, { applyTypography: false });
    expect(await readOriginal(destination)).toEqual(destinationBytes);
  });
  it('never replaces an original when later input differs', async () => {
    await initAstraRepo(doc, '<a/>', {}, data);
    expect(await captureOriginal(doc, Buffer.from('later'))).toMatchObject({ available: false, reason: 'capture-failed' });
    expect(await readOriginal(doc)).toEqual(data);
  });
  it('five concurrent captures elect one complete original without deleting shared bytes', async () => {
    await initAstraRepo(doc, '<a/>', {});
    const inputs = [data, Buffer.from('b'), data, Buffer.from('c'), Buffer.from('b')];
    const infos = await Promise.all(inputs.map((bytes) => captureOriginal(doc, bytes)));
    const result = await readOriginal(doc);
    expect(inputs.some((v) => v.equals(result))).toBe(true);
    expect(await getOriginalInfo(doc)).toMatchObject({ available: true, sha256: originalSha256(result) });
    for (let i = 0; i < inputs.length; i++) expect(infos[i].available).toBe(inputs[i].equals(result));
  });
  it('five same-byte captures all verify one original', async () => {
    await initAstraRepo(doc, '<a/>', {});
    expect((await Promise.all(Array.from({ length: 5 }, () => captureOriginal(doc, data)))).every((x) => x.available)).toBe(true);
  });
  it('reports absence, incomplete publication and corrupt bytes without returning a substitute', async () => {
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'predates-feature' });
    await initAstraRepo(doc, '<a/>', {}, data);
    const od = path.join(getRepoDir(doc), '.git/astra-original');
    const meta = await fsp.readFile(path.join(od, 'meta.json'));
    await fsp.unlink(path.join(od, 'meta.json'));
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'incomplete' });
    await fsp.writeFile(path.join(od, 'meta.json'), meta);
    await fsp.writeFile(path.join(od, `original-${originalSha256(data)}.docx`), Buffer.from('corrupted'));
    expect(await getOriginalInfo(doc)).toMatchObject({ available: false, reason: 'corrupt' });
    await expect(readOriginal(doc)).rejects.toMatchObject({ code: 'E_STATE' });
  });
  it('new documents have no imported original; relative memory labels are not applicable', async () => {
    await initAstraRepo(doc, '<a/>', {});
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'not-applicable' });
    expect(await getOriginalInfo('memory.docx')).toEqual({ available: false, reason: 'not-applicable' });
  });
  it('counts original bytes against quota', async () => {
    vi.stubEnv('ASTRA_HISTORY_MAX_SIDECAR_BYTES', '1');
    await initAstraRepo(doc, '<a/>', {}, data);
    expect(await getOriginalInfo(doc)).toEqual({ available: false, reason: 'too-large' });
  });
  it('keeps capture inside a relocated history root', async () => {
    vi.stubEnv('ASTRA_HISTORY_DIR', path.join(dir, 'histories'));
    await initAstraRepo(doc, '<a/>', {}, data);
    expect(getRepoDir(doc)).toContain('histories');
    expect(await readOriginal(doc)).toEqual(data);
  });
  it('refuses symlinked original data and metadata path traversal', async () => {
    await initAstraRepo(doc, '<a/>', {}, data);
    const od = path.join(getRepoDir(doc), '.git/astra-original');
    await fsp.writeFile(path.join(od, 'meta.json'), JSON.stringify({ version: 1, state: 'complete', sha256: '../../outside', bytes: 1, capturedAt: 'now' }));
    expect(await getOriginalInfo(doc)).toMatchObject({ reason: 'corrupt' });
  });
});
describe('publication helpers', () => {
  it('exclusively publishes one complete output in a five-writer race', async () => {
    const out = path.join(dir, 'out');
    const results = await Promise.allSettled(Array.from({ length: 5 }, () => writeFileExclusive(out, data)));
    expect(results.filter((x) => x.status === 'fulfilled')).toHaveLength(1);
    expect(await fsp.readFile(out)).toEqual(data);
    expect((await fsp.readdir(dir)).filter((n) => n.includes('astra-tmp'))).toEqual([]);
  });
  it('atomically replaces the file and preserves its bytes on a failed rename', async () => {
    const out = path.join(dir, 'out'); await fsp.writeFile(out, 'old');
    await replaceFileAtomic(out, data); expect(await fsp.readFile(out)).toEqual(data);
    vi.spyOn(fsp, 'rename').mockRejectedValueOnce(Object.assign(new Error('injected rename'), { code: 'EIO' }));
    await expect(replaceFileAtomic(out, Buffer.from('new'))).rejects.toMatchObject({ outcome: 'uncertain' });
    expect(await fsp.readFile(out)).toEqual(data);
  });
  it.skipIf(process.platform === 'win32')('refuses read-only and symlink targets', async () => {
    const out = path.join(dir, 'out'); await fsp.writeFile(out, 'old'); await fsp.chmod(out, 0o400);
    await expect(replaceFileAtomic(out, data)).rejects.toMatchObject({ outcome: 'not-replaced' });
    await fsp.chmod(out, 0o600);
    const link = path.join(dir, 'link'); await fsp.symlink(out, link);
    await expect(replaceFileAtomic(link, data)).rejects.toMatchObject({ outcome: 'not-replaced' });
    expect(await fsp.readFile(out, 'utf8')).toBe('old');
  });
  it('a partial temporary write never changes or creates the destination', async () => {
    const out = path.join(dir, 'out');
    await fsp.writeFile(out, 'old');
    const open = fsp.open.bind(fsp);
    const openSpy = vi.spyOn(fsp, 'open');
    openSpy.mockImplementation(async (...args: Parameters<typeof fsp.open>) => {
      const h = await open(...args);
      if (String(args[0]).includes('.astra-tmp-')) {
        vi.spyOn(h, 'writeFile').mockImplementationOnce(async () => {
          await h.write(Buffer.from('partial'));
          throw Object.assign(new Error('injected partial write'), { code: 'EIO' });
        });
      }
      return h;
    });
    await expect(replaceFileAtomic(out, data)).rejects.toMatchObject({ outcome: 'not-replaced' });
    expect(await fsp.readFile(out, 'utf8')).toBe('old');
    const newOut = path.join(dir, 'new');
    await expect(writeFileExclusive(newOut, data)).rejects.toThrow('injected partial write');
    await expect(fsp.stat(newOut)).rejects.toMatchObject({ code: 'ENOENT' });
  });
  it('a failure reported after rename is honestly uncertain and leaves complete new bytes', async () => {
    const out = path.join(dir, 'out'); await fsp.writeFile(out, 'old');
    const rename = fsp.rename.bind(fsp);
    vi.spyOn(fsp, 'rename').mockImplementationOnce(async (...args) => {
      await rename(...args);
      throw Object.assign(new Error('after rename'), { code: 'EIO' });
    });
    await expect(replaceFileAtomic(out, data)).rejects.toMatchObject({ outcome: 'uncertain' });
    expect(await fsp.readFile(out)).toEqual(data);
  });
  it('has no unsafe fallback where hard links are unsupported', async () => {
    const out = path.join(dir, 'out');
    vi.spyOn(fsp, 'link').mockRejectedValueOnce(Object.assign(new Error('unsupported'), { code: 'ENOTSUP' }));
    await expect(writeFileExclusive(out, data)).rejects.toThrow('does not support safe exclusive publication');
    await expect(fsp.stat(out)).rejects.toMatchObject({ code: 'ENOENT' });
  });
});
