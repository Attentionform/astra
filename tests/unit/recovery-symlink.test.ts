import { describe, it, expect } from 'vitest';
import fsp from 'fs/promises';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  discardRecoveredDraft,
  listDraftDocxFiles,
  getDraftRecoveryInfo,
  sweepOrphanDrafts,
} from '@attentionform/astra-engine';

async function mkDir(prefix: string): Promise<string> {
  return fsp.mkdtemp(path.join(os.tmpdir(), prefix));
}

async function writeDocx(p: string, bytes = 64): Promise<void> {
  await fsp.mkdir(path.dirname(p), { recursive: true });
  await fsp.writeFile(p, Buffer.alloc(bytes, 0x41));
}

describe('draft symlink refusal', () => {
  it('list skips symlinked drafts', async () => {
    const drafts = await mkDir('astra-rec-list-');
    const outside = await mkDir('astra-rec-list-out-');
    try {
      const real = path.join(drafts, 'real.docx');
      await writeDocx(real);
      const target = path.join(outside, 'secret.docx');
      await writeDocx(target);
      const link = path.join(drafts, 'link.docx');
      try {
        fs.symlinkSync(target, link);
      } catch {
        return;
      }
      const files = await listDraftDocxFiles(drafts);
      expect(files).toContain(real);
      expect(files).not.toContain(link);
      const info = await getDraftRecoveryInfo(link);
      expect(info.commitCount).toBe(0);
    } finally {
      await fsp.rm(drafts, { recursive: true, force: true });
      await fsp.rm(outside, { recursive: true, force: true });
    }
  });

  it('sweep never deletes a symlinked draft and leaves the target intact', async () => {
    const drafts = await mkDir('astra-rec-sweep-');
    const outside = await mkDir('astra-rec-sweep-out-');
    try {
      const target = path.join(outside, 'secret.docx');
      await writeDocx(target, 128);
      const link = path.join(drafts, 'evil.docx');
      try {
        fs.symlinkSync(target, link);
      } catch {
        return;
      }
      const oldMs = Date.now() - 30 * 24 * 60 * 60 * 1000;
      try {
        await fsp.lutimes(link, new Date(oldMs), new Date(oldMs));
      } catch {
        await fsp.utimes(target, new Date(oldMs), new Date(oldMs));
      }
      const res = await sweepOrphanDrafts(drafts, { nowMs: Date.now() });
      expect(res.deleted).not.toContain(path.resolve(link));
      expect(fs.existsSync(target)).toBe(true);
      expect((await fsp.readFile(target)).length).toBe(128);
      expect(fs.lstatSync(link).isSymbolicLink()).toBe(true);
    } finally {
      await fsp.rm(drafts, { recursive: true, force: true });
      await fsp.rm(outside, { recursive: true, force: true });
    }
  });

  it('discard refuses a symlinked draft', async () => {
    const drafts = await mkDir('astra-rec-discard-');
    const outside = await mkDir('astra-rec-discard-out-');
    try {
      const target = path.join(outside, 'secret.docx');
      await writeDocx(target, 64);
      const link = path.join(drafts, 'evil.docx');
      try {
        fs.symlinkSync(target, link);
      } catch {
        return;
      }
      const res = await discardRecoveredDraft(drafts, link);
      expect(res.discarded).toBe(false);
      expect(fs.existsSync(target)).toBe(true);
      expect(fs.lstatSync(link).isSymbolicLink()).toBe(true);
    } finally {
      await fsp.rm(drafts, { recursive: true, force: true });
      await fsp.rm(outside, { recursive: true, force: true });
    }
  });
});
