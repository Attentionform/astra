import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as os from 'os';
import * as path from 'path';
import { executeExport } from '../../packages/cli/src/export';

const IMAGE_DOC = path.resolve(__dirname, '../../test-gallery/T5_image_doc.docx');

async function makeTempDir(prefix: string): Promise<string> {
  return fsp.mkdtemp(path.join(os.tmpdir(), prefix));
}

describe('export post-mkdir containment', () => {
  it('refuses a symlinked asset dir and writes nothing outside', async () => {
    const dir = await makeTempDir('astra-export-race-');
    const outside = await makeTempDir('astra-export-race-out-');
    try {
      const outMd = path.join(dir, 'out.md');
      await fsp.mkdir(path.join(dir, 'media'), { recursive: true });
      await fsp.rm(path.join(dir, 'media'), { recursive: true, force: true });
      fs.symlinkSync(outside, path.join(dir, 'media'), 'dir');
      await expect(
        executeExport(IMAGE_DOC, 'markdown', outMd, { imageMode: 'file' })
      ).rejects.toThrow(/E_SANDBOX.*realpath/);
      expect(fs.existsSync(outMd)).toBe(false);
      expect(await fsp.readdir(outside)).toEqual([]);
    } finally {
      try {
        await fsp.rm(path.join(dir, 'media'), { force: true });
      } catch { /* ignore */ }
      await fsp.rm(dir, { recursive: true, force: true });
      await fsp.rm(outside, { recursive: true, force: true });
    }
  });

  it('refuses a symlinked destination file before writing', async () => {
    const dir = await makeTempDir('astra-export-dest-');
    const outside = await makeTempDir('astra-export-dest-out-');
    try {
      const target = path.join(outside, 'target.md');
      await fsp.writeFile(target, 'original');
      const link = path.join(dir, 'out.md');
      fs.symlinkSync(target, link);
      await expect(
        executeExport(IMAGE_DOC, 'markdown', link, { imageMode: 'embed' })
      ).rejects.toThrow(/E_SANDBOX/);
      expect(await fsp.readFile(target, 'utf-8')).toBe('original');
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
      await fsp.rm(outside, { recursive: true, force: true });
    }
  });

  it('normal file-mode export still writes assets inside the folder', async () => {
    const dir = await makeTempDir('astra-export-ok-');
    try {
      const outMd = path.join(dir, 'out.md');
      const res = await executeExport(IMAGE_DOC, 'markdown', outMd, { imageMode: 'file' });
      expect(res.ok).toBe(true);
      expect(fs.existsSync(outMd)).toBe(true);
      expect(fs.existsSync(path.join(dir, 'media', 'image-0001.png'))).toBe(true);
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });
});
