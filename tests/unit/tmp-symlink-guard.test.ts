import { describe, it, expect, afterEach } from 'vitest';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  ASTRA_TMPDIR_ENV,
  keepForForensics,
  logForensics,
  resolveAstraTmpBase,
  secureMkdtemp,
} from '@attentionform/astra-engine/tmp';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const prevBase = process.env[ASTRA_TMPDIR_ENV];

afterEach(() => {
  if (prevBase === undefined) delete process.env[ASTRA_TMPDIR_ENV];
  else process.env[ASTRA_TMPDIR_ENV] = prevBase;
});

function captureStderr(fn: () => void): string {
  const orig = process.stderr.write;
  let out = '';
  (process.stderr as unknown as { write: (chunk: unknown) => boolean }).write = (
    chunk: unknown
  ) => {
    out += String(chunk);
    return true;
  };
  try {
    fn();
  } finally {
    process.stderr.write = orig;
  }
  return out;
}

describe('tmp base refuses symlinked override', () => {
  it('throws E_SANDBOX with remedy and leaves the link target mode untouched', () => {
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-tmpguard-'));
    const victim = path.join(scratch, 'victim');
    fs.mkdirSync(victim);
    fs.chmodSync(victim, 0o755);
    const before = fs.statSync(victim).mode & 0o777;
    const beforeEntries = fs.readdirSync(victim);
    const link = path.join(scratch, 'linkbase');
    try {
      fs.symlinkSync(victim, link);
      process.env[ASTRA_TMPDIR_ENV] = link;
      let err: unknown;
      try {
        resolveAstraTmpBase();
      } catch (e) {
        err = e;
      }
      expect(err).toBeInstanceOf(Error);
      expect((err as { code?: string }).code).toBe('E_SANDBOX');
      expect((err as Error).message).toContain('E_SANDBOX');
      expect((err as { remedy?: string }).remedy).toMatch(/symlink/i);
      // Link target never chmodded and nothing planted inside it.
      expect(fs.statSync(victim).mode & 0o777).toBe(before);
      expect(fs.readdirSync(victim)).toEqual(beforeEntries);
      expect(fs.lstatSync(link).isSymbolicLink()).toBe(true);
    } finally {
      try {
        fs.unlinkSync(link);
      } catch {
        /* ignore */
      }
      try {
        fs.rmSync(scratch, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    }
  });

  it('secureMkdtemp under a symlinked override refuses the same way', () => {
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-tmpguard2-'));
    const victim = path.join(scratch, 'victim');
    fs.mkdirSync(victim);
    const link = path.join(scratch, 'linkbase');
    try {
      fs.symlinkSync(victim, link);
      process.env[ASTRA_TMPDIR_ENV] = link;
      expect(() => secureMkdtemp('probe-')).toThrowError(/E_SANDBOX/);
    } finally {
      try {
        fs.unlinkSync(link);
      } catch {
        /* ignore */
      }
      try {
        fs.rmSync(scratch, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    }
  });

  it('a real override dir still resolves and is repaired to 0700', () => {
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-tmpguard3-'));
    const real = path.join(scratch, 'realbase');
    try {
      process.env[ASTRA_TMPDIR_ENV] = real;
      expect(resolveAstraTmpBase()).toBe(path.resolve(real));
      expect(fs.statSync(real).mode & 0o777).toBe(0o700);
    } finally {
      try {
        fs.rmSync(scratch, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    }
  });
});

describe('forensics log carries no absolute path', () => {
  it('logForensics prints relative name plus hash, never the absolute path', () => {
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-tmpfor-'));
    const abs = path.join(scratch, 'kept-artifact.bin');
    try {
      fs.writeFileSync(abs, 'bytes');
      const out = captureStderr(() => logForensics(abs, 'probe reason'));
      expect(out).toContain(path.basename(abs));
      expect(out).toContain('sha256=');
      const expected = crypto
        .createHash('sha256')
        .update(abs)
        .digest('hex')
        .slice(0, 16);
      expect(out).toContain(expected);
      expect(out).not.toContain(abs);
      expect(out).not.toContain(scratch);
      expect(out).not.toContain(os.tmpdir());
    } finally {
      try {
        fs.rmSync(scratch, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    }
  });

  it('keepForForensics shares the same path-free shape', () => {
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-tmpfor2-'));
    const abs = path.join(scratch, 'kept-dir-probe');
    try {
      fs.mkdirSync(abs);
      const out = captureStderr(() => keepForForensics(abs, 'probe reason'));
      expect(out).toContain(path.basename(abs));
      expect(out).not.toContain(abs);
      expect(out).not.toContain(scratch);
    } finally {
      try {
        fs.rmSync(scratch, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    }
  });
});
