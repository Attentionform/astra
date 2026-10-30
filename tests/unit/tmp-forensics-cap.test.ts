import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as os from 'os';
import * as path from 'path';
import {
  ASTRA_TMPDIR_ENV,
  FORENSIC_MAX_BYTES,
  FORENSIC_MAX_COUNT,
  FORENSIC_TTL_MS,
  keepForForensics,
  pruneForensicArtifacts,
} from '@attentionform/astra-engine/tmp';

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

describe('forensic retention bounds', () => {
  it('exposes bounded retention constants', () => {
    expect(FORENSIC_MAX_COUNT).toBe(20);
    expect(FORENSIC_TTL_MS).toBe(24 * 60 * 60 * 1000);
    expect(FORENSIC_MAX_BYTES).toBeGreaterThan(0);
  });

  it('burst keeps stay under the count ceiling with newest retained', async () => {
    const base = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-forcap-'));
    process.env[ASTRA_TMPDIR_ENV] = base;
    try {
      const total = FORENSIC_MAX_COUNT + 10;
      const created: string[] = [];
      for (let i = 0; i < total; i++) {
        const p = path.join(base, `burst-${String(i).padStart(3, '0')}.bin`);
        await fsp.writeFile(p, Buffer.alloc(16, 0x41));
        const t = new Date(Date.now() + i * 10);
        await fsp.utimes(p, t, t);
        created.push(p);
        captureStderr(() => keepForForensics(p, 'burst probe'));
      }
      const survivors = pruneForensicArtifacts(base);
      expect(survivors.length).toBeLessThanOrEqual(FORENSIC_MAX_COUNT);
      const newest = created[created.length - 1];
      expect(survivors).toContain(path.resolve(newest));
      expect(fs.existsSync(newest)).toBe(true);
      const oldest = created[0];
      expect(fs.existsSync(oldest)).toBe(false);
    } finally {
      await fsp.rm(base, { recursive: true, force: true });
    }
  });

  it('expired entries are pruned after the retention window', async () => {
    const base = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-forttl-'));
    process.env[ASTRA_TMPDIR_ENV] = base;
    try {
      const old = path.join(base, 'old.bin');
      await fsp.writeFile(old, 'old');
      const past = new Date(Date.now() - FORENSIC_TTL_MS - 60_000);
      await fsp.utimes(old, past, past);
      captureStderr(() => keepForForensics(old, 'ttl probe'));
      expect(fs.existsSync(old)).toBe(false);
      expect(pruneForensicArtifacts(base)).not.toContain(path.resolve(old));
    } finally {
      await fsp.rm(base, { recursive: true, force: true });
    }
  });

  it('ceiling eviction emits a prune log line without absolute paths', async () => {
    const base = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-forlog-'));
    try {
      const oldest = path.join(base, 'oldest.bin');
      await fsp.writeFile(oldest, 'x');
      const past = new Date(Date.now() - 1000);
      await fsp.utimes(oldest, past, past);
      const ledger = JSON.stringify([path.resolve(oldest)]);
      await fsp.writeFile(path.join(base, '.astra-forensics.json'), ledger);
      const fresh: string[] = [];
      for (let i = 0; i < FORENSIC_MAX_COUNT; i++) {
        const p = path.join(base, `fresh-${i}.bin`);
        await fsp.writeFile(p, 'y');
        const t = new Date(Date.now() + (i + 1) * 20);
        await fsp.utimes(p, t, t);
        fresh.push(path.resolve(p));
      }
      const fullLedger = JSON.stringify([path.resolve(oldest), ...fresh]);
      await fsp.writeFile(path.join(base, '.astra-forensics.json'), fullLedger);
      let out = '';
      const orig = process.stderr.write;
      (process.stderr as unknown as { write: (chunk: unknown) => boolean }).write = (
        chunk: unknown
      ) => {
        out += String(chunk);
        return true;
      };
      try {
        pruneForensicArtifacts(base);
      } finally {
        process.stderr.write = orig;
      }
      expect(out).toContain('tmp forensics prune');
      expect(out).not.toContain(base);
    } finally {
      await fsp.rm(base, { recursive: true, force: true });
    }
  });
});
