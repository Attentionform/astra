import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  validateBuffer,
  selfTest,
  backendInfo,
} from '../src/index.js';
import { WorkerClient } from '../src/worker-client.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.resolve(__dirname, 'fixtures');
const VALID_DOCX = path.join(FIXTURES, 'valid.docx');
const INVALID_SCHEMA_DOCX = path.join(FIXTURES, 'invalid-schema.docx');
const MANY_ERRORS_DOCX = path.join(FIXTURES, 'many-errors.docx');

function getRssBytes(pid: number): number | null {
  if (process.platform !== 'linux') return null;
  try {
    const status = fs.readFileSync(`/proc/${pid}/status`, 'utf8');
    const m = /VmRSS:\s+(\d+)\s+kB/.exec(status);
    return m ? parseInt(m[1], 10) * 1024 : null;
  } catch {
    return null;
  }
}

describe('Worker & Buffer Validation (Milestone M3)', () => {
  it('selfTest succeeds', async () => {
    const res = await selfTest();
    expect(res.ok).toBe(true);
    expect(res.details).toContain('self-test passed');
  });

  it('backendInfo returns expected metadata, resolved package, and sha256', async () => {
    const info = await backendInfo();
    expect(info.sdkVersion).toBe('3.5.1');
    expect(info.workerSupported).toBe(true);
    expect(info.binarySha256).toMatch(/^[a-f0-9]{64}$/);
    expect(fs.existsSync(info.binaryPath)).toBe(true);
    expect(info.platformPackage).toBe('local:native/out');
  });

  it('validates a valid docx buffer without temp files', async () => {
    const bytes = fs.readFileSync(VALID_DOCX);
    const res = await validateBuffer(bytes);
    expect(res.ok).toBe(true);
    expect(res.errorCount).toBe(0);
    expect(res.errors).toEqual([]);
    expect(res.sdkVersion).toBe('3.5.1');
  });

  it('validates an invalid schema docx buffer', async () => {
    const bytes = fs.readFileSync(INVALID_SCHEMA_DOCX);
    const res = await validateBuffer(bytes);
    expect(res.ok).toBe(false);
    expect(res.errorCount).toBeGreaterThan(0);
    expect(res.errors[0]?.errorType).toBe('Schema');
  });

  // Fault test 1: kill the worker process mid-request, then the request succeeds on retry
  it('fault 1: kill worker mid-request succeeds on single retry', async () => {
    const client = new WorkerClient();
    try {
      const bytes = fs.readFileSync(VALID_DOCX);
      // Hook: schedule a kill right after send is called
      const promise = client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes);
      // Wait a tiny moment for process to spawn, then kill it
      setTimeout(() => {
        if (client.pid) {
          try {
            process.kill(client.pid, 'SIGKILL');
          } catch {
            // expected kill error
          }
        }
      }, 5);

      const res = await promise;
      expect(res.ok).toBe(true);
    } finally {
      await client.close();
    }
  });

  // Fault test 2: kill it twice, then the request rejects
  it('fault 2: kill worker twice causes request to reject (fail closed)', async () => {
    const client = new WorkerClient();
    try {
      const bytes = fs.readFileSync(VALID_DOCX);
      const promise = client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes);

      const killInterval = setInterval(() => {
        if (client.pid) {
          try {
            process.kill(client.pid, 'SIGKILL');
          } catch {
            // expected kill error
          }
        }
      }, 2);

      await expect(promise).rejects.toThrow();
      clearInterval(killInterval);
    } finally {
      await client.close();
    }
  });

  // Fault test 3: garbage bytes return ok:false with errorType: 'Exception'
  it('fault 3: garbage bytes return ok:false with errorType: Exception', async () => {
    const garbage = Buffer.from('NOT_A_VALID_ZIP_OR_DOCX_FILE_HEADER');
    const res = await validateBuffer(garbage);
    expect(res.ok).toBe(false);
    expect(res.errors.length).toBeGreaterThan(0);
    expect(res.errors[0]?.errorType).toBe('Exception');
  });

  // Fault test 4: document over maxDocBytes returns a clear error and next request still works
  it('fault 4: document over maxDocBytes returns clear error and next request works', async () => {
    const client = new WorkerClient({ maxDocBytes: 100 });
    try {
      const largeBytes = Buffer.alloc(200, 0x41);
      const errRes = await client.send({ op: 'validate', kind: 'docx', size: largeBytes.length }, largeBytes);
      expect(errRes.ok).toBe(false);
      expect(errRes.errors?.[0]?.description).toContain('exceeds maximum allowed');

      // Next request with valid doc should succeed
      const validBytes = fs.readFileSync(VALID_DOCX);
      const okRes = await client.send(
        { op: 'validate', kind: 'docx', size: validBytes.length, maxDocBytes: 100000 },
        validBytes,
      );
      expect(okRes.ok).toBe(true);
    } finally {
      await client.close();
    }
  });

  // Fault test 5: a document producing more than maxErrors errors returns truncated:true
  it('fault 5: maxErrors returns truncated:true', async () => {
    const bytes = fs.readFileSync(MANY_ERRORS_DOCX);
    const res = await validateBuffer(bytes, { officeVersion: 'Office2007', maxErrors: 1 });
    expect(res.ok).toBe(false);
    expect(res.errors.length).toBe(1);
    expect(res.truncated).toBe(true);
  });

  // Fault test 6: request exceeding timeoutMs rejects and next request works
  it('fault 6: request exceeding timeoutMs rejects and next request works on same client', async () => {
    const client = new WorkerClient({ timeoutMs: 20000 });
    try {
      const bytes = fs.readFileSync(VALID_DOCX);
      await expect(
        client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes, 1),
      ).rejects.toThrow(/timed out after 1ms/);

      // Next request on same client succeeds
      const okRes = await client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes);
      expect(okRes.ok).toBe(true);
    } finally {
      await client.close();
    }
  });

  it('queued burst across recycles spawns expected process count and cleans up', async () => {
    const client = new WorkerClient({ recycleAfter: 25 });
    try {
      const bytes = fs.readFileSync(VALID_DOCX);
      const promises: Promise<any>[] = [];
      for (let i = 0; i < 120; i++) {
        promises.push(client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes));
      }
      const results = await Promise.all(promises);
      for (const res of results) {
        expect(res.ok).toBe(true);
      }
      // ceil(120 / 25) = 5
      expect(client.spawnCount).toBe(5);

      const pids = client.spawnedPids;
      await client.close();

      // Ensure all spawned processes have exited
      for (const pid of pids) {
        expect(() => process.kill(pid, 0)).toThrow();
      }
    } finally {
      await client.close();
    }
  });

  it('timeout with work queued behind it recovers on same client with spawnCount 2', async () => {
    const client = new WorkerClient({ timeoutMs: 20000 });
    try {
      const bytes = fs.readFileSync(VALID_DOCX);
      // Request 1 has tiny timeout
      const p1 = client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes, 1);
      // Requests 2 and 3 queued behind it
      const p2 = client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes);
      const p3 = client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes);

      await expect(p1).rejects.toThrow(/timed out after 1ms/);
      const res2 = await p2;
      const res3 = await p3;
      expect(res2.ok).toBe(true);
      expect(res3.ok).toBe(true);
      expect(client.spawnCount).toBe(2);
    } finally {
      await client.close();
    }
  });

  it('unspawnable binary rejects send without killing process and subsequent send fails cleanly', async () => {
    const tmpNoExec = path.join(os.tmpdir(), `not-exec-${Date.now()}.bin`);
    fs.writeFileSync(tmpNoExec, '#!/bin/sh\nexit 0\n', { mode: 0o644 });
    const client = new WorkerClient({ binaryPath: tmpNoExec });
    try {
      const bytes = fs.readFileSync(VALID_DOCX);
      await expect(
        client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes),
      ).rejects.toThrow(tmpNoExec);

      // Subsequent send also rejects cleanly without hanging
      await expect(
        client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes),
      ).rejects.toThrow(tmpNoExec);
    } finally {
      await client.close();
      try {
        fs.unlinkSync(tmpNoExec);
      } catch {
        // ignore
      }
    }
  });

  it('close() on idle client resolves and rejects queued requests', async () => {
    const client = new WorkerClient();
    // Idle client close must resolve
    await expect(client.close()).resolves.toBeUndefined();

    // Now send on closed client rejects
    const bytes = fs.readFileSync(VALID_DOCX);
    await expect(
      client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes),
    ).rejects.toThrow(/closed/i);
  });

  // Fault test 8: no astra-val-* or other new file appears in temp directory during validateBuffer
  it('fault 8: zero temp files created during validateBuffer calls', async () => {
    const tmpDir = os.tmpdir();
    const filesBefore = new Set(fs.readdirSync(tmpDir).filter((f) => f.startsWith('astra-val-')));

    const bytes = fs.readFileSync(VALID_DOCX);
    for (let i = 0; i < 5; i++) {
      await validateBuffer(bytes);
    }

    const filesAfter = new Set(fs.readdirSync(tmpDir).filter((f) => f.startsWith('astra-val-')));
    const diff = [...filesAfter].filter((f) => !filesBefore.has(f));
    expect(diff).toEqual([]);
  });

  // Fault test 7: sequential validations of a small document complete, resident memory stays within 2x
  it('fault 7: sequential validations leak-check resident memory on same process', async () => {
    const client = new WorkerClient({ recycleAfter: 600 });
    try {
      const bytes = fs.readFileSync(VALID_DOCX);
      let rssAt50 = 0;
      let targetPid = 0;

      for (let i = 1; i <= 450; i++) {
        const res = await client.send({ op: 'validate', kind: 'docx', size: bytes.length }, bytes);
        expect(res.ok).toBe(true);

        if (i === 50 && client.pid) {
          targetPid = client.pid;
          rssAt50 = getRssBytes(client.pid) ?? 0;
        }
      }

      if (process.platform === 'linux' && targetPid > 0 && rssAt50 > 0) {
        expect(client.pid).toBe(targetPid);
        const rssFinal = getRssBytes(targetPid) ?? 0;
        // Final resident memory should stay within 2.0x of value at request 50
        expect(rssFinal).toBeLessThanOrEqual(rssAt50 * 2.0);
      }
    } finally {
      await client.close();
    }
  }, 60000);
});
