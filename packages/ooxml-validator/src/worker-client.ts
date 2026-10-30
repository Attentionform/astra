import { spawn, type ChildProcess, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { resolveEmbeddedBinary } from './resolve-binary.js';
import {
  encodeRequest,
  JsonLineDecoder,
  type WorkerRequestHeader,
  type WorkerResponse,
} from './protocol.js';
import type { ValidatorPoolOptions } from './types.js';

interface PendingRequest {
  id: number;
  header: WorkerRequestHeader;
  bytes?: Uint8Array | Buffer;
  timeoutMs?: number;
  resolve: (res: WorkerResponse) => void;
  reject: (err: Error) => void;
  retries: number;
  timeoutTimer?: NodeJS.Timeout;
}

export class WorkerClient {
  private binaryPath: string;
  private recycleAfter: number;
  private timeoutMs: number;
  private maxDocBytes: number;
  private maxErrors: number;

  private child: ChildProcess | null = null;
  private liveChildren: Set<ChildProcess> = new Set();
  private queue: PendingRequest[] = [];
  private current: PendingRequest | null = null;
  private nextId = 1;
  private requestsHandled = 0;
  private closed = false;
  private _spawnCount = 0;
  private _spawnedPids: number[] = [];

  constructor(opts: ValidatorPoolOptions & { binaryPath?: string } = {}) {
    this.binaryPath = opts.binaryPath ?? resolveEmbeddedBinary();
    this.recycleAfter = opts.recycleAfter ?? 500;
    this.timeoutMs = opts.timeoutMs ?? 20000;
    this.maxDocBytes = opts.maxDocBytes ?? 64 * 1024 * 1024;
    this.maxErrors = opts.maxErrors ?? 200;
  }

  public get pid(): number | undefined {
    return this.child?.pid;
  }

  public get spawnCount(): number {
    return this._spawnCount;
  }

  public get spawnedPids(): number[] {
    return [...this._spawnedPids];
  }

  private ensureWorker(): ChildProcess {
    if (this.child && !this.child.killed && this.child.exitCode === null) {
      return this.child;
    }

    let child: ChildProcessWithoutNullStreams;
    try {
      child = spawn(this.binaryPath, ['--worker'], {
        stdio: ['pipe', 'pipe', 'pipe'],
      });
    } catch (err) {
      // Some systems (Windows, for a file that is not an executable) throw instead of emitting 'error'.
      throw new Error(`Failed to spawn worker process "${this.binaryPath}": ${(err as Error).message}`);
    }
    this._spawnCount++;
    if (child.pid) {
      this._spawnedPids.push(child.pid);
    }
    this.liveChildren.add(child);

    const decoder = new JsonLineDecoder();

    child.on('error', (err) => {
      this.liveChildren.delete(child);
      if (this.child === child) {
        this.child = null;
        const msg = `Failed to spawn worker process "${this.binaryPath}": ${err.message}`;
        this.handleProcessError(new Error(msg));
      }
    });

    child.stdin.on('error', (err) => {
      // Handle EPIPE when worker exits prematurely
      if ((err as NodeJS.ErrnoException).code !== 'EPIPE') {
        console.error('[ooxml-validator-worker] stdin error:', err.message);
      }
    });

    child.stdout.on('data', (chunk: Buffer) => {
      if (this.child !== child) return;
      const responses = decoder.push(chunk);
      for (const resp of responses) {
        this.handleResponse(resp);
      }
    });

    child.stderr.on('data', (chunk: Buffer) => {
      // Logs go to stderr
      const msg = chunk.toString('utf8');
      if (process.env.DEBUG_OOXML_VALIDATOR) {
        process.stderr.write(`[worker-stderr] ${msg}`);
      }
    });

    child.on('exit', (code, signal) => {
      this.liveChildren.delete(child);
      if (this.child === child) {
        this.handleProcessExit(code, signal);
      }
    });

    this.child = child;
    this.requestsHandled = 0;
    return child;
  }

  private refChild() {
    if (this.child) {
      this.child.ref();
      (this.child.stdin as any)?.ref?.();
      (this.child.stdout as any)?.ref?.();
      (this.child.stderr as any)?.ref?.();
    }
  }

  private unrefChild() {
    if (this.child) {
      this.child.unref();
      (this.child.stdin as any)?.unref?.();
      (this.child.stdout as any)?.unref?.();
      (this.child.stderr as any)?.unref?.();
    }
  }

  public send(
    header: Omit<WorkerRequestHeader, 'id'>,
    bytes?: Uint8Array | Buffer,
    overrideTimeoutMs?: number,
  ): Promise<WorkerResponse> {
    if (this.closed) {
      return Promise.reject(new Error('Validator worker is closed'));
    }

    const id = this.nextId++;
    const fullHeader: WorkerRequestHeader = {
      ...header,
      id,
      maxDocBytes: header.maxDocBytes ?? this.maxDocBytes,
      maxErrors: header.maxErrors ?? this.maxErrors,
    };

    return new Promise<WorkerResponse>((resolve, reject) => {
      this.queue.push({
        id,
        header: fullHeader,
        bytes,
        timeoutMs: overrideTimeoutMs,
        resolve,
        reject,
        retries: 0,
      });

      this.processQueue();
    });
  }

  private processQueue() {
    if (this.current !== null || this.closed) {
      return;
    }

    if (this.queue.length === 0) {
      this.unrefChild();
      return;
    }

    const item = this.queue.shift()!;
    this.current = item;

    let child: ChildProcess;
    try {
      child = this.ensureWorker();
      this.refChild();
    } catch (err) {
      this.current = null;
      item.reject(err as Error);
      this.processQueue();
      return;
    }

    // Arm timeout timer
    const timeout = item.timeoutMs ?? this.timeoutMs;
    item.timeoutTimer = setTimeout(() => {
      this.handleTimeout();
    }, timeout);
    item.timeoutTimer.unref();

    const payload = encodeRequest(item.header, item.bytes);
    try {
      if (!child.stdin) {
        throw new Error('Worker stdin is not writable');
      }
      child.stdin.write(payload, (writeErr) => {
        if (writeErr && this.child === child) {
          this.handleWriteError(
            new Error(`Worker stdin write error for "${this.binaryPath}": ${writeErr.message}`),
          );
        }
      });
    } catch (err) {
      this.handleWriteError(err as Error);
    }
  }

  private handleResponse(resp: WorkerResponse) {
    if (!this.current) return;

    if (resp.id === this.current.id || resp.fatal) {
      const active = this.current;
      this.current = null;
      if (active.timeoutTimer) clearTimeout(active.timeoutTimer);

      if (resp.fatal) {
        this.killWorker();
        active.reject(new Error(resp.error || 'Fatal worker error'));
      } else {
        this.requestsHandled++;
        active.resolve(resp);

        if (this.requestsHandled >= this.recycleAfter) {
          this.recycleWorker();
        }
      }

      this.processQueue();
    }
  }

  private handleTimeout() {
    if (!this.current) return;
    const active = this.current;
    this.current = null;

    const timeout = active.timeoutMs ?? this.timeoutMs;
    this.killWorker();
    active.reject(
      new Error(`Validation request ${active.id} timed out after ${timeout}ms`),
    );

    this.processQueue();
  }

  private handleWriteError(err: Error) {
    if (!this.current) return;
    this.killWorker();
    this.retryOrReject(err);
  }

  private handleProcessError(err: Error) {
    this.child = null;
    if (this.current) {
      this.retryOrReject(err);
    }
  }

  private handleProcessExit(code: number | null, signal: string | null) {
    this.child = null;
    if (this.current) {
      const err = new Error(
        `Worker process exited mid-request (code: ${code}, signal: ${signal})`,
      );
      this.retryOrReject(err);
    }
  }

  private retryOrReject(err: Error) {
    const active = this.current;
    this.current = null;
    if (!active) return;

    if (active.timeoutTimer) clearTimeout(active.timeoutTimer);

    if (active.retries === 0) {
      // Retry once on crash/mid-request failure
      active.retries++;
      this.queue.unshift(active);
    } else {
      active.reject(err);
    }

    this.processQueue();
  }

  private recycleWorker() {
    if (this.child) {
      const oldChild = this.child;
      this.child = null;
      try {
        oldChild.stdin?.write(JSON.stringify({ id: 0, op: 'shutdown' }) + '\n');
        const killTimer = setTimeout(() => {
          if (!oldChild.killed && oldChild.exitCode === null) {
            oldChild.kill('SIGTERM');
          }
        }, 500);
        killTimer.unref();
      } catch {
        oldChild.kill('SIGTERM');
      }
    }
  }

  private killWorker() {
    if (this.child) {
      const c = this.child;
      this.child = null;
      this.liveChildren.delete(c);
      try {
        c.kill('SIGKILL');
      } catch {
        // ignore
      }
    }
  }

  public async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;

    if (this.current?.timeoutTimer) {
      clearTimeout(this.current.timeoutTimer);
    }
    if (this.current) {
      this.current.reject(new Error('Validator worker is closed'));
      this.current = null;
    }

    // Drain and reject queued requests
    while (this.queue.length > 0) {
      const item = this.queue.shift()!;
      if (item.timeoutTimer) clearTimeout(item.timeoutTimer);
      item.reject(new Error('Validator worker is closed'));
    }

    const targets = Array.from(this.liveChildren);
    this.child = null;

    if (targets.length === 0) {
      return;
    }

    await Promise.all(
      targets.map((c) => {
        return new Promise<void>((resolve) => {
          if (c.killed || c.exitCode !== null) {
            this.liveChildren.delete(c);
            resolve();
            return;
          }

          c.ref?.();

          let timer: NodeJS.Timeout | null = setTimeout(() => {
            timer = null;
            try {
              c.kill('SIGKILL');
            } catch {
              // ignore
            }
            this.liveChildren.delete(c);
            resolve();
          }, 2000);

          c.once('exit', () => {
            if (timer) {
              clearTimeout(timer);
              timer = null;
            }
            this.liveChildren.delete(c);
            resolve();
          });

          try {
            c.stdin?.write(JSON.stringify({ id: 0, op: 'shutdown' }) + '\n');
          } catch {
            c.kill('SIGTERM');
          }
        });
      }),
    );
  }
}
