import { StringDecoder } from 'node:string_decoder';
import type { OfficeVersion, ValidationError } from './types.js';

export interface WorkerRequestHeader {
  id: number;
  op?: 'validate' | 'ping' | 'version' | 'shutdown' | string;
  kind?: 'docx' | 'pptx' | 'xlsx' | string;
  office?: OfficeVersion | string;
  size?: number;
  maxErrors?: number;
  maxDocBytes?: number;
}

export interface WorkerResponse {
  id: number | null;
  ok: boolean;
  pong?: boolean;
  sdkVersion?: string;
  office?: string;
  durationMs?: number;
  errorCount?: number;
  truncated?: boolean;
  fatal?: boolean;
  error?: string;
  errors?: ValidationError[];
}

export function encodeRequest(header: WorkerRequestHeader, bytes?: Uint8Array | Buffer): Buffer {
  const json = JSON.stringify(header);
  const headerBuf = Buffer.from(json + '\n', 'utf8');

  if (!bytes || bytes.length === 0) {
    return headerBuf;
  }

  const bodyBuf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  return Buffer.concat([headerBuf, bodyBuf]);
}

export class JsonLineDecoder {
  private decoder = new StringDecoder('utf8');
  private buffer = '';

  public push(chunk: Buffer | Uint8Array): WorkerResponse[] {
    this.buffer += this.decoder.write(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    const lines = this.buffer.split('\n');
    // Keep whatever is after the last newline in this.buffer
    this.buffer = lines.pop() ?? '';

    const results: WorkerResponse[] = [];
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      try {
        const parsed = JSON.parse(trimmed) as WorkerResponse;
        results.push(parsed);
      } catch (err) {
        results.push({
          id: null,
          ok: false,
          fatal: true,
          error: `Unparsable response line: ${trimmed}. Error: ${(err as Error).message}`,
        });
      }
    }
    return results;
  }

  public end(): WorkerResponse[] {
    this.buffer += this.decoder.end();
    const remaining = this.buffer.trim();
    this.buffer = '';
    if (!remaining) return [];
    try {
      return [JSON.parse(remaining) as WorkerResponse];
    } catch (err) {
      return [
        {
          id: null,
          ok: false,
          fatal: true,
          error: `Unparsable trailing response line: ${remaining}. Error: ${(err as Error).message}`,
        },
      ];
    }
  }
}
