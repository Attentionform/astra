/**
 * Runs an embedded Windows PowerShell script and reads its line protocol.
 *
 * Used by the Word renderer (word.ts) and the Windows PDF-to-image renderer (windows-pdf.ts).
 *
 * Rules (each one is a requirement, not a preference):
 *  - `powershell.exe` is started by absolute path under %SystemRoot%, never through PATH.
 *  - `-NoProfile -NonInteractive -EncodedCommand`: no script file is installed, no profile runs.
 *    No `-ExecutionPolicy` switch is passed and no policy is changed: whatever the organisation enforces
 *    (execution policy, AppLocker, constrained language, a blocked powershell.exe) is respected, and the
 *    failure is reported so the caller can use its documented fallback.
 *  - `windowsHide: true`: no console window ever appears.
 *  - Inputs travel in environment variables, never spliced into the script text.
 *  - On a timeout only OUR PowerShell child (and its children) is ended. Word is never force-killed.
 *  - One run at a time per queue.
 */

import { spawn } from 'child_process';
import * as path from 'path';

export interface PowerShellRunOptions {
  script: string;
  /** Extra environment variables for the child (inputs for the script). */
  env: Record<string, string>;
  timeoutMs: number;
  /** Tests only: run something else instead of powershell.exe. */
  powershellBin?: string;
  /** Tests only: arguments for `powershellBin` (replace the PowerShell switches). */
  powershellArgs?: string[];
  /** Called for every complete stdout line while the script runs. */
  onLine?: (line: string) => void;
}

export interface PowerShellResult {
  exitCode: number | null;
  lines: string[];
  stderr: string;
  timedOut: boolean;
  /** Set when the process could not be started (missing, blocked, access denied). */
  spawnError?: NodeJS.ErrnoException;
  /** PID of our PowerShell child. */
  pid?: number;
}

const MAX_STDOUT_BYTES = 1024 * 1024;
const MAX_STDERR_BYTES = 64 * 1024;
/** CreateProcess limits a command line to 32767 characters; stay well below it. */
export const MAX_ENCODED_COMMAND_CHARS = 28000;

export function windowsDirectory(env: NodeJS.ProcessEnv = process.env): string {
  return env.SystemRoot ?? env.SYSTEMROOT ?? env.windir ?? env.WINDIR ?? 'C:\\Windows';
}

/** Absolute path of Windows PowerShell 5.1 (present on every supported Windows). */
export function powershellPath(env: NodeJS.ProcessEnv = process.env): string {
  return path.win32.join(windowsDirectory(env), 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
}

export function systemToolPath(name: string, env: NodeJS.ProcessEnv = process.env): string {
  return path.win32.join(windowsDirectory(env), 'System32', name);
}

/** Base64 of the UTF-16LE script, as `-EncodedCommand` wants it. */
export function encodeCommand(script: string): string {
  return Buffer.from(script, 'utf16le').toString('base64');
}

/** The script prelude every ASTra script starts with: quiet, UTF-8 output, errors stop the script. */
export const POWERSHELL_PRELUDE = [
  "$ErrorActionPreference = 'Stop'",
  "$ProgressPreference = 'SilentlyContinue'",
  '[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false',
  'function Say($s) { [Console]::Out.WriteLine($s); [Console]::Out.Flush() }',
  "function Die($code, $e) { Say ('ERR {0} 0x{1:X8} {2}' -f $code, $e.Exception.HResult, ($e.Exception.Message -replace '\\s+', ' ')); exit $code }",
  "if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') { Say 'ERR 12 0x00000000 constrained language mode'; exit 12 }",
].join('\n');

/** Serialises runs: PowerShell and Word are started one at a time per queue. */
export class RunQueue {
  private tail: Promise<unknown> = Promise.resolve();
  run<T>(task: () => Promise<T>): Promise<T> {
    const next = this.tail.then(task, task);
    this.tail = next.catch(() => undefined);
    return next;
  }
}

function killOwnTree(pid: number | undefined, child: { kill(signal?: NodeJS.Signals): boolean }, windows: boolean): void {
  if (windows && pid !== undefined) {
    try {
      // Ends our PowerShell and anything it started. Word is not a child of PowerShell (COM starts it through
      // the service control manager), so this never touches it.
      spawn(systemToolPath('taskkill.exe'), ['/PID', String(pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore' }).on('error', () => undefined);
    } catch {
      /* fall through to the plain kill */
    }
  }
  try {
    child.kill('SIGKILL');
  } catch {
    /* already gone */
  }
}

export function runPowerShell(opts: PowerShellRunOptions): Promise<PowerShellResult> {
  return new Promise((resolve) => {
    const useReal = opts.powershellBin === undefined;
    const bin = opts.powershellBin ?? powershellPath();
    const args = opts.powershellArgs ?? ['-NoProfile', '-NonInteractive', '-EncodedCommand', encodeCommand(opts.script)];
    if (useReal) {
      const encoded = args[3] ?? '';
      if (encoded.length > MAX_ENCODED_COMMAND_CHARS) {
        resolve({ exitCode: null, lines: [], stderr: '', timedOut: false, spawnError: Object.assign(new Error('script too large for -EncodedCommand'), { code: 'E2BIG' }) as NodeJS.ErrnoException });
        return;
      }
    }
    const windows = process.platform === 'win32';
    const lines: string[] = [];
    let stderr = '';
    let stdoutBytes = 0;
    let carry = '';
    let settled = false;
    let timedOut = false;
    let spawnError: NodeJS.ErrnoException | undefined;

    let child: ReturnType<typeof spawn>;
    try {
      child = spawn(bin, args, {
        env: { ...process.env, ...opts.env },
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe'],
      });
    } catch (err: unknown) {
      resolve({ exitCode: null, lines, stderr, timedOut: false, spawnError: err as NodeJS.ErrnoException });
      return;
    }

    const finish = (exitCode: number | null): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (carry.length > 0) {
        lines.push(carry);
        opts.onLine?.(carry);
        carry = '';
      }
      resolve({ exitCode, lines, stderr, timedOut, spawnError, pid: child.pid });
    };

    const timer = setTimeout(() => {
      timedOut = true;
      killOwnTree(child.pid, child, windows);
      // `close` follows the kill; if it never does, do not hang the caller.
      setTimeout(() => finish(null), 2000).unref();
    }, Math.max(1000, opts.timeoutMs));

    child.stdout?.setEncoding('utf8');
    child.stdout?.on('data', (chunk: string) => {
      stdoutBytes += Buffer.byteLength(chunk);
      if (stdoutBytes > MAX_STDOUT_BYTES) return; // a runaway script cannot grow memory without bound
      carry += chunk;
      let nl: number;
      while ((nl = carry.indexOf('\n')) >= 0) {
        const line = carry.slice(0, nl).replace(/\r$/, '');
        carry = carry.slice(nl + 1);
        lines.push(line);
        opts.onLine?.(line);
      }
    });
    child.stderr?.setEncoding('utf8');
    child.stderr?.on('data', (chunk: string) => {
      if (stderr.length < MAX_STDERR_BYTES) stderr += chunk;
    });
    child.on('error', (err: NodeJS.ErrnoException) => {
      spawnError = err;
      finish(null);
    });
    child.on('close', (code) => finish(code));
  });
}

/** Printable, bounded text for an error message (no control characters, no long dumps). */
export function sanitiseDetail(text: string, max = 300): string {
  // eslint-disable-next-line no-control-regex
  const clean = text.replace(/[\u0000-\u001f\u007f]+/g, ' ').replace(/\s+/g, ' ').trim();
  return clean.length > max ? `${clean.slice(0, max)}…` : clean;
}
