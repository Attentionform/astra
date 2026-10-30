// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

import { describe, it, expect } from 'vitest';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

import { OFFICE_VERSIONS } from '../src/index.js';
import { MANY_ERRORS_DOCX, OFFICE2010_ONLY_DOCX, runCli, VALID_PPTX } from './helpers.js';

describe('cli', () => {
  it('cli accepts --officeVersion', async () => {
    const { stdout, code } = await runCli([VALID_PPTX, '--officeVersion', 'Office2019']);
    expect(code).toBe(0);
    expect(JSON.parse(stdout).ok).toBe(true);
  });

  it('--officeVersion changes the verdict, so it really reaches the validator', async () => {
    const as2007 = await runCli([OFFICE2010_ONLY_DOCX, '--officeVersion', 'Office2007']);
    expect(as2007.code).toBe(1);
    expect(JSON.parse(as2007.stdout).ok).toBe(false);

    const as2010 = await runCli([OFFICE2010_ONLY_DOCX, '--officeVersion', 'Office2010']);
    expect(as2010.code).toBe(0);
    expect(JSON.parse(as2010.stdout).ok).toBe(true);
  });

  it('every documented office version is accepted and validates as itself', async () => {
    for (const version of OFFICE_VERSIONS) {
      const { stdout, code } = await runCli([OFFICE2010_ONLY_DOCX, '--officeVersion', version]);
      expect(code).toBe(version === 'Office2007' ? 1 : 0);
      expect(JSON.parse(stdout).ok).toBe(version !== 'Office2007');
    }
  });

  it('cli rejects an unknown flag with a usage message instead of crashing', async () => {
    const { stderr, code } = await runCli([VALID_PPTX, '--office-version', 'Office2019']);
    expect(code).toBe(2);
    expect(stderr).toMatch(/Usage: ooxml-validator/);
    expect(stderr).not.toMatch(/at Object/);
  });

  it('cli rejects an invalid --officeVersion value', async () => {
    for (const bad of ['Office2018', 'office2007', 'OFFICE2007', '1', 'Office2007,Office2010', 'None']) {
      const { stdout, stderr, code } = await runCli([OFFICE2010_ONLY_DOCX, '--officeVersion', bad]);
      expect(code).toBe(2);
      expect(stderr).toMatch(/Invalid --officeVersion/);
      expect(stdout).toBe('');
    }
  });

  it('cli prints usage and exits 2 when no file is given', async () => {
    const { stderr, code } = await runCli([]);
    expect(code).toBe(2);
    expect(stderr).toMatch(/Usage: ooxml-validator/);
  });

  it('cli refuses more than one file instead of silently ignoring the rest', async () => {
    const { stdout, stderr, code } = await runCli([VALID_PPTX, OFFICE2010_ONLY_DOCX]);
    expect(code).toBe(2);
    expect(stderr).toMatch(/Expected exactly one file, got 2/);
    expect(stdout).toBe('');
  });

  it('--help goes to stdout and exits 0', async () => {
    const { stdout, code } = await runCli(['--help']);
    expect(code).toBe(0);
    expect(stdout).toMatch(/Usage: ooxml-validator/);
    expect(stdout).toMatch(/Exit codes:/);
  });

  it('cli exits 1 and reports errorType Exception for a file it cannot open', async () => {
    const { stdout, stderr, code } = await runCli(['/no/such/file.docx']);
    expect(code).toBe(1);
    expect(stderr).toBe('');
    const res = JSON.parse(stdout);
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Exception');
  });

  it('stdout carries only the JSON object; diagnostics go to stderr', async () => {
    const { stdout, stderr } = await runCli([OFFICE2010_ONLY_DOCX, '--officeVersion', 'Office2007']);
    expect(stderr).toBe('');
    expect(() => JSON.parse(stdout)).not.toThrow();
  });

  it('large output survives a pipe intact', async () => {
    const cli = fileURLToPath(new URL('../dist/cli.js', import.meta.url));

    const { stdout, code } = await new Promise<{ stdout: string; code: number | null }>((resolve, reject) => {
      const child = spawn(process.execPath, [cli, MANY_ERRORS_DOCX, '--officeVersion', 'Office2007'], {
        stdio: ['ignore', 'pipe', 'inherit'],
      });
      let out = '';
      child.stdout.setEncoding('utf8');
      child.stdout.on('data', (d) => {
        out += d;
      });
      child.on('error', reject);
      child.on('close', (code) => resolve({ stdout: out, code }));
    });

    expect(code).toBe(1);
    const parsed = JSON.parse(stdout);
    expect(parsed.errors.length).toBeGreaterThanOrEqual(400);
  });
});
