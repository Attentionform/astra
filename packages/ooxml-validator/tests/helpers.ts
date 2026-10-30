// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

import { execFile } from 'node:child_process';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { resolveEmbeddedBinary } from '../src/resolve-binary.js';

const execFileAsync = promisify(execFile);

const CLI = fileURLToPath(new URL('../dist/cli.js', import.meta.url));

export const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url));
export const VALID_PPTX = join(FIXTURES, 'minimal.pptx');
export const OFFICE2010_ONLY_DOCX = join(FIXTURES, 'office2010-only.docx');
export const MANY_ERRORS_DOCX = join(FIXTURES, 'many-errors.docx');
export const VALID_DOCX = join(FIXTURES, 'valid.docx');
export const INVALID_SCHEMA_DOCX = join(FIXTURES, 'invalid-schema.docx');
export const CORRUPT_ZIP_DOCX = join(FIXTURES, 'corrupt-zip.docx');

export function validatorBinary(): string {
  return process.env.OOXML_VALIDATOR_CLI || resolveEmbeddedBinary();
}

export async function runCli(args: string[], env: NodeJS.ProcessEnv = {}) {
  try {
    const { stdout, stderr } = await execFileAsync(process.execPath, [CLI, ...args], {
      env: { ...process.env, ...env },
    });
    return { stdout, stderr, code: 0 };
  } catch (err) {
    const e = err as { stdout?: string; stderr?: string; code?: number };
    return { stdout: e.stdout ?? '', stderr: e.stderr ?? '', code: e.code ?? 1 };
  }
}
