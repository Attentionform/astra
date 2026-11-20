import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { validateDocxBuffer } from '@attentionform/astra-engine';

/** The folder an example writes into: ASTRA_EXAMPLE_OUT, or a new folder under the system temp directory. */
export function outputDir(name) {
  const base = process.env.ASTRA_EXAMPLE_OUT;
  if (base) {
    const dir = join(base, name);
    mkdirSync(dir, { recursive: true });
    return dir;
  }
  return mkdtempSync(join(tmpdir(), `astra-example-${name}-`));
}

/** Validates a .docx buffer against the Office schema, writes it, and stops the script when it is not valid. */
export async function writeValidated(path, buffer) {
  const result = await validateDocxBuffer(buffer);
  if (!result.ok) {
    throw new Error(`${path} is not valid: ${result.errors.map((e) => e.description).join('; ')}`);
  }
  writeFileSync(path, buffer);
  console.log(`wrote ${path} (${buffer.length} bytes, valid)`);
}

/** Stops the script with a clear message when a batch did not apply. */
export function assertApplied(result, label) {
  if (!result.success) {
    throw new Error(`${label} failed: ${JSON.stringify(result.errors)}`);
  }
}
