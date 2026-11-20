// Runs every example in this folder and stops at the first one that fails.
import { mkdirSync, mkdtempSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
// A script that writes a plain file name writes it into its own folder under `out`; the others choose their own folder.
const out = process.env.ASTRA_EXAMPLE_OUT ?? mkdtempSync(join(tmpdir(), 'astra-examples-'));
const scripts = readdirSync(here).filter((f) => /^\d\d-.*\.mjs$/.test(f)).sort();
for (const script of scripts) {
  console.log(`\n== ${script}`);
  const cwd = join(out, script.replace(/\.mjs$/, ''));
  mkdirSync(cwd, { recursive: true });
  const r = spawnSync(process.execPath, [join(here, script)], { stdio: 'inherit', cwd });
  if (r.status !== 0) {
    console.error(`${script} failed (exit ${r.status})`);
    process.exit(1);
  }
}
console.log(`\nall ${scripts.length} examples passed`);
