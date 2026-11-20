// Runs every example in this folder and stops at the first one that fails.
import { readdirSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const scripts = readdirSync(here).filter((f) => /^\d\d-.*\.mjs$/.test(f)).sort();
for (const script of scripts) {
  console.log(`\n== ${script}`);
  const r = spawnSync(process.execPath, [join(here, script)], { stdio: 'inherit' });
  if (r.status !== 0) {
    console.error(`${script} failed (exit ${r.status})`);
    process.exit(1);
  }
}
console.log(`\nall ${scripts.length} examples passed`);
