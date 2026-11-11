// Run the independent oracles (O1-O5) over every file an experiment wrote. Usage: node run/oracle_pass.mjs e1|c1|c2
import fs from 'node:fs';
import path from 'node:path';
import { RESULTS, WORK, readJsonl, jsonl, oracles } from '../lib/exp.mjs';

const exp = process.argv[2];
let files = [];
if (exp === 'e1') {
  for (const r of readJsonl('e1_edit_locality')) {
    if (r.kind === 'noop' && r.save?.ok) files.push(path.join(WORK, 'e1', r.N));
    if (r.kind === 'edit' && r.save?.ok) files.push(path.join(WORK, 'e1', r.E));
  }
} else if (exp === 'c1') files = readJsonl('c1_create').filter((r) => r.file).map((r) => r.file);
else if (exp === 'c2') files = readJsonl('c2_build').filter((r) => r.file).map((r) => r.file);
else throw new Error('unknown experiment');
files = files.filter((f) => fs.existsSync(f));
console.log(exp, 'files', files.length);
const out = jsonl(`oracles_${exp}`);
const CH = 60;
for (let i = 0; i < files.length; i += CH) {
  for (const r of oracles(files.slice(i, i + CH), { lo: true, workers: 10 })) out(r);
  console.log(`  ${Math.min(i + CH, files.length)}/${files.length}`);
}
const all = readJsonl(`oracles_${exp}`);
console.log('pass O1-O5', all.filter((r) => r.pass_O1_O5).length, '/', all.length);
