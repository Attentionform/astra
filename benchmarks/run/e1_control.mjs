// E1 control: plain parse-and-reserialize of every XML part (no product involved).
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { manifest, freshDir, jsonl, diffDocs, CORPUS, ROOT } from '../lib/exp.mjs';
const out = jsonl('e1_control');
const dir = freshDir('e1-control');
for (const d of manifest()) {
  const dst = path.join(dir, d.file);
  execFileSync('python3', [path.join(ROOT, 'lib/control_reserialize.py'), path.join(CORPUS, d.file), dst]);
  const x = diffDocs(path.join(CORPUS, d.file), dst);
  out({ doc: d.file, parts_changed: x.parts_changed.length, parts_changed_semantic: x.parts_changed_semantic.length, parts_total: x.parts_total_before, paragraphs_raw_changed: x.paragraphs_raw_changed_before, paragraphs_content_changed: x.paragraphs_content_changed_before, paragraphs: x.paragraphs_before, file_bytes_before: x.file_bytes_before, file_bytes_after: x.file_bytes_after });
}
console.log('control done');
