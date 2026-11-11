// Fails (exit 1) unless the server under test can open a corpus document and save it. Run before any experiment,
// so a broken setup (for example a missing native validator) stops the job instead of producing rows full of errors.
import fs from 'node:fs';
import path from 'node:path';
import { connect } from '../lib/mcp.mjs';
import { freshDir, CORPUS } from '../lib/exp.mjs';

const dir = freshDir('smoke');
const file = path.join(dir, 'smoke.docx');
fs.copyFileSync(path.join(CORPUS, 'pd-prose-001p.docx'), file);
const s = await connect(dir);
try {
  const op = await s.call('open_document', { path: file });
  if (!op.json?.ok) throw new Error(`open_document failed: ${op.text.slice(0, 300)}`);
  const sv = await s.call('save_document', { sessionId: op.json.sessionId, message: 'smoke' });
  if (!sv.json?.ok) throw new Error(`save_document failed: ${sv.text.slice(0, 300)}`);
  console.log('smoke ok');
} catch (e) {
  console.error(String(e.message || e));
  process.exitCode = 1;
} finally {
  await s.close().catch(() => undefined);
}
