/**
 * Regenerates upstream-golden.json from a reference validator binary.
 *
 *   UPSTREAM_VALIDATOR=/path/to/reference/ooxml-validator npx tsx packages/ooxml-validator/tests/parity/regenerate-golden.ts
 *
 * The reference is the unmodified one-shot program our native code was adapted from (xarsh/ooxml-validator
 * 0.4.0, commit 11e3e8c5430d054c8d9102e3e5f5e15c6697c7b2), built with `dotnet publish -r <rid>
 * --self-contained -p:PublishSingleFile=true`. Regenerate only when the corpus changes or when the
 * DocumentFormat.OpenXml version is raised on purpose; review the diff, since it is the behaviour change.
 */
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { buildCorpus, runCorpus } from './corpus.js';

const binary = process.env.UPSTREAM_VALIDATOR;
if (!binary) {
  console.error('Set UPSTREAM_VALIDATOR to the reference ooxml-validator binary.');
  process.exit(1);
}

const docs = buildCorpus();
const entries = await runCorpus(binary, docs);
const out = {
  _about:
    'Results of the unmodified upstream one-shot validator (xarsh/ooxml-validator 0.4.0, commit 11e3e8c5430d054c8d9102e3e5f5e15c6697c7b2, DocumentFormat.OpenXml 3.5.1) on the corpus built by corpus.ts. Descriptions are sha256 hashes.',
  entries,
};
const target = fileURLToPath(new URL('./upstream-golden.json', import.meta.url));
fs.writeFileSync(target, `${JSON.stringify(out, null, 1)}\n`);
console.log(`wrote ${Object.keys(entries).length} entries to ${target}`);
