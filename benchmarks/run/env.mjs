// Record the environment and the exact package versions/integrity hashes under test.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { ROOT, RESULTS, RELEASE } from '../lib/exp.mjs';

const lock = JSON.parse(fs.readFileSync(path.join(ROOT, 'package-lock.json'), 'utf8')).packages;
const pk = {};
for (const [k, v] of Object.entries(lock)) if (/node_modules\/@attentionform\//.test(k) || /node_modules\/(@modelcontextprotocol\/sdk|gpt-tokenizer|adm-zip)$/.test(k)) pk[k.replace('node_modules/', '')] = { version: v.version, integrity: v.integrity };
const sh = (c, a) => { try { return execFileSync(c, a, { encoding: 'utf8' }).trim().split('\n')[0]; } catch { return 'unavailable'; } };
const env = {
  release_under_test: RELEASE,
  measured_on: new Date().toISOString().slice(0, 10),
  os: `${os.type()} ${os.release()} ${os.arch()}`,
  wsl2: /microsoft/i.test(os.release()),
  cpu: os.cpus()[0].model, cpu_threads: os.cpus().length,
  ram_gb: +(os.totalmem() / 2 ** 30).toFixed(1),
  node: process.version,
  libreoffice: sh('soffice', ['--version']),
  python: sh('python3', ['--version']),
  python_docx: sh('python3', ['-c', 'import docx;print(docx.__version__)']),
  lxml: sh('python3', ['-c', 'from lxml import etree;print(etree.LXML_VERSION)']),
  pdfinfo: (execFileSync('pdfinfo', ['-v'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }) || 'poppler pdfinfo'),
  tokenizer: 'gpt-tokenizer (o200k_base), a proxy for any particular model',
  packages_under_test: pk,
  // Set when the server under test is a local build instead of the published package.
  local_build: process.env.ASTRA_SERVER_BIN
    ? { server_bin: process.env.ASTRA_SERVER_BIN, git_head: process.env.GITHUB_SHA_UNDER_TEST || process.env.GITHUB_SHA || sh('git', ['rev-parse', 'HEAD']), github_run_id: process.env.GITHUB_RUN_ID || null, runner: process.env.RUNNER_NAME || null }
    : null,
};
fs.writeFileSync(path.join(RESULTS, 'env.json'), JSON.stringify(env, null, 1));
console.log(JSON.stringify(env, null, 1).slice(0, 900));
