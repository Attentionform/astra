// Packs @attentionform/astra-engine, installs the tarball in a clean project, and checks the public API from there.
//
//   node scripts/check-packed-consumer.mjs            (after `npm run build --workspace=@attentionform/astra-engine`)
//
// The expected names live in packages/engine/api-surface.json. The checks are:
//   1. JavaScript, ESM and CommonJS: each entry point exports exactly the listed runtime values, nothing more.
//   2. TypeScript, NodeNext and Bundler resolution: every listed value and type imports, and `tsc --noEmit` passes.
//   3. Removed paths stay closed: a deep import such as `@attentionform/astra-engine/git` is refused, `./internal` opens.
//   4. Test hooks are not in the public option types.
//   5. A small program runs from the installed package.
import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const engineDir = join(root, 'packages', 'engine');
const surface = JSON.parse(readFileSync(join(engineDir, 'api-surface.json'), 'utf8'));
const entries = Object.keys(surface).filter((k) => k !== '_note');
const PKG = '@attentionform/astra-engine';
const spec = (entry) => (entry === '.' ? PKG : `${PKG}/${entry.slice(2)}`);

const run = (cmd, args, cwd) => execFileSync(cmd, args, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
const failures = [];
const fail = (msg) => {
  failures.push(msg);
  console.error(`FAIL ${msg}`);
};
const ok = (msg) => console.log(`ok   ${msg}`);

const work = mkdtempSync(join(tmpdir(), 'astra-packed-'));
console.log(`working in ${work}`);

// Pack and install.
const packOut = run('npm', ['pack', '--pack-destination', work, '--json'], engineDir);
const tarball = join(work, JSON.parse(packOut)[0].filename);
// The engine depends on the validator wrapper at the same version. A new version is not on npm until it is released, so the wrapper is
// packed from this checkout and installed with the engine; npm then uses it for the engine's dependency instead of asking the registry.
const wrapperOut = run('npm', ['pack', '--pack-destination', work, '--json'], join(root, 'packages', 'ooxml-validator'));
const wrapperTarball = join(work, JSON.parse(wrapperOut)[0].filename);
const consumer = join(work, 'consumer');
mkdirSync(consumer);
writeFileSync(join(consumer, 'package.json'), JSON.stringify({ name: 'consumer', version: '1.0.0', private: true }));
run('npm', ['install', '--no-audit', '--no-fund', tarball, wrapperTarball, 'typescript@5', '@types/node@22'], consumer);
ok(`installed ${tarball.split('/').pop()} with ${wrapperTarball.split('/').pop()}`);

// 1. JavaScript: exact runtime names, from ESM and from CommonJS.
const esmCheck = `
const surface = ${JSON.stringify(surface)};
const bad = [];
for (const entry of Object.keys(surface).filter((k) => k !== '_note')) {
  const spec = entry === '.' ? '${PKG}' : '${PKG}/' + entry.slice(2);
  const mod = await import(spec);
  const actual = Object.keys(mod).filter((k) => k !== 'default' && k !== 'module.exports' && k !== '__esModule').sort();
  const want = [...surface[entry].values].sort();
  const missing = want.filter((n) => !actual.includes(n));
  const extra = actual.filter((n) => !want.includes(n));
  if (missing.length || extra.length) bad.push(entry + ': missing [' + missing + '] extra [' + extra + ']');
  for (const n of want) if (mod[n] === undefined) bad.push(entry + ': ' + n + ' is undefined');
}
if (bad.length) { console.error(bad.join('\\n')); process.exit(1); }
`;
writeFileSync(join(consumer, 'check-esm.mjs'), esmCheck);
const cjsCheck = `
const surface = ${JSON.stringify(surface)};
const bad = [];
for (const entry of Object.keys(surface).filter((k) => k !== '_note')) {
  const mod = require(entry === '.' ? '${PKG}' : '${PKG}/' + entry.slice(2));
  const actual = Object.keys(mod).sort();
  const want = [...surface[entry].values].sort();
  const missing = want.filter((n) => !actual.includes(n));
  const extra = actual.filter((n) => !want.includes(n));
  if (missing.length || extra.length) bad.push(entry + ': missing [' + missing + '] extra [' + extra + ']');
}
if (bad.length) { console.error(bad.join('\\n')); process.exit(1); }
`;
writeFileSync(join(consumer, 'check-cjs.cjs'), cjsCheck);
for (const file of ['check-esm.mjs', 'check-cjs.cjs']) {
  try {
    run(process.execPath, [file], consumer);
    ok(`JavaScript ${file}: every entry exports exactly its listed values`);
  } catch (err) {
    fail(`JavaScript ${file}\n${err.stderr || err.message}`);
  }
}

// 3. Closed and open paths.
const pathCheck = `
const results = [];
for (const [name, shouldOpen] of [['git', false], ['parser', false], ['ooxml/validator-backend', false], ['internal', true], ['internal/git', true], ['package.json', true]]) {
  try { await import('${PKG}/' + name); results.push([name, true]); }
  catch (e) { results.push([name, false, e.code]); }
}
let bad = 0;
for (const [name, opened, code] of results) {
  const want = ['internal', 'internal/git'].includes(name);
  if (name === 'package.json') continue;
  if (opened !== want) { console.error(name + ': opened=' + opened + ' want=' + want + ' ' + (code || '')); bad++; }
}
process.exit(bad ? 1 : 0);
`;
writeFileSync(join(consumer, 'check-paths.mjs'), pathCheck);
try {
  run(process.execPath, ['check-paths.mjs'], consumer);
  ok('deep imports such as /git and /parser are refused; /internal and /internal/git open');
} catch (err) {
  fail(`paths\n${err.stderr || err.message}`);
}

// 2 and 4. TypeScript.
const lines = [];
entries.forEach((entry, i) => {
  const { values, types } = surface[entry];
  if (values.length) lines.push(`import { ${values.join(', ')} } from '${spec(entry)}';`);
  if (types.length) lines.push(`import type { ${types.join(', ')} } from '${spec(entry)}';`);
  // Use each value so that nothing is dropped as unused; a type import that fails to resolve is already an error.
  if (values.length) lines.push(`export const used${i} = [${values.join(', ')}];`);
});
writeFileSync(join(consumer, 'public-api.ts'), lines.join('\n') + '\n');
writeFileSync(
  join(consumer, 'hooks.ts'),
  `import type { VisualRenderOptions, PdfExportOptions } from '${PKG}';
// @ts-expect-error test hooks are not part of the public options
export const a: VisualRenderOptions = { testHooks: {} };
// @ts-expect-error test hooks are not part of the public options
export const b: PdfExportOptions = { testHooks: {} };
`,
);
writeFileSync(
  join(consumer, 'typed-batch.ts'),
  `import { DocxParser, OperationEngine, parseBatch } from '${PKG}';
import type { BatchPayload, OperationInput } from '${PKG}';
const parser = DocxParser.createNewDocument('blank');
export async function run(): Promise<boolean> {
  const engine = new OperationEngine(await parser.getAST(), parser);
  const ok: OperationInput[] = [
    { action: 'replace', target: 'p_001', text: 'new text' },
    { action: 'insert', type: 'paragraph', anchor: 'p_001', position: 'after', text: 'more' },
  ];
  const payload: BatchPayload = { operations: ok };
  const result = await engine.executeBatch(payload);
  const parsed = parseBatch(JSON.parse('{"operations":[]}'));
  return result.success && (parsed.ok ? parsed.batch.operations.length === 0 : parsed.errors.length > 0);
}
// @ts-expect-error an action that does not exist is a type error
export const bad: OperationInput = { action: 'no-such-action' };
// @ts-expect-error a replace needs a target
export const missing: OperationInput = { action: 'replace', text: 'x' };
`,
);
writeFileSync(
  join(consumer, 'internal.ts'),
  `import { initAstraRepo } from '${PKG}/internal';
import { AdmZip } from '${PKG}/internal';
export const names = [initAstraRepo, AdmZip];
`,
);
for (const [label, compilerOptions] of [
  ['NodeNext (ESM)', { module: 'NodeNext', moduleResolution: 'NodeNext' }],
  ['Bundler', { module: 'ESNext', moduleResolution: 'Bundler' }],
]) {
  writeFileSync(
    join(consumer, 'tsconfig.json'),
    JSON.stringify({
      compilerOptions: { ...compilerOptions, target: 'ES2022', strict: true, noEmit: true, skipLibCheck: false, types: ['node'] },
      files: ['public-api.ts', 'hooks.ts', 'typed-batch.ts', 'internal.ts'],
    }),
  );
  try {
    run(join(consumer, 'node_modules', '.bin', 'tsc'), ['-p', 'tsconfig.json'], consumer);
    ok(`TypeScript ${label}: every listed value and type imports; typed batches check; test hooks are not public; skipLibCheck off`);
  } catch (err) {
    fail(`TypeScript ${label}\n${err.stdout || ''}${err.stderr || err.message}`);
  }
}

// 5. A small program from the installed package.
writeFileSync(
  join(consumer, 'smoke.mjs'),
  `import { DocxParser, composeDocument, TEMPLATE_NAMES, getEngineErrorCode } from '${PKG}';
import { getHistory } from '${PKG}/history';
const parser = DocxParser.createNewDocument('memo');
const result = await composeDocument(parser, '# Title\\n\\nBody text.\\n');
if (!result.success) throw new Error('compose failed');
const bytes = parser.rezip();
if (!(bytes.length > 1000) || !TEMPLATE_NAMES.includes('memo')) throw new Error('unexpected output');
try { DocxParser.createNewDocument('no-such-template'); throw new Error('should have thrown'); }
catch (err) { if (getEngineErrorCode(err) !== 'E_VALIDATION') throw err; }
if (typeof getHistory !== 'function') throw new Error('history entry is empty');
`,
);
try {
  run(process.execPath, ['smoke.mjs'], consumer);
  ok('a document is created, composed and saved from the installed package');
} catch (err) {
  fail(`smoke\n${err.stderr || err.message}`);
}

console.log(readdirSync(work).length ? `\nchecked ${entries.length} entry points` : '');
if (failures.length) {
  console.error(`\n${failures.length} check(s) failed. Working folder kept at ${work}`);
  process.exit(1);
}
console.log('all packed-package checks passed');
