const fs = require('node:fs/promises');
const path = require('node:path');
const engine = require('@attentionform/astra-engine');

const hash = (bytes) => engine.originalSha256(bytes);
const extras = (parser) => Object.fromEntries([...parser.capturePackageSnapshot()].filter(([name]) => engine.isTrackedPackageFile(name)));

async function setup(file) {
  const original = engine.createBlankDocxBuffer('blank');
  const parser = new engine.DocxParser(original);
  await parser.getAST();
  const originalXml = parser.readDocumentXml();
  await engine.initAstraRepo(file, originalXml, extras(parser), original);
  const editedXml = originalXml.replace('Start typing your content here...', 'Locked restore test edit');
  if (editedXml === originalXml) throw new Error('Could not edit the blank document fixture.');
  parser.reloadDocumentXml(editedXml);
  const edited = parser.rezip();
  await fs.writeFile(file, edited);
  await engine.commitChange(file, parser.readDocumentXml(), 'edit for locked restore test', extras(parser));
  const history = await engine.getHistory(file);
  process.stdout.write(JSON.stringify({ ok: true, originalHash: hash(original), editedHash: hash(edited), commits: history.length }));
}

async function inspect(file) {
  const bytes = await fs.readFile(file);
  const history = await engine.getHistory(file);
  process.stdout.write(JSON.stringify({ ok: true, hash: hash(bytes), commits: history.length }));
}

async function restore(file) {
  const result = await engine.restoreOriginal(file);
  process.stdout.write(JSON.stringify({ ok: true, result }));
}

async function main() {
  const [mode, file] = process.argv.slice(2);
  try {
    if (mode === 'setup') await setup(path.resolve(file));
    else if (mode === 'inspect') await inspect(path.resolve(file));
    else if (mode === 'restore') await restore(path.resolve(file));
    else throw new Error(`Unknown worker mode: ${mode}`);
  } catch (error) {
    process.stdout.write(JSON.stringify({ ok: false, code: error && error.code, message: error instanceof Error ? error.message : String(error) }));
    process.exitCode = 1;
  }
}

main();
