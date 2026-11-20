// Preview what a revert would change, and read an older version back as a .docx.
// This uses only the supported entry points: the root and `@attentionform/astra-engine/history`.
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { DocumentManager, DocxParser, OperationEngine, composeDocument, validateDocxBuffer } from '@attentionform/astra-engine';
import { checkoutCommit, getHistory, trackedPredicateFor } from '@attentionform/astra-engine/history';
import { assertApplied, outputDir } from './_shared.mjs';

// A 1 x 1 pixel PNG, to give a later state a package part that the earlier states do not have.
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/**
 * Compares the file on disk with a recorded state, over every part that history records.
 * `added`: the revert would put the part back. `removed`: the revert would delete it. `modified`: it would change it.
 */
async function previewRevert(file, commit) {
  const target = await checkoutCommit(file, commit.oid);
  const live = new DocxParser(readFileSync(file)).capturePackageSnapshot();
  const isTracked = trackedPredicateFor(target);
  const targetParts = new Map(target.extra);
  targetParts.set('word/document.xml', Buffer.from(target.xml, 'utf8'));
  const names = new Set([...targetParts.keys(), ...[...live.keys()].filter((n) => n === 'word/document.xml' || isTracked(n))]);
  const preview = { added: [], removed: [], modified: [], complete: target.complete };
  for (const name of [...names].sort()) {
    const now = live.get(name);
    const then = targetParts.get(name);
    if (now && !then) preview.removed.push(name);
    else if (!now && then) preview.added.push(name);
    else if (now && then && !now.equals(then)) preview.modified.push(name);
  }
  return { preview, target };
}

const file = join(outputDir('09'), 'versions.docx');
const seed = DocxParser.createNewDocument('blank');
assertApplied(await composeDocument(seed, '# Notes\n\nFirst version.\n'), 'compose');
writeFileSync(file, seed.rezip());

// Open it with history, then save three more versions. The last one adds an image.
const manager = new DocumentManager();
await manager.openDocument(file);
for (const text of ['Second paragraph.', 'Third paragraph.']) {
  assertApplied(await composeDocument(manager.getParser(), `${text}\n`), 'compose');
  await manager.saveDocument();
}
const parser = manager.getParser();
const ast = await parser.getAST();
const last = ast.sections.flatMap((s) => s.nodes).at(-1);
const added = await new OperationEngine(ast, parser).executeBatch({
  operations: [{ action: 'image', operation: 'insert', anchor: last.id, fileData: PNG, width: 914400, height: 914400, altText: 'one pixel' }],
});
assertApplied(added, 'image');
await manager.saveDocument();

// The recorded states, newest first: with image, third paragraph, second paragraph, first version.
const commits = await getHistory(file);
console.log(`the history holds ${commits.length} recorded states`);
if (commits.length < 4) throw new Error('expected at least four recorded states');

// Preview a revert of one step. The newest state has an image that the one before it does not.
const { preview } = await previewRevert(file, commits[1]);
console.log(`a revert of 1 step would remove ${preview.removed.join(', ') || 'nothing'}`);
console.log(`  and change ${preview.modified.join(', ') || 'nothing'}`);
if (!preview.removed.some((n) => n.startsWith('word/media/'))) throw new Error('the preview should report the image that the revert removes');
if (!preview.modified.includes('word/document.xml')) throw new Error('the preview should report word/document.xml');

// Read an older state as a document, without touching the file on disk.
const older = await (async () => {
  const { target } = await previewRevert(file, commits[2]);
  const reader = new DocxParser(readFileSync(file));
  const snapshot = reader.capturePackageSnapshot();
  const isTracked = trackedPredicateFor(target);
  for (const key of [...snapshot.keys()]) if (key !== 'word/document.xml' && isTracked(key)) snapshot.delete(key);
  for (const [name, data] of target.extra) snapshot.set(name, data);
  snapshot.set('word/document.xml', Buffer.from(target.xml, 'utf8'));
  reader.restorePackageSnapshot(snapshot, { registerCoreProperties: !target.complete });
  return reader.rezip();
})();

const olderText = new DocxParser(older).getStructuralSummary().nodes.map((n) => n.text).join(' ');
if (!olderText.includes('Second paragraph.') || olderText.includes('Third paragraph.')) {
  throw new Error(`the older version has the wrong content: ${olderText}`);
}
const olderParts = Object.keys(Object.fromEntries(new DocxParser(older).capturePackageSnapshot()));
if (olderParts.some((n) => n.startsWith('word/media/'))) throw new Error('the older version should not carry the image');
console.log('the older version holds the second paragraph, not the third, and no image');

const verdict = await validateDocxBuffer(older);
if (!verdict.ok) throw new Error('the older version is not valid');
console.log('and it is valid');
