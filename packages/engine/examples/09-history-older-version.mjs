// Read an older version of a document back as a .docx, and preview what a revert would change.
// This uses only the supported entry points: the root and `@attentionform/astra-engine/history`.
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { DocumentManager, DocxParser, composeDocument, validateDocxBuffer } from '@attentionform/astra-engine';
import { checkoutCommit, getHistory, trackedPredicateFor } from '@attentionform/astra-engine/history';
import { assertApplied, outputDir } from './_shared.mjs';

const file = join(outputDir('09'), 'versions.docx');
const seed = DocxParser.createNewDocument('blank');
assertApplied(await composeDocument(seed, '# Notes\n\nFirst version.\n'), 'compose');
writeFileSync(file, seed.rezip());

// Open it with history, then save two more versions.
const manager = new DocumentManager();
await manager.openDocument(file);
for (const text of ['Second paragraph.', 'Third paragraph.']) {
  assertApplied(await composeDocument(manager.getParser(), `${text}\n`), 'compose');
  await manager.saveDocument();
}

// The recorded states, newest first.
const commits = await getHistory(file);
console.log(`the history holds ${commits.length} recorded states`);
if (commits.length < 3) throw new Error('expected at least three recorded states');

// Preview a revert of one step: which package files differ between the file now and that state?
const steps = 1;
const target = await checkoutCommit(file, commits[steps].oid);
const live = new DocxParser(readFileSync(file)).capturePackageSnapshot();
const changed = [];
if (live.get('word/document.xml')?.toString('utf8') !== target.xml) changed.push('word/document.xml');
for (const [name, data] of target.extra) if (!live.get(name)?.equals(data)) changed.push(name);
console.log(`a revert of ${steps} step would change: ${changed.join(', ') || 'nothing'}`);
if (!changed.includes('word/document.xml')) throw new Error('the older state should differ from the current file');
if (!target.complete) console.log('note: this state was recorded by an older ASTra version; a restore would not be exact');

// Read that older state as a document, without touching the file on disk.
const parser = new DocxParser(readFileSync(file));
const snapshot = parser.capturePackageSnapshot();
const tracked = trackedPredicateFor(target);
for (const key of [...snapshot.keys()]) if (key !== 'word/document.xml' && tracked(key)) snapshot.delete(key);
for (const [name, data] of target.extra) snapshot.set(name, data);
snapshot.set('word/document.xml', Buffer.from(target.xml, 'utf8'));
parser.restorePackageSnapshot(snapshot, { registerCoreProperties: !target.complete });
const older = parser.rezip();

const olderText = new DocxParser(older).getStructuralSummary().nodes.map((n) => n.text).join(' ');
if (!olderText.includes('Second paragraph.') || olderText.includes('Third paragraph.')) {
  throw new Error(`the older version has the wrong content: ${olderText}`);
}
console.log('the older version holds the second paragraph and not the third');

const verdict = await validateDocxBuffer(older);
if (!verdict.ok) throw new Error('the older version is not valid');
console.log('and it is valid');
