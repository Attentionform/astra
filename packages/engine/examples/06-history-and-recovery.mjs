// Open a file with history, edit and save it, undo, and put the original bytes back.
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { DocumentManager, DocxParser, composeDocument, originalSha256, restoreOriginal } from '@attentionform/astra-engine';
import { assertApplied, outputDir } from './_shared.mjs';

const dir = outputDir('06');
const file = join(dir, 'history.docx');

// A file made outside the library. Opening it records its exact bytes as the original.
const source = DocxParser.createNewDocument('blank');
assertApplied(await composeDocument(source, '# Draft\n\nFirst version.\n'), 'compose');
writeFileSync(file, source.rezip());
const originalHash = originalSha256(readFileSync(file));

const manager = new DocumentManager();
await manager.openDocument(file);
const parser = manager.getParser();

// Edit and save. Every save adds an entry to the history.
assertApplied(await composeDocument(parser, 'A second paragraph.\n'), 'compose');
await manager.saveDocument();
const history = await manager.getVersionHistory();
console.log(`the history holds ${history.length} entries`);
if (history.length < 2) throw new Error('expected the opening and the save in the history');

// Undo goes back one step in the history.
const undone = await manager.undoLastAction();
if (!undone.success) throw new Error('undo failed');
console.log('undo restored the previous version');

// Put the original file back, byte for byte. This is a new history entry, so nothing is lost.
const restored = await restoreOriginal(file);
if (restored.sha256 !== originalHash || originalSha256(readFileSync(file)) !== originalHash) {
  throw new Error('the restored file is not the original');
}
console.log(`restored the original (${restored.sha256.slice(0, 12)}...)`);
