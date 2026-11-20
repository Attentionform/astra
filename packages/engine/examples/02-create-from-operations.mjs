// A new document built only from operations. A heading is a paragraph with a named style applied.
import { join } from 'node:path';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { assertApplied, outputDir, writeValidated } from './_shared.mjs';

const parser = DocxParser.createNewDocument('blank');
const ast = await parser.getAST();
const engine = new OperationEngine(ast, parser);

const result = await engine.executeBatch({
  operations: [
    { action: 'insert', type: 'paragraph', anchor: 'start', position: 'after', text: 'Project notes', id: 'notes_title' },
    { action: 'namedStyle', operation: 'apply', styleName: 'Heading 1', target: 'notes_title' },
    { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Built from operations, with no Markdown.' },
    { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Every operation in the batch applies, or none does.' },
  ],
});
assertApplied(result, 'batch');
console.log(`applied ${result.appliedCount} operations`);

await writeValidated(join(outputDir('02'), 'notes.docx'), parser.rezip());
