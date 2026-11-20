// A new document built only from operations. A heading is a paragraph with a named style applied.
import { join } from 'node:path';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';

// Any zip library works to look inside a .docx; this example uses adm-zip.
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

const bytes = parser.rezip();
// The style is applied by the namedStyle operation. An insert takes no style field, so check the result rather than trust it.
const documentXml = new AdmZip(bytes).readAsText('word/document.xml');
if (!/<w:pStyle w:val="Heading1"\/>[\s\S]*?Project notes/.test(documentXml)) throw new Error('the first paragraph is not a Heading 1');
console.log('the first paragraph is a Heading 1');

await writeValidated(join(outputDir('02'), 'notes.docx'), bytes);
