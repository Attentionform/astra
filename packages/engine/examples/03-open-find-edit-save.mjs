// Open a saved file, find text, edit it, check it and save it.
import { readFileSync, writeFileSync } from 'node:fs';
import { DocxParser, OperationEngine, composeDocument, validateDocxBuffer } from '@attentionform/astra-engine';

// Make a file to open, so that this example needs no input.
const source = DocxParser.createNewDocument('blank');
await composeDocument(source, '# Contract\n\nThe supplier is Acme Corp.\n');
writeFileSync('contract.docx', source.rezip());

// Open it.
const parser = new DocxParser(readFileSync('contract.docx'));
const ast = await parser.getAST();

// Find the paragraph that holds the text.
const hit = ast.sections.flatMap((s) => s.nodes).find((n) => (n.metadata.text ?? '').includes('Acme Corp'));
if (!hit) throw new Error('paragraph not found');

// Edit it. A batch is atomic: when one operation fails, none is applied.
const engine = new OperationEngine(ast, parser);
const result = await engine.executeBatch({
  operations: [{ action: 'replace', target: hit.id, text: 'The supplier is Globex Inc.' }],
});
if (!result.success) throw new Error(`edit failed: ${JSON.stringify(result.errors)}`);

// Check the file before it is saved.
const bytes = parser.rezip();
const check = await validateDocxBuffer(bytes);
if (!check.ok) throw new Error(check.errors.map((e) => e.description).join('; '));

writeFileSync('contract-edited.docx', bytes);
console.log(`wrote contract-edited.docx (${bytes.length} bytes, valid)`);
