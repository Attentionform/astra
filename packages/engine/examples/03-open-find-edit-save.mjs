// Open a saved file, find text, edit it, validate and save.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { DocxParser, OperationEngine, composeDocument } from '@attentionform/astra-engine';
import { assertApplied, outputDir, writeValidated } from './_shared.mjs';

const dir = outputDir('03');

// Make a file to open, so that the example needs no input.
const source = DocxParser.createNewDocument('blank');
assertApplied(await composeDocument(source, '# Contract\n\nThe supplier is Acme Corp.\n'), 'compose');
await writeValidated(join(dir, 'contract.docx'), source.rezip());

// Open it.
const parser = new DocxParser(readFileSync(join(dir, 'contract.docx')));
const ast = await parser.getAST();

// Find the paragraph that holds the text.
const hit = ast.sections.flatMap((s) => s.nodes).find((n) => (n.metadata.text ?? '').includes('Acme Corp'));
if (!hit) throw new Error('paragraph not found');

// Edit it.
const engine = new OperationEngine(ast, parser);
const result = await engine.executeBatch({
  operations: [{ action: 'replace', target: hit.id, text: 'The supplier is Globex Inc.' }],
});
assertApplied(result, 'replace');

await writeValidated(join(dir, 'contract-edited.docx'), parser.rezip());
