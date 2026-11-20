// A comment, tracked edits, and accepting and rejecting them.
import { join } from 'node:path';
import { DocxParser, OperationEngine, composeDocument } from '@attentionform/astra-engine';
import { assertApplied, outputDir, writeValidated } from './_shared.mjs';

const dir = outputDir('05');
const parser = DocxParser.createNewDocument('blank');
assertApplied(await composeDocument(parser, 'The price is 100 dollars. Delivery takes 30 days.\n'), 'compose');

let ast = await parser.getAST();
const paragraph = ast.sections.flatMap((s) => s.nodes).find((n) => (n.metadata.text ?? '').includes('The price'));
let engine = new OperationEngine(ast, parser);

// A comment on the paragraph, and a tracked edit of its text.
assertApplied(
  await engine.executeBatch({
    operations: [
      { action: 'comment', operation: 'insert', target: paragraph.id, text: 'Please confirm the price.', author: 'Reviewer' },
      { action: 'replace', target: paragraph.id, text: 'The price is 120 dollars. Delivery takes 30 days.', trackChanges: true, author: 'Reviewer' },
    ],
  }),
  'comment and tracked edit'
);
const tracked = parser.readDocumentXml();
if (!tracked.includes('<w:ins ') || !tracked.includes('<w:del ')) throw new Error('the edit was not tracked');
await writeValidated(join(dir, 'with-changes.docx'), parser.rezip());

// Keep the file with the tracked changes, so that the same changes can be rejected too.
const withChanges = parser.rezip();

// Accept every tracked change.
ast = await parser.getAST();
engine = new OperationEngine(ast, parser);
assertApplied(await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] }), 'accept');
const accepted = parser.readDocumentXml();
if (accepted.includes('<w:ins ') || accepted.includes('<w:del ')) throw new Error('tracked changes remain after accepting');
const text = (await parser.getAST()).sections.flatMap((s) => s.nodes).map((n) => n.metadata.text ?? '').join('\n');
if (!text.includes('120 dollars') || text.includes('100 dollars')) throw new Error(`unexpected text after accepting: ${text}`);
await writeValidated(join(dir, 'accepted.docx'), parser.rezip());

// Reject every tracked change in a copy of the file that still holds them: the original text comes back.
const rejecting = new DocxParser(withChanges);
engine = new OperationEngine(await rejecting.getAST(), rejecting);
assertApplied(await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'rejectAll' }] }), 'reject');
const rejectedText = (await rejecting.getAST()).sections.flatMap((s) => s.nodes).map((n) => n.metadata.text ?? '').join('\n');
if (!rejectedText.includes('100 dollars') || rejectedText.includes('120 dollars')) throw new Error(`unexpected text after rejecting: ${rejectedText}`);
await writeValidated(join(dir, 'rejected.docx'), rejecting.rezip());
