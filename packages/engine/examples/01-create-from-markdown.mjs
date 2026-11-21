// A new document from a Markdown string.
import { join } from 'node:path';
import { DocxParser, composeDocument } from '@attentionform/astra-engine';
import { assertApplied, outputDir, writeValidated } from './_shared.mjs';

const markdown = `# Quarterly review

Revenue grew in the **second** half, driven by two accounts.

## Results

| Region | Revenue |
|--------|---------|
| North  | 120     |
| South  | 95      |

- Renewals were on plan
- Two new accounts signed
`;

const parser = DocxParser.createNewDocument('blank');
const result = await composeDocument(parser, markdown);
assertApplied(result, 'compose');

await writeValidated(join(outputDir('01'), 'review.docx'), parser.rezip());
