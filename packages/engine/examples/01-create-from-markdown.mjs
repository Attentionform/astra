// A new document from a Markdown string, checked against the Office schema and saved.
import { writeFileSync } from 'node:fs';
import { DocxParser, composeDocument, validateDocxBuffer } from '@attentionform/astra-engine';

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

// An empty document, then the Markdown as headings, paragraphs, a table and a list.
const parser = DocxParser.createNewDocument('blank');
const result = await composeDocument(parser, markdown);
if (!result.success) throw new Error(`compose failed: ${JSON.stringify(result.errors)}`);

// Check the file before it is saved.
const bytes = parser.rezip();
const check = await validateDocxBuffer(bytes);
if (!check.ok) throw new Error(check.errors.map((e) => e.description).join('; '));

writeFileSync('review.docx', bytes);
console.log(`wrote review.docx (${bytes.length} bytes, valid)`);
