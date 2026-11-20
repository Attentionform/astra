// Every built-in template, in Letter and A4, and a compose onto one.
import { join } from 'node:path';
import { DocxParser, TEMPLATE_NAMES, composeDocument } from '@attentionform/astra-engine';
import { assertApplied, outputDir, writeValidated } from './_shared.mjs';

const dir = outputDir('07');

for (const name of TEMPLATE_NAMES) {
  for (const pageSize of ['letter', 'a4']) {
    const parser = DocxParser.createNewDocument(name, { pageSize });
    await writeValidated(join(dir, `${name}-${pageSize}.docx`), parser.rezip());
  }
}

// An untouched template has its placeholder body replaced; its styles, header and footer stay.
const report = DocxParser.createNewDocument('report');
const result = await composeDocument(report, '# Quarterly review\n\nRevenue grew in the second half.\n');
assertApplied(result, 'compose');
for (const warning of result.warnings) console.log(`warning: ${warning}`);
await writeValidated(join(dir, 'report-composed.docx'), report.rezip());
