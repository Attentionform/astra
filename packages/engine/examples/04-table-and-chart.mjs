// A table and a chart.
import { join } from 'node:path';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { assertApplied, outputDir, writeValidated } from './_shared.mjs';

const parser = DocxParser.createNewDocument('blank');
const ast = await parser.getAST();
const engine = new OperationEngine(ast, parser);

const result = await engine.executeBatch({
  operations: [
    {
      action: 'insert',
      type: 'table',
      anchor: 'start',
      position: 'after',
      headers: ['Region', 'Revenue'],
      rows: [
        ['North', '120'],
        ['South', '95'],
      ],
    },
    {
      action: 'chart',
      operation: 'insert',
      anchor: 'end',
      position: 'after',
      chartType: 'column',
      title: 'Revenue by region',
      categories: ['North', 'South'],
      series: [{ name: 'Revenue', values: [120, 95] }],
    },
  ],
});
assertApplied(result, 'batch');

await writeValidated(join(outputDir('04'), 'table-and-chart.docx'), parser.rezip());
