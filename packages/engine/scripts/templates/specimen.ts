/**
 * Style specimen: one document per candidate theme showing every style, so the design direction can be judged
 * in Word. Not a shipped template. Output folder is the first argument.
 */
import * as fs from 'fs';
import * as path from 'path';
import { CANDIDATES } from './themes';
import { contrastChecks } from './design';
import { A4, LETTER, buildPackage, cell, p, para, ph, phPara, pageXofY, run, runningTitle, table } from './lib';

const out = process.argv[2];
if (!out) throw new Error('usage: specimen <outDir>');
fs.mkdirSync(out, { recursive: true });

for (const theme of CANDIDATES) {
  const body = [
    para('Document title goes here', 'Title'),
    para('A subtitle that frames the document in one line', 'Subtitle'),
    p([run('Prepared for '), ph('Client name'), run(' by '), ph('Author'), run(' on '), ph('Date')]),
    para('Executive summary', 'Heading1'),
    para('Body text sets the reading rhythm of the whole document. It uses a comfortable line height, a restrained measure and enough space after each paragraph that a page never reads as a wall of text. The quick brown fox jumps over the lazy dog, then reads the report again.', 'BodyText'),
    p([run('Inline styles: '), run('strong', { style: 'Strong' }), run(', '), run('emphasis', { style: 'Emphasis' }), run(', '), run('code', { style: 'CodeChar' }), run(', and a '), run('link', { style: 'Hyperlink' }), run('.')], { style: 'BodyText' }),
    para('Findings', 'Heading2'),
    para('Second-level heading above a list.', 'BodyText'),
    para('First bullet point', 'ListBullet'),
    para('Second bullet point', 'ListBullet'),
    para('Nested bullet point', 'ListBullet2'),
    para('First numbered step', 'ListNumber'),
    para('Second numbered step', 'ListNumber'),
    para('Third-level heading', 'Heading3'),
    para('A short quotation set apart from the text.', 'Quote'),
    para('A callout draws the eye to one thing the reader must not miss.', 'Callout'),
    para('Table 1. A clean table', 'Caption'),
    table({
      style: 'AstraTable',
      header: true,
      widths: [3600, 1920, 1920, 1920],
      cantSplit: true,
      rows: [
        [cell(run('Item')), cell(run('Quantity'), { align: 'right' }), cell(run('Rate'), { align: 'right' }), cell(run('Amount'), { align: 'right' })],
        [cell(ph('Description')), cell(ph('Qty'), { align: 'right' }), cell(ph('Rate'), { align: 'right' }), cell(ph('Amount'), { align: 'right' })],
        [cell(ph('Description')), cell(ph('Qty'), { align: 'right' }), cell(ph('Rate'), { align: 'right' }), cell(ph('Amount'), { align: 'right' })],
      ],
    }),
    para('', 'NoSpacing'),
    para('Table 2. Banded', 'Caption'),
    table({
      style: 'AstraTableBanded',
      header: true,
      look: { banded: true },
      widths: [3600, 5760],
      rows: [
        [cell(run('Field')), cell(run('Detail'))],
        [cell(run('One')), cell(ph('Detail'))],
        [cell(run('Two')), cell(ph('Detail'))],
        [cell(run('Three')), cell(ph('Detail'))],
      ],
    }),
    para('', 'NoSpacing'),
    para('Heading four', 'Heading4'),
    para('GET /v1/items?limit=20', 'Code'),
    para('Authorization: Bearer <token>', 'Code'),
    para('', 'BodyText'),
    para('Placeholders look like this: ', 'BodyText'),
    phPara('Replace this text'),
  ].join('');
  const buf = buildPackage({
    theme,
    template: 'specimen',
    designVersion: 1,
    body,
    page: theme.name === 'Brick' ? A4 : LETTER,
    parts: { defaultHeader: runningTitle('Specimen'), defaultFooter: pageXofY(run(`${theme.name} candidate`)) },
    title: `${theme.name} specimen`,
  });
  fs.writeFileSync(path.join(out, `specimen-${theme.name.toLowerCase()}.docx`), buf);
  const checks = contrastChecks(theme);
  console.log(theme.name, checks.map((c) => `${c.pair}: ${c.ratio}${c.ok ? '' : ' FAIL'}`).join(' | '));
}
