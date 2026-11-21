import { DESIGN, DESIGN_VERSION } from './themes';
import { A4, LETTER, buildPackage, coverRuleIndent, textWidth, cell, field, marker, p, pageBreak, pageXofY, para, ph, phPara, run, runningTitle, sectionBreak, table, type PageSetup } from './lib';

/** The report: cover, contents, then a body region that holds labelled placeholders. */
export function report(page: PageSetup = LETTER): Buffer {
  const cover = [
    p('', { style: 'NoSpacing', raw: `<w:pBdr><w:top w:val="single" w:sz="36" w:space="0" w:color="${DESIGN.colors.accent}"/></w:pBdr><w:spacing w:before="2400" w:after="240"/><w:ind w:right="${coverRuleIndent(page)}"/>` }),
    // Slot: the title placeholder, wrapped in a bookmark the compose shortcut may fill.
    p(['<w:bookmarkStart w:id="1" w:name="astra_slot_title"/>', ph('Report title'), '<w:bookmarkEnd w:id="1"/>'], { style: 'Title' }),
    phPara('Subtitle: one line that says what the report covers', 'Subtitle'),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:before="1200"/>' }),
    para('Prepared for', 'CoverLabel'),
    phPara('Client or audience', 'CoverValue'),
    para('Prepared by', 'CoverLabel'),
    phPara('Author or team', 'CoverValue'),
    para('Date', 'CoverLabel'),
    phPara('Date', 'CoverValue'),
    sectionBreak(page),
  ];
  const contents = [
    para('Contents', 'TOCHeading'),
    p(field('TOC \\o "1-3" \\h \\z \\u', 'Right-click and choose Update Field to build the contents.', { dirty: true })),
    pageBreak(),
  ];
  const region = [
    para('Executive summary', 'Heading1'),
    phPara('Summarise the purpose, the main findings and the recommendation in a few sentences.', 'BodyText'),
    para('Background', 'Heading1'),
    phPara('Describe the context, scope and sources the report draws on.', 'BodyText'),
    para('Findings', 'Heading1'),
    para('Finding one', 'Heading2'),
    phPara('Explain the finding and the evidence for it.', 'BodyText'),
    para('Table 1. Table title', 'Caption'),
    table({
      style: 'AstraTable',
      header: true,
      cantSplit: true,
      widths: [Math.floor(textWidth(page) / 3), Math.floor(textWidth(page) / 3), textWidth(page) - 2 * Math.floor(textWidth(page) / 3)],
      rows: [
        [cell(ph('Column')), cell(ph('Column')), cell(ph('Column'))],
        [cell(ph('Value')), cell(ph('Value')), cell(ph('Value'))],
      ],
    }),
    p('', { style: 'BodyText' }),
    para('Key point the reader must not miss.', 'Callout'),
    para('Recommendations', 'Heading1'),
    phPara('First recommendation', 'ListNumber'),
    phPara('Second recommendation', 'ListNumber'),
    phPara('Third recommendation', 'ListNumber'),
    para('Appendix', 'Heading1'),
    phPara('Supporting material, data tables and sources.', 'BodyText'),
  ];
  const body = [...cover, ...contents, marker(2, 'astra_body_start'), ...region, marker(3, 'astra_body_end')].join('');
  return buildPackage({
    theme: DESIGN,
    template: 'report',
    designVersion: DESIGN_VERSION,
    body,
    page,
    parts: { defaultHeader: runningTitle(), defaultFooter: pageXofY('', 'SECTIONPAGES'), restartNumbering: true },
  });
}

export const reportA4 = (): Buffer => report(A4);
void run;
