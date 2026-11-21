import { DESIGN, DESIGN_VERSION, MANUSCRIPT } from './themes';
import { NUM_HEADINGS } from './design';
import { LETTER, buildPackage, coverRuleIndent, textWidth, cell, field, marker, p, pageBreak, pageXofY, para, ph, phPara, run, runningTitle, sectionBreak, table, type PageSetup } from './lib';

const D = DESIGN;
const slot = (id: number, name: string, inner: string): string => `<w:bookmarkStart w:id="${id}" w:name="astra_slot_${name}"/>${inner}<w:bookmarkEnd w:id="${id}"/>`;
const rule = (): string => p('', { style: 'NoSpacing', raw: `<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="${D.colors.ink}"/></w:pBdr><w:spacing w:after="240"/>` });
const labelRow = (label: string, value: string): ReturnType<typeof cell>[] => [cell(run(label, { style: 'Strong' })), cell(value)];

/** A blank document: the design system's styles and page setup, one empty paragraph, nothing visible. */
export function blank(page: PageSetup = LETTER): Buffer {
  return buildPackage({ theme: D, template: 'blank', designVersion: DESIGN_VERSION, body: p('', { style: 'Normal' }), page });
}

export function memo(page: PageSetup = LETTER): Buffer {
  const w = textWidth(page);
  const body = [
    para('Memo', 'Title'),
    table({
      style: 'AstraTableGrid',
      widths: [1400, w - 1400],
      textStyle: 'Normal',
      rows: [
        labelRow('To', ph('Recipient')),
        labelRow('From', ph('Sender')),
        labelRow('Date', ph('Date')),
        labelRow('Subject', slot(1, 'title', ph('Subject'))),
        labelRow('Cc', ph('Names, or delete this row')),
      ],
    }),
    rule(),
    marker(2, 'astra_body_start'),
    para('Purpose', 'Heading2'),
    phPara('State in one or two sentences why you are writing and what you need.', 'BodyText'),
    para('Background', 'Heading2'),
    phPara('Give the context the reader needs to act.', 'BodyText'),
    para('Action required', 'Heading2'),
    phPara('First action, with owner and date', 'ListBullet'),
    phPara('Second action, with owner and date', 'ListBullet'),
    marker(3, 'astra_body_end'),
  ].join('');
  return buildPackage({ theme: D, template: 'memo', designVersion: DESIGN_VERSION, body, page, parts: { defaultFooter: pageXofY() } });
}

export function letter(page: PageSetup = LETTER): Buffer {
  const body = [
    phPara('Your name', 'NoSpacing'),
    phPara('Street address', 'NoSpacing'),
    phPara('City, region, postal code', 'NoSpacing'),
    p(ph('Email address · Phone'), { style: 'NoSpacing', raw: '<w:spacing w:after="360"/>' }),
    phPara('Date', 'BodyText'),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:after="120"/>' }),
    phPara('Recipient name', 'NoSpacing'),
    phPara('Recipient address', 'NoSpacing'),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:after="120"/>' }),
    p(slot(1, 'title', run('[Subject line]', { style: 'Placeholder', bold: true })), { style: 'BodyText' }),
    phPara('Dear Name,', 'BodyText'),
    marker(2, 'astra_body_start'),
    phPara('Open with the reason you are writing.', 'BodyText'),
    phPara('Give the detail the reader needs, in short paragraphs.', 'BodyText'),
    phPara('Close with what you ask of the reader or what happens next.', 'BodyText'),
    marker(3, 'astra_body_end'),
    para('Sincerely,', 'BodyText'),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:after="720"/>' }),
    phPara('Your name', 'NoSpacing'),
  ].join('');
  return buildPackage({ theme: D, template: 'letter', designVersion: DESIGN_VERSION, body, page });
}

export function invoice(page: PageSetup = LETTER): Buffer {
  const w = textWidth(page);
  const half = Math.floor(w / 2);
  const body = [
    para('Invoice', 'Title'),
    table({
      style: 'AstraTableMinimal',
      header: false,
      widths: [2000, w - 2000],
      textStyle: 'TableText',
      rows: [labelRow('Invoice number', ph('Invoice number')), labelRow('Issue date', ph('Date')), labelRow('Due date', ph('Date'))],
    }),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:after="240"/>' }),
    table({
      style: 'AstraTableGrid',
      widths: [half, w - half],
      textStyle: 'Normal',
      rows: [
        [cell(para('From', 'CoverLabel'), { blocks: true }), cell(para('Bill to', 'CoverLabel'), { blocks: true })],
        [
          cell([phPara('Your business name', 'NoSpacing'), phPara('Address', 'NoSpacing'), phPara('Email · Phone', 'NoSpacing')].join(''), { blocks: true }),
          cell([phPara('Client name', 'NoSpacing'), phPara('Address', 'NoSpacing'), phPara('Email', 'NoSpacing')].join(''), { blocks: true }),
        ],
      ],
    }),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:after="240"/>' }),
    table({
      style: 'AstraTableMinimal',
      header: true,
      cantSplit: true,
      widths: [w - 1200 - 1700 - 1800, 1200, 1700, 1800],
      rows: [
        [cell(run('Description')), cell(run('Qty'), { align: 'right' }), cell(run('Unit price'), { align: 'right' }), cell(run('Amount'), { align: 'right' })],
        ...[0, 1, 2].map(() => [cell(ph('Item or service')), cell(ph('Qty'), { align: 'right' }), cell(ph('Price'), { align: 'right' }), cell(ph('Amount'), { align: 'right' })]),
      ],
    }),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:after="120"/>' }),
    table({
      style: 'AstraTableGrid',
      widths: [w - 3500, 1700, 1800],
      textStyle: 'TableText',
      rows: [
        [cell(''), cell(run('Subtotal'), { align: 'right' }), cell(ph('Amount'), { align: 'right' })],
        [cell(''), cell(run('Tax'), { align: 'right' }), cell(ph('Amount'), { align: 'right' })],
        [cell(''), cell(run('Total due', { bold: true }), { align: 'right' }), cell(run('[Amount]', { style: 'Placeholder', bold: true }), { align: 'right' })],
      ],
    }),
    para('Payment details', 'Heading2'),
    phPara('Bank, account name, account number or payment link', 'BodyText'),
    marker(1, 'astra_body_start'),
    para('Notes', 'Heading2'),
    phPara('Payment terms, late fees or a thank-you note.', 'BodyText'),
    marker(2, 'astra_body_end'),
  ].join('');
  return buildPackage({ theme: D, template: 'invoice', designVersion: DESIGN_VERSION, body, page, parts: { defaultFooter: pageXofY() } });
}

function coverBlock(page: PageSetup, ruleColor: string): string[] {
  return [
    p('', { style: 'NoSpacing', raw: `<w:pBdr><w:top w:val="single" w:sz="36" w:space="0" w:color="${ruleColor}"/></w:pBdr><w:spacing w:before="2400" w:after="240"/><w:ind w:right="${coverRuleIndent(page)}"/>` }),
    p(slot(1, 'title', ph('Document title')), { style: 'Title' }),
    phPara('Subtitle: one line that says what this covers', 'Subtitle'),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:before="1200"/>' }),
    para('Prepared for', 'CoverLabel'),
    phPara('Client or audience', 'CoverValue'),
    para('Prepared by', 'CoverLabel'),
    phPara('Author or team', 'CoverValue'),
    para('Date', 'CoverLabel'),
    phPara('Date', 'CoverValue'),
    sectionBreak(page),
  ];
}

const contents = (): string[] => [
  para('Contents', 'TOCHeading'),
  p(field('TOC \\o "1-3" \\h \\z \\u', 'Right-click and choose Update Field to build the contents.', { dirty: true })),
  pageBreak(),
];

export function proposal(page: PageSetup = LETTER): Buffer {
  const w = textWidth(page);
  const body = [
    ...coverBlock(page, D.colors.accent),
    ...contents(),
    marker(2, 'astra_body_start'),
    para('Summary', 'Heading1'),
    phPara('Say what you propose, for whom, and the outcome the client gets.', 'BodyText'),
    para('Problem and objectives', 'Heading1'),
    phPara('Describe the problem in the client’s terms and the objectives of the work.', 'BodyText'),
    para('Approach', 'Heading1'),
    phPara('Explain how you will do the work and why this approach fits.', 'BodyText'),
    para('Timeline', 'Heading1'),
    table({
      style: 'AstraTable', header: true, cantSplit: true, widths: [2400, w - 2400 - 2000, 2000],
      rows: [
        [cell(run('Phase')), cell(run('Deliverable')), cell(run('Date'), { align: 'right' })],
        [cell(ph('Phase')), cell(ph('Deliverable')), cell(ph('Date'), { align: 'right' })],
        [cell(ph('Phase')), cell(ph('Deliverable')), cell(ph('Date'), { align: 'right' })],
      ],
    }),
    para('Pricing', 'Heading1'),
    table({
      style: 'AstraTable', header: true, cantSplit: true, widths: [w - 2000, 2000],
      rows: [
        [cell(run('Item')), cell(run('Amount'), { align: 'right' })],
        [cell(ph('Item')), cell(ph('Amount'), { align: 'right' })],
        [cell(ph('Item')), cell(ph('Amount'), { align: 'right' })],
        [cell(run('Total', { bold: true })), cell(run('[Amount]', { style: 'Placeholder', bold: true }), { align: 'right' })],
      ],
    }),
    para('Terms', 'Heading1'),
    phPara('Payment terms, validity of this proposal and any assumptions.', 'BodyText'),
    para('Acceptance', 'Heading1'),
    marker(3, 'astra_body_end'),
    table({
      style: 'AstraTableGrid', widths: [Math.floor(w / 2), w - Math.floor(w / 2)], textStyle: 'Normal',
      rows: [
        [cell(para('For the client', 'CoverLabel'), { blocks: true }), cell(para('For the provider', 'CoverLabel'), { blocks: true })],
        [cell([phPara('Name', 'NoSpacing'), phPara('Signature', 'NoSpacing'), phPara('Date', 'NoSpacing')].join(''), { blocks: true }), cell([phPara('Name', 'NoSpacing'), phPara('Signature', 'NoSpacing'), phPara('Date', 'NoSpacing')].join(''), { blocks: true })],
      ],
    }),
  ].join('');
  return buildPackage({ theme: D, template: 'proposal', designVersion: DESIGN_VERSION, body, page, parts: { defaultHeader: runningTitle(), defaultFooter: pageXofY('', 'SECTIONPAGES'), restartNumbering: true } });
}

export function technicalSpec(page: PageSetup = LETTER): Buffer {
  const w = textWidth(page);
  const body = [
    p(slot(1, 'title', ph('Service name')), { style: 'Title' }),
    phPara('Subtitle: what this specification covers', 'Subtitle'),
    table({
      style: 'AstraTableMinimal', widths: [2000, w - 2000], textStyle: 'TableText',
      rows: [labelRow('Version', ph('Version')), labelRow('Status', ph('Draft, review or approved')), labelRow('Owner', ph('Owner')), labelRow('Last updated', ph('Date'))],
    }),
    marker(2, 'astra_body_start'),
    para('Overview', 'Heading1'),
    phPara('What the service does, who uses it and what it does not cover.', 'BodyText'),
    para('Authentication', 'Heading1'),
    phPara('How a client proves who it is, and what each credential may do.', 'BodyText'),
    para('Endpoints', 'Heading1'),
    p([ph('Method'), run(' '), ph('Path')], { style: 'Heading2' }),
    phPara('What the endpoint does.', 'BodyText'),
    para('Table 1. Parameters', 'Caption'),
    table({
      style: 'AstraTable', header: true, cantSplit: true, widths: [2000, 1700, 1300, w - 5000],
      rows: [
        [cell(run('Name')), cell(run('Type')), cell(run('Required')), cell(run('Description'))],
        [cell(ph('name')), cell(ph('type')), cell(ph('yes or no')), cell(ph('Description'))],
      ],
    }),
    para('Example request', 'Heading3'),
    phPara('Request line and headers', 'Code'),
    para('Example response', 'Heading3'),
    phPara('Status line and body', 'Code'),
    para('Error codes', 'Heading1'),
    table({
      style: 'AstraTable', header: true, cantSplit: true, widths: [1500, 2500, w - 4000],
      rows: [
        [cell(run('Code')), cell(run('Name')), cell(run('Meaning and remedy'))],
        [cell(ph('Code')), cell(ph('Name')), cell(ph('Meaning and remedy'))],
      ],
    }),
    para('Changelog', 'Heading1'),
    table({
      style: 'AstraTable', header: true, cantSplit: true, widths: [1500, 1700, w - 3200],
      rows: [
        [cell(run('Version')), cell(run('Date')), cell(run('Change'))],
        [cell(ph('Version')), cell(ph('Date')), cell(ph('Change'))],
      ],
    }),
    marker(3, 'astra_body_end'),
  ].join('');
  return buildPackage({ theme: D, template: 'technical-spec', designVersion: DESIGN_VERSION, body, page, parts: { defaultHeader: runningTitle(), defaultFooter: pageXofY() } });
}

const numbered = (text: string, level = 0): string => p(run(text), { style: level === 0 ? 'Heading1' : 'Heading2', raw: `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${NUM_HEADINGS}"/></w:numPr>` });

export function researchPaper(page: PageSetup = LETTER): Buffer {
  const body = [
    p(slot(1, 'title', ph('Paper title')), { style: 'Title' }),
    phPara('Author names', 'CoverValue'),
    phPara('Affiliations and contact', 'Subtitle'),
    para('Abstract', 'Heading3'),
    phPara('Summarise the question, the method, the main result and the conclusion in about 150 to 250 words.', 'BodyText'),
    p([run('Keywords: ', { style: 'Strong' }), ph('keyword'), run(', '), ph('keyword'), run(', '), ph('keyword')], { style: 'BodyText' }),
    marker(2, 'astra_body_start'),
    numbered('Introduction'),
    phPara('State the problem, why it matters and what this paper contributes.', 'BodyText'),
    numbered('Methods'),
    phPara('Describe the data, the procedure and how results were checked.', 'BodyText'),
    numbered('Results'),
    phPara('Report what was found. Refer to each table and figure by number.', 'BodyText'),
    numbered('Discussion'),
    phPara('Interpret the results, name the limits and relate them to earlier work.', 'BodyText'),
    numbered('Conclusion'),
    phPara('Restate the contribution and what should happen next.', 'BodyText'),
    para('References', 'Heading1'),
    phPara('Author. Year. Title. Source.', 'Reference'),
    para('Appendix', 'Heading1'),
    phPara('Supplementary material.', 'BodyText'),
    marker(3, 'astra_body_end'),
  ].join('');
  return buildPackage({ theme: D, template: 'research-paper', designVersion: DESIGN_VERSION, body, page, parts: { defaultFooter: pageXofY() } });
}

export function manuscript(page: PageSetup = LETTER): Buffer {
  const M = MANUSCRIPT;
  const body = [
    phPara('Author name', 'NoSpacing'),
    phPara('Street address', 'NoSpacing'),
    phPara('City, region, postal code', 'NoSpacing'),
    phPara('Email address · Phone', 'NoSpacing'),
    p('', { style: 'NoSpacing', raw: '<w:spacing w:before="2800"/>' }),
    p(slot(1, 'title', ph('Title')), { style: 'Title' }),
    p(run('by ', {}) + ph('Author name'), { style: 'Subtitle' }),
    p(ph('Approximate word count'), { style: 'Subtitle' }),
    marker(2, 'astra_body_start'),
    para('Chapter One', 'Heading1'),
    p(ph('Chapter title or opening line'), { style: 'BodyText' }),
    phPara('Begin the chapter here. A first-line indent marks each new paragraph; there is no extra space between paragraphs.', 'BodyText'),
    marker(3, 'astra_body_end'),
  ].join('');
  // Header on every page after the title page: surname / title / page number, right aligned.
  const hdr = p([ph('Surname'), run(' / '), ph('Title'), run(' / '), field('PAGE', '2')], { style: 'Header', raw: '<w:pBdr><w:bottom w:val="nil"/></w:pBdr><w:jc w:val="right"/>' });
  return buildPackage({
    theme: M, template: 'novel', designVersion: DESIGN_VERSION, body, page,
    parts: { defaultHeader: hdr, firstHeader: '<w:p/>', titlePg: true },
  });
}

export function newsletter(page: PageSetup = LETTER): Buffer {
  const body = [
    p('', { style: 'NoSpacing', raw: `<w:pBdr><w:top w:val="single" w:sz="36" w:space="0" w:color="${D.colors.accent}"/></w:pBdr><w:spacing w:before="0" w:after="160"/>` }),
    p(slot(1, 'title', ph('Publication name')), { style: 'Title', raw: '<w:jc w:val="center"/>' }),
    p([ph('Issue'), run('  ·  '), ph('Date')], { style: 'Subtitle', raw: '<w:jc w:val="center"/>' }),
    sectionBreak(page),
    marker(2, 'astra_body_start'),
    para('Lead story headline', 'Heading1'),
    phPara('Open with the most important news in two or three sentences. Keep paragraphs short; columns are narrow.', 'BodyText'),
    para('Second story headline', 'Heading2'),
    phPara('A shorter piece. One idea, one paragraph, one next step.', 'BodyText'),
    para('A pull quote or key fact for the sidebar.', 'Callout'),
    para('In brief', 'Heading2'),
    phPara('Short item', 'ListBullet'),
    phPara('Short item', 'ListBullet'),
    phPara('Short item', 'ListBullet'),
    marker(3, 'astra_body_end'),
  ].join('');
  return buildPackage({ theme: D, template: 'newsletter', designVersion: DESIGN_VERSION, body, page, parts: { defaultFooter: pageXofY(), cols: 2, continuous: true } });
}
