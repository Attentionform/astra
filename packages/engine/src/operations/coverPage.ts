import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node } from '../ast';
import { BaseOperation, createTextElement } from './base';
import { DocxParser, upsertSectPrChild } from '../parser';
import { resolveDeterministicZipDate, resolveReproducibleMode } from '../zip-hardening';
import { generateValidParaId } from '../ooxml/schema-order';
import { toError } from '../errors';

function isCoverReproducible(explicit?: boolean, env?: NodeJS.ProcessEnv): boolean {
  try {
    return resolveReproducibleMode(explicit, env ?? process.env);
  } catch {
    return true;
  }
}

/**
 * Epoch-pinned cover date: deterministic by default (1980-01-01 or
 * SOURCE_DATE_EPOCH when set), wall-clock only when opted out.
 * Exported for tests/live proof.
 */
export function resolveCoverDateText(
  explicitDate?: string,
  explicitOrEnv?: boolean | NodeJS.ProcessEnv,
  env?: NodeJS.ProcessEnv
): string {
  if (explicitDate) return explicitDate;
  let explicit: boolean | undefined;
  let effEnv: NodeJS.ProcessEnv;
  if (typeof explicitOrEnv === 'boolean') {
    explicit = explicitOrEnv;
    effEnv = env ?? process.env;
  } else if (explicitOrEnv && typeof explicitOrEnv === 'object') {
    explicit = undefined;
    effEnv = explicitOrEnv;
  } else {
    explicit = undefined;
    effEnv = process.env;
  }
  if (!isCoverReproducible(explicit, effEnv)) {
    return new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  }
  return resolveDeterministicZipDate(effEnv).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
}

/**
 * Deterministic cover node suffix: content-seeded 8-hex when reproducible
 * (stable across runs, unique per slot), wall-clock + random when opted out.
 */
function coverSuffix(slot: string, title: string, subtitle: string, author: string, org: string, dateText: string): string {
  if (!isCoverReproducible()) return `${Date.now()}_${Math.floor(Math.random() * 1000000)}`;
  return generateValidParaId(`cover:${slot}:${title}:${subtitle}:${author}:${org}:${dateText}`, 0);
}

export interface CoverPageOperation extends BaseOperation {
  action: 'coverPage';
  template?: 'modern_minimalist' | 'corporate_navy' | 'academic' | 'executive';
  title: string;
  subtitle?: string;
  author?: string;
  organization?: string;
  /** Short alias for organization (flat or nested in params). */
  org?: string;
  date?: string;
  topSpacingBefore?: number;
  dividerSpacingAfter?: number;
  metaSpacingAfter?: number;
  params?: Record<string, unknown>;
}

function validateCoverSpacing(name: string, val: unknown): number | undefined {
  if (val === undefined) return undefined;
  if (typeof val !== 'number' || !Number.isInteger(val) || val < 0) {
    throw toError(
      'E_VALIDATION',
      `Invalid ${name}: expected non-negative integer dxa, got ${typeof val === 'number' ? val : JSON.stringify(val)}.`,
      `pass ${name} as a non-negative integer (e.g. 1440 for 1 inch, 720 for 0.5 inch, or 0 to collapse spacing)`
    );
  }
  return val;
}

export function executeCoverPage(op: CoverPageOperation, ast: DocumentAST, parser?: DocxParser): void {
  // OP-16: accept args nested under params:{} or flat (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedTemplate = (pick('template', op.template) ?? op.template ?? 'modern_minimalist') as CoverPageOperation['template'];
  const resolvedTitle = String(pick('title', op.title) ?? op.title ?? '');
  const resolvedSubtitle = pick('subtitle', op.subtitle) as string | undefined;
  const resolvedAuthor = pick('author', op.author) as string | undefined;
  const resolvedOrg = (pick('organization', op.organization) ?? op.org) as string | undefined;
  const resolvedDate = pick('date', op.date) as string | undefined;

  const flatTop = validateCoverSpacing('topSpacingBefore', op.topSpacingBefore);
  const nestedTop = validateCoverSpacing('topSpacingBefore', raw['topSpacingBefore']);
  const resolvedTopSpacingBefore = nestedTop !== undefined ? nestedTop : flatTop !== undefined ? flatTop : 1440;

  const flatDivider = validateCoverSpacing('dividerSpacingAfter', op.dividerSpacingAfter);
  const nestedDivider = validateCoverSpacing('dividerSpacingAfter', raw['dividerSpacingAfter']);
  const resolvedDividerSpacingAfter = nestedDivider !== undefined ? nestedDivider : flatDivider !== undefined ? flatDivider : 720;

  const flatMeta = validateCoverSpacing('metaSpacingAfter', op.metaSpacingAfter);
  const nestedMeta = validateCoverSpacing('metaSpacingAfter', raw['metaSpacingAfter']);
  const resolvedMetaSpacing = nestedMeta !== undefined ? nestedMeta : flatMeta;

  const firstSection = ast.sections[0];
  if (!firstSection || firstSection.nodes.length === 0) {
    throw new Error("Cannot insert cover page into an empty document");
  }

  const doc = firstSection.nodes[0].xml.ownerDocument;
  if (!doc) throw new Error("Document not found on first section");

  // 1. Enable <w:titlePg/> in section properties so header/footer are decoupled on Page 1
  const sectPr = doc.getElementsByTagName('w:sectPr')[0];
  if (sectPr) {
    upsertSectPrChild(sectPr, 'w:titlePg');
  }

  const template = resolvedTemplate || 'modern_minimalist';
  const defaultMetaSpacing = template === 'executive' ? 360 : 720;
  const resolvedMetaSpacingAfter = resolvedMetaSpacing !== undefined ? resolvedMetaSpacing : defaultMetaSpacing;
  const titleText = resolvedTitle;
  const subtitleText = resolvedSubtitle || '';
  const authorText = resolvedAuthor || '';
  const orgText = resolvedOrg || '';
  const dateText = resolveCoverDateText(resolvedDate);

  // Accent theme color
  let accentColor = '4F46E5'; // Indigo default
  if (template === 'corporate_navy') accentColor = '1E3A8A';
  if (template === 'academic') accentColor = '334155';
  if (template === 'executive') accentColor = '047857';

  // 2. Build Cover Page Nodes
  const coverNodes: Node[] = [];
  const coverId = (slot: string): string =>
    `cover_${slot}_${coverSuffix(slot, titleText, subtitleText, authorText, orgText, dateText)}`;
  const coverChecksum = (slot: string): string =>
    isCoverReproducible()
      ? `${slot}_${coverSuffix(slot, titleText, subtitleText, authorText, orgText, dateText)}`
      : `${slot}_${Date.now()}_${Math.floor(Math.random() * 1000000)}`;

  // Top spacing paragraph
  const topP = createWordElement(doc, 'w:p');
  topP.setAttribute('data-astra-id', coverId('top'));
  const topPPr = createWordElement(doc, 'w:pPr');
  const topSpacing = createWordElement(doc, 'w:spacing');
  topSpacing.setAttribute('w:before', String(resolvedTopSpacingBefore));
  topPPr.appendChild(topSpacing);
  topP.appendChild(topPPr);
  const createMeta = (text: string = '') => ({
    text,
    hasCitation: false,
    hasComment: false,
    hasBookmark: false,
    hasTrackedChanges: false,
    hasEquation: false,
    hasField: false,
    hasImage: false,
    hasHyperlink: false
  });

  coverNodes.push({
    id: topP.getAttribute('data-astra-id')!,
    type: 'paragraph',
    xml: topP,
    metadata: createMeta(''),
    version: 1,
    checksum: coverChecksum('top')
  });

  // Title Paragraph
  const titleP = createWordElement(doc, 'w:p');
  titleP.setAttribute('data-astra-id', coverId('title'));
  const titlePPr = createWordElement(doc, 'w:pPr');
  const titleSpacing = createWordElement(doc, 'w:spacing');
  titleSpacing.setAttribute('w:before', '240');
  titleSpacing.setAttribute('w:after', '180');
  titlePPr.appendChild(titleSpacing);
  titleP.appendChild(titlePPr);

  const titleRun = createWordElement(doc, 'w:r');
  const titleRPr = createWordElement(doc, 'w:rPr');
  const titleFont = createWordElement(doc, 'w:rFonts');
  titleFont.setAttribute('w:ascii', 'Calibri');
  const titleBold = createWordElement(doc, 'w:b');
  const titleSz = createWordElement(doc, 'w:sz');
  titleSz.setAttribute('w:val', '64'); // 32pt
  const titleColor = createWordElement(doc, 'w:color');
  titleColor.setAttribute('w:val', accentColor);

  titleRPr.appendChild(titleFont);
  titleRPr.appendChild(titleBold);
  titleRPr.appendChild(titleSz);
  titleRPr.appendChild(titleColor);
  titleRun.appendChild(titleRPr);
  titleRun.appendChild(createTextElement(doc, titleText));
  titleP.appendChild(titleRun);
  coverNodes.push({
    id: titleP.getAttribute('data-astra-id')!,
    type: 'paragraph',
    xml: titleP,
    metadata: createMeta(titleText),
    version: 1,
    checksum: coverChecksum('title')
  });

  // Subtitle Paragraph
  if (subtitleText) {
    const subP = createWordElement(doc, 'w:p');
    subP.setAttribute('data-astra-id', coverId('sub'));
    const subPPr = createWordElement(doc, 'w:pPr');
    const subSpacing = createWordElement(doc, 'w:spacing');
    subSpacing.setAttribute('w:after', '480');
    subPPr.appendChild(subSpacing);
    subP.appendChild(subPPr);

    const subRun = createWordElement(doc, 'w:r');
    const subRPr = createWordElement(doc, 'w:rPr');
    const subSz = createWordElement(doc, 'w:sz');
    subSz.setAttribute('w:val', '28'); // 14pt
    const subColor = createWordElement(doc, 'w:color');
    subColor.setAttribute('w:val', '64748B');
    subRPr.appendChild(subSz);
    subRPr.appendChild(subColor);
    subRun.appendChild(subRPr);
    subRun.appendChild(createTextElement(doc, subtitleText));
    subP.appendChild(subRun);
    coverNodes.push({
      id: subP.getAttribute('data-astra-id')!,
      type: 'paragraph',
      xml: subP,
      metadata: createMeta(subtitleText),
      version: 1,
      checksum: coverChecksum('sub')
    });
  }

  // Accent divider line on cover
  const divP = createWordElement(doc, 'w:p');
  divP.setAttribute('data-astra-id', coverId('div'));
  const divPPr = createWordElement(doc, 'w:pPr');
  const divBdr = createWordElement(doc, 'w:pBdr');
  const bottomBdr = createWordElement(doc, 'w:bottom');
  bottomBdr.setAttribute('w:val', 'single');
  bottomBdr.setAttribute('w:sz', '24');
  bottomBdr.setAttribute('w:space', '1');
  bottomBdr.setAttribute('w:color', accentColor);
  divBdr.appendChild(bottomBdr);
  divPPr.appendChild(divBdr);
  const divSpacing = createWordElement(doc, 'w:spacing');
  divSpacing.setAttribute('w:before', '240');
  divSpacing.setAttribute('w:after', String(resolvedDividerSpacingAfter));
  divPPr.appendChild(divSpacing);
  divP.appendChild(divPPr);
  coverNodes.push({
    id: divP.getAttribute('data-astra-id')!,
    type: 'paragraph',
    xml: divP,
    metadata: createMeta(''),
    version: 1,
    checksum: coverChecksum('div')
  });

  // Metadata Paragraph (Author, Org, Date)
  const metaP = createWordElement(doc, 'w:p');
  metaP.setAttribute('data-astra-id', coverId('meta'));
  const metaPPr = createWordElement(doc, 'w:pPr');
  const metaSpacing = createWordElement(doc, 'w:spacing');
  metaSpacing.setAttribute('w:after', String(resolvedMetaSpacingAfter));
  metaPPr.appendChild(metaSpacing);
  metaP.appendChild(metaPPr);

  const metaRun = createWordElement(doc, 'w:r');
  const metaRPr = createWordElement(doc, 'w:rPr');
  const metaSz = createWordElement(doc, 'w:sz');
  metaSz.setAttribute('w:val', '22'); // 11pt
  const metaColor = createWordElement(doc, 'w:color');
  metaColor.setAttribute('w:val', '475569');
  metaRPr.appendChild(metaSz);
  metaRPr.appendChild(metaColor);
  metaRun.appendChild(metaRPr);

  const metaLines = [authorText, orgText, dateText].filter(Boolean).join(' • ');
  metaRun.appendChild(createTextElement(doc, metaLines || dateText));
  metaP.appendChild(metaRun);
  coverNodes.push({
    id: metaP.getAttribute('data-astra-id')!,
    type: 'paragraph',
    xml: metaP,
    metadata: createMeta(metaLines || dateText),
    version: 1,
    checksum: coverChecksum('meta')
  });

  // Page Break after Cover Page
  const pageBreakP = createWordElement(doc, 'w:p');
  pageBreakP.setAttribute('data-astra-id', coverId('pb'));
  const pbRun = createWordElement(doc, 'w:r');
  const br = createWordElement(doc, 'w:br');
  br.setAttribute('w:type', 'page');
  pbRun.appendChild(br);
  pageBreakP.appendChild(pbRun);
  coverNodes.push({
    id: pageBreakP.getAttribute('data-astra-id')!,
    type: 'paragraph',
    xml: pageBreakP,
    metadata: createMeta(''),
    version: 1,
    checksum: coverChecksum('pb')
  });

  // Insert DOM elements at start of document
  const firstChild = firstSection.nodes[0].xml;
  const parentNode = firstChild.parentNode;
  if (parentNode) {
    coverNodes.forEach(node => {
      parentNode.insertBefore(node.xml, firstChild);
    });
  }

  // Prepend to AST
  firstSection.nodes.unshift(...coverNodes);
  firstSection.version += 1;

  // OP-17: register synthesized nodes in parser.nodeMap so subsequent
  // edits (replace/delete/move/style) can resolve them.
  if (parser) {
    for (const node of coverNodes) {
      try {
        parser.nodeMap?.set(node.id, node.xml);
      } catch { /* best-effort */ }
    }
  }
}
