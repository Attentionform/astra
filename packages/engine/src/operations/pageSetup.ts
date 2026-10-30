import { DocumentAST } from '../ast';
import { BaseOperation } from './base';
import { DocxParser, upsertSectPrChild } from '../parser';
import { sectionPosition } from './section-number';
import { orientPageDimensions, readPageDimensions, readPageOrientation } from '../ooxml/page-size';

export interface PageSetupOperation extends BaseOperation {
  action: 'pageSetup';
  params?: {
    marginTop?: number; // In DXA (1440 = 1 inch)
    marginBottom?: number;
    marginLeft?: number;
    marginRight?: number;
    headerDistance?: number;
    footerDistance?: number;
    /** Short alias for headerDistance (flat or nested in params). */
    header?: number;
    /** Short alias for footerDistance (flat or nested in params). */
    footer?: number;
    orientation?: 'portrait' | 'landscape';
    paperSize?: 'A4' | 'Letter' | 'Legal' | 'A3';
    columns?: number;
    sectionIndex?: number;
  };
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  headerDistance?: number;
  footerDistance?: number;
  /** Short alias for headerDistance (flat or nested in params). */
  header?: number;
  /** Short alias for footerDistance (flat or nested in params). */
  footer?: number;
  orientation?: 'portrait' | 'landscape';
  paperSize?: 'A4' | 'Letter' | 'Legal' | 'A3';
  columns?: number;
  sectionIndex?: number;
}

const PAGE_SIZE_PRESETS: Record<string, { width: number; height: number }> = {
  A3: { width: 16838, height: 23811 },
  A4: { width: 11906, height: 16838 },
  Letter: { width: 12240, height: 15840 },
  Legal: { width: 12240, height: 20160 }
};

/**
 * Mutates the section's <w:sectPr> (w:pgSz / w:pgMar / w:cols) for real —
 * with schema-correct child order — and keeps the AST section metadata in
 * sync so TOON readers observe the change.
 */
export function executePageSetup(op: PageSetupOperation, ast: DocumentAST, parser?: DocxParser): void {
  const sectPrs = parser?.getSectPrElements() ?? [];
  if (sectPrs.length === 0) {
    throw new Error('pageSetup requires a document section (<w:sectPr>); attach a DocxParser to the engine.');
  }

  const p = op.params || op;
  const marginTop = p.marginTop ?? op.marginTop;
  const marginBottom = p.marginBottom ?? op.marginBottom;
  const marginLeft = p.marginLeft ?? op.marginLeft;
  const marginRight = p.marginRight ?? op.marginRight;
  const headerDistance = p.headerDistance ?? p.header ?? op.headerDistance ?? op.header;
  const footerDistance = p.footerDistance ?? p.footer ?? op.footerDistance ?? op.footer;
  const orientation = p.orientation ?? op.orientation;
  const paperSize = p.paperSize ?? op.paperSize;
  const columns = p.columns ?? op.columns;
  const sectionIndex = sectionPosition('pageSetup', p.sectionIndex ?? op.sectionIndex, sectPrs.length, 'omit it to change every section');
  const targets = sectionIndex === undefined ? sectPrs : [sectPrs[sectionIndex]];

  for (const sectPr of targets) {
    if (
      marginTop !== undefined || marginBottom !== undefined ||
      marginLeft !== undefined || marginRight !== undefined ||
      headerDistance !== undefined || footerDistance !== undefined
    ) {
      const pgMar = upsertSectPrChild(sectPr, 'w:pgMar');
      if (marginTop !== undefined) pgMar.setAttribute('w:top', marginTop.toString());
      if (marginBottom !== undefined) pgMar.setAttribute('w:bottom', marginBottom.toString());
      if (marginLeft !== undefined) pgMar.setAttribute('w:left', marginLeft.toString());
      if (marginRight !== undefined) pgMar.setAttribute('w:right', marginRight.toString());
      if (headerDistance !== undefined) pgMar.setAttribute('w:header', headerDistance.toString());
      if (footerDistance !== undefined) pgMar.setAttribute('w:footer', footerDistance.toString());
    }

    if (paperSize !== undefined || orientation !== undefined) {
      const pgSz = upsertSectPrChild(sectPr, 'w:pgSz');

      // The page that is laid out is the size the section already has (or the
      // preset that replaces it), turned to the orientation that applies: the
      // one asked for, otherwise the one the section was in.
      const keptOrientation = readPageOrientation(pgSz);
      let page = readPageDimensions(pgSz);
      if (paperSize !== undefined) {
        const preset = PAGE_SIZE_PRESETS[paperSize];
        if (!preset) throw new Error(`Unknown paper size '${paperSize}'. Supported: ${Object.keys(PAGE_SIZE_PRESETS).join(', ')}`);
        page = preset;
        // A printer paper code left over from the old size would name the wrong paper.
        pgSz.removeAttribute('w:code');
      }

      if (orientation !== undefined) {
        pgSz.setAttribute('w:orient', orientation);
      }
      page = orientPageDimensions(page, orientation ?? keptOrientation);
      pgSz.setAttribute('w:w', page.width.toString());
      pgSz.setAttribute('w:h', page.height.toString());
    }

    if (columns !== undefined) {
      const cols = upsertSectPrChild(sectPr, 'w:cols');
      cols.setAttribute('w:num', columns.toString());
    }
  }

  // Keep AST metadata in sync with the mutated XML.
  const syncSections = sectionIndex === undefined ? ast.sections : ast.sections.slice(sectionIndex, sectionIndex + 1);
  for (const section of syncSections) {
    if (!section.pageSetup) section.pageSetup = {};
    if (marginTop !== undefined) section.pageSetup.marginTop = marginTop;
    if (marginBottom !== undefined) section.pageSetup.marginBottom = marginBottom;
    if (marginLeft !== undefined) section.pageSetup.marginLeft = marginLeft;
    if (marginRight !== undefined) section.pageSetup.marginRight = marginRight;
    if (headerDistance !== undefined) section.pageSetup.headerDistance = headerDistance;
    if (footerDistance !== undefined) section.pageSetup.footerDistance = footerDistance;
    if (orientation !== undefined) section.pageSetup.orientation = orientation;
    if (paperSize !== undefined) section.pageSetup.paperSize = paperSize;
    if (columns !== undefined) section.pageSetup.columns = columns;
    section.version += 1;
  }
}
