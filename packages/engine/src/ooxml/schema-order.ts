/**
 * ISO/IEC 29500 (ECMA-376) Schema Ordering & Canonicalization Engine
 * 
 * Microsoft Word strictly validates XML against the ECMA-376 XSD schemas.
 * When child elements appear out of the sequence specified in the schema
 * (e.g. <w:rFonts> placed after <w:sz> in <w:rPr>), Word considers the document
 * corrupted and refuses to open it.
 * 
 * This module defines the canonical element sequences and provides utilities
 * to insert elements in schema order and canonicalize entire XML trees before saving.
 */

import crypto from 'crypto';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createWordElement } from './ns';
import { assertNoXxe } from './xxe-guard';

/**
 * Strict sequence for child elements of <w:rPr> (Run Properties)
 * Ref: ECMA-376 Part 1, Section 17.3.2.28 (EG_RPrBase)
 */
export const RPR_CHILD_ORDER: readonly string[] = [
  'w:rStyle',
  'w:rFonts',
  'w:b',
  'w:bCs',
  'w:i',
  'w:iCs',
  'w:caps',
  'w:smallCaps',
  'w:strike',
  'w:dstrike',
  'w:outline',
  'w:shadow',
  'w:emboss',
  'w:imprint',
  'w:noProof',
  'w:snapToGrid',
  'w:vanish',
  'w:webHidden',
  'w:color',
  'w:spacing',
  'w:w',
  'w:kern',
  'w:position',
  'w:sz',
  'w:szCs',
  'w:highlight',
  'w:u',
  'w:effect',
  'w:bdr',
  'w:shd',
  'w:fitText',
  'w:vertAlign',
  'w:rtl',
  'w:cs',
  'w:em',
  'w:lang',
  'w:eastAsianLayout',
  'w:specVanish',
  'w:oMath',
  // Word 2010 text effects (MS-DOCX 2.6.1). They follow every w: child and
  // come before w:rPrChange, in this order. Unlisted, they kept their slot
  // while the listed elements moved around them, which put w:lang or w:szCs
  // after a w14 element and made the file invalid.
  'w14:glow',
  'w14:shadow',
  'w14:reflection',
  'w14:textOutline',
  'w14:textFill',
  'w14:scene3d',
  'w14:props3d',
  'w14:ligatures',
  'w14:numForm',
  'w14:numSpacing',
  'w14:stylisticSets',
  'w14:cntxtAlts',
  'w:rPrChange'
];

/**
 * Strict sequence for child elements of <w:pPr> (Paragraph Properties)
 * Ref: ECMA-376 Part 1, Section 17.3.1.25 (pPr)
 */
export const PPR_CHILD_ORDER: readonly string[] = [
  'w:pStyle',
  'w:keepNext',
  'w:keepLines',
  'w:pageBreakBefore',
  'w:framePr',
  'w:widowControl',
  'w:numPr',
  'w:suppressLineNumbers',
  'w:pBdr',
  'w:shd',
  'w:tabs',
  'w:suppressAutoHyphens',
  'w:kinsoku',
  'w:wordWrap',
  'w:overflowPunct',
  'w:topLinePunct',
  'w:autoSpaceDE',
  'w:autoSpaceDN',
  'w:bidi',
  'w:adjustRightInd',
  'w:snapToGrid',
  'w:spacing',
  'w:ind',
  'w:contextualSpacing',
  'w:mirrorIndents',
  'w:suppressOverlap',
  'w:jc',
  'w:textDirection',
  'w:textAlignment',
  'w:textboxTightWrap',
  'w:outlineLvl',
  'w:divId',
  'w:cnfStyle',
  'w:rPr',
  'w:sectPr',
  'w:pPrChange'
];

/**
 * Strict sequence for child elements of <w:tblPr> (Table Properties)
 * Ref: ECMA-376 Part 1, Section 17.4.34 (tblPr)
 */
export const TBLPR_CHILD_ORDER: readonly string[] = [
  'w:tblStyle',
  'w:tblpPr',
  'w:tblOverlap',
  'w:bidiVisual',
  'w:tblStyleRowBandSize',
  'w:tblStyleColBandSize',
  'w:tblW',
  'w:jc',
  'w:tblCellSpacing',
  'w:tblInd',
  'w:tblBorders',
  'w:shd',
  'w:tblLayout',
  'w:tblCellMar',
  'w:tblLook',
  'w:tblCaption',
  'w:tblDescription',
  'w:tblPrChange'
];

/**
 * Strict sequence for child elements of <w:tblBorders> (Table Borders)
 * Ref: ECMA-376 Part 1, Section 17.4.39 (tblBorders)
 */
export const TBLBORDERS_CHILD_ORDER: readonly string[] = [
  'w:top',
  'w:left',
  'w:bottom',
  'w:right',
  'w:insideH',
  'w:insideV',
  'w:tl2br',
  'w:tr2bl'
];

/**
 * Strict sequence for child elements of <w:numPr> (Numbering Properties)
 * Ref: ECMA-376 Part 1, Section 17.3.1.19 (numPr)
 */
export const NUMPR_CHILD_ORDER: readonly string[] = [
  'w:ilvl',
  'w:numId',
  'w:numberingChange',
  'w:ins'
];

/**
 * Strict sequence for child elements of <w:tcPr> (Table Cell Properties)
 * Ref: ECMA-376 Part 1, Section 17.4.70 (tcPr)
 */
export const TCPR_CHILD_ORDER: readonly string[] = [
  'w:cnfStyle',
  'w:tcW',
  'w:gridSpan',
  'w:hMerge',
  'w:vMerge',
  'w:tcBorders',
  'w:shd',
  'w:noWrap',
  'w:tcMar',
  'w:textDirection',
  'w:tcFitText',
  'w:vAlign',
  'w:hideMark',
  'w:headers',
  'w:cellIns',
  'w:cellDel',
  'w:cellMerge',
  'w:tcPrChange'
];

/**
  * Strict sequence for child elements of <w:trPr> (Table Row Properties)
  * Ref: ECMA-376 Part 1, Section 17.4.82 (CT_TrPrBase)
  */
export const TRPR_CHILD_ORDER: readonly string[] = [
  'w:cnfStyle',
  'w:divId',
  'w:gridBefore',
  'w:gridAfter',
  'w:wBefore',
  'w:wAfter',
  'w:cantSplit',
  'w:trHeight',
  'w:tblHeader',
  'w:tblCellSpacing',
  'w:jc',
  'w:hidden',
  'w:ins',
  'w:del',
  'w:trPrChange'
];

/**
 * CT_Settings child order (ECMA-376 Part 1, w:settings sequence), the
 * subset ASTra reads or writes. `updateFields` precedes the note properties
 * and `compat`; `trackRevisions` follows `revisionView`.
 */
export const SETTINGS_CHILD_ORDER: readonly string[] = [
  'w:writeProtection', 'w:view', 'w:zoom', 'w:removePersonalInformation',
  'w:removeDateAndTime', 'w:doNotDisplayPageBoundaries',
  'w:displayBackgroundShape', 'w:printPostScriptOverText', 'w:printFractionalCharacterWidth',
  'w:printFormsData', 'w:embedTrueTypeFonts', 'w:embedSystemFonts',
  'w:saveSubsetFonts', 'w:saveFormsData', 'w:mirrorMargins', 'w:alignBordersAndEdges',
  'w:bordersDoNotSurroundHeader', 'w:bordersDoNotSurroundFooter', 'w:gutterAtTop',
  'w:hideSpellingErrors', 'w:hideGrammaticalErrors', 'w:activeWritingStyle',
  'w:proofState', 'w:formsDesign', 'w:attachedTemplate', 'w:linkStyles',
  'w:stylePaneFormatFilter', 'w:stylePaneSortMethod', 'w:documentType', 'w:mailMerge',
  'w:revisionView', 'w:trackRevisions', 'w:doNotTrackMoves', 'w:doNotTrackFormatting',
  'w:documentProtection', 'w:autoFormatOverride', 'w:styleLockTheme', 'w:styleLockQFSet',
  'w:defaultTabStop', 'w:autoHyphenation', 'w:consecutiveHyphenLimit', 'w:hyphenationZone',
  'w:doNotHyphenateCaps', 'w:showEnvelope', 'w:summaryLength', 'w:clickAndTypeStyle',
  'w:defaultTableStyle', 'w:evenAndOddHeaders', 'w:bookFoldRevPrinting', 'w:bookFoldPrinting',
  'w:bookFoldPrintingSheets', 'w:drawingGridHorizontalSpacing', 'w:drawingGridVerticalSpacing',
  'w:displayHorizontalDrawingGridEvery', 'w:displayVerticalDrawingGridEvery',
  'w:doNotUseMarginsForDrawingGridOrigin', 'w:drawingGridHorizontalOrigin',
  'w:drawingGridVerticalOrigin', 'w:doNotShadeFormData', 'w:noPunctuationKerning',
  'w:characterSpacingControl', 'w:printTwoOnOne', 'w:strictFirstAndLastChars',
  'w:noLineBreaksAfter', 'w:noLineBreaksBefore', 'w:savePreviewPicture',
  'w:doNotValidateAgainstSchema', 'w:saveInvalidXml', 'w:ignoreMixedContent',
  'w:alwaysShowPlaceholderText', 'w:doNotDemarcateInvalidXml', 'w:saveXmlDataOnly',
  'w:useXSLTWhenSaving', 'w:saveThroughXslt', 'w:showXMLTags', 'w:alwaysMergeEmptyNamespace',
  'w:updateFields', 'w:hdrShapeDefaults',
  'w:footnotePr', 'w:endnotePr',
  'w:compat', 'w:docVars', 'w:rsids',
  'm:mathPr', 'w:attachedSchema', 'w:themeFontLang', 'w:clrSchemeMapping',
  'w:doNotIncludeSubdocsInStats', 'w:doNotAutoCompressPictures', 'w:forceUpgrade',
  'w:captions', 'w:readModeInkLockDown', 'w:smartTagType', 'sl:schemaLibrary',
  'w:shapeDefaults', 'w:doNotEmbedSmartTags', 'w:decimalSymbol', 'w:listSeparator',
];

/**
 * Inserts a child element into parent at the correct schema index.
 */
export function appendInSchemaOrder(parent: Element, newChild: Element, order: readonly string[]): void {
  const newIndex = order.indexOf(newChild.tagName);
  if (newIndex === -1) {
    parent.appendChild(newChild);
    return;
  }

  for (let i = 0; i < parent.childNodes.length; i++) {
    const sibling = parent.childNodes[i];
    if (sibling.nodeType === 1) {
      const siblingIndex = order.indexOf((sibling as Element).tagName);
      if (siblingIndex > newIndex) {
        parent.insertBefore(newChild, sibling);
        return;
      }
    }
  }

  parent.appendChild(newChild);
}

/**
 * Copies the elements with a tag name out of the live list xmldom returns.
 * Every read of a live list first checks whether the document changed and
 * walks the whole document again if it did, so a loop that edits the document
 * while reading `.length` costs (elements x document size).
 */
function snapshot(doc: Document, tagName: string): Element[] {
  const list = doc.getElementsByTagName(tagName);
  const out: Element[] = [];
  for (let i = 0; i < list.length; i++) out.push(list[i] as unknown as Element);
  return out;
}

export function sortChildrenBySchema(element: Element, order: readonly string[]): void {
  const recognizedIndices: number[] = [];
  const recognizedElements: Element[] = [];
  const allNodes: Node[] = [];

  for (let i = 0; i < element.childNodes.length; i++) {
    const child = element.childNodes[i];
    allNodes.push(child);
    if (child.nodeType === 1) {
      const idx = order.indexOf((child as Element).tagName);
      if (idx !== -1) {
        recognizedIndices.push(i);
        recognizedElements.push(child as Element);
      }
    }
  }

  if (recognizedElements.length <= 1) return;

  // Already in schema order: leave the document untouched. Removing and
  // appending the children again changes nothing in the output but makes
  // xmldom treat the document as modified.
  let alreadySorted = true;
  for (let i = 1; i < recognizedElements.length; i++) {
    if (order.indexOf(recognizedElements[i - 1].tagName) > order.indexOf(recognizedElements[i].tagName)) {
      alreadySorted = false;
      break;
    }
  }
  if (alreadySorted) return;

  recognizedElements.sort((a, b) => {
    return order.indexOf(a.tagName) - order.indexOf(b.tagName);
  });

  // Now replace the nodes in allNodes at recognizedIndices
  for (let i = 0; i < recognizedIndices.length; i++) {
    allNodes[recognizedIndices[i]] = recognizedElements[i];
  }

  // Clear and re-append
  while (element.firstChild) {
    element.removeChild(element.firstChild);
  }
  for (const node of allNodes) {
    element.appendChild(node);
  }
}

/**
 * Generates a valid 8-character hex string for WordML 2010 paraId (ST_LongHexNumber).
 * WordML 2010 requires: 8 hex chars, positive 31-bit integer (0 < val < 0x80000000).
 * When seed is provided, generates a deterministic paraId via SHA-256 digest.
 */
export function generateValidParaId(seed?: string, counter: number = 0): string {
  if (seed !== undefined && seed !== null && seed !== '') {
    const hash = crypto.createHash('sha256').update(`${seed}:${counter}`).digest();
    const raw = hash.readUInt32BE(0) & 0x7fffffff;
    const val = raw === 0 ? 1 : raw;
    return val.toString(16).toUpperCase().padStart(8, '0');
  }
  const val = Math.floor(Math.random() * 0x7fffffff) + 1;
  return val.toString(16).toUpperCase().padStart(8, '0');
}

/**
 * Ensures that a <w:tbl> contains a compliant <w:tblGrid> between <w:tblPr> and the first <w:tr>.
 */
export function ensureTableGrid(tbl: Element, doc: Document): void {
  // Check if tblGrid already exists
  let hasGrid = false;
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const child = tbl.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === 'w:tblGrid') {
      hasGrid = true;
      break;
    }
  }
  if (hasGrid) return;

  // Determine column count from the maximum number of cells in rows
  let maxCols = 0;
  let firstTr: Element | null = null;
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const child = tbl.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === 'w:tr') {
      if (!firstTr) firstTr = child as Element;
      let tcCount = 0;
      for (let j = 0; j < child.childNodes.length; j++) {
        const grandChild = child.childNodes[j];
        if (grandChild.nodeType === 1 && (grandChild as Element).tagName === 'w:tc') {
          tcCount++;
        }
      }
      if (tcCount > maxCols) maxCols = tcCount;
    }
  }

  if (maxCols === 0) maxCols = 2; // Default to 2 columns if empty table

  const tblGrid = createWordElement(doc, 'w:tblGrid');
  const defaultColWidth = Math.round(9360 / maxCols); // 9360 twip = ~6.5 inches printable width
  for (let c = 0; c < maxCols; c++) {
    const gridCol = createWordElement(doc, 'w:gridCol');
    gridCol.setAttribute('w:w', String(defaultColWidth));
    tblGrid.appendChild(gridCol);
  }

  if (firstTr) {
    tbl.insertBefore(tblGrid, firstTr);
  } else {
    tbl.appendChild(tblGrid);
  }
}

/**
 * Recursively canonicalizes all elements in a DOM tree to adhere strictly to ECMA-376.
 */
export function canonicalizeDocumentDom(doc: Document): void {
  // 1. Canonicalize all Run Properties (<w:rPr>)
  for (const el of snapshot(doc, 'w:rPr')) {
    sortChildrenBySchema(el, RPR_CHILD_ORDER);
  }

  // 2. Canonicalize all Paragraph Properties (<w:pPr>)
  for (const el of snapshot(doc, 'w:pPr')) {
    sortChildrenBySchema(el, PPR_CHILD_ORDER);
  }

  // 3. Canonicalize all Table Properties (<w:tblPr>) & Table Borders (<w:tblBorders>)
  for (const el of snapshot(doc, 'w:tblPr')) {
    sortChildrenBySchema(el, TBLPR_CHILD_ORDER);
  }

  for (const el of snapshot(doc, 'w:tblBorders')) {
    sortChildrenBySchema(el, TBLBORDERS_CHILD_ORDER);
  }

  // 4. Canonicalize all Table Cell Properties (<w:tcPr>) & Table Row Properties (<w:trPr>)
  for (const el of snapshot(doc, 'w:trPr')) {
    sortChildrenBySchema(el, TRPR_CHILD_ORDER);
  }

  for (const el of snapshot(doc, 'w:tcPr')) {
    sortChildrenBySchema(el, TCPR_CHILD_ORDER);
  }

  // 5. Ensure Table Cell Margins (<w:tcMar> and <w:tblCellMar>) have required w:type="dxa"
  for (const tagName of ['w:tcMar', 'w:tblCellMar']) {
    for (const mar of snapshot(doc, tagName)) {
      for (let j = 0; j < mar.childNodes.length; j++) {
        const child = mar.childNodes[j];
        if (child.nodeType === 1) {
          const el = child as Element;
          if (!el.getAttribute('w:type')) {
            el.setAttribute('w:type', 'dxa');
          }
        }
      }
    }
  }

  // 6. Ensure all Tables (<w:tbl>) have a valid <w:tblGrid>
  for (const tbl of snapshot(doc, 'w:tbl')) {
    ensureTableGrid(tbl, doc);
  }

  // 7. Normalize invalid justification ('justify' -> 'both') according to ECMA-376 ST_Jc
  for (const jc of snapshot(doc, 'w:jc')) {
    if (jc.getAttribute('w:val') === 'justify') {
      jc.setAttribute('w:val', 'both');
    }
  }

  // 8. Canonicalize all Numbering Properties (<w:numPr>)
  for (const el of snapshot(doc, 'w:numPr')) {
    sortChildrenBySchema(el, NUMPR_CHILD_ORDER);
  }

  // 9. Sanitize internal id leakage in w14:paraId (e.g. 'p_001', 'p_002')
  for (const p of snapshot(doc, 'w:p')) {
    if (p.hasAttribute('w14:paraId')) {
      const val = p.getAttribute('w14:paraId');
      if (val && /^p_\d+$/i.test(val)) {
        p.removeAttribute('w14:paraId');
      }
    }
  }

  // 10. A-2: deduplicate legacy drawing ids (pristine demo.docx ships dup id 1).
  dedupeDrawingDocPrIds(doc);
}

/**
 * A-2 rezip normalization: enforces unique wp:docPr/pic:cNvPr ids within the
 * part (validator rule Sem_UniqueAttributeValue). First occurrence of each id
 * keeps it; later duplicates are renumbered from max+1. All id attributes
 * within one w:drawing (wp:docPr + pic:cNvPr, conventionally equal) are kept
 * in sync. Valid documents are untouched.
 */
export function dedupeDrawingDocPrIds(doc: Document): void {
  // The closest w:drawing above an element owns it. A drawing inside a text box therefore owns its own
  // ids, and the outer drawing does not renumber them again on every pass.
  const nearestDrawing = (el: Element): Element | null => {
    let node: Node | null = el.parentNode;
    while (node && (node as Element).tagName !== 'w:drawing') node = node.parentNode;
    return node as Element | null;
  };
  const idOf = (el: Element): number => {
    const raw = parseInt(el.getAttribute('id') || '0', 10);
    return isNaN(raw) ? 0 : raw;
  };
  // Seed max from every drawing id in the part.
  let maxId = 0;
  for (const tag of ['wp:docPr', 'pic:cNvPr']) {
    const els = doc.getElementsByTagName(tag);
    for (let i = 0; i < els.length; i++) {
      const v = idOf(els[i] as unknown as Element);
      if (v > maxId) maxId = v;
    }
  }
  if (maxId === 0) return;

  const used = new Set<number>();
  for (const drawing of snapshot(doc, 'w:drawing')) {
    const owned: Element[] = [];
    for (const tag of ['wp:docPr', 'pic:cNvPr']) {
      const els = drawing.getElementsByTagName(tag);
      for (let j = 0; j < els.length; j++) {
        const el = els[j] as unknown as Element;
        if (nearestDrawing(el) === drawing) owned.push(el);
      }
    }
    if (owned.length === 0) continue;
    const current = idOf(owned[0]);
    if (current > 0 && !used.has(current)) {
      used.add(current);
      // Keep the drawing's pair in sync when one side drifted.
      for (let k = 1; k < owned.length; k++) {
        if (idOf(owned[k]) !== current) owned[k].setAttribute('id', String(current));
      }
    } else {
      maxId += 1;
      used.add(maxId);
      for (const el of owned) el.setAttribute('id', String(maxId));
    }
  }
}

/**
 * Canonicalizes an XML string (e.g. document.xml, styles.xml, header1.xml, footer1.xml).
 */
export function canonicalizePartXml(xmlContent: string): string {
  if (!xmlContent || !xmlContent.includes('<')) return xmlContent;
  assertNoXxe(xmlContent, 'XML part');

  let hasError = false;
  const parser = new DOMParser({
    errorHandler: {
      warning: () => {},
      error: () => { hasError = true; },
      fatalError: () => { hasError = true; }
    }
  });
  const doc = parser.parseFromString(xmlContent, 'text/xml');
  const parserError = doc.getElementsByTagName('parsererror');
  if (hasError || parserError.length > 0) {
    return xmlContent;
  }
  canonicalizeDocumentDom(doc);

  const serializer = new XMLSerializer();
  return serializer.serializeToString(doc);
}

export const canonicalizeOOXML = canonicalizeDocumentDom;
