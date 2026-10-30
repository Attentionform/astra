import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node, Section, calculateNodeChecksum } from '../ast';
import { BaseOperation, InsertChainHost } from './base';
import { DocxParser } from '../parser';
import { bodySectPr, insertIntoSection, lastSection, ownerDocumentOf } from '../parser/sections';
import type { StyleGuideConfig } from '../style-guide/types';
import { XMLSerializer } from '@xmldom/xmldom';
import * as crypto from 'crypto';
import { resolveReproducibleMode } from '../zip-hardening';
import { sortChildrenBySchema, TBLPR_CHILD_ORDER, TCPR_CHILD_ORDER } from '../ooxml/schema-order';

function isCalloutReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

function calloutHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

export interface CalloutTheme {
  bg: string;
  border: string;
  titleColor: string;
}

const THEMES: Record<string, CalloutTheme> = {
  indigo: { bg: 'EEF2FF', border: '4F46E5', titleColor: '3730A3' },
  emerald: { bg: 'ECFDF5', border: '059669', titleColor: '065F46' },
  amber: { bg: 'FFFBEB', border: 'D97706', titleColor: '92400E' },
  slate: { bg: 'F8FAFC', border: '475569', titleColor: '1E293B' },
  rose: { bg: 'FFF1F2', border: 'E11D48', titleColor: '9F1239' }
};

export interface CalloutOperation extends BaseOperation {
  action: 'callout';
  anchor?: string;
  title?: string;
  text: string;
  theme?: 'indigo' | 'emerald' | 'amber' | 'slate' | 'rose' | string;
  position?: 'before' | 'after';
  params?: Record<string, unknown>;
}

export interface CalloutOperationResult {
  /** Advisory warnings (e.g. unknown theme fallback). Ops still apply; never a failure. */
  warnings: string[];
}

export async function executeCallout(op: CalloutOperation, ast: DocumentAST, parser?: DocxParser, guide?: StyleGuideConfig): Promise<CalloutOperationResult> {
  // OP-16: accept args nested under params:{} or flat (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedAnchor = pick('anchor', op.anchor) as string | undefined;
  const resolvedTitle = pick('title', op.title) as string | undefined;
  const resolvedText = String(pick('text', op.text) ?? op.text ?? '');
  // Theme precedence (explicit > guide > built-in): an explicit theme always
  // wins (even an unknown one — which falls back loudly below, never
  // silently); otherwise the guide default applies; otherwise legacy indigo.
  const explicitTheme = pick('theme', op.theme) as string | undefined;
  const guideDefault = guide?.callouts?.defaultTheme;
  const resolvedTheme = explicitTheme !== undefined
    ? String(explicitTheme)
    : (guideDefault ?? 'indigo');
  const resolvedPosition = (pick('position', op.position) ?? op.position ?? 'after') as 'before' | 'after';
  const themeKey = resolvedTheme.toLowerCase();
  const warnings: string[] = [];
  let theme = THEMES[themeKey];
  if (!theme) {
    const fallbackKey = (guideDefault ?? 'indigo').toLowerCase();
    theme = THEMES[fallbackKey] ?? THEMES.indigo;
    const effectiveFallback = THEMES[fallbackKey] ? fallbackKey : 'indigo';
    warnings.push(
      `Unknown callout theme '${resolvedTheme}'; fell back to '${effectiveFallback}'. ` +
      `Remedy: use one of indigo, emerald, amber, slate, rose` +
      (guideDefault ? ` (guide default '${guideDefault}')` : '') + '.',
    );
  }

  // With no anchor the callout goes to the end of the document: the last section.
  let targetSection = lastSection(ast.sections);
  let insertIndex = targetSection ? targetSection.nodes.length : 0;
  let anchorNode: Node | undefined;

  if (resolvedAnchor) {
    for (const section of ast.sections) {
      const idx = section.nodes.findIndex(n => n.id === resolvedAnchor);
      if (idx !== -1) {
        targetSection = section;
        insertIndex = resolvedPosition === 'before' ? idx : idx + 1;
        anchorNode = section.nodes[idx];
        break;
      }
    }
  }

  if (!targetSection) throw new Error('No section available for callout box insertion');

  // Find owner document from an existing node xml
  const ownerDoc = targetSection.nodes[0]?.xml?.ownerDocument || ownerDocumentOf(ast.sections);
  if (!ownerDoc) throw new Error('Cannot resolve XML DOM document');

  const calloutId = isCalloutReproducible()
    ? (() => {
        const base = `callout_${calloutHash8(`${resolvedAnchor ?? ''}:${resolvedPosition}:${themeKey}:${resolvedTitle ?? ''}:${resolvedText}`)}`;
        const live = new Set<string>();
        try {
          for (const s of ast.sections ?? []) {
            for (const n of s.nodes ?? []) {
              if (typeof n.id === 'string' && n.id) live.add(n.id);
            }
          }
        } catch { /* best-effort */ }
        let candidate = base;
        let n = 2;
        while (live.has(candidate)) candidate = `${base}_${n++}`;
        return candidate;
      })()
    : `callout_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;

  // Construct ECMA-376 single-cell table container
  const tbl = createWordElement(ownerDoc, 'w:tbl');
  tbl.setAttribute('data-astra-id', calloutId);

  // Table Properties
  const tblPr = createWordElement(ownerDoc, 'w:tblPr');
  const tblW = createWordElement(ownerDoc, 'w:tblW');
  tblW.setAttribute('w:w', '5000');
  tblW.setAttribute('w:type', 'pct');
  tblPr.appendChild(tblW);

  // Borders: thick left accent bar, others none
  const tblBorders = createWordElement(ownerDoc, 'w:tblBorders');
  
  const topB = createWordElement(ownerDoc, 'w:top');
  topB.setAttribute('w:val', 'none');
  tblBorders.appendChild(topB);

  const leftB = createWordElement(ownerDoc, 'w:left');
  leftB.setAttribute('w:val', 'single');
  leftB.setAttribute('w:sz', '32'); // 4pt
  leftB.setAttribute('w:space', '0');
  leftB.setAttribute('w:color', theme.border);
  tblBorders.appendChild(leftB);

  const bottomB = createWordElement(ownerDoc, 'w:bottom');
  bottomB.setAttribute('w:val', 'none');
  tblBorders.appendChild(bottomB);

  const rightB = createWordElement(ownerDoc, 'w:right');
  rightB.setAttribute('w:val', 'none');
  tblBorders.appendChild(rightB);

  tblPr.appendChild(tblBorders);

  // Cell padding margins
  // F2: ECMA-376 CT_TblCellMar sequence is top,left,bottom,right
  // (was top,bottom,left,right — schema-invalid for all themes).
  const tblCellMar = createWordElement(ownerDoc, 'w:tblCellMar');
  const topM = createWordElement(ownerDoc, 'w:top');
  topM.setAttribute('w:w', '180');
  topM.setAttribute('w:type', 'dxa');
  tblCellMar.appendChild(topM);
  const leftM = createWordElement(ownerDoc, 'w:left');
  leftM.setAttribute('w:w', '240');
  leftM.setAttribute('w:type', 'dxa');
  tblCellMar.appendChild(leftM);
  const bottomM = createWordElement(ownerDoc, 'w:bottom');
  bottomM.setAttribute('w:w', '180');
  bottomM.setAttribute('w:type', 'dxa');
  tblCellMar.appendChild(bottomM);
  const rightM = createWordElement(ownerDoc, 'w:right');
  rightM.setAttribute('w:w', '240');
  rightM.setAttribute('w:type', 'dxa');
  tblCellMar.appendChild(rightM);
  tblPr.appendChild(tblCellMar);

  sortChildrenBySchema(tblPr, TBLPR_CHILD_ORDER);
  tbl.appendChild(tblPr);

  // Table Grid
  const tblGrid = createWordElement(ownerDoc, 'w:tblGrid');
  const gridCol = createWordElement(ownerDoc, 'w:gridCol');
  gridCol.setAttribute('w:w', '9000');
  tblGrid.appendChild(gridCol);
  tbl.appendChild(tblGrid);

  // Row and Cell (row never splits across pages, mirrors table inserts)
  const tr = createWordElement(ownerDoc, 'w:tr');
  const trPr = createWordElement(ownerDoc, 'w:trPr');
  trPr.appendChild(createWordElement(ownerDoc, 'w:cantSplit'));
  tr.appendChild(trPr);
  const tc = createWordElement(ownerDoc, 'w:tc');
  const tcPr = createWordElement(ownerDoc, 'w:tcPr');
  
  const shd = createWordElement(ownerDoc, 'w:shd');
  shd.setAttribute('w:val', 'clear');
  shd.setAttribute('w:color', 'auto');
  shd.setAttribute('w:fill', theme.bg);
  tcPr.appendChild(shd);
  sortChildrenBySchema(tcPr, TCPR_CHILD_ORDER);
  tc.appendChild(tcPr);

  // Optional Title Paragraph
  if (resolvedTitle && resolvedTitle.trim()) {
    const titleP = createWordElement(ownerDoc, 'w:p');
    const titlePPr = createWordElement(ownerDoc, 'w:pPr');
    const titleSp = createWordElement(ownerDoc, 'w:spacing');
    titleSp.setAttribute('w:after', '80');
    titlePPr.appendChild(titleSp);
    titleP.appendChild(titlePPr);

    const titleR = createWordElement(ownerDoc, 'w:r');
    const titleRPr = createWordElement(ownerDoc, 'w:rPr');
    titleRPr.appendChild(createWordElement(ownerDoc, 'w:b'));
    const titleColor = createWordElement(ownerDoc, 'w:color');
    titleColor.setAttribute('w:val', theme.titleColor);
    titleRPr.appendChild(titleColor);
    titleR.appendChild(titleRPr);

    const titleT = createWordElement(ownerDoc, 'w:t');
    titleT.textContent = resolvedTitle.trim();
    titleR.appendChild(titleT);
    titleP.appendChild(titleR);
    tc.appendChild(titleP);
  }

  // Body Text Paragraph
  const bodyP = createWordElement(ownerDoc, 'w:p');
  const bodyPPr = createWordElement(ownerDoc, 'w:pPr');
  const bodySp = createWordElement(ownerDoc, 'w:spacing');
  bodySp.setAttribute('w:after', '0');
  bodyPPr.appendChild(bodySp);
  bodyP.appendChild(bodyPPr);

  const bodyR = createWordElement(ownerDoc, 'w:r');
  const bodyRPr = createWordElement(ownerDoc, 'w:rPr');
  const bodyColor = createWordElement(ownerDoc, 'w:color');
  bodyColor.setAttribute('w:val', '334155');
  bodyRPr.appendChild(bodyColor);
  bodyR.appendChild(bodyRPr);

  const bodyT = createWordElement(ownerDoc, 'w:t');
  bodyT.textContent = resolvedText;
  bodyR.appendChild(bodyT);
  bodyP.appendChild(bodyR);
  tc.appendChild(bodyP);

  tr.appendChild(tc);
  tbl.appendChild(tr);

  // Insert into AST and DOM
  const newNode: Node = {
    id: calloutId,
    type: 'table',
    xml: tbl,
    metadata: {
      text: `${resolvedTitle ? resolvedTitle + ': ' : ''}${resolvedText}`,
      callout: { theme: themeKey, title: resolvedTitle },
      hasCitation: false,
      hasComment: false,
      hasBookmark: false,
      hasTrackedChanges: false,
      hasEquation: false,
      hasField: false,
      hasImage: false,
      hasHyperlink: false,
    },
    version: 1,
    checksum: await calculateNodeChecksum(new XMLSerializer().serializeToString(tbl))
  };

  // Insert into AST and DOM. Same-anchor 'after' inserts preserve FIFO
  // order via a per-anchor chain (mirrors insert.ts): re-validated
  // against live DOM/AST at use, dead links dropped so the next insert
  // lands after the latest live link, never after a detached ghost.
  // 'before' stays a direct insert.
  if (anchorNode && anchorNode.xml && anchorNode.xml.parentNode && resolvedPosition === 'after') {
    const host = anchorNode as unknown as InsertChainHost;
    const rawChain: Node[] = Array.isArray(host._insertAfterChain)
      ? (host._insertAfterChain as Node[])
      : [];
    const legacyTail = host._lastInsertAfter as Node | undefined;
    if (rawChain.length === 0 && legacyTail) {
      rawChain.push(legacyTail);
    }
    const liveChain = rawChain.filter((n) => {
      if (!n || !n.xml || !(n.xml as Element).parentNode) return false;
      return ast.sections.some((s: Section) => s.nodes.some((x: Node) => x.id === n.id));
    });
    host._insertAfterChain = liveChain;
    if (legacyTail && !liveChain.some((n) => n.id === legacyTail.id)) {
      delete host._lastInsertAfter;
    }
    const effectiveAnchor: Node = liveChain.length > 0 ? liveChain[liveChain.length - 1] : anchorNode;
    let effSection: Section = targetSection;
    let effIdx = targetSection.nodes.findIndex((x: Node) => x.id === effectiveAnchor.id);
    if (effIdx === -1) {
      for (const s of ast.sections) {
        const i = (s.nodes as Node[]).findIndex((n: Node) => n.id === effectiveAnchor.id);
        if (i !== -1) {
          effSection = s;
          effIdx = i;
          break;
        }
      }
    }
    if (effIdx !== -1) {
      insertIntoSection(ast.sections, effSection, effIdx + 1, [newNode]).version += 1;
    } else {
      insertIntoSection(ast.sections, targetSection, insertIndex, [newNode]).version += 1;
    }
    const effXml = effectiveAnchor.xml as Element;
    effXml.parentNode?.insertBefore(tbl, effXml.nextSibling);
    liveChain.push(newNode);
    host._insertAfterChain = liveChain;
    host._lastInsertAfter = newNode;
  } else {
    insertIntoSection(ast.sections, targetSection, insertIndex, [newNode]).version += 1;

    if (anchorNode && anchorNode.xml && anchorNode.xml.parentNode) {
      anchorNode.xml.parentNode.insertBefore(tbl, anchorNode.xml);
    } else {
      const body = ownerDoc.getElementsByTagName('w:body')[0];
      if (body) {
        // The body's own sectPr, not the first one in the document: a section
        // break paragraph carries one too, and it is not a child of the body.
        body.insertBefore(tbl, bodySectPr(body));
      }
    }
  }

  if (parser) {
    parser.nodeMap?.set(calloutId, tbl);
  }
  return { warnings };
}
