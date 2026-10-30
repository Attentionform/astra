
import { XMLSerializer } from '@xmldom/xmldom';

if (typeof Element !== 'undefined' && !((Element.prototype as unknown as Record<string, unknown>)['toJSON'])) {
  (Element.prototype as unknown as { toJSON: (this: Element) => string })['toJSON'] = function (this: Element) {
    try {
      return new XMLSerializer().serializeToString(this);
    } catch {
      return this.nodeName || '[Element]';
    }
  };
}

export interface DocumentMetadata {
  title?: string;
  creator?: string;
  lastModifiedBy?: string;
  revision?: number;
  created?: Date;
  modified?: Date;
  /** E2: Dublin Core subject (dc:subject) + description (dc:description). */
  subject?: string;
  description?: string;
  /** E2: core-properties keywords (cp:keywords) + category (cp:category). */
  keywords?: string;
  category?: string;
  headerText?: string;
  footerText?: string;
}

export interface NodeMetadata {
  text: string;
  hasCitation: boolean;
  hasComment: boolean;
  hasBookmark: boolean;
  hasTrackedChanges: boolean;
  hasEquation: boolean;
  hasField: boolean;
  hasImage: boolean;
  hasHyperlink: boolean;
  listLevel?: number;
  listStyle?: string;
  namedStyle?: string;
  tableDimensions?: { rows: number; cols: number };
  title?: string;
  headers?: string[];
  rows?: string[][];
  chartData?: { chartType?: string } & Record<string, unknown>;
  categories?: string[];
  series?: unknown[];
  chartType?: string;
  /** Callout/admonition marker written by the callout op and read by exporters. */
  callout?: { theme?: string; title?: string; type?: string; alertType?: string };
  /** TOC generator note (minimum labeling when XML comments are stripped). */
  tocPageSource?: string;
  tocPageEstimate?: boolean;
  tocPageNote?: string;
  /** Explicit alert override (wins over theme/title inference in exporters). */
  alertType?: string;
  /** Alias for alertType (flat alias honored by exporters). */
  calloutType?: string;
}

export interface Node {
  id: string;
  type: 'paragraph' | 'table' | 'list' | 'image' | 'equation' | 'field' | 'chart' | 'chart3d';
  xml: Element;
  metadata: NodeMetadata;
  version: number;
  checksum: string;
}

export interface HeadingNode extends Node {
  level: number;
  text: string;
}

export interface PageSetupProperties {
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  headerDistance?: number;
  footerDistance?: number;
  orientation?: 'portrait' | 'landscape';
  paperSize?: 'A4' | 'Letter' | 'Legal' | 'A3';
  columns?: number;
}

export interface Section {
  id: string;
  heading: HeadingNode | null;
  nodes: Node[];
  pageSetup: PageSetupProperties | null;
  version: number;
  checksum: string;
}

// Map of styleId to Style definition XML or properties
export interface StylesMap {
  [styleId: string]: {
    name: string;
    type: 'paragraph' | 'character' | 'table' | 'numbering';
    xml: Element;
  };
}

export interface NumberingMap {
  abstractNum: Record<string, Element>;
  concreteNum: Record<string, Element>;
}

export interface RelationshipsMap {
  [rId: string]: {
    type: string;
    target: string;
    targetMode?: 'External' | 'Internal';
  };
}

export interface HeaderFooterMap {
  [rId: string]: {
    type: 'default' | 'first' | 'even';
    xml: Element;
    text?: string;
  };
}

export interface MediaMap {
  [fileName: string]: Buffer;
}

export interface DocumentSettings {
  trackChanges: boolean;
  defaultAuthor: string;
}

export interface DocumentAST {
  metadata: DocumentMetadata;
  sections: Section[];
  styles: StylesMap;
  numbering: NumberingMap;
  relationships: RelationshipsMap;
  headers: HeaderFooterMap;
  footers: HeaderFooterMap;
  mediaFiles: MediaMap;
  settings: DocumentSettings;
  nextChangeId: number;
  /** ST-09 G1b parity flags (optional; set by parser for renderer parity gating). */
  sectPrFlags?: { titlePg: boolean; evenAndOddHeaders: boolean };
}

/**
 * Utility to calculate a basic SHA-256 equivalent checksum for XML nodes.
 * Used to track drift and version changes.
 */
export async function calculateNodeChecksum(xmlString: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(xmlString);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

/**
 * Deep clones the primary mutation targets of the AST (sections and their XML nodes).
 * Used for atomic rollbacks to preserve the exact DOM state before a batch.
 *
 * BX-02: every shared dictionary map is deep-copied (not just sections), so a
 * headless no-parser rollback cannot leak mutations back into the snapshot via
 * a shared reference. Element values are cloned with cloneNode(true); Buffers
 * are copied; plain records are spread per-entry.
 */
function cloneXmlElement(xml: Element): Element {
  try {
    if (xml && typeof (xml as unknown as { cloneNode?: unknown }).cloneNode === 'function') {
      return (xml as unknown as { cloneNode: (deep: boolean) => Element }).cloneNode(true) as Element;
    }
  } catch { /* fall through — return original ref as last resort */ }
  return xml;
}

export function cloneAST(ast: DocumentAST): DocumentAST {
  const cloneStyles: StylesMap = {};
  for (const [styleId, def] of Object.entries(ast.styles ?? {})) {
    cloneStyles[styleId] = {
      name: def.name,
      type: def.type,
      xml: cloneXmlElement(def.xml),
    };
  }

  const cloneNumbering: NumberingMap = {
    abstractNum: {},
    concreteNum: {},
  };
  for (const [k, el] of Object.entries(ast.numbering?.abstractNum ?? {})) {
    cloneNumbering.abstractNum[k] = cloneXmlElement(el);
  }
  for (const [k, el] of Object.entries(ast.numbering?.concreteNum ?? {})) {
    cloneNumbering.concreteNum[k] = cloneXmlElement(el);
  }

  const cloneRelationships: RelationshipsMap = {};
  for (const [rId, rel] of Object.entries(ast.relationships ?? {})) {
    cloneRelationships[rId] = { ...rel };
  }

  const cloneHeaderFooterMap = (src: HeaderFooterMap): HeaderFooterMap => {
    const out: HeaderFooterMap = {};
    for (const [rId, entry] of Object.entries(src ?? {})) {
      out[rId] = {
        type: entry.type,
        xml: cloneXmlElement(entry.xml),
        ...(entry.text !== undefined ? { text: entry.text } : {}),
      };
    }
    return out;
  };

  const cloneMedia: MediaMap = {};
  for (const [name, buf] of Object.entries(ast.mediaFiles ?? {})) {
    cloneMedia[name] = Buffer.isBuffer(buf) ? Buffer.from(buf) : buf;
  }

  return {
    ...ast,
    metadata: { ...ast.metadata },
    settings: { ...(ast.settings ?? { trackChanges: false, defaultAuthor: 'ASTra' }) },
    nextChangeId: ast.nextChangeId,
    styles: cloneStyles,
    numbering: cloneNumbering,
    relationships: cloneRelationships,
    headers: cloneHeaderFooterMap(ast.headers),
    footers: cloneHeaderFooterMap(ast.footers),
    mediaFiles: cloneMedia,
    sections: ast.sections.map(section => ({
      ...section,
      heading: section.heading
        ? {
            ...section.heading,
            metadata: {
              ...section.heading.metadata,
              tableDimensions: section.heading.metadata.tableDimensions
                ? { ...section.heading.metadata.tableDimensions }
                : undefined,
            },
            xml: cloneXmlElement(section.heading.xml),
          }
        : null,
      pageSetup: section.pageSetup ? { ...section.pageSetup } : null,
      nodes: section.nodes.map(node => ({
        ...node,
        metadata: {
          ...node.metadata,
          tableDimensions: node.metadata.tableDimensions 
            ? { ...node.metadata.tableDimensions } 
            : undefined
        },
        xml: cloneXmlElement(node.xml)
      }))
    }))
  };
}
