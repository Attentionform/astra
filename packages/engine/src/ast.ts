
import { createHash } from 'crypto';
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

/**
 * The document properties: title, author, who changed it last, the dates, the subject and the keywords.
 */
export interface DocumentMetadata {
  title?: string;
  creator?: string;
  lastModifiedBy?: string;
  revision?: number;
  created?: Date;
  modified?: Date;
  /** The Dublin Core subject (`dc:subject`) and description (`dc:description`). */
  subject?: string;
  description?: string;
  /** The keywords (`cp:keywords`) and category (`cp:category`) of the core properties. */
  keywords?: string;
  category?: string;
  headerText?: string;
  footerText?: string;
}

/**
 * What the parser found out about a node: its text, flags that tell whether it has comments, tracked changes, fields and so on, and the data of tables and charts.
 */
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
  /** Set on a callout: its colour theme, its title and its type. */
  callout?: { theme?: string; title?: string; type?: string; alertType?: string };
  /** How the page numbers of a table of contents were found: `tocPageSource` says where they came from, `tocPageEstimate` is true when they are estimates, and `tocPageNote` has a note. */
  tocPageSource?: string;
  tocPageEstimate?: boolean;
  tocPageNote?: string;
  /** An alert type, such as `note` or `warning`, that wins over the type that the exporters work out from the theme and title. */
  alertType?: string;
  /** The same as `alertType`. */
  calloutType?: string;
}

/**
 * One paragraph, table, list, image, equation, field or chart of a document. `id` is stable, `version` rises with each change, and `checksum` tells whether the content changed.
 */
export interface Node {
  id: string;
  type: 'paragraph' | 'table' | 'list' | 'image' | 'equation' | 'field' | 'chart' | 'chart3d';
  xml: Element;
  metadata: NodeMetadata;
  version: number;
  checksum: string;
}

/**
 * A heading: a node with its level and text.
 */
export interface HeadingNode extends Node {
  level: number;
  text: string;
}

/**
 * The page setup of a section, as read from the document. Sizes are in twips (a twentieth of a point).
 */
export interface PageSetupProperties {
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  headerDistance?: number;
  footerDistance?: number;
  orientation?: 'portrait' | 'landscape';
  /** The named paper that the page size matches (see `PAPER_NAMES`), or `Custom` when it matches none. */
  paperSize?: string;
  /** The extra space on the binding edge, in twips. */
  gutter?: number;
  /** Number of text columns; present only when the section has more than one. */
  columns?: number;
  /** The page size exactly as the section's `w:pgSz` gives it (already turned to the orientation in use). */
  pageWidthTwips?: number;
  pageHeightTwips?: number;
}

/**
 * A part of a document under a heading: its `id`, the `heading`, its `nodes` and its page setup.
 */
export interface Section {
  id: string;
  heading: HeadingNode | null;
  nodes: Node[];
  pageSetup: PageSetupProperties | null;
  version: number;
  checksum: string;
}

// Map of styleId to Style definition XML or properties
/**
 * The styles of a document, by style id: each has a name, a type and its XML.
 */
export interface StylesMap {
  [styleId: string]: {
    name: string;
    type: 'paragraph' | 'character' | 'table' | 'numbering';
    xml: Element;
  };
}

/**
 * What the numbering part looked like when it was read, so that it can be written back unchanged: its root attributes, the elements before and after the definitions, and hashes that tell whether the part is still the one that was read.
 */
export interface NumberingSource {
  rootAttributes: Array<{ name: string; value: string }>;
  leading: string[];
  trailing: string[];
  partHash: string;
  definitionsHash: string;
}

/**
 * The numbering definitions of a document, for lists. `source` is what the part looked like when it was read.
 */
export interface NumberingMap {
  abstractNum: Record<string, Element>;
  concreteNum: Record<string, Element>;
  source?: NumberingSource;
}

/**
 * The relationships of the document part, by id: each has a type, a target and whether it is external.
 */
export interface RelationshipsMap {
  [rId: string]: {
    type: string;
    target: string;
    targetMode?: 'External' | 'Internal';
  };
}

/**
 * The headers or the footers of a document, by relationship id. Each has a type (default, first page or even pages), its XML and its text.
 */
export interface HeaderFooterMap {
  [rId: string]: {
    type: 'default' | 'first' | 'even';
    xml: Element;
    text?: string;
  };
}

/**
 * The images and other media of a document, as bytes by file name.
 */
export interface MediaMap {
  [fileName: string]: Buffer;
}

/**
 * The settings that edits use: whether new changes are tracked, and the author name that they carry.
 */
export interface DocumentSettings {
  trackChanges: boolean;
  defaultAuthor: string;
}

/**
 * The document as typed data. `sections` hold the nodes, and the other fields hold the styles, the numbering, the relationships, the headers and footers, the images and the settings. Operations change this structure and `rezip()` writes it back.
 */
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
  /** Section flags that decide whether the first-page and even-page headers and footers apply. */
  sectPrFlags?: { titlePg: boolean; evenAndOddHeaders: boolean };
}

/**
 * SHA-256 of the serialized XML of a node, as lower-case hex. Used to track drift and version changes.
 *
 * Synchronous on purpose: `crypto.subtle.digest` runs every call as a separate job, which costs about 40
 * microseconds a block, 125 ms for a 3,000-paragraph document, and is paid on every document open and every
 * refresh of the model. `createHash` returns the same string in about 1 microsecond a block.
 */
export function calculateNodeChecksumSync(xmlString: string): string {
  return createHash('sha256').update(xmlString, 'utf8').digest('hex');
}

/** Same value as `calculateNodeChecksumSync`; kept async for the callers that already await it. */
export async function calculateNodeChecksum(xmlString: string): Promise<string> {
  return calculateNodeChecksumSync(xmlString);
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
  const numberingSource = ast.numbering?.source;
  if (numberingSource) {
    cloneNumbering.source = {
      ...numberingSource,
      rootAttributes: numberingSource.rootAttributes.map(attribute => ({ ...attribute })),
      leading: [...numberingSource.leading],
      trailing: [...numberingSource.trailing],
    };
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
