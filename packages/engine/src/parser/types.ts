/** Macro/active-content disposition for open + save (refuse-default). */
export interface DocxParserMacroOptions {
  allowMacros?: boolean;
  stripMacros?: boolean;
}

export interface DocxZipAdapter {
  hasFile(path: string): boolean;
  readText(path: string): string;
  readTextOrNull(path: string): string | null;
  writeText(path: string, content: string): void;
  writeBuffer(path: string, content: Buffer): void;
  /** Raw bytes of one part, or null when absent. Not every adapter provides it. */
  readBuffer?(path: string): Buffer | null;
  /** Names of every part (no directories). Not every adapter provides it. */
  entryNames?(): string[];
  /** Removes one part; absent parts are ignored. Not every adapter provides it. */
  removeFile?(path: string): void;
}

export interface TableMetadata {
  headers?: string[];
  tableDimensions?: { rows: number; cols: number };
  caption?: string;
  rows?: unknown[][];
  [key: string]: unknown;
}

export interface StructuralSummary {
  hasTrackedChanges: boolean;
  stylesXml?: string | null;
  /** The theme (`word/theme/theme1.xml`): its accent colours and fonts. */
  themeXml?: string | null;
  /** The numbering definitions (`word/numbering.xml`) that give lists their levels and bullets. */
  numberingXml?: string | null;
  pageSetup?: { rawXml?: string } & Record<string, unknown>;
  headers?: { id: string; type: string; xml: string }[];
  footers?: { id: string; type: string; xml: string }[];
  /** Section flags that decide whether the first-page and even-page headers and footers apply (`w:titlePg`, `w:evenAndOddHeaders`). */
  sectPrFlags?: { titlePg: boolean; evenAndOddHeaders: boolean };
  charts?: { id: string; xml: string }[];
  mediaFiles?: Record<string, { base64: string; mimeType: string }>;
  relationships?: Record<string, { type: string; target: string; targetMode?: 'External' | 'Internal' }>;
  nodes: {
    id: string;
    type: 'paragraph' | 'table' | 'pageBreak' | 'chart' | 'chart3d' | 'image';
    preview: string;
    text: string;
    rawXml?: string;
    charCount: number;
    isHeading: boolean;
    headingLevel?: number;
    hasTrackedChanges: boolean;
    metadata?: TableMetadata | Record<string, unknown>;
  }[];
  /**
   * The page of each node, as a map from node id to page number, with the page geometry that it came from. It is an estimate unless a renderer supplied the map; `pageMapSource` says which.
   * Prefer a map from a real render when you have one.
   */
  nodePages?: Record<string, number>;
  totalPages?: number;
  pageMapSource?: 'renderer' | 'estimate';
  pageMapEstimate?: boolean;
}

export interface DocumentSnapshot {
  id: string;
  timestamp: number;
  label: string;
  xmlContent: string;
  stylesXmlContent?: string | null;
  astChecksum?: string;
}
