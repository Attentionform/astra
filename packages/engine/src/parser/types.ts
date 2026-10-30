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
  /** ST-04: word/theme/theme1.xml payload (accent map + fonts) for theme resolution. */
  themeXml?: string | null;
  /** ST-05: word/numbering.xml payload (abstractNum→lvl) for list rendering. */
  numberingXml?: string | null;
  pageSetup?: { rawXml?: string } & Record<string, unknown>;
  headers?: { id: string; type: string; xml: string }[];
  footers?: { id: string; type: string; xml: string }[];
  /** ST-09 G1b: sectPr parity flags gating first/even use (w:titlePg, w:evenAndOddHeaders). */
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
   * IN-07: renderer-compatible node→page map embedded at summary time, plus
   * the page geometry params it was derived from. The map uses the same
   * explicit-break semantics as the canvas paginator; line-density fallback
   * is geometry-scaled (not a fixed 45). Consumers (visual_review) must
   * prefer a live renderer map when available; this embedded map carries
   * `pageMapSource: 'estimate'` unless a renderer map was supplied.
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
