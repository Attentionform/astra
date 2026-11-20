// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

export const OFFICE_VERSIONS = [
  'Office2007',
  'Office2010',
  'Office2013',
  'Office2016',
  'Office2019',
  'Office2021',
  'Microsoft365',
] as const;

export type OfficeVersion = (typeof OFFICE_VERSIONS)[number];

/** Document families the validator opens: Word, PowerPoint and Excel packages. */
export const DOCUMENT_KINDS = ['docx', 'pptx', 'xlsx'] as const;

export type DocumentKind = (typeof DOCUMENT_KINDS)[number];

/**
 * File extension (lower case, with the dot) to document family. The native
 * file validator and `createValidator().validateFile()` choose the family from
 * the extension with this table.
 */
export const EXTENSION_KINDS: Readonly<Record<string, DocumentKind>> = {
  '.docx': 'docx',
  '.docm': 'docx',
  '.dotx': 'docx',
  '.dotm': 'docx',
  '.pptx': 'pptx',
  '.pptm': 'pptx',
  '.potx': 'pptx',
  '.potm': 'pptx',
  '.ppsx': 'pptx',
  '.ppsm': 'pptx',
  '.xlsx': 'xlsx',
  '.xlsm': 'xlsx',
  '.xltx': 'xlsx',
  '.xltm': 'xlsx',
  '.xlam': 'xlsx',
};

export interface ErrorNode {
  localName?: string | null;
  namespaceUri?: string | null;
}

export type ValidationErrorClass =
  | 'schema'
  | 'semantic'
  | 'markup-compat'
  | 'package'
  | 'exception'
  | string;

export interface ValidationError {
  /** Human-readable description of the problem. */
  description?: string | null;
  /** Package part the error was found in, e.g. `/word/document.xml`. `null` for `errorType: 'Exception'`. */
  path?: string | null;
  /** XPath of the offending element. `null` for `errorType: 'Exception'`. */
  xPath?: string | null;
  /** Open XML SDK error id, e.g. `Sch_UnexpectedElementContentExpectingComplex`. `null` for `errorType: 'Exception'`. */
  id?: string | null;
  /** `Schema`, `Semantic`, `Package`, `MarkupCompatibility`, or `Exception` when the file could not be opened at all. */
  errorType?: string | null;
  /** Part URI string (e.g. `/word/document.xml`). */
  part?: string | null;
  /** Error classification: 'schema' | 'semantic' | 'markup-compat' | 'package' | 'exception'. */
  class?: ValidationErrorClass | null;
  /** Offending XML element information when present. */
  node?: ErrorNode | null;
  /** Related XML element information when present. */
  relatedNode?: ErrorNode | null;
  [key: string]: unknown;
}

export interface ValidationResult {
  file: string;
  ok: boolean;
  errors: ValidationError[];
  sdkVersion?: string;
  office?: string;
  durationMs?: number;
  truncated?: boolean;
  errorCount?: number;
  /** The Office version the caller asked for (always set by the buffer and worker APIs). */
  requestedOfficeVersion?: string;
  /** The document family the package was opened as (buffer and worker APIs). */
  kind?: DocumentKind;
  [key: string]: unknown;
}

/**
 * Options for the file APIs: the standalone `validateFile()`, `isValid()` and
 * the CLI. The family comes from the file extension, so there is no `kind`.
 */
export interface ValidateOptions {
  /** Open XML schema version to validate against. Defaults to `Microsoft365` for one-shot, `Office2019` for buffer. */
  officeVersion?: OfficeVersion;
  maxErrors?: number;
  timeoutMs?: number;
  maxDocBytes?: number;
}

/** Options for the buffer and worker APIs. */
export interface ValidateBufferOptions extends ValidateOptions {
  /**
   * Document family of the bytes. Defaults to `docx` for a buffer. For
   * `createValidator().validateFile()` the default is inferred from the file
   * extension and an explicit `kind` overrides it.
   */
  kind?: DocumentKind;
}

export interface ValidatorPoolOptions {
  workers?: number;
  recycleAfter?: number;
  timeoutMs?: number;
  maxDocBytes?: number;
  maxErrors?: number;
}

export interface BackendInfo {
  platformPackage: string;
  binaryPath: string;
  sdkVersion: string;
  binarySha256: string;
  workerSupported: boolean;
}

export interface ValidatorInstance {
  validate(bytes: Buffer | Uint8Array, opts?: ValidateBufferOptions): Promise<ValidationResult>;
  validateFile(path: string, opts?: ValidateBufferOptions): Promise<ValidationResult>;
  close(): Promise<void>;
}
