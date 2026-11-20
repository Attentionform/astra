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

/** An Office version that validation can target. */
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

export type ValidationStatus = 'completed' | 'skipped' | 'timed-out' | 'failed-to-start';

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

/** The result of one validation. `ok` is true when the validator reported no errors. */
export interface ValidationResult {
  /** The file that was checked, or a label for the bytes that were checked. */
  file: string;
  /** True when no error was reported. */
  ok: boolean;
  /** The errors that were reported, at most `maxErrors` of them. */
  errors: ValidationError[];
  /** The version of the Open XML SDK that ran. */
  sdkVersion?: string;
  /** The Office version that the document was checked against. */
  office?: string;
  /** How long the run took, in milliseconds. */
  durationMs?: number;
  /** True when `maxErrors` cut the list of errors. */
  truncated?: boolean;
  /** The number of errors that were found. */
  errorCount?: number;
  /** The Office version the caller asked for (always set by the buffer and worker APIs). */
  requestedOfficeVersion?: string;
  /**
   * `completed` when the validator ran and reported. Callers that skip the check, give up waiting
   * for it, or cannot start it report `skipped`, `timed-out` or `failed-to-start` and leave `office` unset.
   */
  validationStatus?: ValidationStatus;
  /** The document family the package was opened as (buffer and worker APIs). */
  kind?: DocumentKind;
  [key: string]: unknown;
}

/**
 * Options for a validation. The document family comes from the file extension, so there is no `kind` here;
 * `ValidateBufferOptions` adds it for bytes.
 */
export interface ValidateOptions {
  /** The Office version to validate against. The default is `Microsoft365` for a file and `Office2019` for bytes. */
  officeVersion?: OfficeVersion;
  /** Stop reporting after this many errors. The result then has `truncated: true`. */
  maxErrors?: number;
  /** Give up after this many milliseconds. */
  timeoutMs?: number;
  /** Refuse a document that is larger than this many bytes. */
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
