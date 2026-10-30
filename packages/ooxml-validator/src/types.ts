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
  [key: string]: unknown;
}

export interface ValidateOptions {
  /** Open XML schema version to validate against. Defaults to `Microsoft365` for one-shot, `Office2019` for buffer. */
  officeVersion?: OfficeVersion;
  maxErrors?: number;
  timeoutMs?: number;
  maxDocBytes?: number;
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
  validate(bytes: Buffer | Uint8Array, opts?: ValidateOptions): Promise<ValidationResult>;
  validateFile(path: string, opts?: ValidateOptions): Promise<ValidationResult>;
  close(): Promise<void>;
}
