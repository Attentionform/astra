import type { ValidationResult, ValidateOptions } from '@attentionform/ooxml-validator';

export type { ValidationResult, ValidateOptions };

export interface ValidatorBackend {
  readonly name: 'native';
  validateFile(path: string, opts?: ValidateOptions): Promise<ValidationResult>;
  validateBuffer(buf: Buffer, opts?: ValidateOptions): Promise<ValidationResult>;
}

export type ValidatorBackendName = 'native';

export const nativeBackend: ValidatorBackend = {
  name: 'native',
  async validateFile(path: string, opts?: ValidateOptions): Promise<ValidationResult> {
    const native = await import('@attentionform/ooxml-validator');
    return native.validateFile(path, opts ? { officeVersion: opts.officeVersion } : undefined);
  },
  async validateBuffer(buf: Buffer, opts?: ValidateOptions): Promise<ValidationResult> {
    const native = await import('@attentionform/ooxml-validator');
    return native.validateBuffer(buf, opts ? { officeVersion: opts.officeVersion, maxErrors: opts.maxErrors } : undefined);
  },
};

export function getValidatorBackend(_overrideName?: string): ValidatorBackend {
  void _overrideName;
  return nativeBackend;
}
