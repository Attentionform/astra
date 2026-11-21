import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { OFFICE_VERSIONS as VALIDATOR_VERSIONS } from '@attentionform/ooxml-validator';
import {
  DEFAULT_OFFICE_VERSION,
  OFFICE_VERSIONS,
  OFFICE_VERSION_ENV,
  isOfficeVersion,
  parseOfficeVersion,
  resolveOfficeVersion,
  validateDocxBuffer,
} from '../src';

const OFFICE2010_ONLY = readFileSync(
  fileURLToPath(new URL('../../ooxml-validator/tests/fixtures/office2010-only.docx', import.meta.url)),
);

describe('office version list', () => {
  it('matches the validator package, in the same order', () => {
    expect([...OFFICE_VERSIONS]).toEqual([...VALIDATOR_VERSIONS]);
  });
  it('defaults to Office2019', () => {
    expect(DEFAULT_OFFICE_VERSION).toBe('Office2019');
  });
});

describe('resolveOfficeVersion', () => {
  it('prefers the explicit option over the environment', () => {
    expect(resolveOfficeVersion('Office2010', { [OFFICE_VERSION_ENV]: 'Office2016' })).toBe('Office2010');
  });
  it('uses the environment variable when nothing is passed', () => {
    expect(resolveOfficeVersion(undefined, { [OFFICE_VERSION_ENV]: ' Microsoft365 ' })).toBe('Microsoft365');
  });
  it('falls back to Office2019', () => {
    expect(resolveOfficeVersion(undefined, {})).toBe('Office2019');
    expect(resolveOfficeVersion(undefined, { [OFFICE_VERSION_ENV]: '' })).toBe('Office2019');
  });
  it('never writes the environment', () => {
    const env: NodeJS.ProcessEnv = {};
    resolveOfficeVersion('Office2013', env);
    expect(env).toEqual({});
  });
  it('rejects an invalid explicit value and an invalid variable, naming the seven versions', () => {
    expect(() => resolveOfficeVersion('Office1999' as never, {})).toThrow(
      /Invalid officeVersion "Office1999"\. Expected one of: Office2007, Office2010, Office2013, Office2016, Office2019, Office2021, Microsoft365/,
    );
    expect(() => resolveOfficeVersion(undefined, { [OFFICE_VERSION_ENV]: 'nope' })).toThrow(/Invalid ASTRA_OFFICE_VERSION "nope"/);
  });
  it('isOfficeVersion and parseOfficeVersion agree', () => {
    expect(isOfficeVersion('Office2021')).toBe(true);
    expect(isOfficeVersion('office2021')).toBe(false);
    expect(parseOfficeVersion('Office2021')).toBe('Office2021');
    expect(() => parseOfficeVersion(7, 'flag')).toThrow(/Invalid flag 7/);
  });
});

describe('validateDocxBuffer reports the target it asked for', () => {
  it('a document that uses an Office 2010 feature fails under Office2007 and passes under each later version', async () => {
    // Only the outcomes observed for each named target are asserted.
    const results: Record<string, boolean> = {};
    for (const version of OFFICE_VERSIONS) {
      const res = await validateDocxBuffer(OFFICE2010_ONLY, { officeVersion: version });
      results[version] = res.ok;
      expect(res.requestedOfficeVersion).toBe(version);
      expect(res.office).toBe(version);
      expect(res.validationStatus).toBe('completed');
    }
    expect(results).toEqual({
      Office2007: false,
      Office2010: true,
      Office2013: true,
      Office2016: true,
      Office2019: true,
      Office2021: true,
      Microsoft365: true,
    });
  });
  it('without an option the default target is used and reported', async () => {
    const res = await validateDocxBuffer(OFFICE2010_ONLY);
    expect(res.requestedOfficeVersion).toBe('Office2019');
    expect(res.office).toBe('Office2019');
  });
});
